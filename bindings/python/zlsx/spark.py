"""PySpark Data Source for xlsx workbooks, backed by zlsx.

Batch read and batch write, Spark 4.0+ / DBR 15.4+ (including
serverless). Install with ``pip install py-zlsx[spark]``; on Databricks
the pyspark half is already there — the plain wheel suffices.

.. code-block:: python

    from zlsx.spark import ZlsxDataSource
    spark.dataSource.register(ZlsxDataSource)

    df = (spark.read.format("zlsx")
          .option("sheet", "Sales")          # name, index, "a,b", or "*"
          .option("rowsPerPartition", 100000)
          .load("/Volumes/cat/schema/vol/*.xlsx"))

    (df.coalesce(1).write.format("zlsx")
       .mode("overwrite")
       .option("sheet", "Report")
       .save("/Volumes/cat/schema/vol/report.xlsx"))

Read options: ``sheet`` (default ``"0"``), ``header`` (default true),
``inferRows`` sample size (default 1000), ``rowsPerPartition`` (default
0 = one partition per file x sheet), ``mode`` (``permissive`` nulls
cells that don't fit the schema, ``failfast`` raises naming the exact
file/sheet/row/column). Write options: ``sheet`` (default "Sheet1"),
``header`` (default true). A ``.xlsx`` target path means single-file
mode and requires a single-partition DataFrame (``coalesce(1)``); any
other path is a directory of ``part-*.xlsx`` files.

Partition planning, schema inference and coercion live in
:mod:`zlsx._tabular` (pure Python, unit-tested without Spark).
"""
from __future__ import annotations

import os
import uuid
from dataclasses import dataclass
from itertools import islice
from typing import Optional

from pyspark.sql.datasource import (
    DataSource,
    DataSourceReader,
    DataSourceWriter,
    InputPartition,
    WriterCommitMessage,
)

from . import _tabular

_READ_KINDS = {
    "boolean": "boolean",
    "tinyint": "bigint",
    "smallint": "bigint",
    "int": "bigint",
    "bigint": "bigint",
    "float": "double",
    "double": "double",
    "string": "string",
    "date": "date",
    "timestamp": "timestamp",
    "timestamp_ntz": "timestamp",
}

_WRITE_OK = set(_READ_KINDS)  # + decimal(p,s), checked by prefix


def _kinds_from_schema(schema, *, what: str) -> list[str]:
    kinds = []
    for f in schema.fields:
        s = f.dataType.simpleString()
        if s in _READ_KINDS:
            kinds.append(_READ_KINDS[s])
        elif s.startswith("decimal(") and what == "write":
            kinds.append("double")
        else:
            raise TypeError(
                f"zlsx cannot {what} column {f.name!r} of type {s}; "
                "supported: boolean, integral, float/double, string, "
                "date, timestamp"
                + (", decimal (written as double)" if what == "write" else "")
            )
    return kinds


def _truthy(options, key: str, default: bool) -> bool:
    return options.get(key, str(default)).strip().lower() in ("true", "1", "yes")


class ZlsxPartition(InputPartition):
    def __init__(self, path: str, sheet_idx: int, sheet_name: str,
                 start: int = 0, end: Optional[int] = None):
        self.path = path
        self.sheet_idx = sheet_idx
        self.sheet_name = sheet_name
        self.start = start
        self.end = end  # None = whole sheet (bulk read path)


class ZlsxDataSource(DataSource):
    """``spark.read.format("zlsx")`` / ``df.write.format("zlsx")``."""

    @classmethod
    def name(cls):
        return "zlsx"

    def schema(self):
        # Only called when the user did not supply a schema. Inference
        # samples inferRows data rows from the FIRST resolved
        # (file, sheet) and widens across the whole sample; with
        # sheet="*" or multi-file input the remaining sheets must be
        # column-compatible (coercion absorbs benign type drift).
        import zlsx

        opts = self.options
        files = _tabular.expand_paths(opts["path"])
        header = _truthy(opts, "header", True)
        infer_rows = int(opts.get("inferRows", "1000"))
        with zlsx.open(files[0]) as book:
            sheet_idx = _tabular.resolve_sheets(
                opts.get("sheet", "0"), book.sheets
            )[0]
            with book.sheet(sheet_idx).rows() as rows:
                it = iter(rows)
                head = next(it, None) if header else None
                sample = list(islice(it, infer_rows))
        if header and head is None:
            raise ValueError(
                f"{files[0]}: sheet {sheet_idx} is empty; cannot infer a "
                "schema (supply one with .schema(...) or set header=false)"
            )
        return _tabular.ddl(_tabular.infer_fields(head, sample))

    def reader(self, schema):
        return ZlsxBatchReader(self.options, schema)

    def writer(self, schema, overwrite):
        return ZlsxBatchWriter(self.options, schema, overwrite)


class ZlsxBatchReader(DataSourceReader):
    def __init__(self, options, schema):
        import zlsx

        self._kinds = _kinds_from_schema(schema, what="read")
        self._header = _truthy(options, "header", True)
        self._mode = options.get("mode", "permissive").strip().lower()
        if self._mode not in ("permissive", "failfast"):
            raise ValueError(
                f"mode must be 'permissive' or 'failfast', got {self._mode!r}"
            )
        rows_per_part = int(options.get("rowsPerPartition", "0"))

        # Driver-side planning: one partition per (file x sheet); with
        # rowsPerPartition also split each sheet into row ranges, which
        # costs one counting pass per sheet here (zlsx parses at C
        # speed — the win on executors dwarfs it).
        self._partitions: list[ZlsxPartition] = []
        for path in _tabular.expand_paths(options["path"]):
            with zlsx.open(path) as book:
                for si in _tabular.resolve_sheets(
                    options.get("sheet", "0"), book.sheets
                ):
                    name = book.sheets[si]
                    if rows_per_part <= 0:
                        self._partitions.append(ZlsxPartition(path, si, name))
                        continue
                    with book.sheet(si).rows() as rows:
                        n = sum(1 for _ in rows)
                    n_data = max(0, n - 1) if self._header else n
                    for start, end in _tabular.plan_row_ranges(
                        n_data, rows_per_part
                    ):
                        self._partitions.append(
                            ZlsxPartition(path, si, name, start, end)
                        )

    def partitions(self):
        return self._partitions

    def read(self, partition):
        import zlsx  # executor-side import

        p = partition
        ctx = f"{p.path} sheet {p.sheet_name!r}"
        kinds, mode, header = self._kinds, self._mode, self._header
        # 1-based Excel display row of the first data row in this range.
        base = p.start + (2 if header else 1)
        with zlsx.open(p.path) as book:
            sheet = book.sheet(p.sheet_idx)
            if p.end is None:
                # Whole sheet: the bulk-FFI matrix path.
                _, data = sheet.read_all(header=header)
                for i, row in enumerate(data):
                    yield tuple(
                        _tabular.coerce_row(row, kinds, mode, ctx, base + i)
                    )
            else:
                # Row range: stream, skipping [0, start) cheaply.
                with sheet.rows() as rows:
                    it = iter(rows)
                    if header:
                        next(it, None)
                    for i, row in enumerate(islice(it, p.start, p.end)):
                        yield tuple(
                            _tabular.coerce_row(row, kinds, mode, ctx, base + i)
                        )


@dataclass
class ZlsxCommit(WriterCommitMessage):
    path: Optional[str]
    rows: int


class ZlsxBatchWriter(DataSourceWriter):
    def __init__(self, options, schema, overwrite: bool):
        self._path = options["path"]
        self._sheet = options.get("sheet", "Sheet1")
        self._header = _truthy(options, "header", True)
        self._names = [f.name for f in schema.fields]
        self._kinds = _kinds_from_schema(schema, what="write")
        self._single = self._path.endswith((".xlsx", ".xlsm"))

        # Driver-side target preparation, before any task launches.
        if self._single:
            if os.path.exists(self._path):
                if not overwrite:
                    raise FileExistsError(
                        f"{self._path} exists; use mode('overwrite') "
                        "(zlsx does not append into an existing workbook)"
                    )
                os.remove(self._path)
        else:
            os.makedirs(self._path, exist_ok=True)
            if overwrite:
                for f in os.listdir(self._path):
                    if f.startswith("part-") and f.endswith((".xlsx", ".xlsm")):
                        os.remove(os.path.join(self._path, f))

    def write(self, iterator):
        import zlsx  # executor-side import
        from pyspark import TaskContext

        pid = TaskContext.get().partitionId()
        it = iter(iterator)
        first = next(it, None)
        if first is None:
            # Empty partition: single-file mode still materialises the
            # workbook (deterministic artifact, header row only) from
            # partition 0; directory mode just skips the part file.
            if not (self._single and pid == 0):
                return ZlsxCommit(path=None, rows=0)

        if self._single:
            if pid != 0 and first is not None:
                raise ValueError(
                    f"single-file target {self._path} needs a "
                    "single-partition DataFrame: call .coalesce(1) or "
                    "write to a directory instead"
                )
            out_path = self._path
        else:
            out_path = os.path.join(
                self._path, f"part-{pid:05d}-{uuid.uuid4().hex[:8]}.xlsx"
            )

        date_fmt = {"date": "yyyy-mm-dd", "timestamp": "yyyy-mm-dd hh:mm:ss"}
        n = 0
        with zlsx.write(out_path) as w:
            styles = [
                w.add_style(zlsx.Style(number_format=date_fmt[k]))
                if k in date_fmt else 0
                for k in self._kinds
            ]
            header_style = w.add_style(zlsx.Style(font_bold=True))
            sheet = w.add_sheet(self._sheet)
            if self._header:
                sheet.write_row(
                    self._names, styles=[header_style] * len(self._names)
                )
            rows_iter = it if first is None else _chain_one(first, it)
            for row in rows_iter:
                sheet.write_row(
                    [_to_cell(v) for v in row], styles=styles
                )
                n += 1
        return ZlsxCommit(path=out_path, rows=n)

    def commit(self, messages):
        pass  # part files are already at their final paths

    def abort(self, messages):
        for m in messages:
            if m is not None and m.path:
                try:
                    os.remove(m.path)
                except OSError:
                    pass


def _chain_one(first, rest):
    yield first
    yield from rest


def _to_cell(v):
    import datetime as dt
    import decimal

    import zlsx

    if v is None or isinstance(v, (bool, int, float, str)):
        return v
    if isinstance(v, (dt.datetime, dt.date)):
        return zlsx.to_excel_serial(v)
    if isinstance(v, decimal.Decimal):
        return float(v)
    raise TypeError(f"cannot write {type(v).__name__} to an xlsx cell")


def register(spark) -> None:
    """Convenience: ``zlsx.spark.register(spark)``."""
    spark.dataSource.register(ZlsxDataSource)
