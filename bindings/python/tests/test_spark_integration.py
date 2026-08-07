"""Integration tests for zlsx.spark against a real local SparkSession.

Skipped wholesale when pyspark is not installed or a local JVM cannot
start (Spark 4 needs Java 17/21). The pure logic underneath is covered
without Spark in test_spark_core.py; the serverless proof lives in
integrations/databricks/.
"""
from __future__ import annotations

import datetime as dt
import os

import pytest

pyspark = pytest.importorskip("pyspark")

import zlsx  # noqa: E402
from zlsx.spark import ZlsxDataSource  # noqa: E402


@pytest.fixture(scope="module")
def spark():
    from pyspark.sql import SparkSession

    try:
        s = (
            SparkSession.builder.master("local[2]")
            .appName("zlsx-spark-tests")
            .config("spark.ui.enabled", "false")
            .getOrCreate()
        )
    except Exception as e:  # noqa: BLE001 — JVM/Java-version failures
        pytest.skip(f"local Spark unavailable: {e}")
    s.dataSource.register(ZlsxDataSource)
    yield s
    s.stop()


@pytest.fixture(scope="module")
def workbook(tmp_path_factory):
    """Two-sheet fixture. `revenue` is int on the first data row and
    float later — the exact case first-row-only inference gets wrong —
    and the last Sales row is ragged (no revenue cell)."""
    path = str(tmp_path_factory.mktemp("wb") / "fixture.xlsx")
    with zlsx.write(path) as w:
        for sheet_name, rows in (
            ("Sales", [
                ["AMER", "widget", 10, 100],
                ["AMER", "flange", 20, 200.5],
                ["EMEA", "widget", 30, 300.25],
                ["EMEA", "gizmo", 40],
            ]),
            ("Extra", [
                ["APAC", "widget", 50, 500.0],
            ]),
        ):
            sh = w.add_sheet(sheet_name)
            sh.write_row(["region", "product", "units", "revenue"])
            for r in rows:
                sh.write_row(r)
    return path


def test_infer_widens_beyond_first_row(spark, workbook):
    df = spark.read.format("zlsx").option("sheet", "Sales").load(workbook)
    schema = {f.name: f.dataType.simpleString() for f in df.schema.fields}
    assert schema == {
        "region": "string", "product": "string",
        "units": "bigint", "revenue": "double",
    }


def test_read_values_and_ragged_row(spark, workbook):
    df = spark.read.format("zlsx").option("sheet", "Sales").load(workbook)
    rows = {(r.region, r.product): (r.units, r.revenue) for r in df.collect()}
    assert len(rows) == 4
    assert rows[("AMER", "widget")] == (10, 100.0)  # int upcast to double
    assert rows[("EMEA", "gizmo")] == (40, None)    # ragged row padded


def test_multi_sheet_star(spark, workbook):
    df = spark.read.format("zlsx").option("sheet", "*").load(workbook)
    assert df.count() == 5
    assert df.rdd.getNumPartitions() == 2  # one per sheet


def test_rows_per_partition_splits_without_loss(spark, workbook):
    df = (spark.read.format("zlsx")
          .option("sheet", "Sales")
          .option("rowsPerPartition", "2")
          .load(workbook))
    assert df.rdd.getNumPartitions() == 2
    assert df.count() == 4
    assert {r.product for r in df.collect()} == {"widget", "flange", "gizmo"}


def test_many_row_range_partitions_lose_and_duplicate_nothing(spark, tmp_path):
    """Every range partition past the first starts with a skip, so an
    off-by-one there drops or repeats rows at partition boundaries
    instead of failing loudly. Enough rows and partitions to catch it,
    including a final short partition."""
    path = str(tmp_path / "wide.xlsx")
    n = 97                      # prime: guarantees a ragged last partition
    with zlsx.write(path) as w:
        s = w.add_sheet("S")
        s.write_row(["idx", "label"])
        for i in range(n):
            s.write_row([i, f"row{i}"])

    df = (spark.read.format("zlsx")
          .option("rowsPerPartition", "10")
          .load(path))
    assert df.rdd.getNumPartitions() == 10          # ceil(97/10)

    got = sorted(r.idx for r in df.collect())
    assert got == list(range(n)), "rows lost or duplicated across boundaries"
    # Labels must stay paired with their index — a skip that lands one
    # row off would keep the count but shear the columns.
    assert all(r.label == f"row{r.idx}" for r in df.collect())


def test_row_range_partitions_match_unpartitioned_read(spark, tmp_path):
    path = str(tmp_path / "cmp.xlsx")
    with zlsx.write(path) as w:
        s = w.add_sheet("S")
        s.write_row(["a", "b"])
        for i in range(50):
            s.write_row([i, i * 1.5])

    whole = spark.read.format("zlsx").load(path).collect()
    split = (spark.read.format("zlsx")
             .option("rowsPerPartition", "7")
             .load(path).collect())
    assert sorted((r.a, r.b) for r in whole) == sorted((r.a, r.b) for r in split)


def test_user_schema_failfast_vs_permissive(spark, workbook):
    ddl = "region string, product string, units bigint, revenue bigint"
    permissive = (spark.read.format("zlsx").schema(ddl)
                  .option("sheet", "Sales").load(workbook))
    by_key = {(r.region, r.product): r.revenue for r in permissive.collect()}
    assert by_key[("AMER", "flange")] is None  # 200.5 doesn't fit bigint

    failfast = (spark.read.format("zlsx").schema(ddl)
                .option("sheet", "Sales").option("mode", "failfast")
                .load(workbook))
    with pytest.raises(Exception, match="does not fit"):
        failfast.collect()


def test_write_single_file_roundtrip(spark, tmp_path):
    out = str(tmp_path / "report.xlsx")
    df = spark.createDataFrame(
        [("a", 1, 1.5, dt.datetime(2024, 1, 1, 12, 0)),
         ("b", 2, 2.5, dt.datetime(2025, 6, 30, 8, 30))],
        "name string, n bigint, x double, ts timestamp",
    )
    df.coalesce(1).write.format("zlsx").mode("overwrite").save(out)

    with zlsx.open(out) as book:
        header, rows = book.sheet(0).read_all(header=True)
    assert header == ["name", "n", "x", "ts"]
    assert len(rows) == 2
    by_name = {r[0]: r for r in rows}
    assert by_name["a"][1] == 1 and by_name["a"][2] == 1.5
    from zlsx._tabular import serial_to_datetime
    assert serial_to_datetime(by_name["a"][3]) == dt.datetime(2024, 1, 1, 12, 0)


def test_write_directory_part_files(spark, tmp_path):
    out = str(tmp_path / "parts")
    df = spark.createDataFrame(
        [(i, f"row{i}") for i in range(10)], "id bigint, label string"
    ).repartition(3)
    df.write.format("zlsx").mode("overwrite").save(out)

    parts = sorted(f for f in os.listdir(out) if f.endswith(".xlsx"))
    assert parts and all(f.startswith("part-") for f in parts)
    total = 0
    for f in parts:
        with zlsx.open(os.path.join(out, f)) as book:
            _, rows = book.sheet(0).read_all(header=True)
            total += len(rows)
    assert total == 10


def test_write_single_file_needs_one_partition(spark, tmp_path):
    out = str(tmp_path / "single.xlsx")
    df = spark.createDataFrame(
        [(i,) for i in range(10)], "id bigint"
    ).repartition(2)
    with pytest.raises(Exception, match="coalesce"):
        df.write.format("zlsx").mode("overwrite").save(out)


def test_write_existing_single_file_requires_overwrite(spark, tmp_path):
    out = str(tmp_path / "existing.xlsx")
    df = spark.createDataFrame([(1,)], "id bigint")
    df.coalesce(1).write.format("zlsx").mode("overwrite").save(out)
    with pytest.raises(Exception, match="overwrite"):
        df.coalesce(1).write.format("zlsx").mode("append").save(out)


# ─── streaming ────────────────────────────────────────────────────────


def _land_wb(path, names):
    """Write a one-sheet workbook with a `name` column, one row per name."""
    with zlsx.write(str(path)) as w:
        sh = w.add_sheet("Sheet1")
        sh.write_row(["name", "n"])
        for i, name in enumerate(names):
            sh.write_row([name, i])


def _drain(spark, zone, ckpt, sink):
    """Run the stream until everything currently in the zone is
    processed, then stop. processAllAvailable works with any source;
    availableNow support for Python stream sources is not assumed."""
    q = (
        spark.readStream.format("zlsx")
        .schema("name string, n bigint")
        .load(str(zone))
        .writeStream.format("parquet")
        .option("path", str(sink))
        .option("checkpointLocation", str(ckpt))
        .trigger(processingTime="1 second")
        .start()
    )
    try:
        q.processAllAvailable()
    finally:
        q.stop()


def test_stream_ingests_existing_then_only_new(spark, tmp_path):
    zone = tmp_path / "zone"
    zone.mkdir()
    ckpt = tmp_path / "ckpt"
    sink = tmp_path / "sink"
    _land_wb(zone / "a.xlsx", ["a1", "a2"])
    _land_wb(zone / "b.xlsx", ["b1"])

    _drain(spark, zone, ckpt, sink)
    got = {r.name for r in spark.read.parquet(str(sink)).collect()}
    assert got == {"a1", "a2", "b1"}

    # New file lands; a restarted stream on the SAME checkpoint must
    # ingest it exactly once and re-ingest nothing.
    _land_wb(zone / "c.xlsx", ["c1", "c2"])
    _drain(spark, zone, ckpt, sink)
    rows = spark.read.parquet(str(sink)).collect()
    assert len(rows) == 5  # 3 old + 2 new, no duplicates
    assert {r.name for r in rows} == {"a1", "a2", "b1", "c1", "c2"}


def test_stream_available_now_trigger(spark, tmp_path):
    """Serverless jobs compute REJECTS infinite triggers, so
    availableNow is the trigger the platform forces. Python stream
    sources don't implement it natively; Spark falls back to
    single-batch execution — verify that path drains and terminates."""
    zone = tmp_path / "zone"
    zone.mkdir()
    _land_wb(zone / "a.xlsx", ["x1", "x2"])

    q = (
        spark.readStream.format("zlsx")
        .schema("name string, n bigint")
        .load(str(zone))
        .writeStream.format("parquet")
        .option("path", str(tmp_path / "sink"))
        .option("checkpointLocation", str(tmp_path / "ckpt"))
        .trigger(availableNow=True)
        .start()
    )
    q.awaitTermination()
    got = {r.name for r in spark.read.parquet(str(tmp_path / "sink")).collect()}
    assert got == {"x1", "x2"}


def test_stream_starting_position_latest_skips_existing(spark, tmp_path):
    from pyspark.sql.types import StructType

    zone = tmp_path / "zone"
    zone.mkdir()
    _land_wb(zone / "pre.xlsx", ["old"])

    from zlsx.spark import ZlsxStreamReader

    schema = StructType.fromDDL("name string, n bigint")
    latest = ZlsxStreamReader({"path": str(zone), "startingPosition": "latest"}, schema)
    assert latest.initialOffset() == latest.latestOffset()
    assert latest.partitions(latest.initialOffset(), latest.latestOffset()) == []

    earliest = ZlsxStreamReader({"path": str(zone)}, schema)
    assert earliest.initialOffset() == {"files": {}}
    parts = earliest.partitions(earliest.initialOffset(), earliest.latestOffset())
    assert len(parts) == 1 and parts[0].path.endswith("pre.xlsx")


# ─── cross-file inference + atomic part landing ───────────────────────


def test_infer_widens_across_files(spark, tmp_path):
    """A landing zone where the type evidence is split across files: the
    first workbook alone says bigint, the pair says double. Before
    multi-file inference this read nulled every fractional value."""
    zone = tmp_path / "zone"
    zone.mkdir()
    for name, value in (("a.xlsx", 10), ("b.xlsx", 20.5)):
        with zlsx.write(str(zone / name)) as w:
            sh = w.add_sheet("S")
            sh.write_row(["region", "revenue"])
            sh.write_row(["AMER", value])

    df = spark.read.format("zlsx").load(str(zone / "*.xlsx"))
    assert dict(df.dtypes)["revenue"] == "double"
    assert sorted(r.revenue for r in df.collect()) == [10.0, 20.5]


def test_infer_files_bound_is_honoured(spark, tmp_path):
    zone = tmp_path / "zone"
    zone.mkdir()
    for name, value in (("a.xlsx", 10), ("b.xlsx", 20.5)):
        with zlsx.write(str(zone / name)) as w:
            sh = w.add_sheet("S")
            sh.write_row(["revenue"])
            sh.write_row([value])

    df = (
        spark.read.format("zlsx")
        .option("inferFiles", "1")
        .load(str(zone / "*.xlsx"))
    )
    # Only a.xlsx sampled, so the schema stays narrow and b.xlsx's
    # fractional value cannot fit — permissive mode nulls it.
    assert dict(df.dtypes)["revenue"] == "bigint"
    got = [r.revenue for r in df.collect()]
    assert sorted(got, key=lambda v: (v is None, v)) == [10, None]


def test_write_leaves_no_temp_files_behind(spark, tmp_path):
    out = tmp_path / "outdir"
    df = spark.createDataFrame([("a", 1), ("b", 2)], "name string, n bigint")
    df.write.format("zlsx").mode("overwrite").save(str(out))

    names = sorted(p.name for p in out.iterdir())
    assert all(n.startswith("part-") and n.endswith(".xlsx") for n in names), names
    assert not any("zlsx-tmp" in n for n in names)

    got = set()
    for p in out.iterdir():
        with zlsx.open(str(p)) as book:
            _, rows = book.sheet(0).read_all(header=True)
            got.update(tuple(r) for r in rows)
    assert got == {("a", 1), ("b", 2)}


# ─── batch recalc (§12.4) ─────────────────────────────────────────────


def _stale_calc_book(path, a2=1, cached=999.0):
    """Header + one data row: A2=`a2`, B2='A2+2' with a STALE cached
    value — the engine computes a2+2, the file claims `cached`."""
    with zlsx.write(str(path)) as w:
        s = w.add_sheet("Calc")
        s.write_row(["a", "b"])
        s.write_row_with_formulas([a2, cached], [None, "A2+2"])


def test_recalc_batch_read_returns_engine_values(spark, tmp_path):
    """The activation contract: without zlsx.recalc the read trusts the
    workbook's cached values; with it, every value is one the engine
    computed — and the source file's bytes are untouched either way."""
    path = str(tmp_path / "calc.xlsx")
    _stale_calc_book(path)
    before = open(path, "rb").read()

    stale = spark.read.format("zlsx").load(path)
    assert [tuple(r) for r in stale.collect()] == [(1, 999.0)]

    live = (spark.read.format("zlsx")
            .option("zlsx.recalc", "true")
            .load(path))
    assert [tuple(r) for r in live.collect()] == [(1, 3.0)]
    assert open(path, "rb").read() == before  # read-only guarantee


def test_recalc_schema_inference_sees_engine_types(spark, tmp_path):
    """Inference runs on the recalced snapshot: a formula cell whose
    stale cache is text but whose computed value is numeric must infer
    numeric under recalc (integral results read back as ints → bigint),
    string without it."""
    path = str(tmp_path / "typed.xlsx")
    _stale_calc_book(path, cached="TBD")

    stale = spark.read.format("zlsx").load(path)
    assert dict(stale.dtypes)["b"] == "string"

    live = (spark.read.format("zlsx")
            .option("zlsx.recalc", "true")
            .load(path))
    assert dict(live.dtypes)["b"] == "bigint"
    assert [tuple(r) for r in live.collect()] == [(1, 3)]


def _recalc_reader(path):
    from pyspark.sql.types import StructType

    from zlsx.spark import ZlsxBatchReader

    schema = StructType.fromDDL("a bigint, b double")
    return ZlsxBatchReader({"path": path, "zlsx.recalc": "true"}, schema)


def test_recalc_partition_carries_identity_and_retry_rederives(tmp_path):
    """Partition inputs carry (digest, resolved context, fingerprint);
    reading the same partition twice — a task retry — yields identical
    rows because everything is re-derived from those inputs."""
    path = str(tmp_path / "calc.xlsx")
    _stale_calc_book(path)
    reader = _recalc_reader(path)
    parts = reader.partitions()
    assert len(parts) == 1
    tag = parts[0].recalc
    assert set(tag) == {"digest", "ctx", "fingerprint"}
    assert tag["fingerprint"] == zlsx.engine_fingerprint()
    assert tag["ctx"]["utc_offset_min"] == 0  # UTC default, like every layer

    rows1 = list(reader.read(parts[0]))
    rows2 = list(reader.read(parts[0]))
    assert rows1 == rows2 == [(1, 3.0)]


def test_recalc_digest_drift_refuses_identically_on_retry(tmp_path):
    """A workbook rewritten between planning and execution refuses by
    digest — and a retry refuses identically instead of silently
    reading the mutant."""
    from zlsx import _tabular

    path = str(tmp_path / "calc.xlsx")
    _stale_calc_book(path)
    reader = _recalc_reader(path)
    parts = reader.partitions()

    _stale_calc_book(path, a2=50)  # drifts after planning

    with pytest.raises(_tabular.SnapshotDriftError, match="immutable") as first:
        list(reader.read(parts[0]))
    with pytest.raises(_tabular.SnapshotDriftError) as retry:
        list(reader.read(parts[0]))
    assert str(retry.value) == str(first.value)


def test_recalc_engine_fingerprint_mismatch_refuses(tmp_path):
    """A partition planned by a different engine build refuses on the
    executor — mixed-version fleets never mix recalc results."""
    from zlsx import _tabular

    path = str(tmp_path / "calc.xlsx")
    _stale_calc_book(path)
    reader = _recalc_reader(path)
    parts = reader.partitions()
    parts[0].recalc = dict(parts[0].recalc, fingerprint="zlsx 0.0.0; bogus")

    with pytest.raises(_tabular.EngineFingerprintMismatch):
        list(reader.read(parts[0]))


def test_recalc_streaming_refused_at_option_validation(spark, tmp_path):
    from pyspark.sql.types import StructType

    from zlsx.spark import ZlsxStreamReader

    zone = tmp_path / "zone"
    zone.mkdir()
    _land_wb(zone / "a.xlsx", ["x1"])
    schema = StructType.fromDDL("name string, n bigint")

    with pytest.raises(ValueError, match="batch-only"):
        ZlsxStreamReader({"path": str(zone), "zlsx.recalc": "true"}, schema)

    # And through the real streaming API: the query must fail with the
    # same refusal, not silently stream stale caches.
    with pytest.raises(Exception, match="batch-only"):
        q = (
            spark.readStream.format("zlsx")
            .schema("name string, n bigint")
            .option("zlsx.recalc", "true")
            .load(str(zone))
            .writeStream.format("parquet")
            .option("path", str(tmp_path / "sink"))
            .option("checkpointLocation", str(tmp_path / "ckpt"))
            .trigger(availableNow=True)
            .start()
        )
        q.awaitTermination()


def test_overwrite_replaces_single_file_completely(spark, tmp_path):
    """Rewriting the same single-file target must leave a valid workbook
    holding only the new rows — not a mix of old and new bytes."""
    out = tmp_path / "report.xlsx"
    wide = spark.createDataFrame(
        [(f"row{i}", i) for i in range(500)], "name string, n bigint"
    )
    wide.coalesce(1).write.format("zlsx").mode("overwrite").save(str(out))

    small = spark.createDataFrame([("only", 1)], "name string, n bigint")
    small.coalesce(1).write.format("zlsx").mode("overwrite").save(str(out))

    with zlsx.open(str(out)) as book:
        _, rows = book.sheet(0).read_all(header=True)
    assert rows == [["only", 1]]
