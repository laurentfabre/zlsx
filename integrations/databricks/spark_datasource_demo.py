"""Query an xlsx workbook in a UC Volume directly with Spark SQL — no Delta copy.

Two paths, tried in order:

  1. PySpark Data Source (Spark 4 / DBR 15.4+ / serverless):
         spark.read.format("zlsx").option("sheet", "Sales").load(path)
  2. Fallback: py-zlsx -> createDataFrame -> temp view

Either way the workbook itself is the queryable object.

Verified on Databricks serverless 2026-07-30 (SUCCESS via path 1). Deploy:

  - Upload the py-zlsx wheel to a UC Volume. Serverless compute is
    Graviton: use the manylinux_2_28_aarch64 wheel from the release assets.
  - Put THIS FILE under /Workspace/... — a spark_python_task cannot read
    its python_file from a Volume path.
  - Submit with an environment spec:
        {"client": "3", "dependencies": ["/Volumes/<cat>/<schema>/<vol>/py_zlsx-<ver>-py3-none-manylinux_2_28_aarch64.whl"]}

Override the workbook path with --path (defaults to the demo fixture
produced by make_fixture.py).
"""
import sys

from pyspark.sql import SparkSession

WB = "/Volumes/workspace/default/zlsx_smoke/fy2026_sales.xlsx"
if "--path" in sys.argv:
    WB = sys.argv[sys.argv.index("--path") + 1]

spark = SparkSession.builder.getOrCreate()

print("=== zlsx direct-query demo ===")
print("workbook:", WB)

import zlsx  # noqa: E402  (from the wheel dependency)

print("py-zlsx import OK")


def sheet_rows(path, sheet):
    with zlsx.open(path) as book:
        return list(book.sheet(sheet).rows())


def ddl_schema(header, first):
    def typ(v):
        # bool before int: bool is an int subclass in Python.
        if isinstance(v, bool):
            return "boolean"
        if isinstance(v, int):
            return "bigint"
        if isinstance(v, float):
            return "double"
        return "string"

    return ", ".join(f"`{h}` {typ(v)}" for h, v in zip(header, first))


used = None
try:
    from pyspark.sql.datasource import DataSource, DataSourceReader

    class ZlsxReader(DataSourceReader):
        def __init__(self, options):
            self.path = options["path"]
            self.sheet_name = options.get("sheet", "0")

        def read(self, partition):
            sheet = int(self.sheet_name) if self.sheet_name.isdigit() else self.sheet_name
            rows = sheet_rows(self.path, sheet)
            for r in rows[1:]:  # skip header row
                yield tuple(r)

    class ZlsxDataSource(DataSource):
        @classmethod
        def name(cls):
            return "zlsx"

        def schema(self):
            # v0 inference: types from the FIRST data row only. Good enough
            # for the demo; the productized source must widen across rows.
            sheet = self.options.get("sheet", "0")
            sheet = int(sheet) if sheet.isdigit() else sheet
            rows = sheet_rows(self.options["path"], sheet)
            return ddl_schema(rows[0], rows[1])

        def reader(self, schema):
            return ZlsxReader(self.options)

    spark.dataSource.register(ZlsxDataSource)
    df = (spark.read.format("zlsx")
          .option("sheet", "Sales")
          .load(WB))
    used = 'PYSPARK DATA SOURCE  spark.read.format("zlsx")'
except Exception as e:  # noqa: BLE001 — fall back, report why
    print(f"data-source path unavailable ({type(e).__name__}: {e}); using temp-view fallback")
    rows = sheet_rows(WB, "Sales")
    df = spark.createDataFrame(rows[1:], schema=ddl_schema(rows[0], rows[1]))
    used = "TEMP VIEW over py-zlsx rows"

df.createOrReplaceTempView("workbook")
print("path used:", used)
print("row count:", df.count())

out = spark.sql("""
    SELECT region,
           ROUND(SUM(revenue), 2)      AS revenue,
           SUM(units)                  AS units,
           COUNT(DISTINCT product)     AS products
    FROM workbook
    GROUP BY region ORDER BY region
""").collect()
print("SQL over the workbook (no Delta):")
for r in out:
    print("  ", r.region, r.revenue, r.units, r.products)
print("=== done ===")
