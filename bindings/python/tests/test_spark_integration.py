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
