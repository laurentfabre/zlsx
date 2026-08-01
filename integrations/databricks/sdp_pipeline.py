"""Lakeflow SDP pipeline: a streaming table fed by the zlsx source.

The declarative half of "Auto Loader for Excel" — workbooks landing in
a Volume zone flow into a governed streaming table, no landing script,
no Delta copy step. Verified 2026-08-01 on a serverless pipeline
(pipeline `zlsx-excel-landing`, workspace catalog): two updates over a
zone that grew from two workbooks to three ingested every row exactly
once.

Deploy notes, all verified:

  - Pipeline spec: {"serverless": true, "catalog": ..., "target": ...,
    "libraries": [{"file": {"path": "/Workspace/.../sdp_pipeline.py"}}],
    "environment": {"dependencies": ["/Volumes/.../py_zlsx-*.whl"]}}
    — the wheel installs from the Volume into the pipeline environment.
  - Register the data source INSIDE this file; the pipeline runs it on
    its own driver.
  - Pipelines compute mounts /Volumes via FUSE, so the source's
    driver-side zone listing works as it does on jobs compute.
  - Diagnostic trap: querying the streaming table right after an update
    on an idle warehouse returns state PENDING until spin-up finishes —
    poll the statement to completion before concluding "0 rows".
"""
import dlt
from pyspark.sql import SparkSession

spark = SparkSession.getActiveSession()

from zlsx.spark import ZlsxDataSource  # noqa: E402

spark.dataSource.register(ZlsxDataSource)

ZONE = "/Volumes/workspace/default/zlsx_smoke/stream_zone/"


@dlt.table(
    name="excel_landing",
    comment="Streaming table over Excel workbooks landing in the Volume "
            "zone — ingested exactly-once by the zlsx streaming source.",
)
def excel_landing():
    return (
        spark.readStream.format("zlsx")
        .schema("name string, n bigint")
        .load(ZONE)
    )
