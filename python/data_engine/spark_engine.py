"""
SparkEngine — wraps PySpark for large-file loading and aggregation.

The Qt frontend communicates with this module via the IPCServer (TCP socket).
Data is transferred as JSON for small chunks, or Apache Arrow IPC for bulk
transfers (when pyarrow is available).
"""
from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)

# ── Optional heavy imports ─────────────────────────────────────────────────────
try:
    from pyspark.sql import SparkSession, DataFrame
    from pyspark.sql import functions as F
    PYSPARK_AVAILABLE = True
except ImportError:
    PYSPARK_AVAILABLE = False
    logger.warning("PySpark not available — falling back to pandas/csv")

try:
    import pyarrow as pa
    import pyarrow.ipc as pa_ipc
    ARROW_AVAILABLE = True
except ImportError:
    ARROW_AVAILABLE = False
    logger.warning("PyArrow not available — using JSON transfer")

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False


class SparkEngine:
    """
    Manages PySpark sessions and cached DataFrames.

    Each loaded file is assigned a numeric df_id so the Qt side can reference
    it in subsequent aggregate / filter calls without re-loading the data.
    """

    def __init__(self) -> None:
        self._spark: SparkSession | None = None
        self._dataframes: dict[int, Any] = {}   # df_id → DataFrame or pd.DataFrame
        self._next_id = 1

    # ── Session management ─────────────────────────────────────────────────────

    def _get_spark(self) -> SparkSession:
        if self._spark is None:
            if not PYSPARK_AVAILABLE:
                raise RuntimeError("PySpark is not installed")
            self._spark = (
                SparkSession.builder
                .appName("OpenSheet")
                .config("spark.driver.memory", "2g")
                .config("spark.sql.adaptive.enabled", "true")
                .getOrCreate()
            )
            self._spark.sparkContext.setLogLevel("WARN")
        return self._spark

    # ── File loading ───────────────────────────────────────────────────────────

    def load_csv(self, path: str, chunk: int = 0, chunk_size: int = 1000) -> dict:
        """
        Load a CSV file using PySpark (or pandas as fallback).

        Returns a dict with:
          df_id      — handle for subsequent operations
          columns    — list of column names
          row_count  — total rows in file
          rows       — list of rows for the requested chunk  (list of lists)
        """
        p = Path(path)
        if not p.exists():
            return {"error": f"File not found: {path}"}

        df_id = self._next_id
        self._next_id += 1

        try:
            if PYSPARK_AVAILABLE:
                spark = self._get_spark()
                df = spark.read.option("header", "true").option("inferSchema", "true").csv(str(p))
                self._dataframes[df_id] = df
                total = df.count()
                columns = df.columns
                chunk_df = df.limit(chunk * chunk_size + chunk_size).collect()
                start = chunk * chunk_size
                rows = [[str(v) for v in row] for row in chunk_df[start:start + chunk_size]]
            elif PANDAS_AVAILABLE:
                df = pd.read_csv(str(p))
                self._dataframes[df_id] = df
                total = len(df)
                columns = list(df.columns)
                slice_ = df.iloc[chunk * chunk_size: (chunk + 1) * chunk_size]
                rows = slice_.values.tolist()
                rows = [[str(v) for v in row] for row in rows]
            else:
                return {"error": "Neither PySpark nor pandas is installed"}

            return {
                "df_id": df_id,
                "columns": columns,
                "row_count": total,
                "rows": rows,
            }
        except Exception as exc:
            logger.exception("load_csv failed")
            return {"error": str(exc)}

    def load_parquet(self, path: str, chunk: int = 0, chunk_size: int = 1000) -> dict:
        """Load a Parquet file using PySpark."""
        p = Path(path)
        if not p.exists():
            return {"error": f"File not found: {path}"}
        if not PYSPARK_AVAILABLE:
            return {"error": "PySpark required for Parquet support"}

        df_id = self._next_id
        self._next_id += 1
        try:
            spark = self._get_spark()
            df = spark.read.parquet(str(p))
            self._dataframes[df_id] = df
            total = df.count()
            columns = df.columns
            chunk_df = df.limit(chunk * chunk_size + chunk_size).collect()
            start = chunk * chunk_size
            rows = [[str(v) for v in row] for row in chunk_df[start:start + chunk_size]]
            return {
                "df_id": df_id,
                "columns": columns,
                "row_count": total,
                "rows": rows,
            }
        except Exception as exc:
            logger.exception("load_parquet failed")
            return {"error": str(exc)}

    # ── Data operations ────────────────────────────────────────────────────────

    def get_chunk(self, df_id: int, chunk: int, chunk_size: int = 1000) -> dict:
        """Fetch a subsequent chunk from an already-loaded DataFrame."""
        df = self._dataframes.get(df_id)
        if df is None:
            return {"error": f"Unknown df_id: {df_id}"}
        try:
            if PYSPARK_AVAILABLE and hasattr(df, "collect"):
                collected = df.limit((chunk + 1) * chunk_size).collect()
                start = chunk * chunk_size
                rows = [[str(v) for v in row] for row in collected[start:]]
            else:
                slice_ = df.iloc[chunk * chunk_size: (chunk + 1) * chunk_size]
                rows = [[str(v) for v in row] for row in slice_.values.tolist()]
            return {"df_id": df_id, "rows": rows}
        except Exception as exc:
            return {"error": str(exc)}

    def aggregate(self, df_id: int, operations: list[dict]) -> dict:
        """
        Run aggregation operations on a DataFrame.

        operations is a list of dicts, each with:
          { "type": "sum"|"count"|"avg"|"min"|"max", "column": "col_name" }

        Returns { "results": [ { "type": ..., "column": ..., "value": ... } ] }
        """
        df = self._dataframes.get(df_id)
        if df is None:
            return {"error": f"Unknown df_id: {df_id}"}
        results = []
        try:
            for op in operations:
                op_type = op.get("type", "").lower()
                col = op.get("column", "")
                if not col:
                    continue
                if PYSPARK_AVAILABLE and hasattr(df, "agg"):
                    fn_map = {
                        "sum": F.sum, "count": F.count,
                        "avg": F.avg, "min": F.min, "max": F.max,
                    }
                    fn = fn_map.get(op_type)
                    if fn:
                        val = df.agg(fn(col)).collect()[0][0]
                    else:
                        continue
                elif PANDAS_AVAILABLE:
                    agg_map = {
                        "sum": "sum", "count": "count",
                        "avg": "mean", "min": "min", "max": "max",
                    }
                    method = agg_map.get(op_type)
                    if method:
                        val = getattr(df[col], method)()
                    else:
                        continue
                else:
                    continue
                results.append({"type": op_type, "column": col, "value": val})
        except Exception as exc:
            return {"error": str(exc)}
        return {"results": results}

    def filter_rows(self, df_id: int, column: str, operator: str, value: str,
                    chunk: int = 0, chunk_size: int = 1000) -> dict:
        """Filter a DataFrame and return a new df_id for the result."""
        df = self._dataframes.get(df_id)
        if df is None:
            return {"error": f"Unknown df_id: {df_id}"}
        try:
            if PYSPARK_AVAILABLE and hasattr(df, "filter"):
                expr = f"`{column}` {operator} {value}"
                filtered = df.filter(expr)
            elif PANDAS_AVAILABLE:
                expr = f"`{column}` {operator} {value}"
                filtered = df.query(expr)
            else:
                return {"error": "No data engine available"}

            new_id = self._next_id
            self._next_id += 1
            self._dataframes[new_id] = filtered
            total = filtered.count() if hasattr(filtered, "count") else len(filtered)
            return {"df_id": new_id, "row_count": total}
        except Exception as exc:
            return {"error": str(exc)}

    def release(self, df_id: int) -> dict:
        """Release a cached DataFrame to free memory."""
        if df_id in self._dataframes:
            del self._dataframes[df_id]
            return {"ok": True}
        return {"error": f"Unknown df_id: {df_id}"}

    def shutdown(self) -> None:
        """Stop the Spark session."""
        self._dataframes.clear()
        if self._spark:
            self._spark.stop()
            self._spark = None
