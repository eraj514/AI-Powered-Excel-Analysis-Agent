import os
import io
import json
import math
import logging
from typing import Any

import pandas as pd
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from groq import Groq

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(title="Excel Analyzer API", version="1.0.0", docs_url="/api/docs", redoc_url="/api/redoc", openapi_url="/api/openapi.json")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

groq_client = Groq(api_key=os.environ["GROQ_API_KEY"])


# ── helpers ──────────────────────────────────────────────────────────────────

def _sanitize(obj: Any) -> Any:
    """Recursively make a value JSON-safe (no NaN / Inf, no numpy scalars)."""
    # numpy integer types → Python int
    if type(obj).__module__ == "numpy" and hasattr(obj, "item"):
        obj = obj.item()
    if isinstance(obj, float):
        if math.isnan(obj) or math.isinf(obj):
            return None
        return obj
    if isinstance(obj, dict):
        return {k: _sanitize(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_sanitize(v) for v in obj]
    return obj


def _analyze_dataframe(df: pd.DataFrame) -> dict:
    """Run pandas analysis and return a structured summary dict."""
    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    categorical_cols = df.select_dtypes(include=["object", "category"]).columns.tolist()
    datetime_cols = df.select_dtypes(include=["datetime", "datetimetz"]).columns.tolist()

    # --- basic stats ---
    stats: dict[str, Any] = {}
    if numeric_cols:
        raw = df[numeric_cols].describe().to_dict()
        stats = _sanitize(raw)

    # --- missing values ---
    missing_counts = df.isnull().sum()
    missing = {
        col: {"count": int(missing_counts[col]),
              "pct": round(float(missing_counts[col]) / len(df) * 100, 2)}
        for col in df.columns
        if missing_counts[col] > 0
    }

    # --- correlations (numeric only, skip if fewer than 2 columns) ---
    correlations: dict = {}
    if len(numeric_cols) >= 2:
        corr_raw = df[numeric_cols].corr().round(4).to_dict()
        correlations = _sanitize(corr_raw)

    # --- categorical value counts (top 10) ---
    value_counts: dict = {}
    for col in categorical_cols[:10]:          # cap at first 10 categorical cols
        vc = df[col].value_counts().head(10)
        value_counts[col] = vc.to_dict()

    # --- numeric distributions (quartiles + skewness) ---
    distributions: dict = {}
    for col in numeric_cols[:20]:              # cap at 20 numeric cols
        series = df[col].dropna()
        if len(series) == 0:
            continue
        distributions[col] = _sanitize({
            "min": series.min(),
            "q1": series.quantile(0.25),
            "median": series.median(),
            "q3": series.quantile(0.75),
            "max": series.max(),
            "mean": series.mean(),
            "std": series.std(),
            "skewness": series.skew(),
        })

    # --- datetime ranges ---
    datetime_ranges: dict = {}
    for col in datetime_cols:
        datetime_ranges[col] = {
            "min": str(df[col].min()),
            "max": str(df[col].max()),
        }

    # --- duplicates ---
    duplicate_count = int(df.duplicated().sum())

    return {
        "shape": {"rows": int(df.shape[0]), "columns": int(df.shape[1])},
        "column_names": df.columns.tolist(),
        "dtypes": {col: str(dtype) for col, dtype in df.dtypes.items()},
        "numeric_columns": numeric_cols,
        "categorical_columns": categorical_cols,
        "datetime_columns": datetime_cols,
        "missing_values": missing,
        "duplicate_rows": duplicate_count,
        "descriptive_stats": stats,
        "correlations": correlations,
        "value_counts": value_counts,
        "distributions": distributions,
        "datetime_ranges": datetime_ranges,
    }


def _get_ai_insights(analysis: dict, sheet_name: str) -> dict:
    """Call OpenAI with the analysis summary and return structured insights."""
    prompt = f"""You are a senior data analyst. You have been given a statistical summary of an Excel sheet named "{sheet_name}".

Produce a concise, actionable analysis in the following JSON structure (return ONLY valid JSON, no markdown):

{{
  "summary": "<2-3 sentence overview of what this dataset is about and its quality>",
  "key_findings": ["<finding 1>", "<finding 2>", "<finding 3>", ...],
  "data_quality_issues": ["<issue 1>", ...],
  "recommendations": ["<recommendation 1>", ...],
  "interesting_patterns": ["<pattern 1>", ...]
}}

Statistical summary:
{json.dumps(analysis, indent=2)}
"""
    response = groq_client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role": "user", "content": prompt}],
        max_tokens=1500,
        temperature=0.3,
    )

    raw_text = response.choices[0].message.content or ""
    # Strip possible markdown fences
    raw_text = raw_text.strip()
    if raw_text.startswith("```"):
        raw_text = raw_text.split("```")[1]
        if raw_text.startswith("json"):
            raw_text = raw_text[4:]
    raw_text = raw_text.strip().rstrip("`").strip()

    try:
        return json.loads(raw_text)
    except json.JSONDecodeError:
        logger.warning("Could not parse AI response as JSON; returning raw text.")
        return {"raw": raw_text}


# ── routes ────────────────────────────────────────────────────────────────────

@app.get("/api/healthz")
async def healthz():
    return {"status": "ok"}


@app.post("/api/analyze")
async def analyze_excel(
    file: UploadFile = File(..., description="Excel file (.xlsx or .xls)"),
    sheet_name: str | None = None,
    ai_insights: bool = True,
):
    """
    Upload an Excel file and receive:
    - pandas statistical analysis for each sheet (or a specific one)
    - optional AI-generated narrative insights via OpenAI
    """
    # --- validate file type ---
    filename = file.filename or ""
    if not filename.lower().endswith((".xlsx", ".xls", ".xlsm", ".xlsb")):
        raise HTTPException(
            status_code=400,
            detail="Only Excel files (.xlsx, .xls, .xlsm, .xlsb) are supported.",
        )

    contents = await file.read()
    if len(contents) == 0:
        raise HTTPException(status_code=400, detail="Uploaded file is empty.")

    # --- read workbook ---
    try:
        excel_file = pd.ExcelFile(io.BytesIO(contents))
    except Exception as exc:
        logger.exception("Failed to parse Excel file")
        raise HTTPException(status_code=422, detail=f"Could not parse Excel file: {exc}")

    available_sheets = excel_file.sheet_names
    sheets_to_analyze = (
        [sheet_name] if sheet_name else available_sheets
    )

    # Validate requested sheet exists
    for sn in sheets_to_analyze:
        if sn not in available_sheets:
            raise HTTPException(
                status_code=400,
                detail=f"Sheet '{sn}' not found. Available: {available_sheets}",
            )

    results: list[dict] = []

    for sn in sheets_to_analyze:
        try:
            df = pd.read_excel(io.BytesIO(contents), sheet_name=sn)
        except Exception as exc:
            logger.exception("Failed to read sheet %s", sn)
            results.append({"sheet": sn, "error": str(exc)})
            continue

        # Try to parse object columns as datetimes
        for col in df.select_dtypes(include="object").columns:
            try:
                df[col] = pd.to_datetime(df[col], infer_datetime_format=True)
            except Exception:
                pass

        analysis = _analyze_dataframe(df)

        sheet_result: dict[str, Any] = {
            "sheet": sn,
            "analysis": analysis,
        }

        if ai_insights:
            try:
                sheet_result["ai_insights"] = _get_ai_insights(analysis, sn)
            except Exception as exc:
                logger.exception("AI insights failed for sheet %s", sn)
                sheet_result["ai_insights"] = {"error": str(exc)}

        results.append(sheet_result)

    return _sanitize({
        "filename": filename,
        "available_sheets": available_sheets,
        "sheets_analyzed": sheets_to_analyze,
        "results": results,
    })


@app.get("/api/sheets")
async def list_sheets(file: UploadFile = File(...)):
    """Return the list of sheet names in an Excel file without analyzing it."""
    contents = await file.read()
    try:
        excel_file = pd.ExcelFile(io.BytesIO(contents))
        return {"sheets": excel_file.sheet_names}
    except Exception as exc:
        raise HTTPException(status_code=422, detail=str(exc))
