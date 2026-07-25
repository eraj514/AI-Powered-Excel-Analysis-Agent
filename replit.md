# Excel Analyzer API

A Python FastAPI backend that accepts Excel file uploads, analyzes them with pandas, and returns statistical insights enriched with AI-generated narrative via OpenAI.

## Run & Operate

- `artifacts/api-server: API Server` workflow — starts the FastAPI server (port 8080, proxied at `/api`)
- Restart the workflow after any changes to `artifacts/api-server/main.py`

## Stack

- Python 3.13
- FastAPI + Uvicorn
- pandas (statistical analysis)
- openpyxl (Excel parsing)
- OpenAI Python SDK (AI insights)

## Where things live

- `artifacts/api-server/main.py` — all route handlers and analysis logic
- `artifacts/api-server/requirements.txt` — Python dependencies
- `artifacts/api-server/.replit-artifact/artifact.toml` — service config (port, paths, run command)

## API Endpoints

| Method | Path | Description |
|--------|------|-------------|
| GET | `/api/healthz` | Health check |
| POST | `/api/analyze` | Upload Excel → get pandas stats + AI insights JSON |
| GET | `/api/sheets` | List sheet names in an uploaded Excel file |
| GET | `/api/docs` | Auto-generated Swagger UI |
| GET | `/api/redoc` | ReDoc API documentation |

### POST `/api/analyze` parameters

- `file` (form, required) — `.xlsx`, `.xls`, `.xlsm`, or `.xlsb`
- `sheet_name` (query, optional) — analyze a specific sheet; defaults to all sheets
- `ai_insights` (query, optional, default `true`) — include OpenAI narrative insights

### Response shape

```json
{
  "filename": "data.xlsx",
  "available_sheets": ["Sheet1"],
  "sheets_analyzed": ["Sheet1"],
  "results": [{
    "sheet": "Sheet1",
    "analysis": {
      "shape": { "rows": 100, "columns": 5 },
      "column_names": [...],
      "dtypes": {...},
      "missing_values": {...},
      "duplicate_rows": 0,
      "descriptive_stats": {...},
      "correlations": {...},
      "value_counts": {...},
      "distributions": {...}
    },
    "ai_insights": {
      "summary": "...",
      "key_findings": [...],
      "data_quality_issues": [...],
      "recommendations": [...],
      "interesting_patterns": [...]
    }
  }]
}
```

## Architecture decisions

- All analysis and AI calls live in a single `main.py` for simplicity — no database, no auth.
- `_sanitize()` recursively replaces `NaN`/`Inf` floats before JSON serialisation (pandas produces these for empty numeric columns).
- AI prompt requests strict JSON output; a fallback strips markdown fences if the model wraps the response.
- Categorical and numeric columns are capped (10 / 20 cols respectively) to keep the AI prompt under token limits for wide spreadsheets.

## Product

Users POST an Excel file to `/api/analyze` and receive a rich JSON response covering shape, data types, missing values, duplicate rows, descriptive statistics, correlations, per-column distributions, and an AI-written summary with key findings, data quality issues, recommendations, and interesting patterns.

## User preferences

_Populate as you build._

## Gotchas

- OpenAI model is `gpt-4o-mini`; change in `_get_ai_insights()` if a more capable model is needed.
- Pass `ai_insights=false` as a query param to skip the OpenAI call (useful for testing without incurring API cost).
- The virtual environment lives at `.pythonlibs/`; packages are managed via `uv` — add new packages with `uv add <pkg>` from the workspace root.
