# fastapi-pptx-service

FastAPI service for PPTX template analysis (and later: deck generation).

## Run (Render)
- Build: `pip install -r requirements.txt`
- Start: `uvicorn main:app --host 0.0.0.0 --port $PORT`

Endpoints:
- `GET /health` → `{"status":"ok"}`
- `POST /template/analyze` (multipart .pptx) → `{ styleProfile: {...} }`
