# Deploy On Render

This project can be deployed as a single Render web service because the FastAPI backend already serves the static frontend from `frontend-static`.

## What gets deployed

- Backend: `backend/main.py`
- Frontend: `frontend-static/index.html`, `frontend-static/styles.css`, `frontend-static/app.js`

## Before you start

1. Push `pdf-extraction-app` to GitHub.
2. Keep the folder structure exactly like this:
   - `backend`
   - `frontend-static`
   - `render.yaml`

## Deploy steps

1. Log in to Render.
2. Click `New` -> `Blueprint`.
3. Connect your GitHub repo.
4. Select the repo that contains `pdf-extraction-app`.
5. Render will detect `render.yaml`.
6. Set the required environment variable:
   - `GEMINI_API_KEY`
7. Optional:
   - `PREVIEW_API_TOKEN`
8. Deploy.

## Start command used

```bash
uvicorn main:app --host 0.0.0.0 --port $PORT
```

## Health check

Render checks:

```text
/api/health
```

## Important note

Uploaded files and generated outputs are stored in temporary workspace storage. On free hosting, that storage is ephemeral, so generated assets should be downloaded by the user and not treated as permanent server storage.
