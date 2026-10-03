# Deploying to Render (free instance)

Use dummy students while testing. Do not put real children's names or marks on a hosted service until the
education department has agreed to it. The server stores nothing (no disk, no database, request contents are
never logged), but the data still passes through Render's machines.

## 1. Put the code on GitHub
The repo root must be this folder, with `app.py`, `config.py`, `render.yaml` and `requirements.txt` at the top.
Before the first push run `git status` and check no roster, workbook, photo of real students or data file is staged.
The `.gitignore` excludes `*.csv`, `*.xlsx`, `data_*.json` and `out/`. The photos in `tests/fixtures/` are dummy data.

## 2. Create the service
Render dashboard > New > Blueprint > choose the repo > Apply. It reads `render.yaml`.
If Render asks for a card at this step, use New > Web Service instead and enter the same values by hand:
build `pip install -r requirements.txt`, start
`gunicorn app:app --workers 1 --threads 2 --timeout 120 --bind 0.0.0.0:$PORT`,
health check path `/health`, and the environment variables below.

| Variable | Value |
|---|---|
| `PYTHON_VERSION` | `3.12.3` (a full patch version; if the build rejects it, choose another 3.12.x) |
| `API_KEYS` | one or more secret keys, comma-separated. `render.yaml` asks Render to generate the first one |
| `ALLOWED_ORIGINS` | `chrome-extension://<your extension id>` (only needed if a web page, not the extension's background script, calls the server) |

Make an extra key for another teacher with `python -c "import secrets; print(secrets.token_urlsafe(32))"` and add it
to `API_KEYS` after a comma. To revoke someone, remove their key and let the service restart.

## 3. Check it
```
curl https://YOUR-SERVICE.onrender.com/health                       # {"ok":true}
curl -X POST https://YOUR-SERVICE.onrender.com/sheets               # 401: no key
curl -X POST https://YOUR-SERVICE.onrender.com/sheets -H "X-API-Key: YOUR_KEY" -H "Content-Type: application/json" \
     -d '{"class":3,"section":"A","school":"TEST SCHOOL","academic_year":"2025-26","exam_type":"Term 1","working_days":37,"roster":[{"id":"1001","name":"TEST ONE"}]}' -o sheets.zip
```

The `/sheets` request may include `"output":"excel"` for a standalone `.xlsx` response or
`"output":"pdf"` for a ZIP containing only PDFs. Omit `output` (or use `"all"`) for both.

## 4. Calling it from the extension
- Call the server from the extension's **background script (service worker)** and list the service URL under
  `host_permissions`. Requests from there are not blocked by the cross-origin rules that apply to a page.
- Keep the API key in `chrome.storage.local`, entered once by the teacher. Never put it in the extension's code.
- **Cold start:** a free instance sleeps after about 15 minutes without traffic, and the next request can take
  up to a minute. Send a `GET /health` as soon as the teacher opens the portal page, show "waking the server up",
  and use a timeout of at least 90 seconds with one retry.
- Endpoints and formats are described at the top of `app.py`.

## 5. Things to know about the free instance
- 512 MB RAM and a small share of one CPU. Measured on a clean install here: about 200 MB peak memory and 0.2 s of CPU
  time to read a 12-megapixel photo, so a photo should take a few seconds on the throttled instance (not measured on Render).
- Free-tier limits change. Check Render's pricing page, and note that the cheapest paid instance removes the sleeping.
- Uploads are limited to 20 MB per request and 30 requests a minute per key (`MAX_UPLOAD_MB`, `RATE_PER_MIN`).
