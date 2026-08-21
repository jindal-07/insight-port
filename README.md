# Insightport 📊

**Facebook page insights exporter** — select your pages, pick a date range,
and export video/reel insights as a CSV plus a refreshable Excel PivotTable,
**each file downloading straight to your browser the moment it's ready**.

Two output files per run: `social_feed_insights_<range>.csv` and
`social_feed_pivot.xlsx`. Results live in a per-run temp directory on the
server and are replaced when the next run starts.

## Features

- **Auto-discovers pages** — lists every page your token can access (`/me/accounts`, paginated)
- **Custom date range** — since/until pickers with presets (7/14/30 days, month-to-date)
- **Filtered export** — video/reel posts longer than 120s, with insights
  (views, engagement, watch time, earnings) and hashtags
- **Pivot workbook** — a real, refreshable Excel PivotTable
  (`page_name → hashtag` with count / earnings / avg views)
- **Live progress log** — streamed from the exporter while it runs, with a Stop button
- **Per-file auto-download** — the CSV downloads as soon as it's final (while
  the pivot is still building), the pivot follows the moment it's written;
  a manual "Download .zip" button bundles both on demand
- **Password protection** — optional HTTP basic auth via `APP_PASSWORD`

## Project layout

| File | Purpose |
|---|---|
| `server.py` | Flask app: page listing, export orchestration, in-memory ZIP downloads |
| `facebook_export.py` | The exporter — fetches feed + insights from the Graph API, writes CSVs |
| `build_pivot_sheet.py` | Builds the refreshable PivotTable workbook from the feed CSV |
| `frontend/index.html` | Single-file web UI (Material 3, no build step) |
| `Dockerfile`, `render.yaml` | Container + Render blueprint for deployment |

## Configuration

Set via environment variables, or a `.env` file next to `server.py` (real env
vars take precedence):

| Variable | Required | Description |
|---|---|---|
| `FB_USER_TOKEN` | ✅ | Facebook user access token with `pages_show_list`, `read_insights`, `pages_read_engagement` |
| `APP_PASSWORD` | recommended in production | Enables basic auth on every route (any username + this password). Unset = open access |
| `HOST` / `PORT` | no | Dev-server bind address (defaults `127.0.0.1:5000`; the Docker image uses gunicorn on `$PORT`, falling back to 7860) |

> ⚠️ Never commit `.env` — it is gitignored. On hosting platforms, set these
> as dashboard secrets.

## Run locally

```bash
pip install -r requirements.txt
echo "FB_USER_TOKEN=your-token" > .env
python server.py        # → http://127.0.0.1:5000
```

## Deploy on Render (free tier)

1. Push this repo to GitHub (private recommended).
2. [dashboard.render.com](https://dashboard.render.com) → **New → Blueprint** →
   select the repo. `render.yaml` configures the service (Docker, free plan,
   health check at `/healthz`).
3. Enter the prompted secrets: `FB_USER_TOKEN` and `APP_PASSWORD`.
4. **Apply** — first build takes a few minutes, then the app is live.

Free-tier notes: the service sleeps after ~15 min idle (first request takes
~30–50 s to wake); the UI's status polling keeps it awake during an export.

## API

| Endpoint | Description |
|---|---|
| `GET /api/pages` | Pages accessible via the token |
| `POST /api/export` | Start an export — `{pages: [...], since, until}` (until is exclusive) |
| `GET /api/export/status?offset=N` | Running state + log lines from `N` |
| `POST /api/export/stop` | Terminate a running export |
| `GET /api/files` | Live manifest of the run's files with per-file `ready` flags |
| `GET /api/download/<name>` | Download one result file (409 while still being written) |
| `GET /api/download-all` | Optional: zip the ready files on the fly |
| `GET /healthz` | Unauthenticated health probe |

## Notes

- The Graph API treats **`until` as exclusive** — the last exported day is the day before it.
- Results live in a per-run temp dir until the next run replaces them or the server restarts.
- Output of `accessible_pages.csv`, `social_pages.csv`, `social_page_insights.csv`,
  and `social_feed.csv` is currently **paused** (see `CSVWriter._filenames`);
  the page-insights API calls are skipped accordingly.
- Single-process by design: the Docker image runs `gunicorn -w 1 --threads 8`.
  Do **not** raise the worker count — export state is per-process.
