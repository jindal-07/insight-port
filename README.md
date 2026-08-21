---
title: Insightport
emoji: 📊
colorFrom: red
colorTo: yellow
sdk: docker
app_port: 7860
pinned: false
---

# Insightport

Facebook page insights exporter — select pages, pick a date range, and export
video/reel insights as CSVs plus a refreshable Excel PivotTable, zipped
straight to your browser. Nothing is saved server-side: exports run in a temp
dir that is zipped into memory and deleted.

## Setup (required)

This Space needs one secret:

| Secret | Value |
|---|---|
| `FB_USER_TOKEN` | A Facebook user access token with `pages_show_list`, `read_insights`, and `pages_read_engagement` permissions |

Add it under **Settings → Variables and secrets → New secret**.

⚠️ Keep this Space **Private** — the app has no login of its own, and the token
grants access to your pages' insights and monetization data.

## How it works

- `GET /api/pages` — lists every page the token can access
- `POST /api/export` — runs `facebook_export.py` as a subprocess with the
  selected pages + date range (env-driven)
- The browser polls `/api/export/status` for the live log, then auto-downloads
  the in-memory ZIP from `/api/download-all`

## Local development

```bash
pip install -r requirements.txt
echo "FB_USER_TOKEN=your-token" > .env
python server.py   # http://127.0.0.1:5000
```
