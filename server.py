#!/usr/bin/env python3
"""
server.py — web frontend for facebook_export.py

Endpoints
---------
GET  /                     the frontend (frontend/index.html)
GET  /api/pages            all pages accessible via the user token (me/accounts)
POST /api/export           start an export {pages: [...], since, until, all_pages}
GET  /api/export/status    running flag + log lines (poll while running)
POST /api/export/stop      terminate a running export
GET  /api/files            output files in facebook_export_csvs/
GET  /api/download/<name>  download one output file

The export itself runs facebook_export.py as a subprocess; page names and the
date range are passed via the FB_EXPORT_* environment variables the script
understands. Token comes from FB_USER_TOKEN (falls back to the value inside
facebook_export.py).

Run:  python server.py   →  http://127.0.0.1:5000
"""

import io
import mimetypes
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import zipfile
from datetime import datetime

import requests
from dotenv import load_dotenv
from flask import Flask, Response, jsonify, request, send_file, send_from_directory

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXPORT_SCRIPT = os.path.join(BASE_DIR, 'facebook_export.py')

# Secrets live in .env (FB_USER_TOKEN). Real env vars take precedence.
load_dotenv(os.path.join(BASE_DIR, '.env'))

API_VERSION = 'v24.0'
GRAPH = f'https://graph.facebook.com/{API_VERSION}'

app = Flask(__name__, static_folder='frontend', static_url_path='')

# ------------------------------------------------------------- basic auth
# Set APP_PASSWORD (env / .env / host secret) to require a password on every
# route. Leave it unset for open access (e.g. local dev or a private Space).
APP_PASSWORD = os.environ.get('APP_PASSWORD')


@app.before_request
def _require_auth():
    if request.path == '/healthz':          # platform health checks skip auth
        return None
    if not APP_PASSWORD:
        return None
    auth = request.authorization
    if auth and auth.password == APP_PASSWORD:
        return None
    return Response('Authentication required.', 401,
                    {'WWW-Authenticate': 'Basic realm="Insightport"'})


@app.get('/healthz')
def healthz():
    return jsonify({'ok': True})


# ---------------------------------------------------------------- token
def get_user_token() -> str | None:
    """FB_USER_TOKEN from the environment (populated from .env at startup)."""
    return os.environ.get('FB_USER_TOKEN') or None


# ---------------------------------------------------------------- export state
class ExportJob:
    """Holds the one running/last-run export subprocess and its log."""

    def __init__(self):
        self.lock = threading.Lock()
        self.proc: subprocess.Popen | None = None
        self.log: list[str] = []
        self.running = False
        self.returncode: int | None = None
        self.params: dict | None = None
        self.started_at: str | None = None
        self.finished_at: str | None = None
        # Results live in a per-run temp dir. Files are served straight from
        # it as soon as they are complete (no ZIP buffering); the dir is
        # replaced when the next run starts.
        self.out_dir: str | None = None
        # Set when the exporter logs that the feed CSV is final (it starts
        # building the pivot) — lets the CSV download before the run ends.
        self.feed_csv_ready = False

    def start(self, params: dict) -> None:
        env = os.environ.copy()
        token = get_user_token()
        if token:
            env['FB_USER_TOKEN'] = token
        if self.out_dir:                       # previous run's results
            shutil.rmtree(self.out_dir, ignore_errors=True)
        self.out_dir = tempfile.mkdtemp(prefix='fb_export_')
        env['FB_EXPORT_OUTPUT_DIR'] = self.out_dir
        env['FB_EXPORT_SINCE'] = params['since']
        env['FB_EXPORT_UNTIL'] = params['until']
        if params.get('all_pages'):
            env['FB_EXPORT_ALL_PAGES'] = '1'
            env.pop('FB_EXPORT_PAGES', None)
        else:
            env.pop('FB_EXPORT_ALL_PAGES', None)
            env['FB_EXPORT_PAGES'] = '|'.join(params['pages'])
        # The exporter prints emoji; force UTF-8 so it doesn't crash on cp1252.
        env['PYTHONIOENCODING'] = 'utf-8'
        env['PYTHONUNBUFFERED'] = '1'

        self.log = []
        self.returncode = None
        self.params = params
        self.started_at = datetime.now().isoformat(timespec='seconds')
        self.finished_at = None
        self.feed_csv_ready = False
        self.running = True

        self.proc = subprocess.Popen(
            [sys.executable, EXPORT_SCRIPT],
            cwd=BASE_DIR,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding='utf-8',
            errors='replace',
            env=env,
        )
        threading.Thread(target=self._pump, daemon=True).start()

    def _pump(self) -> None:
        assert self.proc and self.proc.stdout
        for line in self.proc.stdout:
            line = line.rstrip('\r\n')
            if line:
                with self.lock:
                    self.log.append(line)
                    # The exporter only starts the pivot build after the feed
                    # CSV is fully written, so this line marks it complete.
                    if 'Building pivot workbook' in line:
                        self.feed_csv_ready = True
        self.proc.wait()
        with self.lock:
            self.returncode = self.proc.returncode
            self.running = False
            self.finished_at = datetime.now().isoformat(timespec='seconds')

    def _scan_files(self) -> list[dict]:
        """Result files with per-file readiness. Callers must hold the lock."""
        files = []
        if self.out_dir and os.path.isdir(self.out_dir):
            done = self.returncode is not None
            for name in sorted(os.listdir(self.out_dir)):
                path = os.path.join(self.out_dir, name)
                if not os.path.isfile(path):
                    continue
                ready = done or (self.feed_csv_ready
                                 and name.startswith('social_feed_insights'))
                files.append({'name': name, 'size': os.path.getsize(path),
                              'ready': ready})
        return files

    def stop(self) -> bool:
        if self.proc and self.running:
            self.proc.terminate()
            return True
        return False

    def status(self, since_line: int = 0) -> dict:
        with self.lock:
            return {
                'running': self.running,
                'returncode': self.returncode,
                'params': self.params,
                'started_at': self.started_at,
                'finished_at': self.finished_at,
                'log_offset': len(self.log),
                'log': self.log[since_line:],
                'files': self._scan_files(),
            }


job = ExportJob()


# ---------------------------------------------------------------- routes
@app.get('/')
def index():
    return send_from_directory(app.static_folder, 'index.html')


@app.get('/api/pages')
def api_pages():
    """All pages accessible via the user token, following pagination."""
    token = get_user_token()
    if not token:
        return jsonify({'error': 'No token. Set the FB_USER_TOKEN environment variable.'}), 400

    pages, url = [], f'{GRAPH}/me/accounts'
    params = {'fields': 'id,name,category,picture{url}', 'limit': 100,
              'access_token': token}
    try:
        while url:
            resp = requests.get(url, params=params, timeout=30)
            data = resp.json()
            if 'error' in data:
                msg = data['error'].get('message', 'Unknown Graph API error')
                return jsonify({'error': msg}), 502
            for p in data.get('data', []):
                pages.append({
                    'id': p.get('id', ''),
                    'name': p.get('name', ''),
                    'category': p.get('category', ''),
                    'picture': (p.get('picture') or {}).get('data', {}).get('url', ''),
                })
            url = data.get('paging', {}).get('next')
            params = None  # pagination URL already carries everything
    except requests.RequestException as e:
        return jsonify({'error': f'Request failed: {e}'}), 502

    pages.sort(key=lambda p: p['name'].lower())
    return jsonify({'pages': pages, 'count': len(pages)})


@app.post('/api/export')
def api_export():
    if job.running:
        return jsonify({'error': 'An export is already running.'}), 409

    body = request.get_json(silent=True) or {}
    since = (body.get('since') or '').strip()
    until = (body.get('until') or '').strip()
    pages = body.get('pages') or []
    all_pages = bool(body.get('all_pages'))

    date_re = re.compile(r'^\d{4}-\d{2}-\d{2}$')
    if not date_re.match(since) or not date_re.match(until):
        return jsonify({'error': 'since/until must be YYYY-MM-DD.'}), 400
    try:
        d_since = datetime.strptime(since, '%Y-%m-%d')
        d_until = datetime.strptime(until, '%Y-%m-%d')
    except ValueError:
        return jsonify({'error': 'Invalid date.'}), 400
    if d_since >= d_until:
        return jsonify({'error': 'since must be before until (until is exclusive).'}), 400
    if not all_pages and not pages:
        return jsonify({'error': 'Select at least one page (or choose all pages).'}), 400

    job.start({'since': since, 'until': until,
               'pages': pages, 'all_pages': all_pages})
    return jsonify({'ok': True})


@app.get('/api/export/status')
def api_status():
    offset = request.args.get('offset', default=0, type=int)
    return jsonify(job.status(offset))


@app.post('/api/export/stop')
def api_stop():
    return jsonify({'stopped': job.stop()})


@app.get('/api/files')
def api_files():
    """Live manifest of the current/last run's results with readiness flags."""
    with job.lock:
        return jsonify({'files': job._scan_files(),
                        'finished_at': job.finished_at})


@app.get('/api/download/<path:name>')
def api_download(name):
    """Stream one result file straight from the run's temp dir."""
    name = os.path.basename(name)          # no path traversal
    with job.lock:
        out_dir = job.out_dir
        ready = {f['name'] for f in job._scan_files() if f['ready']}
    if not out_dir:
        return jsonify({'error': 'No results yet. Run an export first.'}), 404
    path = os.path.join(out_dir, name)
    if not os.path.isfile(path):
        return jsonify({'error': f'{name} not found.'}), 404
    if name not in ready:
        return jsonify({'error': f'{name} is still being written.'}), 409
    mime = mimetypes.guess_type(name)[0] or 'application/octet-stream'
    return send_file(path, mimetype=mime, as_attachment=True,
                     download_name=name)


@app.get('/api/download-all')
def api_download_all():
    """Optional: zip the ready files on the fly (manual button)."""
    with job.lock:
        out_dir = job.out_dir
        files = [f for f in job._scan_files() if f['ready']]
    if not out_dir or not files:
        return jsonify({'error': 'No results yet. Run an export first.'}), 404
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            zf.write(os.path.join(out_dir, f['name']), arcname=f['name'])
    buf.seek(0)
    return send_file(buf, mimetype='application/zip', as_attachment=True,
                     download_name=f'insightport_export_{datetime.now():%Y-%m-%d_%H%M%S}.zip')


if __name__ == '__main__':
    host = os.environ.get('HOST', '127.0.0.1')
    port = int(os.environ.get('PORT', 5000))
    print('📦 Results    : streamed per-file to the browser as each is ready')
    print(f'🔑 Token      : {"FB_USER_TOKEN loaded" if get_user_token() else "⚠️ MISSING — add FB_USER_TOKEN to .env"}')
    print(f'🌐 Frontend   : http://{host}:{port}')
    app.run(host=host, port=port, debug=False)
