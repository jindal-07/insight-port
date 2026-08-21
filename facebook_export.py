# Facebook Graph API Data Exporter - CSV Output Version
#
# CHANGES vs the previous version
# -------------------------------
# 1. Only video/reel posts longer than MIN_DURATION_SECONDS (120s) are exported.
#    The duration gate runs BEFORE the insights calls, so short posts cost two
#    cheap lookups instead of six requests.
# 2. New `hashtag` column, derived from `message` in Python (replaces the
#    TEXTBEFORE/TEXTAFTER formula that used to live in column T).
# 3. Token is read from the FB_USER_TOKEN environment variable. Never commit it.
# 4. Output CSVs are truncated at start-up (RESET_OUTPUT) -- flush() appends, so
#    re-running without this silently doubled every row and the pivot with it.
# 5. When the export finishes it builds social_feed_pivot.xlsx: a real,
#    refreshable PivotTable of page_name -> hashtag.

import requests
import pandas as pd
import json
import re
from datetime import datetime, timedelta
import time
import logging
import warnings
import os
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed

warnings.filterwarnings('ignore')
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Load secrets from the .env file next to this script (FB_USER_TOKEN lives
# there). Existing environment variables take precedence over the file.
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env'))

# Output directory for CSV files. The web frontend overrides this with a temp
# dir that is zipped straight to the browser and deleted — nothing persists.
CSV_OUTPUT_DIR = os.environ.get('FB_EXPORT_OUTPUT_DIR', 'facebook_export_csvs')
os.makedirs(CSV_OUTPUT_DIR, exist_ok=True)

print("✅ All packages imported successfully!")

# ---
## Configuration Settings

CONFIG = {
    'API_VERSION': 'v24.0',
    'BASE_URL': 'https://graph.facebook.com',
    # Read from the environment / .env file — never hardcode tokens here.
    'USER_TOKEN': os.environ.get('FB_USER_TOKEN', ''),
    'REQUEST_DELAY': 0.5,
    'MAX_RETRIES': 3,
    'ERROR_REQUEST_DELAY': 2,

    # Threading Configuration
    'MAX_WORKERS': 20,
    'BATCH_SIZE': 100,

    # ---- Post filtering -------------------------------------------------
    'VIDEO_ONLY': True,            # drop photo/text/link posts entirely
    'MIN_DURATION_SECONDS': 120,   # keep only videos STRICTLY longer than this

    # Date Configuration
    'FEED_DAYS_BACK': 30,
    'INSIGHTS_DAYS_BACK': 1,
    'USE_CUSTOM_DATES': True,
    'CUSTOM_SINCE_DATE': '2026-08-01',
    'CUSTOM_UNTIL_DATE': '2026-08-19',
    'FETCH_ALL_POSTS': False,

    # Page Selection Configuration
    'FETCH_ALL_PAGES': False,
    'SPECIFIC_PAGE_IDS': [],
    'SPECIFIC_PAGE_NAMES': ['Alright'],
    # 	Extra page names, kept here for easy re-enabling:
    #   'Alright Bhakti', 'CID Dastak', 'Alright Crime Stories', 'Binge', 'Alright Ehsaas', 'Social Tales', 'Alright Parivar',
    #   'PopCon', 'Love unity vines', 'Living In Trend', 'Alright Tamasha'

    # ---- Output ---------------------------------------------------------
    'RESET_OUTPUT': True,          # wipe previous CSVs so re-runs don't append
    'BUILD_PIVOT': True,           # build the pivot workbook when export ends
    'PIVOT_FILENAME': 'social_feed_pivot.xlsx',

    # Metrics Configuration
    'PAGE_INSIGHTS_METRICS': [
        'page_follows',
        'page_post_engagements',
        'monetization_approximate_earnings'
    ],
    'PAGE_INSIGHTS_PERIOD': 'day',
    'POST_INSIGHTS_METRICS': [
        'post_media_view',
        'post_video_views',
        'post_reactions_by_type_total',
        'post_video_avg_time_watched',
        'monetization_approximate_earnings'
    ],
    'POST_INSIGHTS_PERIOD': 'lifetime',
}

# ---- Environment overrides (set by server.py / the web frontend) ----------
# FB_EXPORT_SINCE / FB_EXPORT_UNTIL : YYYY-MM-DD custom date range
# FB_EXPORT_PAGES                   : pipe-separated exact page names
# FB_EXPORT_ALL_PAGES=1             : export every accessible page
if os.environ.get('FB_EXPORT_SINCE') and os.environ.get('FB_EXPORT_UNTIL'):
    CONFIG['USE_CUSTOM_DATES'] = True
    CONFIG['FETCH_ALL_POSTS'] = False
    CONFIG['CUSTOM_SINCE_DATE'] = os.environ['FB_EXPORT_SINCE']
    CONFIG['CUSTOM_UNTIL_DATE'] = os.environ['FB_EXPORT_UNTIL']

if os.environ.get('FB_EXPORT_ALL_PAGES') == '1':
    CONFIG['FETCH_ALL_PAGES'] = True
elif os.environ.get('FB_EXPORT_PAGES'):
    CONFIG['FETCH_ALL_PAGES'] = False
    CONFIG['SPECIFIC_PAGE_IDS'] = []
    CONFIG['SPECIFIC_PAGE_NAMES'] = [
        n.strip() for n in os.environ['FB_EXPORT_PAGES'].split('|') if n.strip()
    ]

if not CONFIG['USER_TOKEN']:
    raise SystemExit(
        "❌ FB_USER_TOKEN is not set.\n"
        "   Add it to the .env file next to this script:\n"
        "       FB_USER_TOKEN=your-token-here\n"
        "   or set it as an environment variable, then re-run."
    )


def calculate_dates(config):
    """Calculate since and until dates based on configuration.

    Always returns a dict with the same four keys so callers never have to
    branch on the return shape.
    """
    if config['USE_CUSTOM_DATES']:
        return {
            'feed_since':     config['CUSTOM_SINCE_DATE'],
            'feed_until':     config['CUSTOM_UNTIL_DATE'],
            'insights_since': config['CUSTOM_SINCE_DATE'],
            'insights_until': config['CUSTOM_UNTIL_DATE'],
        }

    today = datetime.now()
    return {
        'feed_since':     (today - timedelta(days=config['FEED_DAYS_BACK'])).strftime('%Y-%m-%d'),
        'feed_until':     (today - timedelta(days=1)).strftime('%Y-%m-%d'),
        'insights_since': (today - timedelta(days=config['INSIGHTS_DAYS_BACK'])).strftime('%Y-%m-%d'),
        'insights_until': (today - timedelta(days=1)).strftime('%Y-%m-%d'),
    }


date_config = calculate_dates(CONFIG)


def feed_range_tag():
    """Date-range tag used in the feed insights filename.

    Graph API treats `until` as exclusive for the feed window, so the last day
    actually present in the export is until - 1 day. The filename reflects the
    data the file contains, not the raw request parameter.
    """
    if CONFIG['FETCH_ALL_POSTS']:
        return 'all_posts'
    since = date_config['feed_since']
    last_day = datetime.strptime(date_config['feed_until'], '%Y-%m-%d') - timedelta(days=1)
    return f"{since}_to_{last_day:%Y-%m-%d}"


FEED_INSIGHTS_FILENAME = f"social_feed_insights_{feed_range_tag()}.csv"

print("📋 Configuration loaded")
print(f"🧵 Max Workers: {CONFIG['MAX_WORKERS']}")
print(f"📦 Batch Size: {CONFIG['BATCH_SIZE']}")
print(f"\n🎬 Post filter: video/reel only, duration > {CONFIG['MIN_DURATION_SECONDS']}s")
print(f"\n📂 CSV output directory: {os.path.abspath(CSV_OUTPUT_DIR)}")
print(f"📄 Feed insights file:   {FEED_INSIGHTS_FILENAME}")

print("\n📅 Date Configuration:")
if CONFIG['FETCH_ALL_POSTS']:
    print("   Mode: FETCH ALL POSTS (no date limits)")
elif CONFIG['USE_CUSTOM_DATES']:
    print(f"   Mode: Custom Dates")
    print(f"   Since: {CONFIG['CUSTOM_SINCE_DATE']}, Until: {CONFIG['CUSTOM_UNTIL_DATE']}")
else:
    print(f"   Mode: Days Back")
    print(f"   Feed: {date_config['feed_since']} to {date_config['feed_until']}")
    print(f"   Insights: {date_config['insights_since']} to {date_config['insights_until']}")


# ---
## Hashtag extraction

# Mirrors the old column-T formula:
#   IF(ISNUMBER(SEARCH("#",msg)), "#" & TEXTBEFORE(TEXTAFTER(<msg, newlines->spaces>, "#"), " ", 1, 0, 1), "")
# i.e. the first "#" and everything up to the next whitespace, or "" if no "#".
_HASHTAG_RE = re.compile(r'#\S*')


def extract_hashtag(message):
    """Return the first hashtag in `message`, or '' when there isn't one."""
    if not message:
        return ''
    match = _HASHTAG_RE.search(str(message))
    return match.group(0) if match else ''


# ---
## CSV Writer — replaces MySQLManager

# Explicit column order for the feed file. This is the layout the pivot builder
# and the existing workbook expect: A=date ... T=hashtag.
FEED_COLUMNS = [
    'date', 'page_id', 'page_name', 'post_id', 'published_date', 'post_type',
    'message', 'permalink_url', 'impressions', 'video_views', 'engagements',
    'avg_time_watched', 'approximate_earnings', 'shares', 'views', 'status',
    'duration_seconds', 'video_views_60s', 'custom_label', 'hashtag',
]
# `full_picture` is deliberately omitted -- it isn't in the workbook and the
# URLs bloat the CSV. Add it back here and in add_feed_insights if you need it.

PAGE_COLUMNS = ['page_id', 'name', 'username', 'about', 'profile_pic',
                'category', 'followers_count', 'source', 'status']
PAGE_INSIGHT_COLUMNS = ['page_id', 'date', 'follows', 'impressions',
                        'engagement', 'approximate_earnings', 'status']
LABEL_COLUMNS = ['feed_id', 'attachment_id', 'custom_label', 'post_type', 'status']


class CSVWriter:
    """
    Collects data in memory and writes to CSV files.

    Mirrors the exact column structure that was being inserted into the database:
      - social_pages.csv
      - social_page_insights.csv
      - social_feed_insights_<since>_to_<until-1>.csv
      - social_feed.csv  (custom labels)
    """

    # In-memory stores (list of dicts)
    _pages = []
    _page_insights = []
    _feed_insights = []
    _custom_labels = []

    # ------------------------------------------------------------------ pages
    @classmethod
    def add_page(cls, page_data: dict):
        cls._pages.append({
            'page_id':         page_data.get('page_id'),
            'name':            page_data.get('name', ''),
            'username':        page_data.get('username', ''),
            'about':           page_data.get('about', ''),
            'profile_pic':     page_data.get('profile_pic', ''),
            'category':        page_data.get('category', ''),
            'followers_count': page_data.get('followers_count', 0),
            'source':          'facebook',
            'status':          1,
        })

    # -------------------------------------------------------- page insights
    @classmethod
    def add_page_insights(cls, insights_list: list):
        for insight in insights_list:
            cls._page_insights.append({
                'page_id':              insight.get('page_id'),
                'date':                 insight.get('date'),
                'follows':              insight.get('follows', 0),
                'impressions':          insight.get('impressions', 0),
                'engagement':           insight.get('post_engagements', 0),
                'approximate_earnings': insight.get('monetization_approximate_earnings', 0.0),
                'status':               1,
            })

    # -------------------------------------------------------- feed insights
    @classmethod
    def add_feed_insights(cls, feed_list: list):
        yesterday = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
        for feed in feed_list:
            message = (feed.get('message') or '')[:1000]
            cls._feed_insights.append({
                'date':                 yesterday,
                'page_id':              feed.get('page_id'),
                'page_name':            feed.get('page_name', ''),
                'post_id':              feed.get('post_id'),
                'published_date':       feed.get('created_time'),
                'post_type':            feed.get('post_type', 'text'),
                'message':              message,
                'permalink_url':        feed.get('permalink_url', ''),
                'impressions':          feed.get('impressions', 0),
                'video_views':          feed.get('video_views', 0),
                'engagements':          feed.get('reactions_total', 0),
                'avg_time_watched':     feed.get('video_avg_time_watched', 0.0),
                'approximate_earnings': feed.get('monetization_approximate_earnings', 0.0),
                'shares':               feed.get('shares', 0),
                'views':                feed.get('views', 0),
                'status':               1,
                'duration_seconds':     feed.get('duration_seconds', 0),
                'video_views_60s':      feed.get('video_views_60s', 0),
                'custom_label':         feed.get('custom_label', 'none'),
                # Derived here so the CSV is self-contained -- no Excel formula.
                'hashtag':              extract_hashtag(message),
            })

    # ------------------------------------------------------ custom labels
    @classmethod
    def add_custom_labels(cls, labels_list: list):
        for label in labels_list:
            cls._custom_labels.append({
                'feed_id':       label.get('feed_id'),
                'attachment_id': label.get('attachment_id', 'N/A'),
                'custom_label':  label.get('custom_label', ''),
                'post_type':     label.get('post_type', 'text'),
                'status':        1,
            })

    # --------------------------------------------------------------- reset
    @classmethod
    def reset_files(cls):
        """Delete previous output so a re-run replaces rather than appends.

        flush() opens in append mode; without this a second run would double
        every row and the pivot totals with them.
        """
        removed = []
        for filename in cls._filenames():
            filepath = os.path.join(CSV_OUTPUT_DIR, filename)
            if os.path.exists(filepath):
                os.remove(filepath)
                removed.append(filename)
        if removed:
            logger.info(f"🧹 Cleared previous output: {', '.join(removed)}")

    @staticmethod
    def _filenames():
        return ('social_pages.csv', 'social_page_insights.csv',
                FEED_INSIGHTS_FILENAME, 'social_feed.csv')

    # ------------------------------------------------------------ save all
    @classmethod
    def flush(cls, label: str = ""):
        """
        Append whatever is currently in memory to the CSV files, then clear
        the buffers.  Safe to call repeatedly — uses append mode so rows from
        previous pages are never overwritten.
        """
        datasets = {
            'social_pages.csv':         (cls._pages, PAGE_COLUMNS),
            'social_page_insights.csv': (cls._page_insights, PAGE_INSIGHT_COLUMNS),
            FEED_INSIGHTS_FILENAME:     (cls._feed_insights, FEED_COLUMNS),
            'social_feed.csv':          (cls._custom_labels, LABEL_COLUMNS),
        }

        tag = f" [{label}]" if label else ""
        for filename, (rows, columns) in datasets.items():
            if not rows:
                continue
            filepath = os.path.join(CSV_OUTPUT_DIR, filename)
            df = pd.DataFrame(rows, columns=columns)
            write_header = not os.path.exists(filepath)          # header only on first write
            df.to_csv(filepath, mode='a', index=False, header=write_header)
            logger.info(f"💾 Flushed {len(rows):,} rows → {filename}{tag}")

        # Clear buffers so the next page starts fresh
        cls._pages.clear()
        cls._page_insights.clear()
        cls._feed_insights.clear()
        cls._custom_labels.clear()

    # ------------------------------------------------------------ save_all
    @classmethod
    def save_all(cls):
        """
        Flush any remaining in-memory rows (pages that weren't flushed yet)
        and print a summary of every CSV file that exists on disk.
        """
        cls.flush(label="final")          # flush leftovers

        files_written = []
        for filename in cls._filenames():
            filepath = os.path.join(CSV_OUTPUT_DIR, filename)
            if os.path.exists(filepath):
                # Count real records, not raw lines -- messages contain newlines.
                row_count = len(pd.read_csv(filepath))
                files_written.append((filename, row_count))
                logger.info(f"✅ {filename}  →  {row_count:,} total rows on disk")
        return files_written


print("🔧 CSVWriter class defined successfully!")


# ---
## Facebook API Client Class with Threading Support

class FacebookAPIClient:
    def __init__(self, config):
        self.config = config
        self.base_url = f"{config['BASE_URL']}/{config['API_VERSION']}"
        self.user_token = config['USER_TOKEN']

    def make_request(self, url, params=None, page_token=None):
        """Make API request with retry logic"""
        if params is None:
            params = {}
        if page_token is None:
            params['access_token'] = self.user_token

        for attempt in range(self.config['MAX_RETRIES']):
            try:
                response = requests.get(url, params=params)
                response.raise_for_status()
                time.sleep(self.config['REQUEST_DELAY'])
                return response.json()
            except requests.exceptions.RequestException as e:
                logger.warning(f"Request failed (attempt {attempt + 1}): {e}")
                if attempt == self.config['MAX_RETRIES'] - 1:
                    return None
                time.sleep(2 ** attempt)

    def get_pages_and_tokens(self):
        """Get page info with pagination support"""
        logger.info("Fetching page information...")
        url = f"{self.base_url}/me/accounts"
        params = {
            'fields': 'access_token,id,name,category',
            'limit': 100  # Request more per page to reduce API calls
        }

        all_pages = []
        page_count = 0

        while url:
            # Only pass params on the first request; pagination URLs already contain them
            response = self.make_request(url, params if page_count == 0 else None, page_token=None)
            if not response:
                break

            pages = response.get('data', [])
            all_pages.extend(pages)
            page_count += 1
            logger.info(f"   Fetched {len(all_pages)} pages so far...")

            # Get next page URL from paging
            paging = response.get('paging', {})
            url = paging.get('next')
            params = None  # Pagination URL already has all params including access_token

        logger.info(f"✅ Total pages fetched: {len(all_pages)}")
        return all_pages

    def get_all_accessible_pages(self):
        """
        Fetch ALL pages accessible via the user token, following pagination.
        Returns a list of dicts with page_id and name.
        """
        url = f"{self.base_url}/me/accounts"
        params = {
            'fields': 'id,name,category,tasks',
            'limit': 100,
        }

        all_pages = []
        while url:
            response = self.make_request(
                url,
                params if url == f"{self.base_url}/me/accounts" else None
            )
            if not response:
                break
            all_pages.extend(response.get('data', []))
            url = response.get('paging', {}).get('next')
            params = None

        # NOTE: the previous version made a per-page detail request here and
        # then discarded the result. Removed -- it was one wasted API call per
        # page on every run.
        return [{'page_id': p.get('id', ''), 'name': p.get('name', '')}
                for p in all_pages]

    def list_accessible_pages(self):
        """Print a formatted table of all accessible pages and save to accessible_pages.csv."""
        print("\n" + "=" * 80)
        print("📋  ALL PAGES ACCESSIBLE VIA THIS USER TOKEN")
        print("=" * 80)

        pages = self.get_all_accessible_pages()

        if not pages:
            print("  ❌  No pages found — check that the token has 'pages_show_list' permission.")
            print("=" * 80)
            return []

        col_widths = {'#': 4, 'page_id': 18, 'name': 35}
        header = (
            f"{'#':<{col_widths['#']}}"
            f"{'Page ID':<{col_widths['page_id']}}"
            f"{'Name':<{col_widths['name']}}"
        )
        print(header)
        print("-" * len(header))

        for i, p in enumerate(pages, 1):
            row = (
                f"{i:<{col_widths['#']}}"
                f"{p['page_id']:<{col_widths['page_id']}}"
                f"{p['name'][:34]:<{col_widths['name']}}"
            )
            print(row)

        print("-" * len(header))
        print(f"\n  Total pages found: {len(pages)}")
        print("=" * 80)

        filepath = os.path.join(CSV_OUTPUT_DIR, 'accessible_pages.csv')
        pd.DataFrame(pages).to_csv(filepath, index=False)
        print(f"  💾 Saved to: {os.path.abspath(filepath)}\n")

        return pages

    def get_page_basic_info(self, page_id, page_token):
        url = f"{self.base_url}/{page_id}"
        params = {
            'fields': 'name,username,about,category,followers_count,picture',
            'access_token': page_token
        }
        return self.make_request(url, params, page_token)

    def get_page_insights(self, page_id, page_token):
        since_date = date_config['insights_since']
        until_date = date_config['insights_until']

        url = f"{self.base_url}/{page_id}/insights"
        params = {
            'metric': ','.join(self.config['PAGE_INSIGHTS_METRICS']),
            'period': self.config['PAGE_INSIGHTS_PERIOD'],
            'since': since_date,
            'until': until_date,
            'access_token': page_token
        }

        all_data = []
        while url:
            response = self.make_request(
                url,
                params if url == f"{self.base_url}/{page_id}/insights" else None,
                page_token
            )
            if not response:
                break
            all_data.extend(response.get('data', []))
            url = response.get('paging', {}).get('next')
            params = None

        return all_data

    def get_page_feed(self, page_id, page_token):
        url = f"{self.base_url}/{page_id}/feed"

        if self.config['FETCH_ALL_POSTS']:
            params = {
                'fields': 'id,icon,message,properties,created_time,full_picture,permalink_url',
                'limit': 100,
                'access_token': page_token
            }
        else:
            params = {
                'fields': 'id,icon,message,properties,created_time,full_picture,permalink_url',
                'since': date_config['feed_since'],
                'until': date_config['feed_until'],
                'limit': 100,
                'access_token': page_token
            }

        all_posts = []
        page_count = 0

        while url:
            response = self.make_request(
                url,
                params if url == f"{self.base_url}/{page_id}/feed" else None,
                page_token
            )
            if not response:
                break

            posts = response.get('data', [])
            all_posts.extend(posts)
            page_count += 1

            if self.config['FETCH_ALL_POSTS'] and page_count % 5 == 0:
                logger.info(f"   Fetched {len(all_posts)} posts so far...")

            url = response.get('paging', {}).get('next')
            params = None

        return all_posts

    def determine_post_type(self, icon, permalink_url):
        if not icon:
            return 'text'
        icon_lower = icon.lower()
        link_lower = (permalink_url or '').lower()

        if 'photo' in icon_lower:
            return 'photo'
        elif 'reel' in link_lower:
            return 'reel'
        elif 'video' in icon_lower:
            return 'video'
        return 'text'

    # ---------------------------------------------------------------------
    # Per-post processing
    # ---------------------------------------------------------------------
    def get_video_meta(self, post_id, page_token):
        """Resolve a post to its video node and return (attachment_id, meta).

        Two cheap requests. Runs before any insights call so that posts failing
        the duration gate are abandoned early.
        """
        att_resp = self.make_request(
            f"{self.base_url}/{post_id}",
            {'fields': 'attachments{target}', 'access_token': page_token},
            page_token
        )
        attachments = (att_resp or {}).get('attachments', {}).get('data', [])
        if not attachments:
            return None, {}

        attachment_id = (attachments[0].get('target') or {}).get('id')
        if not attachment_id:
            return None, {}

        meta = self.make_request(
            f"{self.base_url}/{attachment_id}",
            {'fields': 'custom_labels,views,length', 'access_token': page_token},
            page_token
        ) or {}
        return attachment_id, meta

    def process_single_post(self, post, page_id, page_name, page_token, existing_post_ids):
        """Process a single post — thread-safe, no DB calls.

        Returns {'ok': bool, 'reason': str|None, 'feed_data':…, 'label_data':…}.
        `reason` explains why a post was dropped so the caller can tally it.
        """
        post_id = post.get('id')
        try:
            post_type = self.determine_post_type(post.get('icon'), post.get('permalink_url', ''))

            # ---- GATE 1: video/reel only -------------------------------
            if self.config['VIDEO_ONLY'] and post_type not in ('video', 'reel'):
                return {'ok': False, 'reason': 'not_video',
                        'feed_data': None, 'label_data': None}

            # ---- GATE 2: duration --------------------------------------
            # Length lives on the video node, not the post, so resolve the
            # attachment first. Cost: 2 requests vs the ~6 a full fetch needs.
            attachment_id, video_meta = self.get_video_meta(post_id, page_token)
            if not attachment_id:
                return {'ok': False, 'reason': 'no_attachment',
                        'feed_data': None, 'label_data': None}

            try:
                duration = float(video_meta.get('length') or 0)
            except (TypeError, ValueError):
                duration = 0.0

            if duration <= self.config['MIN_DURATION_SECONDS']:
                return {'ok': False, 'reason': 'too_short',
                        'feed_data': None, 'label_data': None}

            # ---- Post qualifies: now spend the expensive calls ----------
            created_time = post.get('created_time', '').split('T')[0]

            insights_url = f"{self.base_url}/{post_id}/insights"
            base_params = {
                'metric': 'post_media_view,post_reactions_by_type_total,monetization_approximate_earnings',
                'period': self.config['POST_INSIGHTS_PERIOD'],
                'access_token': page_token
            }
            base_response = self.make_request(insights_url, base_params, page_token)
            insights_data = base_response.get('data', []) if base_response else []

            try:
                core_video_params = {
                    'metric': 'post_video_views,post_video_avg_time_watched',
                    'period': self.config['POST_INSIGHTS_PERIOD'],
                    'access_token': page_token
                }
                core_resp = self.make_request(insights_url, core_video_params, page_token)
                if core_resp and 'data' in core_resp:
                    insights_data.extend(core_resp['data'])
            except Exception as e:
                logger.error(f"Core video metrics error for {post_id}: {e}")

            try:
                views_60s_params = {
                    'metric': 'post_video_views_60s_excludes_shorter',
                    'period': self.config['POST_INSIGHTS_PERIOD'],
                    'access_token': page_token
                }
                v60_resp = self.make_request(insights_url, views_60s_params, page_token)
                if v60_resp and 'data' in v60_resp:
                    insights_data.extend(v60_resp['data'])
                elif v60_resp and 'error' in v60_resp:
                    logger.warning(f"60s excludes_shorter failed for {post_id}, trying fallback")
                    fallback_params = {
                        'metric': 'post_video_complete_views_60s',
                        'period': self.config['POST_INSIGHTS_PERIOD'],
                        'access_token': page_token
                    }
                    fallback_resp = self.make_request(insights_url, fallback_params, page_token)
                    if fallback_resp and 'data' in fallback_resp:
                        insights_data.extend(fallback_resp['data'])
            except Exception as e:
                logger.error(f"60s views metric error for {post_id}: {e}")

            shares_response = self.make_request(
                f"{self.base_url}/{post_id}",
                {'fields': 'shares', 'access_token': page_token},
                page_token
            )
            shares_count = shares_response.get('shares', {}).get('count', 0) if shares_response else 0

            custom_labels = video_meta.get('custom_labels') or []

            feed_data = {
                'page_id':                            page_id,
                'page_name':                          page_name,
                'post_id':                            post_id,
                'date':                               created_time,
                'created_time':                       created_time,
                'post_type':                          post_type,
                'message':                            post.get('message', ''),
                'permalink_url':                      post.get('permalink_url', ''),
                'impressions':                        0,
                'video_views':                        0,
                'reactions_total':                    0,
                'video_avg_time_watched':             0.0,
                'monetization_approximate_earnings':  0.0,
                'shares':                             shares_count,
                'views':                              video_meta.get('views', 0),
                'duration_seconds':                   duration,
                'video_views_60s':                    0,
                'custom_label':                       custom_labels[0] if custom_labels else 'none',
            }

            for insight in insights_data:
                metric_name = insight.get('name')
                values = insight.get('values', [])
                if not values:
                    continue
                value = values[0].get('value', 0)

                if metric_name == 'post_media_view':
                    feed_data['impressions'] = value
                elif metric_name == 'post_video_views':
                    feed_data['video_views'] = value
                elif metric_name == 'post_reactions_by_type_total':
                    feed_data['reactions_total'] = sum(value.values()) if isinstance(value, dict) else value
                elif metric_name == 'post_video_avg_time_watched':
                    feed_data['video_avg_time_watched'] = round(float(value) / 1000, 2) if value else 0.0
                elif metric_name == 'monetization_approximate_earnings':
                    feed_data['monetization_approximate_earnings'] = round(float(value), 6) if value else 0.0
                elif metric_name in ('post_video_views_60s_excludes_shorter', 'post_video_complete_views_60s'):
                    new_val = int(value) if value else 0
                    if new_val > feed_data['video_views_60s']:
                        feed_data['video_views_60s'] = new_val

            label_data = None
            if post_id not in existing_post_ids:
                label_data = {
                    'feed_id':       post_id,
                    'attachment_id': attachment_id,
                    'custom_label':  custom_labels[0] if custom_labels else 'none',
                    'post_type':     post_type,
                }

            return {'ok': True, 'reason': None,
                    'feed_data': feed_data, 'label_data': label_data}

        except Exception as e:
            logger.error(f"Error processing post {post_id}: {e}")
            return {'ok': False, 'reason': 'error',
                    'feed_data': None, 'label_data': None}


print("🔧 FacebookAPIClient class with threading support defined successfully!")


# ---
## Main Extraction Logic

client = FacebookAPIClient(CONFIG)

# Tracks post IDs already seen in this run (replaces DB existing_feed_ids check)
seen_post_ids: set = set()

# Reasons posts were dropped, aggregated across all pages
skip_tally = Counter()

if CONFIG['RESET_OUTPUT']:
    CSVWriter.reset_files()

# ── Step 1: List ALL pages the token can access ───────────────────────────────
client.list_accessible_pages()

# ── Step 2: Extract & export data ─────────────────────────────────────────────
try:
    pages = client.get_pages_and_tokens()

    if not pages or not pages[0].get('id'):
        print("❌ Could not retrieve page information")
    else:
        # --- Filter pages ---
        if not CONFIG['FETCH_ALL_PAGES']:
            original_count = len(pages)

            if CONFIG['SPECIFIC_PAGE_IDS']:
                pages = [p for p in pages if p['id'] in CONFIG['SPECIFIC_PAGE_IDS']]
                print(f"🔍 Filtered by Page IDs: {len(pages)}/{original_count} pages selected")
            elif CONFIG['SPECIFIC_PAGE_NAMES']:
                pages = [p for p in pages if p['name'] in CONFIG['SPECIFIC_PAGE_NAMES']]
                print(f"🔍 Filtered by Page Names: {len(pages)}/{original_count} pages selected")

            if not pages:
                print("❌ No pages match the specified criteria")
                raise SystemExit

        print(f"✅ Processing {len(pages)} page(s)")

        for page in pages:
            page_id    = page['id']
            page_token = page['access_token']
            page_name  = page['name']

            print(f"\n📄 Processing page: {page_name}")
            print("=" * 60)

            # ---- Page insights ----
            try:
                print("   📈 Fetching page insights...")
                page_insights = client.get_page_insights(page_id, page_token)

                insights_by_date = {}
                for insight in page_insights:
                    metric_name = insight.get('name')
                    for value_data in insight.get('values', []):
                        date = value_data.get('end_time', '').split('T')[0]
                        value = value_data.get('value', 0)

                        try:
                            date_obj = datetime.strptime(date, '%Y-%m-%d')
                            if date_obj >= (datetime.now() - timedelta(days=1)):
                                continue
                        except ValueError:
                            continue

                        if date not in insights_by_date:
                            insights_by_date[date] = {
                                'page_id': page_id, 'date': date,
                                'follows': 0, 'impressions': 0,
                                'post_engagements': 0,
                                'monetization_approximate_earnings': 0.0,
                            }

                        clean_metric = metric_name.replace('page_', '')
                        insights_by_date[date][clean_metric] = value

                if insights_by_date:
                    CSVWriter.add_page_insights(list(insights_by_date.values()))
                    print(f"   ✅ Collected {len(insights_by_date)} page insight records")
            except Exception as e:
                print(f"   ❌ Error collecting page insights: {e}")

            # ---- Feed posts (threaded) ----
            try:
                print("   📰 Fetching feed posts...")
                start_time = time.time()

                feed_posts = client.get_page_feed(page_id, page_token)
                print(f"   📊 Screening {len(feed_posts)} posts with {CONFIG['MAX_WORKERS']} threads "
                      f"(keeping video/reel > {CONFIG['MIN_DURATION_SECONDS']}s)...")

                feed_data_list  = []
                label_data_list = []
                page_skips = Counter()
                processed = 0

                with ThreadPoolExecutor(max_workers=CONFIG['MAX_WORKERS']) as executor:
                    future_to_post = {
                        executor.submit(
                            client.process_single_post,
                            post, page_id, page_name, page_token, seen_post_ids
                        ): post
                        for post in feed_posts
                    }

                    for future in as_completed(future_to_post):
                        result = future.result()
                        if result and result['ok']:
                            feed_data_list.append(result['feed_data'])
                            if result['label_data']:
                                label_data_list.append(result['label_data'])
                                seen_post_ids.add(result['label_data']['feed_id'])
                        elif result:
                            page_skips[result['reason']] += 1

                        processed += 1
                        if processed % 100 == 0:
                            print(f"      Screened {processed}/{len(feed_posts)} posts...", end='\r')

                skip_tally.update(page_skips)
                print(f"\n   ⏱️  Processing completed in {time.time() - start_time:.2f} seconds")
                if page_skips:
                    detail = ", ".join(f"{k}={v}" for k, v in sorted(page_skips.items()))
                    print(f"   ⏭️  Skipped {sum(page_skips.values())} posts  ({detail})")
                print(f"   🎬 Kept {len(feed_data_list)}/{len(feed_posts)} posts")

                if feed_data_list:
                    CSVWriter.add_feed_insights(feed_data_list)
                    print(f"   ✅ Collected {len(feed_data_list)} feed insight records")

                if label_data_list:
                    CSVWriter.add_custom_labels(label_data_list)
                    print(f"   ✅ Collected {len(label_data_list)} custom label records")

            except Exception as e:
                print(f"   ❌ Error processing feed: {e}")

            print(f"   💾 Saving data for '{page_name}' to CSV...")
            CSVWriter.flush(label=page_name)
            print(f"   ✅ Data saved.")

    # ---- Write all CSV files ----
    print("\n💾 Writing CSV files...")
    files_written = CSVWriter.save_all()

    if files_written:
        print("\n📂 Files written:")
        for fname, row_count in files_written:
            filepath = os.path.join(CSV_OUTPUT_DIR, fname)
            print(f"   • {fname}  ({row_count:,} rows)  →  {os.path.abspath(filepath)}")
    else:
        print("   ℹ️  No data was collected — no CSV files written.")

    if skip_tally:
        print("\n⏭️  Posts dropped by the filter:")
        labels = {
            'not_video':     'not a video/reel',
            'too_short':     f"video ≤ {CONFIG['MIN_DURATION_SECONDS']}s",
            'no_attachment': 'no video attachment found',
            'error':         'errored while fetching',
        }
        for reason, count in skip_tally.most_common():
            print(f"   • {labels.get(reason, reason):<28} {count:,}")

    print("\n🎉 Data extraction and CSV export completed!")

    # ---- Build the pivot workbook ----
    if CONFIG['BUILD_PIVOT']:
        feed_csv = os.path.join(CSV_OUTPUT_DIR, FEED_INSIGHTS_FILENAME)
        if not os.path.exists(feed_csv):
            print("\nℹ️  No feed rows survived the filter — skipping pivot build.")
        else:
            print("\n📊 Building pivot workbook...")
            try:
                from build_pivot_sheet import build as build_pivot
                pivot_path = build_pivot(
                    feed_csv,
                    os.path.join(CSV_OUTPUT_DIR, CONFIG['PIVOT_FILENAME']),
                    source_sheet='social_feed_insights',
                    pivot_sheet='Sheet1',
                )
                print(f"   ✅ Pivot workbook → {os.path.abspath(pivot_path)}")
            except ImportError:
                print("   ⚠️  build_pivot_sheet.py not found next to this script — "
                      "pivot not built. Run it manually:")
                print(f"      python build_pivot_sheet.py \"{feed_csv}\"")
            except Exception as e:
                print(f"   ❌ Pivot build failed: {e}")

except Exception as e:
    print(f"❌ Error during extraction: {e}")
    raise

print("\n✅ Process complete!")
