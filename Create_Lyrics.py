from Background_for_Lyrics import creat_powerpoint_background, Background_image
import os
import random
import shutil
from pathlib import Path
import sys
import json
import tempfile

from pptx import Presentation
from pptx.util import Inches, Pt, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN, MSO_AUTO_SIZE

import re
import math

# Used by the direct Genius scraper below. Both ship with the project's venv.
# If either is missing we keep the exact import error so it can be shown to the
# user instead of a vague "packages are required" message.
_SCRAPE_IMPORT_ERROR = ''
try:
    import requests
    from bs4 import BeautifulSoup
except Exception as _exc:  # ImportError, or a broken install
    requests = None
    BeautifulSoup = None
    _SCRAPE_IMPORT_ERROR = str(_exc)

# A real browser User-Agent. Genius (and AZLyrics) serve an anti-bot page with
# no lyrics to requests that look automated, which is the usual reason a valid
# link "isn't found". Pretending to be a normal browser avoids that.
BROWSER_HEADERS = {
    'User-Agent': ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                   'AppleWebKit/537.36 (KHTML, like Gecko) '
                   'Chrome/124.0.0.0 Safari/537.36'),
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.9',
}

# Human-readable reason the most recent fetch failed (shown to the user).
LAST_ERROR = ''
LYRICS_NOT_FOUND_MSG = (
    "Unable to find lyrics. Please try again or paste them manually.")


class AZLyricsBlockedError(RuntimeError):
    """Raised when AZLyrics returns its browser-check / captcha page."""


class GeniusBlockedError(RuntimeError):
    """Raised when Genius/Cloudflare returns its human-verification page."""


GENIUS_BLOCKED_MSG = (
    "Genius is asking for human verification, so the lyrics can't be fetched "
    "from that link right now. Try searching by artist + title instead "
    "(uses LRCLIB), or paste the lyrics manually.")


def _short_err(exc: Exception, limit: int = 160) -> str:
    """One tidy line from an exception, never a wall of page text/headers."""
    msg = ' '.join(str(exc).split())
    return (msg[:limit] + '...') if len(msg) > limit else (msg or type(exc).__name__)


try:
    from azapi import AZlyrics
except ImportError:
    print("The 'azapi' package is required. Install it with:  pip install azapi")
    raise SystemExit(1)

# Genius is an optional backup source used when AZLyrics has no match.
try:
    from lyricsgenius import Genius
except ImportError:
    Genius = None

CUR_DIR = str(Path(__file__).resolve().parent)

# The Genius API token is a private key, so it is kept out of this file.
# It is read from a `.env` file next to this script (or a real environment
# variable). The `.env` file should contain a single line:
#     GENIUS_ACCESS_TOKEN=your_token_here
# Copy `.env.example` to `.env` and paste your token in. Keep `.env` private
# (it is listed in .gitignore so it is never committed).
try:
    from dotenv import load_dotenv
    load_dotenv(os.path.join(CUR_DIR, '.env'))
except ImportError:
    pass

GENIUS_TOKEN = os.environ.get('GENIUS_ACCESS_TOKEN', '')
SETTINGS_FILE = os.path.join(CUR_DIR, 'lyrics_slides_settings.json')


def load_app_settings() -> dict:
    """Load saved GUI choices like output folder and background color modes."""
    try:
        with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
            settings = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return {}

    return settings if isinstance(settings, dict) else {}


def save_app_settings(settings: dict) -> None:
    """Save GUI choices for the next run."""
    try:
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(settings, f, indent=2, sort_keys=True)
    except OSError:
        pass


def _title_from_url(url: str) -> str:
    """Best-effort song title from an AZLyrics/Genius URL slug."""
    if not url:
        return ''
    slug = url.rstrip('/').split('/')[-1]
    slug = slug.replace('.html', '').replace('-lyrics', '').replace('lyrics', '')
    slug = slug.replace('-', ' ').replace('_', ' ').strip()
    return slug.title()


def _detect_source(url: str) -> str:
    """Return 'genius', 'azlyrics', or '' based on the link's domain."""
    u = (url or '').lower()
    if 'genius.com' in u:
        return 'genius'
    if 'azlyrics.com' in u:
        return 'azlyrics'
    return ''


def _split_genius_meta_title(meta_title: str) -> tuple[str, str]:
    """Split Genius metadata like 'Artist - Song' into artist/title."""
    cleaned = (meta_title or '').replace('\xa0', ' ').strip()
    for sep in (' – ', ' — ', ' - '):
        if sep in cleaned:
            artist_name, song_title = cleaned.split(sep, 1)
            return artist_name.strip(), song_title.strip()
    return '', ''


def _require_scraper_packages() -> None:
    """Raise a clear error if requests/BeautifulSoup are unavailable."""
    if requests is not None and BeautifulSoup is not None:
        return

    detail = f": {_SCRAPE_IMPORT_ERROR}" if _SCRAPE_IMPORT_ERROR else ''
    raise RuntimeError(
        "the 'requests' and 'beautifulsoup4' packages are required but could "
        "not be imported by the Python running this script" + detail
        + ". Install them with:  pip install requests beautifulsoup4")


def _scrape_azlyrics_url(url: str) -> dict | None:
    """Scrape lyrics straight off an AZLyrics page.

    AZLyrics stores the lyric text in a plain div with no class/id. The direct
    scraper is used for pasted URLs because azapi can crash on valid pages.
    """
    _require_scraper_packages()

    resp = requests.get(url, headers=BROWSER_HEADERS, timeout=20)
    resp.raise_for_status()

    soup = BeautifulSoup(resp.text, 'html.parser')
    page_title = soup.title.get_text(strip=True) if soup.title else ''
    page_text = soup.get_text(' ', strip=True)
    if (
        'request for access' in page_title.lower()
        or 'b.azlyrics.com' in resp.url
        or 'detected unusual activity' in page_text.lower()
        or 'checking your browser' in page_text.lower()
    ):
        raise AZLyricsBlockedError(
            "AZLyrics blocked the request with a browser check/captcha. "
            "Paste the Genius link for this song instead.")

    title = ''
    b_tag = soup.find('b')
    if b_tag:
        title = b_tag.get_text(strip=True).strip('"')

    candidates = []
    for div in soup.find_all('div'):
        if div.get('class') or div.get('id'):
            continue
        text = div.get_text('\n').strip()
        if not text:
            continue
        if 'Usage of azlyrics.com content' in text:
            continue
        if len(text.splitlines()) < 3:
            continue
        candidates.append(text)

    if not candidates:
        return None

    # The lyrics block is normally the longest unclassified div.
    lyrics = max(candidates, key=len)
    return {'text': lyrics, 'title': title}


def fetch_from_azlyrics(artist: str = '', title: str = '', url: str = '') -> dict | None:
    """Look up a song's lyrics from AZLyrics via the azapi package.

    If `url` is given, the lyrics are pulled straight from that AZLyrics page
    instead of searching by artist/title.
    """
    global LAST_ERROR

    if url:
        try:
            scraped = _scrape_azlyrics_url(url)
        except AZLyricsBlockedError as exc:
            LAST_ERROR = str(exc)
            return None
        except Exception as exc:
            scraped = None
            LAST_ERROR = f"Couldn't load the AZLyrics page ({_short_err(exc)})."

        if scraped and scraped['text'].strip():
            return {
                'title': (title or scraped['title'] or _title_from_url(url)).strip() or 'song',
                'artist': artist.strip(),
                'lyrics': [ln.rstrip() for ln in scraped['text'].splitlines()],
                'source': 'AZLyrics',
            }

    api = AZlyrics('google', accuracy=0.5)
    api.artist = artist
    api.title = title

    try:
        lyrics = api.getLyrics(url=url, save=False) if url else api.getLyrics(save=False)
    except Exception as exc:
        LAST_ERROR = f"AZLyrics lookup failed ({_short_err(exc)})."
        return None

    # azapi returns the lyrics string on success, or an error code on failure.
    if not isinstance(lyrics, str) or not lyrics.strip():
        LAST_ERROR = LAST_ERROR or "AZLyrics returned no lyrics for that song."
        return None

    lines = [ln.rstrip() for ln in lyrics.splitlines()]

    # azapi corrects the artist/title spelling when it finds a match.
    final_title = (api.title or title or _title_from_url(url)).strip() or 'song'
    return {
        'title': final_title,
        'artist': (api.artist or artist).strip(),
        'lyrics': lines,
        'source': 'AZLyrics',
    }


def _scrape_genius_url(url: str) -> dict | None:
    """Scrape lyrics straight off a Genius song page, no API token needed.

    Returns a dict with 'text' and 'title', or None if the page has no lyrics
    container. Raises on network/HTTP errors so the caller can report them.
    """
    _require_scraper_packages()

    resp = requests.get(url, headers=BROWSER_HEADERS, timeout=20)

    # Cloudflare's "are you human?" challenge page — stop with a gentle
    # message instead of letting the raw challenge HTML leak into errors.
    body_start = (resp.text or '')[:3000].lower()
    if (
        resp.headers.get('Cf-Mitigated') == 'challenge'
        or "make sure you're a human" in body_start
        or 'checking your browser' in body_start
        or 'just a moment' in body_start
    ):
        raise GeniusBlockedError(GENIUS_BLOCKED_MSG)

    resp.raise_for_status()

    soup = BeautifulSoup(resp.text, 'html.parser')

    # Genius marks <br> as line breaks; turn them into real newlines first.
    for br in soup.find_all('br'):
        br.replace_with('\n')

    # Drop the "<Song> Lyrics" header block that sits inside the container.
    for bad in soup.find_all('div', class_=re.compile('LyricsHeader')):
        bad.decompose()

    containers = soup.find_all('div', attrs={'data-lyrics-container': 'true'})
    if not containers:
        return None

    text = '\n'.join(c.get_text() for c in containers)

    h1 = soup.find('h1')
    page_title = h1.get_text(strip=True) if h1 else ''
    artist_name = ''

    for key in ('og:title', 'twitter:title'):
        meta = soup.find('meta', attrs={'property': key}) or soup.find('meta', attrs={'name': key})
        meta_artist, meta_title = _split_genius_meta_title(meta.get('content', '') if meta else '')
        if meta_artist:
            artist_name = meta_artist
        if meta_title and not page_title:
            page_title = meta_title
        if artist_name:
            break

    return {'text': text, 'title': page_title, 'artist': artist_name}


def _clean_genius_lyrics(text: str) -> list[str]:
    """Strip the header/footer noise Genius adds around the lyrics body."""
    lines = text.splitlines()

    # The first line is usually a "<Song> Lyrics" header.
    if lines and lines[0].rstrip().endswith('Lyrics'):
        lines = lines[1:]

    cleaned: list[str] = []
    for ln in lines:
        ln = ln.rstrip()
        if ln.strip() == 'You might also like':
            continue
        # The final line often ends with a play count followed by "Embed".
        if ln.endswith('Embed'):
            ln = re.sub(r'\d*Embed$', '', ln).rstrip()
        cleaned.append(ln)

    return cleaned


LRCLIB_API = 'https://lrclib.net/api'
# LRCLIB asks clients to identify themselves; no token or signup needed.
LRCLIB_HEADERS = {'User-Agent': 'PowerpointLyrics/1.0',
                  'Lrclib-Client': 'PowerpointLyrics/1.0'}


def _lrclib_record_to_song(rec: dict, artist: str, title: str) -> dict | None:
    """Convert one LRCLIB API record into the project's song dict, or None."""
    if not isinstance(rec, dict) or rec.get('instrumental'):
        return None
    text = rec.get('plainLyrics') or ''
    if not text.strip():
        return None
    return {
        'title': (rec.get('trackName') or title).strip() or 'song',
        'artist': (rec.get('artistName') or artist).strip(),
        'lyrics': [ln.rstrip() for ln in text.splitlines()],
        'source': 'LRCLIB',
    }


def fetch_from_lrclib(artist: str = '', title: str = '') -> dict | None:
    """Look up a song on LRCLIB (lrclib.net) — free, no token, no anti-bot.

    Tries an exact match first (/api/get), then a fuzzy search (/api/search)
    and picks the first non-instrumental result with plain lyrics.
    """
    global LAST_ERROR
    if requests is None:
        LAST_ERROR = "The 'requests' package isn't installed."
        return None

    if not title:
        LAST_ERROR = 'LRCLIB needs at least a song title.'
        return None

    # 1) Exact match.
    try:
        resp = requests.get(f'{LRCLIB_API}/get',
                            params={'artist_name': artist, 'track_name': title},
                            headers=LRCLIB_HEADERS, timeout=15)
        if resp.ok:
            song = _lrclib_record_to_song(resp.json(), artist, title)
            if song:
                return song
    except Exception as exc:
        LAST_ERROR = f'LRCLIB lookup failed ({exc}).'

    # 2) Fuzzy search.
    try:
        resp = requests.get(f'{LRCLIB_API}/search',
                            params={'track_name': title, 'artist_name': artist}
                                   if artist else {'q': title},
                            headers=LRCLIB_HEADERS, timeout=15)
        resp.raise_for_status()
        results = resp.json()
        if isinstance(results, list):
            for rec in results:
                song = _lrclib_record_to_song(rec, artist, title)
                if song:
                    return song
    except Exception as exc:
        LAST_ERROR = f'LRCLIB search failed ({exc}).'
        return None

    LAST_ERROR = LAST_ERROR or 'No matching song found on LRCLIB.'
    return None


# ---------------------------------------------------------------------------
# Spotify playlist import
#
# Two ways to read a playlist, tried in this order:
#   1. The public embed page (open.spotify.com/embed/playlist/<id>) — no
#      account, token, or login needed. Works for any public playlist/album,
#      but very long playlists may be truncated (~100 tracks).
#   2. The official Web API — used automatically when SPOTIFY_CLIENT_ID and
#      SPOTIFY_CLIENT_SECRET are set in the .env file. Gets every track with
#      clean artist names, no truncation.
#
# Note: Spotify's public API does NOT expose lyrics (the lyrics inside the
# Spotify app come from Musixmatch through a private endpoint), so lyrics are
# still fetched per song from LRCLIB / Genius / AZLyrics as before.
# ---------------------------------------------------------------------------

SPOTIFY_CLIENT_ID = os.environ.get('SPOTIFY_CLIENT_ID', '')
SPOTIFY_CLIENT_SECRET = os.environ.get('SPOTIFY_CLIENT_SECRET', '')


def _spotify_parse_link(text: str) -> tuple[str, str]:
    """Return (kind, id) from a Spotify link/URI, e.g. ('playlist', '37i9…').

    Accepts open.spotify.com URLs (with or without locale prefixes or query
    strings), spotify:playlist:<id> URIs, and bare 22-character IDs.
    Kind is 'playlist' or 'album'; ('', '') when nothing matches.
    """
    t = (text or '').strip()
    if not t:
        return '', ''

    m = re.search(r'open\.spotify\.com/(?:[a-z\-]+/)??(playlist|album)/([A-Za-z0-9]{22})', t)
    if m:
        return m.group(1), m.group(2)

    m = re.match(r'spotify:(playlist|album):([A-Za-z0-9]{22})$', t)
    if m:
        return m.group(1), m.group(2)

    if re.fullmatch(r'[A-Za-z0-9]{22}', t):
        return 'playlist', t

    return '', ''


def _spotify_find_key(obj, key):
    """Depth-first search for the first value of `key` in nested JSON."""
    if isinstance(obj, dict):
        if key in obj:
            return obj[key]
        for v in obj.values():
            found = _spotify_find_key(v, key)
            if found is not None:
                return found
    elif isinstance(obj, list):
        for v in obj:
            found = _spotify_find_key(v, key)
            if found is not None:
                return found
    return None


def _spotify_from_embed(kind: str, spotify_id: str) -> dict | None:
    """Read a playlist/album from its public embed page (no credentials)."""
    _require_scraper_packages()

    url = f'https://open.spotify.com/embed/{kind}/{spotify_id}'
    resp = requests.get(url, headers=BROWSER_HEADERS, timeout=20)
    resp.raise_for_status()

    m = re.search(
        r'<script id="__NEXT_DATA__" type="application/json"[^>]*>(.*?)</script>',
        resp.text, re.S)
    if not m:
        return None

    data = json.loads(m.group(1))
    try:
        entity = data['props']['pageProps']['state']['data']['entity']
    except (KeyError, TypeError):
        entity = None
    track_list = (entity or {}).get('trackList') or _spotify_find_key(data, 'trackList')
    if not track_list:
        return None

    tracks = []
    for item in track_list:
        title = (item.get('title') or '').strip()
        artist = (item.get('subtitle') or '').strip()
        if title:
            tracks.append({'title': title, 'artist': artist})

    name = ((entity or {}).get('name')
            or _spotify_find_key(data, 'name') or '').strip()
    total = (entity or {}).get('trackCount') or len(tracks)
    return {'name': name or 'Spotify playlist', 'tracks': tracks,
            'total': int(total), 'method': 'embed'}


def _spotify_api_token() -> str:
    """Client-credentials token from the official API (no user login)."""
    resp = requests.post(
        'https://accounts.spotify.com/api/token',
        data={'grant_type': 'client_credentials'},
        auth=(SPOTIFY_CLIENT_ID, SPOTIFY_CLIENT_SECRET), timeout=15)
    resp.raise_for_status()
    return resp.json()['access_token']


def _spotify_from_api(kind: str, spotify_id: str) -> dict | None:
    """Read every track via the official Web API (needs client ID + secret)."""
    headers = {'Authorization': f'Bearer {_spotify_api_token()}'}
    base = f'https://api.spotify.com/v1/{kind}s/{spotify_id}'

    resp = requests.get(base, params={'fields': 'name'} if kind == 'playlist' else None,
                        headers=headers, timeout=15)
    resp.raise_for_status()
    name = (resp.json().get('name') or '').strip()

    tracks, url = [], f'{base}/tracks'
    params = {'limit': 100 if kind == 'playlist' else 50}
    while url:
        resp = requests.get(url, params=params, headers=headers, timeout=15)
        resp.raise_for_status()
        page = resp.json()
        for item in page.get('items', []):
            tr = item.get('track', item) or {}
            title = (tr.get('name') or '').strip()
            artists = ', '.join(a.get('name', '') for a in tr.get('artists', [])
                                if a.get('name'))
            if title:
                tracks.append({'title': title, 'artist': artists})
        url, params = page.get('next'), None  # 'next' already carries the query

    if not tracks:
        return None
    return {'name': name or 'Spotify playlist', 'tracks': tracks,
            'total': len(tracks), 'method': 'api'}


def fetch_spotify_playlist(link: str) -> dict | None:
    """Fetch the song list of a Spotify playlist/album link.

    Returns {'name', 'tracks': [{'title', 'artist'}, ...], 'total', 'note'}
    or None with LAST_ERROR set. Uses the official API when credentials are
    in .env, otherwise the public embed page (no account needed).
    """
    global LAST_ERROR
    LAST_ERROR = ''

    kind, spotify_id = _spotify_parse_link(link)
    if not spotify_id:
        LAST_ERROR = ("That doesn't look like a Spotify playlist link. Paste "
                      "one like https://open.spotify.com/playlist/...")
        return None
    if requests is None:
        LAST_ERROR = "The 'requests' package isn't installed."
        return None

    result = None

    # 1) Official API when credentials are available (full track list).
    if SPOTIFY_CLIENT_ID and SPOTIFY_CLIENT_SECRET:
        try:
            result = _spotify_from_api(kind, spotify_id)
        except Exception as exc:
            LAST_ERROR = f'Spotify API lookup failed ({_short_err(exc)}).'

    # 2) Public embed page — no credentials needed.
    if result is None:
        try:
            result = _spotify_from_embed(kind, spotify_id)
        except Exception as exc:
            LAST_ERROR = (LAST_ERROR or
                          f"Couldn't load that Spotify {kind} ({_short_err(exc)}).")
            return None

    if not result or not result['tracks']:
        LAST_ERROR = (LAST_ERROR or
                      f"No tracks found - is the {kind} public?")
        return None

    note = ''
    if result['method'] == 'embed' and result['total'] > len(result['tracks']):
        note = (f"Only the first {len(result['tracks'])} of {result['total']} "
                "tracks could be read without Spotify API credentials. Add "
                "SPOTIFY_CLIENT_ID and SPOTIFY_CLIENT_SECRET to the .env file "
                "to import the full playlist.")
    result['note'] = note
    return result


def fetch_from_genius(artist: str = '', title: str = '', url: str = '') -> dict | None:
    """Lookup from Genius.

    If `url` is given, the lyrics are scraped straight from that Genius page
    (no API token needed) using a browser-like request; lyricsgenius is tried
    as a backup. Searching by artist/title still needs an API token.
    """
    global LAST_ERROR

    if url:
        # 1) Direct scrape with a real browser User-Agent (most reliable).
        try:
            scraped = _scrape_genius_url(url)
        except GeniusBlockedError as exc:
            # Cloudflare wall: the lyricsgenius backup would hit it too and
            # dump the raw challenge page into the error, so stop here.
            LAST_ERROR = str(exc)
            return None
        except Exception as exc:
            scraped = None
            LAST_ERROR = f"Couldn't load the Genius page ({_short_err(exc)})."

        if scraped and scraped['text'].strip():
            return {
                'title': (title or scraped['title'] or _title_from_url(url)).strip() or 'song',
                'artist': (artist or scraped.get('artist', '')).strip(),
                'lyrics': _clean_genius_lyrics(scraped['text']),
                'source': 'Genius Lyrics',
            }

        # 2) Backup: let lyricsgenius try the same URL.
        if Genius is not None:
            try:
                genius = Genius(GENIUS_TOKEN or 'no-token',
                                remove_section_headers=True, timeout=15, retries=2)
                lyrics = genius.lyrics(song_url=url, remove_section_headers=True)
                if isinstance(lyrics, str) and lyrics.strip():
                    scraped_artist = scraped.get('artist', '') if scraped else ''
                    return {
                        'title': (title or _title_from_url(url)).strip() or 'song',
                        'artist': (artist or scraped_artist).strip(),
                        'lyrics': _clean_genius_lyrics(lyrics),
                        'source': 'Genius Lyrics',
                    }
            except Exception as exc:
                LAST_ERROR = f"Couldn't read lyrics from the Genius page ({_short_err(exc)})."

        LAST_ERROR = LAST_ERROR or "That Genius page didn't contain any lyrics text."
        return None

    # --- Search by artist + title (needs the lyricsgenius package + a token) ---
    if Genius is None:
        LAST_ERROR = "The 'lyricsgenius' package isn't installed."
        return None
    if not GENIUS_TOKEN:
        LAST_ERROR = "No Genius API token set, so searching by name isn't possible."
        return None

    try:
        genius = Genius(GENIUS_TOKEN, remove_section_headers=True,
                        timeout=15, retries=2)
        song = genius.search_song(title, artist)
    except Exception as exc:
        LAST_ERROR = f"Genius search failed ({_short_err(exc)})."
        return None

    if song is None or not getattr(song, 'lyrics', '').strip():
        LAST_ERROR = LAST_ERROR or "No matching song found on Genius."
        return None

    return {
        'title': (getattr(song, 'title', None) or title).strip(),
        'artist': (getattr(song, 'artist', None) or artist).strip(),
        'lyrics': _clean_genius_lyrics(song.lyrics),
        'source': 'Genius Lyrics',
    }


def fetch_song(artist: str = '', title: str = '',
               azlyrics_url: str = '', genius_url: str = '') -> dict | None:
    """Fetch lyrics for one song.

    A direct AZLyrics or Genius link takes priority; if none is given (or a
    link fails), it falls back to searching by artist + title. Returns the
    song dict, or None if nothing was found (with LAST_ERROR set to why).
    """
    global LAST_ERROR
    LAST_ERROR = ''

    az = (azlyrics_url or '').strip()
    gen = (genius_url or '').strip()

    # Route each pasted link by its domain, so a Genius link still works even
    # if it was typed into the AZLyrics box (and vice versa).
    if _detect_source(az) == 'genius':
        gen, az = az, ''
    if _detect_source(gen) == 'azlyrics':
        az, gen = gen, ''

    # LRCLIB is the default/first source: a free open API with no token and
    # no anti-bot wall. Search it with the typed artist/title, or — when only
    # a link was pasted — with the words from the link's slug.
    lr_artist, lr_title = artist, title
    if not lr_title and (gen or az):
        # A slug like "artist-song-lyrics" works well as a loose search, so
        # leave the artist box empty to trigger LRCLIB's keyword search.
        lr_artist, lr_title = '', _title_from_url(gen or az)
    if lr_title:
        song = fetch_from_lrclib(lr_artist, lr_title)
        if song:
            return song

    # Fall back to scraping a pasted link directly. Any Genius miss/error moves
    # on to AZLyrics instead of becoming the final user-facing error.
    if gen:
        song = fetch_from_genius(artist, title, url=gen)
        if song:
            return song
        if az:
            song = fetch_from_azlyrics(artist, title, url=az)
            if song:
                return song
        else:
            az_search_title = title or _title_from_url(gen)
            if az_search_title:
                song = fetch_from_azlyrics(artist, az_search_title)
                if song:
                    return song
        LAST_ERROR = LYRICS_NOT_FOUND_MSG
        return None

    if az:
        song = fetch_from_azlyrics(artist, title, url=az)
        if song:
            return song
        LAST_ERROR = LYRICS_NOT_FOUND_MSG
        return None

    # Last resort: Genius/AZLyrics search by name (needs both fields).
    if artist and title:
        song = fetch_from_genius(artist, title)
        if song:
            return song
        song = fetch_from_azlyrics(artist, title)
        if song:
            return song

    if not LAST_ERROR:
        LAST_ERROR = "Enter a song title (artist optional), or paste a Genius/AZLyrics link."
    elif LAST_ERROR != "Enter a song title (artist optional), or paste a Genius/AZLyrics link.":
        LAST_ERROR = LYRICS_NOT_FOUND_MSG
    return None


def song_from_manual_lyrics(artist: str = '', title: str = '',
                            azlyrics_url: str = '', genius_url: str = '',
                            lyrics_text: str = '') -> dict | None:
    """Build a song dict from manually pasted lyrics."""
    lines = [ln.rstrip() for ln in lyrics_text.splitlines()]
    if not any(ln.strip() for ln in lines):
        return None

    final_title = (title or _title_from_url(azlyrics_url or genius_url)).strip() or 'song'
    return {
        'title': final_title,
        'artist': artist.strip(),
        'lyrics': lines,
        'source': 'Manual Paste',
    }


def gather_manual_lyrics_terminal(artist: str = '', title: str = '',
                                  azlyrics_url: str = '',
                                  genius_url: str = '') -> dict | None:
    """Prompt for pasted lyrics in terminal mode."""
    use_manual = input("  Paste lyrics manually? [y/N]: ").strip().lower()
    if use_manual not in ('y', 'yes'):
        return None

    print("  Paste lyrics below. Type DONE on its own line when finished.")
    lines = []
    while True:
        line = input()
        if line.strip() == 'DONE':
            break
        lines.append(line)

    return song_from_manual_lyrics(
        artist, title, azlyrics_url, genius_url, '\n'.join(lines))


def gather_songs() -> list[dict]:
    """Ask for artist/title at the prompt and fetch each song's lyrics.

    Keeps asking until the artist is left blank.
    """
    songs: list[dict] = []
    print("Add songs (leave Artist blank to finish).")
    print("Tip: you can paste an AZLyrics or Genius link instead of artist/title.")
    while True:
        artist = input("Artist: ").strip()
        title = input("Song title: ").strip()
        azlyrics_url = input("AZLyrics link (optional): ").strip()
        genius_url = input("Genius link (optional): ").strip()

        if not (azlyrics_url or genius_url) and not (artist and title):
            # Nothing usable entered -> stop asking.
            break

        song = fetch_song(artist, title, azlyrics_url, genius_url)
        if song is None:
            label = artist_title_or_link(artist, title, azlyrics_url, genius_url)
            reason = LAST_ERROR or "not found on AZLyrics or Genius."
            print(f"  ERROR: {label} - {reason}")
            song = gather_manual_lyrics_terminal(
                artist, title, azlyrics_url, genius_url)
            if song is None:
                continue

        songs.append(song)
        print(f"  Found lyrics from {song['source']}: {song['title']} - {song['artist']}")

    return songs


def artist_title_or_link(artist: str, title: str, azlyrics_url: str, genius_url: str) -> str:
    """A short label for error messages, based on whatever was provided."""
    if artist and title:
        return f"'{title}' by {artist}"
    return azlyrics_url or genius_url or "the song"


def gather_songs_gui() -> tuple[list[dict], str, list[dict]] | None:
    """Pop up a window to enter songs, pick backgrounds, and choose an output folder.

    Returns (songs, output_folder, selected_backgrounds), or None if a graphical
    window can't open (so the caller can fall back to the terminal prompt).
    """
    try:
        import tkinter as tk
        from tkinter import messagebox, filedialog
    except Exception:
        return None

    try:
        root = tk.Tk()
    except Exception:
        return None

    root.title("Lyrics Slides")
    root.minsize(720, 720)
    root.columnconfigure(0, weight=1)

    try:
        from PIL import Image, ImageTk
    except Exception:
        Image = None
        ImageTk = None

    app_settings = load_app_settings()
    saved_bg_colors = app_settings.get('background_colors', {})
    if not isinstance(saved_bg_colors, dict):
        saved_bg_colors = {}
    saved_bg_selected = app_settings.get('background_selected', {})
    if not isinstance(saved_bg_selected, dict):
        saved_bg_selected = {}

    songs: list[dict] = []
    # One StringVar set per song row; rebuilt whenever the count changes.
    song_rows: list[dict] = []
    # Backgrounds the user ticked, captured when "Create slides" is pressed.
    chosen_backgrounds: list[dict] = []
    bg_thumbnail_refs: list = []
    scroll_areas: list[tuple] = []

    def is_descendant(widget, ancestor) -> bool:
        """Return True when widget is inside ancestor in the Tk widget tree."""
        while widget is not None:
            if widget == ancestor:
                return True
            widget = getattr(widget, 'master', None)
        return False

    def register_scroll_area(container, target_canvas):
        """Register a whole section that should scroll under the mouse wheel."""
        scroll_areas.append((container, target_canvas))

    def on_global_mousewheel(event):
        widget = root.winfo_containing(event.x_root, event.y_root)
        for container, target_canvas in reversed(scroll_areas):
            if not is_descendant(widget, container):
                continue

            if getattr(event, 'num', None) == 4:
                units = -1
            elif getattr(event, 'num', None) == 5:
                units = 1
            elif abs(event.delta) >= 120:
                units = int(-event.delta / 120)
            else:
                units = -1 if event.delta > 0 else 1

            target_canvas.yview_scroll(units, "units")
            return "break"
        return None

    root.bind_all("<MouseWheel>", on_global_mousewheel)
    root.bind_all("<Button-4>", on_global_mousewheel)
    root.bind_all("<Button-5>", on_global_mousewheel)

    # ----- Top bar: how many songs to enter -----
    top = tk.Frame(root)
    top.grid(row=0, column=0, padx=10, pady=8, sticky="w")
    tk.Label(top, text="How many songs?").pack(side="left")
    count_var = tk.StringVar(value="1")
    tk.Spinbox(top, from_=1, to=50, width=5, textvariable=count_var).pack(side="left", padx=6)
    tk.Button(top, text="Set", command=lambda: build_rows()).pack(side="left")

    # ----- Scrollable area that holds the per-song input rows -----
    middle = tk.Frame(root)
    middle.grid(row=1, column=0, padx=10, sticky="nsew")
    canvas = tk.Canvas(middle, borderwidth=0, width=660, height=230, highlightthickness=0)
    vsb = tk.Scrollbar(middle, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=vsb.set)
    vsb.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    rows_frame = tk.Frame(canvas)
    canvas.create_window((0, 0), window=rows_frame, anchor="nw")
    rows_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    register_scroll_area(middle, canvas)

    def build_rows():
        """Rebuild the input rows to match the requested song count."""
        try:
            n = int(count_var.get())
        except (TypeError, ValueError):
            n = 1
        n = max(1, min(50, n))

        for child in rows_frame.winfo_children():
            child.destroy()
        song_rows.clear()

        for i in range(n):
            box = tk.LabelFrame(rows_frame, text=f"Song {i + 1}", padx=6, pady=4)
            box.grid(row=i, column=0, padx=4, pady=4, sticky="w")

            artist_var = tk.StringVar()
            title_var = tk.StringVar()
            azlyrics_var = tk.StringVar()
            genius_var = tk.StringVar()

            tk.Label(box, text="Artist:").grid(row=0, column=0, sticky="e", padx=4, pady=2)
            tk.Entry(box, textvariable=artist_var, width=46).grid(row=0, column=1, padx=4, pady=2)
            tk.Label(box, text="Song title:").grid(row=1, column=0, sticky="e", padx=4, pady=2)
            tk.Entry(box, textvariable=title_var, width=46).grid(row=1, column=1, padx=4, pady=2)
            tk.Label(box, text="AZLyrics link (optional):").grid(row=2, column=0, sticky="e", padx=4, pady=2)
            tk.Entry(box, textvariable=azlyrics_var, width=46).grid(row=2, column=1, padx=4, pady=2)
            tk.Label(box, text="Genius link (optional):").grid(row=3, column=0, sticky="e", padx=4, pady=2)
            tk.Entry(box, textvariable=genius_var, width=46).grid(row=3, column=1, padx=4, pady=2)

            song_rows.append({
                'artist': artist_var,
                'title': title_var,
                'azlyrics': azlyrics_var,
                'genius': genius_var,
            })

    # ----- Background picker -----
    images_dir = os.path.join(CUR_DIR, 'background_images')
    bg_vars: dict = {}  # filename -> BooleanVar
    bg_color_vars: dict = {}  # filename -> StringVar: Auto, Black, or White

    def list_image_files():
        """Image files in the background_images folder (png/jpg/jpeg), skipping hidden ones."""
        try:
            files = os.listdir(images_dir)
        except FileNotFoundError:
            return []
        out = []
        for f in files:
            if f.startswith(','):
                continue
            low = f.lower()
            if low.endswith('.png') or low.endswith('.jpg') or low.endswith('.jpeg'):
                out.append(f)
        return sorted(out, key=str.lower)

    bg_outer = tk.LabelFrame(root, text="Backgrounds to use")
    bg_outer.grid(row=2, column=0, padx=10, pady=8, sticky="nsew")

    bg_inner = tk.Frame(bg_outer)
    bg_inner.pack(fill="both", expand=True)
    bg_canvas = tk.Canvas(bg_inner, borderwidth=0, width=660, height=260, highlightthickness=0)
    bg_vsb = tk.Scrollbar(bg_inner, orient="vertical", command=bg_canvas.yview)
    bg_canvas.configure(yscrollcommand=bg_vsb.set)
    bg_vsb.pack(side="right", fill="y")
    bg_canvas.pack(side="left", fill="both", expand=True)
    bg_list_frame = tk.Frame(bg_canvas)
    bg_canvas.create_window((0, 0), window=bg_list_frame, anchor="nw")
    bg_list_frame.bind(
        "<Configure>",
        lambda e: bg_canvas.configure(scrollregion=bg_canvas.bbox("all")))
    register_scroll_area(bg_outer, bg_canvas)

    def refresh_backgrounds():
        """Rebuild the checkbox list, keeping any current selections."""
        prev = {name: var.get() for name, var in bg_vars.items()}
        prev_colors = {name: var.get() for name, var in bg_color_vars.items()}
        for child in bg_list_frame.winfo_children():
            child.destroy()
        bg_vars.clear()
        bg_color_vars.clear()
        bg_thumbnail_refs.clear()
        files = list_image_files()
        if not files:
            tk.Label(bg_list_frame, text="(no images yet — use 'Add image...')",
                     fg="gray").grid(row=0, column=0, sticky="w", padx=4, pady=2)
            return
        tk.Label(bg_list_frame, text="Use").grid(row=0, column=0, padx=4, pady=(2, 4), sticky="w")
        tk.Label(bg_list_frame, text="Preview").grid(row=0, column=1, padx=4, pady=(2, 4), sticky="w")
        tk.Label(bg_list_frame, text="Background").grid(row=0, column=2, padx=4, pady=(2, 4), sticky="w")
        tk.Label(bg_list_frame, text="Text").grid(row=0, column=3, padx=4, pady=(2, 4), sticky="w")
        for i, f in enumerate(files):
            row = i + 1
            var = tk.BooleanVar(value=prev.get(f, bool(saved_bg_selected.get(f, True))))
            saved_color = saved_bg_colors.get(f, "Auto")
            if saved_color not in ("Auto", "Black", "White"):
                saved_color = "Auto"
            color_var = tk.StringVar(value=prev_colors.get(f, saved_color))

            tk.Checkbutton(bg_list_frame, variable=var, anchor="w").grid(
                row=row, column=0, sticky="w", padx=4, pady=3)
            preview_loaded = False
            if Image is not None and ImageTk is not None:
                try:
                    with Image.open(os.path.join(images_dir, f)) as im:
                        im.thumbnail((104, 62))
                        photo = ImageTk.PhotoImage(im.copy())
                    bg_thumbnail_refs.append(photo)
                    tk.Label(bg_list_frame, image=photo, width=110, height=66).grid(
                        row=row, column=1, sticky="w", padx=4, pady=3)
                    preview_loaded = True
                except Exception:
                    preview_loaded = False
            if not preview_loaded:
                tk.Label(bg_list_frame, text="No preview", width=14, fg="gray").grid(
                    row=row, column=1, sticky="w", padx=4, pady=3)

            tk.Label(bg_list_frame, text=f, anchor="w", width=34).grid(
                row=row, column=2, sticky="w", padx=4, pady=3)
            tk.OptionMenu(bg_list_frame, color_var, "Auto", "Black", "White").grid(
                row=row, column=3, sticky="w", padx=4, pady=3)
            bg_vars[f] = var
            bg_color_vars[f] = color_var

    def select_all_bg():
        for var in bg_vars.values():
            var.set(True)

    def unselect_all_bg():
        for var in bg_vars.values():
            var.set(False)

    def add_image():
        path = filedialog.askopenfilename(
            title="Choose an image to add to your backgrounds",
            filetypes=[("Background images", "*.png *.jpg *.jpeg"), ("All files", "*.*")])
        if not path:
            return
        os.makedirs(images_dir, exist_ok=True)
        dest = os.path.join(images_dir, os.path.basename(path))
        try:
            shutil.copy2(path, dest)
        except Exception as exc:
            messagebox.showerror("Couldn't add image", str(exc))
            return
        refresh_backgrounds()

    bg_btns = tk.Frame(bg_outer)
    bg_btns.pack(fill="x", pady=4)
    tk.Button(bg_btns, text="Select all", command=select_all_bg).pack(side="left", padx=4)
    tk.Button(bg_btns, text="Unselect all", command=unselect_all_bg).pack(side="left", padx=4)
    tk.Button(bg_btns, text="Add image...", command=add_image).pack(side="left", padx=4)

    status_var = tk.StringVar(value="")
    tk.Label(root, textvariable=status_var, fg="gray").grid(row=3, column=0)

    # ----- Output folder picker -----
    saved_output_folder = app_settings.get('output_folder', CUR_DIR)
    if not isinstance(saved_output_folder, str) or not saved_output_folder.strip():
        saved_output_folder = CUR_DIR
    output_dir_var = tk.StringVar(value=saved_output_folder)

    def browse_output():
        chosen = filedialog.askdirectory(
            initialdir=output_dir_var.get() or CUR_DIR,
            title="Choose where to save the PowerPoints")
        if chosen:
            output_dir_var.set(chosen)

    def gather_manual_lyrics_gui(artist: str = '', title: str = '',
                                 azlyrics_url: str = '',
                                 genius_url: str = '',
                                 reason: str = '') -> dict | None:
        """Open a modal paste box for lyrics when web lookup fails."""
        label = artist_title_or_link(artist, title, azlyrics_url, genius_url)
        if not messagebox.askyesno(
                "Lyrics not found",
                f"{label}\n\n{reason}\n\nPaste the lyrics manually?"):
            return None

        result = {'song': None}
        win = tk.Toplevel(root)
        win.title("Manual lyrics")
        win.transient(root)
        win.grab_set()

        title_var = tk.StringVar(value=(title or _title_from_url(azlyrics_url or genius_url)))
        artist_var = tk.StringVar(value=artist)

        tk.Label(win, text="Song title:").grid(row=0, column=0, sticky="e", padx=8, pady=6)
        tk.Entry(win, textvariable=title_var, width=42).grid(row=0, column=1, sticky="we", padx=8, pady=6)
        tk.Label(win, text="Artist:").grid(row=1, column=0, sticky="e", padx=8, pady=6)
        tk.Entry(win, textvariable=artist_var, width=42).grid(row=1, column=1, sticky="we", padx=8, pady=6)
        tk.Label(win, text="Lyrics:").grid(row=2, column=0, sticky="ne", padx=8, pady=6)
        lyrics_box = tk.Text(win, width=58, height=18, wrap="word")
        lyrics_box.grid(row=2, column=1, sticky="nsew", padx=8, pady=6)

        btns = tk.Frame(win)
        btns.grid(row=3, column=0, columnspan=2, pady=8)

        def save_manual():
            song = song_from_manual_lyrics(
                artist_var.get(), title_var.get(), azlyrics_url, genius_url,
                lyrics_box.get("1.0", "end-1c"))
            if song is None:
                messagebox.showwarning("No lyrics", "Paste at least one lyric line.")
                return
            result['song'] = song
            win.destroy()

        tk.Button(btns, text="Use pasted lyrics", command=save_manual).pack(side="left", padx=4)
        tk.Button(btns, text="Cancel", command=win.destroy).pack(side="left", padx=4)

        win.columnconfigure(1, weight=1)
        win.rowconfigure(2, weight=1)
        lyrics_box.focus_set()
        root.wait_window(win)
        return result['song']

    output_frame = tk.Frame(root)
    output_frame.grid(row=4, column=0, sticky="we", padx=10, pady=8)
    output_frame.columnconfigure(1, weight=1)
    tk.Label(output_frame, text="Output folder:").grid(row=0, column=0, sticky="w", padx=(0, 8))
    tk.Entry(output_frame, textvariable=output_dir_var).grid(row=0, column=1, sticky="we", padx=4)
    tk.Button(output_frame, text="Browse...", command=browse_output).grid(row=0, column=2, padx=(8, 0))

    def create_slides():
        """Fetch every filled-in row, then close the window."""
        entries = []
        for r in song_rows:
            artist = r['artist'].get().strip()
            title = r['title'].get().strip()
            azlyrics_url = r['azlyrics'].get().strip()
            genius_url = r['genius'].get().strip()
            if (azlyrics_url or genius_url) or (artist and title):
                entries.append((artist, title, azlyrics_url, genius_url))

        if not entries:
            messagebox.showwarning(
                "Nothing to create",
                "Fill in at least one song (artist + title, or a link).")
            return

        create_btn.config(state="disabled")
        songs.clear()
        not_found = []
        for i, (artist, title, azlyrics_url, genius_url) in enumerate(entries, 1):
            status_var.set(f"Searching {i} of {len(entries)}...")
            root.update_idletasks()
            song = fetch_song(artist, title, azlyrics_url, genius_url)
            if song is None:
                label = artist_title_or_link(artist, title, azlyrics_url, genius_url)
                reason = LAST_ERROR or "not found on AZLyrics or Genius."
                song = gather_manual_lyrics_gui(
                    artist, title, azlyrics_url, genius_url, reason)
                if song is None:
                    not_found.append(f"{label} - {reason}")
                    continue

            songs.append(song)
            print(f"Found lyrics from {song['source']}: {song['title']} - {song['artist']}")

        create_btn.config(state="normal")

        if not_found:
            messagebox.showwarning(
                "Some songs not found",
                "These were not found on AZLyrics or Genius:\n\n- "
                + "\n- ".join(not_found))

        if not songs:
            status_var.set("")
            return

        chosen_backgrounds[:] = [
            {'image': name, 'color_mode': bg_color_vars[name].get()}
            for name, var in bg_vars.items()
            if var.get()
        ]
        app_settings['output_folder'] = output_dir_var.get().strip() or CUR_DIR
        app_settings['background_colors'] = {
            name: var.get()
            for name, var in bg_color_vars.items()
        }
        app_settings['background_selected'] = {
            name: var.get()
            for name, var in bg_vars.items()
        }
        save_app_settings(app_settings)
        root.destroy()

    create_btn = tk.Button(root, text="Create slides", command=create_slides)
    create_btn.grid(row=5, column=0, pady=(4, 12))

    build_rows()           # start with one song row
    refresh_backgrounds()  # load the available background images
    root.mainloop()
    return songs, (output_dir_var.get().strip() or CUR_DIR), chosen_backgrounds


def new_pres(title: str, artist: str) -> None:
    """Create a new presentation and add the title slide for one song."""
    global P, W, H
    P = Presentation()
    slide0 = P.slides.add_slide(P.slide_layouts[0])
    slide0.placeholders[0].text = title
    slide0.placeholders[1].text = artist
    W, H = P.slide_width, P.slide_height


def is_marker(line: str) -> bool:
    """Return True for section markers like [Chorus], (Verse 1), or lines containing 'summary'."""
    if not line:
        return False
    bracketed = (line[0] in ('[', '(')) and (line[-1] in (']', ')'))
    return bracketed or ('summary' in line.lower())


def split_lengths(n: int) -> list[int]:
    """Split n lines into slide-sized chunks.

    Rules:
      - 3, 4, 5 stay on one slide
      - 6 -> 3+3
      - 7 -> 4+3
      - 8 -> 4+4
      - 9 -> 5+4
      - 10 -> 4+3+3 (etc.)
      - 11 -> 4+4+3 (etc.)

    Uses 4s and 3s for long groups, and uses a 5 when it avoids an awkward remainder.
    """
    if n <= 0:
        return []
    if n <= 5:
        return [n]

    out: list[int] = []
    remaining = n

    # Greedy with remainder-fixing using 5s when remainder would be 1.
    while remaining > 0:
        if remaining in (3, 4, 5):
            out.append(remaining)
            break
        if remaining == 6:
            out.extend([3, 3])
            break
        if remaining == 7:
            out.extend([4, 3])
            break
        if remaining == 8:
            out.extend([4, 4])
            break
        if remaining == 10:
            out.extend([4, 3, 3])
            break
        if remaining == 11:
            out.extend([4, 4, 3])
            break

        r = remaining % 4
        if r == 0:
            out.append(4)
            remaining -= 4
        elif r == 1:
            # Prefer a 5 so the rest is divisible by 4
            out.append(5)
            remaining -= 5
        else:
            # r == 2 or r == 3 -> take 3, then split the rest into 3/4/5-line slides.
            out.append(3)
            remaining -= 3

    return out


def split_long_lyric_lines(lines: list[str], max_chars: int = 36) -> list[str]:
    """Split long lyric lines before deciding how many lines go on a slide."""
    out: list[str] = []
    for line in lines:
        out.extend(split_long_lyric_line(line, max_chars))

    return out


def split_long_lyric_line(line: str, max_chars: int = 36) -> list[str]:
    """Split one long lyric line into balanced display lines."""
    if len(line) <= max_chars:
        return [line]

    parenthetical_split = split_line_at_parentheses(line, max_chars)
    if parenthetical_split is not None:
        return parenthetical_split

    return split_line_balanced(line, max_chars)


def split_line_balanced(line: str, max_chars: int) -> list[str]:
    """Wrap a lyric line with balanced lengths and no avoidable orphan words."""
    words = line.split()
    if len(words) <= 1:
        return [line]

    lengths = [len(word) for word in words]
    prefix = [0]
    for length in lengths:
        prefix.append(prefix[-1] + length)

    def segment_length(start: int, end: int) -> int:
        return prefix[end] - prefix[start] + (end - start - 1)

    total_len = segment_length(0, len(words))
    min_lines = max(2, math.ceil(total_len / max_chars))
    max_lines = min(len(words), min_lines + 2)
    best_parts: list[str] | None = None
    best_cost: float | None = None

    for line_count in range(min_lines, max_lines + 1):
        target = total_len / line_count
        dp: list[list[tuple[float, list[int]] | None]] = [
            [None] * (line_count + 1) for _ in range(len(words) + 1)
        ]
        dp[0][0] = (0.0, [])

        for end in range(1, len(words) + 1):
            for used in range(1, line_count + 1):
                best_here: tuple[float, list[int]] | None = None
                for start in range(used - 1, end):
                    prev = dp[start][used - 1]
                    if prev is None:
                        continue

                    length = segment_length(start, end)
                    overflow = max(0, length - max_chars)
                    word_count = end - start
                    orphan_penalty = 900 if word_count == 1 and len(words) > line_count else 0
                    cost = (
                        prev[0]
                        + ((length - target) ** 2)
                        + (overflow * overflow * 200)
                        + orphan_penalty
                    )
                    if best_here is None or cost < best_here[0]:
                        best_here = (cost, prev[1] + [start])

                dp[end][used] = best_here

        result = dp[len(words)][line_count]
        if result is None:
            continue

        overflow = sum(
            max(0, segment_length(start, end) - max_chars)
            for start, end in zip(result[1], result[1][1:] + [len(words)])
        )
        cost = result[0] + line_count * 10 + overflow * 10000
        if best_cost is None or cost < best_cost:
            best_cost = cost
            parts = []
            for start, end in zip(result[1], result[1][1:] + [len(words)]):
                parts.append(' '.join(words[start:end]))
            best_parts = parts

    return best_parts or [line]


def split_line_at_parentheses(line: str, max_chars: int) -> list[str] | None:
    """Prefer splitting a long line around parentheses when both sides fit."""
    break_points = []
    for match in re.finditer(r'\([^)]*\)', line):
        break_points.append(match.start())
        break_points.append(match.end())

    # Try the most balanced parenthesis break first.
    for pos in sorted(set(break_points), key=lambda p: abs(len(line) / 2 - p)):
        left = line[:pos].strip()
        right = line[pos:].strip()
        if not left or not right:
            continue
        if len(left) <= max_chars and len(right) <= max_chars:
            return [left, right]

    return None


def split_lyric_group_for_slides(lines: list[str], max_chars: int = 36) -> list[list[str]]:
    """Split a lyric group into slide chunks without separating wrapped lines."""
    units = [split_long_lyric_line(line, max_chars) for line in lines]
    if not units:
        return []

    best: list[tuple[float, list[list[str]]] | None] = [None] * (len(units) + 1)
    best[0] = (0.0, [])

    for end in range(1, len(units) + 1):
        count = 0
        for start in range(end - 1, -1, -1):
            count += len(units[start])
            if count > 5:
                break

            prev = best[start]
            if prev is None:
                continue

            if count == 4:
                chunk_cost = 0
            elif count in (3, 5):
                chunk_cost = 8
            elif count == 2:
                chunk_cost = 90
            else:
                chunk_cost = 220

            if end != len(units) and count < 3:
                chunk_cost += 500

            chunk: list[str] = []
            for unit in units[start:end]:
                chunk.extend(unit)

            candidate = (prev[0] + chunk_cost, prev[1] + [chunk])
            if best[end] is None or candidate[0] < best[end][0]:
                best[end] = candidate

    if best[-1] is None:
        return [split_long_lyric_lines(lines, max_chars)]
    return best[-1][1]


def add_lyrics_slide(lines: list[str]) -> None:
    """Add one lyric slide with a textbox configuration based on line count."""
    global P, W, H

    n = len(lines)
    # Tune these to your taste. The goal is to keep the block visually centered.
    # Each tuple: (left, top, width, height, font_size)
    layout = {
        1: (Cm(0.59), Cm(6.97), Cm(24.21), Cm(4.00), Pt(40)),
        2: (Cm(0.59), Cm(6.22), Cm(24.21), Cm(5.50), Pt(38)),
        3: (Cm(0.59), Cm(5.56), Cm(24.21), Cm(7.92), Pt(36)),
        4: (Cm(0.59), Cm(4.23), Cm(24.21), Cm(10.06), Pt(36)),
        5: (Cm(0.59), Cm(3.06), Cm(24.21), Cm(12.93), Pt(34)),
    }

    # Fallback if something unexpected happens
    left, top, width, height, font_size = layout.get(n, (Cm(2.03), Cm(2.78), W - Inches(1.6), H - Inches(3.0), Pt(36)))

    slide = P.slides.add_slide(P.slide_layouts[6])
    tx = slide.shapes.add_textbox(left, top, width, height)
    tf = tx.text_frame
    tf.clear()
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    # Wrap any line that is wider than the textbox onto extra lines, and shrink
    # the text just enough to stay inside the box if wrapping makes it too tall.
    # Slides whose lines already fit are left at their normal size.
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

    for idx, line in enumerate(lines):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.text = line
        p.font.size = font_size
        p.font.bold = True
        p.font.color.rgb = RGBColor(0, 0, 0)
        p.alignment = PP_ALIGN.CENTER

        # Line spacing: keep the first line default, apply 2.0 to all following lines
        if idx == 0:
            p.line_spacing = 1.0
        else:
            p.line_spacing = 2.0

        p.space_after = Pt(0)


def safe_filename(name: str) -> str:
    """Make a filesystem-safe filename (keeps spaces, removes forbidden chars)."""
    bad = ['/', '\\', ':', '*', '?', '"', '<', '>', '|']
    out = name.strip()
    for ch in bad:
        out = out.replace(ch, '-')
    return out if out else 'song'


def _background_info(prs, background: dict | None,
                     selected_backgrounds: list[dict] | None):
    """Resolve which background image/text-color/layout to use for one song.

    `background` is a specific choice like {'image': name, 'color_mode': mode}.
    If it's None (or its image is 'random'), one is picked at random from
    `selected_backgrounds`, falling back to a random image from the folder.
    The image's saved layout (position/stretch from the UI's "Adjust" editor,
    stored in lyrics_slides_settings.json) is always applied.
    """
    images_dir = CUR_DIR + '/background_images'

    chosen = None
    if background and background.get('image') and background['image'] != 'random':
        chosen = background
    elif selected_backgrounds:
        chosen = random.choice(selected_backgrounds)

    if chosen is None:
        # Random pick from the whole folder.
        try:
            files = [f for f in os.listdir(images_dir)
                     if not f.startswith((',', '.'))
                     and f.lower().endswith(('.png', '.jpg', '.jpeg'))]
        except FileNotFoundError:
            files = []
        if not files:
            return Background_image(images_dir, prs)  # keep original error path
        chosen = {'image': random.choice(files), 'color_mode': 'Auto'}

    transforms = load_app_settings().get('background_transforms', {}) or {}
    transform = transforms.get(chosen['image'])

    color_mode = chosen.get('color_mode', 'Auto')
    if color_mode == 'Black':
        return Background_image(
            images_dir, prs, source=chosen['image'],
            color=(0, 0, 0), auto_color=False, transform=transform)
    if color_mode == 'White':
        return Background_image(
            images_dir, prs, source=chosen['image'],
            color=(255, 255, 255), auto_color=False, transform=transform)
    return Background_image(
        images_dir, prs, source=chosen['image'], transform=transform)


def change_powerpoint_background(pptx_path: str, output_path: str,
                                 background: dict | None = None,
                                 selected_backgrounds: list[dict] | None = None) -> str:
    """Apply a new saved/random background to an existing .pptx file."""
    prs = Presentation(pptx_path)
    info_pic = _background_info(prs, background, selected_backgrounds)
    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    creat_powerpoint_background(pptx_path, output_path, info_pic)
    return output_path


def build_song_presentation(song: dict, output_base: str,
                            background: dict | None = None,
                            selected_backgrounds: list[dict] | None = None) -> str:
    """Build the finished .pptx for one song and return its file path."""
    title = song['title']
    artist = song['artist']
    lyric_lines: list[str] = song['lyrics']

    new_pres(title, artist)

    # 1) Build groups separated by empty lines (within this song)
    groups: list[list[str]] = []
    current: list[str] = []

    for raw in lyric_lines:
        if raw == '':
            if current:
                groups.append(current)
                current = []
            continue

        if is_marker(raw):
            continue

        current.append(raw)

    if current:
        groups.append(current)

    # 2) For each group, split into slide-sized chunks and render
    for g in groups:
        for chunk in split_lyric_group_for_slides(g):
            add_lyrics_slide(chunk)

    out_name = safe_filename(title)
    bg_path = os.path.join(output_base, f'{out_name}.pptx')

    temp_pptx = tempfile.NamedTemporaryFile(
        suffix='.pptx', prefix='lyrics_slides_', delete=False)
    temp_pptx.close()
    pptx_path = temp_pptx.name
    P.save(str(pptx_path))

    info_pic = _background_info(P, background, selected_backgrounds)
    try:
        creat_powerpoint_background(pptx_path, bg_path, info_pic)
    finally:
        try:
            os.remove(pptx_path)
        except OSError:
            pass

    return bg_path


def main() -> None:
    gui_result = gather_songs_gui()
    if gui_result is None:
        print("Graphical window not available, using the terminal instead.")
        songs = gather_songs()
        output_base = CUR_DIR
        selected_backgrounds = []
    else:
        songs, output_base, selected_backgrounds = gui_result

    if not songs:
        print("No songs to create. Exiting.")
        raise SystemExit(0)

    # Save the finished PowerPoints directly under the chosen output folder.
    os.makedirs(output_base, exist_ok=True)

    # Log created presentations to a file (matches the original behavior).
    console = sys.stdout
    sys.stdout = open(CUR_DIR + "/songs done.txt", "a")

    for song in songs:
        note = f" (lyrics from {song['source']})" if song.get('source') else ''
        print(safe_filename(song['title']) + note)
        build_song_presentation(
            song, output_base, selected_backgrounds=selected_backgrounds)

    print('Presentation created =', len(songs))
    print()
    sys.stdout.close()
    sys.stdout = console
    print(f"Done. Created {len(songs)} presentation(s).")


if __name__ == '__main__':
    main()
