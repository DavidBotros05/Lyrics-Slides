"""Cloud web UI for the lyrics-slides maker (iPhone/iPad friendly).

Differences from the local app.py:
- Binds to 0.0.0.0 and reads the port from the PORT env var (cloud hosts
  like Render set this automatically).
- No native folder dialogs and no auto-opening browser.
- Finished .pptx files are NOT saved to a folder on this machine; they are
  kept in a temp dir and served back to the phone as a download
  (GET /download/<token>).

Run locally with:  python app.py   (defaults to port 8765)
"""

import base64
import json
import os
import secrets
import tempfile
import threading
import time
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlparse

import Create_Lyrics as core
from Background_for_Lyrics import (DEFAULT_TRANSFORMS, IDENTITY_TRANSFORM,
                                   pick_text_color)

CUR_DIR = str(Path(__file__).resolve().parent)
IMAGES_DIR = os.path.join(CUR_DIR, 'background_images')
UI_FILE = os.path.join(CUR_DIR, 'ui.html')
PORT = int(os.environ.get('PORT', '8765'))

# python-pptx slide building in Create_Lyrics uses module globals, so only
# one presentation may be built at a time.
_BUILD_LOCK = threading.Lock()

ALLOWED_IMAGE_EXTS = ('.png', '.jpg', '.jpeg')

# Finished presentations waiting to be downloaded: token -> (path, created_at).
_DOWNLOADS: dict[str, tuple[str, float]] = {}
_DOWNLOADS_LOCK = threading.Lock()
_DOWNLOAD_DIR = tempfile.mkdtemp(prefix='lyrics_slides_out_')
_DOWNLOAD_TTL = 60 * 60  # files stay downloadable for 1 hour

PPTX_MIME = ('application/vnd.openxmlformats-officedocument'
             '.presentationml.presentation')


def _register_download(path: str) -> str:
    token = secrets.token_urlsafe(16)
    now = time.time()
    with _DOWNLOADS_LOCK:
        # prune expired files
        for tok, (p, ts) in list(_DOWNLOADS.items()):
            if now - ts > _DOWNLOAD_TTL:
                _DOWNLOADS.pop(tok, None)
                try:
                    os.remove(p)
                except OSError:
                    pass
        _DOWNLOADS[token] = (path, now)
    return token


def list_background_images() -> list[str]:
    try:
        files = os.listdir(IMAGES_DIR)
    except FileNotFoundError:
        return []
    out = [f for f in files
           if not f.startswith(',') and not f.startswith('.')
           and f.lower().endswith(ALLOWED_IMAGE_EXTS)]
    return sorted(out, key=str.lower)


def fetch_song_for_request(data: dict) -> dict | None:
    """Fetch lyrics from a request payload (artist/title and/or one URL)."""
    artist = (data.get('artist') or '').strip()
    title = (data.get('title') or '').strip()
    url = (data.get('url') or '').strip()
    return core.fetch_song(artist, title, azlyrics_url='', genius_url=url)


class Handler(BaseHTTPRequestHandler):
    protocol_version = 'HTTP/1.1'

    # ---------- helpers ----------
    def _send_json(self, payload: dict, status: int = 200) -> None:
        body = json.dumps(payload).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_file(self, path: str, content_type: str,
                   download_name: str | None = None) -> None:
        try:
            with open(path, 'rb') as f:
                body = f.read()
        except OSError:
            self.send_error(404)
            return
        self.send_response(200)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(body)))
        if download_name:
            safe = download_name.replace('"', '')
            self.send_header('Content-Disposition',
                             f'attachment; filename="{safe}"')
        self.end_headers()
        self.wfile.write(body)

    def _read_json(self) -> dict:
        length = int(self.headers.get('Content-Length', 0) or 0)
        if length <= 0:
            return {}
        try:
            return json.loads(self.rfile.read(length).decode('utf-8'))
        except (json.JSONDecodeError, UnicodeDecodeError):
            return {}

    def log_message(self, fmt, *args):
        pass

    # ---------- GET ----------
    def do_GET(self):
        path = urlparse(self.path).path

        if path in ('/', '/index.html'):
            self._send_file(UI_FILE, 'text/html; charset=utf-8')
            return

        if path == '/api/backgrounds':
            settings = core.load_app_settings()
            colors = settings.get('background_colors', {}) or {}
            selected = settings.get('background_selected', {}) or {}
            transforms = settings.get('background_transforms', {}) or {}
            items = []
            for name in list_background_images():
                mode = colors.get(name, 'Auto')
                if mode not in ('Auto', 'Black', 'White'):
                    mode = 'Auto'
                transform = (transforms.get(name)
                             or DEFAULT_TRANSFORMS.get(name)
                             or IDENTITY_TRANSFORM)
                auto = pick_text_color(os.path.join(IMAGES_DIR, name))
                items.append({
                    'name': name,
                    'color_mode': mode,
                    'selected': bool(selected.get(name, True)),
                    'transform': transform,
                    'auto_color': 'white' if auto == (255, 255, 255) else 'black',
                })
            self._send_json({'ok': True, 'backgrounds': items})
            return

        if path.startswith('/download/'):
            token = path[len('/download/'):].strip('/')
            with _DOWNLOADS_LOCK:
                entry = _DOWNLOADS.get(token)
            if not entry or not os.path.isfile(entry[0]):
                self.send_error(404)
                return
            mime = ('application/zip' if entry[0].lower().endswith('.zip')
                    else PPTX_MIME)
            self._send_file(entry[0], mime,
                            download_name=os.path.basename(entry[0]))
            return

        if path.startswith('/backgrounds/'):
            name = os.path.basename(unquote(path[len('/backgrounds/'):]))
            full = os.path.join(IMAGES_DIR, name)
            if not name.lower().endswith(ALLOWED_IMAGE_EXTS) or not os.path.isfile(full):
                self.send_error(404)
                return
            ext = name.rsplit('.', 1)[-1].lower()
            ctype = 'image/png' if ext == 'png' else 'image/jpeg'
            self._send_file(full, ctype)
            return

        self.send_error(404)

    # ---------- POST ----------
    def do_POST(self):
        path = urlparse(self.path).path
        data = self._read_json()

        if path == '/api/preview':
            self.handle_preview(data)
        elif path == '/api/spotify_playlist':
            self.handle_spotify_playlist(data)
        elif path == '/api/create_one':
            self.handle_create_one(data)
        elif path == '/api/download_all':
            self.handle_download_all(data)
        elif path == '/api/save_settings':
            self.handle_save_settings(data)
        elif path == '/api/upload_background':
            self.handle_upload_background(data)
        else:
            self.send_error(404)

    def handle_preview(self, data: dict) -> None:
        song = fetch_song_for_request(data)
        if song is None:
            self._send_json({'ok': False,
                             'error': core.LAST_ERROR or core.LYRICS_NOT_FOUND_MSG})
            return
        self._send_json({
            'ok': True,
            'title': song['title'],
            'artist': song['artist'],
            'source': song['source'],
            'lyrics': '\n'.join(song['lyrics']),
        })

    def handle_spotify_playlist(self, data: dict) -> None:
        """Return every song (title + artist) in a Spotify playlist/album."""
        link = (data.get('link') or '').strip()
        result = core.fetch_spotify_playlist(link)
        if result is None:
            self._send_json({'ok': False,
                             'error': core.LAST_ERROR or
                                      "Couldn't read that Spotify playlist."})
            return
        self._send_json({
            'ok': True,
            'name': result['name'],
            'tracks': result['tracks'],
            'total': result['total'],
            'note': result.get('note', ''),
        })

    def handle_create_one(self, data: dict) -> None:
        """Create the .pptx for one song and return a download link."""
        manual_lyrics = data.get('lyrics') or ''
        if manual_lyrics.strip():
            song = core.song_from_manual_lyrics(
                artist=(data.get('artist') or '').strip(),
                title=(data.get('title') or '').strip(),
                genius_url=(data.get('url') or '').strip(),
                lyrics_text=manual_lyrics)
            if song is None:
                self._send_json({'ok': False, 'error': 'The lyrics box is empty.'})
                return
            if data.get('source'):
                song['source'] = data['source']
        else:
            song = fetch_song_for_request(data)
            if song is None:
                self._send_json({
                    'ok': False,
                    'error': core.LAST_ERROR or core.LYRICS_NOT_FOUND_MSG,
                    'needs_lyrics': True,
                })
                return

        background = None
        bg_name = (data.get('background') or '').strip()
        all_images = list_background_images()
        settings = core.load_app_settings()
        saved_colors = settings.get('background_colors', {}) or {}

        if bg_name and bg_name != 'random' and bg_name in all_images:
            background = {'image': bg_name,
                          'color_mode': saved_colors.get(bg_name, 'Auto')}

        saved_selected = settings.get('background_selected', {}) or {}
        selected_backgrounds = [
            {'image': n, 'color_mode': saved_colors.get(n, 'Auto')}
            for n in all_images if saved_selected.get(n, True)
        ]

        try:
            with _BUILD_LOCK:
                out_path = core.build_song_presentation(
                    song, _DOWNLOAD_DIR, background=background,
                    selected_backgrounds=selected_backgrounds)
        except Exception as exc:
            self._send_json({'ok': False,
                             'error': f'Failed to build the PowerPoint ({exc}).'})
            return

        token = _register_download(out_path)
        self._send_json({
            'ok': True,
            'file': os.path.basename(out_path),
            'download': f'/download/{token}',
            'title': song['title'],
            'artist': song['artist'],
            'source': song['source'],
        })

    def handle_download_all(self, data: dict) -> None:
        """Bundle the finished .pptx files into one named .zip download."""
        tokens = data.get('tokens') or []
        if not isinstance(tokens, list):
            tokens = []
        raw_name = (data.get('name') or '').strip()
        name = core.safe_filename(raw_name) if raw_name else 'Lyrics Slides'

        paths = []
        with _DOWNLOADS_LOCK:
            for tok in tokens:
                entry = _DOWNLOADS.get(str(tok))
                if entry and os.path.isfile(entry[0]) and entry[0] not in paths:
                    paths.append(entry[0])
        if not paths:
            self._send_json({'ok': False,
                             'error': 'No finished PowerPoints to bundle - '
                                      'create the slides first.'})
            return

        zip_path = os.path.join(_DOWNLOAD_DIR, f'{name}.zip')
        n = 2
        while os.path.exists(zip_path):
            zip_path = os.path.join(_DOWNLOAD_DIR, f'{name} ({n}).zip')
            n += 1

        try:
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                used = set()
                for p in paths:
                    arc = os.path.basename(p)
                    stem, ext = os.path.splitext(arc)
                    k = 2
                    while arc in used:
                        arc = f'{stem} ({k}){ext}'
                        k += 1
                    used.add(arc)
                    zf.write(p, arc)
        except OSError as exc:
            self._send_json({'ok': False,
                             'error': f"Couldn't build the zip ({exc})."})
            return

        token = _register_download(zip_path)
        self._send_json({
            'ok': True,
            'download': f'/download/{token}',
            'file': os.path.basename(zip_path),
            'count': len(paths),
        })

    def handle_save_settings(self, data: dict) -> None:
        settings = core.load_app_settings()
        if isinstance(data.get('background_colors'), dict):
            settings['background_colors'] = {
                str(k): (v if v in ('Auto', 'Black', 'White') else 'Auto')
                for k, v in data['background_colors'].items()
            }
        if isinstance(data.get('background_selected'), dict):
            settings['background_selected'] = {
                str(k): bool(v) for k, v in data['background_selected'].items()
            }
        if isinstance(data.get('background_transforms'), dict):
            clean = {}
            for name, t in data['background_transforms'].items():
                if not isinstance(t, dict):
                    continue
                try:
                    clean[str(name)] = {
                        'x': max(-5.0, min(5.0, float(t.get('x', 0.0)))),
                        'y': max(-5.0, min(5.0, float(t.get('y', 0.0)))),
                        'w': max(0.05, min(5.0, float(t.get('w', 1.0)))),
                        'h': max(0.05, min(5.0, float(t.get('h', 1.0)))),
                    }
                except (TypeError, ValueError):
                    continue
            settings['background_transforms'] = clean
        core.save_app_settings(settings)
        self._send_json({'ok': True})

    def handle_upload_background(self, data: dict) -> None:
        name = os.path.basename((data.get('name') or '').strip())
        content = data.get('data') or ''
        if not name.lower().endswith(ALLOWED_IMAGE_EXTS):
            self._send_json({'ok': False,
                             'error': 'Only .png, .jpg or .jpeg images.'})
            return
        try:
            raw = base64.b64decode(content.split(',', 1)[-1])
        except Exception:
            self._send_json({'ok': False, 'error': "Couldn't read the image data."})
            return
        if not raw or len(raw) > 30 * 1024 * 1024:
            self._send_json({'ok': False, 'error': 'Image is empty or over 30 MB.'})
            return
        os.makedirs(IMAGES_DIR, exist_ok=True)
        try:
            with open(os.path.join(IMAGES_DIR, name), 'wb') as f:
                f.write(raw)
        except OSError as exc:
            self._send_json({'ok': False, 'error': str(exc)})
            return
        self._send_json({'ok': True, 'name': name})


def main() -> None:
    server = ThreadingHTTPServer(('0.0.0.0', PORT), Handler)
    print(f'Lyrics Slides (cloud) listening on port {PORT}')
    server.serve_forever()


if __name__ == '__main__':
    main()
