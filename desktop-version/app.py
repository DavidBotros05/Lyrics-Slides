"""Local web UI for the lyrics-slides maker.

Run with:
    python app.py
Then the browser opens at http://127.0.0.1:8765 automatically.

Uses only the Python standard library for the server, and reuses the
lyric-fetching / slide-building code in Create_Lyrics.py.
"""

import base64
import json
import os
import subprocess
import sys
import threading
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlparse

import Create_Lyrics as core
from Background_for_Lyrics import (DEFAULT_TRANSFORMS, IDENTITY_TRANSFORM,
                                   pick_text_color)

CUR_DIR = str(Path(__file__).resolve().parent)
IMAGES_DIR = os.path.join(CUR_DIR, 'background_images')
UI_FILE = os.path.join(CUR_DIR, 'ui.html')
PORT = 8765

# python-pptx slide building in Create_Lyrics uses module globals, so only
# one presentation may be built at a time.
_BUILD_LOCK = threading.Lock()

ALLOWED_IMAGE_EXTS = ('.png', '.jpg', '.jpeg')

# ---- auto-shutdown when the browser tab goes away -------------------------
# The page pings /api/ping every couple of seconds and sends a final
# "closing" beacon when the tab is closed. If the tab disappears (or the
# user clicks Quit in the UI), the server stops itself - no Ctrl+C needed.
_ALIVE = {'last_ping': 0.0, 'closing_at': 0.0}
_QUIT = threading.Event()

# Close grace: short, but long enough that a page refresh (which also fires
# the closing beacon) can reconnect first. Idle grace: long, so a throttled
# background tab (browsers slow timers to ~1/min) doesn't get killed.
CLOSE_GRACE_SECS = 4.0
IDLE_GRACE_SECS = 150.0


def _note_ping(closing: bool) -> None:
    now = time.time()
    if closing:
        _ALIVE['closing_at'] = now
    else:
        _ALIVE['last_ping'] = now
        _ALIVE['closing_at'] = 0.0  # tab is back (refresh / reopened)


def _watchdog(server) -> None:
    while not _QUIT.wait(1.0):
        now = time.time()
        if _ALIVE['closing_at'] and now - _ALIVE['closing_at'] > CLOSE_GRACE_SECS:
            print('\nBrowser tab closed - stopped.')
            break
        if _ALIVE['last_ping'] and now - _ALIVE['last_ping'] > IDLE_GRACE_SECS:
            print('\nNo browser tab found - stopped.')
            break
    else:
        print('\nQuit from the browser - stopped.')
    server.shutdown()


def list_background_images() -> list[str]:
    try:
        files = os.listdir(IMAGES_DIR)
    except FileNotFoundError:
        return []
    out = [f for f in files
           if not f.startswith(',') and not f.startswith('.')
           and f.lower().endswith(ALLOWED_IMAGE_EXTS)]
    return sorted(out, key=str.lower)


# Only one native folder dialog at a time.
_DIALOG_LOCK = threading.Lock()


def pick_folder_dialog() -> str | None:
    """Open a native 'choose folder' dialog.

    Returns the chosen path, '' if the user cancelled, or None if no
    dialog could be shown at all.
    """
    if sys.platform == 'darwin':
        script = ('tell application "System Events" to activate\n'
                  'POSIX path of (choose folder with prompt '
                  '"Choose where to save the PowerPoints")')
        try:
            out = subprocess.run(['osascript', '-e', script],
                                 capture_output=True, text=True, timeout=600)
            if out.returncode == 0:
                return out.stdout.strip()
            if 'User canceled' in (out.stderr or '') or out.returncode == 1:
                return ''
        except Exception:
            pass

    # Fallback (and Windows/Linux): a tkinter dialog in its own process so it
    # never interferes with the web server.
    code = (
        "import tkinter as tk\n"
        "from tkinter import filedialog\n"
        "root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)\n"
        "print(filedialog.askdirectory("
        "title='Choose where to save the PowerPoints') or '')\n")
    try:
        out = subprocess.run([sys.executable, '-c', code],
                             capture_output=True, text=True, timeout=600)
        if out.returncode == 0:
            return out.stdout.strip()
    except Exception:
        pass
    return None


def fetch_song_for_request(data: dict) -> dict | None:
    """Fetch lyrics from a request payload (artist/title and/or one URL)."""
    artist = (data.get('artist') or '').strip()
    title = (data.get('title') or '').strip()
    url = (data.get('url') or '').strip()

    # fetch_song() routes any pasted link to the right scraper by domain.
    return core.fetch_song(artist, title, azlyrics_url='', genius_url=url)


class Handler(BaseHTTPRequestHandler):
    protocol_version = 'HTTP/1.1'  # keep-alive: loads thumbnails reliably/faster

    # ---------- helpers ----------
    def _send_json(self, payload: dict, status: int = 200) -> None:
        body = json.dumps(payload).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_file(self, path: str, content_type: str) -> None:
        try:
            with open(path, 'rb') as f:
                body = f.read()
        except OSError:
            self.send_error(404)
            return
        self.send_response(200)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(body)))
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

    def log_message(self, fmt, *args):  # keep the terminal quiet
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

        if path == '/api/settings':
            settings = core.load_app_settings()
            self._send_json({
                'ok': True,
                'output_folder': settings.get('output_folder', CUR_DIR) or CUR_DIR,
                'storage_mode': core.clean_storage_mode(settings.get('storage_mode')),
                'previous_powerpoints_action': (
                    core.clean_previous_powerpoints_action(
                        settings.get('previous_powerpoints_action'))),
            })
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

        if path == '/api/ping':
            _note_ping(closing='closing=1' in (urlparse(self.path).query or ''))
            self._send_json({'ok': True})
        elif path == '/api/quit':
            self._send_json({'ok': True})
            _QUIT.set()
        elif path == '/api/pick_folder':
            if not _DIALOG_LOCK.acquire(blocking=False):
                self._send_json({'ok': False,
                                 'error': 'A folder dialog is already open.'})
                return
            try:
                chosen = pick_folder_dialog()
            finally:
                _DIALOG_LOCK.release()
            if chosen is None:
                self._send_json({'ok': False,
                                 'error': "Couldn't open a folder dialog - "
                                          'type the path instead.'})
            else:
                self._send_json({'ok': True, 'path': chosen,
                                 'cancelled': chosen == ''})
        elif path == '/api/preview':
            self.handle_preview(data)
        elif path == '/api/spotify_playlist':
            self.handle_spotify_playlist(data)
        elif path == '/api/create_one':
            self.handle_create_one(data)
        elif path == '/api/save_settings':
            self.handle_save_settings(data)
        elif path == '/api/upload_background':
            self.handle_upload_background(data)
        else:
            self.send_error(404)

    def handle_preview(self, data: dict) -> None:
        """Fetch lyrics so the user can check/edit them before creating slides."""
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
        """Create the .pptx for one song."""
        output_folder = (data.get('output_folder') or '').strip()
        output_folder = os.path.expanduser(output_folder) or CUR_DIR
        storage_mode = core.clean_storage_mode(data.get('storage_mode'))
        previous_action = core.clean_previous_powerpoints_action(
            data.get('previous_powerpoints_action'))
        try:
            os.makedirs(output_folder, exist_ok=True)
            core.prepare_selected_output_folder(output_folder, previous_action)
        except OSError as exc:
            self._send_json({'ok': False,
                             'error': f"Can't use that output folder ({exc})."})
            return

        library_path = (data.get('library_path') or '').strip()
        if library_path:
            try:
                existing = core.place_existing_song_in_output(
                    library_path, output_folder, storage_mode)
            except Exception as exc:
                self._send_json({'ok': False,
                                 'error': f"Couldn't place that library song ({exc})."})
                return
            self._send_json({
                'ok': True,
                'path': existing['path'],
                'file': existing['file'],
                'title': existing['title'],
                'artist': existing['artist'],
                'source': existing['source'],
                'delivery': existing.get('delivery', ''),
            })
            return

        manual_lyrics = data.get('lyrics') or ''
        if manual_lyrics.strip():
            # Lyrics were previewed/edited (or pasted) in the UI - use them as-is.
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
            if not data.get('skip_library'):
                existing = core.resolve_existing_song_from_all_songs(
                    artist=(data.get('artist') or '').strip(),
                    title=(data.get('title') or '').strip(),
                    url=(data.get('url') or '').strip(),
                    output_base=output_folder,
                    storage_mode=storage_mode)
                if existing.get('status') == 'placed':
                    result = existing['result']
                    self._send_json({
                        'ok': True,
                        'path': result['path'],
                        'file': result['file'],
                        'title': result['title'],
                        'artist': result['artist'],
                        'source': result['source'],
                        'delivery': result.get('delivery', ''),
                    })
                    return
                if existing.get('status') == 'ambiguous':
                    self._send_json({
                        'ok': False,
                        'needs_choice': True,
                        'choices': existing.get('choices', []),
                        'error': 'Multiple saved versions match this song.',
                    })
                    return

            song = fetch_song_for_request(data)
            if song is None:
                self._send_json({
                    'ok': False,
                    'error': core.LAST_ERROR or core.LYRICS_NOT_FOUND_MSG,
                    'needs_lyrics': True,  # UI offers the manual-paste editor
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

        # "Random" pulls from the images ticked in the UI (saved settings).
        saved_selected = settings.get('background_selected', {}) or {}
        selected_backgrounds = [
            {'image': n, 'color_mode': saved_colors.get(n, 'Auto')}
            for n in all_images if saved_selected.get(n, True)
        ]

        try:
            with _BUILD_LOCK:
                out_path = core.build_song_presentation(
                    song, output_folder, background=background,
                    selected_backgrounds=selected_backgrounds,
                    storage_mode=storage_mode)
        except Exception as exc:
            self._send_json({'ok': False,
                             'error': f'Failed to build the PowerPoint ({exc}).'})
            return

        self._send_json({
            'ok': True,
            'path': out_path,
            'file': os.path.basename(out_path),
            'title': song['title'],
            'artist': song['artist'],
            'source': song['source'],
            'delivery': 'linked' if storage_mode == core.STORAGE_ALL_SONGS_ALIAS else
                        ('copied' if storage_mode == core.STORAGE_BOTH_COPIES else 'output_only'),
        })

    def handle_save_settings(self, data: dict) -> None:
        settings = core.load_app_settings()
        if isinstance(data.get('output_folder'), str) and data['output_folder'].strip():
            settings['output_folder'] = data['output_folder'].strip()
        if isinstance(data.get('storage_mode'), str):
            settings['storage_mode'] = core.clean_storage_mode(data.get('storage_mode'))
        if isinstance(data.get('previous_powerpoints_action'), str):
            settings['previous_powerpoints_action'] = (
                core.clean_previous_powerpoints_action(
                    data.get('previous_powerpoints_action')))
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
        """Save a new background image sent as base64 from the browser."""
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
    port = PORT
    server = None
    for attempt in range(10):
        try:
            server = ThreadingHTTPServer(('127.0.0.1', port), Handler)
            break
        except OSError:
            port += 1
    if server is None:
        raise SystemExit('No free port found between 8765 and 8774.')

    url = f'http://127.0.0.1:{port}'
    print(f'Lyrics Slides is running at {url}')
    print('It stops by itself when you close the tab (or use the Quit button).')
    threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    threading.Thread(target=_watchdog, args=(server,), daemon=True).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\nStopped.')


if __name__ == '__main__':
    main()
