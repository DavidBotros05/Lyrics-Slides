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
ALLOWED_PPTX_EXTS = ('.pptx',)

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


def short_error(exc: Exception, limit: int = 180) -> str:
    msg = ' '.join(str(exc).split())
    return (msg[:limit] + '...') if len(msg) > limit else (msg or type(exc).__name__)


# Only one native folder dialog at a time.
_DIALOG_LOCK = threading.Lock()


def _applescript_string(value: str) -> str:
    return '"' + value.replace('\\', '\\\\').replace('"', '\\"') + '"'


def pick_folder_dialog(prompt: str = 'Choose where to save the PowerPoints') -> str | None:
    """Open a native 'choose folder' dialog.

    Returns the chosen path, '' if the user cancelled, or None if no
    dialog could be shown at all.
    """
    if sys.platform == 'darwin':
        script = ('tell application "System Events" to activate\n'
                  'POSIX path of (choose folder with prompt '
                  f'{_applescript_string(prompt)})')
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
        f"title={prompt!r}) or '')\n")
    try:
        out = subprocess.run([sys.executable, '-c', code],
                             capture_output=True, text=True, timeout=600)
        if out.returncode == 0:
            return out.stdout.strip()
    except Exception:
        pass
    return None


def pick_powerpoint_dialog() -> str | None:
    """Open a native 'choose .pptx file' dialog."""
    prompt = 'Choose a PowerPoint to change'
    if sys.platform == 'darwin':
        script = ('tell application "System Events" to activate\n'
                  'POSIX path of (choose file with prompt '
                  f'{_applescript_string(prompt)})')
        try:
            out = subprocess.run(['osascript', '-e', script],
                                 capture_output=True, text=True, timeout=600)
            if out.returncode == 0:
                return out.stdout.strip()
            if 'User canceled' in (out.stderr or '') or out.returncode == 1:
                return ''
        except Exception:
            pass

    code = (
        "import tkinter as tk\n"
        "from tkinter import filedialog\n"
        "root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)\n"
        "print(filedialog.askopenfilename("
        "title='Choose a PowerPoint to change', "
        "filetypes=[('PowerPoint', '*.pptx')]) or '')\n")
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
            by_folder = settings.get('previous_powerpoints_actions_by_folder')
            if not isinstance(by_folder, dict):
                by_folder = {}
            clean_by_folder = {
                str(folder): core.clean_previous_powerpoints_action(action)
                for folder, action in by_folder.items()
                if isinstance(folder, str)
            }
            self._send_json({
                'ok': True,
                'output_folder': settings.get('output_folder', CUR_DIR) or CUR_DIR,
                'storage_mode': core.clean_storage_mode(settings.get('storage_mode')),
                'last_spotify_playlist': (
                    settings.get('last_spotify_playlist', '')
                    if isinstance(settings.get('last_spotify_playlist'), str)
                    else ''),
                'last_apple_music_playlist': (
                    settings.get('last_apple_music_playlist', '')
                    if isinstance(settings.get('last_apple_music_playlist'), str)
                    else ''),
                'last_playlist': (
                    settings.get('last_playlist')
                    or settings.get('last_apple_music_playlist')
                    or settings.get('last_spotify_playlist')
                    or ''),
                'previous_powerpoints_action': (
                    core.clean_previous_powerpoints_action(
                        settings.get('previous_powerpoints_action'))),
                'previous_powerpoints_actions_by_folder': clean_by_folder,
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
                prompt = (data.get('prompt') or 'Choose where to save the PowerPoints')
                chosen = pick_folder_dialog(str(prompt))
            finally:
                _DIALOG_LOCK.release()
            if chosen is None:
                self._send_json({'ok': False,
                                 'error': "Couldn't open a folder dialog - "
                                          'type the path instead.'})
            else:
                self._send_json({'ok': True, 'path': chosen,
                                 'cancelled': chosen == ''})
        elif path == '/api/pick_powerpoint':
            if not _DIALOG_LOCK.acquire(blocking=False):
                self._send_json({'ok': False,
                                 'error': 'A file dialog is already open.'})
                return
            try:
                chosen = pick_powerpoint_dialog()
            finally:
                _DIALOG_LOCK.release()
            if chosen is None:
                self._send_json({'ok': False,
                                 'error': "Couldn't open a file dialog - "
                                          'type the path instead.'})
            else:
                self._send_json({'ok': True, 'path': chosen,
                                 'cancelled': chosen == ''})
        elif path == '/api/preview':
            self.handle_preview(data)
        elif path == '/api/playlist':
            self.handle_playlist(data)
        elif path == '/api/spotify_playlist':
            self.handle_spotify_playlist(data)
        elif path == '/api/apple_music_playlist':
            self.handle_apple_music_playlist(data)
        elif path == '/api/create_one':
            self.handle_create_one(data)
        elif path == '/api/prepare_output_folder':
            self.handle_prepare_output_folder(data)
        elif path == '/api/change_backgrounds':
            self.handle_change_backgrounds(data)
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
        settings = core.load_app_settings()
        settings['last_spotify_playlist'] = link[:2048]
        core.save_app_settings(settings)
        self._send_json({
            'ok': True,
            'name': result['name'],
            'tracks': result['tracks'],
            'total': result['total'],
            'note': result.get('note', ''),
        })

    def handle_playlist(self, data: dict) -> None:
        """Detect the playlist service, save the link, and return its songs."""
        link = (data.get('link') or '').strip()
        result = core.fetch_playlist(link)
        if result is None:
            self._send_json({'ok': False,
                             'error': core.LAST_ERROR or
                                      "Couldn't read that playlist."})
            return
        settings = core.load_app_settings()
        settings['last_playlist'] = link[:2048]
        core.save_app_settings(settings)
        self._send_json({
            'ok': True,
            'name': result['name'],
            'tracks': result['tracks'],
            'total': result['total'],
            'note': result.get('note', ''),
        })

    def handle_apple_music_playlist(self, data: dict) -> None:
        """Return every song (title + artist) in an Apple Music playlist."""
        link = (data.get('link') or '').strip()
        result = core.fetch_apple_music_playlist(link)
        if result is None:
            self._send_json({'ok': False,
                             'error': core.LAST_ERROR or
                                      "Couldn't read that Apple Music playlist."})
            return
        settings = core.load_app_settings()
        settings['last_apple_music_playlist'] = link[:2048]
        core.save_app_settings(settings)
        self._send_json({
            'ok': True,
            'name': result['name'],
            'tracks': result['tracks'],
            'total': result['total'],
            'note': result.get('note', ''),
        })

    def _background_options_for_request(self, data: dict) -> tuple[dict | None, list[dict]]:
        bg_name = (data.get('background') or '').strip()
        all_images = list_background_images()
        settings = core.load_app_settings()
        saved_colors = settings.get('background_colors', {}) or {}

        background = None
        if bg_name and bg_name != 'random' and bg_name in all_images:
            background = {
                'image': bg_name,
                'color_mode': saved_colors.get(bg_name, 'Auto'),
            }

        saved_selected = settings.get('background_selected', {}) or {}
        selected_backgrounds = [
            {'image': n, 'color_mode': saved_colors.get(n, 'Auto')}
            for n in all_images if saved_selected.get(n, True)
        ]
        return background, selected_backgrounds

    def _pptx_sources_for_path(self, source_path: str,
                               include_subfolders: bool = True) -> tuple[Path | None, list[Path], str | None]:
        source = Path(os.path.expanduser(source_path)).resolve(strict=False)
        if source.is_file():
            if source.suffix.lower() not in ALLOWED_PPTX_EXTS:
                return None, [], 'Choose a .pptx file.'
            return source.parent, [source], None
        if source.is_dir():
            iterator = source.rglob('*') if include_subfolders else source.iterdir()
            files = sorted(
                p for p in iterator
                if p.is_file()
                and p.suffix.lower() in ALLOWED_PPTX_EXTS
                and not p.name.startswith('~$')
            )
            if not files:
                return source, [], 'No .pptx files were found in that folder.'
            return source, files, None
        return None, [], "That PowerPoint path or folder doesn't exist."

    def handle_change_backgrounds(self, data: dict) -> None:
        """Change backgrounds for one .pptx file or every .pptx in a folder."""
        source_items = []
        raw_items = data.get('source_items')
        if isinstance(raw_items, list):
            for item in raw_items:
                if not isinstance(item, dict):
                    continue
                path = item.get('path')
                if not isinstance(path, str) or not path.strip():
                    continue
                source_items.append({
                    'path': path.strip(),
                    'include_subfolders': bool(item.get('include_subfolders', True)),
                })
        else:
            raw_sources = data.get('source_paths')
            if not isinstance(raw_sources, list):
                raw_sources = [data.get('source_path')]
            include_subfolders = bool(data.get('include_subfolders', True))
            for path in raw_sources:
                if isinstance(path, str) and path.strip():
                    source_items.append({
                        'path': path.strip(),
                        'include_subfolders': include_subfolders,
                    })

        if not source_items:
            self._send_json({'ok': False,
                             'error': 'Choose at least one PowerPoint or source folder first.'})
            return

        source_groups = []
        source_errors = []
        seen_sources = set()
        for item in source_items:
            source_path = item['path']
            source_probe = Path(os.path.expanduser(source_path)).resolve(strict=False)
            source_is_dir = source_probe.is_dir()
            source_root, sources, error = self._pptx_sources_for_path(
                source_path, include_subfolders=item['include_subfolders'])
            if error:
                source_errors.append({'source': source_path, 'error': error})
                continue
            unique_sources = []
            for src in sources:
                key = str(src.resolve(strict=False))
                if key in seen_sources:
                    continue
                seen_sources.add(key)
                unique_sources.append(src)
            if unique_sources:
                source_groups.append({
                    'path': source_path,
                    'root': source_root,
                    'is_dir': source_is_dir,
                    'sources': unique_sources,
                })

        if not source_groups:
            error = source_errors[0]['error'] if source_errors else 'No .pptx files were found.'
            self._send_json({'ok': False, 'error': error, 'failed': source_errors})
            return

        output_mode = (data.get('output_mode') or 'overwrite').strip()
        if output_mode not in ('overwrite', 'copy_to_folder'):
            output_mode = 'overwrite'

        output_folder_raw = (data.get('output_folder') or '').strip()
        output_folder = Path(os.path.expanduser(output_folder_raw))
        if output_mode == 'copy_to_folder':
            if not output_folder_raw:
                self._send_json({'ok': False,
                                 'error': 'Choose where to save the changed PowerPoints.'})
                return
            try:
                output_folder.mkdir(parents=True, exist_ok=True)
            except OSError as exc:
                self._send_json({'ok': False,
                                 'error': f"Can't use that output folder ({exc})."})
                return

        background, selected_backgrounds = self._background_options_for_request(data)
        delete_original = bool(data.get('delete_original')) and output_mode == 'copy_to_folder'

        changed = []
        failed = list(source_errors)
        multiple_source_groups = len(source_groups) > 1
        for group in source_groups:
            source_root = group['root']
            source_is_dir = group['is_dir']
            for src in group['sources']:
                if output_mode == 'overwrite':
                    dest = src
                elif source_is_dir and source_root:
                    rel = src.relative_to(source_root)
                    dest = output_folder / source_root.name / rel if multiple_source_groups else output_folder / rel
                else:
                    dest = output_folder / src.name

                try:
                    with _BUILD_LOCK:
                        core.change_powerpoint_background(
                            str(src), str(dest),
                            background=background,
                            selected_backgrounds=selected_backgrounds)
                    if delete_original and src.resolve(strict=False) != dest.resolve(strict=False):
                        src.unlink()
                    changed.append({'source': str(src), 'path': str(dest), 'file': dest.name})
                except Exception as exc:
                    failed.append({'source': str(src), 'error': short_error(exc)})

        self._send_json({
            'ok': not failed,
            'changed': changed,
            'failed': failed,
            'changed_count': len(changed),
            'failed_count': len(failed),
            'deleted_originals': delete_original,
        }, status=200 if not failed else 500)

    def handle_create_one(self, data: dict) -> None:
        """Create the .pptx for one song."""
        output_folder = (data.get('output_folder') or '').strip()
        output_folder = os.path.expanduser(output_folder) or CUR_DIR
        storage_mode = core.clean_storage_mode(data.get('storage_mode'))
        previous_action = core.clean_previous_powerpoints_action(
            data.get('previous_powerpoints_action'))
        try:
            os.makedirs(output_folder, exist_ok=True)
        except OSError as exc:
            self._send_json({'ok': False,
                             'error': f"Can't use that output folder ({exc})."})
            return

        library_path = (data.get('library_path') or '').strip()
        protected_paths = core.existing_output_paths_for_songs(
            output_folder,
            data.get('preserve_songs') if isinstance(data.get('preserve_songs'), list)
            else [],
        )
        if library_path:
            protected_paths.append(library_path)
        elif not data.get('skip_output'):
            existing = core.resolve_existing_song_from_output(
                artist=(data.get('artist') or '').strip(),
                title=(data.get('title') or '').strip(),
                url=(data.get('url') or '').strip(),
                output_base=output_folder)
            if existing.get('status') == 'ambiguous':
                self._send_json({
                    'ok': False,
                    'needs_choice': True,
                    'choice_source': 'output',
                    'choices': existing.get('choices', []),
                    'error': 'Multiple versions already exist in the chosen folder.',
                })
                return
            if existing.get('status') == 'placed':
                result = existing['result']
                protected_paths.append(result['path'])
                try:
                    core.prepare_selected_output_folder(
                        output_folder, previous_action, protected_paths)
                except OSError as exc:
                    self._send_json({'ok': False,
                                     'error': f"Can't use that output folder ({exc})."})
                    return
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

        try:
            core.prepare_selected_output_folder(
                output_folder, previous_action, protected_paths)
        except OSError as exc:
            self._send_json({'ok': False,
                             'error': f"Can't use that output folder ({exc})."})
            return

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
                        'choice_source': 'all_songs',
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

    def handle_prepare_output_folder(self, data: dict) -> None:
        """Immediately apply the selected previous-PowerPoint action."""
        output_folder = os.path.expanduser(
            (data.get('output_folder') or '').strip())
        if not output_folder:
            self._send_json({'ok': False,
                             'error': 'Choose an output folder first.'})
            return
        if not os.path.isdir(output_folder):
            self._send_json({'ok': False,
                             'error': "That output folder doesn't exist."})
            return

        action = core.clean_previous_powerpoints_action(
            data.get('previous_powerpoints_action'))
        if action == core.PREVIOUS_POWERPOINTS_KEEP:
            self._send_json({'ok': True, 'changed_count': 0,
                             'action': action})
            return

        try:
            with _BUILD_LOCK:
                changed = core.prepare_selected_output_folder(
                    output_folder, action, preserve_paths=[], force=True)
        except Exception as exc:
            self._send_json({
                'ok': False,
                'error': f"Couldn't prepare the output folder ({short_error(exc)}).",
            })
            return

        self._send_json({'ok': True, 'changed_count': changed,
                         'action': action})

    def handle_save_settings(self, data: dict) -> None:
        settings = core.load_app_settings()
        if isinstance(data.get('output_folder'), str) and data['output_folder'].strip():
            settings['output_folder'] = data['output_folder'].strip()
        if isinstance(data.get('storage_mode'), str):
            settings['storage_mode'] = core.clean_storage_mode(data.get('storage_mode'))
        if (isinstance(data.get('last_spotify_playlist'), str) and
                data['last_spotify_playlist'].strip()):
            settings['last_spotify_playlist'] = (
                data['last_spotify_playlist'].strip()[:2048])
        if (isinstance(data.get('last_apple_music_playlist'), str) and
                data['last_apple_music_playlist'].strip()):
            settings['last_apple_music_playlist'] = (
                data['last_apple_music_playlist'].strip()[:2048])
        if (isinstance(data.get('last_playlist'), str) and
                data['last_playlist'].strip()):
            settings['last_playlist'] = data['last_playlist'].strip()[:2048]
        if isinstance(data.get('previous_powerpoints_action'), str):
            settings['previous_powerpoints_action'] = (
                core.clean_previous_powerpoints_action(
                    data.get('previous_powerpoints_action')))
        if isinstance(data.get('previous_powerpoints_actions_by_folder'), dict):
            clean_by_folder = {}
            for folder, action in data['previous_powerpoints_actions_by_folder'].items():
                if not isinstance(folder, str) or not folder.strip():
                    continue
                clean_by_folder[folder.strip()] = (
                    core.clean_previous_powerpoints_action(action))
            settings['previous_powerpoints_actions_by_folder'] = clean_by_folder
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
