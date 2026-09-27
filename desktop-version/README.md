# Lyrics Slides Desktop

Create lyric PowerPoints locally, save them to a folder you choose, and reuse presentations from your song library. The app runs on your computer and opens its interface in your browser.

## Setup

Requires Python 3.10 or newer. From this folder:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

On Windows, use `.venv\Scripts\activate` to activate the environment.

The browser opens automatically, normally at [localhost:8765](http://127.0.0.1:8765). If that port is occupied, the app chooses another available port.

## Workflow

1. Choose the output folder for your presentations.
2. Add songs individually or import a public Spotify or Apple Music playlist.
3. Preview and edit lyrics, then choose backgrounds and text colors.
4. Create the PowerPoints. Existing matching presentations can be reused, with a choice offered when multiple versions match.

Storage options let you save to the output folder, keep copies in the song library, or link to library presentations. You can also update backgrounds on existing PowerPoints and choose how to handle presentations left in the output folder from a previous session.

## Configuration

Optional `GENIUS_ACCESS_TOKEN` and `APPLE_MUSIC_DEVELOPER_TOKEN` values can be supplied as environment variables or in a `.env` file in this folder. Apple Music uses public-page track information when no developer token is available.

The song-library location is configured by `ALL_SONGS_DIR` in `Create_Lyrics.py`. Set it to the folder you use for reusable presentations before using library storage. Choose output-only storage to save presentations without using the library.

Bundled backgrounds live in `background_images/`. Background choices and folder preferences are remembered between sessions.

## Tests

```bash
python -m unittest discover -v
```
