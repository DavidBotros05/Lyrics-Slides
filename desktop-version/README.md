# Lyrics Slides Desktop Version

This is the local desktop version of Lyrics Slides. It runs a small web server on
your computer, opens the app in your browser, and saves generated PowerPoint
files directly to a folder you choose.

## Setup

From this folder:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

If you use a Genius API token, copy `.env.example` to `.env` and put your token
there. The app also works without a token by using the other lyric sources and
manual pasted lyrics.

## Run

```bash
python app.py
```

The app opens at `http://127.0.0.1:8765`.

## Files

- `app.py`: local browser-based app.
- `ui.html`: local app interface.
- `Create_Lyrics.py`: lyric fetching and slide creation.
- `Background_for_Lyrics.py`: background fitting and text color helpers.
- `background_images/`: bundled slide backgrounds.
- `To Add Background/` and `Created_with_background/`: optional folders for the
  older batch background workflow.

Generated PowerPoints, local settings, `.env`, and virtual environments are
ignored by git.
