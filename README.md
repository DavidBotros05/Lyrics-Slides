# Lyrics Slides

Lyrics Slides creates PowerPoint lyric slides from artist/title searches, pasted
lyrics, Genius/AZLyrics links, or Spotify playlist imports.

This repository contains both versions of the project:

- Root folder: the cloud web service deployed on Render.
- `desktop-version/`: the local desktop version for running on a Mac or PC.

## Cloud Version

Render deploys the root of this repo. The important cloud files are:

- `app.py`
- `ui.html`
- `Create_Lyrics.py`
- `Background_for_Lyrics.py`
- `background_images/`
- `requirements.txt`
- `render.yaml`

Render uses:

```bash
pip install -r requirements.txt
python app.py
```

Set `GENIUS_ACCESS_TOKEN` in Render as an environment variable. Do not commit a
real `.env` file.

## Desktop Version

The desktop app lives in `desktop-version/`. It runs locally, opens the browser
automatically, and saves generated PowerPoints to a folder selected on the
computer.

See `desktop-version/README.md` for local setup instructions.

## Background Images

Permanent background images should be committed to the relevant
`background_images/` folder:

- `background_images/` for the cloud app.
- `desktop-version/background_images/` for the desktop app.

Backgrounds uploaded through the cloud website may disappear when the Render
service restarts because Render's free filesystem is ephemeral.
