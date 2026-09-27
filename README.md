# Lyrics Slides

Turn songs into PowerPoint presentations for worship services, rehearsals, and group singing. Lyrics Slides brings song selection, lyric editing, and slide backgrounds into one browser-based workspace.

[Open Lyrics Slides](https://lyrics-slides.onrender.com)

## Features

- Find lyrics by song title and artist, or use a Genius or AZLyrics link.
- Import song lists from public Spotify and Apple Music playlists.
- Paste or edit lyrics before creating slides.
- Choose background images, adjust their positioning, and select readable text colors.
- Generate PowerPoint presentations with lyrics split across slides.
- Update backgrounds on existing PowerPoint presentations.

## Web and desktop versions

| | Web | Desktop |
| --- | --- | --- |
| Interface | Browser on a computer, tablet, or phone | Browser connected to a local Python app |
| Finished slides | Download individual PowerPoints or a ZIP | Save directly to a chosen folder |
| Existing presentations | Upload files to change backgrounds | Browse local files and reuse saved songs |
| Location | Repository root | [`desktop-version/`](desktop-version/) |

The desktop version also supports a reusable song library, matching existing presentations by title and artist, and organizing previous presentations in the selected output folder.

## Quick start

Requires Python 3.10 or newer.

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

On Windows, activate the environment with `.venv\Scripts\activate` instead.

Open [localhost:8765](http://localhost:8765). To use local folder selection and song-library features, follow the [desktop setup](desktop-version/README.md).

## Create a presentation

1. Add songs by title and artist, paste a lyrics link, or import a playlist.
2. Preview the lyrics and make any edits.
3. Choose a background or select a group for random backgrounds.
4. Create the slides, then download them or save them to your desktop output folder.

## Configuration

`GENIUS_ACCESS_TOKEN` enables Genius API access. `APPLE_MUSIC_DEVELOPER_TOKEN` enables Apple Music API imports, including pagination for longer playlists. Both can be supplied as environment variables or in a local `.env` file.

Without an Apple Music token, imports use the track information available on the public playlist page. Availability of automatic lyrics and playlist imports depends on the source service; pasted lyrics can be used when a source is unavailable.

The web app uses `PORT` when supplied and defaults to `8765`. See [deployment](DEPLOY.md) for hosting details.

## Development

```bash
python -m unittest discover -v
cd desktop-version
python -m unittest discover -v
```

The apps use Python's HTTP server, python-pptx for presentations, and Pillow for background processing.
