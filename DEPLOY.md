# Web deployment

The repository root contains the web version of Lyrics Slides. It serves the browser interface and provides PowerPoint and ZIP downloads.

## Render

Create a Python web service connected to this repository, using the root directory. The included `render.yaml` defines the service configuration:

- Build command: `pip install -r requirements.txt`
- Start command: `python app.py`
- Port: supplied by the host through `PORT`

Configure `GENIUS_ACCESS_TOKEN` for Genius API access and optionally `APPLE_MUSIC_DEVELOPER_TOKEN` for complete Apple Music playlist imports. After deployment, open the service URL shown in the hosting dashboard.

## Backgrounds and downloads

Add bundled backgrounds to `background_images/` to include them in deployments. Runtime uploads and saved preferences require persistent storage to survive replacement of the service filesystem.

Generated presentations are temporary downloads. Save the PowerPoint or ZIP after creating it; download links are not permanent and are lost when the service restarts.

## Mobile use

Open the service URL in a mobile browser, create slides, and save the downloads to the device's file manager. On iOS, the site can also be added to the home screen from Safari's Share menu.

External lyric providers can restrict requests from hosting services. The lyric editor accepts pasted text when automatic fetching is unavailable.
