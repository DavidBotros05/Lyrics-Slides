# Put Lyrics Slides on the internet (free) — use it from your iPhone/iPad

After this one-time setup (~15 min on your Mac), the app lives at a web
address like `https://lyrics-slides.onrender.com`. Open it in Safari at
church, create the PowerPoint, and it downloads straight to your phone's
Files app. Your Mac can be off.

## Step 1 — Put this repo on GitHub

1. Go to https://github.com and sign up (free) if you don't have an account.
2. Click **+** (top right) → **New repository**. Name it `lyrics-slides`,
   set it to **Private**, click **Create repository**.
3. Push this repository to GitHub. The Render cloud service lives at the repo
   root, and the local desktop app lives in `desktop-version/`.

> Never upload your real `.env` file — the Genius token goes into Render in
> Step 2 instead. (The `.gitignore` here excludes it anyway.)

## Step 2 — Deploy on Render (free)

1. Go to https://render.com and sign up — choose **Sign up with GitHub**.
2. Click **New +** → **Web Service**, and pick your `lyrics-slides` repo.
3. Render reads `render.yaml` and fills everything in. If it asks:
   - Build command: `pip install -r requirements.txt`
   - Start command: `python app.py`
   - Instance type: **Free**
4. Under **Environment Variables**, add:
   - Key: `GENIUS_ACCESS_TOKEN`
   - Value: your token (it's in the `.env` file in the Creation folder
     on your Mac)
5. Click **Create Web Service** and wait a few minutes for the first
   deploy to finish.
6. Your app is now at the URL shown at the top, e.g.
   `https://lyrics-slides.onrender.com`.

## Step 3 — On your iPhone/iPad

1. Open that URL in Safari.
2. Tap **Share** → **Add to Home Screen** — now it looks and opens like
   an app.
3. Create slides as usual; tap the green **Download** button for each
   song. Files land in **Files app → Downloads**, and from there you can
   AirDrop, email, or USB them to the church computer.

## Things to know

- **First load can take ~1 minute.** The free plan puts the app to sleep
  after 15 minutes of no use; it wakes up on the next visit. Open it
  while walking into church and it'll be ready.
- **Background images:** ones you upload through the website disappear
  whenever the app restarts. To add a background permanently, add the
  image file to the `background_images` folder in your GitHub repo
  (GitHub → repo → `background_images` → Add file → Upload). Render
  redeploys automatically.
- **Lyrics fetching:** AZLyrics/Genius sometimes block cloud servers. If
  a fetch fails, use **Preview / edit lyrics** and paste the lyrics in —
  slide creation itself always works.
- **The URL is public.** Anyone who has the exact link can use the app,
  so don't post it anywhere public.
- The desktop app is included in `desktop-version/` and can be downloaded from
  the same GitHub repo.
