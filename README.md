# Sheet Cards (React + Google Sheets)

A React flashcard app that:

- Loads cards from `Card Data!A:Z`
- Loads stats from `Card Progress!A:Z`
- Uses adaptive multiple-choice options based on mastery
- Tracks stats in memory as you answer
- Writes stat updates back using `spreadsheets.values.batchUpdate`

## 1) Google Cloud setup

1. Create a Google Cloud project.
2. Enable the Google Sheets API.
3. Configure OAuth consent screen.
4. Create an OAuth 2.0 Client ID for a web app.
5. Add your dev origin (for example `http://localhost:5173`) to Authorized JavaScript origins.

Required scope:

- `https://www.googleapis.com/auth/drive.file`

## 2) Sheet format

Use the in-app `Initialize Sheet Template` button. It creates two tabs:

- `Card Data` headers:
  - `question,answer,pronunciation,tags,question_explanation,answer_explanation`
- `Card Progress` headers:
  - `question,answer,pronunciation,times_seen,times_correct,times_wrong,streak,last_seen_at,last_result,mastery`

Optional columns in `Card Data`:
- `pronunciation` (optional)
- `tags` (optional)
- `question_explanation` (optional)
- `answer_explanation` (optional)


## 3) Run locally

App owner setup (one-time):

```bash
cp .env.example .env.local
```

Set `VITE_GOOGLE_CLIENT_ID` in `.env.local`.
Set `VITE_GOOGLE_API_KEY` in `.env.local` (used by Google Picker).
Set `VITE_GOOGLE_CLOUD_PROJECT_NUMBER` in `.env.local` (required for Picker with `drive.file`).

Then run:

```bash
npm install
npm run dev
```

Open the printed local URL (usually `http://localhost:5173`).

If your environment does not detect file changes reliably (common on some Windows/network drive setups), use:

```bash
npm run dev:poll
```

User flow in the app:

1. `Home`: click `Connect Google`.
2. `Sheet`: click `Pick Sheet From Drive` (recommended) or paste a sheet URL, then `Load Cards`.
3. Click `Start Study Round`.
4. `Study`: answer cards.
5. `Round Stats` and `Stats`: review performance, sync if needed, then start next round.

Recent/last sheet URLs are remembered in local browser storage.
The OAuth client ID is app-owned config and is not entered by end users.
The app includes `/privacy.html` and `/terms.html` pages for policy links.

## 4) Production setup (Namecheap + Google)

### 4.1 Prepare production env

Copy and fill a production env file:

```bash
cp .env.production.example .env.production
```

Set:

- `VITE_GOOGLE_CLIENT_ID`
- `VITE_GOOGLE_API_KEY`
- `VITE_GOOGLE_CLOUD_PROJECT_NUMBER`

`.env.production` is ignored by git.

### 4.2 Build for production

```bash
npm install
npm run build:prod
```

This creates `dist/`.

### 4.3 Upload to Namecheap hosting

1. Open Namecheap cPanel `File Manager`.
2. Go to `public_html` (or your addon-domain root folder).
3. Upload the **contents** of `dist/` (not the `dist` folder itself).
4. Verify `index.html` is at the web root.
5. Enable SSL so the site is served over `https://`.

Notes:

- `public/.htaccess` is included in production build to support SPA routing and cache headers on Apache hosting.
- `privacy.html` and `terms.html` are deployed as static pages.

### 4.4 Configure Google for the live domain

In Google Cloud:

1. OAuth Client (Web App) -> add Authorized JavaScript origins:
   - `https://yourdomain.com`
   - `https://www.yourdomain.com` (if used)
   - keep `http://localhost:5173` for local dev
2. API Key restrictions (HTTP referrers):
   - `https://yourdomain.com/*`
   - `https://www.yourdomain.com/*` (if used)
3. OAuth consent screen:
   - add your authorized domain
   - set app homepage/privacy/terms URLs
   - publish app when ready

### 4.5 Deploy on GitHub Pages (alternative)

This repo includes a workflow at `.github/workflows/deploy-github-pages.yml`.

One-time setup in GitHub repository settings:

1. `Settings -> Pages`
2. `Source`: set to `GitHub Actions`
3. `Settings -> Secrets and variables -> Actions -> New repository secret`
4. Add:
   - `VITE_GOOGLE_CLIENT_ID`
   - `VITE_GOOGLE_API_KEY`
   - `VITE_GOOGLE_CLOUD_PROJECT_NUMBER`

Deploy flow:

1. Push to `main` (or run workflow manually from Actions tab).
2. Workflow builds and publishes `dist/` to Pages.
3. URL will be:
   - user site repo (`<owner>.github.io`): `https://<owner>.github.io/`
   - project repo: `https://<owner>.github.io/<repo>/`

Google Cloud allowlist for GitHub Pages:

1. OAuth Authorized JavaScript origins:
   - `https://<owner>.github.io`
2. API key referrer restrictions:
   - `https://<owner>.github.io/*`
   - if project page: `https://<owner>.github.io/<repo>/*`

## 5) Study behavior

- Mastery controls answer choices:
  - `< 0.40` => 2 choices
  - `0.40-0.79` => 4 choices
  - `>= 0.80` => 6 choices
- Study direction modes:
  - `Front Only` (question -> answer)
  - `Back Only` (answer -> question)
  - `Random` (mixes both per card)
- After correct answer behavior:
  - `Auto Next` with selectable delay (shows answer before moving on)
  - `Manual Next` to stay on the card until you click next
- Distractors are picked from matching tags first, then full deck fallback.
- `tags` is optional and comma-separated (example: `math,algebra`). It improves distractor relevance.
- Narration buttons use the free browser Web Speech API:
  - `Read Question`
  - `Read Answers`
  - `Stop Voice`
- Stats update on card completion:
  - `seen_count += 1`
  - first-try correct: `correct_count += 1`, `streak += 1`
  - any miss before completion: `wrong_count += 1`, `streak = 0`
  - `last_seen_at`, `last_result`, `mastery` recalculated

## 6) Sync behavior

- Pending writes are queued in memory.
- Auto-flush every 10 answers.
- Manual flush via `Sync Pending`.
- Also tries to flush when tab goes to background.

Writes use `spreadsheets.values.batchUpdate` with row-targeted ranges mapped from the header columns.
