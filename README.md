# Social Media Scraper — Economic Discourse Research

Academic thesis project that collects and transcribes social media posts about economic topics into a structured xlsx file.

---

## Before You Run — Critical Things to Know

### 1. The duplicate detection relies on column D being intact

The script reads every URL in **column D (Link to post)** of the target sheet on startup and puts them in a set. Before saving any new post it checks `if url in already_scraped`. This means:

- **Do not delete or edit URLs in column D** — if you do, the script will re-scrape those posts and create duplicates
- If you want to clear bad data, delete the **entire row**, not just the URL cell
- If you want a full fresh start on a sheet, delete all data rows (keep rows 1–4 which are account info and headers)

### 2. The msToken expires

The TikTok session token (`msToken`) embedded in the script expires **2026-05-30**. After that date the script will fail silently or return no results. To refresh it:

1. Log into tiktok.com in your browser with the research account (`cupcakebarry32`)
2. Open DevTools (`F12`) → Storage → Cookies → `https://www.tiktok.com` (not `.tiktok.com`)
3. Wait — there are **two** `msToken` entries. Use the one with domain `.tiktok.com` (not `www.tiktok.com`)
4. Copy the full value and update `MS_TOKEN` in `tiktok_scraper.py`

### 3. A Chromium window will open — do not close it

The script uses Playwright which launches its own Chromium browser (separate from Firefox). This window is the scraper's working environment. You can minimize it but **never close it** while the script is running or it will crash.

### 4. The Whisper model downloads on first use

The first time a video needs audio transcription, Whisper downloads its model (~140 MB). This happens once and is then cached. Do not interrupt the script during this download.

### 5. Ctrl+C is safe

The xlsx is saved after every keyword completes. If you stop mid-run you only lose the keyword currently being scraped. Everything before that is already written to disk.

---

## How Content Is Collected

For each video the script tries to get content in this order:

1. **TikTok auto-captions** — TikTok generates ASR (speech-to-text) subtitles for many videos. If an English subtitle track is found in the API response, it is downloaded and parsed. Fast and free.
2. **yt-dlp + Whisper** — if no captions exist, the audio is downloaded and transcribed locally using OpenAI Whisper. Slower (~5–30s per video depending on length). All temp audio files are deleted immediately after transcription.
3. **Skip** — if both fail, the video is skipped entirely. No fallback to the written caption/description.

The terminal shows `(tiktok-captions)` or `(whisper)` next to each collected post so you know the source.

---

## How Language Filtering Works

The script uses `langdetect` on the content text. If the detected language is not English the video is skipped. This correctly rejects Spanish, Portuguese, French, etc. — the old ASCII-ratio approach failed because those languages use mostly ASCII characters too.

---

## How Account Category Is Classified

Automatically assigned based on keywords found in the account's username, display name, and bio:

| Category | Signal |
|---|---|
| News outlet | Keywords: news, media, reuters, bbc, cnn, bloomberg, etc. |
| Business account | Verified account OR keywords: corp, bank, invest, capital, etc. |
| Dedicated account | Keywords: economy, finance, market, gdp, econ, trading, etc. |
| Personal account | Default — none of the above matched |

This is a heuristic, not perfect. Spot-check a sample of rows and correct manually if needed.

---

## Running the Script

```fish
cd ~/Documents/code/my-project

# default run — 40 posts × 10 keywords on TikTok
uv run tiktok_scraper.py

# custom number of posts per keyword
uv run tiktok_scraper.py -n 20

# custom keywords
uv run tiktok_scraper.py -k "economy" "inflation" "recession"

# everything custom
uv run tiktok_scraper.py -p tiktok -n 10 -k "housing market" "job market"

# see all options
uv run tiktok_scraper.py --help
```

---

## Output Structure (TikTok sheet)

Rows 1–4 are reserved (account credentials + column headers). Data starts at row 5.

| Column | Content |
|---|---|
| A — Number | Sequential row index |
| B — Keyword | Which keyword triggered this result |
| C — Content of post | Transcript or caption text |
| D — Link to post | Full TikTok URL ← duplicate detection reads this |
| E — Category of source | News outlet / Personal / Dedicated / Business |
| F — Likes | Like count at time of scraping |
| G — Reposts | Share count |
| H — Comments | Comment count |
| I — Date posted | When the video was published |
| J — Date gathered | Date the script ran |

---

## Progress Output Explained

```
[3/10 — 20%] Inflation (#inflation)
  [1/40] (whisper) @user | transcript preview...
  [2/40] (tiktok-captions) @user2 | transcript preview...
  Saved. Collected 40, skipped 11 | keyword took 284s | overall 30% | elapsed 0:14:22 | ETA ~0:33:21
```

- `[3/10 — 20%]` — which keyword out of total, overall % complete
- `(whisper)` / `(tiktok-captions)` — how content was obtained
- `skipped N` — non-English or untranscribable videos ignored
- `ETA` — estimated time remaining based on average time per completed keyword

---

## Keywords (defaults)

| Keyword | Hashtag searched |
|---|---|
| Economy | #economy |
| Recession | #recession |
| Economic Growth | #economicgrowth |
| Growth | #growth |
| Housing market | #housingmarket |
| Job market | #jobmarket |
| Inflation | #inflation |
| Global economy | #globaleconomy |
| EconDev | #econdev |
| Personal finance | #personalfinance |

---

## Files

| File | Purpose |
|---|---|
| `tiktok_scraper.py` | Main scraper — CLI entry point |
| `transcriber.py` | Reusable yt-dlp + Whisper microservice |
| `Data gathering example(1).xlsx` | The research data file (gitignored — contains credentials) |

---

## Gotchas

- **Temperature** — the script is CPU-intensive when Whisper is transcribing. 77–85°C is normal on a laptop. Keep the laptop on a hard flat surface with vents clear.
- **Session length** — a full default run (10 keywords × 40 posts) takes roughly 1–2 hours depending on how many videos need Whisper.
- **EconDev hashtag** — `#econdev` returns a lot of unrelated content (e.g. `#encode.dev`, ultrasound clinics). The English filter removes most of it but expect a higher skip rate for this keyword.
- **msToken and account** — the scraper does not visually log in. The token is injected as a cookie silently. The Chromium window looking "logged out" is normal.
