# Daily Prep Tracker

A one-page daily checklist for backend interview prep: gym, DSA, Spring Boot, system design, job hunt and revision, with a DSA streak and a 5-week grid.

Live: https://prep-tracker-nu-blond.vercel.app

## How sync works

- Ticking a box updates the page instantly and saves to the browser (localStorage).
- In the background, changes are batched (400 ms) and written to a Google Sheet, one row per day.
- The page pulls from the sheet on open, on tab focus and every 20 s, so other devices catch up.
- Offline ticks are queued and sent when the connection returns.

## Google Sheet setup (one time)

1. Open the sheet, then **Extensions → Apps Script**.
2. Replace the code with `apps-script/Code.gs` and save.
3. Pick **setup** in the function dropdown and click **Run**. It styles the Log tab and builds a Summary tab with charts.
4. **Deploy → New deployment → Web app**. Execute as: *Me*. Who has access: *Anyone*.
5. Authorize, copy the web app URL (ends in `/exec`).
6. Paste it into `SHEET_URL` at the top of the `<script>` in `index.html`, commit, redeploy.

Anyone with the `/exec` URL can write to the sheet, so don't share it.
