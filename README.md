# Flash Me Baby

German–English flashcards with Tinder-style swiping, points, streaks, streak freezes and live Google Sheet sync. About 11,000 words from A1 to C2, plus IT & Management, Software Development and Cybersecurity topics.

## Run it on GitHub Pages

1. Create a repository and upload everything in this folder (all files sit side by side, no folders).
2. Repository **Settings → Pages → Source: Deploy from a branch**, branch `main`, folder `/ (root)`.
3. Open `https://<your-user>.github.io/<repo>/`.

## Install it like an app

- **Android (Chrome):** ⋮ menu → **Install app** (or **Add to Home screen**).
- **Desktop (Chrome / Edge):** install icon in the address bar, or ⋮ → **Cast, save and share → Install page as app**.
- **iPhone (Safari):** Share → **Add to Home Screen**.

Installing only works from the `https://` GitHub Pages address, not from a local file.

## Google Sheet sync

1. In your Google Sheet: **Extensions → Apps Script**, paste `flash-me-baby-sheet-sync.gs`.
2. Set your own long random `TOKEN` in the editor only. **Never commit your real token.** Keep the `CHANGE-ME` placeholder in this repo.
3. Run `setup`, then **Deploy → New deployment → Web app** (Execute as: Me, Who has access: Anyone).
4. In the app: **Settings → Google Sheet sync**, paste the `/exec` URL and your token, tap **Connect**.

## Updating

After changing `index.html`, bump `VERSION` in `sw.js` so installed copies refresh.
