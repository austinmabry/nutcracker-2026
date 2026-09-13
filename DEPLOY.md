# Nutcracker 2026 – My Schedule (static site)

Everything in this folder is static. Host it anywhere that serves plain files over https
(GitHub Pages, Cloudflare Pages, Netlify, or a folder behind your Cloudflare Tunnel).

## GitHub Pages (about 5 minutes)
1. Create a new public repo, e.g. `nutcracker-2026`.
2. Put the contents of this folder at the repo root (`index.html`, `ics/`, `.nojekyll`).
3. Repo → Settings → Pages → Source: "Deploy from a branch", branch `main`, folder `/ (root)`. Save.
4. After a minute the site is live at `https://<user>.github.io/nutcracker-2026/`.
   Share that link. Parents pick their role/cast and tap the button for their phone.

## How the buttons work
- **Add to iPhone Calendar** → `webcal://…/ics/Role-CastX.ics`. iOS opens Calendar and asks to subscribe.
  The dates then auto-refresh (iOS checks the file roughly daily).
- **Add to Google Calendar** → `https://calendar.google.com/calendar/r?cid=webcal://…`. Google adds it as a
  subscribed ("From URL") calendar; it appears in the Google Calendar app on Android and iPhone after sync.
- **Download .ics** → one-time import, for anything else (Outlook, Samsung Calendar, desktop).

## When the schedule changes
Regenerate the `ics/` files (or edit them by hand) and push. Subscribed phones pick up the change on their next refresh
— nobody has to re-import anything. Parents who used "Download .ics" will need to import again.

## Files
- `index.html` – the picker page (all data embedded, no external dependencies)
- `ics/<Role>-Cast<A|B>.ics` – 44 pre-built calendars, one per role and cast
- `.nojekyll` – tells GitHub Pages to serve files as-is
