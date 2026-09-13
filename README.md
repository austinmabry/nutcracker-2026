# Nutcracker 2026 – My Schedule (auto-synced)

A phone-friendly page where a parent picks their dancer's role and cast and subscribes their calendar.
Every 6 hours a GitHub Action downloads the Ballet Arkansas master Google Sheet, rebuilds the page and
the 44 calendar files, and publishes them — so subscribed calendars follow the master automatically.

```
sync.py                 downloads + parses the master sheet, validates, writes docs/
template.html           the page; sync.py injects the data into it
docs/                   the published site (GitHub Pages serves this folder)
docs/ics/*.ics          one calendar per role + cast
CHANGES.md              appended by every sync that changed something (what was added / removed)
.github/workflows/      the schedule
```

## One-time setup (about 10 minutes)

1. **Create the repo.** github.com → **+** → New repository → name `nutcracker-2026`, **Public**, no README. Create.
2. **Upload this folder's contents** (Add file → Upload files, drag everything in including the `.github` and
   `docs` folders — or `git init` / `git push` from a terminal). Commit to `main`.
   *Note: the GitHub web uploader can't upload the hidden `.github` folder by drag-and-drop from some file
   managers. If it's missing after upload: Add file → Create new file → name it
   `.github/workflows/sync.yml` and paste the contents of that file.*
3. **Turn on Pages.** Settings → Pages → Source: *Deploy from a branch* → Branch `main`, folder **`/docs`** → Save.
4. **Let the Action push.** Settings → Actions → General → *Workflow permissions* → **Read and write permissions** → Save.
5. **Run it once by hand.** Actions tab → *Sync from BA master schedule* → Run workflow. It should finish green
   in about a minute and (if the master differs from the bundled copy) commit an update to `docs/`.
6. Open `https://<your-user>.github.io/nutcracker-2026/`. That's the link.

From then on it runs on its own every 6 hours. You don't need to do anything.

## What happens when Ballet Arkansas changes the sheet

- The next run picks it up, rewrites the calendar files, and commits. Phones that *subscribed* pull the
  change on their next refresh (iPhone: roughly daily by default; Google Calendar: up to ~24 h).
- `CHANGES.md` gets a dated entry listing exactly which events were added or removed. If you want a heads-up,
  watch the repo (Watch → Custom → Pushes) or just glance at the file.
- If the sheet's layout changes in a way the parser doesn't understand, the run **fails instead of publishing
  guesses** — the site keeps its last good version and GitHub emails you a failed-workflow notice. Open the
  failed run, read the `SYNC ABORTED – …` line, and either fix `sync.py` or send the error to Claude.

## How it decides what things are

- **Roles**: `tag()` in `sync.py` maps the sheet's wording ("All Soldiers", "Battle", "Party", "ACT I", …) to the
  22 community-cast roles. If BA introduces a new group name the run will abort with
  `no role recognised in '…'` and tell you which text needs a rule.
- **1st rehearsal (mandatory)**: the first time a role + cast appears in the weekend grid.
- **Mandatory w/ Company**: anything in the weekend grid after Thanksgiving, or marked "w/ Company".
- **Costume fitting**: any entry containing "Fitting". The backup fitting date is shown to every role.
- **Production week**: parsed line-by-line from the December cells (call times, time ranges, section headings).
- **Cast**: "Cast A" / "Cast B" in the entry text; otherwise both casts.

## Testing locally

```
pip install -r requirements.txt
python sync.py --check                 # download + parse + validate, write nothing
python sync.py --file some_copy.xlsx   # build docs/ from a local copy of the sheet
```

## Keep it private

Ballet Arkansas asked that this not be shared as a separate source. It's set up for one household's
convenience; the master schedule, the weekly emails and the BA portal remain the official word.
