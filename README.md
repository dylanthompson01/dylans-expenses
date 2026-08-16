# Dylan's Expenses

A single-file expense tracker for a UCF student, built around 529 plan eligibility.
Snap a receipt, let AI read it, and it lands in Google Sheets with the image in Drive.

**Live:** https://dylanthompson01.github.io/dylans-expenses/

## What it does

- **Receipt scanning** — photos are parsed by Gemini into place, amount, date, category, and notes
- **529 eligibility** — every expense is classified Yes / No / Not Sure against IRS Qualified
  Higher Education Expenses, including the groceries-vs-restaurants distinction that decides
  whether food counts as room & board
- **Google sync** — expenses append to a Google Sheet, receipt images upload to Drive, both
  through a Google Apps Script backend; `localStorage` caches everything so the app still
  works offline
- **Five views** — Upload, Dashboard, Cards, Table, and a Food tab that splits groceries from
  eating out

## Stack

Pure HTML/CSS/JavaScript in one `index.html`. No frameworks, no npm, no build step.

| Dependency | Purpose |
| --- | --- |
| Chart.js 4.4.1 + datalabels 2.2.0 | All charts |
| Bootstrap Icons 1.11.3 | Icons (pinned) |
| Gemini `2.5-flash-lite` → `2.0-flash-lite` | Receipt parsing, with fallback on overload |
| Google Apps Script | Sheets + Drive backend |

## Running it

Open `index.html` in a browser. That's the whole setup — there is nothing to install
and nothing to compile.

## Deploying

Pushes to `main` publish automatically via GitHub Pages.

```
git add .
git commit -m "your message"
git push
```

## A note on the API key

The Gemini key is embedded in `index.html`. This is unavoidable for a backend-less
client-side app — any key the browser uses is a key the user can read. It is restricted by
HTTP referrer in Google Cloud Console, so a copied key won't work from another domain.
Treat it as public and keep the referrer restriction in place.
