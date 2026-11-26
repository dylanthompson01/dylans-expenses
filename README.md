# Dylan's Expenses

An expense tracker for a UCF student, built around 529 plan eligibility.
Snap a receipt, let AI read it, and it lands in Google Sheets with the image in Drive.

**Live:** https://dylanthompson01.github.io/dylans-expenses/

## What it does

- **Receipt scanning**: Gemini reads photos into place, amount, date, category and notes,
  three at a time
- **529 eligibility**: every expense is classified Yes / No / Not Sure against IRS Qualified
  Higher Education Expenses, including the groceries-vs-restaurants split that decides
  whether food counts as room & board
- **Google sync**: expenses live in a Google Sheet and receipt images in Drive, both through
  a Google Apps Script backend. Every save waits for Google to confirm it.
- **Edit and delete**: tap any expense. Deletes move the row to the Sheet's "Deleted" tab.
- **Budgets**: monthly limits per category (Groceries and Eating Out are separate), with a
  mark showing where you'd be if you spent evenly through the month
- **Reimbursements**: everything waiting to be paid back, multi-select to mark it paid, and
  a read-only link for Dad that shows only that list
- **Passcode**: the backend refuses every request without it
- **Installable**: add it to the iPhone home screen or install it on a Mac as an app

## Stack

Plain HTML/CSS/JavaScript in `index.html`. No frameworks, no npm, no build step.
The only other files are what installing as an app needs: `manifest.webmanifest`,
`sw.js` and `icons/`.

| Dependency | Purpose |
| --- | --- |
| Chart.js 4.4.1 + datalabels 2.2.0 | Charts |
| Bootstrap Icons 1.11.3 | Icons |
| Gemini `2.5-flash-lite` → `2.0-flash-lite` | Receipt reading (called from Apps Script) |
| Google Apps Script (`apps-script/Code.gs`) | Sheets + Drive backend |

## Setup (Apps Script)

1. Open the expenses Sheet → **Extensions → Apps Script**, replace the code with
   `apps-script/Code.gs`, and save.
2. **Project Settings → Script Properties**:
   - `GEMINI_KEY`: your Gemini API key
   - `PASSCODE`: what you'll type in the app (6+ characters)
3. **Deploy → Manage deployments → ✏️ → Version: New version → Deploy**. Editing the
   existing deployment keeps the same URL.

No secrets live in this repo. The page is public, so anything in it is public.

## Deploying the site

Pushes to `main` publish automatically via GitHub Pages.
