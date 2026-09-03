# VCU Technology Services — Finance Dashboard

**Live site:** https://vcu-ts-finance-dashboard.netlify.app/

A live-updating dashboard covering TS expenditures, budgets, contracts, and purchase orders, built to sit alongside the department's Tableau workbooks.

## How it stays current

Data flows automatically, no manual steps required:

1. **`AppsScript_Code.gs`** runs hourly (Google Apps Script, `refreshAll` trigger) — pulls from the underlying Google Sheets/Excel sources, cleans and merges them, and writes CSVs into the "TS Finance Dashboard - Data Sources" Drive folder.
2. That same script pushes those CSVs into this repo's `data/` folder.
3. **Netlify** watches this repo and redeploys the live site automatically whenever `data/` changes.

This repo is **private** on purpose — the CSVs contain real department spending figures, so keeping the repo private (while still serving the site publicly via Netlify) avoids that data being publicly browsable/forkable on GitHub. No restricted data (e.g. the Chart of Accounts, which includes a financial manager's name) is ever committed here — that ID is stored only in the Apps Script project's Script Properties.

## If something looks stale or wrong

- Check the Apps Script project's **Executions** page for recent `refreshAll` runs and their status.
- A failed run (including a GitHub push failure, e.g. an expired token) triggers Apps Script's built-in failure-notification email automatically.
- The dashboard's header shows "source current through [date]" — computed from the data itself, so if it looks old, the underlying pipeline run is the place to look, not the website.
