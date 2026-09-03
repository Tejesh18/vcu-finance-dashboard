# VCU Technology Services — Finance Dashboard

**Live site:** https://tejesh18.github.io/vcu-finance-dashboard/

A live-updating dashboard covering TS expenditures, budgets, contracts, and purchase orders, built to sit alongside the department's Tableau workbooks.

## How it stays current

Data flows automatically, no manual steps required:

1. **`AppsScript_Code.gs`** runs hourly (Google Apps Script, `refreshAll` trigger) — pulls from the underlying Google Sheets/Excel sources, cleans and merges them, and writes CSVs into the "TS Finance Dashboard - Data Sources" Drive folder.
2. That same script pushes those CSVs into this repo's `data/` folder.
3. **GitHub Pages** serves the site directly from this repo — it picks up new commits automatically within a minute or two, no separate build step or deploy quota to worry about.

This repo is currently **public**, which means the CSVs (real department spending figures) are technically browsable/forkable on GitHub, not just viewable through the dashboard. This was a deliberate, pragmatic call for now: a private-repo setup (via Netlify) ran into that host's monthly deploy-quota limit at this pipeline's hourly refresh rate. Worth revisiting once there's time to either reduce the refresh frequency, move to a host without that constraint (e.g. Firebase Hosting), or decide the public tradeoff is fine long-term. No *restricted* data (e.g. the Chart of Accounts, which includes a financial manager's name) is ever committed here regardless — that ID is stored only in the Apps Script project's Script Properties.

## If something looks stale or wrong

- Check the Apps Script project's **Executions** page for recent `refreshAll` runs and their status.
- A failed run (including a GitHub push failure, e.g. an expired token) triggers Apps Script's built-in failure-notification email automatically.
- The dashboard's header shows "source current through [date]" — computed from the data itself, so if it looks old, the underlying pipeline run is the place to look, not the website.
- If GitHub has fresh commits in `data/` but the live site doesn't reflect them, check whether the hosting side (currently GitHub Pages) is actually redeploying — this exact disconnect happened once already with Netlify running out of its free deploy quota.
