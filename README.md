# LeadSquared FY27 Pipeline Dashboard (automated)

A self-updating version of the pipeline dashboard you originally built as a
Claude Artifact. Instead of manually exporting CSVs from LeadSquared and
uploading them every day, this version **pulls the data automatically** on a
daily schedule and serves a live dashboard via GitHub Pages.

```
LeadSquared API
      │  (scripts/fetch_leadsquared.py — runs daily in GitHub Actions)
      ▼
dashboard/public/data/*.json   ← committed/deployed snapshots
      │  (React app fetches these on load)
      ▼
GitHub Pages dashboard         ← always shows the latest data, zero uploads
```

## What changed from the artifact

| Artifact (manual)                         | This version (automated)                          |
| ----------------------------------------- | ------------------------------------------------- |
| `window.storage` / `window.fs` (Claude)   | browser `localStorage` + `fetch('/data/*.json')`  |
| Upload a CSV every day                    | Daily GitHub Actions job fetches via the API      |
| Lives only in a Claude chat               | Hosted as a static site on GitHub Pages           |
| Manual upload is the only path            | Manual upload still works as a fallback / override |

The dashboard logic (Pipeline Number View, MRR Performance, Renewal Risks,
all filters, drill-downs, targets) is unchanged.

## One-time setup

1. **Enable GitHub Pages**: repo → Settings → Pages → Build and deployment →
   Source = **GitHub Actions**.

2. **Add LeadSquared API credentials** as repository secrets
   (Settings → Secrets and variables → Actions → New repository secret):

   | Secret | Value |
   | ------ | ----- |
   | `LSQ_ACCESS_KEY` | LeadSquared API Access Key |
   | `LSQ_SECRET_KEY` | LeadSquared API Secret Key |
   | `LSQ_HOST` | Your region host, e.g. `https://api-in21.leadsquared.com` |
   | `LSQ_OPPORTUNITY_EVENT` | (optional) Opportunity activity/event code |
   | `LSQ_MRR_EVENT` | (optional) MRR transaction activity code |
   | `LSQ_RENEWAL_EVENT` | (optional) Renewal activity code |

   Find your Access/Secret keys in LeadSquared → **Settings → API and Webhooks
   → LeadSquared API**. The host depends on your account region (check the URL
   you log in to, e.g. `in21`, `us11`, etc.).

3. **Adjust field mapping if needed.** LeadSquared's field names vary per
   tenant (custom Opportunity types/fields). `scripts/fetch_leadsquared.py`
   flattens API responses into rows whose keys match the original CSV headers
   the dashboard expects. Run it locally first with real keys to confirm the
   columns line up:

   ```bash
   pip install -r scripts/requirements.txt
   LSQ_ACCESS_KEY=... LSQ_SECRET_KEY=... LSQ_HOST=https://api-inXX.leadsquared.com \
     python scripts/fetch_leadsquared.py
   ```

   This writes `dashboard/public/data/*.json`. Open the dashboard locally to
   verify (see below), then commit.

## Run the dashboard locally

```bash
cd dashboard
npm install
npm run dev      # http://localhost:5173
```

The app auto-loads `public/data/*.json`. The **Sync** button (top-right)
re-fetches them on demand. You can still drag-drop / upload CSVs manually to
override for a session.

## Schedule

The daily job runs at **02:00 UTC** (`~07:30 IST`). Change the `cron` in
`.github/workflows/deploy.yml` to suit. You can also trigger it any time from
the **Actions** tab → "Fetch data & deploy dashboard" → **Run workflow**.

## Project layout

```
dashboard/                 Vite + React + Tailwind app
  src/App.tsx              the full dashboard (adapted from your artifact)
  src/storage.ts           localStorage shim + data URL helper
  public/data/*.json       data snapshots (overwritten by the fetch script)
scripts/
  fetch_leadsquared.py     pulls Opportunities / MRR / Renewals from the API
  requirements.txt
.github/workflows/
  deploy.yml               daily fetch → build → deploy to Pages
```
