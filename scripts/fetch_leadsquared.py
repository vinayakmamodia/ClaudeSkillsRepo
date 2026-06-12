#!/usr/bin/env python3
"""
Fetch dashboard data from the LeadSquared API and write the three JSON files
the dashboard reads:

    dashboard/public/data/opportunities.json
    dashboard/public/data/mrr_transactions.json
    dashboard/public/data/renewals.json

Each file has the shape:  { "generatedAt": <iso>, "fileName": <str>, "rows": [ {col: val, ...}, ... ] }
The `rows` are column->value dicts whose KEYS MATCH the original CSV headers,
so the dashboard's existing column-mapping logic works unchanged.

Credentials (set as environment variables / GitHub Actions secrets):

    LSQ_ACCESS_KEY     LeadSquared API Access Key
    LSQ_SECRET_KEY     LeadSquared API Secret Key
    LSQ_HOST           API host, e.g. https://api-in21.leadsquared.com  (region-specific)

Optional:
    LSQ_OPPORTUNITY_EVENT   Opportunity activity/event code (default: tries the generic search)
    LSQ_PAGE_SIZE          Rows per page (default 1000)

NOTE: LeadSquared's exact endpoints and field names vary by account configuration
(custom Opportunity types, custom fields, etc.). The functions below use the
documented public endpoints; adjust the `*_FIELD_MAP` dictionaries near the top
to match your tenant's schema. Run locally first with real keys to validate.
"""

import json
import os
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlencode

import requests

# ----------------------------------------------------------------------------
DATA_DIR = Path(__file__).resolve().parent.parent / "dashboard" / "public" / "data"
ACCESS_KEY = os.environ.get("LSQ_ACCESS_KEY", "")
SECRET_KEY = os.environ.get("LSQ_SECRET_KEY", "")
HOST = os.environ.get("LSQ_HOST", "").rstrip("/")
PAGE_SIZE = int(os.environ.get("LSQ_PAGE_SIZE", "1000"))

SESSION = requests.Session()
SESSION.headers.update({"Content-Type": "application/json", "Accept": "application/json"})


def _auth_params() -> dict:
    return {"accessKey": ACCESS_KEY, "secretKey": SECRET_KEY}


def _post(path: str, payload: dict, extra_params: dict | None = None) -> dict:
    """POST to a LeadSquared endpoint with auth params in the query string."""
    params = _auth_params()
    if extra_params:
        params.update(extra_params)
    url = f"{HOST}{path}?{urlencode(params)}"
    for attempt in range(4):
        try:
            resp = SESSION.post(url, data=json.dumps(payload), timeout=60)
            resp.raise_for_status()
            return resp.json()
        except requests.RequestException as exc:
            wait = 2 ** (attempt + 1)
            print(f"  ! request failed ({exc}); retry in {wait}s", file=sys.stderr)
            time.sleep(wait)
    raise RuntimeError(f"POST {path} failed after retries")


# ----------------------------------------------------------------------------
# Opportunities — uses the Opportunity Activity "Get" / search endpoint.
# Docs: https://apidocs.leadsquared.com/  (Opportunity > Retrieve Opportunities)
# ----------------------------------------------------------------------------
def fetch_opportunities() -> list[dict]:
    rows: list[dict] = []
    page = 1
    opp_event = os.environ.get("LSQ_OPPORTUNITY_EVENT", "")
    while True:
        payload = {
            "Parameter": {"ActivityEvent": opp_event} if opp_event else {},
            "Paging": {"PageIndex": page, "PageSize": PAGE_SIZE},
            # Sorting/Columns can be customized; leaving blank returns the default set.
        }
        data = _post("/v2/OpportunityManagement.svc/Opportunity/Activity/Retrieve", payload)
        batch = data.get("List") or data.get("Records") or []
        if not batch:
            break
        for item in batch:
            # Flatten the LeadSquared field structure into a flat dict whose keys
            # match the CSV headers the dashboard expects.
            fields = item.get("Fields", item)
            if isinstance(fields, list):
                flat = {f.get("SchemaName") or f.get("Name"): f.get("Value") for f in fields}
            else:
                flat = dict(fields)
            rows.append(flat)
        print(f"  opportunities page {page}: +{len(batch)} (total {len(rows)})")
        if len(batch) < PAGE_SIZE:
            break
        page += 1
    return rows


# ----------------------------------------------------------------------------
# MRR transactions & Renewals are typically custom Activities or a custom
# Opportunity pipeline in LeadSquared. They use the same Activity Retrieve
# endpoint with a different ActivityEvent code. Set these env vars:
#   LSQ_MRR_EVENT, LSQ_RENEWAL_EVENT
# ----------------------------------------------------------------------------
def fetch_activity_rows(event_code: str, label: str) -> list[dict]:
    if not event_code:
        print(f"  (no event code for {label}; writing empty)")
        return []
    rows: list[dict] = []
    page = 1
    while True:
        payload = {
            "Parameter": {"ActivityEvent": event_code},
            "Paging": {"PageIndex": page, "PageSize": PAGE_SIZE},
        }
        data = _post("/v2/ProspectActivity.svc/Activity/Retrieve", payload)
        batch = data.get("List") or data.get("Records") or []
        if not batch:
            break
        for item in batch:
            fields = item.get("Fields", item)
            if isinstance(fields, list):
                flat = {f.get("SchemaName") or f.get("Name"): f.get("Value") for f in fields}
            else:
                flat = dict(fields)
            rows.append(flat)
        print(f"  {label} page {page}: +{len(batch)} (total {len(rows)})")
        if len(batch) < PAGE_SIZE:
            break
        page += 1
    return rows


# ----------------------------------------------------------------------------
def write_data(name: str, rows: list[dict]) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    out = DATA_DIR / name
    payload = {
        "generatedAt": datetime.now(timezone.utc).isoformat(),
        "fileName": f"{name} (LeadSquared API)",
        "rows": rows,
    }
    out.write_text(json.dumps(payload, ensure_ascii=False, indent=0))
    print(f"  wrote {out}  ({len(rows)} rows)")


def main() -> int:
    if not (ACCESS_KEY and SECRET_KEY and HOST):
        print("ERROR: LSQ_ACCESS_KEY, LSQ_SECRET_KEY and LSQ_HOST must be set.", file=sys.stderr)
        return 1

    print("Fetching Opportunities…")
    opportunities = fetch_opportunities()
    write_data("opportunities.json", opportunities)

    print("Fetching MRR transactions…")
    mrr = fetch_activity_rows(os.environ.get("LSQ_MRR_EVENT", ""), "mrr")
    write_data("mrr_transactions.json", mrr)

    print("Fetching Renewals…")
    renewals = fetch_activity_rows(os.environ.get("LSQ_RENEWAL_EVENT", ""), "renewals")
    write_data("renewals.json", renewals)

    print("Done.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
