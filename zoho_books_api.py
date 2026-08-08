"""
Zoho Books API client — pulls Invoice, Bill, Expense, and Credit Note data
directly for a given period, replacing the earlier file-export approach.

Field names below are based on Zoho's published API docs, but Zoho's schema
has small regional variations (KSA edition vs. UK/US edition), and some
fields only show up once your org has VAT enabled for Saudi Arabia. Before
relying on this in production: call each *_for_period() function once
against your real organization, print the raw response, and confirm the
field names marked "verify" below match what your org actually returns —
I have not been able to test this against a live Knights Telecom
organization.

Auth: self-client OAuth (see https://www.zoho.com/books/api/v3/oauth/ ->
"Self Client"). You generate a refresh token once; this module exchanges it
for short-lived access tokens automatically on every cold start.

Environment variables required:
  ZOHO_CLIENT_ID
  ZOHO_CLIENT_SECRET
  ZOHO_REFRESH_TOKEN
  ZOHO_ORG_ID
  ZOHO_API_DOMAIN     default https://www.zohoapis.sa (KSA data center) —
                      use https://www.zohoapis.com if your org is on the
                      global/.com data center instead. Check the URL you
                      see logged into Zoho Books with: books.zoho.sa vs
                      books.zoho.com.
  ZOHO_ACCOUNTS_DOMAIN  default https://accounts.zoho.sa (must match the
                      data center above)
"""
import os
import time
import calendar
from datetime import date

import requests

_token_cache = {"access_token": None, "expires_at": 0}


def _get_access_token():
    now = time.time()
    if _token_cache["access_token"] and now < _token_cache["expires_at"] - 60:
        return _token_cache["access_token"]

    accounts_domain = os.environ.get("ZOHO_ACCOUNTS_DOMAIN", "https://accounts.zoho.sa")
    resp = requests.post(f"{accounts_domain}/oauth/v2/token", params={
        "refresh_token": os.environ["ZOHO_REFRESH_TOKEN"],
        "client_id": os.environ["ZOHO_CLIENT_ID"],
        "client_secret": os.environ["ZOHO_CLIENT_SECRET"],
        "grant_type": "refresh_token",
    })
    resp.raise_for_status()
    data = resp.json()
    if "access_token" not in data:
        raise RuntimeError(f"Zoho token refresh failed: {data}")
    _token_cache["access_token"] = data["access_token"]
    _token_cache["expires_at"] = now + data.get("expires_in", 3600)
    return _token_cache["access_token"]


def _headers():
    return {"Authorization": f"Zoho-oauthtoken {_get_access_token()}"}


def _api_root():
    return os.environ.get("ZOHO_API_DOMAIN", "https://www.zohoapis.sa") + "/books/v3"


def _org_id():
    return os.environ["ZOHO_ORG_ID"]


def _period_bounds(year, month):
    year, month = int(year), int(month)
    last_day = calendar.monthrange(year, month)[1]
    return date(year, month, 1).isoformat(), date(year, month, last_day).isoformat()


def _list_all(endpoint, extra_params, list_key):
    """Handles Zoho's page_context pagination, returns the full list across all pages."""
    results = []
    page = 1
    while True:
        params = {"organization_id": _org_id(), "page": page, "per_page": 200, **extra_params}
        resp = requests.get(f"{_api_root()}/{endpoint}", headers=_headers(), params=params)
        resp.raise_for_status()
        data = resp.json()
        results.extend(data.get(list_key, []))
        ctx = data.get("page_context", {})
        if not ctx.get("has_more_page"):
            break
        page += 1
    return results


def _get_detail(endpoint, record_id, detail_key):
    resp = requests.get(
        f"{_api_root()}/{endpoint}/{record_id}",
        headers=_headers(), params={"organization_id": _org_id()},
    )
    resp.raise_for_status()
    return resp.json()[detail_key]


# ---------------- Public: one function per Books module, per period ---------
def get_invoices_for_period(year, month):
    date_start, date_end = _period_bounds(year, month)
    listed = _list_all("invoices", {"date_start": date_start, "date_end": date_end}, "invoices")
    # The list endpoint omits line_items — fetch each invoice's detail to get them.
    return [_get_detail("invoices", inv["invoice_id"], "invoice") for inv in listed]


def get_creditnotes_for_period(year, month):
    date_start, date_end = _period_bounds(year, month)
    listed = _list_all("creditnotes", {"date_start": date_start, "date_end": date_end}, "creditnotes")
    return [_get_detail("creditnotes", cn["creditnote_id"], "creditnote") for cn in listed]


def get_bills_for_period(year, month):
    date_start, date_end = _period_bounds(year, month)
    listed = _list_all("bills", {"date_start": date_start, "date_end": date_end}, "bills")
    return [_get_detail("bills", b["bill_id"], "bill") for b in listed]


def get_expenses_for_period(year, month):
    date_start, date_end = _period_bounds(year, month)
    listed = _list_all("expenses", {"date_start": date_start, "date_end": date_end}, "expenses")
    return [_get_detail("expenses", e["expense_id"], "expense") for e in listed]
