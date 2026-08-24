"""
alphacore_api.py
-----------------
Live connection to the alphacore.live dashboard (ThingsBoard-based), used by
app.py's Phase-2 "auto-fetch from dashboard" mode instead of requiring a
manual alarm-file upload every month.

WHAT'S CONFIRMED WORKING (carried over from alphacore_download.py, which was
tested against the real dashboard earlier):
  - Login: POST /api/auth/login -> JWT token, sent as X-Authorization: Bearer <token>
  - Alarm listing: GET /api/v2/alarms?...&startTime=..&endTime=.. (whole
    tenant, all clients combined, for a date range) - statusList=ACTIVE and
    statusList=CLEARED fetched separately and merged.

WHAT IS **NOT** CONFIRMED YET (this file's job is to find out, the first
time you run it for real):
  - Which endpoint lists all customers/clients (there are several plausible
    ThingsBoard-style paths - see CUSTOMER_LIST_CANDIDATES below). This
    module tries each one in turn and remembers whichever one works in
    clients_config.json (via clients_store) so it only has to probe once.
  - Whether /api/v2/alarms accepts a customerId filter directly. To stay
    safe, this module does NOT depend on that - it fetches the month's
    alarms for the WHOLE tenant (confirmed working) and filters to one
    client's alarms itself, by matching each alarm's customerId against
    that client's stored customer_id. This is slightly less efficient than
    a server-side filter would be, but it doesn't depend on an unconfirmed
    parameter name.

HOW TO USE
----------
Set your password as an environment variable before running Streamlit,
exactly as with alphacore_download.py:
    Windows CMD:        set ALPHACORE_PASS=your_password
    Windows PowerShell:  $env:ALPHACORE_PASS="your_password"
    Mac/Linux:           export ALPHACORE_PASS=your_password

If the customer-list probe (fetch_customers) fails against the real
dashboard, run this file directly for a verbose report:
    python alphacore_api.py
and send me the printed output - I'll fix the candidate URL/param based on
the real error message, since I can't reach alphacore.live myself to test it.
"""

import os
import sys
import datetime
import requests

try:
    import streamlit as st
except ImportError:
    st = None

IST = datetime.timezone(datetime.timedelta(hours=5, minutes=30))

BASE_URL = "https://alphacore.live"
LOGIN_URL = f"{BASE_URL}/api/auth/login"
ALARMS_URL = f"{BASE_URL}/api/v2/alarms"


def _credential(name, default=None):
    """Read a login credential from an OS env var first (local/CMD usage,
    unchanged), then from Streamlit Secrets if available (cloud deployment,
    where there's no terminal to run `set`/`export` in). Resolved at call
    time (not at import time) so it works whichever way the app is started."""
    val = os.environ.get(name)
    if val:
        return val
    if st is not None:
        try:
            val = st.secrets.get(name, default)
            if val:
                return val
        except Exception:
            pass
    return default

PAGE_SIZE = 100
STATUSES_TO_FETCH = ["ACTIVE", "CLEARED"]
MAX_SANE_ALARMS = 20000

COMMON_HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"),
    "Accept": "application/json, text/plain, */*",
    "Content-Type": "application/json",
}

# Tried in order; the first one that returns a real list of customers wins,
# and that path gets cached (see clients_store / fetch_customers below) so
# future runs don't re-probe every candidate.
CUSTOMER_LIST_CANDIDATES = [
    "/api/customers",
    "/api/customer/customers",
    "/api/tenant/customers",
    "/api/v2/customers",
    "/api/v2/customer",
]


class AlphacoreAuthError(Exception):
    pass


class AlphacoreApiError(Exception):
    pass


def login():
    username = _credential("ALPHACORE_USER", "trushi@alphacore.co.in")
    password = _credential("ALPHACORE_PASS")
    if not password:
        raise AlphacoreAuthError(
            "ALPHACORE_PASS is not set. Locally: set it as an environment variable before "
            "running Streamlit. On Streamlit Cloud: add it under the app's Secrets instead."
        )
    resp = requests.post(LOGIN_URL, json={"username": username, "password": password}, headers=COMMON_HEADERS, timeout=20)
    if resp.status_code != 200:
        raise AlphacoreAuthError(f"Login failed (status {resp.status_code}): {resp.text[:300]}")
    return resp.json()["token"]


def _auth_headers(token):
    headers = dict(COMMON_HEADERS)
    headers["X-Authorization"] = f"Bearer {token}"
    return headers


def _try_customer_path(token, path):
    headers = _auth_headers(token)
    items = []
    page = 0
    while True:
        params = {"pageSize": 100, "page": page, "sortProperty": "title", "sortOrder": "ASC"}
        resp = requests.get(f"{BASE_URL}{path}", headers=headers, params=params, timeout=20)
        if resp.status_code != 200:
            return None, f"status {resp.status_code}: {resp.text[:200]}"
        body = resp.json()
        data = body.get("data", body if isinstance(body, list) else None)
        if not isinstance(data, list):
            return None, f"unexpected response shape: {str(body)[:200]}"
        items.extend(data)
        if not data or not body.get("hasNext", False):
            break
        page += 1
        if page > 50:
            break
    return items, None


def fetch_customers(token, cached_path=None):
    """
    Returns (customers, working_path). customers is a list of
    {"customerId": str, "name": str}. Tries `cached_path` first (the
    endpoint that worked last time, if known) before falling back to
    probing every candidate - so a known-good setup does one HTTP call,
    not five.
    """
    paths_to_try = ([cached_path] if cached_path else []) + [p for p in CUSTOMER_LIST_CANDIDATES if p != cached_path]
    errors = []
    for path in paths_to_try:
        items, err = _try_customer_path(token, path)
        if items is not None:
            customers = []
            for it in items:
                cid = it.get("id", {}).get("id") if isinstance(it.get("id"), dict) else it.get("id")
                title = it.get("title") or it.get("name") or ""
                if cid and title:
                    customers.append({"customerId": cid, "name": title})
            return customers, path
        errors.append(f"{path} -> {err}")
    raise AlphacoreApiError(
        "Could not list customers from any known endpoint. Tried:\n" + "\n".join(errors) +
        "\n\nRun `python alphacore_api.py` directly and share the output so the right endpoint can be added."
    )


def month_range_ms(year, month):
    start = datetime.datetime(year, month, 1, tzinfo=IST)
    end = datetime.datetime(year + 1, 1, 1, tzinfo=IST) if month == 12 else datetime.datetime(year, month + 1, 1, tzinfo=IST)
    return int(start.timestamp() * 1000), int(end.timestamp() * 1000)


def _fetch_alarms_page(token, page, status, start_ts, end_ts):
    headers = _auth_headers(token)
    params = {
        "pageSize": PAGE_SIZE, "page": page, "sortProperty": "createdTime", "sortOrder": "DESC",
        "statusList": status, "startTime": start_ts, "endTime": end_ts,
    }
    return requests.get(ALARMS_URL, headers=headers, params=params, timeout=30)


def fetch_all_alarms_for_month(token, year, month, progress_cb=None):
    """
    Fetches every alarm (all clients combined) for one calendar month - this
    part mirrors alphacore_download.py, which was confirmed working. Does
    NOT filter by client; see fetch_alarms_for_client() below for that.
    """
    start_ts, end_ts = month_range_ms(year, month)
    all_alarms, seen_ids = [], set()
    for status in STATUSES_TO_FETCH:
        page = 0
        while True:
            resp = _fetch_alarms_page(token, page, status, start_ts, end_ts)
            if resp.status_code != 200:
                raise AlphacoreApiError(f"Alarm fetch failed (status={status}, page={page}): "
                                         f"{resp.status_code}: {resp.text[:300]}")
            body = resp.json()
            data = body.get("data", [])
            for alarm in data:
                alarm_id = alarm.get("id", {}).get("id") if isinstance(alarm.get("id"), dict) else alarm.get("id")
                if alarm_id in seen_ids:
                    continue
                seen_ids.add(alarm_id)
                all_alarms.append(alarm)
            if progress_cb:
                progress_cb(status, page, len(all_alarms))
            if not data or not body.get("hasNext", False) or len(all_alarms) > MAX_SANE_ALARMS:
                break
            page += 1
    return all_alarms


def alarm_customer_id(alarm):
    cid = alarm.get("customerId")
    if isinstance(cid, dict):
        return cid.get("id")
    return cid


def filter_alarms_for_customer(all_alarms, customer_id):
    """Pure filter step, split out from fetch_alarms_for_client() so a
    caller (see app.py) can fetch+cache the whole-tenant month once and
    reuse it for every client's report that month, instead of re-fetching
    the whole tenant from scratch for each client."""
    return [a for a in all_alarms if alarm_customer_id(a) == customer_id]


def fetch_alarms_for_client(token, customer_id, year, month, progress_cb=None):
    """
    Fetches the whole tenant's alarms for the month (confirmed-working
    endpoint) and filters down to just this client's customer_id locally -
    see the module docstring for why it's done this way instead of trusting
    an unconfirmed server-side customer filter param.
    """
    all_alarms = fetch_all_alarms_for_month(token, year, month, progress_cb=progress_cb)
    return filter_alarms_for_customer(all_alarms, customer_id)


def flatten(alarm):
    row = {}
    for k, v in alarm.items():
        if isinstance(v, dict):
            for sk, sv in v.items():
                row[f"{k}_{sk}"] = sv
        else:
            row[k] = v
    for tkey in ("createdTime", "startTs", "endTs", "ackTs", "clearTs"):
        if tkey in row and isinstance(row[tkey], (int, float)) and row[tkey] > 0:
            row[tkey + "_readable"] = datetime.datetime.fromtimestamp(row[tkey] / 1000, tz=IST).strftime("%Y-%m-%d %H:%M:%S")
    return row


if __name__ == "__main__":
    print("Logging in...")
    try:
        tok = login()
    except AlphacoreAuthError as e:
        sys.exit(str(e))
    print("Login OK.\n")

    print("Probing customer-list endpoints...")
    try:
        customers, working_path = fetch_customers(tok)
        print(f"SUCCESS with {working_path} - found {len(customers)} customers:")
        for c in customers:
            print(f"  {c['customerId']}  {c['name']}")
    except AlphacoreApiError as e:
        print(str(e))