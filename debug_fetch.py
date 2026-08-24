"""
debug_fetch.py
---------------
Quick standalone diagnostic — run this RIGHT NOW to find out exactly why
the app's dashboard fetch returned 0 alarms even though you can see 207 in
the dashboard's History view for GSP in July 2026.

HOW TO USE
----------
In the SAME CMD window where ALPHACORE_USER / ALPHACORE_PASS are already
set (or set them again first), run:

    python debug_fetch.py 2026 7

It will:
  1. Log in.
  2. Fetch ALL alarms tenant-wide for that month (no customer filter) —
     this is the same call alphacore_download.py already confirmed works.
  3. Print how many came back in total.
  4. Print every distinct customerId found, with how many alarms each one
     has — so you can match GSP's real count (207) against one of them.
  5. Tell you whether the Customer ID currently saved in
     clients_config.json for GSP matches any of those.
  6. Print the raw JSON of the first alarm, so we can see exactly what
     field holds the customer reference (in case it's not called
     "customerId" in this API's response).

Paste the FULL output back — that tells us exactly what to fix (usually
either: the saved Customer ID is wrong/mistyped, or the alarm objects use a
different field name than expected).
"""

import sys
import json
import datetime

import alphacore_api
import clients_store

IST = alphacore_api.IST


def main():
    if len(sys.argv) < 3:
        sys.exit("Usage: python debug_fetch.py <year> <month>   e.g. python debug_fetch.py 2026 7")
    year, month = int(sys.argv[1]), int(sys.argv[2])

    print(f"Logging in as {alphacore_api.USERNAME} ...")
    try:
        token = alphacore_api.login()
    except alphacore_api.AlphacoreAuthError as e:
        sys.exit(f"LOGIN FAILED: {e}")
    print("Login OK.\n")

    start_ts, end_ts = alphacore_api.month_range_ms(year, month)
    print(f"Fetching all alarms for {year}-{month:02d} "
          f"(startTime={start_ts}, endTime={end_ts}) ...")

    def progress(status, page, total):
        print(f"  status={status} page={page} running_total={total}")

    alarms = alphacore_api.fetch_all_alarms_for_month(token, year, month, progress_cb=progress)
    print(f"\nTOTAL ALARMS RETURNED (all clients combined): {len(alarms)}\n")

    if not alarms:
        print("Zero alarms came back for the WHOLE TENANT this month — the date-range filter "
              "itself is likely being ignored or mismatched by this API (same risk flagged in "
              "alphacore_download.py's comments). This is not a customer_id problem.")
        return

    print("Raw JSON of the first alarm (to check field names):")
    print(json.dumps(alarms[0], indent=2, default=str)[:3000])
    print()

    counts = {}
    for a in alarms:
        cid = alphacore_api.alarm_customer_id(a)
        counts[cid] = counts.get(cid, 0) + 1

    print("Distinct customerId values found, with alarm counts:")
    for cid, count in sorted(counts.items(), key=lambda x: -x[1]):
        print(f"  {cid!r}: {count} alarms")

    store = clients_store.load_store()
    gsp_entry = store.get("GSP", {})
    saved_id = gsp_entry.get("customer_id")
    print(f"\nCustomer ID currently saved for GSP in clients_config.json: {saved_id!r}")
    if saved_id in counts:
        print(f"  -> MATCHES a real customerId above, with {counts[saved_id]} alarms this month.")
    else:
        print("  -> Does NOT match any customerId seen in this month's data. "
              "That's why the app got 0 results — pick the correct ID from the list above "
              "(the one whose count is close to 207) and re-enter it in the app.")


if __name__ == "__main__":
    main()