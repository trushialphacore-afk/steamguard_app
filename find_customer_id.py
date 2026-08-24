"""
find_customer_id.py
--------------------
Easier replacement for "dig through the browser's Network tab" — this finds
a client's correct dashboard Customer ID for you and saves it directly,
in ONE command.

WHY THIS EXISTS
----------------
The dashboard account doesn't have permission to list customers
(/api/customers returns a 403), so the app can't look up a client's
Customer ID automatically. But it CAN fetch every alarm for the whole
tenant for a given month (that always works) — and every alarm carries the
real customerId. So instead of hunting through DevTools, this script:

  1. Logs in.
  2. Fetches all alarms for the month you give it (all clients combined).
  3. Groups them by customerId and counts how many alarms each one has.
  4. Shows you the list, ranked by count.
  5. You tell it which one is the right client (usually the one whose count
     matches what you see in the dashboard's History view) and it SAVES it
     to clients_config.json immediately — no separate save step needed.

HOW TO USE
----------
In the same CMD window where ALPHACORE_PASS is set:

    python find_customer_id.py BKT 2026 6

If you already know roughly how many alarms that client should have this
month (e.g. you saw "1991" in the dashboard), pass it as a 4th argument and
the script will point out the closest match for you:

    python find_customer_id.py BKT 2026 6 1991

Then just confirm when prompted, and BKT's Customer ID is saved — you won't
need to touch this again for BKT.
"""

import sys
import alphacore_api
import clients_store


def main():
    if len(sys.argv) < 4:
        sys.exit("Usage: python find_customer_id.py <client_name> <year> <month> [expected_alarm_count]\n"
                  "Example: python find_customer_id.py BKT 2026 6 1991")

    client_name = sys.argv[1]
    year = int(sys.argv[2])
    month = int(sys.argv[3])
    expected_count = int(sys.argv[4]) if len(sys.argv) > 4 else None

    print(f"Logging in as {alphacore_api.USERNAME} ...")
    try:
        token = alphacore_api.login()
    except alphacore_api.AlphacoreAuthError as e:
        sys.exit(f"LOGIN FAILED: {e}")
    print("Login OK.\n")

    print(f"Fetching all alarms for {year}-{month:02d} (whole tenant, all clients combined) ...")

    def progress(status, page, total):
        print(f"  status={status} page={page} running_total={total}")

    alarms = alphacore_api.fetch_all_alarms_for_month(token, year, month, progress_cb=progress)
    print(f"\nTOTAL ALARMS THIS MONTH (all clients combined): {len(alarms)}\n")

    if not alarms:
        sys.exit("Zero alarms came back for the WHOLE TENANT this month. That's not a customer_id "
                  "problem — either the date range is wrong or the API itself returned nothing. "
                  "Double check the year/month and try again.")

    counts = {}
    for a in alarms:
        cid = alphacore_api.alarm_customer_id(a)
        counts[cid] = counts.get(cid, 0) + 1

    ranked = sorted(counts.items(), key=lambda x: -x[1])

    print(f"Found {len(ranked)} distinct customer IDs this month:\n")
    for i, (cid, count) in enumerate(ranked, start=1):
        marker = ""
        if expected_count is not None and abs(count - expected_count) <= max(2, expected_count * 0.02):
            marker = "   <-- closest match to the count you gave"
        print(f"  [{i}] {cid}   ({count} alarms){marker}")

    print()
    choice = input(f"Which number is '{client_name}'? Type the [number] shown in brackets above and press "
                    f"Enter (or press Enter to skip without saving): ").strip()
    if not choice:
        print("Nothing saved. Re-run this script once you know which one is correct.")
        return

    # Accept either the list index (e.g. "2") OR the raw customer ID pasted
    # directly (e.g. "f0e19830-ba03-11f0-9333-67f47b798878") — people
    # naturally paste the ID itself, so both should just work.
    matching_ids = [cid for cid, _ in ranked if cid == choice]
    if matching_ids:
        chosen_cid = matching_ids[0]
    elif "-" in choice and len(choice) > 20:
        # Looks like a raw customer ID (UUID-style) that just wasn't one of
        # the ones seen this month — save it anyway rather than blocking.
        print(f"Note: {choice!r} wasn't among this month's customer IDs above, but saving it as given.")
        chosen_cid = choice
    else:
        try:
            idx = int(choice)
            chosen_cid = ranked[idx - 1][0]
        except (ValueError, IndexError):
            sys.exit("That wasn't a valid number from the list above, and it didn't look like a customer ID "
                      "either. Nothing saved — try again with just the [number] in brackets.")

    entry = clients_store.upsert_client(client_name, customer_id=chosen_cid)
    print(f"\nSaved. '{client_name}' -> customer_id = {chosen_cid!r}")
    print(f"Full entry: {entry}")
    print(f"\n'{client_name}' will now auto-fetch in the app — you won't be asked for this again.")


if __name__ == "__main__":
    main()