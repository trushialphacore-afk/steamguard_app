"""
find_customer_by_trap_name.py
------------------------------
When find_customer_id.py's alarm-COUNT list doesn't have a number that
matches what you see on the dashboard (e.g. you see "83 alarms" for a
client but no customerId in the list shows 83), guessing by count won't
work. This script finds the right customerId a different way: by matching
the TRAP NAME itself (e.g. every JK_Tyres trap probably has "JK" or
"Tyre" somewhere in its name) instead of guessing from a count.

HOW TO USE
----------
In the same terminal where ALPHACORE_PASS is set:

    python find_customer_by_trap_name.py JK 2026 7

(keyword, year, month — keyword is matched case-insensitively against every
trap/device name that raised an alarm that month; try a few different
keywords if the first one finds nothing — e.g. "JK", "Tyre", "Balamban" —
whatever text you'd expect to see in that client's trap names.)

It prints every distinct trap name that matched, grouped by which
customerId owns it, with alarm counts. If everything under one keyword
belongs to a single customerId, that's your answer — save it (this script
will offer to save it for you if only one customerId is found).

If NOTHING matches your keyword at all, two likely explanations:
  1. Wrong month — try the month you actually saw "83 alarms" in on the
     dashboard (not necessarily the report month).
  2. This client's traps aren't tagged with a customerId in the dashboard
     at all (they'd show up under customerId "None" instead) — in that
     case customerId-based fetching won't work for this client and we'll
     need a different approach (ask me and I'll help figure out a fix).
"""

import sys
import alphacore_api
import clients_store


def main():
    if len(sys.argv) != 4:
        sys.exit("Usage: python find_customer_by_trap_name.py <keyword> <year> <month>\n"
                  "Example: python find_customer_by_trap_name.py JK 2026 7")

    keyword = sys.argv[1].lower()
    year = int(sys.argv[2])
    month = int(sys.argv[3])

    print(f"Logging in as {alphacore_api.USERNAME} ...")
    try:
        token = alphacore_api.login()
    except alphacore_api.AlphacoreAuthError as e:
        sys.exit(f"LOGIN FAILED: {e}")
    print("Login OK.\n")

    print(f"Fetching all alarms for {year}-{month:02d} (whole tenant) ...")

    def progress(status, page, total):
        print(f"  status={status} page={page} running_total={total}")

    alarms = alphacore_api.fetch_all_alarms_for_month(token, year, month, progress_cb=progress)
    print(f"\nTOTAL ALARMS THIS MONTH: {len(alarms)}\n")

    matches = []
    for a in alarms:
        name = a.get("originatorName") or a.get("name") or ""
        if keyword in str(name).lower():
            cid = alphacore_api.alarm_customer_id(a)
            matches.append((cid, name))

    if not matches:
        print(f"No trap names containing '{keyword}' found in {year}-{month:02d}.\n")
        print("Try: (a) a different keyword, (b) a different month (the month you actually saw the "
              "count on the dashboard), or (c) this client's traps might have no customerId set at all "
              "(check with a keyword you're SURE is right, across a couple of months) — tell me what you "
              "find and I'll help from there.")
        return

    by_customer = {}
    for cid, name in matches:
        by_customer.setdefault(cid, {"count": 0, "sample_names": set()})
        by_customer[cid]["count"] += 1
        by_customer[cid]["sample_names"].add(name)

    ranked = sorted(by_customer.items(), key=lambda x: -x[1]["count"])
    print(f"Found {len(matches)} matching alarms across {len(ranked)} distinct customer ID(s):\n")
    for i, (cid, info) in enumerate(ranked, start=1):
        samples = ", ".join(list(info["sample_names"])[:3])
        print(f"  [{i}] {cid}   ({info['count']} matching alarms) — e.g. {samples}")

    print()
    if len(ranked) == 1:
        only_cid = ranked[0][0]
        client_name = input(f"Only one customerId matched '{keyword}' — save it as which client name? "
                             f"(type the exact client name, or press Enter to skip): ").strip()
        if client_name:
            entry = clients_store.upsert_client(client_name, customer_id=only_cid)
            print(f"\nSaved. '{client_name}' -> customer_id = {only_cid!r}")
            print(f"Full entry: {entry}")
        else:
            print("Nothing saved.")
    else:
        print("More than one customerId matched this keyword — look at the sample trap names above to "
              "tell which group is really your client, then save the right one with:\n"
              "  python update_customer_id.py <client_name> <customer_id>")


if __name__ == "__main__":
    main()