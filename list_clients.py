"""
list_clients.py
-----------------
Prints every client currently saved in clients_config.json — exact name
spelling, and whether its Customer ID / Traps Monitored / Investment are
already set or still missing.

HOW TO USE
----------
    python list_clients.py

(no arguments needed, no login/internet needed — just reads the local
clients_config.json file.)
"""

import clients_store


def main():
    store = clients_store.load_store()
    if not store:
        print("No clients saved yet in clients_config.json.")
        return

    print(f"{len(store)} client(s) saved:\n")
    print(f"{'Client name':<25} {'Customer ID':<10} {'Traps':<10} {'Investment':<10}")
    print("-" * 60)
    for name, entry in sorted(store.items()):
        cid_status = "set" if entry.get("customer_id") else "MISSING"
        traps_status = str(entry.get("traps_monitored")) if entry.get("traps_monitored") is not None else "MISSING"
        inv_status = str(entry.get("investment")) if entry.get("investment") is not None else "MISSING"
        print(f"{name:<25} {cid_status:<10} {traps_status:<10} {inv_status:<10}")

    print("\n(Exact spelling above is exactly how each name is stored — this must match what you pick "
          "in the app's Client dropdown for its saved details to be found.)")


if __name__ == "__main__":
    main()