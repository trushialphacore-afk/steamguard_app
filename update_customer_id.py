"""
update_customer_id.py
----------------------
One-off helper: save/overwrite a client's dashboard Customer ID in
clients_config.json without having to edit the JSON file by hand.

HOW TO USE
----------
    python update_customer_id.py GSP 3fe9a6a0-dcd4-11f0-b708-e5039c453036

(client name, then the Customer ID — same two things you'd otherwise type
into the app's "Dashboard Customer ID" box.)
"""

import sys
import clients_store

if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit("Usage: python update_customer_id.py <client_name> <customer_id>")
    name, cid = sys.argv[1], sys.argv[2]
    entry = clients_store.upsert_client(name, customer_id=cid)
    print(f"Saved. '{name}' now maps to customer_id={cid!r}")
    print(f"Full entry: {entry}")