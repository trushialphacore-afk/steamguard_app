"""
remove_client.py
------------------
Delete one client's entry from clients_config.json entirely — e.g. a
duplicate name, or a client that's no longer tracked (like MRF, which is on
monthly rental so it's out of scope for this one-time-investment tool).

HOW TO USE
----------
    python remove_client.py MRF_Tyres

This does NOT touch app.py's dropdown list — if the client is also in
FALLBACK_CLIENTS there, remove its name from that list too so it stops
showing up as an option.
"""

import sys
import clients_store

if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit("Usage: python remove_client.py <client_name>\nExample: python remove_client.py MRF_Tyres")
    name = sys.argv[1]
    removed = clients_store.remove_client(name)
    if removed:
        print(f"Removed '{name}' from clients_config.json.")
    else:
        print(f"'{name}' wasn't found in clients_config.json — nothing to remove.")