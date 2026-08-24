"""
set_client_details.py
-----------------------
Set a client's Traps Monitored and/or Investment WITHOUT opening the app
or generating a report — just saves straight to clients_config.json.

HOW TO USE
----------
Set traps only:
    python set_client_details.py JK_Tyres --traps 40

Set investment only:
    python set_client_details.py JK_Tyres --investment 250000

Set both at once:
    python set_client_details.py JK_Tyres --traps 40 --investment 250000

Whichever ones you have ready, pass those flags — leave out the ones you
don't know yet, and they'll just stay MISSING until you run this again
later with that value. Check anytime with:
    python list_clients.py
"""

import sys
import clients_store


def main():
    args = sys.argv[1:]
    if not args:
        sys.exit("Usage: python set_client_details.py <client_name> [--traps N] [--investment N]\n"
                  "Example: python set_client_details.py JK_Tyres --traps 40 --investment 250000")

    name = args[0]
    traps = None
    investment = None

    i = 1
    while i < len(args):
        if args[i] == "--traps" and i + 1 < len(args):
            try:
                traps = int(args[i + 1])
            except ValueError:
                sys.exit(f"'--traps' value '{args[i + 1]}' isn't a whole number.")
            i += 2
        elif args[i] == "--investment" and i + 1 < len(args):
            try:
                investment = float(args[i + 1])
            except ValueError:
                sys.exit(f"'--investment' value '{args[i + 1]}' isn't a number.")
            i += 2
        else:
            sys.exit(f"Didn't understand '{args[i]}'. Use --traps N and/or --investment N.")

    if traps is None and investment is None:
        sys.exit("Nothing to set — pass --traps N and/or --investment N.")

    entry = clients_store.upsert_client(name, traps_monitored=traps, investment=investment)
    print(f"Saved for '{name}':")
    if traps is not None:
        print(f"  Traps Monitored = {traps}")
    if investment is not None:
        print(f"  Investment = Rs {investment:,.0f} (one-time total)")
    print(f"\nFull entry: {entry}")
    print("\nCheck all clients anytime with: python list_clients.py")


if __name__ == "__main__":
    main()