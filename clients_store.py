"""
clients_store.py
-----------------
Small persistent store for things that must NOT be re-entered every time a
report is generated:

  1. customer_id       - the dashboard's internal customerId (UUID) for each
                          client, so the app can auto-fetch that client's
                          alarms without you having to look up IDs by hand.
  2. traps_monitored   - the client's TRUE installed trap count (e.g. 25 for
                          GSP). Fixed per client, not per month.
  3. investment        - the client's one-time total SteamGuard investment
                          (e.g. Rs 5,64,000 for GSP). Fixed per client, not
                          per month.
  4. ROI assumptions   - pressure_bar, orifice_mm, cost_per_ton,
                          detect_minutes: read once from the client's ROI
                          sheet and reused every month after that, instead
                          of re-uploading the ROI sheet each time.

STORAGE
-------
Backed by a JSON file (clients_config.json) next to this script — safe to
open/edit by hand if needed:

{
  "GSP": {"customer_id": "1234-...-uuid", "traps_monitored": 25, "investment": 564000,
          "pressure_bar": 14, "orifice_mm": 3, "cost_per_ton": 3200, "detect_minutes": 15}
}

CLOUD DEPLOYMENT (Streamlit Community Cloud)
---------------------------------------------
A file written by the app on Streamlit Cloud is NOT guaranteed to survive a
restart/redeploy of the app (the container's disk is not permanent). To make
saved clients survive that, without anyone needing to run a command or edit
a file by hand, this module can optionally sync clients_config.json straight
to the app's own GitHub repo using the GitHub Contents API:

  - Add these to the app's Streamlit Cloud "Secrets" (Settings -> Secrets):
        GITHUB_TOKEN = "ghp_..."      # a GitHub Personal Access Token with
                                       # "repo" (or fine-grained Contents:
                                       # Read & Write) permission on the app's
                                       # repo
        GITHUB_REPO  = "your-org/steamguard_app"
        # optional, default "main":
        # GITHUB_BRANCH = "main"
  - Once those two secrets are set, every save_store() call (which
    upsert_client() triggers) automatically commits the updated
    clients_config.json straight to that GitHub repo - no manual GitHub
    edit, no CMD, nothing to remember. The next time the app restarts (even
    on a totally fresh container) it reads the latest version back from
    GitHub first.
  - If those secrets are NOT set (e.g. running locally on your laptop, same
    as always), this module behaves exactly as before: a plain local
    clients_config.json file next to this script. Nothing changes for local
    use.
  - If the GitHub sync call ever fails for any reason (bad token, network
    hiccup), it silently falls back to saving the local file too, so a
    client is never lost even if the "no restart" persistence hiccups for a
    moment.
"""

import base64
import json
import os

import requests

try:
    import streamlit as st
except ImportError:
    st = None

STORE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "clients_config.json")

ROI_ASSUMPTION_FIELDS = ["pressure_bar", "orifice_mm", "cost_per_ton", "detect_minutes"]


def _secret(name, default=None):
    """Read a config value from an OS env var first (local/CMD usage,
    unchanged), then from Streamlit Secrets if available (cloud usage)."""
    val = os.environ.get(name)
    if val:
        return val
    if st is not None:
        try:
            return st.secrets.get(name, default)
        except Exception:
            pass
    return default


def _github_config():
    token = _secret("GITHUB_TOKEN")
    repo = _secret("GITHUB_REPO")
    if not token or not repo:
        return None
    return {
        "token": token,
        "repo": repo,
        "branch": _secret("GITHUB_BRANCH", "main"),
        "path": _secret("GITHUB_CLIENTS_FILE", "clients_config.json"),
    }


def _github_headers(token):
    return {"Authorization": f"token {token}", "Accept": "application/vnd.github+json"}


def _github_get_file(cfg):
    url = f"https://api.github.com/repos/{cfg['repo']}/contents/{cfg['path']}"
    resp = requests.get(url, headers=_github_headers(cfg["token"]), params={"ref": cfg["branch"]}, timeout=15)
    if resp.status_code == 404:
        return {}, None
    resp.raise_for_status()
    body = resp.json()
    content = base64.b64decode(body["content"]).decode("utf-8")
    try:
        return json.loads(content), body["sha"]
    except json.JSONDecodeError:
        return {}, body["sha"]


def _github_put_file(cfg, store, sha):
    url = f"https://api.github.com/repos/{cfg['repo']}/contents/{cfg['path']}"
    payload = {
        "message": "Update clients_config.json via SteamGuard Report Generator",
        "content": base64.b64encode(json.dumps(store, indent=2, ensure_ascii=False).encode("utf-8")).decode("ascii"),
        "branch": cfg["branch"],
    }
    if sha:
        payload["sha"] = sha
    resp = requests.put(url, headers=_github_headers(cfg["token"]), json=payload, timeout=15)
    resp.raise_for_status()


def load_store():
    cfg = _github_config()
    if cfg:
        try:
            store, _ = _github_get_file(cfg)
            return store
        except Exception:
            pass  # fall through to the local file as a safety net
    if not os.path.exists(STORE_PATH):
        return {}
    with open(STORE_PATH, "r", encoding="utf-8") as f:
        try:
            return json.load(f)
        except json.JSONDecodeError:
            return {}


def save_store(store):
    cfg = _github_config()
    synced_to_github = False
    if cfg:
        try:
            _, sha = _github_get_file(cfg)
            _github_put_file(cfg, store, sha)
            synced_to_github = True
        except Exception:
            pass  # still save locally below so nothing is lost
    # Always keep a local copy too - cheap, and it's the only copy at all
    # when no GitHub secrets are configured (e.g. running locally).
    with open(STORE_PATH, "w", encoding="utf-8") as f:
        json.dump(store, f, indent=2, ensure_ascii=False)
    return synced_to_github


def get_client(name):
    return load_store().get(name)


def upsert_client(name, customer_id=None, traps_monitored=None, investment=None,
                   pressure_bar=None, orifice_mm=None, cost_per_ton=None, detect_minutes=None):
    """
    Update (or create) one client's entry. Only overwrites the fields you
    actually pass in - e.g. calling upsert_client("GSP", traps_monitored=25)
    won't clobber an already-known customer_id, investment, or ROI
    assumptions.
    """
    store = load_store()
    entry = store.get(name, {})
    if customer_id is not None:
        entry["customer_id"] = customer_id
    if traps_monitored is not None:
        entry["traps_monitored"] = traps_monitored
    if investment is not None:
        entry["investment"] = investment
    if pressure_bar is not None:
        entry["pressure_bar"] = pressure_bar
    if orifice_mm is not None:
        entry["orifice_mm"] = orifice_mm
    if cost_per_ton is not None:
        entry["cost_per_ton"] = cost_per_ton
    if detect_minutes is not None:
        entry["detect_minutes"] = detect_minutes
    store[name] = entry
    save_store(store)
    return entry


def get_saved_roi_assumptions(name):
    """Returns whichever of the 4 ROI-sheet-derived assumptions are already
    saved for this client (dict may be partial/empty). Used by the app to
    decide whether the ROI sheet upload can be skipped this time."""
    entry = get_client(name) or {}
    return {k: entry[k] for k in ROI_ASSUMPTION_FIELDS if k in entry and entry[k] is not None}


def remove_client(name):
    """Delete one client's entry entirely (e.g. a duplicate, or a client
    that's no longer tracked here). Returns True if something was removed,
    False if that name wasn't in the store."""
    store = load_store()
    if name in store:
        del store[name]
        save_store(store)
        return True
    return False


def merge_dashboard_customers(customers):
    """
    customers: list of {"customerId": ..., "name": ...} from the live
    dashboard's customer-list API (see alphacore_api.fetch_customers()).

    Adds any brand-new client name found on the dashboard to the store
    (with customer_id filled in, traps_monitored left blank until you set
    it once in the app), and refreshes customer_id for existing ones in
    case it changed. Never removes a client that's already in the store,
    even if it didn't come back in this particular API response (so a
    temporary API hiccup can't silently wipe out saved trap counts).
    """
    store = load_store()
    added = []
    for c in customers:
        name = (c.get("name") or "").strip()
        cid = c.get("customerId")
        if not name or not cid:
            continue
        if name not in store:
            store[name] = {"customer_id": cid}
            added.append(name)
        else:
            store[name]["customer_id"] = cid
    save_store(store)
    return added  # list of newly-discovered client names this refresh