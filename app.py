"""
SteamGuard Report Generator — dashboard auto-fetch

Flow: pick a client + month -> alarms for that month are fetched
automatically from the alphacore.live dashboard (no file upload) -> upload
the client's ROI Excel -> enter the investment amount -> Generate. Produces
both a detailed Excel workbook and a short management-level PowerPoint.

Run with:  streamlit run app.py

ONE-TIME SETUP for the live dashboard connection:
    Windows CMD:        set ALPHACORE_PASS=your_password
    Windows PowerShell:  $env:ALPHACORE_PASS="your_password"
    Mac/Linux:           export ALPHACORE_PASS=your_password
(Same password alphacore_download.py already used.) Also set ALPHACORE_USER
if your dashboard login isn't trushi@alphacore.co.in.

Each client needs its dashboard Customer ID entered once (the app will ask
for it the first time you select that client) — after that it's remembered
in clients_config.json and every future report for that client auto-fetches.
"""

import os
import csv
import tempfile
import datetime
import openpyxl
import streamlit as st
import pandas as pd

from report_engine import build_report
from pptx_report import build_pptx_report
import clients_store
import alphacore_api

st.set_page_config(page_title="SteamGuard Report Generator", page_icon="♨️", layout="centered")

# ---------------- Look & feel ----------------
# Alphacore navy/blue brand palette (same colors used in the PPTX):
# NAVY = #1F4E78, BLUE = #2E75B6, LIGHT_BLUE = #DDEBF7
st.markdown("""
<style>
    /* Trim the big default gap above the title */
    .block-container { padding-top: 2rem; padding-bottom: 2rem; }

    /* Section "cards" — subtle bordered box around each numbered step */
    div[data-testid="stVerticalBlockBorderWrapper"] {
        background-color: #F7FAFD;
        border: 1px solid #DDEBF7;
        border-radius: 12px;
        padding: 0.5rem 0.5rem 1rem 0.5rem;
        margin-bottom: 1rem;
    }

    /* Primary button — brand navy, rounded, bit of shadow */
    button[kind="primary"] {
        background-color: #1F4E78 !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 0.75rem 1rem !important;
        font-weight: 700 !important;
        letter-spacing: 0.03em;
        box-shadow: 0 2px 6px rgba(31, 78, 120, 0.35);
        transition: background-color 0.15s ease-in-out;
    }
    button[kind="primary"]:hover {
        background-color: #2E75B6 !important;
    }

    /* Secondary download buttons — outlined blue */
    button[kind="secondary"] {
        border-radius: 10px !important;
        border-color: #2E75B6 !important;
        color: #1F4E78 !important;
        font-weight: 600 !important;
    }
    button[kind="secondary"]:hover {
        border-color: #1F4E78 !important;
        color: #1F4E78 !important;
        background-color: #DDEBF7 !important;
    }

    /* Section headers */
    h3 { color: #1F4E78 !important; }
</style>
""", unsafe_allow_html=True)

title_col1, title_col2 = st.columns([0.06, 0.94])
with title_col1:
    st.markdown("<div style='font-size:2.4rem;'>♨️</div>", unsafe_allow_html=True)
with title_col2:
    st.markdown("<h1 style='color:#1F4E78; margin-bottom:0;'>SteamGuard Report Generator</h1>", unsafe_allow_html=True)
st.caption("Select a client + month, confirm assumptions, and generate the monthly Excel + PowerPoint report.")
st.markdown("<hr style='margin-top:0.5rem; margin-bottom:1.5rem; border-color:#DDEBF7;'>", unsafe_allow_html=True)

FALLBACK_CLIENTS = [
    # Only clients where SteamGuard is actually deployed — demo/trial-only
    # plants (USV, Patanjali, Laxmi Organics, Sudarshan, Alivus, Shilpa,
    # Sun Pharma Mohali, Deepak Fertilizers, Torrent Pharma Dahej,
    # Clean Science, Ceat Nashik) are intentionally left out of this list.
    # MRF_Tyres is also excluded — it's on a monthly-rental arrangement,
    # which this tool no longer models (investment is one-time-total only).
    "JK_Tyres", "GSP", "BKT", "HIKAL", "Rallis",
    "SRF", "Marico_Jalgaon", "Apollo", "Alkem", "ipca", "Amneal_Matoda",
    "Amneal_Palli", "CEAT_Bhandup",
]
MONTHS = [datetime.date(2026, m, 1).strftime("%B %Y") for m in range(1, 13)]

DEFAULTS = {
    "pressure_bar": 14.0,
    "orifice_mm": 3.0,
    "cost_per_ton": 3200.0,
    "detect_minutes": 15.0,
    "investment_monthly": 40000.0,
    "traps_monitored": 25,
}

ROI_KEYWORDS = {
    "pressure_bar": ["pressure"],
    "orifice_mm": ["orifice"],
    "cost_per_ton": ["cost per ton", "steam cost", "per ton", "cost/ton"],
    "detect_minutes": ["detection time", "detection"],
}


@st.cache_data(ttl=3600, show_spinner=False)
def _cached_fetch_all_alarms(year, month):
    """Fetches the WHOLE tenant's alarms for one calendar month once, then
    Streamlit reuses that result (up to an hour) for every subsequent
    client report generated for that same month — instead of re-fetching
    every single client's report from scratch, which was the main reason
    report generation felt slow. Cache key is just (year, month); a fresh
    login is done under the hood on the first call for that month."""
    token = alphacore_api.login()
    return alphacore_api.fetch_all_alarms_for_month(token, year, month)


def extract_assumptions_from_roi(uploaded_file):
    """Best-effort scan of the uploaded ROI Excel (see report history for
    the two layouts this handles). Never trusted silently — results are
    always shown to the user with their source."""
    found = {}
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str) or not cell.value.strip():
                    continue
                label_text = cell.value.strip().lower()
                for key, kws in ROI_KEYWORDS.items():
                    if key in found:
                        continue
                    if any(kw in label_text for kw in kws):
                        for r_off in range(1, 6):
                            below = ws.cell(row=cell.row + r_off, column=cell.column)
                            if isinstance(below.value, (int, float)):
                                found[key] = float(below.value)
                                break
        for row in ws.iter_rows():
            label_cell = None
            for cell in row:
                if isinstance(cell.value, str) and cell.value.strip():
                    label_cell = cell
                    break
            if label_cell is None:
                continue
            label_text = label_cell.value.strip().lower()
            for key, kws in ROI_KEYWORDS.items():
                if key in found:
                    continue
                if any(kw in label_text for kw in kws):
                    for cell in row:
                        if isinstance(cell.value, (int, float)) and cell.column != label_cell.column:
                            found[key] = float(cell.value)
                            break
    return found


# ---------------- Client list ----------------
# Simple static dropdown, same as before — just the client names that are
# already known. (The live "refresh from dashboard" attempt was removed:
# that endpoint needs a permission on the alphacore.live account that isn't
# available right now, so it wasn't worth the extra button/complexity.)
step1 = st.container(border=True)
with step1:
    st.subheader("🏭 1. Client & month")

    CLIENTS = FALLBACK_CLIENTS

    col1, col2 = st.columns(2)
    with col1:
        client_choice = st.selectbox("Client", CLIENTS + ["Other (type below)"])
        client_name = st.text_input("Client name (if 'Other')", "") if client_choice.startswith("Other") else client_choice
    with col2:
        month_label = st.selectbox("Report month", MONTHS, index=6)  # default July

    target_year_month = datetime.datetime.strptime(month_label, "%B %Y")
    filter_year_month = (target_year_month.year, target_year_month.month)

    # Alarm data is always auto-fetched from the dashboard for the selected
    # client + month — no file upload needed. The one thing this needs that
    # can't be auto-discovered (the dashboard account doesn't have permission
    # to list all customers) is this client's dashboard Customer ID. Once
    # that's entered here for a client, it's saved and never needs re-entering.
    saved_client = clients_store.get_client(client_name) if client_name else None
    customer_id_known = bool((saved_client or {}).get("customer_id"))
    if client_name and not customer_id_known:
        st.info(
            f"🆕 First time fetching '{client_name}' — need its dashboard Customer ID once, then it's saved "
            f"forever for this client. Use the finder below — no terminal/CMD needed, works the same whether "
            f"you're running this locally or on the deployed website."
        )
        find_col1, find_col2 = st.columns([0.6, 0.4])
        with find_col1:
            expected_count_str = st.text_input(
                f"(Optional) alarm count you see for '{client_name}' on the dashboard, to auto-highlight the match",
                key=f"expected_{client_name}",
            )
        with find_col2:
            st.write("")
            st.write("")
            find_clicked = st.button(f"🔍 Find {client_name}'s Customer ID", key=f"find_{client_name}")

        if find_clicked:
            with st.spinner(f"Fetching {target_year_month.year}-{target_year_month.month:02d} alarms for the "
                             f"whole tenant, so every client's Customer ID can be shown..."):
                try:
                    all_month_alarms = _cached_fetch_all_alarms(target_year_month.year, target_year_month.month)
                    counts = {}
                    for a in all_month_alarms:
                        cid = alphacore_api.alarm_customer_id(a)
                        counts[cid] = counts.get(cid, 0) + 1
                    ranked = sorted(counts.items(), key=lambda x: -x[1])
                    st.session_state[f"cid_options_{client_name}"] = ranked
                except (alphacore_api.AlphacoreAuthError, alphacore_api.AlphacoreApiError) as e:
                    st.error(f"Couldn't fetch alarms to find the Customer ID: {e}")

        ranked = st.session_state.get(f"cid_options_{client_name}")
        if ranked:
            try:
                expected_count = int(expected_count_str) if expected_count_str.strip() else None
            except ValueError:
                expected_count = None

            def _label(item):
                cid, count = item
                marker = ""
                if expected_count is not None and abs(count - expected_count) <= max(2, expected_count * 0.02):
                    marker = "  ⭐ closest match to the count you gave"
                return f"{count} alarms — {cid}{marker}"

            options = ["(select one)"] + [_label(item) for item in ranked]
            default_idx = 0
            if expected_count is not None:
                for i, (cid, count) in enumerate(ranked, start=1):
                    if abs(count - expected_count) <= max(2, expected_count * 0.02):
                        default_idx = i
                        break
            chosen_label = st.selectbox(
                f"Found {len(ranked)} distinct customer IDs for {target_year_month.year}-{target_year_month.month:02d} "
                f"— pick the one matching '{client_name}'",
                options, index=default_idx, key=f"cid_select_{client_name}",
            )
            if chosen_label != "(select one)":
                chosen_idx = options.index(chosen_label) - 1
                chosen_cid = ranked[chosen_idx][0]
                if st.button(f"✅ Save this as {client_name}'s Customer ID", key=f"save_cid_{client_name}"):
                    clients_store.upsert_client(client_name, customer_id=chosen_cid)
                    st.success(f"✅ Saved — '{client_name}' will auto-fetch from now on.")
                    saved_client = clients_store.get_client(client_name)
                    customer_id_known = True
                    del st.session_state[f"cid_options_{client_name}"]
                    st.rerun()

        with st.expander("Or paste a Customer ID directly, if you already have it"):
            new_customer_id = st.text_input(
                f"Dashboard Customer ID for {client_name}", key=f"cid_manual_{client_name}"
            )
            if new_customer_id.strip():
                clients_store.upsert_client(client_name, customer_id=new_customer_id.strip())
                st.success(f"✅ Saved — '{client_name}' will auto-fetch from now on.")
                saved_client = clients_store.get_client(client_name)
                customer_id_known = True

    # Safety net for today: if the live auto-fetch isn't returning the right
    # data yet (wrong Customer ID, API quirk, etc.), you can tick this to fall
    # back to uploading the alarm export by hand instead — so a demo/report
    # never has to wait on the dashboard connection being fully sorted out.
    use_manual_upload = st.checkbox("⚠️ Having trouble with auto-fetch? Upload the alarm file manually instead")
    alarm_file = None
    if use_manual_upload:
        alarm_file = st.file_uploader("This month's alarm export (CSV or Excel)", type=["csv", "xlsx", "xls"])

# ---------------- ROI + Investment ----------------
step2 = st.container(border=True)
with step2:
    st.subheader("💰 2. ROI sheet & investment")

    # ROI assumptions (pressure/orifice/cost-per-ton/detection-time) are
    # fixed per client, same as Investment and Traps Monitored — once
    # they've been read from a client's ROI sheet and a report has been
    # generated, they're saved and reused automatically every month after
    # that, so the ROI sheet only needs to be uploaded again if something
    # about it actually changes.
    saved_roi = clients_store.get_saved_roi_assumptions(client_name) if client_name else {}
    has_saved_roi = bool(saved_roi)

    if has_saved_roi:
        st.success(
            f"✅ Using saved ROI assumptions for **{client_name}** from a previous report "
            f"(no need to re-upload the ROI sheet). Upload it below only if the numbers changed."
        )
        roi_file = st.file_uploader(
            "Client's ROI Excel (optional — already saved for this client)", type=["xlsx", "xls"]
        )
    else:
        roi_file = st.file_uploader(
            "Client's ROI Excel (first time for this client — will be saved automatically)",
            type=["xlsx", "xls"],
        )

    # ---------------- Investment + Traps monitored ----------------
    col_inv, col_traps = st.columns(2)
    with col_inv:
        # Every client's investment is a one-time total cost for the whole
        # SteamGuard engagement (e.g. GSP's Rs 5,64,000) — never a recurring
        # monthly rental. ROI is shown as an estimated payback period in months
        # rather than a misleading "x" multiple of one month's avoidable loss.
        investment_type = "one_time"
        # Fixed per client (like Traps Monitored below) — auto-filled from
        # clients_config.json once it's been entered for this client, so it
        # doesn't need re-typing every month. Still editable if it ever changes.
        default_investment = (saved_client or {}).get("investment", DEFAULTS["investment_monthly"])
        investment = st.number_input(
            "💵 Investment — one-time total (Rs)", value=float(default_investment), step=1000.0,
            help="The client's total one-time SteamGuard investment (whole engagement, not per month). "
                 "Saved automatically once you generate a report for this client — auto-fills next time.",
        )
    with col_traps:
        # Traps monitored is the client's TRUE installed base — fixed per
        # client, not per month. Auto-filled from clients_config.json (saved
        # once per client) so it doesn't need re-typing every report.
        default_traps = (saved_client or {}).get("traps_monitored", DEFAULTS["traps_monitored"])
        traps_override = st.number_input(
            "🔧 Traps monitored (true installed base)",
            value=int(default_traps), step=1, min_value=1,
            help="The client's actual total installed trap count (e.g. 25 for GSP) — fixed per client, "
                 "not per month. Saved automatically once you generate a report for this client.",
        )

    detected = {}
    if roi_file is not None:
        try:
            detected = extract_assumptions_from_roi(roi_file)
        except Exception as e:
            st.warning(f"Couldn't auto-read the ROI sheet ({e}) — using defaults below. "
                       "You can still set values manually in Advanced.")

    # Precedence: freshly-uploaded ROI sheet (detected) > previously-saved
    # values for this client (saved_roi) > hardcoded defaults.
    assumption_values = {**DEFAULTS, **saved_roi, **detected}

    if roi_file is not None:
        st.markdown("**📋 Assumptions detected from the ROI sheet:**")
        rows = []
        for key, label in [
            ("pressure_bar", "Steam header pressure (bar)"),
            ("orifice_mm", "Orifice diameter (mm)"),
            ("cost_per_ton", "Cost per ton of steam (Rs)"),
        ]:
            source = "from ROI sheet" if key in detected else (
                "saved from before" if key in saved_roi else "default (not found in sheet)")
            rows.append({"Assumption": label, "Value": assumption_values[key], "Source": source})
        st.table(pd.DataFrame(rows))
        detectable_keys = ["pressure_bar", "orifice_mm", "cost_per_ton"]
        if sum(1 for k in detectable_keys if k in detected) < len(detectable_keys):
            st.caption("Some values weren't found in the ROI sheet and are using saved/default values — "
                       "open Advanced below to correct them if needed.")
    elif has_saved_roi:
        st.caption("📋 Using the saved assumptions shown in Advanced below (from this client's last ROI sheet).")

    with st.expander("⚙️ Advanced: edit assumptions manually"):
        a1, a2 = st.columns(2)
        with a1:
            pressure = st.number_input("Steam header pressure (bar)", value=assumption_values["pressure_bar"], step=0.5)
            orifice = st.number_input("Orifice diameter (mm)", value=assumption_values["orifice_mm"], step=0.5)
        with a2:
            cost_per_ton = st.number_input("Cost per ton of steam (Rs)", value=assumption_values["cost_per_ton"], step=100.0)
        # SteamGuard detection time is used internally (Loss_Projection sheet's
        # residual-loss calc) but is intentionally not shown anywhere in the UI.
        detect_min = assumption_values["detect_minutes"]

st.write("")
generate = st.button("⚡ GENERATE REPORT", type="primary", use_container_width=True)


def fetch_alarms_to_csv(client_name, year, month, tmpdir):
    """Live dashboard fetch path — the only alarm-data source now. Requires
    clients_config.json to already know this client's customer_id (entered
    once, right above, the first time this client is used)."""
    entry = clients_store.get_client(client_name)
    if not entry or "customer_id" not in entry:
        raise ValueError(
            f"Don't have a dashboard Customer ID for '{client_name}' yet — enter it in the box above first."
        )
    all_month_alarms = _cached_fetch_all_alarms(year, month)
    alarms = alphacore_api.filter_alarms_for_customer(all_month_alarms, entry["customer_id"])
    if not alarms:
        raise ValueError(
            f"The dashboard returned 0 alarms for {client_name} in {year}-{month:02d} — the saved "
            f"Customer ID for '{client_name}' is almost certainly wrong (this happened with GSP too). "
            f"Fix it by running, in a terminal:  "
            f"python find_customer_id.py {client_name} {year} {month}  "
            f"— it lists every client's real alarm counts for that month so you can pick the correct one "
            f"and it saves it automatically. Then just click GENERATE REPORT again."
        )
    rows = [alphacore_api.flatten(a) for a in alarms]
    all_keys = []
    for r in rows:
        for k in r:
            if k not in all_keys:
                all_keys.append(k)
    path = os.path.join(tmpdir, "alarms.csv")
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=all_keys)
        writer.writeheader()
        writer.writerows(rows)
    return path


def save_upload_to_csv(uploaded_file, tmpdir):
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        path = os.path.join(tmpdir, "alarms.csv")
        with open(path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        return path
    else:
        df = pd.read_excel(uploaded_file)
        path = os.path.join(tmpdir, "alarms.csv")
        df.to_csv(path, index=False)
        return path


if generate:
    if not client_name:
        st.error("Please select or type a client name.")
    elif use_manual_upload and alarm_file is None:
        st.error("Please upload the alarm export file for this month.")
    elif not use_manual_upload and not customer_id_known:
        st.error(f"Enter '{client_name}'s dashboard Customer ID above first, so its alarms can be auto-fetched.")
    else:
        spinner_msg = ("Building report from the uploaded file..." if use_manual_upload else
                        f"Fetching {month_label} alarms from the dashboard and building the report...")
        with st.spinner(spinner_msg):
            try:
                with tempfile.TemporaryDirectory() as tmpdir:
                    if use_manual_upload:
                        csv_path = save_upload_to_csv(alarm_file, tmpdir)
                    else:
                        csv_path = fetch_alarms_to_csv(client_name, filter_year_month[0], filter_year_month[1], tmpdir)

                    out_path = os.path.join(tmpdir, f"{client_name}_{month_label.replace(' ', '_')}_SteamGuard_Report.xlsx")
                    pptx_path = os.path.join(tmpdir, f"{client_name}_{month_label.replace(' ', '_')}_Management_Report.pptx")

                    final_assumptions = {
                        "pressure_bar": pressure,
                        "orifice_mm": orifice,
                        "cost_per_ton": cost_per_ton,
                        "detect_minutes": detect_min,
                        "investment_monthly": investment,
                        "investment_type": investment_type,
                        "traps_monitored": int(traps_override),
                    }

                    saved_path, summary = build_report(
                        client_name, csv_path, out_path, final_assumptions,
                        filter_year_month=filter_year_month,
                    )
                    build_pptx_report(
                        summary, month_label, pptx_path,
                        investment_confirmed=(investment != DEFAULTS["investment_monthly"]),
                    )

                    # Persist this client's true trap count, investment, AND
                    # ROI assumptions for next time — all fixed per client,
                    # not per month, so once entered/uploaded correctly they
                    # shouldn't need re-typing/re-uploading on future reports.
                    clients_store.upsert_client(
                        client_name, traps_monitored=int(traps_override), investment=investment,
                        pressure_bar=pressure, orifice_mm=orifice,
                        cost_per_ton=cost_per_ton, detect_minutes=detect_min,
                    )

                    with open(saved_path, "rb") as f:
                        report_bytes = f.read()
                    with open(pptx_path, "rb") as f:
                        pptx_bytes = f.read()

                # Stashed in session_state rather than shown directly here:
                # clicking EITHER download button below triggers a Streamlit
                # rerun, and on that rerun `generate` (from st.button) is
                # False again — so this whole block wouldn't re-run and the
                # OTHER download button would vanish. Session state survives
                # the rerun, so both buttons (and the success message) are
                # rendered from it below, outside this "if generate" block.
                st.session_state["last_report"] = {
                    "client_name": client_name,
                    "report_bytes": report_bytes,
                    "report_filename": os.path.basename(out_path),
                    "pptx_bytes": pptx_bytes,
                    "pptx_filename": os.path.basename(pptx_path),
                    "summary": summary,
                    "month_label": month_label,
                }

            except ValueError as e:
                st.error(str(e))
            except (alphacore_api.AlphacoreAuthError, alphacore_api.AlphacoreApiError) as e:
                st.error(f"Dashboard fetch failed: {e}")
            except Exception as e:
                st.error(f"Something went wrong: {e}")

# Rendered from session_state (see comment above) so both download buttons
# stay visible no matter which one — or neither — was just clicked.
if st.session_state.get("last_report"):
    r = st.session_state["last_report"]
    summary = r["summary"]

    results_box = st.container(border=True)
    with results_box:
        st.markdown("#### 📊 Report ready")
        st.success(f"✅ Last successfully generated report — **{r.get('client_name', '')}**, {r['month_label']}: "
                   f"{summary['total_events']} events "
                   f"(out of {summary['total_rows_in_file']} rows in the source data), "
                   f"{summary['traps_flagged']} traps flagged, "
                   f"{summary['date_range'][0]} to {summary['date_range'][1]}.")
        if r.get('client_name') != client_name or r.get('month_label') != month_label:
            st.caption("⚠️ This is from your last successful run, not your current client/month selection above — "
                       "if the current attempt failed (see red box, if any), this older report is still available "
                       "to download below, but it is NOT for what you just tried to generate.")

        if summary.get('other_type_count', 0) > 0:
            st.warning(f"{summary['other_type_count']} events had a 'type' value we didn't recognize "
                       f"and were bucketed as Other (not counted as Leak or Block): "
                       f"{summary['unrecognized_types']}. These won't show up in Leak/Block totals — "
                       f"tell me these type values and I'll add them to the mapping.")

        dl1, dl2 = st.columns(2)
        with dl1:
            st.download_button(
                "📥 Download Excel Report",
                data=r["report_bytes"],
                file_name=r["report_filename"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_excel",
            )
        with dl2:
            st.download_button(
                "📥 Download Management PPTX",
                data=r["pptx_bytes"],
                file_name=r["pptx_filename"],
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
                key="dl_pptx",
            )
        st.caption("💡 Open the Excel file and press Ctrl+Alt+F9 once if any numbers look blank — "
                   "Excel recalculates formulas automatically on open, but this forces it.")