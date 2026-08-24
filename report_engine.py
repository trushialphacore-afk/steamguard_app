"""
report_engine.py
-----------------
Core report-building logic, refactored from the original build_client.py
into a single callable function so it can be driven by the Streamlit app
(or still run standalone from the command line for testing).

Expected input CSV format: the same columns produced by our
alphacore_download.py export (createdTime, startTs, endTs, type,
originatorName, severity, status, ...). If you export alarms straight from
the alphacore.live "Alarms" table instead, the column names will be
different — see convert_dashboard_export() in app.py for that path.
"""

import openpyxl, datetime, csv, re
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint

# Brand palette (same colors used in the PPTX — NAVY/BLUE/LIGHT_BLUE) plus a
# warm accent for "Block" events, used to make the Management_View charts
# look consistent and less default-Excel-blue.
CHART_NAVY = '1F4E78'
CHART_BLUE = '2E75B6'
CHART_LIGHT_BLUE = 'DDEBF7'
CHART_MID_BLUE = '9DC3E6'
CHART_ACCENT_RED = 'C0504D'
import pandas as pd

FONT_NAME = 'Arial'
BLUE = Font(name=FONT_NAME, color='0000FF')
BLACK = Font(name=FONT_NAME, color='000000')
GREEN = Font(name=FONT_NAME, color='008000')
HDR_FILL = PatternFill('solid', fgColor='1F4E78')
HDR_FONT = Font(name=FONT_NAME, bold=True, color='FFFFFF')
TITLE_FONT = Font(name=FONT_NAME, bold=True, size=14, color='1F4E78')
SUB_FONT = Font(name=FONT_NAME, italic=True, size=10, color='595959')
SECTION_FONT = Font(name=FONT_NAME, bold=True, size=12, color='FFFFFF')
SECTION_FILL = PatternFill('solid', fgColor='2E75B6')
thin = Side(style='thin', color='BFBFBF')
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

# Keys are normalized: lowercased, with spaces/underscores/hyphens stripped.
# Matching goes through _normalize_type() below, so both 'trap_fullleak' and
# 'Trap Full Leak' (say) land on the same 'trapfullleak' key. Add new raw
# strings here (normalized form) if a future file still lands in "Other".
TYPE_MAP = {
    'trapblocked': ('Block', 'Block'),
    'blocked': ('Block', 'Block'),
    'trapfullleak': ('Leak', 'Full'),
    'fullleak': ('Leak', 'Full'),
    'steamtrappartialleak': ('Leak', 'Partial'),
    'partialleak': ('Leak', 'Partial'),
    'trappartialleak': ('Leak', 'Partial'),
}


def _normalize_type(raw):
    return str(raw or '').strip().lower().replace('_', '').replace(' ', '').replace('-', '')


def map_type(raw):
    return TYPE_MAP.get(_normalize_type(raw), ('Other', 'Other'))

REQUIRED_COLUMNS = ['createdTime', 'startTs', 'endTs', 'type', 'severity', 'status']

# Columns seen when exporting straight from the alphacore.live "Alarms" table
# UI (Export button), instead of via alphacore_download.py's API export.
DASHBOARD_COLUMNS = ['Created time', 'Type', 'Originator', 'Duration (hrs)', 'Severity', 'Status']

IST_OFFSET = datetime.timedelta(hours=5, minutes=30)


def parse_duration_to_hours(s):
    """
    Parse the dashboard export's free-text 'Duration (hrs)' string into a
    float number of hours. The dashboard's own "Alarms" table UI export does
    NOT put a number in this column — it puts a human-readable string like
    '1 hr ', '6 hrs 19 min 55 sec ', '22 day 16 hrs 12 min 46 sec ', or an
    empty string for zero/ongoing durations.

    IMPORTANT: do not swap this for pd.to_numeric() — that silently coerces
    every one of these strings to 0/NaN, which was the confirmed root cause
    of a "0 leak hours found" bug (real leak durations were all being read
    as zero because they were text, not numbers).
    """
    if s is None:
        return 0.0
    s = str(s).strip()
    if not s or s.lower() == 'nan':
        return 0.0
    if s.startswith('-'):
        # The dashboard export uses a fixed sentinel '-1 hrs -1 min -1 sec '
        # for events whose duration wasn't tracked/measured (seen only on
        # Cleared/Unacknowledged rows) — it is a placeholder, not a real
        # negative duration, so treat it as unknown/zero rather than
        # parsing the digits and getting a spurious positive value.
        return 0.0
    total_hours = 0.0
    m = re.search(r'(\d+)\s*day', s)
    if m:
        total_hours += int(m.group(1)) * 24
    m = re.search(r'(\d+)\s*hr', s)
    if m:
        total_hours += int(m.group(1))
    m = re.search(r'(\d+)\s*min', s)
    if m:
        total_hours += int(m.group(1)) / 60
    m = re.search(r'(\d+)\s*sec', s)
    if m:
        total_hours += int(m.group(1)) / 3600
    return total_hours


def normalize_alarm_csv(input_path, out_path=None):
    """
    Accepts either:
      (a) the alphacore_download.py API export (createdTime/startTs/endTs/type/...), or
      (b) a raw export from the dashboard's "Alarms" table UI (Created time/
          Duration (hrs)/Originator/Type/Severity/Status).
    Returns a path to a CSV in format (a), which build_report() expects.
    Raises ValueError with a clear message if neither format matches.
    """
    with open(input_path, newline='', encoding='utf-8') as f:
        header = next(csv.reader(f))

    if all(c in header for c in REQUIRED_COLUMNS):
        return input_path  # already in the expected format

    if all(c in header for c in DASHBOARD_COLUMNS):
        df = pd.read_csv(input_path)
        created = pd.to_datetime(df['Created time'])  # naive, already IST wall-clock time
        created_ms = ((created - IST_OFFSET) - pd.Timestamp("1970-01-01")) // pd.Timedelta('1ms')
        duration_ms = (df['Duration (hrs)'].apply(parse_duration_to_hours) * 3600000).astype('int64')

        out = pd.DataFrame({
            'createdTime': created_ms,
            'startTs': created_ms,
            'endTs': created_ms + duration_ms,
            'type': df['Type'],
            'originatorName': df['Originator'],
            'severity': df['Severity'],
            'status': df['Status'],
        })
        out_path = out_path or (input_path + '.normalized.csv')
        out.to_csv(out_path, index=False)
        return out_path

    raise ValueError(
        "This file doesn't match either expected format.\n\n"
        f"Expected either the API export columns: {REQUIRED_COLUMNS}\n"
        f"or the dashboard 'Alarms' table export columns: {DASHBOARD_COLUMNS}\n\n"
        f"Found columns: {header}"
    )


def style_header(ws, row, ncols, start_col=1):
    for c in range(start_col, start_col + ncols):
        cell = ws.cell(row=row, column=c)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = BORDER


def add_table(ws, name, ref, style='TableStyleMedium2'):
    tab = Table(displayName=name, ref=ref)
    tab.tableStyleInfo = TableStyleInfo(name=style, showRowStripes=True, showFirstColumn=False)
    ws.add_table(tab)


def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def validate_csv_columns(csv_path):
    """Quick sanity check so a wrong-format upload fails loudly, not silently."""
    with open(csv_path, newline='', encoding='utf-8') as f:
        header = next(csv.reader(f))
    missing = [c for c in REQUIRED_COLUMNS if c not in header]
    if missing:
        raise ValueError(
            f"This file doesn't look like an alphacore alarm export — missing columns: {missing}. "
            f"Found columns: {header[:12]}..."
        )


def build_report(client_name, csv_path, out_path, assumptions=None, filter_year_month=None):
    """
    Build the 8-sheet SteamGuard workbook for one client/month.

    assumptions: optional dict, any subset of:
        pressure_bar        (default 14)
        orifice_mm           (default 3)
        cost_per_ton         (default 3200)
        full_severity        (default 1.0)
        partial_severity     (default 0.5)
        detect_minutes       (default 15)
        investment_monthly   (default 40000)   <- update from the client's ROI sheet
        traps_monitored      (default: auto-counted from alarm data)

    filter_year_month: optional (year, month) tuple. If given, only alarm
        rows whose Created time falls in that calendar month are used —
        everything else in the uploaded file is dropped before building the
        report (a wider file covering several months is fine to upload).
    """
    assumptions = assumptions or {}
    a_pressure = assumptions.get('pressure_bar', 14)
    a_orifice = assumptions.get('orifice_mm', 3)
    a_cost = assumptions.get('cost_per_ton', 3200)
    a_full_sev = assumptions.get('full_severity', 1)
    a_partial_sev = assumptions.get('partial_severity', 0.5)
    a_detect_min = assumptions.get('detect_minutes', 15)
    a_investment = assumptions.get('investment_monthly', 40000)
    # 'monthly'  -> a_investment is a recurring monthly rental cost; ROI is
    #               shown as a multiple (x) of this month's avoidable loss.
    # 'one_time' -> a_investment is a single upfront total cost (e.g. GSP's
    #               Rs 5,64,000 one-time investment); showing that as an "x
    #               multiple" of ONE month's avoidable loss would be
    #               meaningless (the number would look tiny and alarming),
    #               so instead we show an estimated payback period in months
    #               = investment / this month's avoidable monthly loss.
    a_investment_type = assumptions.get('investment_type', 'one_time')
    a_traps_override = assumptions.get('traps_monitored')  # None -> auto count

    csv_path = normalize_alarm_csv(csv_path)
    validate_csv_columns(csv_path)

    with open(csv_path, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        raw_rows = list(reader)

    total_rows_in_file = len(raw_rows)
    unrecognized_types = set()

    records = []
    for r in raw_rows:
        typ = r.get('type', '')
        category, sev_type = map_type(typ)
        if category == 'Other' and typ:
            unrecognized_types.add(typ)
        try:
            created_ms = int(r['createdTime'])
            start_ms = int(r['startTs']) if r['startTs'] else created_ms
            end_ms = int(r['endTs']) if r['endTs'] else start_ms
        except (ValueError, KeyError):
            continue
        dur_hours = max(0.0, (end_ms - start_ms) / 3600000.0)
        created_dt = datetime.datetime.utcfromtimestamp(created_ms / 1000) + datetime.timedelta(hours=5, minutes=30)
        if filter_year_month is not None:
            if (created_dt.year, created_dt.month) != tuple(filter_year_month):
                continue
        records.append({
            'created': created_dt,
            'end': datetime.datetime.utcfromtimestamp(end_ms / 1000) + datetime.timedelta(hours=5, minutes=30),
            'type': typ,
            'category': category,
            'sev_type': sev_type,
            'trap_id': r.get('originatorName', '').strip() or r.get('originator_id', ''),
            'severity': r.get('severity', ''),
            'status': r.get('status', ''),
            'duration_hours': dur_hours,
        })

    if not records:
        if filter_year_month is not None:
            raise ValueError(
                f"No alarm rows found for {filter_year_month[0]}-{filter_year_month[1]:02d} in this file — "
                f"check that the selected month actually has data in the upload (file has {total_rows_in_file} rows total)."
            )
        raise ValueError("No usable alarm rows found in this file after parsing — check the date range and column values.")

    # De-duplicate: the dashboard's alarm export can carry two rows for the
    # SAME physical alarm — an "Acknowledged" snapshot and an
    # "Unacknowledged" snapshot of it (same trap, same created time, same
    # type), almost always with the real duration on only one of the two and
    # 0/blank on the other. Counting both would double-count that event in
    # Total/Leak/Block Events and in Traps Flagged. Keep one row per
    # (trap_id, created time, type), using the longest duration seen for it.
    deduped = {}
    duplicate_rows_removed = 0
    for rec in records:
        key = (rec['trap_id'], rec['created'], rec['type'])
        if key in deduped:
            duplicate_rows_removed += 1
            if rec['duration_hours'] > deduped[key]['duration_hours']:
                deduped[key] = rec
        else:
            deduped[key] = rec
    records = list(deduped.values())

    records.sort(key=lambda x: (x['trap_id'], x['created']))
    n = len(records)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ---------------- Assumptions_Flags ----------------
    af = wb.create_sheet('Assumptions_Flags')
    set_widths(af, [3, 42, 16, 62])
    af['B2'] = f'{client_name.replace("_", " ")} - SteamGuard Monthly Report - Assumptions & Inputs'
    af['B2'].font = TITLE_FONT
    af['B3'] = 'Change the blue cells to re-run every formula in this workbook. Data source: alphacore.live alarm export (ThingsBoard-based), 1 calendar month.'
    af['B3'].font = SUB_FONT
    af.merge_cells('B3:D3')

    r = 5
    af.cell(row=r, column=2, value='1.  CALCULATION INPUTS  (edit blue cells)').font = SECTION_FONT
    for c in range(2, 5):
        af.cell(row=r, column=c).fill = SECTION_FILL
    r += 1
    hdr_row = r
    for c, t in zip([2, 3, 4], ['Input', 'Value', 'Source / Note']):
        af.cell(row=r, column=c, value=t)
    style_header(af, r, 3, start_col=2)
    r += 1

    traps_monitored_default = a_traps_override if a_traps_override is not None else len(set(rec['trap_id'] for rec in records))
    traps_note = ('Manually set from the app.' if a_traps_override is not None else
                  'Auto-counted from distinct traps that raised an alarm this month - update if the true installed base differs.')

    investment_label = ('Investment - one-time total (Rs)' if a_investment_type == 'one_time'
                         else 'Investment - monthly rental (Rs)')
    investment_note = ('One-time total SteamGuard cost - payback period below is estimated from this '
                        'month\'s avoidable-loss run-rate.' if a_investment_type == 'one_time' else
                        'From the app / ROI sheet - confirm against the client\'s actual contracted SteamGuard cost.')

    inputs = [
        ('Steam header pressure (bar)', a_pressure, 'From the app / ROI sheet - default Napier-model assumption if unconfirmed.'),
        ('Orifice diameter (mm)', a_orifice, 'From the app / ROI sheet - standard trap orifice if unconfirmed.'),
        ('Cost per Ton of steam (Rs)', a_cost, 'From the app / ROI sheet - update with the client\'s actual steam cost if known.'),
        ('Full-leak severity (% of orifice flow)', a_full_sev, 'Full blow-through = 100%.'),
        ('Partial-leak severity (% of full leak)', a_partial_sev, 'Consistent with the Alphacore steam-loss model used in prior case studies (50%).'),
        ('SteamGuard detection time (minutes)', a_detect_min, 'Per SteamGuard spec: real-time alert, resolved within ~15 minutes.'),
        (investment_label, a_investment, investment_note),
        ('Traps monitored (installed base)', traps_monitored_default, traps_note),
    ]
    for label, val, note in inputs:
        af.cell(row=r, column=2, value=label).font = BLACK
        vc = af.cell(row=r, column=3, value=val)
        vc.font = BLUE
        if isinstance(val, float) and val < 1:
            vc.number_format = '0%'
        nc = af.cell(row=r, column=4, value=note)
        nc.alignment = Alignment(wrap_text=True, vertical='top')
        af.row_dimensions[r].height = 30
        for c in range(2, 5):
            af.cell(row=r, column=c).border = BORDER
        r += 1

    INPUT_ROWS = {}
    labels_order = [x[0] for x in inputs]
    start_input_row = hdr_row + 1
    for i, lbl in enumerate(labels_order):
        INPUT_ROWS[lbl] = start_input_row + i

    PRESSURE_ROW = INPUT_ROWS['Steam header pressure (bar)']
    ORIFICE_ROW = INPUT_ROWS['Orifice diameter (mm)']
    COST_ROW = INPUT_ROWS['Cost per Ton of steam (Rs)']
    FULL_SEV_ROW = INPUT_ROWS['Full-leak severity (% of orifice flow)']
    PARTIAL_SEV_ROW = INPUT_ROWS['Partial-leak severity (% of full leak)']
    DETECT_MIN_ROW = INPUT_ROWS['SteamGuard detection time (minutes)']
    INVESTMENT_ROW = INPUT_ROWS[investment_label]
    TRAPS_MONITORED_ROW = INPUT_ROWS['Traps monitored (installed base)']

    r += 1
    af.cell(row=r, column=2, value='2.  COMPUTED STEAM-LOSS RATES  (Napier orifice-flow formula)').font = SECTION_FONT
    for c in range(2, 5):
        af.cell(row=r, column=c).fill = SECTION_FILL
    r += 1
    for c, t in zip([2, 3, 4], ['Rate', 'Value (Ton/hr)', 'Formula']):
        af.cell(row=r, column=c, value=t)
    style_header(af, r, 3, start_col=2)
    r += 1
    FULL_RATE_ROW = r
    af.cell(row=r, column=2, value='Full blow-through leak rate (Ton/hr)')
    af.cell(row=r, column=3, value=f'=0.24725*($C${PRESSURE_ROW}+1)*($C${ORIFICE_ROW}^2)*$C${FULL_SEV_ROW}/1000').font = BLACK
    af.cell(row=r, column=4, value='0.24725 x (Pressure+1) x Orifice^2 x Severity / 1000')
    r += 1
    PARTIAL_RATE_ROW = r
    af.cell(row=r, column=2, value='Partial leak rate (Ton/hr)')
    af.cell(row=r, column=3, value=f'=$C${FULL_RATE_ROW}*$C${PARTIAL_SEV_ROW}').font = BLACK
    af.cell(row=r, column=4, value='Full-leak rate x partial severity %')
    r += 1
    for rr in [FULL_RATE_ROW, PARTIAL_RATE_ROW]:
        af.cell(row=rr, column=3).number_format = '0.00000'
        for c in range(2, 5):
            af.cell(row=rr, column=c).border = BORDER

    af.freeze_panes = 'B6'

    # ---------------- Alarm_Data ----------------
    ad = wb.create_sheet('Alarm_Data')
    headers = ['Alarm_ID', 'Created_Time', 'End_Time', 'Trap_ID', 'Fault_Type_Raw',
               'Fault_Category', 'Fault_Severity', 'Duration_Hours', 'Alarm_Severity',
               'Status', 'Event_Date']
    for c, h in enumerate(headers, start=1):
        ad.cell(row=1, column=c, value=h)
    style_header(ad, 1, len(headers))

    for i, rec in enumerate(records):
        row = i + 2
        ad.cell(row=row, column=1, value=i + 1)
        c2 = ad.cell(row=row, column=2, value=rec['created']); c2.number_format = 'yyyy-mm-dd hh:mm:ss'
        c3 = ad.cell(row=row, column=3, value=rec['end']); c3.number_format = 'yyyy-mm-dd hh:mm:ss'
        ad.cell(row=row, column=4, value=rec['trap_id'])
        ad.cell(row=row, column=5, value=rec['type'])
        ad.cell(row=row, column=6, value=rec['category'])
        ad.cell(row=row, column=7, value=rec['sev_type'])
        dc = ad.cell(row=row, column=8, value=round(rec['duration_hours'], 4)); dc.number_format = '0.0000'
        ad.cell(row=row, column=9, value=rec['severity'])
        ad.cell(row=row, column=10, value=rec['status'])
        ec = ad.cell(row=row, column=11, value=rec['created'].date()); ec.number_format = 'yyyy-mm-dd'

    last_row = n + 1
    add_table(ad, 'tbl_AlarmData', f'A1:K{last_row}')
    set_widths(ad, [10, 19, 19, 34, 20, 14, 12, 14, 13, 14, 12])
    ad.freeze_panes = 'A2'
    AD = 'tbl_AlarmData'

    # ---------------- Trap_Analysis ----------------
    trap_ids = sorted(set(rec['trap_id'] for rec in records))
    ta = wb.create_sheet('Trap_Analysis')
    ta_headers = ['Trap_ID', 'Total_Events', 'Leak_Events', 'Block_Events', 'Full_Leak_Events',
                  'Partial_Leak_Events', 'Total_Leak_Hours', 'Total_Block_Hours',
                  'Avg_Event_Duration_Hours', 'First_Alarm_Date', 'Last_Alarm_Date', 'Priority']
    for c, h in enumerate(ta_headers, start=1):
        ta.cell(row=1, column=c, value=h)
    style_header(ta, 1, len(ta_headers))

    for i, tid in enumerate(trap_ids):
        row = i + 2
        ta.cell(row=row, column=1, value=tid)
        ta.cell(row=row, column=2, value=f'=COUNTIF({AD}[Trap_ID],$A{row})')
        ta.cell(row=row, column=3, value=f'=COUNTIFS({AD}[Trap_ID],$A{row},{AD}[Fault_Category],"Leak")')
        ta.cell(row=row, column=4, value=f'=COUNTIFS({AD}[Trap_ID],$A{row},{AD}[Fault_Category],"Block")')
        ta.cell(row=row, column=5, value=f'=COUNTIFS({AD}[Trap_ID],$A{row},{AD}[Fault_Severity],"Full")')
        ta.cell(row=row, column=6, value=f'=COUNTIFS({AD}[Trap_ID],$A{row},{AD}[Fault_Severity],"Partial")')
        c7 = ta.cell(row=row, column=7, value=f'=SUMIFS({AD}[Duration_Hours],{AD}[Trap_ID],$A{row},{AD}[Fault_Category],"Leak")'); c7.number_format = '0.0'
        c8 = ta.cell(row=row, column=8, value=f'=SUMIFS({AD}[Duration_Hours],{AD}[Trap_ID],$A{row},{AD}[Fault_Category],"Block")'); c8.number_format = '0.0'
        c9 = ta.cell(row=row, column=9, value=f'=IFERROR((G{row}+H{row})/B{row},0)'); c9.number_format = '0.00'
        c10 = ta.cell(row=row, column=10, value=f'=_xlfn.MINIFS({AD}[Created_Time],{AD}[Trap_ID],$A{row})'); c10.number_format = 'yyyy-mm-dd hh:mm'
        c11 = ta.cell(row=row, column=11, value=f'=_xlfn.MAXIFS({AD}[Created_Time],{AD}[Trap_ID],$A{row})'); c11.number_format = 'yyyy-mm-dd hh:mm'
        ta.cell(row=row, column=12, value=f'=IF(G{row}>=100,"Critical",IF(G{row}>=40,"High",IF(G{row}>0,"Medium",IF(H{row}>0,"Low - Block Only","Healthy"))))')

    ta_last = len(trap_ids) + 1
    ta.cell(row=1, column=13, value='Leak_Hours_Rank')
    style_header(ta, 1, 1, start_col=13)
    for i in range(len(trap_ids)):
        row = i + 2
        ta.cell(row=row, column=13, value=f'=RANK(G{row},$G$2:$G${ta_last})+COUNTIF($G$2:G{row},G{row})-1')

    add_table(ta, 'tbl_TrapAnalysis', f'A1:M{ta_last}')
    set_widths(ta, [34, 12, 12, 12, 14, 16, 15, 15, 18, 18, 18, 16, 14])
    ta.freeze_panes = 'A2'
    TA = 'tbl_TrapAnalysis'

    # ---------------- KPI_Data ----------------
    kd = wb.create_sheet('KPI_Data')
    set_widths(kd, [46, 18, 14, 60])
    kd['A1'] = 'KPI'; kd['B1'] = 'Value'; kd['C1'] = 'Unit'; kd['D1'] = 'Formula / Source Note'
    style_header(kd, 1, 4)

    kpi_defs = [
        ('traps_monitored', 'Traps Monitored', lambda k: f'=Assumptions_Flags!$C${TRAPS_MONITORED_ROW}', 'traps',
         'From Assumptions_Flags input - update if the true installed base differs.', '0'),
        ('total_events', 'Total Events', lambda k: f'=SUBTOTAL(103,{AD}[Alarm_ID])', 'events', 'Count of all rows in Alarm_Data.', '0'),
        ('leak_events', 'Leak Events', lambda k: f'=COUNTIF({AD}[Fault_Category],"Leak")', 'events', 'Fault_Category = Leak (Partial + Full).', '0'),
        ('block_events', 'Block Events', lambda k: f'=COUNTIF({AD}[Fault_Category],"Block")', 'events', 'Fault_Category = Block.', '0'),
        ('traps_flagged', 'Traps Flagged', lambda k: f'=SUBTOTAL(103,{TA}[Trap_ID])', 'traps', 'Distinct traps that raised >=1 alarm this month.', '0'),
        ('healthy_traps', 'Healthy Traps', lambda k: f'=MAX(0,B{k["traps_monitored"]}-B{k["traps_flagged"]})', 'traps', 'Traps Monitored - Traps Flagged.', '0'),
        ('leak_hours', 'Leak Hours', lambda k: f'=SUMIF({AD}[Fault_Category],"Leak",{AD}[Duration_Hours])', 'hours', 'Sum of Duration_Hours for Leak events.', '0.0'),
        ('partial_hours', 'Partial Leak Hours', lambda k: f'=SUMIFS({AD}[Duration_Hours],{AD}[Fault_Severity],"Partial")', 'hours', '', '0.0'),
        ('full_hours', 'Full Leak Hours', lambda k: f'=SUMIFS({AD}[Duration_Hours],{AD}[Fault_Severity],"Full")', 'hours', '', '0.0'),
        ('steam_lost', 'Steam Lost (this month)', lambda k: f'=B{k["partial_hours"]}*Assumptions_Flags!$C${PARTIAL_RATE_ROW}+B{k["full_hours"]}*Assumptions_Flags!$C${FULL_RATE_ROW}', 'Tons',
         'Partial hours x partial rate + full hours x full rate.', '0.000'),
        ('steam_cost', 'Steam Cost per Ton', lambda k: f'=Assumptions_Flags!$C${COST_ROW}', 'Rs/Ton', '', '#,##0'),
        ('days_monitored', 'Days Monitored', lambda k: f'=_xlfn.MAXIFS({AD}[Event_Date],{AD}[Event_Date],">=0")-_xlfn.MINIFS({AD}[Event_Date],{AD}[Event_Date],">=0")+1', 'days', 'Span of the alarm export.', '0'),
        ('annual_steam_loss', 'Annual Steam Loss (run-rate)', lambda k: f'=B{k["steam_lost"]}/B{k["days_monitored"]}*365', 'Tons/yr', 'Steam Lost scaled to a full year.', '0.0'),
        ('annual_steam_value', 'Annual Steam Value (run-rate)', lambda k: f'=B{k["annual_steam_loss"]}*B{k["steam_cost"]}', 'Rs', 'Annual Steam Loss x Cost per Ton.', '#,##0'),
        ('investment', f'Investment ({"one-time total" if a_investment_type == "one_time" else "monthly rental"})',
         lambda k: f'=Assumptions_Flags!$C${INVESTMENT_ROW}', 'Rs', '', '#,##0'),
        ('avoidable_loss', 'Avoidable Loss (30-day check interval)', lambda k: "='Loss_Projection'!G4", 'Rs',
         'Net Loss Avoided at the 30-day scenario in Loss_Projection (all leaking traps running continuously for 30 days, minus SteamGuard\'s ~15-min detection window).', '#,##0'),
    ]
    if a_investment_type == 'one_time':
        kpi_defs.append(
            ('roi', 'Estimated Payback Period', lambda k: f'=B{k["investment"]}/B{k["avoidable_loss"]}', 'months',
             'One-time investment / this month\'s avoidable-loss run-rate - how many months of savings at this rate '
             'would recover the investment. Not comparable to a monthly-rental ROI multiple.', '0.0" months"')
        )
    else:
        kpi_defs.append(
            ('roi', 'ROI (avoidable loss / investment)', lambda k: f'=B{k["avoidable_loss"]}/B{k["investment"]}', 'x', '', '0.0"x"')
        )

    kpi_key_rows = {}
    r = 2
    for key, *_ in kpi_defs:
        kpi_key_rows[key] = r
        r += 1

    r = 2
    for key, label, formula_fn, unit, note, fmt in kpi_defs:
        kd.cell(row=r, column=1, value=label)
        formula = formula_fn(kpi_key_rows)
        vc = kd.cell(row=r, column=2, value=formula)
        vc.number_format = fmt
        vc.font = GREEN if ('Assumptions_Flags' in formula or 'Loss_Projection' in formula) else BLACK
        kd.cell(row=r, column=3, value=unit)
        nc = kd.cell(row=r, column=4, value=note)
        nc.alignment = Alignment(wrap_text=True, vertical='top')
        kd.row_dimensions[r].height = 30 if note else 16
        for c in range(1, 5):
            kd.cell(row=r, column=c).border = BORDER
        r += 1
    kd.freeze_panes = 'A2'
    KPI_ROWS = kpi_key_rows

    # ---------------- Loss_Projection ----------------
    lp = wb.create_sheet('Loss_Projection')
    lp_headers = ['Scenario_Days', 'Leaking_Traps_Count', 'Steam_Lost_Tons', 'Potential_Loss_Rs',
                  'Residual_Steam_Tons', 'Residual_Loss_Rs', 'Net_Loss_Avoided_Rs', 'Net_Steam_Saved_Tons']
    for c, h in enumerate(lp_headers, start=1):
        lp.cell(row=1, column=c, value=h)
    style_header(lp, 1, len(lp_headers))
    lp['A9'] = ('Model: every currently-leaking trap (Leak_Events > 0 in Trap_Analysis) treated as leaking '
                'continuously at the partial-leak rate until found on the next manual round of N days. '
                'Residual columns net out the small loss still incurred during SteamGuard\'s own detection window.')
    lp['A9'].font = SUB_FONT
    lp.merge_cells('A9:H9')
    lp.row_dimensions[9].height = 30

    for i, d in enumerate([7, 15, 30, 90]):
        row = i + 2
        lp.cell(row=row, column=1, value=d)
        lp.cell(row=row, column=2, value=f'=COUNTIF({TA}[Leak_Events],">0")')
        c3 = lp.cell(row=row, column=3, value=f'=$B{row}*Assumptions_Flags!$C${PARTIAL_RATE_ROW}*$A{row}*24'); c3.number_format = '0.0'
        c4 = lp.cell(row=row, column=4, value=f'=$C{row}*Assumptions_Flags!$C${COST_ROW}'); c4.number_format = '#,##0'
        c5 = lp.cell(row=row, column=5, value=f'=$B{row}*Assumptions_Flags!$C${PARTIAL_RATE_ROW}*(Assumptions_Flags!$C${DETECT_MIN_ROW}/60)'); c5.number_format = '0.0000'
        c6 = lp.cell(row=row, column=6, value=f'=$E{row}*Assumptions_Flags!$C${COST_ROW}'); c6.number_format = '#,##0'
        c7 = lp.cell(row=row, column=7, value=f'=$D{row}-$F{row}'); c7.number_format = '#,##0'
        c8 = lp.cell(row=row, column=8, value=f'=$C{row}-$E{row}'); c8.number_format = '0.0'
        for c in range(1, 9):
            lp.cell(row=row, column=c).border = BORDER

    add_table(lp, 'tbl_LossProjection', 'A1:H5')
    set_widths(lp, [15, 20, 16, 17, 18, 16, 19, 19])

    # ---------------- Daily_Trend ----------------
    ad.cell(row=1, column=12, value='Trap_Date_Key')
    ad.cell(row=1, column=13, value='Is_First_Trap_Event_On_Date')
    style_header(ad, 1, 2, start_col=12)
    for i in range(n):
        row = i + 2
        ad.cell(row=row, column=12, value=f'=D{row}&"|"&TEXT(K{row},"yyyy-mm-dd")')
        ad.cell(row=row, column=13, value=f'=IF(COUNTIF($L$2:L{row},L{row})=1,1,0)')
    del ad.tables['tbl_AlarmData']
    add_table(ad, 'tbl_AlarmData', f'A1:M{last_row}')
    set_widths(ad, [10, 19, 19, 34, 20, 14, 12, 14, 13, 14, 12, 34, 14])

    dt = wb.create_sheet('Daily_Trend')
    dt_headers = ['Date', 'Total_Events', 'Leak_Events', 'Block_Events', 'Affected_Traps', 'Leak_Hours']
    for c, h in enumerate(dt_headers, start=1):
        dt.cell(row=1, column=c, value=h)
    style_header(dt, 1, len(dt_headers))

    all_dates = sorted(set(rec['created'].date() for rec in records))
    for i, d in enumerate(all_dates):
        row = i + 2
        dc = dt.cell(row=row, column=1, value=datetime.datetime(d.year, d.month, d.day)); dc.number_format = 'yyyy-mm-dd'
        dt.cell(row=row, column=2, value=f'=COUNTIFS({AD}[Event_Date],$A{row})')
        dt.cell(row=row, column=3, value=f'=COUNTIFS({AD}[Event_Date],$A{row},{AD}[Fault_Category],"Leak")')
        dt.cell(row=row, column=4, value=f'=COUNTIFS({AD}[Event_Date],$A{row},{AD}[Fault_Category],"Block")')
        dt.cell(row=row, column=5, value=f'=SUMIFS({AD}[Is_First_Trap_Event_On_Date],{AD}[Event_Date],$A{row})')
        c6 = dt.cell(row=row, column=6, value=f'=SUMIFS({AD}[Duration_Hours],{AD}[Event_Date],$A{row},{AD}[Fault_Category],"Leak")'); c6.number_format = '0.0'
        for c in range(1, 7):
            dt.cell(row=row, column=c).border = BORDER

    dt_last = len(all_dates) + 1
    add_table(dt, 'tbl_DailyTrend', f'A1:F{dt_last}')
    set_widths(dt, [14, 14, 14, 14, 16, 14])
    dt.freeze_panes = 'A2'

    # ---------------- Management_View ----------------
    mv = wb.create_sheet('Management_View')
    set_widths(mv, [30, 16, 3, 30, 16, 3, 22, 16, 16])
    mv['A1'] = f'{client_name.replace("_", " ")} - SteamGuard Monthly Overview'
    mv['A1'].font = TITLE_FONT
    mv.merge_cells('A1:I1')
    mv['A2'] = 'Monthly monitoring summary — charts below show where leaks are concentrated and their cost impact.'
    mv['A2'].font = SUB_FONT
    mv.merge_cells('A2:I2')

    # ---- Bottom-line callout: puts the cost-savings numbers Sir wants
    # front-and-center, ahead of the detailed KPI table below, instead of
    # them being buried in a list of ~11 metrics. ----
    banner_formula = (
        '="💰 BOTTOM LINE:   Annual steam value at risk ₹"&TEXT(KPI_Data!B'
        + str(KPI_ROWS['annual_steam_value']) + ',"#,##0")'
        + '&"   |   Avoidable loss (30-day) ₹"&TEXT(KPI_Data!B' + str(KPI_ROWS['avoidable_loss']) + ',"#,##0")'
    )
    if a_investment_type == 'one_time':
        banner_formula += '&"   |   Est. payback "&TEXT(KPI_Data!B' + str(KPI_ROWS['roi']) + ',"0.0")&" months"'
    else:
        banner_formula += '&"   |   ROI "&TEXT(KPI_Data!B' + str(KPI_ROWS['roi']) + ',"0.0")&"x"'
    mv['A3'] = banner_formula
    mv['A3'].font = Font(name=FONT_NAME, bold=True, size=13, color='B8714A')
    mv.merge_cells('A3:I3')
    mv.row_dimensions[3].height = 26

    mv['A4'] = 'KEY METRICS'
    mv['A4'].font = SECTION_FONT
    mv['A4'].fill = SECTION_FILL
    for c in range(1, 3):
        mv.cell(row=4, column=c).fill = SECTION_FILL

    kpi_display = [
        ('Traps Monitored', 'traps_monitored', '0'),
        ('Traps Flagged', 'traps_flagged', '0'),
        ('Healthy Traps', 'healthy_traps', '0'),
        ('Total Events', 'total_events', '0'),
        ('Leak Events', 'leak_events', '0'),
        ('Block Events', 'block_events', '0'),
        ('Steam Lost (Tons, this month)', 'steam_lost', '0.00'),
        ('Annual Steam Value (Rs run-rate)', 'annual_steam_value', '#,##0'),
        (f'Investment ({"one-time" if a_investment_type == "one_time" else "Rs/month"})', 'investment', '#,##0'),
        ('Avoidable Loss - 30 day (Rs)', 'avoidable_loss', '#,##0'),
        ('Est. Payback Period (months)', 'roi', '0.0" months"') if a_investment_type == 'one_time'
        else ('ROI (x)', 'roi', '0.0"x"'),
    ]
    r = 5
    for label, key, fmt in kpi_display:
        mv.cell(row=r, column=1, value=label)
        c2 = mv.cell(row=r, column=2, value=f'=KPI_Data!B{KPI_ROWS[key]}')
        c2.number_format = fmt
        c2.font = Font(name=FONT_NAME, bold=True, size=12, color='1F4E78')
        for c in range(1, 3):
            mv.cell(row=r, column=c).border = BORDER
        r += 1

    n_top = min(10, len(trap_ids))
    mv['D4'] = f'TOP {n_top} TRAPS BY LEAK HOURS'
    mv['D4'].font = SECTION_FONT
    mv['D4'].fill = SECTION_FILL
    for c in range(4, 6):
        mv.cell(row=4, column=c).fill = SECTION_FILL
    mv.cell(row=5, column=4, value='Trap_ID')
    mv.cell(row=5, column=5, value='Leak_Hours')
    style_header(mv, 5, 2, start_col=4)
    for i in range(n_top):
        row = 6 + i
        k = i + 1
        mv.cell(row=row, column=4, value=f'=INDEX({TA}[Trap_ID],MATCH({k},{TA}[Leak_Hours_Rank],0))')
        c5 = mv.cell(row=row, column=5, value=f'=INDEX({TA}[Total_Leak_Hours],MATCH({k},{TA}[Leak_Hours_Rank],0))')
        c5.number_format = '0.0'
        for c in range(4, 6):
            mv.cell(row=row, column=c).border = BORDER
    for c in range(4, 6):
        mv.cell(row=5, column=c).border = BORDER

    mv['G4'] = 'EVENT MIX'
    mv['G4'].font = SECTION_FONT
    mv['G4'].fill = SECTION_FILL
    for c in range(7, 9):
        mv.cell(row=4, column=c).fill = SECTION_FILL
    mv['G5'] = 'Category'
    mv['H5'] = 'Events'
    style_header(mv, 5, 2, start_col=7)
    mv['G6'] = 'Leak Events'
    mv['H6'] = f'=KPI_Data!B{KPI_ROWS["leak_events"]}'
    mv['G7'] = 'Block Events'
    mv['H7'] = f'=KPI_Data!B{KPI_ROWS["block_events"]}'
    for rr in (6, 7):
        for c in range(7, 9):
            mv.cell(row=rr, column=c).border = BORDER

    mv['G9'] = 'LOSS IF LEAKS RUN UNDETECTED'
    mv['G9'].font = SECTION_FONT
    mv['G9'].fill = SECTION_FILL
    for c in range(7, 9):
        mv.cell(row=9, column=c).fill = SECTION_FILL
    mv['G10'] = 'Check Interval (days)'
    mv['H10'] = 'Potential Loss (Rs)'
    style_header(mv, 10, 2, start_col=7)
    for i, d in enumerate([7, 15, 30, 90]):
        row = 11 + i
        mv.cell(row=row, column=7, value=f"='Loss_Projection'!A{i+2}")
        c8 = mv.cell(row=row, column=8, value=f"='Loss_Projection'!D{i+2}")
        c8.number_format = '#,##0'
        for c in range(7, 9):
            mv.cell(row=row, column=c).border = BORDER

    bar1 = BarChart()
    bar1.type = 'col'
    bar1.style = 10
    bar1.title = f'Top {n_top} Traps by Leak Hours'
    bar1.y_axis.title = 'Leak Hours'
    bar1.x_axis.title = 'Trap'
    data = Reference(mv, min_col=5, min_row=5, max_row=5 + n_top)
    cats = Reference(mv, min_col=4, min_row=6, max_row=5 + n_top)
    bar1.add_data(data, titles_from_data=True)
    bar1.set_categories(cats)
    bar1.height = 9
    bar1.width = 16
    bar1.legend = None
    bar1.gapWidth = 40
    s1 = bar1.series[0]
    s1.graphicalProperties.solidFill = CHART_BLUE
    s1.graphicalProperties.line.noFill = True
    s1.dLbls = DataLabelList(showVal=True, showCatName=False, showSerName=False,
                              showLegendKey=False, showPercent=False, numFmt='0.0', dLblPos='outEnd')
    bar1.y_axis.majorGridlines = None
    bar1.x_axis.delete = False
    bar1.y_axis.delete = False
    mv.add_chart(bar1, 'A18')

    pie1 = PieChart()
    pie1.style = 10
    pie1.title = 'Leak vs Block Events'
    pdata = Reference(mv, min_col=8, min_row=5, max_row=7)
    pcats = Reference(mv, min_col=7, min_row=6, max_row=7)
    pie1.add_data(pdata, titles_from_data=True)
    pie1.set_categories(pcats)
    pie1.height = 9
    pie1.width = 13
    s2 = pie1.series[0]
    s2.dLbls = DataLabelList(showVal=True, showPercent=True, showCatName=False,
                              showSerName=False, showLegendKey=False)
    leak_pt = DataPoint(idx=0)
    leak_pt.graphicalProperties.solidFill = CHART_BLUE
    block_pt = DataPoint(idx=1)
    block_pt.graphicalProperties.solidFill = CHART_ACCENT_RED
    s2.data_points = [leak_pt, block_pt]
    # Stacked one below another (not side-by-side) with a wide, fixed row
    # gap between each — anchoring them side-by-side at nearby columns
    # previously made them overlap, since a chart this wide (16-18cm) spans
    # far more columns than its anchor cell alone suggests.
    mv.add_chart(pie1, 'A39')

    bar2 = BarChart()
    bar2.type = 'col'
    bar2.style = 10
    bar2.title = 'Potential Loss if Undetected (Rs)'
    bar2.y_axis.title = 'Rs'
    bar2.x_axis.title = 'Days Undetected'
    bdata = Reference(mv, min_col=8, min_row=10, max_row=14)
    bcats = Reference(mv, min_col=7, min_row=11, max_row=14)
    bar2.add_data(bdata, titles_from_data=True)
    bar2.set_categories(bcats)
    bar2.height = 9
    bar2.width = 16
    bar2.legend = None
    bar2.gapWidth = 40
    s3 = bar2.series[0]
    s3.dLbls = DataLabelList(showVal=True, showCatName=False, showSerName=False,
                              showLegendKey=False, showPercent=False, numFmt='#,##0', dLblPos='outEnd')
    # Escalating colors — light blue for the shortest blind-spot window,
    # deepening to the warm accent red for the longest (90-day) one, so the
    # growing risk reads visually, not just numerically.
    escalation_colors = [CHART_LIGHT_BLUE, CHART_MID_BLUE, CHART_BLUE, CHART_ACCENT_RED]
    points = []
    for i, color in enumerate(escalation_colors):
        dp = DataPoint(idx=i)
        dp.graphicalProperties.solidFill = color
        points.append(dp)
    s3.data_points = points
    bar2.y_axis.majorGridlines = None
    bar2.x_axis.delete = False
    bar2.y_axis.delete = False
    mv.add_chart(bar2, 'A60')

    # ---------------- Daily Trend chart (new) ----------------
    # Leak hours logged per day across the month — makes it visually obvious
    # whether leaks are a steady background problem or concentrated in a few
    # bad days, which the KPI table alone doesn't show.
    line1 = LineChart()
    line1.style = 12
    line1.title = 'Leak Hours Logged Per Day'
    line1.y_axis.title = 'Leak Hours'
    line1.x_axis.title = 'Date'
    ldata = Reference(dt, min_col=6, min_row=1, max_row=dt_last)
    lcats = Reference(dt, min_col=1, min_row=2, max_row=dt_last)
    line1.add_data(ldata, titles_from_data=True)
    line1.set_categories(lcats)
    line1.height = 9
    line1.width = 16
    line1.legend = None
    ls = line1.series[0]
    ls.graphicalProperties.line.solidFill = CHART_NAVY
    ls.graphicalProperties.line.width = 20000  # EMUs, ~1.6pt
    ls.smooth = False
    line1.y_axis.majorGridlines = None
    line1.x_axis.delete = False
    line1.y_axis.delete = False
    mv.add_chart(line1, 'A81')

    mv.sheet_view.showGridLines = False
    # Landscape + fit-to-width so printing/PDF-exporting this sheet doesn't
    # split the tables and charts across multiple pages.
    mv.page_setup.orientation = 'landscape'
    mv.page_setup.fitToWidth = 1
    mv.page_setup.fitToHeight = 0
    mv.sheet_properties.pageSetUpPr.fitToPage = True
    # Terracotta tab color so the primary/first sheet stands out at a glance
    # (matches the PPT's accent color).
    mv.sheet_properties.tabColor = 'B8714A'

    # Loss_Projection stays in the workbook (KPI_Data and Management_View's
    # "Loss if leaks run undetected" table both pull numbers from it via
    # formulas), but it's hidden from the visible tabs — it's backing data,
    # not something that needs to be shown/explained on its own.
    lp.sheet_state = 'hidden'

    # Final visible tab order, as requested: Management_View first, then
    # KPI_Data, Assumptions_Flags, Daily_Trend, Trap_Analysis, Alarm_Data.
    # Loss_Projection is hidden backing data, so it's tucked at the very end.
    desired_order = ['Management_View', 'KPI_Data', 'Assumptions_Flags',
                      'Daily_Trend', 'Trap_Analysis', 'Alarm_Data', 'Loss_Projection']
    wb._sheets = [wb[name] for name in desired_order]

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                f = cell.font
                if f.name != FONT_NAME:
                    cell.font = Font(name=FONT_NAME, size=f.size, bold=f.bold, italic=f.italic,
                                      color=f.color, underline=f.underline)

    wb.save(out_path)

    # ---------------- Plain-Python KPI values (mirrors the Excel formulas
    # above) so callers (e.g. the PPTX management-report generator) can get
    # the same numbers without needing Excel/LibreOffice to recalculate the
    # workbook's formulas first. ----------------
    full_rate = 0.24725 * (a_pressure + 1) * (a_orifice ** 2) * a_full_sev / 1000
    partial_rate = full_rate * a_partial_sev
    leak_hours = sum(rec['duration_hours'] for rec in records if rec['category'] == 'Leak')
    partial_hours = sum(rec['duration_hours'] for rec in records if rec['sev_type'] == 'Partial')
    full_hours = sum(rec['duration_hours'] for rec in records if rec['sev_type'] == 'Full')
    block_hours = sum(rec['duration_hours'] for rec in records if rec['category'] == 'Block')
    steam_lost_tons = partial_hours * partial_rate + full_hours * full_rate
    days_span = (all_dates[-1] - all_dates[0]).days + 1
    annual_steam_loss_tons = (steam_lost_tons / days_span * 365) if days_span else 0.0
    annual_steam_value_rs = annual_steam_loss_tons * a_cost
    leaking_traps = len(set(rec['trap_id'] for rec in records if rec['category'] == 'Leak'))
    healthy_traps = max(0, traps_monitored_default - len(trap_ids))

    loss_scenarios = {}
    for d in (7, 15, 30, 90):
        steam_lost_scenario = leaking_traps * partial_rate * d * 24
        potential_loss = steam_lost_scenario * a_cost
        residual_tons = leaking_traps * partial_rate * (a_detect_min / 60)
        residual_loss = residual_tons * a_cost
        loss_scenarios[d] = {
            'steam_lost_tons': steam_lost_scenario,
            'potential_loss_rs': potential_loss,
            'net_loss_avoided_rs': potential_loss - residual_loss,
        }
    avoidable_loss_30day = loss_scenarios[30]['net_loss_avoided_rs']

    # ---------------- Per-trap breakdown (repeated faulty traps) ----------
    # Sir's feedback: a slide/table identifying WHICH traps keep recurring
    # (leaking vs blocked) and how much each one is costing per month - not
    # just an aggregate number. Built once here so both the Excel (Trap
    # Analysis sheet, via formulas) and the PPTX (via this plain-Python
    # list) show the same numbers.
    trap_breakdown = []
    for tid in trap_ids:
        trap_recs = [rec for rec in records if rec['trap_id'] == tid]
        leak_ev = sum(1 for rec in trap_recs if rec['category'] == 'Leak')
        block_ev = sum(1 for rec in trap_recs if rec['category'] == 'Block')
        partial_h = sum(rec['duration_hours'] for rec in trap_recs if rec['sev_type'] == 'Partial')
        full_h = sum(rec['duration_hours'] for rec in trap_recs if rec['sev_type'] == 'Full')
        leak_h = sum(rec['duration_hours'] for rec in trap_recs if rec['category'] == 'Leak')
        block_h = sum(rec['duration_hours'] for rec in trap_recs if rec['category'] == 'Block')
        steam_lost_trap = partial_h * partial_rate + full_h * full_rate
        if leak_ev > 0 and block_ev > 0:
            status = 'Leak + Block'
        elif leak_ev > 0:
            status = 'Leak'
        elif block_ev > 0:
            status = 'Block'
        else:
            status = 'Other'
        trap_breakdown.append({
            'trap_id': tid,
            'total_events': len(trap_recs),
            'leak_events': leak_ev,
            'block_events': block_ev,
            'leak_hours': leak_h,
            'block_hours': block_h,
            'steam_lost_tons': steam_lost_trap,
            'monthly_cost_rs': steam_lost_trap * a_cost,
            'status': status,
        })
    trap_breakdown.sort(key=lambda x: (-x['leak_hours'], -x['total_events']))

    roi_x = None
    payback_months = None
    if a_investment_type == 'one_time':
        payback_months = (a_investment / avoidable_loss_30day) if avoidable_loss_30day else None
    else:
        roi_x = (avoidable_loss_30day / a_investment) if a_investment else None

    summary = {
        'client_name': client_name,
        'total_events': n,
        'total_rows_in_file': total_rows_in_file,
        'traps_flagged': len(trap_ids),
        'traps_monitored': traps_monitored_default,
        'healthy_traps': healthy_traps,
        'date_range': (all_dates[0].isoformat(), all_dates[-1].isoformat()),
        'other_type_count': sum(1 for rec in records if rec['category'] == 'Other'),
        'unrecognized_types': sorted(unrecognized_types),
        'duplicate_rows_removed': duplicate_rows_removed,
        'leak_events': sum(1 for rec in records if rec['category'] == 'Leak'),
        'block_events': sum(1 for rec in records if rec['category'] == 'Block'),
        'leak_hours': leak_hours,
        'partial_hours': partial_hours,
        'full_hours': full_hours,
        'block_hours': block_hours,
        'steam_lost_tons': steam_lost_tons,
        'annual_steam_loss_tons': annual_steam_loss_tons,
        'annual_steam_value_rs': annual_steam_value_rs,
        'cost_per_ton': a_cost,
        'investment': a_investment,
        'investment_type': a_investment_type,
        'avoidable_loss_30day': avoidable_loss_30day,
        'loss_scenarios': loss_scenarios,
        'roi_x': roi_x,
        'payback_months': payback_months,
        'trap_breakdown': trap_breakdown,
    }
    return out_path, summary


if __name__ == "__main__":
    import sys
    client = sys.argv[1] if len(sys.argv) > 1 else "TestClient"
    csv_in = sys.argv[2] if len(sys.argv) > 2 else f"clients/{client}_alarms.csv"
    out = sys.argv[3] if len(sys.argv) > 3 else f"{client}_SteamGuard_Report.xlsx"
    path, info = build_report(client, csv_in, out)
    print("Saved:", path)
    print(info)
