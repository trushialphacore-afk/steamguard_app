"""
pptx_report.py
--------------
Auto-generates a short, management-level PowerPoint (4 slides) straight from
the same summary dict build_report() already returns — no separate data
entry needed.

DESIGN: matches the look of Alphacore's hand-designed case-study decks
(e.g. the JK Tyre Banmore one) — clean white background, the Alphacore
triangle logo + wordmark in the top-left corner of every slide (and the
full logo centered on the closing slide), a terracotta accent color for
section labels/headline numbers, and light mint stat tiles — rather than
the old solid-navy header-bar look. Needs the two logo PNGs shipped in
./assets/ next to this file (alphacore_icon.png, alphacore_logo_full.png).
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

ASSETS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets')
ICON_PATH = os.path.join(ASSETS_DIR, 'alphacore_icon.png')
FULL_LOGO_PATH = os.path.join(ASSETS_DIR, 'alphacore_logo_full.png')

# Palette sampled directly from Alphacore's own case-study template so the
# auto-generated deck matches it, not the old navy/blue theme.
TERRACOTTA = RGBColor(0xB8, 0x71, 0x4A)   # section labels, big numbers, accent
CHARCOAL = RGBColor(0x2C, 0x38, 0x30)      # headlines, bold labels
SAGE_GREY = RGBColor(0x54, 0x6B, 0x5E)     # subtitles / body text
CAPTION_GREY = RGBColor(0x7A, 0x8A, 0x80)  # small captions
TILE_BG = RGBColor(0xF4, 0xFA, 0xF6)       # stat tile fill
TILE_BORDER = RGBColor(0xD2, 0xE8, 0xD8)   # stat tile border
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ALERT_RED = RGBColor(0xB0, 0x2A, 0x2A)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)


def _blank_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])  # blank layout


def _textbox(slide, left, top, width, height, text, size=14, bold=False, italic=False,
             color=CHARCOAL, align=PP_ALIGN.LEFT, font_name='Arial', anchor=None,
             wrap=True, letter_spacing_caps=False):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = wrap
    if anchor is not None:
        tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text.upper() if letter_spacing_caps else text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = font_name
    return box


def _header(slide, eyebrow_text):
    """Small top-left icon + ALPHACORE wordmark + tagline, then a terracotta
    eyebrow label — replaces the old full-width navy header bar."""
    if os.path.exists(ICON_PATH):
        slide.shapes.add_picture(ICON_PATH, Inches(0.45), Inches(0.32), height=Inches(0.42))
        wordmark_left = Inches(1.0)
    else:
        wordmark_left = Inches(0.45)
    _textbox(slide, wordmark_left, Inches(0.3), Inches(4), Inches(0.35),
              'ALPHACORE', size=16, bold=True, color=CHARCOAL, letter_spacing_caps=True)
    _textbox(slide, wordmark_left, Inches(0.62), Inches(4.5), Inches(0.3),
              'Industrial IoT  ·  Plant Intelligence', size=10, color=CAPTION_GREY)
    _textbox(slide, Inches(0.45), Inches(1.05), Inches(11), Inches(0.35),
              eyebrow_text, size=12, bold=True, color=TERRACOTTA, letter_spacing_caps=True)


def _footer(slide, client_name, page_no, total_pages):
    _textbox(slide, Inches(0.45), Inches(7.1), Inches(9), Inches(0.3),
              f'Alphacore SteamGuard™ · {client_name.replace("_", " ")} — Monthly Report',
              size=9, color=CAPTION_GREY)
    _textbox(slide, Inches(12.3), Inches(7.1), Inches(0.8), Inches(0.3),
              f'{page_no} / {total_pages}', size=9, color=CAPTION_GREY, align=PP_ALIGN.RIGHT)


def _stat_tile(slide, left, top, width, height, big_text, small_text, big_size=32):
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    box.fill.solid()
    box.fill.fore_color.rgb = TILE_BG
    box.line.color.rgb = TILE_BORDER
    box.line.width = Pt(1)
    box.shadow.inherit = False
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.15)
    tf.margin_right = Inches(0.15)
    p1 = tf.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    r1 = p1.add_run()
    r1.text = big_text
    r1.font.size = Pt(big_size)
    r1.font.bold = True
    r1.font.color.rgb = TERRACOTTA
    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run()
    r2.text = small_text
    r2.font.size = Pt(11)
    r2.font.bold = True
    r2.font.color.rgb = CHARCOAL
    return box


def _add_event_mix_pie(slide, left, top, width, height, leak_events, block_events):
    """Small pie chart for the Leak vs Block event split, in the terracotta /
    sage palette so it matches the rest of the deck."""
    chart_data = CategoryChartData()
    chart_data.categories = [f'Leak Events ({leak_events})', f'Block Events ({block_events})']
    chart_data.add_series('Events', (leak_events, block_events))
    gframe = slide.shapes.add_chart(XL_CHART_TYPE.PIE, left, top, width, height, chart_data)
    chart = gframe.chart
    chart.has_title = False
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.RIGHT
    chart.legend.include_in_layout = False
    chart.legend.font.size = Pt(11)
    chart.legend.font.color.rgb = CHARCOAL
    plot = chart.plots[0]
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.show_value = False
    dl.show_percentage = True
    dl.show_category_name = False
    dl.number_format = '0%'
    dl.number_format_is_linked = False
    dl.font.size = Pt(12)
    dl.font.bold = True
    dl.font.color.rgb = WHITE
    points = plot.series[0].points
    points[0].format.fill.solid()
    points[0].format.fill.fore_color.rgb = TERRACOTTA
    points[1].format.fill.solid()
    points[1].format.fill.fore_color.rgb = SAGE_GREY
    return gframe


def _add_cost_impact_chart(slide, left, top, width, height, loss_scenarios):
    """Bar chart of potential loss by blind-spot duration — light mint to
    terracotta escalation, matching the tile palette used across the deck."""
    days = [7, 15, 30, 90]
    chart_data = CategoryChartData()
    chart_data.categories = [f'{d} days' for d in days]
    chart_data.add_series('Potential Loss (Rs)', tuple(loss_scenarios[d]['potential_loss_rs'] for d in days))
    gframe = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, left, top, width, height, chart_data)
    chart = gframe.chart
    chart.has_title = False
    chart.has_legend = False
    plot = chart.plots[0]
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.show_value = True
    dl.number_format = '#,##0'
    dl.number_format_is_linked = False
    dl.font.size = Pt(11)
    dl.font.bold = True
    dl.font.color.rgb = CHARCOAL
    plot.gap_width = 60
    escalation_colors = [RGBColor(0xD2, 0xE8, 0xD8), RGBColor(0x9C, 0xC2, 0xA8),
                          RGBColor(0xE0, 0xA9, 0x86), TERRACOTTA]
    points = plot.series[0].points
    for pt, color in zip(points, escalation_colors):
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = color
    chart.category_axis.tick_labels.font.size = Pt(11)
    chart.category_axis.tick_labels.font.color.rgb = CHARCOAL
    chart.value_axis.tick_labels.font.size = Pt(10)
    chart.value_axis.tick_labels.font.color.rgb = CAPTION_GREY
    chart.value_axis.has_major_gridlines = False
    return gframe


def _add_cumulative_savings_chart(slide, left, top, width, height, monthly_avoidable_loss, investment):
    """12-month 'cumulative savings vs investment' line chart, matching the
    reference Alphacore case-study deck (JK Tyre Banmore, page 6) — a rising
    cumulative-savings line compared against a flat investment reference
    line, so the payback point is visually obvious rather than a single
    number buried in a stat tile."""
    months = [f'M{i}' for i in range(1, 13)]
    cumulative = [monthly_avoidable_loss * i for i in range(1, 13)]
    flat_investment = [investment] * 12

    chart_data = CategoryChartData()
    chart_data.categories = months
    chart_data.add_series('Cumulative savings (Rs)', tuple(cumulative))
    chart_data.add_series('Investment (Rs)', tuple(flat_investment))
    gframe = slide.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, left, top, width, height, chart_data)
    chart = gframe.chart
    chart.has_title = False
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.legend.font.size = Pt(11)
    chart.legend.font.color.rgb = CHARCOAL

    savings_series = chart.plots[0].series[0]
    savings_series.format.line.color.rgb = TERRACOTTA
    savings_series.format.line.width = Pt(2.5)
    savings_series.marker.style = 8  # circle
    savings_series.marker.format.fill.solid()
    savings_series.marker.format.fill.fore_color.rgb = TERRACOTTA
    savings_series.marker.format.line.color.rgb = TERRACOTTA

    investment_series = chart.plots[0].series[1]
    investment_series.format.line.color.rgb = SAGE_GREY
    investment_series.format.line.width = Pt(1.5)
    investment_series.format.line.dash_style = 2  # dashed reference line
    investment_series.marker.style = -4142  # none

    chart.category_axis.tick_labels.font.size = Pt(10)
    chart.category_axis.tick_labels.font.color.rgb = CAPTION_GREY
    chart.value_axis.tick_labels.font.size = Pt(10)
    chart.value_axis.tick_labels.font.color.rgb = CAPTION_GREY
    chart.value_axis.has_major_gridlines = False
    chart.value_axis.tick_labels.number_format = '#,##0'
    chart.value_axis.tick_labels.number_format_is_linked = False
    return gframe


def _add_trap_table(slide, left, top, width, max_height, trap_rows):
    """Table identifying which specific traps keep recurring (leaking vs
    blocked) and what each one costs per month — the thing a headline
    aggregate number doesn't show on its own. Fault type gets its own
    plain-text 'Status' column (not just a color) so it reads unambiguously
    without needing a color-key legend.

    Row heights are fixed (not the whole table stretched to fill
    max_height) — with python-pptx, a table's total height gets divided
    evenly across however many rows it has, so with only 1-2 qualifying
    traps that made the header and data rows balloon to ~2in each. Fixing
    the per-row height keeps it readable regardless of row count."""
    HEADER_H = Inches(0.45)
    ROW_H = Inches(0.5)
    n_rows = len(trap_rows) + 1  # + header
    n_cols = 7
    table_h = min(max_height, HEADER_H + ROW_H * len(trap_rows))
    gframe = slide.shapes.add_table(n_rows, n_cols, left, top, width, table_h)
    table = gframe.table
    table.rows[0].height = HEADER_H
    for r in range(1, n_rows):
        table.rows[r].height = ROW_H

    col_widths = [0.22, 0.14, 0.11, 0.11, 0.12, 0.14, 0.16]
    for i, frac in enumerate(col_widths):
        table.columns[i].width = Emu(int(width * frac))

    headers = ['Trap ID', 'Status', 'Leak Events', 'Block Events', 'Leak Hours',
               'Steam Loss (Tons)', 'Est. Cost / Month (Rs)']
    for c, h in enumerate(headers):
        cell = table.cell(0, c)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = CHARCOAL
        cell.margin_top = Pt(2)
        cell.margin_bottom = Pt(2)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER if c > 0 else PP_ALIGN.LEFT
        run = p.runs[0]
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = WHITE

    status_color = {'Leak': TERRACOTTA, 'Block': SAGE_GREY, 'Leak + Block': ALERT_RED, 'Other': CAPTION_GREY}
    status_label = {'Leak': 'Leak', 'Block': 'Block', 'Leak + Block': 'Leak + Block', 'Other': 'Other'}
    for r, row in enumerate(trap_rows, start=1):
        status = row['status']
        values = [
            row['trap_id'],
            status_label.get(status, status),
            str(row['leak_events']),
            str(row['block_events']),
            f"{row['leak_hours']:.1f}",
            f"{row['steam_lost_tons']:.3f}",
            _fmt_rs(row['monthly_cost_rs']),
        ]
        for c, val in enumerate(values):
            cell = table.cell(r, c)
            cell.text = val
            cell.margin_top = Pt(2)
            cell.margin_bottom = Pt(2)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE if r % 2 else TILE_BG
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if c > 0 else PP_ALIGN.LEFT
            run = p.runs[0]
            run.font.size = Pt(11.5)
            if c == 1:
                run.font.color.rgb = status_color.get(status, CHARCOAL)
                run.font.bold = True
            else:
                run.font.color.rgb = CHARCOAL
                run.font.bold = (c == 0)
    return gframe


def _fmt_rs(v):
    if v is None:
        return "—"
    if abs(v) >= 100000:
        return f"₹{v/100000:.2f} L"
    return f"₹{v:,.0f}"


def build_pptx_report(summary, month_label, out_path, investment_confirmed=True):
    """
    summary: the dict returned by report_engine.build_report()'s second
        return value — this function reads it directly, so it always
        reflects the exact same numbers as the Excel workbook (no separate
        recalculation needed).
    month_label: e.g. "July 2026" — just for the title/subtitle text.
    out_path: where to save the .pptx.
    investment_confirmed: pass False if the Investment figure is still a
        placeholder/default rather than a confirmed client value — a
        caption is added to the ROI slide flagging that so a management
        deck never silently shows an unverified number as fact.
    """
    client_name = summary['client_name']
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    total_pages = 5
    tile_w, gap = Inches(2.85), Inches(0.25)

    # ---------------- Slide 1: Overview / hero stats ----------------
    s1 = _blank_slide(prs)
    _header(s1, 'Monitoring Overview')
    _textbox(s1, Inches(0.45), Inches(1.45), Inches(11.5), Inches(0.6),
              f'{client_name.replace("_", " ")} — {month_label}', size=28, bold=True, color=CHARCOAL)
    _textbox(s1, Inches(0.45), Inches(2.05), Inches(12.3), Inches(0.5),
              f'SteamGuard™ monthly monitoring summary', size=15, italic=True, color=SAGE_GREY)
    _textbox(s1, Inches(0.45), Inches(2.55), Inches(12.3), Inches(0.7),
              f'SteamGuard™ continuously monitored {summary["traps_monitored"]} steam traps at '
              f'{client_name.replace("_", " ")} in {month_label}, flagging {summary["traps_flagged"]} traps with '
              f'leak or block events out of the full installed base.', size=13, color=CHARCOAL)

    tiles = [
        (str(summary['traps_monitored']), 'steam traps monitored'),
        (str(summary['traps_flagged']), 'traps flagged this month'),
        (f'{summary["leak_hours"]:.0f}', 'leak hours logged'),
        (f'{summary["steam_lost_tons"]:.2f} T', 'steam lost this month'),
    ]
    x = Inches(0.45)
    for big, small in tiles:
        _stat_tile(s1, x, Inches(3.4), tile_w, Inches(1.6), big, small)
        x += tile_w + gap

    _textbox(s1, Inches(0.45), Inches(5.35), Inches(6.3), Inches(0.4),
              'Event mix', size=14, bold=True, color=TERRACOTTA, letter_spacing_caps=True)
    _textbox(s1, Inches(0.45), Inches(5.8), Inches(6.3), Inches(0.9),
              f'Leak events: {summary["leak_events"]}   ·   Block events: {summary["block_events"]}   ·   '
              f'Healthy traps: {summary["healthy_traps"]}', size=13, color=CHARCOAL)
    _add_event_mix_pie(s1, Inches(7.1), Inches(5.15), Inches(5.2), Inches(1.85),
                        summary['leak_events'], summary['block_events'])
    if summary['leak_events'] == 0 and summary['block_events'] > 0:
        # A month where every flagged event was a blockage, not a leak, is
        # a real and useful finding on its own — but the steam-loss/cost
        # slides that follow will show all-zero numbers, which can read as
        # "the tool is broken" if it isn't called out explicitly.
        _textbox(s1, Inches(0.45), Inches(6.55), Inches(12.3), Inches(0.45),
                  'ℹ No leak events this month — only blockages, which carry no steam-loss cost under this '
                  'model (the maintenance action needed is different: clear the blockage, not seal a leak).',
                  size=11, color=SAGE_GREY)
    _footer(s1, client_name, 1, total_pages)

    # ---------------- Slide 2: Detection & cost impact ----------------
    s2 = _blank_slide(prs)
    _header(s2, 'Detection & Cost Impact')
    _textbox(s2, Inches(0.45), Inches(1.45), Inches(12.3), Inches(0.6),
              'Steam lost, valued and projected', size=24, bold=True, color=CHARCOAL)
    if summary['leak_hours'] == 0:
        cost_note = 'No leak-type events this month (only blockages, if any) — steam-loss cost is ₹0.'
    else:
        cost_note = (f'Steam lost this month is valued at {_fmt_rs(summary["steam_lost_tons"] * summary["cost_per_ton"])} '
                     f'at ₹{summary["cost_per_ton"]:,.0f}/tonne — projected to '
                     f'{_fmt_rs(summary["annual_steam_value_rs"])} a year at this run-rate.')
    _textbox(s2, Inches(0.45), Inches(2.1), Inches(12.3), Inches(0.6), cost_note, size=14, color=CHARCOAL)

    tiles2 = [
        (f'{summary["steam_lost_tons"]:.2f} T', 'steam lost this month'),
        (_fmt_rs(summary['steam_lost_tons'] * summary['cost_per_ton']), 'cost this month'),
        (_fmt_rs(summary['annual_steam_value_rs']), 'annualised value at risk'),
        (_fmt_rs(summary['avoidable_loss_30day']), 'avoidable loss (30-day check)'),
    ]
    x = Inches(0.45)
    for big, small in tiles2:
        _stat_tile(s2, x, Inches(2.85), tile_w, Inches(1.5), big, small, big_size=24)
        x += tile_w + gap

    _textbox(s2, Inches(0.45), Inches(4.6), Inches(10), Inches(0.4),
              'Cost impact by blind-spot duration', size=14, bold=True, color=TERRACOTTA, letter_spacing_caps=True)

    _add_cost_impact_chart(s2, Inches(0.45), Inches(5.05), Inches(12.3), Inches(1.9),
                            summary['loss_scenarios'])
    _footer(s2, client_name, 2, total_pages)

    # ---------------- Slide 3: ROI / Investment ----------------
    s3 = _blank_slide(prs)
    _header(s3, 'Return on Investment')
    investment_label = 'one-time total' if summary.get('investment_type') == 'one_time' else 'monthly rental'
    headline = (f'{_fmt_rs(summary["investment"])} invested ({investment_label}) — ' +
                (f'estimated payback in {summary["payback_months"]:.1f} months' if summary.get('payback_months')
                 else f'ROI of {summary["roi_x"]:.1f}x this month\'s avoidable loss' if summary.get('roi_x')
                 else 'payback not calculable this month (no leaking traps)'))
    _textbox(s3, Inches(0.45), Inches(1.45), Inches(12.3), Inches(0.7), headline, size=22, bold=True, color=CHARCOAL)

    tiles3 = [
        (_fmt_rs(summary['investment']), f'Investment ({investment_label})'),
        (_fmt_rs(summary['avoidable_loss_30day']), 'avoidable loss / month (30-day)'),
        ((f'{summary["payback_months"]:.1f} mo' if summary.get('payback_months') else
          f'{summary["roi_x"]:.1f}x' if summary.get('roi_x') else '—'),
         'payback period' if summary.get('payback_months') else 'ROI multiple'),
        (str(summary['traps_flagged']), 'traps flagged this month'),
    ]
    x = Inches(0.45)
    for big, small in tiles3:
        _stat_tile(s3, x, Inches(2.5), tile_w, Inches(1.7), big, small, big_size=26)
        x += tile_w + gap

    if summary.get('investment_type') == 'one_time' and summary.get('avoidable_loss_30day'):
        _textbox(s3, Inches(0.45), Inches(4.45), Inches(10), Inches(0.4),
                  'Cumulative savings vs investment (12-month projection)', size=14, bold=True,
                  color=TERRACOTTA, letter_spacing_caps=True)
        _add_cumulative_savings_chart(s3, Inches(0.45), Inches(4.8), Inches(12.3), Inches(1.85),
                                        summary['avoidable_loss_30day'], summary['investment'])
        caption_top = Inches(6.75)
    else:
        caption_top = Inches(4.55)

    if not investment_confirmed:
        _textbox(s3, Inches(0.45), caption_top, Inches(12.3), Inches(0.35),
                  '⚠ Investment figure shown is a placeholder/default — confirm the client\'s actual '
                  'contracted amount before sharing this deck externally.', size=11, color=ALERT_RED)
    _footer(s3, client_name, 3, total_pages)

    # ---------------- Slide 4: Recurring problem traps ----------------
    s4a = _blank_slide(prs)
    _header(s4a, 'Recurring Problem Traps')
    _textbox(s4a, Inches(0.45), Inches(1.45), Inches(12.3), Inches(0.6),
              'Which traps keep coming back — and what each one costs', size=22, bold=True, color=CHARCOAL)
    # Only genuinely recurring/serious traps make this slide — a trap that
    # logged a couple of minutes of leak time isn't worth putting in front
    # of management, so this is filtered by a leak-hours threshold rather
    # than just "top 10 regardless of how small the numbers are".
    LEAK_HOURS_THRESHOLD = 10
    trap_rows = summary.get('trap_breakdown') or []
    top_traps = [t for t in trap_rows if t['leak_hours'] > LEAK_HOURS_THRESHOLD][:12]
    if top_traps:
        _textbox(s4a, Inches(0.45), Inches(2.05), Inches(12.3), Inches(0.5),
                  f'{len(top_traps)} traps with more than {LEAK_HOURS_THRESHOLD} leak hours this month, out of '
                  f'{summary["traps_flagged"]} traps flagged — see the Status column for whether each one is '
                  f'leaking, blocked, or both.',
                  size=13, color=SAGE_GREY)
        _add_trap_table(s4a, Inches(0.45), Inches(2.65), Inches(12.3), Inches(4.3), top_traps)
    else:
        _textbox(s4a, Inches(0.45), Inches(2.5), Inches(12.3), Inches(0.6),
                  f'No trap logged more than {LEAK_HOURS_THRESHOLD} leak hours this month.',
                  size=14, color=SAGE_GREY)
    _footer(s4a, client_name, 4, total_pages)

    # ---------------- Slide 5: Closing ----------------
    s4 = _blank_slide(prs)
    _textbox(s4, Inches(1), Inches(1.7), Inches(11.3), Inches(0.7),
              'Thank you,', size=32, bold=True, color=CHARCOAL, align=PP_ALIGN.CENTER)
    if os.path.exists(FULL_LOGO_PATH):
        logo_w = Inches(2.4)
        s4.shapes.add_picture(FULL_LOGO_PATH, (SLIDE_W - logo_w) // 2, Inches(2.5), width=logo_w)
    _textbox(s4, Inches(1), Inches(5.55), Inches(11.3), Inches(0.5),
              'Efficiency through Innovation delivered at Alphacore Technologies Pvt Ltd',
              size=14, italic=True, color=SAGE_GREY, align=PP_ALIGN.CENTER)
    _textbox(s4, Inches(1), Inches(6.05), Inches(11.3), Inches(0.5),
              'SteamGuard™ · Online Steam Trap Monitoring · www.alphacore.co.in',
              size=12, color=CAPTION_GREY, align=PP_ALIGN.CENTER)

    prs.save(out_path)
    return out_path
