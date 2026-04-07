from __future__ import annotations

from pathlib import Path


WORKDIR = Path("/Users/jacheflin/Documents/Playground")
OUTDIR = WORKDIR / "assets" / "nocatee_visuals"

BG = "#f6f7fb"
CARD = "#ffffff"
TEXT = "#163047"
MUTED = "#5f7080"
GRID = "#dbe3eb"
BLUE = "#2f6bff"
BLUE_SOFT = "#89aefc"
TEAL = "#22b8b3"
GREEN = "#2fa26b"
ORANGE = "#f08a24"
RED = "#df5f59"
GOLD = "#c58a14"
SLATE = "#9aa9b8"
INK = "#0f2235"


def fmt_pct(value: float) -> str:
    return f"{value:.1f}%"


def svg_text(x: float, y: float, text: str, size: int = 14, weight: int = 400, fill: str = TEXT,
             anchor: str = "start") -> str:
    return (
        f'<text x="{x:.2f}" y="{y:.2f}" font-family="Inter, Segoe UI, Arial, sans-serif" '
        f'font-size="{size}" font-weight="{weight}" fill="{fill}" text-anchor="{anchor}">{text}</text>'
    )


def stat_card(x: float, y: float, width: float, height: float, title: str, value: str, sub: str,
              accent: str) -> str:
    return f"""
    <g>
      <rect x="{x}" y="{y}" width="{width}" height="{height}" rx="18" fill="{CARD}" stroke="#dfe6ee"/>
      <rect x="{x + 18}" y="{y + 18}" width="10" height="{height - 36}" rx="5" fill="{accent}" opacity="0.9"/>
      {svg_text(x + 44, y + 34, title, size=13, weight=600, fill=MUTED)}
      {svg_text(x + 44, y + 72, value, size=32, weight=700, fill=INK)}
      {svg_text(x + 44, y + 96, sub, size=12, fill=MUTED)}
    </g>
    """


def line_path(values: list[float], left: float, top: float, width: float, height: float,
              v_min: float, v_max: float) -> str:
    step = width / (len(values) - 1)
    points = []
    for idx, value in enumerate(values):
        x = left + idx * step
        ratio = (value - v_min) / (v_max - v_min)
        y = top + height - ratio * height
        points.append((x, y))
    return "M " + " L ".join(f"{x:.2f} {y:.2f}" for x, y in points)


def point_y(value: float, top: float, height: float, v_min: float, v_max: float) -> float:
    ratio = (value - v_min) / (v_max - v_min)
    return top + height - ratio * height


def occupancy_svg() -> str:
    months = ["Mar-26", "Apr-26", "May-26", "Jun-26"]
    budget = [59.30, 61.60, 66.10, 69.10]
    dlr = [60.11, 62.36, 63.48, 64.61]
    post = [60.11, 64.05, 65.73, 66.86]

    width = 1400
    height = 900
    left = 110
    top = 280
    chart_w = 1180
    chart_h = 420
    v_min = 58
    v_max = 70

    x_step = chart_w / (len(months) - 1)

    grid = []
    for tick in [58, 60, 62, 64, 66, 68, 70]:
        y = point_y(tick, top, chart_h, v_min, v_max)
        grid.append(f'<line x1="{left}" y1="{y:.2f}" x2="{left + chart_w}" y2="{y:.2f}" stroke="{GRID}" stroke-dasharray="4 8"/>')
        grid.append(svg_text(left - 18, y + 5, fmt_pct(tick), size=12, fill=MUTED, anchor="end"))

    labels = []
    for idx, month in enumerate(months):
        x = left + idx * x_step
        labels.append(f'<line x1="{x:.2f}" y1="{top}" x2="{x:.2f}" y2="{top + chart_h}" stroke="#eef2f6"/>')
        labels.append(svg_text(x, top + chart_h + 28, month, size=13, weight=600, fill=MUTED, anchor="middle"))

    circles = []
    for values, color in [(budget, GREEN), (dlr, BLUE), (post, ORANGE)]:
        for idx, value in enumerate(values):
            x = left + idx * x_step
            y = point_y(value, top, chart_h, v_min, v_max)
            circles.append(f'<circle cx="{x:.2f}" cy="{y:.2f}" r="6" fill="{color}" stroke="{CARD}" stroke-width="3"/>')

    point_labels = []
    for idx, value in enumerate(dlr):
        x = left + idx * x_step
        y = point_y(value, top, chart_h, v_min, v_max)
        point_labels.append(svg_text(x, y - 14, fmt_pct(value), size=11, weight=700, fill=BLUE, anchor="middle"))

    point_labels.append(svg_text(left + 2 * x_step, point_y(post[2], top, chart_h, v_min, v_max) - 14,
                                 "Scenario 65.7%", size=11, weight=700, fill=ORANGE, anchor="middle"))

    return f"""<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}">
  <rect width="{width}" height="{height}" fill="{BG}"/>
  {svg_text(72, 66, "Nocatee Occupancy vs Budget", size=32, weight=700, fill=INK)}
  {svg_text(72, 96, "Uses the March 24, 2026 DLR budget path and a clearly labeled post-DLR scenario for April and May signings.", size=15, fill=MUTED)}

  {stat_card(72, 126, 290, 116, "Current Occupied (03/24/2026)", "60.11%", "Ahead of March budget by 81 bps", BLUE)}
  {stat_card(386, 126, 290, 116, "Current Leased (03/24/2026)", "65.73%", "10 leased-not-occupied units in the cushion", TEAL)}
  {stat_card(700, 126, 290, 116, "April Budget vs DLR", "+0.76 pts", "62.36% projected vs 61.60% budget", GREEN)}
  {stat_card(1014, 126, 314, 116, "May Budget Gap", "2.62 pts", "Before post-03/24 signings were added", RED)}

  <rect x="72" y="260" width="1256" height="486" rx="24" fill="{CARD}" stroke="#dfe6ee"/>
  {svg_text(96, 302, "Budget line vs live DLR projection", size=19, weight=700, fill=INK)}
  {svg_text(96, 328, "The data supports a March/April budget defense. The real pressure point is May and later if fresh leasing does not continue.", size=13, fill=MUTED)}

  {''.join(grid)}
  {''.join(labels)}
  <path d="{line_path(budget, left, top, chart_w, chart_h, v_min, v_max)}" fill="none" stroke="{GREEN}" stroke-width="4"/>
  <path d="{line_path(dlr, left, top, chart_w, chart_h, v_min, v_max)}" fill="none" stroke="{BLUE}" stroke-width="4"/>
  <path d="{line_path(post, left, top, chart_w, chart_h, v_min, v_max)}" fill="none" stroke="{ORANGE}" stroke-width="4" stroke-dasharray="10 10"/>
  {''.join(circles)}
  {''.join(point_labels)}

  <rect x="100" y="638" width="16" height="16" rx="4" fill="{GREEN}"/>
  {svg_text(126, 651, "Monthly budget occupancy", size=12, fill=MUTED)}
  <rect x="306" y="638" width="16" height="16" rx="4" fill="{BLUE}"/>
  {svg_text(332, 651, "March 24 DLR projection", size=12, fill=MUTED)}
  <rect x="532" y="638" width="16" height="16" rx="4" fill="{ORANGE}"/>
  {svg_text(558, 651, "Post-03/24 scenario if the 4 visible April/May signings were incremental", size=12, fill=MUTED)}

  <rect x="96" y="770" width="1208" height="82" rx="18" fill="#fff8ef" stroke="#f0d8b2"/>
  {svg_text(120, 804, "Key read", size=16, weight=700, fill=GOLD)}
  {svg_text(120, 832, "On March 24, 2026 the property was ahead of March budget and projected ahead of April budget. May was the first clear forward gap, and even that gap narrows materially once the post-DLR signings are layered in.", size=14, fill=TEXT)}
</svg>
"""


def funnel_svg() -> str:
    width = 1400
    height = 860
    metrics = [
        ("Leads", 160, 174),
        ("Tours", 70, 85),
        ("Applications", 12, 17),
    ]
    max_value = 190
    left = 130
    top = 280
    chart_w = 1120
    chart_h = 420
    group_gap = chart_w / len(metrics)
    bar_w = 130

    lead_to_tour_2025 = 43.75
    lead_to_tour_2026 = 48.85
    tour_to_app_2025 = 17.14
    tour_to_app_2026 = 20.00

    grid = []
    for tick in [0, 40, 80, 120, 160]:
        y = top + chart_h - (tick / max_value) * chart_h
        grid.append(f'<line x1="{left}" y1="{y:.2f}" x2="{left + chart_w}" y2="{y:.2f}" stroke="{GRID}" stroke-dasharray="4 8"/>')
        grid.append(svg_text(left - 16, y + 5, str(tick), size=12, fill=MUTED, anchor="end"))

    bars = []
    labels = []
    for idx, (name, v2025, v2026) in enumerate(metrics):
        group_x = left + idx * group_gap + 80
        h_2025 = (v2025 / max_value) * chart_h
        h_2026 = (v2026 / max_value) * chart_h
        x1 = group_x
        x2 = group_x + bar_w + 18
        y1 = top + chart_h - h_2025
        y2 = top + chart_h - h_2026
        bars.append(f'<rect x="{x1}" y="{y1:.2f}" width="{bar_w}" height="{h_2025:.2f}" rx="18" fill="{SLATE}" opacity="0.95"/>')
        bars.append(f'<rect x="{x2}" y="{y2:.2f}" width="{bar_w}" height="{h_2026:.2f}" rx="18" fill="{BLUE}" opacity="0.95"/>')
        labels.append(svg_text(x1 + bar_w / 2, top + chart_h + 26, "Q1 2025", size=12, fill=MUTED, anchor="middle"))
        labels.append(svg_text(x2 + bar_w / 2, top + chart_h + 26, "YTD 2026", size=12, fill=MUTED, anchor="middle"))
        labels.append(svg_text(group_x + bar_w + 9, top + chart_h + 56, name, size=14, weight=700, fill=INK, anchor="middle"))
        labels.append(svg_text(x1 + bar_w / 2, y1 - 12, str(v2025), size=12, weight=700, fill=SLATE, anchor="middle"))
        labels.append(svg_text(x2 + bar_w / 2, y2 - 12, str(v2026), size=12, weight=700, fill=BLUE, anchor="middle"))

    return f"""<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}">
  <rect width="{width}" height="{height}" fill="{BG}"/>
  {svg_text(72, 66, "Nocatee Leasing Funnel", size=32, weight=700, fill=INK)}
  {svg_text(72, 96, "Q1 2025 compared with year-to-date 2026 through late March. This is the cleanest visual for the 'lead volume and activity' rebuttal.", size=15, fill=MUTED)}

  {stat_card(72, 126, 300, 116, "Lead Growth", "+8.8%", "174 YTD 2026 vs 160 in Q1 2025", BLUE)}
  {stat_card(396, 126, 300, 116, "Tour Growth", "+21.4%", "85 YTD 2026 vs 70 in Q1 2025", TEAL)}
  {stat_card(720, 126, 300, 116, "Application Growth", "+41.7%", "17 YTD 2026 vs 12 in Q1 2025", GREEN)}
  {stat_card(1044, 126, 284, 116, "Lead Target", "+19.2%", "174 leads vs the 146 benchmark", ORANGE)}

  <rect x="72" y="260" width="1256" height="486" rx="24" fill="{CARD}" stroke="#dfe6ee"/>
  {svg_text(96, 302, "Top-of-funnel and mid-funnel volume", size=19, weight=700, fill=INK)}
  {svg_text(96, 328, "The volume story is better in 2026, which helps push back on the idea that demand generation itself is missing.", size=13, fill=MUTED)}

  {''.join(grid)}
  {''.join(bars)}
  {''.join(labels)}

  <rect x="98" y="610" width="220" height="92" rx="18" fill="#f8fbff" stroke="#dfe6ee"/>
  {svg_text(122, 642, "Lead to tour conversion", size=13, weight=600, fill=MUTED)}
  {svg_text(122, 680, f"{lead_to_tour_2025:.1f}% -> {lead_to_tour_2026:.1f}%", size=28, weight=700, fill=INK)}
  {svg_text(122, 704, "43.8% in 2025 vs 48.9% in 2026", size=12, fill=MUTED)}

  <rect x="348" y="610" width="220" height="92" rx="18" fill="#f8fbff" stroke="#dfe6ee"/>
  {svg_text(372, 642, "Tour to app conversion", size=13, weight=600, fill=MUTED)}
  {svg_text(372, 680, f"{tour_to_app_2025:.1f}% -> {tour_to_app_2026:.1f}%", size=28, weight=700, fill=INK)}
  {svg_text(372, 704, "17.1% in 2025 vs 20.0% in 2026", size=12, fill=MUTED)}

  <rect x="598" y="610" width="706" height="92" rx="18" fill="#f2fbf8" stroke="#cae8da"/>
  {svg_text(622, 642, "Suggested talking line", size=13, weight=700, fill=GREEN)}
  {svg_text(622, 676, "Traffic is not the missing piece at Nocatee. Through late March 2026, the property was ahead of Q1 lead target, ahead of Q1 2025 on leads and tours, and materially ahead on applications.", size=14, fill=TEXT)}
</svg>
"""


def pricing_response_svg() -> str:
    plans = [
        {
            "name": "Cypress",
            "leases": 5,
            "discount": "-5.8%",
            "action": "Strongest response after 02/20 pricing reset",
            "remaining": None,
            "accent": TEAL,
        },
        {
            "name": "Pelican",
            "leases": 0,
            "discount": "-7.5%",
            "action": "8 remaining; review again as lost leader",
            "remaining": "8 remaining",
            "accent": RED,
        },
        {
            "name": "Osprey",
            "leases": 1,
            "discount": "-14.5%",
            "action": "5 remaining; one response, but needs another review",
            "remaining": "5 remaining",
            "accent": ORANGE,
        },
        {
            "name": "Mayfair",
            "leases": 0,
            "discount": "-16.9%",
            "action": "7 remaining; guest suite merchandising goes live 04/06",
            "remaining": "7 remaining",
            "accent": GOLD,
        },
    ]

    width = 1400
    height = 900
    cards = []
    max_leases = 5
    for idx, plan in enumerate(plans):
        x = 72 + idx * 314
        y = 260
        bar_w = (plan["leases"] / max_leases) * 220
        remaining = plan["remaining"] or "DLR showed the strongest lease response in this set"
        cards.append(f"""
        <g>
          <rect x="{x}" y="{y}" width="282" height="520" rx="24" fill="{CARD}" stroke="#dfe6ee"/>
          <rect x="{x + 24}" y="{y + 24}" width="56" height="10" rx="5" fill="{plan['accent']}" opacity="0.9"/>
          {svg_text(x + 24, y + 68, plan["name"], size=28, weight=700, fill=INK)}
          {svg_text(x + 24, y + 98, "Leases since 02/20/2026 pricing change", size=12, fill=MUTED)}
          {svg_text(x + 24, y + 146, str(plan["leases"]), size=54, weight=700, fill=INK)}
          <rect x="{x + 24}" y="{y + 176}" width="220" height="18" rx="9" fill="#edf2f7"/>
          <rect x="{x + 24}" y="{y + 176}" width="{bar_w}" height="18" rx="9" fill="{plan['accent']}"/>
          {svg_text(x + 24, y + 236, "Current DLR discount to budgeted in-place rent", size=12, fill=MUTED)}
          {svg_text(x + 24, y + 278, plan["discount"], size=34, weight=700, fill=plan['accent'])}
          {svg_text(x + 24, y + 334, remaining, size=15, weight=600, fill=TEXT)}
          <rect x="{x + 24}" y="{y + 372}" width="234" height="116" rx="18" fill="#f8fbff" stroke="#dfe6ee"/>
          {svg_text(x + 42, y + 410, "Operating read", size=13, weight=700, fill=MUTED)}
          {svg_text(x + 42, y + 446, plan["action"], size=15, fill=TEXT)}
        </g>
        """)

    return f"""<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}">
  <rect width="{width}" height="{height}" fill="{BG}"/>
  {svg_text(72, 66, "Pricing Response by Targeted Floor Plan", size=32, weight=700, fill=INK)}
  {svg_text(72, 96, "The visual case here is not that every price move worked. It is that pricing was active, and the response pattern is now clear by plan.", size=15, fill=MUTED)}

  {stat_card(72, 126, 360, 108, "Adjusted Set Leases Since 02/20", "6", "5 Cypress + 1 Osprey", BLUE)}
  {stat_card(456, 126, 420, 108, "What this disproves", "No-response narrative", "The changes produced measurable leasing response, even if uneven by plan", GREEN)}
  {stat_card(900, 126, 428, 108, "What still needs action", "Pelican and Mayfair", "Those are the clearest next-round pricing and merchandising pain points", RED)}

  {''.join(cards)}
</svg>
"""


def peer_svg() -> str:
    peers = [
        ("RISE at Nocatee", 69.7, 71.9, 28.1, BLUE),
        ("Olea eTown", 84.9, 85.4, 14.6, SLATE),
        ("RISE Glen Kernan", 13.0, 14.6, 85.4, SLATE),
        ("Olea Beach Haven", 96.6, 96.0, 4.0, SLATE),
    ]
    avg_leased = 66.1
    avg_pre = 67.0
    avg_exp = 33.0

    width = 1400
    height = 980
    left = 270
    bar_w = 900
    bar_h = 22

    sections = [
        ("Leased %", 260, 1, avg_leased, True),
        ("Pre-Leased %", 500, 2, avg_pre, True),
        ("Exposure %", 740, 3, avg_exp, False),
    ]

    groups = []
    for title, top, metric_idx, avg_value, higher_is_better in sections:
        groups.append(svg_text(72, top - 24, title, size=22, weight=700, fill=INK))
        groups.append(svg_text(72, top + 4, f"Active-adult average: {fmt_pct(avg_value)}", size=13, fill=MUTED))
        for idx, (name, leased, pre, exposure, color) in enumerate(peers):
            value = [leased, pre, exposure][metric_idx - 1]
            y = top + idx * 46
            width_scaled = (value / 100) * bar_w
            groups.append(svg_text(72, y + 16, name, size=14, weight=700 if name == "RISE at Nocatee" else 500,
                                   fill=INK if name == "RISE at Nocatee" else TEXT))
            groups.append(f'<rect x="{left}" y="{y}" width="{bar_w}" height="{bar_h}" rx="11" fill="#edf2f7"/>')
            groups.append(f'<rect x="{left}" y="{y}" width="{width_scaled:.2f}" height="{bar_h}" rx="11" fill="{color}" opacity="0.95"/>')
            groups.append(svg_text(left + width_scaled + 12, y + 16, fmt_pct(value), size=13, weight=700, fill=color))
        avg_x = left + (avg_value / 100) * bar_w
        groups.append(f'<line x1="{avg_x:.2f}" y1="{top - 8}" x2="{avg_x:.2f}" y2="{top + 160}" stroke="{GREEN}" stroke-width="2" stroke-dasharray="6 6"/>')
        label_y = top - 12 if higher_is_better else top + 168
        groups.append(svg_text(avg_x + 6, label_y, "Peer avg", size=11, weight=700, fill=GREEN))

    return f"""<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}">
  <rect width="{width}" height="{height}" fill="{BG}"/>
  {svg_text(72, 66, "Active-Adult Peer Positioning", size=32, weight=700, fill=INK)}
  {svg_text(72, 96, "This is the cleanest market-facing visual: Nocatee is above the active-adult average on leased and pre-leased percentage and below the average on exposure.", size=15, fill=MUTED)}

  {stat_card(72, 126, 390, 108, "Leased vs active-adult average", "+3.6 pts", "69.7% at Nocatee vs 66.1% average", BLUE)}
  {stat_card(486, 126, 390, 108, "Pre-leased vs active-adult average", "+4.9 pts", "71.9% at Nocatee vs 67.0% average", TEAL)}
  {stat_card(900, 126, 428, 108, "Exposure vs active-adult average", "-4.9 pts", "28.1% at Nocatee vs 33.0% average", GREEN)}

  {''.join(groups)}

  <rect x="72" y="898" width="1256" height="52" rx="18" fill="#f2fbf8" stroke="#cae8da"/>
  {svg_text(96, 930, "Suggested talking line: Nocatee is not the weakest active-adult lease-up in the set. It is outperforming the active-adult benchmark on both leased and pre-leased percentage while carrying less exposure than the peer average.", size=14, fill=TEXT)}
</svg>
"""


def one_slide_svg() -> str:
    months = ["Mar", "Apr", "May", "Jun"]
    budget = [59.30, 61.60, 66.10, 69.10]
    dlr = [60.11, 62.36, 63.48, 64.61]
    post = [60.11, 64.05, 65.73, 66.86]

    width = 1366
    height = 768
    chart_left = 74
    chart_top = 330
    chart_w = 748
    chart_h = 286
    v_min = 58
    v_max = 70
    x_step = chart_w / (len(months) - 1)

    def p_y(value: float) -> float:
        return point_y(value, chart_top, chart_h, v_min, v_max)

    grid = []
    for tick in [58, 60, 62, 64, 66, 68, 70]:
        y = p_y(tick)
        grid.append(f'<line x1="{chart_left}" y1="{y:.2f}" x2="{chart_left + chart_w}" y2="{y:.2f}" stroke="{GRID}" stroke-dasharray="4 8"/>')
        grid.append(svg_text(chart_left - 12, y + 4, fmt_pct(tick), size=11, fill=MUTED, anchor="end"))

    month_labels = []
    for idx, month in enumerate(months):
        x = chart_left + idx * x_step
        month_labels.append(f'<line x1="{x:.2f}" y1="{chart_top}" x2="{x:.2f}" y2="{chart_top + chart_h}" stroke="#edf2f7"/>')
        month_labels.append(svg_text(x, chart_top + chart_h + 24, month, size=13, weight=700, fill=MUTED, anchor="middle"))

    chart_marks = []
    for values, color in [(budget, GREEN), (dlr, BLUE), (post, ORANGE)]:
        for idx, value in enumerate(values):
            x = chart_left + idx * x_step
            y = p_y(value)
            chart_marks.append(f'<circle cx="{x:.2f}" cy="{y:.2f}" r="5" fill="{color}" stroke="{CARD}" stroke-width="3"/>')

    pricing_rows = [
        ("Cypress", "5", "-", "Best"),
        ("Pelican", "0", "8", "Review"),
        ("Osprey", "1", "5", "Review"),
        ("Mayfair", "0", "7", "Guest suite"),
    ]

    pricing_lines = []
    row_y_start = 574
    for idx, row in enumerate(pricing_rows):
        y = row_y_start + idx * 38
        pricing_lines.append(f'<line x1="934" y1="{y - 22}" x2="1288" y2="{y - 22}" stroke="#edf2f7"/>')
        pricing_lines.append(svg_text(950, y, row[0], size=14, weight=700, fill=INK))
        pricing_lines.append(svg_text(1098, y, row[1], size=14, weight=700, fill=BLUE, anchor="middle"))
        pricing_lines.append(svg_text(1186, y, row[2], size=14, weight=700, fill=TEXT, anchor="middle"))
        pricing_lines.append(svg_text(1272, y, row[3], size=13, fill=MUTED, anchor="end"))

    return f"""<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}">
  <rect width="{width}" height="{height}" fill="{BG}"/>
  <rect x="24" y="24" width="1318" height="720" rx="28" fill="{CARD}" stroke="#dfe6ee"/>

  {svg_text(54, 72, "RISE at Nocatee | Leasing Snapshot", size=28, weight=700, fill=INK)}
  {svg_text(54, 98, "DLR 03/24/2026 + post-03/24 signed move-ins + 2026 budget + pricing updates through 03/27/2026", size=13, fill=MUTED)}

  {stat_card(54, 122, 230, 78, "Mar vs budget", "+0.8 pts", "60.1% vs 59.3%", BLUE)}
  {stat_card(300, 122, 230, 78, "Apr vs budget", "+0.8 pts", "62.4% vs 61.6%", GREEN)}
  {stat_card(546, 122, 230, 78, "Leads vs target", "+19%", "174 vs 146", TEAL)}
  {stat_card(792, 122, 260, 78, "Price response", "6 leases", "5 Cypress, 1 Osprey", ORANGE)}
  {stat_card(1068, 122, 244, 78, "May scenario", "-0.4 pts", "65.7% vs 66.1%", RED)}

  <rect x="42" y="220" width="810" height="494" rx="22" fill="#fbfcfe" stroke="#e5ebf2"/>
  {svg_text(62, 254, "Occupancy vs budget", size=20, weight=700, fill=INK)}
  {svg_text(62, 276, "Budget, 03/24 DLR projection, and post-03/24 signed scenario", size=12, fill=MUTED)}
  {''.join(grid)}
  {''.join(month_labels)}
  <path d="{line_path(budget, chart_left, chart_top, chart_w, chart_h, v_min, v_max)}" fill="none" stroke="{GREEN}" stroke-width="4"/>
  <path d="{line_path(dlr, chart_left, chart_top, chart_w, chart_h, v_min, v_max)}" fill="none" stroke="{BLUE}" stroke-width="4"/>
  <path d="{line_path(post, chart_left, chart_top, chart_w, chart_h, v_min, v_max)}" fill="none" stroke="{ORANGE}" stroke-width="4" stroke-dasharray="8 8"/>
  {''.join(chart_marks)}
  {svg_text(chart_left + x_step * 1, p_y(dlr[1]) - 14, "62.4%", size=11, weight=700, fill=BLUE, anchor="middle")}
  {svg_text(chart_left + x_step * 2, p_y(post[2]) - 14, "65.7%", size=11, weight=700, fill=ORANGE, anchor="middle")}

  <rect x="62" y="640" width="14" height="14" rx="4" fill="{GREEN}"/>
  {svg_text(84, 652, "Budget", size=12, fill=MUTED)}
  <rect x="154" y="640" width="14" height="14" rx="4" fill="{BLUE}"/>
  {svg_text(176, 652, "03/24 DLR", size=12, fill=MUTED)}
  <rect x="270" y="640" width="14" height="14" rx="4" fill="{ORANGE}"/>
  {svg_text(292, 652, "Signed scenario", size=12, fill=MUTED)}

  <rect x="62" y="670" width="770" height="28" rx="14" fill="#f2fbf8" stroke="#cae8da"/>
  {svg_text(82, 689, "Ahead of Mar/Apr budget. May is the first forward gap.", size=12, fill=TEXT)}

  <rect x="878" y="220" width="434" height="212" rx="22" fill="#fbfcfe" stroke="#e5ebf2"/>
  {svg_text(900, 252, "Funnel: 2026 vs Q1 2025", size=20, weight=700, fill=INK)}
  {svg_text(1068, 294, "Q1 2025", size=11, weight=700, fill=MUTED, anchor="middle")}
  {svg_text(1172, 294, "YTD 2026", size=11, weight=700, fill=MUTED, anchor="middle")}
  {svg_text(1282, 294, "Delta", size=11, weight=700, fill=MUTED, anchor="end")}
  {svg_text(918, 336, "Leads", size=15, weight=700, fill=TEXT)}
  {svg_text(1068, 336, "160", size=15, weight=700, fill=SLATE, anchor="middle")}
  {svg_text(1172, 336, "174", size=15, weight=700, fill=BLUE, anchor="middle")}
  {svg_text(1282, 336, "+8.8%", size=15, weight=700, fill=GREEN, anchor="end")}
  {svg_text(918, 380, "Tours", size=15, weight=700, fill=TEXT)}
  {svg_text(1068, 380, "70", size=15, weight=700, fill=SLATE, anchor="middle")}
  {svg_text(1172, 380, "85", size=15, weight=700, fill=BLUE, anchor="middle")}
  {svg_text(1282, 380, "+21.4%", size=15, weight=700, fill=GREEN, anchor="end")}
  {svg_text(918, 424, "Apps", size=15, weight=700, fill=TEXT)}
  {svg_text(1068, 424, "12", size=15, weight=700, fill=SLATE, anchor="middle")}
  {svg_text(1172, 424, "17", size=15, weight=700, fill=BLUE, anchor="middle")}
  {svg_text(1282, 424, "+41.7%", size=15, weight=700, fill=GREEN, anchor="end")}

  <rect x="878" y="450" width="434" height="264" rx="22" fill="#fbfcfe" stroke="#e5ebf2"/>
  {svg_text(900, 482, "Pricing response since 02/20", size=20, weight=700, fill=INK)}
  <line x1="934" y1="522" x2="1288" y2="522" stroke="#edf2f7"/>
  {svg_text(950, 540, "Plan", size=11, weight=700, fill=MUTED)}
  {svg_text(1098, 540, "Leased", size=11, weight=700, fill=MUTED, anchor="middle")}
  {svg_text(1186, 540, "Left", size=11, weight=700, fill=MUTED, anchor="middle")}
  {svg_text(1272, 540, "Next step", size=11, weight=700, fill=MUTED, anchor="end")}
  {''.join(pricing_lines)}
</svg>
"""


def dashboard_html() -> str:
    return """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>Nocatee Investor Visual Pack</title>
  <style>
    :root {
      --bg: #f4f6fb;
      --card: #ffffff;
      --text: #163047;
      --muted: #5f7080;
      --border: #dfe6ee;
      --blue: #2f6bff;
      --teal: #22b8b3;
      --green: #2fa26b;
      --orange: #f08a24;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "Inter", "Segoe UI", Arial, sans-serif;
      background: radial-gradient(circle at top left, #ffffff 0%, var(--bg) 55%, #ecf1f8 100%);
      color: var(--text);
    }
    .wrap {
      max-width: 1320px;
      margin: 0 auto;
      padding: 28px 20px 56px;
    }
    .hero {
      background: linear-gradient(135deg, #183047 0%, #274967 58%, #2f6bff 100%);
      color: #fff;
      border-radius: 28px;
      padding: 28px 30px;
      box-shadow: 0 24px 48px rgba(20, 44, 68, 0.14);
      margin-bottom: 20px;
    }
    .hero h1 {
      margin: 0 0 10px;
      font-size: 34px;
      line-height: 1.05;
    }
    .hero p {
      margin: 0;
      max-width: 900px;
      color: rgba(255, 255, 255, 0.84);
      font-size: 16px;
      line-height: 1.5;
    }
    .summary {
      display: grid;
      grid-template-columns: repeat(4, 1fr);
      gap: 14px;
      margin: 20px 0 24px;
    }
    .pill {
      background: var(--card);
      border: 1px solid var(--border);
      border-radius: 22px;
      padding: 16px 18px;
      box-shadow: 0 16px 30px rgba(28, 47, 68, 0.06);
    }
    .pill .label {
      color: var(--muted);
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0.08em;
      margin-bottom: 8px;
    }
    .pill .value {
      font-size: 28px;
      font-weight: 700;
      margin-bottom: 4px;
    }
    .pill .sub {
      color: var(--muted);
      font-size: 13px;
      line-height: 1.4;
    }
    .grid {
      display: grid;
      grid-template-columns: 1fr;
      gap: 18px;
    }
    .panel {
      background: var(--card);
      border: 1px solid var(--border);
      border-radius: 28px;
      padding: 18px;
      box-shadow: 0 16px 30px rgba(28, 47, 68, 0.06);
    }
    .panel h2 {
      margin: 0 0 6px;
      font-size: 20px;
    }
    .panel p {
      margin: 0 0 14px;
      color: var(--muted);
      font-size: 14px;
      line-height: 1.5;
    }
    .panel img {
      width: 100%;
      display: block;
      border-radius: 20px;
      border: 1px solid var(--border);
      background: #fff;
    }
    .foot {
      margin-top: 22px;
      padding: 18px 20px;
      border-radius: 22px;
      background: #fff9ef;
      border: 1px solid #f0d8b2;
      color: var(--text);
      font-size: 14px;
      line-height: 1.6;
    }
    @media (max-width: 980px) {
      .summary { grid-template-columns: repeat(2, 1fr); }
    }
    @media (max-width: 640px) {
      .summary { grid-template-columns: 1fr; }
      .hero h1 { font-size: 28px; }
    }
  </style>
</head>
<body>
  <div class="wrap">
    <section class="hero">
      <h1>RISE at Nocatee Investor Visual Pack</h1>
      <p>
        Four visuals built from the reconciled March 2026 DLR, budget path, funnel counts,
        pricing response notes, and active-adult peer data. These are designed to support an
        investor-facing narrative, not to replace the source reports.
      </p>
    </section>

    <section class="summary">
      <div class="pill">
        <div class="label">Current Occupancy</div>
        <div class="value">60.11%</div>
        <div class="sub">Ahead of March budget as of 03/24/2026</div>
      </div>
      <div class="pill">
        <div class="label">Current Leased</div>
        <div class="value">65.73%</div>
        <div class="sub">Shows a 10-unit leased-not-occupied cushion</div>
      </div>
      <div class="pill">
        <div class="label">Lead Volume</div>
        <div class="value">174</div>
        <div class="sub">19.2% above the 146-lead benchmark</div>
      </div>
      <div class="pill">
        <div class="label">Adjusted Set Response</div>
        <div class="value">6 leases</div>
        <div class="sub">5 Cypress and 1 Osprey since 02/20/2026</div>
      </div>
    </section>

    <section class="grid">
      <article class="panel">
        <h2>1. Occupancy vs budget</h2>
        <p>The fastest way to rebut the “already behind” narrative is to anchor on month-end budget definitions and the March 24 DLR projection.</p>
        <img src="nocatee_occupancy_budget_vs_projection.svg" alt="Occupancy vs budget chart"/>
      </article>

      <article class="panel">
        <h2>2. Funnel strength</h2>
        <p>This visual makes the demand-generation case quickly: 2026 is stronger than 2025 on the top and middle of the funnel.</p>
        <img src="nocatee_funnel_yoy.svg" alt="Leasing funnel year over year"/>
      </article>

      <article class="panel">
        <h2>3. Pricing response by floor plan</h2>
        <p>The goal here is not to prove every pricing move worked. It is to show pricing has been active and the next actions are already identified.</p>
        <img src="nocatee_pricing_response.svg" alt="Pricing response by floor plan"/>
      </article>

      <article class="panel">
        <h2>4. Active-adult peer position</h2>
        <p>This is the market check: Nocatee is above the active-adult average on leased and pre-leased percentage and below average on exposure.</p>
        <img src="nocatee_active_adult_peer_positioning.svg" alt="Active adult peer positioning"/>
      </article>
    </section>

    <section class="foot">
      If you want a tighter board-room version next, the best upgrade would be a single-slide summary that pulls one chart, three callout bullets, and a small appendix table of the post-03/24 move-ins by floor plan and date.
    </section>
  </div>
</body>
</html>
"""


def main() -> None:
    OUTDIR.mkdir(parents=True, exist_ok=True)

    (OUTDIR / "nocatee_occupancy_budget_vs_projection.svg").write_text(occupancy_svg())
    (OUTDIR / "nocatee_funnel_yoy.svg").write_text(funnel_svg())
    (OUTDIR / "nocatee_pricing_response.svg").write_text(pricing_response_svg())
    (OUTDIR / "nocatee_active_adult_peer_positioning.svg").write_text(peer_svg())
    (OUTDIR / "nocatee_one_slide.svg").write_text(one_slide_svg())
    (OUTDIR / "nocatee_investor_visual_pack.html").write_text(dashboard_html())


if __name__ == "__main__":
    main()
