from __future__ import annotations

from pathlib import Path


WORKDIR = Path("/Users/jacheflin/Documents/Playground")
SVG_PATH = WORKDIR / "first-year-lease-velocity-comparison.svg"
NOTES_PATH = WORKDIR / "first-year-lease-velocity-notes.md"

# Hand-traced from the screenshot labels the user provided.
# To keep the comparison fair, the chart uses the first 44 observed weeks
# for both communities because that is the common legible window.
VIERA = [
    1, 0, 3, 4, 1, 1, 4, 3, 2, 1, 6, 1, 1, 2, 2, 2, 4, 3, 1, 3, 1, 0,
    0, 0, 0, 0, 1, 1, 2, 2, 9, 1, 4, 3, 1, 4, 1, 3, 1, 0, 1, 1, 0, 0,
]
PARASOL = [
    0, 3, 2, 2, 0, 2, 1, 2, 1, 1, 1, 1, 0, 2, 0, 0, 1, 0, 1, 4, 1, 1,
    0, 1, 2, 1, 0, 2, 0, 0, 1, 0, 1, 2, 2, 1, 1, 3, 2, 2, 2, 1, 0, 5,
]


def cumulative(values: list[int]) -> list[int]:
    total = 0
    output = []
    for value in values:
        total += value
        output.append(total)
    return output


def line_path(values: list[float], left: float, top: float, width: float, height: float, v_max: float) -> str:
    points = []
    x_step = width / (len(values) - 1)
    for idx, value in enumerate(values):
        x = left + idx * x_step
        y = top + height - (value / v_max) * height
        points.append((x, y))
    return "M " + " L ".join(f"{x:.2f} {y:.2f}" for x, y in points)


def circle_points(values: list[float], left: float, top: float, width: float, height: float, v_max: float, color: str) -> str:
    parts = []
    x_step = width / (len(values) - 1)
    for idx, value in enumerate(values):
        x = left + idx * x_step
        y = top + height - (value / v_max) * height
        parts.append(f'<circle cx="{x:.2f}" cy="{y:.2f}" r="2.5" fill="{color}" />')
    return "\n".join(parts)


def bars(values: list[int], left: float, top: float, width: float, height: float, v_max: float, color: str, offset: float) -> str:
    bar_gap = width / len(values)
    bar_width = max(bar_gap * 0.38, 4)
    parts = []
    for idx, value in enumerate(values):
        x = left + idx * bar_gap + offset
        bar_height = (value / v_max) * height
        y = top + height - bar_height
        parts.append(
            f'<rect x="{x:.2f}" y="{y:.2f}" width="{bar_width:.2f}" height="{bar_height:.2f}" '
            f'rx="2" fill="{color}" opacity="0.8" />'
        )
    return "\n".join(parts)


def axis_labels(left: float, top: float, width: float, height: float, y_ticks: list[int], x_labels: list[tuple[int, str]], v_max: float) -> str:
    parts = []
    for tick in y_ticks:
        y = top + height - (tick / v_max) * height
        parts.append(f'<line x1="{left}" y1="{y:.2f}" x2="{left + width}" y2="{y:.2f}" stroke="#d7dee6" stroke-dasharray="4 6" />')
        parts.append(f'<text x="{left - 10}" y="{y + 4:.2f}" text-anchor="end" font-size="11" fill="#617180">{tick}</text>')
    for week, label in x_labels:
        x = left + ((week - 1) / (44 - 1)) * width
        parts.append(f'<line x1="{x:.2f}" y1="{top}" x2="{x:.2f}" y2="{top + height}" stroke="#eef2f6" />')
        parts.append(f'<text x="{x:.2f}" y="{top + height + 22}" text-anchor="middle" font-size="11" fill="#617180">{label}</text>')
    return "\n".join(parts)


def stat_card(x: int, y: int, w: int, h: int, title: str, left_value: str, right_value: str) -> str:
    return f"""
    <g>
      <rect x="{x}" y="{y}" width="{w}" height="{h}" rx="12" fill="#f7f9fb" stroke="#d9e0e7" />
      <text x="{x + 16}" y="{y + 24}" font-size="12" fill="#617180">{title}</text>
      <text x="{x + 16}" y="{y + 52}" font-size="24" font-weight="700" fill="#4f5bd5">{left_value}</text>
      <text x="{x + 16}" y="{y + 70}" font-size="11" fill="#617180">RISE Viera</text>
      <text x="{x + w - 16}" y="{y + 52}" text-anchor="end" font-size="24" font-weight="700" fill="#5ecfa4">{right_value}</text>
      <text x="{x + w - 16}" y="{y + 70}" text-anchor="end" font-size="11" fill="#617180">Parasol</text>
    </g>
    """


def main() -> None:
    viera_cum = cumulative(VIERA)
    parasol_cum = cumulative(PARASOL)

    weekly_max = max(max(VIERA), max(PARASOL))
    cumulative_max = max(viera_cum[-1], parasol_cum[-1])

    svg = f"""<svg xmlns="http://www.w3.org/2000/svg" width="1400" height="980" viewBox="0 0 1400 980">
  <rect width="1400" height="980" fill="#ffffff" />
  <text x="64" y="64" font-size="30" font-weight="700" fill="#183247">First-Year Lease Velocity Comparison</text>
  <text x="64" y="92" font-size="14" fill="#617180">Reconstructed from the screenshot labels provided. Fair comparison uses the first 44 observed weeks for both communities.</text>
  <text x="64" y="116" font-size="13" fill="#617180">Start points used: RISE Viera 7/22/2024 and Parasol Melbourne 4/24/2023.</text>

  {stat_card(64, 144, 300, 92, "Leases Across First 44 Observed Weeks", str(sum(VIERA)), str(sum(PARASOL)))}
  {stat_card(384, 144, 300, 92, "Peak Weekly Leases", str(max(VIERA)), str(max(PARASOL)))}
  {stat_card(704, 144, 300, 92, "Weeks With 3+ Leases", str(sum(1 for value in VIERA if value >= 3)), str(sum(1 for value in PARASOL if value >= 3)))}
  {stat_card(1024, 144, 300, 92, "Weeks With 4+ Leases", str(sum(1 for value in VIERA if value >= 4)), str(sum(1 for value in PARASOL if value >= 4)))}

  <text x="64" y="286" font-size="18" font-weight="700" fill="#183247">Weekly Leases Since Opening</text>
  <text x="64" y="306" font-size="12" fill="#617180">Weekly bars show the relative intensity of leasing activity in the comparable first-year window.</text>
  {axis_labels(84, 330, 1230, 220, [0, 2, 4, 6, 8, 10], [(1, "Wk 1"), (8, "Wk 8"), (16, "Wk 16"), (24, "Wk 24"), (32, "Wk 32"), (40, "Wk 40"), (44, "Wk 44")], 10)}
  {bars(VIERA, 90, 330, 1220, 220, 10, "#4f5bd5", 0)}
  {bars(PARASOL, 90, 330, 1220, 220, 10, "#5ecfa4", 11)}

  <rect x="970" y="290" width="14" height="14" rx="3" fill="#4f5bd5" />
  <text x="992" y="302" font-size="12" fill="#617180">RISE Viera</text>
  <rect x="1088" y="290" width="14" height="14" rx="3" fill="#5ecfa4" />
  <text x="1110" y="302" font-size="12" fill="#617180">Parasol Melbourne</text>

  <text x="64" y="612" font-size="18" font-weight="700" fill="#183247">Cumulative Leases Since Opening</text>
  <text x="64" y="632" font-size="12" fill="#617180">This view makes the pace difference easier to see: Viera separates from Parasol early and stays ahead.</text>
  {axis_labels(84, 656, 1230, 220, [0, 20, 40, 60, 80, 100], [(1, "Wk 1"), (8, "Wk 8"), (16, "Wk 16"), (24, "Wk 24"), (32, "Wk 32"), (40, "Wk 40"), (44, "Wk 44")], 100)}
  <path d="{line_path(viera_cum, 84, 656, 1230, 220, 100)}" fill="none" stroke="#4f5bd5" stroke-width="4" />
  <path d="{line_path(parasol_cum, 84, 656, 1230, 220, 100)}" fill="none" stroke="#5ecfa4" stroke-width="4" />
  {circle_points(viera_cum, 84, 656, 1230, 220, 100, "#4f5bd5")}
  {circle_points(parasol_cum, 84, 656, 1230, 220, 100, "#5ecfa4")}

  <text x="1318" y="{656 + 220 - (viera_cum[-1] / 100) * 220 + 4:.2f}" font-size="12" fill="#4f5bd5" text-anchor="end">{viera_cum[-1]}</text>
  <text x="1318" y="{656 + 220 - (parasol_cum[-1] / 100) * 220 - 8:.2f}" font-size="12" fill="#5ecfa4" text-anchor="end">{parasol_cum[-1]}</text>

  <rect x="64" y="912" width="1270" height="36" rx="10" fill="#f7f9fb" stroke="#d9e0e7" />
  <text x="82" y="935" font-size="12" fill="#617180">Use this as a directional call visual, not as a system-export exhibit. If you can pull the raw first-year lease export for both communities, I can replace the reconstruction with an exact chart.</text>
</svg>
"""

    SVG_PATH.write_text(svg)

    notes = """# First-Year Lease Velocity Notes

This visual compares the first 44 observed weeks for RISE Viera and Parasol Melbourne using the screenshot labels provided in the thread.

Why 44 weeks:
- That is the common legible first-year window I could reconstruct from both screenshots without inventing missing values.

Directional read:
- RISE Viera: 81 leases across the first 44 observed weeks, peak week of 9, 13 weeks with 3+ leases, and 7 weeks with 4+ leases.
- Parasol Melbourne: 55 leases across the first 44 observed weeks, peak week of 5, 4 weeks with 3+ leases, and 2 weeks with 4+ leases.

Suggested talking line:
- Based on the comparable first-year window visible in the historical charts, Viera is leasing faster than Parasol, not slower.
"""
    NOTES_PATH.write_text(notes)


if __name__ == "__main__":
    main()
