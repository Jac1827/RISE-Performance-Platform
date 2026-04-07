from __future__ import annotations

import json
import math
from pathlib import Path
from xml.sax.saxutils import escape as xml_escape
from zipfile import ZIP_DEFLATED, ZipFile


WORKDIR = Path("/Users/jacheflin/Documents/Playground")
JSON_PATH = Path("/Users/jacheflin/Downloads/rise_vp_weekly_snapshot_2026-03-13.json")
HTML_PATH = WORKDIR / "active-adult-performance-brief.html"
RTF_PATH = WORKDIR / "active-adult-performance-brief.rtf"
DOCX_PATH = WORKDIR / "active-adult-performance-brief.docx"

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
UNITS = {
    "Viera": 166,
    "Nocatee": 178,
    "Glen Kernan Park": 308,
}


def pct(value: float) -> str:
    return f"{value:.1f}%"


def whole(value: float) -> str:
    return f"{value:,.0f}"


def esc(text: str) -> str:
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def rtf_escape(text: str) -> str:
    return (
        str(text)
        .replace("\\", r"\\")
        .replace("{", r"\{")
        .replace("}", r"\}")
        .replace("\n", r"\line ")
    )


def hyperlink(url: str, label: str) -> str:
    return (
        r"{\field{\*\fldinst{HYPERLINK \"" + rtf_escape(url) + r"\"}}"
        r"{\fldrslt{\ul\cf2 " + rtf_escape(label) + r"}}}"
    )


def para(text: str = "", *, size: int = 22, bold: bool = False, center: bool = False, space_after: int = 120) -> str:
    align = r"\qc" if center else r"\ql"
    bold_on = r"\b " if bold else ""
    bold_off = r"\b0" if bold else ""
    return rf"\pard{align}\sa{space_after}\f0\fs{size} {bold_on}{rtf_escape(text)}{bold_off}\par" + "\n"


def rich_para(parts: list[str], *, size: int = 22, bold: bool = False, center: bool = False, space_after: int = 120) -> str:
    align = r"\qc" if center else r"\ql"
    prefix = rf"\pard{align}\sa{space_after}\f0\fs{size} "
    if bold:
        prefix += r"\b "
    suffix = r"\b0" if bold else ""
    return prefix + "".join(parts) + suffix + r"\par" + "\n"


def table_row(cells: list[str], widths: list[int], *, header: bool = False) -> str:
    pieces = [r"\trowd\trgaph70\trleft0"]
    edge = 0
    for width in widths:
        edge += width
        pieces.append(rf"\cellx{edge}")
    for cell in cells:
        cell_text = rtf_escape(cell)
        style_on = r"\b " if header else ""
        style_off = r"\b0" if header else ""
        pieces.append(rf"\intbl\fs18 {style_on}{cell_text}{style_off}\cell")
    pieces.append(r"\row")
    return "".join(pieces) + "\n"


def community_metrics(name: str, entry: dict) -> dict:
    units = UNITS[name]
    current_month = int(entry.get("currentMonth", 2))
    row = (entry.get("monthlyData") or [{}])[current_month]
    occupied = int(entry.get("currentOccupied") or 0)
    leased = int(entry.get("currentLeased") or 0)
    occupancy_goal = float(entry.get("occupancyGoal") or 0)
    leased_goal = float(entry.get("leasedGoal") or 0)
    occ_pct = occupied / units * 100 if units else 0
    leased_pct = leased / units * 100 if units else 0
    move_outs = int(row.get("moveOuts") or 0)
    move_ins = int(row.get("moveIns") or 0)
    guest_cards = int(row.get("guestCards") or 0)
    tours = int(row.get("tours") or 0)
    applications = int(row.get("applications") or 0)
    leases_signed = int(row.get("leasesSignedActual") or 0)
    q1_absorption = 0
    for quarter_row in (entry.get("monthlyData") or [])[:3]:
        q1_absorption += int(quarter_row.get("moveIns") or 0) - int(quarter_row.get("moveOuts") or 0)
    return {
        "name": name,
        "investor": entry.get("investorName") or "",
        "month_label": MONTHS[current_month],
        "units": units,
        "occupied": occupied,
        "leased": leased,
        "occupancy_goal": occupancy_goal,
        "leased_goal": leased_goal,
        "occupancy_pct": occ_pct,
        "leased_pct": leased_pct,
        "occupancy_gap": occ_pct - occupancy_goal,
        "leased_gap": leased_pct - leased_goal,
        "units_to_occ_goal": max(math.ceil(units * occupancy_goal / 100) - occupied, 0),
        "units_to_leased_goal": max(math.ceil(units * leased_goal / 100) - leased, 0),
        "move_outs": move_outs,
        "move_ins": move_ins,
        "month_absorption": move_ins - move_outs,
        "q1_absorption": q1_absorption,
        "guest_cards": guest_cards,
        "tours": tours,
        "applications": applications,
        "leases_signed": leases_signed,
        "lead_to_tour": (tours / guest_cards * 100) if guest_cards else None,
        "tour_to_app": (applications / tours * 100) if tours else None,
        "lead_to_lease": (leases_signed / guest_cards * 100) if guest_cards else None,
        "app_to_lease": (leases_signed / applications * 100) if applications else None,
    }


def format_conv(value: float | None) -> str:
    return "—" if value is None else pct(value)


def docx_run(text: str, *, bold: bool = False, size: int = 22, color: str | None = None, underline: bool = False) -> str:
    props = []
    if bold:
        props.append("<w:b/>")
    if underline:
        props.append('<w:u w:val="single"/>')
    if color:
        props.append(f'<w:color w:val="{color}"/>')
    props.append(f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>')
    rpr = f"<w:rPr>{''.join(props)}</w:rPr>" if props else ""
    return f'<w:r>{rpr}<w:t xml:space="preserve">{xml_escape(text)}</w:t></w:r>'


def docx_para(runs: list[str], *, center: bool = False, spacing_after: int = 120) -> str:
    jc = '<w:jc w:val="center"/>' if center else ""
    return f'<w:p><w:pPr>{jc}<w:spacing w:after="{spacing_after}"/></w:pPr>{"".join(runs)}</w:p>'


def docx_cell(text: str, width: int, *, bold: bool = False) -> str:
    paragraph = docx_para([docx_run(text, bold=bold, size=18)], spacing_after=0)
    return f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/></w:tcPr>{paragraph}</w:tc>'


def docx_table(rows: list[list[str]], widths: list[int]) -> str:
    grid = "".join(f'<w:gridCol w:w="{width}"/>' for width in widths)
    border = (
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="6" w:space="0" w:color="C9D6DF"/>'
        '<w:left w:val="single" w:sz="6" w:space="0" w:color="C9D6DF"/>'
        '<w:bottom w:val="single" w:sz="6" w:space="0" w:color="C9D6DF"/>'
        '<w:right w:val="single" w:sz="6" w:space="0" w:color="C9D6DF"/>'
        '<w:insideH w:val="single" w:sz="6" w:space="0" w:color="C9D6DF"/>'
        '<w:insideV w:val="single" w:sz="6" w:space="0" w:color="C9D6DF"/>'
        '</w:tblBorders>'
    )
    row_xml = []
    for idx, row in enumerate(rows):
        cells = "".join(docx_cell(cell, width, bold=idx == 0) for cell, width in zip(row, widths))
        row_xml.append(f"<w:tr>{cells}</w:tr>")
    return f'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>{border}</w:tblPr><w:tblGrid>{grid}</w:tblGrid>{"".join(row_xml)}</w:tbl>'


def build_docx(document_xml: str, hyperlink_rels: list[tuple[str, str]]) -> None:
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>"""
    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"""
    doc_rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    for rel_id, url in hyperlink_rels:
        doc_rels.append(
            f'<Relationship Id="{rel_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
            f'Target="{xml_escape(url)}" TargetMode="External"/>'
        )
    doc_rels.append("</Relationships>")
    app_props = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Codex</Application>
</Properties>"""
    core_props = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:dcmitype="http://purl.org/dc/dcmitype/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Active Adult Performance and Greystar Counter Brief</dc:title>
  <dc:creator>Codex</dc:creator>
</cp:coreProperties>"""
    with ZipFile(DOCX_PATH, "w", compression=ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", root_rels)
        archive.writestr("docProps/app.xml", app_props)
        archive.writestr("docProps/core.xml", core_props)
        archive.writestr("word/document.xml", document_xml)
        archive.writestr("word/_rels/document.xml.rels", "".join(doc_rels))


def main() -> None:
    data = json.loads(JSON_PATH.read_text())
    portfolio = data["portfolioData"]
    metrics = [community_metrics(name, portfolio[name]) for name in ["Viera", "Nocatee", "Glen Kernan Park"]]

    exported_at = data.get("exportedAt", "")
    detail_map = {
        "Viera": "Updated March snapshot shows Viera slightly behind March budget but still within a manageable unit gap for a new active-adult lease-up.",
        "Nocatee": "Nocatee is further off its current budget curve, but it remains in early active-adult lease-up rather than a mature stabilized position.",
        "Glen Kernan Park": "Glen Kernan Park is still a phased-delivery community, so its performance should be framed as an in-delivery lease-up, not a stabilized comp.",
    }

    html = f"""<!DOCTYPE html><html><body><h1>Active Adult Performance and Greystar Counter Brief</h1><p>Prepared March 13, 2026 using the internal snapshot export dated {esc(exported_at)} and current public market sources.</p></body></html>"""
    HTML_PATH.write_text(html)

    table_widths = [1600, 900, 1100, 1100, 1700, 2500, 1800]
    rtf_parts = [
        r"{\rtf1\ansi\deff0",
        r"{\fonttbl{\f0 Helvetica;}{\f1 Courier;}}",
        r"{\colortbl;\red31\green41\blue55;\red14\green94\blue168;\red92\green107\blue119;}",
        r"\paperw12240\paperh15840\margl720\margr720\margt720\margb720",
    ]
    rtf_parts.append(para("Active Adult Performance and Greystar Counter Brief", size=32, bold=True, center=True, space_after=80))
    rtf_parts.append(para(f"Prepared March 13, 2026 using the internal snapshot export dated {exported_at}.", size=18, center=True, space_after=180))
    rtf_parts.append(para("Bottom line: the updated snapshot shows three communities in different lease-up stages, not a clean case that the current operator lacks 55+ capability. One of the three assets is still in phased delivery, and Greystar does not appear to have meaningful active-adult operating depth in Jacksonville or Ponte Vedra.", size=22, space_after=180))

    rtf_parts.append(para("Executive Takeaways", size=26, bold=True, space_after=120))
    takeaways = [
        f"Updated Viera data shows {pct(metrics[0]['occupancy_pct'])} occupied and {pct(metrics[0]['leased_pct'])} leased, only {whole(metrics[0]['units_to_occ_goal'])} occupied units and {whole(metrics[0]['units_to_leased_goal'])} leased units off March targets.",
        "Nocatee and Glen Kernan Park should not be judged against stabilized 55+ comps. Nocatee remains an early lease-up and Glen Kernan Park remains a phased-delivery asset.",
        "NIC reported 91.9% active-adult occupancy in 4Q25 across the sector, while stabilized active-adult communities open at least two years averaged about 95% occupied.",
        "Industry sources note that upscale active-adult communities often lease more slowly because many prospects need to sell a home first, so raw traffic volume is a weak standalone scorecard.",
        "I could not verify any Greystar-managed active-adult rental community in Jacksonville or Ponte Vedra. The only current Greystar-managed public comp I confirmed in the broader scope is Parasol Melbourne.",
    ]
    for line in takeaways:
        rtf_parts.append(para(f"- {line}", size=22, space_after=70))

    rtf_parts.append(para("Internal Operating Snapshot", size=26, bold=True, space_after=100))
    rtf_parts.append(table_row(["Community", "Units", "Occ %", "Leased %", "Gap to Goal", "March Funnel", "Operating Read"], table_widths, header=True))
    for item in metrics:
        gap_text = f"{item['occupancy_gap']:+.1f} occ / {item['leased_gap']:+.1f} leased"
        funnel_text = f"{whole(item['guest_cards'])} GC / {whole(item['tours'])} tours / {whole(item['applications'])} apps / {whole(item['leases_signed'])} leases"
        read_text = f"{whole(item['units_to_occ_goal'])} occ and {whole(item['units_to_leased_goal'])} leased units to goal"
        rtf_parts.append(
            table_row(
                [
                    item["name"],
                    whole(item["units"]),
                    f"{pct(item['occupancy_pct'])} ({whole(item['occupied'])})",
                    f"{pct(item['leased_pct'])} ({whole(item['leased'])})",
                    gap_text,
                    funnel_text,
                    read_text,
                ],
                table_widths,
            )
        )
    rtf_parts.append(para("", size=8, space_after=60))

    rtf_parts.append(para("Property Detail", size=26, bold=True, space_after=100))
    for item in metrics:
        rtf_parts.append(para(item["name"], size=24, bold=True, space_after=70))
        detail_lines = [
            detail_map[item["name"]],
            f"{item['month_label']} snapshot: {whole(item['occupied'])} occupied ({pct(item['occupancy_pct'])}) and {whole(item['leased'])} leased ({pct(item['leased_pct'])}) on {whole(item['units'])} units.",
            f"Gap to current budget: {whole(item['units_to_occ_goal'])} occupied units and {whole(item['units_to_leased_goal'])} leased units to target.",
            f"Current funnel: {whole(item['guest_cards'])} guest cards, {whole(item['tours'])} tours, {whole(item['applications'])} applications, and {whole(item['leases_signed'])} signed leases.",
            f"Conversion view: lead-to-tour {format_conv(item['lead_to_tour'])}, tour-to-app {format_conv(item['tour_to_app'])}, lead-to-lease {format_conv(item['lead_to_lease'])}, app-to-lease {format_conv(item['app_to_lease'])}.",
            f"Absorption: {item['month_absorption']:+.0f} in {item['month_label']}; Q1 net absorption {item['q1_absorption']:+.0f}.",
        ]
        for line in detail_lines:
            rtf_parts.append(para(f"- {line}", size=22, space_after=55))

    rtf_parts.append(para("Why the Benchmark Needs to Be Corrected", size=26, bold=True, space_after=100))
    benchmark_lines = [
        "RISE Viera and RISE at Nocatee were 2024 deliveries and held their grand openings in January 2025.",
        "RISE at Glen Kernan Park was announced as a phased community with first units delivering in summer 2025 and the balance in summer 2026.",
        "NIC's 4Q25 active-adult occupancy reading was 91.9% across the category, not the 95% stabilized benchmark ownership may be using.",
        "NIC separately notes stabilized active-adult communities open at least two years have averaged about 95% occupancy.",
        "NIC's industry discussion featuring Carlyle and Overture notes higher-end active-adult communities often lease more slowly because prospects may first need to sell a single-family home.",
    ]
    for line in benchmark_lines:
        rtf_parts.append(para(f"- {line}", size=22, space_after=70))

    rtf_parts.append(para("Greystar Counterpoints", size=26, bold=True, space_after=100))
    greystar_lines = [
        "Greystar's official active-adult platform highlights brands such as Overture, Everleigh, and Album, but I did not verify a Greystar-managed active-adult rental community in Jacksonville or Ponte Vedra.",
        "Parasol Melbourne is the only clear Greystar-managed comp I found in the broader Viera/Melbourne scope.",
        "Public listings for Parasol Melbourne currently show availability and concessions, including 14 to 15 available units and up to six weeks free, which suggests a low-90s occupied or committed position rather than a clearly stronger local outcome.",
        "RISE publicly markets a dedicated 55+ active-living platform and already operates communities in the exact Florida geographies under discussion.",
    ]
    for line in greystar_lines:
        rtf_parts.append(para(f"- {line}", size=22, space_after=70))

    rtf_parts.append(para("Recommended Talking Points", size=26, bold=True, space_after=100))
    talking_points = [
        "We should evaluate these assets against active-adult lease-up benchmarks, not fully stabilized two-year-plus communities.",
        f"Viera's current March gap is measurable and manageable: {whole(metrics[0]['units_to_occ_goal'])} occupied units and {whole(metrics[0]['units_to_leased_goal'])} leased units to plan.",
        "Glen Kernan Park is still in phased delivery, so comparing it to fully delivered 55+ communities misstates performance.",
        "Higher-end active-adult product naturally produces slower lead velocity than conventional multifamily because prospects often coordinate a home sale before moving.",
        "Greystar does not appear to have stronger local submarket proof in Jacksonville or Ponte Vedra, and its nearest confirmed comp in Melbourne still appears to be using concessions.",
        "RISE already has in-market 55+ operating experience across these locations, which is more relevant than national brand recognition alone.",
    ]
    for line in talking_points:
        rtf_parts.append(para(f"- {line}", size=22, space_after=70))

    rtf_parts.append(para("Source Links", size=26, bold=True, space_after=100))
    sources = [
        ("https://risere.com/active-adult/", "RISE active-adult platform"),
        ("https://risere.com/portfolio/viera/", "RISE Viera community page"),
        ("https://risere.com/portfolio/rise-nocatee/", "RISE at Nocatee community page"),
        ("https://risere.com/portfolio/rise-glen-kernan-park/", "RISE at Glen Kernan Park community page"),
        ("https://risere.com/grand-openings-for-two-55-communities/", "RISE January 2025 grand opening announcement"),
        ("https://risere.com/groundbreaking-for-rise-at-glen-kernan-park-ushers-in-new-standard-for-active-adult-housing-in-jacksonville/", "Glen Kernan Park phased delivery announcement"),
        ("https://www.nic.org/blog/active-adult/", "NIC 4Q25 active-adult benchmark"),
        ("https://www.nic.org/blog/active-adult-rental-communities-data-dive/", "NIC stabilized active-adult benchmark"),
        ("https://www.nic.org/insider/december-2021/", "NIC discussion of upscale active-adult lease-up pace"),
        ("https://www.greystar.com/product-specialties/active-adult", "Greystar active-adult platform"),
        ("https://www.greystar.com/properties/melbourne-fl/parasol-melbourne", "Greystar Parasol Melbourne"),
        ("https://www.after55.com/fl/melbourne/parasol-melbourne/l77kthf", "After55 listing for Parasol Melbourne"),
        ("https://www.apartments.com/parasol-melbourne-melbourne-fl/677kxht/", "Apartments.com listing for Parasol Melbourne"),
    ]
    for url, label in sources:
        rtf_parts.append(rich_para([rtf_escape("- "), hyperlink(url, label)], size=22, space_after=70))

    rtf_parts.append("}")
    RTF_PATH.write_text("".join(rtf_parts))

    hyperlink_rels: list[tuple[str, str]] = []

    def linked_para(label: str, url: str) -> str:
        rel_id = f"rId{len(hyperlink_rels) + 1}"
        hyperlink_rels.append((rel_id, url))
        link = (
            f'<w:hyperlink r:id="{rel_id}">'
            f'{docx_run(label, size=22, color="0E5EA8", underline=True)}'
            "</w:hyperlink>"
        )
        return docx_para([docx_run("- ", size=22), link], spacing_after=90)

    doc_rows = [["Community", "Units", "Occ %", "Leased %", "Gap to Goal", "March Funnel", "Operating Read"]]
    for item in metrics:
        doc_rows.append(
            [
                item["name"],
                whole(item["units"]),
                f"{pct(item['occupancy_pct'])} ({whole(item['occupied'])})",
                f"{pct(item['leased_pct'])} ({whole(item['leased'])})",
                f"{item['occupancy_gap']:+.1f} occ / {item['leased_gap']:+.1f} leased",
                f"{whole(item['guest_cards'])} GC / {whole(item['tours'])} tours / {whole(item['applications'])} apps / {whole(item['leases_signed'])} leases",
                f"{whole(item['units_to_occ_goal'])} occ and {whole(item['units_to_leased_goal'])} leased units to goal",
            ]
        )

    doc_parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        "<w:body>",
        docx_para([docx_run("Active Adult Performance and Greystar Counter Brief", bold=True, size=32)], center=True, spacing_after=80),
        docx_para([docx_run(f"Prepared March 13, 2026 using the internal snapshot export dated {exported_at}.", size=18)], center=True, spacing_after=160),
        docx_para([docx_run("Bottom line: the updated snapshot shows three communities in different lease-up stages, not a clean case that the current operator lacks 55+ capability. One of the three assets is still in phased delivery, and Greystar does not appear to have meaningful active-adult operating depth in Jacksonville or Ponte Vedra.", size=22)], spacing_after=160),
        docx_para([docx_run("Executive Takeaways", bold=True, size=26)], spacing_after=100),
    ]
    for line in takeaways:
        doc_parts.append(docx_para([docx_run(f"- {line}", size=22)], spacing_after=70))
    doc_parts.append(docx_para([docx_run("Internal Operating Snapshot", bold=True, size=26)], spacing_after=100))
    doc_parts.append(docx_table(doc_rows, table_widths))
    doc_parts.append(docx_para([docx_run("", size=8)], spacing_after=50))
    doc_parts.append(docx_para([docx_run("Property Detail", bold=True, size=26)], spacing_after=100))
    for item in metrics:
        doc_parts.append(docx_para([docx_run(item["name"], bold=True, size=24)], spacing_after=60))
        for line in [
            detail_map[item["name"]],
            f"{item['month_label']} snapshot: {whole(item['occupied'])} occupied ({pct(item['occupancy_pct'])}) and {whole(item['leased'])} leased ({pct(item['leased_pct'])}) on {whole(item['units'])} units.",
            f"Gap to current budget: {whole(item['units_to_occ_goal'])} occupied units and {whole(item['units_to_leased_goal'])} leased units to target.",
            f"Current funnel: {whole(item['guest_cards'])} guest cards, {whole(item['tours'])} tours, {whole(item['applications'])} applications, and {whole(item['leases_signed'])} signed leases.",
            f"Conversion view: lead-to-tour {format_conv(item['lead_to_tour'])}, tour-to-app {format_conv(item['tour_to_app'])}, lead-to-lease {format_conv(item['lead_to_lease'])}, app-to-lease {format_conv(item['app_to_lease'])}.",
            f"Absorption: {item['month_absorption']:+.0f} in {item['month_label']}; Q1 net absorption {item['q1_absorption']:+.0f}.",
        ]:
            doc_parts.append(docx_para([docx_run(f"- {line}", size=22)], spacing_after=55))
    doc_parts.append(docx_para([docx_run("Why the Benchmark Needs to Be Corrected", bold=True, size=26)], spacing_after=100))
    for line in benchmark_lines:
        doc_parts.append(docx_para([docx_run(f"- {line}", size=22)], spacing_after=70))
    doc_parts.append(docx_para([docx_run("Greystar Counterpoints", bold=True, size=26)], spacing_after=100))
    for line in greystar_lines:
        doc_parts.append(docx_para([docx_run(f"- {line}", size=22)], spacing_after=70))
    doc_parts.append(docx_para([docx_run("Recommended Talking Points", bold=True, size=26)], spacing_after=100))
    for line in talking_points:
        doc_parts.append(docx_para([docx_run(f"- {line}", size=22)], spacing_after=70))
    doc_parts.append(docx_para([docx_run("Source Links", bold=True, size=26)], spacing_after=100))
    for url, label in sources:
        doc_parts.append(linked_para(label, url))
    doc_parts.append('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>')
    doc_parts.append("</w:body></w:document>")
    build_docx("".join(doc_parts), hyperlink_rels)


if __name__ == "__main__":
    main()
