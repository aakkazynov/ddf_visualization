"""Build data.json and data.js from the files in research/.

Inputs:
  research/ddf_pooled_programs_funders.json       13 pooled programs and 80 funders (from the DDF Funders Network page)
  research/ddf_international_cofunded_programs.xlsx  co-funded programs that sent money outside the US

Outputs (next to this script):
  data.json  for pages served over http (fetch("data.json"))
  data.js    sets window.DDF_DATA, so a page opened by double-click can use <script src="data.js"></script>

Run:  python build_data.py
"""
import json
import re
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
RESEARCH = ROOT / "research"
NETWORK_JSON = RESEARCH / "ddf_pooled_programs_funders.json"
INTL_XLSX = RESEARCH / "ddf_international_cofunded_programs.xlsx"
HEADER_ROW = 4  # every tab in the workbook has a title block above the header row

# The 13 network programs all fund work in the United States (confirmed by the project owner).
NETWORK_COUNTRIES = ["USA"]
US = {"iso3": "USA", "name": "United States", "region": "North America"}

# Standard ISO 3166-1 numeric codes; the page's world map (world-atlas TopoJSON) keys countries by these.
ISO_NUMERIC = {
    "USA": "840", "ETH": "231", "MOZ": "508", "GHA": "288", "UGA": "800", "CHN": "156", "FRA": "250",
    "ARG": "032", "BRA": "076", "CHL": "152", "COL": "170", "CUB": "192", "GTM": "320", "HTI": "332",
    "HND": "340", "MEX": "484", "URY": "858", "DZA": "012", "TZA": "834", "EGY": "818", "IDN": "360",
    "PAK": "586", "UKR": "804",
}


def read_tab(wb, name, columns, id_prefix):
    """Read the rows of one tab, keeping only `columns` ({header: output key}).

    The header is matched from the left, so the country summary beside the
    Countries table is ignored. Rows whose first cell doesn't start with
    `id_prefix` (blank rows, the Total row, notes) are skipped.
    """
    ws = wb[name]
    headers = {}
    for cell in ws[HEADER_ROW]:
        if cell.value is None:
            break
        headers.setdefault(cell.value, cell.column - 1)
    missing = [h for h in columns if h not in headers]
    if missing:
        raise SystemExit(f"{INTL_XLSX.name} / {name}: missing column(s) {missing}")

    rows = []
    for row in ws.iter_rows(min_row=HEADER_ROW + 1, values_only=True):
        if not (isinstance(row[0], str) and row[0].startswith(id_prefix)):
            continue
        rows.append({key: row[headers[h]] for h, key in columns.items()})
    return rows


def split_list(text, sep):
    return [part.strip() for part in (text or "").split(sep) if part.strip()]


def parse_amount(text):
    """Headline amount in $ millions from text like "~$11.6M (...)" or "$166,683 (...)".

    Returns (value, approximate); (None, False) when the text doesn't start with a dollar figure,
    e.g. "Not disclosed". Only the leading figure is used, so nothing is added up or estimated here.
    """
    m = re.match(r"\s*(~)?\$([\d,]+(?:\.\d+)?)\s*([MB])?", text or "")
    if not m:
        return None, False
    value = float(m.group(2).replace(",", ""))
    unit = m.group(3)
    millions = value * 1000 if unit == "B" else value if unit == "M" else value / 1e6
    return round(millions, 3), bool(m.group(1))


def start_year(text):
    m = re.search(r"(19|20)\d{2}", text or "")
    if not m:
        raise SystemExit(f"No year found in {text!r}")
    return int(m.group(0))


def build_network():
    src = json.loads(NETWORK_JSON.read_text(encoding="utf-8"))
    programs = [dict(p, countries=list(NETWORK_COUNTRIES)) for p in src["programs"]]
    return {
        "source": src["source"],
        "ddf": src["ddf"],
        "programs": programs,
        "funders": src["funders"],
    }


def build_international():
    wb = load_workbook(INTL_XLSX, read_only=True)

    programs = read_tab(wb, "Programs", {
        "Program ID": "id",
        "Program": "name",
        "DDF Program Area": "ddf_program",
        "Years": "years",
        "Administering / Implementing Org(s)": "administrator",
        "Total Pool": "total_pool",
        "DDF Contribution": "ddf_contribution",
        "Meets Outside-US Test": "meets_outside_us_test",
        "Confidence": "confidence",
        "Description": "description",
        "Caveats": "caveats",
        "Source URL(s)": "sources",
        "Also Funds US Work": "funds_us_work",
        "US Work Evidence": "us_work_evidence",
    }, "P")
    funders = read_tab(wb, "Funders", {
        "Funder ID": "id",
        "Funder (short name)": "name",
        "Full Official Name": "full_name",
        "Type": "type",
        "HQ Country": "hq_country",
        "Website": "website",
    }, "F")
    links = read_tab(wb, "Program-Funder Links", {
        "Program ID": "program_id",
        "Funder ID": "funder_id",
        "Role": "role",
        "Contribution (if known)": "contribution",
        "Confidence": "confidence",
    }, "P")
    countries = read_tab(wb, "Countries", {
        "Program ID": "program_id",
        "Country": "country",
        "ISO3": "iso3",
        "Region": "region",
        "Detail": "detail",
    }, "P")

    # The workbook's count columns are Excel formulas with no saved values,
    # so the counts are worked out here from the link and country rows.
    program_ids = {p["id"] for p in programs}
    funder_ids = {f["id"] for f in funders}
    for row in links:
        if row["program_id"] not in program_ids or row["funder_id"] not in funder_ids:
            raise SystemExit(f"Link points to an unknown program or funder: {row}")
    for row in countries:
        if row["program_id"] not in program_ids:
            raise SystemExit(f"Country row points to an unknown program: {row}")

    for p in programs:
        if p["funds_us_work"] not in ("Yes", "No"):
            raise SystemExit(f"{p['id']}: 'Also Funds US Work' must be Yes or No, got {p['funds_us_work']!r}")
        p["funds_us_work"] = p["funds_us_work"] == "Yes"
        p["start_year"] = start_year(p["years"])
        p["amount_usd_millions"], p["amount_is_approximate"] = parse_amount(p["total_pool"])
        p["sources"] = split_list(p["sources"], ";")
        p["funder_ids"] = [l["funder_id"] for l in links if l["program_id"] == p["id"]]
        p["countries"] = [c["iso3"] for c in countries if c["program_id"] == p["id"]]
    for f in funders:
        f["program_ids"] = [l["program_id"] for l in links if l["funder_id"] == f["id"]]
    for row in countries:
        row["detail"] = row["detail"] or None

    return {"programs": programs, "funders": funders, "links": links, "countries": countries}


def build_country_index(intl):
    """One entry per country that any program funds, with region and map code."""
    seen = {US["iso3"]: dict(US)}
    for row in intl["countries"]:
        seen.setdefault(row["iso3"], {"iso3": row["iso3"], "name": row["country"], "region": row["region"]})
    missing = sorted(k for k in seen if k not in ISO_NUMERIC)
    if missing:
        raise SystemExit(f"Add ISO numeric codes for: {missing}")
    for c in seen.values():
        c["iso_numeric"] = ISO_NUMERIC[c["iso3"]]
    return sorted(seen.values(), key=lambda c: (c["region"], c["name"]))


def main():
    intl = build_international()
    data = {
        "generated": date.today().isoformat(),
        "inputs": [str(p.relative_to(ROOT)).replace("\\", "/") for p in (NETWORK_JSON, INTL_XLSX)],
        "network": build_network(),
        "international": intl,
        "countries": build_country_index(intl),
    }

    (ROOT / "data.json").write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    # ASCII-only so accented names survive even if the page doesn't declare UTF-8.
    js = "// Generated by build_data.py. Do not edit by hand.\nwindow.DDF_DATA = " + json.dumps(data, indent=2) + ";\n"
    (ROOT / "data.js").write_text(js, encoding="utf-8")

    net, intl = data["network"], data["international"]
    print(f"network: {len(net['programs'])} programs, {len(net['funders'])} funders + DDF")
    print(f"international: {len(intl['programs'])} programs, {len(intl['funders'])} funders (incl. DDF), "
          f"{len(intl['links'])} links, {len({c['iso3'] for c in intl['countries']})} countries")
    print("wrote data.json and data.js")


if __name__ == "__main__":
    main()
