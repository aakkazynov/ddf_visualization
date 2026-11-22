"""Build data.json, data.js and DDF_Pooled_Fund_Programs.xlsx from research/*.json.

Usage:  python build_data.py
Inputs: research/programs.json  {"programs": [...], "excluded": [...], "leads": [...]}
        research/funders.json   {"funders": [...]}
"""
import json
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

HERE = Path(__file__).parent
RES = HERE / "research"
UPDATED = "November 2026"
DDF = "Doris Duke Foundation"

# Program-area palette (validated for color-vision deficiency on the dark background)
AREAS = [
    {"name": "Arts", "color": "#E53935"},                 # red
    {"name": "Medical Research", "color": "#3F51B5"},     # blue
    {"name": "Environment", "color": "#00897B"},          # green
    {"name": "Child Well-being", "color": "#FFA726"},     # orange
    {"name": "Building Bridges", "color": "#BA68C8"},     # purple
    {"name": "Cross-program", "color": "#4DD0E1"},        # cyan
]

REGIONS = {
    "North America": ["United States", "Canada", "Mexico"],
    "Latin America & Caribbean": ["Brazil", "Colombia", "Peru", "Chile", "Argentina", "Guatemala", "Haiti", "Honduras",
                                  "Ecuador", "Bolivia", "Puerto Rico", "Dominican Republic", "Jamaica", "Nicaragua", "El Salvador"],
    "Europe": ["United Kingdom", "France", "Germany", "Netherlands", "Switzerland", "Italy", "Spain", "Belgium", "Ireland",
               "Sweden", "Norway", "Denmark", "Poland", "Ukraine", "Greece", "Portugal", "Bosnia and Herzegovina"],
    "Africa": ["Ghana", "Mozambique", "Rwanda", "Tanzania", "Zambia", "Kenya", "Uganda", "Nigeria", "South Africa", "Ethiopia",
               "Malawi", "Senegal", "Mali", "Burkina Faso", "Democratic Republic of the Congo", "Zimbabwe", "Botswana",
               "Cameroon", "Côte d'Ivoire", "Sierra Leone", "Liberia", "Madagascar", "Niger", "Benin", "Togo", "Sudan",
               "South Sudan", "Somalia", "Egypt", "Morocco", "Tunisia", "Namibia", "Lesotho", "Eswatini", "Guinea", "Chad"],
    "Middle East": ["Lebanon", "Jordan", "Israel", "Palestine", "Turkey", "Iraq", "Syria", "Saudi Arabia",
                    "United Arab Emirates", "Qatar", "Iran", "Yemen", "Oman", "Kuwait"],
    "Asia": ["India", "China", "Indonesia", "Pakistan", "Bangladesh", "Philippines", "Vietnam", "Thailand", "Malaysia",
             "Nepal", "Japan", "South Korea", "Afghanistan", "Myanmar", "Cambodia", "Sri Lanka", "Uzbekistan", "Kazakhstan",
             "Kyrgyzstan", "Tajikistan", "Singapore", "Laos", "Mongolia"],
    "Oceania": ["Australia", "New Zealand", "Papua New Guinea", "Fiji"],
}
REGION_OF = {c: r for r, cs in REGIONS.items() for c in cs}


ALIASES = {"Andrew W. Mellon Foundation": "Mellon Foundation",
           "Doris Duke Charitable Foundation": DDF}


def slug(s):
    return "".join(c if c.isalnum() else "-" for c in s.lower()).strip("-").replace("--", "-")


def load():
    """The 13 programs / 80 funders of the reference visualization (research/original.json).
    None of these programs funds work outside the US. The earlier research files
    (programs.json, international.json, funders.json) are kept for reference only."""
    o = json.loads((RES / "original.json").read_text(encoding="utf-8"))
    full = lambda f: DDF if f == "DDF" else f
    programs = [{
        "id": slug(p["n"]), "name": p["n"], "short_name": p["n"], "type": None,
        "area": p["pr"], "year": p["y"], "year_end": None,
        "amount_usd": p["amt"] * 1e6 if p["amt"] is not None else None,
        "amount_note": p["basis"], "funders": [full(f) for f in p["f"]], "manager": None,
        "scope": "US", "countries": ["United States"], "description": "",
        "sources": [p["src"]] if p["src"] else [], "confidence": None,
    } for p in o["P"]]
    names = list(dict.fromkeys(f for p in o["P"] for f in p["f"]))
    funders = [{
        "name": full(n), "short": "Doris Duke" if n == "DDF" else n,
        "endowment_usd": o["E"][n][0] * 1e9 if n in o["E"] else None,
        "endowment_note": o["E"][n][1] if n in o["E"] else "",
        "website": o["WEB"].get(n), "website_checked": n in o["WEB"],
    } for n in names]
    return {"programs": programs, "excluded": [], "leads": [],
            "caveat": "Program list, funders, endowments and websites follow the reference DDF Funders Network (ddf-constellation.vercel.app)."}, funders


def load_research():
    progs = json.loads((RES / "programs.json").read_text(encoding="utf-8"))
    intl = RES / "international.json"
    if intl.exists():
        by_id = {p["id"]: i for i, p in enumerate(progs["programs"])}
        for p in json.loads(intl.read_text(encoding="utf-8"))["programs"]:
            if p["id"] in by_id:
                progs["programs"][by_id[p["id"]]] = p      # updated record wins
            else:
                progs["programs"].append(p)
    for p in progs["programs"]:
        p["funders"] = list(dict.fromkeys(ALIASES.get(f, f) for f in p["funders"]))
    funders = json.loads((RES / "funders.json").read_text(encoding="utf-8"))["funders"]
    return progs, funders


def main():
    pj, funders = load()
    programs = pj["programs"]
    area_names = {a["name"] for a in AREAS}
    for p in programs:
        if p["area"] not in area_names:
            raise SystemExit(f"Unknown area {p['area']!r} in {p['id']}")
        p.setdefault("countries", [])
        if not p["countries"]:
            p["countries"] = ["Not disclosed"]
    fnames = {f["name"] for f in funders}
    missing = sorted({n for p in programs for n in p["funders"]} - fnames)
    if missing:
        print("WARNING funders without a funders.json row:", missing)
    used = {n for p in programs for n in p["funders"]}
    funders = [f for f in funders if f["name"] in used]

    countries = sorted({c for p in programs for c in p["countries"]})
    unknown = [c for c in countries if c not in REGION_OF and c != "Not disclosed"]
    if unknown:
        print("WARNING countries without a region:", unknown)
    country_rows = [{"name": c, "region": REGION_OF.get(c, "Other")} for c in countries]

    data = {
        "meta": {"updated": UPDATED, "caveat": pj.get("caveat", "")},
        "areas": AREAS,
        "programs": programs,
        "funders": funders,
        "countries": country_rows,
    }
    (HERE / "data.json").write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    (HERE / "data.js").write_text("window.DDF_DATA = " + json.dumps(data, ensure_ascii=False) + ";\n", encoding="utf-8")

    # ---- workbook ----
    wb = Workbook()
    hdr = Font(bold=True, color="FFFFFF")
    fill = PatternFill("solid", fgColor="22262C")

    def sheet(title, cols, rows, first=False):
        ws = wb.active if first else wb.create_sheet()
        ws.title = title
        ws.append(cols)
        for c in ws[1]:
            c.font, c.fill = hdr, fill
        for r in rows:
            ws.append(r)
        for i, col in enumerate(cols, 1):
            width = max(len(str(col)), *(len(str(r[i - 1])) for r in rows)) if rows else len(col)
            ws.column_dimensions[ws.cell(1, i).column_letter].width = min(60, max(10, width + 2))
        ws.freeze_panes = "A2"
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.alignment = Alignment(vertical="top", wrap_text=False)

    for area in [a["name"] for a in AREAS]:
        pass
    sheet("Programs",
          ["Program", "Short name", "DDF program area", "Type", "Year", "Year end", "Amount (USD)", "Amount meaning",
           "Scope", "Countries", "# funders", "Funders", "Manager", "Description", "Confidence", "Sources"],
          [[p["name"], p.get("short_name", ""), p["area"], p["type"], p.get("year"), p.get("year_end"), p.get("amount_usd"),
            p.get("amount_note", ""), p["scope"], "; ".join(p["countries"]), len(p["funders"]), "; ".join(p["funders"]),
            p.get("manager") or "", p.get("description", ""), p.get("confidence", ""), " ".join(p.get("sources", []))]
           for p in sorted(programs, key=lambda p: (p["area"], p.get("year") or 0))], first=True)
    sheet("Funders",
          ["Funder", "Type", "HQ country", "Endowment / assets (USD)", "Endowment note", "Endowment source", "Website",
           "# programs with DDF", "Programs"],
          [[f["name"], f.get("type", ""), f.get("hq_country", ""), f.get("endowment_usd"), f.get("endowment_note", ""),
            f.get("endowment_source", ""), f.get("website", ""),
            sum(f["name"] in p["funders"] for p in programs),
            "; ".join(p["name"] for p in programs if f["name"] in p["funders"])] for f in sorted(funders, key=lambda f: f["name"])])
    sheet("Program-Funder links", ["Program", "DDF program area", "Funder"],
          [[p["name"], p["area"], n] for p in programs for n in p["funders"]])
    sheet("Countries", ["Program", "Scope", "Country", "Region"],
          [[p["name"], p["scope"], c, REGION_OF.get(c, "Other")] for p in programs for c in p["countries"]])
    sheet("Excluded items", ["Item", "Reason"], [[e.get("name", ""), e.get("reason", "")] for e in pj.get("excluded", [])])
    sheet("Leads", ["Lead", "Note"], [[l.get("lead", ""), l.get("note", "")] for l in pj.get("leads", [])])
    wb.save(HERE / "DDF_Pooled_Fund_Programs.xlsx")
    print(f"OK: {len(programs)} programs, {len(funders)} funders, {len(countries)} countries")


if __name__ == "__main__":
    main()
