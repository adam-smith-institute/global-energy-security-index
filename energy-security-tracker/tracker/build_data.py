"""Build the tracker's data files from the final index CSV.

    python3 build_data.py            # data.js only (world.js is kept)
    python3 build_data.py --world    # also re-download the map geometry

Reads  ../Energy_Security_Index_2026_final.csv   (written by build_index.py)
Writes data.js   window.ESI = {edition, dataYear, countries: [...]}
       world.js  window.WORLD = <world-atlas 50m TopoJSON, Antarctica removed>

Re-run whenever the CSV changes, then republish the tracker and run build_standalone.py.
"""
import csv, json, os, sys, urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
CSV = os.path.join(HERE, "..", "Energy_Security_Index_2026_final.csv")
WORLD_URL = "https://cdn.jsdelivr.net/npm/world-atlas@2/countries-50m.json"

countries = []
with open(CSV, encoding="utf-8") as f:
    for r in csv.DictReader(f):
        n = lambda k, d=4: round(float(r[k]), d) if r[k] != "" else None
        countries.append({
            "iso3": r["iso3"],
            "id": r["iso_numeric"],
            "name": r["country"],
            "rank": int(r["rank"]),
            "overall": n("overall"),
            "s1": n("s1"),
            "s2": n("s2"),
            "s3": n("s3"),
            "s4": n("s4"),
            "imp": n("net_import_dependence", 6),
            "hhi": n("partner_hhi", 6),
            "wgiDom": n("domestic_wgi"),
            "wgiSup": n("supplier_wgi"),
            "covPartner": n("partner_coverage_pct"),
            "covWgi": n("supplier_wgi_coverage_pct"),
            "suppliers": int(float(r["supplier_count"])) if r["supplier_count"] else None,
            "gdppc": n("gdp_per_capita_2015usd", 0),
        })

with open(os.path.join(HERE, "data.js"), "w", encoding="utf-8") as f:
    f.write("window.ESI=")
    json.dump({"edition": 2026, "dataYear": 2024, "countries": countries}, f, separators=(",", ":"), ensure_ascii=False)
    f.write(";\n")

world_js = os.path.join(HERE, "world.js")
if "--world" in sys.argv or not os.path.exists(world_js):
    with urllib.request.urlopen(WORLD_URL) as r:
        world = json.load(r)
    geoms = world["objects"]["countries"]["geometries"]
    world["objects"]["countries"]["geometries"] = [g for g in geoms if g.get("id") != "010"]
    with open(world_js, "w") as f:
        f.write("window.WORLD=")
        json.dump(world, f, separators=(",", ":"))
        f.write(";\n")

print(f"{len(countries)} countries -> data.js")
