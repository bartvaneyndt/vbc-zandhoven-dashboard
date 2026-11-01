import json
import urllib.request
import xlrd

# De directe URL van Volleyscores
EXCEL_URL = "https://www.volleyscores.be/index.php?v=2&isActiveSeason=1&t=Rankings%20AH-1717%20VBC%20Zandhoven&a=re&se=13&pi=&si=&ti=&ci=11082&mm=&ssi=&st=&w=%&f=&lng=nl"
LOCAL_FILE = "volleyscores_rankings.xls"

print("1. Bezig met downloaden van het Excel-bestand via Volleyscores...")
try:
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    req = urllib.request.Request(EXCEL_URL, headers=headers)
    with urllib.request.urlopen(req) as response, open(LOCAL_FILE, "wb") as out_file:
        out_file.write(response.read())
    print("Download succesvol!")
except Exception as e:
    print(f"Fout bij het downloaden: {e}")
    exit(1)

print("2. Bezig met inlezen van alle volledige ranglijsten...")
wb = xlrd.open_workbook(LOCAL_FILE, ignore_workbook_corruption=True)
all_rankings = {}

for sheet_name in wb.sheet_names():
    sheet = wb.sheet_by_name(sheet_name)
    current_cat = sheet_name
    teams_in_table = []
    
    for r in range(sheet.nrows):
        row_vals = [sheet.cell_value(r, c) for c in range(sheet.ncols)]
        val0 = str(row_vals[0]).strip()

        # Categorie of reeks herkennen
        if val0 and not val0.startswith("0") and not val0.startswith("1") and val0 not in ["Reserve", "Ploeg", ""]:
            # Als we al teams hadden verzameld voor een vorige categorie, slaan we die eerst op
            if teams_in_table:
                all_rankings[current_cat] = teams_in_table
                teams_in_table = []
            current_cat = val0

        # Controleer of dit een rij is met een positie en teamnaam (meestal begint positie met cijfer of bevat cijfer)
        if val0 and (val0.endswith(".") or val0[:-1].isdigit()):
            team_name = str(row_vals[1]).strip() if len(row_vals) > 1 else ""
            if team_name and team_name != "Ploeg":
                teams_in_table.append({
                    "pos": val0,
                    "team": team_name,
                    "ptn": row_vals[2] if len(row_vals) > 2 else "",
                    "wed": row_vals[3] if len(row_vals) > 3 else "",
                    "is_zandhoven": "Zandhoven" in team_name # Handig om jouw teams straks op het scherm te accentueren!
                })

    # Sla de laatste categorie van de sheet ook op
    if teams_in_table:
        all_rankings[current_cat] = teams_in_table

# 3. Sla de volledige resultaten op in teams_data.json
output_json = "teams_data.json"
with open(output_json, "w", encoding="utf-8") as f:
    json.dump(all_rankings, f, ensure_ascii=False, indent=4)

print(f"Klaar! Volledige tabellen opgeslagen in '{output_json}'.")
