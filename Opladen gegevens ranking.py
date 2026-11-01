import json
import urllib.request
import xlrd

# De directe URL van Volleyscores die je hebt opgegeven
EXCEL_URL = "https://www.volleyscores.be/index.php?v=2&isActiveSeason=1&t=Rankings%20AH-1717%20VBC%20Zandhoven&a=re&se=13&pi=&si=&ti=&ci=11082&mm=&ssi=&st=&w=%&f=&lng=nl"
LOCAL_FILE = "volleyscores_rankings.xls"

print("1. Bezig met downloaden van het Excel-bestand via Volleyscores...")
try:
  # We voegen een User-Agent toe zodat de website denkt dat het een webbrowser is
  headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
  req = urllib.request.Request(EXCEL_URL, headers=headers)

  with urllib.request.urlopen(req) as response, open(
      LOCAL_FILE, "wb"
  ) as out_file:
    out_file.write(response.read())
  print("Download succesvol!")
except Exception as e:
  print(f"Fout bij het downloaden: {e}")
  exit(1)

print("2. Bezig met inlezen en filteren van VBC Zandhoven teams...")
wb = xlrd.open_workbook(LOCAL_FILE, ignore_workbook_corruption=True)
zandhoven_teams = []

for sheet_name in wb.sheet_names():
  sheet = wb.sheet_by_name(sheet_name)
  current_cat = sheet_name
  for r in range(sheet.nrows):
    row_vals = [sheet.cell_value(r, c) for c in range(sheet.ncols)]
    val0 = str(row_vals[0]).strip()

    # Categorie of reeks herkennen
    if (
        val0
        and not val0.startswith("0")
        and not val0.startswith("1")
        and val0 not in ["Reserve", "Ploeg", ""]
    ):
      current_cat = val0

    # Zoek of 'Zandhoven' in de rij voorkomt
    for val in row_vals:
      if "Zandhoven" in str(val):
        zandhoven_teams.append({
            "category": current_cat,
            "pos": str(row_vals[0]).strip(),
            "team": str(val).strip(),
            "ptn": row_vals[2] if len(row_vals) > 2 else "",
            "wed": row_vals[3] if len(row_vals) > 3 else "",
        })

# 3. Sla de resultaten op in een JSON-bestand
output_json = "teams_data.json"
with open(output_json, "w", encoding="utf-8") as f:
  json.dump(zandhoven_teams, f, ensure_ascii=False, indent=4)

print(
    f"Klaar! {len(zandhoven_teams)} teams gevonden en opgeslagen in"
    f" '{output_json}'."
)