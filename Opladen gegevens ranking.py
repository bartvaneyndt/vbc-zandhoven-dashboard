import json
import urllib.request
import xlrd
import datetime

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

# Reeksen/categorieën die je bewust wilt uitsluiten
EXCLUDE_CATEGORIES = [
    "Nationale 3 Dames C (reserven)"
]

for sheet_name in wb.sheet_names():
    sheet = wb.sheet_by_name(sheet_name)
    current_cat = sheet_name
    teams_in_table = []
    
    for r in range(sheet.nrows):
        row_vals = [sheet.cell_value(r, c) for c in range(sheet.ncols)]
        val0 = str(row_vals[0]).strip()

        # Categorie of reeks herkennen
        if val0 and not val0.startswith("0") and not val0.startswith("1") and val0 not in ["Reserve", "Ploeg", ""]:
            if teams_in_table:
                if current_cat.strip() not in EXCLUDE_CATEGORIES:
                    all_rankings[current_cat] = teams_in_table
                teams_in_table = []
            current_cat = val0

        # Controleer of dit een rij is met een positie en teamnaam
        if val0 and (val0.endswith(".") or val0[:-1].isdigit()):
            team_name = str(row_vals[1]).strip() if len(row_vals) > 1 else ""
            if team_name and team_name != "Ploeg":
                teams_in_table.append({
                    "pos": val0,
                    "team": team_name,
                    "ptn": row_vals[2] if len(row_vals) > 2 else "",
                    "wed": row_vals[3] if len(row_vals) > 3 else "",
                    "is_zandhoven": "Zandhoven" in team_name
                })

    if teams_in_table and current_cat.strip() not in EXCLUDE_CATEGORIES:
        all_rankings[current_cat] = teams_in_table

# -------------------------------------------------------------
# 3. Herorganiseren van de Dames-reeksen
# -------------------------------------------------------------
preferred_order = [
    "Nationale 1 Dames",
    "Nationale 3 Dames C",
    "Dames promo 2 A",
    "Dames promo 2 B",
    "Dames promo 4 (1e gewest Herentals)",
    "Dames promo 5 (2e gewest Herentals)"
]

ordered_rankings = {}

for cat in preferred_order:
    for key in list(all_rankings.keys()):
        if key.strip() == cat.strip():
            ordered_rankings[key] = all_rankings[key]
            break

for key, value in all_rankings.items():
    if key not in ordered_rankings:
        ordered_rankings[key] = value

# -------------------------------------------------------------
# 4. Automatisch genereren van opgesplitste HTML-slides (3 kolommen, 1 rij)
# -------------------------------------------------------------
categories = list(ordered_rankings.items())
chunk_size = 3  # Maximaal 3 tabellen per pagina (1 rij van 3 kolommen)
chunks = [categories[i:i + chunk_size] for i in range(0, len(categories), chunk_size)]

base_css = """
<style>
    body {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        background-color: #121212;
        color: #ffffff;
        margin: 0;
        padding: 20px;
        width: 1920px;
        box-sizing: border-box;
    }
    h1 {
        text-align: center;
        color: #ff0000;
        margin-bottom: 5px;
    }
    .subtitle {
        text-align: center;
        color: #aaaaaa;
        margin-bottom: 20px;
        font-size: 1.1rem;
    }
    .container {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 20px;
        max-width: 1850px;
        margin: 0 auto;
    }
    .category-card {
        background-color: #1e1e1e;
        border-radius: 10px;
        padding: 15px;
        box-shadow: 0 4px 10px rgba(0,0,0,0.5);
    }
    h2 {
        font-size: 1.1rem;
        color: #ffaa00;
        margin-top: 0;
        border-bottom: 2px solid #333;
        padding-bottom: 6px;
    }
    table {
        width: 100%;
        border-collapse: collapse;
        margin-top: 8px;
        font-size: 0.9rem;
    }
    th, td {
        padding: 6px 8px;
        text-align: left;
        border-bottom: 1px solid #2c2c2c;
    }
    th {
        background-color: #2a2a2a;
        color: #ccc;
    }
    .zandhoven-row {
        background-color: rgba(255, 102, 0, 0.25);
        font-weight: bold;
        color: #ff9944;
    }
    .center {
        text-align: center;
    }
    .footer-info {
        text-align: center;
        color: #666;
        margin-top: 15px;
        font-size: 0.85rem;
    }
</style>
"""

timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
total_chunks = len(chunks)

for index, chunk in enumerate(chunks):
    file_name = "index.html" if index == 0 else f"slide{index + 1}.html"
    slide_num = index + 1
    
    html_lines = [
        "<!DOCTYPE html>",
        '<html lang="nl">',
        "<head>",
        '    <meta charset="UTF-8">',
        '    <meta name="viewport" content="width=1920, initial-scale=1.0">',
        f"    <title>VBC Zandhoven - Live Standen (Deel {slide_num})</title>",
        base_css,
        "</head>",
        "<body>",
        "",
        "    <h1>VBC Zandhoven</h1>",
        f'    <div class="subtitle">Live Ranglijsten & Competitiestanden — Slide {slide_num} van {total_chunks}</div>',
        "",
        '    <div class="container">'
    ]

    for category, teams in chunk:
        html_lines.append('        <div class="category-card">')
        html_lines.append(f'            <h2>{category}</h2>')
        html_lines.append('            <table>')
        html_lines.append('                <thead>')
        html_lines.append('                    <tr>')
        html_lines.append('                        <th style="width: 35px;">P</th>')
        html_lines.append('                        <th>Ploeg</th>')
        html_lines.append('                        <th class="center" style="width: 45px;">Ptn</th>')
        html_lines.append('                        <th class="center" style="width: 45px;">Wed</th>')
        html_lines.append('                    </tr>')
        html_lines.append('                </thead>')
        html_lines.append('                <tbody>')

        for team in teams:
            row_class = ' class="zandhoven-row"' if team['is_zandhoven'] else ''
            html_lines.append(f'                    <tr{row_class}>')
            html_lines.append(f'                        <td>{team["pos"]}</td>')
            html_lines.append(f'                        <td>{team["team"]}</td>')
            html_lines.append(f'                        <td class="center">{team["ptn"]}</td>')
            html_lines.append(f'                        <td class="center">{team["wed"]}</td>')
            html_lines.append('                    </tr>')

        html_lines.append('                </tbody>')
        html_lines.append('            </table>')
        html_lines.append('        </div>')

    html_lines.append('    </div>')
    html_lines.append(f'    <div class="footer-info">Laatste update: {timestamp}</div>')
    html_lines.append("</body>")
    html_lines.append("</html>")

    with open(file_name, "w", encoding="utf-8") as f:
        f.write("\n".join(html_lines))

print(f"Klaar! {total_chunks} HTML-slides succesvol gegenereerd.")
