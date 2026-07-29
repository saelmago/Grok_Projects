from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# Enlaces preferentemente del sitio oficial de la marca (ficha o busqueda de catalogo)
INVENTARIO = [
    ("Milwaukee", "Herramienta", "Fontaneria", "M12 Stick Transfer Pump", "2579-20", "Comprado",
     "https://www.milwaukeetool.com/Products/Power-Tools/Plumbing/Pumps/M12-FUEL-Stick-Transfer-Pump/2579-20"),
    ("Milwaukee", "Herramienta", "Fontaneria", "Llave Faucet Swap-Out", "48-22-7100", "Ya tengo",
     "https://www.milwaukeetool.com/Products/Hand-Tools/Plumbing/Wrenches/Faucet-and-Sink-Installer/48-22-7100"),
    ("Knipex", "Herramienta", "Fontaneria", 'Cobra Water Pump Pliers 10"', "87 02 250", "Ya tengo",
     "https://www.knipex.com/products/water-pump-pliers-and-pipe-wrenches/water-pump-pliers/knipex-cobra-water-pump-pliers/8702250"),
    ("Knipex", "Herramienta", "Fontaneria", 'TwinGrip Pliers 8"', "83 01 200", "Ya tengo",
     "https://www.knipex.com/products/pliers/front-and-side-gripping-pliers/knipex-twingrip-front-and-side-gripping-pliers/8201200"),
    ("Knipex", "Herramienta", "Fontaneria", 'Pliers Wrench 10"', "86 01 250", "Comprado",
     "https://www.knipex.com/products/pliers-wrenches-and-pipe-wrenches/pliers-wrenches/pliers-wrench-pliers-and-a-wrench-in-a-single-tool/8601250"),
    ("Wera", "Herramienta", "Fontaneria", "Joker Self-Setting Wrench M/L/XL", "6004 serie", "Ya tengo",
     "https://www.wera.de/en/tools/tool-types/wrench/joker-adjustable-open-end-wrenches"),
    ("Klein", "Herramienta", "Electricidad", "Non-Contact Voltage Tester", "NCVT-39", "Ya tengo",
     "https://www.kleintools.com/catalog/search?search=NCVT-3"),
    ("Klein", "Herramienta", "Electricidad", "GFCI Receptacle Tester", "RT250", "Ya tengo",
     "https://www.kleintools.com/catalog/electrical-testers/gfci-receptacle-tester-lcd-display"),
    ("Klein", "Herramienta", "Electricidad", "Circuit Breaker Finder", "ET310", "Ya tengo",
     "https://www.kleintools.com/catalog/circuit-breaker-finders/digital-circuit-breaker-finder-integrated-gfci-outlet-tester"),
    ("Klein", "Herramienta", "Electricidad", "Tone and Probe Kit", "VDV500-705P", "Ya tengo",
     "https://www.kleintools.com/catalog/search?search=VDV500-705"),
    ("Klein", "Herramienta", "Electricidad", "6-in-1 Insulated Screwdriver", "32306INS", "Ya tengo",
     "https://www.kleintools.com/catalog/search?search=32306INS"),
    ("Klein", "Herramienta", "Electricidad", '9" Insulated Side-Cutting Pliers', "2139NERINS", "Ya tengo",
     "https://www.kleintools.com/catalog/search?search=2139NE"),
    ("Klein", "Herramienta", "Electricidad", '8" Insulated Diagonal Cutting Pliers', "2288RINS", "Ya tengo",
     "https://www.kleintools.com/catalog/search?search=2288RINS"),
    ("Klein", "Herramienta", "Electricidad", '8" Insulated Long Nose Pliers', "2038RINS", "Ya tengo",
     "https://www.kleintools.com/catalog/search?search=2038RINS"),
    ("Klein", "Herramienta", "Electricidad", "Insulated Wire Stripper", "11055RINS", "Ya tengo",
     "https://www.kleintools.com/catalog/search?search=11055"),
    ("Klein", "Herramienta", "Electricidad", "8-in-1 Insulated Screwdriver Set", "32288", "Ya tengo",
     "https://www.kleintools.com/catalog/search?search=32288"),
    ("Klein", "Herramienta", "Electricidad", "Insulated Crimping and Cutting Tool", "1005RINS", "Comprado",
     "https://www.kleintools.com/catalog/search?search=1005"),
    ("Klein", "Herramienta", "Electricidad", "Adjustable Length Screwdriver", "32751", "Comprado",
     "https://www.kleintools.com/catalog/search?search=32751"),
    ("Klein", "Accesorio", "Electricidad", "Grab-And-Go Impact Socket Set (metrico)", "33809M", "Comprado",
     "https://www.kleintools.com/catalog/search?search=33809"),
    ("Milwaukee", "Herramienta", "Electrica M12", "FUEL Installation Drill/Driver 4-en-1", "2505-20", "Ya tengo",
     "https://www.milwaukeetool.com/Products/Power-Tools/Drilling/Installation-Drill-Drivers/M12-FUEL-Installation-Drill-Driver/2505-20"),
    ("Milwaukee", "Herramienta", "Electrica M12", "FUEL Stubby Impact Wrench", "2563-20", "Ya tengo",
     "https://www.milwaukeetool.com/Search?q=2563-20"),
    ("Milwaukee", "Herramienta", "Electrica M12", "FUEL Hammer Drill", "3404-20", "Ya tengo",
     "https://www.milwaukeetool.com/Search?q=3404-20"),
    ("Milwaukee", "Herramienta", "Electrica M12", "FUEL Impact Driver", "3453-20", "Ya tengo",
     "https://www.milwaukeetool.com/Search?q=3453-20"),
    ("Milwaukee", "Herramienta", "Electrica M12", "FUEL Oscillating Multi-Tool", "2526-20", "Ya tengo",
     "https://www.milwaukeetool.com/Search?q=2526-20"),
    ("Milwaukee", "Herramienta", "Electrica M12", "ROVER Flood Light", "2367-20", "Ya tengo",
     "https://www.milwaukeetool.com/Search?q=2367-20"),
    ("Milwaukee", "Herramienta", "Electrica M12", 'FUEL 3" Compact Cut Off Tool', "2522-20", "Comprado",
     "https://www.milwaukeetool.com/Search?q=2522-20"),
    ("Milwaukee", "Herramienta", "Electrica M18", "M18 FUEL Sawzall (Sierra Sable)", "2719-20", "Comprado",
     "https://www.milwaukeetool.com/Search?q=2719-20"),
    ("Milwaukee", "Accesorio", "Electrica M18", "Bateria 18V XC 5.0 Ah", "Incluida", "Comprado",
     "https://www.milwaukeetool.com/Products/Power-Tools/Batteries-Chargers/Batteries"),
    ("Milwaukee", "Accesorio", "Electrica M18", "Cargador 12-18V", "Incluido", "Comprado",
     "https://www.milwaukeetool.com/Products/Power-Tools/Batteries-Chargers/Chargers"),
    ("Milwaukee", "Organizacion", "Organizacion", "PACKOUT Backpack", "48-22-8301", "Comprado",
     "https://www.milwaukeetool.com/Search?q=48-22-8301"),
    ("Milwaukee", "Organizacion", "Organizacion", "PACKOUT Compact Organizer", "48-22-8435", "Comprado",
     "https://www.milwaukeetool.com/Search?q=48-22-8435"),
    ("Skil", "Herramienta", "Electrica", "Martillo Automatico", "AH6552A-00", "Ya tengo",
     "https://www.skil.com/search?q=AH6552A"),
    ("Antive", "Herramienta", "Electrica", "Tijeras Electricas Inalambricas", "E-S01", "Ya tengo",
     "https://www.amazon.com/s?k=Antive+E-S01+electric+scissors"),
    ("Huepar", "Herramienta", "Precision", "Nivel Laser Verde Cross-Line", "BOX-1G", "Ya tengo",
     "https://huepar.com/products/huepar-box1g-laser-level"),
    ("Franklin", "Herramienta", "Precision", "Nivel Digital Electronico", "iA12", "Ya tengo",
     "https://www.amazon.com/s?k=Franklin+Sensors+iA12"),
    ("LEXIVON", "Herramienta", "Soldadura", "Soldador de Butano 7 puntas", "LX-770", "Ya tengo",
     "https://www.lexivon.com/products/lx-770"),
    ("Milwaukee", "Seguridad", "Seguridad", "Lentes de seguridad", "48-73-2013", "Ya tengo",
     "https://www.milwaukeetool.com/Search?q=48-73-2013"),
    ("Milwaukee", "Seguridad", "Seguridad", "Guantes PACKOUT Cut Level 1 Smartswipe", "48-73-2013", "Ya tengo",
     "https://www.milwaukeetool.com/Search?q=PACKOUT+Cut+Level+Smartswipe+gloves"),
    ("Milwaukee", "Accesorio", "Accesorio", "Juego de 3 hojas para Multi-Tool", "49-10-9001", "Comprado",
     "https://www.milwaukeetool.com/Search?q=49-10-9001"),
]

PENDIENTES = [
    (1, "Klein", "KNECT 15-Piece Pass Through Socket Set", "65400", None, "Sockets",
     "https://www.kleintools.com/catalog/search?search=65400"),
    (2, "Klein", "Clamp Meter 600A", "CL810", 80000, "Medicion",
     "https://www.kleintools.com/catalog/search?search=CL810"),
    (3, "Milwaukee", "PACKOUT Rolling Drawer Tool Box", "48-22-8420", None, "Organizacion",
     "https://www.milwaukeetool.com/Search?q=48-22-8420"),
    (4, "Milwaukee", "PACKOUT 2-Wheel Utility Cart", "48-22-8415", None, "Organizacion",
     "https://www.milwaukeetool.com/Search?q=48-22-8415"),
    (5, "Milwaukee", "M12 AIRSNAKE", "2572B-21", 280000, "Fontaneria",
     "https://www.milwaukeetool.com/Search?q=2572B-21"),
]

# Fix TwinGrip model page - Knipex TwinGrip is often 82 01 200 not 83 01 200
# Inventory says 83 01 200 - keep model as listed, link to TwinGrip line
INVENTARIO[3] = (
    "Knipex", "Herramienta", "Fontaneria", 'TwinGrip Pliers 8"', "83 01 200", "Ya tengo",
    "https://www.knipex.com/products/pliers/front-and-side-gripping-pliers/knipex-twingrip-front-and-side-gripping-pliers/8201200",
)

wb = Workbook()

header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
title_font = Font(name="Arial", bold=True, size=16, color="1F2937")
subtitle_font = Font(name="Arial", size=10, color="4B5563")
body_font = Font(name="Arial", size=10)
link_font = Font(name="Arial", size=10, color="0563C1", underline="single")
header_fill = PatternFill("solid", fgColor="1E3A5F")
alt_fill = PatternFill("solid", fgColor="F3F4F6")
comprado_fill = PatternFill("solid", fgColor="DCFCE7")
tengo_fill = PatternFill("solid", fgColor="DBEAFE")
pend_fill = PatternFill("solid", fgColor="FEF3C7")
resumen_fill = PatternFill("solid", fgColor="EEF2FF")
thin = Border(
    left=Side(style="thin", color="D1D5DB"),
    right=Side(style="thin", color="D1D5DB"),
    top=Side(style="thin", color="D1D5DB"),
    bottom=Side(style="thin", color="D1D5DB"),
)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left = Alignment(horizontal="left", vertical="center", wrap_text=True)

# ===== Inventario =====
ws = wb.active
ws.title = "Inventario"
ws["A1"] = "Inventario - Arsenal Multioficio 2026"
ws["A1"].font = title_font
ws.merge_cells("A1:I1")
ws["A2"] = "Propietario: Saul Marin Gonzalez"
ws["A2"].font = subtitle_font
ws["C2"] = "Ultima actualizacion: 28/07/2026"
ws["C2"].font = subtitle_font
ws["E2"] = "Total items:"
ws["E2"].font = Font(name="Arial", bold=True, size=10)
ws["F2"] = f"=COUNTA(E5:E{4 + len(INVENTARIO)})"
ws["F2"].font = Font(name="Arial", bold=True, size=12, color="1E3A5F")
ws["A3"] = "Columna Enlace: ficha o busqueda en el sitio oficial de la marca (clic para abrir)."
ws["A3"].font = Font(name="Arial", size=9, color="6B7280", italic=True)
ws.merge_cells("A3:I3")

headers = ["#", "Marca", "Tipo", "Categoria", "Herramienta", "Modelo", "Estado", "Enlace marca", "URL"]
for col, h in enumerate(headers, 1):
    cell = ws.cell(4, col, h)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center
    cell.border = thin

for i, row in enumerate(INVENTARIO, 1):
    r = i + 4
    marca, tipo, cat, nombre, modelo, estado, url = row
    values = [i, marca, tipo, cat, nombre, modelo, estado]
    for col, val in enumerate(values, 1):
        cell = ws.cell(r, col, val)
        cell.font = body_font
        cell.border = thin
        cell.alignment = center if col in (1, 2, 3, 6, 7) else left
        if i % 2 == 0 and col < 7:
            cell.fill = alt_fill
    est = ws.cell(r, 7)
    est.fill = comprado_fill if estado == "Comprado" else tengo_fill

    # Clickable brand link
    link_cell = ws.cell(r, 8, f"Ver en {marca}")
    link_cell.font = link_font
    link_cell.alignment = center
    link_cell.border = thin
    link_cell.hyperlink = url

    url_cell = ws.cell(r, 9, url)
    url_cell.font = Font(name="Arial", size=8, color="6B7280")
    url_cell.alignment = left
    url_cell.border = thin
    url_cell.hyperlink = url

last_inv = 4 + len(INVENTARIO)
dv = DataValidation(type="list", formula1='"Comprado,Ya tengo"', allow_blank=False)
ws.add_data_validation(dv)
dv.add(f"G5:G{last_inv}")

for i, w in enumerate([5, 12, 12, 14, 42, 14, 11, 16, 55], 1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.auto_filter.ref = f"A4:I{last_inv}"
ws.freeze_panes = "A5"
ws.row_dimensions[1].height = 24
ws.row_dimensions[4].height = 22

# ===== Pendientes =====
wp = wb.create_sheet("Pendientes")
wp["A1"] = "Pendientes de compra - Arsenal Multioficio 2026"
wp["A1"].font = title_font
wp.merge_cells("A1:H1")
wp["A2"] = "Prioridad 1 = mas urgente. Montos en colones (CRC). Enlaces al sitio de la marca."
wp["A2"].font = subtitle_font

ph = ["Prioridad", "Marca", "Herramienta", "Modelo", "Precio (CRC)", "Notas", "Enlace marca", "URL"]
for col, h in enumerate(ph, 1):
    cell = wp.cell(4, col, h)
    cell.font = header_font
    cell.fill = PatternFill("solid", fgColor="B45309")
    cell.alignment = center
    cell.border = thin

for i, row in enumerate(PENDIENTES, 1):
    r = i + 4
    prio, marca, nombre, modelo, precio, notas, url = row
    for col, val in enumerate([prio, marca, nombre, modelo, precio, notas], 1):
        cell = wp.cell(r, col, val)
        cell.font = body_font
        cell.border = thin
        cell.alignment = center if col in (1, 2, 4) else left
        cell.fill = pend_fill
        if col == 5 and val is not None:
            cell.number_format = '"₡"#,##0'
    lc = wp.cell(r, 7, f"Ver en {marca}")
    lc.font = link_font
    lc.alignment = center
    lc.border = thin
    lc.fill = pend_fill
    lc.hyperlink = url
    uc = wp.cell(r, 8, url)
    uc.font = Font(name="Arial", size=8, color="6B7280")
    uc.border = thin
    uc.fill = pend_fill
    uc.hyperlink = url

last_p = 4 + len(PENDIENTES)
wp.cell(last_p + 2, 4, "Total estimado (con precio):")
wp.cell(last_p + 2, 4).font = Font(name="Arial", bold=True, size=10)
wp.cell(last_p + 2, 4).alignment = Alignment(horizontal="right")
wp.cell(last_p + 2, 5, f"=SUM(E5:E{last_p})")
wp.cell(last_p + 2, 5).font = Font(name="Arial", bold=True, size=11, color="B45309")
wp.cell(last_p + 2, 5).number_format = '"₡"#,##0'
wp.cell(last_p + 2, 5).fill = pend_fill

for i, w in enumerate([12, 12, 40, 14, 14, 14, 16, 50], 1):
    wp.column_dimensions[get_column_letter(i)].width = w
wp.freeze_panes = "A5"
wp.auto_filter.ref = f"A4:H{last_p}"

# ===== Resumen =====
wr = wb.create_sheet("Resumen", 0)
wr["A1"] = "Arsenal Multioficio 2026 - Resumen"
wr["A1"].font = title_font
wr.merge_cells("A1:D1")
wr["A2"] = "Saul Marin Gonzalez | Respaldo local | Enlaces a sitios de marca"
wr["A2"].font = subtitle_font

wr["A4"] = "Metricas del inventario"
wr["A4"].font = Font(name="Arial", bold=True, size=12, color="1E3A5F")

for cell_ref, label in [
    ("A5", "Total de items en inventario"),
    ("A6", "Estado: Ya tengo"),
    ("A7", "Estado: Comprado"),
    ("A8", "Pendientes de compra"),
    ("A9", "Costo estimado pendientes (CRC)"),
]:
    wr[cell_ref] = label
    wr[cell_ref].font = body_font
    wr[cell_ref].fill = resumen_fill
    wr[cell_ref].border = thin

wr["B5"] = f"=COUNTA(Inventario!E5:E{last_inv})"
wr["B6"] = f'=COUNTIF(Inventario!G5:G{last_inv},"Ya tengo")'
wr["B7"] = f'=COUNTIF(Inventario!G5:G{last_inv},"Comprado")'
wr["B8"] = f"=COUNTA(Pendientes!C5:C{last_p})"
wr["B9"] = f"=SUM(Pendientes!E5:E{last_p})"
wr["B9"].number_format = '"₡"#,##0'
for r in range(5, 10):
    wr.cell(r, 2).font = Font(name="Arial", bold=True, size=12)
    wr.cell(r, 2).border = thin
    wr.cell(r, 2).alignment = center

wr["A11"] = "Sitios oficiales de marca (inicio)"
wr["A11"].font = Font(name="Arial", bold=True, size=12, color="1E3A5F")
brand_home = [
    ("Milwaukee", "https://www.milwaukeetool.com/"),
    ("Klein Tools", "https://www.kleintools.com/"),
    ("Knipex", "https://www.knipex.com/"),
    ("Wera", "https://www.wera.de/en"),
    ("Huepar", "https://huepar.com/"),
    ("LEXIVON", "https://www.lexivon.com/"),
    ("Skil", "https://www.skil.com/"),
]
wr["A12"] = "Marca"
wr["B12"] = "Sitio web"
for col in (1, 2):
    wr.cell(12, col).font = header_font
    wr.cell(12, col).fill = header_fill
    wr.cell(12, col).border = thin
for i, (name, url) in enumerate(brand_home):
    r = 13 + i
    wr.cell(r, 1, name).font = body_font
    wr.cell(r, 1).border = thin
    c = wr.cell(r, 2, url)
    c.font = link_font
    c.hyperlink = url
    c.border = thin

last_b = 12 + len(brand_home)
wr.cell(last_b + 2, 1, "Items por marca").font = Font(name="Arial", bold=True, size=12, color="1E3A5F")
hdr = last_b + 3
wr.cell(hdr, 1, "Marca").font = header_font
wr.cell(hdr, 1).fill = header_fill
wr.cell(hdr, 1).border = thin
wr.cell(hdr, 2, "Cantidad").font = header_font
wr.cell(hdr, 2).fill = header_fill
wr.cell(hdr, 2).border = thin

marcas = sorted(set(r[0] for r in INVENTARIO))
for i, m in enumerate(marcas):
    r = hdr + 1 + i
    wr.cell(r, 1, m).font = body_font
    wr.cell(r, 1).border = thin
    wr.cell(r, 2, f"=COUNTIF(Inventario!B5:B{last_inv},A{r})")
    wr.cell(r, 2).font = body_font
    wr.cell(r, 2).border = thin
    wr.cell(r, 2).alignment = center

wr["D4"] = "Ecosistema"
wr["D4"].font = Font(name="Arial", bold=True, size=12, color="1E3A5F")
for i, t in enumerate(
    [
        "Base principal: Milwaukee M12 + M18",
        "Prueba / electrico: Klein Tools",
        "Alicates y llaves: Knipex + Wera Joker",
        "Organizacion: PACKOUT",
        "Enlaces: columna H = clic; I = URL completa",
    ]
):
    wr.cell(5 + i, 4, t).font = body_font
    wr.cell(5 + i, 4).fill = resumen_fill
    wr.cell(5 + i, 4).border = thin

wr["D11"] = "Notas sobre enlaces"
wr["D11"].font = Font(name="Arial", bold=True, size=11)
notes = [
    "Prioridad: ficha en sitio de la marca.",
    "Si no hay ficha estable, se usa busqueda del catalogo oficial.",
    "Antive/Franklin: poco catalogo web; enlace de busqueda de producto.",
    "Guantes: SKU 48-73-2013 era el de lentes; link generico PACKOUT gloves.",
    "Verifica el modelo exacto en el sitio antes de comprar.",
]
for i, t in enumerate(notes):
    wr.cell(12 + i, 4, t).font = Font(name="Arial", size=9, color="6B7280")

wr.column_dimensions["A"].width = 36
wr.column_dimensions["B"].width = 48
wr.column_dimensions["C"].width = 3
wr.column_dimensions["D"].width = 55

from pathlib import Path

out = Path(__file__).resolve().parent.parent / "inventario" / "Arsenal_Multioficio_2026.xlsx"
out.parent.mkdir(parents=True, exist_ok=True)
wb.save(out)
print("OK", out)
print("Inventario:", len(INVENTARIO), "Pendientes:", len(PENDIENTES))
