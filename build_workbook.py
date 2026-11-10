"""Builds irish_property_2024_2025.xlsx from the Property Price Register (PPR).

Sheets
  Data     : 2024-2025 full-market-price sales (date, county, region, price, type)
  Regions  : county -> region lookup table (used by XLOOKUP)
  Summary  : COUNTIFS / AVERAGEIFS / SUMIFS by county and year, plus a chart
  Notes    : what each formula does
"""
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

SRC = "../irish-property-sql/data/raw/PPR-ALL.csv"
OUT = "irish_property_2024_2025.xlsx"

REGION = {
    "Dublin": "Dublin", "Cork": "South", "Kerry": "South", "Waterford": "South", "Tipperary": "South",
    "Limerick": "Mid-West", "Clare": "Mid-West", "Galway": "West", "Mayo": "West", "Roscommon": "West",
    "Sligo": "North-West", "Leitrim": "North-West", "Donegal": "North-West", "Kildare": "Mid-East",
    "Meath": "Mid-East", "Wicklow": "Mid-East", "Louth": "Mid-East", "Wexford": "South-East",
    "Kilkenny": "South-East", "Carlow": "South-East", "Laois": "Midlands", "Offaly": "Midlands",
    "Westmeath": "Midlands", "Longford": "Midlands", "Cavan": "Border", "Monaghan": "Border",
}

raw = pd.read_csv(SRC, encoding="latin-1", low_memory=False)
raw.columns = ["date", "address", "county", "eircode", "price", "not_full_market",
               "vat_exclusive", "description", "size"]
raw["date"] = pd.to_datetime(raw["date"], dayfirst=True)
raw["price"] = raw["price"].astype(str).str.replace(r"[^\d.]", "", regex=True).astype(float)
df = raw[(raw["date"].dt.year.isin([2024, 2025])) & (raw["not_full_market"] == "No")].copy()
df["year"] = df["date"].dt.year
df["type"] = df["description"].str.contains("New", case=False).map({True: "New", False: "Second-hand"})
df = df[["date", "year", "county", "price", "type"]].sort_values("date")

wb = Workbook()
bold = Font(bold=True)
head = PatternFill("solid", fgColor="DDDDDD")

# --- Data sheet (region filled by XLOOKUP formula, not hard-coded) ---
ws = wb.active
ws.title = "Data"
cols = ["Date", "Year", "County", "Region", "Price (EUR)", "Type"]
ws.append(cols)
for i, r in enumerate(df.itertuples(index=False), start=2):
    ws.append([r.date.date(), r.year, r.county,
               f'=IFERROR(XLOOKUP(C{i},Regions!$A:$A,Regions!$B:$B),INDEX(Regions!$B:$B,MATCH(C{i},Regions!$A:$A,0)))', r.price, r.type])
n = len(df) + 1
for c in range(1, 7):
    ws.cell(1, c).font = bold
    ws.cell(1, c).fill = head
ws.column_dimensions["A"].width = 12
ws.column_dimensions["C"].width = 12
ws.column_dimensions["D"].width = 12
ws.column_dimensions["E"].width = 13
for row in ws.iter_rows(min_row=2, min_col=5, max_col=5):
    row[0].number_format = "#,##0"
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:F{n}"

# --- Regions lookup ---
wr = wb.create_sheet("Regions")
wr.append(["County", "Region"])
for k, v in sorted(REGION.items()):
    wr.append([k, v])
wr["A1"].font = wr["B1"].font = bold

# --- Summary with live formulas ---
s = wb.create_sheet("Summary")
s.append(["County", "Sales 2024", "Sales 2025", "Avg price 2024", "Avg price 2025",
          "Change %", "Sales over EUR 500k (2025)", "New-build share 2025"])
counties = sorted(df["county"].unique())
for j, cty in enumerate(counties, start=2):
    D = f"Data!$C$2:$C${n}"; Y = f"Data!$B$2:$B${n}"; P = f"Data!$E$2:$E${n}"; T = f"Data!$F$2:$F${n}"
    s.append([cty,
              f"=COUNTIFS({D},A{j},{Y},2024)",
              f"=COUNTIFS({D},A{j},{Y},2025)",
              f"=AVERAGEIFS({P},{D},A{j},{Y},2024)",
              f"=AVERAGEIFS({P},{D},A{j},{Y},2025)",
              f"=E{j}/D{j}-1",
              f'=COUNTIFS({D},A{j},{Y},2025,{P},">500000")',
              f'=COUNTIFS({D},A{j},{Y},2025,{T},"New")/C{j}'])
last = len(counties) + 1
for c in range(1, 9):
    s.cell(1, c).font = bold
    s.cell(1, c).fill = head
    s.column_dimensions[get_column_letter(c)].width = 16
for r in range(2, last + 1):
    for c in (4, 5):
        s.cell(r, c).number_format = "#,##0"
    s.cell(r, 6).number_format = "0.0%"
    s.cell(r, 8).number_format = "0.0%"
s.freeze_panes = "B2"

chart = BarChart()
chart.title = "Average sale price by county, 2025 (EUR)"
chart.y_axis.title = "EUR"
chart.add_data(Reference(s, min_col=5, min_row=1, max_row=last), titles_from_data=True)
chart.set_categories(Reference(s, min_col=1, min_row=2, max_row=last))
chart.width, chart.height = 26, 11
s.add_chart(chart, f"A{last + 3}")

# --- Notes ---
nt = wb.create_sheet("Notes")
for line in [
    "Source: Property Price Register (propertypriceregister.ie), sales in 2024 and 2025, full market price only.",
    "Data!D (Region): XLOOKUP(county, Regions!A:A, Regions!B:B) finds each county's region; INDEX/MATCH is the fallback for older Excel versions.",
    "Caution: averages are pulled up by a few very expensive sales (e.g. Wicklow). A median would be more robust; new-build prices exclude VAT.",
    "Summary!B-C: COUNTIFS counts sales that match county AND year.",
    "Summary!D-E: AVERAGEIFS averages price for rows that match county AND year.",
    "Summary!F: change in average price, 2025 vs 2024.",
    "Summary!G: COUNTIFS with a price condition (\">500000\").",
    "Summary!H: share of 2025 sales that were new builds.",
    "Exercise: insert a PivotTable from Data (Rows = Region, Columns = Year, Values = Average of Price).",
]:
    nt.append([line])
nt.column_dimensions["A"].width = 110

wb.save(OUT)
print(OUT, "rows:", len(df), "counties:", len(counties))
