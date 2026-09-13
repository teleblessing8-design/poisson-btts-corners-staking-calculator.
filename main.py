import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Create a new workbook and select active sheet
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "10-Game Matrix & CLV"

# Ensure grid lines are visible
ws.views.sheetView[0].showGridLines = True

# --- 1. SET UP GLOBAL VARIABLES ---
ws['A2'] = "ANALYST GLOBAL PARAMETERS"
ws['A2'].font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
ws['A2'].fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")

ws['A4'] = "Total Bankroll:"
ws['B4'] = 500000  # ₦500,000 Starting Bankroll
ws['B4'].number_format = '₦#,##0.00'

ws['A5'] = "Kelly Multiplier:"
ws['B5'] = 0.25  # Quarter-Kelly

# --- 2. SET UP KPI DASHBOARD ---
ws['N4'] = "TOTAL PERFORMANCE"
ws['N4'].font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
ws['N4'].fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")

ws['N5'] = "Total Profit/Loss:"
ws['O5'] = "=SUM(P10:P19)-SUM(M10:M19)"
ws['O5'].number_format = '₦#,##0.00'

ws['N6'] = "Average CLV Edge:"
ws['O6'] = "=AVERAGE(Q10:Q19)"
ws['O6'].number_format = '0.00%'

# Style global and KPI blocks
for row in:
    ws[f'A{row}'].font = Font(name="Arial", size=10, bold=True)
    ws[f'N{row}'].font = Font(name="Arial", size=10, bold=True)
    ws[f'O{row}'].font = Font(name="Arial", size=10, bold=True, color="0070C0" if row==5 else "000000")

# --- 3. MATRIX DATA & HEADERS ---
headers = [
    "Game #", "Matchup", "Home λ (xG)", "Away λ (xG)", 
    "P(Home=0)", "P(Away=0)", "P(BTTS Yes)", "Fair Odds", 
    "Bookie Odds", "Edge (EV)", "Raw Kelly", "Fractional Kelly", 
    "Actual Stake", "Closing Odds", "Bet Status", "Net Return", "CLV Edge %"
]

header_fill = PatternFill(start_color="333333", end_color="333333", fill_type="solid")
header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")

for col_num, header_title in enumerate(headers, 1):
    cell = ws.cell(row=9, column=col_num)
    cell.value = header_title
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Example match data placeholders
sample_matches = [
    ("Enyimba vs Kano Pillars", 1.85, 1.20, 1.95),
    ("Remo Stars vs Rivers United", 1.40, 0.95, 2.15),
    ("Shooting Stars vs Rangers Intl", 1.10, 1.10, 2.10),
    ("Bendel Insurance vs Lobi Stars", 2.10, 0.80, 1.85),
    ("Plateau United vs Kwara United", 1.65, 1.35, 1.75),
    ("Sunshine Stars vs Heartland", 1.25, 0.75, 2.30),
    ("Akwa United vs Katsina United", 1.50, 1.10, 1.90),
    ("Bayelsa United vs Abia Warriors", 1.70, 1.45, 1.65),
    ("El-Kanemi vs Gombe United", 2.00, 0.90, 1.80),
    ("Sporting Lagos vs Niger Tornadoes", 1.30, 1.30, 1.95)
]

# Thin borders for columns
thin_side = Side(border_style="thin", color="D9D9D9")
border_box = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

# --- 4. POPULATE DYNAMIC FORMULAS ---
for i, match in enumerate(sample_matches):
    r = 10 + i
    
    # Inputs
    ws.cell(row=r, column=1, value=i+1)                    # Game ID
    ws.cell(row=r, column=2, value=match[0])               # Matchup
    ws.cell(row=r, column=3, value=match[1])               # Home λ
    ws.cell(row=r, column=4, value=match[2])               # Away λ
    ws.cell(row=r, column=9, value=match[3])               # Taken Bookie Odds
    ws.cell(row=r, column=14, value=match[3] - 0.15)       # Closing Odds Dummy (Mock drop)
    ws.cell(row=r, column=15, value="Win" if i % 2 == 0 else "Loss") # Dummy Outcome Status
    
    # Active Formulas
    ws.cell(row=r, column=5, value=f"=EXP(-C{r})")         # P(Home=0)
    ws.cell(row=r, column=6, value=f"=EXP(-D{r})")         # P(Away=0)
    ws.cell(row=r, column=7, value=f"=(1-E{r})*(1-F{r})")   # P(BTTS Yes)
    ws.cell(row=r, column=8, value=f"=1/G{r}")             # Implied Fair Odds
    ws.cell(row=r, column=10, value=f"=(G{r}*I{r})-1")      # Expected Value (Edge)
    ws.cell(row=r, column=11, value=f"=IF(J{r}>0, J{r}/(I{r}-1), 0)") # Raw Kelly
    ws.cell(row=r, column=12, value=f"=K{r}*$B$5")          # Fractional Kelly
    ws.cell(row=r, column=13, value=f"=L{r}*$B$4")          # Actual Stake in Naira
    ws.cell(row=r, column=16, value=f'=IF(O{r}="Win", M{r}*I{r}, IF(O{r}="Loss", 0, M{r}))') # Returns
    ws.cell(row=r, column=17, value=f"=IF(N{r}>0, (I{r}/N{r})-1, 0)") # CLV Edge %

    # Apply Formats & Alignment
    for c in range(1, 18):
        cell = ws.cell(row=r, column=c)
        cell.font = Font(name="Arial", size=10)
        cell.border = border_box
        if c in:
            cell.alignment = Alignment(horizontal="center")
        elif c > 2:
            cell.alignment = Alignment(horizontal="right")
            
        # Number formats
        if c in:
            cell.number_format = '0.00%'
        elif c in:
            cell.number_format = '0.00'
        elif c in:
            cell.number_format = '₦#,##0.00'

# Auto-adjust column widths for flawless visualization
for col in ws.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

# Save the configured sheet
file_name = "btts_poisson_clv_naira_calculator.xlsx"
wb.save(file_name)
print(f"Success! Your master spreadsheet has been generated as: {file_name}")

