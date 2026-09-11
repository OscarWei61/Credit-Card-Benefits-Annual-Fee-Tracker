#!/usr/bin/env python3
"""
Generates credit_card_tracker.xlsx from the design in PLAN.md.

Design source of truth: PLAN.md (sections 3-6). Every formula below is copied from
the spec's formula map; cell addresses are contractual (Dashboard reads them).

Usage:  ./.venv/bin/python build_tracker.py
"""
import json
import os
import sys
import shutil
from datetime import date
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

OUT = "Oscar_credit_tracker.xlsx"
YEAR = 2026
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# ---------------------------------------------------------------- styling ---
BOLD = Font(bold=True)
TITLE = Font(bold=True, size=14, color="1F3864")
HDR_FILL = PatternFill("solid", start_color="D9E1F2")
TOT_FILL = PatternFill("solid", start_color="F2F2F2")
GREEN = PatternFill("solid", start_color="C6EFCE")
ORANGE = PatternFill("solid", start_color="FFD8A8")
AMBER = PatternFill("solid", start_color="FFE699")
GREY_FONT = Font(color="A6A6A6")
RED_FONT = Font(color="C00000", bold=True)
GREEN_FONT = Font(color="006100", bold=True)
RED_BG = PatternFill("solid", start_color="FFC7CE")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
MONEY = '"$"#,##0'
PCT = "0%"
DATEFMT = "yyyy-mm-dd"

# --------------------------------------------------------------- card data ---
# benefit rows: (name, basis, cycle, value_per_cycle, note)
#   basis  SC = statement credit (counts in net benefit)
#          EST = estimated value (counts only in the "incl. estimated" view)
#          ACCESS = perk with no dollar value (counts uses only)
# `value`: None = leave blank for the user to fill (nothing invented here)
BAG_NOTE = ("SET YOUR VALUE in 'Value/cycle' (D) = how much you expect to save on checked bags this year "
            "(default 350 = about 10 bags at $35). Then type the fee of each bag into the MONTH you checked it - "
            "they add up with no monthly limit; the row is only capped at D. Raise D if you hit the cap (the cell "
            "turns orange). This card covers you + up to %s, so a family trip saves more per bag. "
            "Basis = EST, so it lands in the 'incl. estimated' totals; switch Basis to SC to put it in the SC headline.")

OSCAR_CARDS = [
    dict(tab="CITI-COSTCO", issuer="Citi", product="Costco Anywhere Visa Card by Citi",
         opened="2025-09-19", af=0, waived=False, au_af=0, status="Active",
         af_note="No annual fee; requires a paid Costco membership. This card has no statement credits, so there is nothing to track here.",
         cal=[], cmy=[], multi=[]),
    dict(tab="AMEX-ASPIRE", issuer="Amex", product="Hilton Honors Aspire Card",
         opened="2026-03-29", af=550, waived=False, au_af=0, status="Active", af_note="",
         cal=[("Hilton Resort Credit", "SC", "SEMI-ANNUAL", 200,
               "$200 Jan-Jun + $200 Jul-Dec; participating Hilton Resorts, charged to the room; prepaid/non-refundable rates excluded"),
              ("Airline Fee Credit", "SC", "QUARTERLY", 50,
               "$50 per quarter = $200/yr. Scope CONFLICTS between sources: Amex terms say incidental fees (checked bag, change/cancel, internet); one aggregator says airfare also counts. Confirm on your Amex benefits page"),
              ("CLEAR+ Credit", "SC", "CALENDAR-YEAR", 219,
               "Auto-renewing CLEAR+ membership charged to this card")],
         cmy=[], multi=[]),
    dict(tab="CHASE-CSP", issuer="Chase", product="Chase Sapphire Preferred",
         opened="2025-03-27", af=95, waived=False, au_af=0, status="Active",
         af_note="The June 2026 refresh added the $120 Global Entry credit and doubled the hotel credit to $100 - automatic for existing cardmembers.",
         cal=[("DoorDash non-restaurant credit", "SC", "MONTHLY", 10,
               "$10/mo promo through 12/31/2027; requires DashPass activated"),
              ],
         cmy=[("Chase Travel hotel credit", "SC", "CARDMEMBER-YEAR", 100,
               "OFFICIAL (Chase.com): '$100 Annual Chase Travel Hotel Credit' - resets every year on the cardmember's anniversary and unused amounts are forfeited. Prepaid hotels booked through Chase Travel only. No activation needed")],
         multi=[("Global Entry / TSA PreCheck / NEXUS credit", "SC", "MULTI-YEAR", 120, 4,
                 "OFFICIAL (Chase press release 2026-06-15): '$120 Global Entry, TSA PreCheck, or NEXUS credit every four years'. Added 2026-06-15, no activation needed. Enter the amount in the 'Used?' column when you claim it")]),
    dict(tab="BOA-ATMOS-ASCENT", issuer="Bank of America", product="Atmos Rewards Ascent Visa Signature",
         opened="2025-09-19", af=95, waived=False, au_af=0, status="Active",
         af_note="Formerly the Alaska Airlines Visa Signature. No statement credits on this card - the only tracked benefit is the free checked bag.",
         cal=[("Free checked bag", "EST", "CALENDAR-YEAR", 350, BAG_NOTE % "6 guests")],
         cmy=[], multi=[]),
    dict(tab="RH-GOLD", issuer="Robinhood", product="Robinhood Gold Card",
         opened="2026-06-01", af=60, waived=False, au_af=0, status="Active",
         af_note="Annual fee = $60 = the required Robinhood Gold subscription at $5/mo ($50/yr if billed annually - set this to what you actually pay). The card itself is $0. No statement credits.",
         cal=[], cmy=[], multi=[]),
    dict(tab="AMEX-DELTA-BLUE", issuer="Amex", product="Delta SkyMiles Blue Card",
         opened="2025-06-23", af=0, waived=False, au_af=0, status="Active",
         af_note="No annual fee. No statement credits and no free checked bag on this card.",
         cal=[], cmy=[], multi=[]),
    dict(tab="CITI-AA-PLAT", issuer="Citi", product="Citi / AAdvantage Platinum Select World Elite Mastercard",
         opened="2025-12-24", af=99, waived=True, au_af=0, status="Active",
         af_note="First-year annual fee is waived ($0 intro). Untick 'First-year AF waived?' when the 1st anniversary hits (Dec 2026).",
         cal=[("Free checked bag", "EST", "CALENDAR-YEAR", 350, BAG_NOTE % "4 companions")],
         cmy=[],
         multi=[("Turo credit", "SC", "ONE-TIME", 180, 1,
                 "Up to $30 statement credit per completed Turo trip, max $180 total, for trips completed 2025-10-19 through 2026-10-18. EXPIRES 2026-10-18 - link this card to your Turo account, then enter the amount captured")]),
    dict(tab="CHASE-CFU", issuer="Chase", product="Chase Freedom Unlimited",
         opened="2024-12-09", af=0, waived=False, au_af=0, status="Active",
         af_note="No annual fee. No statement credits on this card.", cal=[], cmy=[], multi=[]),
    dict(tab="AMEX-GREEN", issuer="Amex", product="Green Card from American Express (now 'Classic Green')",
         opened="2025-09-19", af=150, waived=False, au_af=0, status="Active",
         af_note="The $100 LoungeBuddy credit ended 2025-01-13, before this card was opened.",
         cal=[("CLEAR+ Credit", "SC", "CALENDAR-YEAR", 219,
               "Was $189, then $199; CLEAR raised its membership price to $219 on 2026-07-01 and Amex lifted the credit to match. Confirm the current amount on your Amex benefits page")],
         cmy=[], multi=[]),
    dict(tab="BOA-CUSTOMIZED-CASH", issuer="Bank of America", product="Bank of America Customized Cash Rewards",
         opened="2024-08-17", af=0, waived=False, au_af=0, status="Active",
         af_note="No annual fee. This card has NO statement credits.", cal=[], cmy=[], multi=[]),
    dict(tab="AMEX-BCE", issuer="Amex", product="Blue Cash Everyday Card",
         opened="2024-08-14", af=0, waived=False, au_af=0, status="Active",
         af_note="No annual fee.",
         cal=[("Disney Streaming Credit", "SC", "MONTHLY", 7,
               "OFFICIAL (Amex): '$84 Disney Streaming Credit ... up to a $7 monthly statement credit' at DisneyPlus.com / Hulu.com / Stream.ESPN.com; enrollment required"),
              ("Home Chef credit", "SC", "MONTHLY", 15,
               "PENDING - Amex's current Blue Cash Everyday page shows only the $84 Disney credit and does NOT list Home Chef. Confirm in the Amex app, then keep or delete this row")],
         cmy=[], multi=[]),
]


# Myra's cards (profile: myra)
PARTNER_CARDS = [
    dict(tab="CHASE-UNITED-QUEST", issuer="Chase", product="United Quest Card",
         opened="2026-01-01", af=350, waived=False, au_af=0, status="Active",
         af_note="Annual fee $350 in 2026 (was $250 before). Open date is a PLACEHOLDER - change B5 and B10 to the real approval date.",
         cal=[("Rideshare credit", "SC", "MONTHLY", 8,
               "OFFICIAL (Chase.com): up to $100 per calendar year - $8 back each month Jan-Nov and up to $12 in December. Enrollment required. December is worth more, so type 12 there if you use it"),
              ("Instacart credits", "SC", "MONTHLY", 15,
               "OFFICIAL (Chase.com): one $10 plus one $5 Instacart credit each month, up to $180 per calendar year. Benefits end 12/31/2027"),
              ("Free checked bag", "EST", "CALENDAR-YEAR", 220,
               "OFFICIAL (Chase.com): the primary cardholder and one companion on the same reservation each get their 1st and 2nd checked bags free on United-operated flights - up to $50 for the first bag and $60 for the second, each way. D = your expected annual saving (220 = one roundtrip, two people, two bags each). Type each trip's saving into that month; capped at D. Ticket must be bought with the card")],
         cmy=[("United TravelBank credit", "SC", "CARDMEMBER-YEAR", 200,
               "OFFICIAL (Chase.com full-terms page): '$200 Annual United Travel Credit: One $200 annual United TravelBank cash credit will automatically be added to your MileagePlus account. Annual means the year beginning with your account open date through the first statement date after your account open anniversary...' Deposited within 4 weeks of the anniversary. Must be used for a United or United Express flight; expires 12 months after issue and cannot be combined with miles or travel certificates"),
              ("Renowned Hotels and Resorts credit", "SC", "CARDMEMBER-YEAR", 150,
               "OFFICIAL (Chase.com): up to $150 back each anniversary year on prepaid hotel accommodation booked directly through Renowned Hotels and Resorts"),
              ("JSX credit", "SC", "CARDMEMBER-YEAR", 150,
               "OFFICIAL (Chase.com): up to $150 back each anniversary year when you book flights directly with JSX"),
              ("Avis/Budget TravelBank cash", "SC", "CARDMEMBER-YEAR", 80,
               "OFFICIAL (Chase.com): $40 in United TravelBank cash for your 1st and 2nd Avis or Budget rental booked via cars.united.com, up to $80 each anniversary year")],
         multi=[("Global Entry / TSA PreCheck / NEXUS credit", "SC", "MULTI-YEAR", 120, 4,
                 "OFFICIAL (Chase.com): up to $120 every four years. Enter the amount in the 'Used?' column when you claim it")]),
    dict(tab="CHASE-BOUNDLESS", issuer="Chase", product="Marriott Bonvoy Boundless",
         opened="2026-01-01", af=95, waived=False, au_af=0, status="Active",
         af_note="Annual fee $95. Open date is a PLACEHOLDER - change B5 and B10 to the real approval date.",
         cal=[("Airline credit", "SC", "SEMI-ANNUAL", 50,
               "OFFICIAL (Chase.com): 'Unlock up to $100 in Airline Credits' - a $50 statement credit after you spend $250 or more DIRECTLY WITH AIRLINES in Jan-Jun, and another $50 after $250 or more in Jul-Dec (max $100 through 12/31/2026). EXISTING cardmembers must activate the offer once at chase.com/mybonus/marriott-airline-credit or by phone. New cardmembers who applied 6/4/2026-12/31/2026 are auto-enrolled and their second half runs Jan-Jun 2027"),
              ("DoorDash non-restaurant discount", "EST", "QUARTERLY", 10,
               "OFFICIAL (Chase.com): $10 off one qualifying non-restaurant DoorDash order each calendar quarter (up to $40/yr) once you activate the complimentary DashPass. It is a checkout discount, not a statement credit, hence EST")],
         cmy=[("Annual Free Night Award", "EST", "CARDMEMBER-YEAR", None,
               "OFFICIAL (Chase.com): one free night each account anniversary year, valid at Marriott properties up to 35,000 points (you may top up with your own points). Deposited up to 8 weeks after the anniversary and expires 12 months later. NOT a statement credit - enter what the night would have cost you, or delete this row if you only want statement credits")],
         multi=[]),
    dict(tab="CHASE-CSP", issuer="Chase", product="Chase Sapphire Preferred",
         opened="2026-01-01", af=95, waived=False, au_af=0, status="Active",
         af_note="Annual fee $95. The June 2026 refresh applies automatically to existing cardmembers. Open date is a PLACEHOLDER - change B5 and B10.",
         cal=[("DoorDash non-restaurant credit", "SC", "MONTHLY", 10,
               "$10/mo promo through 12/31/2027; requires DashPass activated"),
              ],
         cmy=[("Chase Travel hotel credit", "SC", "CARDMEMBER-YEAR", 100,
               "OFFICIAL (Chase.com): '$100 Annual Chase Travel Hotel Credit' - resets every year on the cardmember's anniversary and unused amounts are forfeited. Prepaid hotels booked through Chase Travel only. No activation needed")],
         multi=[("Global Entry / TSA PreCheck / NEXUS credit", "SC", "MULTI-YEAR", 120, 4,
                 "OFFICIAL (Chase press release 2026-06-15): '$120 Global Entry, TSA PreCheck, or NEXUS credit every four years'. Added 2026-06-15, no activation needed. Enter the amount in the 'Used?' column when you claim it")]),
]

PROFILES = {
    "oscar":   ("Oscar_credit_tracker.xlsx", OSCAR_CARDS),
    "myra":    ("Myra_credit_tracker.xlsx", PARTNER_CARDS),
}

# ------------------------------------------------------------------ helpers ---
def q(sheet):
    """Quote a sheet name for cross-sheet references."""
    return f"'{sheet}'" if any(c in sheet for c in " -") else sheet

CYCLE_END = ('=IF($C{r}="MONTHLY",EOMONTH(TODAY(),0),'
             'IF($C{r}="QUARTERLY",EOMONTH(DATE(\'Ref\'!$B$1,ROUNDUP(MONTH(TODAY())/3,0)*3,1),0),'
             'IF($C{r}="SEMI-ANNUAL",IF(MONTH(TODAY())<=6,DATE(\'Ref\'!$B$1,6,30),DATE(\'Ref\'!$B$1,12,31)),'
             'DATE(\'Ref\'!$B$1,12,31))))')

def cal_row(ws, r, name, basis, cycle, value, note):
    """Block A row formulas (PLAN.md 3.3)."""
    ws[f"A{r}"] = name
    ws[f"B{r}"] = basis
    ws[f"C{r}"] = cycle
    if value is not None:
        ws[f"D{r}"] = value
    ws[f"E{r}"] = (f'=IF($C{r}="MONTHLY",12,IF($C{r}="QUARTERLY",4,'
                   f'IF($C{r}="SEMI-ANNUAL",2,IF($C{r}="CALENDAR-YEAR",1,""))))')
    ws[f"F{r}"] = f'=IF(OR($D{r}="",$E{r}=""),"",$D{r}*$E{r})'
    ws[f"G{r}"] = f'=IF($D{r}="","",IF(ISNUMBER($F{r}),MIN($X{r},$F{r}),$X{r}))'
    ws[f"H{r}"] = f'=IF($F{r}="","",MAX(0,$F{r}-$G{r}))'
    ws[f"I{r}"] = f'=SUMPRODUCT(ISNUMBER($L{r}:$W{r})*($L{r}:$W{r}>0))'
    ws[f"J{r}"] = CYCLE_END.format(r=r)
    ws[f"K{r}"] = f'=IFERROR($J{r}-TODAY(),"")'
    ws[f"X{r}"] = (f'=IF($D{r}="",SUM($L{r}:$W{r}),'
                   f'SUM($L{r}:$W{r})-SUMIF($L{r}:$W{r},">"&$D{r})'
                   f'+(COUNT($L{r}:$W{r})-COUNTIF($L{r}:$W{r},"<="&$D{r}))*$D{r})')
    ws[f"Y{r}"] = (f'=IF(AND(ISNUMBER($H{r}),$H{r}>0,ISNUMBER($K{r}),$K{r}>=0),$K{r},9999)')
    ws[f"Z{r}"] = note
    ws[f"D{r}"].number_format = MONEY
    for c in "FGH":
        ws[f"{c}{r}"].number_format = MONEY

def cmy_row(ws, r, name, basis, cycle, value, note):
    """Block B row (cardmember-year)."""
    cal_row(ws, r, name, basis, cycle, value, note)
    ws[f"E{r}"] = 1
    ws[f"F{r}"] = f'=IF($D{r}="","",$D{r})'
    ws[f"J{r}"] = '=IFERROR(EDATE($B$14,12)-1,"")'

def multi_row(ws, r, name, basis, cycle, value, years, note):
    """Block C row (multi-year / one-time)."""
    ws[f"A{r}"] = name
    ws[f"B{r}"] = basis
    ws[f"C{r}"] = cycle
    if value is not None:
        ws[f"D{r}"] = value
    if years is not None:
        ws[f"E{r}"] = years
    ws[f"F{r}"] = f'=IF(AND(ISNUMBER($D{r}),ISNUMBER($E{r})),$D{r}/$E{r},0)'
    ws[f"J{r}"] = f'=IF($I{r}="y",$D{r},IF(ISNUMBER($I{r}),MIN($I{r},$D{r}),0))'
    ws[f"K{r}"] = note
    ws[f"D{r}"].number_format = MONEY
    ws[f"F{r}"].number_format = MONEY
    ws[f"J{r}"].number_format = MONEY

def add_amount_dv(ws, ranges):
    """Month cells take a dollar amount. Non-numeric input is blocked, because the
    whole point is that a credit is rarely used to the full amount."""
    dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0",
                        allow_blank=True, showErrorMessage=True,
                        errorTitle="Enter an amount",
                        error=("Type the dollar amount you used in this cycle (0 or more), "
                               "or leave the cell blank. Anything above the credit amount "
                               "is automatically capped at the credit amount."),
                        promptTitle="Amount used",
                        prompt=("Type the dollar amount used in this cycle. Blank = not used. "
                                "Values above the credit amount are capped automatically."))
    ws.add_data_validation(dv)
    for rng in ranges:
        dv.add(rng)

def add_list_dv(ws, ref, items, prompt=None):
    dv = DataValidation(type="list", formula1=f'"{",".join(items)}"',
                        allow_blank=True, showErrorMessage=True,
                        errorTitle="Pick from the list",
                        error=f"Allowed: {', '.join(items)}")
    if prompt:
        dv.promptTitle = "Allowed values"
        dv.prompt = prompt
    ws.add_data_validation(dv)
    dv.add(ref)

def add_mark_cf(ws, ranges):
    """Green = amount used, orange = over the per-cycle cap (value is capped),
    red = two uses inside one quarter/half, red fill = non-numeric entry.
    `ranges` is a list of (first_row, last_row) month-grid ranges."""
    for first, last in ranges:
        rng = f"L{first}:W{last}"
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'AND(L{first}<>"",NOT(ISNUMBER(L{first})))'],
            fill=RED_BG, font=RED_FONT, stopIfTrue=True))
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'AND(ISNUMBER(L{first}),$D{first}>0,L{first}>$D{first})'],
            fill=ORANGE, stopIfTrue=True))
        for a, b, _ in (("L", "N", "Q1"), ("O", "Q", "Q2"), ("R", "T", "Q3"), ("U", "W", "Q4")):
            ws.conditional_formatting.add(rng, FormulaRule(
                formula=[f'AND($C{first}="QUARTERLY",COUNTIF(${a}${first}:${b}${first},">0")>1)'],
                font=RED_FONT, stopIfTrue=False))
        for a, b, _ in (("L", "Q", "H1"), ("R", "W", "H2")):
            ws.conditional_formatting.add(rng, FormulaRule(
                formula=[f'AND($C{first}="SEMI-ANNUAL",COUNTIF(${a}${first}:${b}${first},">0")>1)'],
                font=RED_FONT, stopIfTrue=False))
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'AND(ISNUMBER(L{first}),L{first}>0)'], fill=GREEN, stopIfTrue=False))


# ------------------------------------------------------------ card sheet ---
BASIS_STYLE = {
    "SC":     ("1E7B34", "C6EFCE"),   # green  - real statement credits
    "EST":    ("7F6000", "FFE699"),   # amber  - your own estimates
    "ACCESS": ("404040", "D9D9D9"),   # grey   - perks, no dollars
    "MULTI":  ("1F4E79", "DDEBF7"),   # blue   - one-time / every-N-years
}
GRID_HDR = ["Benefit", "Basis", "Cycle", "Value/cycle", "Cyc/yr", "Max annual",
            "Used", "Remaining", "Uses", "Cycle end", "Days left"]
MULTI_HDR = ["Benefit", "Basis", "Cycle", "Value", "Yrs/cycle", "Annualized",
             "Date used", "Expiry / next eligible", "Used? (amount or y)",
             "Value captured", "Notes"]
CYCLES_CAL = ["MONTHLY", "QUARTERLY", "SEMI-ANNUAL", "CALENDAR-YEAR"]
CYCLES_MULTI = ["MULTI-YEAR", "ONE-TIME"]
BASES = ["SC", "EST", "ACCESS"]


def _band(ws, r, text, kind):
    fg, bg = BASIS_STYLE[kind]
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=26)
    c = ws.cell(row=r, column=1, value=text)
    c.font = Font(bold=True, color=fg, size=11)
    c.fill = PatternFill("solid", start_color=bg)
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[r].height = 20


def _title(basis, ann, label):
    who = {"SC": "STATEMENT CREDITS (SC)", "EST": "ESTIMATED SAVINGS (EST)",
           "ACCESS": "ACCESS PERKS (no dollar value)"}[basis]
    tail = {"SC": "NET BENEFIT COMES FROM HERE.",
            "EST": "Counts only in the 'incl. estimated' totals.",
            "ACCESS": "A number typed here logs a USE only and never adds dollars."}[basis]
    if ann:
        return ("SECTION " + label + " \u00b7 " + who +
                " THAT RESET ON YOUR CARD ANNIVERSARY \u00b7 the month columns follow YOUR anniversary month, "
                "not January. " + tail)
    if basis == "SC":
        return ("SECTION " + label + " \u00b7 " + who +
                " \u00b7 real money on your statement. " + tail + " Type the amount you used into the month column.")
    return "SECTION " + label + " \u00b7 " + who + " \u00b7 " + tail


def _mtitle(basis, label):
    return ("SECTION " + label + " \u00b7 ONE-TIME / EVERY-N-YEARS ITEMS (" + basis +
            ") \u00b7 no month grid - enter the amount in the 'Used?' column.")


SECTION_LABEL = {"SC": "1", "EST": "2", "ACCESS": "3"}
MULTI_LABEL = {"SC": "4", "EST": "4b", "ACCESS": "4c"}


def _ann_label(k):
    inner = "INDEX('Ref'!$J$3:$J$14,MONTH(EDATE($B$14," + str(k) + ")))"
    yy = 'TEXT(EDATE($B$14,' + str(k) + '),"yy")'
    return '=IFERROR(' + inner + '&"-"&' + yy + ',"")'


def _head(ws, r, cols, ann=False):
    for i, h in enumerate(cols):
        c = ws.cell(row=r, column=1 + i, value=h)
        c.font, c.fill, c.border = BOLD, HDR_FILL, BOX
    if cols == GRID_HDR:
        for k in range(12):
            val = _ann_label(k) if ann else "=INDEX('Ref'!$J$3:$J$14,COLUMN()-11)"
            c = ws.cell(row=r, column=12 + k, value=val)
            c.font, c.fill, c.border = BOLD, HDR_FILL, BOX
        z = ws.cell(row=r, column=26, value="Notes / source")
        z.font, z.fill = BOLD, HDR_FILL


def _subtotal(ws, r, first, last, cols, label="Subtotal"):
    c = ws.cell(row=r, column=1, value=label)
    c.font, c.fill = BOLD, TOT_FILL
    for col in cols:
        cc = ws.cell(row=r, column=ord(col) - 64,
                     value="=SUM($" + col + "$" + str(first) + ":$" + col + "$" + str(last) + ")")
        cc.font, cc.fill, cc.number_format = BOLD, TOT_FILL, MONEY


def grid_row(ws, r, entry, ann=False):
    name, basis, cycle, value, note = entry
    ws["A" + str(r)] = name
    ws["B" + str(r)] = basis
    ws["C" + str(r)] = cycle
    if value is not None:
        ws["D" + str(r)] = value
    if ann:
        ws["E" + str(r)] = 1
        ws["F" + str(r)] = '=IF($D{r}="","",$D{r})'.format(r=r)
        ws["J" + str(r)] = '=IFERROR(EDATE($B$14,12)-1,"")'
    else:
        ws["E" + str(r)] = ('=IF($C{r}="MONTHLY",12,IF($C{r}="QUARTERLY",4,'
                            'IF($C{r}="SEMI-ANNUAL",2,IF($C{r}="CALENDAR-YEAR",1,""))))').format(r=r)
        ws["F" + str(r)] = '=IF(OR($D{r}="",$E{r}=""),"",$D{r}*$E{r})'.format(r=r)
        ws["J" + str(r)] = CYCLE_END.format(r=r)
    ws["G" + str(r)] = '=IF($D{r}="","",IF(ISNUMBER($F{r}),MIN($X{r},$F{r}),$X{r}))'.format(r=r)
    ws["H" + str(r)] = '=IF($F{r}="","",MAX(0,$F{r}-$G{r}))'.format(r=r)
    ws["I" + str(r)] = '=SUMPRODUCT(ISNUMBER($L{r}:$W{r})*($L{r}:$W{r}>0))'.format(r=r)
    ws["K" + str(r)] = '=IFERROR($J{r}-TODAY(),"")'.format(r=r)
    ws["X" + str(r)] = ('=IF($D{r}="",SUM($L{r}:$W{r}),'
                        'SUM($L{r}:$W{r})-SUMIF($L{r}:$W{r},">"&$D{r})'
                        '+(COUNT($L{r}:$W{r})-COUNTIF($L{r}:$W{r},"<="&$D{r}))*$D{r})').format(r=r)
    ws["Y" + str(r)] = ('=IF(AND(ISNUMBER($H{r}),$H{r}>0,ISNUMBER($K{r}),$K{r}>=0),'
                        '$K{r},9999)').format(r=r)
    ws["Z" + str(r)] = note
    for col in ("D", "F", "G", "H"):
        ws[col + str(r)].number_format = MONEY


def multi_row(ws, r, entry):
    name, basis, cycle, value, years, note = entry
    ws["A" + str(r)] = name
    ws["B" + str(r)] = basis
    ws["C" + str(r)] = cycle
    if value is not None:
        ws["D" + str(r)] = value
    if years is not None:
        ws["E" + str(r)] = years
    ws["F" + str(r)] = '=IF(AND(ISNUMBER($D{r}),ISNUMBER($E{r})),$D{r}/$E{r},0)'.format(r=r)
    ws["J" + str(r)] = '=IF($I{r}="y",$D{r},IF(ISNUMBER($I{r}),MIN($I{r},$D{r}),0))'.format(r=r)
    ws["K" + str(r)] = note
    for col in ("D", "F", "J"):
        ws[col + str(r)].number_format = MONEY


TEMPLATE_ROWS = dict(
    cal=[("(example) Monthly credit", "SC", "MONTHLY", 10, "Replace or delete this row"),
         ("(example) Annual statement credit", "SC", "CALENDAR-YEAR", 100, "Replace or delete this row"),
         ("(example) Estimated perk value", "EST", "CALENDAR-YEAR", None, "EST rows count only in the incl.-estimated totals"),
         ("(example) Free checked bag", "ACCESS", "MONTHLY", None, "ACCESS: a number here logs a use, never dollars")],
    cmy=[("(example) Credit that resets on your card anniversary", "SC", "CARDMEMBER-YEAR", 100,
          "Lives in the anniversary section; the month columns follow B14")],
    multi=[("(example) Global Entry / TSA credit", "SC", "MULTI-YEAR", 120, 4, "Every 4 years"),
           ("(example) One-off promo credit", "EST", "ONE-TIME", 50, 1, "One-time")])


def build_card_sheet(wb, tab, d=None, template=False):
    ws = wb.create_sheet(tab)
    ws.sheet_properties.tabColor = "FFD966" if template else "8EAADB"
    data = d or TEMPLATE_ROWS

    for r, lbl in enumerate(["Card", "Issuer", "Product (exact)", "Card status",
                             "Open date (approval)", "Months held", "Years held",
                             "Annual fee (USD)", "First-year AF waived?",
                             "AF posting date (current cycle)",
                             "Authorized user AF (total, USD)",
                             "Total AF charged (this cycle)", "Next AF date (projected)",
                             "Cardmember year start (current cycle)"], start=1):
        ws["A" + str(r)] = lbl
        ws["A" + str(r)].font = BOLD

    if d:
        ws["B1"], ws["B2"], ws["B3"] = d["tab"], d["issuer"], d["product"]
        ws["B4"], ws["B5"] = d["status"], date.fromisoformat(d["opened"])
        ws["B8"], ws["B11"] = d["af"], d["au_af"]
        ws["B9"] = "y" if d["waived"] else None
        ws["B10"] = date.fromisoformat(d["opened"])
        ws["Z2"] = d["af_note"]
    ws["B6"] = '=IF($B$5="","",DATEDIF($B$5,TODAY(),"m"))'
    ws["B7"] = '=IF($B$6="","",ROUND($B$6/12,1))'
    ws["B12"] = '=IF($B$9="y",0,$B$8)+$B$11'
    ws["B13"] = '=EDATE($B$10,12)'
    ws["B14"] = '=IFERROR(EDATE($B$10,12*DATEDIF($B$10,TODAY(),"y")),"")'
    for c in ("B5", "B10", "B13", "B14"):
        ws[c].number_format = DATEFMT
    for c in ("B8", "B11", "B12"):
        ws[c].number_format = MONEY

    kpi = [("Credits used YTD - statement credits (SC)", MONEY),
           ("Credits used YTD - incl. estimated (SC+EST)", MONEY),
           ("Max annual value (statement credits + estimated)", MONEY),
           ("Remaining annual credit value (incl. est.)", MONEY),
           ("Remaining - statement credits only", MONEY),
           ("Net benefit - statement credits", MONEY),
           ("Net benefit - incl. estimated", MONEY),
           ("Result", None),
           ("AF breakeven gap", MONEY),
           ("Credit utilization %", PCT),
           ("Credits at risk (unused, <=30 days left)", None),
           ("Nearest expiry (days)", None),
           ("Amounts you entered (cells)", None),
           ("Entered total before capping", MONEY)]
    for i, (label, fmt) in enumerate(kpi):
        ws["D" + str(3 + i)] = label
        ws["D" + str(3 + i)].font = BOLD
    ws["G3"] = '=HYPERLINK("#\'Dashboard\'!A1","< back to Dashboard")'

    # ---- sections, grouped by Basis ----
    sections = []
    r = 19
    plan = []
    for basis in BASES:
        rows = [e for e in data["cal"] if e[1] == basis]
        if rows:
            plan.append((basis, rows, "cal", SECTION_LABEL[basis]))
        rows = [e for e in data["cmy"] if e[1] == basis]
        if rows:
            plan.append((basis, rows, "ann", SECTION_LABEL[basis] + "b"))
    for basis in BASES:
        rows = [e for e in data["multi"] if e[1] == basis]
        if rows:
            plan.append((basis, rows, "multi", MULTI_LABEL[basis]))

    for basis, rows, kind, label in plan:
        if kind == "multi":
            _band(ws, r, _mtitle(basis, label), "MULTI")
            r += 1
            _head(ws, r, MULTI_HDR)
            r += 1
            first = r
            for e in rows:
                multi_row(ws, r, e)
                r += 1
            last = r - 1
            _subtotal(ws, r, first, last, ("D", "F", "J"), "Subtotal - Section " + label)
            r += 2
        else:
            _band(ws, r, _title(basis, kind == "ann", label), basis)
            r += 1
            _head(ws, r, GRID_HDR, ann=(kind == "ann"))
            r += 1
            first = r
            for e in rows:
                grid_row(ws, r, e, ann=(kind == "ann"))
                r += 1
            last = r - 1
            _subtotal(ws, r, first, last, ("F", "G", "H", "I"), "Subtotal - Section " + label)
            r += 2
        sections.append((kind, first, last))

    grid = [(a, b) for k, a, b in sections if k in ("cal", "ann")]
    mult = [(a, b) for k, a, b in sections if k == "multi"]

    def join(terms, fallback="0"):
        return "+".join(terms) if terms else fallback

    def sumif_g(ranges, basis):
        return ['SUMIF($B${0}:$B${1},"{2}",$G${0}:$G${1})'.format(a, b, basis) for a, b in ranges]

    def sumifs_m(ranges, basis):
        return ['SUMIFS($F${0}:$F${1},$B${0}:$B${1},"{2}",$I${0}:$I${1},"y")'.format(a, b, basis)
                for a, b in ranges]

    used_sc = sumif_g(grid, "SC") + sumifs_m(mult, "SC")
    used_est = sumif_g(grid, "EST") + sumifs_m(mult, "EST")
    max_all = (["SUM($F${0}:$F${1})".format(a, b) for a, b in grid] +
               ["SUM($F${0}:$F${1})".format(a, b) for a, b in mult])
    y_ranges = ["$Y${0}:$Y${1}".format(a, b) for a, b in grid]
    atrisk = ["SUMPRODUCT(($Y${0}:$Y${1}>=0)*($Y${0}:$Y${1}<=30))".format(a, b) for a, b in grid]
    months = ["L{0}:W{1}".format(a, b) for a, b in grid]
    mj = ["J{0}:J{1}".format(a, b) for a, b in mult]

    ws["E3"] = "=" + join(used_sc)
    ws["E4"] = "=" + join(used_sc + used_est)
    ws["E5"] = "=" + join(max_all)
    ws["E6"] = '=MAX(0,$E$5-$E$4)'
    sc_max = (['SUMIF($B${0}:$B${1},"SC",$F${0}:$F${1})'.format(a, b) for a, b in grid] +
              ['SUMIFS($F${0}:$F${1},$B${0}:$B${1},"SC")'.format(a, b) for a, b in mult])
    ws["H4"] = "=" + join(sc_max)
    ws["E7"] = '=MAX(0,$H$4-$E$3)'
    ws["E8"] = '=$E$3-$B$12'
    ws["E9"] = '=$E$4-$B$12'
    ws["E10"] = '=IF($E$8>=0,"POSITIVE","NEGATIVE")'
    ws["E11"] = '=MAX(0,-$E$8)'
    ws["E12"] = '=IF($E$5=0,"",$E$4/$E$5)'
    ws["E13"] = "=" + join(atrisk)
    ws["H3"] = "=MIN(" + ",".join(y_ranges) + ")" if y_ranges else "=9999"
    ws["E14"] = '=IF($H$3>=9999,"",$H$3)'
    ws["E15"] = "=" + join(["COUNT(" + x + ")" for x in months] +
                           ['COUNTIF($I${0}:$I${1},">0")'.format(a, b) for a, b in mult])
    ws["E16"] = "=" + join(["SUM(" + x + ")" for x in months] +
                           ["SUM($J${0}:$J${1})".format(a, b) for a, b in mult])
    for c in ("E3", "E4", "E5", "E6", "E7", "E8", "E9", "E11", "E16"):
        ws[c].number_format = MONEY
    ws["E12"].number_format = PCT
    ws["E8"].font = BOLD

    ws["A17"] = ("WHERE TO TYPE: put the amount you used into the month column of that row. "
                 "SC rows are real statement credits and drive your net benefit. EST rows are values you set "
                 "yourself (baggage fees saved) and only count in the 'incl. estimated' totals. "
                 "'Amounts you entered' above counts how many month cells you have filled in.")
    ws["A17"].font = Font(bold=True, size=9, color="1F3864")
    ws.merge_cells("A17:K17")

    # ---- validation ----
    add_amount_dv(ws, months)
    basis_ranges = ["B{0}:B{1}".format(a, b) for k, a, b in sections]
    for rng in basis_ranges:
        add_list_dv(ws, rng, BASES)
    for k, a, b in sections:
        if k == "cal":
            add_list_dv(ws, "C{0}:C{1}".format(a, b), CYCLES_CAL)
        elif k == "ann":
            add_list_dv(ws, "C{0}:C{1}".format(a, b), ["CARDMEMBER-YEAR"])
        else:
            add_list_dv(ws, "C{0}:C{1}".format(a, b), CYCLES_MULTI)
            add_list_dv(ws, "I{0}:I{1}".format(a, b), ["y"])
    add_list_dv(ws, "B9", ["y"])
    add_list_dv(ws, "B4", ["Active", "Closed", "Product-changed", "Authorized-user-only"])
    add_list_dv(ws, "B2", ["Amex", "Chase", "Citi", "Capital One", "Bank of America",
                           "Discover", "Wells Fargo", "US Bank", "Barclays", "Robinhood", "Other"])

    # ---- conditional formatting ----
    add_mark_cf(ws, grid)
    for a, b in grid:
        ws.conditional_formatting.add("A{0}:I{1}".format(a, b), FormulaRule(
            formula=['AND($A' + str(a) + '<>"",$C' + str(a) + '="")'], font=GREY_FONT, stopIfTrue=False))
    ws.conditional_formatting.add("E10", FormulaRule(
        formula=['$E$10="POSITIVE"'], fill=GREEN, font=GREEN_FONT, stopIfTrue=False))
    ws.conditional_formatting.add("E10", FormulaRule(
        formula=['$E$10="NEGATIVE"'], fill=RED_BG, font=RED_FONT, stopIfTrue=False))

    # ---- layout ----
    for col, w in {"A": 36, "B": 10, "C": 15, "D": 13, "E": 8, "F": 12, "G": 12,
                   "H": 12, "I": 10, "J": 11, "K": 9, "Z": 62}.items():
        ws.column_dimensions[col].width = w
    for col in range(12, 24):
        ws.column_dimensions[get_column_letter(col)].width = 7
    for col in ("J", "K", "X", "Y"):
        ws.column_dimensions[col].hidden = True
    ws.freeze_panes = "L19"
    return [(tab, a, b, k) for k, a, b in sections]


# --------------------------------------------------------------- Dashboard ---
def build_dashboard(wb, cards):
    d = wb.create_sheet("Dashboard", 0)
    d.sheet_properties.tabColor = "1F3864"
    d["A1"] = "Credit Card Benefits & Annual Fee Tracker"
    d["A1"].font = TITLE
    d["A2"] = "Portfolio rollup as of"
    d["B2"] = "=TODAY()"
    d["B2"].number_format = DATEFMT

    kpi = {
        "B3": ("Total annual fees charged (active cards)", '=SUMIF($G$9:$G$28,"Active",$H$9:$H$28)', MONEY),
        "D3": ("Credits used - statement credits", '=SUMIF($G$9:$G$28,"Active",$J$9:$J$28)', MONEY),
        "F3": ("Credits used - incl. estimated", '=SUMIF($G$9:$G$28,"Active",$K$9:$K$28)', MONEY),
        "H3": ("Remaining credit value", '=SUMIF($G$9:$G$28,"Active",$L$9:$L$28)', MONEY),
        "B4": ("Portfolio net benefit - statement credits", '=SUMIF($G$9:$G$28,"Active",$M$9:$M$28)', MONEY),
        "D4": ("Portfolio net benefit - incl. estimated", '=SUMIF($G$9:$G$28,"Active",$N$9:$N$28)', MONEY),
        "F4": ("Portfolio result", '=IF($B$4>=0,"POSITIVE","NEGATIVE")', None),
        "H4": ("Credit utilization %", '=IF($I$29=0,"",$F$3/$I$29)', PCT),
        "B5": ("Cards tracked / active", '=COUNTA($A$9:$A$28)&" / "&COUNTIF($G$9:$G$28,"Active")', None),
        "D5": ("Total months held (active)", '=SUMIF($G$9:$G$28,"Active",$E$9:$E$28)', None),
        "F5": ("Average months held (active)", '=IFERROR(AVERAGEIF($G$9:$G$28,"Active",$E$9:$E$28),"")', None),
        "H5": ("Credits at risk (<=30 days)", '=SUMIF($G$9:$G$28,"Active",$Q$9:$Q$28)', None),
        "B6": ("Nearest expiry across active cards (days)", '=IF(MIN($S$9:$S$28)>=9999,"",MIN($S$9:$S$28))', None),
    }
    for cell, (label, formula, fmt) in kpi.items():
        row = int(cell[1:])
        col = cell[0]
        d[cell] = formula
        lab_col = {"B": "A", "D": "C", "F": "E", "H": "G"}[col]
        d[f"{lab_col}{row}"] = label
        d[f"{lab_col}{row}"].font = BOLD
        if fmt:
            d[f"{col}{row}"].number_format = fmt
    d["B4"].font = BOLD
    d["F4"].font = BOLD
    d.conditional_formatting.add("F4", FormulaRule(formula=['$F$4="POSITIVE"'], fill=GREEN, font=GREEN_FONT))
    d.conditional_formatting.add("F4", FormulaRule(formula=['$F$4="NEGATIVE"'], fill=RED_BG, font=RED_FONT))
    d.conditional_formatting.add("H5", FormulaRule(formula=['$H$5>0'], fill=AMBER, stopIfTrue=False))

    headers = ["Card (tab)", "Issuer", "Product", "Open date", "Months held", "Years held",
               "Card status", "AF this cycle", "Max annual (SC+est.)", "Used - SC",
               "Used - incl. est.", "Remaining", "Net - SC", "Net - incl. est.", "Result",
               "Util %", "At risk", "Nearest expiry (d)"]
    for i, h in enumerate(headers):
        c = d.cell(row=8, column=1 + i, value=h)
        c.font, c.fill, c.border = BOLD, HDR_FILL, BOX

    refmap = {"B": "$B$2", "C": "$B$3", "D": "$B$5", "E": "$B$6", "F": "$B$7",
              "G": "$B$4", "H": "$B$12", "I": "$E$5", "J": "$E$3", "K": "$E$4",
              "L": "$E$6", "M": "$E$8", "N": "$E$9", "O": "$E$10", "P": "$E$12",
              "Q": "$E$13", "R": "$E$14"}
    for i, card in enumerate(cards):
        r = 9 + i
        tab = card["tab"]
        d[f"A{r}"] = f'=HYPERLINK("#\'{tab}\'!A1","{tab}")'
        for col, ref in refmap.items():
            d[f"{col}{r}"] = f"={q(tab)}!{ref}"
        for col in "HJ K L M N".replace(" ", ""):
            d[f"{col}{r}"].number_format = MONEY
        d[f"P{r}"].number_format = PCT
        d[f"D{r}"].number_format = DATEFMT
    # hidden helper column S: days-left if active
    for r in range(9, 29):
        d[f"S{r}"] = f'=IF(AND($G{r}="Active",ISNUMBER($R{r})),$R{r},9999)'
        d[f"S{r}"].font = GREY_FONT

    for col, r_ in {"H": "H", "I": "I", "J": "J", "K": "K", "L": "L",
                    "M": "M", "N": "N", "Q": "Q"}.items():
        c = d[f"{col}29"]
        c.value = f'=SUMIF($G$9:$G$28,"Active",${r_}$9:${r_}$28)'
        c.font, c.fill = BOLD, TOT_FILL
        if col != "Q":
            c.number_format = MONEY
    d["E29"] = '=IFERROR(AVERAGEIF($G$9:$G$28,"Active",$E$9:$E$28),"")'
    d["E29"].font, d["E29"].fill = BOLD, TOT_FILL
    d["O29"] = '=IF($M$29>=0,"POSITIVE","NEGATIVE")'
    d["P29"] = '=IF($I$29=0,"",$K$29/$I$29)'
    d["P29"].number_format = PCT
    for c in ("O29", "P29"):
        d[c].font, d[c].fill = BOLD, TOT_FILL
    d.conditional_formatting.add("O9:O28", FormulaRule(
        formula=['$O9="POSITIVE"'], fill=GREEN, font=GREEN_FONT))
    d.conditional_formatting.add("O9:O28", FormulaRule(
        formula=['$O9="NEGATIVE"'], fill=RED_BG, font=RED_FONT))
    d.conditional_formatting.add("A9:R28", FormulaRule(
        formula=['$G9="Closed"'], font=GREY_FONT))
    d.conditional_formatting.add("Q9:Q28", FormulaRule(
        formula=['AND($G9="Active",$Q9>0)'], fill=AMBER))

    d["A31"] = "How to use"
    d["A31"].font = BOLD
    notes = [
        "Enter the AMOUNT you used in the month cell (e.g. 150 of a 200 credit). Green = counted. Blank = not used.",
        "If you enter more than the credit is worth, the workbook counts only the credit amount - a single cell is capped at 'Value/cycle', and the row total is capped at 'Max annual'.",
        "Orange cell = you entered more than that cycle's credit (capped automatically). Red banner = two uses inside the same quarter/half.",
        "Net benefit (statement credits) = used statement credits - annual fees charged. The 'incl. estimated' columns add perk values you set yourself.",
        "Cardmember-year credits reset on the card's anniversary, not Jan 1 - their month columns run from the AF posting month in column A10 of the card sheet.",
        "To add a card: copy the TEMPLATE tab, rename it, then copy a row in this table and Find/Replace the tab name.",
        "Card tabs group rows into sections by BASIS: SC (real statement credits - these move your net benefit), EST (your estimates), ACCESS (perks, no dollars).",
        "TYPE YOUR USAGE ON THE CARD TAB, in the month column. The 'Benefits' tab is read-only - typing there breaks the link.",
        "If a number you typed did not change the totals: check the row's Basis (ACCESS rows never add dollars) and check you typed in the MONTH column, not in 'Value/cycle'.",
        "Amber on 'At risk' means an unused credit expires within 30 days.",
    ]
    for i, n in enumerate(notes):
        d[f"A{32+i}"] = "- " + n

    for col, w in {"A": 20, "B": 22, "C": 34, "D": 12, "E": 11, "F": 10, "G": 15,
                   "H": 13, "I": 14, "J": 13, "K": 14, "L": 12, "M": 12, "N": 14,
                   "O": 12, "P": 9, "Q": 9, "R": 14}.items():
        d.column_dimensions[col].width = w
    d.column_dimensions["S"].hidden = True
    d.freeze_panes = "B9"
    return d

# ------------------------------------------------------- Benefits index ---
def build_benefits_index(wb, index_rows):
    """One flat, live list of every benefit on every card, so nothing hides inside a
    section. Every cell is a formula pointing back at the card tabs."""
    ws = wb.create_sheet("Benefits", 1)
    ws.sheet_properties.tabColor = "2E75B6"
    ws["A1"] = "Every benefit, every card - one live list, grouped by Basis"
    ws["A1"].font = TITLE
    ws["A2"] = ("All values, usage and reset dates are formulas pointing at the card tabs. "
                "Enter your usage on the CARD TAB - do not type in this sheet, it will break the link.")
    ws["A2"].font = Font(bold=True, color="C00000")
    headers = ["Card", "Benefit", "Basis", "Cycle", "Resets on", "Value/cycle", "Max annual",
               "Used", "Remaining", "Status", "Note / source", "Open"]
    for i, h in enumerate(headers):
        c = ws.cell(row=4, column=1 + i, value=h)
        c.font, c.fill, c.border = BOLD, HDR_FILL, BOX

    ir = 5
    for tab, a, b, kind in index_rows:
        for cr in range(a, b + 1):
            t = q(tab)
            ws["A" + str(ir)] = "=" + t + "!$B$1"
            ws["B" + str(ir)] = "=" + t + "!$A" + str(cr)
            ws["C" + str(ir)] = "=" + t + "!$B" + str(cr)
            ws["D" + str(ir)] = "=" + t + "!$C" + str(cr)
            if kind == "multi":
                ws["E" + str(ir)] = '=IF(' + t + '!$H' + str(cr) + '="","",' + t + '!$H' + str(cr) + ')'
                ws["F" + str(ir)] = "=" + t + "!$D" + str(cr)
                ws["G" + str(ir)] = "=" + t + "!$E" + str(cr)
                ws["H" + str(ir)] = "=" + t + "!$J" + str(cr)
                ws["I" + str(ir)] = ('=IF(ISNUMBER(' + t + '!$D' + str(cr) + '),MAX(0,' + t +
                                      '!$D' + str(cr) + '-' + t + '!$J' + str(cr) + '),"")')
            else:
                ws["E" + str(ir)] = '=IFERROR(' + t + '!$J' + str(cr) + ',"")'
                ws["F" + str(ir)] = "=" + t + "!$D" + str(cr)
                ws["G" + str(ir)] = '=IF(' + t + '!$F' + str(cr) + '="","",' + t + '!$F' + str(cr) + ')'
                ws["H" + str(ir)] = '=IF(' + t + '!$G' + str(cr) + '="","",' + t + '!$G' + str(cr) + ')'
                ws["I" + str(ir)] = '=IF(' + t + '!$H' + str(cr) + '="","",' + t + '!$H' + str(cr) + ')'
            ws["J" + str(ir)] = ('=IF(ISNUMBER(SEARCH("PENDING",' + t + '!$Z' + str(cr) + '&"")),'
                                 '"PENDING","")')
            ws["K" + str(ir)] = '=IF(' + t + '!$Z' + str(cr) + '="","",' + t + '!$Z' + str(cr) + ')'
            ws["L" + str(ir)] = '=HYPERLINK("#' + "'" + tab + "'!A" + str(cr) + '","open")'
            for col in ("F", "G", "H", "I"):
                ws[col + str(ir)].number_format = MONEY
            ws["E" + str(ir)].number_format = DATEFMT
            ir += 1
    last = ir - 1
    ws.auto_filter.ref = "A4:L" + str(last)
    ws.conditional_formatting.add("A5:E" + str(last), FormulaRule(
        formula=['$C5="ACCESS"'], font=GREY_FONT, stopIfTrue=False))
    ws.conditional_formatting.add("H5:I" + str(last), FormulaRule(
        formula=['AND(ISNUMBER($H5),$H5>0)'], fill=GREEN, stopIfTrue=False))
    ws.conditional_formatting.add("J5:J" + str(last), FormulaRule(
        formula=['$J5="PENDING"'], fill=AMBER, font=RED_FONT, stopIfTrue=False))
    for col, w in {"A": 19, "B": 46, "C": 10, "D": 14, "E": 11, "F": 12, "G": 12,
                   "H": 11, "I": 11, "J": 10, "K": 70, "L": 7}.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A5"
    return ws


# -------------------------------------------------------------------- Ref ---
def build_ref(wb):
    ws = wb.create_sheet("Ref")
    ws.sheet_properties.tabColor = "BFBFBF"
    ws["A1"], ws["B1"] = "Tracking year", YEAR
    ws["B1"].font = Font(bold=True, size=12, color="C00000")
    ws["A2"] = "Change B1 to roll the whole workbook to the next year"
    ws["A2"].font = GREY_FONT

    def put(title, items, start):
        ws[f"A{start-1}"] = title
        ws[f"A{start-1}"].font = BOLD
        for i, it in enumerate(items):
            ws[f"A{start+i}"] = it

    put("Issuers", ["Amex", "Chase", "Citi", "Capital One", "Bank of America", "Discover",
                    "Wells Fargo", "US Bank", "Barclays", "Robinhood", "Other"], 3)
    put("Basis", ["SC", "EST", "ACCESS"], 16)
    put("Cycles", ["MONTHLY", "QUARTERLY", "SEMI-ANNUAL", "CALENDAR-YEAR",
                   "CARDMEMBER-YEAR", "MULTI-YEAR", "ONE-TIME"], 22)
    put("Card status", ["Active", "Closed", "Product-changed", "Authorized-user-only"], 31)

    ws["J1"] = "Short month names (English)"
    ws["J1"].font = BOLD
    for i, m in enumerate(MONTHS):
        ws[f"J{3+i}"] = m

    ws["A36"] = "Benefit catalog - typical reset cycles. Verify every amount on the issuer's own terms before relying on it."
    ws["A36"].font = BOLD
    ws.append([])
    cat = [("Benefit type", "Typical cycle", "Typical value", "Verification status"),
           ("Hilton Resort Credit (Aspire)", "SEMI-ANNUAL ($200 Jan-Jun / $200 Jul-Dec)", "$400/yr", "Third-party - confirm"),
           ("Airline Fee Credit (Aspire)", "QUARTERLY ($50 per quarter)", "$200/yr", "PENDING - incidental fees vs airfare"),
           ("CLEAR+ Credit (Aspire)", "CALENDAR-YEAR", "$219/yr", "Third-party - confirm"),
           ("Chase Travel hotel credit (CSP)", "CARDMEMBER-YEAR", "$100/yr", "OFFICIAL - Chase.com"),
           ("DoorDash non-restaurant credit (CSP)", "MONTHLY", "$10/mo through 12/31/2027", "Third-party - confirm"),
           ("Instacart credit (CSP)", "MONTHLY", "unknown", "PENDING"),
           ("Global Entry/TSA/NEXUS (CSP)", "MULTI-YEAR (4 yr)", "up to $120", "Third-party - confirm"),
           ("CLEAR+ Credit (Amex Green)", "CALENDAR-YEAR", "$219", "Third-party - confirm"),
           ("Turo credit (Citi AA Plat)", "ONE-TIME, expires 2026-10-18", "up to $30/trip, max $180", "Issuer/Turo page"),
           ("Disney Streaming Credit (BCE)", "MONTHLY", "$7/mo up to $84/yr", "OFFICIAL - Amex.com"),
           ("Home Chef credit (BCE)", "MONTHLY", "up to $15/mo", "PENDING"),
           ("Free checked bag (Citi AA / Atmos)", "PER USE - you set the value", "your call, e.g. $35/bag", "You set it"),
           ("Robinhood Gold membership", "YEARLY (monthly billing)", "$5/mo = $60/yr", "Counted in the annual fee"),
           ]
    start = ws.max_row + 1
    for i, row in enumerate(cat):
        for j, v in enumerate(row):
            c = ws.cell(row=start + i, column=1 + j, value=v)
            if i == 0:
                c.font, c.fill = BOLD, HDR_FILL
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 44
    ws.column_dimensions["C"].width = 30
    ws.column_dimensions["D"].width = 30
    return ws

# --------------------------------------------- entry preservation --------
STATE_FILE = ".tracker_build_state_oscar.json"


def load_state():
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, encoding="utf-8") as fh:
                return json.load(fh)
        except Exception:
            return {}
    return {}


def build_state(cards):
    """Remember the defaults this build wrote, so a later harvest can tell a
    user-edited value apart from an untouched default."""
    st = {}
    for card in cards:
        per = {}
        for key in ("cal", "cmy", "multi"):
            for e in card[key]:
                per[e[0]] = {"d": e[3], "cycle": e[2]}
        st[card["tab"]] = per
    return st


def save_state(state):
    with open(STATE_FILE, "w", encoding="utf-8") as fh:
        json.dump(state, fh, indent=1, ensure_ascii=False)


def harvest_entries(path, prev_state):
    """Read what the user typed in an existing workbook, keyed by (sheet, benefit
    name, month position) so it survives the layout being rebuilt."""
    if not os.path.exists(path):
        return None
    old = load_workbook(path)
    cards, ident = {}, {}
    for sh in old.sheetnames:
        ws = old[sh]
        if sh == "Ref":
            if ws["B1"].value is not None:
                ident["Ref!B1"] = ws["B1"].value
            continue
        if sh in ("Dashboard", "Benefits", "TEMPLATE"):
            continue
        per_prev = prev_state.get(sh, {})
        card = {}
        for r in range(18, ws.max_row + 1):
            name = ws.cell(row=r, column=1).value
            basis = ws.cell(row=r, column=2).value
            cycle = ws.cell(row=r, column=3).value
            if not isinstance(name, str) or basis not in ("SC", "EST", "ACCESS"):
                continue
            months = {}
            for c in range(12, 24):
                v = ws.cell(row=r, column=c).value
                if isinstance(v, (int, float)):
                    months[c - 11] = v
            d = ws.cell(row=r, column=4).value
            prev_d = (per_prev.get(name) or {}).get("d")
            d_user = d if (isinstance(d, (int, float)) and d != prev_d) else None
            used = None
            if cycle in ("MULTI-YEAR", "ONE-TIME"):      # only here is I a user input
                v = ws.cell(row=r, column=9).value
                if isinstance(v, (str, int, float)) and not (isinstance(v, str) and v.startswith("=")):
                    used = v
            card[name] = {"months": months, "d": d_user, "used": used}
        cards[sh] = card
        for cell in ("B1", "B2", "B3", "B4", "B5", "B8", "B9", "B10", "B11"):
            v = ws[cell].value
            if v is not None:
                ident[sh + "!" + cell] = v
    return {"cards": cards, "identity": ident}


def apply_entries(wb, harv):
    """Write the harvested values back into the freshly built workbook."""
    if not harv:
        return 0, []
    restored, missing = 0, []
    for sh, card in harv["cards"].items():
        if sh not in wb.sheetnames:
            missing.append(sh)
            continue
        ws = wb[sh]
        byname = {}
        for r in range(18, ws.max_row + 1):
            name = ws.cell(row=r, column=1).value
            if isinstance(name, str) and ws.cell(row=r, column=2).value in ("SC", "EST", "ACCESS"):
                byname.setdefault(name, []).append(r)
        for name, data in card.items():
            rows = byname.get(name)
            if not rows:
                if data["months"] or data["used"] or data["d"] is not None:
                    missing.append(sh + " / " + name)
                continue
            r = rows.pop(0)
            for idx, v in data["months"].items():
                ws.cell(row=r, column=11 + idx).value = v
                restored += 1
            if data["d"] is not None:
                ws.cell(row=r, column=4).value = data["d"]
                restored += 1
            if data["used"] is not None:
                ws.cell(row=r, column=9).value = data["used"]
                restored += 1
    for key, v in harv["identity"].items():
        sh, cell = key.split("!")
        if sh in wb.sheetnames:
            wb[sh][cell] = v
    return restored, missing


# ------------------------------------------------------------------- main ---
def main():
    global OUT, STATE_FILE
    profile = sys.argv[1] if len(sys.argv) > 1 else "oscar"
    if profile not in PROFILES:
        print("unknown profile. use: " + " | ".join(PROFILES))
        return
    OUT, cards = PROFILES[profile]
    STATE_FILE = ".tracker_build_state_" + profile + ".json"
    print("profile: " + profile + "  ->  " + OUT)
    prev_state = load_state()
    harv = harvest_entries(OUT, prev_state)
    if harv and os.path.exists(OUT):
        shutil.copy(OUT, OUT.replace(".xlsx", ".backup.xlsx"))
    wb = Workbook()
    wb.remove(wb.active)
    wb.calculation.fullCalcOnLoad = True

    build_dashboard(wb, cards)
    index_rows = []
    for card in cards:
        index_rows += build_card_sheet(wb, card["tab"], card)
    build_card_sheet(wb, "TEMPLATE", None, template=True)
    build_benefits_index(wb, index_rows)
    build_ref(wb)

    restored, missing = apply_entries(wb, harv)
    wb.save(OUT)
    save_state(build_state(cards))
    print(f"wrote {OUT}")
    print(f"sheets: {wb.sheetnames}")
    if harv:
        print(f"preserved your entries: {restored} cell value(s) written back")
        if missing:
            print("  NOT restored (row was renamed or removed):")
            for m in missing:
                print("   -", m)

if __name__ == "__main__":
    main()
