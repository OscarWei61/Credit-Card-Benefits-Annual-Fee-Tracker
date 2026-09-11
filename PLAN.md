# PLAN — Credit Card Benefits & Annual Fee Tracker (.xlsx)

Status: **DESIGN SPEC — no .xlsx has been generated.** The build is a separate, later phase.
Owner: oscarwei · Spec date: 2026-09-11 · Primary tracking year: **2026**
Deliverable of this phase: this document only.

---

## 0. How to read this spec

Every metric in this document is specified as **cell address + literal formula**, and the spec is
implemented by `build_tracker.py` (one generator, several account profiles). Anything card-specific
lives in the generator's card data, not in this document.

> **Status, open action items and the change log live in [`PROGRESS.md`](PROGRESS.md).** Read that
> first when picking the project back up; this file stays the design reference.

Design principles the spec commits to (do not re-litigate):

1. **Issuer-official sources only** — a benefit exists if the issuer's own site (including its official
   press releases) lists it. A credit that appears only on aggregator sites is treated as non-existent
   (this is how the Sapphire Preferred Instacart credit was removed).
2. **Scope** — only statement credits, plus one exception: the airline **free checked bag**, carried as
   an `EST` row whose value the user sets.
3. **No `IF` inside any aggregate** — array-`IF` silently evaluates to `0` in some Excel versions
   (Section 3.3).
4. **Rebuilds must not destroy user input** (Section 9.8).

Legend used in formula blocks:

- `TEMPLATE` = the tab that is copied to create a new card tab.
- `Ref` = reference/lookup tab (Section 4).
- `Dashboard` = page 1.
- Ranges are written relative to the `TEMPLATE`/card tab unless prefixed.

---

## 1. Objectives and non-goals

### In scope
- Page-1 `Dashboard`: one row per card + portfolio rollups + per-card months held.
- One tab per card, holding that card's benefits, monthly checkbox grid, and per-card math.
- Automatic computation of: credits used, credits remaining, annual fees charged, net benefit
  (per card and portfolio-wide), months/years held, credit utilization, and credits at risk of expiring.
- Month-granular tracking, driven by one `Tracking year` cell so rollover to 2027 is a one-cell edit.
- Mixed reset-cycle model (calendar-year / cardmember-year / monthly / quarterly / semi-annual /
  multi-year / one-time).

### Out of scope (this phase and the build)
- Generating any `.xlsx` or `.xlsm` file. No macros/VBA.
- Sign-up bonus / welcome offer tracking.
- Points-and-miles valuation, transfer partners, redemption optimization.
- Credit score, 5/24 or other application-rule tracking.
- Tax or financial advice.
- Non-US cards. USD only, no FX conversion.
- A Google Sheets-specific build (the design must merely *survive* a Sheets/LibreOffice import).

---

## 2. Workbook architecture

| # | Tab | Purpose | Manual input? |
|---|-----|---------|---------------|
| 1 | `Dashboard` | Page 1: portfolio KPIs + one row per card | Only the tab-name column + card status |
| 2 | `Benefits` | Flat live index of EVERY benefit on EVERY card, filterable | Nothing - read-only, all formulas |
| 2 | `TEMPLATE` | Copy source for a new card tab | Yes (setup only) |
| 3..N | `<ISSUER>-<CODE>` | One per card, e.g. `AMEX-PLAT`, `CHASE-CSR` | Yes: identity block + benefit rows + the amounts you used |
| N+1 | `Ref` | Tracking year, dropdown lists, benefit catalog, month-label driver | Yes (year + catalog) |

**Build profiles.** One generator produces one workbook per account:

| Profile | Output file | Cards |
|---------|-------------|-------|
| `oscar` | `Oscar_credit_tracker.xlsx` | 11 |
| `myra` | `Myra_credit_tracker.xlsx` | 3 |

`./.venv/bin/python build_tracker.py <profile>` — each profile keeps its own `*.backup.xlsx` and
`.tracker_build_state_<profile>.json`, so rebuilding one never touches the other. Adding an account =
adding a card list plus a `PROFILES` entry in `build_tracker.py`.

Tab-naming rule: uppercase, `ISSUER-PRODUCTCODE`, ≤ 31 chars, none of `[ ] : * ? / \`.
Tab colour: `Dashboard` = dark blue, `Ref` = grey, `TEMPLATE` = yellow, card tabs = issuer colour.

Single source of truth: **card identity (open date, annual fee, AF posting date) is entered only on
the card tab.** `Dashboard` is 100% formula-driven and duplicates no input.

---

## 3. Card tab layout (`TEMPLATE`)

### 3.1 Identity block (rows 1–13)

| Cell | Label | Content | Type / validation |
|------|-------|---------|-------------------|
| A1 | Card | e.g. `Amex Platinum` | text |
| A2 | Issuer | e.g. `Amex` | dropdown from `Ref!$A$3:$A$12` |
| A3 | Product (exact) | e.g. `The Platinum Card` | text |
| A4 | Card status | `Active` / `Closed` / `Product-changed` | dropdown |
| A5 | Open date (approval) | `2023-08-15` | date |
| A6 | Months held | `=IF($B$5="","",DATEDIF($B$5,TODAY(),"m"))` | formula, integer |
| A7 | Years held | `=IF($B$6="","",ROUND($B$6/12,1))` | formula |
| A8 | Annual fee (USD) | `895` | number |
| A9 | First-year AF waived? | blank or `y` | dropdown `y` (error alert off) |
| A10 | AF posting date (current cycle) | `2026-08-15` | date — **drives the anniversary grid** |
| A11 | Authorized user AF (total, USD) | `0` | number |
| A12 | Total AF charged (this cycle) | `=IF($B$9="y",0,$B$8)+$B$11` | formula, currency |
| A13 | Next AF date (projected) | `=EDATE($B$10,12)` | formula, date |
| A14 | Cardmember year start (current cycle) | `=IFERROR(EDATE($B$10,12*DATEDIF($B$10,TODAY(),"y")),"")` | formula, date — **anchors Block B** |

Column A width 26, column B width 18. B5/B10/B13/B14 `yyyy-mm-dd`; B8/B11/B12 `"$"#,##0`.

### 3.2 KPI block (rows 3–14, columns D:E)

Labels in `D`, formulas in `E`. This block is the **only** thing `Dashboard` reads, so its addresses
are part of the contract.

| Cell | Label | Formula |
|------|-------|---------|
| D3 | Credits used YTD — statement credits | *(E3 below)* |
| D4 | Credits used YTD — incl. estimated | *(E4 below)* |
| D5 | Max annual credit value (annualized) | *(E5 below)* |
| D6 | Remaining annual credit value (incl. est.) | *(E6 below)* |
| D7 | Remaining — statement credits only | *(E7 below)* |
| D8 | Net benefit — statement credits | *(E8 below)* |
| D9 | Net benefit — incl. estimated | *(E9 below)* |
| D10 | Result | *(E10 below)* |
| D11 | AF breakeven gap | *(E11 below)* |
| D12 | Credit utilization % | *(E12 below)* |
| D13 | Credits at risk (unused, ≤30 days left) | *(E13 below)* |
| D14 | Nearest expiry (days) | *(E14 below)* |
| G3 *(hidden helper)* | Min days-left, Block A | `=MIN($Y$19:$Y$42)` |
| G4 *(hidden helper)* | Min days-left, Block B | `=MIN($Y$45:$Y$56)` |
| G5 *(hidden helper)* | Min of the two | `=MIN($H$3,$H$4)` |

```excel
E3  =SUMIF($B$19:$B$42,"SC",$G$19:$G$42)+SUMIF($B$45:$B$56,"SC",$G$45:$G$56)
    +SUMIFS($F$59:$F$66,$B$59:$B$66,"SC",$I$59:$I$66,"y")
E4  =SUMIF($B$19:$B$42,"SC",$G$19:$G$42)+SUMIF($B$19:$B$42,"EST",$G$19:$G$42)
    +SUMIF($B$45:$B$56,"SC",$G$45:$G$56)+SUMIF($B$45:$B$56,"EST",$G$45:$G$56)
    +SUMIFS($F$59:$F$66,$B$59:$B$66,"SC",$I$59:$I$66,"y")
    +SUMIFS($F$59:$F$66,$B$59:$B$66,"EST",$I$59:$I$66,"y")
E5  =$F$43+$F$57+SUM($F$59:$F$66)
E6  =MAX(0,$E$5-$E$4)
E7  =MAX(0,$E$5-$E$3)
E8  =$E$3-$B$12
E9  =$E$4-$B$12
E10 =IF($E$8>=0,"POSITIVE","NEGATIVE")
E11 =MAX(0,-$E$8)
E12 =IF($E$5=0,"",$E$4/$E$5)
E13 =SUMPRODUCT(($Y$19:$Y$42>=0)*($Y$19:$Y$42<=30))+SUMPRODUCT(($Y$45:$Y$56>=0)*($Y$45:$Y$56<=30))
E14 =IF($H$5>=9999,"",$H$5)
H3  =MIN($Y$19:$Y$42)
H4  =MIN($Y$45:$Y$56)
H5  =MIN($H$3,$H$4)
```

**Grid C is deliberately excluded from `E3`/`E4` except through the annualized `F` column.** An earlier
draft summed grid C's `J` (Value captured = the full `$120` for a Global Entry credit) *and* then added
annualized `F`, which double-counted it. `J` is a cash-flow reference only; annual totals use `F`.

Two Excel traps this block deliberately avoids (both confirmed empirically, Section 8):

1. **Never do arithmetic on a text-capable range inside `SUMPRODUCT`.** `H` legitimately contains `""`
   for empty rows; `COUNTIFS(range,">0")` counts text cells (Excel sorts any text above any number), and
   engines can raise `#VALUE!` on `("" > 0)`. `E13` therefore reads the numeric-only `Y` column, where
   "at risk" is just `Y >= 0 AND Y <= 30`.
2. **`MINIFS` returns `0`, not an error, when nothing matches** — a card with no unused credits would
   report "expires today" (observed: block B returned `0` where the correct answer was "nothing").
   `MIN(IF(…))` would work but needs Ctrl+Shift+Enter in legacy Excel. Hence the hidden `Y` column,
   which pre-reduces each row to "days left, or `9999` if no value remains", leaving a plain `MIN` over
   a purely numeric range: no array entry, no `AGGREGATE`/`MINIFS` dependency, and executable by any
   recalculation engine.

`E8`/`E9` are the definition of net benefit: **used credit value − total AF charged**.
Multi-year credits enter net benefit **annualized** (Section 5.4) so one Global Entry credit does not
create a fake +$120 year.

Also on the card tab:
- `G3` = `=HYPERLINK("#'Dashboard'!A1","◀ back to Dashboard")`
- Conditional formatting on `E10`: green fill if `$E$10="POSITIVE"`, red fill if `$E$10="NEGATIVE"`.

### 3.3 Benefit sections — grouped by Basis

**Revised 2026-09-11 at the user's request** ("把每一個 page 改成用 basis 類別來區分 section").
Rows are no longer grouped by reset cycle; they are grouped **by Basis**, because Basis is what decides
whether a number moves the totals — which is the thing the user actually needs to see.

Section order on every card tab (a section is emitted only if it has rows):

| Section | Contents | Month grid | Band colour |
|---------|----------|-----------|-------------|
| `1` | SC — statement credits, calendar cycle | Jan–Dec of `Ref!B1` | green |
| `1b` | SC — statement credits, anniversary cycle | anniversary months from `B14` | green |
| `2` | EST — your own estimates, calendar cycle | Jan–Dec | amber |
| `2b` | EST — anniversary cycle | anniversary months from `B14` | amber |
| `3` | ACCESS — perks with no dollar value | Jan–Dec | grey |
| `4` / `4b` / `4c` | one-time / every-N-years items, split by basis | none — mark the `Used?` column | blue |

Each section = merged band row (title) + column-header row + data rows + a `Subtotal` row + one blank
spacer row. **All row addresses are therefore dynamic per card** and are computed at build time; the KPI
formulas in 3.2 are assembled from those ranges rather than hard-coded. Every section still carries:

| Col | Header | Formula pattern |
|-----|--------|-----------------|
| A | Benefit | text |
| B | Basis | dropdown `SC` / `EST` / `ACCESS` |
| C | Cycle | dropdown, per section type |
| D | Value/cycle | number |
| E | Cyc/yr | `=IF($C…="MONTHLY",12,IF($C…="QUARTERLY",4,IF($C…="SEMI-ANNUAL",2,IF($C…="CALENDAR-YEAR",1,""))))` (anniversary sections: literal `1`) |
| F | Max annual | calendar: `=IF(OR($D…="",$E…=""),"",$D…*$E…)` · anniversary: `=IF($D…="","",$D…)` |
| G | Used | `=IF($D…="","",IF(ISNUMBER($F…),MIN($X…,$F…),$X…))` |
| H | Remaining | `=IF($F…="","",MAX(0,$F…-$G…))` |
| I | Uses | `=SUMPRODUCT(ISNUMBER($L…:$W…)*($L…:$W…>0))` |
| J/K | *(hidden)* cycle end / days left | `=IFERROR(EDATE($B$14,12)-1,"")` for anniversary sections, `CYCLE_END` otherwise / `=IFERROR($J…-TODAY(),"")` |
| L–W | month cells | amount entry (3.6) |
| X/Y | *(hidden)* raw used / days-left sentinel | 3.6 |
| Z | Notes / source | text, carries `PENDING` flags |

One-time sections use `MULTI_HDR` (`Benefit | Basis | Cycle | Value | Yrs/cycle | Annualized | Date used |
Expiry / next eligible | Used? (amount or y) | Value captured | Notes`) with
`F = =IF(AND(ISNUMBER($D…),ISNUMBER($E…)),$D…/$E…,0)` and
`J = =IF($I…="y",$D…,IF(ISNUMBER($I…),MIN($I…,$D…),0))`.

**Changing a row's Basis reclassifies its maths without moving it** — the KPI `SUMIF`s span every
section's range, so a row switched from ACCESS to EST starts counting immediately even while it sits in
the ACCESS section. Moving the row is only cosmetic.

#### The `X` column must never use `IF` inside `SUMPRODUCT`

`X` (raw used) originally read
`=SUMPRODUCT(IF(ISNUMBER($L19:$W19),IF($L19:$W19>$D19,$D19,$L19:$W19),0))`. **That evaluates to `0` in the
user's Excel** (the `IF` is not evaluated in array context without Ctrl+Shift+Enter, so `Used` stayed `$0`,
`E3` stayed `0`, and the whole Dashboard refused to move — the user reported exactly this on 2026-09-11).
The sibling `I` column uses the multiplication idiom `SUMPRODUCT(ISNUMBER(r)* (r>0))`, which *did* work in
the same workbook, so the failure is specific to array-`IF`.

It now uses only plain aggregates, which keeps the per-cell cap and is text-safe:

```excel
X19 = IF($D19="", SUM($L19:$W19),
             SUM($L19:$W19) - SUMIF($L19:$W19,">"&$D19)
             + (COUNT($L19:$W19) - COUNTIF($L19:$W19,"<="&$D19)) * $D19)
```

`SUM` ignores text; `SUMIF(">D")` matches text cells but sums them as 0; the `COUNT - COUNTIF("<="&D)`
pair counts only *numeric* cells above `D` (text compares greater than any number, so `COUNTIF("<="&D)`
safely excludes it). Verified: `50` → `X=50`; `200` against a `$100` cap → `X=150`, `G=100` (row cap);
a typed `y` → ignored. **Rule from here on: no `IF` inside any aggregate in this workbook.** The spec's
verification engine evaluates array-`IF` permissively, so this class of bug is invisible to it — prefer
constructs whose Excel behaviour is unambiguous.

### 3.4 Subtotals and the KPI wiring

Every section gets its own `Subtotal` row: `=SUM($F$first:$F$last)` and likewise for `G`, `H`, `I`
(one-time sections subtotal `D`, `F`, `J`). The card KPI block is then built from the section ranges:

```excel
E3  = <Σ over every section range> SUMIF($B$first:$B$last,"SC",$G$first:$G$last)
      + Σ SUMIFS($F…,$B…,"SC",$I…,"y")            (one-time sections only)
E4  = E3 + the same two terms for "EST"
E5  = Σ SUM($F$first:$F$last) over all sections
E6  = MAX(0,$E$5-$E$4)      E7 = MAX(0,$E$5-$E$3)
E8  = $E$3-$B$12            E9 = $E$4-$B$12
E10 = IF($E$8>=0,"POSITIVE","NEGATIVE")
E11 = MAX(0,-$E$8)          E12 = IF($E$5=0,"",$E$4/$E$5)
E13 = Σ SUMPRODUCT(($Y$first:$Y$last>=0)*($Y$first:$Y$last<=30))
H3  = MIN(<every Y range>)      E14 = IF($H$3>=9999,"",$H$3)
E15 = Σ COUNT(<month ranges>) + Σ COUNTIF($I$first:$I$last,">0")     <- amounts entered (cells)
E16 = Σ SUM(<month ranges>)  + Σ SUM($J$first:$J$last)               <- entered total, before capping
```

`E15`/`E16` are the **diagnostics added after the user reported "我 update 了我使用的 sc 但統計都沒變"**.
They separate "did my typing register at all?" from "is this row counted?": if `E15` rises but `E3`
does not, the row's Basis is `ACCESS` (perks deliberately carry no dollars). That ambiguity was the
most likely cause of the report and is now self-diagnosing on the sheet.

### 3.5 Cycle-end helper (column J, grid sections)

```excel
J19 =IF($C19="MONTHLY",EOMONTH(TODAY(),0),
     IF($C19="QUARTERLY",EOMONTH(DATE('Ref'!$B$1,ROUNDUP(MONTH(TODAY())/3,0)*3,1),0),
     IF($C19="SEMI-ANNUAL",IF(MONTH(TODAY())<=6,DATE('Ref'!$B$1,6,30),DATE('Ref'!$B$1,12,31)),
     DATE('Ref'!$B$1,12,31))))
```

This is what makes "credits at risk" possible: cycle end − today ≤ 30 days while value is unused.
It encodes the researched reset semantics — monthly resets on the 1st (window = calendar month),
quarterly = calendar quarters, semi-annual = **Jan–Jun / Jul–Dec** (not "six months from signup"),
calendar-year = Dec 31.

### 3.6 Amount-entry cells (columns L–W in Blocks A and B)

Revised per user feedback: a credit is rarely used to its full amount (a $200 hotel credit might be
$150), so **the cell takes the dollar amount actually used — not a `y` checkbox**. Any entry above the
credit amount is **capped at the credit amount** rather than counted at face value.

- Data validation: `type="decimal"`, `operator="greaterThanOrEqual"`, `formula1=0`, `allow_blank=True`,
  `showErrorMessage=True` with an input prompt ("type the dollar amount used in this cycle; blank = not
  used; values above the credit are capped automatically"). Non-numeric input such as `y` is rejected.
- Cell semantics: blank = unused · any number ≥ 0 = USD used in that cycle.
- **Two levels of capping**, both automatic:
  1. *Per cell (per cycle):* an entry is capped at `D` (`Value per cycle`), so $250 typed against a $200
     semi-annual credit counts as $200.
  2. *Per row (per year):* the row total is capped at `F` (`Max annual`), so the annual maximum for a
     $7/month credit stays $84 no matter how many months are filled in.
- Conditional formatting (applied to `$L$19:$W$42` and `$L$45:$W$56`), in this priority order:

| # | Name | Formula (written for the top-left cell of the range) | Format | Stop? |
|---|------|------------------------------------------------------|--------|-------|
| 1 | Non-numeric entry | `=AND(L19<>"",NOT(ISNUMBER(L19)))` | red fill | yes |
| 2 | Over the cycle cap (capped) | `=AND(ISNUMBER(L19),$D19>0,L19>$D19)` | orange fill | yes |
| 3 | Over-claim Q1 | `=AND($C19="QUARTERLY",COUNTIF($L19:$N19,">0")>1)` | red text | no |
| 4 | Over-claim Q2 | `=AND($C19="QUARTERLY",COUNTIF($O19:$Q19,">0")>1)` | red text | no |
| 5 | Over-claim Q3 | `=AND($C19="QUARTERLY",COUNTIF($R19:$T19,">0")>1)` | red text | no |
| 6 | Over-claim Q4 | `=AND($C19="QUARTERLY",COUNTIF($U19:$W19,">0")>1)` | red text | no |
| 7 | Over-claim H1 | `=AND($C19="SEMI-ANNUAL",COUNTIF($L19:$Q19,">0")>1)` | red text | no |
| 8 | Over-claim H2 | `=AND($C19="SEMI-ANNUAL",COUNTIF($R19:$W19,">0")>1)` | red text | no |
| 9 | Used | `=AND(ISNUMBER(L19),L19>0)` | green fill | no |

`COUNTIF(range,">0")` is safe here (unlike on the `H` column) because the cells hold numbers or blanks
only — non-numeric input is blocked by the validation and flagged by rule 1. Rules 3–8 catch the one case
capping cannot: two separate uses inside the same quarter or half, where the per-cell cap would otherwise
let both through.

Block C's `I` column keeps both options (`y` or a number) because a one-time benefit is normally used in
full: `J59 = =IF($I59="y",$D59,IF(ISNUMBER($I59),MIN($I59,$D59),0))`, so an amount over `D59` is capped too.
Its dropdown allows `y` and free text.

Rows 44–56 need the same rules with the row-44 anchors and their own `$C` values; rules 3–8 are inert
there because `$C` never equals QUARTERLY/SEMI-ANNUAL in Block B.

### 3.7 Layout, freezing, protection

- `ws.freeze_panes = "L19"` on card tabs (identity + KPI + the row-17 "where to type" legend stay visible).
- Column widths: A 30, B 9, C 15, D 12, E 9, F 12, G 12, H 12, I 9, J 11, K 9, L–W 7, Z 62
  (as built: A 34, B 10, C 15, D 13, E 8, F 12, G 12, H 12, I 10, J 11, K 9, L–W 7, Z 62).
- Hide columns J, K, X and Y (`column_dimensions[...].hidden = True`). Block A and B both hide J/K/Y;
  X is hidden in both as well; Block C has no hidden columns.
- Row 18 / 44 / 58 headers: bold, wrap, grey fill, thin bottom border. Subtotal rows: bold, top border.
- No sheet protection (the user edits freely); protection is a possible later addition.
- Sheet-level conditional formatting also greys out rows whose `$C` is blank, to keep unused slots quiet:
  `=AND($A19<>"",$C19="")` → light grey font, applied to `$A$19:$I$42`.

---

## 4. `Ref` tab

| Cell | Content |
|------|---------|
| A1 | `Tracking year` |
| B1 | `2026` ← **the single rollover cell** |
| A3:A12 | Issuer list: `Amex, Chase, Citi, Capital One, Bank of America, Discover, Wells Fargo, US Bank, Barclays, Other` |
| A14:A18 | Basis list: `SC, EST, ACCESS` |
| A20:A26 | Cycle list: `MONTHLY, QUARTERLY, SEMI-ANNUAL, CALENDAR-YEAR, CARDMEMBER-YEAR, MULTI-YEAR, ONE-TIME` |
| A28:A31 | Status list: `Active, Closed, Product-changed, Authorized-user-only` |
| A33:D46 | Benefit catalog (Section 5.5) |
| J1 | Label: `Short month names (English)` |
| J3:J14 | `Jan` `Feb` `Mar` `Apr` `May` `Jun` `Jul` `Aug` `Sep` `Oct` `Nov` `Dec` — single source for every month label |

Month-label driver used by every card tab:

```excel
L18 =INDEX('Ref'!$J$3:$J$14,COLUMN()-COLUMN($L$18)+1)   → fill right to W18
```

`Tracking year` also feeds `DATE('Ref'!$B$1,…)` inside every cycle-end helper, so **changing B1 rolls
the whole workbook to the next year** (Section 10).

**Why `INDEX` and not `TEXT(date,"mmm")`:** `TEXT` format codes follow the Excel *locale*, so on a
zh-TW install `TEXT(DATE(2026,1,1),"mmm")` renders `1月`, not `Jan`. Since the workbook content must
be English, month names come from the typed `Ref!J3:J14` list. `TEXT(...,"yy")` remains safe because
numeric format codes are locale-neutral.

---

## 5. Benefit taxonomy and reset-cycle model

> **What is actually in use today:** only `SC` rows and `EST` rows exist in either workbook. The only
> `EST` rows are the airline free-checked-bag rows (their dollar value is the user's own estimate).
> The `ACCESS` machinery below is implemented and tested but no `ACCESS` rows remain, because the scope
> was narrowed to statement credits + the checked bag. Cycle classes in use: `MONTHLY`, `QUARTERLY`,
> `SEMI-ANNUAL`, `CALENDAR-YEAR`, `CARDMEMBER-YEAR`, `MULTI-YEAR`, `ONE-TIME`.

### 5.1 The eight cycle classes

| Class | Window | Resets | Month grid | Used-value rule |
|-------|--------|--------|-----------|-----------------|
| `MONTHLY` | calendar month | 1st of month | Block A, 12 cols | 1 mark per month |
| `QUARTERLY` | calendar quarter | Jan/Apr/Jul/Oct 1 | Block A, 12 cols | 1 mark per quarter |
| `SEMI-ANNUAL` | Jan–Jun / Jul–Dec | Jan 1 & Jul 1 | Block A, 12 cols | 1 mark per half |
| `CALENDAR-YEAR` | Jan 1 – Dec 31 | Jan 1 | Block A, 12 cols | 1 mark per year |
| `CARDMEMBER-YEAR` | anniversary → +12 months − 1 day | on anniversary | Block B, anniversary-relative | 1 mark per cardmember year |
| `MULTI-YEAR` | e.g. 4 years | on next-eligible date | Block C | annualized value |
| `ONE-TIME` | once ever | n/a | Block C | full value in the year used |
| `ACCESS` | no dollar value | n/a | any (Basis = ACCESS) | counts uses only; value 0 |

`Basis` is orthogonal to `Cycle` and controls whether a row counts toward net benefit:
`SC` (statement credit, face value) → counts · `EST` (user-set estimated value, e.g. free checked
bag) → counts only in the "incl. estimated" view · `ACCESS` (lounge/Priority Pass, no value) →
counts uses, contributes $0.

### 5.2 The two "year" definitions and why both appear

The user chose a mixed model, so exactly one ambiguity must be understood:

- `Dashboard` "Used YTD" is the **calendar year** (`Ref!B1`) for Blocks A and C, and the **current
  cardmember year** for Block B (the window that contains today).
- `Months held` (A6) is always `TODAY() − open date`.
- Net benefit therefore answers "am I ahead *this cycle*", which is the correct question, but a
  cardmember-year card that opened in August will read "net" against an August–July window, not Jan–Dec.
  Documented in the tab note rather than flattened, because preserving the real reset date is the point.

### 5.3 Partial usage

A `y` mark means "used the full per-cycle value". A typed number means partial. This matters for
Amex-style monthly credits where $15 is rarely fully used; without it the net-benefit number would be
systematically optimistic. Rows where partial use is likely are marked `SC` with the real value in D.

### 5.4 Multi-year amortization

`F = D / E` (e.g. $120 / 4 years = $30/yr). Net benefit (`E8`/`E9`) adds the annualized figure only
when the row is marked used. The full $120 is still visible via `J` (Value captured) for cash-flow
viewing, but it never inflates the annual net benefit.

### 5.5 Benefit catalog seed (typical cycles — ALL VALUES PENDING VERIFICATION)

Starting library for the per-card fills. **Verification status: PENDING.** The dollar figures below come
from secondary sources and issuer marketing pages current as of 2026; each must be confirmed on the
issuer's own terms page for the specific card and date before being written into a card tab. Where a
figure cannot be confirmed it goes into the row's Notes as `PENDING` and D is left blank rather than guessed.

| Benefit type | Typical cycle | Typical value | Source checked | Status |
|---|---|---|---|---|
| Airline fee credit | CALENDAR-YEAR | card-dependent | issuer benefits page | PENDING |
| Hotel credit | SEMI-ANNUAL | e.g. $300 / half | issuer benefits page | PENDING |
| Uber / rideshare cash | MONTHLY | e.g. $15/mo (+$20 Dec) | issuer + Uber terms | PENDING |
| Digital entertainment | MONTHLY | e.g. $25/mo | issuer benefits page | PENDING |
| Dining / Resy | QUARTERLY | e.g. $100/quarter | issuer benefits page | PENDING |
| Retail (e.g. lululemon) | QUARTERLY | e.g. $75/quarter | issuer benefits page | PENDING |
| Saks Fifth Avenue | SEMI-ANNUAL | e.g. $50 / half | issuer benefits page | PENDING |
| Equinox | CALENDAR-YEAR | e.g. $300/yr | issuer benefits page | PENDING |
| Walmart+ membership | MONTHLY | membership cost | issuer benefits page | PENDING |
| CLEAR+ credit | CALENDAR-YEAR | membership cost | issuer benefits page | PENDING |
| Global Entry / TSA PreCheck | MULTI-YEAR (4 yr) | e.g. up to $120 | Amex/Capital One benefit pages | PENDING |
| Travel credit (e.g. CSR $300) | CARDMEMBER-YEAR | e.g. $300 | Chase card page | PENDING |
| Free checked bag | ACCESS or EST | per-trip estimate | airline card terms | PENDING |
| Lounge access / Priority Pass | ACCESS | $0 (count uses) | issuer benefits page | PENDING |
| Cell phone protection | ACCESS | claim-based | issuer benefits page | PENDING |

Sources already reviewed for cycle semantics (not for per-card amounts):
Amex Platinum benefits overview and card-benefits pages; Chase Sapphire Reserve card page and
Ultimate Rewards agreement; issuer/card-guide comparisons of reset schedules.

### 5.6 What the user must supply vs. what gets researched

**User supplies (per card):** issuer · product name (exact) · open/approval date · annual fee ·
first-year AF waived? · AF posting month · authorized-user AF · card status.

**Researched by me (per card, from issuer terms):** every benefit row's name, basis, cycle, value per
cycle, redemption channel, and the Notes/expiry wording. Unverifiable items are flagged `PENDING`.

---

## 6. `Dashboard` (page 1) — layout

```
Row 1        Title: "Credit Card Benefits & Annual Fee Tracker"
Row 2        Portfolio rollup as of  =TODAY()
Rows 3–6     Portfolio KPI cards
Row 8        Table header
Rows 9–28    One row per card (20 slots)
Row 29       Totals (Active cards only)
Rows 31–34   Legend / notes / link to Ref
```

### 6.1 Portfolio KPI block (rows 3–6)

| Cell | Label | Formula |
|------|-------|---------|
| B3 | Total annual fees charged (active cards) | `=SUMIF($G$9:$G$28,"Active",$H$9:$H$28)` |
| D3 | Credits used — statement credits | `=SUMIF($G$9:$G$28,"Active",$J$9:$J$28)` |
| F3 | Credits used — incl. estimated | `=SUMIF($G$9:$G$28,"Active",$K$9:$K$28)` |
| H3 | Remaining credit value | `=SUMIF($G$9:$G$28,"Active",$L$9:$L$28)` |
| B4 | Portfolio net benefit — SC | `=SUMIF($G$9:$G$28,"Active",$M$9:$M$28)` |
| D4 | Portfolio net benefit — incl. est. | `=SUMIF($G$9:$G$28,"Active",$N$9:$N$28)` |
| F4 | Portfolio result | `=IF($B$4>=0,"POSITIVE","NEGATIVE")` |
| H4 | Credit utilization % | `=IF($I$29=0,"",$F$3/$I$29)` |
| B5 | Cards tracked / active | `=COUNTA($A$9:$A$28)&" / "&COUNTIF($G$9:$G$28,"Active")` |
| D5 | Total months held (all active) | `=SUMIF($G$9:$G$28,"Active",$E$9:$E$28)` |
| F5 | Average months held | `=IFERROR(AVERAGEIF($G$9:$G$28,"Active",$E$9:$E$28),"")` |
| H5 | Credits at risk (≤30 days) | `=SUMIF($G$9:$G$28,"Active",$Q$9:$Q$28)` |
| B6 | Nearest expiry across active cards (days) | `=IF(MIN($S$9:$S$28)>=9999,"",MIN($S$9:$S$28))` |

Green fill on `F4` when `POSITIVE`, red fill when `NEGATIVE`; amber fill on `H5` when > 0.

Column `S` is a hidden helper implementing the same "`9999` if nothing qualifies" pattern as the card
tabs' `Y` column, so no `AGGREGATE`/`MINIFS` and no Ctrl+Shift+Enter array entry is needed anywhere in
the workbook:

```excel
S9 =IF(AND($G9="Active",ISNUMBER($R9)),$R9,9999)      → fill down to S28
```

### 6.2 Per-card table

Header row 8; data rows 9–28. Column A is the **only** manual column (the card tab name);
every other cell is a formula pointing at that card's tab. To add a card: copy an existing row down,
then Find/Replace the sheet name inside that row (the doc deliberately uses direct references instead
of `INDIRECT`, because `INDIRECT` is volatile and silently breaks on tab rename — an optional
`INDIRECT` variant is noted in Section 9.4).

| Col | Header | Formula pattern (row 9, card tab `AMEX-PLAT`) |
|-----|--------|-----------------------------------------------|
| A | Card (tab) | manual `AMEX-PLAT`, display via `=HYPERLINK("#'AMEX-PLAT'!A1","AMEX-PLAT")` where needed |
| B | Issuer | `='AMEX-PLAT'!$B$2` |
| C | Product | `='AMEX-PLAT'!$B$3` |
| D | Open date | `='AMEX-PLAT'!$B$5` |
| E | Months held | `='AMEX-PLAT'!$B$6` |
| F | Years held | `='AMEX-PLAT'!$B$7` |
| G | Card status | `='AMEX-PLAT'!$B$4` |
| H | AF this cycle | `='AMEX-PLAT'!$B$12` |
| I | Max annual credit value | `='AMEX-PLAT'!$E$5` |
| J | Used YTD — SC | `='AMEX-PLAT'!$E$3` |
| K | Used YTD — incl. est. | `='AMEX-PLAT'!$E$4` |
| L | Remaining | `='AMEX-PLAT'!$E$6` |
| M | Net benefit — SC | `='AMEX-PLAT'!$E$8` |
| N | Net benefit — incl. est. | `='AMEX-PLAT'!$E$9` |
| O | Result | `='AMEX-PLAT'!$E$10` |
| P | Utilization % | `='AMEX-PLAT'!$E$12` |
| Q | Credits at risk | `='AMEX-PLAT'!$E$13` |
| R | Nearest expiry (days) | `='AMEX-PLAT'!$E$14` |
| S | *(hidden helper)* Days-left if active | `=IF(AND($G9="Active",ISNUMBER($R9)),$R9,9999)` |

Number formats: H–N `"$"#,##0`; P `0%`; D date; E/F integers/1-decimal; M/N red when negative.
Conditional formatting: `$O9="POSITIVE"` green fill, `"NEGATIVE"` red fill; `$Q9>0` amber fill;
`$G9="Closed"` grey-strikethrough on the whole row.

### 6.3 Totals row (row 29)

```excel
E29 =IFERROR(AVERAGEIF($G$9:$G$28,"Active",$E$9:$E$28),"")
H29 =SUMIF($G$9:$G$28,"Active",$H$9:$H$28)
I29 =SUMIF($G$9:$G$28,"Active",$I$9:$I$28)
J29 =SUMIF($G$9:$G$28,"Active",$J$9:$J$28)
K29 =SUMIF($G$9:$G$28,"Active",$K$9:$K$28)
L29 =SUMIF($G$9:$G$28,"Active",$L$9:$L$28)
M29 =SUMIF($G$9:$G$28,"Active",$M$9:$M$28)
N29 =SUMIF($G$9:$G$28,"Active",$N$9:$N$28)
O29 =IF($M$29>=0,"POSITIVE","NEGATIVE")
P29 =IF($I$29=0,"",$K$29/$I$29)
Q29 =SUMIF($G$9:$G$28,"Active",$Q$9:$Q$28)
```

`SUMIF` on status means closed cards stop polluting the portfolio totals while their history stays visible.

---

## 7. Requirements traceability

| # | Original requirement | Where addressed |
|---|----------------------|-----------------|
| 1 | English spreadsheet content | Sections 3–6: all labels/headers are English |
| 2 | Page 1 aggregates all cards | Section 6.2 per-card table + 6.1/6.3 rollups |
| 3 | Time since each card was opened | `B5`+`B6`/`B7` per card, `E`/`F`+`D5`/`F5` on Dashboard |
| 4 | One tab per card | Section 2; created from `TEMPLATE` |
| 5 | Per-benefit detail (hotel credit, free checked bag, …) | Sections 3.3, 5.5 |
| 6 | Date-aware benefits | Sections 5.1, 3.5, Block B anniversary grid |
| 7 | Month-granular tracking | Blocks A/B month grids, `Ref!B1`-driven labels |
| 8 | Checkbox so use is easy to mark | Section 3.6 (`y` dropdown / number, green fill) |
| 9 | Auto-computed used credit amount | `X` raw → `G` used → `G43`/`G57` → `E3`/`E4` → `J`/`K` on Dashboard |
| 10 | Auto-computed annual fee | `B12` → `H` on Dashboard → `H29` |
| 11 | Is my benefit position positive or negative | `E8`/`E9`/`E10` → `M`/`N`/`O` → `B4`/`F4`/`O29` |
| 12 | Page 1 live calculation of fees + used benefits | Entire `Dashboard` is formula-driven; no hard-coded numbers |

---

## 8. Verification of this spec

| Check | Result |
|-------|--------|
| Every Dashboard metric has a cell + literal formula | Pass — Sections 6.1, 6.2, 6.3 |
| Every card-tab metric has a cell + literal formula | Pass — Sections 3.2, 3.3, 3.4 |
| Every checkbox column has DV + CF + consuming formula | Pass — Section 3.6 (8 CF rules + DV + `X`/`G`) |
| Every benefit row is classified into one reset cycle | Pass — Section 5.1; `C` column dropdown enforced |
| Monthly granularity with correct real-world reset semantics | Pass — Section 3.5 encodes month/quarter/half/calendar windows |
| Card-inventory intake fields complete enough to fill card tabs | Pass — Section 5.6 |
| No `.xlsx`/`.xlsm` generated in this phase | Pass — this document is the only artifact |
| Implementation can proceed without new design decisions | Pass — Section 9 checklist is literal |
| Static audit of the spec itself: 87 formulas parsed; every cell reference resolves inside a declared layout block; the 17-cell `Dashboard`←card-tab read contract is complete and every target cell exists; only whitelisted Excel functions used; parentheses/quotes balanced (incl. the multi-line `J19`) | Pass — `Ref!J3:J14` and the two hidden helpers `H3`/`H4` were the only cells added to close the audit |
| Semantic behaviour of the formulas — **executed** | Pass. A faithful prototype of the card tab + `Dashboard` was built in a scratch directory and recalculated with the `formulas` engine. Scenario: Uber `y`×3, Resy `y`×2, hotel `y`×1, checked-bag `y`×1, travel-credit `y`×1, Global Entry used, AF `$895`. Verified cell-by-cell — `G19=45` `G21=200` `G22=300` `G23=35` `G45=300` `F43=1800` `F57=300` `F59=30` `B12=895` `E3=875` `E4=910` `E5=2130` `E6=1220` `E7=1255` `E8=-20` `E9=15` `E10=NEGATIVE` `E11=20` `E13=3` `E14=19`, and on `Dashboard` `H29=895` `J29=875` `K29=910` `M29=-20` `O29=NEGATIVE` `Q29=3` `R9=19`. Cross-sheet references into the hyphenated quoted tab name `'AMEX-PLAT'!` resolved correctly, the anniversary grid rendered `Aug-26 … Jul-27`, and `DATEDIF` gave the expected 36 months. `Dashboard!B6` was executed separately against a `Active/19, Closed/5, Active/111, Active/blank, Active/None` fixture → `19` (closed row and blanks correctly ignored, `S` sentinels `[19, 9999, 111, 9999, 9999]`), and against an all-closed fixture → `""` rather than a misleading "expires today" |
| Four defects this execution caught (all fixed above) | (1) `E3`/`E4` summed grid C's full `J` value while `E8`/`E9` added annualized `F` as well — multi-year credits were double-counted; annual totals now use `F` only. (2) `MINIFS` returned `0` rather than an error for a block with nothing unused, which would have shown "expires today"; replaced by the hidden `Y` helper + plain `MIN`. (3) `F45 = =$D45` returns **`0`** for an empty row (a bare reference to a blank cell is `0`, not `""`), so `H45 = MAX(0,$F45-$G45)` evaluated `0 - ""` → **`#VALUE!`** on every unused cardmember-year slot; fixed to `=IF($D45="","",$D45)`. This one only surfaced when the real 11-card workbook was recalculated — the synthetic probe had no text-containing `H` range in Block B. **(4) Block B was anchored on `B10` (the AF posting date), so any card whose anniversary had already passed this year showed a cycle-end date in the PAST and its credits silently dropped out of the at-risk logic** — caught 2026-09-11 when the user challenged the CSP hotel credit; now anchored on the new `B14` (current cardmember-year start), verified as CSP `2026-03-27 → 2027-03-26` with labels `Mar-26 … Feb-27`, Aspire `2026-03-29 → 2027-03-28`, Atmos/Green `2025-09-19 → 2026-09-18` |
| Built workbook re-executed (11 cards, real data) | Pass. Aggregates: total AF charged this cycle `$950`, max annualized credit value `$2,203`, portfolio net `-$950` with nothing ticked, `11 / 11` active. Per-card spot checks: Aspire `B12=$550`, RH-Gold `B12=$60` (includes the Gold membership), Citi AA `B12=$0` (first-year waiver ticked), Green `E5=$189`, BCE `E5=$264`, Atmos/Costco net `$0`. Behaviour tests: partial use (enter `150` against the Aspire $200 resort credit → row used `$150`, `E8` moves by +$150, result POSITIVE); multi-year (Global Entry marked → `F59=$30` annualized, `E3=$30`, `E8=-$65`); portfolio rollup `B4=$-301`. Zero `#VALUE!`/`#DIV/0!` cells anywhere |
| Capping behaviour (user-requested revision, 2026-09-11) | Pass. Month cells take a **dollar amount** — the `y` checkbox was removed at the user's request — and an entry above the credit value is capped at the credit value at **two** levels. Executed against the built workbook: `150` + `250` on the Aspire `$200` semi-annual resort credit → row used `$350` (the 250 capped to 200), remaining `$50`; `80` on the `$50` quarterly flight credit → `$50`; `10`/`7`/`3` on the BCE `$7/month` Disney credit → `$17` (the 10 capped to 7) with `Uses = 3`, and the row can never exceed `Max annual` `$84`. Totals propagated correctly to `E3`/`E8` and the portfolio rollup (`E3=$619`, `E8=+$69` Aspire; `E3=$29` BCE) |
| Engine coverage limits (not spec gaps) | The `formulas` engine implements every function used **except `AGGREGATE`**, which the earlier draft used in `H3`/`H4` (and `Dashboard!B6`). That unverifiability motivated the switch to the `Y` (card tab) and `S` (Dashboard) "`9999` sentinel" helpers, which are plain `MIN` over numeric ranges — so the spec now needs no `AGGREGATE`, no `MINIFS`, and no Ctrl+Shift+Enter array entry at all. The engine also mishandles `HYPERLINK` (and renames its `#NAME?` nodes inconsistently between runs); `HYPERLINK` is used only for the tab navigation links, is standard Excel, and carries no calculation risk. Both functions are fine in Excel 2010+/LibreOffice/Sheets, so this was a portability and verifiability improvement, not a bug fix |

---

## 9. Implementation checklist (for the later build phase)

9.1 **Create workbook** `credit_card_tracker.xlsx`; `wb.calculation.fullCalcOnLoad = True`
(set it, or viewers that read cached values show blanks/zeros instead of computed results).

9.2 **Build `Ref`** first: `Tracking year` cell `B1`, the four dropdown lists, the benefit catalog,
and the month-label formula row.

9.3 **Build `TEMPLATE`**: identity block, KPI block, three benefit blocks, subtotal rows, hidden
columns, freeze panes, then apply, in this order per block:
```python
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill

mark = DataValidation(type="list", formula1='"y"', allow_blank=True,
                      showErrorMessage=False)          # dropdown + free-text numbers
ws.add_data_validation(mark); mark.add("L19:W42"); mark.add("L45:W56")

green = PatternFill("solid", start_color="C6EFCE")
ws.conditional_formatting.add("L19:W42",
    FormulaRule(formula=['OR(L19="y",AND(ISNUMBER(L19),L19>0))'], fill=green))
ws.conditional_formatting.add("L45:W56",
    FormulaRule(formula=['OR(L45="y",AND(ISNUMBER(L45),L45>0))'], fill=green))
# + the 7 guard/over-value rules from Section 3.6, each added with the correct range
```
Repeat the CF block for `L45:W56`. Keep rules 1–7 above rule 8 (openpyxl appends, so add the green
rule last, and set `stopIfTrue` on the "Over value" rule).

9.4 **Build `Dashboard`**: KPI block, header row, one formula row per card, totals row, conditional
formats. *Optional* alternative to Find/Replace when adding cards: replace direct references with
`=INDIRECT("'"&$A9&"'!$E$8")`-style formulas driven by column A. Trade-off: adding a card becomes a
one-row edit, but formulas are volatile and break silently if a tab is renamed. Default is direct
references.

9.5 **Copy `TEMPLATE`** once per card from the user's inventory, rename, and fill identity + benefits.

9.6 **Sanity pass**: no `#REF!`/`#VALUE!` anywhere; `Dashboard!H29` equals the sum of the card tabs'
`B12`; marking one `y` in each cycle type moves `E3` by the expected amount; LibreOffice and Google
Sheets open the file with dropdowns and colours intact (Sheets drops some CF priorities — the green
"used" rule is the one that must survive).

9.7 **Deliver**: the `.xlsx`, plus a short README describing how to add a card and roll the year.

---

## 9.8 Rebuild safety (added 2026-09-11)

`build_tracker.py` regenerates the whole workbook, so from this date it **harvests the user's entries from
the existing `credit_card_tracker.xlsx` first and writes them back afterwards**, keyed by
`(sheet, benefit name, month index)` so it survives sections moving around:

- preserved: month amounts, a `Value/cycle` the user changed, the `Used?` column on one-time rows, the
  identity block (`B1`–`B5`, `B8`–`B11`) and `Ref!B1`;
- a `Value/cycle` is only preserved when it differs from the default the last build wrote — the defaults
  are recorded in `.tracker_build_state.json`, so generator default changes still take effect;
- column `I` (a formula on grid rows) is only harvested for `MULTI-YEAR`/`ONE-TIME` rows, where it is a
  real input;
- the previous file is copied to `credit_card_tracker.backup.xlsx` before every rebuild;
- **not** preserved: a row the user renames (the benefit name is the key) and a row the user deletes
  outright (the generator recreates the whole structure). Both need a change in `CARDS` instead.

## 10. Year rollover 2026 → 2027

1. `File > Save As` → `credit_card_tracker_2027.xlsx`; keep the 2026 copy as a frozen archive.
2. On `Ref`, set `Tracking year` = `2027`. Month labels, cycle-end dates, and calendar-year windows
   all move automatically (single-cell rollover — this is why `Ref!B1` exists).
3. Clear the mark ranges on every card tab: `L19:W42`, `L45:W56`, and grid C's `G`/`I` columns.
   (Manual in this design; a small openpyxl reset script is an obvious later addition but is out of
   scope here.)
4. Update per card: `AF posting date (current cycle)` (B10) to the new anniversary, `Annual fee` (B8)
   if it changed, `First-year AF waived?` (B9) back to blank, and `Card status` (B4).
5. Re-verify benefit values: issuers refresh credits annually; re-check Section 5.5 rows and update
   `D` / `E` / `C` where terms changed. Record the check date in the row Notes.
6. Add newly opened cards: copy `TEMPLATE`, rename, fill, then add a Dashboard row.
7. Block B re-labels itself from B10 — no manual month editing.

**One-time 2026 backfill:** for the 2026 workbook, mark Jan–Aug usage month by month from statements.
Where a past month's usage is unknown, leave the cell blank and write `UNVERIFIED Jan–Aug` in the row
Notes, so YTD is never overstated.

---

## 11. Open items and blockers

> **The live action list is in `PROGRESS.md` §5** (with owners and deadlines). Only design-level open
> questions are kept here.

| # | Item | Status |
|---|------|--------|
| ~~B1~~ | ~~User's card inventory~~ | **Done** — both accounts are built (`PROGRESS.md` §4) |
| B2 | Benefit amounts and terms | Researched per card. Chase cards are fully backed by issuer pages; **Amex, Bank of America and Citi pages cannot be read by any fetch path available here** (JS-rendered / no content / 404), so those rows rest on aggregators and are flagged in `PROGRESS.md` §6 |
| Q1 | Valuation basis for non-statement benefits (free checked bag, lounge access) | Default selected: show **both** views (SC-only and incl. estimated), with `Basis` per row. Switchable by setting those rows to `ACCESS` (excluded entirely) if preferred |
| Q2 | Multi-year architecture | Default selected: single-year workbook + copy-on-rollover (Section 10). The "2026 and 2027 tabs side by side" alternative would require parameterising every cycle window and roughly doubles the formula surface; not recommended |
| Q3 | Cards opened mid-year show a negative net benefit because the AF is charged in full while credits accrue over 12 months | Mitigated by also showing `Months held` and `AF breakeven gap` (`E11`); a "monthly run-rate net benefit" metric is available on request |
| Q4 | Retentions / product changes mid-year | Modeled as `Card status` + a new card tab with its own open date; no retroactive splitting |
| A1 | Currency | USD only, no FX conversion |
| A2 | Credit value definition | Face value of the statement credit, regardless of whether the spend would have happened anyway (the "incl. estimated" view is the place for subjective value) |
| R1 | Excel-specific *rendering* remains unverified: dropdown behaviour, conditional-format paint order, `HYPERLINK` navigation, cell formats | Arithmetic is now executed and verified (Section 8); only presentation is unobserved. Closed by Section 9.6's smoke test as the first build-phase action |
| R2 | Excel-locale rendering (`TEXT` codes, list separators, date formats) | Month labels neutralised via `Ref!J3:J14` + `INDEX` (Section 4). Formula argument separators in this doc are commas (en-US convention); a locale using semicolons will have them converted automatically by Excel on paste |

---

## 12. Build record

> **The running change log is in `PROGRESS.md` §9.** This section records only what the spec built and
the design defects found while building it.

The spec was implemented on 2026-09-11 by `build_tracker.py`. Sheet order: `Dashboard`, `Benefits`,
N card tabs, `TEMPLATE`, `Ref`. Current output files: `Oscar_credit_tracker.xlsx` (11 cards) and
`Myra_credit_tracker.xlsx` (3 cards).

| Tab | Issuer | Product | Opened | AF this cycle | Notes |
|-----|--------|---------|--------|---------------|-------|
| CITI-COSTCO | Citi | Costco Anywhere Visa | 2025-09-19 | $0 | needs a paid Costco membership |
| AMEX-ASPIRE | Amex | Hilton Honors Aspire | 2026-03-29 | $550 | resort $200/half, flight $50/quarter, CLEAR $219/yr |
| CHASE-CSP | Chase | Sapphire Preferred | 2025-03-27 | $95 | hotel $100/anniversary, DoorDash $10/mo, Global Entry 4-yr |
| BOA-ATMOS-ASCENT | Bank of America | Atmos Rewards Ascent Visa Signature | 2025-09-19 | $95 | companion fare needs $6k prior-year spend |
| RH-GOLD | Robinhood | Gold Card | 2026-06-01 | $60 | $60 = the required Gold membership ($5/mo); card itself is $0 |
| AMEX-DELTA-BLUE | Amex | Delta SkyMiles Blue | 2025-06-23 | $0 | |
| CITI-AA-PLAT | Citi | AAdvantage Platinum Select | 2025-12-24 | **$0** | first-year AF waiver ticked; untick at the 2026-12 anniversary |
| CHASE-CFU | Chase | Freedom Unlimited | 2024-12-09 | $0 | |
| AMEX-GREEN | Amex | Green ("Classic Green") | 2025-09-19 | $150 | LoungeBuddy credit ended 2025-01-13 |
| BOA-CUSTOMIZED-CASH | Bank of America | Customized Cash Rewards | 2024-08-17 | $0 | |
| AMEX-BCE | Amex | Blue Cash Everyday | 2024-08-14 | $0 | Disney $7/mo, Home Chef $15/mo |

**Basis-major re-layout (2026-09-11).** Card tabs were regrouped from cycle-major into Basis-major
sections (1 SC / 1b SC-anniversary / 2 EST / 2b EST-anniversary / 3 ACCESS / 4–4c one-time), each with its
own band, subtotal and colour; the `Benefits` tab was added as a flat live index of all 57 benefits;
and `E15`/`E16` diagnostics were added after the user reported that entering usage did not move the
totals. Re-verified by execution: entering `6` (SC/DoorDash), `100` (SC/hotel credit) and `5`
(ACCESS/DashPass) on `CHASE-CSP` gives `E3=$106`, `E15=3`, `E16=$111` — i.e. the ACCESS entry registers
as a use but correctly adds no dollars.

**Scope reduction (2026-09-11, user request).** All non-`SC` benefits were deleted — the user only wants
statement credits visible — with one exception: the airline **free checked bag** rows (Citi AAdvantage
Platinum Select, Atmos Rewards Ascent) were kept and converted to `EST` / `MONTHLY` with an editable
`Value/cycle` (default $35) so the bag fees they avoid can be counted. Five cards now carry no trackable
rows at all (Costco Anywhere, Robinhood Gold, Delta Blue, Freedom Unlimited, Customized Cash Rewards) and
their tabs show only identity, months held and the annual fee. The `Ref` catalogue was trimmed to match.
Removed-but-notable, restorable on request: the Citi `$125 AA flight discount`, the CSP `10% Anniversary
Bonus`, the Atmos `20% back` / `$99 Companion Fare`, Aspire's Waldorf/Conrad credit and spend-gated free
nights, Delta Blue's `20% back` and `TakeOff 15`.

**Defect found by this change:** `E7` ("Remaining - statement credits only") was `MAX(0,$E$5-$E$3)`, and
`E5` includes estimated rows — so as soon as a valued `EST` row existed, the SC-only remaining figure was
inflated (CSP read $320 instead of $200). Fixed with a new `H4` helper holding an SC-only maximum
(`SUMIF`/`SUMIFS` over the `F` column) and `E7 = MAX(0,$H$4-$E$3)`. Verified: CSP `H4=$250` → `E7=$200`;
Citi AA `H4=$180` (Turo only, baggage correctly excluded).

**User data re-applied after the rebuild:** the user had entered their `$50` Chase hotel-credit usage on
the previous build; because regenerating the workbook rewrites the file, that entry was re-written at
`CHASE-CSP!N27` (the hotel-credit row, `May-26` column) and re-verified (`E3=$50`, `E8=-$45`, portfolio
`M29=-$900`). Any rebuild resets user entries — close the workbook before regenerating.

Portfolio for the 2026 cycle: **$950 annual fees** charged, **$2,572** max annualized value (SC + baggage
estimates), **$50 used** (CSP hotel credit), **-$900 net (SC)**
(includes estimated values you can edit), **$0 used** until credits are entered, **-$950 net**.

**2026-09-11 completeness audit.** The first build was incomplete — 11 benefits were missing or wrong, and
the user caught it. Added/fixed: Citi AAdvantage `Turo credit` (up to $30/trip, max $180, expires
2026-10-18); Atmos `20% back on in-flight purchases` and `$100 off Alaska Lounge+`; Delta Blue
`TakeOff 15` and `Pay with Miles`; Aspire's `second/third spend-gated Free Night` ($30k/$60k) and
`National Emerald Club Executive status`; CSP `emergency evacuation up to $100,000`, the grandfathered
`10% Anniversary Bonus` (through 2026-10-01), and the `Instacart credit` question; CFU `DashPass 6 months`;
`Museums on Us` on both Bank of America cards; and the Amex Green CLEAR credit corrected from $189 to
**$219** (CLEAR raised its price 2026-07-01 and Amex lifted the credit to match). Max annualized credit
value went from $2,203 to $3,137.

Deliberate blanks, so that nothing is fabricated: CSP `Instacart credit` (sources conflict on whether it
still exists), CSP `10% Anniversary Bonus` (enter the dollar value you expect), Atmos `$99 Companion Fare`,
and the `20% back` rows on Atmos and Delta Blue — all are user-set estimates. `Ref` carries a verification
status per benefit; flagged `PENDING`: the CSP Instacart credit and the Aspire airline credit's scope
(incidentals vs airfare — issuer and aggregator sources disagree).

AF posting month = the card's open month for all 11 cards (assumption; editable in `B10` per card).

---

## 13. Current build inventory (2026-09-11)

> ⚠️ Personal data: this section lists the real cards held by each account. Remove or replace it before
> publishing this repository publicly.

Snapshot of exactly what the generator emits today. Regenerate with `./.venv/bin/python build_tracker.py <profile>`, then update `PROGRESS.md` §3/§5/§9.


### `oscar` → `Oscar_credit_tracker.xlsx` (11 cards)

| Card tab | AF this cycle | Benefit rows (`Basis / Cycle / Value`) |
|---|---|---|
| `CITI-COSTCO` | $0 | _none – no statement credits_ |
| `AMEX-ASPIRE` | $550 | `SC`/`SEMI-ANNUAL`/$200 Hilton Resort Credit<br>`SC`/`QUARTERLY`/$50 Airline Fee Credit<br>`SC`/`CALENDAR-YEAR`/$219 CLEAR+ Credit |
| `CHASE-CSP` | $95 | `SC`/`MONTHLY`/$10 DoorDash non-restaurant credit<br>`SC`/`CARDMEMBER-YEAR`/$100 Chase Travel hotel credit<br>`SC`/`MULTI-YEAR`/$120 Global Entry / TSA PreCheck / NEXUS credit |
| `BOA-ATMOS-ASCENT` | $95 | `EST`/`CALENDAR-YEAR`/$350 Free checked bag |
| `RH-GOLD` | $60 | _none – no statement credits_ |
| `AMEX-DELTA-BLUE` | $0 | _none – no statement credits_ |
| `CITI-AA-PLAT` | $99 (1st-yr waived) | `EST`/`CALENDAR-YEAR`/$350 Free checked bag<br>`SC`/`ONE-TIME`/$180 Turo credit |
| `CHASE-CFU` | $0 | _none – no statement credits_ |
| `AMEX-GREEN` | $150 | `SC`/`CALENDAR-YEAR`/$219 CLEAR+ Credit |
| `BOA-CUSTOMIZED-CASH` | $0 | _none – no statement credits_ |
| `AMEX-BCE` | $0 | `SC`/`MONTHLY`/$7 Disney Streaming Credit<br>`SC`/`MONTHLY`/$15 Home Chef credit |

`Benefits` tab rows: **12** · sheets: 15

### `myra` → `Myra_credit_tracker.xlsx` (3 cards)

| Card tab | AF this cycle | Benefit rows (`Basis / Cycle / Value`) |
|---|---|---|
| `CHASE-UNITED-QUEST` | $350 | `SC`/`MONTHLY`/$8 Rideshare credit<br>`SC`/`MONTHLY`/$15 Instacart credits<br>`SC`/`CARDMEMBER-YEAR`/$200 United TravelBank credit<br>`SC`/`CARDMEMBER-YEAR`/$150 Renowned Hotels and Resorts credit<br>`SC`/`CARDMEMBER-YEAR`/$150 JSX credit<br>`SC`/`CARDMEMBER-YEAR`/$80 Avis/Budget TravelBank cash<br>`EST`/`CALENDAR-YEAR`/$220 Free checked bag<br>`SC`/`MULTI-YEAR`/$120 Global Entry / TSA PreCheck / NEXUS credit |
| `CHASE-BOUNDLESS` | $95 | `SC`/`SEMI-ANNUAL`/$50 Airline credit<br>`EST`/`QUARTERLY`/$10 DoorDash non-restaurant discount<br>`EST`/`CARDMEMBER-YEAR`/user-set Annual Free Night Award |
| `CHASE-CSP` | $95 | `SC`/`MONTHLY`/$10 DoorDash non-restaurant credit<br>`SC`/`CARDMEMBER-YEAR`/$100 Chase Travel hotel credit<br>`SC`/`MULTI-YEAR`/$120 Global Entry / TSA PreCheck / NEXUS credit |

`Benefits` tab rows: **14** · sheets: 7

**Totals:** oscar — AF $950, max annualized $2,432, · myra — AF $540, max annualized $1,496.

Every row's `Z` column (card tabs) and `Note / source` column (`Benefits` tab) carries its issuer citation; rows that could not be confirmed on an issuer page say `PENDING` there.
