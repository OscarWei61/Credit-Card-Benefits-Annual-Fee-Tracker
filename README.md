# Credit Card Benefits & Annual Fee Tracker

A generator that turns a short Python list of your credit cards into a fully formula-driven Excel
workbook: a portfolio dashboard, one tab per card, and a checkbox-style month grid for every recurring
statement credit so you can see — at a glance — **which credits expire soon and whether each card is
actually paying for itself.**

```python
dict(tab="CHASE-CSP", issuer="Chase", product="Chase Sapphire Preferred",
     opened="2025-03-27", af=95, waived=False, au_af=0, status="Active",
     cal=[("DoorDash non-restaurant credit", "SC", "MONTHLY", 10, "needs DashPass activated")],
     cmy=[("Chase Travel hotel credit", "SC", "CARDMEMBER-YEAR", 100, "prepaid via Chase Travel")],
     multi=[("Global Entry / TSA PreCheck credit", "SC", "MULTI-YEAR", 120, 4, "every 4 years")])
```

Run one command and you get a workbook with 100+ conditional-formatting rules, data validation,
month-by-month credit tracking and a portfolio rollup — no macros, no add-ins, no cloud service.

---

## Why this exists

Modern US credit cards pay for themselves through a pile of small statement credits — $10/month here,
$50/quarter there, $100 per card anniversary somewhere else. They reset on **four different schedules**
(monthly, quarterly, semi-annual, calendar year, and cardmember-anniversary), they silently expire, and
a spreadsheet that tracks them properly is genuinely annoying to build by hand:

- a credit that resets on your **card anniversary** can't live in a Jan–Dec grid,
- a `$200` hotel credit is rarely used for exactly `$200`, so a checkbox loses information,
- and once you have ten cards, "am I up or down on annual fees?" is impossible to answer in your head.

This project solves those three problems and nothing else.

---

## What you get

| Tab | What it is |
|---|---|
| `Dashboard` | Page 1. Portfolio KPIs + one row per card (annual fee, credits used, credits remaining, net benefit, months held, credits expiring soon) |
| `Benefits` | A flat, filterable index of **every benefit on every card** — one row each, live-linked to the card tabs |
| One tab per card | Identity block, 14 KPIs, and benefit rows grouped into coloured **sections by basis** |
| `TEMPLATE` | A blank card tab to copy when you add a card by hand |
| `Ref` | The tracking year (one cell rolls the whole workbook to next year), dropdown lists, month names, and a benefit catalogue |

### Features

- **Amount entry, not checkboxes.** Type the dollar amount you actually used into that month's cell.
  Green = counted, orange = you typed more than the credit is worth (it is capped automatically), red =
  non-numeric entry.
- **Two levels of capping.** Each cell is capped at the credit's per-cycle value, and each row is capped
  at its annual maximum, so the tracker can never overstate a benefit.
- **Correct reset calendars.** Credits that reset on your card anniversary get their own section whose
  month columns run from *your* anniversary month, not January.
- **Two net-benefit numbers.** A conservative one that counts only real statement credits, and an
  "including estimates" one for things like free checked bags whose value you set yourself.
- **"Expiring soon" detection.** Any unused credit whose cycle ends within 30 days is counted and
  highlighted, so you stop forfeiting the quarterly ones.
- **Input diagnostics.** Two KPI cells tell you how many month cells you have filled in and the raw total
  you typed — so if a number doesn't move the totals you can tell instantly whether the entry registered.
- **Rebuilds don't destroy your data.** The generator reads your existing workbook first, harvests what
  you typed, backs the old file up, and writes your values back into the freshly built one.
- **No macros.** Plain formulas, data validation and conditional formatting. Works in Excel and survives
  opening in LibreOffice / Google Sheets.

---

## Quick start

### 1. Requirements

- Python 3.9+ (developed on 3.9.6)
- `openpyxl` (3.1.x)

```bash
git clone <this repo>
cd credit_card_tracker
python3 -m venv .venv
./.venv/bin/pip install openpyxl
```

### 2. Describe your cards

Everything about your wallet lives in **one file**: `build_tracker.py`. Add your cards as a list of
dicts and register a profile that names your output file.

At the top of the file you'll find example card lists. Replace them with your own (see
[Card data reference](#card-data-reference) for every field), then edit `PROFILES`:

```python
PROFILES = {
    "me": ("My_credit_tracker.xlsx", MY_CARDS),
}
```

A minimal starting list looks like this:

```python
MY_CARDS = [
    dict(tab="CHASE-CSP", issuer="Chase", product="Chase Sapphire Preferred",
         opened="2025-03-27",          # YYYY-MM-DD, the approval date
         af=95,                        # annual fee in USD, per year
         waived=False,                 # True if the first year's fee is waived
         au_af=0,                      # total authorized-user fees, if any
         status="Active",
         af_note="",                   # shows up in the notes column
         cal=[                         # credits that reset on Jan 1 / monthly / quarterly
            ("DoorDash non-restaurant credit", "SC", "MONTHLY", 10, "needs DashPass activated"),
         ],
         cmy=[                         # credits that reset on YOUR card anniversary
            ("Chase Travel hotel credit", "SC", "CARDMEMBER-YEAR", 100, "prepaid via Chase Travel"),
         ],
         multi=[                       # one-time and every-N-years credits
            ("Global Entry / TSA PreCheck credit", "SC", "MULTI-YEAR", 120, 4, "every 4 years"),
         ]),
    # ... one dict per card
]
```

> **Tip:** don't want to track a card's benefits? Give it empty lists — the tab still shows months held
> and the annual fee, which is what feeds the portfolio rollup.

### 3. Generate

```bash
./.venv/bin/python build_tracker.py me        # -> My_credit_tracker.xlsx
```

Exact cell addresses, formulas and the caps described above are all produced at build time; nothing is
hard-coded to the sample cards.

### 4. Open it

Open the `.xlsx` in Excel. It recalculates on open (the workbook is flagged `fullCalcOnLoad`), because
files written by `openpyxl` contain no cached results. If you ever see blank formula cells, press `F9`.

---

## The daily loop

1. When you use a credit, type the amount into the month cell of that row (columns `L`–`W`).
2. The card tab updates `Used`, `Remaining`, and the KPI block.
3. The `Dashboard` updates the portfolio totals and flags anything expiring within 30 days.
4. Anything you don't use just goes unused — that's the point, you want to *see* it.

**Colour legend on the month grid**

| Colour | Meaning |
|---|---|
| 🟩 Green fill | An amount is counted |
| 🟧 Orange fill | You entered more than that cycle's credit — the value is capped to the credit |
| 🟥 Red fill | Non-numeric entry (the data validation normally blocks this) |
| Red text | Two uses inside the same quarter / half-year, which capping cannot detect |
| Grey text | An unfilled slot |

**Where the numbers live** — every card tab exposes the same 14 KPIs, and the Dashboard reads only these:

| Cell | KPI |
|---|---|
| `E3` | Credits used YTD — statement credits |
| `E4` | Credits used YTD — including your estimates |
| `E5` | Maximum annual value (per-card total) |
| `E6` / `E7` | Remaining — including estimates / statement credits only |
| `E8` / `E9` | **Net benefit** — statement credits / including estimates |
| `E10` | Result, `POSITIVE` or `NEGATIVE` |
| `E11` | How much more you need to use to break even on the annual fee |
| `E12` | Credit utilization % |
| `E13` / `E14` | Credits at risk (≤30 days left) / nearest expiry in days |
| `E15` / `E16` | Amounts you entered (cell count) / entered total before capping |

---

## Card data reference

### Card fields

| Field | Type | Notes |
|---|---|---|
| `tab` | str | Sheet name. Uppercase `ISSUER-CODE`, ≤31 chars, avoid `[ ] : * ? / \`. The Dashboard links to it by name |
| `issuer` | str | Shown on the Dashboard; the issuer dropdown list lives on `Ref` |
| `product` | str | Exact product name |
| `opened` | `"YYYY-MM-DD"` | Approval date. Drives **months held** and the anniversary calendar |
| `af` | number | Annual fee per year |
| `waived` | bool | `True` if the first year's fee is waived — the workbook then charges `$0` this cycle |
| `au_af` | number | Total authorized-user fees, added to the annual fee |
| `status` | str | `Active`, `Closed`, `Product-changed`, or `Authorized-user-only`. **Only `Active` cards count toward portfolio totals** |
| `af_note` | str | Free text, appears in the notes column |
| `cal` / `cmy` / `multi` | list | The benefit rows — see below |

### Which list does a benefit go in?

| List | Use it for | Month columns show |
|---|---|---|
| `cal` | Credits that reset on the **calendar**: `MONTHLY`, `QUARTERLY`, `SEMI-ANNUAL`, `CALENDAR-YEAR` | Jan–Dec of the tracking year |
| `cmy` | Credits that reset on **your card anniversary** (`CARDMEMBER-YEAR`) | Your anniversary month onwards |
| `multi` | `MULTI-YEAR` (e.g. Global Entry every 4 years) and `ONE-TIME` items | No month grid — you enter the amount captured |

### Benefit row shape

```python
# cal / cmy  ->  (name, basis, cycle, value_per_cycle, note)
("Hilton Resort Credit", "SC", "SEMI-ANNUAL", 200, "$200 Jan-Jun and $200 Jul-Dec")

# multi      ->  (name, basis, cycle, value, years_per_cycle, note)
("Global Entry / TSA PreCheck credit", "SC", "MULTI-YEAR", 120, 4, "every 4 years")
```

Use `None` for `value_per_cycle` if you deliberately want it blank (the row then tracks usage but no
money) — that is how the reference card lists mark values they can't verify.

### Basis — what the row counts toward

| Basis | Meaning | Lands in |
|---|---|---|
| `SC` | A real statement credit | Both the conservative and the "including estimates" totals |
| `EST` | Value you set yourself (free checked bag, free night, in-flight rebate) | Only the "including estimates" totals |
| `ACCESS` | A perk with no dollar value (lounge access, status, insurance) | Neither — the row only logs how often you used it |

### Cycle — how often it resets

| Cycle | Window | Notes |
|---|---|---|
| `MONTHLY` | Calendar month | 12 cells per year |
| `QUARTERLY` | Calendar quarters | Jan/Apr/Jul/Oct |
| `SEMI-ANNUAL` | Jan–Jun and Jul–Dec | *Not* "six months from signup" |
| `CALENDAR-YEAR` | Jan 1 – Dec 31 | |
| `CARDMEMBER-YEAR` | Your anniversary → +12 months − 1 day | Goes in `cmy`; the anchor is the AF posting month |
| `MULTI-YEAR` | Every N years | Goes in `multi`; the row is annualised as `value ÷ years` |
| `ONE-TIME` | Once | Goes in `multi` |

### A complete worked example

```python
dict(tab="AMEX-ASPIRE", issuer="Amex", product="Hilton Honors Aspire Card",
     opened="2026-03-29", af=550, waived=False, au_af=0, status="Active", af_note="",
     cal=[
        ("Hilton Resort Credit", "SC", "SEMI-ANNUAL", 200, "$200 each half; participating resorts only"),
        ("Airline Fee Credit",   "SC", "QUARTERLY",    50, "$50 per quarter"),
        ("CLEAR+ Credit",        "SC", "CALENDAR-YEAR",219, "auto-renewing CLEAR+ membership"),
        ("Free checked bag",     "EST","CALENDAR-YEAR",350, "set your own expected saving"),
     ],
     cmy=[
        ("Anniversary Free Night Reward", "EST", "CARDMEMBER-YEAR", 300, "on card renewal"),
     ],
     multi=[]),
```

---

## How the maths works

```
you type an amount in a month cell
  → the row's raw total        (each cell capped at Value/cycle)
  → the row's Used             (whole row capped at Max annual)
  → the card's E3 / E4                                                        (credits used)
  → the card's E8 / E9  =  credits used − annual fee charged                  (net benefit)
  → the Dashboard's per-card row (J/K/M/N)
  → the Dashboard totals row and the portfolio KPIs
```

- **`Max annual`** for a calendar-cycle row is `Value/cycle × cycles per year`; for an anniversary or
  one-time row it is just the value.
- **Net benefit** is `credits used − annual fee actually charged this cycle` (first-year waivers handled).
- **Multi-year credits are annualised** (`$120 ÷ 4 = $30/yr`) so a single Global Entry reimbursement
  doesn't create a fake $120 year.
- The headline `POSITIVE` / `NEGATIVE` verdict uses the **conservative** number, so you can't talk
  yourself into a card with aspirational valuations.

---

## Rebuilding safely

`build_tracker.py` regenerates the whole workbook, so it protects your data on every run:

1. it copies the current output to `<name>.backup.xlsx`,
2. it harvests everything you typed — month amounts, any `Value/cycle` you changed, the `Used?` column of
   one-time rows, the identity block, and the tracking year —
3. it rebuilds, then writes those values back, matching rows **by benefit name** so values survive the
   layout moving around.

To distinguish "you changed this value" from "this is still my default", the last-written defaults are
stored in `.tracker_build_state_<profile>.json`.

**Two things are not preserved**, because the generator recreates the structure:

- renaming a benefit row (the name is the matching key), and
- deleting a row by hand (it comes back — remove it from the card data instead).

**Close the workbook before rebuilding.** Excel leaves a `~$<name>.xlsx` lock file while a file is open,
and your unsaved edits live in Excel's memory, not on disk.

---

## Year rollover

1. Save the old file as an archive, or let the backup be your archive.
2. Change **`Ref!B1`** to the new year — every month label, cycle-end date and calendar-year window
   follows that one cell.
3. Clear the month cells on each card tab (the generator's preservation will otherwise carry them over).
4. Update each card's annual fee (`B8`) and AF posting date (`B10`) if they changed.
5. Re-verify the credit amounts — issuers change them almost every year.

---

## Multiple people

Each profile keeps its own output file, its own backup and its own state file, so rebuilding one never
touches the other:

```python
PROFILES = {
    "alex":  ("Alex_credit_tracker.xlsx",  ALEX_CARDS),
    "sam":   ("Sam_credit_tracker.xlsx",   SAM_CARDS),
}
```

```bash
./.venv/bin/python build_tracker.py alex
./.venv/bin/python build_tracker.py sam
```

---

## Repo layout

| File | What it is |
|---|---|
| `build_tracker.py` | The generator. Your card data, the formulas, the layout — everything editable lives here |
| `PLAN.md` | Design spec: sheet architecture, the full formula map, cell contracts, verification notes, and a current inventory of what the generator emits |
| `PROGRESS.md` | Project status: what's built, open action items, per-benefit source verification, adopted rules, change log |
| `My_credit_tracker.xlsx` | (generated) your workbook |
| `My_credit_tracker.backup.xlsx` | (generated) automatic pre-rebuild backup |
| `.tracker_build_state_<profile>.json` | (generated) remembers the last-written defaults |

---

## Design notes and gotchas

These are the non-obvious things that cost real debugging time — worth reading before you extend it.

**Never put `IF` inside an aggregate.** `SUMPRODUCT(IF(ISNUMBER(range), ...))` evaluates to `0` in some
Excel versions, because `IF` is not evaluated in array context without `Ctrl+Shift+Enter`. That silently
produced `$0` for every "used" figure. The generator only uses plain `SUM` / `SUMIF` / `COUNT` /
`COUNTIF`, or the multiplication idiom `SUMPRODUCT(ISNUMBER(r)*(r>0))`, which behaves identically
everywhere.

**A bare reference to an empty cell returns `0`, not `""`.** `F45 = =$D45` on an unused row made
`MAX(0, F45 - G45)` compute `0 - ""`, i.e. `#VALUE!` on every empty slot. Guard with
`=IF($D45="","",$D45)`.

**Compare amounts in a numeric-only helper column.** `COUNTIF(range,">0")` counts *text* cells, because
Excel sorts any text above any number. At-risk detection therefore reads a hidden helper column that is
numeric by construction (`days left`, or `9999` when nothing remains) instead of testing text-bearing
columns directly.

**`TEXT(date,"mmm")` follows your Excel locale.** On a Chinese-language install it renders `1月`, not
`Jan`. Month names come from a typed list on `Ref` and are fetched with `INDEX`, so labels stay English
everywhere.

**Anchor anniversary cycles on the *current* cycle start, not the card's original open date.** Otherwise a
card whose anniversary has already passed this year shows a cycle that ended in the past, and its credits
drop out of the "expiring soon" logic entirely.

**Verify credit amounts against the issuer's own site.** Aggregators are useful for discovery but go
stale, and limited-time cardmember offers often appear *only* on the issuer's page. Benefits that can't be
confirmed on an issuer page are marked `PENDING` in the notes column rather than guessed at. Some issuer
pages are JavaScript-rendered and cannot be read by simple fetchers at all; those rows say so.

**Keep the notes column.** Every benefit row carries a `Notes / source` cell recording where its amount
came from. When a figure changes next year, that column is what tells you what to re-check.

---

## Disclaimer

This is a personal bookkeeping tool, not financial advice. Credit card terms change frequently — verify
every amount, reset schedule and eligibility rule against the issuer's own current terms before acting on
anything the spreadsheet tells you. Benefit values you enter yourself (`EST` rows) are your own estimates.

Your workbook contains your card list, annual fees and usage. It is excluded by `.gitignore` — **do not
commit it** to a public repository.
