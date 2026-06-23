# sheets-excel-master

A Claude skill that turns Claude into a senior spreadsheet architect — expert in Microsoft Excel, Google Sheets, and VBA macro development.

Built for developers, government encoders, business analysts, and anyone who needs production-ready spreadsheet output without re-prompting.

---

## What It Does

- **Data Entry Form Design** — layouts, validation, protection, color coding, named ranges
- **Formula Engineering** — Excel 2016 through 365, Google Sheets, XLOOKUP, INDEX/MATCH, QUERY, ARRAYFORMULA, LAMBDA
- **VBA Macro Development** — full macros with error handling, dynamic row detection, no hard-coded paths
- **Dashboard & Reporting** — 3-layer sheet structure, PivotTables, KPI cards, chart specs
- **Data Validation & Protection** — dropdowns, dependent lists, sheet locks, protected ranges
- **Google Sheets Specific** — Apps Script, IMPORTRANGE, QUERY, named ranges, locale-aware formulas

---

## Installation

1. Download or clone this skill folder
2. Go to **claude.ai → Sidebar → Customize → Skills → Upload a Skill**
3. Upload the `sheets-excel-master` folder

---

## Usage

Invoke naturally in Claude:

```
Build me a data entry form for employee records in Excel
```

```
Write a VBA macro that loops through all rows and highlights duplicates in column B
```

```
Google Sheets formula to sum sales where region = "NCR" and status = "Paid"
```

```
Design a monthly dashboard with KPI cards and a bar chart — source data is in Sheet1
```

```
Create a dependent dropdown in Excel — Region → Province
```

Or invoke explicitly:

```
/sheets-excel-master

Build me a data entry format for tracking civil service eligibilities — fields: name, position, eligibility type, date granted, expiry date, remarks
```

---

## Output Format

Every output is labeled `[Excel]`, `[Google Sheets]`, or `[VBA]`. Formulas come with a plain-English explanation. VBA comes with inline comments and a header block. Layouts come as tables, not prose.

---

## Notes

- Locale-aware: will flag comma vs semicolon separator issues for regional Excel installs (PH, EU)
- Version-aware: flags when a function requires Excel 2019+, 2021, or 365
- Platform-safe: never mixes Excel and Sheets syntax without labeling
