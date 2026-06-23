# Swerdlow Multi-Phase Cockpit Model

Scalable input template / "cockpit" for a large multi-phase real estate
development program, built to `Cockpit_Model_Build_Spec.txt`.

This pass builds the **input + assumptions architecture and a single stabilized
(untrended) proforma summary**. It deliberately does **not** build a monthly
cash-flow engine, financing/refinance, or an equity waterfall — those are out of
scope and the structure is laid out so they can be added later without rework.

## Deliverables

| File | What it is |
|------|------------|
| `Swerdlow_Cockpit_Model.xlsx` | The workbook: 9 tabs, formulas, named ranges, dropdowns, conditional formatting, row grouping, sheet protection. |
| `vba/Module1_Update.bas` | `UPDATE_Phase` + FIFO versioning cap (auto-generated field map). |
| `vba/Module2_Assets.bas` | `Add_Asset`, `Validate_Inputs`. |
| `vba/Module3_Support.bas` | `Request_Support` (Calendly escalation). |
| `build_cockpit_model.py` | The generator. Re-run it to rebuild the workbook **and** the `.bas` files from one source of truth. |

> **Why `.xlsx` + `.bas` instead of `.xlsm`?** Pure-Python tooling cannot author
> Excel's compiled VBA project (`vbaProject.bin`). The spec explicitly sanctions
> this fallback. The 60-second conversion to a macro-enabled `.xlsm` is below.

## Enabling the macros (`.xlsx` → `.xlsm`)

1. Open `Swerdlow_Cockpit_Model.xlsx` in Excel → **File ▸ Save As ▸ Excel
   Macro-Enabled Workbook (`.xlsm`)**.
2. **Alt+F11** (VBA editor) → **File ▸ Import File…** → import all three `.bas`
   modules from `vba/`.
3. Back in Excel, assign the macros to the three placeholder buttons on the
   Cockpit control bar (right-click ▸ Assign Macro, or Developer ▸ Insert ▸
   Button): `UPDATE_Phase`, `Add_Asset`, `Request_Support`.
4. In `Module3_Support` set `CALENDLY_URL` to your real scheduling link.

## Tabs (left → right)

`Read me · Cockpit · Master Assumptions · Database · All Phases · Phase Detail ·
Monthly CF` plus hidden support tabs `Lists` and `Archive`.

**Operating loop:** the user edits one Phase on the **Cockpit** → clicks
**UPDATE** → a flat snapshot row is appended to the **Database** (system of
record). *All Phases*, *Phase Detail*, and *Monthly CF* read the latest record
per Phase from the Database.

## Conventions

- **Blue font + light-blue fill** = `[INPUT]` (you type here). **Black** =
  `[CALC]` formula. **Green** = `[LINK]` cross-tab reference.
- Number formats: currency `$#,##0;($#,##0);"-"`, percent `0.0%`, PSF/PUPM
  `$#,##0.00`, counts `#,##0`; negatives in parentheses, zeros show as `-`.
- Named ranges for every input block / master driver; dropdowns sourced from
  `Lists`; collapsible row groups per Cockpit section; all cells locked except
  Cockpit inputs.

## Key decisions baked in (spec §7)

Two master rent inputs (AMI drives LIHTC/Live Local; FMR-by-bedroom drives
PBV/RAD), market rents entered directly · untrended proforma only · Database =
flat snapshot ledger, FIFO-capped at 20 (configurable via `MaxRecordsPerPhase`),
purged rows archived · development budget **in** scope, financing/waterfall
**out** · no RE-tax abatement line · 15-row caps on unit mix and commercial
rent roll.

## Validation performed

- Opens with **zero formula errors**: 1,153 formulas across all tabs audited —
  no undefined names (`#NAME?`/`#REF!`), no error literals; every variable-
  denominator division is `IFERROR`-guarded; dates are real serials so `EDATE`
  cannot throw `#VALUE!`.
- The **Rent Basis Engine** (spec §4.1b) was evaluated numerically and matches
  hand calcs exactly (LIHTC 1BR = $1,127.50/mo; PBV 2BR = $1,885.00/mo).
- Full calculation is set to run on load.

## Intentional placeholders (labeled, not silently faked)

Cells showing `-` on *All Phases* / *Phase Detail* are program-level $ roll-ups
that recompute from each phase's stored raw inputs; those recompute formulas are
wired alongside the monthly CF engine in the next pass. The `Monthly CF` tab
contains only the date-header spine this pass, as specified.

## Rebuilding

```bash
pip install openpyxl
python3 build_cockpit_model.py
```
