# Swerdlow Asset Management & Underwriting Platform

The reshaped workbook built from your `Example.xlsx` and the
`Management_Platform Doc`. Every input/assumption lives on **one tab**, with the
blue **[INPUT]** block sitting beside the green **[Current]** block.

## Files

| File | What it is |
|------|------------|
| `Swerdlow_Management_Platform.xlsx` | The workbook (9 tabs). |
| `vba_platform/Module1_Publish.bas` | `Publish_Project` + version control (Current/Historical) + FIFO archive. |
| `vba_platform/Module2_Projects.bas` | `Add_Project`, `Validate_Inputs`. |
| `vba_platform/Module3_Support.bas` | `Request_Support` (Calendly). |
| `build_console_model.py` | The generator (one source of truth for the workbook + `.bas`). |

## The loop

1. Pick a **Project** from the single selector at the top of the **Cockpit** (Management Console).
2. The **green "Current"** column (cols G–J) shows the latest assumptions published to the database.
3. Edit the **blue input** fields (cols B–E).
4. Click **Publish** → it validates, writes the blue values to the database as a
   new **Current** record (the project's prior Current row → **Historical**), and
   the green values refresh.

The database is the single source of truth; the console only edits.

## Layout (matches your example)

- Blue input values in cols **B–E**, green current mirror in cols **G–J**.
- Scalars: input in **D**, current in **I**. Tables (unit mix): **C/D/E** ↔ **H/I/J**.
- Sections: General Inputs · Key Dates · Commercial Assumptions · Multifamily Unit
  Mix & Rental Revenue (GPR→Effective Rental Revenue) · Other Income · EGR ·
  Operating Expenses · NOI / OpEx Ratio · Development Budget.
- Both sides compute in parallel (`E33=Blended_PUPM*Units_Total*12` on input,
  `J33=Blended_PUPM_Cur*Units_Total_Cur*12` on current), exactly like your
  `E33 / J33` pattern.

## Single Project selector (per your instruction)

There is one selector and it is **Project** — the phase toggle was removed.
**Phase No.** remains a stored input field. The relational database is keyed by
Project; "Current" = the latest published record for the selected Project.
Dashboards roll up by Project. (If phases later need to be independently
selectable records, the DB key can be switched to Project + Phase without
reworking the console.)

## Tabs

`Read me · Cockpit (console) · Master Assumptions · Database · Portfolio Dashboard
· Project Dashboard · Monthly CF` + hidden `Lists`, `Archive`.

## Enabling the macros (`.xlsx` → `.xlsm`)

1. Open in Excel → **Save As → Excel Macro-Enabled Workbook (`.xlsm`)**.
2. **Alt+F11** → **File ▸ Import File…** → import all three `.bas` from `vba_platform/`.
3. Assign `Publish_Project` to the red **PUBLISH ASSUMPTIONS** button, and
   `Add_Project` / `Request_Support` to their buttons.
4. Set `CALENDLY_URL` in `Module3_Support`.

## Version control (per the platform doc)

Each Publish records User, Timestamp, Notes, and Status. Latest = **Current**;
prior records → **Historical**. A FIFO cap (`MaxRecordsPerProject`, default 20)
archives the oldest record per project to the hidden **Archive** tab for
auditability.

## Validation performed

- 564 formulas audited: **no undefined names, no error literals**; divisions
  `IFERROR`-guarded; real date serials.
- The input-side proforma was evaluated numerically and ties out exactly
  (blended $1,375 PUPM → GPR $66,000 → EGR $67,500 → NOI $31,235, 54% OpEx
  ratio, $59M dev cost). The green/current side runs the same formulas over
  `INDEX` lookups (separately verified).
- Couldn't do one full-file recalc in real Excel (no working Excel/LibreOffice in
  the build sandbox); the genuine final smoke-test — open, recalc, import macros —
  is yours to run.

## Rebuilding

```bash
pip install openpyxl
python3 build_console_model.py
```
