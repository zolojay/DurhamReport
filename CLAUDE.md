# SCR Catalyst Test Report Workbook

## What This Is
A professional Excel/VBA workbook (.xlsm) for SCR catalyst testing at Cormetech's Durham lab. Replaces a legacy Excel workbook (v56) with a clean, form-driven, auditable system.

## Authoritative Documents
- `PRD.md` — Single architectural source of truth. All design decisions are final.
- `SCR_Calculations_Reference.md` — All engineering formulas, K-value logic, H₂O correction chains, dry correction, injection rates, steady-state checks. Every formula in the workbook must match this reference exactly.

## Build Approach
- Python (openpyxl) scripts generate the workbook shell: sheets, Excel Tables, named ranges, formatting, conditional formatting, data validation
- VBA modules exported as .bas files for: navigation, state management, form logic, audit metadata, protection toggling, reset routines
- UserForm code-behind exported as .frm text files
- UserForm visual layout must be designed manually in Excel VBA editor after import

## Build Phases (execute in order)
1. **Workbook shell** — sheets, visibility states, table shells, named ranges, constants, lists, control structure
2. **Geometry engine** — geometry input sheets, Geometry Calc formulas, resolved named outputs (GC_ prefix)
3. **Home and Setup Summary** — setup workflow, status block, review outputs, signoff display
4. **Hidden tables and forms** — normalized backing tables, frmNOx, frmNH3, frmSO2, frmSO3, frmPhysical, frmDP
5. **Verification and browser** — frmSetupVerify, frmRecordBrowser, row-specific pairing
6. **Dashboards** — Activity, Conversion, DP with summary formulas and steady-state sections
7. **Protection and testing** — protection logic, reset workflow, validation

## Key Architecture Rules
- **Geometry Calc is the sole calculation engine.** Downstream sheets consume resolved outputs only (GC_ named ranges). No sheet re-derives geometry independently.
- **Source-resolution happens once.** Geometry Calc resolves flow, AV, injection rates. Activity/Conversion reference those outputs.
- **Backing tables stay VeryHidden always.** No workflow state ever unhides them.
- **All data entry goes through UserForms.** No direct typing into hidden tables.
- **Audit metadata auto-populates via VBA.** RecordID, TestID, timestamps, technician — all written by form save routines.
- **H2O_Reference = 18%** on Constants sheet. Part of the H₂O fallback chain. Not technician-editable.
- **Dual setup verification required** before Begin Test enables.

## 3-State Workflow Model
- **State 1 (Setup):** Home + Specifications visible. Technician enters identity, conditions, geometry.
- **State 2 (Review):** Setup Summary unhides. Geometry resolved. Two-person verification required.
- **State 3 (Testing):** Activity, Conversion, DP unhide. Form-driven data entry begins.
- **Reset:** Clears all data, returns to State 1 with confirmation.

## Color Palette
- Section Header: #0E1638 (white text)
- Sub-Section: #2C3A8C (white text)
- Column Header: #4C5CE0 (white text)
- Input Cell: #FFF9C4
- Formula Output: #E8E8E8
- Label: #E8EAF6
- Pass: #C8E6C9
- Warning: #FFE0B2
- Fail: #FFCDD2
- Font: Calibri throughout (16/14/12/10-11 pt hierarchy)

## Named Range Conventions
- `GC_` prefix for all Geometry Calc resolved outputs
- `Ctrl_` prefix for all Control sheet state fields
- `Spec_` prefix for Specifications thresholds
- Table names: `tbl_NOx`, `tbl_NH3`, `tbl_SO2`, `tbl_SO3`, `tbl_Temperature`, `tbl_Flow`, `tbl_O2`, `tbl_DP`

## UserForm Inventory
- frmNOx, frmNH3, frmSO2, frmSO3, frmPhysical, frmDP — data entry
- frmRecordBrowser — record review and pass-row pairing (replaces raw sheet access)
- frmSetupVerify — dual verification workflow

## What NOT to Build
- No Dev Tools sheet
- No Calculators sheet
- No Formula Reference tab in the production workbook
- No backup/debug/temporary sheets
- No duplicate sheets

## Flow Rule
- AV entered + UGS blank → derive UGS
- UGS entered + AV blank → derive AV
- Both entered → warn, use AV as authoritative
- Neither entered → withhold flow/injection calculations

## File Structure
```
scr-workbook/
├── CLAUDE.md                      # This file
├── PRD.md                         # Architecture spec
├── SCR_Calculations_Reference.md  # Engineering formulas
├── build/
│   └── build_workbook.py          # openpyxl build script
├── vba/
│   ├── modNavigation.bas          # Sheet navigation & visibility
│   ├── modState.bas               # 3-state workflow management
│   ├── modAudit.bas               # Audit metadata helpers
│   ├── modProtection.bas          # Protection toggle routines
│   ├── modReset.bas               # New Test / Reset logic
│   ├── modHelpers.bas             # Shared utility functions
│   ├── frmNOx.frm                 # NOx entry form
│   ├── frmNH3.frm                 # NH3 entry form
│   ├── frmSO2.frm                 # SO2 entry form
│   ├── frmSO3.frm                 # SO3 entry form
│   ├── frmPhysical.frm            # Physical params entry form
│   ├── frmDP.frm                  # DP entry form
│   ├── frmRecordBrowser.frm       # Record browser / pair selector
│   └── frmSetupVerify.frm         # Setup verification form
└── output/
    └── SCR_TestReport.xlsm        # Built workbook
```

## Testing Notes
- Claude Code cannot open Excel or test runtime behavior
- After each phase, the developer imports VBA into Excel and tests manually
- UserForm visual layout (control placement, sizing) must be done in VBA editor
- Claude Code generates the code-behind; the developer designs the form layout
