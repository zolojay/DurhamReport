# Architecture Redesign PRD — Complete Build Specification

**Revision 5 — March 2026**

---

## 1. Vision & Design Principles

Build a professional Excel-based SCR catalyst test report workbook for daily lab use. The workbook must preserve the required engineering calculations and reporting workflow while giving technicians a clean, controlled interface. Data entry must be form-driven, record storage must be table-backed and hidden, and all core calculation logic must be centralized.

### Core Principles

- **Clean technician UX** — Visible sheets expose only intended workflow fields. No raw storage, no scratch zones, no engine logic on technician-facing sheets.
- **Centralized calculation logic** — Geometry Calc is the sole engine. Source-resolution happens once. Downstream formulas consume resolved outputs and never re-interpret user choices independently.
- **Normalized hidden storage** — Each measurement type is stored in its own hidden Excel Table. All records carry audit metadata written automatically by VBA.
- **Parallel dashboard design** — Activity and Conversion use the same overall layout pattern so operators do not have to learn two unrelated page designs.
- **Controlled review and pairing** — Technicians must be able to see all candidate records and choose official pairings without ever browsing raw backing sheets.
- **Protection and traceability** — Visible sheets are protected except for intended input cells. Hidden sheets stay VeryHidden. Setup and geometry require dual verification before testing begins.
- **No production clutter** — No duplicate sheets, no backup tabs, no debug sheets, and no "temporary" helper structures left in the production workbook.

---

## 2. Final Architecture Decisions

These decisions are now fixed for the build.

- Backing data sheets remain VeryHidden throughout normal operation.
- Begin Test does not unhide backing data sheets.
- View All does not navigate to raw data sheets.
- View All becomes a record browser / pair selector UserForm.
- Home is split from a dedicated Setup Summary sheet.
- Setup requires two-person verification before testing begins.
- Activity and Conversion both display setup verification initials at the top.
- H2O_Reference is stored on Constants and defaults to 18%.
- Activity and Conversion both unlock on Begin Test.
- Dev Tools is excluded from the production workbook.
- Calculators is excluded from the production workbook unless later requested as a separate utility workbook.
- Formula Reference should not live in the production workbook; maintain it externally or in a separate engineering copy.

---

## 3. Workbook Workflow (3-State Model)

### State 1 — Setup

Visible sheets on open:
- Home
- Specifications

Technician enters:
- identity fields
- gas start times
- test conditions
- sample type

Technician clicks **Set Up Geometry**:
- unhides the matching geometry sheet only
- keeps other geometry sheets hidden
- writes active geometry type to Control
- navigates to the geometry sheet

### State 2 — Setup Review

After geometry entry, technician returns to Home.

Workbook now:
- resolves geometry outputs in Geometry Calc
- populates compact readiness/status items on Home
- unhides Setup Summary
- allows setup verification workflow

At this stage:
- setup can be reviewed
- geometry/flow/injection outputs can be reviewed
- two-person verification must be completed
- Begin Test remains disabled until required conditions are met

### State 3 — Testing Active

When Begin Test is clicked and validation passes:
- Activity is unhidden
- Conversion is unhidden
- DP is unhidden
- backing sheets remain VeryHidden

Technicians then:
- enter new measurements only through UserForms
- review candidate records through compact visible tables or the record browser form
- choose official pairings and result rows on the dashboards

### New Test / Reset

New Test returns workbook to State 1 and must:
- clear all rows from backing tables but preserve headers and formulas
- clear Home inputs and geometry inputs
- clear signoff state
- hide geometry sheets, Setup Summary, Activity, Conversion, and DP
- return to Home
- require confirmation before execution

---

## 4. Production Sheet Inventory

### Visible at open
- Home
- Specifications

### Conditionally visible
- HC Geometry
- Plate Geometry
- Corrugated Geometry
- Setup Summary
- Activity
- Conversion
- DP

### VeryHidden
- NOx Data
- NH3 Data
- SO2 Data
- SO3 Data
- Temperature Data
- Flow Data
- O2 Data
- DP Data
- Geometry Calc
- Control
- Product Specs
- Lists
- Constants

### Excluded from production build
- Dev Tools
- Calculators
- Formula Reference (as an in-workbook tab)

---

## 5. Color Palette & Formatting Standards

| Role | Color | Usage |
|------|-------|-------|
| Section Header | #0E1638 | Major titles |
| Sub-Section Header | #2C3A8C | Sub-sections |
| Column Header | #4C5CE0 | Table headers |
| Input Cell | #FFF9C4 | Editable cells |
| Formula Output | #E8E8E8 | Calculated/read-only cells |
| Label / Descriptor | #E8EAF6 | Labels and row headers |
| Pass / OK | #C8E6C9 | Pass indicators |
| Warning / Fail | #FFE0B2 / #FFCDD2 | Warnings and failures |

### Font standards
- Calibri throughout
- 16 pt major title
- 14 pt utility title
- 12 pt section title
- 10–11 pt body

### UI style intent

The workbook should use:
- restrained color
- clear section bands
- compact result cards/blocks
- clean table spacing
- modal forms for detailed actions

---

## 6. Home Sheet

### Purpose
Home is the setup and control surface, not the place for all informational output.

### 6.1 Identity
- LRF #
- Load ID
- Project Name
- Date
- Sample Type
- optional Active Technician default
- Gas start times:
  - SO₂ On
  - NH₃ On

### 6.2 Test Conditions

Transposed layout:

| Variable | Activity | Conversion |
|----------|----------|------------|
| AV | input | input |
| UGS | input | input |
| Temperature | input | input |
| H₂O | input | input |
| O₂ | input | input |
| SO₂ | input | input |
| SO₃ | input | input |
| NOx | input | input |
| MR | input | input |

Bottom rows:
- Flow Source / Status
- Flow (Nm³/h)

These are formula outputs from Geometry Calc.

### 6.3 Workflow buttons
- Set Up Geometry
- Open Setup Summary
- Verify Setup
- Begin Test
- New Test

### 6.4 Compact status block

Must show simple readiness items:
- Sample type selected
- Geometry entered
- Geometry resolved
- Flow resolved
- Primary verification complete
- Secondary verification complete
- Ready to begin test
- Activity data present
- Conversion data present
- DP data present

### Must not contain
- full geometry/injection review panels
- slip prediction
- large output dashboards
- backing table content
- helper columns

### Active Technician default
Home may include an Active Technician dropdown used only as a default for forms. It does not replace technician selection on individual forms, because multiple technicians may contribute records during the same workbook lifecycle.

---

## 7. Setup Summary Sheet

### Purpose
Setup Summary is the read-only review page for setup correctness before testing begins.

### Visibility
- hidden at workbook open
- becomes visible after geometry has been entered or after Set Up Geometry returns usable outputs

### 7.1 Geometry & Flow Results
- Avg Adjusted FFA
- Total Adjusted Area
- Active Layers
- Flow (Nm³/h) — Activity and Conversion
- Flow (scfm) — Activity and Conversion
- AV Used — Activity and Conversion
- Flow source/status — Activity and Conversion

### 7.2 Injection Rates
- SO₃ injection — Activity and Conversion
- NH₃ injection — Activity and Conversion
- SO₂ injection — Activity and Conversion
- Combustion NH₃ estimate where applicable

### 7.3 Slip Prediction
- Expected K input
- Predicted slip
- Predicted outlet NOx
- Predicted DeNOx

### 7.4 Setup verification block

Read-only display of:
- Setup entered by
- Setup entered at
- Verification 1 initials
- Verification 1 timestamp
- Verification 2 initials
- Verification 2 timestamp
- signoff status

### 7.5 Action buttons
- Verify Setup
- Back to Home

### Must not contain
- raw geometry entry
- backing tables
- editable signoff cells typed directly by users

---

## 8. Geometry Input Sheets

Three technician-facing input sheets:
- HC Geometry
- Plate Geometry
- Corrugated Geometry

### Shared rule
These sheets are for raw input only. Any displayed calculations are convenience outputs only. Authoritative calculations live in Geometry Calc.

### 8.1 HC Geometry

Inputs per layer:
- Product Type
- AP Override
- Length
- Cells A
- Cells B
- Width A
- Width B
- Plugged Cells

Display only:
- layer status
- total cells
- perhaps a small resolved output block if useful

### 8.2 Plate Geometry

Inputs:
- Total plates
- Box 1 / Box 2 measurements
- 13 plates per box
- Length
- Thickness
- Width A
- Width B

Allowed local convenience:
- average row for operator review

Authoritative geometry outputs still come from Geometry Calc.

### 8.3 Corrugated Geometry

Inputs per layer:
- SSA
- Length
- Width
- Height
- Total Cells
- Plugged Cells

All calculations must occur in Geometry Calc.

---

## 9. Geometry Calc Engine

### Purpose
Geometry Calc is the sole engine for:
- geometry rollups
- flow derivation
- injection rates
- source resolution
- shared outputs used by dashboards

### Required sections
- Active geometry / mode selection
- Unified geometry outputs
- Plate calculations
- Honeycomb calculations
- Corrugated calculations
- Flow calculations
- Injection rate calculations
- Shared resolved outputs
- optional slip prediction support

### Source-resolution pattern
User choices may vary, but downstream formulas must consume resolved outputs only.

Examples of resolved outputs:
- GC_AvgAdjFFA
- GC_TotalAdjArea
- GC_ActiveLayers
- GC_Flow_Act_Nm3h
- GC_Flow_Conv_Nm3h
- GC_Flow_Act_scfm
- GC_Flow_Conv_scfm
- GC_AV_Act
- GC_AV_Conv
- GC_Flow_Act_Status
- GC_Flow_Conv_Status

### Flow rule
Use current calculation behavior:
- if AV entered and UGS blank, derive UGS
- if UGS entered and AV blank, derive AV
- if both entered, warn and use AV as authoritative
- if neither entered, withhold flow/injection calculations

> This remains aligned with the calculations reference. See `SCR_Calculations_Reference.md`

---

## 10. Specifications, Lists, and Constants

### 10.1 Specifications
Visible sheet for editable operating thresholds and tolerances.

Must contain:
- pre-test validation tolerances
- steady-state thresholds
- any user-editable operating limits approved for routine use

### 10.2 Lists
VeryHidden sheet for:
- dropdown sources
- helper lists
- dynamic-array spill bridges where needed
- form list support

### 10.3 Constants
VeryHidden sheet for:
- physical constants
- protected calculation constants
- H2O_Reference

**Fixed decision:** Constants!H2O_Reference = 18%. This value is part of the H₂O-corrected K logic and the fallback chain. It should not be a casual technician-editable value.

> See `SCR_Calculations_Reference.md`

---

## 11. Audit Metadata Standard

Every hidden backing table must include these audit columns:

| Column | Type | Auto-populated | Notes |
|--------|------|----------------|-------|
| RecordID | integer | Yes | sequential per table |
| TestID | text | Yes | based on workbook/test identity |
| EntryType | text | Yes | Form / Manual / Import |
| DateEntered | date | Yes | record creation date |
| TimeEntered | time | Yes | record creation time |
| EnteredBy | text | Yes | technician chosen on form |
| LastModifiedDate | date | Yes | updated on edit |
| LastModifiedTime | time | Yes | updated on edit |
| LastModifiedBy | text | Yes | updated on edit |
| Excluded | boolean | No | default FALSE |
| ExcludeReason | text | No | optional |
| Notes | text | No | optional |

**Rule:** Audit metadata lives in hidden tables only. Visible dashboards show only the operational information needed for decision-making.

---

## 12. Setup Control & Verification Data

Use a single-row control structure on Control for workbook-level setup state.

Recommended named fields or one control table:
- Ctrl_TestID
- Ctrl_WorkbookState
- Ctrl_ActiveGeometryType
- Ctrl_TestStartTimestamp
- Ctrl_SetupEnteredBy
- Ctrl_SetupEnteredAt
- Ctrl_Verified1By
- Ctrl_Verified1At
- Ctrl_Verified2By
- Ctrl_Verified2At
- Ctrl_SetupSignoffComplete

### Verification rule
- Verification 1 can be completed by the person who entered setup.
- Verification 2 must be a second technician.
- Begin Test cannot proceed until both verifications are complete.

---

## 13. Hidden Backing Tables

Each hidden data sheet has one Excel Table, style None.

### 13.1 tbl_NOx — NOx Data

Columns:
- audit columns
- TestType
- SamplePoint (Inlet / Outlet)
- Analyzer (FTIR / NOxAn)
- FTIR_NO_Wet
- FTIR_NO2_Wet
- FTIR_H2O_Pct
- NOxAn_NO_Dry
- NOxAn_NO2_Dry
- DryNOxUsed

Formula rule:
- FTIR uses (NO + NO2) then dry-corrects
- NOx analyzer uses dry values directly

### 13.2 tbl_NH3 — NH3 Data

Columns:
- audit columns
- TestType
- SamplePoint
- Method (FTIR / IC)
- FTIR_NH3_Wet
- FTIR_H2O_Pct
- IC_Result
- Dilution
- MeterVol_L
- MeterTemp_C
- BaroP_mmHg
- DryNH3

### 13.3 tbl_SO2 — SO2 Data

Columns:
- audit columns
- TestType
- Stage (Pre / Test / Post)
- Method (FTIR / IC)
- FTIR_SO2_Wet
- FTIR_H2O_Pct
- IC_Result
- Dilution
- MeterVol_L
- MeterTemp_C
- BaroP_mmHg
- DrySO2

### 13.4 tbl_SO3 — SO3 Data

Columns:
- audit columns
- TestType
- SamplePoint
- PullVol_L
- IC_Result
- Dilution
- MeterTemp_C
- RoomP_mmHg
- CorrectedGasVol
- DryMoles
- SO3Moles
- DrySO3

### 13.5 tbl_Temperature — Temperature Data

Columns:
- audit columns
- TestType
- Context (Pre-Test / During Test)
- Value_C

### 13.6 tbl_Flow — Flow Data

Columns:
- audit columns
- TestType
- Context
- Value_scfm

### 13.7 tbl_O2 — O2 Data

Columns:
- audit columns
- TestType
- Context
- Value_Pct

### 13.8 tbl_DP — DP Data

Columns:
- audit columns
- TestType
- DP_S2
- DP_S3
- DP_S4
- DP_S5
- DPTotal
- TheoryDP
- PctTheory
- Status

---

## 14. Setup Verification Workflow

### Purpose
Because setup correctness is critical, the workbook must force formal review of conditions and geometry before testing starts.

### UI behavior
Verify Setup opens a small modal UserForm.

### Form fields
- verifier initials / technician
- verification type:
  - Verification 1
  - Verification 2
- optional comments
- read-only summary of:
  - test ID
  - sample type
  - geometry status
  - flow status

### Rules
- Verification 1 writes Ctrl_Verified1By and timestamp
- Verification 2 writes Ctrl_Verified2By and timestamp
- Verification 2 must be a different technician from Verification 1 unless an admin override is intentionally designed later
- Begin Test remains disabled until:
  - sample type selected
  - geometry resolved
  - flow resolved
  - Verification 1 complete
  - Verification 2 complete

### Visible display
A compact read-only banner must appear on both Activity and Conversion:
- Setup Check 1: XX
- Setup Check 2: YY
- Verified: timestamp or latest verification time

---

## 15. Activity Dashboard

### Purpose
Clean NOx-removal dashboard for entry, review, pairing, selection, and final reporting.

### Layout order
1. setup/signoff banner
2. current test context
3. pre-test validation
4. result/summary table
5. steady-state section
6. compact reading tables

### 15.1 Setup/signoff banner
Read-only top band showing:
- test ID
- sample type
- flow used
- setup verification initials

### 15.2 Pre-Test Validation
Display most recent or resolved pre-test values for:
- Temperature
- Flow
- O₂
- NOx
- SO₂
- NH₃ (when MR > 0)
- MR (when MR > 0)

Columns:
- Parameter
- Actual
- Target
- LSL
- USL
- Status

Buttons:
- Log Physical
- Log NOx
- Log SO2
- Log NH3 (when MR > 0)

### 15.3 Results / Summary Table
Use a real Excel Table. Six rows is acceptable as the default pass area.

Required columns:
- Use
- Inlet NOx RecordID
- Inlet NOx
- Outlet NOx RecordID
- Outlet NOx
- DeNOx %
- K
- NH3 RecordID
- NH3 Used
- NH3 Sample Point
- MR Actual
- Cormetech K
- H2O Used
- H2O Corr K
- H2O Corr Corm-K
- Browse / Pair

### 15.4 Pairing workflow
The technician must be able to:
- browse inlet candidates
- browse outlet candidates
- browse NH₃ candidates where relevant
- preview the computed result for the chosen set
- apply those selections to a specific pass row

This is done through the record browser form, not raw sheet access.

### 15.5 Steady-State Check
Based on the most recent 4 included (Use = y) rows.

Show:
- K mean
- StdDev
- normalized slope
- steady-state result

### 15.6 Compact reading tables
Visible compact tables at the bottom:
- Inlet NOx
- Outlet NOx
- SO₂
- NH₃

Each block includes:
- \+ Add
- Browse / Pair

The compact visible tables are reference views only, not full storage.

### 15.7 H₂O used rule
The H₂O correction must follow the established fallback chain:
1. Activity outlet pass data
2. Activity inlet pass data
3. SO₂ gas validation
4. NH₃ gas validation
5. H2O_Reference

> This remains aligned with the calculations reference. See `SCR_Calculations_Reference.md`

---

## 16. Conversion Dashboard

### Purpose
Clean SO₂→SO₃ conversion dashboard for source selection, candidate review, pairing, and final result selection.

### Layout order
1. setup/signoff banner
2. current test context
3. pre-test validation
4. result/summary table
5. steady-state section
6. compact reading tables

### 16.1 Setup/signoff banner
Same pattern as Activity.

### 16.2 Pre-Test Validation
Display:
- Temperature
- Flow
- O₂
- NOx
- SO₂
- SO₃
- NH₃ (when MR > 0)
- MR (when MR > 0)

Buttons:
- Log Physical
- Log SO3
- Log SO2
- Log NOx
- Log NH3 (when applicable)

### 16.3 Results / Summary Table
Use a real Excel Table.

Required columns:
- Use
- Outlet SO3 RecordID
- Outlet SO3
- Inlet SO3 Source
- Inlet SO3
- SO2 Source
- SO2 Used
- Difference
- Conversion %
- NH3 RecordID (when MR > 0)
- MR Actual (when MR > 0)
- Browse / Pair

### 16.4 Source selection behavior
Inlet SO3 Source may be:
- Average
- a specific inlet SO₃ RecordID

SO2 Source may be:
- validation value
- test average
- a specific SO₂ RecordID

The user still makes the selection, but the engine resolves the official value centrally.

### 16.5 Pairing workflow
The technician must be able to:
- browse outlet SO₃ candidates
- browse inlet SO₃ candidates
- browse SO₂ candidates or averages
- preview the computed conversion result
- apply those selections to a specific result row

### 16.6 Steady-State Check
Uses current conversion steady-state logic and thresholds from Specifications.

### 16.7 Compact reading tables
Visible compact tables:
- SO₃ Outlet
- SO₃ Inlet
- SO₂
- NOx
- NH₃ (when applicable)

Each includes:
- \+ Add
- Browse / Pair

---

## 17. DP Dashboard

### Purpose
Dedicated DP review and entry page.

### Layout
- setup/signoff banner
- Log DP button
- visible DP table

### Visible columns
- RecordID
- Test Type
- DP@S2
- DP@S3
- DP@S4
- DP@S5
- DP Total
- Theory DP
- % Theory
- Status

**Rule:** DP entry is form-based only.

---

## 18. Record Browser / Pair Selector UserForm

### Purpose
This form replaces the old "View All" raw-sheet workflow.

### Name
Recommended: `frmRecordBrowser`

### General behavior
The form opens in context from the current dashboard and current row.

Examples:
- Activity row 2 → open browser for pass row 2
- Conversion row 4 → open browser for conversion row 4

### Activity mode layout

Sections:
- Inlet NOx candidates
- Outlet NOx candidates
- NH₃ candidates (when applicable)
- preview block

Preview block shows:
- inlet NOx used
- outlet NOx used
- DeNOx %
- K
- NH₃ used
- MR Actual
- Cormetech K
- H₂O used

Buttons:
- Apply to Selected Row
- Clear Row
- Cancel

### Conversion mode layout

Sections:
- Outlet SO₃ candidates
- Inlet SO₃ candidates
- SO₂ source candidates
- preview block

Preview block shows:
- outlet SO₃
- inlet SO₃
- SO₂ used
- difference
- conversion %

Buttons:
- Apply to Selected Row
- Clear Row
- Cancel

### Filter options
Support filters such as:
- included only / all
- inlet / outlet
- time order
- technician
- method

**Rule:** This form is the official detailed review surface. It prevents technicians from flying blind without exposing raw storage sheets.

---

## 19. Other UserForms

Recommended final form inventory:
- frmNOx
- frmNH3
- frmSO2
- frmSO3
- frmPhysical
- frmDP
- frmRecordBrowser
- frmSetupVerify

### Shared form behavior
All entry forms should support:
- Save
- Save & New
- Duplicate Prior
- Exclude / Restore (where meaningful)
- Cancel

Each form auto-populates:
- date
- time
- default technician from Home if available
- test ID
- test type / context
- audit metadata

But the operator must still be able to choose the actual technician for that record.

---

## 20. Visible Summary Tables vs Hidden Storage

**Rule:** Where the operator sees rows, use real Excel Tables or tightly controlled visible tabular regions.

Examples:
- Activity result set
- Conversion result set
- compact reading views where helpful

**Rule:** Visible summary tables are not the raw store. Hidden normalized tables remain the source of truth.

---

## 21. Data Validation Strategy

### Summary row editing
The dashboard result rows may support either:
- direct dropdown selection for fast editing
- or browser-form selection

### Recommendation
Keep both:
- lightweight dropdowns for fast correction
- browser form for guided review and pairing

### Lists sheet use
Use Lists helper arrays/ranges for:
- filtered RecordID dropdowns
- technician lists
- stage lists
- sample point lists
- source-mode options

---

## 22. Protection Strategy

### Visible sheets
- lock everything by default
- unlock only intended input cells
- protect sheets with password
- unlock summary selector cells where editing is intended

### Hidden sheets
- keep backing tables VeryHidden
- never unhide them during normal technician workflow
- allow VBA to write to them under controlled routines

### Forms
- forms are the preferred route for new records
- avoid direct row typing into hidden tables

---

## 23. Implementation Phases

### Phase 1 — Workbook shell
- create production sheet set
- initial visibility states
- table shells
- named ranges
- constants
- lists
- control structure

### Phase 2 — Geometry engine
- build geometry input sheets
- build Geometry Calc
- publish resolved named outputs

### Phase 3 — Home and Setup Summary
- build clean setup workflow
- status block
- setup review outputs
- signoff display

### Phase 4 — Hidden tables and forms
- build normalized hidden tables
- build frmNOx, frmNH3, frmSO2, frmSO3, frmPhysical, frmDP

### Phase 5 — Verification and browser workflow
- build frmSetupVerify
- build frmRecordBrowser
- wire row-specific pairing actions

### Phase 6 — Dashboards
- Activity dashboard
- Conversion dashboard
- DP dashboard
- summary formulas
- steady-state sections

### Phase 7 — Protection and testing
- protection logic
- reset workflow
- end-to-end validation
- formula audit
- operator walkthrough testing

---

## 24. Closed Decisions

These are no longer open questions.

- H2O_Reference = 18% on Constants
- Activity and Conversion both unlock on Begin Test
- backing sheets remain VeryHidden
- View All is replaced by browser-form review/pairing
- Home status block is included
- dual setup verification is required
- Activity and Conversion display setup verification initials
- Home and Setup Summary are separate
- production workbook excludes dev/debug sheets

---

## 25. Acceptance Standard

The workbook is ready to build when the engineer treats this PRD as the single architectural source.

The workbook is successful when:
- technicians work from clean visible sheets only
- setup is dual-verified before testing starts
- all data entry is controlled and auditable
- backing storage is hidden and normalized
- Activity and Conversion are parallel and intuitive
- technicians can review and pair readings without browsing raw sheets
- calculations are centralized and trustworthy
- the finished workbook feels like a controlled professional tool, not a legacy spreadsheet

**Next step:** lock the exact table names, named ranges, and form field names before VBA development begins.
