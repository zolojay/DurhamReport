"""
Phase 1 — Build Workbook Shell
Creates the SCR Catalyst Test Report workbook with:
  - All production sheets with correct visibility states
  - Excel Table shells on backing data sheets
  - Named ranges (Ctrl_, GC_, Spec_ prefixes)
  - Constants sheet with H2O_Reference = 18%
  - Lists sheet with dropdown sources
  - Control sheet with single-row control structure
  - Color palette formatting and font standards
"""

import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName

# ── Output path ──────────────────────────────────────────────────────────────
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "..", "output")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "SCR_TestReport.xlsm")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ── Color palette ────────────────────────────────────────────────────────────
COLORS = {
    "section_header":  "0E1638",
    "sub_section":     "2C3A8C",
    "column_header":   "4C5CE0",
    "input_cell":      "FFF9C4",
    "formula_output":  "E8E8E8",
    "label":           "E8EAF6",
    "pass":            "C8E6C9",
    "warning":         "FFE0B2",
    "fail":            "FFCDD2",
    "white":           "FFFFFF",
}

FILLS = {k: PatternFill(start_color=v, end_color=v, fill_type="solid") for k, v in COLORS.items()}

# ── Font definitions ─────────────────────────────────────────────────────────
FONT_TITLE_16     = Font(name="Calibri", size=16, bold=True, color=COLORS["white"])
FONT_UTILITY_14   = Font(name="Calibri", size=14, bold=True, color=COLORS["white"])
FONT_SECTION_12   = Font(name="Calibri", size=12, bold=True, color=COLORS["white"])
FONT_BODY_11      = Font(name="Calibri", size=11)
FONT_BODY_10      = Font(name="Calibri", size=10)
FONT_BODY_BOLD_11 = Font(name="Calibri", size=11, bold=True)
FONT_HEADER_COL   = Font(name="Calibri", size=11, bold=True, color=COLORS["white"])
FONT_DEFAULT      = Font(name="Calibri", size=11)

THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)

# ── Sheet definitions ────────────────────────────────────────────────────────
# (name, visibility)  visibility: "visible", "hidden", "veryHidden"
SHEET_DEFS = [
    ("Home",                "visible"),
    ("Specifications",      "visible"),
    ("HC Geometry",         "hidden"),
    ("Plate Geometry",      "hidden"),
    ("Corrugated Geometry", "hidden"),
    ("Setup Summary",       "hidden"),
    ("Activity",            "hidden"),
    ("Conversion",          "hidden"),
    ("DP",                  "hidden"),
    ("NOx Data",            "veryHidden"),
    ("NH3 Data",            "veryHidden"),
    ("SO2 Data",            "veryHidden"),
    ("SO3 Data",            "veryHidden"),
    ("Temperature Data",    "veryHidden"),
    ("Flow Data",           "veryHidden"),
    ("O2 Data",             "veryHidden"),
    ("DP Data",             "veryHidden"),
    ("Geometry Calc",       "veryHidden"),
    ("Control",             "veryHidden"),
    ("Product Specs",       "veryHidden"),
    ("Lists",               "veryHidden"),
    ("Constants",           "veryHidden"),
]

# ── Audit columns (shared by all backing tables) ────────────────────────────
AUDIT_COLUMNS = [
    "RecordID", "TestID", "EntryType", "DateEntered", "TimeEntered",
    "EnteredBy", "LastModifiedDate", "LastModifiedTime", "LastModifiedBy",
    "Excluded", "ExcludeReason", "Notes",
]

# ── Backing table definitions ────────────────────────────────────────────────
BACKING_TABLES = {
    "NOx Data": {
        "table_name": "tbl_NOx",
        "columns": AUDIT_COLUMNS + [
            "TestType", "SamplePoint", "Analyzer",
            "FTIR_NO_Wet", "FTIR_NO2_Wet", "FTIR_H2O_Pct",
            "NOxAn_NO_Dry", "NOxAn_NO2_Dry", "DryNOxUsed",
        ],
    },
    "NH3 Data": {
        "table_name": "tbl_NH3",
        "columns": AUDIT_COLUMNS + [
            "TestType", "SamplePoint", "Method",
            "FTIR_NH3_Wet", "FTIR_H2O_Pct",
            "IC_Result", "Dilution", "MeterVol_L", "MeterTemp_C", "BaroP_mmHg",
            "DryNH3",
        ],
    },
    "SO2 Data": {
        "table_name": "tbl_SO2",
        "columns": AUDIT_COLUMNS + [
            "TestType", "Stage", "Method",
            "FTIR_SO2_Wet", "FTIR_H2O_Pct",
            "IC_Result", "Dilution", "MeterVol_L", "MeterTemp_C", "BaroP_mmHg",
            "DrySO2",
        ],
    },
    "SO3 Data": {
        "table_name": "tbl_SO3",
        "columns": AUDIT_COLUMNS + [
            "TestType", "SamplePoint",
            "PullVol_L", "IC_Result", "Dilution", "MeterTemp_C", "RoomP_mmHg",
            "CorrectedGasVol", "DryMoles", "SO3Moles", "DrySO3",
        ],
    },
    "Temperature Data": {
        "table_name": "tbl_Temperature",
        "columns": AUDIT_COLUMNS + [
            "TestType", "Context", "Value_C",
        ],
    },
    "Flow Data": {
        "table_name": "tbl_Flow",
        "columns": AUDIT_COLUMNS + [
            "TestType", "Context", "Value_scfm",
        ],
    },
    "O2 Data": {
        "table_name": "tbl_O2",
        "columns": AUDIT_COLUMNS + [
            "TestType", "Context", "Value_Pct",
        ],
    },
    "DP Data": {
        "table_name": "tbl_DP",
        "columns": AUDIT_COLUMNS + [
            "TestType",
            "DP_S2", "DP_S3", "DP_S4", "DP_S5",
            "DPTotal", "TheoryDP", "PctTheory", "Status",
        ],
    },
}

# ── Control sheet fields ─────────────────────────────────────────────────────
CONTROL_FIELDS = [
    ("Ctrl_TestID",                ""),
    ("Ctrl_WorkbookState",         "Setup"),
    ("Ctrl_ActiveGeometryType",    ""),
    ("Ctrl_TestStartTimestamp",    ""),
    ("Ctrl_SetupEnteredBy",        ""),
    ("Ctrl_SetupEnteredAt",        ""),
    ("Ctrl_Verified1By",           ""),
    ("Ctrl_Verified1At",           ""),
    ("Ctrl_Verified2By",           ""),
    ("Ctrl_Verified2At",           ""),
    ("Ctrl_SetupSignoffComplete",  "FALSE"),
]

# ── Constants ────────────────────────────────────────────────────────────────
CONSTANTS = [
    ("H2O_Reference",               0.18,    "H₂O reference for fallback chain (18%)"),
    ("MolarMass_NO",                30.01,    "g/mol"),
    ("MolarMass_NO2",               46.01,    "g/mol"),
    ("MolarMass_NH3",               17.03,    "g/mol"),
    ("MolarMass_SO2",               64.07,    "g/mol"),
    ("MolarMass_SO3",               80.06,    "g/mol"),
    ("STP_Temp_K",                 273.15,    "Standard temperature (K)"),
    ("STP_Pressure_mmHg",          760.0,     "Standard pressure (mmHg)"),
    ("IdealGasVol_L",               22.414,   "Ideal gas molar volume at STP (L/mol)"),
    ("Nm3_to_scfm_Factor",         0.58858,   "Conversion factor Nm³/h → scfm"),
]

# ── Lists sheet content ──────────────────────────────────────────────────────
LISTS_DATA = {
    "SampleType":       ["Honeycomb", "Plate", "Corrugated"],
    "GeometryType":     ["HC", "Plate", "Corrugated"],
    "TestType":         ["Activity", "Conversion"],
    "SamplePoint":      ["Inlet", "Outlet"],
    "NOxAnalyzer":      ["FTIR", "NOxAn"],
    "NH3Method":        ["FTIR", "IC"],
    "SO2Method":        ["FTIR", "IC"],
    "SO2Stage":         ["Pre", "Test", "Post"],
    "TempContext":      ["Pre-Test", "During Test"],
    "FlowContext":      ["Pre-Test", "During Test"],
    "O2Context":        ["Pre-Test", "During Test"],
    "EntryType":        ["Form", "Manual", "Import"],
    "YesNo":            ["Yes", "No"],
    "TrueFalse":        ["TRUE", "FALSE"],
    "InletSO3Source":   ["Average", "RecordID"],
    "SO2Source":        ["Validation", "Test Average", "RecordID"],
    "Technicians":      ["Tech1", "Tech2", "Tech3", "Tech4"],
}

# ── Specifications defaults ──────────────────────────────────────────────────
SPEC_FIELDS = [
    ("Spec_Temp_Tolerance_C",           5,     "Temperature tolerance (±°C)"),
    ("Spec_Flow_Tolerance_Pct",         5,     "Flow tolerance (±%)"),
    ("Spec_O2_Tolerance_Pct",           0.5,   "O₂ tolerance (±%)"),
    ("Spec_NOx_Tolerance_Pct",          5,     "NOx tolerance (±%)"),
    ("Spec_SO2_Tolerance_Pct",          10,    "SO₂ tolerance (±%)"),
    ("Spec_NH3_Tolerance_Pct",          10,    "NH₃ tolerance (±%)"),
    ("Spec_MR_Tolerance",               0.02,  "MR tolerance (±)"),
    ("Spec_SS_MinPoints",               4,     "Steady-state min points"),
    ("Spec_SS_K_StdDev_Max",            0.05,  "K StdDev threshold for steady-state"),
    ("Spec_SS_NormSlope_Max",           0.02,  "Normalized slope threshold"),
    ("Spec_SS_Conv_StdDev_Max",         2.0,   "Conversion StdDev threshold (%)"),
    ("Spec_SS_Conv_NormSlope_Max",      0.02,  "Conversion normalized slope threshold"),
    ("Spec_DP_PctTheory_Warning",       120,   "DP % Theory warning threshold"),
    ("Spec_DP_PctTheory_Fail",          150,   "DP % Theory fail threshold"),
]

# ── GC named ranges (placeholders for Phase 2) ──────────────────────────────
GC_NAMED_RANGES = [
    "GC_AvgAdjFFA",
    "GC_TotalAdjArea",
    "GC_ActiveLayers",
    "GC_Flow_Act_Nm3h",
    "GC_Flow_Conv_Nm3h",
    "GC_Flow_Act_scfm",
    "GC_Flow_Conv_scfm",
    "GC_AV_Act",
    "GC_AV_Conv",
    "GC_Flow_Act_Status",
    "GC_Flow_Conv_Status",
    "GC_SO3_Inj_Act",
    "GC_SO3_Inj_Conv",
    "GC_NH3_Inj_Act",
    "GC_NH3_Inj_Conv",
    "GC_SO2_Inj_Act",
    "GC_SO2_Inj_Conv",
]


# ═══════════════════════════════════════════════════════════════════════════════
# Helper functions
# ═══════════════════════════════════════════════════════════════════════════════

def apply_section_header(ws, row, col_start, col_end, text, level="section"):
    """Write a merged section-header band."""
    if level == "title":
        font = FONT_TITLE_16
        fill = FILLS["section_header"]
    elif level == "sub":
        font = FONT_UTILITY_14
        fill = FILLS["sub_section"]
    else:
        font = FONT_SECTION_12
        fill = FILLS["section_header"]

    cell = ws.cell(row=row, column=col_start, value=text)
    cell.font = font
    cell.fill = fill
    cell.alignment = Alignment(horizontal="left", vertical="center")
    if col_end > col_start:
        ws.merge_cells(start_row=row, start_column=col_start,
                        end_row=row, end_column=col_end)
    # Fill background for merged range
    for c in range(col_start + 1, col_end + 1):
        ws.cell(row=row, column=c).fill = fill


def write_label_value_pair(ws, row, label_col, label_text, value_col, value=None):
    """Write a label cell and an adjacent value cell with standard formatting."""
    lc = ws.cell(row=row, column=label_col, value=label_text)
    lc.font = FONT_BODY_BOLD_11
    lc.fill = FILLS["label"]
    lc.alignment = Alignment(horizontal="right", vertical="center")
    lc.border = THIN_BORDER

    vc = ws.cell(row=row, column=value_col, value=value)
    vc.font = FONT_BODY_11
    vc.fill = FILLS["input_cell"]
    vc.alignment = Alignment(horizontal="left", vertical="center")
    vc.border = THIN_BORDER
    return vc


def write_table_headers(ws, row, col_start, headers):
    """Write column-header row with standard formatting."""
    for i, h in enumerate(headers):
        cell = ws.cell(row=row, column=col_start + i, value=h)
        cell.font = FONT_HEADER_COL
        cell.fill = FILLS["column_header"]
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER


def create_excel_table(ws, table_name, headers, header_row, col_start, data_rows=1):
    """Create an Excel Table with headers and empty data rows."""
    col_end = col_start + len(headers) - 1
    end_row = header_row + data_rows

    write_table_headers(ws, header_row, col_start, headers)

    # Write empty data rows with borders
    for r in range(header_row + 1, end_row + 1):
        for c in range(col_start, col_end + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = FONT_BODY_10
            cell.border = THIN_BORDER

    ref = "{}{}:{}{}".format(
        get_column_letter(col_start), header_row,
        get_column_letter(col_end), end_row,
    )
    table = Table(displayName=table_name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleLight1", showFirstColumn=False,
        showLastColumn=False, showRowStripes=False, showColumnStripes=False,
    )
    ws.add_table(table)
    return table


def set_column_widths(ws, widths):
    """Set column widths from a dict {col_letter: width}."""
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


# ═══════════════════════════════════════════════════════════════════════════════
# Build functions for each sheet
# ═══════════════════════════════════════════════════════════════════════════════

def build_home(ws):
    """Build Home sheet — identity, test conditions, workflow buttons, status block."""
    set_column_widths(ws, {"A": 3, "B": 22, "C": 20, "D": 20, "E": 20, "F": 3})

    # ── Title ────────────────────────────────────────────────────────────
    apply_section_header(ws, 1, 2, 5, "SCR Catalyst Test Report", level="title")
    ws.row_dimensions[1].height = 30

    # ── Identity section ─────────────────────────────────────────────────
    apply_section_header(ws, 3, 2, 5, "Test Identity", level="section")
    identity_fields = [
        ("LRF #", None),
        ("Load ID", None),
        ("Project Name", None),
        ("Date", None),
        ("Sample Type", None),
        ("Active Technician", None),
        ("SO₂ Gas On", None),
        ("NH₃ Gas On", None),
    ]
    for i, (label, default) in enumerate(identity_fields):
        write_label_value_pair(ws, 4 + i, 2, label, 3, default)

    # ── Test Conditions ──────────────────────────────────────────────────
    apply_section_header(ws, 13, 2, 5, "Test Conditions", level="section")

    # Column headers: Variable | Activity | Conversion
    for col, text in [(2, "Variable"), (3, "Activity"), (4, "Conversion")]:
        cell = ws.cell(row=14, column=col, value=text)
        cell.font = FONT_HEADER_COL
        cell.fill = FILLS["column_header"]
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER

    condition_vars = ["AV", "UGS", "Temperature", "H₂O", "O₂", "SO₂", "SO₃", "NOx", "MR"]
    for i, var in enumerate(condition_vars):
        r = 15 + i
        lbl = ws.cell(row=r, column=2, value=var)
        lbl.font = FONT_BODY_BOLD_11
        lbl.fill = FILLS["label"]
        lbl.border = THIN_BORDER
        lbl.alignment = Alignment(horizontal="right", vertical="center")
        for c in [3, 4]:
            cell = ws.cell(row=r, column=c)
            cell.fill = FILLS["input_cell"]
            cell.font = FONT_BODY_11
            cell.border = THIN_BORDER

    # Bottom rows: Flow Source/Status and Flow (Nm³/h) — formula outputs
    for i, label in enumerate(["Flow Source / Status", "Flow (Nm³/h)"]):
        r = 24 + i
        lbl = ws.cell(row=r, column=2, value=label)
        lbl.font = FONT_BODY_BOLD_11
        lbl.fill = FILLS["label"]
        lbl.border = THIN_BORDER
        lbl.alignment = Alignment(horizontal="right", vertical="center")
        for c in [3, 4]:
            cell = ws.cell(row=r, column=c)
            cell.fill = FILLS["formula_output"]
            cell.font = FONT_BODY_11
            cell.border = THIN_BORDER

    # ── Workflow Buttons placeholder ─────────────────────────────────────
    apply_section_header(ws, 27, 2, 5, "Workflow", level="section")
    buttons = ["Set Up Geometry", "Open Setup Summary", "Verify Setup", "Begin Test", "New Test"]
    for i, btn in enumerate(buttons):
        cell = ws.cell(row=28 + i, column=2, value=f"[ {btn} ]")
        cell.font = FONT_BODY_BOLD_11
        cell.fill = FILLS["label"]
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells(start_row=28 + i, start_column=2, end_row=28 + i, end_column=3)

    # ── Compact Status Block ─────────────────────────────────────────────
    apply_section_header(ws, 34, 2, 5, "Readiness Status", level="section")
    status_items = [
        "Sample type selected",
        "Geometry entered",
        "Geometry resolved",
        "Flow resolved",
        "Primary verification complete",
        "Secondary verification complete",
        "Ready to begin test",
        "Activity data present",
        "Conversion data present",
        "DP data present",
    ]
    for i, item in enumerate(status_items):
        r = 35 + i
        lbl = ws.cell(row=r, column=2, value=item)
        lbl.font = FONT_BODY_11
        lbl.fill = FILLS["label"]
        lbl.border = THIN_BORDER
        lbl.alignment = Alignment(horizontal="right", vertical="center")
        val = ws.cell(row=r, column=3, value="—")
        val.font = FONT_BODY_11
        val.fill = FILLS["formula_output"]
        val.border = THIN_BORDER
        val.alignment = Alignment(horizontal="center", vertical="center")

    ws.sheet_properties.tabColor = "0E1638"


def build_specifications(ws):
    """Build Specifications sheet with editable thresholds."""
    set_column_widths(ws, {"A": 3, "B": 40, "C": 16, "D": 40, "E": 3})

    apply_section_header(ws, 1, 2, 4, "Specifications & Thresholds", level="title")
    ws.row_dimensions[1].height = 30

    # Headers
    for col, text in [(2, "Parameter"), (3, "Value"), (4, "Description")]:
        cell = ws.cell(row=3, column=col, value=text)
        cell.font = FONT_HEADER_COL
        cell.fill = FILLS["column_header"]
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER

    for i, (name, value, desc) in enumerate(SPEC_FIELDS):
        r = 4 + i
        nc = ws.cell(row=r, column=2, value=name)
        nc.font = FONT_BODY_11
        nc.fill = FILLS["label"]
        nc.border = THIN_BORDER

        vc = ws.cell(row=r, column=3, value=value)
        vc.font = FONT_BODY_11
        vc.fill = FILLS["input_cell"]
        vc.border = THIN_BORDER
        vc.alignment = Alignment(horizontal="center")

        dc = ws.cell(row=r, column=4, value=desc)
        dc.font = FONT_BODY_10
        dc.fill = FILLS["label"]
        dc.border = THIN_BORDER


def build_geometry_input(ws, geo_type):
    """Build a geometry input sheet shell (HC, Plate, or Corrugated)."""
    set_column_widths(ws, {"A": 3, "B": 18, "C": 14, "D": 14, "E": 14,
                            "F": 14, "G": 14, "H": 14, "I": 14, "J": 3})

    apply_section_header(ws, 1, 2, 9, f"{geo_type} Geometry Input", level="title")
    ws.row_dimensions[1].height = 30

    if geo_type == "HC":
        headers = ["Layer", "Product Type", "AP Override", "Length",
                    "Cells A", "Cells B", "Width A", "Width B", "Plugged Cells"]
        write_table_headers(ws, 3, 2, headers)
        for r in range(4, 10):  # 6 layer rows
            ws.cell(row=r, column=2, value=r - 3)
            for c in range(2, 11):
                cell = ws.cell(row=r, column=c)
                cell.fill = FILLS["input_cell"] if c > 2 else FILLS["label"]
                cell.font = FONT_BODY_11
                cell.border = THIN_BORDER

    elif geo_type == "Plate":
        headers = ["Measurement", "Box 1", "Box 2"]
        write_table_headers(ws, 3, 2, headers)
        plate_rows = ["Total Plates", "Length", "Thickness",
                       "Width A", "Width B"]
        for i, label in enumerate(plate_rows):
            r = 4 + i
            ws.cell(row=r, column=2, value=label).font = FONT_BODY_BOLD_11
            ws.cell(row=r, column=2).fill = FILLS["label"]
            ws.cell(row=r, column=2).border = THIN_BORDER
            for c in [3, 4]:
                cell = ws.cell(row=r, column=c)
                cell.fill = FILLS["input_cell"]
                cell.font = FONT_BODY_11
                cell.border = THIN_BORDER

    elif geo_type == "Corrugated":
        headers = ["Layer", "SSA", "Length", "Width", "Height",
                    "Total Cells", "Plugged Cells"]
        write_table_headers(ws, 3, 2, headers)
        for r in range(4, 10):  # 6 layer rows
            ws.cell(row=r, column=2, value=r - 3)
            for c in range(2, 9):
                cell = ws.cell(row=r, column=c)
                cell.fill = FILLS["input_cell"] if c > 2 else FILLS["label"]
                cell.font = FONT_BODY_11
                cell.border = THIN_BORDER

    # Navigation placeholder
    r_nav = 12
    cell = ws.cell(row=r_nav, column=2, value="[ Back to Home ]")
    cell.font = FONT_BODY_BOLD_11
    cell.fill = FILLS["label"]
    cell.border = THIN_BORDER
    cell.alignment = Alignment(horizontal="center")


def build_setup_summary(ws):
    """Build Setup Summary sheet shell — read-only review page."""
    set_column_widths(ws, {"A": 3, "B": 30, "C": 20, "D": 20, "E": 3})

    apply_section_header(ws, 1, 2, 4, "Setup Summary", level="title")
    ws.row_dimensions[1].height = 30

    # ── Geometry & Flow Results ──────────────────────────────────────────
    apply_section_header(ws, 3, 2, 4, "Geometry & Flow Results", level="section")
    geo_items = [
        "Avg Adjusted FFA", "Total Adjusted Area", "Active Layers",
    ]
    for i, label in enumerate(geo_items):
        write_label_value_pair(ws, 4 + i, 2, label, 3)

    # Flow for Activity and Conversion
    flow_header_row = 8
    for col, text in [(2, "Flow Parameter"), (3, "Activity"), (4, "Conversion")]:
        cell = ws.cell(row=flow_header_row, column=col, value=text)
        cell.font = FONT_HEADER_COL
        cell.fill = FILLS["column_header"]
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center")

    flow_params = ["Flow (Nm³/h)", "Flow (scfm)", "AV Used", "Flow Source/Status"]
    for i, param in enumerate(flow_params):
        r = flow_header_row + 1 + i
        ws.cell(row=r, column=2, value=param).font = FONT_BODY_BOLD_11
        ws.cell(row=r, column=2).fill = FILLS["label"]
        ws.cell(row=r, column=2).border = THIN_BORDER
        for c in [3, 4]:
            cell = ws.cell(row=r, column=c)
            cell.fill = FILLS["formula_output"]
            cell.font = FONT_BODY_11
            cell.border = THIN_BORDER

    # ── Injection Rates ──────────────────────────────────────────────────
    apply_section_header(ws, 14, 2, 4, "Injection Rates", level="section")
    inj_params = ["SO₃ Injection", "NH₃ Injection", "SO₂ Injection", "Combustion NH₃ Est."]
    for col, text in [(2, "Parameter"), (3, "Activity"), (4, "Conversion")]:
        cell = ws.cell(row=15, column=col, value=text)
        cell.font = FONT_HEADER_COL
        cell.fill = FILLS["column_header"]
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center")

    for i, param in enumerate(inj_params):
        r = 16 + i
        ws.cell(row=r, column=2, value=param).font = FONT_BODY_BOLD_11
        ws.cell(row=r, column=2).fill = FILLS["label"]
        ws.cell(row=r, column=2).border = THIN_BORDER
        for c in [3, 4]:
            cell = ws.cell(row=r, column=c)
            cell.fill = FILLS["formula_output"]
            cell.font = FONT_BODY_11
            cell.border = THIN_BORDER

    # ── Slip Prediction ──────────────────────────────────────────────────
    apply_section_header(ws, 21, 2, 4, "Slip Prediction", level="section")
    slip_items = ["Expected K", "Predicted Slip", "Predicted Outlet NOx", "Predicted DeNOx"]
    for i, label in enumerate(slip_items):
        write_label_value_pair(ws, 22 + i, 2, label, 3)

    # ── Setup Verification Block ─────────────────────────────────────────
    apply_section_header(ws, 27, 2, 4, "Setup Verification", level="section")
    verify_items = [
        "Setup Entered By", "Setup Entered At",
        "Verification 1 Initials", "Verification 1 Timestamp",
        "Verification 2 Initials", "Verification 2 Timestamp",
        "Signoff Status",
    ]
    for i, label in enumerate(verify_items):
        r = 28 + i
        write_label_value_pair(ws, r, 2, label, 3)
        ws.cell(row=r, column=3).fill = FILLS["formula_output"]

    # ── Action buttons ───────────────────────────────────────────────────
    r_btn = 36
    for i, btn in enumerate(["Verify Setup", "Back to Home"]):
        cell = ws.cell(row=r_btn + i, column=2, value=f"[ {btn} ]")
        cell.font = FONT_BODY_BOLD_11
        cell.fill = FILLS["label"]
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center")


def build_dashboard_shell(ws, dashboard_type):
    """Build Activity, Conversion, or DP dashboard shell."""
    set_column_widths(ws, {"A": 3, "B": 18, "C": 14, "D": 14, "E": 14,
                            "F": 14, "G": 14, "H": 14, "I": 14, "J": 14,
                            "K": 14, "L": 14, "M": 14, "N": 14, "O": 14,
                            "P": 14, "Q": 14})

    apply_section_header(ws, 1, 2, 10, f"{dashboard_type} Dashboard", level="title")
    ws.row_dimensions[1].height = 30

    # Setup/signoff banner placeholder
    apply_section_header(ws, 3, 2, 10, "Setup Verification", level="sub")
    banner_items = ["Test ID", "Sample Type", "Flow Used", "Setup Check 1", "Setup Check 2", "Verified"]
    for i, label in enumerate(banner_items):
        write_label_value_pair(ws, 4 + i, 2, label, 3)
        ws.cell(row=4 + i, column=3).fill = FILLS["formula_output"]

    if dashboard_type == "DP":
        # DP is simpler — just a log button and table placeholder
        apply_section_header(ws, 11, 2, 10, "DP Records", level="section")
        cell = ws.cell(row=12, column=2, value="[ Log DP ]")
        cell.font = FONT_BODY_BOLD_11
        cell.fill = FILLS["label"]
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center")
        # DP visible table headers placeholder
        dp_headers = ["RecordID", "Test Type", "DP@S2", "DP@S3", "DP@S4",
                       "DP@S5", "DP Total", "Theory DP", "% Theory", "Status"]
        write_table_headers(ws, 14, 2, dp_headers)
    else:
        # Pre-Test Validation placeholder
        apply_section_header(ws, 11, 2, 10, "Pre-Test Validation", level="section")

        # Results/Summary table placeholder
        r_results = 20
        apply_section_header(ws, r_results, 2, 16, f"{dashboard_type} Results", level="section")

        # Steady-State placeholder
        r_ss = 30
        apply_section_header(ws, r_ss, 2, 10, "Steady-State Check", level="section")

        # Compact reading tables placeholder
        r_compact = 36
        apply_section_header(ws, r_compact, 2, 10, "Reading Tables", level="section")


def build_geometry_calc(ws):
    """Build Geometry Calc engine sheet shell — VeryHidden, holds all GC_ outputs."""
    set_column_widths(ws, {"A": 3, "B": 30, "C": 20, "D": 20, "E": 3})

    apply_section_header(ws, 1, 2, 4, "Geometry Calc Engine", level="title")

    # Section placeholders for Phase 2
    sections = [
        (3,  "Active Geometry / Mode Selection"),
        (8,  "Unified Geometry Outputs"),
        (16, "Honeycomb Calculations"),
        (24, "Plate Calculations"),
        (32, "Corrugated Calculations"),
        (40, "Flow Calculations"),
        (50, "Injection Rate Calculations"),
        (60, "Shared Resolved Outputs"),
    ]
    for r, title in sections:
        apply_section_header(ws, r, 2, 4, title, level="section")

    # Write GC_ output placeholders in the Shared Resolved Outputs section
    for i, name in enumerate(GC_NAMED_RANGES):
        r = 61 + i
        ws.cell(row=r, column=2, value=name).font = FONT_BODY_BOLD_11
        ws.cell(row=r, column=2).fill = FILLS["label"]
        ws.cell(row=r, column=2).border = THIN_BORDER
        cell = ws.cell(row=r, column=3)
        cell.fill = FILLS["formula_output"]
        cell.border = THIN_BORDER


def build_control(ws):
    """Build Control sheet — single-row state structure."""
    set_column_widths(ws, {"A": 3, "B": 35, "C": 30, "D": 3})

    apply_section_header(ws, 1, 2, 3, "Workbook Control", level="title")

    for i, (field, default) in enumerate(CONTROL_FIELDS):
        r = 3 + i
        ws.cell(row=r, column=2, value=field).font = FONT_BODY_BOLD_11
        ws.cell(row=r, column=2).fill = FILLS["label"]
        ws.cell(row=r, column=2).border = THIN_BORDER
        vc = ws.cell(row=r, column=3, value=default)
        vc.font = FONT_BODY_11
        vc.fill = FILLS["formula_output"]
        vc.border = THIN_BORDER


def build_constants(ws):
    """Build Constants sheet — physical constants and protected values."""
    set_column_widths(ws, {"A": 3, "B": 30, "C": 18, "D": 40, "E": 3})

    apply_section_header(ws, 1, 2, 4, "Constants", level="title")

    for col, text in [(2, "Name"), (3, "Value"), (4, "Description")]:
        cell = ws.cell(row=3, column=col, value=text)
        cell.font = FONT_HEADER_COL
        cell.fill = FILLS["column_header"]
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center")

    for i, (name, value, desc) in enumerate(CONSTANTS):
        r = 4 + i
        ws.cell(row=r, column=2, value=name).font = FONT_BODY_BOLD_11
        ws.cell(row=r, column=2).fill = FILLS["label"]
        ws.cell(row=r, column=2).border = THIN_BORDER

        vc = ws.cell(row=r, column=3, value=value)
        vc.font = FONT_BODY_11
        vc.fill = FILLS["formula_output"]
        vc.border = THIN_BORDER
        vc.alignment = Alignment(horizontal="center")

        dc = ws.cell(row=r, column=4, value=desc)
        dc.font = FONT_BODY_10
        dc.fill = FILLS["label"]
        dc.border = THIN_BORDER


def build_lists(ws):
    """Build Lists sheet — dropdown sources and helper lists."""
    set_column_widths(ws, {"A": 3})

    apply_section_header(ws, 1, 2, 2 + len(LISTS_DATA) - 1,
                          "Lists & Dropdown Sources", level="title")

    col = 2
    for list_name, items in LISTS_DATA.items():
        ws.column_dimensions[get_column_letter(col)].width = 20
        # Header
        hdr = ws.cell(row=3, column=col, value=list_name)
        hdr.font = FONT_HEADER_COL
        hdr.fill = FILLS["column_header"]
        hdr.border = THIN_BORDER
        hdr.alignment = Alignment(horizontal="center")
        # Items
        for j, item in enumerate(items):
            cell = ws.cell(row=4 + j, column=col, value=item)
            cell.font = FONT_BODY_11
            cell.border = THIN_BORDER
        col += 1


def build_product_specs(ws):
    """Build Product Specs sheet shell — VeryHidden lookup table."""
    set_column_widths(ws, {"A": 3, "B": 20, "C": 14, "D": 14, "E": 14, "F": 14})

    apply_section_header(ws, 1, 2, 6, "Product Specifications", level="title")

    headers = ["Product Type", "AP (m²/m³)", "Wall Thickness (mm)",
               "Pitch (mm)", "Channel Type"]
    write_table_headers(ws, 3, 2, headers)

    # Placeholder rows
    for r in range(4, 9):
        for c in range(2, 7):
            cell = ws.cell(row=r, column=c)
            cell.font = FONT_BODY_11
            cell.border = THIN_BORDER


def build_backing_table(ws, table_def):
    """Build a hidden backing-table sheet with its Excel Table."""
    table_name = table_def["table_name"]
    columns = table_def["columns"]

    # Set reasonable column widths
    for i in range(len(columns)):
        ws.column_dimensions[get_column_letter(i + 1)].width = 16

    create_excel_table(ws, table_name, columns, header_row=1, col_start=1, data_rows=1)


# ═══════════════════════════════════════════════════════════════════════════════
# Named ranges
# ═══════════════════════════════════════════════════════════════════════════════

def create_named_ranges(wb):
    """Create all named ranges for the workbook."""

    # ── Control named ranges ─────────────────────────────────────────────
    ctrl_ws = wb["Control"]
    for i, (field, _) in enumerate(CONTROL_FIELDS):
        r = 3 + i
        ref = f"'{ctrl_ws.title}'!$C${r}"
        dn = DefinedName(field, attr_text=ref)
        wb.defined_names.add(dn)

    # ── Specifications named ranges ──────────────────────────────────────
    spec_ws = wb["Specifications"]
    for i, (name, _, _) in enumerate(SPEC_FIELDS):
        r = 4 + i
        ref = f"'{spec_ws.title}'!$C${r}"
        dn = DefinedName(name, attr_text=ref)
        wb.defined_names.add(dn)

    # ── Constants named ranges ───────────────────────────────────────────
    const_ws = wb["Constants"]
    for i, (name, _, _) in enumerate(CONSTANTS):
        r = 4 + i
        ref = f"'{const_ws.title}'!$C${r}"
        dn = DefinedName(name, attr_text=ref)
        wb.defined_names.add(dn)

    # ── GC_ named ranges (placeholders pointing to Geometry Calc) ────────
    gc_ws = wb["Geometry Calc"]
    for i, name in enumerate(GC_NAMED_RANGES):
        r = 61 + i
        ref = f"'{gc_ws.title}'!$C${r}"
        dn = DefinedName(name, attr_text=ref)
        wb.defined_names.add(dn)

    # ── Lists named ranges ───────────────────────────────────────────────
    lists_ws = wb["Lists"]
    col = 2
    for list_name, items in LISTS_DATA.items():
        col_letter = get_column_letter(col)
        end_row = 3 + len(items)
        ref = f"'{lists_ws.title}'!${col_letter}$4:${col_letter}${end_row}"
        dn = DefinedName(f"List_{list_name}", attr_text=ref)
        wb.defined_names.add(dn)
        col += 1


# ═══════════════════════════════════════════════════════════════════════════════
# Main build
# ═══════════════════════════════════════════════════════════════════════════════

def build():
    wb = Workbook()

    # Remove default sheet
    wb.remove(wb.active)

    # Create all sheets
    sheets = {}
    for name, visibility in SHEET_DEFS:
        ws = wb.create_sheet(title=name)
        if visibility == "hidden":
            ws.sheet_state = "hidden"
        elif visibility == "veryHidden":
            ws.sheet_state = "veryHidden"
        sheets[name] = ws

    # ── Build each sheet ─────────────────────────────────────────────────
    build_home(sheets["Home"])
    build_specifications(sheets["Specifications"])
    build_geometry_input(sheets["HC Geometry"], "HC")
    build_geometry_input(sheets["Plate Geometry"], "Plate")
    build_geometry_input(sheets["Corrugated Geometry"], "Corrugated")
    build_setup_summary(sheets["Setup Summary"])
    build_dashboard_shell(sheets["Activity"], "Activity")
    build_dashboard_shell(sheets["Conversion"], "Conversion")
    build_dashboard_shell(sheets["DP"], "DP")
    build_geometry_calc(sheets["Geometry Calc"])
    build_control(sheets["Control"])
    build_constants(sheets["Constants"])
    build_lists(sheets["Lists"])
    build_product_specs(sheets["Product Specs"])

    # Build backing tables
    for sheet_name, table_def in BACKING_TABLES.items():
        build_backing_table(sheets[sheet_name], table_def)

    # ── Named ranges ─────────────────────────────────────────────────────
    create_named_ranges(wb)

    # ── Save ─────────────────────────────────────────────────────────────
    wb.save(OUTPUT_FILE)
    print(f"Workbook saved to: {OUTPUT_FILE}")
    print(f"  Sheets: {len(wb.sheetnames)}")
    print(f"  Named ranges: {len(wb.defined_names)}")

    # Summary
    visible = [n for n, v in SHEET_DEFS if v == "visible"]
    hidden = [n for n, v in SHEET_DEFS if v == "hidden"]
    very_hidden = [n for n, v in SHEET_DEFS if v == "veryHidden"]
    print(f"\n  Visible ({len(visible)}):     {', '.join(visible)}")
    print(f"  Hidden ({len(hidden)}):      {', '.join(hidden)}")
    print(f"  VeryHidden ({len(very_hidden)}): {', '.join(very_hidden)}")

    tables = [t["table_name"] for t in BACKING_TABLES.values()]
    print(f"\n  Tables ({len(tables)}): {', '.join(tables)}")


if __name__ == "__main__":
    build()
