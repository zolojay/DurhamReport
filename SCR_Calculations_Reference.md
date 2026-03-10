# SCR Calculations Reference

**Revision 1 — March 2026**

This document is the single authoritative reference for every engineering formula
used in the SCR Catalyst Test Report workbook. Every formula in Geometry Calc,
Activity, Conversion, and DP must match this reference exactly.

---

## 1. Geometry Calculations

### 1.1 Honeycomb (HC) Geometry

**Inputs per layer** (from HC Geometry sheet):
- Product Type → lookup AP from Product Specs
- AP Override (m²/m³) — if provided, overrides lookup
- Length (mm)
- Cells A (count)
- Cells B (count)
- Width A (mm)
- Width B (mm)
- Plugged Cells (count)

**Derived per layer:**

```
Resolved_AP = IF(AP_Override <> "", AP_Override, VLOOKUP(ProductType, ProductSpecs, AP_col))
Total_Cells = Cells_A × Cells_B
Open_Cells = Total_Cells − Plugged_Cells
Pitch = Width_A / Cells_A                          (mm)
Wall_Thickness = lookup from Product Specs or default 1.0 mm
Channel_Width = Pitch − Wall_Thickness              (mm)
Hydraulic_Diameter = Channel_Width                   (mm, square channels)
FFA_single = (Channel_Width / Pitch)²               (dimensionless, fraction)
Cross_Section_Area = (Width_A / 1000) × (Width_B / 1000)   (m²)
Adjusted_FFA = FFA_single × (Open_Cells / Total_Cells)     (dimensionless)
Layer_Area = Cross_Section_Area × Adjusted_FFA              (m², free-flow area)
Layer_Volume = Cross_Section_Area × (Length / 1000)         (m³, total volume)
Layer_Surface_Area = Resolved_AP × Layer_Volume             (m², catalytic surface)
```

**Layer active rule:** A layer is active if Length > 0 AND Cells_A > 0 AND Cells_B > 0.

**Rollup across layers:**
```
Active_Layers = COUNT of active layers
Avg_Adjusted_FFA = AVERAGE(Adjusted_FFA) across active layers    (m²)
Total_Adjusted_Area = SUM(Layer_Surface_Area) across active layers (m²)
```

### 1.2 Plate Geometry

**Inputs** (from Plate Geometry sheet):
- Total Plates (integer)
- Box 1: 13 plates × (Length, Thickness, Width A, Width B) in mm
- Box 2: 13 plates × (Length, Thickness, Width A, Width B) in mm

**Derived:**
```
Avg_Length = AVERAGE(all non-blank lengths across Box1 + Box2)    (mm)
Avg_Thickness = AVERAGE(all non-blank thicknesses)                (mm)
Avg_WidthA = AVERAGE(all non-blank Width A values)                (mm)
Avg_WidthB = AVERAGE(all non-blank Width B values)                (mm)

Plate_Spacing = Avg_Thickness                                     (mm, approx)
Hydraulic_Diameter = 2 × Plate_Spacing                            (mm)
Channel_Height = Avg_WidthA / 1000                                (m)

Active_Layers = Total_Plates − 1    (channels between plates)
Single_Channel_Area = (Avg_Length / 1000) × (Channel_Height)      (m²)
FFA = Active_Layers × (Plate_Spacing / 1000) × (Avg_WidthB / 1000)  (m²)
Total_Surface_Area = 2 × Active_Layers × Single_Channel_Area     (m², both sides)

Avg_Adjusted_FFA = FFA                                            (m²)
Total_Adjusted_Area = Total_Surface_Area                          (m²)
```

### 1.3 Corrugated Geometry

**Inputs per layer** (from Corrugated Geometry sheet):
- SSA (m²/m³) — specific surface area
- Length (mm)
- Width (mm)
- Height (mm)
- Total Cells (count)
- Plugged Cells (count)

**Derived per layer:**
```
Open_Cells = Total_Cells − Plugged_Cells
Open_Fraction = Open_Cells / Total_Cells
Layer_Volume = (Length / 1000) × (Width / 1000) × (Height / 1000)   (m³)
Layer_Surface_Area = SSA × Layer_Volume × Open_Fraction              (m²)
Cross_Section = (Width / 1000) × (Height / 1000)                    (m²)
Layer_FFA = Cross_Section × Open_Fraction                            (m²)
```

**Layer active rule:** A layer is active if Length > 0 AND Width > 0 AND Height > 0.

**Rollup:**
```
Active_Layers = COUNT of active layers
Avg_Adjusted_FFA = AVERAGE(Layer_FFA) across active layers    (m²)
Total_Adjusted_Area = SUM(Layer_Surface_Area)                 (m²)
Active_Volume = SUM(Layer_Volume) across active layers        (m³)
```

---

## 2. Flow and Injection

### 2.1 Flow Derivation (AV / UGS mutual derivation)

The flow rule from the PRD:
- AV entered + UGS blank → derive UGS and Flow from AV
- UGS entered + AV blank → derive AV and Flow from UGS
- Both entered → warn, use AV as authoritative
- Neither entered → withhold flow/injection calculations

**From AV (area velocity):**
```
AV is in Nm/h (normal meters per hour)
Flow_Nm3h = AV × Total_Adjusted_Area                 (Nm³/h)
UGS_derived = Flow_Nm3h / Avg_Adjusted_FFA / 3600    (Nm/s)
```

Wait — AV in SCR context is defined as:
```
AV = Flow / Total_Adjusted_Area                       (Nm/h)
```
So:
```
Flow_Nm3h = AV × Total_Adjusted_Area                 (Nm³/h)
```

**From UGS (gas superficial velocity):**
```
UGS is in Nm/s
Flow_Nm3h = UGS × Avg_Adjusted_FFA × 3600            (Nm³/h)
AV_derived = Flow_Nm3h / Total_Adjusted_Area          (Nm/h)
```

**Flow unit conversion:**
```
Flow_scfm = Flow_Nm3h × Nm3_to_scfm_Factor           (scfm)
```
Where `Nm3_to_scfm_Factor` = 0.58858 (from Constants sheet).

**Flow status:**
```
"AV → Flow"       when AV entered, UGS blank
"UGS → Flow"      when UGS entered, AV blank
"AV (both given)"  when both entered
"Awaiting input"   when neither entered
```

### 2.2 Injection Rate Calculations

All injection rates use the resolved flow (Nm³/h) for the corresponding test type.

**SO₃ Injection (mL/min):**
```
Target_SO3_ppmvd = SO3 condition from Home (ppmvd)
SO3_Inj_mLmin = (Target_SO3_ppmvd / 1e6) × Flow_Nm3h × (1e6 / 60)
              = Target_SO3_ppmvd × Flow_Nm3h / 60
```
Simplified: SO3_Inj = Target_SO3 × Flow_Nm3h / 60

**NH₃ Injection (mL/min):**
```
Target_NH3_ppmvd = MR × Target_NOx     (derived from molar ratio and inlet NOx)
  — OR directly from conditions if NH3 is specified
NH3_Inj_mLmin = Target_NH3_ppmvd × Flow_Nm3h / 60
```
When MR > 0 and NOx target is given:
```
NH3_target = MR × NOx_target            (ppmvd)
NH3_Inj = NH3_target × Flow_Nm3h / 60   (mL/min)
```

**SO₂ Injection (mL/min):**
```
Target_SO2_ppmvd = SO2 condition from Home (ppmvd)
SO2_Inj_mLmin = Target_SO2_ppmvd × Flow_Nm3h / 60
```

---

## 3. Dry Correction

When FTIR reports wet-basis values:
```
Dry_value = Wet_value / (1 − H2O_fraction)
```
Where H2O_fraction is the FTIR-reported H₂O as a decimal (e.g., 8% → 0.08).

---

## 4. NOx Calculations

### FTIR path:
```
NOx_wet = FTIR_NO_Wet + FTIR_NO2_Wet        (ppmv wet)
DryNOxUsed = NOx_wet / (1 − FTIR_H2O_Pct)  (ppmvd)
```

### NOx Analyzer path:
```
DryNOxUsed = NOxAn_NO_Dry + NOxAn_NO2_Dry   (ppmvd, already dry)
```

---

## 5. DeNOx and K-value (Activity)

```
DeNOx_pct = (1 − Outlet_NOx / Inlet_NOx) × 100
K = −ln(1 − DeNOx_pct / 100) × AV_used
  = −ln(Outlet_NOx / Inlet_NOx) × AV_used
```

Where AV_used = GC_AV_Act (resolved area velocity for activity).

### Cormetech K (NH₃-corrected):
```
Corm_K = K × (1 + MR_actual)    — simplified Cormetech correction
```
Where MR_actual = NH3_used / Inlet_NOx.

### H₂O-corrected K:
```
H2O_K = K × (H2O_Reference / H2O_used)
```
Where H2O_Reference = 0.18 (from Constants), H2O_used follows the fallback chain.

---

## 6. H₂O Fallback Chain (Activity)

Priority order for H2O_used:
1. Activity outlet pass FTIR H₂O
2. Activity inlet pass FTIR H₂O
3. SO₂ gas validation FTIR H₂O
4. NH₃ gas validation FTIR H₂O
5. H2O_Reference (0.18)

---

## 7. Conversion (SO₂→SO₃)

```
Difference = Outlet_SO3 − Inlet_SO3         (ppmvd)
Conversion_pct = Difference / SO2_used × 100 (%)
```

Where:
- Outlet_SO3 = selected outlet SO₃ record value
- Inlet_SO3 = resolved from source (average or specific record)
- SO2_used = resolved from source (validation, test average, or specific record)

---

## 8. Steady-State Checks

### Activity:
```
K_Mean = AVERAGE(K values for Use=Yes rows)
K_StdDev = STDEV(K values for Use=Yes rows)
Normalized_Slope = slope of K vs row index, normalized by K_Mean
Steady_State = (K_StdDev ≤ Spec_SS_K_StdDev_Max) AND
               (|Normalized_Slope| ≤ Spec_SS_NormSlope_Max) AND
               (Points ≥ Spec_SS_MinPoints)
```

### Conversion:
```
Conv_Mean = AVERAGE(Conv% for Use=Yes rows)
Conv_StdDev = STDEV(Conv% for Use=Yes rows)
Normalized_Slope = slope of Conv% vs row index, normalized by Conv_Mean
Steady_State = (Conv_StdDev ≤ Spec_SS_Conv_StdDev_Max) AND
               (|Normalized_Slope| ≤ Spec_SS_Conv_NormSlope_Max) AND
               (Points ≥ Spec_SS_MinPoints)
```

---

## 9. Physical Constants (from Constants sheet)

| Name | Value | Unit |
|------|-------|------|
| H2O_Reference | 0.18 | fraction (18%) |
| MolarMass_NO | 30.01 | g/mol |
| MolarMass_NO2 | 46.01 | g/mol |
| MolarMass_NH3 | 17.03 | g/mol |
| MolarMass_SO2 | 64.07 | g/mol |
| MolarMass_SO3 | 80.06 | g/mol |
| STP_Temp_K | 273.15 | K |
| STP_Pressure_mmHg | 760.0 | mmHg |
| IdealGasVol_L | 22.414 | L/mol |
| Nm3_to_scfm_Factor | 0.58858 | conversion |

---

## 10. SO₃ Dry Calculation

```
CorrectedGasVol = PullVol_L × (STP_Temp_K / (MeterTemp_C + 273.15)) × (RoomP_mmHg / STP_Pressure_mmHg)
DryMoles = CorrectedGasVol / IdealGasVol_L
SO3Moles = (IC_Result × Dilution) / (MolarMass_SO3 × 1000)
DrySO3_ppmv = (SO3Moles / DryMoles) × 1e6
```

---

## 11. DP Calculations

```
DPTotal = DP_S2 + DP_S3 + DP_S4 + DP_S5
PctTheory = (DPTotal / TheoryDP) × 100
Status:
  - "Pass" if PctTheory ≤ Spec_DP_PctTheory_Warning
  - "Warning" if PctTheory > Warning AND ≤ Fail
  - "Fail" if PctTheory > Spec_DP_PctTheory_Fail
```

---

## 12. NH₃ Dry Calculation

### FTIR path:
```
DryNH3 = FTIR_NH3_Wet / (1 − FTIR_H2O_Pct)    (ppmvd)
```

### IC path:
```
CorrectedGasVol = MeterVol_L × (STP_Temp_K / (MeterTemp_C + 273.15)) × (BaroP_mmHg / STP_Pressure_mmHg)
DryMoles = CorrectedGasVol / IdealGasVol_L
NH3Moles = (IC_Result × Dilution) / (MolarMass_NH3 × 1000)
DryNH3 = (NH3Moles / DryMoles) × 1e6    (ppmvd)
```

---

## 13. SO₂ Dry Calculation

### FTIR path:
```
DrySO2 = FTIR_SO2_Wet / (1 − FTIR_H2O_Pct)    (ppmvd)
```

### IC path:
```
CorrectedGasVol = MeterVol_L × (STP_Temp_K / (MeterTemp_C + 273.15)) × (BaroP_mmHg / STP_Pressure_mmHg)
DryMoles = CorrectedGasVol / IdealGasVol_L
SO2Moles = (IC_Result × Dilution) / (MolarMass_SO2 × 1000)
DrySO2 = (SO2Moles / DryMoles) × 1e6    (ppmvd)
```
