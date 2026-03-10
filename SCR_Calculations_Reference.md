# SCR Catalyst Test System — Calculations & Formulas Reference

**Extracted from Blueprint v1.4 • March 2026 • Confidential — Cormetech**

---

## 1. Geometry Engine

Only the geometry model matching the selected Sample Type is active. All models publish unified outputs: Avg Adjusted FFA (m²), Total Adjusted Area (m²), Active Layer Count, and Geometry Status.

### 1.1 Plate Geometry

| Item | Formula / Rule | Notes |
|------|---------------|-------|
| Avg Length (mm) | Group2 present: AVG(Group1 avg length, Group2 avg length); else Group1 avg length | Average of group averages, not pooled |
| Avg Width (mm) | Average of all non-blank Width A and Width B values across both groups | Widths only; excludes length/thickness |
| Avg Thickness (mm) | Group2 present: AVG(Group1 avg thickness, Group2 avg thickness); else Group1 avg thickness | Average of group averages |
| Plate FFA (m²) | 150 × AvgWidth / 1,000,000 | 150 mm = standard plate face |
| Plate Total Area (m²) | TotalPlates × AvgLength × AvgWidth × 2 / 1,000,000 | ×2 for two active faces |

### 1.2 Honeycomb Geometry

Supports up to 4 independent layers. Each layer can reference a different Product Type from the Product Specs table.

| Item | Formula / Rule | Notes |
|------|---------------|-------|
| AP Used | AP Override if entered; else ProductSpec.AP | |
| Calc Width A | WidthA if entered; else StandardWidth × (ActualCellsA / SpecCellsA) | |
| Calc Width B | WidthB if entered; else StandardWidth × (ActualCellsB / SpecCellsB) | |
| Avg Width | AVG(CalcWidthA, CalcWidthB) | |
| Available Cells | (ActualCellsA × ActualCellsB) − PluggedCells | Blank PluggedCells = 0 |
| FFA (m²) | CalcWidthA × CalcWidthB / 1,000,000 | |
| Volume (m³) | FFA × Length / 1,000 | |
| Area (m²) | AP_Used × Volume | |
| Adj FFA (m²) | FFA × AvailableCells / (ActualCellsA × ActualCellsB) | |
| Adj Volume (m³) | AdjFFA × Length / 1,000 | |
| Adj Area (m²) | AP_Used × AdjVolume | |

**Honeycomb Rollup**

| Rollup | Formula |
|--------|---------|
| Avg Adjusted FFA | Average of Adj FFA across layers where Status = OK |
| Total Adjusted Area | Sum of Adj Area across layers where Status = OK |
| Active Layer Count | Count of layers where Status = OK |

### 1.3 Corrugated Geometry

Supports up to 4 independent layers. Inputs per layer: SSA (m²/m³), Length, Width, Height (mm), Total Cells, Plugged Cells.

| Item | Formula / Rule | Notes |
|------|---------------|-------|
| Layer Volume (m³) | Length × Width × Height / 10⁹ | OK layers only |
| Layer Adj Area (m²) | SSA × Length × Width × Height / 10⁹ × (Total−Plugged)/Total | OK layers only |
| Layer Adj FFA (m²) | Width × Height / 10⁶ × (Total−Plugged)/Total | OK layers only |
| Layer Unadj FFA (m²) | Width × Height / 10⁶ | |
| Total Volume | Sum of layer Volume for OK layers | |
| Total Adj Area | Sum of layer Adj Area for OK layers | |
| Avg Adj FFA | Sum of layer Adj FFA / count(OK layers) | |
| Total FFA (unadj) | Sum of layer Unadj FFA for OK layers | |
| Active Layer Count | Count of OK layers | |

---

## 2. Flow & Injection Engine

Computed separately for Activity and Conversion modes. K and Cormetech-K use AV Used from this engine.

### 2.1 Flow Formulas

| Item | Formula / Rule | Notes |
|------|---------------|-------|
| If AV entered | Flow (Nm³/h) = AV × TotalAdjArea; UGS = Flow / (3600 × AvgAdjFFA) | |
| If UGS entered | Flow (Nm³/h) = UGS × 3600 × AvgAdjFFA; AV Used = Flow / TotalAdjArea | |
| Flow (scfm) | Flow (Nm³/h) × Nm3h_to_scfm | Default const 0.62210 |
| Flow (L/min) | Flow (Nm³/h) × 1000 / 60 | |

### 2.2 Injection Rate Formulas

| Item | Formula / Rule | Notes |
|------|---------------|-------|
| SO3 Injection | (((Flow_Lmin × SO3_target / 10⁶) / SO3_MolarVol_ft3 / SO3_Density) / (H2SO4_Pct / 100)) × SO3_Empirical | |
| NH3 Injection | (NOx_target × MR × Flow_Lmin / 10⁶) × NH3SO2_Empirical | Inactive when MR = 0 |
| SO2 Injection | (SO2_target × Flow_Lmin / 10⁶) × NH3SO2_Empirical | |
| Combustion NH3 | A × NOx² + B × NOx + C | Valid NOx 50–500 ppmvd |

---

## 3. Gas Validation

### 3.1 SO2 Validation

| Item | Formula / Rule | Notes |
|------|---------------|-------|
| FTIR SO2 dry | SO2_wet / (1 − H2O / 100) | |
| IC Vapor Pressure | VP_A × EXP(−VP_B / (MeterTemp + 273.15)) × VP_mmHgFactor | |
| IC Corrected Gas Vol | MeterVol × (273.15 / (273.15 + MeterTemp)) × ((BaroP − VaporP) / 760) | |
| IC SO2 dry (ppmvd) | IC_Result × (Dilution / CorrectedGasVol) × (MolarVol_STP / MW_H2SO4) | |
| OOS Flag | ABS(SO2_dry − Target) > Target × Tol_SO2 | |
| Conv Pre Avg | AVG of included Conv-mode SO2 dry where Stage = Pre | |
| Conv Post Avg | AVG of included Conv-mode SO2 dry where Stage = Post | |
| Conv Validation Value | AVG(Pre Avg, Post Avg) when both exist; else whichever exists | |
| Conv Test Avg | AVG of included Conv-mode SO2 dry where Stage = Test | |

### 3.2 SO3 Validation

| Item | Formula / Rule | Notes |
|------|---------------|-------|
| Vapor Pressure | VP_A × EXP(−VP_B / (MeterTemp + 273.15)) × VP_mmHgFactor | |
| Corrected Gas Vol | PullVol × (273.15 / (273.15 + MeterTemp)) × ((RoomP − VaporP) / 760) | |
| Dry Moles | CorrectedGasVol × MolesPerLiter_STP | |
| SO3 Moles | (IC_Result × Dilution / 10⁶) / MW_H2SO4 | |
| SO3 dry (ppmvd) | SO3_Moles / DryMoles × 10⁶ | |
| OOS Flag | ABS(SO3_dry − Target) > Target × Tol_SO3 | |
| Conv Average | AVG of included Conv-mode SO3 dry records | Used for 'Average' inlet SO3 source |

### 3.3 NH3 Validation

| Item | Formula / Rule | Notes |
|------|---------------|-------|
| FTIR NH3 dry | NH3_wet / (1 − H2O / 100) | |
| IC Vapor Pressure | VP_A × EXP(−VP_B / (MeterTemp + 273.15)) × VP_mmHgFactor | |
| IC Corrected Gas Vol | MeterVol × (273.15 / (273.15 + MeterTemp)) × ((BaroP − VaporP) / 760) | |
| IC NH3 dry (ppmvd) | IC_Result × (Dilution / CorrectedGasVol) × (MolarVol_STP / MW_NH3) | MW_NH3 = 17.031 g/mol |
| Target for OOS | MR × NOx_target (same mode) | |
| OOS Flag | Inlet only, MR > 0: ABS(NH3_dry − target) > Tol_MR × NOx_target | |

### 3.4 NOx Validation

| Item | Formula / Rule | Notes |
|------|---------------|-------|
| FTIR total | FTIR_NO + FTIR_NO2 | Both components required |
| FTIR dry | FTIR_total / (1 − H2O / 100) | |
| NOx dry used | FTIR dry if source = FTIR; else Analyzer value | |
| OOS Flag | ABS(NOx_dry_used − Target) > Target × Tol_NO | |

---

## 4. Pre-Test Validation

Applied to both Activity and Conversion modes. Conversion adds SO3 row; NH3/MR rows inactive when MR = 0.

| Parameter | Tolerance Rule |
|-----------|---------------|
| Temperature | Actual within Target ± Tol_Temp (°C absolute) |
| Flow (scfm) | Actual within Target × (1 ± Tol_Flow) |
| O2 | Actual within Target ± Tol_O2 (% absolute) |
| NOx | Actual within Target × (1 ± Tol_NO) |
| SO2 | Actual within Target × (1 ± Tol_SO2) |
| SO3 (Conv only) | Actual within Target × (1 ± Tol_SO3) |
| NH3 | Actual within Target ± (Tol_MR × NOx_target) |
| Molar Ratio | Actual within Target ± Tol_MR |

---

## 5. Activity Testing (NOx Removal)

### 5.1 Pass Formulas

| Item | Formula / Rule | Notes |
|------|---------------|-------|
| FTIR NOx total | FTIR_NO + FTIR_NO2 | When both present |
| FTIR NOx dry | FTIR_total / (1 − H2O / 100) | |
| NOx dry used | FTIR dry if source = FTIR; else Analyzer value | |

### 5.2 Activity Result-Set Formulas

| Item | Formula / Rule | Notes |
|------|---------------|-------|
| DeNOx | (NOx_In − NOx_Out) / NOx_In | Blank when NOx_In is 0 or blank |
| MR Actual (Inlet basis) | NH3 / NOx_In | |
| MR Actual (Outlet basis) | DeNOx + (NH3 / NOx_In) | |
| K | −AV_Used × LN(1 − DeNOx) | Blank if (1−DeNOx) ≤ 0 |
| H2O Used | Last non-blank: outlet passes → inlet passes → SO2 gas val → NH3 gas val → H2O ref | Fallback chain |
| H2O Corrected K | K × (H2O_Reference / H2O_Used)^(−0.05) | H2O_Ref default 18% |

**Cormetech K (Three-Branch Model)**

Cormetech K is only defined for MR Actual ≤ 1.4. Above that, return blank / not evaluable.

| Branch | Formula / Rule | Notes |
|--------|---------------|-------|
| Branch 1: MR ≤ 1 | 0.5 × AV_Used × LN(MR / ((MR − DeNOx) × (1 − DeNOx))) | NH3-limited regime |
| Branch 2: 1 < MR ≤ 1.4, T ≥ 335°C | AV_Used × (0.3 × LN(MR / ((MR − DeNOx) × (1 − DeNOx))) − 0.4 × LN(1 − DeNOx)) | |
| Branch 3: 1 < MR ≤ 1.4, T < 335°C | f(T) = EXP((8.087 − 5651) / (Temp + 273.15)); then AV_Used × (f(T) × LN(MR / ((MR − DeNOx) × (1 − DeNOx))) − (1 − 2×f(T)) × LN(1 − DeNOx)) | Temp-dependent blend |
| H2O Corr Cormetech K | Cormetech_K × (H2O_Reference / H2O_Used)^(−0.05) | Same correction as K |

### 5.3 Activity Steady-State Formulas

Evaluated on the most recent 4 included result sets (Use = Yes).

| Item | Formula / Rule | Notes |
|------|---------------|-------|
| K Mean | AVG(K1, K2, K3, K4) | |
| Normalized Slope | ABS((4×(1×K1 + 2×K2 + 3×K3 + 4×K4) − 10×SUM(K1:K4)) / 20) / K_Mean | Pass ≤ SS_Act_Trend_Pct (0.01) |
| StdDev | STDEV.S(K1, K2, K3, K4) | Pass ≤ K_Mean × SS_Act_StdDev_Pct (0.02) |
| Steady State | STEADY only when both Trend and StdDev pass | |

---

## 6. Conversion Testing (SO2 → SO3)

### 6.1 Outlet SO3 Sample Formulas

| Item | Formula / Rule | Notes |
|------|---------------|-------|
| Vapor Pressure | VP_A × EXP(−VP_B / (MeterTemp + 273.15)) × VP_mmHgFactor | |
| Corrected Gas Vol | PullVol × (273.15 / (273.15 + MeterTemp)) × ((RoomP − VaporP) / 760) | |
| Dry Moles | CorrectedGasVol × MolesPerLiter_STP | |
| SO3 Moles | (IC_Result × Dilution / 10⁶) / MW_H2SO4 | |
| Outlet SO3 (ppmvd) | SO3_Moles / DryMoles × 10⁶ | |

### 6.2 Conversion Result-Set Formulas

| Item | Formula / Rule | Notes |
|------|---------------|-------|
| Difference (ppm) | Outlet_SO3 − Inlet_SO3 | |
| Conversion % | Difference / SO2_Used | Blank if SO2_Used is 0 or blank |

### 6.3 Conversion Steady-State Formulas

Thresholds scale with average SO2 because conversion measurements are noisier at lower SO2.

| Item | Formula / Rule | Notes |
|------|---------------|-------|
| Conv Mean | AVG(Conv1, Conv2, Conv3, Conv4) | |
| SO2 Avg | AVG(SO2_1, SO2_2, SO2_3, SO2_4) | |
| Normalized Slope | ABS((4×(1×Conv1+2×Conv2+3×Conv3+4×Conv4) − 10×SUM) / 20) / ABS(Conv_Mean) | Blank if Conv_Mean = 0 |
| Slope Threshold | MIN(SS_Conv_Trend_Coef × (1000/SO2_Avg), SS_Conv_Trend_Cap) | Defaults 0.01, 0.05 |
| StdDev | STDEV.S(Conv1, Conv2, Conv3, Conv4) | |
| StdDev Threshold | MIN(SS_Conv_StdDev_Coef × |Conv_Mean| × (1000/SO2_Avg), SS_Conv_StdDev_Cap × |Conv_Mean|) | Defaults 0.10, 0.50 |
| Steady State | STEADY only when both Trend and StdDev pass | |

---

## 7. Differential Pressure Validation

| Item | Formula / Rule | Notes |
|------|---------------|-------|
| DP Total | Sum of 4 station readings | Blank if all blank |
| % Theory | (FlowStraightenerDP × 100) / TheoryDP | |
| Pass / Fail | PASS when % Theory within 100 ± Tol_DP_Pct | Do not divide threshold by 100 again |

---

## 8. Utility Calculators

### 8.1 Slip Prediction Model

| Item | Formula / Rule | Notes |
|------|---------------|-------|
| Calc1 | EXP(2 × K / AV) | Blank if AV ≤ 0 |
| Calc2 | −(1 + MR) × Calc1 | |
| Calc3 | (Calc1 − 1) × MR | |
| Predicted Slip | NOx_In × (MR − ((−Calc2 − SQRT(MAX(Calc2² − 4×Calc1×Calc3, 0))) / (2×Calc1))) | Clamp negative discriminant to 0 |
| Predicted Outlet NOx | NOx_In × (1 − MR) + PredictedSlip | |
| Predicted DeNOx | (NOx_In − PredictedOutletNOx) / NOx_In | |

### 8.2 SO3 Collection Volume Estimator

| Item | Formula / Rule | Notes |
|------|---------------|-------|
| Pull Volume (L) | 50 × ((−0.000005×T²) − (0.00000003×T) + 1.0002) × Dilution / MW_H2SO4 × MolarVol_STP / SO3_target | EPA Method 8A helper |

### 8.3 Combustion NH3 Estimator

| Item | Formula / Rule | Notes |
|------|---------------|-------|
| Combustion NH3 | A × NOx² + B × NOx + C | Valid for NOx 50–500 ppmvd |

---

## 9. Physical Constants

| Constant | Default Value | Notes |
|----------|--------------|-------|
| VP_A | 2,229,000,000 | Water vapor pressure coefficient |
| VP_B | 5,385 | Water vapor pressure coefficient |
| VP_mmHgFactor | 0.750062 mmHg/Pa | Pressure conversion |
| MolesPerLiter_STP | 0.04462 mol/L | Reciprocal molar volume at STP |
| MW_H2SO4 | 96.06 g/mol | |
| MW_NH3 | 17.031 g/mol | |
| STP_T_K | 273.15 K | |
| STP_P_mmHg | 760 mmHg | |
| MolarVol_STP | 22.4914 L/mol | |
| Nm3h_to_scfm | 0.62210 | scf@60°F per Nm³@0°C |
| SO3_MolarVol_ft3 | 28.315 ft³/mol | |
| SO3_Density | 0.01870 | SO3 density/conversion factor |
| H2SO4_Pct | 20% | Default acid conc for SO3 injection |
| SO3_Empirical | 0.8275 | SO3 injection empirical factor |
| NH3SO2_Empirical | 0.8333 | NH3/SO2 injection empirical factor |
| CombNH3_A | 2.6e-06 | Combustion NH3 quadratic coeff |
| CombNH3_B | 0.0056 | Combustion NH3 linear coeff |
| CombNH3_C | −0.4328 | Combustion NH3 constant |

---

## 10. Configurable Parameters (Tolerances & Thresholds)

| Parameter | Default | Notes |
|-----------|---------|-------|
| Tol_Temp | 3 °C (absolute) | Temperature tolerance |
| Tol_Flow | 0.01 (relative fraction) | Flow tolerance |
| Tol_O2 | 0.2 % (absolute) | O2 tolerance |
| Tol_NO | 0.02 (relative fraction) | NOx tolerance |
| Tol_SO2 | 0.04 (relative fraction) | SO2 tolerance |
| Tol_SO3 | 0.20 (relative fraction) | SO3 tolerance |
| Tol_MR | 0.02 (absolute) | Molar ratio tolerance |
| Tol_DP_Pct | 5 % (absolute) | DP % theory deviation around 100 |
| SS_Act_StdDev_Pct | 0.02 (fraction of mean) | Activity SS std-dev threshold |
| SS_Act_Trend_Pct | 0.01 (slope/mean) | Activity SS trend threshold |
| SS_Conv_StdDev_Coef | 0.10 × (1000/SO2) | Conversion SS std-dev coeff |
| SS_Conv_StdDev_Cap | 0.50 × \|mean\| | Conversion SS std-dev cap |
| SS_Conv_Trend_Coef | 0.01 × (1000/SO2) | Conversion SS trend coeff |
| SS_Conv_Trend_Cap | 0.05 (slope/mean) | Conversion SS trend cap |
| H2O_Reference | 18% | Reference H2O for corrected K calcs |

---

## 11. AV / UGS Mutual-Derivation Rules

| State | Behavior |
|-------|----------|
| AV entered, UGS blank | Compute UGS from AV and geometry outputs |
| UGS entered, AV blank | Compute AV from UGS and geometry outputs |
| Both entered | Warn and use AV as authoritative |
| Neither entered | Withhold flow/injection calculations |

---

## 12. H2O Used Fallback Chain

For H2O-corrected K calculations, H2O Used is the last non-blank value from this priority chain:

| Priority | Source |
|----------|--------|
| 1 (highest) | Activity outlet passes |
| 2 | Activity inlet passes |
| 3 | Gas Validation SO2 records |
| 4 | Gas Validation NH3 records |
| 5 (fallback) | H2O_Reference (configurable, default 18%) |
