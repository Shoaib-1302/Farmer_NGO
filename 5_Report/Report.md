# 🌳 Makkala Vikasa & YIDS — Plantation Report

*Generated July 04, 2026 · Source: `2_Cleaned_Data/Cleaned_Master.xlsx`*

This report follows the same two-part layout as a standard MRV (monitoring, reporting &
verification) dashboard — an **Organization Report** (portfolio-level totals) and a
**Monitoring / QA Report** (this batch's records, broken down by species and quality
status). Where the source data doesn't contain a field the reference layout expects
(e.g. DBH, standing-tree counts), it is shown as **N/A** rather than estimated —
consistent with how the reference dashboard itself handles missing pipeline data.

---

## 1. Organization Report

| | | | | | |
|---|---|---|---|---|---|
| **6,483 ac** (2,624 Ha) <br>Total Acreage (KML-measured) | **4,513** <br>Number of Plot Records | **N/A** <br>Trees Contracted / Planted | **N/A** <br>Standing Trees | **21** <br>Tree Species Recorded | **N/A** <br>Avg. Survival Rate |
| **N/A** <br>Avg Trees per Ha | **N/A** <br>Avg Trees per Plot | **N/A** <br>DBH Trees | **N/A** <br>Small Trees | **2** <br>Number of Projects (NGOs) | **N/A** <br>Avg. Mortality Rate |
| **3,946** <br>Plots with Matched KML | **567** <br>Plots Missing KML | **N/A** <br>DBH Trees (verified) | **N/A** <br>Small Trees (verified) | **62** <br>Management Units (Blocks) | **81.4 tCO₂e/yr*** <br>Estimated Sequestration |

<sub>*Illustrative estimate (species count × literature-average kg CO₂/tree/year) — the
source data has no DBH/girth measurements, so no field-measured carbon figure exists yet.
Equivalent to **0.031 tCO₂e/Ha/yr** at current planting density. See §4 for the
per-species assumptions.</sub>

**Why so many N/A fields:** the reference layout is built for a system that tracks tree
counts, DBH, and survival over monitoring cycles. The two NGO Kobo exports provided here
capture *plot-level* plantation records (one row per farmer plot, with a species list and
a claimed vs. GPS-measured area) but do not include individual tree counts, girth
measurements, or repeat-visit survival tracking. Everything markable from the source data
**is** populated above; everything else is honestly flagged rather than backfilled.

### Organization Report — by NGO

| NGO | Records | Farmers | KML Match % | Claimed Area (ac) | KML-Measured Area (ac) | Reliability Score |
|---|---|---|---|---|---|---|
| NGO_1 (Makkala Vikasa) | 1,563 | 1,083 | 63.7% | 12,875.9 | 1,773.1 | **73 / 100** |
| NGO_2 (YIDS) | 2,950 | 1,845 | 100.0% | 10,375.9 | 4,710.4 | **88 / 100** |
| **Total** | **4,513** | **2,928** | **87.4%** | **23,251.8** | **6,483.5** | — |

---

## 2. Monitoring / QA Report — Full Dataset

**Cycle:** Single-batch review (no repeat monitoring cycles in source data)
**Sampling rate:** 100% (all submitted records reviewed, not a 10% sample)

| Target | Total | Clean ("Approved") | Warning ("To Validate") | Critical ("Rejected") |
|---|---|---|---|---|
| 4,513 | **4,513** | 🟢 **634** | 🟡 **2,561** | 🔴 **1,318** |

| Metric | Value |
|---|---|
| KML Match Rate | 87.4% |
| Clean Rate | 14.0% |
| Area Mismatch Rate (>15%) | 60.9% |
| Plots w/ Geometry Overlap | 210 |
| Suspicious Duplicate-GPS Records | 12 |

*(Severity tiers replace the reference layout's Approved/Rejected/To-Validate workflow —
see `6_Report/Methodology.md` §3 for exactly which checks map to which tier.)*

### Species Distribution & Estimated Carbon

| Species | Plot Count | Distribution | Est. tCO₂e/yr* |
|---|---:|---|---:|
| Mango | 1,728 | `████████████████████████` | 34.560 |
| Arecanut | 1,588 | `██████████████████████░░` | 12.704 |
| Coconut | 1,472 | `████████████████████░░░░` | 22.080 |
| Other / Unspecified | 227 | `███░░░░░░░░░░░░░░░░░░░░░` | 2.724 |
| Guava | 187 | `███░░░░░░░░░░░░░░░░░░░░░` | 1.870 |
| Pomegranate | 144 | `██░░░░░░░░░░░░░░░░░░░░░░` | 0.864 |
| Mahogany | 114 | `██░░░░░░░░░░░░░░░░░░░░░░` | 2.508 |
| Lemon | 71 | `█░░░░░░░░░░░░░░░░░░░░░░░` | 0.568 |
| Jamun | 61 | `█░░░░░░░░░░░░░░░░░░░░░░░` | 1.098 |
| Jackfruit | 51 | `█░░░░░░░░░░░░░░░░░░░░░░░` | 1.020 |
| RedSandal | 27 | `░░░░░░░░░░░░░░░░░░░░░░░░` | 0.405 |
| Butterfruit | 19 | `░░░░░░░░░░░░░░░░░░░░░░░░` | 0.285 |
| Oak | 15 | `░░░░░░░░░░░░░░░░░░░░░░░░` | 0.300 |
| WoodApple | 8 | `░░░░░░░░░░░░░░░░░░░░░░░░` | 0.096 |
| Rosewood | 5 | `░░░░░░░░░░░░░░░░░░░░░░░░` | 0.090 |
| Mugnahalli | 5 | `░░░░░░░░░░░░░░░░░░░░░░░░` | 0.060 |
| Gooseberry | 4 | `░░░░░░░░░░░░░░░░░░░░░░░░` | 0.040 |
| Tamarind | 4 | `░░░░░░░░░░░░░░░░░░░░░░░░` | 0.088 |
| Teak | 3 | `░░░░░░░░░░░░░░░░░░░░░░░░` | 0.060 |
| Bamboo | 2 | `░░░░░░░░░░░░░░░░░░░░░░░░` | 0.010 |
| Palm | 1 | `░░░░░░░░░░░░░░░░░░░░░░░░` | 0.010 |
| **Total** | | | **81.44** |

<sub>*Species counts are plots where that species is present, not tree counts (source
data doesn't record per-species tree quantity). Carbon figures use the same illustrative
kg CO₂/tree/year assumptions as the dashboard's Species & Carbon tab — confidence: **Low
(proxy-based)**, since no DBH/girth was measured in the field.</sub>

### Top 5 Problem Villages (by flagged-record count)

| Village | Total Records | Flagged | Issue Rate | Main Problem |
|---|---:|---:|---:|---|
| Sanganahalli | 46 | 43 | 93.5% | Area Mismatch |
| Mudimadu | 73 | 67 | 91.8% | Area Mismatch |
| Kallenahalli | 50 | 45 | 90.0% | Area Mismatch |
| Seebi | 43 | 37 | 86.0% | Area Mismatch |
| Yaramadanahalli | 75 | 61 | 81.3% | Area Mismatch |

Full list of 20 in `2_Cleaned_Data/Cleaned_Master.xlsx → Top_Problem_Villages`.

---

## 3. Reading this report alongside the reference layout

| Reference field | Where it lives here |
|---|---|
| Target / Total / Approved / Rejected / To Validate | §2 table — mapped to Clean/Warning/Critical severity, not a literal approval workflow |
| Survival Rate / Mortality Rate | Not available — no repeat-visit tree-survival tracking in source data |
| Species bar chart + tCO₂e sequestered | §2 species table (merged into one table instead of two side-by-side panels) |
| Total Acreage / Number of Plots / Avg per Ha / per Plot | §1 Organization Report |
| DBH Trees / Small Trees | Not available — no girth/height measurements in source data |
| Number of Projects / Management Units | §1 — mapped to NGOs and standardized Blocks respectively |
| Projected vs. Actual tCO₂e | Only one figure exists here (illustrative estimate) — there's no separate "projected" baseline in the source data to compare it against |

## 4. Carbon assumption table

| Species | kg CO₂ / tree / year (assumed) |
|---|---:|
| Mahogany, Tamarind | 22 |
| Mango, Jackfruit, Oak, Teak | 20 |
| Jamun, Rosewood | 18 |
| Coconut, RedSandal, Butterfruit | 15 |
| Mugnahalli, Other/Unspecified, WoodApple | 12 |
| Guava, Gooseberry, Palm | 10 |
| Arecanut, Lemon | 8 |
| Bamboo | 5 |
| Pomegranate | 6 |

Replace this table with field-measured DBH-based allometric equations once girth data is
collected, for a defensible measured carbon figure instead of this proxy.

---
*Full row-level detail behind every number in this report is in
`2_Cleaned_Data/Cleaned_Master.xlsx`, `GIS_Validation_Report.xlsx`, and
`Data_Quality_Report.xlsx`. Methodology and thresholds: `6_Report/Methodology.md`.*
