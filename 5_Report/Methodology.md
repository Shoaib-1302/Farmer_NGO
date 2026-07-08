# Methodology — Makkala Vikasa & YIDS Farmer Plantation Data

## 1. Data sources
- **NGO_1 (Makkala Vikasa):** 1,563 Kobo Collect records (`1_Raw_Data/NGO_1.xlsx`) +
  996 KML field-boundary polygons (`1_Raw_Data/NGO_1.kmz`).
- **NGO_2 (YIDS):** 2,950 Kobo Collect records (`1_Raw_Data/NGO_2.xlsx`) + 2,950 KML
  field-boundary polygons (`1_Raw_Data/NGO_2.kmz`).
- Each Excel row carries a `_uuid` field that is expected to match the `<name>` of one
  `<Placemark>` in the corresponding NGO's KML file.

## 2. Pipeline

```
   Kobo Collect          Kobo Collect
   (NGO_1 .xlsx)         (NGO_1 .kmz)
        │                     │
        ▼                     ▼
 ┌─────────────────┐   ┌──────────────────┐
 │ data_cleaning.py │   │  kml_matching.py │◄── unzips .kmz, parses polygons
 │ UUID/phone/name  │──▶│  UUID match,     │    (geodesic area + centroid),
 │ validation,      │   │  GPS validation, │    suspicious-GPS-cluster scan
 │ error-note       │   │  area check      │
 │ standardization  │   └──────────────────┘
 └─────────────────┘             │
        │                        ▼
        │              ┌──────────────────────┐
        └─────────────▶│ overlap_analysis.py  │◄── STRtree spatial index for
                        │ polygon overlaps,    │    O(n log n) overlap detection
                        │ 3-tier severity,     │
                        │ reliability score,   │
                        │ priority actions     │
                        └──────────────────────┘
                                  │
                                  ▼
                     2_Cleaned_Data/*.xlsx
                                  │
                                  ▼
                     dashboard_app.py (Streamlit)
```

Text form: **Kobo → Raw Excel/KML → Cleaning → GIS Validation → Overlap & Severity
Scoring → QA Flags → Dashboard**. Each of the 3 pipeline scripts reads the previous
script's output (via a small interim cache) and hands off a richer dataset to the next.

## 3. Pipeline stages

### Stage 1 — `data_cleaning.py`
- Loads both Excel exports into one common schema (uniform column names across NGOs).
- **UUID checks:** format validated against the standard UUID v4 pattern; duplicates
  detected by exact match.
- **Phone validation:** Indian 10-digit mobile format (must start 6–9); country-code
  prefixes stripped before validation; duplicate valid numbers flagged across records.
- **Name standardization:** Block / GP / Village names are normalized (lowercased,
  whitespace-collapsed) then fuzzy-clustered using `difflib.SequenceMatcher` at a 0.84
  similarity threshold — spelling variants of the same place collapse into one cluster,
  and the most frequent spelling in each cluster becomes the suggested canonical name.
- **Missing-value audit** on the fields that matter most for downstream analysis.
- **Error-note standardization:** the NGOs had already logged 725 distinct free-text
  QA phrasings. A keyword-rule classifier maps each note to one of ~9 standard
  categories, written to `Data_Quality_Report.xlsx`.
- **Standardization-impact tracking:** before/after counts for every standardization
  step are captured for the `Standardization_Impact` sheet (see §5 below).

### Stage 2 — `kml_matching.py`
- Unzips each `.kmz` and parses every `<Placemark>`'s `<coordinates>` into a Shapely
  `Polygon`. Area & centroid are computed geodesically on the WGS84 ellipsoid
  (`pyproj.Geod`), not a flat-plane approximation.
- **Geometry validity:** self-intersecting polygons are detected and auto-repaired
  with `.buffer(0)` for area calculation, but flagged in the output.
- **UUID matching:** Excel UUID ↔ KML placemark name, joined per-NGO.
- **GPS-vs-KML cross-check:** geodesic distance between the Kobo-recorded lat/long and
  the KML polygon's centroid; flagged if > 300 m.
- **Area cross-check:** percentage difference between KML-measured area and the
  farmer/surveyor-claimed area; flagged if magnitude > 15%.
- **Region sanity check:** GPS points outside a Karnataka-region bounding box
  (10–20°N, 72–80°E) are flagged as likely column-shift export errors.
- **Suspicious duplicate-GPS clusters (new):** records are rounded to ~1m coordinate
  precision and grouped; a group of ≥3 records from ≥2 *different* farmer names at the
  same location — after excluding already-invalid out-of-region points — is flagged as
  a likely bulk-copied/lazy pin-drop rather than a genuine coincidence. Written to
  `GIS_Validation_Report.xlsx → Suspicious_GPS_Clusters`.

### Stage 3 — `overlap_analysis.py`
- **Overlap detection:** a `shapely.STRtree` spatial index over each NGO's polygons,
  testing candidate pairs for intersection; flags any pair whose overlap covers ≥10%
  of either plot's area (roughly O(n log n) rather than a naive O(n²) all-pairs scan).
- **3-tier severity engine (new):** every check across all three scripts collapses
  into one of 11 operational issue categories, then into a severity tier:

  | Tier | Categories |
  |---|---|
  | 🔴 **Critical** | UUID Issue (invalid/duplicate), Missing KML, GPS Issue (missing/out-of-region/>300m from KML centroid), Geometry Overlap ≥20%, Suspicious GPS Duplicate |
  | 🟡 **Warning** | Phone Issue, Area Mismatch (>15%), Name Standardization Needed, Year Planted Issue, Geometry Self-Intersecting, Geometry Overlap 10–20% |
  | 🟢 **Clean** | none of the above |

  A record can carry several issue categories at once (`issue_categories` column lists
  every one that applies); `severity` takes the worst tier present.
- **NGO-wise Data Reliability Score (new):** `100 − weighted penalty`, where the
  penalty combines invalid-GPS %, duplicate-UUID %, missing-KML %, area-mismatch %,
  overlap %, and invalid-phone % (weights: 0.25 / 0.20 / 0.20 / 0.15 / 0.10 / 0.10).
- **Top problematic villages (new):** villages ranked by flagged-record count, with
  each village's single most common issue category surfaced.
- **Priority actions (new):** a ranked, quantified action list generated directly from
  the numbers above (see §7).

### Dashboard — `dashboard_app.py` / `5_Dashboard/app.py`
- Streamlit app reading `Cleaned_Master.xlsx` directly. 8 tabs: Overview, Data Quality,
  GIS Validation, Species & Carbon, Map, Villages, Farmer Table, Priority Actions.
- **Map:** a real zoomable/clusterable Folium map (`streamlit-folium` +
  `MarkerCluster`) with one cluster layer per severity tier (green/amber/red) plus a
  dedicated, un-clustered layer for suspicious-GPS-duplicate points so that finding
  doesn't get buried inside a cluster bubble.
- **Carbon estimates** are labeled with an explicit confidence tier (Low, proxy-based)
  and an assumptions table listing the exact kg CO₂/tree/year figure used per species —
  the source data has no DBH/girth measurements, so no measured biomass/carbon figure
  can be derived from it.

## 4. Key thresholds (all adjustable at the top of the relevant script)
| Check | Threshold |
|---|---|
| Phone format | 10-digit Indian mobile, starts 6–9 |
| Name-clustering similarity | 0.84 (difflib ratio) |
| GPS-vs-KML-centroid distance | > 300 m |
| Claimed-vs-KML area mismatch | > 15% |
| Plot overlap (flagged) | ≥ 10% of either plot's area |
| Plot overlap (Critical tier) | ≥ 20% of either plot's area |
| Valid GPS region | lat 10–20°N, lon 72–80°E |
| Suspicious GPS cluster | ≥3 records, ≥2 distinct farmers, same ~1m location, inside valid region |

## 5. Before / after cleaning impact (from `Data_Quality_Report.xlsx → Standardization_Impact`)
| Metric | Raw | After standardization |
|---|---|---|
| Block name spelling variants | 102 | 62 canonical names |
| Village name spelling variants | 1,463 | 898 canonical names |
| GP name spelling variants | 705 | 387 canonical names |
| Free-text "Error" note phrasings (partner logs) | 725 | 9 standardized categories |
| Duplicate UUID rows | 14 rows in 7 groups | 7 unique plots if 1 kept per group |

## 6. Headline results (from the sample data provided)
- 4,513 total records (1,563 NGO_1 + 2,950 NGO_2); 3,946 (87.4%) matched to a KML polygon.
- **Severity split:** 14.0% Clean, 56.7% Warning, 29.2% Critical.
- **Data Reliability Score:** NGO_1 (Makkala Vikasa) = **73/100**, NGO_2 (YIDS) = **88/100**.
  NGO_1's lower score is driven almost entirely by GPS/KML-match problems concentrated
  in one submission batch, not by the NGO's fieldwork generally.
- Total claimed area (~23,250 acres) is roughly 3.6× the KML-measured total
  (~6,480 acres) — the single largest discrepancy in the dataset (Area Mismatch,
  2,750 records / 61%).
- 114 KML polygon pairs overlap ≥10%; 207 individual plots overlap ≥20% (Critical tier).
- 400 records have GPS coordinates far outside the project region (column-shift export
  error), concentrated entirely in NGO_1.
- **3 suspicious duplicate-GPS clusters (12 records, all NGO_1):** groups of 3–5
  *different* farmer names all recorded at the exact same coordinate to ~1m precision
  — a strong bulk-copy/lazy-pin-drop signal, not a coincidence.
- 988 records share a phone number with a different farmer name.

## 7. Priority actions (ranked, from `Cleaned_Master.xlsx → Priority_Actions`)
1. **[Critical]** Investigate 12 records in suspicious duplicate-GPS clusters.
2. **[Critical]** Reverify 567 records with no matching KML polygon.
3. **[Critical]** Investigate 114 polygon-overlap pairs (207 plots overlap >20%).
4. **[Critical]** Correct 671 invalid/missing/out-of-region GPS points.
5. **[Critical]** Resolve 14 duplicate UUID records.
6. **[Warning]** Standardize 1,463 village-name spellings (canonical suggestions ready).
7. **[Warning]** Review 988 records sharing a phone number with a different farmer.

Full row-level detail for every number above is in the corresponding sheet of
`2_Cleaned_Data/Cleaned_Master.xlsx`, `Duplicate_Records.xlsx`, `Data_Quality_Report.xlsx`,
and `GIS_Validation_Report.xlsx`.
