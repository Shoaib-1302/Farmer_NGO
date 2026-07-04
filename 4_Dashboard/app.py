"""
3_Scripts/dashboard_app.py   (identical copy also lives at 5_Dashboard/app.py for deployment)

Interactive Streamlit dashboard for the Makkala Vikasa (NGO_1) & YIDS (NGO_2) plantation
data. Reads directly from the 2_Cleaned_Data/*.xlsx outputs produced by data_cleaning.py,
kml_matching.py and overlap_analysis.py — run those three scripts first.

Run:
    streamlit run dashboard_app.py
"""
from pathlib import Path
from collections import Counter

import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import folium
from folium.plugins import MarkerCluster, Fullscreen, MiniMap
from streamlit_folium import st_folium

# ---------------------------------------------------------------------------
# PATHS
# ---------------------------------------------------------------------------
THIS_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = THIS_DIR.parent
CLEANED_DIR = PROJECT_ROOT / '2_Cleaned_Data'

st.set_page_config(page_title='Makkala Vikasa & YIDS | Plantation Intelligence',
                    layout='wide', page_icon='🌳', initial_sidebar_state='expanded')

# ---------------------------------------------------------------------------
# THEME / STYLING
# ---------------------------------------------------------------------------
PRIMARY = '#2E7D4F'       # forest green
PRIMARY_DARK = '#1F4E3D'
ACCENT = '#D9A441'        # harvest gold
BG = '#F6F8F5'
CARD_BG = '#FFFFFF'
RED = '#E2725B'
AMBER = '#E8A93B'
GREEN = '#3FA34D'
PURPLE = '#8B5FBF'
INK = '#1B2B22'
MUTED = '#6B7A6F'

SEVERITY_COLORS = {'Clean': GREEN, 'Warning': AMBER, 'Critical': RED}

PLOTLY_LAYOUT_DEFAULTS = dict(
    font=dict(family='Helvetica Neue, Arial, sans-serif', color=INK, size=12),
    paper_bgcolor='rgba(0,0,0,0)',
    plot_bgcolor='rgba(0,0,0,0)',
    colorway=[PRIMARY, ACCENT, '#6A9FBF', RED, '#B98FD6', '#7fbf7f', '#c9a227', '#5f9ea0'],
    margin=dict(t=50, l=10, r=10, b=10),
    title=dict(font=dict(size=15, color=PRIMARY_DARK)),
    legend=dict(bgcolor='rgba(0,0,0,0)'),
)

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Fraunces:wght@500;600&display=swap');

    html, body, [class*="css"]  {{ font-family: 'Inter', sans-serif; }}
    .stApp {{ background: {BG}; }}

    /* Force dark, legible text for every plain markdown block in the main content area
       (not the sidebar) — this is the fix for text that "disappears" under a dark
       browser/OS theme, since Streamlit's default text color otherwise follows that. */
    section.main [data-testid="stMarkdownContainer"] p,
    section.main [data-testid="stMarkdownContainer"] li,
    section.main [data-testid="stMarkdownContainer"] span,
    section.main [data-testid="stMarkdownContainer"] div {{
        color: {INK} !important;
    }}

    .gp-header {{
        background: linear-gradient(120deg, {PRIMARY_DARK} 0%, {PRIMARY} 100%);
        border-radius: 18px; padding: 28px 34px; margin-bottom: 22px; color: white;
        box-shadow: 0 8px 24px rgba(31,78,61,0.18);
    }}
    .gp-header, .gp-header * {{ color: white !important; }}
    .gp-header .eyebrow {{
        text-transform: uppercase; letter-spacing: .12em; font-size: 12px;
        color: {ACCENT} !important; font-weight: 600; margin-bottom: 6px;
    }}
    .gp-header h1 {{ font-family: 'Fraunces', serif; font-weight: 600; font-size: 30px; margin: 0 0 8px 0; }}
    .gp-header p {{ margin:0; color: #E4EFE6 !important; font-size: 14px; max-width: 780px; line-height:1.5;}}

    .kpi-card {{
        background: {CARD_BG}; border-radius: 14px; padding: 16px 18px;
        border: 1px solid #E4E9E2; box-shadow: 0 2px 10px rgba(31,78,61,0.06); height: 100%;
    }}
    .kpi-card .kpi-label {{ font-size: 12px; color: {MUTED} !important; font-weight:500; margin-bottom:6px; }}
    .kpi-card .kpi-value {{ font-family: 'Fraunces', serif; font-size: 26px; font-weight:600; color: {PRIMARY_DARK} !important; }}
    .kpi-card.critical .kpi-value {{ color: {RED} !important; }}
    .kpi-card.warning .kpi-value {{ color: {AMBER} !important; }}
    .kpi-card .kpi-sub {{ font-size: 11px; color: {MUTED} !important; margin-top: 4px; }}

    .section-title {{ font-family: 'Fraunces', serif; font-size: 20px; color: {PRIMARY_DARK} !important; margin: 6px 0 2px 0; font-weight:600; }}
    .section-sub {{ color: {MUTED} !important; font-size: 13px; margin-bottom: 14px; }}
    .section-sub b, .section-sub code {{ color: {INK} !important; }}

    .stTabs [data-baseweb="tab-list"] {{ gap: 4px; }}
    .stTabs [data-baseweb="tab"] {{
        background-color: white; border-radius: 10px 10px 0 0; padding: 10px 18px;
        border: 1px solid #E4E9E2; border-bottom: none; font-weight: 500;
    }}
    .stTabs [data-baseweb="tab"] p {{ color: {MUTED} !important; }}
    .stTabs [aria-selected="true"] p {{ color: {PRIMARY_DARK} !important; font-weight:600; }}

    section[data-testid="stSidebar"] {{ background: {PRIMARY_DARK}; }}
    section[data-testid="stSidebar"] * {{ color: #EAF1EA !important; }}
    section[data-testid="stSidebar"] .stMultiSelect [data-baseweb="tag"] {{ background: {PRIMARY} !important; }}
    section[data-testid="stSidebar"] .stMultiSelect [data-baseweb="tag"] span {{ color: white !important; }}

    div[data-testid="stMetricValue"] {{ color: {PRIMARY_DARK} !important; }}
    div[data-testid="stMetricLabel"] {{ color: {MUTED} !important; }}

    .legend-chip {{ display:inline-block; padding:3px 10px; border-radius:20px; font-size:11px; margin-right:8px; font-weight:600; }}
    .legend-clean {{ background:#E4F4E7; color:{GREEN} !important; }}
    .legend-warning {{ background:#FCF1DC; color:{AMBER} !important; }}
    .legend-critical {{ background:#FCE6E1; color:{RED} !important; }}
    .legend-purple {{ background:#F0E9F9; color:{PURPLE} !important; }}

    .conf-high {{ color:{GREEN} !important; font-weight:600; }}
    .conf-low {{ color:{RED} !important; font-weight:600; }}

    .action-row {{
        background:white; border-radius:10px; padding:12px 16px; margin-bottom:8px;
        border-left:4px solid {MUTED}; box-shadow:0 1px 6px rgba(0,0,0,0.04);
        font-size:14px; line-height:1.5;
    }}
    .action-row, .action-row *, .action-row p {{ color: {INK} !important; }}
    .action-row b {{ color: {INK} !important; font-weight:700; }}
    .action-row.critical {{ border-left-color: {RED}; }}
    .action-row.warning {{ border-left-color: {AMBER}; }}
    .action-row .legend-chip {{ font-weight:600; }}

    /* Dataframes / tables: force a light background + dark text regardless of OS theme */
    [data-testid="stDataFrame"], [data-testid="stTable"] {{
        background: white !important; color: {INK} !important;
    }}
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# DATA LOADING (cached)
# ---------------------------------------------------------------------------
@st.cache_data
def load_all():
    master_path = CLEANED_DIR / 'Cleaned_Master.xlsx'
    if not master_path.exists():
        return None
    xl = pd.ExcelFile(master_path)
    df = pd.read_excel(xl, sheet_name='Master_Cleaned')
    df['farmer_name'] = df['farmer_name'].astype(str)
    sheets = {name: pd.read_excel(xl, sheet_name=name) for name in
              ['NGO_Summary', 'Reliability_Score', 'Issue_Category_Breakdown',
               'Top_Problem_Villages', 'Priority_Actions']}
    dq_path = CLEANED_DIR / 'Data_Quality_Report.xlsx'
    standardization_impact = None
    if dq_path.exists():
        try:
            standardization_impact = pd.read_excel(dq_path, sheet_name='Standardization_Impact')
        except Exception:
            standardization_impact = None
    gis_path = CLEANED_DIR / 'GIS_Validation_Report.xlsx'
    suspicious_clusters = None
    plot_overlaps = None
    if gis_path.exists():
        try:
            suspicious_clusters = pd.read_excel(gis_path, sheet_name='Suspicious_GPS_Clusters')
        except Exception:
            suspicious_clusters = None
        try:
            plot_overlaps = pd.read_excel(gis_path, sheet_name='Plot_Overlaps')
        except Exception:
            plot_overlaps = None
    return df, sheets, standardization_impact, suspicious_clusters, plot_overlaps


CANON_MAP = {
    'mango': 'Mango', 'arecanut': 'Arecanut', 'coconut': 'Coconut', 'guava': 'Guava',
    'pomegranate': 'Pomegranate', 'mahogany': 'Mahogany', 'jamun': 'Jamun', 'jackfruit': 'Jackfruit',
    'lemon': 'Lemon', 'redsandal': 'RedSandal', 'red': 'RedSandal', 'sandal': 'RedSandal',
    'butterfruit': 'Butterfruit', 'oak': 'Oak', 'rosewood': 'Rosewood', 'mugnahalli': 'Mugnahalli',
    'gooseberry': 'Gooseberry', 'tamarind': 'Tamarind', 'teak': 'Teak', 'woodapple': 'WoodApple',
    'wood': 'WoodApple', 'apple': 'WoodApple', 'bamboo': 'Bamboo', 'palm': 'Palm',
    'other': 'Other/Unspecified', 'others': 'Other/Unspecified',
}
# Illustrative annual CO2 sequestration estimate (kg/tree/yr) — literature-average proxy,
# NOT a measured figure (source data has no DBH/girth measurements).
SEQ_RATE = {
    'Mango': 20, 'Coconut': 15, 'Arecanut': 8, 'Guava': 10, 'Pomegranate': 6, 'Mahogany': 22,
    'Jamun': 18, 'Jackfruit': 20, 'Lemon': 8, 'RedSandal': 15, 'Butterfruit': 15, 'Oak': 20,
    'Rosewood': 18, 'Mugnahalli': 12, 'Gooseberry': 10, 'Tamarind': 22, 'Teak': 20,
    'WoodApple': 12, 'Bamboo': 5, 'Palm': 10, 'Other/Unspecified': 12,
}


def species_counts(species_series):
    counter = Counter()
    for s in species_series.dropna():
        toks = str(s).replace(',', ' ').split()
        canon_in_row = {CANON_MAP.get(t.strip().lower(), 'Other/Unspecified') for t in toks}
        for c in canon_in_row:
            counter[c] += 1
    return counter


def kpi_card(col, label, value, sub=None, tier=None):
    cls = 'kpi-card' + (f' {tier}' if tier else '')
    sub_html = f'<div class="kpi-sub">{sub}</div>' if sub else ''
    col.markdown(f"""
        <div class="{cls}">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            {sub_html}
        </div>
    """, unsafe_allow_html=True)


def style_fig(fig, height=380, **kwargs):
    layout = dict(PLOTLY_LAYOUT_DEFAULTS)
    layout['height'] = height
    layout.update(kwargs)
    fig.update_layout(**layout)
    return fig


# ---------------------------------------------------------------------------
# LOAD
# ---------------------------------------------------------------------------
loaded = load_all()

st.markdown("""
<div class="gp-header">
    <div class="eyebrow">Greenipath &middot; Farmer Plantation Programme</div>
    <h1>🌳 Makkala Vikasa &amp; YIDS — Plantation Intelligence Dashboard</h1>
    <p>A live, filterable view of the cleaned and GIS-validated master dataset — severity-scored
    data quality, spatial validation, species mix, and estimated carbon potential across both NGOs.</p>
</div>
""", unsafe_allow_html=True)

if loaded is None:
    st.error(
        'Cleaned_Master.xlsx not found in 2_Cleaned_Data/. Run the three pipeline scripts '
        'first:\n\n```\ncd 3_Scripts\npython3 data_cleaning.py\npython3 kml_matching.py\n'
        'python3 overlap_analysis.py\n```')
    st.stop()

df, sheets, standardization_impact, suspicious_clusters, plot_overlaps = loaded
ngo_summary_full = sheets['NGO_Summary']
reliability_full = sheets['Reliability_Score']
issue_cat_full = sheets['Issue_Category_Breakdown']
top_villages_full = sheets['Top_Problem_Villages']
priority_actions = sheets['Priority_Actions']

if 'severity' not in df.columns:
    st.error('This Cleaned_Master.xlsx was built with an older version of overlap_analysis.py '
              '(no "severity" column). Re-run the 3 pipeline scripts to regenerate it.')
    st.stop()

# ---------------------------------------------------------------------------
# SIDEBAR FILTERS
# ---------------------------------------------------------------------------
st.sidebar.markdown("### 🌿 Filters")
ngo_options = sorted(df['ngo'].dropna().unique().tolist())
selected_ngos = st.sidebar.multiselect('NGO', ngo_options, default=ngo_options)

severity_options = ['Clean', 'Warning', 'Critical']
selected_severity = st.sidebar.multiselect('Severity', severity_options, default=severity_options)

block_options = sorted([b for b in df['block_name_clean'].dropna().unique() if b != 'INVALID/JUNK VALUE'])
selected_blocks = st.sidebar.multiselect('Block', block_options, default=[])

st.sidebar.markdown("---")
st.sidebar.caption("Data source: `2_Cleaned_Data/Cleaned_Master.xlsx`")
st.sidebar.caption(f"{len(df):,} total records loaded")

fdf = df[df['ngo'].isin(selected_ngos) & df['severity'].isin(selected_severity)]
if selected_blocks:
    fdf = fdf[fdf['block_name_clean'].isin(selected_blocks)]

# ---------------------------------------------------------------------------
# KPI ROW — reordered for "tell the story fast": volume, health, biggest risks, coverage
# ---------------------------------------------------------------------------
n_critical = int((fdf['severity'] == 'Critical').sum())
n_missing_kml = int((~fdf['has_kml_match']).sum())
n_overlap_pairs = int(fdf['has_geometry_overlap'].sum())

k1, k2, k3, k4, k5, k6 = st.columns(6)
kpi_card(k1, 'Total Records', f'{len(fdf):,}')
kpi_card(k2, 'Clean Rate', f"{(fdf['severity']=='Clean').mean()*100:.1f}%")
kpi_card(k3, 'Critical Errors', f'{n_critical:,}', sub=f"{n_critical/len(fdf)*100:.1f}% of records",
          tier='critical' if n_critical > 0 else None)
kpi_card(k4, 'KML Match Rate', f"{fdf['has_kml_match'].mean()*100:.1f}%")
kpi_card(k5, 'Missing KML %', f"{(~fdf['has_kml_match']).mean()*100:.1f}%",
          tier='warning' if (~fdf['has_kml_match']).mean() > 0.05 else None)
kpi_card(k6, 'Plots w/ Geometry Overlap', f'{n_overlap_pairs:,}',
          tier='warning' if n_overlap_pairs > 0 else None)

st.write("")

tabs = st.tabs(['📊 Overview', '🔍 Data Quality', '🛰️ GIS Validation', '🌱 Species & Carbon',
                 '🗺️ Map', '📍 Villages', '👨‍🌾 Farmer Table', '✅ Priority Actions'])
(tab_overview, tab_quality, tab_gis, tab_species, tab_map, tab_villages, tab_farmers, tab_actions) = tabs

# ---------------------------------------------------------------------------
# TAB: Overview
# ---------------------------------------------------------------------------
with tab_overview:
    st.markdown('<div class="section-title">Programme Snapshot</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Severity distribution and reliability score across both NGOs</div>',
                unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        sev_counts = fdf['severity'].value_counts().reindex(['Clean', 'Warning', 'Critical']).fillna(0)
        fig = px.pie(values=sev_counts.values, names=sev_counts.index, hole=.6,
                     color=sev_counts.index, color_discrete_map=SEVERITY_COLORS,
                     title='Records by Severity')
        fig.update_traces(textinfo='label+percent', textfont_size=12)
        st.plotly_chart(style_fig(fig), use_container_width=True)
    with col2:
        ngo_counts = fdf['ngo'].value_counts()
        fig = px.pie(values=ngo_counts.values, names=ngo_counts.index, hole=.6,
                     title='Records by NGO', color_discrete_sequence=[PRIMARY, ACCENT])
        fig.update_traces(textinfo='label+percent', textfont_size=12)
        st.plotly_chart(style_fig(fig), use_container_width=True)

    st.markdown('<div class="section-title">Data Reliability Score</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">100 minus weighted penalties for invalid GPS, duplicate UUIDs, '
                'missing KML matches, area mismatch, geometry overlap, and invalid phone numbers — an '
                'executive-readable summary of how trustworthy each NGO\'s submission batch is.</div>',
                unsafe_allow_html=True)
    rcols = st.columns(len(reliability_full))
    for i, (_, r) in enumerate(reliability_full.iterrows()):
        score = int(r['reliability_score'])
        tier = 'critical' if score < 60 else ('warning' if score < 80 else None)
        kpi_card(rcols[i], f"{r['ngo']} Reliability Score", f'{score}/100',
                  sub=f"{int(r['records']):,} records", tier=tier)
    st.dataframe(reliability_full, use_container_width=True, hide_index=True)

    st.markdown('<div class="section-title">Before / After Cleaning Impact</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">What standardization actually changed</div>', unsafe_allow_html=True)
    if standardization_impact is not None:
        st.dataframe(standardization_impact, use_container_width=True, hide_index=True)
    else:
        st.info('Run data_cleaning.py to generate the Standardization_Impact sheet.')

# ---------------------------------------------------------------------------
# TAB: Data Quality
# ---------------------------------------------------------------------------
with tab_quality:
    st.markdown('<div class="section-title">Issue Categories</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">What should we fix first? — operational breakdown of every automated '
                'QA check (a record can appear in more than one category)</div>', unsafe_allow_html=True)

    cats_in_view = fdf.loc[fdf['severity'] != 'Clean', 'issue_categories'].str.split('; ').explode()
    cat_counts = cats_in_view.value_counts()
    cat_df = cat_counts.reset_index()
    cat_df.columns = ['category', 'count']
    critical_set = {'UUID Issue', 'Missing KML', 'GPS Issue', 'Geometry Overlap', 'Suspicious GPS Duplicate'}
    cat_df['tier'] = cat_df['category'].apply(lambda c: 'Critical' if c in critical_set else 'Warning')
    fig = px.bar(cat_df, x='count', y='category', orientation='h', color='tier',
                 color_discrete_map={'Critical': RED, 'Warning': AMBER},
                 labels={'count': 'Records', 'category': ''})
    fig.update_yaxes(autorange='reversed', categoryorder='total ascending')
    st.plotly_chart(style_fig(fig, height=420, margin=dict(t=20, l=220, r=20, b=40)), use_container_width=True)

    st.markdown('<div class="section-title">Filter by Issue Category</div>', unsafe_allow_html=True)
    chosen_cat = st.selectbox('Show records with issue category:', ['(all flagged records)'] + cat_df['category'].tolist())
    if chosen_cat == '(all flagged records)':
        detail = fdf[fdf['severity'] != 'Clean']
    else:
        detail = fdf[fdf['issue_categories'].str.contains(chosen_cat, na=False)]
    st.dataframe(detail[['ngo', 'uuid', 'farmer_name', 'severity', 'issue_categories']].head(500),
                 use_container_width=True, hide_index=True, height=340)

    st.markdown('<div class="section-title">Partner-Logged Error Notes (standardized)</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">725 free-text phrasings from the NGOs\' own field QA, classified into '
                'consistent categories</div>', unsafe_allow_html=True)
    ec = fdf['error_category'].value_counts()
    ec = ec[ec.index != 'No Note / Not Reviewed'].head(10)
    fig = px.bar(x=ec.values, y=ec.index, orientation='h', labels={'x': 'Records', 'y': ''},
                 color_discrete_sequence=[ACCENT])
    fig.update_yaxes(autorange='reversed')
    st.plotly_chart(style_fig(fig, height=320, margin=dict(t=20, l=220, r=20, b=40)), use_container_width=True)

# ---------------------------------------------------------------------------
# TAB: GIS Validation
# ---------------------------------------------------------------------------
with tab_gis:
    st.markdown('<div class="section-title">Claimed vs. KML-Measured Area</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Each point is one plot. Points on the dotted line agree exactly.'
                '</div>', unsafe_allow_html=True)
    scatter_df = fdf.dropna(subset=['claimed_area', 'kml_area_acres'])
    scatter_df = scatter_df.sample(min(1500, len(scatter_df)), random_state=1)
    max_axis = min(30, max(scatter_df['claimed_area'].max(), scatter_df['kml_area_acres'].max()))
    fig = px.scatter(scatter_df, x='claimed_area', y='kml_area_acres', color='ngo',
                      opacity=0.55, color_discrete_sequence=[PRIMARY, ACCENT],
                      labels={'claimed_area': 'Claimed area (ac)', 'kml_area_acres': 'KML-measured area (ac)'})
    fig.add_trace(go.Scatter(x=[0, max_axis], y=[0, max_axis], mode='lines',
                              line=dict(dash='dot', color=MUTED), name='Perfect agreement'))
    fig.update_xaxes(range=[0, max_axis])
    fig.update_yaxes(range=[0, max_axis])
    st.plotly_chart(style_fig(fig, height=440), use_container_width=True)

    colA, colB, colC = st.columns(3)
    kpi_card(colA, 'Area Mismatch (>15%)', f"{fdf['area_mismatch_flag'].mean()*100:.1f}%", tier='warning')
    kpi_card(colB, 'Plots with Geometry Overlap', f'{n_overlap_pairs:,}',
              tier='critical' if n_overlap_pairs > 0 else None)
    kpi_card(colC, 'GPS-KML Distance Mismatch', f"{fdf['gps_kml_mismatch'].mean()*100:.1f}%", tier='warning')

    st.markdown('<div class="section-title">Suspicious Duplicate-GPS Clusters</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Multiple <i>different</i> farmers recorded at the exact same coordinate '
                '(~1m precision) — a strong signal of bulk-copied/lazy pin-drops rather than genuine surveys.'
                '</div>', unsafe_allow_html=True)
    if suspicious_clusters is not None and len(suspicious_clusters):
        st.dataframe(suspicious_clusters, use_container_width=True, hide_index=True)
    else:
        st.success('No suspicious duplicate-GPS clusters detected.')

# ---------------------------------------------------------------------------
# TAB: Species & Carbon
# ---------------------------------------------------------------------------
with tab_species:
    st.markdown('<div class="section-title">Species Distribution &amp; Estimated Carbon Potential</div>',
                unsafe_allow_html=True)
    counts = species_counts(fdf['species'])
    top = dict(sorted(counts.items(), key=lambda x: -x[1])[:15])
    col1, col2 = st.columns(2)
    with col1:
        fig = px.bar(x=list(top.keys()), y=list(top.values()), title='Species Distribution (plot count)',
                     color_discrete_sequence=[PRIMARY])
        st.plotly_chart(style_fig(fig, height=400), use_container_width=True)
    with col2:
        carbon = {k: round(v * SEQ_RATE.get(k, 12) / 1000, 1) for k, v in counts.items()}
        carbon_top = dict(sorted(carbon.items(), key=lambda x: -x[1])[:12])
        fig = px.bar(x=list(carbon_top.keys()), y=list(carbon_top.values()),
                     title='Estimated Annual CO2 Sequestration (tonnes/yr, illustrative)',
                     color_discrete_sequence=[ACCENT])
        st.plotly_chart(style_fig(fig, height=400), use_container_width=True)

    st.markdown('<div class="section-title">Confidence Levels</div>', unsafe_allow_html=True)
    conf_df = pd.DataFrame([
        {'Metric': 'Species count / distribution', 'Confidence': 'High',
         'Basis': 'Directly recorded in the field by surveyors'},
        {'Metric': 'Plot area (KML-measured)', 'Confidence': 'High',
         'Basis': 'Geodesic calculation from GPS-traced polygon boundaries'},
        {'Metric': 'Carbon sequestration estimate', 'Confidence': 'Low (proxy-based)',
         'Basis': 'Literature-average kg CO2/tree/year by species — no DBH/girth measured in the field'},
    ])
    st.dataframe(conf_df, use_container_width=True, hide_index=True)

    st.markdown('<div class="section-title">Carbon Estimate Assumptions</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Every figure above is <code>tree count x this rate</code> — replace with '
                'field-measured DBH-based allometric equations once available, for a defensible measured figure.'
                '</div>', unsafe_allow_html=True)
    assump_df = pd.DataFrame(sorted(SEQ_RATE.items(), key=lambda x: -x[1]),
                              columns=['Species', 'Assumed kg CO2 / tree / year'])
    st.dataframe(assump_df, use_container_width=True, hide_index=True, height=300)

# ---------------------------------------------------------------------------
# TAB: Map  (3-tier severity clustering + a distinct, un-clustered layer for
#            suspicious duplicate-GPS points, plus an optional overlay showing
#            every overlapping-plot pair as a connecting line)
# ---------------------------------------------------------------------------
with tab_map:
    st.markdown('<div class="section-title">Interactive Plot Map</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-sub">'
        '<span class="legend-chip legend-clean">● Clean</span>'
        '<span class="legend-chip legend-warning">● Warning</span>'
        '<span class="legend-chip legend-critical">● Critical</span>'
        '<span class="legend-chip legend-purple">★ Suspicious duplicate GPS</span>'
        'Zoom in to break clusters apart into individual plots; zoom out to re-cluster. '
        'Click a marker for details.</div>', unsafe_allow_html=True)

    has_overlap_data = plot_overlaps is not None and 'uuid_1' in plot_overlaps.columns and len(plot_overlaps)
    show_overlaps = st.checkbox(
        f'🔴 Highlight polygon overlaps ({len(plot_overlaps):,} pairs)' if has_overlap_data
        else 'No polygon overlaps to highlight',
        value=has_overlap_data, disabled=not has_overlap_data)

    map_df = fdf[fdf['kobo_lat'].between(10, 20) & fdf['kobo_lon'].between(72, 80)].copy()

    @st.cache_resource(show_spinner=False)
    def build_cluster_map(records_tuple, ngo_key, severity_key, block_key, overlay_key, overlap_lines_tuple):
        m = folium.Map(location=[13.4, 78.1], zoom_start=8, tiles='CartoDB positron', control_scale=True)
        clusters = {
            'Clean': MarkerCluster(name='Clean', overlay=True).add_to(m),
            'Warning': MarkerCluster(name='Warning', overlay=True).add_to(m),
            'Critical': MarkerCluster(name='Critical', overlay=True).add_to(m),
        }
        suspicious_group = folium.FeatureGroup(name='Suspicious GPS Duplicate', overlay=True).add_to(m)
        for r in records_tuple:
            color = SEVERITY_COLORS.get(r['severity'], MUTED)
            popup_html = (f"<b>{r['farmer_name']}</b><br>"
                          f"NGO: {r['ngo']}<br>Block: {r['block_name_clean']}<br>"
                          f"Village: {r['village_name_clean']}<br>Species: {r['species']}<br>"
                          f"Severity: <b>{r['severity']}</b><br>"
                          f"Issues: {r['issue_categories']}")
            if r.get('in_suspicious_gps_cluster'):
                folium.RegularPolygonMarker(
                    location=[r['kobo_lat'], r['kobo_lon']], number_of_sides=5, radius=8,
                    color=PURPLE, fill=True, fill_color=PURPLE, fill_opacity=0.95,
                    popup=folium.Popup(popup_html, max_width=260),
                ).add_to(suspicious_group)
            else:
                folium.CircleMarker(
                    location=[r['kobo_lat'], r['kobo_lon']], radius=6, color=color,
                    weight=1, fill=True, fill_color=color, fill_opacity=0.85,
                    popup=folium.Popup(popup_html, max_width=260),
                ).add_to(clusters[r['severity']])

        if overlap_lines_tuple:
            overlap_group = folium.FeatureGroup(name='Polygon Overlaps', overlay=True).add_to(m)
            for ov in overlap_lines_tuple:
                sev_color = RED if max(ov['overlap_pct_of_plot1'], ov['overlap_pct_of_plot2']) >= 20 else AMBER
                popup = (f"Overlap: {ov['overlap_pct_of_plot1']}% of plot1 / "
                         f"{ov['overlap_pct_of_plot2']}% of plot2<br>"
                         f"UUID 1: {ov['uuid_1']}<br>UUID 2: {ov['uuid_2']}")
                folium.PolyLine(
                    [[ov['centroid1_lat'], ov['centroid1_lon']], [ov['centroid2_lat'], ov['centroid2_lon']]],
                    color=sev_color, weight=2.5, opacity=0.8, popup=popup,
                ).add_to(overlap_group)

        folium.LayerControl(collapsed=False).add_to(m)
        Fullscreen(position='topleft').add_to(m)
        MiniMap(toggle_display=True).add_to(m)
        return m

    records = tuple(map_df[['kobo_lat', 'kobo_lon', 'ngo', 'severity', 'farmer_name',
                             'block_name_clean', 'village_name_clean', 'species',
                             'issue_categories', 'in_suspicious_gps_cluster']].to_dict('records'))
    overlap_lines = tuple()
    if show_overlaps and has_overlap_data:
        overlap_lines = tuple(plot_overlaps.dropna(
            subset=['centroid1_lat', 'centroid1_lon', 'centroid2_lat', 'centroid2_lon']
        ).to_dict('records'))
    fmap = build_cluster_map(records, tuple(selected_ngos), tuple(selected_severity), tuple(selected_blocks),
                              show_overlaps, overlap_lines)
    st_folium(fmap, use_container_width=True, height=560, returned_objects=[])

    st.caption(f'Showing all {len(map_df):,} geolocated plots (out of {len(fdf):,} filtered records; '
               f'{(len(fdf) - len(map_df)):,} excluded for missing/invalid GPS coordinates).')

# ---------------------------------------------------------------------------
# TAB: Villages
# ---------------------------------------------------------------------------
with tab_villages:
    st.markdown('<div class="section-title">Top Problematic Villages</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Ranked by number of flagged records — where should field teams go first?'
                '</div>', unsafe_allow_html=True)
    st.dataframe(top_villages_full, use_container_width=True, hide_index=True, height=420)

    if len(top_villages_full):
        fig = px.bar(top_villages_full.head(15), x='issue_count', y='village', orientation='h',
                     color='main_problem', labels={'issue_count': 'Flagged records', 'village': ''})
        fig.update_yaxes(autorange='reversed', categoryorder='total ascending')
        st.plotly_chart(style_fig(fig, height=460, margin=dict(t=20, l=180, r=20, b=40)), use_container_width=True)

# ---------------------------------------------------------------------------
# TAB: Farmer Table
# ---------------------------------------------------------------------------
with tab_farmers:
    st.markdown('<div class="section-title">Farmer-wise Summary</div>', unsafe_allow_html=True)
    farmer_summary = fdf.groupby(['ngo', 'farmer_name'], dropna=False).agg(
        plots=('uuid', 'count'),
        total_claimed_area=('claimed_area', 'sum'),
        total_kml_area=('kml_area_acres', 'sum'),
        clean_plots=('severity', lambda s: (s == 'Clean').sum()),
        critical_plots=('severity', lambda s: (s == 'Critical').sum()),
        block=('block_name_clean', 'first'),
        village=('village_name_clean', 'first'),
    ).reset_index()
    farmer_summary['clean_rate_pct'] = (farmer_summary['clean_plots'] / farmer_summary['plots'] * 100).round(1)
    farmer_summary = farmer_summary.sort_values('plots', ascending=False)
    st.dataframe(farmer_summary, use_container_width=True, hide_index=True, height=460)
    st.download_button('⬇ Download farmer summary as CSV', farmer_summary.to_csv(index=False),
                        file_name='farmer_summary.csv', mime='text/csv')

# ---------------------------------------------------------------------------
# TAB: Priority Actions
# ---------------------------------------------------------------------------
with tab_actions:
    st.markdown('<div class="section-title">Immediate Actions</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Ranked by severity and records affected — what to fix first.</div>',
                unsafe_allow_html=True)
    for _, r in priority_actions.sort_values('priority').iterrows():
        cls = 'critical' if r['severity'] == 'Critical' else 'warning'
        st.markdown(f"""
        <div class="action-row {cls}">
            <b>#{int(r['priority'])}</b> — {r['action']}
            <span class="legend-chip legend-{'critical' if cls=='critical' else 'warning'}">
                {r['severity']} · {int(r['records_affected']):,} records
            </span>
        </div>
        """, unsafe_allow_html=True)

st.markdown(
    f'<div style="text-align:center;color:{MUTED};font-size:12px;margin-top:30px;">'
    'Greenipath Assignment — Data Scientist Assessment Dashboard</div>', unsafe_allow_html=True)
