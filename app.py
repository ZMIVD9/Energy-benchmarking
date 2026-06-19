"""
Energy Model Benchmarking & QA/QC Platform
===========================================
Run:  streamlit run app.py
Deps: pip install streamlit pandas plotly openpyxl reportlab

IMPORTANT: Keep benchmarks.xlsx in the same folder as app.py
"""

import io
import os
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import date

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Energy Benchmarking Platform",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .main { background-color: #f8fafc; }
    .block-container { padding-top: 1.5rem; }
    h1 { color: #0f4c81; font-size: 1.6rem !important; }
    h2 { color: #0f4c81; font-size: 1.2rem !important; }
    h3 { font-size: 1rem !important; }
    .flag-pass { background:#f0fdf4; border-left:4px solid #16a34a; padding:10px 14px; border-radius:6px; margin:6px 0; color:#14532d !important; font-size:13px; }
    .flag-warn { background:#fffbeb; border-left:4px solid #d97706; padding:10px 14px; border-radius:6px; margin:6px 0; color:#78350f !important; font-size:13px; }
    .flag-fail { background:#fef2f2; border-left:4px solid #dc2626; padding:10px 14px; border-radius:6px; margin:6px 0; color:#7f1d1d !important; font-size:13px; }
    .flag-info { background:#eff6ff; border-left:4px solid #3b82f6; padding:10px 14px; border-radius:6px; margin:6px 0; color:#1e3a8a !important; font-size:13px; }
    .flag-pass b, .flag-warn b, .flag-fail b, .flag-info b { font-size:15px; }
    .bm-card   { background:white; border-radius:10px; padding:16px 20px; border:1px solid #e2e8f0; margin-bottom:12px; }
    .bm-label  { font-size:11px; font-weight:600; color:#64748b; text-transform:uppercase; letter-spacing:.05em; }
    .bm-value  { font-size:22px; font-weight:700; color:#0f4c81; line-height:1.2; }
    .bm-sub    { font-size:12px; color:#94a3b8; margin-top:2px; }
    .success-box { background:#f0fdf4; border:1px solid #bbf7d0; border-radius:8px; padding:12px 16px; margin:8px 0; }
</style>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────────────────────
# Google Sheet ID — replace this with your own Sheet ID
# Get it from your sheet URL: docs.google.com/spreadsheets/d/YOUR_ID_HERE/edit
SHEET_ID = st.secrets.get("SHEET_ID", "YOUR_SHEET_ID_HERE")
SHEET_NAME = "Benchmarks"
SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={SHEET_NAME}"
# NECB end-use savings benchmarks live on their own tab. Set this to match the tab name
# in your sheet (the screenshot shows "NECB Savings"; the spec calls it "NECB").
NECB_SHEET_NAME = "NECB Savings"
NECB_SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={NECB_SHEET_NAME.replace(' ', '%20')}"
NECB_TOL = 10  # percentage-point tolerance: savings within 10 pts of (or above) the target → pass
MODEL_TYPES      = ["Proposed","Baseline","Existing"]
PROJECT_PHASES   = ["Concept","Schematic Design","Design Development","100% Design","As-Built"]
SOFTWARE_OPTIONS = ["IES VE","EnergyPlus","OpenStudio","eQUEST","Manual / Excel template","Other"]
DHW_BUILDINGS    = ["School","Office","Hospital","Residential","Community Centre","Library"]
END_USE_COLORS   = ["#ef4444","#3b82f6","#8b5cf6","#f59e0b","#06b6d4","#10b981","#f97316"]

# ── GHG emission factors (Canada 2026 provincial values) ──────────────────────
# Natural gas energy content used to convert g CO₂/m³ → kg CO₂/kWh (HHV ≈ 38 MJ/m³).
NG_KWH_PER_M3 = 10.55
# Marketable natural-gas CO₂ factor by province/territory, g CO₂/m³ (Table 1.3, 2026).
NG_CO2_G_PER_M3 = {
    "British Columbia":1966, "Alberta":1962, "Saskatchewan":1920, "Manitoba":1915,
    "Ontario":1921, "Quebec":1926, "New Brunswick":1919, "Nova Scotia":1919,
    "Prince Edward Island":1919, "Newfoundland and Labrador":1919,
    "Yukon":1966, "Northwest Territories":1966, "Nunavut":1966,
}
# Electricity consumption intensity by province/territory, g CO₂e/kWh (Table 5.3, 2026).
ELEC_CO2E_G_PER_KWH = {
    "British Columbia":18, "Alberta":438, "Saskatchewan":631, "Manitoba":2.5,
    "Ontario":59, "Quebec":1.9, "New Brunswick":234, "Nova Scotia":581,
    "Prince Edward Island":234, "Newfoundland and Labrador":17,
    "Yukon":74, "Northwest Territories":420, "Nunavut":800,
}
# Map the cities used in the tool to their province/territory.
CITY_PROVINCE = {
    "Edmonton":"Alberta", "Calgary":"Alberta", "Red Deer":"Alberta", "Lethbridge":"Alberta",
    "Vancouver":"British Columbia", "Victoria":"British Columbia", "Kelowna":"British Columbia",
    "Toronto":"Ontario", "Ottawa":"Ontario", "Hamilton":"Ontario", "London":"Ontario",
    "Winnipeg":"Manitoba", "Regina":"Saskatchewan", "Saskatoon":"Saskatchewan",
    "Montreal":"Quebec", "Quebec City":"Quebec",
    "Halifax":"Nova Scotia", "Fredericton":"New Brunswick",
    "St. John's":"Newfoundland and Labrador", "Charlottetown":"Prince Edward Island",
    "Whitehorse":"Yukon", "Yellowknife":"Northwest Territories", "Iqaluit":"Nunavut",
}

def province_of_city(city):
    """Province for a city: prefer the sheet's Province column, fall back to the built-in map."""
    try:
        for (bt, c, z, sub), v in BENCHMARKS.items():
            if c == city and v.get("province"):
                return v["province"]
    except NameError:
        pass
    return CITY_PROVINCE.get(city)

def gas_factor_for(city):
    """Default natural-gas factor (kg CO₂/kWh) for a city, converted from g CO₂/m³."""
    g = NG_CO2_G_PER_M3.get(province_of_city(city) or "")
    return round(g / NG_KWH_PER_M3 / 1000, 4) if g else 0.0

def elec_factor_for(city):
    """Default electricity factor (kg CO₂e/kWh) for a city."""
    g = ELEC_CO2E_G_PER_KWH.get(province_of_city(city) or "")
    return round(g / 1000, 4) if g is not None else 0.0

def average_benchmarks(bm_list):
    """Average a list of benchmark dicts into a single aggregated benchmark
    (used for province-wide views, e.g. all Schools in Alberta)."""
    n = len(bm_list)
    def mean(key):
        vals = [b[key] for b in bm_list if b.get(key) is not None]
        return round(sum(vals) / len(vals), 1) if vals else 0
    def mean_pos(key):  # average only non-zero values (0 = not provided, e.g. TEDI)
        vals = [b[key] for b in bm_list if b.get(key)]
        return round(sum(vals) / len(vals), 1) if vals else 0
    def mean_array(key):
        arrs = [b[key] for b in bm_list if b.get(key)]
        if not arrs:
            return []
        L = min(len(a) for a in arrs)
        return [round(sum(a[i] for a in arrs) / len(arrs), 1) for i in range(L)]
    med_eui  = mean("median_eui")
    med_tedi = mean_pos("median_tedi")
    return {
        "median_eui":  med_eui,
        "good_eui":    round(med_eui * 0.85, 1),
        "high_eui":    round(med_eui * 1.15, 1),
        "median_tedi": med_tedi,
        "good_tedi":   round(med_tedi * 0.85, 1) if med_tedi else 0,
        "high_tedi":   round(med_tedi * 1.15, 1) if med_tedi else 0,
        "median_ghgi": mean("median_ghgi"),
        "heat_pct": mean("heat_pct"), "cool_pct": mean("cool_pct"), "fan_pct": mean("fan_pct"),
        "ltg_pct": mean("ltg_pct"), "dhw_pct": mean("dhw_pct"), "recept_pct": mean("recept_pct"),
        "pumps_pct": mean("pumps_pct"), "elec_pct": mean("elec_pct"),
        "gas_pct": mean("gas_pct"), "heat_rej_pct": mean("heat_rej_pct"),
        "misc_pct": mean("misc_pct"),
        "pct_data":      mean_array("pct_data"),
        "tedi_pct_data": mean_array("tedi_pct_data"),
        "subtype": "All", "_n": n,
    }

AUTO_MAP = {
    # ── Main energy sources ──
    "electricity_kwh":  ["electricity","total electricity","elec","electricity_kwh","electricity (kwh)",
                         "total elec","elec_kwh","annual electricity","total_kwh","total electricity use"],
    "gas_kwh":          ["gas","natural gas","gas consumption","gas_kwh","naturalgas","natural gas (kwh)",
                         "gas (kwh)","annual gas","ng_kwh","natural_gas","natural_gas_kwh"],
    "area_m2":          ["area","floor area","gfa","area_m2","floor area (m2)","area (m2)",
                         "gross floor area","conditioned area","heated floor area","area_m2"],
    # ── Heating ──
    "heating_kwh":      ["heating","heat","heating_kwh","heating (kwh)","htg","htg energy",
                         "space heating","space_heating","space heating (kwh)","space_heating_kwh",
                         "annual heating","zone heating","heat energy"],
    # ── Cooling ──
    "cooling_kwh":      ["cooling","cool","cooling_kwh","cooling (kwh)","clg","clg energy",
                         "space cooling","space_cooling","space cooling (kwh)","space_cooling_kwh",
                         "annual cooling","zone cooling","heat rejection","heat_rejection",
                         "heat rejection (kwh)","heat_rejection_kwh"],
    # ── Fans — three separate components ──
    "central_fan_kwh":  ["central fan","central_fan","central_fan_kwh","central fan (kwh)",
                         "interior central fan","interior_central_fan","interior_central_fan_kwh",
                         "interior central fans","interior_central_fans","interior_central_fans_kwh",
                         "supply fan","supply_fan","supply_fan_kwh","ahu fan","ahu_fan",
                         "air handling unit","air handling","central air handling",
                         "fans","fan","fans_kwh","fan energy","supply fans","total fan energy"],
    "local_fan_kwh":    ["local fan","local_fan","local_fan_kwh","local fan (kwh)",
                         "interior local fan","interior_local_fan","interior_local_fan_kwh",
                         "interior local fans","interior_local_fans","interior_local_fans_kwh",
                         "fan coil","fan coil unit","fcu","fan_coil","fan_coil_kwh",
                         "zone fan","local ventilation","unit ventilator","local unit"],
    "exhaust_fan_kwh":  ["exhaust fan","exhaust_fan","exhaust_fan_kwh","exhaust fan (kwh)",
                         "exhaust fans","exhaust_fans","exhaust energy","exhaust fan energy",
                         "exhaust_fan_energy","building exhaust","toilet exhaust",
                         "kitchen exhaust","parking exhaust","general exhaust"],
    # ── Lighting ──
    "lighting_kwh":     ["lighting","lights","lighting_kwh","lighting (kwh)","ltg",
                         "interior lighting","interior_lighting","interior_lighting_kwh",
                         "interior lighting (kwh)","annual lighting","zone lighting",
                         "lighting energy","interior lights"],
    # ── DHW ──
    "dhw_kwh":          ["dhw","domestic hot water","dhw_kwh","dhw (kwh)","hot water",
                         "service hot water","shw","dhw energy","dhw_energy",
                         "domestic hot water (kwh)","hot water energy"],
    # ── Pumps ──
    "pumps_kwh":        ["pumps","pump","pumps_kwh","pump energy","pumps (kwh)","heating pumps",
                         "pump_kwh","pumps energy","chw pump","hw pump","condenser pump",
                         "heating pump","cooling pump","pumping energy"],
    # ── Receptacle / plug loads ──
    "receptacle_kwh":   ["receptacle","receptacles","receptacle_kwh","receptacle (kwh)",
                         "plug loads","plug_loads","plug load","plug_load_kwh",
                         "equipment","equipment (kwh)","equipment_kwh","process load",
                         "interior equipment","interior_equipment","interior_equipment_kwh",
                         "miscellaneous","misc load","misc_load","other loads"],
    # ── Interior ceiling / heat rejection misc ──
    "heat_rejection_kwh":["heat rejection","heat_rejection","heat_rejection_kwh",
                          "heat rejection (kwh)","cooling tower","condenser heat",
                          "interior ceiling","interior_ceiling","interior_ce"],
    # ── Exterior lighting ──
    "ext_lighting_kwh": ["exterior lighting","exterior_lighting","exterior_lighting_kwh",
                         "ext lighting","ext_lighting","outdoor lighting","site lighting"],
    # ── Process / other ──
    "process_kwh":      ["process","process load","process_kwh","process (kwh)",
                         "process energy","other","other energy","other_kwh"],
    # ── Other fuels / biomass / district energy (counted in total energy + GHGI) ──
    "other_fuel_kwh":   ["biomass","biomass_kwh","other fuel","other_fuel","other_fuel_kwh",
                         "other resource","other resources","district heating","district energy",
                         "oil","propane","wood","wood pellets","other fuel (kwh)"],
    # ── TEDI — Thermal Energy Demand Intensity (already kWh/m²·yr, NOT divided by area) ──
    "tedi":             ["tedi","tedi (kwh/m2)","tedi_kwh_m2","tedi kwh/m2","tedi (kwh/m2·yr)",
                         "thermal energy demand intensity","thermal demand intensity","tedi_kwh/m2"],
    # ── Unmet (out-of-range comfort) hours — counts, NOT divided by area ──
    "unmet_hours_heating": ["unmet_hours_heating","unmet hours heating","heating unmet hours","unmet heating hours"],
    "unmet_hours_cooling": ["unmet_hours_cooling","unmet hours cooling","cooling unmet hours","unmet cooling hours"],
    "unmet_hours_total":   ["unmet_hours_total","unmet hours total","total unmet hours","unmet hours","unmet_hours"],
}

FIELD_LABELS = {
    "electricity_kwh":   "Electricity (kWh/yr)",
    "gas_kwh":           "Natural Gas (kWh/yr)",
    "area_m2":           "Floor Area (m²)",
    "heating_kwh":       "Space Heating (kWh/yr)",
    "cooling_kwh":       "Space Cooling (kWh/yr)",
    "central_fan_kwh":   "Interior Central Fan / AHU (kWh/yr)",
    "local_fan_kwh":     "Interior Local Fan (kWh/yr)",
    "exhaust_fan_kwh":   "Exhaust Fan (kWh/yr)",
    "lighting_kwh":      "Interior Lighting (kWh/yr)",
    "dhw_kwh":           "DHW (kWh/yr)",
    "pumps_kwh":         "Pumps (kWh/yr)",
    "receptacle_kwh":    "Receptacle / Plug Loads (kWh/yr)",
    "heat_rejection_kwh":"Heat Rejection (kWh/yr)",
    "ext_lighting_kwh":  "Exterior Lighting (kWh/yr)",
    "process_kwh":       "Process / Other (kWh/yr)",
    "other_fuel_kwh":    "Other Fuel / Biomass (kWh/yr)",
    "tedi":              "TEDI (kWh/m²·yr)",
    "unmet_hours_heating": "Unmet Hours — Heating",
    "unmet_hours_cooling": "Unmet Hours — Cooling",
    "unmet_hours_total":   "Unmet Hours — Total",
}

# ── Load benchmarks from Google Sheets ────────────────────────────────────────
@st.cache_data(ttl=60)  # cache for 60 seconds — edits in Google Sheets appear within 1 minute
def load_benchmarks():
    """Read from Google Sheets — keyed by (building_type, city, zone, subtype).
    Subtype allows multiple benchmarks for the same building type + city
    e.g. School · Edmonton · Zone 7 · Boiler+VAV  vs  School · Edmonton · Zone 7 · Heat Pump+DOAS
    """
    try:
        df = pd.read_csv(SHEET_URL, dtype=str)
    except Exception as e:
        st.error(f"❌ Could not load benchmark data from Google Sheets. Check your SHEET_ID. Error: {e}")
        st.stop()
    df.columns = [c.strip() for c in df.columns]
    benchmarks = {}
    for _, row in df.iterrows():
        try:
            btype   = str(row["Building Type"]).strip()
            city    = str(row["City"]).strip()
            zone    = str(row["Climate Zone"]).strip()
            # Province from the sheet (accepts the common "Provience" misspelling);
            # falls back to the built-in city→province map only if the column is blank.
            province = str(row.get("Province", row.get("Provience", "")) or "").strip() or CITY_PROVINCE.get(city, "")
            # Subtype is optional — falls back to "General" if column missing or empty
            subtype = str(row.get("Subtype", "") or "").strip() or "General"
            pct_raw = str(row["Percentile Data (comma separated)"]).strip()
            pct_data = [float(x.strip()) for x in pct_raw.split(",") if x.strip()]
            median = float(row["Median EUI"])
            try:
                med_tedi = float(row.get("Median TEDI", "") or 0)
            except (TypeError, ValueError):
                med_tedi = 0.0
            # Optional TEDI percentile distribution. If a dedicated column isn't provided
            # but a Median TEDI is, approximate the shape from the EUI distribution scaled
            # to the TEDI median so the chart still renders.
            tedi_pct_raw = str(row.get("TEDI Percentile Data (comma separated)", "") or "").strip()
            tedi_pct_data = [float(x.strip()) for x in tedi_pct_raw.split(",") if x.strip()]
            if not tedi_pct_data and med_tedi and pct_data and median:
                tedi_pct_data = [round(v * med_tedi / median, 1) for v in pct_data]
            benchmarks[(btype, city, zone, subtype)] = {
                "median_eui":  median,
                "good_eui":    round(median * 0.85, 1),   # −15% of median
                "high_eui":    round(median * 1.15, 1),   # +15% of median
                "median_tedi": med_tedi,
                "good_tedi":   round(med_tedi * 0.85, 1) if med_tedi else 0,
                "high_tedi":   round(med_tedi * 1.15, 1) if med_tedi else 0,
                "median_ghgi": float(row["Median GHGI"]),
                "heat_pct":    float(row["Heating %"]),
                "cool_pct":    float(row["Cooling %"]),
                "fan_pct":     float(row["Fan %"]),
                "ltg_pct":     float(row["Lighting %"]),
                "dhw_pct":     float(row["DHW %"]),
                "recept_pct":  float(row["Receptacle %"]),
                "pumps_pct":   float(row["Pumps %"]),
                "elec_pct":    float(row.get("Electricity %")   or row.get("Electricity")  or 0),
                "gas_pct":     float(row.get("NaturalGas %")    or row.get("Natural Gas %") or row.get("NaturalGas") or 0),
                "heat_rej_pct":float(row.get("Heat Rejection %") or row.get("HeatRejection %") or 0),
                "misc_pct":    float(row.get("Miscellaneous %")  or row.get("Miscellaneous") or 0),
                "pct_data":    pct_data,
                "tedi_pct_data": tedi_pct_data,
                "subtype":     subtype,
                "province":    province,
                "project_year":   str(row.get("Project Year/Date", row.get("Project Year", "")) or "").strip(),
                "audit_new":      str(row.get("Audit/New", "") or "").strip(),
                "heating_system": str(row.get("Heating System", "") or "").strip(),
            }
        except Exception:
            continue
    return benchmarks

def reload_benchmarks():
    """Clear cache and reload from Google Sheets."""
    load_benchmarks.clear()
    st.rerun()

# ── NECB end-use savings benchmarks (separate tab) ───────────────────────────
# Maps each internal energy field to the NECB end-use name(s) used on the sheet.
NECB_ENDUSES = [
    ("electricity_kwh",    "Electricity",       ["electricity","elec","total electricity"]),
    ("gas_kwh",            "Natural Gas",       ["natural gas","gas","naturalgas","natural_gas"]),
    ("lighting_kwh",       "Interior Lighting", ["interior lighting","lighting","lights","ltg","interior_lighting"]),
    ("heating_kwh",        "Space Heating",     ["space heating","heating","heat","space_heating"]),
    ("cooling_kwh",        "Space Cooling",     ["space cooling","cooling","cool","space_cooling"]),
    ("pumps_kwh",          "Pumps",             ["pumps","pump"]),
    ("heat_rejection_kwh", "Heat Rejection",    ["heat rejection","heat_rejection","heat rej"]),
    ("central_fan_kwh",    "Central Fan",       ["central fan","interior central fans","central fans","ahu","interior_central_fans"]),
    ("local_fan_kwh",      "Local Fan",         ["local fan","interior local fans","local fans","interior_local_fans"]),
    ("exhaust_fan_kwh",    "Exhaust Fan",       ["exhaust fan","exhaust fans","exhaust_fans"]),
    ("dhw_kwh",            "DHW",               ["dhw","domestic hot water","hot water","service hot water"]),
    ("receptacle_kwh",     "Receptacle",        ["receptacle","receptacle loads","plug loads","receptacles"]),
    ("total_kwh",          "Total Energy",      ["total energy","total","total energy use","total_kwh","total kwh"]),
]

# Fallback savings targets (used if the NECB tab can't be read or is missing a value).
DEFAULT_NECB_SAVINGS = {
    "electricity_kwh":30, "gas_kwh":40, "lighting_kwh":76, "heating_kwh":76, "cooling_kwh":30,
    "pumps_kwh":40, "heat_rejection_kwh":76, "central_fan_kwh":76, "local_fan_kwh":30,
    "exhaust_fan_kwh":40, "dhw_kwh":76, "receptacle_kwh":76, "total_kwh":40,
}

def _match_necb_key(name):
    """Resolve a sheet end-use label to an internal field key."""
    n = str(name).strip().lower()
    for key, _label, aliases in NECB_ENDUSES:   # exact match first
        if n == key or n in aliases:
            return key
    for key, _label, aliases in NECB_ENDUSES:   # then substring
        if any(a in n for a in aliases):
            return key
    return None

# NECB Savings tab end-use columns → internal field keys, matched on a normalized
# (lowercase, alphanumeric-only, trailing "kwh"/"%" stripped) form so it tolerates the
# per-end-use names (electricity, natural_gas, space_heating, …), the legacy "Heating %"
# names, and small spelling/casing differences.
NECB_NORM_ALIASES = {
    "electricity": ["electricity_kwh"], "elec": ["electricity_kwh"], "totalelectricity": ["electricity_kwh"],
    "naturalgas": ["gas_kwh"], "gas": ["gas_kwh"], "natgas": ["gas_kwh"],
    "interiorlighting": ["lighting_kwh"], "lighting": ["lighting_kwh"], "lights": ["lighting_kwh"],
    "spaceheating": ["heating_kwh"], "heating": ["heating_kwh"], "heat": ["heating_kwh"],
    "spacecooling": ["cooling_kwh"], "cooling": ["cooling_kwh"], "cool": ["cooling_kwh"],
    "pumps": ["pumps_kwh"], "pump": ["pumps_kwh"],
    "heatrejection": ["heat_rejection_kwh"], "heatrej": ["heat_rejection_kwh"],
    "interiorcentralfans": ["central_fan_kwh"], "interiorcentralfan": ["central_fan_kwh"],
    "centralfans": ["central_fan_kwh"], "centralfan": ["central_fan_kwh"],
    "interiorlocalfans": ["local_fan_kwh"], "interiorlocalfan": ["local_fan_kwh"],
    "localfans": ["local_fan_kwh"], "localfan": ["local_fan_kwh"],
    "exhaustfans": ["exhaust_fan_kwh"], "exhaustfan": ["exhaust_fan_kwh"],
    "fan": ["central_fan_kwh", "local_fan_kwh", "exhaust_fan_kwh"],   # legacy combined "Fan %"
    "fans": ["central_fan_kwh", "local_fan_kwh", "exhaust_fan_kwh"],
    "dhw": ["dhw_kwh"], "domestichotwater": ["dhw_kwh"], "hotwater": ["dhw_kwh"],
    "receptacle": ["receptacle_kwh"], "receptacles": ["receptacle_kwh"],
    "receptacleloads": ["receptacle_kwh"], "plugloads": ["receptacle_kwh"],
    "total": ["total_kwh"], "totalenergy": ["total_kwh"], "totalkwh": ["total_kwh"],
}

def _necb_norm(s):
    n = "".join(ch for ch in str(s).lower() if ch.isalnum())
    if n.endswith("kwh"):
        n = n[:-3]
    return n

def _necb_col_keys(colname):
    """Internal field key(s) a NECB-sheet end-use column maps to, or None for metadata cols."""
    return NECB_NORM_ALIASES.get(_necb_norm(colname))

def _necb_code_from_version(v):
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return None
    try:    s = str(int(float(s)))   # "2020" / "2020.0" → "2020"
    except Exception: pass
    return s if s.upper().startswith("NECB") else f"NECB {s}"

@st.cache_data(ttl=60)
def load_necb_savings():
    """Read the NECB Savings tab (wide format: one row per building type / city / NECB
    version, with per-end-use savings columns). Returns {"rows":[...], "codes":[...]}."""
    empty = {"rows": [], "codes": []}
    try:
        df = pd.read_csv(NECB_SHEET_URL, dtype=str)
    except Exception:
        return empty
    df.columns = [str(c).strip() for c in df.columns]
    low = {c.lower(): c for c in df.columns}
    bt_col   = next((low[c] for c in ["building type","buildingtype","type"] if c in low), None)
    city_col = next((low[c] for c in ["city"] if c in low), None)
    prov_col = next((low[c] for c in ["province","provience"] if c in low), None)
    ver_col  = next((low[c] for c in ["necb version","version","compliance code","code","necb"] if c in low), None)
    # Any column that resolves to an end-use key is treated as a savings target.
    sav_cols = {c: _necb_col_keys(c) for c in df.columns if _necb_col_keys(c)}
    if not sav_cols:
        return empty
    rows, codes = [], []
    for _, r in df.iterrows():
        code = (_necb_code_from_version(r[ver_col]) if ver_col else None) or "NECB"
        if code not in codes:
            codes.append(code)
        savings = {}
        for actual, keys in sav_cols.items():
            raw = str(r[actual]).replace("%", "").replace(",", "").strip()
            if raw == "" or raw.lower() == "nan":
                continue
            try:
                pctv = float(raw)
            except Exception:
                continue
            if pctv != pctv:   # NaN guard
                continue
            for key in keys:
                savings[key] = pctv
        rows.append({
            "building_type": (str(r[bt_col]).strip()   if bt_col   else ""),
            "city":          (str(r[city_col]).strip() if city_col else ""),
            "province":      (str(r[prov_col]).strip() if prov_col else ""),
            "code":          code,
            "savings":       savings,
        })
    return {"rows": rows, "codes": codes or ["NECB 2020"]}

def _avg_savings(maps):
    keys = set().union(*[m.keys() for m in maps]) if maps else set()
    out = {}
    for k in keys:
        vals = [m[k] for m in maps if k in m]
        if vals:
            out[k] = round(sum(vals) / len(vals), 1)
    return out

def necb_savings_for(necb, building_type, city, code):
    """Resolve the savings target map for a building type / city / code, with fallbacks:
    exact (type+city+code) → type+code (avg across cities) → type → any."""
    rows = necb.get("rows", [])
    bt = (building_type or "").strip().lower()
    ct = (city or "").strip().lower()
    exact = [r["savings"] for r in rows if r["building_type"].lower()==bt and r["city"].lower()==ct and r["code"]==code]
    if exact:
        return exact[0]
    bc = [r["savings"] for r in rows if r["building_type"].lower()==bt and r["code"]==code]
    if bc:
        return _avg_savings(bc)
    b = [r["savings"] for r in rows if r["building_type"].lower()==bt]
    if b:
        return _avg_savings(b)
    return _avg_savings([r["savings"] for r in rows]) if rows else {}

def build_necb_rows(vals, ref_vals, savings_map):
    """Per-end-use proposed-vs-reference savings rows for the NECB QA table."""
    def g(d, k):
        try: return float((d or {}).get(k) or 0)
        except Exception: return 0.0
    def total(d): return g(d,"electricity_kwh") + g(d,"gas_kwh") + g(d,"other_fuel_kwh")
    rows = []
    for key, label, _aliases in NECB_ENDUSES:
        prop = total(vals)     if key == "total_kwh" else g(vals, key)
        ref  = total(ref_vals) if key == "total_kwh" else g(ref_vals, key)
        bench = savings_map.get(key)
        if ref:
            savings = (ref - prop) / ref * 100
            sav_txt = f"{savings:+.0f}%"
            auto = "⚪" if bench is None else ("🟢" if savings >= bench - NECB_TOL else "🔴")
        else:
            sav_txt = "n/a"; auto = "⚪"
        rows.append({
            "End Use": label,
            "Proposed": round(prop, 1),
            "Reference Model": round(ref, 1),
            "Savings": sav_txt,
            "Benchmark": (f"{bench:.0f}%" if bench is not None else "—"),
            "Auto Flag": auto,
        })
    return rows


def append_to_google_sheet(sheet_name: str, row: list):
    """
    Append a row to a Google Sheet using the Sheets API via requests.
    Requires SHEET_ID and GOOGLE_SERVICE_ACCOUNT in st.secrets.
    Falls back gracefully if credentials are not configured.
    """
    try:
        import json, requests
        creds_raw = st.secrets.get("GOOGLE_SERVICE_ACCOUNT", None)
        if not creds_raw:
            return False, "No Google service account configured in secrets."

        creds = json.loads(creds_raw) if isinstance(creds_raw, str) else dict(creds_raw)

        # Get OAuth2 token using service account
        import time, base64, hashlib
        from urllib.parse import urlencode

        # Build JWT
        header = base64.urlsafe_b64encode(
            json.dumps({"alg":"RS256","typ":"JWT"}).encode()
        ).rstrip(b"=").decode()

        now = int(time.time())
        claim = base64.urlsafe_b64encode(json.dumps({
            "iss": creds["client_email"],
            "scope": "https://www.googleapis.com/auth/spreadsheets",
            "aud": "https://oauth2.googleapis.com/token",
            "iat": now, "exp": now + 3600
        }).encode()).rstrip(b"=").decode()

        from cryptography.hazmat.primitives import hashes, serialization
        from cryptography.hazmat.primitives.asymmetric import padding

        private_key = serialization.load_pem_private_key(
            creds["private_key"].encode(), password=None
        )
        sig_input = f"{header}.{claim}".encode()
        signature = base64.urlsafe_b64encode(
            private_key.sign(sig_input, padding.PKCS1v15(), hashes.SHA256())
        ).rstrip(b"=").decode()

        jwt = f"{header}.{claim}.{signature}"

        token_resp = requests.post("https://oauth2.googleapis.com/token", data={
            "grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer",
            "assertion": jwt,
        })
        access_token = token_resp.json().get("access_token")
        if not access_token:
            return False, f"Could not get access token: {token_resp.text}"

        # Append row to sheet
        url = (f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}"
               f"/values/{sheet_name}!A1:append?valueInputOption=USER_ENTERED"
               f"&insertDataOption=INSERT_ROWS")
        resp = requests.post(url,
            headers={"Authorization": f"Bearer {access_token}"},
            json={"values": [row]}
        )
        if resp.status_code == 200:
            return True, "Success"
        else:
            return False, f"Sheets API error: {resp.text}"
    except Exception as e:
        return False, str(e)

# ── Helper functions ───────────────────────────────────────────────────────────
def guess_mapping(headers):
    mapping = {}
    for field, aliases in AUTO_MAP.items():
        match = next((h for h in headers if h.strip().lower() in aliases), "")
        mapping[field] = match
    return mapping

def calc_percentile(eui, pct_data):
    above = sum(1 for v in sorted(pct_data) if v > eui)
    return round((above / len(pct_data)) * 100)

def calculate_kpis(vals, area_override=None, ef_elec=0.0, ef_gas=0.0, ef_other=0.0):
    area    = float(area_override) if area_override else float(vals.get("area_m2") or 1)
    elec    = float(vals.get("electricity_kwh")    or 0)
    gas     = float(vals.get("gas_kwh")            or 0)
    other_fuel = float(vals.get("other_fuel_kwh")  or 0)   # biomass / district energy / oil etc.
    heat    = float(vals.get("heating_kwh")        or 0)
    cool    = float(vals.get("cooling_kwh")        or 0)
    central_fan  = float(vals.get("central_fan_kwh")  or 0)
    local_fan    = float(vals.get("local_fan_kwh")    or 0)
    exhaust_fan  = float(vals.get("exhaust_fan_kwh")  or 0)
    fans         = central_fan + local_fan + exhaust_fan  # total fans = central + local + exhaust
    ltg     = float(vals.get("lighting_kwh")       or 0)
    dhw     = float(vals.get("dhw_kwh")            or 0)
    pumps   = float(vals.get("pumps_kwh")          or 0)
    recept  = float(vals.get("receptacle_kwh")     or 0)
    heat_rej= float(vals.get("heat_rejection_kwh") or 0)
    ext_ltg = float(vals.get("ext_lighting_kwh")   or 0)
    process = float(vals.get("process_kwh")        or 0)
    tedi    = float(vals.get("tedi")               or 0)   # already kWh/m²·yr — not divided by area
    # Unmet (out-of-range comfort) hours — raw counts, NOT divided by area
    unmet_h = float(vals.get("unmet_hours_heating") or 0)
    unmet_c = float(vals.get("unmet_hours_cooling") or 0)
    unmet_t = float(vals.get("unmet_hours_total")   or 0)
    total   = elec + gas + other_fuel
    def eui(v): return round(v / area, 1) if area else 0

    # GHGI uses user-supplied emission factors (kgCO₂e/kWh). If none are provided,
    # GHGI is left undefined (None) rather than shown as a misleading value.
    ef_elec, ef_gas, ef_other = float(ef_elec or 0), float(ef_gas or 0), float(ef_other or 0)
    factors_on = (ef_elec > 0) or (ef_gas > 0) or (ef_other > 0)
    ghgi = round((elec*ef_elec + gas*ef_gas + other_fuel*ef_other) / area, 1) if (factors_on and area) else None

    return {
        "area": area, "total_energy": total,
        "total_eui":      eui(total),
        "tedi":           round(tedi, 1),
        "unmet_heating":  int(round(unmet_h)),
        "unmet_cooling":  int(round(unmet_c)),
        "unmet_total":    int(round(unmet_t)),
        "elec_eui":       eui(elec),
        "gas_eui":        eui(gas),
        "other_fuel_eui": eui(other_fuel),
        "heat_eui":       eui(heat),
        "cool_eui":       eui(cool),
        "fan_eui":         eui(fans),
        "central_fan_eui": eui(central_fan),
        "local_fan_eui":   eui(local_fan),
        "exhaust_fan_eui": eui(exhaust_fan),
        "ltg_eui":        eui(ltg),
        "dhw_eui":        eui(dhw),
        "pumps_eui":      eui(pumps),
        "recept_eui":     eui(recept),
        "heat_rej_eui":   eui(heat_rej),
        "ext_ltg_eui":    eui(ext_ltg),
        "process_eui":    eui(process),
        "ghgi":           ghgi,
    }

def generate_flags(kpis, bm, building_type):
    flags = []
    if not bm:
        flags.append(("info","ℹ️","No benchmark found for this combination. KPIs calculated but no comparison available."))
        return flags
    # All thresholds use ±15% of median EUI
    med       = bm["median_eui"]
    good      = bm["good_eui"]    # median * 0.85
    high      = bm["high_eui"]    # median * 1.15

    def end_med(pct): return med * pct / 100
    def end_good(pct): return end_med(pct) * 0.85
    def end_high(pct): return end_med(pct) * 1.15

    fan_med  = end_med(bm["fan_pct"])
    heat_med = end_med(bm["heat_pct"])
    cool_med = end_med(bm["cool_pct"])
    ltg_med  = end_med(bm["ltg_pct"])

    # Total EUI
    if kpis["total_eui"] > high:
        flags.append(("fail","✗",
            f"Total EUI ({kpis['total_eui']} kWh/m²·yr) is more than 15% above the benchmark median "
            f"({med} kWh/m²·yr). High flag threshold: {high} kWh/m²·yr."))
    elif kpis["total_eui"] < good * 0.6:
        flags.append(("warn","⚠",
            f"Total EUI ({kpis['total_eui']} kWh/m²·yr) is unusually low — confirm all end-uses are modelled."))
    elif kpis["total_eui"] <= good:
        flags.append(("pass","✓",
            f"Total EUI ({kpis['total_eui']} kWh/m²·yr) is below the benchmark median ({med} kWh/m²·yr) — good performance."))
    else:
        flags.append(("pass","✓",
            f"Total EUI ({kpis['total_eui']} kWh/m²·yr) is within ±15% of the benchmark median ({med} kWh/m²·yr)."))

    # Fan energy
    if fan_med > 0 and kpis["fan_eui"] > end_high(bm["fan_pct"]):
        flags.append(("fail","✗",
            f"Fan energy ({kpis['fan_eui']} kWh/m²·yr) is more than 15% above benchmark median "
            f"({round(fan_med,1)} kWh/m²·yr) — verify AHU schedules and fan sizing."))
    elif kpis["fan_eui"] == 0:
        flags.append(("warn","⚠","Fan energy is zero — confirm fan systems are included in the model."))
    else:
        flags.append(("pass","✓",f"Fan energy ({kpis['fan_eui']} kWh/m²·yr) is within expected range."))

    # Cooling
    if kpis["cool_eui"] == 0:
        flags.append(("warn","⚠","No cooling energy — confirm whether a mechanical cooling system exists."))
    elif kpis["cool_eui"] < cool_med * 0.25:
        flags.append(("warn","⚠",
            f"Cooling energy ({kpis['cool_eui']} kWh/m²·yr) is very low compared to benchmark — verify cooling system modelling."))
    elif kpis["cool_eui"] > end_high(bm["cool_pct"]):
        flags.append(("fail","✗",
            f"Cooling energy ({kpis['cool_eui']} kWh/m²·yr) is more than 15% above benchmark median "
            f"({round(cool_med,1)} kWh/m²·yr)."))

    # Heating
    if heat_med > 0 and kpis["heat_eui"] > end_high(bm["heat_pct"]):
        flags.append(("fail","✗",
            f"Heating energy ({kpis['heat_eui']} kWh/m²·yr) is more than 15% above benchmark median "
            f"({round(heat_med,1)} kWh/m²·yr) — review envelope and heating schedules."))
    elif heat_med > 0 and kpis["heat_eui"] <= end_high(bm["heat_pct"]):
        flags.append(("pass","✓",f"Heating energy ({kpis['heat_eui']} kWh/m²·yr) is within expected range."))

    # Simultaneous heating + cooling
    if kpis["heat_eui"] > end_high(bm["heat_pct"]) and kpis["cool_eui"] > end_high(bm["cool_pct"]):
        flags.append(("warn","⚠","Both heating and cooling are above the high threshold — possible simultaneous heating/cooling or control issue."))

    # DHW
    if building_type in DHW_BUILDINGS and kpis["dhw_eui"] == 0:
        flags.append(("fail","✗",f"DHW energy is zero for a {building_type} — domestic hot water is typically required."))
    elif kpis["dhw_eui"] > 0:
        flags.append(("pass","✓",f"DHW energy present ({kpis['dhw_eui']} kWh/m²·yr)."))

    # Lighting
    if ltg_med > 0 and kpis["ltg_eui"] > end_high(bm["ltg_pct"]):
        flags.append(("fail","✗",
            f"Lighting energy ({kpis['ltg_eui']} kWh/m²·yr) is more than 15% above benchmark median "
            f"({round(ltg_med,1)} kWh/m²·yr) — verify LPD values against NECB."))

    # GHGI (only when emission factors were provided)
    if kpis.get("ghgi") is not None:
        if kpis["ghgi"] > bm["median_ghgi"] * 1.15:
            flags.append(("fail","✗",
                f"GHGI ({kpis['ghgi']} kgCO₂e/m²·yr) is more than 15% above benchmark median "
                f"({bm['median_ghgi']} kgCO₂e/m²·yr) — review fuel mix and emission factors."))
        else:
            flags.append(("pass","✓",
                f"GHGI ({kpis['ghgi']} kgCO₂e/m²·yr) is within acceptable range (benchmark median: {bm['median_ghgi']})."))
    return flags

def status_color(val, good, median, high):
    if val <= good:   return "🟢"
    if val <= median: return "🟡"
    if val <= high:   return "🟠"
    return "🔴"

def build_comparison_rows(kpis, bm):
    """Per-metric comparison rows with the threshold-vs-median QA Status (the Auto Flag)."""
    def be(pct): return round(bm["median_eui"] * pct / 100, 1)
    def bs(val, pct):
        med = bm["median_eui"] * pct / 100
        return status_color(val, med*0.85, med, med*1.15)
    def pdiff(your, med):  # percentage difference vs benchmark median
        return f"{(your-med)/med*100:+.0f}%" if med else "n/a"
    rows = [
        {"End Use":"Total EUI (kWh/m²·yr)","Your Model":kpis["total_eui"],"Benchmark Median":bm["median_eui"],
         "Difference":pdiff(kpis["total_eui"], bm["median_eui"]),
         "QA Status":status_color(kpis["total_eui"],bm["good_eui"],bm["median_eui"],bm["high_eui"])},
    ]
    # TEDI sits right next to Total EUI when both a benchmark and a model value are present.
    if bm.get("median_tedi") and kpis.get("tedi", 0) > 0:
        rows.append(
            {"End Use":"TEDI (kWh/m²·yr)","Your Model":kpis["tedi"],"Benchmark Median":bm["median_tedi"],
             "Difference":pdiff(kpis["tedi"], bm["median_tedi"]),
             "QA Status":status_color(kpis["tedi"],bm["good_tedi"],bm["median_tedi"],bm["high_tedi"])})
    rows += [
        {"End Use":"Heating EUI (kWh/m²·yr)","Your Model":kpis["heat_eui"],"Benchmark Median":be(bm["heat_pct"]),
         "Difference":pdiff(kpis["heat_eui"], be(bm["heat_pct"])),"QA Status":bs(kpis["heat_eui"],bm["heat_pct"])},
        {"End Use":"Cooling EUI (kWh/m²·yr)","Your Model":kpis["cool_eui"],"Benchmark Median":be(bm["cool_pct"]),
         "Difference":pdiff(kpis["cool_eui"], be(bm["cool_pct"])),"QA Status":bs(kpis["cool_eui"],bm["cool_pct"])},
        {"End Use":"Fan EUI (kWh/m²·yr)","Your Model":kpis["fan_eui"],"Benchmark Median":be(bm["fan_pct"]),
         "Difference":pdiff(kpis["fan_eui"], be(bm["fan_pct"])),"QA Status":bs(kpis["fan_eui"],bm["fan_pct"])},
        {"End Use":"Lighting EUI (kWh/m²·yr)","Your Model":kpis["ltg_eui"],"Benchmark Median":be(bm["ltg_pct"]),
         "Difference":pdiff(kpis["ltg_eui"], be(bm["ltg_pct"])),"QA Status":bs(kpis["ltg_eui"],bm["ltg_pct"])},
        {"End Use":"DHW EUI (kWh/m²·yr)","Your Model":kpis["dhw_eui"],"Benchmark Median":be(bm["dhw_pct"]),
         "Difference":pdiff(kpis["dhw_eui"], be(bm["dhw_pct"])),"QA Status":bs(kpis["dhw_eui"],bm["dhw_pct"])},
        {"End Use":"Receptacle EUI (kWh/m²·yr)","Your Model":kpis.get("recept_eui",0),"Benchmark Median":be(bm["recept_pct"]),
         "Difference":pdiff(kpis.get("recept_eui",0), be(bm["recept_pct"])),"QA Status":bs(kpis.get("recept_eui",0),bm["recept_pct"])},
        {"End Use":"Pumps EUI (kWh/m²·yr)","Your Model":kpis["pumps_eui"],"Benchmark Median":be(bm["pumps_pct"]),
         "Difference":pdiff(kpis["pumps_eui"], be(bm["pumps_pct"])),"QA Status":bs(kpis["pumps_eui"],bm["pumps_pct"])},
    ]
    if kpis.get("ghgi") is not None:
        rows.append(
            {"End Use":"GHGI (kgCO₂e/m²·yr)","Your Model":kpis["ghgi"],"Benchmark Median":bm["median_ghgi"],
             "Difference":pdiff(kpis["ghgi"], bm["median_ghgi"]),
             "QA Status":status_color(kpis["ghgi"],bm["median_ghgi"]*0.85,bm["median_ghgi"],bm["median_ghgi"]*1.15)})
    return rows

def flags_from_comparison(rows, kpis, bm, building_type):
    """Build the QA/QC Flags list directly from the Auto Flag column so the two never
    disagree. Special-cases a missing required system (e.g. DHW = 0), which a pure
    threshold check would mistakenly read as "low = good"."""
    LEVEL = {"🟢":"pass", "🟡":"pass", "🟠":"warn", "🔴":"fail"}
    ICON  = {"pass":"✓", "warn":"⚠", "fail":"✗"}
    HINT  = {
        "Total EUI":  "review overall model inputs, schedules and envelope.",
        "TEDI":       "review envelope, airtightness and ventilation heat recovery.",
        "Heating":    "review envelope inputs and heating schedules.",
        "Cooling":    "verify cooling system sizing and controls.",
        "Fan":        "verify AHU schedules and fan sizing.",
        "Lighting":   "verify LPD values against NECB.",
        "DHW":        "review DHW demand and system efficiency.",
        "Receptacle": "verify plug-load assumptions.",
        "Pumps":      "verify pump sizing and run hours.",
        "GHGI":       "review fuel mix and emission factors.",
    }
    def hint_for(name):
        return next((v for k, v in HINT.items() if name.startswith(k)), "")

    flags = []
    for r in rows:
        name, status = r["End Use"], r["QA Status"]
        your, bench, diff = r["Your Model"], r["Benchmark Median"], r["Difference"]

        # Missing required systems — thresholds alone would call a 0 value "good".
        if name.startswith("DHW") and kpis["dhw_eui"] == 0 and building_type in DHW_BUILDINGS:
            flags.append(("fail","✗", f"DHW energy is zero for a {building_type} — domestic hot water is typically required."))
            continue
        if name.startswith("Cooling") and kpis["cool_eui"] == 0:
            flags.append(("warn","⚠", "No cooling energy — confirm whether a mechanical cooling system exists."))
            continue
        if name.startswith("Total EUI") and kpis["total_eui"] < bm["good_eui"] * 0.6:
            flags.append(("warn","⚠", f"Total EUI ({kpis['total_eui']} kWh/m²·yr) is unusually low — confirm all end-uses are modelled."))
            continue

        level = LEVEL.get(status, "warn")
        if level == "fail":
            msg = f"{name}: {your} vs benchmark median {bench} ({diff}) — more than 15% above median; {hint_for(name)}"
        elif level == "warn":
            msg = f"{name}: {your} vs benchmark median {bench} ({diff}) — near the +15% high threshold; review."
        else:
            msg = f"{name}: {your} vs benchmark median {bench} ({diff}) — within the expected range."
        flags.append((level, ICON[level], msg))

    # TEDI not provided but a TEDI benchmark exists.
    if bm.get("median_tedi") and kpis.get("tedi", 0) == 0:
        flags.append(("warn","⚠", "TEDI not provided — add Thermal Energy Demand Intensity to compare against the benchmark."))

    # Cross-check: heating and cooling both above the high threshold.
    heat_red = any(r["End Use"].startswith("Heating") and r["QA Status"] == "🔴" for r in rows)
    cool_red = any(r["End Use"].startswith("Cooling") and r["QA Status"] == "🔴" for r in rows)
    if heat_red and cool_red:
        flags.append(("warn","⚠", "Both heating and cooling are above the high threshold — possible simultaneous heating/cooling or control issue."))
    return flags

def make_pie(labels, values, title):
    fig = go.Figure(go.Pie(
        labels=labels, values=values, hole=0.42,
        marker=dict(colors=END_USE_COLORS, line=dict(color="#fff", width=2)),
        textinfo="label+percent", textfont=dict(size=12),
        hovertemplate="<b>%{label}</b><br>%{value} kWh/m²·yr<br>%{percent}<extra></extra>",
    ))
    fig.update_layout(
        title=dict(text=title, font=dict(size=13, color="#0f4c81")),
        showlegend=True, legend=dict(orientation="v", x=1.02, y=0.5, font=dict(size=11)),
        margin=dict(t=40,b=10,l=10,r=10), height=320,
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
    )
    return fig

def build_pdf_report(meta, kpis, bm, flags, pct, comparison_df):
    """Build a one-file PDF QA/QC report (replaces the previous Excel export).

    Emoji status dots and subscript characters are mapped to plain text so they
    render reliably in the PDF's base fonts. The reviewer's overridden QA Status
    and Comments (from the editable Benchmark Comparison table) are included.
    """
    from reportlab.lib.pagesizes import letter
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from xml.sax.saxutils import escape as _xml

    STATUS_WORD = {"🟢": "Pass", "🟡": "OK", "🟠": "Watch", "🔴": "Fail"}
    FLAG_WORD   = {"pass": "PASS", "warn": "REVIEW", "fail": "FAIL", "info": "INFO"}

    def desub(x):                       # subscript 2 -> 2 (base fonts can't render ₂)
        return str(x).replace("₂", "2")

    def destatus(x):                    # emoji dot -> word
        s = desub(x)
        for k, v in STATUS_WORD.items():
            s = s.replace(k, v)
        return s.strip()

    def para(x, style):                 # XML-safe paragraph
        return Paragraph(_xml(desub(x)), style)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=letter,
        topMargin=1.4*cm, bottomMargin=1.4*cm, leftMargin=1.4*cm, rightMargin=1.4*cm,
        title=f"QA/QC Report - {meta['project_name'] or 'Project'}",
    )
    styles = getSampleStyleSheet()
    h1    = ParagraphStyle("h1", parent=styles["Heading1"], textColor=colors.HexColor("#0f4c81"), fontSize=17, spaceAfter=2)
    h2    = ParagraphStyle("h2", parent=styles["Heading2"], textColor=colors.HexColor("#0f4c81"), fontSize=12, spaceBefore=8, spaceAfter=4)
    body  = styles["BodyText"]
    small = ParagraphStyle("small", parent=body, fontSize=9, textColor=colors.HexColor("#64748b"))
    cell  = ParagraphStyle("cell", parent=body, fontSize=8, leading=10)

    elems = []

    # Title + metadata line
    elems.append(para(meta["project_name"] or "Project Results", h1))
    sub = meta['building_type']
    if meta.get('city'):
        sub += f" · {meta['city']}"
    sub += f" · Climate Zone {meta['climate_zone']}"
    if meta.get("subtype", "General") not in ("", "General", "All"):
        sub += f" · {meta['subtype']}"
    if not meta.get('city'):
        sub += " · zone average"
    sub += f" · {meta['model_type']}"
    if meta.get("phase"):
        sub += f" · {meta['phase']}"
    sub += f" · {meta['software']} · {meta['date']}"
    elems.append(para(sub, small))
    elems.append(Spacer(1, 8))

    # Overall result
    fc = {"pass": 0, "warn": 0, "fail": 0}
    for f in flags:
        fc[f[0]] = fc.get(f[0], 0) + 1
    overall = "Issues Found" if fc["fail"] > 0 else "Review Required" if fc["warn"] > 0 else "All Clear"
    pct_line = f"  |  Benchmark percentile: {pct}th (lower = better)" if pct else ""
    elems.append(Paragraph(
        desub(f"<b>Overall QA/QC:</b> {overall} — Pass {fc['pass']} · Review {fc['warn']} · Fail {fc['fail']}{pct_line}"),
        body))
    elems.append(Spacer(1, 10))

    # Energy summary
    elems.append(para("Energy Summary", h2))
    summ = [
        ["Metric", "Value", "Benchmark median"],
        ["Total EUI (kWh/m2/yr)",  kpis["total_eui"], bm["median_eui"]  if bm else "-"],
        ["TEDI (kWh/m2/yr)",       kpis.get("tedi", 0) or "-", (bm.get("median_tedi") or "-") if bm else "-"],
        ["GHGI (kgCO2e/m2/yr)",    kpis.get("ghgi") if kpis.get("ghgi") is not None else "-",      bm["median_ghgi"] if bm else "-"],
        ["Electricity EUI",        kpis["elec_eui"],  "-"],
        ["Gas EUI",                kpis["gas_eui"],   "-"],
        ["Floor area (m2)",        round(kpis["area"]), "-"],
        ["Unmet hours (total)",    kpis.get("unmet_total", 0), "-"],
        ["Unmet hours (heating / cooling)", f"{kpis.get('unmet_heating',0)} / {kpis.get('unmet_cooling',0)}", "-"],
    ]
    t = Table(summ, hAlign="LEFT", colWidths=[7*cm, 4*cm, 5*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0f4c81")),
        ("TEXTCOLOR",  (0, 0), (-1, 0), colors.white),
        ("FONTSIZE",   (0, 0), (-1, -1), 9),
        ("GRID",       (0, 0), (-1, -1), 0.5, colors.HexColor("#e2e8f0")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
        ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    elems.append(t)

    # Benchmark comparison with reviewer overrides + comments
    if bm and comparison_df is not None:
        elems.append(para("Benchmark Comparison & Reviewer Notes", h2))
        header = ["End Use", "Your Model", "Bench. Median", "Diff.", "Auto Flag", "QA Status", "Comment"]
        data = [header]
        for _, rr in comparison_df.iterrows():
            data.append([
                para(rr["End Use"], cell),
                rr["Your Model"],
                rr["Benchmark Median"],
                str(rr["Difference"]),
                destatus(rr["Auto Flag"]),
                destatus(rr["QA Status"]),
                para(rr.get("Comment", "") or "", cell),
            ])
        ct = Table(data, hAlign="LEFT", repeatRows=1,
                   colWidths=[4.0*cm, 2.0*cm, 2.4*cm, 1.5*cm, 1.7*cm, 1.7*cm, 4.0*cm])
        ct.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0f4c81")),
            ("TEXTCOLOR",  (0, 0), (-1, 0), colors.white),
            ("FONTSIZE",   (0, 0), (-1, -1), 8),
            ("GRID",       (0, 0), (-1, -1), 0.5, colors.HexColor("#e2e8f0")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
            ("VALIGN",     (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))
        elems.append(ct)

    # QA/QC flags
    elems.append(para("QA/QC Flags", h2))
    for level, icon, msg in flags:
        safe_msg = _xml(desub(msg))
        elems.append(Paragraph(f"<b>[{FLAG_WORD.get(level, level.upper())}]</b> {safe_msg}", body))
        elems.append(Spacer(1, 2))

    doc.build(elems)
    return buf.getvalue()

# ── Session state ─────────────────────────────────────────────────────────────
if "step"    not in st.session_state: st.session_state.step    = 1
if "vals"    not in st.session_state: st.session_state.vals    = {}
if "results" not in st.session_state: st.session_state.results = None
if "page"    not in st.session_state: st.session_state.page    = "QA/QC Tool"
if "bm_reload" not in st.session_state: st.session_state.bm_reload = 0

# ── Load benchmarks ───────────────────────────────────────────────────────────
BENCHMARKS = load_benchmarks()
ALL_BUILDING_TYPES = sorted(set(k[0] for k in BENCHMARKS)) or ["School","Office","Retail","Hospital","Residential","Warehouse"]
CITIES             = sorted(set(k[1] for k in BENCHMARKS)) or ["Edmonton","Calgary","Vancouver","Toronto"]
CLIMATE_ZONES      = sorted(set(k[2] for k in BENCHMARKS)) or ["4","5","6","7","8"]
# Subtypes derived dynamically per selection in Step 3

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚡ Energy Benchmarking")
    st.markdown("**QA/QC Platform** | MVP v1.0")
    st.divider()
    page = st.radio("Navigate", ["QA/QC Tool", "📚 Benchmark Explorer", "⚙️ Manage Benchmarks"], label_visibility="collapsed")
    st.session_state.page = page
    st.divider()
    if page == "QA/QC Tool":
        steps = ["1. Upload / Enter Data","2. Map Columns","3. Building Info","4. Results"]
        for i, s in enumerate(steps, 1):
            if i < st.session_state.step:    st.markdown(f"✅ {s}")
            elif i == st.session_state.step:  st.markdown(f"**→ {s}**")
            else:                             st.markdown(f"&nbsp;&nbsp;&nbsp;{s}")
        st.divider()
        if st.button("🔄 Start Over", use_container_width=True):
            for k in ["step","vals","results","headers","csv_df","mapping","meta","ref_csv_df","ref_vals","compliance_code"]:
                if k in st.session_state: del st.session_state[k]
            st.session_state.step = 1
            st.rerun()
    st.divider()
    st.markdown(f"**{len(BENCHMARKS)} benchmarks loaded**")



# ══════════════════════════════════════════════════════════════════════════════
#  PAGE: MANAGE BENCHMARKS
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.page == "⚙️ Manage Benchmarks":
    st.markdown("# ⚙️ Manage Benchmark Database")
    st.markdown("All benchmark data is stored in **Google Sheets**. Edit the sheet directly and changes appear in the app within 1 minute.")
    st.info(f"📊 [Open Google Sheet to edit benchmarks](https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit)", icon="📊")
    st.divider()

    tab_view, tab_howto = st.tabs(["📋 View All", "✏️ How to Edit"])

    # ── VIEW ──
    with tab_view:
        st.markdown("### All Benchmarks")
        if st.button("🔄 Refresh from Google Sheets", use_container_width=False):
            load_benchmarks.clear()
            st.rerun()
        rows = []
        for (btype, city, zone, subtype), bm in BENCHMARKS.items():
            rows.append({
                "Building Type": btype, "Province": bm.get("province",""), "City": city, "Zone": zone,
                "Subtype": subtype,
                "Median EUI": bm["median_eui"],
                "Median TEDI": bm.get("median_tedi", 0),
                "Good (−15%)": bm["good_eui"],
                "High Flag (+15%)": bm["high_eui"],
                "GHGI": bm["median_ghgi"],
                "Heating %": bm["heat_pct"], "Cooling %": bm["cool_pct"],
                "Fan %": bm["fan_pct"], "Lighting %": bm["ltg_pct"],
                "DHW %": bm["dhw_pct"], "Receptacle %": bm["recept_pct"],
                "Pumps %": bm["pumps_pct"], "Electricity %": bm.get("elec_pct",0),
                "NaturalGas %": bm.get("gas_pct",0), "Heat Rejection %": bm.get("heat_rej_pct",0),
                "Miscellaneous %": bm.get("misc_pct",0),
            })
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
        st.caption(f"Total: {len(rows)} benchmark records — refreshes automatically every 60 seconds")

    # ── HOW TO EDIT ──
    with tab_howto:
        st.markdown("### How to add or edit benchmarks")
        st.markdown(f"**[👉 Click here to open the Google Sheet]"
                    f"(https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit)**")
        st.markdown("""
**To add a new benchmark:**
1. Open the Google Sheet using the link above
2. Scroll to the next empty row
3. Fill in all columns — Building Type, City, Climate Zone, EUI values, percentages, and percentile data
4. Save — the app picks up the change within 60 seconds

**To edit an existing benchmark:**
1. Open the Google Sheet
2. Find the row you want to change
3. Edit the cell directly
4. Save — changes appear in the app within 60 seconds

**To delete a benchmark:**
1. Open the Google Sheet
2. Right-click the row number → Delete row
3. Save

**Column guide:**
- Percentile Data: 10 numbers separated by commas e.g. `95,110,125,141,155,165,180,195,210,230`
- All percentages should sum to approximately 100%
        """)


# ══════════════════════════════════════════════════════════════════════════════
#  PAGE: BENCHMARK EXPLORER
# ══════════════════════════════════════════════════════════════════════════════
elif st.session_state.page == "📚 Benchmark Explorer":
    st.markdown("# 📚 Benchmark Explorer")
    st.markdown("Browse the benchmark database by building type and location — no upload required.")
    st.divider()

    # ── Row 1: required filters ───────────────────────────────────────────────
    cf1, cf2, cf3 = st.columns(3)
    with cf1:
        bx_type = st.selectbox("Building Type", ALL_BUILDING_TYPES)
    all_zones = sorted({k[2] for k,v in BENCHMARKS.items() if k[0]==bx_type and k[2]})
    with cf2:
        bx_zone = st.selectbox("Climate Zone", all_zones if all_zones else ["—"])
    # HVAC System — required; list all distinct values for this type + zone
    all_hvac = sorted({str(v.get("heating_system","")) for k,v in BENCHMARKS.items()
                       if k[0]==bx_type and k[2]==bx_zone and v.get("heating_system")})
    all_hvac_label = "All HVAC systems"
    with cf3:
        bx_hvac = st.selectbox("HVAC System", [all_hvac_label] + all_hvac)

    # ── Row 2: optional refinement filters (all in one row) ──────────────────
    opt1, opt2, opt3, opt4 = st.columns(4)

    provs = sorted({v["province"] for k,v in BENCHMARKS.items()
                    if k[0]==bx_type and k[2]==bx_zone
                    and v.get("province") and str(v["province"]).strip().lower() not in ("","nan")})
    all_provs_label = "All provinces"
    with opt1:
        bx_prov = st.selectbox("Province (optional)", [all_provs_label] + provs)

    def _city_filter(k, v):
        if k[0]!=bx_type or k[2]!=bx_zone: return False
        if bx_prov != all_provs_label and v.get("province")!=bx_prov: return False
        if bx_hvac != all_hvac_label and str(v.get("heating_system",""))!=bx_hvac: return False
        return True
    cities = sorted({k[1] for k,v in BENCHMARKS.items() if _city_filter(k,v)})
    all_cities_label = "All cities"
    with opt2:
        bx_city = st.selectbox("City (optional)", [all_cities_label] + cities)

    def _bm_filter(k, v):
        if not _city_filter(k,v): return False
        if bx_city != all_cities_label and k[1]!=bx_city: return False
        return True
    filtered_bms = {k:v for k,v in BENCHMARKS.items() if _bm_filter(k,v)}

    proj_years  = sorted({str(v.get("project_year","")) for v in filtered_bms.values() if v.get("project_year")})
    audit_types = sorted({str(v.get("audit_new",""))    for v in filtered_bms.values() if v.get("audit_new")})
    all_label = "All"
    with opt3:
        bx_year  = st.selectbox("Project Year (optional)", [all_label] + proj_years)
    with opt4:
        bx_audit = st.selectbox("Audit / New (optional)",  [all_label] + audit_types)

    # ── Resolve matching benchmarks ─────────────────────────────────────────
    def _final_filter(k, v):
        if not _bm_filter(k,v): return False
        if bx_year  != all_label and str(v.get("project_year","")) != bx_year:  return False
        if bx_audit != all_label and str(v.get("audit_new",""))    != bx_audit: return False
        return True
    final_bms = {k:v for k,v in BENCHMARKS.items() if _final_filter(k,v)}

    if not final_bms:
        scope = f"{bx_type} · Climate Zone {bx_zone}"
        if bx_hvac  != all_hvac_label:  scope += f" · {bx_hvac}"
        if bx_prov  != all_provs_label: scope += f" · {bx_prov}"
        if bx_city  != all_cities_label: scope += f" · {bx_city}"
        st.warning(f"No benchmark data found for **{scope}** with the selected filters. Try broadening your selection.")
        st.stop()

    agg = average_benchmarks(list(final_bms.values()))
    scope_parts = [f"Climate Zone {bx_zone}"]
    if bx_hvac  != all_hvac_label:  scope_parts.append(bx_hvac)
    if bx_prov  != all_provs_label:  scope_parts.append(bx_prov)
    if bx_city  != all_cities_label: scope_parts.append(bx_city)
    if bx_year  != all_label:        scope_parts.append(f"Year {bx_year}")
    if bx_audit != all_label:        scope_parts.append(bx_audit)
    scope_str = " · ".join(scope_parts)
    matches = {(bx_type, bx_city if bx_city!=all_cities_label else "", bx_zone, "filtered"): agg}
    if agg["_n"] == 1:
        st.info(f"Showing 1 benchmark for **{bx_type} · {scope_str}**.")
    else:
        st.info(f"Showing the **average of {agg['_n']} benchmarks** for **{bx_type} · {scope_str}**. Refine with the optional filters above.")

    for (btype, bcity, bzone, bsubtype), bm in matches.items():
        st.markdown(f"## 🏢 {btype} · {scope_str}")

        m1, m2, m3 = st.columns(3)
        m1.metric("📊 Median EUI",          f"{bm['median_eui']} kWh/m²·yr",
                  help="Typical performance for this building type and location.")
        m2.metric("🟢 Good Practice (−15%)", f"{bm['good_eui']} kWh/m²·yr",
                  help="15% below median — used as the pass threshold in QA/QC.")
        m3.metric("🔴 High Flag (+15%)",     f"{bm['high_eui']} kWh/m²·yr",
                  help="15% above median — used as the fail threshold in QA/QC.")

        if bm.get("median_tedi"):
            t1, t2, t3 = st.columns(3)
            t1.metric("🔥 Median TEDI",          f"{bm['median_tedi']} kWh/m²·yr",
                      help="Thermal Energy Demand Intensity — heating + ventilation demand per m². As important as EUI.")
            t2.metric("🟢 Good Practice (−15%)", f"{bm['good_tedi']} kWh/m²·yr",
                      help="15% below median TEDI — pass threshold in QA/QC.")
            t3.metric("🔴 High Flag (+15%)",     f"{bm['high_tedi']} kWh/m²·yr",
                      help="15% above median TEDI — fail threshold in QA/QC.")

        # ── end-use data ──────────────────────────────────────────────────────
        eu_labels = ["Heating","Cooling","Fans","Lighting","DHW","Receptacle","Pumps","Heat Rejection","Miscellaneous"]
        eu_pcts   = [bm["heat_pct"], bm["cool_pct"], bm["fan_pct"], bm["ltg_pct"],
                     bm["dhw_pct"], bm["recept_pct"], bm["pumps_pct"],
                     bm.get("heat_rej_pct",0), bm.get("misc_pct",0)]
        med_vals  = [round(bm["median_eui"]*p/100,1) for p in eu_pcts]
        good_vals = [round(bm["good_eui"]*p/100,1)   for p in eu_pcts]
        high_vals = [round(bm["high_eui"]*p/100,1)   for p in eu_pcts]

        # ── energy source data ─────────────────────────────────────────────────
        src_labels = ["Electricity","Natural Gas"]
        src_pcts   = [bm.get("elec_pct",0), bm.get("gas_pct",0)]
        src_med    = [round(bm["median_eui"]*p/100,1) for p in src_pcts]

        cp1, cp2 = st.columns(2)
        with cp1:
            subtype_label = f" ({bsubtype})" if bsubtype not in ("General", "All", "filtered") else ""
            # Left: energy source pie
            src_l = [l for l,v in zip(src_labels, src_med) if v>0]
            src_v = [v for v in src_med if v>0]
            if src_l:
                src_colors = ["#3b82f6","#ef4444"][:len(src_l)]
                fig_src = go.Figure(go.Pie(
                    labels=src_l, values=src_v, hole=0.42,
                    marker=dict(colors=src_colors, line=dict(color="#fff",width=2)),
                    textinfo="label+percent", textfont=dict(size=12),
                    hovertemplate="<b>%{label}</b><br>%{value} kWh/m²·yr<br>%{percent}<extra></extra>",
                ))
                fig_src.update_layout(
                    title=dict(text=f"Energy Source Split — {btype} · {bcity}{subtype_label}", font=dict(size=13,color="#0f4c81")),
                    showlegend=True, legend=dict(orientation="v",x=1.02,y=0.5,font=dict(size=11)),
                    margin=dict(t=40,b=10,l=10,r=10), height=320,
                    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                )
                st.plotly_chart(fig_src, use_container_width=True)
            else:
                st.caption("No Electricity % / NaturalGas % data in the sheet for this benchmark.")
        with cp2:
            # Right: end-use pie (includes Miscellaneous)
            pie_l = [l for l,v in zip(eu_labels, med_vals) if v>0]
            pie_v = [v for v in med_vals if v>0]
            st.plotly_chart(make_pie(pie_l, pie_v, f"Median End-Use Split — {btype} · {bcity}{subtype_label}"), use_container_width=True)

        # ── bar chart — end-uses only; heat rejection + miscellaneous included ──
        bar_pairs = [(l,v) for l,v in zip(eu_labels, med_vals) if v>0]
        bar_labels_f = [l for l,v in bar_pairs]
        bar_med_f    = [v for l,v in bar_pairs]
        fig_bar = go.Figure()
        fig_bar.add_trace(go.Bar(name="Median EUI", x=bar_labels_f, y=bar_med_f,
                                 marker_color="#0f4c81", opacity=0.85,
                                 error_y=dict(type="data", symmetric=True,
                                              array=[round(v*0.15,1) for v in bar_med_f],
                                              color="#94a3b8", thickness=1.5, width=4)))
        fig_bar.update_layout(template="plotly_white", height=320,
                              title=dict(text="End-Use Median EUI (bars show ±15% range)", font=dict(size=13,color="#0f4c81")),
                              yaxis_title="EUI (kWh/m²·yr)", showlegend=False, margin=dict(t=40,b=10))
        st.plotly_chart(fig_bar, use_container_width=True)

        st.markdown("#### End-Use Breakdown Table")
        st.caption("Good and High Flag are calculated as ±15% of median.")
        st.dataframe(pd.DataFrame({
            "End Use": eu_labels,
            "Share (%)": eu_pcts,
            "Median (kWh/m²·yr)": med_vals,
        }), use_container_width=True, hide_index=True)

        st.markdown("#### Portfolio Percentile Distribution — EUI")
        sorted_pcts = sorted(bm["pct_data"])
        bar_colors  = ["#16a34a" if v<=bm["good_eui"] else "#d97706" if v<=bm["median_eui"] else "#ea580c" if v<=bm["high_eui"] else "#dc2626" for v in sorted_pcts]
        fig_pct = go.Figure(go.Bar(x=[f"{i*10}th" for i in range(1,11)], y=sorted_pcts, marker_color=bar_colors,
            hovertemplate="<b>%{x} percentile</b><br>EUI: %{y} kWh/m²·yr<extra></extra>"))
        fig_pct.add_hline(y=bm["good_eui"],   line_dash="dot",  line_color="#16a34a", line_width=1.5)
        fig_pct.add_hline(y=bm["median_eui"], line_dash="dash", line_color="#d97706", line_width=2)
        fig_pct.add_hline(y=bm["high_eui"],   line_dash="dot",  line_color="#dc2626", line_width=1.5)
        fig_pct.update_layout(template="plotly_white", height=300, xaxis_title="Percentile", yaxis_title="EUI (kWh/m²·yr)", margin=dict(t=20,b=10))
        st.plotly_chart(fig_pct, use_container_width=True)
        st.caption(f"🟩 Good (−15%): {bm['good_eui']}  ·  🟧 Median: {bm['median_eui']}  ·  🟥 High flag (+15%): {bm['high_eui']}  kWh/m²·yr")

        if bm.get("median_tedi") and bm.get("tedi_pct_data"):
            st.markdown("#### Portfolio Percentile Distribution — TEDI")
            sorted_tedi = sorted(bm["tedi_pct_data"])
            tedi_colors = ["#16a34a" if v<=bm["good_tedi"] else "#d97706" if v<=bm["median_tedi"] else "#ea580c" if v<=bm["high_tedi"] else "#dc2626" for v in sorted_tedi]
            fig_tedi = go.Figure(go.Bar(x=[f"{i*10}th" for i in range(1,11)], y=sorted_tedi, marker_color=tedi_colors,
                hovertemplate="<b>%{x} percentile</b><br>TEDI: %{y} kWh/m²·yr<extra></extra>"))
            fig_tedi.add_hline(y=bm["good_tedi"],   line_dash="dot",  line_color="#16a34a", line_width=1.5)
            fig_tedi.add_hline(y=bm["median_tedi"], line_dash="dash", line_color="#d97706", line_width=2)
            fig_tedi.add_hline(y=bm["high_tedi"],   line_dash="dot",  line_color="#dc2626", line_width=1.5)
            fig_tedi.update_layout(template="plotly_white", height=300, xaxis_title="Percentile", yaxis_title="TEDI (kWh/m²·yr)", margin=dict(t=20,b=10))
            st.plotly_chart(fig_tedi, use_container_width=True)
            st.caption(f"🟩 Good (−15%): {bm['good_tedi']}  ·  🟧 Median: {bm['median_tedi']}  ·  🟥 High flag (+15%): {bm['high_tedi']}  kWh/m²·yr")

        st.divider()


# ══════════════════════════════════════════════════════════════════════════════
#  PAGE: QA/QC TOOL
# ══════════════════════════════════════════════════════════════════════════════
else:
    st.markdown("# ⚡ Energy Model Benchmarking & QA/QC Platform")
    st.markdown("Upload simulation results → Calculate KPIs → Compare to benchmarks → Get QA/QC flags")
    st.divider()

    if st.session_state.step == 1:
        st.markdown("## Step 1 — Upload Simulation Results or Enter Manually")
        tab_upload, tab_manual = st.tabs(["📂 Upload CSV","⌨️ Enter Manually"])
        with tab_upload:
            uploaded = st.file_uploader("Proposed Model CSV (required)", type=["csv"], label_visibility="collapsed")
            if uploaded:
                df = pd.read_csv(uploaded)
                st.session_state.csv_df  = df
                st.session_state.headers = list(df.columns)
                st.success(f"✅ {uploaded.name} — {len(df)} rows, {len(df.columns)} columns")
                st.dataframe(df.head(3), use_container_width=True)

                # ── Optional: Reference model + NECB savings check ──
                st.markdown("---")
                st.markdown("**Optional — NECB savings check**")
                st.caption("Add a Reference (baseline) model to compare end-use savings against NECB targets. "
                           "Leave blank to run the tool exactly as before.")
                necb_codes = load_necb_savings().get("codes") or ["NECB 2020", "NECB 2017"]
                prev_code = st.session_state.get("compliance_code")
                st.session_state.compliance_code = st.selectbox(
                    "Compliance Code", necb_codes,
                    index=necb_codes.index(prev_code) if prev_code in necb_codes else 0)
                ref_up = st.file_uploader("Reference Model CSV (optional)", type=["csv"], key="ref_uploader")
                if ref_up:
                    rdf = pd.read_csv(ref_up)
                    st.session_state.ref_csv_df = rdf
                    st.success(f"✅ Reference: {ref_up.name} — {len(rdf)} rows, {len(rdf.columns)} columns")
                elif st.session_state.get("ref_csv_df") is not None:
                    st.caption("✅ Reference model already loaded — re-upload above to replace it.")

                if st.button("Next: Map Columns →", type="primary"):
                    st.session_state.step = 2; st.rerun()
        with tab_manual:
            c1,c2,c3 = st.columns(3)
            manual = {}
            with c1:
                st.markdown("**Energy Sources**")
                manual["electricity_kwh"]   = st.number_input("Electricity (kWh/yr)",        min_value=0.0, step=1000.0, format="%.0f")
                manual["gas_kwh"]           = st.number_input("Natural Gas (kWh/yr)",        min_value=0.0, step=1000.0, format="%.0f")
                manual["other_fuel_kwh"]    = st.number_input("Other Fuel / Biomass (kWh/yr)", min_value=0.0, step=1000.0, format="%.0f",
                                                              help="Biomass, district energy, oil, propane, etc. Counted in total energy and GHGI.")
                manual["area_m2"]           = st.number_input("Floor Area (m²)",              min_value=0.0, step=100.0,  format="%.0f")
                manual["tedi"]              = st.number_input("TEDI (kWh/m²·yr)",             min_value=0.0, step=10.0,   format="%.1f",
                                                              help="Thermal Energy Demand Intensity — enter as an intensity (already per m²).")
            with c2:
                st.markdown("**HVAC End Uses**")
                manual["heating_kwh"]       = st.number_input("Space Heating (kWh/yr)",      min_value=0.0, step=1000.0, format="%.0f")
                manual["cooling_kwh"]       = st.number_input("Space Cooling (kWh/yr)",      min_value=0.0, step=1000.0, format="%.0f")
                manual["central_fan_kwh"]  = st.number_input("Interior Central Fan / AHU (kWh/yr)", min_value=0.0, step=1000.0, format="%.0f")
                manual["local_fan_kwh"]    = st.number_input("Interior Local Fan (kWh/yr)",        min_value=0.0, step=1000.0, format="%.0f")
                manual["exhaust_fan_kwh"]  = st.number_input("Exhaust Fan (kWh/yr)",               min_value=0.0, step=1000.0, format="%.0f")
                manual["pumps_kwh"]         = st.number_input("Pumps (kWh/yr)",              min_value=0.0, step=500.0,  format="%.0f")
                manual["heat_rejection_kwh"]= st.number_input("Heat Rejection (kWh/yr)",     min_value=0.0, step=500.0,  format="%.0f")
            with c3:
                st.markdown("**Other End Uses**")
                manual["lighting_kwh"]      = st.number_input("Interior Lighting (kWh/yr)",  min_value=0.0, step=1000.0, format="%.0f")
                manual["dhw_kwh"]           = st.number_input("DHW (kWh/yr)",                min_value=0.0, step=500.0,  format="%.0f")
                manual["receptacle_kwh"]    = st.number_input("Receptacle / Plug Loads (kWh/yr)", min_value=0.0, step=500.0, format="%.0f")
                manual["ext_lighting_kwh"]  = st.number_input("Exterior Lighting (kWh/yr)",  min_value=0.0, step=500.0,  format="%.0f")
                manual["process_kwh"]       = st.number_input("Process / Other (kWh/yr)",    min_value=0.0, step=500.0,  format="%.0f")
                st.markdown("**Unmet Hours**")
                manual["unmet_hours_heating"] = st.number_input("Unmet Hours — Heating", min_value=0.0, step=1.0, format="%.0f",
                                                                help="Count of occupied hours outside the heating setpoint range (not divided by area).")
                manual["unmet_hours_cooling"] = st.number_input("Unmet Hours — Cooling", min_value=0.0, step=1.0, format="%.0f")
                manual["unmet_hours_total"]   = st.number_input("Unmet Hours — Total",   min_value=0.0, step=1.0, format="%.0f")
            if st.button("Next: Building Info →", type="primary"):
                st.session_state.vals = manual; st.session_state.ref_vals = None; st.session_state.step = 3; st.rerun()

    elif st.session_state.step == 2:
        st.markdown("## Step 2 — Map Columns")
        headers = st.session_state.headers
        auto_map = guess_mapping(headers)
        options  = ["— not mapped —"] + headers
        mapping  = {}
        auto_matched = sum(1 for v in auto_map.values() if v)
        st.info(f"Auto-matched {auto_matched} of {len(FIELD_LABELS)} fields. Review and fix any that say '— not mapped —'.")

        # Group fields into sections for clarity
        sections = {
            "⚡ Energy Sources": ["electricity_kwh","gas_kwh","other_fuel_kwh","area_m2","tedi"],
            "🔥 HVAC End Uses":  ["heating_kwh","cooling_kwh","central_fan_kwh","local_fan_kwh","exhaust_fan_kwh","pumps_kwh","heat_rejection_kwh"],
            "💡 Other End Uses": ["lighting_kwh","dhw_kwh","receptacle_kwh","ext_lighting_kwh","process_kwh"],
            "⏱️ Unmet Hours":    ["unmet_hours_heating","unmet_hours_cooling","unmet_hours_total"],
        }
        for section_title, keys in sections.items():
            st.markdown(f"**{section_title}**")
            cols = st.columns(3)
            for i, key in enumerate(keys):
                label = FIELD_LABELS.get(key, key)
                with cols[i % 3]:
                    default_idx = options.index(auto_map[key]) if auto_map.get(key) and auto_map[key] in options else 0
                    sel = st.selectbox(label, options, index=default_idx, key=f"map_{key}")
                    mapping[key] = sel if sel != "— not mapped —" else None
            st.markdown("")
        cb,cn = st.columns([1,4])
        with cb:
            if st.button("← Back"): st.session_state.step=1; st.rerun()
        with cn:
            if st.button("Next: Building Info →", type="primary"):
                df = st.session_state.csv_df
                vals = {key: pd.to_numeric(df[col], errors="coerce").sum() if col else 0 for key,col in mapping.items()}
                st.session_state.vals=vals; st.session_state.mapping=mapping
                # Reference model (optional) — same export format, so reuse the same mapping
                rdf = st.session_state.get("ref_csv_df")
                if rdf is not None:
                    st.session_state.ref_vals = {
                        key: (pd.to_numeric(rdf[col], errors="coerce").sum() if (col and col in rdf.columns) else 0)
                        for key, col in mapping.items()
                    }
                else:
                    st.session_state.ref_vals = None
                st.session_state.step=3; st.rerun()

    elif st.session_state.step == 3:
        st.markdown("## Step 3 — Building Information")
        # ── Required project metadata (visible & mandatory) ──
        m1, m2, m3 = st.columns(3)
        with m1:
            project_name = st.text_input("Project Name", placeholder="e.g. New School A")
        with m2:
            software     = st.selectbox("Simulation Software *", SOFTWARE_OPTIONS)
        with m3:
            model_type   = st.selectbox("Model Type", MODEL_TYPES)
        phase = ""   # Project Phase removed; kept blank for downstream compatibility

        # ── Benchmark selection — identical filtering to the Benchmark Explorer ──
        st.markdown("**Benchmark to compare against** — same filters as the Benchmark Explorer. "
                    "Pick a city to compare against that exact benchmark model, or leave City on the zone average.")
        b1, b2, b3, b4 = st.columns(4)
        with b1:
            building_type = st.selectbox("Building Type", ALL_BUILDING_TYPES)
        provs = sorted({v["province"] for k,v in BENCHMARKS.items() if k[0]==building_type and v.get("province")})
        with b2:
            province = st.selectbox("Province", provs if provs else ["—"])
        zones_avail = sorted({k[2] for k,v in BENCHMARKS.items()
                              if k[0]==building_type and v.get("province")==province and k[2]})
        with b3:
            climate_zone = st.selectbox("Climate Zone", zones_avail if zones_avail else ["—"])
        cities_in_zone = sorted({k[1] for k,v in BENCHMARKS.items()
                                 if k[0]==building_type and v.get("province")==province and k[2]==climate_zone})
        zone_avg_label = f"Climate Zone {climate_zone} average"
        with b4:
            city_sel = st.selectbox("City (optional)", [zone_avg_label] + cities_in_zone,
                                    help="Leave on the zone average, or pick a city to compare against that exact benchmark model.")

        # Resolve the benchmark baseline used by all comparison tables, charts and QA checks
        if city_sel == zone_avg_label:
            zone_bms = [v for k,v in BENCHMARKS.items()
                        if k[0]==building_type and v.get("province")==province and k[2]==climate_zone]
            bm = average_benchmarks(zone_bms) if zone_bms else None
            city, subtype = "", "All"
            if bm:
                st.info(f"Comparing against the **Climate Zone {climate_zone} average** for {building_type} in {province} "
                        f"— average of {bm['_n']} benchmark(s). Select a city to use a specific benchmark model.")
            else:
                st.warning("No benchmark found for this combination. KPIs will be calculated but no comparison will be shown.")
        else:
            city = city_sel
            available_subtypes = sorted({k[3] for k in BENCHMARKS
                                         if k[0]==building_type and k[1]==city and k[2]==climate_zone})
            if len(available_subtypes) > 1:
                subtype = st.selectbox("Benchmark Model (subtype)", available_subtypes,
                                       help="Pick the exact benchmark model to compare against.")
            elif len(available_subtypes) == 1:
                subtype = available_subtypes[0]
            else:
                subtype = "General"
            bm = BENCHMARKS.get((building_type, city, climate_zone, subtype))
            if bm:
                lbl = f"{building_type} · {city} · Climate Zone {climate_zone}" + (f" · {subtype}" if subtype not in ("General","All") else "")
                st.info(f"Comparing against benchmark model: **{lbl}**.")
            else:
                st.warning("No benchmark found for this exact combination. KPIs will be calculated but no comparison will be shown.")

        # ── GHG emission factors (defaults follow the selected province) ──
        elec_opts  = {p: round(g/1000, 4)              for p, g in ELEC_CO2E_G_PER_KWH.items()}
        gas_opts   = {p: round(g/NG_KWH_PER_M3/1000,4) for p, g in NG_CO2_G_PER_M3.items()}
        other_opts = {"None": 0.0, "Propane": 0.2197, "Butane": 0.2229}
        CUSTOM = "Custom value…"
        elec_default = round(ELEC_CO2E_G_PER_KWH.get(province, 0)/1000, 4)
        gas_default  = round(NG_CO2_G_PER_M3.get(province, 0)/NG_KWH_PER_M3/1000, 4) if NG_CO2_G_PER_M3.get(province) else 0.0

        st.markdown("**GHG Emission Factors (kg CO₂e/kWh)** — select a province/fuel value (Canada 2026 tables) or choose *Custom value…* to enter your own. Choose 0 / None to skip GHGI.")
        ef1, ef2, ef3 = st.columns(3)
        with ef1:
            opts = [f"{p} — {v:.4f}" for p, v in elec_opts.items()] + [CUSTOM]
            di = list(elec_opts).index(province) if province in elec_opts else 0
            sel = st.selectbox("Electricity", opts, index=di, key=f"ef_elec_sel_{province}",
                               help="Provincial electricity consumption intensity (Table 5.3, 2026).")
            ef_elec = (st.number_input("Electricity — custom", min_value=0.0, step=0.001, format="%.4f",
                                       value=elec_default, key=f"ef_elec_cust_{province}")
                       if sel == CUSTOM else elec_opts[sel.rsplit(" — ", 1)[0]])
        with ef2:
            opts = [f"{p} — {v:.4f}" for p, v in gas_opts.items()] + [CUSTOM]
            di = list(gas_opts).index(province) if province in gas_opts else 0
            sel = st.selectbox("Natural Gas", opts, index=di, key=f"ef_gas_sel_{province}",
                               help="Marketable natural-gas CO₂ (Table 1.3, 2026, g CO₂/m³) ÷ 10.55 kWh/m³.")
            ef_gas = (st.number_input("Natural Gas — custom", min_value=0.0, step=0.001, format="%.4f",
                                      value=gas_default, key=f"ef_gas_cust_{province}")
                      if sel == CUSTOM else gas_opts[sel.rsplit(" — ", 1)[0]])
        with ef3:
            opts = [f"{p} — {v:.4f}" for p, v in other_opts.items()] + [CUSTOM]
            sel = st.selectbox("Other Resources / Biomass", opts, index=0, key=f"ef_other_sel_{province}",
                               help="Propane/Butane converted from Table 3.3; or pick Custom for biomass/other.")
            ef_other = (st.number_input("Other — custom", min_value=0.0, step=0.001, format="%.4f",
                                        value=0.0, key=f"ef_other_cust_{province}")
                        if sel == CUSTOM else other_opts[sel.rsplit(" — ", 1)[0]])

        cb,cn = st.columns([1,4])
        with cb:
            if st.button("← Back"):
                st.session_state.step = 2 if "headers" in st.session_state else 1; st.rerun()
        with cn:
            if st.button("Calculate KPIs & View Results →", type="primary"):
                if not software:
                    st.error("⚠️ Simulation Software is required.")
                else:
                    kpis   = calculate_kpis(st.session_state.vals, ef_elec=ef_elec, ef_gas=ef_gas, ef_other=ef_other)
                    flags  = generate_flags(kpis, bm, building_type)
                    pct    = calc_percentile(kpis["total_eui"], bm["pct_data"]) if bm else None
                    st.session_state.results = {"kpis":kpis,"bm":bm,"flags":flags,"percentile":pct,
                        "meta":{"project_name":project_name,"building_type":building_type,"city":city,
                                "province":province,"climate_zone":climate_zone,"subtype":subtype,"software":software,
                                "model_type":model_type,"phase":phase,"date":str(date.today()),
                                "ef_elec":ef_elec,"ef_gas":ef_gas,"ef_other":ef_other}}
                    st.session_state.step=4; st.rerun()

    elif st.session_state.step == 4 and st.session_state.results:
        r=st.session_state.results; kpis=r["kpis"]; bm=r["bm"]; meta=r["meta"]; flags=r["flags"]; pct=r["percentile"]

        # Build the per-metric comparison rows once, and derive the QA/QC Flags from the
        # same Auto Flag statuses so the flag list and the comparison table always agree.
        comp_rows = build_comparison_rows(kpis, bm) if bm else None
        flags = flags_from_comparison(comp_rows, kpis, bm, meta["building_type"]) if bm else flags

        # ── Project header banner ──
        fc={"pass":0,"warn":0,"fail":0}
        for f in flags: fc[f[0]]=fc.get(f[0],0)+1
        overall       = "Issues Found"    if fc["fail"]>0 else "Review Required" if fc["warn"]>0 else "All Clear"
        overall_icon  = "🔴"              if fc["fail"]>0 else "🟡"              if fc["warn"]>0 else "🟢"
        overall_color = "#fef2f2"         if fc["fail"]>0 else "#fffbeb"         if fc["warn"]>0 else "#f0fdf4"
        overall_border= "#dc2626"         if fc["fail"]>0 else "#d97706"         if fc["warn"]>0 else "#16a34a"
        proj_name     = meta["project_name"] or "Project Results"
        pct_text      = f"&nbsp;&nbsp;|&nbsp;&nbsp; Benchmark percentile: <b>{pct}th</b> (lower = better)" if pct else ""
        flag_text     = f"Pass: {fc['pass']}  &nbsp; Review: {fc['warn']}  &nbsp; Fail: {fc['fail']}"
        subtype_disp  = f" &middot; {meta.get('subtype','')}" if meta.get('subtype','') not in ("","General","All") else ""
        phase_disp    = f" &middot; {meta['phase']}" if meta.get('phase') else ""
        city_disp     = f"{meta['city']} &middot; " if meta.get('city') else ""
        scope_disp    = "" if meta.get('city') else " &middot; zone average"
        info_line     = (f"{meta['building_type']} &middot; {city_disp}"
                         f"Climate Zone {meta['climate_zone']}{scope_disp}{subtype_disp} &middot; {meta['model_type']}"
                         f"{phase_disp} &middot; {meta['software']}")
        banner_html = (
            f'<div style="background:{overall_color};border-left:5px solid {overall_border};'
            f'border-radius:10px;padding:18px 22px;margin-bottom:16px">'
            f'<div style="font-size:22px;font-weight:700;color:#1e293b">{proj_name}</div>'
            f'<div style="font-size:13px;color:#64748b;margin-top:4px">{info_line}</div>'
            f'<div style="margin-top:12px;display:flex;gap:8px;align-items:center;flex-wrap:wrap">'
            f'<span style="font-size:14px;font-weight:700;color:{overall_border}">{overall_icon} {overall}</span>'
            f'<span style="font-size:13px;color:#64748b">{pct_text}</span>'
            f'</div>'
            f'<div style="margin-top:6px;font-size:13px;color:#94a3b8">{flag_text}</div>'
            f'</div>'
        )
        st.markdown(banner_html, unsafe_allow_html=True)

        # ── Section 1: Energy Summary ──
        st.markdown("### 📊 Energy Summary")
        st.caption("How much energy does this building use per square metre per year?")

        # Row 1 — main totals
        r1c1,r1c2,r1c3,r1c4,r1c5 = st.columns(5)
        with r1c1:
            delta_eui = f"Benchmark median: {bm['median_eui']} kWh/m²·yr (±15% = {bm['good_eui']}–{bm['high_eui']})" if bm else "No benchmark"
            color_eui = "#f0fdf4" if bm and kpis["total_eui"]<=bm["good_eui"] else "#fef2f2" if bm and kpis["total_eui"]>bm["high_eui"] else "#fffbeb"
            st.markdown(f'''<div style="background:{color_eui};border-radius:10px;padding:14px 16px;border:1px solid #e2e8f0">
                <div style="font-size:11px;font-weight:600;color:#64748b;text-transform:uppercase;letter-spacing:.05em">Total EUI</div>
                <div style="font-size:28px;font-weight:700;color:#0f4c81;line-height:1.1">{kpis["total_eui"]}</div>
                <div style="font-size:12px;color:#64748b">kWh/m²·yr</div>
                <div style="font-size:12px;color:#94a3b8;margin-top:4px">{delta_eui}</div>
            </div>''', unsafe_allow_html=True)
        with r1c2:
            has_tedi_bm = bool(bm and bm.get("median_tedi"))
            tedi_val = kpis.get("tedi", 0)
            if has_tedi_bm and tedi_val > 0:
                color_tedi = "#f0fdf4" if tedi_val<=bm["good_tedi"] else "#fef2f2" if tedi_val>bm["high_tedi"] else "#fffbeb"
                tedi_sub = f"Benchmark median: {bm['median_tedi']} kWh/m²·yr (±15% = {bm['good_tedi']}–{bm['high_tedi']})"
            else:
                color_tedi = "#f8fafc"
                tedi_sub = "No TEDI benchmark" if not has_tedi_bm else "TEDI not provided"
            st.markdown(f'''<div style="background:{color_tedi};border-radius:10px;padding:14px 16px;border:1px solid #e2e8f0">
                <div style="font-size:11px;font-weight:600;color:#64748b;text-transform:uppercase;letter-spacing:.05em">TEDI</div>
                <div style="font-size:28px;font-weight:700;color:#0f4c81;line-height:1.1">{tedi_val if tedi_val>0 else "—"}</div>
                <div style="font-size:12px;color:#64748b">kWh/m²·yr</div>
                <div style="font-size:12px;color:#94a3b8;margin-top:4px">{tedi_sub}</div>
            </div>''', unsafe_allow_html=True)
        with r1c3:
            ghgi_val = kpis.get("ghgi")
            ghgi_sub = ("Benchmark median: " + str(bm["median_ghgi"]) + " kgCO₂e/m²·yr" if bm else "No benchmark") \
                       if ghgi_val is not None else "Enter emission factors in Step 3"
            st.markdown(f'''<div style="background:#f8fafc;border-radius:10px;padding:14px 16px;border:1px solid #e2e8f0">
                <div style="font-size:11px;font-weight:600;color:#64748b;text-transform:uppercase;letter-spacing:.05em">GHGI</div>
                <div style="font-size:28px;font-weight:700;color:#0f4c81;line-height:1.1">{ghgi_val if ghgi_val is not None else "—"}</div>
                <div style="font-size:12px;color:#64748b">kgCO₂e / m²·yr</div>
                <div style="font-size:12px;color:#94a3b8;margin-top:4px">{ghgi_sub}</div>
            </div>''', unsafe_allow_html=True)
        with r1c4:
            st.markdown(f'''<div style="background:#f8fafc;border-radius:10px;padding:14px 16px;border:1px solid #e2e8f0">
                <div style="font-size:11px;font-weight:600;color:#64748b;text-transform:uppercase;letter-spacing:.05em">Electricity EUI</div>
                <div style="font-size:28px;font-weight:700;color:#3b82f6;line-height:1.1">{kpis["elec_eui"]}</div>
                <div style="font-size:12px;color:#64748b">kWh/m²·yr</div>
                <div style="font-size:12px;color:#94a3b8;margin-top:4px">Floor area: {round(kpis["area"]):,} m²</div>
            </div>''', unsafe_allow_html=True)
        with r1c5:
            st.markdown(f'''<div style="background:#f8fafc;border-radius:10px;padding:14px 16px;border:1px solid #e2e8f0">
                <div style="font-size:11px;font-weight:600;color:#64748b;text-transform:uppercase;letter-spacing:.05em">Gas EUI</div>
                <div style="font-size:28px;font-weight:700;color:#ef4444;line-height:1.1">{kpis["gas_eui"]}</div>
                <div style="font-size:12px;color:#64748b">kWh/m²·yr</div>
                <div style="font-size:12px;color:#94a3b8;margin-top:4px">Total energy: {round(kpis["total_energy"]/1000):,} MWh/yr</div>
            </div>''', unsafe_allow_html=True)

        st.markdown("")

        # Row 2 — featured cards: Other Fuel / Biomass + Unmet Hours.
        # (The full per-end-use breakdown now lives below the benchmark comparison table.)
        r2c1, r2c2, _r2c3, _r2c4, _r2c5 = st.columns(5)
        with r2c1:
            st.markdown(f'''<div style="background:#f8fafc;border-radius:10px;padding:12px 14px;border:1px solid #e2e8f0;border-top:3px solid #84cc16;margin-bottom:8px">
                <div style="font-size:11px;font-weight:600;color:#64748b;text-transform:uppercase;letter-spacing:.05em">Other Fuel / Biomass EUI</div>
                <div style="font-size:22px;font-weight:700;color:#84cc16;line-height:1.1">{kpis.get("other_fuel_eui",0)}</div>
                <div style="font-size:11px;color:#94a3b8;margin-top:2px">kWh/m²·yr</div>
            </div>''', unsafe_allow_html=True)
        with r2c2:
            uh_sub = f'Heating: {kpis.get("unmet_heating",0):,} · Cooling: {kpis.get("unmet_cooling",0):,}'
            st.markdown(f'''<div style="background:#f8fafc;border-radius:10px;padding:12px 14px;border:1px solid #e2e8f0;border-top:3px solid #e11d48;margin-bottom:8px">
                <div style="font-size:11px;font-weight:600;color:#64748b;text-transform:uppercase;letter-spacing:.05em">Unmet Hours</div>
                <div style="font-size:22px;font-weight:700;color:#e11d48;line-height:1.1">{kpis.get("unmet_total",0):,}</div>
                <div style="font-size:11px;color:#94a3b8;margin-top:2px">hours / yr</div>
                <div style="font-size:11px;color:#94a3b8">{uh_sub}</div>
            </div>''', unsafe_allow_html=True)

        st.divider()

        # ── Section 2: Charts ──
        st.markdown("### 📈 Charts")
        cc1,cc2 = st.columns(2)
        with cc1:
            eu_l=["Heating","Cooling","Fans","Lighting","DHW","Receptacle","Pumps"]
            eu_v=[kpis["heat_eui"],kpis["cool_eui"],kpis["fan_eui"],kpis["ltg_eui"],kpis["dhw_eui"],kpis.get("recept_eui",0),kpis["pumps_eui"]]
            pl=[l for l,v in zip(eu_l,eu_v) if v>0]; pv=[v for v in eu_v if v>0]
            st.plotly_chart(make_pie(pl,pv,"End-Use Energy Split"), use_container_width=True)
        with cc2:
            if bm:
                sorted_pcts=sorted(bm["pct_data"])
                bar_colors=["#16a34a" if v<=bm["good_eui"] else "#d97706" if v<=bm["median_eui"] else "#ea580c" if v<=bm["high_eui"] else "#dc2626" for v in sorted_pcts]
                fig_pct=go.Figure(go.Bar(x=[f"{i*10}th" for i in range(1,11)],y=sorted_pcts,marker_color=bar_colors,
                    hovertemplate="<b>%{x} percentile</b><br>EUI: %{y} kWh/m²·yr<extra></extra>"))
                fig_pct.add_hline(y=kpis["total_eui"],line_dash="dash",line_color="#0f4c81",line_width=2.5,
                    annotation_text=f"  Your model: {kpis['total_eui']}",annotation_font=dict(color="#0f4c81",size=12))
                fig_pct.update_layout(template="plotly_white",height=320,
                    title=dict(text="Where does your building rank? (lower = better)",font=dict(size=13,color="#0f4c81")),
                    xaxis_title="Percentile rank of similar buildings",yaxis_title="EUI (kWh/m²·yr)",margin=dict(t=45,b=10))
                st.plotly_chart(fig_pct,use_container_width=True)

        if bm:
            eu_l2=["Heating","Cooling","Central Fan","Local Fan","Exhaust Fan","Lighting","DHW","Receptacle","Pumps"]
            your_eu=[kpis["heat_eui"],kpis["cool_eui"],kpis.get("central_fan_eui",0),kpis.get("local_fan_eui",0),
                     kpis.get("exhaust_fan_eui",0),kpis["ltg_eui"],kpis["dhw_eui"],kpis.get("recept_eui",0),kpis["pumps_eui"]]
            fan_split = bm["fan_pct"] / 3  # split benchmark fan % evenly across central/local/exhaust
            bm_good=[round(bm["good_eui"]*p/100,1)  for p in [bm["heat_pct"],bm["cool_pct"],fan_split,fan_split,fan_split,bm["ltg_pct"],bm["dhw_pct"],bm["recept_pct"],bm["pumps_pct"]]]
            bm_med =[round(bm["median_eui"]*p/100,1) for p in [bm["heat_pct"],bm["cool_pct"],fan_split,fan_split,fan_split,bm["ltg_pct"],bm["dhw_pct"],bm["recept_pct"],bm["pumps_pct"]]]
            fig_cmp = go.Figure()
            # Benchmark median bars with ±15% error bars
            fig_cmp.add_trace(go.Bar(
                name="Benchmark Median", x=eu_l2, y=bm_med,
                marker_color="#94a3b8", opacity=0.6,
                error_y=dict(type="data", symmetric=True,
                             array=[round(v*0.15,1) for v in bm_med],
                             color="#64748b", thickness=1.5, width=5),
            ))
            # Your model bars, labelled with % difference vs benchmark median
            pct_labels = [f"{(y-m)/m*100:+.0f}%" if m else "" for y, m in zip(your_eu, bm_med)]
            fig_cmp.add_trace(go.Bar(
                name="Your Model", x=eu_l2, y=your_eu,
                marker_color="#0f4c81",
                text=pct_labels, textposition="outside", textfont=dict(size=11, color="#0f4c81"),
                cliponaxis=False,
            ))
            fig_cmp.update_layout(
                barmode="group", template="plotly_white", height=340,
                title=dict(text="Your Model vs Benchmark Median (labels = % vs median, error bars = ±15%)", font=dict(size=13,color="#0f4c81")),
                yaxis_title="EUI (kWh/m²·yr)",
                legend=dict(orientation="h", y=-0.25), margin=dict(t=50,b=10)
            )
            st.plotly_chart(fig_cmp, use_container_width=True)

        st.divider()

        # ── Section 3: Benchmark Table ──
        comparison_edited = None
        if bm:
            st.markdown("### 📋 Benchmark Comparison")
            st.caption("QA/QC thresholds: 🟢 ≤ −15% of median &nbsp; 🟡 Within ±15% of median &nbsp; 🔴 > +15% of median")

            rows = comp_rows   # built once at the top of Step 4 (same statuses drive the QA/QC Flags)

            # Editable comparison: reviewer can override the QA Status and add a Comment.
            # "Auto Flag" preserves the tool's original 4-level computed status for audit;
            # the editable QA Status is a binary pass/fail decision (🟢 / 🔴).
            comp_df = pd.DataFrame(rows).rename(columns={"QA Status": "Auto Flag"})
            comp_df["QA Status"] = comp_df["Auto Flag"].apply(lambda s: "🔴" if s == "🔴" else "🟢")  # default: red only for fails
            comp_df["Comment"]   = ""
            comp_df = comp_df[["End Use","Your Model","Benchmark Median","Difference","Auto Flag","QA Status","Comment"]]

            st.caption("Set the **QA Status** of any row to 🟢 Pass or 🔴 Fail (e.g. pass a flagged item after review) and add a **Comment**. The **Auto Flag** column keeps the tool's original result for audit, and your overrides + comments are saved into the exported PDF.")
            comparison_edited = st.data_editor(
                comp_df, use_container_width=True, hide_index=True, key="bm_review",
                column_config={
                    "End Use":          st.column_config.TextColumn(disabled=True),
                    "Your Model":       st.column_config.NumberColumn(disabled=True),
                    "Benchmark Median": st.column_config.NumberColumn(disabled=True),
                    "Difference":       st.column_config.TextColumn("Difference (%)", disabled=True,
                                            help="Percent difference of your model vs the benchmark median."),
                    "Auto Flag":        st.column_config.TextColumn(
                                            "Auto Flag", disabled=True,
                                            help="Status the tool calculated automatically — kept for audit."),
                    "QA Status":        st.column_config.SelectboxColumn(
                                            "QA Status", options=["🟢","🔴"], required=True,
                                            help="Reviewer decision — 🟢 Pass or 🔴 Fail."),
                    "Comment":          st.column_config.TextColumn(
                                            "Comment", help="Reviewer notes / justification for any override."),
                },
            )
            binary_default = comp_df["QA Status"]
            n_over = int((comparison_edited["QA Status"].values != binary_default.values).sum())
            if n_over:
                st.caption(f"✏️ {n_over} status value(s) overridden by reviewer.")

        st.divider()

        # ── NECB Savings QA (only when a Reference model was also uploaded) ──
        ref_vals = st.session_state.get("ref_vals")
        if ref_vals is not None:
            code = st.session_state.get("compliance_code", "NECB 2020")
            necb_all = load_necb_savings()
            sheet_savings = necb_savings_for(necb_all, meta.get("building_type"), meta.get("city"), code)
            from_sheet = bool(sheet_savings)
            # Use the sheet's targets exactly (blank cells → no target). Defaults only when no row matches.
            savings_map = sheet_savings if from_sheet else dict(DEFAULT_NECB_SAVINGS)
            necb_rows = build_necb_rows(st.session_state.vals, ref_vals, savings_map)

            st.markdown(f"### 🎯 Code Compliance Check — {code}")
            if from_sheet:
                src_note = f"targets loaded from the **{NECB_SHEET_NAME}** sheet for {meta.get('building_type','?')} · {meta.get('city','?')}"
            else:
                src_note = f"using built-in defaults (no matching row found on the **{NECB_SHEET_NAME}** tab)"
            st.caption(
                f"Proposed vs Reference savings per end use against NECB targets ({src_note}). "
                f"**Savings = (Reference − Proposed) / Reference × 100.** "
                f"Auto Flag: 🟢 savings within {NECB_TOL} pts of (or above) target · 🔴 below target · ⚪ n/a (reference = 0). "
                f"Override **QA Status** and add a **Comment** as needed."
            )
            necb_df = pd.DataFrame(necb_rows)
            necb_df["QA Status"] = necb_df["Auto Flag"]
            necb_df["Comment"]   = ""
            necb_df = necb_df[["End Use","Proposed","Reference Model","Savings","Benchmark","Auto Flag","QA Status","Comment"]]
            necb_edited = st.data_editor(
                necb_df, use_container_width=True, hide_index=True, key="necb_review",
                column_config={
                    "End Use":         st.column_config.TextColumn(disabled=True),
                    "Proposed":        st.column_config.NumberColumn("Proposed (kWh)",        disabled=True, format="%.1f"),
                    "Reference Model": st.column_config.NumberColumn("Reference Model (kWh)", disabled=True, format="%.1f"),
                    "Savings":         st.column_config.TextColumn(disabled=True),
                    "Benchmark":       st.column_config.TextColumn(disabled=True, help="NECB target savings for this end use."),
                    "Auto Flag":       st.column_config.TextColumn(disabled=True, help="Computed automatically; kept for audit."),
                    "QA Status":       st.column_config.SelectboxColumn(options=["🟢","🔴","⚪"], required=True, help="Reviewer override (🟢 pass / 🔴 fail / ⚪ n/a)."),
                    "Comment":         st.column_config.TextColumn(help="Justification or notes for any override."),
                },
            )
            n_g  = int((necb_edited["QA Status"]=="🟢").sum())
            n_r  = int((necb_edited["QA Status"]=="🔴").sum())
            n_na = int((necb_edited["QA Status"]=="⚪").sum())
            n_ov = int((necb_edited["QA Status"].values != necb_df["QA Status"].values).sum())
            st.caption(f"🟢 {n_g} meet target · 🔴 {n_r} below target · ⚪ {n_na} n/a"
                       + (f"  ·  ✏️ {n_ov} overridden by reviewer" if n_ov else ""))
            st.divider()

        # ── Section 4: QA/QC Flags ──
        st.markdown("### 🚩 QA/QC Flags")
        st.caption("Each flag mirrors the Auto Flag column above (your model vs benchmark median).")
        for level,icon,msg in flags:
            st.markdown(f'<div class="flag-{level}"><b>{icon}</b> {msg}</div>',unsafe_allow_html=True)

        st.divider()

        # ── Export ──
        st.markdown("### 📥 Export Report")
        pdf_bytes = build_pdf_report(meta, kpis, bm, flags, pct, comparison_edited)

        cd1,cd2=st.columns(2)
        with cd1:
            st.download_button("📄 Download PDF Report", data=pdf_bytes,
                file_name=f"QA_QC_{(meta['project_name'] or 'report').replace(' ','_')}_{meta['date']}.pdf",
                mime="application/pdf", use_container_width=True)
        with cd2:
            if st.button("← Run Another Project",use_container_width=True):
                for k in ["step","vals","results","headers","csv_df","mapping","ref_csv_df","ref_vals","compliance_code"]:
                    if k in st.session_state: del st.session_state[k]
                st.session_state.step=1; st.rerun()

        st.divider()

        # ── Add to Benchmark ──
        st.markdown("### 🏛️ Add This Model to Benchmark Database")
        # This section reads from the QA Status column. With every row overridden to 🟢
        # there are no fails and no reviews, so neither the error nor the warning shows.
        if comparison_edited is not None:
            fail_count = int((comparison_edited["QA Status"] == "🔴").sum())
            warn_count = int(comparison_edited["QA Status"].isin(["🟡", "🟠"]).sum())
            pass_count = int((comparison_edited["QA Status"] == "🟢").sum())
        else:
            fail_count = sum(1 for f in flags if f[0]=="fail")
            warn_count = sum(1 for f in flags if f[0]=="warn")
            pass_count = sum(1 for f in flags if f[0]=="pass")

        if fail_count > 0:
            st.error(f"❌ This model has **{fail_count} unresolved QA/QC fail(s)** (🔴) in the comparison table above. Override them to 🟢 after review — or fix the model — before adding to the benchmark database.")
        else:
            if warn_count > 0:
                st.warning(f"⚠️ This model has **{warn_count} review flag(s)**. You can still add it to the database, but review the flags first.")

            with st.expander("ℹ️ What does adding to the benchmark do?", expanded=False):
                st.markdown("""
When you add this model to the benchmark database, it contributes to the pool of real projects used for future comparisons.

**What gets recorded:**
- Project name, building type, city, climate zone, subtype
- Total EUI, all end-use EUIs, GHGI
- Floor area, model type, phase, date, software

**How it updates the benchmark:**
- Your project's EUI is added to the percentile dataset for that building type/city/subtype
- The median EUI and end-use percentages in the Benchmarks sheet are **not changed automatically** — a senior reviewer can periodically recalculate and update those based on accumulated project data
- All submitted projects are visible in the **Project Results** tab of the Google Sheet for full audit trail
                """)

            with st.form("add_to_benchmark_form"):
                st.markdown("**Confirm submission details:**")
                fc1, fc2, fc3 = st.columns(3)
                with fc1:
                    confirm_name    = st.text_input("Project Name",    value=meta["project_name"])
                with fc2:
                    confirm_subtype = st.text_input("Benchmark Subtype", value=meta.get("subtype","General"),
                                                    help="e.g. Boiler + VAV · Medium, Heat Pump + DOAS")
                    confirm_notes   = st.text_area("Notes (optional)", placeholder="e.g. NECB 2020 proposed model, 100% design stage", height=80)
                with fc3:
                    st.markdown("**Summary of values being submitted:**")
                    st.markdown(f"""
| Field | Value |
|---|---|
| Building Type | {meta['building_type']} |
| City | {meta['city']} · Zone {meta['climate_zone']} |
| Total EUI | **{kpis['total_eui']} kWh/m²·yr** |
| GHGI | {kpis['ghgi'] if kpis.get('ghgi') is not None else '—'} kgCO₂e/m²·yr |
| Floor Area | {round(kpis['area']):,} m² |
| QA/QC | ✅ {pass_count} pass · ⚠️ {warn_count} review |
                    """)

                submitted = st.form_submit_button("✅ Add to Benchmark Database", type="primary", use_container_width=True)

                if submitted:
                    if not confirm_name.strip():
                        st.error("Please enter a project name.")
                    else:
                        # Build the row to append to "Project Results" sheet
                        project_row = [
                            confirm_name.strip(),
                            meta["building_type"],
                            meta["city"],
                            meta["climate_zone"],
                            confirm_subtype.strip(),
                            meta["software"],
                            meta["model_type"],
                            meta["phase"],
                            meta["date"],
                            round(kpis["area"]),
                            kpis["total_eui"],
                            kpis.get("tedi", 0),
                            kpis["elec_eui"],
                            kpis["gas_eui"],
                            kpis["heat_eui"],
                            kpis["cool_eui"],
                            kpis["fan_eui"],
                            kpis["ltg_eui"],
                            kpis["dhw_eui"],
                            kpis.get("recept_eui", 0),
                            kpis["pumps_eui"],
                            kpis.get("ghgi") if kpis.get("ghgi") is not None else "",
                            f"{pct}th" if pct else "N/A",
                            f"{pass_count} pass · {warn_count} review · {fail_count} fail",
                            confirm_notes.strip(),
                        ]

                        ok, msg = append_to_google_sheet("Project Results", project_row)
                        if ok:
                            st.success(f"✅ **{confirm_name}** has been added to the benchmark database. Thank you!")
                            st.info("💡 A senior reviewer can periodically update the Benchmarks sheet median EUI and percentile data using the accumulated project results.")
                            load_benchmarks.clear()
                        else:
                            # Fallback — show the data so user can manually add it
                            st.warning(f"⚠️ Could not write to Google Sheets automatically: {msg}")
                            st.markdown("**Please manually copy this row into the 'Project Results' sheet in Google Sheets:**")
                            headers_pr = ["Project Name","Building Type","City","Climate Zone","Subtype",
                                          "Software","Model Type","Phase","Date","Area (m²)","Total EUI","TEDI","Elec EUI",
                                          "Gas EUI","Heating EUI","Cooling EUI","Fan EUI","Lighting EUI","DHW EUI",
                                          "Receptacle EUI","Pumps EUI","GHGI","Percentile","QA Flags","Notes"]
                            st.dataframe(pd.DataFrame([dict(zip(headers_pr, project_row))]),
                                        use_container_width=True, hide_index=True)
