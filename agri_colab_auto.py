# ============================================================================
# AGRI INTELLIGENCE WORKBOOK - ALL-INDIA TEHSILS, AUTO-FETCHED  (Google Colab)
# ----------------------------------------------------------------------------
# No upload needed. Paste each "CELL n" block into its own Colab cell and run
# them in order. Cells marked OPTIONAL can be skipped.
#
# SOURCE PRIORITY (every field tries these in order; later tiers only fill what earlier tiers left blank):
#   TIER 1  Government of India records / assessments  (LGD, IMD, CGWB, MI Census, DES APY, Agmarknet*)
#   TIER 2  Government-linked / public-interest models (CoRE Stack, agro-climatic zones)
#   TIER 3  Global agency / third-party models = FALLBACK (Open-Meteo, Copernicus DEM, MODIS/SMAP/SoilGrids)
#   TIER 4  News signals (PIB/Google RSS) - context only, never fills numbers          (*needs free key)
#
#   CELL 1  setup
#   CELL 2  schema + settings
#   CELL 3  [T1] tehsil list + boundaries (LGD, direct download from bharatlas.com)
#   CELL 4  [T1/T2] agro-climatic zone, CGWB groundwater, MI-census wells, CoRE Stack cropping
#   CELL 5  OPTIONAL  merge data you already have
#   CELL 6  [T3] elevation (Copernicus DEM via Open-Meteo)
#   CELL 7  [T1 -> T3 fallback] IMD gridded weather; Open-Meteo only where IMD has no value
#   CELL 8  [T1] OPTIONAL  Agmarknet mandi prices (free data.gov.in key, else skipped)
#   CELL 9  [T1] DES APY district area / production / yield -> YIELD_PRODUCTION + 5_YEAR_HISTORY
#   CELL 10 [T4] PIB + Google News RSS (context only, labelled unverified)
#   CELL 11 [T3] satellite + soil + cropland (Google Earth Engine; service-account key auto-detected)
#   CELL 12 [derived] decision layer: AGRI_INTELLIGENCE, ALERTS, TRACTOR_OPPORTUNITY (rule-based, transparent)
#   CELL 12B [validation] RECORD_METADATA, DATA_VALIDATION (govt vs satellite), DATA_RELIABILITY score,
#            DATA_GAP_REGISTER, AI_ESTIMATION_REGISTER (Gemini -> GPT -> Claude, else regex; never fills main tabs)
#   CELL 13 build the workbook (+ DATA_LINEAGE, SOURCE_RECOVERY_LOG; Source + Source Date on every tab)
#   CELL 14 create the Google Sheet(s) in your Drive (link printed)
#   CELL 15 trend module - collect 5-year monthly series (IMD per year, CHIRPS fallback; Earth Engine)
#   CELL 16 trend module - build tabs A-G + long file + Google Sheet
#   CELL 17 final summary: all Google Sheet links, files, reliability, gaps
#
# EVIDENCE PILLARS: (1) Government records  (2) Google Earth Engine / satellite datasets as independent checks.
# District figures are never relabelled as tehsil facts: RECORD_METADATA shows Source Level vs Output Level.
#
# Honest limits: ISRO VEDAS / Bhuvan / MOSDAC, Soil Health Card, India-WRIS and
# CGWB have no open anonymous API, so they are NOT pulled. Anything without a
# real source stays blank - nothing is invented.
# ============================================================================


# ===== CELL 1 : SETUP =======================================================
import subprocess, sys
subprocess.run([sys.executable, "-m", "pip", "install", "-q", "imdlib", "geopandas",
                "pyarrow", "shapely", "xarray", "netCDF4", "openpyxl", "xlrd"],
               check=False, capture_output=True)

import csv, json, math, os, re, time, zipfile
from concurrent.futures import ThreadPoolExecutor
from datetime import date, timedelta

import numpy as np
import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


# ===== CELL 2 : SCHEMA + SETTINGS ===========================================
OUTPUT_FILE = "Agri_Intelligence_India.xlsx"
TODAY = date.today()
FONT = "Arial"

# --- settings you may edit ---------------------------------------------------
BHARATLAS = "https://bharatlas.com/api/v1"   # open API, no key (LGD boundaries)
TEHSIL_SOURCE_URL = "https://bharatlas.com/api/dl/admin/subdistricts/LGD_Subdistricts.parquet"  # LGD sub-districts (6,471 polygons)
GEE_PROJECT = "agri-intelligence-515106"   # your Earth Engine Cloud project; used for the interactive-login fallback when no JSON key is found
GEE_KEY_FILE = ""           # leave blank: auto-finds focus-cargo-446109-q1-20a07590f5db*.json (or any service-account key) in /content
IMD_LAG_DAYS = 2            # IMD real-time grids arrive ~1-2 days late
NORMAL_YEARS = 10           # years of IMD history used for the rainfall normal
DRY_DAY_MM = 2.5            # a day with less rain than this counts as "dry"
HEAVY_RAIN_7D_MM = 100      # forecast 7-day rain above this -> "Heavy rain watch"
SIMPLIFY_DEG = 0.004        # polygon simplification for Earth Engine (~400 m)
FRESHNESS_DAYS = {"SOIL_INTELLIGENCE": 730, "WATER_IRRIGATION": 400,
                  "WEATHER_LIVE": 4, "MARKET_ECONOMICS": 4}
STALE_MULTIPLIER = 2        # STALE = older than max but within this x max
SAT_TARGET_PIXELS = 30      # satellite/soil Confidence = share of this many 250 m cropland pixels the tehsil has (1.0 = 30+)
MIN_NORMAL_MM = 50          # rainfall-deviation alerts only where the 30-day normal is at least this (avoids dry-season false alarms)
# Decision-layer weights (rule-based, editable). Missing components are skipped and the rest re-weighted.
SCORE_WEIGHTS = {"Crop Health": 0.30, "Water Score": 0.20, "Weather Score": 0.15, "Soil Score": 0.15, "Yield Outlook": 0.20, "Market Score": 0.15}
TRACTOR_WEIGHTS = {"Crop Area": 0.25, "Cropping Intensity": 0.15, "Crop Health": 0.15, "Crop Economy": 0.10,
                   "Farm Prosperity": 0.10, "Mechanisation": 0.05, "Historical Tractor Demand": 0.15, "Replacement Potential": 0.05}
CROP_DURATION = {"Paddy": 130, "Wheat": 125, "Maize": 105, "Groundnut": 110, "Cotton": 170, "Soyabean": 100, "Bajra": 80,
                 "Sugarcane": 365, "Black gram": 75, "Onion": 130, "Potato": 100}   # typical days (REFERENCE DEFAULTS, editable)
TRACTOR_INPUT_FILE = "tractor_inputs.csv"   # optional CSV: State,District,Tehsil + any of the four business columns above
# TYPICAL national sowing windows (REFERENCE DEFAULTS - not retrieved data; they differ by state/variety. Edit to your agronomy).
IDEAL_SOWING = {
    "Paddy":      {"kharif": "15 Jun - 31 Jul", "rabi": "15 Nov - 15 Jan", "winter": "15 Nov - 15 Jan", "summer": "15 Jan - 28 Feb", "autumn": "15 Apr - 31 May"},
    "Wheat":      {"rabi": "01 Nov - 25 Nov"},
    "Maize":      {"kharif": "15 Jun - 15 Jul", "rabi": "15 Oct - 15 Nov", "summer": "01 Feb - 28 Feb"},
    "Groundnut":  {"kharif": "15 Jun - 15 Jul", "rabi": "01 Nov - 15 Dec", "summer": "15 Jan - 15 Feb"},
    "Cotton":     {"kharif": "01 May - 30 Jun"},
    "Soyabean":   {"kharif": "15 Jun - 15 Jul"},
    "Bajra":      {"kharif": "01 Jul - 31 Jul", "summer": "15 Feb - 15 Mar"},
    "Sugarcane":  {"annual": "Feb - Mar (spring) / Sep - Oct (autumn)", "kharif": "Feb - Mar (spring) / Sep - Oct (autumn)"},
    "Black gram": {"kharif": "01 Jul - 31 Jul", "rabi": "15 Oct - 15 Nov", "summer": "15 Feb - 15 Mar"},
    "Onion":      {"kharif": "Jun - Jul", "rabi": "15 Oct - 30 Nov"},
    "Potato":     {"rabi": "15 Oct - 15 Nov"},
}
# crops pulled from Agmarknet: your crop name -> Agmarknet commodity name
CROP_COMMODITY = {"Paddy": "Paddy(Dhan)(Common)", "Wheat": "Wheat", "Maize": "Maize",
                  "Groundnut": "Groundnut", "Cotton": "Cotton", "Soyabean": "Soyabean",
                  "Bajra": "Bajra(Pearl Millet/Cumbu)", "Onion": "Onion", "Potato": "Potato"}
DATA_GOV_IN_KEY = ""   # OPTIONAL free key from data.gov.in (Agmarknet prices). Blank = skipped, nothing prompts.
AUTO_DOWNLOAD = False         # True = also push files to your browser (can stall on big files; use the Files panel instead)
GEE_ASK_UPLOAD = True          # no key found -> show an upload dialog for the JSON (if the browser blocks it: Drive search, then interactive login)
FORCE_REFRESH_WEATHER = False   # True = ignore today's saved IMD result and recompute
NEWS_DAYS = 14          # PIB / Google News RSS look-back
GEMINI_MODEL = "gemini-2.5-flash"   # used only if you add a Colab Secret GEMINI_API_KEY (free key: aistudio.google.com); tags news headlines, never fills numbers
NEWS_GOOGLE = True      # Google News RSS per state (context signals only)
MSP = {}   # fill yourself, e.g. {"Paddy": <Rs/qtl>} - left blank: not fetched from any API
# --- validation / reliability / AI recovery layer ---------------------------
GEE_SEARCH_DRIVE = True         # mount Google Drive and look for the service-account JSON there
RUN_SENTINEL2 = True            # Sentinel-2 10 m NDVI/EVI/NDWI + cloud cover over cropland (last S2_DAYS days); ~1-2 h, checkpointed
S2_DAYS = 30
S2_SCALE = 100                  # metres for zonal statistics (10 m pixels averaged at 100 m)
S2_MIN_CLEAR_OBS = 3            # fewer cloud-free observations than this -> flagged "Not comparable", never compared
RUN_CLIMATE_CHECK = True        # CHIRPS rainfall + ERA5-Land temperature as independent checks of IMD (Earth Engine)
AI_RECOVERY = True              # search Google News for missing state-level figures; AI (or regex) extracts - NEVER written into main tabs
AI_MAX_CALLS = 80               # hard cap on AI API calls per run (cost control)
OPENAI_MODEL = "gpt-4o-mini"    # only used if Colab Secret OPENAI_API_KEY exists (secondary) - edit if renamed
ANTHROPIC_MODEL = "claude-haiku-5-5"   # only used if Colab Secret ANTHROPIC_API_KEY exists (reviewer) - edit if renamed
SHEETS_CELL_LIMIT = 9_500_000   # Google Sheets max is 10 million cells; above this the validation tabs go to a 2nd Sheet
RELIABILITY_WEIGHTS = {"Source authority": 25, "Actual vs estimated": 25, "Latest data availability": 20,
                       "Government-satellite agreement": 20, "Geographic matching": 10}
VAL_STORE = {}       # values kept only for validation (proxies, independent observations), aligned to tehsil order
RECOVERY_LOG = []    # every source / API attempt -> SOURCE_RECOVERY_LOG tab
SHEET_URLS = {}      # every Google Sheet created -> printed in CELL 17

GROUPS = {"Foundation": "2E7D32", "Live intelligence": "1565C0",
          "Economics/history": "EF6C00", "Decision layer": "C62828", "Control": "455A64"}

SHEETS = [('TEHSIL_MASTER',
  'Foundation',
  ['State',
   'District',
   'Tehsil',
   'Tehsil Code',
   'Latitude',
   'Longitude',
   'Total Area Ha',
   'Cultivable Area Ha',
   'Net Sown Area Ha',
   'Gross Cropped Area Ha',
   'Cropping Intensity %',
   'Agro Climatic Zone',
   'Elevation']),
 ('SOIL_INTELLIGENCE',
  'Foundation',
  ['State',
   'District',
   'Tehsil',
   'Soil Type',
   'Soil Texture',
   'pH',
   'EC',
   'Organic Carbon',
   'N',
   'P',
   'K',
   'Soil Depth',
   'Water Holding Capacity',
   'Soil Moisture',
   'Soil Data Date',
   'Source',
   'Freshness',
   'Confidence']),
 ('WATER_IRRIGATION',
  'Foundation',
  ['State',
   'District',
   'Tehsil',
   'Irrigated Area Ha',
   'Rainfed Area Ha',
   'Irrigation %',
   'Canal',
   'Borewell',
   'Open Well',
   'Tank',
   'Reservoir',
   'Groundwater Level',
   'Groundwater Trend',
   'Water Stress',
   'Data Date',
   'Source',
   'Freshness']),
 ('WEATHER_LIVE',
  'Live intelligence',
  ['Date',
   'State',
   'District',
   'Tehsil',
   'Rainfall mm',
   'Normal Rainfall mm',
   'Rainfall Deviation %',
   'Rainfall 7D',
   'Rainfall 30D',
   'Dry Days',
   'Temperature',
   'Humidity',
   'Forecast 7D',
   'Extreme Weather Alert',
   'Data Date',
   'Source',
   'Freshness']),
 ('SATELLITE_INTELLIGENCE',
  'Live intelligence',
  ['Date',
   'State',
   'District',
   'Tehsil',
   'Crop',
   'NDVI',
   'NDVI 5Y Avg',
   'NDVI Anomaly %',
   'EVI',
   'NDWI',
   'Soil Moisture',
   'LST',
   'Vegetation Trend',
   'Cloud Cover %',
   'Satellite Date',
   'Satellite Source',
   'Confidence']),
 ('CROP_SOWN',
  'Live intelligence',
  ['Season',
   'State',
   'District',
   'Tehsil',
   'Crop',
   'Variety',
   'Sown Area Ha',
   'Sowing %',
   'Sowing Start',
   'Sowing End',
   'Ideal Sowing Window',
   'Delay Days',
   'Crop Source',
   'Report Date',
   'Freshness',
   'Confidence']),
 ('CROP_STAGE',
  'Live intelligence',
  ['Date',
   'State',
   'District',
   'Tehsil',
   'Crop',
   'Sowing Date',
   'Expected Stage',
   'Estimated Current Stage',
   'Stage Progress %',
   'Days From Normal',
   'Satellite Signal',
   'Weather Signal',
   'Stage Confidence',
   'Source',
   'Data Date']),
 ('CROP_HEALTH',
  'Live intelligence',
  ['Date',
   'State',
   'District',
   'Tehsil',
   'Crop',
   'NDVI Score',
   'Moisture Score',
   'Rainfall Score',
   'Temperature Score',
   'Crop Stage Score',
   'Overall Crop Health',
   '5Y Comparison',
   'Health Trend',
   'Status',
   'Confidence',
   'Reason']),
 ('PEST_DISEASE',
  'Live intelligence',
  ['Date',
   'State',
   'District',
   'Tehsil',
   'Crop',
   'Pest/Disease',
   'Incidence',
   'Severity',
   'Affected Area',
   'Weather Favourability',
   'Alert Level',
   'Source',
   'Report Date',
   'Freshness',
   'Confidence']),
 ('YIELD_PRODUCTION',
  'Economics/history',
  ['Season',
   'State',
   'District',
   'Tehsil',
   'Crop',
   'Area Sown Ha',
   'Historical Yield',
   'Current Expected Yield',
   'Production Estimate',
   'Yield vs 5Y Avg %',
   'Yield Trend',
   'Confidence',
   'Data Date',
   'Source']),
 ('MARKET_ECONOMICS',
  'Economics/history',
  ['Date',
   'State',
   'District',
   'Tehsil',
   'Crop',
   'Mandi Arrival',
   'Min Price',
   'Modal Price',
   'Max Price',
   'MSP',
   'Price vs 5Y Avg %',
   'Price Trend',
   'Estimated Crop Value',
   'Source',
   'Freshness']),
 ('5_YEAR_HISTORY',
  'Economics/history',
  ['Date',
   'Year',
   'State',
   'District',
   'Tehsil',
   'Crop',
   'Parameter',
   'Value',
   '5Y Average',
   '5Y Min',
   '5Y Max',
   'Current vs 5Y %',
   'YoY %',
   'Trend',
   'Source']),
 ('CROP_CALENDAR',
  'Economics/history',
  ['State',
   'District',
   'Tehsil',
   'Crop',
   'Season',
   'Ideal Sowing Start',
   'Ideal Sowing End',
   'Germination',
   'Vegetative',
   'Flowering',
   'Grain/Fruit Formation',
   'Maturity',
   'Harvest Window',
   'Typical Duration']),
 ('FARMER_AGRI_ECOSYSTEM',
  'Economics/history',
  ['State',
   'District',
   'Tehsil',
   'Farmers',
   'Small/Marginal Farmers %',
   'Average Holding',
   'FPOs',
   'Warehouses',
   'Cold Storage',
   'Mandis',
   'Mills',
   'Seed Dealers',
   'Fertilizer Dealers',
   'Other Agri Infrastructure',
   'Data Year',
   'Source']),
 ('EXTREME_EVENTS',
  'Economics/history',
  ['Date',
   'State',
   'District',
   'Tehsil',
   'Event',
   'Start Date',
   'End Date',
   'Severity',
   'Area Affected',
   'Crop Affected',
   'Historical Frequency',
   '5Y Comparison',
   'Current Risk',
   'Source']),
 ('AGRI_INTELLIGENCE',
  'Decision layer',
  ['Date',
   'State',
   'District',
   'Tehsil',
   'Crop',
   'Crop Health',
   'Water Score',
   'Weather Score',
   'Soil Score',
   'Yield Outlook',
   'Market Score',
   'Overall Agri Score',
   'Risk',
   'Opportunity',
   'Confidence',
   'Key Reason',
   'Recommended Action']),
 ('ALERTS',
  'Decision layer',
  ['Alert Date',
   'State',
   'District',
   'Tehsil',
   'Crop',
   'Alert Type',
   'Severity',
   'Trigger',
   'Current Value',
   'Normal Value',
   'Deviation',
   'Days Valid',
   'Recommended Action',
   'Status',
   'Confidence']),
 ('TRACTOR_OPPORTUNITY',
  'Decision layer',
  ['Date',
   'State',
   'District',
   'Tehsil',
   'Crop Area',
   'Crop Health',
   'Farm Prosperity',
   'Mechanisation',
   'Historical Tractor Demand',
   'Replacement Potential',
   'Crop Economy',
   'Tractor Opportunity Score',
   'Opportunity Reason']),
 ('DATA_FRESHNESS_CONTROL',
  'Control',
  ['Source',
   'Parameter',
   'Report Date',
   'Retrieved Date',
   'Age Days',
   'Maximum Allowed Days',
   'Freshness Status',
   'Use For Live?',
   'Source Priority',
   'Confidence']),
 ('SOURCE_REGISTRY',
  'Control',
  ['Source',
   'Data Type',
   'Geography',
   'API/MCP Available',
   'Update Frequency',
   'Historical Available',
   'Priority',
   'Reliability',
   'Access Status',
   'Last Successful Update'])]

# every data tab gets a visible source + source date/year column (the control tabs 19-20 are sources themselves)
_NEEDS_SOURCE = {"TEHSIL_MASTER", "CROP_HEALTH", "CROP_CALENDAR", "AGRI_INTELLIGENCE", "ALERTS", "TRACTOR_OPPORTUNITY"}
_STAMP_TABS = [t[0] for t in SHEETS[:18]]
SHEETS = [(n, g, h + (["Source"] if n in _NEEDS_SOURCE else []) + (["Source Date"] if n in _STAMP_TABS else []))
          for n, g, h in SHEETS]

# --- recommendation layer: new tabs ADDED after the original ones (original tabs/columns unchanged) ---
GROUPS["Validation"] = "6A1B9A"
VALIDATION_TABS = ["RECORD_METADATA", "DATA_VALIDATION", "DATA_RELIABILITY", "DATA_GAP_REGISTER",
                   "SOURCE_RECOVERY_LOG", "AI_ESTIMATION_REGISTER"]
SHEETS += [
    ("RECORD_METADATA", "Validation", ["State", "District", "Tehsil", "Tab", "Parameter", "Value", "Unit", "Source", "Source URL",
                                       "Source Level", "Output Level", "Data Type", "Reference Period", "Publication Date",
                                       "Last Updated", "Confidence", "Validation Status"]),
    ("DATA_VALIDATION", "Validation", ["State", "District", "Tehsil", "Parameter", "Government / Primary Value", "Primary Source",
                                       "Primary Period", "Independent Value", "Independent Source", "Independent Period", "Unit",
                                       "Difference", "Difference %", "Agreement", "Possible Explanation", "Checked On"]),
    ("DATA_RELIABILITY", "Validation", ["State", "District", "Tehsil", "Source Authority", "Actual vs Estimated",
                                        "Latest Data Availability", "Govt-Satellite Agreement", "Geographic Matching",
                                        "Data Reliability Score", "Reliability Band", "Dimensions Applied", "Validation Note",
                                        "Parameters With Data", "Parameters Expected"]),
    ("DATA_GAP_REGISTER", "Validation", ["State", "District", "Tehsil", "Tab", "Parameter", "Gap Type", "Rows Affected",
                                         "Latest Source Date", "Recovery Attempted", "Recovery Result", "Next Best Source", "Note"]),
    ("SOURCE_RECOVERY_LOG", "Validation", ["Timestamp", "Step", "Source", "URL / Endpoint", "Purpose", "Status", "Records",
                                           "Fallback Used", "Message"]),
    ("AI_ESTIMATION_REGISTER", "Validation", ["Date", "State", "District", "Tehsil", "Parameter", "Extracted Value", "Unit",
                                              "Reference Period", "Evidence (quote)", "Evidence URL", "Publisher", "Publication Date",
                                              "Method", "Confidence", "Validation Status", "Used In Main Tabs"]),
]

SHEET_NAMES = [s[0] for s in SHEETS]
HEADERS = {s[0]: s[2] for s in SHEETS}
TEHSIL_LEVEL = ["TEHSIL_MASTER", "SOIL_INTELLIGENCE", "WATER_IRRIGATION",
                "WEATHER_LIVE", "SATELLITE_INTELLIGENCE", "FARMER_AGRI_ECOSYSTEM"]

# Formulas written ONLY into empty cells ([Header] = same-row cell)
FORMULAS = {
    "TEHSIL_MASTER": {
        "Cropping Intensity %": 'IF(OR([Gross Cropped Area Ha]="",[Net Sown Area Ha]=""),"",IFERROR([Gross Cropped Area Ha]/[Net Sown Area Ha]*100,""))'},
    "SOIL_INTELLIGENCE": {
        "Freshness": f'IF([Soil Data Date]="","",IF(TODAY()-[Soil Data Date]<={FRESHNESS_DAYS["SOIL_INTELLIGENCE"]},"VALID","STALE"))'},
    "WATER_IRRIGATION": {
        "Irrigation %": 'IF(OR([Irrigated Area Ha]="",[Rainfed Area Ha]=""),"",IFERROR([Irrigated Area Ha]/([Irrigated Area Ha]+[Rainfed Area Ha])*100,""))',
        "Freshness": f'IF([Data Date]="","",IF(TODAY()-[Data Date]<={FRESHNESS_DAYS["WATER_IRRIGATION"]},"VALID","STALE"))'},
    "WEATHER_LIVE": {   # deviation = 30-day actual vs 30-day normal (IMD style departure)
        "Rainfall Deviation %": 'IF(OR([Rainfall 30D]="",[Normal Rainfall mm]=""),"",IFERROR(([Rainfall 30D]-[Normal Rainfall mm])/[Normal Rainfall mm]*100,""))',
        "Freshness": f'IF([Data Date]="","",IF(TODAY()-[Data Date]<={FRESHNESS_DAYS["WEATHER_LIVE"]},"VALID","STALE"))'},
    "SATELLITE_INTELLIGENCE": {
        "NDVI Anomaly %": 'IF(OR([NDVI]="",[NDVI 5Y Avg]=""),"",IFERROR(([NDVI]-[NDVI 5Y Avg])/[NDVI 5Y Avg]*100,""))'},
    "CROP_SOWN": {
        "Freshness": 'IF([Report Date]="","",IF(TODAY()-[Report Date]<=30,"VALID","STALE"))'},
    "MARKET_ECONOMICS": {
        "Freshness": f'IF([Date]="","",IF(TODAY()-[Date]<={FRESHNESS_DAYS["MARKET_ECONOMICS"]},"VALID","STALE"))'},
}

META = {}        # run facts used by the lineage tab
DATA = {}        # tab name -> DataFrame
T = None         # tehsil table (one row per tehsil, + geometry)
FRESH_RUN = {}   # source -> date of the newest data actually fetched
NOTES = {}       # (sheet, header) -> comment text written on the header cell


def ram(tag=""):
    """Print memory use - a crashed/recycled runtime shows up in the browser as 'Failed to fetch'."""
    try:
        import psutil
        m = psutil.virtual_memory()
        print(f"  [memory{(' ' + tag) if tag else ''}] {m.used / 1e9:.1f} GB used of {m.total / 1e9:.1f} GB ({m.percent:.0f}%)")
    except Exception:
        pass


def norm(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def clean_name(s):
    return re.sub(r"\s+", " ", str(s)).strip()


def key_of(df):
    return df["State"].map(norm) + "|" + df["District"].map(norm) + "|" + df["Tehsil"].map(norm)


def get_json(url, params=None, tries=4, timeout=60):
    for i in range(tries):
        try:
            r = requests.get(url, params=params, timeout=timeout)
            if r.status_code == 200:
                return r.json()
            time.sleep(8 * (i + 1) if r.status_code == 429 else 2 * (i + 1))
        except (requests.RequestException, ValueError):
            time.sleep(2 * (i + 1))
    return None


def uniq_key(df):
    k = key_of(df)
    return k + "#" + k.groupby(k).cumcount().astype(str)      # duplicate tehsil names stay distinct


def _secret(name):
    """Read a key from the environment or a Colab Secret (key icon in the left bar). Returns '' if absent."""
    v = os.environ.get(name, "")
    if not v:
        try:
            from google.colab import userdata
            v = userdata.get(name) or ""
        except Exception:
            v = ""
    return v.strip()


def rlog(step, source, url="", purpose="", status="OK", records=None, fallback="", msg=""):
    """One line per source / API attempt -> SOURCE_RECOVERY_LOG tab."""
    RECOVERY_LOG.append({"Timestamp": pd.Timestamp.now(tz="Asia/Kolkata").strftime("%Y-%m-%d %H:%M"), "Step": step,
                         "Source": source, "URL / Endpoint": str(url or "")[:200], "Purpose": purpose, "Status": status,
                         "Records": records, "Fallback Used": fallback, "Message": str(msg or "")[:250]})


# --- AI chain: Gemini (primary) -> GPT -> Claude, whichever Colab Secrets exist. AI only STRUCTURES text it is given. ---
AI_PROVIDERS = (("Gemini", "GEMINI_API_KEY"), ("GPT", "OPENAI_API_KEY"), ("Claude", "ANTHROPIC_API_KEY"))
_AI_N = [0]


def ai_available(exclude=()):
    return any(_secret(s) for n, s in AI_PROVIDERS if n not in exclude)


def _parse_json(text):
    t = re.sub(r"^```(?:json)?|```$", "", str(text).strip(), flags=re.M).strip()
    for op, cl in (("{", "}"), ("[", "]")):
        i, j = t.find(op), t.rfind(cl)
        if 0 <= i < j:
            try:
                return json.loads(t[i:j + 1])
            except ValueError:
                pass
    return None


def _ai_call(name, key, prompt):
    if name == "Gemini":
        for model in dict.fromkeys([GEMINI_MODEL, "gemini-2.5-flash", "gemini-2.0-flash"]):
            r = requests.post(f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent",
                              params={"key": key}, timeout=90,
                              json={"contents": [{"parts": [{"text": prompt}]}],
                                    "generationConfig": {"temperature": 0, "responseMimeType": "application/json"}})
            if r.status_code == 404:
                continue
            r.raise_for_status()
            return r.json()["candidates"][0]["content"]["parts"][0]["text"], model
        raise RuntimeError("no Gemini model name accepted - edit GEMINI_MODEL")
    if name == "GPT":
        r = requests.post("https://api.openai.com/v1/chat/completions", timeout=90, headers={"Authorization": f"Bearer {key}"},
                          json={"model": OPENAI_MODEL, "temperature": 0, "response_format": {"type": "json_object"},
                                "messages": [{"role": "user", "content": prompt}]})
        r.raise_for_status()
        return r.json()["choices"][0]["message"]["content"], OPENAI_MODEL
    r = requests.post("https://api.anthropic.com/v1/messages", timeout=90,
                      headers={"x-api-key": key, "anthropic-version": "2023-06-01", "content-type": "application/json"},
                      json={"model": ANTHROPIC_MODEL, "max_tokens": 2000, "temperature": 0,
                            "messages": [{"role": "user", "content": prompt}]})
    r.raise_for_status()
    return "".join(b.get("text", "") for b in r.json().get("content", [])), ANTHROPIC_MODEL


def ai_json(prompt, purpose, exclude=()):
    """Ask for JSON {"items":[...]}. Returns (items, provider) or (None, None). Capped by AI_MAX_CALLS; every call logged."""
    for name, sec in AI_PROVIDERS:
        key = _secret(sec)
        if not key or name in exclude:
            continue
        if _AI_N[0] >= AI_MAX_CALLS:
            rlog("AI", name, "", purpose, "SKIPPED", msg=f"AI_MAX_CALLS={AI_MAX_CALLS} reached")
            return None, None
        _AI_N[0] += 1
        try:
            text, model = _ai_call(name, key, prompt)
            obj = _parse_json(text)
            items = obj.get("items") if isinstance(obj, dict) else obj
            if not isinstance(items, list):
                raise ValueError("reply was not the expected JSON")
            rlog("AI", f"{name} ({model})", "", purpose, "OK", len(items))
            return items, name
        except Exception as e:
            body = (getattr(getattr(e, "response", None), "text", "") or "")[:120]
            print(f"  AI {name} failed: {type(e).__name__} {body}")
            rlog("AI", name, "", purpose, "FAILED", fallback="next provider", msg=f"{type(e).__name__} {body}")
    return None, None


def install(sheet, df):
    """Put fetched rows into a tab. Values already there (yours) win; blanks are filled."""
    df = df.reindex(columns=HEADERS[sheet]).reset_index(drop=True)
    old = DATA[sheet]
    if old.empty:
        DATA[sheet] = df
        return
    a, b = old.copy(), df.copy()
    a.index, b.index = uniq_key(old), uniq_key(df)
    have = set(a.index)
    order = list(a.index) + [k for k in b.index if k not in have]
    out = a.combine_first(b).reindex(order).reset_index(drop=True)
    DATA[sheet] = out.reindex(columns=HEADERS[sheet])


# ===== CELL 3 : TEHSIL LIST + BOUNDARIES (LGD via bharatlas.com) ============
import geopandas as gpd


def find_urls(obj, found=None):
    """Recursively collect download-looking URLs from a layer's JSON entry."""
    found = [] if found is None else found
    if isinstance(obj, dict):
        for v in obj.values():
            find_urls(v, found)
    elif isinstance(obj, list):
        for v in obj:
            find_urls(v, found)
    elif isinstance(obj, str) and obj.startswith("http") and re.search(
            r"\.(parquet|geojson|json|gpkg|zip)(\?|$)", obj, re.I):
        found.append(obj)
    return found


def discover_tehsil_url():
    data = get_json(f"{BHARATLAS}/layers")
    if data is None:
        return None
    layers = data if isinstance(data, list) else (
        data.get("layers") or data.get("data") or data.get("items") or [])
    cands = []
    for L in layers:
        text = json.dumps(L).lower()
        lid = str(L.get("id") or L.get("layer_id") or L.get("slug") or L.get("name") or "")
        if re.search(r"sub.?district", text) and "village" not in lid.lower() \
                and "block" not in lid.lower():
            cands.append((0 if lid.lower().startswith("lgd") else 1, lid, L))
    cands.sort(key=lambda x: x[0])
    print("Candidate sub-district layers:", [c[1] for c in cands[:6]])
    for _, lid, L in cands:
        urls = sorted(find_urls(L), key=lambda u: (".parquet" not in u.lower(), u))
        if urls:
            print("Using layer:", lid, "->", urls[0])
            return urls[0]
    return None


def download(url, path):
    with requests.get(url, stream=True, timeout=600, headers={"User-Agent": "Mozilla/5.0"}) as r:
        r.raise_for_status()
        with open(path, "wb") as f:
            for chunk in r.iter_content(1 << 20):
                f.write(chunk)
    return path


def read_geo(path):
    if path.lower().endswith(".parquet"):
        try:
            return gpd.read_parquet(path)
        except Exception:                       # plain parquet with a WKB geometry column
            from shapely import wkb
            df = pd.read_parquet(path)
            g = next((c for c in df.columns if norm(c) in ("geometry", "geom", "wkb", "wkbgeometry")), None)
            geom = df[g].map(lambda b: wkb.loads(bytes(b)) if b is not None else None)
            return gpd.GeoDataFrame(df.drop(columns=[g]), geometry=geom, crs=4326)
    return gpd.read_file(path)


def pick(cols, *must, exclude=()):
    for c in cols:
        n = norm(c)
        if all(m in n for m in must) and not any(e in n for e in exclude):
            return c
    return None


def to_tehsil_table(gdf):
    cols = [c for c in gdf.columns if c != gdf.geometry.name]
    c_state = (pick(cols, "stname") or pick(cols, "state", "name")
               or pick(cols, "state", exclude=("code", "id")))
    c_dist = (pick(cols, "dtname") or pick(cols, "district", "name")
              or pick(cols, "district", exclude=("code", "id")))
    c_teh = (pick(cols, "sdtname") or pick(cols, "subdistrict", "name") or pick(cols, "tehsil")
             or pick(cols, "taluk") or pick(cols, "subdistrict", exclude=("code", "id")))
    c_code = (pick(cols, "sdtcode") or pick(cols, "subdistrict", "lgd")
              or pick(cols, "subdistrict", "code", exclude=("census",)))
    if not (c_state and c_dist and c_teh):
        raise ValueError(f"Could not find state/district/sub-district columns. Columns: {cols}")
    g = gdf.to_crs(4326)
    t = pd.DataFrame({"State": g[c_state].map(clean_name), "District": g[c_dist].map(clean_name),
                      "Tehsil": g[c_teh].map(clean_name),
                      "Tehsil Code": g[c_code].astype(str) if c_code else None})
    pts = g.geometry.representative_point()
    t["Latitude"], t["Longitude"] = pts.y.round(5), pts.x.round(5)
    t["Total Area Ha"] = (g.to_crs(6933).geometry.area / 10000).round(0)
    t["geometry"] = g.geometry.values
    return t


def finalise_table(t):
    t = t[(t["Tehsil"] != "") & (t["State"] != "")].copy()
    t = t.drop_duplicates(subset=["State", "District", "Tehsil", "Tehsil Code"])
    dup = t.duplicated(["State", "District", "Tehsil"], keep=False)
    t.loc[dup, "Tehsil"] = t.loc[dup, "Tehsil"] + " [" + t.loc[dup, "Tehsil Code"].astype(str) + "]"
    return t.sort_values(["State", "District", "Tehsil"]).reset_index(drop=True)


url = TEHSIL_SOURCE_URL or discover_tehsil_url()
if not url:
    raise SystemExit(
        "Could not find the LGD sub-district file. Open https://bharatlas.com , find "
        "'Sub-districts (2024) LGD', copy its parquet link into TEHSIL_SOURCE_URL (CELL 2), "
        "then re-run CELL 2 and CELL 3.")
ext = ".parquet" if ".parquet" in url.lower() else ".geojson"
local = download(url, "tehsils" + ext)
g_raw = read_geo(local)
print("Downloaded layer:", url)
print("  rows:", len(g_raw), "| columns:", [c for c in g_raw.columns][:25])
if len(g_raw) == 0:
    raise SystemExit("The downloaded layer is empty. Copy a working LGD sub-district Parquet/GeoJSON "
                     "link from https://bharatlas.com into TEHSIL_SOURCE_URL (CELL 2) and re-run CELL 2-3.")
T = finalise_table(to_tehsil_table(g_raw))
if len(T) == 0:
    raise SystemExit("No tehsils found in the layer - see the column list printed above.")
_ONLY = [x.strip().upper() for x in os.environ.get("AGRI_STATES", "").split(",") if x.strip()]   # [GitHub] test runs only
if _ONLY:
    T = T[T["State"].str.upper().isin(_ONLY)].reset_index(drop=True)
    print(f"TEST MODE - AGRI_STATES={_ONLY}: {len(T):,} tehsils kept (leave AGRI_STATES empty for all India)")
META["lgd_rows"], META["lgd_kept"] = len(g_raw), len(T)
META["lgd_dup"] = int(T["Tehsil"].str.contains(r" \[").sum())
rlog("CELL 3", "LGD sub-districts via bharatlas.com", url, "Tehsil list + boundaries", "OK", len(T),
     msg=f"{len(g_raw):,} polygons in layer, {len(T):,} kept")
if len(T) < 3000:
    print(f"WARNING: only {len(T):,} rows - this may be the wrong layer (districts/blocks?). "
          "India has roughly 7,000 sub-districts; set TEHSIL_SOURCE_URL to the right layer.")

for _s in SHEET_NAMES:
    DATA[_s] = pd.DataFrame(columns=HEADERS[_s])
for _s in TEHSIL_LEVEL:
    df = pd.DataFrame(columns=HEADERS[_s], index=range(len(T)))
    df["State"], df["District"], df["Tehsil"] = T["State"], T["District"], T["Tehsil"]
    DATA[_s] = df
DATA["TEHSIL_MASTER"]["Tehsil Code"] = T["Tehsil Code"]
DATA["TEHSIL_MASTER"]["Latitude"] = T["Latitude"]
DATA["TEHSIL_MASTER"]["Longitude"] = T["Longitude"]
DATA["TEHSIL_MASTER"]["Total Area Ha"] = T["Total Area Ha"]
NOTES[("TEHSIL_MASTER", "Latitude")] = "Representative point of the LGD sub-district boundary."
NOTES[("TEHSIL_MASTER", "Total Area Ha")] = "Computed from the LGD boundary polygon (equal-area projection)."

print(f"Loaded {len(T):,} tehsils in {T['State'].nunique()} states/UTs, "
      f"{T.groupby(['State','District']).ngroups} districts.")
print(T[["State", "District", "Tehsil", "Latitude", "Longitude", "Total Area Ha"]].head())


# ===== CELL 4 : GOVERNMENT LAYERS (CGWB, MI Census, CoRE Stack, agro-climatic zones) ====
# Real government/ICAR/Planning-Commission layers published as open Parquet on bharatlas.com,
# matched to each tehsil by location. Column names are auto-detected and printed - check them.
GOV = {
    "agro": "https://bharatlas.com/api/dl/agriculture/agro-climatic-zones/Agro_Climatic_Zones.parquet",
    "cgwb": "https://bharatlas.com/api/dl/water/groundwater/CGWB_Groundwater_Extraction.parquet",
    "mi6": "https://bharatlas.com/api/dl/water/mi6-wells-districts/MI6_Wells_Districts_2017_18.parquet",
    "corestack": "https://bharatlas.com/api/dl/agriculture/corestack-lulc-blocks/CoREStack_LULC_Blocks_2023_24.parquet",
}


def load_layer(key):
    path = f"{key}.parquet"
    if not os.path.exists(path):
        download(GOV[key], path)
    g = read_geo(path).to_crs(4326)
    print(f"{key}: {len(g):,} rows | columns: {[c for c in g.columns if c != g.geometry.name][:30]}")
    return g


def tehsil_points():
    return gpd.GeoDataFrame({"ti": range(len(T))},
                            geometry=gpd.points_from_xy(T["Longitude"], T["Latitude"]), crs=4326)


def join_points(layer, cols):
    """Value of `cols` from the layer polygon that contains each tehsil point (nearest within 15 km if none)."""
    gname = layer.geometry.name
    sub = layer[list(cols) + [gname]]
    pts = tehsil_points()
    j = gpd.sjoin(pts, sub, how="left", predicate="within").drop_duplicates("ti").set_index("ti").sort_index()
    miss = list(j.index[j["index_right"].isna()])
    if len(miss):
        try:
            left = pts.loc[miss, ["ti", "geometry"]].to_crs(6933)
            right = sub.reset_index(drop=True).to_crs(6933)
            nn = gpd.sjoin_nearest(left, right, how="left", distance_col="_dist_m")
            nn = nn.drop_duplicates("ti").set_index("ti")
            nn = nn[nn["_dist_m"] <= 15000]
            for c in cols:
                j.loc[nn.index, c] = nn[c].values
        except Exception as e:
            print("  nearest fallback skipped:", type(e).__name__, str(e)[:120])
    return j[list(cols)].reindex(range(len(T)))


def text_cols(layer):
    return [c for c in layer.columns if c != layer.geometry.name
            and (layer[c].dtype == object or pd.api.types.is_string_dtype(layer[c]))]


def num_cols(layer):
    return [c for c in layer.columns if c != layer.geometry.name and pd.api.types.is_numeric_dtype(layer[c])]


def label_col(layer):
    cands = []
    for c in text_cols(layer):
        n = layer[c].nunique(dropna=True)
        if 2 <= n <= 60:
            cands.append((0 if "zone" in norm(c) else 1 if "name" in norm(c) else 2, n, c))
    return sorted(cands)[0][2] if cands else None


def stress_from_category(v):
    t = str(v).lower()
    if "semi" in t: return "Medium"
    if "over" in t or "critical" in t: return "High"
    if "safe" in t: return "Low"
    return None


def category_col(layer):
    best, score = None, 0
    for c in text_cols(layer):
        sc = layer[c].dropna().astype(str).str.lower().str.contains("safe|critical|over").sum()
        if sc > score:
            best, score = c, sc
    return best


def pick_num(layer, *keys):
    """Numeric columns whose name has all `keys`; prefer one named 'total', else all of them."""
    c = [x for x in num_cols(layer) if all(k in norm(x) for k in keys)]
    tot = [x for x in c if "total" in norm(x)]
    return tot[:1] or c


def agro_zone():
    L = load_layer("agro")
    col = label_col(L)
    soil = next((c for c in text_cols(L) if norm(c) == "soil"), None)
    cols = [col] + ([soil] if soil else [])
    j = join_points(L, cols)
    DATA["TEHSIL_MASTER"]["Agro Climatic Zone"] = j[col].values
    NOTES[("TEHSIL_MASTER", "Agro Climatic Zone")] = f"Planning Commission agro-climatic zone (layer column '{col}')."
    print("Agro-climatic zone filled:", int(j[col].notna().sum()), "tehsils")
    if soil:
        S = DATA["SOIL_INTELLIGENCE"]
        S["Soil Type"] = j[soil].values
        S["Source"] = "Planning Commission agro-climatic zone (zone-level dominant soil)"
        NOTES[("SOIL_INTELLIGENCE", "Soil Type")] = "ZONE-level dominant soil of the agro-climatic zone, not a tehsil soil survey."
        print("Soil type (zone level) filled:", int(j[soil].notna().sum()), "tehsils")


def groundwater_stress():
    L = load_layer("cgwb")
    cat = category_col(L)
    stage = next((c for c in num_cols(L) if "stage" in norm(c)), None)
    cols = [c for c in (cat, stage) if c]
    j = join_points(L, cols)
    W = DATA["WATER_IRRIGATION"]
    W["Water Stress"] = j[cat].map(stress_from_category).values
    W["Source"] = "CGWB 2020 groundwater assessment (block/taluk level)"
    NOTES[("WATER_IRRIGATION", "Water Stress")] = (
        f"From CGWB category column '{cat}': Safe=Low, Semi-Critical=Medium, Critical/Over-Exploited=High."
        + (f" Stage of extraction column: '{stage}'." if stage else ""))
    print("Groundwater stress filled:", int(W["Water Stress"].notna().sum()), "tehsils; category column:", cat)


def choose_well_cols(cols):
    """Exact MI-census columns first; filtered fallback so shares / in-use / sub-types are never double counted."""
    bad = re.compile(r"share|in_use|energy|pucca|kutcha|cum_bore|total|n_|census")
    dug = ["dugwell"] if "dugwell" in cols else [c for c in cols if "dug" in c.lower() and not bad.search(c.lower())]
    tube = [c for c in ("shallow_tw", "medium_tw", "deep_tw") if c in cols]
    if not tube:
        tube = [c for c in cols if re.search(r"tw|tube", c.lower()) and not bad.search(c.lower())]
    return dug, tube


def apportion(values, group, weight):
    """Split a district total over its tehsils in proportion to `weight` (estimate, sums back to the district total)."""
    w = weight.fillna(0).astype(float)
    tot = w.groupby(group).transform("sum")
    size = group.groupby(group).transform("size")
    share = np.where(tot > 0, w / tot.where(tot > 0, 1), 1.0 / size)
    return (values * share).round(0)


def wells_by_district():
    L = load_layer("mi6")
    dug, tube = choose_well_cols(num_cols(L))
    print("  dugwell columns used:", dug, "| tubewell columns used:", tube)
    cols = dug + [c for c in tube if c not in dug]
    if not cols:
        print("  no well-count columns detected - skipped")
        return
    key = "dist_lgd" if "dist_lgd" in L.columns else None
    j = join_points(L, cols + ([key] if key else []))
    W, M = DATA["WATER_IRRIGATION"], DATA["TEHSIL_MASTER"]
    group = (j[key].astype(str) if key else (T["State"] + "|" + T["District"])).where(j[cols[0]].notna(), "none|" + T.index.astype(str))
    weight = M["Net Sown Area Ha"].astype(float).where(M["Net Sown Area Ha"].notna(), M["Total Area Ha"].astype(float))
    if dug:
        W["Open Well"] = apportion(j[dug].sum(axis=1, min_count=1), group, weight).values
    if tube:
        W["Borewell"] = apportion(j[tube].sum(axis=1, min_count=1), group, weight).values
    NOTES[("WATER_IRRIGATION", "Open Well")] = (f"ESTIMATE: district dugwell total (Minor Irrigation Census 2017-18, cols {dug}) split over the district's "
                                               "tehsils in proportion to cropland area (CoRE Stack). Tehsil counts are not published.")
    NOTES[("WATER_IRRIGATION", "Borewell")] = (f"ESTIMATE: district tubewell total (shallow+medium+deep, MI Census 2017-18, cols {tube}) split over the "
                                              "district's tehsils in proportion to cropland area (CoRE Stack).")
    print("Wells (district totals apportioned to tehsils by cropland share) filled.")


def choose_crop_cols(cols):
    ha = lambda c: re.search(r"(^|_)ha($|_)", c.lower()) and not re.search(r"pct|percent", c.lower())
    single = [c for c in cols if "single" in norm(c) and ha(c)]
    double = [c for c in cols if "doubl" in norm(c) and ha(c)]
    triple = [c for c in cols if "tripl" in norm(c) and ha(c)]
    return single, double, triple


def corestack_cropping():
    B = load_layer("corestack")
    sc, dc, tc = choose_crop_cols(num_cols(B))
    print("  single:", sc, "| double:", dc, "| triple:", tc)
    if not (sc and dc and tc):
        print("  could not detect cropping columns - skipped (see column list above)")
        return
    vals = pd.DataFrame({"s": B[sc].astype(float).sum(axis=1, min_count=1),
                         "d": B[dc].astype(float).sum(axis=1, min_count=1),
                         "t": B[tc].astype(float).sum(axis=1, min_count=1),
                         "o": B["orchard_ha"].astype(float) if "orchard_ha" in B else 0.0,
                         "cl": B["cropland_ha"].astype(float) if "cropland_ha" in B else np.nan,
                         "snk": (B[[c for c in sc if "nonkharif" in norm(c)]].astype(float).sum(axis=1, min_count=1)
                                 if any("nonkharif" in norm(c) for c in sc) else 0.0),
                         "wp": B["water_perennial_ha"].astype(float) if "water_perennial_ha" in B else np.nan})
    if "cropland_ha" in B:
        gap = (vals.sum(axis=1) - B["cropland_ha"]).abs().median()
        print(f"  check: median |single+double+triple - cropland_ha| per block = {gap:,.0f} ha")
    pts = B[[B.geometry.name]].copy()
    pts[B.geometry.name] = B.geometry.representative_point()
    tg = gpd.GeoDataFrame({"ti": range(len(T))}, geometry=T["geometry"].values, crs=4326)
    j = gpd.sjoin(gpd.GeoDataFrame(vals.assign(geometry=pts.geometry.values), crs=4326), tg,
                  how="inner", predicate="within")
    agg = j.groupby("ti")[["s", "d", "t", "o", "snk", "wp", "cl"]].sum(min_count=1).reindex(range(len(T)))
    M = DATA["TEHSIL_MASTER"]
    M["Net Sown Area Ha"] = (agg["s"] + agg["d"] + agg["t"]).round(0).values
    VAL_STORE["core_nsa"] = (agg["s"] + agg["d"] + agg["t"]).to_numpy(float)      # kept for govt-vs-satellite validation
    M["Gross Cropped Area Ha"] = (agg["s"] + 2 * agg["d"] + 3 * agg["t"]).round(0).values
    cult = (agg["cl"] + agg["o"]).where(agg["cl"].notna(), agg["s"] + agg["d"] + agg["t"] + agg["o"])
    M["Cultivable Area Ha"] = cult.round(0).values
    NOTES[("TEHSIL_MASTER", "Cultivable Area Ha")] = ("PROXY: CoRE Stack 2023-24 cropland (incl. fallow cropland) + orchard area, satellite-based. "
                                                      "True cultivable land also includes fallow and culturable waste, which is not open data.")
    NOTES[("TEHSIL_MASTER", "Net Sown Area Ha")] = ("CoRE Stack 2023-24 (10 m satellite model, NOT revenue records): single (kharif + "
                                                    "non-kharif) + double + triple cropped area, summed from blocks whose centre lies in the tehsil.")
    NOTES[("TEHSIL_MASTER", "Gross Cropped Area Ha")] = "single + 2 x double + 3 x triple cropped area (CoRE Stack, block->tehsil by centre point)."
    print("Net sown / gross cropped filled:", int(agg["s"].notna().sum()), "tehsils")
    fill_water_proxy(agg)


def fill_water_proxy(agg):
    """FALLBACK for irrigated/rainfed area: cropping-pattern proxy from CoRE Stack (replaced by official DES figures in CELL 9
    where the district matches). Non-monsoon cropping (rabi-only, double, triple) needs irrigation or stored water."""
    W = DATA["WATER_IRRIGATION"]
    W["Irrigated Area Ha"] = (agg["snk"] + agg["d"] + agg["t"]).round(0).values
    VAL_STORE["irr_proxy"] = (agg["snk"] + agg["d"] + agg["t"]).to_numpy(float)   # satellite proxy kept for validation
    W["Rainfed Area Ha"] = (agg["s"] - agg["snk"]).round(0).values
    W["Reservoir"] = agg["wp"].round(0).values
    W["Data Date"] = date(2024, 6, 30)
    W["_irr_src"] = np.where(agg["s"].notna().values, "CoRE", "")
    NOTES[("WATER_IRRIGATION", "Irrigated Area Ha")] = ("PROXY: CoRE Stack 2023-24 non-monsoon cropping (rabi-only + double + triple cropped area). "
                                                       "Replaced by official DES Land Use Statistics (district, apportioned) where available.")
    NOTES[("WATER_IRRIGATION", "Rainfed Area Ha")] = "PROXY: CoRE Stack single-kharif cropped area, or DES net sown minus irrigated where DES applies."
    NOTES[("WATER_IRRIGATION", "Reservoir")] = "Area (ha) of perennial surface water bodies from CoRE Stack 2023-24 (satellite) - not storage capacity."
    NOTES[("WATER_IRRIGATION", "Data Date")] = "Agricultural year end of the main source (CoRE 2023-24 = 30-Jun-2024, or DES year end). Per-field source dates are in DATA_LINEAGE."


ram("before government layers")
for _step, _src, _url in ((agro_zone, "Planning Commission agro-climatic zones", GOV["agro"]),
                          (groundwater_stress, "CGWB groundwater assessment 2020", GOV["cgwb"]),
                          (corestack_cropping, "CoRE Stack LULC 2023-24", GOV["corestack"]),
                          (wells_by_district, "Minor Irrigation Census 2017-18", GOV["mi6"])):
    try:
        _step()
        rlog("CELL 4", _src, _url, _step.__name__, "OK")
    except Exception as e:
        print(f"{_step.__name__} failed: {type(e).__name__}: {e}")
        rlog("CELL 4", _src, _url, _step.__name__, "FAILED", msg=f"{type(e).__name__}: {e}")


# ===== CELL 5 : OPTIONAL - MERGE DATA YOU ALREADY HAVE ======================
# Skip if you have no files. Otherwise set EXISTING_FILES = ["my.xlsx", ...]
# (sheet names / CSV names = tab names). Your values win over fetched ones.
EXISTING_FILES = []


def load_existing(paths):
    found = {}
    for p in paths:
        if p.lower().endswith((".xlsx", ".xls")):
            for name, df in pd.read_excel(p, sheet_name=None).items():
                if name.strip().upper() in HEADERS:
                    found[name.strip().upper()] = df
        elif p.lower().endswith(".csv"):
            stem = os.path.splitext(os.path.basename(p))[0].strip().upper()
            if stem in HEADERS:
                found[stem] = pd.read_csv(p)
    return found


for sheet, df in load_existing(EXISTING_FILES).items():
    df = df.dropna(how="all").reindex(columns=HEADERS[sheet])
    if sheet in TEHSIL_LEVEL:
        old = DATA[sheet]
        DATA[sheet] = pd.concat([df, old[~key_of(old).isin(set(key_of(df)))]], ignore_index=True)
    else:
        DATA[sheet] = pd.concat([DATA[sheet], df], ignore_index=True)
    print(f"  merged {sheet}: {len(DATA[sheet]):,} rows")


# ===== CELL 6 : ELEVATION (Open-Meteo / Copernicus DEM) =====================
ELEV_CACHE = "elevation_cache.json"          # elevation never changes, so it is cached permanently
_ec = json.load(open(ELEV_CACHE)) if os.path.exists(ELEV_CACHE) else {}
_ekey = lambda k: f"{T.loc[k, 'Latitude']},{T.loc[k, 'Longitude']}"


def fetch_elevation(idx, batch=100):
    out = {}
    for i in range(0, len(idx), batch):
        part = idx[i:i + batch]
        ch = T.loc[part]
        d = get_json("https://api.open-meteo.com/v1/elevation", {
            "latitude": ",".join(map(str, ch["Latitude"])),
            "longitude": ",".join(map(str, ch["Longitude"]))}, tries=4)
        vals = (d or {}).get("elevation")
        if vals and len(vals) == len(part):
            out.update(dict(zip(part, vals)))
            for k, v in zip(part, vals):
                _ec[_ekey(k)] = v
            json.dump(_ec, open(ELEV_CACHE, "w"))
        time.sleep(1.0)
    return out


if "Elevation" not in T:
    T["Elevation"] = np.nan
for k in T.index[T["Elevation"].isna()]:                    # fill from the cache first
    if _ekey(k) in _ec:
        T.loc[k, "Elevation"] = _ec[_ekey(k)]
for _round in (1, 2):                       # second pass retries batches that were rate-limited
    todo = list(T.index[T["Elevation"].isna()])
    if not todo:
        break
    print(f"Elevation pass {_round}: {len(todo):,} tehsils to fetch")
    for k, v in fetch_elevation(todo).items():
        T.loc[k, "Elevation"] = v
    time.sleep(5)
DATA["TEHSIL_MASTER"]["Elevation"] = T["Elevation"].values
NOTES[("TEHSIL_MASTER", "Elevation")] = "Metres; Copernicus DEM via Open-Meteo at the representative point."
print("Elevation filled:", int(T["Elevation"].notna().sum()), "of", len(T))
rlog("CELL 6", "Copernicus DEM via Open-Meteo", "https://api.open-meteo.com/v1/elevation", "Elevation",
     "OK" if T["Elevation"].notna().all() else "PARTIAL", int(T["Elevation"].notna().sum()))


# ===== CELL 7 : WEATHER (IMD gridded + Open-Meteo forecast) =================
# Rainfall: IMD 0.25-degree gridded; temperature: IMD 1-degree. Nearest grid cell
# to each tehsil point. Needs a few hundred MB of downloads from imdpune.gov.in.
import contextlib, io
import imdlib as imd
import xarray as xr

WORK = "imd_data"
os.makedirs(WORK, exist_ok=True)


def imd_realtime(var, start, end):
    """IMD real-time daily grids. A day whose file is missing on the IMD server is SKIPPED (logged), instead of failing the
    whole 30-day window. Raises only if no day at all could be read."""
    opener = getattr(imd, "open_real_data", None)
    if opener is None:
        raise RuntimeError("Your imdlib has no open_real_data - run: pip install -U imdlib")
    try:
        with contextlib.redirect_stdout(io.StringIO()):          # imdlib prints one line per file
            imd.get_real_data(var, start.isoformat(), end.isoformat(), WORK)
    except Exception:
        pass                                                     # partial download - per-day loop below
    try:                                                         # fast path: every day present
        ds = opener(var, start.isoformat(), end.isoformat(), WORK).get_xarray()
        return ds[var].where(ds[var] > -900)
    except Exception:
        pass
    arrs, miss = [], []
    for k in range((end - start).days + 1):
        d = (start + timedelta(days=k)).isoformat()
        for attempt in range(2):
            try:
                with contextlib.redirect_stdout(io.StringIO()):
                    if attempt:
                        imd.get_real_data(var, d, d, WORK)
                    arrs.append(opener(var, d, d, WORK).get_xarray()[var].load())
                break
            except Exception:
                if attempt:
                    miss.append(d)
    if not arrs:
        raise RuntimeError(f"IMD {var}: no daily file readable for {start}..{end}")
    if miss:
        print(f"  IMD {var}: {len(miss)} day(s) not on the server, skipped: {', '.join(miss[:6])}{' ...' if len(miss) > 6 else ''}")
        rlog("CELL 7", f"IMD real-time {var}", "https://www.imdpune.gov.in/", "daily grids", "PARTIAL", len(arrs),
             msg=f"missing days: {', '.join(miss[:10])}")
    da = xr.concat(arrs, dim="time").sortby("time")
    return da.where(da > -900)


def at_points(da, lats, lons):
    s = da.sel(lat=xr.DataArray(lats, dims="pt"), lon=xr.DataArray(lons, dims="pt"),
               method="nearest")
    return s.transpose("time", "pt")


def imd_weather():
    lats, lons = T["Latitude"].values, T["Longitude"].values
    end = TODAY - timedelta(days=IMD_LAG_DAYS)
    start = end - timedelta(days=29)
    print(f"IMD real-time grids {start} to {end} ...")
    rain_da = at_points(imd_realtime("rain", start, end), lats, lons)
    rain = rain_da.values                                                       # (time, pt)
    tmax_da = at_points(imd_realtime("tmax", end - timedelta(days=13), end), lats, lons)   # 14 days: overlap with ERA5-Land check
    tmax = tmax_da.values
    tmin = at_points(imd_realtime("tmin", end - timedelta(days=13), end), lats, lons).values
    last_t = np.nanmean([tmax[-1], tmin[-1]], axis=0)
    VAL_STORE["imd_rain_daily"] = rain.astype("float32")                        # kept for the CHIRPS cross-check
    VAL_STORE["imd_dates"] = [pd.Timestamp(x).date() for x in rain_da["time"].values]
    VAL_STORE["imd_tmean_daily"] = np.nanmean([tmax, tmin], axis=0).astype("float32")   # for the ERA5-Land cross-check
    VAL_STORE["imd_t_dates"] = [pd.Timestamp(x).date() for x in tmax_da["time"].values]
    res = pd.DataFrame({
        "Rainfall mm": rain[-1], "Rainfall 7D": np.nansum(rain[-7:], axis=0),
        "Rainfall 30D": np.nansum(rain, axis=0),
        "Dry Days": np.sum(rain[-7:] < DRY_DAY_MM, axis=0),
        "Temperature": last_t})
    res.loc[np.isnan(rain[-1]), ["Rainfall mm", "Rainfall 7D", "Rainfall 30D", "Dry Days"]] = np.nan

    # 30-day normal: mean over past years of the SAME calendar window
    y1 = TODAY.year - 1
    while True:
        y0 = y1 - NORMAL_YEARS + 1
        try:
            print(f"  IMD archive {y0 - 1}-{y1}: downloading/reading (a few minutes) ...")
            with contextlib.redirect_stdout(io.StringIO()):
                imd.get_data("rain", y0 - 1, y1, fn_format="yearwise", file_dir=WORK)
            ds = imd.open_data("rain", y0 - 1, y1, "yearwise", WORK).get_xarray()
            break
        except Exception as e:                     # latest year may not be published yet
            print("  IMD archive not ready up to", y1, "-", type(e).__name__, "; trying", y1 - 1)
            y1 -= 1
            if y1 < TODAY.year - 4:
                raise
    print("  sampling the archive at every tehsil point ...")
    arch = at_points(ds["rain"].where(ds["rain"] > -900), lats, lons).compute()
    totals = []
    for y in range(y0, y1 + 1):
        try:
            e = date(y, end.month, end.day)
        except ValueError:
            e = date(y, end.month, 28)
        s = e - timedelta(days=29)
        totals.append(arch.sel(time=slice(str(s), str(e))).sum("time", skipna=True).values)
    res["Normal Rainfall mm"] = np.mean(totals, axis=0).round(1)
    NOTES[("WEATHER_LIVE", "Normal Rainfall mm")] = (
        f"Mean 30-day rainfall total for the same calendar window, IMD gridded {y0}-{y1}. "
        "Rainfall Deviation % compares Rainfall 30D with this normal.")
    res["Data Date"] = end
    META["imd_years"] = f"{y0}-{y1}"
    FRESH_RUN["IMD gridded"] = end
    return res.round(1)


def forecast_7d(batch=25, pause=8.0, passes=4, wait=75):
    """Forecast per DISTRICT centre. Open-Meteo's free limit is per IP and Colab IPs are shared, so batches
    can get HTTP 429: failed batches are retried after a wait, and results are cached for today."""
    g = T.groupby(["State", "District"], sort=False).agg(
        lat=("Latitude", "mean"), lon=("Longitude", "mean")).reset_index()
    keys = (g["State"] + "|" + g["District"]).tolist()
    cache_file = f"forecast_{TODAY}.json"
    res = json.load(open(cache_file)) if os.path.exists(cache_file) else {}
    if res:
        print(f"  forecast cache: {len(res):,} districts already fetched today")
    starts = list(range(0, len(g), batch))
    for pass_no in range(passes):
        todo = [i for i in starts if not all(k in res for k in keys[i:i + batch])]
        if not todo:
            break
        if pass_no:
            print(f"  pass {pass_no + 1}: {len(todo)} batches left - waiting {wait}s for the rate limit to clear ...")
            time.sleep(wait)
        for n, i in enumerate(todo, 1):
            ch = g.iloc[i:i + batch]
            p = {"latitude": ",".join(str(round(float(x), 3)) for x in ch["lat"]),
                 "longitude": ",".join(str(round(float(x), 3)) for x in ch["lon"]),
                 "daily": "precipitation_sum,relative_humidity_2m_mean",
                 "forecast_days": 7, "timezone": "Asia/Kolkata"}
            d = get_json("https://api.open-meteo.com/v1/forecast", p, tries=2)
            if d is None:
                p["daily"] = "precipitation_sum"
                d = get_json("https://api.open-meteo.com/v1/forecast", p, tries=2)
            locs = d if isinstance(d, list) else ([d] if d else [])
            for k, loc in zip(keys[i:i + batch], locs):
                dd = loc.get("daily", {})
                r7 = round(sum(x or 0 for x in dd.get("precipitation_sum", [])), 1)
                h = [x for x in (dd.get("relative_humidity_2m_mean") or []) if x is not None]
                res[k] = [r7, round(float(np.mean(h)), 0) if h else None]
            json.dump(res, open(cache_file, "w"))
            if n % 5 == 0 or n == len(todo):
                print(f"  forecast: {len(res):,}/{len(g):,} districts done")
            time.sleep(pause)
    left = len(g) - len(res)
    if left:
        print(f"  {left:,} districts still have no forecast (rate-limited). Re-run this cell later - it only fetches what is missing.")
    FRESH_RUN["Open-Meteo"] = TODAY
    tk = [f"{a}|{b}" for a, b in zip(T["State"], T["District"])]
    return [(res.get(k) or [None, None])[0] for k in tk], [(res.get(k) or [None, None])[1] for k in tk]


def fallback_weather(W, batch=50, pause=6.0, cap=3000):
    """TIER-3 FALLBACK: only tehsils where IMD gave no value get Open-Meteo observed/modelled rain + temperature."""
    for c in ("Rainfall mm", "Rainfall 7D", "Rainfall 30D", "Dry Days", "Temperature"):
        if c not in W:
            W[c] = np.nan
    if "Data Date" not in W:
        W["Data Date"] = None
    W["_src"] = np.where(W["Rainfall 30D"].notna(), "IMD", "")
    miss = list(W.index[W["Rainfall 30D"].isna()])
    if not miss:
        return W
    print(f"IMD had no value for {len(miss):,} tehsils -> Open-Meteo fallback (cap {cap:,} per run)")
    for i in range(0, min(len(miss), cap), batch):
        idx = miss[i:i + batch]
        ch = T.loc[idx]
        d = get_json("https://api.open-meteo.com/v1/forecast", {
            "latitude": ",".join(str(round(float(x), 3)) for x in ch["Latitude"]),
            "longitude": ",".join(str(round(float(x), 3)) for x in ch["Longitude"]),
            "daily": "precipitation_sum,temperature_2m_max,temperature_2m_min",
            "past_days": 30, "forecast_days": 1, "timezone": "Asia/Kolkata"}, tries=6)
        locs = d if isinstance(d, list) else ([d] if d else [])
        for k, loc in zip(idx, locs):
            dd = loc.get("daily", {})
            p = [x or 0 for x in dd.get("precipitation_sum", [])]
            if len(p) < 31:
                continue
            h = p[-31:-1]                                   # the 30 complete days before today
            W.loc[k, ["Rainfall mm", "Rainfall 7D", "Rainfall 30D", "Dry Days"]] = [
                round(h[-1], 1), round(sum(h[-7:]), 1), round(sum(h), 1),
                sum(1 for x in h[-7:] if x < DRY_DAY_MM)]
            tx, tn = dd["temperature_2m_max"][-2], dd["temperature_2m_min"][-2]
            if tx is not None and tn is not None:
                W.loc[k, "Temperature"] = round((tx + tn) / 2, 1)
            W.loc[k, "Data Date"] = TODAY - timedelta(days=1)
            W.loc[k, "_src"] = "Open-Meteo"
        print(f"  fallback {min(i + batch, len(miss), cap):,}/{min(len(miss), cap):,}")
        time.sleep(pause)
    FRESH_RUN["Open-Meteo"] = TODAY
    return W


W_CACHE = f"imd_W_{TODAY}.pkl"
if os.path.exists(W_CACHE) and not FORCE_REFRESH_WEATHER:
    _obj = pd.read_pickle(W_CACHE)
    W, _m, _f = _obj[:3]
    if len(_obj) > 3:
        VAL_STORE.update(_obj[3])
    META.update(_m)
    FRESH_RUN.update(_f)
    print("IMD weather loaded from today's saved result (set FORCE_REFRESH_WEATHER = True to recompute)")
    rlog("CELL 7", "IMD gridded (cached today)", "https://www.imdpune.gov.in/", "Rainfall / temperature", "OK",
         int(W["Rainfall 30D"].notna().sum()) if "Rainfall 30D" in W else 0)
else:
    try:
        W = imd_weather()
        pd.to_pickle((W, dict(META), {k: v for k, v in FRESH_RUN.items() if k == "IMD gridded"},
                      {k: v for k, v in VAL_STORE.items() if k.startswith("imd_")}), W_CACHE)
        rlog("CELL 7", "IMD gridded (imdlib)", "https://www.imdpune.gov.in/", "Rainfall / temperature", "OK",
             int(W["Rainfall 30D"].notna().sum()))
    except Exception as e:
        print("IMD step failed:", type(e).__name__, e,
              "\n  (IMD servers can be slow/down - re-run this cell later.)")
        rlog("CELL 7", "IMD gridded (imdlib)", "https://www.imdpune.gov.in/", "Rainfall / temperature", "FAILED",
             fallback="Open-Meteo for all tehsils", msg=f"{type(e).__name__}: {e}")
        W = pd.DataFrame(index=range(len(T)))
W = fallback_weather(W)
if "_src" in W and (W["_src"] == "Open-Meteo").any():
    rlog("CELL 7", "Open-Meteo (fallback)", "https://api.open-meteo.com/v1/forecast", "Rainfall where IMD had no value",
         "OK", int((W["_src"] == "Open-Meteo").sum()), fallback="Tier 3")
f7, hum = forecast_7d()
rlog("CELL 7", "Open-Meteo forecast", "https://api.open-meteo.com/v1/forecast", "7-day forecast (district centre)",
     "OK" if all(v is not None for v in f7) else "PARTIAL", sum(v is not None for v in f7))
W["Forecast 7D"], W["Humidity"] = f7[:len(T)], hum[:len(T)]
W["Extreme Weather Alert"] = np.where(pd.Series(f7[:len(T)], dtype=float) > HEAVY_RAIN_7D_MM,
                                      "Heavy rain watch", "None")
W["Source"] = np.where(W["_src"] == "IMD", "IMD gridded (govt: rain 0.25deg, temp 1deg) + Open-Meteo forecast",
                       np.where(W["_src"] == "Open-Meteo", "Open-Meteo FALLBACK (IMD had no value) + Open-Meteo forecast",
                                "Open-Meteo forecast only (no IMD / fallback value)"))
W = W.drop(columns=["_src"])
W["Date"] = TODAY
if "Data Date" not in W:
    W["Data Date"] = None
for c in ("State", "District", "Tehsil"):
    W[c] = T[c]
install("WEATHER_LIVE", W)
NOTES[("WEATHER_LIVE", "Forecast 7D")] = "Total forecast rainfall (mm), next 7 days, at the DISTRICT centre - Open-Meteo model, not IMD."
print("Weather rows with IMD rainfall:", int(W.get("Rainfall 30D", pd.Series(dtype=float)).notna().sum()))


# ===== CELL 8 : OPTIONAL - MANDI PRICES (Agmarknet via data.gov.in) =========
# Free key: register at data.gov.in -> My Account -> API key. The portal is
# sometimes slow/blocked; re-run later if it returns nothing.
AG_RESOURCE = "https://api.data.gov.in/resource/9ef84268-d588-465a-a308-a864a43d0070"


def fetch_agmarknet(key):
    rows = []
    for crop, commodity in CROP_COMMODITY.items():
        offset = 0
        while True:
            d = get_json(AG_RESOURCE, {"api-key": key, "format": "json", "limit": 1000,
                                       "offset": offset, "filters[commodity]": commodity})
            recs = (d or {}).get("records", [])
            for r in recs:
                r["Crop"] = crop
            rows += recs
            if len(recs) < 1000:
                break
            offset += 1000
        print(f"  {crop}: {sum(1 for r in rows if r['Crop'] == crop):,} mandi records")
    return pd.DataFrame(rows)


def market_rows(raw):
    if raw.empty:
        return pd.DataFrame(columns=HEADERS["MARKET_ECONOMICS"])
    for c in ("min_price", "modal_price", "max_price"):
        raw[c] = pd.to_numeric(raw[c], errors="coerce")
    raw["d"] = pd.to_datetime(raw["arrival_date"], format="%d/%m/%Y", errors="coerce")
    g = raw.groupby(["state", "district", "Crop"]).agg(
        **{"Min Price": ("min_price", "mean"), "Modal Price": ("modal_price", "mean"),
           "Max Price": ("max_price", "mean"), "Date": ("d", "max")}).reset_index()
    out = pd.DataFrame({"Date": g["Date"].dt.date, "State": g["state"].map(clean_name),
                        "District": g["district"].map(clean_name),
                        "Tehsil": "(district mandis)", "Crop": g["Crop"],
                        "Min Price": g["Min Price"].round(0), "Modal Price": g["Modal Price"].round(0),
                        "Max Price": g["Max Price"].round(0),
                        "MSP": g["Crop"].map(MSP), "Source": "Agmarknet via data.gov.in"})
    FRESH_RUN["Agmarknet"] = out["Date"].max()
    return out


AG_KEY = DATA_GOV_IN_KEY or _secret("DATA_GOV_IN_KEY")
if AG_KEY:
    DATA["MARKET_ECONOMICS"] = market_rows(fetch_agmarknet(AG_KEY)).reindex(
        columns=HEADERS["MARKET_ECONOMICS"])
    NOTES[("MARKET_ECONOMICS", "Modal Price")] = "Rs/quintal, average of mandi quotes in the district (latest arrival date)."
    print("Market rows:", len(DATA["MARKET_ECONOMICS"]))
    rlog("CELL 8", "Agmarknet via data.gov.in", AG_RESOURCE, "Mandi prices", "OK" if len(DATA["MARKET_ECONOMICS"]) else "NO RESULT",
         len(DATA["MARKET_ECONOMICS"]))
else:
    rlog("CELL 8", "Agmarknet via data.gov.in", AG_RESOURCE, "Mandi prices", "SKIPPED", fallback="India Data Portal copy (CELL 9)",
         msg="no DATA_GOV_IN_KEY")
    print("No data.gov.in key - mandi prices are read from the India Data Portal copy of Agmarknet in CELL 9 instead (no key needed). "
          "MARKET_ECONOMICS stays empty; PIB/Google news signals still run in CELL 9.")


# ===== CELL 9 : GOVERNMENT CROP STATISTICS - DES "APY" (district area / production / yield) ===
# Ministry of Agriculture & Farmers Welfare (Directorate of Economics & Statistics): district,
# season and crop-wise area, production and yield (1997-98 onward), served openly by the India
# Data Portal. TIER 1 (official statistics) - DISTRICT level only; tehsil figures are not published.
CKAN_HOSTS = ["https://ckandev.indiadataportal.com", "https://indiadataportal.com"]
APY_RESOURCES = ["f980409d-49a2-42ae-9eb0-182365005c04", "ead44f5f-6471-48ec-a488-4b5894302aaa"]
APY_YEARS = 7            # latest N reported years kept
APY_CROP_REGEX = {"Paddy": r"^rice", "Wheat": r"^wheat", "Maize": r"^maize", "Groundnut": r"^groundnut",
                  "Cotton": r"^cotton", "Soyabean": r"^soya?bean|^soybean", "Bajra": r"^bajra",
                  "Sugarcane": r"^sugarcane", "Black gram": r"^urad", "Onion": r"^onion", "Potato": r"^potato"}


def ckan(host, action, params):
    d = get_json(f"{host}/api/3/action/{action}", params, tries=3, timeout=180)
    return d["result"] if d and d.get("success") else None


APY_PACKAGE = "area-production-yield-apy"


def fcol(fields, *keys):
    return next((f for f in fields if all(k in norm(f) for k in keys)), None)


def apy_resources():
    """All datastore resources of the APY dataset (it is split into several files, e.g. by year range)."""
    out = []
    for host in CKAN_HOSTS:
        pkg = ckan(host, "package_show", {"id": APY_PACKAGE})
        if pkg:
            out = [(host, r["id"], r.get("name", "")) for r in pkg.get("resources", [])
                   if r.get("datastore_active", True)]
        if out:
            return out
    for host in CKAN_HOSTS:                                    # fallback: ids found earlier
        out = [(host, rid, "") for rid in APY_RESOURCES
               if ckan(host, "datastore_search", {"resource_id": rid, "limit": 1})]
        if out:
            return out
    return []


def apy_download():
    res = apy_resources()
    if not res:
        print("APY unavailable (India Data Portal not reachable or datastore off) - skipped.")
        return None
    print(f"APY: {len(res)} resource(s):", [(n[:40] or rid[:8]) for _, rid, n in res])
    meta = []
    for host, rid, name in res:
        probe = ckan(host, "datastore_search", {"resource_id": rid, "limit": 1})
        if not probe or not probe.get("fields"):
            continue
        fields = [f["id"] for f in probe["fields"]]
        c_crop = fcol(fields, "crop", "name") or fcol(fields, "crop")
        c_year = fcol(fields, "year")
        if not (c_crop and c_year):
            continue
        dist = lambda col: [r[col] for r in (ckan(host, "datastore_search", {
            "resource_id": rid, "fields": col, "distinct": "true", "limit": 2000}) or {}).get("records", [])]
        years = dist(c_year)
        crops = [c for c in dist(c_crop) if any(re.search(rx, str(c).lower()) for rx in APY_CROP_REGEX.values())]
        meta.append(dict(host=host, rid=rid, fields=fields, c_crop=c_crop, c_year=c_year, years=years, crops=crops))
    if meta:
        META["apy_url"] = f"{meta[0]['host']}/api/3/action/datastore_search?resource_id={meta[0]['rid']}"
    all_years = sorted({y for m in meta for y in m["years"]})
    use_years = all_years[-APY_YEARS:]
    print(f"  years available: {all_years[:1]} ... {all_years[-1:]} | using: {use_years}")
    frames = []
    for m in meta:
        yrs = [y for y in m["years"] if y in use_years]
        if not yrs or not m["crops"]:
            continue
        rows, offset = [], 0
        while True:
            r = ckan(m["host"], "datastore_search", {"resource_id": m["rid"], "limit": 30000, "offset": offset,
                                                      "filters": json.dumps({m["c_crop"]: m["crops"], m["c_year"]: yrs})})
            recs = (r or {}).get("records", [])
            rows += recs
            if len(recs) < 30000:
                break
            offset += 30000
        print(f"  resource {m['rid'][:8]}: {len(rows):,} records for {yrs}")
        if rows:
            frames.append((pd.DataFrame(rows), m["fields"]))
    return frames


def apy_clean(raw, fields):
    c = {k: fcol(fields, *v) for k, v in {"state": ("state", "name"), "dist": ("district", "name"),
         "crop": ("crop", "name"), "year": ("year",), "season": ("season",), "au": ("area", "unit"),
         "pu": ("production", "unit")}.items()}
    c["area"] = next((f for f in fields if norm(f) == "area"), fcol(fields, "area"))
    c["prod"] = next((f for f in fields if norm(f) == "production"), fcol(fields, "production"))
    df = pd.DataFrame({"State": raw[c["state"]].map(clean_name), "District": raw[c["dist"]].map(clean_name),
                       "cropraw": raw[c["crop"]].astype(str), "Year": raw[c["year"]].astype(str),
                       "Season": raw[c["season"]].astype(str).str.strip() if c["season"] else "All",
                       "area": pd.to_numeric(raw[c["area"]], errors="coerce"),
                       "prod": pd.to_numeric(raw[c["prod"]], errors="coerce")})
    if c["au"]:
        df["area"] *= raw[c["au"]].astype(str).str.lower().str.contains("thousand|'000|000").map({True: 1000, False: 1})
    if c["pu"]:
        pu = raw[c["pu"]].astype(str).str.lower()
        keep = pu.str.contains("tonne|ton")
        df["prod"] *= pu.str.contains("thousand|'000|000").map({True: 1000, False: 1})
        print(f"  dropped {int((~keep).sum()):,} rows whose production is not in tonnes (bales, nuts ...)")
        df = df[keep.values]
    df["Crop"] = df["cropraw"].map(lambda x: next((k for k, rx in APY_CROP_REGEX.items()
                                                  if re.search(rx, x.lower())), None))
    return df.dropna(subset=["Crop", "area", "prod"])


def apy_tables(frames):
    df = pd.concat([apy_clean(raw, fields) for raw, fields in frames], ignore_index=True)
    g = df.groupby(["State", "District", "Crop", "Year"], as_index=False).agg(area=("area", "sum"), prod=("prod", "sum"))
    g = g[g["area"] > 0]
    g["yld"] = g["prod"] / g["area"] * 1000                      # kg per hectare
    gs = df.groupby(["State", "District", "Crop", "Season", "Year"], as_index=False).agg(area=("area", "sum"))
    return g, gs


import difflib

_ALIAS = {"orissa": "odisha", "uttaranchal": "uttarakhand", "pondicherry": "puducherry", "nctofdelhi": "delhi"}


def _n(s):
    t = re.sub(r"[^a-z0-9]", "", str(s).lower().replace("&", "and"))
    return _ALIAS.get(t, t)


NAME_MAP = {}     # (state, district) as written by a source -> (LGD State, LGD District) or None


def match_district_names(pairs):
    """Map source (state, district) spellings to the LGD names used in this workbook (exact, then fuzzy within the state)."""
    idx = {}
    for st, di in T[["State", "District"]].drop_duplicates().itertuples(index=False):
        idx.setdefault(_n(st), {"name": st, "d": {}})["d"][_n(di)] = di
    skeys = list(idx)
    for st, di in set(pairs):
        if (st, di) in NAME_MAP:
            continue
        ns_ = _n(st)
        sk = ns_ if ns_ in idx else (difflib.get_close_matches(ns_, skeys, 1, 0.75) or [None])[0]
        if sk is None:
            NAME_MAP[(st, di)] = None
            continue
        dd = idx[sk]["d"]
        nd = _n(di)
        dk = nd if nd in dd else (difflib.get_close_matches(nd, list(dd), 1, 0.8) or [None])[0]
        NAME_MAP[(st, di)] = (idx[sk]["name"], dd[dk]) if dk else None
    return NAME_MAP


def tehsil_shares():
    """Each tehsil's share of its district's cropland - used to split district totals (estimates).
    Tehsils with no cropland figure get the district's typical cropland ratio x their area (not raw area)."""
    M = DATA["TEHSIL_MASTER"]
    net = M["Net Sown Area Ha"].astype(float).where(lambda x: x > 0)
    tot = M["Total Area Ha"].astype(float)
    d = pd.DataFrame({"State": T["State"].values, "District": T["District"].values, "Tehsil": T["Tehsil"].values,
                      "net": net.values, "tot": tot.values})
    ratio = (d["net"] / d["tot"]).groupby([d["State"], d["District"]]).transform("median")
    ratio = ratio.fillna((d["net"] / d["tot"]).median()).fillna(0.3)
    d["w"] = d["net"].where(d["net"].notna(), ratio * d["tot"])
    d["share"] = d["w"] / d.groupby(["State", "District"])["w"].transform("sum")
    d["pos"] = np.arange(len(d))
    return d


def season_key(s):
    t = str(s).lower()
    return next((k for k in ("kharif", "rabi", "summer", "autumn", "winter") if k in t), "annual")


def ideal_window(crop, season):
    return IDEAL_SOWING.get(crop, {}).get(season_key(season))


def apy_crop_sown(gs):
    """CROP_SOWN baseline from DES APY: latest REPORTED (final) area per district/crop/season, split to tehsils by cropland
    share (ESTIMATE). Sowing % = that area as % of the mean of up to 5 earlier years (area vs normal). Not live sowing progress."""
    gs = gs[gs["area"] > 0].sort_values("Year")
    rec = []
    for (st, di, cr, se), grp in gs.groupby(["State", "District", "Crop", "Season"], sort=False):
        a = grp["area"].tolist()
        base = a[:-1][-5:]
        pct = round(a[-1] / float(np.mean(base)) * 100, 1) if len(base) >= 3 and np.mean(base) > 0 else None
        rec.append((st, di, cr, se, grp["Year"].iloc[-1], a[-1], pct))
    d = pd.DataFrame(rec, columns=["S0", "D0", "Crop", "Season", "Year", "area", "pct"])
    mp = match_district_names(list(zip(d["S0"], d["D0"])))
    d["lg"] = [mp.get(k) for k in zip(d["S0"], d["D0"])]
    ok, no = d[d["lg"].notna()].copy(), d[d["lg"].isna()].copy()
    print(f"  districts matched to LGD names: {ok[['S0','D0']].drop_duplicates().shape[0]:,} matched, "
          f"{no[['S0','D0']].drop_duplicates().shape[0]:,} kept at district level")
    base_note = "DES APY final reported area (district)"
    outs = []
    if len(ok):
        ok["State"] = ok["lg"].map(lambda x: x[0])
        ok["District"] = ok["lg"].map(lambda x: x[1])
        ok = ok.groupby(["State", "District", "Crop", "Season", "Year"], as_index=False).agg(area=("area", "sum"), pct=("pct", "mean"))
        m = ok.merge(tehsil_shares()[["State", "District", "Tehsil", "share"]], on=["State", "District"], how="inner")
        outs.append(pd.DataFrame({
            "Season": m["Season"].astype(str) + " " + m["Year"], "State": m["State"], "District": m["District"],
            "Tehsil": m["Tehsil"], "Crop": m["Crop"], "Sown Area Ha": (m["area"] * m["share"]).round(1),
            "Sowing %": m["pct"].round(1),
            "Ideal Sowing Window": [ideal_window(c, s_) for c, s_ in zip(m["Crop"], m["Season"])],
            "Crop Source": base_note + ", split to tehsil by cropland share (ESTIMATE)",
            "Report Date": m["Year"].map(fy_end), "Confidence": 0.5}))
    if len(no):
        outs.append(pd.DataFrame({
            "Season": no["Season"].astype(str) + " " + no["Year"], "State": no["S0"], "District": no["D0"],
            "Tehsil": "(district level)", "Crop": no["Crop"], "Sown Area Ha": no["area"].round(0), "Sowing %": no["pct"],
            "Ideal Sowing Window": [ideal_window(c, s_) for c, s_ in zip(no["Crop"], no["Season"])],
            "Crop Source": base_note, "Report Date": no["Year"].map(fy_end), "Confidence": 0.9}))
    return pd.concat(outs, ignore_index=True)


LANDUSE_PACKAGE = "land-use-statistics"
IRR_RESOURCE_HINT = "ef174105-0886-45f2-bd0d-679c45e05845"      # 'Sources of Irrigation' (DES Land Use Statistics)
IRR_MIN_ROWS = 300                                                # a usable year has at least this many districts


def irrigation_resource():
    for host in CKAN_HOSTS:
        pkg = ckan(host, "package_show", {"id": LANDUSE_PACKAGE})
        res = [r for r in (pkg or {}).get("resources", []) if r.get("datastore_active", True)]
        hit = next((r for r in res if r["id"] == IRR_RESOURCE_HINT), None) or next(
            (r for r in res if "irrigation" in r.get("name", "").lower() and "source" in r.get("name", "").lower()), None)
        if hit:
            return host, hit["id"]
    for host in CKAN_HOSTS:
        if ckan(host, "datastore_search", {"resource_id": IRR_RESOURCE_HINT, "limit": 1}):
            return host, IRR_RESOURCE_HINT
    return None


def ckan_all(host, rid, filters):
    out, off = [], 0
    while True:
        r = ckan(host, "datastore_search", {"resource_id": rid, "limit": 30000, "offset": off,
                                            "filters": json.dumps(filters)})
        recs = (r or {}).get("records", [])
        out += recs
        if len(recs) < 30000:
            return out
        off += 30000


IRR_MIN_YEAR = 2010        # older district tables are ignored (the portal copy may stop in the early 2000s)


def irrigation_by_source():
    """TIER 1: DES Land Use Statistics - area irrigated by source per district (long format: type / classification / value)."""
    got = irrigation_resource()
    if not got:
        print("Irrigation by source (DES Land Use Statistics) unavailable - CoRE proxy kept in WATER_IRRIGATION.")
        return
    host, rid = got
    META["irr_url"] = f"{host}/api/3/action/datastore_search?resource_id={rid}"
    probe = ckan(host, "datastore_search", {"resource_id": rid, "limit": 1}) or {}
    fields = [f["id"] for f in probe.get("fields", [])]
    c_year = fcol(fields, "year")
    c_state = fcol(fields, "state", "name") or fcol(fields, "state")
    c_dist = fcol(fields, "district", "name") or fcol(fields, "district")
    c_type, c_cls = fcol(fields, "irrigated", "type"), fcol(fields, "area", "classification")
    c_sub, c_val = fcol(fields, "sub", "classification"), next((f for f in fields if norm(f) == "irrigatedarea"), None)
    if not (c_year and c_state and c_dist and c_cls and c_val):
        print("  unexpected irrigation table layout - skipped. Fields:", fields)
        return
    years = sorted((r[c_year] for r in (ckan(host, "datastore_search", {
        "resource_id": rid, "fields": c_year, "distinct": "true", "limit": 300}) or {}).get("records", [])), reverse=True)
    print("  irrigation table years (latest first):", years[:6])
    year = next((y for y in years if int(str(y)[:4]) >= IRR_MIN_YEAR), None)
    if year is None:
        print(f"  newest year in this portal table is {years[0] if years else 'n/a'} (< {IRR_MIN_YEAR}) - too old to use; CoRE proxy kept.")
        return
    df = pd.DataFrame(ckan_all(host, rid, {c_year: [year]}))
    if len(df) < IRR_MIN_ROWS:
        print(f"  only {len(df)} rows for {year} - skipped")
        return
    df["v"] = pd.to_numeric(df[c_val], errors="coerce")
    typ = df[c_type].astype(str).str.lower() if c_type else pd.Series("net", index=df.index)
    if (typ.str.contains("net")).any():
        df = df[typ.str.contains("net")]
    cls = df[c_cls].astype(str).str.lower()
    sub = df[c_sub].astype(str).str.lower() if c_sub else pd.Series("", index=df.index)
    df = df.assign(cl=cls, sb=sub, S0=df[c_state].map(clean_name), D0=df[c_dist].map(clean_name))

    def bucket(g):
        out = {}
        for key, kw in (("canal", "canal"), ("tank", "tank"), ("well", "well"), ("other", "other")):
            x = g[g["cl"].str.contains(kw)]
            tot = x[x["sb"].str.contains("total")]
            out[key] = (tot["v"] if len(tot) else x["v"]).sum(min_count=1)
        t = g[g["cl"].str.contains(r"total|\ball\b|net irrigated")]
        out["net"] = t["v"].max() if len(t) else np.nan
        return pd.Series(out)
    d = df.groupby(["S0", "D0"]).apply(bucket, include_groups=False).reset_index()
    d["net"] = d["net"].where(d["net"].notna(), d[["canal", "tank", "well", "other"]].sum(axis=1, min_count=1))
    scale = 1000 if d["net"].median() < 3000 else 1
    print(f"  year {year} | {len(d):,} districts | units assumed: {'thousand hectares (x1000)' if scale == 1000 else 'hectares'} "
          f"(median net irrigated {d['net'].median():,.1f})")
    mp = match_district_names(list(zip(d["S0"], d["D0"])))
    d["lg"] = [mp.get(k) for k in zip(d["S0"], d["D0"])]
    ok = d[d["lg"].notna()].copy()
    ok["State"], ok["District"] = ok["lg"].map(lambda x: x[0]), ok["lg"].map(lambda x: x[1])
    agg = ok.groupby(["State", "District"], as_index=False)[["canal", "tank", "net"]].sum(min_count=1)
    sh = tehsil_shares().merge(agg, on=["State", "District"], how="left").sort_values("pos").reset_index(drop=True)
    for c in ("canal", "tank", "net"):
        sh[c] = sh[c] * scale * sh["share"]
    has = sh["net"].notna().to_numpy()
    VAL_STORE["des_irr"] = sh["net"].to_numpy(float)
    W, M = DATA["WATER_IRRIGATION"], DATA["TEHSIL_MASTER"]

    def put(col, vals):
        cur = W[col].to_numpy(dtype=object)
        W[col] = np.where(has, vals.to_numpy(dtype=float).round(0), cur)

    put("Irrigated Area Ha", sh["net"])
    put("Canal", sh["canal"])
    put("Tank", sh["tank"])
    put("Rainfed Area Ha", (M["Net Sown Area Ha"].astype(float).reset_index(drop=True) - sh["net"]).clip(lower=0))
    W["Data Date"] = np.where(has, fy_end(year), W["Data Date"].to_numpy(dtype=object) if "Data Date" in W else None)
    W["_irr_src"] = np.where(has, "DES", W["_irr_src"].to_numpy(dtype=object) if "_irr_src" in W else "")
    META["irr_year"] = str(year)
    txt = f"DES Land Use Statistics {year}: district net area irrigated split to tehsils by cropland share (ESTIMATE). "
    NOTES[("WATER_IRRIGATION", "Canal")] = txt + "Canal = net area irrigated by canals (ha)."
    NOTES[("WATER_IRRIGATION", "Tank")] = txt + "Tank = net area irrigated by tanks (ha)."
    NOTES[("WATER_IRRIGATION", "Irrigated Area Ha")] = txt + "Where the district did not match, a CoRE Stack cropping-pattern proxy is used instead."
    print(f"Irrigation by source applied to {int(has.sum()):,} of {len(sh):,} tehsils (rest keep the CoRE proxy)")


def fy_end(label):
    m = re.match(r"(\d{4})", str(label))
    return date(int(m.group(1)) + 1, 3, 31) if m else None


def apy_to_sheets(g):
    yp, hist = [], []
    for (st, di, cr), grp in g.groupby(["State", "District", "Crop"]):
        grp = grp.sort_values("Year")
        yrs = list(grp["Year"])
        latest = grp.iloc[-1]
        base = grp.iloc[:-1].tail(5)
        h_avg = base["yld"].mean() if len(base) >= 3 else np.nan
        yp.append({"Season": f"{latest['Year']} (all seasons)", "State": st, "District": di,
                   "Tehsil": "(district level)", "Crop": cr, "Area Sown Ha": round(latest["area"], 0),
                   "Historical Yield": round(h_avg, 0) if h_avg == h_avg else None,
                   "Production Estimate": round(latest["prod"], 0),
                   "Yield vs 5Y Avg %": round((latest["yld"] / h_avg - 1) * 100, 1) if h_avg == h_avg else None,
                   "Yield Trend": (None if h_avg != h_avg else "Up" if latest["yld"] > h_avg * 1.03
                                   else "Down" if latest["yld"] < h_avg * 0.97 else "Flat"),
                   "Data Date": fy_end(latest["Year"]),
                   "Source": f"DES APY via India Data Portal - district, reported {latest['Year']}"})
        for param, col in (("Yield kg/ha", "yld"), ("Area ha", "area"), ("Production t", "prod")):
            vals = grp[col].tolist()
            b = vals[-6:-1] if len(vals) > 5 else vals[:-1]
            avg, mn, mx = (np.mean(b), min(b), max(b)) if b else (np.nan, np.nan, np.nan)
            for k, (y, v) in enumerate(zip(yrs, vals)):
                prev = vals[k - 1] if k else np.nan
                yoy = (v / prev - 1) * 100 if prev == prev and prev else np.nan
                hist.append({"Date": fy_end(y), "Year": y, "State": st, "District": di, "Tehsil": "(district level)",
                             "Crop": cr, "Parameter": param, "Value": round(v, 1),
                             "5Y Average": round(avg, 1) if avg == avg else None,
                             "5Y Min": round(mn, 1) if mn == mn else None, "5Y Max": round(mx, 1) if mx == mx else None,
                             "Current vs 5Y %": round((v / avg - 1) * 100, 1) if avg == avg and avg else None,
                             "YoY %": round(yoy, 1) if yoy == yoy else None,
                             "Trend": None if yoy != yoy else "Up" if yoy > 2 else "Down" if yoy < -2 else "Flat",
                             "Source": "DES APY via India Data Portal (district)"})
    return pd.DataFrame(yp), pd.DataFrame(hist)


STAGES = [("Germination", 0.00, 0.10), ("Vegetative", 0.10, 0.45), ("Flowering", 0.45, 0.65),
          ("Grain/Fruit Formation", 0.65, 0.85), ("Maturity", 0.85, 1.00)]
_MON = {m: i for i, m in enumerate(["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], 1)}


def _window(txt):
    m = re.match(r"\s*(\d{1,2}) ([A-Za-z]{3}) - (\d{1,2}) ([A-Za-z]{3})", str(txt or ""))
    if not m:
        return None
    y = TODAY.year
    a, b = date(y, _MON[m.group(2).title()], int(m.group(1))), date(y, _MON[m.group(4).title()], int(m.group(3)))
    return (a, b) if b >= a else (a, date(y + 1, b.month, b.day))


def _fmt(d):
    return f"{d.day:02d} {d.strftime('%b')}"


def build_crop_calendar():
    """CROP_CALENDAR from REFERENCE DEFAULTS (typical national sowing windows + typical durations; edit in CELL 2).
    One row per state x crop x season found in CROP_SOWN. Stage windows = fractions of the typical duration."""
    cs = DATA["CROP_SOWN"]
    if cs is None or cs.empty:
        return
    u = cs[["State", "Crop", "Season"]].dropna().copy()
    u["sk"] = u["Season"].map(season_key)
    u = u.drop_duplicates(["State", "Crop", "sk"])
    rows = []
    for st, cr, sk in zip(u["State"], u["Crop"], u["sk"]):
        win, dur = IDEAL_SOWING.get(cr, {}).get(sk), CROP_DURATION.get(cr)
        w = _window(win)
        rec = {"State": st, "District": "(all districts)", "Tehsil": "(state default)", "Crop": cr, "Season": sk.title(),
               "Typical Duration": f"{dur} days" if dur else None,
               "Source": "Reference default (typical national calendar) - NOT state/tehsil specific; edit IDEAL_SOWING / CROP_DURATION in CELL 2"}
        if w and dur:
            s0, s1 = w
            rec["Ideal Sowing Start"], rec["Ideal Sowing End"] = _fmt(s0), _fmt(s1)
            for name, f0, f1 in STAGES:
                rec[name] = f"{_fmt(s0 + timedelta(days=round(f0 * dur)))} - {_fmt(s0 + timedelta(days=round(f1 * dur)))}"
            rec["Harvest Window"] = f"{_fmt(s0 + timedelta(days=dur))} - {_fmt(s1 + timedelta(days=dur + 14))}"
        elif win:
            rec["Ideal Sowing Start"] = win
        rows.append(rec)
    DATA["CROP_CALENDAR"] = pd.DataFrame(rows).reindex(columns=HEADERS["CROP_CALENDAR"])
    NOTES[("CROP_CALENDAR", "Germination")] = "Fractions of the typical crop duration counted from the start of the ideal sowing window (0-10% / 10-45% / 45-65% / 65-85% / 85-100%)."
    print(f"CROP_CALENDAR: {len(rows):,} state-crop-season rows (reference defaults)")


AGC_DISTRICT_HINT = "c0f168be-3532-4d08-b908-473ded89bd8b"      # 'District level AgCensus Non Crop' (India Data Portal)


def _tn(s):
    t = re.sub(r"\[.*?\]", "", str(s)).lower()
    t = re.sub(r"\b(tehsil|tahsil|taluka|taluk|mandal|block|subdistrict|sub district|circle|sub-division|subdivision)\b", "", t)
    return re.sub(r"[^a-z0-9]", "", t.replace("&", "and"))


def match_tehsils(keys):
    """(state, district, tehsil) as written by the census -> LGD names (district exact/fuzzy, then tehsil fuzzy inside it)."""
    mp = match_district_names(list({(a, b) for a, b, _ in keys}))
    idx = {}
    for st, di, te in T[["State", "District", "Tehsil"]].itertuples(index=False):
        idx.setdefault((st, di), {})[_tn(te)] = te
    out = {}
    for a, b, c in set(keys):
        lg = mp.get((a, b))
        cand = idx.get(lg, {}) if lg else {}
        n = _tn(c)
        k = n if n in cand else (difflib.get_close_matches(n, list(cand), 1, 0.8) or [None])[0] if cand else None
        out[(a, b, c)] = (lg[0], lg[1], cand[k]) if (lg and k) else None
    return out


def agcensus_tables():
    host = CKAN_HOSTS[0]
    found = {}
    res = ckan(host, "package_search", {"q": "agricultural census non crop", "rows": 25}) or {}
    for pk in res.get("results", []):
        for r in pk.get("resources", []):
            nm = (r.get("name") or "").lower()
            if ("non crop" in nm or "non-crop" in nm) and r.get("datastore_active", True):
                for lvl in ("tehsil", "district"):
                    if lvl in nm:
                        found.setdefault(lvl, (host, r["id"]))
    found.setdefault("district", (host, AGC_DISTRICT_HINT))
    print("  Agriculture Census (non-crop) tables found:", {k: v[1][:8] for k, v in found.items()})
    return found


def agc_level(host, rid, level):
    """Read one Agriculture Census non-crop table -> holdings, area, small+marginal share (+ irrigated area by source if present)."""
    probe = ckan(host, "datastore_search", {"resource_id": rid, "limit": 2}) or {}
    fields = [f["id"] for f in probe.get("fields", [])]
    print(f"  AgCensus {level} fields:", fields)
    META[f"agc_url_{level}"] = f"{host}/api/3/action/datastore_search?resource_id={rid}"
    pick1 = lambda *names: next((n for n in names if n in fields), None)
    c_h = pick1("hol_no", "holding_num") or next((f for f in fields if "holding" in norm(f) and "num" in norm(f)), None)
    c_a = pick1("hol_ar", "holding_area") or next((f for f in fields if "holding" in norm(f) and "area" in norm(f)), None)
    c_soc = next((f for f in fields if "social" in norm(f)), None)
    cls_cands = list(dict.fromkeys([f for f in ("farm_size_category", "farm_size_class") if f in fields] +
                                   [f for f in fields if "size" in norm(f)]))
    c_y = fcol(fields, "year")
    c_s = fcol(fields, "state", "name") or fcol(fields, "state")
    c_d = fcol(fields, "district", "name") or fcol(fields, "district")
    c_t = fcol(fields, "tehsil", "name") or fcol(fields, "tehsil") if level == "tehsil" else None
    if not all((c_h, c_a, c_s, c_d)) or (level == "tehsil" and not c_t):
        print("  required fields not found - skipped")
        return None
    dist = lambda col: [r[col] for r in (ckan(host, "datastore_search", {
        "resource_id": rid, "fields": col, "distinct": "true", "limit": 500}) or {}).get("records", [])]
    c_cls, keep_cls = None, []
    for c_ in cls_cands:                                   # use the size column whose VALUES are words (All / Marginal / Small)
        vals = dist(c_)
        print(f"  {c_} values: {vals[:12]}")
        k_ = [v for v in vals if re.search(r"\ball\b|total|marginal|small", str(v).lower()) and "semi" not in str(v).lower()
              or re.search(r"\ball\b|total", str(v).lower())]
        if k_:
            c_cls, keep_cls = c_, k_
            break
    c_cls = c_cls or (cls_cands[0] if cls_cands else None)
    has_all = any(re.search(r"\ball\b|total", str(v).lower()) for v in keep_cls)
    filt = {c_cls: keep_cls} if has_all else {}             # no 'All' class -> read every size class and add them up
    if c_soc:
        soc = [v for v in dist(c_soc) if re.search(r"\ball\b|total", str(v).lower())]
        if soc:
            filt[c_soc] = soc
    if c_y:
        filt[c_y] = [sorted(dist(c_y))[-1]]
    print(f"  filters: {filt}")
    df = pd.DataFrame(ckan_all(host, rid, filt))
    if df.empty:
        print("  no rows after filtering - skipped")
        return None
    num_ = lambda c: pd.to_numeric(df[c], errors="coerce")
    df["h"], df["a"] = num_(c_h), num_(c_a)
    cl = df[c_cls].astype(str).str.lower() if c_cls else pd.Series("", index=df.index)
    df["tot"] = cl.str.contains(r"\ball\b|total")
    if not df["tot"].any():                                  # size classes are disjoint -> their sum is the total
        df["tot"] = True
        print("  no 'All/Total' size class - totals = sum of all size classes")
    df["sm"] = cl.str.contains("marginal|small") & ~cl.str.contains("semi") & ~cl.str.contains(r"\ball\b|total")
    if c_soc and not df[c_soc].astype(str).str.lower().str.contains(r"\ball\b|total").any():
        print("  no 'All' social group - totals = sum of social groups")
    df["S0"], df["D0"] = df[c_s].map(clean_name), df[c_d].map(clean_name)
    keys = ["S0", "D0"]
    if level == "tehsil":
        df["T0"] = df[c_t].map(clean_name)
        keys.append("T0")
    # irrigation / land-use columns of the Agriculture Census non-crop tables (hectares)
    irr = {k: [f for f in names if f in fields] for k, names in
           {"canal": ["canal_ar"], "tank": ["tank_ar"], "net": ["net_irr_ar"], "nsown": ["net_swn_ar"]}.items()}
    irr = {k: v for k, v in irr.items() if v}
    if "net" in irr and not (pd.to_numeric(df[irr["net"][0]], errors="coerce").fillna(0) > 0).any() and {"w_ir_ar", "p_ir_ar"} <= set(fields):
        irr["net"] = ["w_ir_ar", "p_ir_ar"]
        print("  net_irr_ar is empty in this table -> net irrigated = wholly irrigated area (w_ir_ar) + irrigated area of partly irrigated holdings (p_ir_ar)")
    if irr:
        print("  irrigation / land-use columns used:", irr)
    agg = {"holdings": ("h", "sum"), "area": ("a", "sum")}
    t = df[df["tot"]].groupby(keys, as_index=False).agg(**agg)
    m = df[df["sm"]].groupby(keys, as_index=False).agg(sm_h=("h", "sum"))
    out = t.merge(m, on=keys, how="left")
    out["sm_share"] = out["sm_h"] / out["holdings"] * 100
    for k, cols in irr.items():
        if cols:
            tt = df[df["tot"]].copy()
            tt[k] = tt[cols].apply(pd.to_numeric, errors="coerce").sum(axis=1, min_count=1)
            out = out.merge(tt.groupby(keys, as_index=False)[k].sum(min_count=1), on=keys, how="left")
    return out


def agcensus_farmers():
    """TIER 1: Agriculture Census 2015-16 (non-crop): holdings (farmers), small+marginal share, average holding.
    Uses the TEHSIL table where names match; remaining tehsils get the district total split by cropland share (ESTIMATE)."""
    tabs = agcensus_tables()
    base = tehsil_shares()[["State", "District", "Tehsil", "pos", "share"]]
    n = len(base)
    farmers = np.full(n, np.nan)
    smv, avg = np.full(n, np.nan), np.full(n, np.nan)
    lvl = np.full(n, "", dtype=object)
    irr_vals = {k: np.full(n, np.nan) for k in ("canal", "tank", "net", "nsown")}
    irr_lvl = np.full(n, "", dtype=object)

    if "district" in tabs:
        d = agc_level(*tabs["district"], "district")
        if d is not None and len(d):
            scale = 1000 if d["holdings"].median() < 2000 else 1
            mp = match_district_names(list(zip(d["S0"], d["D0"])))
            d["lg"] = [mp.get(k) for k in zip(d["S0"], d["D0"])]
            ok = d[d["lg"].notna()].copy()
            ok["State"], ok["District"] = ok["lg"].map(lambda x: x[0]), ok["lg"].map(lambda x: x[1])
            cols = [c for c in ("canal", "tank", "net", "nsown") if c in ok]
            g = ok.groupby(["State", "District"], as_index=False).agg(
                holdings=("holdings", "sum"), area=("area", "sum"), sm=("sm_share", "mean"), **{c: (c, "sum") for c in cols})
            sh = base.merge(g, on=["State", "District"], how="left").sort_values("pos").reset_index(drop=True)
            has = sh["holdings"].notna().to_numpy()
            farmers = np.where(has, (sh["holdings"] * scale * sh["share"]).to_numpy(), np.nan)
            smv = np.where(has, sh["sm"].to_numpy(), np.nan)
            avg = np.where(has, (sh["area"] / sh["holdings"]).to_numpy(), np.nan)
            lvl = np.where(has, "district", "")
            for c in cols:
                v = (sh[c] * sh["share"]).to_numpy()
                irr_vals[c] = np.where(has, v, np.nan)
                irr_lvl = np.where(has, "district", irr_lvl)
            print(f"  district table applied to {int(has.sum()):,} tehsils (split estimate)")
    if "tehsil" in tabs:
        t = agc_level(*tabs["tehsil"], "tehsil")
        if t is not None and len(t):
            mt = match_tehsils(list(zip(t["S0"], t["D0"], t["T0"])))
            t["lg"] = [mt.get(k) for k in zip(t["S0"], t["D0"], t["T0"])]
            ok = t[t["lg"].notna()].copy()
            ok["State"], ok["District"], ok["Tehsil"] = (ok["lg"].map(lambda x: x[0]), ok["lg"].map(lambda x: x[1]),
                                                         ok["lg"].map(lambda x: x[2]))
            cols = [c for c in ("canal", "tank", "net", "nsown") if c in ok]
            g = ok.groupby(["State", "District", "Tehsil"], as_index=False).agg(
                holdings=("holdings", "sum"), area=("area", "sum"), sm=("sm_share", "mean"), **{c: (c, "sum") for c in cols})
            sh = base.merge(g, on=["State", "District", "Tehsil"], how="left").sort_values("pos").reset_index(drop=True)
            has = sh["holdings"].notna().to_numpy()
            farmers = np.where(has, sh["holdings"].to_numpy(), farmers)
            smv = np.where(has, sh["sm"].to_numpy(), smv)
            avg = np.where(has, (sh["area"] / sh["holdings"]).to_numpy(), avg)
            lvl = np.where(has, "tehsil", lvl)
            for c in cols:
                irr_vals[c] = np.where(has, sh[c].to_numpy(), irr_vals[c])
                irr_lvl = np.where(has, "tehsil", irr_lvl)
            print(f"  tehsil table matched {int(has.sum()):,} of {n:,} LGD tehsils (real tehsil values)")
    got = lvl != ""
    if not got.any():
        print("Agriculture Census not applied - FARMER_AGRI_ECOSYSTEM stays empty.")
        return
    F = DATA["FARMER_AGRI_ECOSYSTEM"]
    VAL_STORE["agc_farmer_lvl"] = np.asarray(lvl, dtype=object)
    F["Farmers"] = np.where(got, np.round(farmers), None)
    F["Small/Marginal Farmers %"] = np.where(got, np.round(smv, 1), None)
    F["Average Holding"] = np.where(got, np.round(avg, 2), None)
    F["Data Year"] = np.where(got, 2016, None)
    F["Source"] = np.where(lvl == "tehsil", "Agriculture Census 2015-16 (tehsil table; names matched to LGD)",
                           np.where(lvl == "district", "Agriculture Census 2015-16 (district total split to tehsil by cropland share = ESTIMATE)", None))
    FRESH_RUN["Agriculture Census"] = date(2016, 3, 31)
    NOTES[("FARMER_AGRI_ECOSYSTEM", "Farmers")] = "Operational holdings (proxy for farmers), Agriculture Census 2015-16: tehsil table where names matched, else district total split by cropland share (ESTIMATE)."
    NOTES[("FARMER_AGRI_ECOSYSTEM", "Average Holding")] = "Hectares per operational holding."
    print(f"FARMER_AGRI_ECOSYSTEM filled for {int(got.sum()):,} of {n:,} tehsils")
    if irr_lvl.size and (irr_lvl != "").any() and not np.isnan(irr_vals["net"]).all():
        W = DATA["WATER_IRRIGATION"]
        hasi = ~np.isnan(irr_vals["net"])
        def put(col, arr):
            cur = W[col].to_numpy(dtype=object)
            W[col] = np.where(hasi, np.round(arr), cur)
        put("Irrigated Area Ha", irr_vals["net"])
        if not np.isnan(irr_vals["nsown"]).all():
            put("Rainfed Area Ha", np.clip(irr_vals["nsown"] - irr_vals["net"], 0, None))
        if not np.isnan(irr_vals["canal"]).all():
            put("Canal", irr_vals["canal"])
        if not np.isnan(irr_vals["tank"]).all():
            put("Tank", irr_vals["tank"])
        W["Data Date"] = np.where(hasi, date(2016, 3, 31), W["Data Date"].to_numpy(dtype=object))
        W["_irr_src"] = np.where(hasi, np.where(irr_lvl == "tehsil", "AGC-T", "AGC-D"),
                                 W["_irr_src"].to_numpy(dtype=object) if "_irr_src" in W else "")
        VAL_STORE["agc_irr"], VAL_STORE["agc_nsown"] = irr_vals["net"].astype(float), irr_vals["nsown"].astype(float)
        META["agc_irrigation"] = True
        M = DATA["TEHSIL_MASTER"]
        if not np.isnan(irr_vals["nsown"]).all():
            cur = M["Net Sown Area Ha"].astype(float).to_numpy()
            M["Net Sown Area Ha"] = np.where(np.isnan(cur), irr_vals["nsown"], cur)
            NOTES[("TEHSIL_MASTER", "Net Sown Area Ha")] = ("CoRE Stack 2023-24 (satellite) cropped area; gaps filled from Agriculture Census 2015-16 net sown area "
                                                            "(tehsil table or district split).")
        print(f"  Irrigated area from Agriculture Census applied to {int(hasi.sum()):,} tehsils")


try:
    got = apy_download()
    if got:
        G, GS = apy_tables(got)
        match_district_names(list(zip(G["State"], G["District"])))          # needed later by the decision layer
        YP, HI = apy_to_sheets(G)
        DATA["YIELD_PRODUCTION"] = YP.reindex(columns=HEADERS["YIELD_PRODUCTION"])
        DATA["5_YEAR_HISTORY"] = HI.reindex(columns=HEADERS["5_YEAR_HISTORY"])
        DATA["CROP_SOWN"] = apy_crop_sown(GS).reindex(columns=HEADERS["CROP_SOWN"])
        META["apy_years"] = f"{G['Year'].min()} to {G['Year'].max()}"
        META["apy_latest"] = str(G["Year"].max())
        FRESH_RUN["DES APY"] = fy_end(G["Year"].max())
        print(f"Latest OFFICIAL APY year published: {G['Year'].max()} (financial year ends {fy_end(G['Year'].max()):%d-%b-%Y}). "
              "Newer years are picked up automatically when the Ministry publishes them.")
        if int(str(G["Year"].max())[:4]) < TODAY.year - 6:
            print(f"WARNING: newest APY year found is {G['Year'].max()} - the portal copy looks outdated; treat as history only.")
        NOTES[("YIELD_PRODUCTION", "Production Estimate")] = "Latest REPORTED production (tonnes), not a forecast. DISTRICT level (DES APY)."
        NOTES[("YIELD_PRODUCTION", "Historical Yield")] = "kg/ha: mean of the up-to-5 reported years before the latest year."
        NOTES[("5_YEAR_HISTORY", "Date")] = (f"Financial-year end of the reported year. The latest official APY year is {G['Year'].max()}; "
                                            "there is no later published data yet.")
        NOTES[("CROP_SOWN", "Sown Area Ha")] = ("Final reported area, latest reported year per season (DES APY). Tehsil rows = district area split by "
                                               "cropland share (ESTIMATE). NOT current-season sowing progress.")
        NOTES[("CROP_SOWN", "Sowing %")] = "Latest reported area as % of the mean of up to 5 earlier years (district level, repeated per tehsil)."
        NOTES[("CROP_SOWN", "Ideal Sowing Window")] = "REFERENCE DEFAULT (typical national window, edit IDEAL_SOWING in CELL 2) - not retrieved data."
        NOTES[("CROP_SOWN", "Confidence")] = "Rule-based: 0.9 = official district value; 0.5 = tehsil share estimated by cropland split."
        print(f"APY: {len(YP):,} district-crop rows, {len(HI):,} history rows, {len(DATA['CROP_SOWN']):,} crop-sown rows, years {META['apy_years']}")
        rlog("CELL 9", "DES APY via India Data Portal", META.get("apy_url"), "District area / production / yield", "OK", len(YP),
             msg=f"years {META['apy_years']}")
    else:
        rlog("CELL 9", "DES APY via India Data Portal", CKAN_HOSTS[0], "District area / production / yield", "NO RESULT")
except Exception as e:
    print("APY step failed:", type(e).__name__, e)
    rlog("CELL 9", "DES APY via India Data Portal", CKAN_HOSTS[0], "District area / production / yield", "FAILED", msg=f"{type(e).__name__}: {e}")

try:
    irrigation_by_source()
    rlog("CELL 9", "DES Land Use Statistics (sources of irrigation)", META.get("irr_url", CKAN_HOSTS[0]), "Irrigated area by source",
         "OK" if META.get("irr_year") else "NO RESULT", fallback="" if META.get("irr_year") else "CoRE Stack proxy kept")
except Exception as e:
    print("Irrigation-by-source step failed:", type(e).__name__, e)
    rlog("CELL 9", "DES Land Use Statistics (sources of irrigation)", CKAN_HOSTS[0], "Irrigated area by source", "FAILED",
         fallback="CoRE Stack proxy kept", msg=f"{type(e).__name__}: {e}")

MKT_REGEX = {"Paddy": r"^paddy|^rice\b", "Wheat": r"^wheat", "Maize": r"^maize", "Groundnut": r"groundnut|ground nut", "Cotton": r"^cotton|^kapas",
             "Soyabean": r"^soya?\s?bean", "Bajra": r"^bajra|pearl millet", "Black gram": r"black gram|urd|urad", "Onion": r"^onion",
             "Potato": r"^potato", "Tur": r"arhar|tur\b|red gram", "Bengal gram": r"bengal gram|gram\b.*whole|^chana", "Mustard": r"mustard|rapeseed",
             "Jowar": r"^jowar|sorghum", "Ragi": r"^ragi|finger millet"}


def market_resources():
    for host in CKAN_HOSTS:
        out = []
        for pkg_name in ("agriculture-marketing", "apmc-arrivals-and-prices"):
            pkg = ckan(host, "package_show", {"id": pkg_name})
            out += [(host, r["id"], r.get("name", "")) for r in (pkg or {}).get("resources", []) if r.get("datastore_active", True)]
        if out:
            return out
    return []


def market_from_portal(window_days=7, past_years=5):
    """TIER 1 (no key): Agmarknet daily mandi prices/arrivals as published on the India Data Portal (Ministry of Agriculture data).
    District average of reporting mandis over the last `window_days`, compared with the same dates in the previous years."""
    best = None
    for host, rid, name in market_resources():
        probe = ckan(host, "datastore_search", {"resource_id": rid, "limit": 1, "sort": "report_date desc"}) or {}
        recs = probe.get("records", [])
        fields = [f["id"] for f in probe.get("fields", [])]
        if recs and "modal_price" in fields and recs[0].get("report_date"):
            d = str(recs[0]["report_date"])[:10]
            if best is None or d > best[0]:
                best = (d, host, rid, name, str(recs[0]["report_date"])[10:], fields)
    if not best:
        print("Mandi prices: India Data Portal price table not reachable - MARKET_ECONOMICS stays empty (or add a data.gov.in key).")
        return
    latest, host, rid, name, dsuf, mfields = best
    latest = date.fromisoformat(latest)
    ds = lambda d: d.isoformat() + dsuf                       # the table stores dates as e.g. 2025-10-26T00:00:00 - filters must match exactly
    c_com = "commodity" if "commodity" in mfields else next((f for f in mfields if "commodity" in f.lower()), "commodity")
    print(f"  price table fields: {mfields[:15]}")
    META["mkt_url"] = f"{host}/api/3/action/datastore_search?resource_id={rid}"
    print(f"Mandi prices from portal table '{name[:40]}' - latest report date {latest}")
    cs = [r.get(c_com) for r in (ckan(host, "datastore_search", {
        "resource_id": rid, "fields": c_com, "distinct": "true", "limit": 1000}) or {}).get("records", [])]
    if not cs:                                               # 'distinct' query refused by the portal -> read the commodity names from the latest days
        for k in range(window_days):
            cs += [r_.get(c_com) for r_ in ckan_all(host, rid, {"report_date": [ds(latest - timedelta(days=k))]})]
        cs = sorted({c for c in cs if c})
    use = [c for c in cs if any(re.search(rx, str(c).lower()) for rx in MKT_REGEX.values())]
    print(f"  commodities in table: {len(cs)} | matched: {use[:20]}{' ...' if len(use) > 20 else ''}")
    if not use:
        return

    def day_rows(d):
        return ckan_all(host, rid, {"report_date": [ds(d)], c_com: use})

    def frame(rows, tag):
        f = pd.DataFrame(rows)
        if f.empty:
            return f
        for c in ("min_price", "modal_price", "max_price", "arrivals_tonnes"):
            f[c] = pd.to_numeric(f.get(c), errors="coerce")
        f["Crop"] = f[c_com].map(lambda x: next((k for k, rx in MKT_REGEX.items() if re.search(rx, str(x).lower())), None))
        f["S0"], f["D0"] = f["state_name"].map(clean_name), f["district_name"].map(clean_name)
        f["tag"] = tag
        return f.dropna(subset=["Crop"])

    cur = [frame(day_rows(latest - timedelta(days=i)), i) for i in range(window_days)]
    cur = pd.concat([f for f in cur if len(f)], ignore_index=True) if any(len(f) for f in cur) else pd.DataFrame()
    if cur.empty:
        print("  no price rows in the last week - skipped")
        return
    hist = []
    for k in range(1, past_years + 1):
        try:
            d0 = date(latest.year - k, latest.month, latest.day)
        except ValueError:
            d0 = date(latest.year - k, latest.month, 28)
        for off in (0, 1):
            f = frame(day_rows(d0 - timedelta(days=off)), -k)
            if len(f):
                hist.append(f)
    keys = ["S0", "D0", "Crop"]
    now = cur.groupby(keys).agg(**{"Min Price": ("min_price", "mean"), "Modal Price": ("modal_price", "mean"), "Max Price": ("max_price", "mean"),
                                  "Mandi Arrival": ("arrivals_tonnes", "sum"), "mkts": ("market_center", "nunique")}).reset_index()
    recent = cur[cur["tag"] < 3].groupby(keys)["modal_price"].mean().rename("r")
    older = cur[cur["tag"] >= 3].groupby(keys)["modal_price"].mean().rename("o")
    now = now.merge(recent, on=keys, how="left").merge(older, on=keys, how="left")
    chg = (now["r"] / now["o"] - 1) * 100
    now["Price Trend"] = np.where(chg > 1, "Up", np.where(chg < -1, "Down", np.where(chg.notna(), "Flat", None)))
    if hist:
        h = pd.concat(hist, ignore_index=True).groupby(keys)["modal_price"].mean().rename("avg5")
        now = now.merge(h, on=keys, how="left")
        now["Price vs 5Y Avg %"] = ((now["Modal Price"] / now["avg5"] - 1) * 100).round(1)
    else:
        now["Price vs 5Y Avg %"] = np.nan
    out = pd.DataFrame({
        "Date": latest, "State": now["S0"], "District": now["D0"], "Tehsil": "(district mandis)", "Crop": now["Crop"],
        "Mandi Arrival": now["Mandi Arrival"].round(1), "Min Price": now["Min Price"].round(0), "Modal Price": now["Modal Price"].round(0),
        "Max Price": now["Max Price"].round(0), "MSP": now["Crop"].map(MSP), "Price vs 5Y Avg %": now["Price vs 5Y Avg %"],
        "Price Trend": now["Price Trend"], "Estimated Crop Value": (now["Mandi Arrival"] * 10 * now["Modal Price"]).round(0),
        "Source": f"Agmarknet (Ministry of Agriculture) via India Data Portal - district average of reporting mandis, {window_days} days to {latest}",
        "Source Date": latest})
    DATA["MARKET_ECONOMICS"] = out.reindex(columns=HEADERS["MARKET_ECONOMICS"])
    FRESH_RUN["Agmarknet"] = latest
    NOTES[("MARKET_ECONOMICS", "Modal Price")] = "Rs per quintal; mean of daily modal prices over the last 7 reported days across the district's reporting mandis."
    NOTES[("MARKET_ECONOMICS", "Mandi Arrival")] = "Tonnes reported over the last 7 days (district total of reporting mandis)."
    NOTES[("MARKET_ECONOMICS", "Price vs 5Y Avg %")] = "Current modal price vs the average modal price on the same dates in the previous years available (2019 onward)."
    NOTES[("MARKET_ECONOMICS", "Estimated Crop Value")] = "Arrival tonnes x 10 quintals x modal price (Rs) for the 7-day window."
    print(f"MARKET_ECONOMICS: {len(out):,} district-crop rows, prices to {latest} (no API key needed)")


for _fn, _src in ((build_crop_calendar, "Reference crop calendar (CELL 2 defaults)"),
                  (agcensus_farmers, "Agriculture Census 2015-16 via India Data Portal"),
                  (lambda: market_from_portal() if DATA['MARKET_ECONOMICS'].empty else None, "Agmarknet via India Data Portal")):
    try:
        _fn()
        rlog("CELL 9", _src, META.get("mkt_url", "") if "Agmarknet" in _src else META.get("agc_url_district", "") if "Census" in _src else "",
             "", "OK")
    except Exception as e:
        print("step failed (see message):", type(e).__name__, e)
        rlog("CELL 9", _src, "", "", "FAILED", msg=f"{type(e).__name__}: {e}")


# ===== CELL 10 : NEWS SIGNALS - PIB + Google News RSS (context layer, NOT numeric data) ===
# PIB press releases and Google News headlines are text, not measurements. They are added
# to EXTREME_EVENTS as clearly-labelled "unverified news signals" with date, link and
# publisher, so you can spot floods/droughts/pest outbreaks/policy changes worth checking.
# They are NEVER used to fill numeric fields (rainfall, area, price, yield ...).
import xml.etree.ElementTree as ET
from email.utils import parsedate_to_datetime
from urllib.parse import quote_plus

PIB_FEEDS = ["https://pib.gov.in/RssMain.aspx?ModId=6&Lang=1&Regid=3", "https://pib.gov.in/RssMain.aspx?ModId=6&Lang=2&Regid=3"]  # PIB releases (both language codes)
NEWS_KEYWORDS = re.compile(
    r"crop|kharif|rabi|monsoon|rainfall|flood|drought|dry spell|heatwave|cold wave|hailstorm|cyclone|"
    r"pest|locust|armyworm|fertili[sz]er|seed|\bMSP\b|mandi|paddy|wheat|sugarcane|cotton|farmer|"
    r"agricultur|irrigation|tractor|fodder|horticulture|"
    r"किसान|फसल|कृषि|बारिश|वर्षा|सूखा|बाढ़|खरीफ|रबी|एमएसपी|मंडी|सिंचाई|मौसम", re.I)
CROP_REGEX = {"Paddy": r"paddy|\brice\b|\bdhan\b", "Wheat": r"wheat", "Maize": r"maize|\bcorn\b",
              "Sugarcane": r"sugarcane", "Cotton": r"cotton", "Groundnut": r"groundnut|peanut",
              "Soyabean": r"soy", "Millets": r"bajra|jowar|millet|ragi", "Pulses": r"pulse|arhar|\btur\b|moong|urad|lentil|chana",
              "Onion": r"onion", "Potato": r"potato", "Tomato": r"tomato"}


def _date(txt):
    try:
        return parsedate_to_datetime(txt).date()
    except Exception:
        return None


def fetch_rss(url):
    try:
        r = requests.get(url, timeout=30, headers={"User-Agent": "Mozilla/5.0"})
        r.raise_for_status()
        root = ET.fromstring(r.content)
    except Exception as e:
        print("  RSS unavailable:", url[:70], type(e).__name__)
        return []
    out = []
    for it in root.iter("item"):
        src = it.find("source")
        out.append({"title": (it.findtext("title") or "").strip(),
                    "link": (it.findtext("link") or "").strip(),
                    "date": _date(it.findtext("pubDate")),
                    "desc": re.sub(r"<[^>]+>", " ", it.findtext("description") or ""),
                    "publisher": (src.text if src is not None else "").strip()})
    return out


def news_rows(items, origin, forced_state=None):
    states = sorted(T["State"].unique(), key=len, reverse=True)
    dist_by_state = {st: sorted(g["District"].unique(), key=len, reverse=True)
                     for st, g in T.groupby("State")}
    rows, cutoff = [], TODAY - timedelta(days=NEWS_DAYS)
    for it in items:
        text = f"{it['title']} {it['desc']}"
        if not NEWS_KEYWORDS.search(text) or (it["date"] and it["date"] < cutoff):
            continue
        nt = norm(text)
        state = forced_state or next((st for st in states if len(norm(st)) > 4 and norm(st) in nt),
                                     "India (national)")
        district = next((d for d in dist_by_state.get(state, []) if len(norm(d)) >= 5 and norm(d) in nt), "")
        crop = next((c for c, rx in CROP_REGEX.items() if re.search(rx, text, re.I)), "")
        src = (f"{origin} | {it['publisher']} | " if it["publisher"] else f"{origin} | ") + it["link"] + " | UNVERIFIED news signal"
        rows.append({"Date": it["date"] or TODAY, "State": state, "District": district, "Tehsil": "",
                     "Event": "[News] " + it["title"], "Start Date": it["date"] or TODAY,
                     "Crop Affected": crop, "Source": src})
    return rows


def run_news():
    pib_items = []
    for feed in PIB_FEEDS:
        pib_items += fetch_rss(feed)
    pib_items = list({it["link"] or it["title"]: it for it in pib_items}.values())
    print(f"PIB feeds returned {len(pib_items)} items" + (f"; first: {pib_items[0]['title'][:70]}" if pib_items else " (empty or blocked)"))
    rows = news_rows(pib_items, "PIB RSS")
    print(f"PIB: {len(rows)} agriculture/weather related releases in the last {NEWS_DAYS} days")
    if NEWS_GOOGLE:
        for st in sorted(T["State"].unique()):
            nm = st.title().replace("&", "and")
            q = f'"{nm}" (farmers OR crop OR rainfall OR flood OR drought OR pest OR mandi) when:{NEWS_DAYS}d'
            url = f"https://news.google.com/rss/search?q={quote_plus(q)}&hl=en-IN&gl=IN&ceid=IN:en"
            got = news_rows(fetch_rss(url)[:15], "Google News RSS", forced_state=st)
            rows += got
            time.sleep(1.0)
        print(f"Google News: total {len(rows)} signals after all states")
    if not rows:
        return
    df = pd.DataFrame(rows).drop_duplicates(subset=["Event", "State"])
    old = DATA["EXTREME_EVENTS"]
    DATA["EXTREME_EVENTS"] = pd.concat([old, df.reindex(columns=HEADERS["EXTREME_EVENTS"])], ignore_index=True)
    FRESH_RUN["News RSS"] = TODAY
    NOTES[("EXTREME_EVENTS", "Event")] = "Rows starting with [News] are headlines from PIB / Google News RSS - unverified signals, not measurements."
    print("EXTREME_EVENTS news rows:", len(df))


def gemini_tag_news(max_items=120, chunk=15):
    """OPTIONAL: tag headlines with event type / severity / crop / district. AI chain: Gemini (primary) -> GPT -> Claude
    (whichever Colab Secret exists: GEMINI_API_KEY / OPENAI_API_KEY / ANTHROPIC_API_KEY). Reads ONLY headline text, every
    answer is validated against known crops and LGD districts, results labelled AI-tagged (unverified). Never touches numbers."""
    if not ai_available():
        print("AI tagging skipped (add a Colab Secret GEMINI_API_KEY - or OPENAI_API_KEY / ANTHROPIC_API_KEY).")
        rlog("CELL 10", "AI news tagging", "", "Tag headlines", "SKIPPED", msg="no AI key")
        return
    ev = DATA["EXTREME_EVENTS"]
    for c_ in ("Event", "Severity", "Crop Affected", "District", "Source"):
        ev[c_] = ev[c_].astype(object)                      # all-empty columns are float; text must fit
    idx = [i for i in ev.index if str(ev.at[i, "Event"]).startswith("[News]")][:max_items]
    districts = {norm(st): {norm(d): d for d in g["District"].unique()} for st, g in T.groupby("State")}
    ok_types = {"flood", "drought", "heavy_rain", "heatwave", "cold_wave", "pest", "hailstorm", "cyclone", "price", "policy", "other"}
    blank = lambda v: v is None or str(v) in ("", "nan", "None") or (isinstance(v, float) and v != v)
    done = 0
    for k in range(0, len(idx), chunk):
        part = idx[k:k + chunk]
        lines = "\n".join(f"{n}. [{ev.at[i, 'State']}] {str(ev.at[i, 'Event'])[7:200]}" for n, i in enumerate(part))
        prompt = ("For each numbered Indian agriculture news headline return JSON {\"items\":[{\"n\":<number>,"
                  '"event":"flood|drought|heavy_rain|heatwave|cold_wave|pest|hailstorm|cyclone|price|policy|other",'
                  '"severity":"Low|Medium|High|Unknown","crop":"<crop named in the headline or null>","district":"<district named in the headline or null>"}]}. '
                  "Use ONLY words that appear in the headline. If something is not stated, use null or Unknown. Do not guess.\n" + lines)
        out, prov = ai_json(prompt, "news tagging")
        if out is None:
            print("  stopping AI tagging (check the API key secrets / model names in CELL 2).")
            break
        for o in out:
            try:
                i = part[int(o["n"])]
            except Exception:
                continue
            et = str(o.get("event", "other")).lower()
            et = et if et in ok_types else "other"
            sev = o.get("severity") if o.get("severity") in ("Low", "Medium", "High") else None
            crop = o.get("crop") if o.get("crop") in list(CROP_REGEX) + list(CROP_COMMODITY) else None
            dist = districts.get(norm(ev.at[i, "State"]), {}).get(norm(o.get("district") or ""))
            ev.at[i, "Event"] = ev.at[i, "Event"].replace("[News]", f"[News|{et}]", 1)
            if sev and blank(ev.at[i, "Severity"]):
                ev.at[i, "Severity"] = sev
            if crop and blank(ev.at[i, "Crop Affected"]):
                ev.at[i, "Crop Affected"] = crop
            if dist and blank(ev.at[i, "District"]):
                ev.at[i, "District"] = dist
            ev.at[i, "Source"] = str(ev.at[i, "Source"]) + f" | AI-tagged ({prov}, unverified)"
            done += 1
        time.sleep(1.0)
    print(f"AI tagging applied to {done} headlines (event type, severity, crop, district - all unverified).")


try:
    run_news()
    rlog("CELL 10", "PIB RSS + Google News RSS", PIB_FEEDS[0], "News signals (context only)", "OK",
         int(DATA["EXTREME_EVENTS"]["Event"].astype(str).str.startswith("[News").sum()) if len(DATA["EXTREME_EVENTS"]) else 0)
    gemini_tag_news()
except Exception as e:
    print("News layer failed:", type(e).__name__, e)
    rlog("CELL 10", "PIB RSS + Google News RSS", PIB_FEEDS[0], "News signals", "FAILED", msg=f"{type(e).__name__}: {e}")


# ===== CELL 11 : OPTIONAL - SATELLITE + SOIL + CROPLAND (Earth Engine) =======
# Needs a free Google Earth Engine account + a Cloud project registered for EE
# (service-account JSON in /content). Per-tehsil zonal statistics over the real boundary
# polygons, masked to cropland. Takes roughly 30-90 minutes for ~7,000 tehsils.
import glob


def _load_info(obj):
    try:
        d = obj if isinstance(obj, dict) else json.loads(obj)
        ok = d.get("type") == "service_account" and d.get("client_email") and d.get("private_key")
        return d if ok else None
    except Exception:
        return None


def find_gee_key():
    """Look for a Google service-account key: given path, anywhere under /content (incl. sub-folders and Drive),
    or a Colab Secret. Never prints key contents."""
    cands = [GEE_KEY_FILE] if GEE_KEY_FILE else []
    cands += glob.glob("/content/**/*focus-cargo*", recursive=True)
    cands += [f for f in glob.glob("/content/**/*.json", recursive=True)
              if not re.search(r"gee_|forecast_|elevation_cache|geocode|sample_data|/drive/", f)]
    cands += glob.glob("/content/drive/MyDrive/*.json") + glob.glob("/content/drive/MyDrive/*/*.json") + glob.glob("*.json")
    for name in ("GEE_KEY", "GEE_JSON"):                    # [GitHub] Actions secret passed as an environment variable
        raw = os.environ.get(name, "").strip()
        info = _load_info(raw) if raw else None
        if info:
            return f"environment variable '{name}'", info
    seen = set()
    for c in cands:
        if c in seen or not os.path.isfile(c):
            continue
        seen.add(c)
        try:
            info = _load_info(open(c, encoding="utf-8", errors="ignore").read())
        except Exception:
            info = None
        if info:
            return c, info
    try:                                                    # Colab Secrets (key icon in the left bar)
        from google.colab import userdata
        for name in ("GEE_KEY", "GEE_JSON", "EE_KEY", "EARTHENGINE_KEY", "GOOGLE_SERVICE_ACCOUNT"):
            try:
                raw = userdata.get(name)
            except Exception as e:
                if "Access" in type(e).__name__:
                    print(f"  Secret '{name}' exists but 'Notebook access' is OFF - switch it on in the key panel.")
                continue
            info = _load_info(raw)
            if info:
                return f"Colab secret '{name}'", info
            if raw:
                print(f"  Secret '{name}' is set but is not a valid service-account JSON - paste the WHOLE file contents (starting with {{ ).")
    except ImportError:
        pass
    return None, None


def ask_gee_upload():
    """Ask for the key: Colab upload dialog. Wrapped so a blocked dialog ('Failed to fetch') does not stop the run."""
    try:
        from google.colab import files
    except ImportError:
        return None, None
    print(">>> Earth Engine key not found. Choose your service-account JSON (e.g. focus-cargo-446109-q1-....json) "
          "in the dialog below, or press Cancel to skip ...")
    try:
        up = files.upload()
    except Exception as e:
        print(f"  upload dialog failed ({type(e).__name__}) - trying Google Drive next")
        return None, None
    for name, data in (up or {}).items():
        info = _load_info(data.decode("utf-8", "ignore"))
        if info:
            with open(os.path.join("/content", os.path.basename(name)), "wb") as f:     # keep it for CELL 15 in this session
                f.write(data)
            return f"uploaded {name}", info
        print(f"  {name} is not a service-account key (needs type, client_email, private_key)")
    return None, None


def search_drive_for_key():
    """Mount Google Drive (one permission pop-up) and look for the service-account JSON there."""
    try:
        from google.colab import drive
        if not os.path.isdir("/content/drive/MyDrive"):
            drive.mount("/content/drive")
    except Exception as e:
        print(f"  Drive not mounted ({type(e).__name__})")
        return None, None
    pats = ["/content/drive/MyDrive/**/*focus-cargo*", "/content/drive/MyDrive/*.json", "/content/drive/MyDrive/*/*.json"]
    for pat in pats:
        for c in glob.glob(pat, recursive=True):
            try:
                info = _load_info(open(c, encoding="utf-8", errors="ignore").read())
            except Exception:
                info = None
            if info:
                return c, info
    print("  no service-account JSON found in your Drive (put it in My Drive to be found next time)")
    return None, None


def ensure_gee_key():
    """Secret / file -> upload dialog -> Google Drive. Sets GEE_KEYFILE, GEE_INFO. Never prints key contents."""
    global GEE_KEYFILE, GEE_INFO
    if globals().get("GEE_INFO"):
        return
    GEE_KEYFILE, GEE_INFO = find_gee_key()
    if not GEE_INFO:
        shown = sorted(os.listdir("/content"))[:25] if os.path.isdir("/content") else []
        print("No Earth Engine key in Secrets or /content. /content contains:", shown)
        if GEE_ASK_UPLOAD:
            GEE_KEYFILE, GEE_INFO = ask_gee_upload()
    if not GEE_INFO and GEE_SEARCH_DRIVE:
        GEE_KEYFILE, GEE_INFO = search_drive_for_key()
    if not GEE_INFO:
        print("  -> No JSON key. Best: add a Colab Secret named GEE_KEY (key icon in the left bar; paste the WHOLE JSON text, switch "
              f"'Notebook access' on). This run falls back to an interactive Google login for project {GEE_PROJECT!r}.")


GEE_INFO = globals().get("GEE_INFO")
GEE_KEYFILE = globals().get("GEE_KEYFILE")
ensure_gee_key()
ram("before Earth Engine")
RUN_GEE = bool(GEE_INFO or GEE_PROJECT)
print("Earth Engine key:", (f"{GEE_KEYFILE} (project {GEE_INFO.get('project_id')}, account {GEE_INFO.get('client_email')})")
      if GEE_INFO else "none - interactive login will be offered")
rlog("CELL 11", "Earth Engine credentials", "", "Satellite / soil access", "OK" if GEE_INFO else "PARTIAL",
     fallback="" if GEE_INFO else "interactive login", msg=str(GEE_KEYFILE or "no key file"))


def gee_init():
    import ee
    if GEE_INFO:
        from google.oauth2 import service_account
        creds = service_account.Credentials.from_service_account_info(
            GEE_INFO, scopes=["https://www.googleapis.com/auth/earthengine",
                              "https://www.googleapis.com/auth/cloud-platform"])
        ee.Initialize(creds, project=GEE_PROJECT or GEE_INFO.get("project_id"))
    else:
        if os.environ.get("CI"):               # [GitHub] no browser for an interactive login
            raise RuntimeError("Earth Engine key missing: add the GitHub secret GEE_KEY (whole service-account JSON)")
        ee.Authenticate()                      # interactive login fallback
        ee.Initialize(project=GEE_PROJECT)
    return ee
def usda_texture(sand, silt, clay):
    if any(pd.isna(x) for x in (sand, silt, clay)):
        return None
    if silt + 1.5 * clay < 15: return "Sand"
    if silt + 1.5 * clay < 30 and silt + 2 * clay >= 15: return "Loamy sand"
    if clay >= 40 and sand <= 45 and silt < 40: return "Clay"
    if clay >= 40 and silt >= 40: return "Silty clay"
    if clay >= 35 and sand > 45: return "Sandy clay"
    if clay >= 27 and sand <= 20: return "Silty clay loam"
    if clay >= 27 and sand <= 45: return "Clay loam"
    if clay >= 20 and sand > 45 and silt < 28: return "Sandy clay loam"
    if silt >= 80 and clay < 12: return "Silt"
    if silt >= 50 and (clay >= 12 or silt < 80): return "Silt loam"
    if clay >= 7 and silt < 50 and sand <= 52: return "Loam"
    return "Sandy loam"


GEE_FAILS = {}      # tag -> number of polygons that could not be computed


def gee_zonal(ee, image, scale, geoms, chunk=100, tag="x", tile=4):
    """Zonal means per polygon, in chunks, checkpointed to disk so a Colab disconnect can resume.
    A chunk that fails twice (timeout / memory) is split in half, down to 5 polygons, before it is given up."""
    ck = f"gee_{tag}_{TODAY}.json"
    out = {int(k): v for k, v in json.load(open(ck)).items()} if os.path.exists(ck) else {}
    if out:
        print(f"  resuming '{tag}' from checkpoint: {len(out):,} tehsils already done")
    failed = []

    def run(lo, hi):
        feats = [ee.Feature(ee.Geometry(geoms[j], "EPSG:4326", False), {"k": j}) for j in range(lo, hi)]
        err = None
        for attempt in range(2):
            try:
                res = image.reduceRegions(collection=ee.FeatureCollection(feats), reducer=ee.Reducer.mean(),
                                          scale=scale, tileScale=tile).getInfo()
                for f in res["features"]:
                    out[int(f["properties"]["k"])] = f["properties"]
                return
            except Exception as e:
                err = e
                time.sleep(4 * (attempt + 1))
        transient = re.search(r"time|memory|too many|rate|capacity|quota|429|50\d|computation|aborted|deadline|connection", str(err), re.I)
        if hi - lo > 5 and transient:
            mid = (lo + hi) // 2
            run(lo, mid)
            run(mid, hi)
        else:
            failed.append((lo, hi, f"{type(err).__name__}: {str(err)[:120]}"))

    for i in range(0, len(geoms), chunk):
        if all(j in out for j in range(i, min(i + chunk, len(geoms)))):
            continue
        run(i, min(i + chunk, len(geoms)))
        json.dump(out, open(ck, "w"))
        if (i // chunk) % 5 == 0:
            print(f"  [{tag}] zonal stats {min(i + chunk, len(geoms)):,}/{len(geoms):,}")
    if failed:
        n_f = sum(h - l for l, h, _ in failed)
        GEE_FAILS[tag] = n_f
        print(f"  [{tag}] {n_f:,} polygons failed (first error: {failed[0][2]})")
        rlog("Earth Engine", f"zonal statistics '{tag}'", "", "", "PARTIAL", len(out), msg=f"{n_f} polygons failed: {failed[0][2]}")
    return pd.DataFrame.from_dict(out, orient="index").reindex(range(len(geoms)))


def ee_last_date(ee, col):
    return pd.Timestamp(ee.Date(col.sort("system:time_start", False).first().get("system:time_start")).getInfo()["value"], unit="ms").date()


def point_squares(d=0.15):
    return [{"type": "Polygon", "coordinates": [[[lo - d, la - d], [lo + d, la - d], [lo + d, la + d], [lo - d, la + d], [lo - d, la - d]]]}
            for la, lo in zip(T["Latitude"], T["Longitude"])]


def cover_confidence(valid):
    """Data-coverage score (NOT a statistical confidence): share of SAT_TARGET_PIXELS 250 m cropland pixels the tehsil has.
    1.0 = 30+ cropland pixels (reliable zonal mean); small/unmapped tehsils score lower; 0.3 if cropland area is unknown."""
    cult = DATA["TEHSIL_MASTER"]["Cultivable Area Ha"].astype(float).reset_index(drop=True)
    conf = np.minimum(1.0, (cult / 6.25) / SAT_TARGET_PIXELS).fillna(0.3).round(2)
    return np.where(np.asarray(valid), conf.to_numpy(), np.nan)


def col(df, name):
    return df[name] if name in df else pd.Series(np.nan, index=range(len(T)))


def run_gee():
    from shapely.geometry import mapping
    try:
        ee = gee_init()
    except Exception as e:
        print("Earth Engine login failed:", type(e).__name__, str(e)[:300])
        rlog("CELL 11", "Earth Engine login", "https://earthengine.google.com", "", "FAILED", msg=f"{type(e).__name__}: {str(e)[:200]}")
        print("  Check: (1) Earth Engine API enabled on the project, (2) the project is registered for Earth Engine,\n"
              "  (3) the service account has an Earth Engine role (e.g. Earth Engine Resource Viewer) in IAM.")
        return
    geoms = [mapping(g.simplify(SIMPLIFY_DEG, preserve_topology=True)) for g in T["geometry"]]
    lc = ee.ImageCollection("MODIS/061/MCD12Q1").sort("system:time_start", False).first().select("LC_Type1")
    crop = lc.eq(12).Or(lc.eq(14))         # cropland + cropland/natural mosaic
    rlog("CELL 11", "Earth Engine login", "https://earthengine.google.com", "", "OK")
    for step in (gee_satellite, gee_soil, gee_cropland, gee_groundwater, gee_climate_check, gee_sentinel2):
        t0 = time.time()
        try:
            step(ee, geoms, crop)
            rlog("CELL 11", step.__name__, "Earth Engine", "", "OK", msg=f"{(time.time() - t0) / 60:.1f} min")
        except Exception as e:
            print(f"{step.__name__} failed: {type(e).__name__}: {e}")
            rlog("CELL 11", step.__name__, "Earth Engine", "", "FAILED", msg=f"{type(e).__name__}: {e}")


def gee_satellite(ee, geoms, crop):
    veg = ee.ImageCollection("MODIS/061/MOD13Q1").filterDate(str(TODAY - timedelta(days=60)),
                                                             str(TODAY + timedelta(days=1)))
    ordered = veg.sort("system:time_start", False)
    latest, prev = ee.Image(ordered.first()), ee.Image(ordered.toList(2).get(1))
    sat_date = latest.date().format("YYYY-MM-dd").getInfo()
    doy = ee.Date(sat_date).getRelative("day", "year").getInfo() + 1
    hist = (ee.ImageCollection("MODIS/061/MOD13Q1")
            .filterDate(f"{TODAY.year - 5}-01-01", f"{TODAY.year}-01-01")
            .filter(ee.Filter.calendarRange(max(1, doy - 16), min(366, doy + 16), "day_of_year"))
            .select("NDVI").mean())
    recent = lambda cid: ee.ImageCollection(cid).filterDate(
        str(TODAY - timedelta(days=30)), str(TODAY + timedelta(days=1))).sort("system:time_start", False).first()
    ndwi_img = ee.Image(recent("MODIS/061/MOD09A1")).normalizedDifference(["sur_refl_b02", "sur_refl_b06"])
    lst = ee.Image(recent("MODIS/061/MOD11A2")).select("LST_Day_1km").multiply(0.02).subtract(273.15)
    bands = [latest.select("NDVI").multiply(0.0001).rename("ndvi"),
             latest.select("EVI").multiply(0.0001).rename("evi"),
             prev.select("NDVI").multiply(0.0001).rename("ndvi_prev"),
             hist.multiply(0.0001).rename("ndvi5"), ndwi_img.rename("ndwi"), lst.rename("lst")]
    for cid in ("NASA/SMAP/SPL4SMGP/008", "NASA/SMAP/SPL4SMGP/007"):     # SMAP is optional
        try:
            sm = (ee.ImageCollection(cid).filterDate(str(TODAY - timedelta(days=5)), str(TODAY + timedelta(days=1)))
                  .select("sm_surface").mean().rename("sm"))
            sm.bandNames().getInfo()
            bands.append(sm)
            break
        except Exception:
            continue
    else:
        print("  SMAP soil moisture unavailable - skipped")
    sat = gee_zonal(ee, ee.Image.cat(bands).updateMask(crop), 250, geoms, tag="sat")
    FRESH_RUN["MODIS (via GEE)"] = date.fromisoformat(sat_date)
    ndvi = col(sat, "ndvi")
    S = pd.DataFrame({"Date": TODAY, "State": T["State"], "District": T["District"],
                      "Tehsil": T["Tehsil"], "Crop": "All cropland (MODIS land-cover mask)",
                      "NDVI": ndvi.round(3), "NDVI 5Y Avg": col(sat, "ndvi5").round(3),
                      "EVI": col(sat, "evi").round(3), "NDWI": col(sat, "ndwi").round(3),
                      "Soil Moisture": col(sat, "sm").round(3), "LST": col(sat, "lst").round(1),
                      "Satellite Date": date.fromisoformat(sat_date),
                      "Satellite Source": "MODIS MOD13Q1/MOD09A1/MOD11A2 (+SMAP) via Earth Engine"})
    d = ndvi - col(sat, "ndvi_prev")
    S["Vegetation Trend"] = np.where(d > 0.02, "Improving", np.where(d < -0.02, "Declining", "Stable"))
    S.loc[ndvi.isna(), "Vegetation Trend"] = None
    S["Confidence"] = cover_confidence(ndvi.notna())
    NOTES[("SATELLITE_INTELLIGENCE", "Confidence")] = "Data-coverage score: share of 30 cropland 250 m pixels available in the tehsil (not a statistical confidence)."
    install("SATELLITE_INTELLIGENCE", S)
    NOTES[("SATELLITE_INTELLIGENCE", "Soil Moisture")] = "SMAP surface soil moisture (m3/m3), ~11 km - coarse."
    print("Satellite filled:", int(ndvi.notna().sum()), "tehsils; image date", sat_date)


def gee_soil(ee, geoms, crop):               # ISRIC SoilGrids 2.0, 0-5 cm, modelled
    props = ["phh2o", "soc", "clay", "sand", "silt", "nitrogen"]
    img = ee.Image.cat([ee.Image(f"projects/soilgrids-isric/{p}_mean").select(f"{p}_0-5cm_mean").rename(p)
                        for p in props])
    sg = gee_zonal(ee, img.updateMask(crop), 250, geoms, tag="soil")
    O = pd.DataFrame({"State": T["State"], "District": T["District"], "Tehsil": T["Tehsil"],
                      "pH": (col(sg, "phh2o") / 10).round(2),
                      "Organic Carbon": (col(sg, "soc") / 100).round(2),     # dg/kg -> %
                      "N": (col(sg, "nitrogen") / 100).round(2),             # cg/kg -> g/kg (total N)
                      "Source": "ISRIC SoilGrids 2.0 (modelled, 250 m)"})
    O["Soil Texture"] = [usda_texture(a / 10, b / 10, c / 10) for a, b, c in
                         zip(col(sg, "sand"), col(sg, "silt"), col(sg, "clay"))]
    O["Confidence"] = cover_confidence(col(sg, "phh2o").notna())
    NOTES[("SOIL_INTELLIGENCE", "Confidence")] = "Data-coverage score: share of 30 cropland 250 m pixels available in the tehsil (SoilGrids is itself modelled)."
    install("SOIL_INTELLIGENCE", O)
    NOTES[("SOIL_INTELLIGENCE", "N")] = "TOTAL nitrogen g/kg (SoilGrids), not available-N kg/ha. P and K are not available from open grids."
    NOTES[("SOIL_INTELLIGENCE", "Organic Carbon")] = "Percent; SoilGrids soil organic carbon 0-5 cm (modelled)."
    print("Soil filled:", int(O["pH"].notna().sum()), "tehsils")


def gee_groundwater(ee, geoms, crop):
    """Groundwater STORAGE trend from NASA GLDAS-2.2 (CLSM, GRACE-assimilated): last 12 months vs the 12 before, ~25 km model.
    A regional signal only - not well levels. Each tehsil is represented by a ~30 km square around its centre."""
    gl = ee.ImageCollection("NASA/GLDAS/V022/CLSM/G025/DA1D").select("GWS_tavg")
    end = TODAY - timedelta(days=60)                       # GLDAS latency
    recent = gl.filterDate(str(end - timedelta(days=365)), str(end)).mean()
    prev = gl.filterDate(str(end - timedelta(days=730)), str(end - timedelta(days=365))).mean()
    diff = recent.subtract(prev).rename("gws_change")
    d = 0.15
    sq = [{"type": "Polygon", "coordinates": [[[lo - d, la - d], [lo + d, la - d], [lo + d, la + d], [lo - d, la + d], [lo - d, la - d]]]}
          for la, lo in zip(T["Latitude"], T["Longitude"])]
    gw = gee_zonal(ee, diff, 25000, sq, chunk=500, tag="gw")
    ch = col(gw, "mean")
    W = DATA["WATER_IRRIGATION"]
    W["Groundwater Trend"] = np.where(ch.isna(), None, np.where(ch > 10, "Rising", np.where(ch < -10, "Falling", "Stable")))
    FRESH_RUN["GLDAS (via GEE)"] = end
    NOTES[("WATER_IRRIGATION", "Groundwater Trend")] = ("NASA GLDAS-2.2 CLSM (GRACE-assimilated) groundwater-storage change, last 12 months vs previous 12 "
                                                       "(Rising > +10 mm, Falling < -10 mm), ~25 km model grid. Regional signal, NOT well-level data.")
    print("Groundwater trend filled:", int(ch.notna().sum()), "tehsils; change mm: median", round(float(ch.median()), 1))


def gee_cropland(ee, geoms, crop):           # ESA WorldCover 2021 cropland share
    wc = ee.ImageCollection("ESA/WorldCover/v200").first().select("Map")
    cf = gee_zonal(ee, wc.eq(40).rename("crop"), 100, geoms, tag="crop")
    ha = (col(cf, "mean") * T["Total Area Ha"]).round(0)
    cur = DATA["TEHSIL_MASTER"]["Cultivable Area Ha"]
    DATA["TEHSIL_MASTER"]["Cultivable Area Ha"] = cur.where(cur.notna(), pd.Series(ha.values, index=cur.index))
    NOTES[("TEHSIL_MASTER", "Cultivable Area Ha")] = (
        "PROXY: CoRE Stack cropland+orchard where available, else ESA WorldCover 2021 cropland share x boundary area.")
    print("Cropland area filled:", int(ha.notna().sum()), "tehsils")


URL_CAT = "https://developers.google.com/earth-engine/datasets/catalog/"


def gee_climate_check(ee, geoms, crop):
    """INDEPENDENT CHECK of IMD (not a replacement): CHIRPS rainfall and ERA5-Land air temperature over the SAME days
    as the IMD values, ~33 km square around each tehsil centre. Results go to DATA_VALIDATION only."""
    if not RUN_CLIMATE_CHECK:
        return
    sq = point_squares(0.15)
    rd, rdates = VAL_STORE.get("imd_rain_daily"), VAL_STORE.get("imd_dates")
    if rd is not None and rdates:
        ch = ee.ImageCollection("UCSB-CHG/CHIRPS/DAILY").select("precipitation")
        last = ee_last_date(ee, ch)
        days = [d for d in rdates if d <= last]
        if len(days) >= 10:
            s0, s1 = days[0], days[-1]
            z = gee_zonal(ee, ch.filterDate(str(s0), str(s1 + timedelta(days=1))).sum(), 5566, sq, chunk=500, tag="chirps")
            ix = [rdates.index(d) for d in days]
            sub = rd[ix]
            VAL_STORE["rain_gov"] = np.where(np.isnan(sub).all(axis=0), np.nan, np.nansum(sub, axis=0))
            VAL_STORE["rain_ind"] = pd.to_numeric(col(z, "mean"), errors="coerce").to_numpy(float)
            VAL_STORE["rain_period"] = f"{s0} to {s1} ({len(days)} days)"
            print(f"  CHIRPS rainfall check: {np.isfinite(VAL_STORE['rain_ind']).sum():,} tehsils, {VAL_STORE['rain_period']}")
        else:
            print(f"  CHIRPS check skipped: CHIRPS data end {last}, only {len(days)} overlapping days with IMD")
            rlog("CELL 11", "CHIRPS daily", URL_CAT + "UCSB-CHG_CHIRPS_DAILY", "IMD rainfall check", "SKIPPED", msg=f"CHIRPS ends {last}")
    td, tdates = VAL_STORE.get("imd_tmean_daily"), VAL_STORE.get("imd_t_dates")
    if td is not None and tdates:
        er = ee.ImageCollection("ECMWF/ERA5_LAND/DAILY_AGGR").select("temperature_2m")
        last = ee_last_date(ee, er)
        days = [d for d in tdates if d <= last]
        if len(days) >= 2:
            s0, s1 = days[0], days[-1]
            z = gee_zonal(ee, er.filterDate(str(s0), str(s1 + timedelta(days=1))).mean().subtract(273.15), 11132, sq, chunk=500, tag="era5t")
            ix = [tdates.index(d) for d in days]
            VAL_STORE["temp_gov"] = np.nanmean(td[ix], axis=0)
            VAL_STORE["temp_ind"] = pd.to_numeric(col(z, "mean"), errors="coerce").to_numpy(float)
            VAL_STORE["temp_period"] = f"{s0} to {s1} ({len(days)} days, daily mean)"
            print(f"  ERA5-Land temperature check: {np.isfinite(VAL_STORE['temp_ind']).sum():,} tehsils, {VAL_STORE['temp_period']}")
        else:
            print(f"  ERA5-Land check skipped: ERA5 data end {last}, only {len(days)} overlapping days with IMD")
            rlog("CELL 11", "ERA5-Land daily", URL_CAT + "ECMWF_ERA5_LAND_DAILY_AGGR", "IMD temperature check", "SKIPPED", msg=f"ERA5 ends {last}")


def gee_sentinel2(ee, geoms, crop):
    """Sentinel-2 (10 m) NDVI / EVI / NDWI median of cloud-free observations over ESA WorldCover cropland, last S2_DAYS days,
    plus cloud share (fills SATELLITE_INTELLIGENCE 'Cloud Cover %') and the number of clear observations.
    Independent observation for the MODIS values (DATA_VALIDATION); MODIS columns are not overwritten."""
    if not RUN_SENTINEL2:
        print("  Sentinel-2 skipped (RUN_SENTINEL2 = False)")
        return
    start, end = TODAY - timedelta(days=S2_DAYS), TODAY + timedelta(days=1)
    wc = ee.ImageCollection("ESA/WorldCover/v200").first().select("Map").eq(40)

    def prep(img):
        scl = img.select("SCL")
        cloud = scl.eq(3).Or(scl.eq(8)).Or(scl.eq(9)).Or(scl.eq(10))       # shadow, cloud medium/high, cirrus
        valid = scl.gt(1)
        clear = valid.And(cloud.Not())
        b = img.select(["B2", "B4", "B8", "B11"]).multiply(0.0001)
        ndvi = b.normalizedDifference(["B8", "B4"]).rename("s2ndvi")
        evi = b.expression("2.5*(N-R)/(N+6*R-7.5*B+1)", {"N": b.select("B8"), "R": b.select("B4"), "B": b.select("B2")}).rename("s2evi")
        ndwi = b.normalizedDifference(["B8", "B11"]).rename("s2ndwi")
        return (ee.Image.cat([ndvi, evi, ndwi]).updateMask(clear)
                .addBands(cloud.updateMask(valid).toFloat().rename("cloud")).addBands(clear.toFloat().rename("nclear")))

    c = ee.ImageCollection("COPERNICUS/S2_SR_HARMONIZED").filterDate(str(start), str(end)).map(prep)
    img = ee.Image.cat([c.select(["s2ndvi", "s2evi", "s2ndwi"]).median(), c.select("cloud").mean(),
                        c.select("nclear").sum()]).updateMask(wc)
    z = gee_zonal(ee, img, S2_SCALE, geoms, chunk=25, tag="s2", tile=8)
    s2 = pd.DataFrame({k: pd.to_numeric(col(z, k), errors="coerce").to_numpy(float) for k in ("s2ndvi", "s2evi", "s2ndwi", "cloud", "nclear")})
    VAL_STORE["s2"] = s2
    VAL_STORE["s2_period"] = f"{start} to {TODAY} (median of cloud-free scenes)"
    S = DATA["SATELLITE_INTELLIGENCE"]
    if len(S) == len(T):
        S["Cloud Cover %"] = np.round(s2["cloud"].to_numpy() * 100, 0)
        NOTES[("SATELLITE_INTELLIGENCE", "Cloud Cover %")] = (f"Sentinel-2: share of cropland pixel-observations flagged cloud/shadow/cirrus (SCL) in the last {S2_DAYS} days. "
                                                             "High values = satellite indicators less reliable. Sentinel-2 vs MODIS comparison is in DATA_VALIDATION.")
    FRESH_RUN["Sentinel-2 (via GEE)"] = TODAY
    print(f"Sentinel-2 filled: {int(np.isfinite(s2['s2ndvi']).sum()):,} tehsils; low-observation tehsils (< {S2_MIN_CLEAR_OBS} clear): "
          f"{int((s2['nclear'] < S2_MIN_CLEAR_OBS).sum()):,}")


if RUN_GEE:
    run_gee()
else:
    print("Skipped Earth Engine: no service-account JSON found in /content (upload it, then re-run this cell).")


# ===== CELL 12 : DECISION LAYER - AGRI_INTELLIGENCE, ALERTS, TRACTOR_OPPORTUNITY (rule-based) ====
# Transparent rules, no hidden model. Scores use only inputs that exist; missing components are skipped and the rest re-weighted.
# Market score is omitted until you have price data. Tractor business inputs (prosperity, mechanisation, past demand, replacement)
# come from YOUR data: put a CSV named tractor_inputs.csv (State,District,Tehsil + those columns) in /content and re-run.
def _aligned(df, base):
    """Reorder df to the tehsil order of `base` (keys survive duplicate names)."""
    d = df.copy()
    d.index = uniq_key(d)
    return d.reindex(uniq_key(base)).reset_index(drop=True)


def _clip(x):
    return np.clip(x, 0, 100)


def decision_layer():
    M = DATA["TEHSIL_MASTER"]
    if M.empty:
        print("No tehsil data - decision layer skipped")
        return
    sat, wx, wt, so = (_aligned(DATA[k], M) for k in ("SATELLITE_INTELLIGENCE", "WEATHER_LIVE", "WATER_IRRIGATION", "SOIL_INTELLIGENCE"))
    num = lambda df, c: pd.to_numeric(df[c], errors="coerce") if c in df else pd.Series(np.nan, index=df.index)
    anom = (num(sat, "NDVI") / num(sat, "NDVI 5Y Avg") - 1) * 100
    dev = (num(wx, "Rainfall 30D") - num(wx, "Normal Rainfall mm")) / num(wx, "Normal Rainfall mm").where(num(wx, "Normal Rainfall mm") > 0) * 100
    f7 = num(wx, "Forecast 7D")
    sc = pd.DataFrame(index=M.index)
    sc["Crop Health"] = _clip(75 + 2 * anom)
    stress = wt["Water Stress"].map({"Low": 90, "Medium": 60, "High": 30}) if "Water Stress" in wt else pd.Series(np.nan, index=M.index)
    sc["Water Score"] = _clip(stress.fillna(60) + dev.clip(-50, 50).fillna(0) * 0.3).where(stress.notna() | dev.notna())
    alert = wx["Extreme Weather Alert"].astype(str).eq("Heavy rain watch") if "Extreme Weather Alert" in wx else False
    sc["Weather Score"] = _clip(80 - dev.abs().fillna(0) * 0.4 - num(wx, "Dry Days").fillna(0) * 2 - alert * 15).where(dev.notna() | num(wx, "Dry Days").notna())
    ph = num(so, "pH")
    ph_s = 100 - 30 * (ph.sub(6.75).abs() - 0.75).clip(lower=0)
    oc_s = _clip(num(so, "Organic Carbon") / 0.75 * 100)
    sc["Soil Score"] = pd.concat([ph_s, oc_s], axis=1).mean(axis=1)
    sat_conf = num(sat, "Confidence")

    # ---- tehsil x crop rows from the tehsil-level CROP_SOWN split (latest reported year) ----
    cs = DATA["CROP_SOWN"]
    cs = cs[(cs["Tehsil"] != "(district level)") & cs["Tehsil"].notna()].copy() if len(cs) else cs
    yo = DATA["YIELD_PRODUCTION"]
    ymap = {}
    if len(yo):
        for st, di, cr, pv in zip(yo["State"], yo["District"], yo["Crop"], pd.to_numeric(yo["Yield vs 5Y Avg %"], errors="coerce")):
            lg = NAME_MAP.get((st, di))
            if lg:
                ymap[(lg[0], lg[1], cr)] = pv
    if len(cs):
        last = cs["Report Date"].max()
        cs = cs[cs["Report Date"] == last]
        rows = cs.groupby(["State", "District", "Tehsil", "Crop"], as_index=False)["Sown Area Ha"].sum()
        rows = rows[rows["Sown Area Ha"] > 0]
    else:
        rows = pd.DataFrame(columns=["State", "District", "Tehsil", "Crop", "Sown Area Ha"])
    base_key = (M["State"].map(norm) + "|" + M["District"].map(norm) + "|" + M["Tehsil"].map(norm))
    first_pos = dict(zip(base_key[~base_key.duplicated()], np.flatnonzero(~base_key.duplicated())))
    rows["pos"] = (rows["State"].map(norm) + "|" + rows["District"].map(norm) + "|" + rows["Tehsil"].map(norm)).map(first_pos)
    rows = rows.dropna(subset=["pos"]).copy()
    rows["pos"] = rows["pos"].astype(int)
    pctv = pd.Series([ymap.get((a, b, c), np.nan) for a, b, c in zip(rows["State"], rows["District"], rows["Crop"])], index=rows.index)
    rows["Yield Outlook"] = _clip(60 + 3 * pctv)
    rows["yo_label"] = np.where(pctv > 3, "Positive", np.where(pctv < -3, "Negative", np.where(pctv.notna(), "Stable", "n/a")))
    mkt = DATA["MARKET_ECONOMICS"]
    mscore = {}
    if len(mkt):
        mp_ = match_district_names(list(zip(mkt["State"], mkt["District"])))
        for st, di, cr, pv in zip(mkt["State"], mkt["District"], mkt["Crop"], pd.to_numeric(mkt["Price vs 5Y Avg %"], errors="coerce")):
            lg = mp_.get((st, di))
            if lg and pv == pv:
                mscore[(lg[0], lg[1], cr)] = float(_clip(60 + 2 * pv))
    rows["Market Score"] = [mscore.get((a, b, c), np.nan) for a, b, c in zip(rows["State"], rows["District"], rows["Crop"])]
    comp = pd.concat([sc.iloc[rows["pos"].to_numpy()].reset_index(drop=True), rows[["Yield Outlook", "Market Score"]].reset_index(drop=True)], axis=1)
    w = pd.Series(SCORE_WEIGHTS)
    comp = comp[list(w.index)]
    avail = comp.notna()
    wsum = (avail * w).sum(axis=1)
    overall = ((comp.fillna(0) * w).sum(axis=1) / wsum.where(wsum > 0)).round(0)
    conf = (wsum / w.sum() * sat_conf.iloc[rows["pos"].to_numpy()].fillna(0.5).to_numpy()).round(2)
    risk = np.where(overall < 55, "High", np.where(overall < 70, "Medium", "Low"))
    opp = np.where(overall >= 75, "High", np.where(overall >= 60, "Medium", "Low"))
    weakest = comp.fillna(np.inf).idxmin(axis=1).where(avail.any(axis=1))
    an_v, dv_v = anom.iloc[rows["pos"].to_numpy()].to_numpy(), dev.iloc[rows["pos"].to_numpy()].to_numpy()
    fmtv = lambda v: "n/a" if v != v else f"{v:+.0f}%"
    reason = [f"Weakest: {wk} {comp.at[i, wk]:.0f}; NDVI vs 5Y {fmtv(a)}; 30D rain vs normal {fmtv(d)}" if isinstance(wk, str) else "Insufficient inputs"
              for i, (wk, a, d) in enumerate(zip(weakest, an_v, dv_v))]
    action = np.where(risk == "High", "Field verification + advisory", np.where(risk == "Medium", "Close monitoring", "Normal monitoring"))
    imd_end = FRESH_RUN.get("IMD gridded")
    AI = pd.DataFrame({
        "Date": TODAY, "State": rows["State"].values, "District": rows["District"].values, "Tehsil": rows["Tehsil"].values,
        "Crop": rows["Crop"].values, "Crop Health": comp["Crop Health"].round(0).values, "Water Score": comp["Water Score"].round(0).values,
        "Weather Score": comp["Weather Score"].round(0).values, "Soil Score": comp["Soil Score"].round(0).values,
        "Yield Outlook": rows["yo_label"].values, "Market Score": comp["Market Score"].round(0).values, "Overall Agri Score": overall.values,
        "Risk": np.where(overall.isna(), None, risk), "Opportunity": np.where(overall.isna(), None, opp), "Confidence": conf,
        "Key Reason": reason, "Recommended Action": np.where(overall.isna(), None, action),
        "Source": "Derived (rule-based) from IMD, MODIS/Earth Engine, CGWB, soil, DES APY, Agmarknet prices (when available)",
        "Source Date": max([d for d in (imd_end, FRESH_RUN.get("MODIS (via GEE)"), FRESH_RUN.get("DES APY")) if d] or [TODAY])})
    DATA["AGRI_INTELLIGENCE"] = AI.reindex(columns=HEADERS["AGRI_INTELLIGENCE"])
    NOTES[("AGRI_INTELLIGENCE", "Overall Agri Score")] = ("Weighted mean of Crop Health, Water, Weather, Soil, Yield Outlook (weights in SCORE_WEIGHTS, CELL 2); "
                                                         "missing components skipped. Crop = latest reported crops of the tehsil (APY split). Rule-based, not a forecast model.")
    NOTES[("AGRI_INTELLIGENCE", "Crop Health")] = "75 + 2 x NDVI anomaly % (MODIS vs 5-year average), clipped 0-100. Tehsil-wide cropland NDVI, not crop-specific."
    print(f"AGRI_INTELLIGENCE: {len(AI):,} tehsil-crop rows (score available for {int(overall.notna().sum()):,})")

    # ---- CROP_HEALTH (same tehsil x crop rows): satellite + weather scores; no stage score (needs sowing dates) ----
    pr = rows["pos"].to_numpy()
    smv = num(sat, "Soil Moisture").iloc[pr].to_numpy()
    ndw = num(sat, "NDWI").iloc[pr].to_numpy()
    lst = num(sat, "LST").iloc[pr].to_numpy()
    moist = np.where(~np.isnan(smv), _clip((smv - 0.10) / 0.25 * 100), np.where(~np.isnan(ndw), _clip((ndw + 0.10) / 0.50 * 100), np.nan))
    rain_s = _clip(100 - np.abs(dv_v) * 0.8)
    temp_s = _clip(100 - 8 * np.maximum(0, np.abs(lst - 26) - 6))
    parts = pd.DataFrame({"n": comp["Crop Health"].to_numpy(), "m": moist, "r": rain_s, "t": temp_s})
    oh = parts.mean(axis=1).round(0)
    trend = sat["Vegetation Trend"].iloc[pr].to_numpy() if "Vegetation Trend" in sat else np.full(len(pr), None)
    cmp_ = np.where(np.isnan(an_v), None, np.where(an_v > 5, "Above 5Y avg", np.where(an_v < -5, "Below 5Y avg", "Near 5Y avg")))
    status = np.where(oh.isna(), None, np.where(oh >= 75, "Good", np.where(oh >= 55, "Watch", "Stressed")))
    sc_ = sat_conf.iloc[pr].fillna(0.5).to_numpy()
    d_h = max([d for d in (imd_end, FRESH_RUN.get("MODIS (via GEE)")) if d] or [TODAY])
    CH = pd.DataFrame({
        "Date": TODAY, "State": rows["State"].values, "District": rows["District"].values, "Tehsil": rows["Tehsil"].values,
        "Crop": rows["Crop"].values, "NDVI Score": parts["n"].round(0).values, "Moisture Score": np.round(moist, 0),
        "Rainfall Score": np.round(rain_s, 0), "Temperature Score": np.round(temp_s, 0), "Crop Stage Score": np.nan,
        "Overall Crop Health": oh.values, "5Y Comparison": cmp_, "Health Trend": trend, "Status": status, "Confidence": np.round(sc_, 2),
        "Reason": [f"NDVI vs 5Y {fmtv(a)}; 30D rain vs normal {fmtv(d)}; LST {l:.0f} C" if l == l else f"NDVI vs 5Y {fmtv(a)}; 30D rain vs normal {fmtv(d)}"
                   for a, d, l in zip(an_v, dv_v, lst)],
        "Source": "Derived (rule-based) from MODIS NDVI/NDWI/LST + SMAP (Earth Engine) and IMD rainfall; cropland-wide per tehsil, not crop-specific",
        "Source Date": d_h})
    DATA["CROP_HEALTH"] = CH.reindex(columns=HEADERS["CROP_HEALTH"])
    NOTES[("CROP_HEALTH", "NDVI Score")] = "75 + 2 x NDVI anomaly % vs the 5-year average, clipped 0-100 (tehsil-wide cropland, not crop-specific)."
    NOTES[("CROP_HEALTH", "Crop Stage Score")] = "Not scored: needs observed sowing dates / crop stage (no open source)."
    print(f"CROP_HEALTH: {len(CH):,} rows")

    # ---- CROP_STAGE: CALENDAR-BASED expectation from reference sowing windows (not observed) ----
    def sow_date(crop, sk):
        w, dur = _window(IDEAL_SOWING.get(crop, {}).get(sk)), CROP_DURATION.get(crop)
        if not (w and dur):
            return None
        for yr in (TODAY.year - 1, TODAY.year):
            a = date(yr, w[0].month, w[0].day)
            b = date(yr, w[1].month, w[1].day)
            b = b if b >= a else date(yr + 1, w[1].month, w[1].day)
            mid = a + (b - a) / 2
            if mid <= TODAY <= mid + timedelta(days=dur + 14):
                return mid
        return None
    if len(cs):
        c2 = cs.copy()
        c2["sk"] = c2["Season"].map(season_key)
        act = {k: sow_date(*k) for k in c2[["Crop", "sk"]].drop_duplicates().itertuples(index=False, name=None)}
        c2["sow"] = [act.get(k) for k in zip(c2["Crop"], c2["sk"])]
        c2 = c2[c2["sow"].notna()].drop_duplicates(["State", "District", "Tehsil", "Crop", "sk"]).copy()
        c2["pos"] = (c2["State"].map(norm) + "|" + c2["District"].map(norm) + "|" + c2["Tehsil"].map(norm)).map(first_pos)
        c2 = c2.dropna(subset=["pos"])
        if len(c2):
            ps = c2["pos"].astype(int).to_numpy()
            dur_ = c2["Crop"].map(CROP_DURATION).astype(float).to_numpy()
            days = np.array([(TODAY - d_).days for d_ in c2["sow"]], dtype=float)
            prog = np.clip(days / dur_ * 100, 0, 100)
            names = [n_ for n_, _, _ in STAGES]
            edges = [f0 * 100 for _, f0, _ in STAGES]
            stage = np.array([names[max(i for i, e in enumerate(edges) if p_ >= e)] for p_ in prog], dtype=object)
            stage = np.where(days > dur_, "Harvest", stage)
            dvs = dev.iloc[ps].to_numpy()
            wsig = np.where(np.isnan(dvs), None, np.where(dvs < -25, "Dry", np.where(dvs > 60, "Wet", "Favourable")))
            CS = pd.DataFrame({
                "Date": TODAY, "State": c2["State"].values, "District": c2["District"].values, "Tehsil": c2["Tehsil"].values,
                "Crop": c2["Crop"].values, "Sowing Date": c2["sow"].values, "Expected Stage": stage, "Estimated Current Stage": stage,
                "Stage Progress %": np.round(prog, 0), "Days From Normal": np.nan,
                "Satellite Signal": sat["Vegetation Trend"].iloc[ps].to_numpy() if "Vegetation Trend" in sat else None,
                "Weather Signal": wsig, "Stage Confidence": 0.3,
                "Source": "CALENDAR-BASED expectation from reference sowing windows + typical durations (NOT observed); signals from MODIS / IMD",
                "Data Date": TODAY, "Source Date": FRESH_RUN.get("MODIS (via GEE)") or TODAY})
            DATA["CROP_STAGE"] = CS.reindex(columns=HEADERS["CROP_STAGE"])
            NOTES[("CROP_STAGE", "Sowing Date")] = "Midpoint of the REFERENCE sowing window (IDEAL_SOWING, CELL 2) - not an observed sowing date."
            NOTES[("CROP_STAGE", "Stage Confidence")] = "0.3 = calendar expectation only. Replace with observed sowing dates for a real estimate."
            print(f"CROP_STAGE: {len(CS):,} tehsil-crop rows currently in season (calendar-based)")

    # ---- PEST_DISEASE: news signals about pests (unverified) ----
    ev = DATA["EXTREME_EVENTS"]
    if len(ev):
        txt = ev["Event"].astype(str)
        pm = txt.str.contains(r"pest|armyworm|locust|whitefly|blight|disease|aphid|bollworm|hopper|rust|wilt|mosaic|infest|attack|keeda|कीट|रोग", case=False, regex=True)
        pe = ev[pm]
        if len(pe):
            PD = pd.DataFrame({
                "Date": pe["Date"].values, "State": pe["State"].values, "District": pe["District"].values, "Tehsil": "",
                "Crop": pe["Crop Affected"].values, "Pest/Disease": txt[pm].str.replace(r"^\[News[^\]]*\]\s*", "", regex=True).str[:140].values,
                "Alert Level": pe["Severity"].values, "Source": pe["Source"].values, "Report Date": pe["Date"].values, "Confidence": 0.3,
                "Source Date": pe["Date"].values})
            DATA["PEST_DISEASE"] = PD.reindex(columns=HEADERS["PEST_DISEASE"])
            NOTES[("PEST_DISEASE", "Pest/Disease")] = "News headline mentioning a pest or disease - UNVERIFIED signal, not surveillance data."
            print(f"PEST_DISEASE: {len(PD):,} news-based pest/disease signals")

    # ---- ALERTS (thresholds are explicit) ----
    al = []
    base = pd.DataFrame({"State": M["State"], "District": M["District"], "Tehsil": M["Tehsil"], "anom": anom, "dev": dev, "f7": f7,
                         "norm": num(wx, "Normal Rainfall mm"), "sat_conf": sat_conf})
    rules = [
        ("Drought risk", "Red", (base["dev"] <= -50) & (base["norm"] >= MIN_NORMAL_MM), "dev", 0, f"30-day rain <= -50% of normal (normal >= {MIN_NORMAL_MM} mm)", "Check irrigation/contingency plan", 0.8),
        ("Drought risk", "Orange", (base["dev"] > -50) & (base["dev"] <= -25) & (base["norm"] >= MIN_NORMAL_MM), "dev", 0, f"30-day rain -50% to -25% of normal (normal >= {MIN_NORMAL_MM} mm)", "Monitor soil moisture", 0.8),
        ("Crop stress", "Orange", base["anom"] <= -15, "anom", 0, "Cropland NDVI <= -15% vs 5-year average", "Field verification", None),
        ("Heavy rain watch", "Orange", base["f7"] > HEAVY_RAIN_7D_MM, "f7", HEAVY_RAIN_7D_MM, f"7-day forecast rain > {HEAVY_RAIN_7D_MM} mm", "Drainage / harvest planning", 0.6),
        ("Heavy rain watch", "Yellow", (base["f7"] > 0.6 * HEAVY_RAIN_7D_MM) & (base["f7"] <= HEAVY_RAIN_7D_MM), "f7", HEAVY_RAIN_7D_MM, f"7-day forecast rain {int(0.6 * HEAVY_RAIN_7D_MM)}-{HEAVY_RAIN_7D_MM} mm", "Watch forecast", 0.6),
        ("Crop health improvement", "Green", base["anom"] >= 10, "anom", 0, "Cropland NDVI >= +10% vs 5-year average", "No action", None),
    ]
    for typ, sev, mask, col_, nrm, trig, act, cf in rules:
        b = base[mask.fillna(False)]
        if b.empty:
            continue
        al.append(pd.DataFrame({
            "Alert Date": TODAY, "State": b["State"], "District": b["District"], "Tehsil": b["Tehsil"], "Crop": "All cropland",
            "Alert Type": typ, "Severity": sev, "Trigger": trig, "Current Value": b[col_].round(1), "Normal Value": nrm,
            "Deviation": (b[col_] - nrm).round(1), "Days Valid": 7, "Recommended Action": act, "Status": "Active",
            "Confidence": cf if cf is not None else b["sat_conf"].fillna(0.5),
            "Source": "Derived (rule-based) from " + ("IMD gridded rainfall + IMD normal" if col_ == "dev" else "Open-Meteo forecast (Tier 3)" if col_ == "f7" else "MODIS NDVI via Earth Engine"),
            "Source Date": max([d for d in (imd_end, FRESH_RUN.get("MODIS (via GEE)"), FRESH_RUN.get("Open-Meteo")) if d] or [TODAY])}))
    if al:
        DATA["ALERTS"] = pd.concat(al, ignore_index=True).reindex(columns=HEADERS["ALERTS"])
        print("ALERTS:", DATA["ALERTS"].groupby(["Alert Type", "Severity"]).size().to_dict())
    else:
        print("ALERTS: no rule triggered")

    # ---- TRACTOR_OPPORTUNITY: PARTIAL score from available inputs (+ optional business CSV) ----
    ncols = pd.DataFrame(index=M.index)
    gca, nsa = num(M, "Gross Cropped Area Ha"), num(M, "Net Sown Area Ha")
    ncols["Crop Area"] = gca.rank(pct=True) * 100
    ncols["Cropping Intensity"] = _clip((gca / nsa.where(nsa > 0) - 1) * 100)
    hp = pd.Series(comp["Crop Health"].to_numpy(), index=rows["pos"].to_numpy())
    hy = pd.Series(comp["Yield Outlook"].to_numpy(), index=rows["pos"].to_numpy())
    wgt = pd.Series(rows["Sown Area Ha"].to_numpy(), index=rows["pos"].to_numpy())
    def wmean(s):
        ok = s.notna() & wgt.notna()
        return (s[ok] * wgt[ok]).groupby(level=0).sum() / wgt[ok].groupby(level=0).sum()
    ncols["Crop Health"] = wmean(hp).reindex(range(len(M))).to_numpy() if len(hp) else np.nan
    ncols["Crop Economy"] = wmean(hy).reindex(range(len(M))).to_numpy() if len(hy) else np.nan
    for c in ("Farm Prosperity", "Mechanisation", "Historical Tractor Demand", "Replacement Potential"):
        ncols[c] = np.nan
    if os.path.exists(TRACTOR_INPUT_FILE):
        ti = pd.read_csv(TRACTOR_INPUT_FILE)
        ti["_k"] = ti["State"].map(norm) + "|" + ti["District"].map(norm) + "|" + ti["Tehsil"].map(norm)
        ti = ti.drop_duplicates("_k").set_index("_k")
        for c in ("Farm Prosperity", "Mechanisation", "Historical Tractor Demand", "Replacement Potential"):
            if c in ti:
                ncols[c] = base_key.map(ti[c]).to_numpy()
        print(f"  tractor inputs loaded from {TRACTOR_INPUT_FILE}")
    tw = pd.Series(TRACTOR_WEIGHTS)[list(ncols.columns)]
    av = ncols.notna()
    tsum = (av * tw).sum(axis=1)
    score = ((ncols.fillna(0) * tw).sum(axis=1) / tsum.where(tsum > 0)).round(0)
    used = av.apply(lambda r: ", ".join(c for c in ncols.columns if r[c]), axis=1)
    miss = av.apply(lambda r: ", ".join(c for c in ncols.columns if not r[c]), axis=1)
    band = np.where(score >= 75, "HIGH", np.where(score >= 60, "MEDIUM", "LOW"))
    TO = pd.DataFrame({
        "Date": TODAY, "State": M["State"], "District": M["District"], "Tehsil": M["Tehsil"], "Crop Area": gca,
        "Crop Health": ncols["Crop Health"].round(0), "Farm Prosperity": ncols["Farm Prosperity"], "Mechanisation": ncols["Mechanisation"],
        "Historical Tractor Demand": ncols["Historical Tractor Demand"], "Replacement Potential": ncols["Replacement Potential"],
        "Crop Economy": ncols["Crop Economy"].round(0), "Tractor Opportunity Score": score,
        "Opportunity Reason": np.where(score.isna(), "No inputs", [f"{b} - PARTIAL score from: {u}. Missing: {m_}" for b, u, m_ in zip(band, used, miss)]),
        "Source": "Derived (rule-based, PARTIAL) from CoRE Stack cropped area, MODIS health, DES APY outlook" + (" + tractor_inputs.csv" if os.path.exists(TRACTOR_INPUT_FILE) else ""),
        "Source Date": max([d for d in (date(2024, 6, 30), FRESH_RUN.get("MODIS (via GEE)"), FRESH_RUN.get("DES APY")) if d])})
    DATA["TRACTOR_OPPORTUNITY"] = TO.reindex(columns=HEADERS["TRACTOR_OPPORTUNITY"])
    NOTES[("TRACTOR_OPPORTUNITY", "Tractor Opportunity Score")] = ("PARTIAL: only inputs that exist are used (see Opportunity Reason). Add tractor_inputs.csv "
                                                                  "(Farm Prosperity, Mechanisation, Historical Tractor Demand, Replacement Potential, 0-100) for the full score.")
    print(f"TRACTOR_OPPORTUNITY: {len(TO):,} tehsils, score for {int(score.notna().sum()):,}")


try:
    decision_layer()
    rlog("CELL 12", "Decision layer (rule-based)", "", "AGRI_INTELLIGENCE, CROP_HEALTH, CROP_STAGE, ALERTS, TRACTOR", "OK",
         len(DATA["AGRI_INTELLIGENCE"]))
except Exception as e:
    print("Decision layer failed:", type(e).__name__, e)
    rlog("CELL 12", "Decision layer (rule-based)", "", "", "FAILED", msg=f"{type(e).__name__}: {e}")


# ===== CELL 12B : VALIDATION LAYER - record metadata, govt vs satellite checks, reliability score, gap register, AI recovery =====
# Two independent evidence pillars: (1) government records, (2) Google Earth Engine / satellite datasets.
# Satellite values CHECK government values - they never replace them. District figures keep "Source Level = District".
# AI only extracts figures that are written in a fetched text snippet (its quote must match the snippet). Those results
# stay in AI_ESTIMATION_REGISTER and are NEVER copied into the main tabs.
URL_CAT = "https://developers.google.com/earth-engine/datasets/catalog/"
URL_IMD = "https://www.imdpune.gov.in/"
TIER_SCORE = {1: 100, 2: 75, 3: 60, 4: 20, 5: 40}
DTYPE_SCORE = {"Actual": 100, "Satellite-derived": 75, "Model": 60, "Estimated": 50, "Proxy": 35}
AGREE_SCORE = {"Matched": 100, "Partially matched": 60, "Not matched": 20}
VAL_RES = {}      # validation key -> agreement per tehsil (used by RECORD_METADATA)


def _tab(tab):
    """Tab aligned to TEHSIL_MASTER order (row i = tehsil i)."""
    M = DATA["TEHSIL_MASTER"]
    df = DATA.get(tab)
    if df is None or df.empty:
        return pd.DataFrame(index=range(len(M)), columns=HEADERS.get(tab, []))
    return df.reset_index(drop=True) if tab == "TEHSIL_MASTER" else _aligned(df, M)


def _fnum(df, c, n):
    return pd.to_numeric(df[c], errors="coerce").to_numpy(float) if c in df else np.full(n, np.nan)


def _vec(v, n):
    a = np.asarray(v, dtype=float) if v is not None else np.full(n, np.nan)
    return a if a.shape == (n,) else np.full(n, np.nan)


def record_specs():
    imd_end, sat = FRESH_RUN.get("IMD gridded"), FRESH_RUN.get("MODIS (via GEE)")
    g = globals().get("GOV", {})
    S = lambda **k: k
    return [
        S(tab="TEHSIL_MASTER", col="Total Area Ha", p="Total area", u="ha", src="LGD sub-district boundary (bharatlas.com)", url=TEHSIL_SOURCE_URL,
          lvl="Tehsil", dt="Actual (official boundary)", per="LGD 2024", pub=date(2024, 12, 31), maxd=730, tier=1, val=None),
        S(tab="TEHSIL_MASTER", col="Net Sown Area Ha", p="Net sown area", u="ha", src="CoRE Stack LULC (satellite model)", url=g.get("corestack"),
          lvl="Block (summed to tehsil)", dt="Satellite-derived", per="2023-24", pub=date(2024, 6, 30), maxd=730, tier=2, val="nsa"),
        S(tab="TEHSIL_MASTER", col="Cultivable Area Ha", p="Cultivable area", u="ha", src="CoRE Stack cropland + orchard / ESA WorldCover", url=g.get("corestack"),
          lvl="Block / 10 m grid", dt="Proxy (satellite land cover)", per="2023-24", pub=date(2024, 6, 30), maxd=730, tier=2, val=None),
        S(tab="SOIL_INTELLIGENCE", col="Soil Type", p="Soil type", u="", src="Planning Commission agro-climatic zone", url=g.get("agro"),
          lvl="Agro-climatic zone", dt="Actual (zone classification)", per="static", pub=None, maxd=None, tier=1, val=None),
        S(tab="SOIL_INTELLIGENCE", col="pH", p="Soil pH 0-5 cm", u="pH", src="ISRIC SoilGrids 2.0 (Earth Engine)", url="https://www.isric.org/explore/soilgrids",
          lvl="250 m grid", dt="Model", per="2020 release", pub=date(2020, 12, 31), maxd=None, tier=3, val=None),
        S(tab="WATER_IRRIGATION", col="Irrigated Area Ha", p="Irrigated area", u="ha", src="CoRE Stack non-monsoon cropping (PROXY)", url=g.get("corestack"),
          lvl="Block", dt="Proxy (satellite cropping pattern)", per="2023-24", pub=date(2024, 6, 30), maxd=1095, tier=2, val="irr", special="irr"),
        S(tab="WATER_IRRIGATION", col="Water Stress", p="Groundwater stress category", u="", src="CGWB groundwater resource assessment", url=g.get("cgwb"),
          lvl="Block / taluk", dt="Actual (assessment)", per="2020", pub=date(2020, 3, 31), maxd=1825, tier=1, val="gw"),
        S(tab="WATER_IRRIGATION", col="Borewell", p="Borewells", u="count", src="Minor Irrigation Census 2017-18", url=g.get("mi6"),
          lvl="District", dt="Estimated (district split by cropland share)", per="2017-18", pub=date(2018, 3, 31), maxd=3650, tier=1, val=None),
        S(tab="WATER_IRRIGATION", col="Groundwater Trend", p="Groundwater storage trend", u="", src="NASA GLDAS-2.2 CLSM (Earth Engine)",
          url=URL_CAT + "NASA_GLDAS_V022_CLSM_G025_DA1D", lvl="~25 km model grid", dt="Model", per="last 12 vs previous 12 months",
          pub=FRESH_RUN.get("GLDAS (via GEE)"), maxd=180, tier=3, val="gw"),
        S(tab="WEATHER_LIVE", col="Rainfall 30D", p="Rainfall 30 days", u="mm", src="IMD gridded rainfall 0.25 deg", url=URL_IMD,
          lvl="0.25 deg grid (nearest cell)", dt="Actual (gridded observation)", per=f"30 days to {imd_end}", pub=imd_end, maxd=7, tier=1, val="rain", special="wx"),
        S(tab="WEATHER_LIVE", col="Temperature", p="Mean temperature (latest day)", u="deg C", src="IMD gridded temperature 1 deg", url=URL_IMD,
          lvl="1 deg grid (nearest cell)", dt="Actual (gridded observation)", per=str(imd_end), pub=imd_end, maxd=7, tier=1, val="temp", special="wx"),
        S(tab="SATELLITE_INTELLIGENCE", col="NDVI", p="NDVI cropland", u="index", src="MODIS MOD13Q1 (Earth Engine)", url=URL_CAT + "MODIS_061_MOD13Q1",
          lvl="250 m pixels (cropland mean)", dt="Satellite-derived", per=str(sat), pub=sat, maxd=30, tier=3, val="ndvi"),
        S(tab="FARMER_AGRI_ECOSYSTEM", col="Farmers", p="Operational holdings", u="count", src="Agriculture Census 2015-16", url=META.get("agc_url_tehsil", ""),
          lvl="Tehsil", dt="Actual", per="2015-16", pub=date(2016, 3, 31), maxd=3650, tier=1, val=None, special="agc"),
    ]


# ------------------------------------------------------------------ govt vs satellite / independent checks
def _agree(g, i, tight, loose, relative=True, small=None):
    d = i - g
    dp = np.where(np.abs(g) > 0, d / np.where(np.abs(g) > 0, np.abs(g), 1) * 100, np.nan)
    a = np.abs(dp) if relative else np.abs(d)
    res = np.where(np.isnan(d), None, np.where(a <= tight, "Matched", np.where(a <= loose, "Partially matched", "Not matched")))
    if small is not None:                                            # both totals tiny -> compare in absolute terms
        tiny = (np.abs(g) < small[0]) & (np.abs(i) < small[0]) & (np.abs(d) <= small[1])
        res = np.where(tiny, "Matched", res)
    return d, dp, res.astype(object)


def build_validation():
    M = DATA["TEHSIL_MASTER"]
    n = len(M)
    keys = M[["State", "District", "Tehsil"]].reset_index(drop=True)
    W, SAT = _tab("WATER_IRRIGATION"), _tab("SATELLITE_INTELLIGENCE")
    out = []

    def add(vkey, param, g, gsrc, gper, iv, isrc, iper, unit, d, dp, agr, expl):
        agr = np.asarray(agr, dtype=object)
        VAL_RES[vkey] = agr if vkey not in VAL_RES else np.where(pd.isna(VAL_RES[vkey]), agr, VAL_RES[vkey])
        m = np.array([a is not None for a in agr])
        if not m.any():
            return
        rnd = lambda x: np.round(np.asarray(x, dtype=float), 3) if np.asarray(x).dtype.kind == "f" else x
        out.append(pd.DataFrame({
            "State": keys["State"].to_numpy()[m], "District": keys["District"].to_numpy()[m], "Tehsil": keys["Tehsil"].to_numpy()[m],
            "Parameter": param, "Government / Primary Value": rnd(g)[m] if np.ndim(g) else g,
            "Primary Source": np.asarray(gsrc, dtype=object)[m] if np.ndim(gsrc) else gsrc,
            "Primary Period": np.asarray(gper, dtype=object)[m] if np.ndim(gper) else gper,
            "Independent Value": rnd(iv)[m], "Independent Source": isrc, "Independent Period": iper, "Unit": unit,
            "Difference": rnd(d)[m] if d is not None else None, "Difference %": np.round(dp, 1)[m] if dp is not None else None,
            "Agreement": agr[m], "Possible Explanation": np.asarray(expl, dtype=object)[m] if np.ndim(expl) else expl, "Checked On": TODAY}))

    # 1 rainfall: IMD (govt) vs CHIRPS (satellite+gauge), same days
    g, i = _vec(VAL_STORE.get("rain_gov"), n), _vec(VAL_STORE.get("rain_ind"), n)
    if np.isfinite(g).any() and np.isfinite(i).any():
        d, dp, a = _agree(g, i, 20, 40, small=(10, 5))
        add("rain", "Rainfall total (same days)", g, "IMD gridded rainfall 0.25 deg", VAL_STORE.get("rain_period"), i,
            "CHIRPS daily (UCSB, satellite + gauge) via Earth Engine", VAL_STORE.get("rain_period"), "mm", d, dp, a,
            "Different grids and gauge networks (IMD 0.25 deg nearest cell vs CHIRPS ~5 km averaged over ~33 km). IMD value is kept.")
    # 2 temperature: IMD vs ERA5-Land
    g, i = _vec(VAL_STORE.get("temp_gov"), n), _vec(VAL_STORE.get("temp_ind"), n)
    if np.isfinite(g).any() and np.isfinite(i).any():
        d, dp, a = _agree(g, i, 1.5, 3.0, relative=False)
        add("temp", "Mean air temperature (same days)", g, "IMD gridded temperature 1 deg", VAL_STORE.get("temp_period"), i,
            "ERA5-Land reanalysis (ECMWF) via Earth Engine", VAL_STORE.get("temp_period"), "deg C", d, dp, a,
            "IMD 1 deg grid vs ERA5-Land ~11 km model; elevation differences matter in hills. IMD value is kept.")
    # 3 irrigated area: official (DES / Agriculture Census) vs satellite cropping-pattern proxy
    code = W["_irr_src"].astype(str).to_numpy() if "_irr_src" in W else np.full(n, "")
    gov = np.where(np.isin(code, ["DES", "AGC-T", "AGC-D"]), _fnum(W, "Irrigated Area Ha", n), np.nan)
    i = _vec(VAL_STORE.get("irr_proxy"), n)
    if np.isfinite(gov).any() and np.isfinite(i).any():
        d, dp, a = _agree(gov, i, 15, 35)
        gsrc = np.where(code == "DES", f"DES Land Use Statistics {META.get('irr_year', '')} (district split = estimate)",
                        np.where(code == "AGC-T", "Agriculture Census 2015-16 (tehsil)", "Agriculture Census 2015-16 (district split = estimate)"))
        gper = np.where(code == "DES", META.get("irr_year", ""), "2015-16")
        add("irr", "Irrigated area", gov, gsrc, gper, i, "CoRE Stack non-monsoon cropping (satellite PROXY)", "2023-24", "ha", d, dp, a,
            "Different reference periods and definitions (net area irrigated vs satellite rabi/double/triple cropping). "
            "Government value kept; difference recorded for investigation.")
    # 4 net sown area: Agriculture Census vs CoRE Stack satellite
    g, i = _vec(VAL_STORE.get("agc_nsown"), n), _vec(VAL_STORE.get("core_nsa"), n)
    if np.isfinite(g).any() and np.isfinite(i).any():
        d, dp, a = _agree(g, i, 15, 35)
        add("nsa", "Net sown area", g, "Agriculture Census 2015-16 (tehsil, or district split = estimate)", "2015-16", i,
            "CoRE Stack LULC 10 m satellite model", "2023-24", "ha", d, dp, a,
            "8-year gap between reference periods, crop-definition and boundary differences; record the gap, do not overwrite.")
    # 5 groundwater: CGWB stress category vs GLDAS storage trend (direction consistency only)
    ws = W["Water Stress"].astype(object).to_numpy() if "Water Stress" in W else np.full(n, None)
    gt = W["Groundwater Trend"].astype(object).to_numpy() if "Groundwater Trend" in W else np.full(n, None)

    def gwa(s_, t_):
        if s_ not in ("Low", "Medium", "High") or t_ not in ("Rising", "Stable", "Falling"):
            return None
        if (s_ == "High" and t_ == "Falling") or (s_ == "Low" and t_ in ("Rising", "Stable")) or (s_ == "Medium" and t_ in ("Stable", "Falling")):
            return "Matched"
        if (s_ == "Low" and t_ == "Falling") or (s_ == "High" and t_ == "Rising"):
            return "Not matched"
        return "Partially matched"
    a = np.array([gwa(s_, t_) for s_, t_ in zip(ws, gt)], dtype=object)
    if any(x is not None for x in a):
        add("gw", "Groundwater (direction consistency)", ws, "CGWB 2020 stage-of-extraction category", "2020", gt,
            "NASA GLDAS-2.2 storage trend (~25 km model)", "last 12 vs previous 12 months", "category", None, None, a,
            "Different measures (extraction stage vs storage change, 2020 vs recent). Checks direction only - not a level comparison.")
    # 6-8 MODIS vs Sentinel-2 (satellite vs satellite)
    s2 = VAL_STORE.get("s2")
    if isinstance(s2, pd.DataFrame) and len(s2) == n:
        low = (s2["nclear"].to_numpy() < S2_MIN_CLEAR_OBS)
        for vkey, mcol, scol, lab in (("ndvi", "NDVI", "s2ndvi", "NDVI"), ("evi", "EVI", "s2evi", "EVI"), ("ndwi", "NDWI", "s2ndwi", "NDWI")):
            g, i = _fnum(SAT, mcol, n), s2[scol].to_numpy(float)
            if not (np.isfinite(g).any() and np.isfinite(i).any()):
                continue
            d, dp, a = _agree(g, i, 0.05, 0.12, relative=False)
            a = np.where(np.isfinite(g) & np.isfinite(i) & low, "Not comparable", a).astype(object)
            expl = np.where(low, f"Fewer than {S2_MIN_CLEAR_OBS} cloud-free Sentinel-2 observations - not compared (cloud cover).",
                            "MODIS 250 m 16-day composite vs Sentinel-2 10 m median; different dates inside the window and cropland masks (MODIS vs WorldCover).")
            add(vkey, f"{lab} (MODIS vs Sentinel-2)", g, "MODIS (Earth Engine)", str(FRESH_RUN.get("MODIS (via GEE)", "")), i,
                "Sentinel-2 SR harmonized (Earth Engine)", VAL_STORE.get("s2_period"), "index", d, dp, a, expl)
    if not out:
        print("DATA_VALIDATION: no comparison possible this run (needs IMD + Earth Engine steps).")
        return pd.DataFrame(columns=HEADERS["DATA_VALIDATION"])
    V = pd.concat(out, ignore_index=True).reindex(columns=HEADERS["DATA_VALIDATION"])
    print("DATA_VALIDATION:", V.groupby(["Parameter", "Agreement"]).size().unstack(fill_value=0).to_string())
    return V


# ------------------------------------------------------------------ per-record metadata
def _conf(dt, lvl, stale, agr):
    lv = ["Low", "Medium", "High"]
    k = 2 if dt.startswith("Actual") else 1 if dt.startswith(("Satellite", "Estimated", "Model")) else 0
    if dt.startswith("Model") and ("25 km" in lvl or "zone" in lvl.lower()):
        k = 0
    if agr == "Matched" and k == 1:
        k = 2
    if agr == "Not matched":
        k = max(0, k - 1)
    if stale:
        k = max(0, k - 1)
    return lv[k]


def record_metadata():
    M = DATA["TEHSIL_MASTER"]
    n = len(M)
    keys = M[["State", "District", "Tehsil"]].reset_index(drop=True)
    frames = []
    specs = record_specs()
    for sp in specs:
        df = _tab(sp["tab"])
        if sp["col"] not in df:
            continue
        v = df[sp["col"]].reset_index(drop=True)
        has = (v.notna() & ~v.astype(str).isin(["", "None", "nan"])).to_numpy()
        if not has.any():
            continue
        f = lambda x, dt_=object: np.full(n, x, dtype=dt_)
        src, lvl, dt, per, pub, url = f(sp["src"]), f(sp["lvl"]), f(sp["dt"]), f(sp["per"]), f(sp["pub"]), f(sp["url"] or "")
        tier = f(sp["tier"], float)
        sk = sp.get("special")
        if sk == "irr" and "_irr_src" in df:
            code = df["_irr_src"].astype(str).to_numpy()
            iy = META.get("irr_year")
            for c_, vals in {"DES": (f"DES Land Use Statistics {iy or ''}", "District", "Estimated (district split by cropland share)", iy or "",
                                     fy_end(iy) if iy else None, 1, META.get("irr_url", "")),
                             "AGC-T": ("Agriculture Census 2015-16 (tehsil table)", "Tehsil", "Actual", "2015-16", date(2016, 3, 31), 1,
                                       META.get("agc_url_tehsil", "")),
                             "AGC-D": ("Agriculture Census 2015-16 (district table)", "District", "Estimated (district split by cropland share)",
                                       "2015-16", date(2016, 3, 31), 1, META.get("agc_url_district", ""))}.items():
                m = code == c_
                src[m], lvl[m], dt[m], per[m] = vals[0], vals[1], vals[2], vals[3]
                pub[m], tier[m], url[m] = vals[4], vals[5], vals[6]
        if sk == "wx" and "Source" in df:
            fb = df["Source"].astype(str).str.contains("FALLBACK").to_numpy()
            src[fb], lvl[fb], dt[fb], url[fb] = "Open-Meteo (FALLBACK model - IMD had no value)", "~10 km model grid", "Model", "https://open-meteo.com"
            tier[fb], pub[fb] = 3, TODAY - timedelta(days=1)
        if sk == "agc":
            lv = VAL_STORE.get("agc_farmer_lvl")
            if lv is not None and len(lv) == n:
                d_ = lv == "district"
                lvl[d_], dt[d_], url[d_] = "District", "Estimated (district split by cropland share)", META.get("agc_url_district", "")
        agr = VAL_RES.get(sp["val"]) if sp["val"] else None
        agr = agr if agr is not None and len(agr) == n else np.full(n, None, dtype=object)
        age = np.array([(TODAY - p_).days if isinstance(p_, date) else np.nan for p_ in pub], dtype=float)
        mx = sp["maxd"]
        lat = (np.full(n, 100.0) if mx is None else
               np.where(np.isnan(age), 60.0, np.where(age <= mx, 100.0, np.where(age <= mx * STALE_MULTIPLIER, 60.0, 30.0))))
        stale = lat < 100
        status = np.array([("Not comparable (see DATA_VALIDATION)" if a_ == "Not comparable" else a_) if a_ is not None
                           else "Unverified (no independent source this run)" for a_ in agr], dtype=object)
        conf = np.array([_conf(dt[k], lvl[k], stale[k], agr[k]) for k in range(n)], dtype=object)
        fr = pd.DataFrame({
            "State": keys["State"], "District": keys["District"], "Tehsil": keys["Tehsil"], "Tab": sp["tab"], "Parameter": sp["p"],
            "Value": v.to_numpy(dtype=object), "Unit": sp["u"], "Source": src, "Source URL": url, "Source Level": lvl, "Output Level": "Tehsil",
            "Data Type": dt, "Reference Period": per, "Publication Date": pub, "Last Updated": TODAY, "Confidence": conf,
            "Validation Status": status, "_pos": np.arange(n), "_tier": tier, "_lat": lat, "_param": sp["p"]})
        frames.append(fr[has])
    if not frames:
        return pd.DataFrame(), pd.DataFrame(columns=HEADERS["RECORD_METADATA"]), specs
    RM = pd.concat(frames, ignore_index=True)
    RM["_dts"] = RM["Data Type"].map(lambda x: next((s_ for k, s_ in DTYPE_SCORE.items() if str(x).startswith(k)), 35))
    RM["_auth"] = RM["_tier"].map(lambda t_: TIER_SCORE.get(int(t_), 40) if t_ == t_ else 40)
    print(f"RECORD_METADATA: {len(RM):,} records | data types: {RM['Data Type'].str.split(' ').str[0].value_counts().to_dict()}")
    return RM, RM.reindex(columns=HEADERS["RECORD_METADATA"]), specs


# ------------------------------------------------------------------ reliability score (separate from all business scores)
def reliability(RM, specs, V):
    M = DATA["TEHSIL_MASTER"].reset_index(drop=True)
    n = len(M)
    keys = M[["State", "District", "Tehsil"]]
    dims = pd.DataFrame(index=range(n))
    if len(RM):
        g = RM.groupby("_pos")
        dims["Source authority"] = g["_auth"].mean().reindex(range(n))
        dims["Actual vs estimated"] = g["_dts"].mean().reindex(range(n))
        n_exp = len({sp["p"] for sp in specs})
        dims["Latest data availability"] = (g["_lat"].sum() / n_exp).reindex(range(n)).fillna(0.0)    # missing parameters count as 0
        have = g["_param"].nunique().reindex(range(n)).fillna(0).astype(int)
    else:
        for c in ("Source authority", "Actual vs estimated"):
            dims[c] = np.nan
        dims["Latest data availability"], have, n_exp = 0.0, pd.Series(0, index=range(n)), len(specs)
    pos_of = {k: i for i, k in enumerate(key_of(M))}
    agr_s = pd.Series(np.nan, index=range(n))
    cmp_txt = pd.Series(None, index=range(n), dtype=object)
    if len(V):
        vv = V[V["Agreement"].isin(list(AGREE_SCORE))].copy()
        vv["pos"] = key_of(vv).map(pos_of)
        vv["sc"] = vv["Agreement"].map(AGREE_SCORE)
        agr_s = vv.groupby("pos")["sc"].mean().reindex(range(n))
        cnt = vv.groupby(["pos", "Agreement"]).size().unstack(fill_value=0).reindex(range(n))
        cmp_txt = pd.Series([None if r.isna().all() else
                             f"{int(r.sum())} comparisons: {int(r.get('Matched', 0))} matched, {int(r.get('Partially matched', 0))} partial, "
                             f"{int(r.get('Not matched', 0))} not matched" for _, r in cnt.iterrows()], index=range(n))
    dims["Government-satellite agreement"] = agr_s.to_numpy()
    code_ok = ~M["Tehsil Code"].astype(str).isin(["", "None", "nan"])
    uniq = ~M["Tehsil"].astype(str).str.contains(r" \[")
    matched_d = {v for v in NAME_MAP.values() if v} if "NAME_MAP" in globals() else set()
    if matched_d:
        dm = pd.Series([(a_, b_) in matched_d for a_, b_ in zip(M["State"], M["District"])])
        dims["Geographic matching"] = (code_ok * 50 + uniq * 20 + dm * 30).astype(float).to_numpy()
    else:
        dims["Geographic matching"] = ((code_ok * 50 + uniq * 20) / 70 * 100).astype(float).to_numpy()
    w = pd.Series(RELIABILITY_WEIGHTS, dtype=float)[list(dims.columns)]
    ok = dims.notna()
    wsum = (ok * w).sum(axis=1)
    score = ((dims.fillna(0) * w).sum(axis=1) / wsum.where(wsum > 0)).round(0)
    band = np.where(score >= 85, "High reliability", np.where(score >= 70, "Good reliability",
                    np.where(score >= 50, "Moderate reliability", np.where(score.notna(), "Low reliability", None))))
    applied = ok.sum(axis=1).astype(str) + "/5" + ok.apply(lambda r: "" if r.all() else " (not applicable: " + ", ".join(c for c in dims.columns if not r[c]) + ")", axis=1)
    note = np.where(dims["Government-satellite agreement"].isna(),
                    "Validation incomplete: no independent comparison was possible for this tehsil - dimension excluded, not penalised",
                    cmp_txt.to_numpy())
    R = pd.DataFrame({"State": keys["State"], "District": keys["District"], "Tehsil": keys["Tehsil"],
                      "Source Authority": dims["Source authority"].round(0), "Actual vs Estimated": dims["Actual vs estimated"].round(0),
                      "Latest Data Availability": dims["Latest data availability"].round(0),
                      "Govt-Satellite Agreement": dims["Government-satellite agreement"].round(0), "Geographic Matching": dims["Geographic matching"].round(0),
                      "Data Reliability Score": score, "Reliability Band": band, "Dimensions Applied": applied, "Validation Note": note,
                      "Parameters With Data": have.to_numpy(), "Parameters Expected": n_exp})
    print("DATA_RELIABILITY:", pd.Series(band).value_counts().to_dict())
    return R.reindex(columns=HEADERS["DATA_RELIABILITY"])


# ------------------------------------------------------------------ AI-assisted recovery of missing figures (register only)
RECOVERY_TOPICS = [   # parameter, (tab, field) gaps it addresses, Google News query, extraction kind
    ("Current-season sowing area", [("CROP_SOWN", "Sowing Start"), ("CROP_SOWN", "Sowing End"), ("CROP_SOWN", "Delay Days")],
     '"{st}" {season} sowing area hectare', "area"),
    ("Groundwater level (depth to water)", [("WATER_IRRIGATION", "Groundwater Level")],
     '"{st}" groundwater level metres below ground', "depth"),
]
AREA_RX = re.compile(r"(\d[\d,]*(?:\.\d+)?)\s*(lakh|lakhs|crore|thousand)?\s*(hectares?|ha)\b", re.I)
DEPTH_RX = re.compile(r"(\d+(?:\.\d+)?)\s*(m|metres?|meters?|mbgl)\b\s*(?:bgl|below ground)?", re.I)


def _season_now():
    return "kharif" if 5 <= TODAY.month <= 9 else "summer" if TODAY.month in (3, 4) else "rabi"


def _regex_extract(kind, text):
    out = []
    for mt in (AREA_RX if kind == "area" else DEPTH_RX).finditer(text):
        if kind == "depth" and not re.search(r"bgl|below ground|groundwater|water level", text[max(0, mt.start() - 60):mt.end() + 30], re.I):
            continue
        v = float(mt.group(1).replace(",", ""))
        unit = "ha" if kind == "area" else "m below ground"
        if kind == "area":
            v *= {"lakh": 1e5, "lakhs": 1e5, "crore": 1e7, "thousand": 1e3}.get((mt.group(2) or "").lower(), 1)
        out.append((v, unit, text[max(0, mt.start() - 80):mt.end() + 40].strip()))
    return out


def ai_recover_gaps():
    """For gaps with no open dataset: search Google News (state level), extract stated figures with AI (Gemini -> GPT -> Claude)
    or, with no AI key, with a regex. A figure is kept only if its quote is found in the snippet. A second AI model (if a 2nd
    key exists) reviews each figure. Everything stays in AI_ESTIMATION_REGISTER - never copied into the main tabs."""
    cols = HEADERS["AI_ESTIMATION_REGISTER"]
    if not AI_RECOVERY:
        rlog("CELL 12B", "AI recovery", "", "", "SKIPPED", msg="AI_RECOVERY = False")
        return pd.DataFrame(columns=cols)
    season, rows = _season_now(), []
    dist_by_state = {st: {norm(d): d for d in g["District"].unique()} for st, g in T.groupby("State")}
    for param, _gaps, tmpl, kind in RECOVERY_TOPICS:
        for st in sorted(T["State"].unique()):
            nm = st.title().replace("&", "and")
            url = ("https://news.google.com/rss/search?q=" + quote_plus(tmpl.format(st=nm, season=season) + " when:60d")
                   + "&hl=en-IN&gl=IN&ceid=IN:en")
            dn = [k_ for k_ in dist_by_state.get(st, {}) if len(k_) >= 5]
            items = [it for it in fetch_rss(url)[:12] if re.search(r"\d", it["title"] + " " + it["desc"])
                     and (norm(nm) in norm(it["title"] + it["desc"]) or any(k_ in norm(it["title"] + it["desc"]) for k_ in dn))]   # must name this state / a district
            time.sleep(1.0)
            rlog("CELL 12B", "Google News RSS (discovery)", url, f"recover {param} - {st}", "OK" if items else "NO RESULT", len(items))
            if not items:
                continue
            cands = []
            if ai_available():
                snip = "\n".join(f"{k}. {it['title']} -- {it['desc'][:300]}" for k, it in enumerate(items))
                prompt = (f"Numbered Indian news snippets are below. Extract ONLY figures that a snippet explicitly states for '{param}' in {nm}. "
                          'Return JSON {"items":[{"n":<snippet number>,"value":<the number as a plain number>,"unit":"<unit as written>",'
                          '"period":"<season/year/date stated, or null>","level":"state|district|tehsil","district":"<district named, or null>",'
                          '"quote":"<exact words copied from the snippet that contain the number>"}]}. If no snippet states such a figure return {"items":[]}. '
                          "Never estimate, infer, combine numbers or use outside knowledge.\n" + snip)
                got, prov = ai_json(prompt, f"extract {param} - {st}")
                for o in got or []:
                    try:
                        it = items[int(o["n"])]
                        val = float(str(o.get("value")).replace(",", ""))
                    except Exception:
                        continue
                    quote = str(o.get("quote") or "")
                    if len(norm(quote)) < 6 or norm(quote) not in norm(it["title"] + " " + it["desc"]):
                        continue                                           # quote not in the snippet -> rejected
                    dist = dist_by_state.get(st, {}).get(norm(o.get("district") or ""))
                    cands.append([it, val, o.get("unit"), o.get("period"), quote, (o.get("level") or "state").title(), dist, f"AI extraction ({prov})", prov])
            else:
                for it in items:
                    for val, unit, quote in _regex_extract(kind, it["title"] + " " + it["desc"]):
                        cands.append([it, val, unit, None, quote, "State", None, "Regex extraction (no AI key - non-AI fallback)", None])
            review = {}
            first = {c[8] for c in cands if c[8]}
            if cands and ai_available(exclude=first):
                lines = "\n".join(f"{k}. value={c[1]} {c[2] or ''} | quote: {c[4]}" for k, c in enumerate(cands))
                rv, rprov = ai_json('For each numbered line say whether the quote explicitly states that value for "' + param + '" in ' + nm +
                                    '. Return JSON {"items":[{"n":<number>,"supported":true|false}]}. Use only the quote.\n' + lines,
                                    f"review {param} - {st}", exclude=first)
                review = {int(r_["n"]): (bool(r_.get("supported")), rprov) for r_ in rv or [] if str(r_.get("n", "")).isdigit()}
            for k, (it, val, unit, per, quote, lvl, dist, method, prov) in enumerate(cands):
                ok, rprov = review.get(k, (None, None))
                status = ("Unverified - AI-extracted from a news snippet; check the linked source" if ok is None and prov else
                          "Unverified - regex-extracted from a news snippet; check the linked source" if ok is None else
                          f"Second model ({rprov}) confirms the quote states this value - still not an official record" if ok else
                          f"Second model ({rprov}) does NOT confirm - treat as doubtful")
                rows.append({"Date": TODAY, "State": st, "District": dist or "", "Tehsil": "", "Parameter": param, "Extracted Value": val,
                             "Unit": unit, "Reference Period": per, "Evidence (quote)": quote[:300], "Evidence URL": it["link"],
                             "Publisher": it["publisher"], "Publication Date": it["date"], "Method": method,
                             "Confidence": "Medium" if ok else "Low", "Validation Status": status, "Used In Main Tabs": "No - review before use"})
    R = pd.DataFrame(rows).reindex(columns=cols).drop_duplicates(["State", "Parameter", "Extracted Value", "Evidence URL"])
    print(f"AI_ESTIMATION_REGISTER: {len(R):,} candidate figures (state-level news; never copied into main tabs)")
    return R


# ------------------------------------------------------------------ gap register
GAP_SKIP = {"State", "District", "Tehsil", "Tehsil Code", "Source", "Source Date", "Date", "Freshness", "Confidence", "Data Date",
            "Satellite Date", "Satellite Source", "Crop", "Soil Data Date", "Data Year"}
NEXT_SOURCE = {
    "Groundwater Level": "India-WRIS / CGWB station data (registration) or state groundwater department reports",
    "EC": "Soil Health Card portal / soil-testing labs (no open API)", "P": "Soil Health Card portal (no open API)",
    "K": "Soil Health Card portal (no open API)", "Soil Depth": "NBSS&LUP soil survey (no open API)",
    "Water Holding Capacity": "NBSS&LUP soil survey (no open API)", "Soil Moisture": "Earth Engine SMAP (CELL 11) - needs the key",
    "FPOs": "SFAC / NABARD FPO lists", "Warehouses": "WDRA registry", "Cold Storage": "NHB cold-storage directory",
    "Mandis": "Agmarknet / e-NAM market list", "Mills": "State industry registries", "Seed Dealers": "State agriculture licence lists",
    "Fertilizer Dealers": "iFMS / state licence lists", "Other Agri Infrastructure": "State registries",
    "Humidity": "Open-Meteo forecast (rate-limited) - re-run CELL 7 later", "Forecast 7D": "Open-Meteo forecast - re-run CELL 7 later",
    "NDVI": "Earth Engine (CELL 11) - needs the key", "Cloud Cover %": "Sentinel-2 via Earth Engine (CELL 11, RUN_SENTINEL2)",
    "Canal": "DES Land Use Statistics / Agriculture Census (district match)", "Tank": "DES Land Use Statistics / Agriculture Census (district match)",
    "Mills": "State industry registries",
}
TAB_NEED = {
    "CROP_SOWN": "Current-season sowing: state agriculture department reports / Bhuvan-VEDAS (registration)",
    "CROP_STAGE": "Observed sowing dates (state reports) + Sentinel-2 time series", "PEST_DISEASE": "State pest-surveillance reports (no open API)",
    "MARKET_ECONOMICS": "Agmarknet (India Data Portal copy or data.gov.in key)", "YIELD_PRODUCTION": "DES APY (CELL 9)",
    "5_YEAR_HISTORY": "DES APY (CELL 9)", "CROP_CALENDAR": "CROP_SOWN must load first", "EXTREME_EVENTS": "PIB / Google News RSS (CELL 10)",
}


def gap_register(RM):
    M = DATA["TEHSIL_MASTER"].reset_index(drop=True)
    n = len(M)
    rows, parts = [], []
    ai = DATA.get("AI_ESTIMATION_REGISTER")
    rec = {}
    for param, gaps, _, _ in RECOVERY_TOPICS:
        k = int((ai["Parameter"] == param).sum()) if ai is not None and len(ai) else 0
        for g_ in gaps:
            rec[g_] = ("Yes - state-level news search + " + ("AI" if ai_available() else "regex") + " extraction",
                       f"{k} candidate figures in AI_ESTIMATION_REGISTER (unverified)" if k else "No stated figure found") if AI_RECOVERY else ("No", "")
    for tab in TEHSIL_LEVEL:
        df = _tab(tab)
        for c in HEADERS[tab]:
            if c in GAP_SKIP or c in FORMULAS.get(tab, {}):
                continue
            v = df[c] if c in df else pd.Series([None] * n)
            miss = (v.isna() | v.astype(str).isin(["", "None", "nan"])).to_numpy()
            k = int(miss.sum())
            if not k:
                continue
            ra, rr = rec.get((tab, c), ("No", ""))
            base = {"Tab": tab, "Parameter": c, "Recovery Attempted": ra, "Recovery Result": rr,
                    "Next Best Source": NEXT_SOURCE.get(c, "See DATA_LINEAGE / SOURCE_REGISTRY")}
            if k == n:
                rows.append({**base, "State": "(all)", "District": "(all)", "Tehsil": f"ALL {n:,} tehsils", "Gap Type": "Missing - no value for any tehsil",
                             "Rows Affected": n, "Note": "No open source fetched this run (never filled with invented values)."})
            else:
                ix = np.flatnonzero(miss)
                parts.append(pd.DataFrame({**base, "State": M["State"].to_numpy()[ix], "District": M["District"].to_numpy()[ix],
                                           "Tehsil": M["Tehsil"].to_numpy()[ix], "Gap Type": "Missing", "Rows Affected": 1,
                                           "Note": f"{k:,} of {n:,} tehsils missing this field"}))
    for tab in [t[0] for t in SHEETS[:18] if t[0] not in TEHSIL_LEVEL]:
        if DATA.get(tab) is None or DATA[tab].empty:
            rows.append({"State": "(all)", "District": "(all)", "Tehsil": "(tab)", "Tab": tab, "Parameter": "(whole tab)", "Gap Type": "Tab empty",
                         "Rows Affected": 0, "Recovery Attempted": "No", "Next Best Source": TAB_NEED.get(tab, "See SOURCE_REGISTRY")})
    if len(RM):
        for (tab, p_, dtp), g in RM.groupby(["Tab", "Parameter", "Data Type"]):
            if str(dtp).startswith(("Estimated", "Proxy")):
                rows.append({"State": "(all)", "District": "(all)", "Tehsil": "(see RECORD_METADATA)", "Tab": tab, "Parameter": p_,
                             "Gap Type": "Estimate only (district split)" if str(dtp).startswith("Estimated") else "Proxy only (not an official / observed value)",
                             "Rows Affected": len(g), "Latest Source Date": max([d for d in g["Publication Date"] if isinstance(d, date)], default=None),
                             "Recovery Attempted": "No", "Next Best Source": "Tehsil-level official record (state revenue / agriculture department)",
                             "Note": str(dtp)})
        st_ = RM[RM["_lat"] < 100]
        for (tab, p_), g in st_.groupby(["Tab", "Parameter"]):
            rows.append({"State": "(all)", "District": "(all)", "Tehsil": "(see RECORD_METADATA)", "Tab": tab, "Parameter": p_, "Gap Type": "Stale",
                         "Rows Affected": len(g), "Latest Source Date": max([d for d in g["Publication Date"] if isinstance(d, date)], default=None),
                         "Recovery Attempted": "No", "Next Best Source": "Newer release of the same source",
                         "Note": "Source date older than the freshness limit for this parameter"})
    s2 = VAL_STORE.get("s2")
    if isinstance(s2, pd.DataFrame) and len(s2) == n:
        k = int((s2["nclear"] < S2_MIN_CLEAR_OBS).sum())
        if k:
            rows.append({"State": "(all)", "District": "(all)", "Tehsil": "(see DATA_VALIDATION)", "Tab": "SATELLITE_INTELLIGENCE",
                         "Parameter": "Sentinel-2 indicators", "Gap Type": "Incomplete - cloud cover", "Rows Affected": k,
                         "Next Best Source": "Wait for cloud-free scenes / longer window (S2_DAYS)",
                         "Note": f"Fewer than {S2_MIN_CLEAR_OBS} cloud-free observations in {S2_DAYS} days"})
    # boundary reconciliation (LGD)
    nocode = int(M["Tehsil Code"].astype(str).isin(["", "None", "nan"]).sum())
    for gt, k, note in (("Boundary - polygons dropped", META.get("lgd_rows", n) - n, "Layer rows without a name, or exact duplicates"),
                        ("Boundary - duplicate names", META.get("lgd_dup", 0), "Same tehsil name twice in a district: LGD code appended in [ ]"),
                        ("Boundary - missing LGD code", nocode, "Cannot be matched by code"),
                        ("Boundary - coverage vs reference count", max(0, 7092 - n),
                         f"LGD 2024 snapshot has {n:,} tehsils vs your reference 7,092 - newer tehsils are not in the boundary layer")):
        if k and k > 0:
            rows.append({"State": "(all)", "District": "(all)", "Tehsil": "(LGD)", "Tab": "TEHSIL_MASTER", "Parameter": "Tehsil boundary",
                         "Gap Type": gt, "Rows Affected": int(k), "Recovery Attempted": "No",
                         "Next Best Source": "LGD directory download (lgdirectory.gov.in) + state revenue boundaries", "Note": note})
    G = pd.concat([pd.DataFrame(rows)] + parts, ignore_index=True).reindex(columns=HEADERS["DATA_GAP_REGISTER"])
    print(f"DATA_GAP_REGISTER: {len(G):,} rows | {G.groupby('Gap Type')['Rows Affected'].sum().to_dict()}")
    return G


RM_INT, REC_SPECS = pd.DataFrame(), []
for _name, _fn in (("DATA_VALIDATION", build_validation),):
    try:
        DATA[_name] = _fn()
        rlog("CELL 12B", "Validation engine (govt vs satellite)", "", "DATA_VALIDATION", "OK", len(DATA[_name]))
    except Exception as e:
        print(f"{_name} failed: {type(e).__name__}: {e}")
        rlog("CELL 12B", _name, "", "", "FAILED", msg=f"{type(e).__name__}: {e}")
try:
    RM_INT, DATA["RECORD_METADATA"], REC_SPECS = record_metadata()
except Exception as e:
    print("RECORD_METADATA failed:", type(e).__name__, e)
    rlog("CELL 12B", "RECORD_METADATA", "", "", "FAILED", msg=f"{type(e).__name__}: {e}")
for _name, _fn in (("AI_ESTIMATION_REGISTER", ai_recover_gaps), ("DATA_GAP_REGISTER", lambda: gap_register(RM_INT)),
                   ("DATA_RELIABILITY", lambda: reliability(RM_INT, REC_SPECS or record_specs(), DATA["DATA_VALIDATION"]))):
    try:
        DATA[_name] = _fn()
        rlog("CELL 12B", _name, "", "", "OK", len(DATA[_name]))
    except Exception as e:
        print(f"{_name} failed: {type(e).__name__}: {e}")
        rlog("CELL 12B", _name, "", "", "FAILED", msg=f"{type(e).__name__}: {e}")
NOTES[("DATA_RELIABILITY", "Data Reliability Score")] = (
    "Weighted: source authority 25, actual vs estimated 25, latest data 20, govt-satellite agreement 20, geographic matching 10 "
    "(RELIABILITY_WEIGHTS, CELL 2). Dimensions that cannot be assessed are excluded and disclosed, not penalised. Separate from all business scores.")
NOTES[("RECORD_METADATA", "Source Level")] = "Level at which the source publishes the figure. Output Level is always Tehsil: if they differ, the tehsil value is an ESTIMATE."
NOTES[("DATA_VALIDATION", "Agreement")] = "Matched / Partially matched / Not matched within the tolerance for that parameter. Government values are never overwritten by satellite values."
NOTES[("AI_ESTIMATION_REGISTER", "Used In Main Tabs")] = "AI / regex extractions are leads only. They are never copied into the main tabs."


# ===== CELL 13 : BUILD + DOWNLOAD THE WORKBOOK ===============================
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FRESH_POLICY = [  # Source, Parameter, Max Allowed Days, Priority  (edit to your policy)
    ("IMD gridded", "Rainfall / temperature", 4, 1), ("Open-Meteo", "Forecast, elevation", 3, 3),
    ("MODIS (via GEE)", "NDVI / EVI / LST", 20, 1), ("Agmarknet", "Mandi price", 4, 1),
    ("SoilGrids", "Soil (modelled)", 3650, 3), ("LGD", "Tehsil boundaries", 365, 1),
    ("Soil Health Card", "Soil nutrients (not auto-fetched)", 730, 2),
    ("CGWB (via bharatlas)", "Groundwater stage of extraction", 1825, 2),
    ("MI Census (via bharatlas)", "Wells by district", 3650, 3),
    ("CoRE Stack (via bharatlas)", "Cropping intensity (blocks)", 400, 2),
    ("India-WRIS", "Groundwater LEVEL (not auto-fetched)", 180, 2),
    ("State Agri Dept", "Sowing report (not auto-fetched)", 30, 1),
    ("Agriculture Census", "Holdings (Agriculture Census 2015-16)", 3650, 3),
    ("Sentinel-2 (via GEE)", "NDVI / EVI / NDWI 10 m, cloud cover", 20, 1),
    ("GLDAS (via GEE)", "Groundwater storage trend (model)", 180, 3),
    ("DES APY", "District area / production / yield", 730, 1)]

REGISTRY = [  # best understanding - VERIFY each before relying on it
    ("LGD via bharatlas.com", "Tehsil list + boundaries (6,471 polygons)", "India", "Yes (parquet)", "Ongoing", "No", 1, "High", "Auto"),
    ("IMD gridded (imdlib)", "Rainfall / temperature grids", "India", "Yes", "Daily", "Yes", 1, "High", "Auto"),
    ("Open-Meteo", "Forecast, elevation", "Global", "Yes", "Hourly", "Yes", 3, "Medium", "Auto"),
    ("Agmarknet (data.gov.in)", "Mandi prices", "India", "Yes (key)", "Daily", "Limited", 1, "High", "Auto (key)"),
    ("MODIS / SMAP / WorldCover (Earth Engine)", "Vegetation, LST, cropland", "Global", "Yes (login)", "8-16 days", "Yes", 1, "High", "Auto (login)"),
    ("ISRIC SoilGrids (Earth Engine)", "Soil (modelled)", "Global", "Yes (login)", "Static", "No", 3, "Medium", "Auto (login)"),
    ("ISRO VEDAS / Bhuvan / MOSDAC", "Crop, LULC, satellite", "India", "No (registration)", "Varies", "Yes", 1, "High", "Manual"),
    ("Soil Health Card portal", "Soil nutrients", "India", "No", "Periodic", "Yes", 2, "Medium", "Manual"),
    ("CGWB groundwater assessment (via bharatlas)", "Stage of extraction (block/taluk, 2020)", "India", "Yes (parquet)", "Multi-year", "No", 2, "High", "Auto"),
    ("Minor Irrigation Census (via bharatlas)", "Wells by district (2017-18)", "India", "Yes (parquet)", "Census", "No", 3, "High", "Auto"),
    ("CoRE Stack (via bharatlas)", "Cropping intensity per block (2023-24)", "India", "Yes (parquet)", "Yearly", "No", 2, "Medium", "Auto"),
    ("Planning Commission agro-climatic zones (via bharatlas)", "Zone per tehsil", "India", "Yes (parquet)", "Static", "No", 3, "Medium", "Auto"),
    ("India-WRIS / CGWB water LEVEL data", "Groundwater level", "India", "No", "Quarterly", "Yes", 2, "High", "Manual"),
    ("State Agriculture Depts", "Sowing / pest reports", "State", "No", "Weekly", "Yes", 1, "High", "Manual"),
    ("PIB RSS + Google News RSS", "News headlines (context only)", "India", "Yes (RSS)", "Continuous", "No", 4, "Low - unverified", "Auto"),
    ("Agriculture Census (India Data Portal)", "Holdings / farmers, irrigated + net sown area", "India", "Yes (CKAN)", "5-yearly", "Yes", 3, "High", "Auto"),
    ("Sentinel-2 SR harmonized (Earth Engine)", "NDVI / EVI / NDWI 10 m, cloud cover", "Global", "Yes (login)", "5 days", "Yes", 1, "High", "Auto (login)"),
    ("CHIRPS daily (Earth Engine)", "Independent rainfall check of IMD", "Global", "Yes (login)", "Daily (~3-6 wk lag)", "Yes", 3, "Medium", "Auto (login)"),
    ("ERA5-Land (Earth Engine)", "Independent temperature check, soil moisture", "Global", "Yes (login)", "Daily / monthly", "Yes", 3, "Medium", "Auto (login)"),
    ("NASA GLDAS-2.2 (Earth Engine)", "Groundwater storage trend", "Global", "Yes (login)", "Daily (~2 mo lag)", "Yes", 3, "Medium", "Auto (login)"),
    ("Gemini / GPT / Claude APIs", "Text extraction + tagging only (never measurements)", "n/a", "Yes (key)", "On demand", "No", 4, "Low - unverified", "Auto (key)"),
    ("Soil Health Card / NBSS&LUP", "Soil-test values, soil survey", "India", "No", "Periodic", "Yes", 1, "High", "Manual"),
    ("VAHAN / FADA / OEM dealer records", "Tractor registrations / sales", "India", "No (access)", "Monthly", "Yes", 1, "High", "Manual (tractor_inputs.csv)")]

DROPDOWNS = {
    "ALERTS": {"Severity": ["Red", "Orange", "Yellow", "Green"], "Status": ["Active", "Resolved", "Expired"]},
    "AGRI_INTELLIGENCE": {"Risk": ["Low", "Medium", "High"], "Opportunity": ["Low", "Medium", "High"]},
}


def to_cell(v):
    if v is None or v is pd.NaT:
        return None
    if isinstance(v, float) and math.isnan(v):
        return None
    if isinstance(v, pd.Timestamp):
        return v.date()
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.floating,)):
        return None if np.isnan(v) else float(v)
    return v


HEADER_ROW = 4        # rows 1-3 of every tab hold a short write-up; the column headers are on row 4

TAB_INFO = {  # tab -> (what it is, how to read it)
    "TEHSIL_MASTER": ("One row per tehsil (sub-district)", "Identity, location, area and land-use baseline. Net sown / gross cropped / cultivable areas are satellite-based (CoRE Stack), not revenue records."),
    "SOIL_INTELLIGENCE": ("Soil of each tehsil", "Soil Type is the agro-climatic-zone dominant soil (coarse). pH, organic carbon, nitrogen and texture are modelled (SoilGrids) on cropland pixels - not lab data."),
    "WATER_IRRIGATION": ("Irrigation and groundwater", "Irrigated/Canal/Tank = district official figures split to tehsils (estimates) or a satellite proxy where the district did not match. Water Stress = CGWB 2020. Groundwater Trend is a coarse (~25 km) model signal."),
    "WEATHER_LIVE": ("Latest rainfall, temperature and 7-day outlook", "Rainfall/temperature = IMD gridded observations (rows marked FALLBACK use Open-Meteo). Deviation compares the last 30 days with the 10-year IMD normal. Forecast is a model at district centre."),
    "SATELLITE_INTELLIGENCE": ("Vegetation, water and heat from satellites", "MODIS values averaged over cropland pixels of the tehsil. Confidence = share of ~30 cropland pixels available (coverage score, not statistics)."),
    "CROP_SOWN": ("Sown area by crop and season", "Official DES district area (latest reported year) split to tehsils by cropland share = ESTIMATE. Sowing % = area vs 5-year mean. Not live sowing progress."),
    "CROP_STAGE": ("Expected crop growth stage right now", "CALENDAR-BASED: reference sowing windows + typical durations tell which crops should be in which stage today (confidence 0.3). Not observed - add real sowing dates to improve."),
    "CROP_HEALTH": ("Satellite and weather health scores per tehsil and crop", "Scores 0-100 from MODIS NDVI / moisture / temperature and IMD rainfall (tehsil-wide cropland, not crop-specific). Status: Good >= 75, Watch >= 55, else Stressed."),
    "PEST_DISEASE": ("Pest and disease signals", "UNVERIFIED news headlines that mention pests or diseases (no open surveillance data exists). Treat as leads to check, not measurements."),
    "YIELD_PRODUCTION": ("Official yield and production by district and crop", "Latest reported year (DES APY) versus the mean of up to 5 earlier years. District level; reported figures, not a forecast."),
    "MARKET_ECONOMICS": ("Mandi prices and arrivals", "Agmarknet daily data as published on the India Data Portal (no key needed): district average of reporting mandis over the last 7 reported days, Rs per quintal, vs the same dates in earlier years. With a data.gov.in key the live feed is used instead."),
    "5_YEAR_HISTORY": ("Yield, area and production history", "DES APY district data; the latest official year is shown in the cell note on the Date header. Newer years appear automatically once published."),
    "CROP_CALENDAR": ("Typical crop calendar by state, crop and season", "REFERENCE DEFAULTS (typical national windows and durations, editable in CELL 2) - not tehsil-specific retrieved data."),
    "FARMER_AGRI_ECOSYSTEM": ("Farmers and agri infrastructure", "Operational holdings, small/marginal share and average holding from the Agriculture Census 2015-16 (tehsil table where names matched, else district split = estimate). FPOs, warehouses, cold stores, dealers have no open source."),
    "EXTREME_EVENTS": ("Recent weather, pest and policy news signals", "Headlines from PIB / Google News RSS: UNVERIFIED context, not measurements. Optional AI tagging marks event type, severity, crop and district."),
    "AGRI_INTELLIGENCE": ("Tehsil x crop opportunity and risk score", "Rule-based 0-100 scores from NDVI, rainfall, groundwater stress, soil, yield trend and mandi price trend. Weights are in SCORE_WEIGHTS (CELL 2); a component with no data is skipped."),
    "ALERTS": ("Active rule-based alerts", "Drought, crop stress, heavy rain and positive signals. Each row states its trigger and thresholds; severity colours: Red, Orange, Yellow, Green."),
    "TRACTOR_OPPORTUNITY": ("Tractor opportunity score (PARTIAL)", "Uses only inputs that exist (area, cropping intensity, crop health, yield trend). Add tractor_inputs.csv for demand, prosperity, mechanisation and replacement to complete it."),
    "DATA_FRESHNESS_CONTROL": ("Is each source fresh enough?", "Age in days against the maximum allowed; VALID / STALE / EXPIRED decides whether a source feeds live scores. Edit the limits in column F."),
    "SOURCE_REGISTRY": ("Where data comes from", "Every source, whether it has an open API, how often it updates and how it is accessed. Manual = no open API."),
    "RECORD_METADATA": ("Metadata for every key tehsil value", "One row per tehsil x key parameter: value, source, source URL, SOURCE LEVEL vs OUTPUT LEVEL, data type (Actual / Estimated / Satellite-derived / Model / Proxy), reference period, publication date, confidence and validation status. A district figure shown at tehsil level is always marked Estimated."),
    "DATA_VALIDATION": ("Government vs satellite / independent cross-checks", "IMD vs CHIRPS rainfall and ERA5-Land temperature (same days), official irrigated / net sown area vs satellite, CGWB stress vs GLDAS trend, MODIS vs Sentinel-2. Government values are never replaced - the difference is recorded with a likely explanation."),
    "DATA_RELIABILITY": ("Data Reliability Score per tehsil (separate from business scores)", "Source authority 25, actual vs estimated 25, latest data 20, govt-satellite agreement 20, geographic matching 10. Bands: 85-100 High, 70-84 Good, 50-69 Moderate, below 50 Low. Dimensions that cannot be assessed are excluded and disclosed."),
    "DATA_GAP_REGISTER": ("Missing, stale, estimate-only and boundary gaps", "Per tehsil where a field is partly missing; one summary row where a field is missing everywhere, estimate/proxy only, or stale. Shows whether recovery was attempted and the next best source."),
    "SOURCE_RECOVERY_LOG": ("Every source / API attempt of this run", "Step, source, endpoint, purpose, status (OK / PARTIAL / FAILED / SKIPPED / NO RESULT), records and fallback used. Read this first when a tab is empty."),
    "AI_ESTIMATION_REGISTER": ("Figures extracted from news text by AI (or regex)", "Leads only: each value has its exact quote, link, publisher and method (Gemini / GPT / Claude / regex). A second model reviews when a 2nd key exists. NEVER copied into the main tabs."),
    "DATA_LINEAGE": ("Per-field source, date and coverage", "One line per field: source, source date or vintage, authenticity tier and how many rows are filled. NOT FETCHED marks fields without an open source."),
}


def banner(ws, name, ncols, color, line3=""):
    title, desc = TAB_INFO.get(name, (name, ""))
    last = get_column_letter(min(max(ncols, 8), 14))
    ws["A1"] = f"{name.replace('_', ' ')}  -  {title}"
    ws["A2"] = desc
    ws["A3"] = line3
    spec = {1: (13, True, 24, color), 2: (10, False, 42, "F3F6F4"), 3: (9, False, 30, "F3F6F4")}
    for r, (size, bold, h, fill) in spec.items():
        ws.merge_cells(f"A{r}:{last}{r}")
        c = ws[f"A{r}"]
        c.font = Font(name=FONT, size=size, bold=bold, color="FFFFFF" if r == 1 else "000000")
        c.fill = PatternFill("solid", fgColor=fill)
        c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[r].height = h


def tab_stats(name):
    df = DATA.get(name)
    if df is None or df.empty:
        return 0, 0
    keys = {"State", "District", "Tehsil", "Tehsil Code", "Source", "Source Date"}
    return len(df), int(df.drop(columns=[c for c in df.columns if c in keys], errors="ignore").notna().any(axis=1).sum())


def source_line(name):
    """Short 'sources + latest date + missing fields' line for the banner of a data tab."""
    df = DATA.get(name)
    srcs, dates, missing = [], [], []
    for t, fields, source, vintage, grade, method, asof in lineage_defs():
        if t != name:
            continue
        ok = df is not None and any(f in df and df[f].notna().any() for f in fields)
        if ok:
            srcs.append(_short(source))
            if asof:
                dates.append(asof)
        else:
            missing += [f for f in fields if f in HEADERS.get(name, [])]
    txt = "Sources: " + ("; ".join(dict.fromkeys(srcs)) if srcs else "none loaded")
    if dates:
        txt += f"  |  Latest source date: {max(dates):%Y-%m-%d}"
    if missing:
        txt += "  |  Not available: " + ", ".join(dict.fromkeys(missing))[:140]
    return txt[:420]


def write_readme(wb):
    ws = wb.create_sheet("README", 0)
    ws.sheet_properties.tabColor = "000000"
    ws["A1"] = f"AGRI INTELLIGENCE - ALL-INDIA TEHSIL WORKBOOK   |   built {TODAY:%Y-%m-%d}"
    ws["A2"] = ("How to read: every tab starts with a 3-line write-up (what it is, how to read it, sources and latest source date), column headers are on row 4. "
                "'Source' names the source and 'Source Date' is the latest source date in that row (year-only sources are shown as year-end). "
                "ESTIMATE = official district figure split to tehsils by cropland share. PROXY / REFERENCE DEFAULT = not retrieved data. "
                "Tier 1 = Government of India record, 2 = government-linked model, 3 = global/third-party fallback, 4 = news (context only), 5 = derived here.")
    for r, h in ((1, 24), (2, 78)):
        ws.merge_cells(f"A{r}:F{r}")
        ws[f"A{r}"].alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[r].height = h
    ws["A1"].font = Font(name=FONT, size=14, bold=True)
    ws["A2"].font = Font(name=FONT, size=10)
    hdr = ["Tab", "What it contains", "How to read it", "Rows", "Rows with data", "Sources and latest date"]
    for c, h in enumerate(hdr, 1):
        cell = ws.cell(row=4, column=c, value=h)
        cell.font = Font(name=FONT, bold=True, color="FFFFFF", size=10)
        cell.fill = PatternFill("solid", fgColor="455A64")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    r = 5
    for name, group, headers in SHEETS + [("DATA_LINEAGE", "Control", [])]:
        title, desc = TAB_INFO.get(name, (name, ""))
        n, f = tab_stats(name) if name not in ("DATA_FRESHNESS_CONTROL", "SOURCE_REGISTRY", "DATA_LINEAGE") else ("", "")
        if name in SPLIT_OUT:
            title = title + "  [in the separate validation Sheet - see CELL 14 links]"
        line = banner_line(name) if name in HEADERS else ""
        for c, v in enumerate([name, title, desc, n, f, line], 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = Font(name=FONT, size=9, bold=(c == 1))
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=r, column=1).fill = PatternFill("solid", fgColor=GROUPS[group])
        ws.cell(row=r, column=1).font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
        ws.row_dimensions[r].height = 62
        r += 1
    for col, w in zip("ABCDEF", (26, 34, 70, 9, 12, 70)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A5"


def style_header(ws, headers, color):
    fill = PatternFill("solid", fgColor=color)
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=HEADER_ROW, column=c, value=h)
        cell.font = Font(name=FONT, bold=True, color="FFFFFF", size=10)
        cell.fill, cell.border = fill, BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(c)].width = max(14, len(h) + 4)
    ws.row_dimensions[HEADER_ROW].height = 32
    ws.freeze_panes = f"A{HEADER_ROW + 1}"
    ws.auto_filter.ref = f"A{HEADER_ROW}:{get_column_letter(len(headers))}{HEADER_ROW}"


def _clean(v):
    if v is None or v is pd.NaT or v is pd.NA:
        return None
    if isinstance(v, float) and v != v:
        return None
    if isinstance(v, pd.Timestamp):
        return v.date()
    return v


def write_data_sheet(ws, name, headers):
    """Fast writer: row numbers tracked here (ws.max_row rescans every cell = O(n^2)), default font set once."""
    df = DATA[name].reindex(columns=headers)
    if df.empty:
        return
    letters = {h: get_column_letter(i + 1) for i, h in enumerate(headers)}
    fm = FORMULAS.get(name, {})
    pieces = {h: re.split(r"\[([^\]]+)\]", t) for h, t in fm.items()}
    fcols = [(i, h) for i, h in enumerate(headers) if h in fm]
    date_cols = [i for i, h in enumerate(headers)
                 if any(isinstance(v, date) for v in df[h].dropna().head(20))]
    r = HEADER_ROW + 1
    for rec in df.itertuples(index=False, name=None):
        row = [_clean(v) for v in rec]
        for i, h in fcols:
            if row[i] is None:
                row[i] = "=" + "".join(p if k % 2 == 0 else f"{letters[p]}{r}" for k, p in enumerate(pieces[h]))
        ws.append(row)
        r += 1
    for i in date_cols:                                   # number format only for date columns
        for rr in range(HEADER_ROW + 1, r):
            ws.cell(row=rr, column=i + 1).number_format = "yyyy-mm-dd"


def write_freshness(ws, headers):
    ws["L1"], ws["M1"] = "Stale multiplier", STALE_MULTIPLIER
    ws["L1"].font = Font(name=FONT, bold=True, size=10)
    ws["M1"].font = Font(name=FONT, color="0000FF", bold=True, size=10)
    ws["M1"].fill = PatternFill("solid", fgColor="FFFF00")
    ws["M1"].comment = Comment("Assumption (editable): VALID if Age <= Max Allowed Days; "
                               "STALE if Age <= Max x this; otherwise EXPIRED.", "Agri Engine")
    ws.column_dimensions["L"].width = 18
    n = max(500, len(FRESH_POLICY) + 50)
    for i in range(n):
        r = i + HEADER_ROW + 1
        if i < len(FRESH_POLICY):
            s, p, mx, pr = FRESH_POLICY[i]
            ws[f"A{r}"], ws[f"B{r}"], ws[f"F{r}"], ws[f"I{r}"] = s, p, mx, pr
            if s in FRESH_RUN:
                ws[f"C{r}"], ws[f"D{r}"] = FRESH_RUN[s], TODAY
        ws[f"E{r}"] = f'=IF(OR(C{r}="",D{r}=""),"",D{r}-C{r})'
        ws[f"G{r}"] = (f'=IF(OR(E{r}="",F{r}=""),"",IF(E{r}<=F{r},"VALID",'
                       f'IF(E{r}<=F{r}*$M$1,"STALE","EXPIRED")))')
        ws[f"H{r}"] = f'=IF(G{r}="","",IF(G{r}="VALID","YES","NO"))'
        for col in "CD":
            ws[f"{col}{r}"].number_format = "yyyy-mm-dd"
        for col in "ABCDEFGHIJ":
            ws[f"{col}{r}"].font = Font(name=FONT, size=10)
    for text, color in (("VALID", "C6EFCE"), ("STALE", "FFEB9C"), ("EXPIRED", "FFC7CE")):
        ws.conditional_formatting.add(f"G{HEADER_ROW + 1}:G{n + HEADER_ROW}", CellIsRule(
            operator="equal", formula=[f'"{text}"'], fill=PatternFill("solid", bgColor=color)))


def write_registry(ws, headers):
    upd = {"IMD": "IMD gridded", "Open-Meteo": "Open-Meteo", "Agmarknet": "Agmarknet", "MODIS": "MODIS (via GEE)"}
    for row in REGISTRY:
        last = next((FRESH_RUN[v] for k, v in upd.items() if k in row[0] and v in FRESH_RUN), None)
        ws.append(list(row[:9]) + [last])
    ws.append([])
    ws.append(["Note: 'Access Status' Auto = fetched by this notebook; Manual = no open API, add data yourself. "
               "Availability is best understanding - verify."])


LINEAGE_COLS = ["Tab", "Field", "Source", "Source Date / Vintage", "Authenticity Grade",
                "Method / Notes", "Rows Filled", "Total Rows", "Coverage %", "Priority Tier"]
TIER = {1: "1 - Government of India record / assessment", 2: "2 - Government-linked / public-interest model",
        3: "3 - Global agency / third-party model (FALLBACK)", 4: "4 - News signal (context only)",
        5: "5 - Derived here / reference default (not source data)"}


ASOF_RULES = [  # (text found in the source name (lower case), function giving the source's as-of DATE; year-only sources = year end)
    ("land use statistics", lambda: fy_end(META["irr_year"]) if META.get("irr_year") else None),
    ("lgd sub-district", lambda: date(2024, 12, 31)),
    ("core stack", lambda: date(2024, 6, 30)),
    ("cgwb", lambda: date(2020, 3, 31)),
    ("minor irrigation census", lambda: date(2018, 3, 31)),
    ("imd gridded", lambda: FRESH_RUN.get("IMD gridded")),
    ("gldas", lambda: FRESH_RUN.get("GLDAS (via GEE)")),
    ("open-meteo", lambda: FRESH_RUN.get("Open-Meteo")),
    ("agmarknet", lambda: FRESH_RUN.get("Agmarknet")),
    ("modis", lambda: FRESH_RUN.get("MODIS (via GEE)")),
    ("soilgrids", lambda: date(2020, 12, 31)),
    ("worldcover", lambda: date(2021, 12, 31)),
    ("des apy", lambda: FRESH_RUN.get("DES APY")),
    ("agriculture census", lambda: date(2016, 3, 31)),
    ("rss", lambda: FRESH_RUN.get("News RSS")),
    ("sentinel-2", lambda: FRESH_RUN.get("Sentinel-2 (via GEE)")),
]


def asof_date(source):
    t = source.lower()
    for key, fn in ASOF_RULES:
        if key in t:
            return fn()
    return None


def lineage_defs():
    """(tab, fields, source, vintage text, grade, method, as-of DATE or None)"""
    return [tuple(t) + (asof_date(t[2]),) for t in _lineage_base()]


def _lineage_base():
    imd_end = FRESH_RUN.get("IMD gridded")
    return [
        ("TEHSIL_MASTER", ["Tehsil Code", "Latitude", "Longitude", "Total Area Ha"], "LGD sub-district boundaries via bharatlas.com",
         "LGD 2024 snapshot", "Official boundary record", "Lat/Lon = representative point of the polygon; area from polygon (equal-area projection)."),
        ("TEHSIL_MASTER", ["Agro Climatic Zone"], "Planning Commission agro-climatic zones via bharatlas.com",
         "Static delineation (undated)", "Official classification", "Zone polygon containing the tehsil centre point."),
        ("TEHSIL_MASTER", ["Net Sown Area Ha", "Gross Cropped Area Ha"], "CoRE Stack land-use / cropping intensity via bharatlas.com",
         "2023-24", "Satellite model (NOT revenue records)", "Single + double + triple cropped area of blocks whose centre lies in the tehsil."),
        ("TEHSIL_MASTER", ["Elevation"], "Copernicus DEM via Open-Meteo", "Static (DEM)", "Satellite model", "Elevation at the centre point."),
        ("WATER_IRRIGATION", ["Water Stress"], "CGWB groundwater resource assessment via bharatlas.com",
         "2020 assessment", "Government assessment (block/taluk)", "Safe=Low, Semi-Critical=Medium, Critical/Over-Exploited=High."),
        ("WATER_IRRIGATION", ["Open Well", "Borewell"], "Minor Irrigation Census via bharatlas.com",
         "2017-18", "Official census (DISTRICT totals split to tehsils - ESTIMATE)", "Split by cropland share; Borewell = shallow + medium + deep tubewells."),
        ("WATER_IRRIGATION", ["Canal", "Tank"], "DES Land Use Statistics 'Sources of Irrigation' via India Data Portal",
         META.get("irr_year", "NOT FETCHED"), "Official statistics (DISTRICT split to tehsils - ESTIMATE)", "Net area irrigated by source, split by cropland share."),
        ("WATER_IRRIGATION", ["Irrigated Area Ha", "Rainfed Area Ha"], "DES Land Use Statistics (else CoRE Stack cropping-pattern PROXY)",
         META.get("irr_year", "2023-24 (CoRE proxy)"), "Official statistics where district matched, else satellite proxy", "See cell notes; proxy = non-monsoon cropping."),
        ("WATER_IRRIGATION", ["Reservoir"], "CoRE Stack perennial water bodies via bharatlas.com", "2023-24",
         "Satellite model", "Area (ha) of perennial surface water."),
        ("WATER_IRRIGATION", ["Groundwater Level"], "CGWB / India-WRIS station data", "NOT FETCHED",
         "No open API", "Needs India-WRIS / CGWB station time-series (depth to water)."),
        ("WATER_IRRIGATION", ["Groundwater Trend"], "NASA GLDAS-2.2 CLSM (GRACE-assimilated) via Earth Engine",
         str(FRESH_RUN.get("GLDAS (via GEE)", "NOT FETCHED")), "Third-party model (regional ~25 km)", "Storage change last 12 vs previous 12 months."),
        ("CROP_CALENDAR", ["Ideal Sowing Start", "Germination", "Typical Duration"], "Reference default (typical national calendar)",
         "typical windows (assumption)", "Reference default (NOT retrieved data)", "Editable IDEAL_SOWING / CROP_DURATION."),
        ("FARMER_AGRI_ECOSYSTEM", ["Farmers", "Small/Marginal Farmers %", "Average Holding"], "Agriculture Census 2015-16 via India Data Portal",
         "2015-16", "Official census (DISTRICT, split to tehsils - ESTIMATE)", "Holdings by size class."),
        ("FARMER_AGRI_ECOSYSTEM", ["FPOs", "Warehouses", "Cold Storage", "Mandis", "Mills", "Seed Dealers", "Fertilizer Dealers"],
         "No open source (SFAC / WDRA / NHB / state lists)", "NOT FETCHED", "No open API", "Add from registries when available."),
        ("CROP_SOWN", ["Sowing %"], "DES APY (area as % of 5-year mean)", META.get("apy_years", "NOT FETCHED"),
         "Official statistics (DISTRICT, repeated per tehsil)", "Latest reported area vs mean of up to 5 earlier years."),
        ("CROP_SOWN", ["Ideal Sowing Window"], "Reference default (typical national window)", "typical windows (assumption)",
         "Reference default (NOT retrieved data)", "State/variety specific windows differ."),
        ("CROP_SOWN", ["Sowing Start", "Sowing End", "Delay Days"], "State agriculture dept. sowing reports / Bhuvan-VEDAS", "NOT FETCHED",
         "No open API", "Needs current-season tehsil sowing progress."),
        ("CROP_SOWN", ["Confidence"], "Computed here (rule-based)", f"computed {TODAY}", "Derived data-quality score", "0.9 official district value; 0.5 tehsil estimate."),
        ("SATELLITE_INTELLIGENCE", ["Cloud Cover %"], "Sentinel-2 SR harmonized via Earth Engine (SCL cloud classes)",
         str(FRESH_RUN.get("Sentinel-2 (via GEE)", "NOT FETCHED")), "Satellite observations", f"Share of cropland observations flagged cloud/shadow, last {S2_DAYS} days."),
        ("SATELLITE_INTELLIGENCE", ["Confidence"], "Computed here from cropland-pixel coverage", str(TODAY),
         "Derived data-quality score", "Share of 30 cropland 250 m pixels available."),
        ("SOIL_INTELLIGENCE", ["Confidence"], "Computed here from cropland-pixel coverage", str(TODAY),
         "Derived data-quality score", "Share of 30 cropland 250 m pixels available."),
        ("CROP_HEALTH", ["Overall Crop Health"], "Derived in this notebook (rule-based) from MODIS/SMAP via Earth Engine and IMD", str(TODAY),
         "Derived score (rule-based)", "Tehsil-wide cropland signal."),
        ("CROP_STAGE", ["Expected Stage"], "Reference default calendar (derived here)", "calendar-based (assumption)",
         "Reference default (NOT observed)", "Needs observed sowing dates for a real estimate."),
        ("PEST_DISEASE", ["Pest/Disease"], "PIB RSS / Google News RSS headlines (unverified)", str(TODAY),
         "UNVERIFIED news signal", "Headlines mentioning pests/diseases."),
        ("AGRI_INTELLIGENCE", ["Overall Agri Score"], "Derived in this notebook (rule-based) from IMD, MODIS, CGWB, soil, DES APY", str(TODAY),
         "Derived score (rule-based)", "Weights in SCORE_WEIGHTS; market score omitted (no price data)."),
        ("ALERTS", ["Alert Type"], "Derived in this notebook (rule-based) from IMD, MODIS, forecast", str(TODAY),
         "Derived alert (rule-based)", "Thresholds listed in CELL 12."),
        ("TRACTOR_OPPORTUNITY", ["Tractor Opportunity Score"], "Derived in this notebook (PARTIAL: area, intensity, health, economy)", str(TODAY),
         "Derived score (rule-based, partial)", "Business inputs need your data (tractor_inputs.csv)."),
        ("WEATHER_LIVE", ["Rainfall mm", "Rainfall 7D", "Rainfall 30D", "Dry Days", "Temperature"], "IMD gridded rainfall 0.25deg / temperature 1deg",
         f"up to {imd_end}" if imd_end else "NOT FETCHED", "Government observations (gridded)", "Nearest grid cell to the tehsil centre."),
        ("WEATHER_LIVE", ["Normal Rainfall mm", "Rainfall Deviation %"], "IMD gridded rainfall archive",
         META.get("imd_years", "NOT FETCHED"), "Government observations (derived)", "Mean 30-day total for the same calendar window."),
        ("WEATHER_LIVE", ["Forecast 7D", "Humidity", "Extreme Weather Alert"], "Open-Meteo forecast model",
         str(FRESH_RUN["Open-Meteo"]) if "Open-Meteo" in FRESH_RUN else "NOT FETCHED", "Third-party model (not IMD)", "District-centre value repeated for each tehsil."),
        ("MARKET_ECONOMICS", ["Min Price", "Modal Price", "Max Price"], "Agmarknet via India Data Portal / data.gov.in",
         f"latest arrival {FRESH_RUN['Agmarknet']}" if "Agmarknet" in FRESH_RUN else "NOT FETCHED (no API key)", "Official market records (district average)", "Rs/quintal."),
        ("SATELLITE_INTELLIGENCE", ["NDVI", "NDVI 5Y Avg", "EVI", "NDWI", "LST"], "MODIS via Earth Engine",
         str(FRESH_RUN["MODIS (via GEE)"]) if "MODIS (via GEE)" in FRESH_RUN else "NOT FETCHED", "Satellite observations", "Cropland-masked zonal mean over the tehsil polygon. Needs CELL 10."),
        ("SOIL_INTELLIGENCE", ["pH", "Organic Carbon", "N", "Soil Texture"], "ISRIC SoilGrids 2.0 via Earth Engine",
         "2020 release" if DATA["SOIL_INTELLIGENCE"]["pH"].notna().any() else "NOT FETCHED", "Modelled (not lab data)", "0-5 cm, cropland-masked. N = total nitrogen. Needs CELL 10."),
        ("SOIL_INTELLIGENCE", ["Soil Type"], "Planning Commission agro-climatic zones via bharatlas.com",
         "Static delineation (undated)", "Official classification (ZONE level)", "Dominant soil of the zone containing the tehsil centre."),
        ("TEHSIL_MASTER", ["Cultivable Area Ha"], "CoRE Stack cropland + orchard (gaps: ESA WorldCover via Earth Engine)",
         "2023-24 (gaps: WorldCover 2021)", "Satellite model (lower-bound proxy)", "Not revenue-record cultivable land."),
        ("CROP_SOWN", ["Sown Area Ha"], "DES APY (Ministry of Agriculture) via India Data Portal",
         META.get("apy_years", "NOT FETCHED"), "Official statistics (DISTRICT, final reported)", "Latest reported year per season; NOT current sowing progress."),
        ("YIELD_PRODUCTION", ["Area Sown Ha", "Historical Yield", "Production Estimate", "Yield vs 5Y Avg %"], "DES APY (Ministry of Agriculture) via India Data Portal",
         META.get("apy_years", "NOT FETCHED"), "Official statistics (DISTRICT level)", "Latest reported year vs mean of up to 5 prior years."),
        ("5_YEAR_HISTORY", ["Value", "5Y Average", "Current vs 5Y %"], "DES APY (Ministry of Agriculture) via India Data Portal",
         META.get("apy_years", "NOT FETCHED"), "Official statistics (DISTRICT level)", "Yield kg/ha, area ha, production t per district and crop."),
        ("EXTREME_EVENTS", ["Event"], "PIB RSS / Google News RSS (optional AI tagging, unverified)",
         f"last {NEWS_DAYS} days to {FRESH_RUN['News RSS']}" if "News RSS" in FRESH_RUN else "NOT FETCHED", "UNVERIFIED news signal", "Headlines with link and publisher; context only."),
    ]


GAPS = {  # what is still needed for tabs/fields that stayed empty
    "SOIL_INTELLIGENCE": "Soil Health Card portal (no API) or Earth Engine login for SoilGrids",
    "CROP_SOWN": "Current-season sowing progress: state agriculture department reports / Bhuvan-VEDAS (registration). APY gives final district area only",
    "CROP_STAGE": "Derived from CROP_SOWN + satellite - needs sowing data first",
    "PEST_DISEASE": "State pest-surveillance reports (NIPHM / state portals) - no open API",
    "YIELD_PRODUCTION": "DES APY via India Data Portal (district only) - CELL 9; tehsil crop-cutting data is not open",
    "MARKET_ECONOMICS": "Free data.gov.in key (Agmarknet)",
    "5_YEAR_HISTORY": "DES APY (CELL 9) for yield/area/production; other parameters need your stored runs",
    "FARMER_AGRI_ECOSYSTEM": "Agriculture Census / state registries (no open tehsil API)",
}



def tier_of(source, grade):
    t = (source + " " + grade).lower()
    if "unverified" in t or "rss" in t:
        return 4
    if "reference default" in t or "derived" in t or "computed here" in t:
        return 5
    if "core stack" in t:
        return 2
    if any(k in t for k in ("fallback", "open-meteo", "copernicus", "earth engine", "soilgrids", "modis", "worldcover")):
        return 3
    return 1


def write_lineage(wb):
    ws = wb.create_sheet("DATA_LINEAGE")
    ws.sheet_properties.tabColor = GROUPS["Control"]
    banner(ws, "DATA_LINEAGE", len(LINEAGE_COLS), GROUPS["Control"])
    style_header(ws, LINEAGE_COLS, GROUPS["Control"])
    body = Font(name=FONT, size=10)
    r = HEADER_ROW + 1
    for tab, fields, source, vintage, grade, method, asof in lineage_defs():
        for f in fields:
            df = DATA.get(tab)
            filled = int(df[f].notna().sum()) if df is not None and f in df else 0
            total = len(df) if df is not None else 0
            if filled == 0 and not vintage.startswith("NOT"):
                vintage = "NOT FETCHED"
            row = [tab, f, source, vintage if filled else "NOT FETCHED", grade, method,
                   filled, total, f'=IFERROR(G{r}/H{r},"")',
                   TIER[tier_of(source, grade)] if filled else "n/a - not fetched / no open source"]
            ws.append(row)
            for ci in range(1, 11):
                ws.cell(row=r, column=ci).font = body
            ws.cell(row=r, column=9).number_format = "0%"
            r += 1
    for col, w in zip("ABCDEFGHIJ", (24, 24, 44, 26, 36, 70, 12, 12, 12, 44)):
        ws.column_dimensions[col].width = w
    for text, color in (("NOT FETCHED", "FFC7CE"),):
        ws.conditional_formatting.add(f"D{HEADER_ROW + 1}:D{r}", CellIsRule(
            operator="equal", formula=[f'"{text}"'], fill=PatternFill("solid", bgColor=color)))


def gap_report():
    print("\nGAP REPORT - tabs/fields still empty and what they need:")
    for tab, need in GAPS.items():
        df = DATA.get(tab)
        n = 0 if df is None or df.empty else int(df.drop(columns=[c for c in ("State", "District", "Tehsil") if c in df],
                                                        errors="ignore").notna().any(axis=1).sum())
        if n == 0:
            print(f"  {tab:<24} -> {need}")


def _short(src):
    return src.split(" via ")[0].split(" (")[0][:48]


def stamp_sources():
    """Fill 'Source' (where blank) and 'Source Date' (a DATE) on every data tab. Row-level where the source differs."""
    defs = lineage_defs()
    d_imd, d_sat = FRESH_RUN.get("IMD gridded"), FRESH_RUN.get("MODIS (via GEE)")
    for tab in [t[0] for t in SHEETS[:18]]:
        df = DATA.get(tab)
        if df is None or df.empty:
            continue
        hdrs = HEADERS[tab]
        keycols = {"State", "District", "Tehsil", "Tehsil Code", "Source", "Source Date"}
        has = df.reindex(columns=[h for h in hdrs if h not in keycols]).notna().any(axis=1).to_numpy()
        used = [(source, asof) for t, fields, source, vintage, grade, method, asof in defs
                if t == tab and any(f in df and df[f].notna().any() for f in fields)]
        if not used:
            continue
        n = len(df)
        srcs = "; ".join(dict.fromkeys(a for a, _ in used))
        dts = [d for _, d in used if d]
        tab_date = max(dts) if dts else None
        src_col = np.full(n, srcs, dtype=object)
        dat_col = np.full(n, tab_date, dtype=object)
        if tab == "WEATHER_LIVE" and "Data Date" in df:
            dat_col = df["Data Date"].map(lambda v: v if hasattr(v, "year") else tab_date).to_numpy()
            if "Source" in df:
                src_col = df["Source"].astype(object).where(df["Source"].notna(), srcs).to_numpy()
        elif tab == "WATER_IRRIGATION":
            des = (df["_irr_src"].astype(str) == "DES").to_numpy() if "_irr_src" in df else np.zeros(n, bool)
            agc = df["_irr_src"].astype(str).str.startswith("AGC").to_numpy() if "_irr_src" in df else np.zeros(n, bool)
            common = "CGWB 2020 groundwater assessment; MI Census 2017-18 wells (split to tehsils); CoRE Stack 2023-24 water bodies; GLDAS groundwater trend"
            src_col = np.where(agc, "Agriculture Census 2015-16 irrigated area; " + common,
                               np.where(des, f"DES Land Use Statistics {META.get('irr_year', '')} (district, split to tehsils); " + common,
                                        "CoRE Stack 2023-24 cropping-pattern PROXY for irrigated/rainfed; " + common))
        elif tab == "SATELLITE_INTELLIGENCE" and "Satellite Date" in df:
            dat_col = df["Satellite Date"].map(lambda v: v if hasattr(v, "year") else tab_date).to_numpy()
        elif tab == "EXTREME_EVENTS" and "Date" in df:
            dat_col = df["Date"].map(lambda v: v if hasattr(v, "year") else tab_date).to_numpy()
        if "Source" in hdrs:
            cur = df["Source"] if "Source" in df else pd.Series([None] * n, index=df.index)
            keep = (cur.notna() & (cur.astype(str) != "")).to_numpy() & (tab not in ("WATER_IRRIGATION", "WEATHER_LIVE"))
            df["Source"] = np.where(has, np.where(keep, cur.astype(object).to_numpy(), src_col),
                                    cur.astype(object).to_numpy() if "Source" in df else None)
        df["Source Date"] = np.where(has, dat_col, None)
    for _t in DATA:                                       # helper columns (names starting with _) never go to the outputs
        if len(DATA[_t].columns) and any(str(c).startswith("_") for c in DATA[_t].columns):
            DATA[_t] = DATA[_t].drop(columns=[c for c in DATA[_t].columns if str(c).startswith("_")])


def banner_line(name):
    if name in ("DATA_FRESHNESS_CONTROL", "SOURCE_REGISTRY"):
        return ""
    if name in VALIDATION_TABS:
        return f"Rows: {len(DATA.get(name, [])):,}  |  built {TODAY:%Y-%m-%d}  |  validation layer (CELL 12B / log of all steps)"
    return source_line(name)


def build_workbook(path=OUTPUT_FILE, skip=(), only=None):
    from openpyxl.utils.indexed_list import IndexedList
    t0 = time.time()
    wb = Workbook()
    wb._fonts = IndexedList([Font(name=FONT, size=10)])      # workbook-wide default font (no per-cell styling needed)
    wb.remove(wb.active)
    for name, group, headers in SHEETS:
        if name in skip or (only is not None and name not in only):
            continue
        ws = wb.create_sheet(name)
        ws.sheet_properties.tabColor = GROUPS[group]
        banner(ws, name, len(headers), GROUPS[group], banner_line(name))
        style_header(ws, headers, GROUPS[group])
        if name == "DATA_FRESHNESS_CONTROL":
            write_freshness(ws, headers)
        elif name == "SOURCE_REGISTRY":
            write_registry(ws, headers)
        else:
            write_data_sheet(ws, name, headers)
            if len(DATA[name]):
                print(f"  wrote {name}: {len(DATA[name]):,} rows  ({time.time() - t0:.0f}s)")
        for col_name, values in DROPDOWNS.get(name, {}).items():
            L = get_column_letter(headers.index(col_name) + 1)
            dv = DataValidation(type="list", formula1='"' + ",".join(values) + '"', allow_blank=True)
            ws.add_data_validation(dv)
            dv.add(f"{L}{HEADER_ROW + 1}:{L}200000")
    if only is None:
        write_lineage(wb)
        write_readme(wb)
    NOTES[("TEHSIL_MASTER", "Source Date")] = "Latest source date among the fields in the row (year-only sources = year end). Per-field dates: DATA_LINEAGE."
    for (sheet, header), text in NOTES.items():
        if sheet in HEADERS and header in HEADERS[sheet] and sheet in wb.sheetnames:
            wb[sheet].cell(row=HEADER_ROW, column=HEADERS[sheet].index(header) + 1).comment = Comment(text, "Agri Engine")
    print("  saving workbook ...")
    wb.save(path)
    print(f"Saved {path} in {time.time() - t0:.0f}s")
    for name in SHEET_NAMES:
        if name in ("DATA_FRESHNESS_CONTROL", "SOURCE_REGISTRY") or name in skip or (only is not None and name not in only):
            continue
        df = DATA[name]
        filled = int(df.drop(columns=[c for c in ("State", "District", "Tehsil") if c in df],
                             errors="ignore").notna().any(axis=1).sum()) if len(df) else 0
        print(f"  {name:<26} {len(df):>6,} rows, {filled:>6,} with data")


def _gap_after_build():
    gap_report()


def save_csvs(zip_path="Agri_Intelligence_CSV.zip"):
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for name in SHEET_NAMES:
            if name in DATA and len(DATA[name]):
                df_ = DATA[name]
                z.writestr(f"{name}.csv", df_.drop(columns=[c for c in df_.columns if str(c).startswith("_")]).to_csv(index=False))
    return zip_path


if len(DATA["TEHSIL_MASTER"]) == 0:
    raise SystemExit("No tehsils are loaded, so the workbook would be empty. Re-run CELL 3 first "
                     "and check its printed output (rows / columns / 'Loaded N tehsils').")
VALIDATION_FILE = "Agri_Validation_India.xlsx"
DATA["SOURCE_RECOVERY_LOG"] = pd.DataFrame(RECOVERY_LOG).reindex(columns=HEADERS["SOURCE_RECOVERY_LOG"])
for _t in SHEET_NAMES:                                    # every tab exists, even if a cell was skipped
    DATA.setdefault(_t, pd.DataFrame(columns=HEADERS[_t]))


def sheet_cells(names):
    return sum(len(DATA[n_]) * len(HEADERS[n_]) for n_ in names if n_ in DATA) + 2000 * len(names)


ram("before build")
stamp_sources()
_cells = sheet_cells(SHEET_NAMES)
SPLIT_OUT = set(VALIDATION_TABS) if _cells > SHEETS_CELL_LIMIT else set()
print(f"Workbook size: {_cells:,} cells (Google Sheets limit 10,000,000)"
      + (" -> validation tabs go to a separate workbook / Sheet" if SPLIT_OUT else ""))
build_workbook(skip=SPLIT_OUT)
if SPLIT_OUT:
    build_workbook(VALIDATION_FILE, only=SPLIT_OUT)
gap_report()
csv_zip = save_csvs()
print(f"Files are in the Colab Files panel: {OUTPUT_FILE}{', ' + VALIDATION_FILE if SPLIT_OUT else ''}, {csv_zip}")
if AUTO_DOWNLOAD:
    try:
        from google.colab import files
        files.download(OUTPUT_FILE)
        files.download(csv_zip)
    except ImportError:
        pass


# ===== CELL 14 : CREATE THE GOOGLE SHEET IN YOUR DRIVE ======================
# Uploads the workbook and lets Google convert it to a native Google Sheet
# (formulas, colours, dropdowns and header notes are kept). Opens a Google
# permission pop-up once. Result: a link printed below.
SHEET_TITLE = f"Agri Intelligence India {TODAY}"


def upload_to_google_sheets(xlsx_path=OUTPUT_FILE, title=SHEET_TITLE):
    try:
        from google.colab import auth
        auth.authenticate_user()
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaFileUpload
        drive = build("drive", "v3")
        meta = {"name": title, "mimeType": "application/vnd.google-apps.spreadsheet"}
        media = MediaFileUpload(
            xlsx_path, resumable=True,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        req = drive.files().create(body=meta, media_body=media, fields="id,webViewLink")
        resp = None
        while resp is None:
            status, resp = req.next_chunk()
            if status:
                print(f"  uploaded {int(status.progress() * 100)}%")
        print("Google Sheet created:", resp["webViewLink"])
        SHEET_URLS[title] = resp["webViewLink"]
        rlog("Google Sheets", "Google Drive API (convert xlsx)", resp["webViewLink"], title, "OK")
        return resp["webViewLink"]
    except ImportError:
        return upload_headless(xlsx_path, title)               # [GitHub] service account, fixed Sheet IDs
    except Exception as e:
        print(f"Drive upload failed ({type(e).__name__}: {str(e)[:150]}).")
        SHEET_URLS[title] = None
        rlog("Google Sheets", "Google Drive API (convert xlsx)", "", title, "FAILED", fallback="download the xlsx", msg=f"{type(e).__name__}: {str(e)[:150]}")
        print(f"Your workbook is saved as {xlsx_path} in the Colab Files panel (left bar): right-click -> Download, "
              "then drag it into drive.google.com and open it with Google Sheets.")
    return None


def _sheet_target(title):
    """[GitHub] Which fixed Google Sheet a workbook overwrites (IDs come from GitHub secrets / variables)."""
    for prefix, env in (("Agri Trend Module - E_SOWING", "SHEET_ID_TREND_SOWING"), ("Agri Trend Module", "SHEET_ID_TREND"),
                        ("Agri Validation India", "SHEET_ID_VALIDATION"), ("Agri Intelligence India", "SHEET_ID_MAIN")):
        if title.startswith(prefix):
            return env, os.environ.get(env, "").strip()
    return "", ""


def upload_headless(xlsx_path, title):
    """[GitHub] Outside Colab: convert the xlsx INTO an existing Google Sheet (owned by you, shared with the service account
    as Editor), so the link never changes and storage counts to your Drive. Same content as the Colab upload."""
    env, fid = _sheet_target(title)
    raw = os.environ.get("SHEETS_KEY", "").strip() or os.environ.get("GEE_KEY", "").strip()
    folder = os.environ.get("SHEETS_FOLDER_ID", "").strip()
    if not raw or not (fid or folder):
        print(f"Google Sheet not updated: set the secrets SHEETS_KEY (or GEE_KEY) and {env or 'SHEET_ID_*'} - "
              f"{xlsx_path} was not uploaded (see the run log).")
        SHEET_URLS[title] = None
        rlog("Google Sheets", "Drive API (service account)", "", title, "SKIPPED", msg=f"missing {env} / key")
        return None
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaFileUpload
        creds = service_account.Credentials.from_service_account_info(json.loads(raw), scopes=["https://www.googleapis.com/auth/drive"])
        drive = build("drive", "v3", credentials=creds, cache_discovery=False)
        media = MediaFileUpload(xlsx_path, resumable=True, chunksize=8 * 1024 * 1024,
                                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        if fid:
            req = drive.files().update(fileId=fid, body={"name": title}, media_body=media, fields="id,webViewLink", supportsAllDrives=True)
        else:
            req = drive.files().create(body={"name": title, "mimeType": "application/vnd.google-apps.spreadsheet", "parents": [folder]},
                                       media_body=media, fields="id,webViewLink", supportsAllDrives=True)
        resp = None
        while resp is None:
            status, resp = req.next_chunk()
            if status:
                print(f"  uploaded {int(status.progress() * 100)}%")
        print("Google Sheet updated:", resp["webViewLink"])
        SHEET_URLS[title] = resp["webViewLink"]
        rlog("Google Sheets", "Drive API (service account)", resp["webViewLink"], title, "OK")
        return resp["webViewLink"]
    except Exception as e:
        body = (getattr(getattr(e, "content", b""), "decode", lambda *a: "")("utf-8", "ignore") or str(e))[:200]
        print(f"Google Sheet update failed ({type(e).__name__}: {body}). Check that {env} is shared with the service account as Editor. "
              f"{xlsx_path} was not uploaded (see the run log).")
        SHEET_URLS[title] = None
        rlog("Google Sheets", "Drive API (service account)", "", title, "FAILED", msg=f"{type(e).__name__}: {body}")
        return None


SHEET_URL = upload_to_google_sheets()
if SPLIT_OUT:
    VALIDATION_SHEET_URL = upload_to_google_sheets(VALIDATION_FILE, f"Agri Validation India {TODAY}")
print("\n================ GOOGLE SHEET LINKS ================")
for _k, _v in SHEET_URLS.items():
    print(f"  {_k}: {_v or 'upload failed - download the xlsx from the Files panel'}")


# ===== CELL 15 : TREND MODULE - COLLECT 5-YEAR MONTHLY SERIES (rain, dry days, NDVI, stress, soil moisture, groundwater, irrigation proxy) =====
# For the forecasting module: 5 completed calendar years (Year 1..5) month by month + the current year to date (YTD), per tehsil.
#   rain / cdd      IMD gridded rainfall (Tier 1): monthly total, longest dry spell
#   ndvi / stress   MODIS NDVI over cropland pixels (Earth Engine): monthly mean + share of cropland pixels below 85% of the 5-year normal
#   irrig           IRRIGATION PROXY: share of cropland pixels green (NDVI > 0.40) in months with < 40 mm rain (no official monthly series exists)
#   soil            ERA5-Land soil moisture, top layer (model, ~11 km)
#   gw              GLDAS-2.2 groundwater STORAGE (GRACE-assimilated model, ~25 km) - not well levels
# Heavy Earth Engine steps are checkpointed: if Colab disconnects, re-run this cell and it resumes.
import warnings
warnings.filterwarnings("ignore", category=RuntimeWarning)
if globals().get("T") is None or len(T) == 0:
    raise SystemExit("Tehsils are not loaded in this session. Run CELL 1 to CELL 3 first (CELL 11 for the Earth Engine key), then this cell.")
if not globals().get("GEE_INFO") and "ensure_gee_key" in globals():
    ensure_gee_key()                                   # asks for the JSON (upload / Drive) if CELL 11 did not find it

TREND_YEARS = 5
TREND_SCALE = 1000            # metres; 2000 = about 2x faster but coarser, 500 = finer but about 3x slower
TREND_RUN_IRRIGATION = True   # False skips the irrigation-proxy passes (saves roughly a third of the Earth Engine time)
NDVI_STRESS_RATIO = 0.85      # a cropland pixel is 'stressed' if NDVI < 85% of its 5-year normal for that month
STRESS_FRAC_THRESHOLD = 0.30  # a tehsil-month counts as 'stressed' if >= 30% of cropland pixels are stressed
IRR_NDVI, IRR_RAIN_MM = 0.40, 40.0
YEARS = [TODAY.year - k for k in range(TREND_YEARS, 0, -1)]
CUR, CM = TODAY.year, TODAY.month
MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
NT = len(T)
TR_KEYS = ("rain", "cdd", "ndvi", "stress", "irrig", "soil", "gw")
TR = {k: np.full((NT, 6, 12), np.nan, dtype="float32") for k in TR_KEYS}       # (tehsil, year index 0-4 = Year 1..5, 5 = current YTD, month)
TR_FILE = f"trend_{TODAY}.npz"
TR_SRC_FILE = f"trend_src_{TODAY}.json"
TR_SRC = json.load(open(TR_SRC_FILE)) if os.path.exists(TR_SRC_FILE) else {}  # "rain|<year>" -> IMD / CHIRPS
FV = {}                                                                        # forecast validation per district
if os.path.exists(TR_FILE):
    _z = np.load(TR_FILE)
    for _k in TR_KEYS:
        if _k in _z and _z[_k].shape == TR[_k].shape:
            TR[_k] = _z[_k]
    print("Trend cache loaded:", {k: int(np.isfinite(v).sum()) for k, v in TR.items()})


def tr_save():
    np.savez_compressed(TR_FILE, **TR)
    json.dump(TR_SRC, open(TR_SRC_FILE, "w"))


def tr_has(k):
    return bool(np.isfinite(TR[k]).any())


# ---------------------------------------------------------------- IMD monthly rainfall + longest dry spell
def month_stats(arr, times, dry_mm=None):
    """arr (days, points), times DatetimeIndex -> monthly total and longest dry run per (year, month)."""
    dry_mm = DRY_DAY_MM if dry_mm is None else dry_mm
    out = {}
    for (y, m), idx in pd.Series(np.arange(len(times))).groupby([times.year, times.month]):
        sl = arr[idx.to_numpy()]
        allnan = np.isnan(sl).all(axis=0)
        tot = np.nansum(sl, axis=0)
        run = np.zeros(sl.shape[1])
        best = np.zeros(sl.shape[1])
        for d in range(sl.shape[0]):
            isdry = sl[d] < dry_mm
            run = np.where(isdry, run + 1, 0)
            best = np.maximum(best, run)
        out[(y, m)] = (np.where(allnan, np.nan, tot), np.where(allnan, np.nan, best))
    return out


def trend_imd():
    """IMD monthly rainfall + longest dry spell, ONE YEAR AT A TIME: a year whose IMD file is not published / not reachable no longer
    stops the others. Missing months are filled later from CHIRPS (Earth Engine) and labelled."""
    import imdlib as imd
    lats, lons = T["Latitude"].values, T["Longitude"].values
    for y in YEARS:
        yi = YEARS.index(y)
        if np.isfinite(TR["rain"][:, yi, :]).any():
            continue
        try:
            print(f"  IMD archive {y} (daily -> monthly) ...")
            with contextlib.redirect_stdout(io.StringIO()):
                imd.get_data("rain", y, y, fn_format="yearwise", file_dir=WORK)
            ds = imd.open_data("rain", y, y, "yearwise", WORK).get_xarray()
            pts = at_points(ds["rain"].where(ds["rain"] > -900), lats, lons).compute()
            for (yy, m), (tot, cdd) in month_stats(pts.values, pd.DatetimeIndex(pts["time"].values)).items():
                if yy == y:
                    TR["rain"][:, yi, m - 1], TR["cdd"][:, yi, m - 1] = tot, cdd
            TR_SRC[f"rain|{y}"] = "IMD"
            tr_save()
            rlog("CELL 15", f"IMD gridded rainfall {y}", URL_IMD if "URL_IMD" in globals() else "", "Trend: monthly rainfall", "OK")
        except Exception as e:
            print(f"  IMD {y} failed ({type(e).__name__}: {str(e)[:100]}) - CHIRPS fallback will be tried in the Earth Engine step")
            rlog("CELL 15", f"IMD gridded rainfall {y}", "", "Trend: monthly rainfall", "FAILED", fallback="CHIRPS via Earth Engine",
                 msg=f"{type(e).__name__}: {str(e)[:150]}")
    if not np.isfinite(TR["rain"][:, 5, :]).any():
        end = TODAY - timedelta(days=IMD_LAG_DAYS)
        try:
            print(f"  IMD year-to-date {CUR}-01-01 .. {end} (daily real-time files, a few minutes) ...")
            da = imd_realtime("rain", date(CUR, 1, 1), end)
            for (y, m), (tot, cdd) in month_stats(at_points(da, lats, lons).values, pd.DatetimeIndex(da["time"].values)).items():
                if y == CUR:
                    TR["rain"][:, 5, m - 1], TR["cdd"][:, 5, m - 1] = tot, cdd
            TR_SRC[f"rain|{CUR}"] = "IMD"
            tr_save()
            rlog("CELL 15", f"IMD real-time rainfall {CUR} YTD", "", "Trend: monthly rainfall", "OK")
        except Exception as e:
            print(f"  IMD YTD failed ({type(e).__name__}: {str(e)[:100]}) - CHIRPS fallback will be tried")
            rlog("CELL 15", f"IMD real-time rainfall {CUR} YTD", "", "Trend: monthly rainfall", "FAILED", fallback="CHIRPS", msg=str(e)[:150])
    print("  IMD monthly rainfall / dry spells:", int(np.isfinite(TR["rain"]).any(axis=(1, 2)).sum()), "tehsils with data")


def trend_forecast_validation(days=30):
    """Share of days where the archived day-ahead forecast (Open-Meteo) agreed with IMD on rain / no rain, per district."""
    fvf = f"fv_{TODAY}.json"
    if os.path.exists(fvf):
        FV.update(json.load(open(fvf)))
        print("  forecast validation: cached for", len(FV), "districts")
        return
    end = TODAY - timedelta(days=IMD_LAG_DAYS)
    start = end - timedelta(days=days - 1)
    da = imd_realtime("rain", start, end)
    obs = at_points(da, T["Latitude"].values, T["Longitude"].values).values          # (days, tehsil)
    dts = [pd.Timestamp(x).date() for x in da["time"].values]
    g = T.groupby(["State", "District"], sort=False)
    keys = list(g.groups.keys())
    cent = np.array([[T.loc[ix, "Latitude"].mean(), T.loc[ix, "Longitude"].mean()] for ix in g.groups.values()])
    dobs = np.array([np.nanmean(obs[:, list(ix)], axis=1) for ix in g.groups.values()])      # (districts, days)
    for i in range(0, len(keys), 50):
        sl = slice(i, i + 50)
        d = get_json("https://historical-forecast-api.open-meteo.com/v1/forecast", {
            "latitude": ",".join(str(round(float(x), 3)) for x in cent[sl, 0]),
            "longitude": ",".join(str(round(float(x), 3)) for x in cent[sl, 1]),
            "start_date": start.isoformat(), "end_date": end.isoformat(), "daily": "precipitation_sum", "timezone": "Asia/Kolkata"}, tries=3)
        locs = d if isinstance(d, list) else ([d] if d else [])
        for k, loc in zip(range(i, i + 50), locs):
            fc = np.array([x if x is not None else np.nan for x in loc.get("daily", {}).get("precipitation_sum", [])], dtype=float)
            ob = dobs[k][:len(fc)]
            ok = ~np.isnan(fc) & ~np.isnan(ob)
            if ok.sum() >= 10:
                agree = ((fc[ok] >= DRY_DAY_MM) == (ob[ok] >= DRY_DAY_MM)).mean() * 100
                FV["|".join(keys[k])] = [round(float(agree), 1), int(ok.sum())]
        time.sleep(3)
    json.dump(FV, open(fvf, "w"))
    print(f"  forecast validation for {len(FV):,} of {len(keys):,} districts (rain / no-rain agreement over {days} days)")


# ---------------------------------------------------------------- Earth Engine monthly series
def _fill(arr, df, prefix, yi):
    for c in df.columns:
        m = re.fullmatch(prefix + r"(\d{2})", str(c))
        if m:
            arr[:, yi, int(m.group(1)) - 1] = pd.to_numeric(df[c], errors="coerce").to_numpy("float32")


def _has_images(ee, col, y, m):
    s = ee.Date.fromYMD(y, m, 1)
    return col.filterDate(s, s.advance(1, "month")).size().getInfo() > 0


def g_ndvi(ee, geoms, sq, crop):
    if tr_has("ndvi"):
        print("  NDVI / stress: cached")
        return
    col = ee.ImageCollection("MODIS/061/MOD13Q1").select("NDVI")
    base = [col.filter(ee.Filter.calendarRange(m, m, "month")).filterDate(f"{YEARS[0]}-01-01", f"{YEARS[-1] + 1}-01-01")
            .mean().multiply(0.0001) for m in range(1, 13)]
    for yi, y in enumerate(YEARS + [CUR]):
        bands = []
        for m in range(1, 13 if y != CUR else CM + 1):
            if y == CUR and not _has_images(ee, col, y, m):
                continue
            s = ee.Date.fromYMD(y, m, 1)
            img = col.filterDate(s, s.advance(1, "month")).mean().multiply(0.0001)
            bands += [img.rename(f"n{m:02d}"), img.lt(base[m - 1].multiply(NDVI_STRESS_RATIO)).rename(f"s{m:02d}")]
        if not bands:
            continue
        df = gee_zonal(ee, ee.Image.cat(bands).updateMask(crop), TREND_SCALE, geoms, chunk=100, tag=f"trndvi{y}")
        _fill(TR["ndvi"], df, "n", yi)
        _fill(TR["stress"], df, "s", yi)
        tr_save()
        print(f"  NDVI / stress {y}: done")


def g_irrig(ee, geoms, sq, crop):
    if tr_has("irrig"):
        print("  irrigation proxy: cached")
        return
    col = ee.ImageCollection("MODIS/061/MOD13Q1").select("NDVI")
    rain = ee.ImageCollection("UCSB-CHG/CHIRPS/DAILY").select("precipitation")
    for yi, y in enumerate(YEARS + [CUR]):
        bands, bands_names = [], []
        for m in range(1, 13 if y != CUR else CM):                      # skip the running month (rain not complete)
            if y == CUR and not (_has_images(ee, col, y, m) and _has_images(ee, rain, y, m)):
                print(f"  irrigation proxy {y}-{m:02d}: skipped (NDVI or CHIRPS not yet published)")
                continue
            s = ee.Date.fromYMD(y, m, 1)
            e = s.advance(1, "month")
            nd = col.filterDate(s, e).mean().multiply(0.0001)
            bands.append(nd.gt(IRR_NDVI).And(rain.filterDate(s, e).sum().lt(IRR_RAIN_MM)).rename(f"i{m:02d}"))
            bands_names.append(f"i{m:02d}")
        if not bands:
            continue
        df = gee_zonal(ee, ee.Image.cat(bands).updateMask(crop), TREND_SCALE, geoms, chunk=100, tag=f"trirr{y}")
        if len(bands) == 1 and "mean" in df.columns:
            df = df.rename(columns={"mean": bands_names[0]})
        _fill(TR["irrig"], df, "i", yi)
        tr_save()
        print(f"  irrigation proxy {y}: done")


def _monthly_coarse(ee, col, key, tag, sq, scale, avail, target=None, agg="mean"):
    bands, names = [], []
    for yi, y in enumerate(YEARS + [CUR]):
        for m in range(1, 13):
            if (y, m) in avail:
                s = ee.Date.fromYMD(y, m, 1)
                c_ = col.filterDate(s, s.advance(1, "month"))
                names.append(f"{key}{yi}{m:02d}")
                bands.append((c_.sum() if agg == "sum" else c_.mean()).rename(names[-1]))
    if not bands:
        return
    df = gee_zonal(ee, ee.Image.cat(bands), scale, sq, chunk=500, tag=tag)
    if len(names) == 1 and "mean" in df.columns:          # single-band results come back as 'mean'
        df = df.rename(columns={"mean": names[0]})
    tgt = target or ("soil" if key == "s" else "gw")
    for c in df.columns:
        m_ = re.fullmatch(key + r"(\d)(\d{2})", str(c))
        if m_:
            TR[tgt][:, int(m_.group(1)), int(m_.group(2)) - 1] = pd.to_numeric(df[c], errors="coerce").to_numpy("float32")


def g_soil(ee, geoms, sq, crop):
    if tr_has("soil"):
        print("  soil moisture: cached")
        return
    col = ee.ImageCollection("ECMWF/ERA5_LAND/MONTHLY_AGGR").select("volumetric_soil_water_layer_1")
    t = col.filterDate(f"{YEARS[0]}-01-01", str(TODAY)).aggregate_array("system:time_start").getInfo()
    avail = {(pd.Timestamp(x, unit="ms").year, pd.Timestamp(x, unit="ms").month) for x in t}
    _monthly_coarse(ee, col, "s", "trsoil", sq, 11132, avail)
    tr_save()
    print("  soil moisture: done; latest month", max(avail) if avail else "n/a")


def g_gw(ee, geoms, sq, crop):
    if tr_has("gw"):
        print("  groundwater storage: cached")
        return
    col = ee.ImageCollection("NASA/GLDAS/V022/CLSM/G025/DA1D").select("GWS_tavg")
    last = pd.Timestamp(ee.Date(col.sort("system:time_start", False).first().get("system:time_start")).getInfo()["value"], unit="ms")
    avail = {(y, m) for y in YEARS + [CUR] for m in range(1, 13) if (y, m) < (last.year, last.month)}
    _monthly_coarse(ee, col, "g", "trgw", sq, 25000, avail)
    tr_save()
    print("  groundwater storage: done; complete months to", last.strftime("%Y-%m"))


def g_rain_fallback(ee, geoms, sq, crop):
    """CHIRPS monthly rainfall (Earth Engine) ONLY for months IMD could not supply; labelled per year. Dry-spell months stay blank."""
    need = {(y, m) for yi, y in enumerate(YEARS + [CUR]) for m in range(1, 13)
            if (y, m) < (CUR, CM) and not np.isfinite(TR["rain"][:, yi, m - 1]).any()}
    if not need:
        print("  CHIRPS fallback: not needed (IMD complete)")
        return
    ch = ee.ImageCollection("UCSB-CHG/CHIRPS/DAILY").select("precipitation")
    last = ee_last_date(ee, ch)
    avail = {(y, m) for (y, m) in need if date(y, m, 1) + timedelta(days=31) <= last + timedelta(days=1)}
    print(f"  CHIRPS fallback for {len(avail)} months without IMD (CHIRPS data to {last})")
    _monthly_coarse(ee, ch, "r", "trchirps", sq, 5566, avail, target="rain", agg="sum")
    for y in sorted({y for y, _ in avail}):
        TR_SRC[f"rain|{y}"] = "CHIRPS (fallback)" if TR_SRC.get(f"rain|{y}") is None else TR_SRC[f"rain|{y}"] + " + CHIRPS months"
    tr_save()
    rlog("CELL 15", "CHIRPS daily (Earth Engine)", URL_CAT + "UCSB-CHG_CHIRPS_DAILY" if "URL_CAT" in globals() else "", "Trend: rainfall months IMD could not supply",
         "OK", len(avail), fallback="Tier 3 for these months")


def trend_gee():
    try:
        from shapely.geometry import mapping
        ee = gee_init()
    except Exception as e:
        print("Earth Engine unavailable for the trend module:", type(e).__name__, str(e)[:200])
        rlog("CELL 15", "Earth Engine login", "", "Trend series", "FAILED", msg=f"{type(e).__name__}: {str(e)[:200]}")
        return
    geoms = [mapping(g.simplify(SIMPLIFY_DEG, preserve_topology=True)) for g in T["geometry"]]
    d = 0.15
    sq = [{"type": "Polygon", "coordinates": [[[lo - d, la - d], [lo + d, la - d], [lo + d, la + d], [lo - d, la + d], [lo - d, la - d]]]}
          for la, lo in zip(T["Latitude"], T["Longitude"])]
    lc = ee.ImageCollection("MODIS/061/MCD12Q1").sort("system:time_start", False).first().select("LC_Type1")
    crop = lc.eq(12).Or(lc.eq(14))
    steps = (g_rain_fallback, g_soil, g_gw, g_ndvi) + ((g_irrig,) if TREND_RUN_IRRIGATION else ())
    for step in steps:                                                  # cheap coarse series first, long NDVI passes last
        t0 = time.time()
        try:
            step(ee, geoms, sq, crop)
            print(f"  {step.__name__} took {(time.time() - t0) / 60:.1f} min")
            rlog("CELL 15", step.__name__, "Earth Engine", "Trend series", "OK", msg=f"{(time.time() - t0) / 60:.1f} min")
        except Exception as e:
            print(f"  {step.__name__} failed after {(time.time() - t0) / 60:.1f} min: {type(e).__name__}: {str(e)[:160]}")
            rlog("CELL 15", step.__name__, "Earth Engine", "Trend series", "FAILED", msg=f"{type(e).__name__}: {str(e)[:160]}")


ram("before trend collection")
for _fn in (trend_imd, trend_forecast_validation):
    try:
        _fn()
    except Exception as e:
        print(f"  {_fn.__name__} failed: {type(e).__name__}: {str(e)[:160]}")
        rlog("CELL 15", _fn.__name__, "", "Trend", "FAILED", msg=f"{type(e).__name__}: {str(e)[:160]}")
trend_gee()
_filled = {k: int(np.isfinite(v).any(axis=(1, 2)).sum()) for k, v in TR.items()}
print("Trend series - tehsils with data per series:", _filled, "| rainfall source by year:", TR_SRC)
if not any(_filled.values()):
    print("NOTHING COLLECTED. Causes, in order: (1) IMD archive not reachable AND (2) Earth Engine login failed - see the FAILED lines above. "
          "Fix the Earth Engine key (CELL 11 asks for it) and re-run this cell; it resumes from checkpoints.")


# ===== CELL 16 : TREND MODULE - DERIVE INDICATORS, BUILD TREND TABS A-G + LONG FILE (+ Google Sheet) =====
# A Soil  B Water & irrigation  C Weather  D Satellite  E Crop sowing  F Crop stage  G Crop health & loss
# Output 1: Agri_Trend_Module.xlsx (wide: Year 1..5 month by month + current-year months + YTD summaries, one row per tehsil)
# Output 2: Agri_Trend_Long.parquet (tidy: tehsil, parameter, year, month, value) - the file to feed a forecasting model.
TREND_FILE = "Agri_Trend_Module.xlsx"
TREND_LONG = "Agri_Trend_Long.parquet"
TREND_TO_SHEETS = True
STAGE_NAMES = ["Fallow / low vegetation", "Sowing / emergence", "Vegetative", "Peak / flowering", "Maturity / harvest"]
NEXT_STAGE = {0: 1, 1: 2, 2: 3, 3: 4, 4: 0}
# TYPICAL agronomic requirements (REFERENCE DEFAULTS, editable): pH (optimum lo, hi, absolute lo, hi), annual rain mm (same), preferred textures.
CROP_REQ = {
    "Paddy":      dict(ph=(5.5, 7.5, 4.5, 8.5), rain=(1000, 2000, 700, 3000), tex="clay|clay loam|silty clay|silt"),
    "Wheat":      dict(ph=(6.0, 7.5, 5.5, 8.5), rain=(400, 900, 250, 1200), tex="loam|clay loam|silt loam"),
    "Maize":      dict(ph=(5.8, 7.5, 5.0, 8.5), rain=(600, 1100, 400, 1500), tex="loam|sandy loam|silt loam"),
    "Groundnut":  dict(ph=(6.0, 7.0, 5.0, 8.0), rain=(500, 1000, 400, 1300), tex="sandy loam|loamy sand|sand"),
    "Cotton":     dict(ph=(6.0, 8.0, 5.5, 8.5), rain=(600, 1000, 400, 1400), tex="clay|clay loam|silty clay"),
    "Soyabean":   dict(ph=(6.0, 7.0, 5.5, 8.0), rain=(700, 1200, 500, 1500), tex="loam|clay loam|silt loam"),
    "Bajra":      dict(ph=(6.5, 8.0, 5.5, 9.0), rain=(300, 700, 250, 900), tex="sand|loamy sand|sandy loam"),
    "Sugarcane":  dict(ph=(6.5, 7.5, 5.5, 8.5), rain=(1000, 1800, 700, 2500), tex="loam|clay loam|clay"),
    "Black gram": dict(ph=(6.5, 7.5, 5.5, 8.5), rain=(600, 1000, 400, 1300), tex="loam|clay loam"),
    "Onion":      dict(ph=(6.0, 7.5, 5.5, 8.0), rain=(500, 800, 350, 1100), tex="sandy loam|loam"),
    "Potato":     dict(ph=(5.2, 6.5, 4.5, 7.5), rain=(500, 1000, 300, 1300), tex="sandy loam|loam"),
}
LASTM = min(CM, 12) - 1                                   # index of the current month (0-based)


def _div(a, b):
    with np.errstate(all="ignore"):
        return a / np.where(b == 0, np.nan, b)


def last_month_with(arr_year):
    idx = np.flatnonzero(np.isfinite(arr_year).any(axis=0))
    return int(idx[-1]) if len(idx) else -1


def ytd(arr, how, lm):
    """Current-year YTD value, the mean of the same months in Years 1-5, and % difference."""
    n = arr.shape[0]
    if lm < 0:
        z = np.full(n, np.nan)
        return z, z, z
    cur, past = arr[:, 5, :lm + 1], arr[:, :5, :lm + 1]
    if how == "sum":
        f = lambda a, ax: np.where(np.isnan(a).all(axis=ax), np.nan, np.nansum(a, axis=ax))
    elif how == "max":
        f = lambda a, ax: np.where(np.isnan(a).all(axis=ax), np.nan, np.nanmax(np.where(np.isnan(a), -np.inf, a), axis=ax))
    else:
        f = lambda a, ax: np.nanmean(a, axis=ax)
    cv, pv = f(cur, 1), f(past, 2)
    pavg = np.nanmean(pv, axis=1)
    return cv, pavg, (_div(cv, pavg) - 1) * 100


def latest_percentile(arr, lm):
    n = arr.shape[0]
    if lm < 0:
        return np.full(n, np.nan)
    cur, past = arr[:, 5, lm], arr[:, :5, lm]
    valid = np.isfinite(past).sum(axis=1)
    pct = (past <= cur[:, None]).sum(axis=1) / np.where(valid > 0, valid, np.nan) * 100
    return np.where(np.isfinite(cur), pct, np.nan)


def wide(arr, label, how, digits, lm):
    """Year 1..5 month by month, current-year months so far, then YTD / 5-year-average YTD / % difference."""
    cols = {}
    for yi, y in enumerate(YEARS + [CUR]):
        for m in range(12 if yi < 5 else lm + 1):
            cols[f"{label} {y}-{MON[m]}"] = np.round(arr[:, yi, m], digits)
    cv, pavg, pct = ytd(arr, how, lm)
    cols[f"{label} YTD {CUR}"] = np.round(cv, digits)
    cols[f"{label} YTD 5Y avg"] = np.round(pavg, digits)
    cols[f"{label} YTD vs 5Y avg %"] = np.round(pct, 1)
    if lm >= 0:
        cur = arr[:, 5, lm]
        prev = arr[:, 5, lm - 1] if lm >= 1 else arr[:, 4, 11]
        ly = arr[:, 4, lm]
        pc = lambda a, b: np.where(np.isfinite(a) & (b > 0), (a / np.where(b > 0, b, np.nan) - 1) * 100, np.nan)
        cols[f"{label} latest month ({MON[lm]} {CUR})"] = np.round(cur, digits)
        cols[f"{label} MoM change (vs previous month)"] = np.round(cur - prev, digits)
        cols[f"{label} MoM % (vs previous month)"] = np.round(pc(cur, prev), 1)
        cols[f"{label} vs same month last year %"] = np.round(pc(cur, ly), 1)
    return pd.DataFrame(cols)


def phenology_codes(nd):
    """Month-wise stage code from the NDVI curve: 0 fallow, 1 sowing/emergence, 2 vegetative, 3 peak/flowering, 4 maturity/harvest, -1 no data."""
    n = nd.shape[0]
    clim = np.nanmean(nd[:, :5, :], axis=1)
    cmin, cmax = np.nanmin(clim, axis=1), np.nanmax(clim, axis=1)
    amp = cmax - cmin
    s = nd.reshape(n, 72)
    prev = np.concatenate([np.full((n, 1), np.nan, dtype=s.dtype), s[:, :-1]], axis=1)
    slope = s - prev
    nrm = _div(s - cmin[:, None], amp[:, None])
    c = np.where(nrm < 0.25, np.where(slope > 0.02, 1, 0), np.where(nrm < 0.6, np.where(slope >= 0, 2, 4), np.where(slope >= -0.02, 3, 4)))
    c = np.where((amp < 0.08)[:, None], 0, c)
    ok = np.isfinite(s) & (np.isfinite(nrm) | (amp < 0.08)[:, None])
    return np.where(ok, c, -1).astype(np.int8).reshape(n, 6, 12), nrm.reshape(n, 6, 12)


def run_lengths(row, code):
    out, cur = [], 0
    for v in row:
        if v == code:
            cur += 1
        elif cur:
            out.append(cur)
            cur = 0
    if cur:
        out.append(cur)
    return out


def stage_summary(codes, lm):
    n = codes.shape[0]
    cur = codes[:, 5, lm] if lm >= 0 else np.full(n, -1)
    months_in = np.zeros(n)
    typ = np.full(n, np.nan)
    for i in range(n):
        c = cur[i]
        if c < 0:
            continue
        k = 0
        for j in range(lm, -1, -1):
            if codes[i, 5, j] == c:
                k += 1
            else:
                break
        months_in[i] = k
        rl = [r for yr in range(5) for r in run_lengths(codes[i, yr], c)]
        typ[i] = np.mean(rl) if rl else np.nan
    return cur, months_in, typ


def season_timing(nrm, lm):
    """Satellite proxy for sowing / harvest months per season and year (cropland-wide, not crop-specific)."""
    n = nrm.shape[0]
    s = nrm.reshape(n, 72)
    res = {k: np.full((n, 6), np.nan) for k in ("k_sow", "k_harv", "r_sow", "r_harv")}
    ar = np.arange(n)
    for yi in range(6):
        for ks, kh, lo, hi in (("k_sow", "k_harv", yi * 12 + 5, yi * 12 + 10), ("r_sow", "r_harv", yi * 12 - 1, yi * 12 + 3)):
            lo_c, hi_c = max(lo, 0), min(hi, 71)
            win = s[:, lo_c:hi_c + 1]
            if win.shape[1] == 0:
                continue
            filled = np.where(np.isnan(win), -9, win)
            pk = filled.argmax(axis=1) + lo_c
            ok = (filled.max(axis=1) >= 0.5)
            sow = np.full(n, np.nan)
            for off in range(1, 5):
                pos = pk - off
                v = np.where(pos >= 0, s[ar, np.clip(pos, 0, 71)], np.nan)
                sow = np.where(ok & np.isnan(sow) & (v < 0.25), pos, sow)
            sow = np.where(ok & np.isnan(sow), pk - 2, sow)
            harv = np.full(n, np.nan)
            for off in range(1, 5):
                pos = pk + off
                v = np.where(pos <= 71, s[ar, np.clip(pos, 0, 71)], np.nan)
                harv = np.where(ok & np.isnan(harv) & (v < 0.5), pos, harv)
            known = (pk + 2 <= yi * 12 + lm) if yi == 5 else np.ones(n, bool)
            harv = np.where(ok & np.isnan(harv) & known, pk + 2, harv)
            res[ks][:, yi] = np.where(ok & (sow >= 0), sow % 12 + 1, np.nan)
            res[kh][:, yi] = np.where(ok & np.isfinite(harv), harv % 12 + 1, np.nan)
    return res


def consecutive_stress(sf, lm, thr):
    series = np.concatenate([sf[:, 4, :], sf[:, 5, :lm + 1]], axis=1) if lm >= 0 else sf[:, 4, :]
    run = np.zeros(sf.shape[0], int)
    alive = np.ones(sf.shape[0], bool)
    for j in range(series.shape[1] - 1, -1, -1):
        hit = alive & (series[:, j] >= thr)
        run += hit
        alive = hit
    return run


def slope_rows(y):
    n, k = y.shape
    x = np.arange(k, dtype=float)
    m = np.isfinite(y)
    cnt = m.sum(1)
    yy = np.where(m, y, 0.0)
    xm = (m * x).sum(1) / np.where(cnt > 0, cnt, np.nan)
    ym = yy.sum(1) / np.where(cnt > 0, cnt, np.nan)
    num = (m * (x - xm[:, None]) * (yy - ym[:, None])).sum(1)
    den = (m * (x - xm[:, None]) ** 2).sum(1)
    return np.where(cnt >= 3, _div(num, den), np.nan), ym


def trap(v, lo, hi, alo, ahi):
    v = np.asarray(v, dtype=float)
    return np.where(np.isnan(v), np.nan, np.where((v >= lo) & (v <= hi), 100.0, np.where((v >= alo) & (v < lo), (v - alo) / (lo - alo) * 100,
                    np.where((v > hi) & (v <= ahi), (ahi - v) / (ahi - hi) * 100, 0.0))))


def suitability(ph, rain_annual, texture):
    out = {}
    for crop, rq in CROP_REQ.items():
        s_ph, s_rain = trap(ph, *rq["ph"]), trap(rain_annual, *rq["rain"])
        tex = pd.Series(texture).fillna("").astype(str).str.lower()
        s_tex = np.where(tex == "", np.nan, np.where(tex.str.contains(rq["tex"]), 100.0, 60.0))
        parts = np.vstack([s_ph, s_rain, s_tex]).T
        w = np.array([0.35, 0.35, 0.30])
        ok = np.isfinite(parts)
        den = (ok * w).sum(1)
        out[crop] = np.where(den > 0, np.round((np.nan_to_num(parts) * w).sum(1) / np.where(den > 0, den, np.nan), 0), np.nan)
    return out


def mname(v):
    return np.array([MON[int(x) - 1] if x == x else None for x in v], dtype=object)


def build_trend_frames():
    n = NT
    keys = pd.DataFrame({"State": T["State"].values, "District": T["District"].values, "Tehsil": T["Tehsil"].values})
    M = DATA["TEHSIL_MASTER"]
    cult = M["Cultivable Area Ha"].astype(float).to_numpy()
    lm_n, lm_r = last_month_with(TR["ndvi"][:, 5, :]), last_month_with(TR["rain"][:, 5, :])
    lm_s, lm_g = last_month_with(TR["soil"][:, 5, :]), last_month_with(TR["gw"][:, 5, :])
    lm_i = last_month_with(TR["irrig"][:, 5, :])
    stress_ha = TR["stress"] * cult[:, None, None]
    irrig_ha = TR["irrig"] * cult[:, None, None]
    gw_anom = TR["gw"] - np.nanmean(TR["gw"].reshape(n, -1), axis=1)[:, None, None]
    ARR = {"rain": TR["rain"], "cdd": TR["cdd"], "ndvi": TR["ndvi"], "stress_area": stress_ha, "irrig_area": irrig_ha, "soil": TR["soil"], "gw_anom": gw_anom}
    F = {}

    # ---------- A: soil ----------
    a = pd.concat([keys, wide(TR["soil"], "Soil moisture m3/m3", "mean", 3, lm_s)], axis=1)
    a["Soil moisture percentile (latest month vs 5Y)"] = np.round(latest_percentile(TR["soil"], lm_s), 0)
    rain_ann = np.nansum(np.nanmean(TR["rain"][:, :5, :], axis=1), axis=1)
    rain_ann = np.where(np.isfinite(TR["rain"][:, :5, :]).any(axis=(1, 2)), rain_ann, np.nan)
    S = DATA["SOIL_INTELLIGENCE"]
    ph = pd.to_numeric(S["pH"], errors="coerce").to_numpy() if "pH" in S and len(S) == n else np.full(n, np.nan)
    tex = S["Soil Texture"].to_numpy() if "Soil Texture" in S and len(S) == n else np.full(n, None)
    sut = suitability(ph, rain_ann, tex)
    for crop, v in sut.items():
        a[f"Suitability {crop} (0-100)"] = v
    sm = pd.DataFrame(sut)
    a["Best-suited crops (top 3)"] = [", ".join(sm.columns[np.argsort(-np.nan_to_num(r, nan=-1))[:3]]) if np.isfinite(r).any() else None for r in sm.to_numpy()]
    a["Source"] = "ERA5-Land soil moisture (model); suitability = reference rules on SoilGrids pH/texture + IMD rainfall (editable CROP_REQ)"
    a["Source Date"] = FRESH_RUN.get("MODIS (via GEE)") or TODAY
    F["A_SOIL"] = a

    # ---------- B: water & irrigation ----------
    b = pd.concat([keys, wide(irrig_ha, "Irrigated area proxy ha", "mean", 0, lm_i), wide(gw_anom, "Groundwater storage anomaly mm", "mean", 1, lm_g)], axis=1)
    W = DATA["WATER_IRRIGATION"]
    for col in ("Irrigated Area Ha", "Rainfed Area Ha", "Water Stress", "Groundwater Trend"):
        b[f"Latest {col}"] = W[col].to_numpy() if col in W and len(W) == n else None
    peak = np.nanmax(np.where(np.isnan(irrig_ha[:, :5, :]), -1, irrig_ha[:, :5, :]).reshape(n, -1), axis=1)
    b["Rainfed area proxy ha (cultivable - peak irrigated proxy)"] = np.round(np.where(peak >= 0, np.clip(cult - peak, 0, None), np.nan), 0)
    b["Groundwater LEVEL"] = "not available (no open station data); storage anomaly shown instead"
    b["Source"] = "Irrigation = satellite PROXY (MODIS NDVI > 0.40 in months with < 40 mm CHIRPS rain); groundwater = GLDAS-2.2 storage anomaly (model); latest columns from the main workbook"
    b["Source Date"] = FRESH_RUN.get("GLDAS (via GEE)") or TODAY
    F["B_WATER"] = b

    # ---------- C: weather ----------
    c = pd.concat([keys, wide(TR["rain"], "Rainfall mm", "sum", 1, lm_r), wide(TR["cdd"], "Longest dry spell days", "max", 0, lm_r)], axis=1)
    mean_m = np.nanmean(TR["rain"][:, :5, :], axis=1)
    c["Rainy months (>=50 mm) in a typical year"] = (mean_m >= 50).sum(axis=1)
    c["Rainfall months (typical wet season)"] = [", ".join(MON[j] for j in np.flatnonzero(r >= 50)) or "none >= 50 mm" for r in mean_m]
    ann = np.nansum(TR["rain"][:, :5, :], axis=2)
    c["Annual rainfall mm (5Y mean)"] = np.round(np.nanmean(np.where(ann > 0, ann, np.nan), axis=1), 0)
    c["Annual rainfall variability CV %"] = np.round(_div(np.nanstd(np.where(ann > 0, ann, np.nan), axis=1), np.nanmean(np.where(ann > 0, ann, np.nan), axis=1)) * 100, 0)
    dk = (keys["State"] + "|" + keys["District"]).to_numpy()
    c["Forecast validation % (rain / no-rain agreement, last 30 days)"] = [FV.get(k_, [np.nan, 0])[0] for k_ in dk]
    c["Validation days"] = [FV.get(k_, [np.nan, 0])[1] for k_ in dk]
    _rs = "; ".join(f"{k.split('|')[1]}: {v}" for k, v in sorted(TR_SRC.items()) if k.startswith("rain|"))
    c["Source"] = ("IMD gridded rainfall (0.25 deg) monthly totals" + (f" [by year - {_rs}]" if _rs else "") +
                   "; dry spell = days with < 2.5 mm (IMD only); validation = Open-Meteo archived day-ahead forecast vs IMD at district centre")
    c["Source Date"] = FRESH_RUN.get("IMD gridded") or TODAY
    F["C_WEATHER"] = c

    # ---------- D: satellite ----------
    d = pd.concat([keys, wide(TR["ndvi"], "NDVI", "mean", 3, lm_n)], axis=1)
    pc = latest_percentile(TR["ndvi"], lm_n)
    d["NDVI percentile (latest month vs 5Y)"] = np.round(pc, 0)
    ytdp = np.nanmean(np.vstack([latest_percentile(TR["ndvi"], j) for j in range(lm_n + 1)]), axis=0) if lm_n >= 0 else np.full(n, np.nan)
    d["NDVI percentile (YTD average)"] = np.round(ytdp, 0)
    tot_ha = M["Total Area Ha"].astype(float).to_numpy()
    d["Crop pixel coverage % (cultivable / total area)"] = np.round(np.minimum(100, _div(cult, tot_ha) * 100), 1)
    d["Cropland 250 m pixels (est.)"] = np.round(cult / 6.25, 0)
    d["Source"] = "MODIS MOD13Q1 NDVI (Earth Engine), cropland pixels only (MODIS land cover); coverage from CoRE Stack cultivable area"
    d["Source Date"] = FRESH_RUN.get("MODIS (via GEE)") or TODAY
    F["D_SATELLITE"] = d

    # ---------- F: stage (and timing used by E) ----------
    codes, nrm = phenology_codes(TR["ndvi"])
    f = keys.copy()
    for yi, y in enumerate(YEARS + [CUR]):
        for m in range(12 if yi < 5 else lm_n + 1):
            cc = codes[:, yi, m]
            f[f"Stage {y}-{MON[m]}"] = np.where(cc >= 0, np.array(STAGE_NAMES, dtype=object)[np.clip(cc, 0, 4)], None)
    cur, months_in, typ = stage_summary(codes, lm_n)
    f["Current crop stage (satellite phenology)"] = np.where(cur >= 0, np.array(STAGE_NAMES, dtype=object)[np.clip(cur, 0, 4)], None)
    f["Expected next stage"] = np.where(cur >= 0, np.array(STAGE_NAMES, dtype=object)[np.vectorize(lambda x: NEXT_STAGE.get(int(x), 0))(np.clip(cur, 0, 4))], None)
    f["Months in current stage so far"] = np.where(cur >= 0, months_in, np.nan)
    f["Typical duration of this stage (months, Years 1-5)"] = np.round(typ, 1)
    tim = season_timing(nrm, lm_n)
    for key, lab in (("k_sow", "Kharif sowing month"), ("k_harv", "Kharif harvest month"), ("r_sow", "Rabi sowing month"), ("r_harv", "Rabi harvest month")):
        for yi, y in enumerate(YEARS + [CUR]):
            if yi == 5 and key.endswith("harv") and not np.isfinite(tim[key][:, 5]).any():
                continue
            f[f"{lab} {y}"] = mname(tim[key][:, yi])
    f["Source"] = "Derived from MODIS NDVI curve (cropland-wide satellite phenology proxy, not crop-specific observations)"
    f["Source Date"] = FRESH_RUN.get("MODIS (via GEE)") or TODAY
    F["F_STAGE"] = f

    # ---------- G: crop health & loss ----------
    g = pd.concat([keys, wide(stress_ha, "Area under stress ha", "mean", 0, lm_n)], axis=1)
    sf_now = TR["stress"][:, 5, lm_n] if lm_n >= 0 else np.full(n, np.nan)
    g["Area under stress now ha"] = np.round(np.where(np.isfinite(sf_now), sf_now * cult, np.nan), 0)
    g["Share of cultivable area under stress %"] = np.round(sf_now * 100, 0)
    g["Consecutive stress months (to latest)"] = consecutive_stress(TR["stress"], lm_n, STRESS_FRAC_THRESHOLD)
    ann_max = np.concatenate([np.nanmax(np.where(np.isnan(stress_ha[:, :5, :]), -1, stress_ha[:, :5, :]), axis=2),
                              np.nanmax(np.where(np.isnan(stress_ha[:, 5:, :]), -1, stress_ha[:, 5:, :]), axis=2)], axis=1)
    ann_max = np.where(ann_max < 0, np.nan, ann_max)
    sl, mu = slope_rows(ann_max)
    rel = _div(sl, np.maximum(mu, 1.0))
    g["Crop-loss trend (peak stressed area, Years 1-5 + YTD)"] = np.where(np.isnan(rel), "n/a", np.where(rel > 0.10, "Rising", np.where(rel < -0.10, "Falling", "Stable")))
    # reasons for the latest stress
    if lm_n >= 0:
        rdev = (TR["rain"][:, 5, lm_n] - np.nanmean(TR["rain"][:, :5, lm_n], axis=1)) / np.where(np.nanmean(TR["rain"][:, :5, lm_n], axis=1) > 0, np.nanmean(TR["rain"][:, :5, lm_n], axis=1), np.nan) * 100
        spc = latest_percentile(TR["soil"], min(lm_n, lm_s)) if lm_s >= 0 else np.full(n, np.nan)
        cdd_now = TR["cdd"][:, 5, lm_n]
    else:
        rdev, spc, cdd_now = (np.full(n, np.nan),) * 3
    pest = DATA["PEST_DISEASE"]
    pest_d = {(norm(a_), norm(b_)) for a_, b_ in zip(pest["State"], pest["District"])} if len(pest) else set()
    reasons = []
    for i in range(n):
        if not (sf_now[i] >= STRESS_FRAC_THRESHOLD):
            reasons.append(None)
            continue
        r = []
        if rdev[i] <= -25:
            r.append(f"Rainfall deficit ({rdev[i]:+.0f}% vs 5Y)")
        if rdev[i] >= 75:
            r.append(f"Excess rain ({rdev[i]:+.0f}% vs 5Y) - waterlogging/flood risk")
        if spc[i] <= 20:
            r.append("Low soil moisture (bottom 20% of 5Y)")
        if cdd_now[i] >= 15:
            r.append(f"Long dry spell ({cdd_now[i]:.0f} days)")
        if (norm(keys.at[i, "State"]), norm(keys.at[i, "District"])) in pest_d:
            r.append("Pest/disease reports in district (news, unverified)")
        reasons.append("; ".join(r) if r else "No clear driver in rainfall / soil data - verify in the field")
    g["Reason for loss / stress (rule-based)"] = reasons
    flagged = np.where(sf_now >= STRESS_FRAC_THRESHOLD)[0]
    aff = np.full(n, None, dtype=object)
    if len(flagged):
        fd = pd.DataFrame({"i": flagged, "dk": dk[flagged], "ha": (sf_now * cult)[flagged], "t": keys["Tehsil"].to_numpy()[flagged]})
        for _, grp in fd.groupby("dk"):
            top = grp.sort_values("ha", ascending=False).head(5)
            txt = "; ".join(f"{t_} ({h:,.0f} ha)" for t_, h in zip(top["t"], top["ha"]))
            aff[grp["i"].to_numpy()] = txt
    g["Affected locations (district top 5 stressed tehsils)"] = aff
    g["Source"] = f"MODIS NDVI: share of cropland pixels below {NDVI_STRESS_RATIO:.0%} of their 5-year normal x cultivable area (proxy); reasons use IMD rain, ERA5 soil moisture, dry spells, news"
    g["Source Date"] = FRESH_RUN.get("MODIS (via GEE)") or TODAY
    F["G_HEALTH_LOSS"] = g

    # ---------- E: crop sowing (APY history, district split to tehsils) + satellite timing ----------
    F["E_SOWING"] = sowing_frame(keys, tim, lm_n)
    return F, ARR


def sowing_frame(keys, tim, lm_n):
    n = NT
    cols_t = {}
    for key, lab in (("k_sow", "Satellite kharif sowing month"), ("k_harv", "Satellite kharif harvest month"), ("r_sow", "Satellite rabi sowing month"), ("r_harv", "Satellite rabi harvest month")):
        cols_t[lab] = mname(np.array([np.nanmedian(tim[key][i, :5]) if np.isfinite(tim[key][i, :5]).any() else np.nan for i in range(n)]))
    nd = TR["ndvi"]
    jj = list(range(5, lm_n + 1)) if lm_n >= 5 else []
    cur_g = np.nanmean(nd[:, 5, jj], axis=1) if jj else np.full(n, np.nan)
    past_g = np.nanmean(nd[:, :5, :][:, :, jj], axis=(1, 2)) if jj else np.full(n, np.nan)
    cols_t["Current-season greenness vs 5Y % (Jun to latest)"] = np.round((_div(cur_g, past_g) - 1) * 100, 1)
    tdf = pd.concat([keys, pd.DataFrame(cols_t)], axis=1)
    tdf["pos"] = np.arange(n)
    if "GS" not in globals() or GS is None or len(GS) == 0:
        print("  E_SOWING: APY history (GS) not in memory - crop-level area history skipped; satellite timing only")
        tdf["Source"] = "Satellite phenology proxy (MODIS NDVI)"
        tdf["Source Date"] = FRESH_RUN.get("MODIS (via GEE)") or TODAY
        return tdf.drop(columns="pos")
    gs = GS.copy()
    yrs = sorted(gs["Year"].unique())[-5:]
    gs = gs[gs["Year"].isin(yrs)]
    gs["sk"] = gs["Season"].map(season_key)
    base = ["State", "District", "Crop"]
    tot = gs.groupby(base + ["Year"])["area"].sum().unstack("Year").reindex(columns=yrs)
    kh = gs[gs["sk"] == "kharif"].groupby(base + ["Year"])["area"].sum().unstack("Year").reindex(columns=yrs)
    rb = gs[gs["sk"] == "rabi"].groupby(base + ["Year"])["area"].sum().unstack("Year").reindex(columns=yrs)
    wide_ = pd.concat([tot.add_prefix("Area all seasons ha "), kh.add_prefix("Kharif area ha "), rb.add_prefix("Rabi area ha ")], axis=1).reset_index()
    mp = match_district_names(list(zip(wide_["State"], wide_["District"])))
    wide_["lg"] = [mp.get(k_) for k_ in zip(wide_["State"], wide_["District"])]
    ok = wide_[wide_["lg"].notna()].copy()
    ok["State"], ok["District"] = ok["lg"].map(lambda x: x[0]), ok["lg"].map(lambda x: x[1])
    val_cols = [c for c in ok.columns if c.startswith(("Area all", "Kharif area", "Rabi area"))]
    ok = ok.groupby(["State", "District", "Crop"], as_index=False)[val_cols].sum(min_count=1)
    sh = tehsil_shares()[["State", "District", "Tehsil", "share", "pos"]]
    m = ok.merge(sh, on=["State", "District"], how="inner")
    for c in val_cols:
        m[c] = (m[c] * m["share"]).round(1)
    m = m.drop(columns=["share"]).merge(tdf.drop(columns=["State", "District", "Tehsil"]), on="pos", how="left")
    last = yrs[-1]
    m["Latest reported year"] = last
    m[f"Previous-year sowing ha ({last})"] = m[f"Area all seasons ha {last}"]
    cc = DATA["CROP_CALENDAR"]
    hw = {(a_, b_, c_): d_ for a_, b_, c_, d_ in zip(cc["State"], cc["Crop"], cc["Season"], cc["Harvest Window"])} if len(cc) else {}
    m["Reference harvest window - kharif (crop calendar)"] = [hw.get((s_, c_, "Kharif")) for s_, c_ in zip(m["State"], m["Crop"])]
    m["Reference harvest window - rabi (crop calendar)"] = [hw.get((s_, c_, "Rabi")) for s_, c_ in zip(m["State"], m["Crop"])]
    m["Source"] = "DES APY area by year/season (district, split to tehsils by cropland share = ESTIMATE); timing = MODIS phenology proxy; windows = reference calendar"
    m["Source Date"] = FRESH_RUN.get("DES APY") or TODAY
    first = ["State", "District", "Tehsil", "Crop"]
    return m.drop(columns=["pos"]).reindex(columns=first + [c for c in m.columns if c not in first])


TREND_TABS = {
    "A_SOIL": ("Soil moisture and crop suitability", "Soil moisture month by month for Years 1-5 and this year to date (ERA5-Land model, ~11 km), percentile of the latest month, and a 0-100 suitability score per crop from pH, texture and rainfall using editable reference rules (CROP_REQ)."),
    "B_WATER": ("Irrigated area, rainfed area and groundwater", "Irrigated area is a satellite PROXY (no official monthly series exists). Groundwater LEVEL has no open source; the GLDAS storage anomaly (mm) shows the direction of change. Latest official/proxy areas from the main workbook are included."),
    "C_WEATHER": ("Rainfall, dry spells and forecast validation", "IMD monthly rainfall for Years 1-5 and YTD, longest dry spell per month, typical wet-season months, variability, and how often the archived day-ahead forecast matched IMD on rain / no rain over the last 30 days."),
    "D_SATELLITE": ("NDVI history, percentile and crop pixel coverage", "Monthly cropland NDVI (MODIS) for Years 1-5 and YTD, where the latest month ranks against the same month in the previous 5 years, and how much of the tehsil is cropland."),
    "E_SOWING": ("Sowing history, crop-wise area and harvest windows", "Crop-wise area for the last 5 official years (DES APY, split to tehsils = estimate), satellite sowing/harvest timing, current-season greenness versus normal, and reference harvest windows. Official sowing is yearly/seasonal - there is no monthly series."),
    "F_STAGE": ("Crop stage by month, current and next stage", "Cropland-wide stage inferred from the NDVI curve (fallow, sowing, vegetative, peak, maturity) for every month of Years 1-5 and YTD, with current stage, expected next stage and typical stage duration. Not crop-specific."),
    "G_HEALTH_LOSS": ("Area under stress, consecutive stress, loss trend and reasons", "Share of cropland pixels below 85% of their 5-year NDVI normal x cultivable area, months of continuous stress, loss trend, rule-based reasons (rain, soil moisture, dry spell, pest news) and the most affected tehsils per district."),
}


def long_frame(ARR):
    meta = {"rain": ("Rainfall", "mm", "IMD gridded"), "cdd": ("Longest dry spell", "days", "IMD gridded"), "ndvi": ("NDVI", "index", "MODIS MOD13Q1 (cropland)"),
            "stress_area": ("Area under stress", "ha", "MODIS NDVI vs 5Y normal x cultivable area"),
            "irrig_area": ("Irrigated area (proxy)", "ha", "MODIS NDVI + CHIRPS proxy x cultivable area"),
            "soil": ("Soil moisture", "m3/m3", "ERA5-Land (model)"), "gw_anom": ("Groundwater storage anomaly", "mm", "GLDAS-2.2 CLSM (model)")}
    parts = []
    yrs = np.array(YEARS + [CUR])
    for k, (lab, unit, src) in meta.items():
        nn = ARR[k].shape[0]
        ser = ARR[k].reshape(nn, 72)
        prev = np.concatenate([np.full((nn, 1), np.nan, ser.dtype), ser[:, :-1]], axis=1).reshape(nn, 6, 12)
        lyr = np.concatenate([np.full((nn, 12), np.nan, ser.dtype), ser[:, :-12]], axis=1).reshape(nn, 6, 12)
        cur = ARR[k]
        with np.errstate(all="ignore"):
            mom = cur - prev
            momp = np.where(prev > 0, (cur / prev - 1) * 100, np.nan)
            yoyp = np.where(lyr > 0, (cur / lyr - 1) * 100, np.nan)
        ti, yi, mi = np.nonzero(np.isfinite(ARR[k]))
        parts.append(pd.DataFrame({"State": T["State"].to_numpy()[ti], "District": T["District"].to_numpy()[ti], "Tehsil": T["Tehsil"].to_numpy()[ti],
                                   "Parameter": lab, "Year": yrs[yi], "Month": mi + 1, "Value": cur[ti, yi, mi].astype("float32"), "Unit": unit,
                                   "MoM change": mom[ti, yi, mi].astype("float32"), "MoM %": momp[ti, yi, mi].astype("float32"),
                                   "YoY %": yoyp[ti, yi, mi].astype("float32"), "Source": src}))
    out = pd.concat(parts, ignore_index=True)
    for c in ("State", "District", "Tehsil", "Parameter", "Unit", "Source"):
        out[c] = out[c].astype("category")
    return out


def write_trend_workbook(F):
    TAB_INFO.update({k: v for k, v in TREND_TABS.items()})
    wb = Workbook()
    wb.remove(wb.active)
    from openpyxl.utils.indexed_list import IndexedList
    wb._fonts = IndexedList([Font(name=FONT, size=10)])
    ws = wb.create_sheet("TREND_README")
    ws["A1"] = f"AGRI TREND MODULE - inputs for detailed forecasting   |   built {TODAY:%Y-%m-%d}"
    ws["A1"].font = Font(name=FONT, size=14, bold=True)
    rows_ = [
        ("Layout", f"One row per tehsil (E_SOWING: per tehsil x crop). Year 1..5 = {YEARS[0]}..{YEARS[-1]} month by month, then {CUR} months so far, then YTD, 5-year-average YTD and % difference, plus latest-month value, month-on-month change and change vs the same month last year. The tidy long file Agri_Trend_Long.parquet (tehsil, parameter, year, month, value) is the better input for a model."),
        ("Authentic vs derived", "Official: IMD rainfall, DES APY areas. Models/satellite: MODIS NDVI, ERA5-Land soil moisture, GLDAS groundwater storage. PROXY: irrigated area, stress area, stage, sowing/harvest timing. Reference rules: crop suitability, harvest windows. News: pest reasons (unverified)."),
        ("Not available openly", "Monthly official irrigated area, groundwater LEVEL (m below ground), observed crop stage and observed sowing progress per tehsil. These cells are not invented; proxies are labelled."),
        ("Resolution", f"Earth Engine statistics use {TREND_SCALE} m sampling over cropland pixels; soil moisture ~11 km and groundwater ~25 km are coarse (regional signal). Year-to-date uses months up to {MON[CM - 1]} {CUR}."),
        ("Re-run", "Series are cached (trend_*.npz and gee_*.json checkpoints). Delete them to refresh; re-running continues after a disconnect."),
    ]
    for r, (a_, b_) in enumerate(rows_, start=3):
        ws.cell(row=r, column=1, value=a_).font = Font(name=FONT, bold=True)
        c = ws.cell(row=r, column=2, value=b_)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 58
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 24, 130
    for name, df in F.items():
        HEADERS[name] = list(df.columns)
        DATA[name] = df
        w2 = wb.create_sheet(name)
        w2.sheet_properties.tabColor = "1565C0"
        banner(w2, name, len(df.columns), "1565C0", f"Rows: {len(df):,}  |  Source: {df['Source'].iloc[0] if len(df) else ''}"[:420])
        style_header(w2, list(df.columns), "1565C0")
        write_data_sheet(w2, name, list(df.columns))
        print(f"  wrote {name}: {len(df):,} rows x {len(df.columns)} columns")
    wb.save(TREND_FILE)
    print(f"Saved {TREND_FILE} ({os.path.getsize(TREND_FILE) / 1e6:.0f} MB)")


if not any(tr_has(k) for k in TR_KEYS):
    print("No trend series were collected - nothing to build. FAILED steps from CELL 15:")
    for r_ in RECOVERY_LOG:
        if r_["Step"] == "CELL 15" and r_["Status"] == "FAILED":
            print(f"  {r_['Source']}: {r_['Message']}")
else:
    ram("before trend build")
    F_TREND, ARR_TREND = build_trend_frames()
    try:
        long_df = long_frame(ARR_TREND)
        long_df.to_parquet(TREND_LONG, index=False)
        print(f"Saved {TREND_LONG}: {len(long_df):,} rows")
    except Exception as e:
        print("Parquet failed (", type(e).__name__, ") - writing CSV.gz instead")
        long_df.to_csv("Agri_Trend_Long.csv.gz", index=False)
    F_TREND["SOURCE_RECOVERY_LOG"] = pd.DataFrame(RECOVERY_LOG).reindex(columns=HEADERS["SOURCE_RECOVERY_LOG"])   # full log incl. trend steps
    _tc = sum(len(df_) * len(df_.columns) for df_ in F_TREND.values())
    print(f"Trend workbook size: {_tc:,} cells (Google Sheets limit 10,000,000)")
    write_trend_workbook(F_TREND)
    TREND_SHEET_URL = None
    if TREND_TO_SHEETS:
        if _tc > SHEETS_CELL_LIMIT:
            print("  too large for one Google Sheet - uploading E_SOWING separately")
            _e = F_TREND.pop("E_SOWING")
            write_trend_workbook(F_TREND)
            TREND_SHEET_URL = upload_to_google_sheets(TREND_FILE, f"Agri Trend Module {TODAY}")
            write_trend_workbook({"E_SOWING": _e})
            os.replace(TREND_FILE, "Agri_Trend_Sowing.xlsx")
            upload_to_google_sheets("Agri_Trend_Sowing.xlsx", f"Agri Trend Module - E_SOWING {TODAY}")
            F_TREND["E_SOWING"] = _e
            write_trend_workbook(F_TREND)                 # full local file again
        else:
            TREND_SHEET_URL = upload_to_google_sheets(TREND_FILE, f"Agri Trend Module {TODAY}")


# ===== CELL 17 : FINAL SUMMARY - all Google Sheet links, files, reliability, gaps =====
print("=" * 70)
print(f"AGRI INTELLIGENCE RUN SUMMARY  |  {TODAY}")
print("=" * 70)
print("GOOGLE SHEETS:")
for _k, _v in SHEET_URLS.items():
    print(f"  {_k}:\n    {_v or 'upload failed - download the xlsx from the Files panel and open it in Google Sheets'}")
if not SHEET_URLS:
    print("  none created yet - run CELL 14 (and CELL 16 for the trend module)")
print("FILES (Colab Files panel):", ", ".join(f for f in (OUTPUT_FILE, "Agri_Validation_India.xlsx", "Agri_Intelligence_CSV.zip",
                                                         "Agri_Trend_Module.xlsx", "Agri_Trend_Sowing.xlsx", "Agri_Trend_Long.parquet",
                                                         "Agri_Trend_Long.csv.gz") if os.path.exists(f)))
_R = DATA.get("DATA_RELIABILITY")
if _R is not None and len(_R):
    print("DATA RELIABILITY:", _R["Reliability Band"].value_counts().to_dict(),
          f"| median score {pd.to_numeric(_R['Data Reliability Score'], errors='coerce').median():.0f}")
_V = DATA.get("DATA_VALIDATION")
if _V is not None and len(_V):
    print("VALIDATION (govt vs satellite):", _V["Agreement"].value_counts().to_dict())
_G = DATA.get("DATA_GAP_REGISTER")
if _G is not None and len(_G):
    print("GAPS:", _G.groupby("Gap Type")["Rows Affected"].sum().sort_values(ascending=False).head(8).to_dict())
_A = DATA.get("AI_ESTIMATION_REGISTER")
print("AI / regex candidate figures (review only):", 0 if _A is None else len(_A), f"| AI calls used: {_AI_N[0]} of {AI_MAX_CALLS}")
_fails = [r_ for r_ in RECOVERY_LOG if r_["Status"] == "FAILED"]
print(f"SOURCE ATTEMPTS: {len(RECOVERY_LOG)} logged, {len(_fails)} failed" + (":" if _fails else ""))
for r_ in _fails[:15]:
    print(f"  [{r_['Step']}] {r_['Source']}: {r_['Message'][:110]}")
