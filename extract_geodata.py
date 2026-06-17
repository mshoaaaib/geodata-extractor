#!/usr/bin/env python3
"""
GeoData Extractor — Excel Batch Mode

Reads an Excel file with flood event data and extracts 15 geospatial features
per row, saving results to a new Excel file.

Required columns (case-insensitive):
  lat / latitude
  lon / longitude / long
  flood_date / Flood_Date / FloodDate   (YYYY-MM-DD)
  end_date   / End_Date   / EndDate     (YYYY-MM-DD)

Features extracted:
  Static (coordinates only):
    elevation (m)       — OpenTopoData SRTM 30m
    slope (°)           — Horn's method from 3×3 elevation grid
    aspect (°)          — Horn's method, 0=N clockwise
    curvature           — Laplacian (positive = concave upward)
    twi                 — Topographic Wetness Index (local approximation)
    dt_river (m)        — Haversine to nearest OSM river/stream
    dt_drainage (m)     — Haversine to nearest OSM canal/drain/ditch
    dt_roads (m)        — Haversine to nearest OSM road

  Date-sensitive:
    ndvi                — Sentinel Hub Statistical API, ≤45 d before Flood_Date
    ndbi                — same window as NDVI
    impervious_pct      — Copernicus HRL (Europe) or built-up estimate via SH
    population_per_km2  — WorldPop ArcGIS ImageServer identify
    rainfall_mm         — Open-Meteo ERA5, cumulative over flood duration
    soil_texture        — USDA class, OpenLandMap/SoilGrids 250m
    lulc_class          — ESA WorldCover 2021 (10m) / MODIS MCD12Q1 fallback

Install:
  pip install pandas openpyxl requests numpy

Usage:
  python3 extract_geodata.py floods.xlsx
"""

import sys
import os
import math
import time
import io
from datetime import datetime, timedelta
from typing import Optional

import numpy as np
import pandas as pd
import requests

try:
    import ee
    _GEE_PROJECT = os.environ.get("GEE_PROJECT", "fleet-furnace-348411")
    ee.Initialize(project=_GEE_PROJECT)
    _GEE_AVAILABLE = True
except Exception as _gee_err:
    _GEE_AVAILABLE = False
    print(f"[GEE] Not available — impervious surface will fall back: {_gee_err}")

# ── API Credentials (set via environment variables or .streamlit/secrets.toml) ─
OPENTOPOGRAPHY_API_KEY    = os.environ.get("OPENTOPOGRAPHY_API_KEY", "")
SENTINELHUB_CLIENT_ID     = os.environ.get("SH_CLIENT_ID", "")
SENTINELHUB_CLIENT_SECRET = os.environ.get("SH_CLIENT_SECRET", "")

# ── Settings ──────────────────────────────────────────────────────────────────
DEM_DATASET      = "SRTMGL1"
BBOX_BUFFER      = 0.15      # degrees buffer for OSM queries (~16 km)
SH_BBOX_BUFFER   = 0.02      # degrees buffer for Sentinel Hub NDVI/NDBI (~2 km)
SH_RES_DEG       = 0.0001    # Sentinel Hub resx/resy in degrees (≈ 11 m); must be
                              # in CRS units (degrees) when bbox CRS is EPSG:4326
TERRAIN_CELL_M   = 30.0      # grid spacing for finite-difference terrain (≈ SRTM res)
HS_CELL_M        = 90.0      # HydroSHEDS 03-arc-second cell size in metres
NDVI_DAYS_BEFORE = 60        # days before Flood_Date to start Sentinel-2 search
MAX_CLOUD_COVER  = 30        # % maximum cloud cover for Sentinel-2
TWI_CAP          = 20.0      # cap TWI at this value for flat terrain

# ── Column name aliases ───────────────────────────────────────────────────────
_COL_ALIASES = {
    "year":  {"year"},
    "month": {"month"},
    "day":   {"day"},
    "location":   {"location", "area", "area_name", "place", "city",
                   "district", "region", "address", "name", "site"},
    "lat":        {"lat", "latitude", "y", "lat_approx"},
    "lon":        {"lon", "long", "longitude", "x", "lot_approx"},
    "flood_date": {"flood_date", "flooddate", "start_date", "startdate", "date_start", "date"},
    "end_date":   {"end_date", "enddate", "flood_end", "date_end", "end"},
}

# ── Feature column names in output ───────────────────────────────────────────
FEATURE_COLS = [
    "elevation", "slope", "aspect", "curvature", "twi",
    "dt_river", "dt_drainage", "dt_roads",
    "ndvi", "ndbi", "impervious_pct", "population_per_km2", "rainfall_mm",
    "soil_texture", "lulc_class",
]

# Human-readable column headers with units shown in parentheses
FEATURE_COL_LABELS = {
    "elevation":          "Elevation (m)",
    "slope":              "Slope (°)",
    "aspect":             "Aspect (°)",
    "curvature":          "Curvature (1/m)",
    "twi":                "TWI (unitless)",
    "dt_river":           "DTRiver (m)",
    "dt_drainage":        "DTDrainage (m)",
    "dt_roads":           "DTRoads (m)",
    "ndvi":               "NDVI (index, -1 to 1)",
    "ndbi":               "NDBI (index, -1 to 1)",
    "impervious_pct":     "Impervious Surface (%)",
    "population_per_km2": "Population Density (people/km²)",
    "rainfall_mm":        "Cumulative Rainfall (mm)",
    "soil_texture":       "Soil Texture (USDA class)",
    "lulc_class":         "Land Use/Land Cover (ESA WorldCover)",
}

OSM_HEADERS = {"User-Agent": "GeoDataExtractor/1.0 (research)"}

NAN = float("nan")


# ─────────────────────────────────────────────────────────────────────────────
# Utilities
# ─────────────────────────────────────────────────────────────────────────────

def resolve_col(df: pd.DataFrame, key: str) -> Optional[str]:
    """Return the first column in df whose normalised name matches key's aliases."""
    aliases = _COL_ALIASES.get(key, {key.lower()})
    for col in df.columns:
        if col.strip().lower().replace(" ", "_") in aliases:
            return col
    return None


def bbox_from_point(lat: float, lon: float, buf: float = BBOX_BUFFER) -> dict:
    return {"south": lat - buf, "north": lat + buf,
            "west":  lon - buf, "east":  lon + buf}


def haversine_m(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    R = 6_371_000.0
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    a = (math.sin(math.radians(lat2 - lat1) / 2) ** 2
         + math.cos(phi1) * math.cos(phi2)
         * math.sin(math.radians(lon2 - lon1) / 2) ** 2)
    return R * 2 * math.asin(math.sqrt(a))


def nearest_osm_dist(lat: float, lon: float, elements: list) -> float:
    """Minimum Haversine distance from (lat, lon) to any node in OSM way elements."""
    best = float("inf")
    for el in elements:
        for node in el.get("geometry", []):
            d = haversine_m(lat, lon, node["lat"], node["lon"])
            if d < best:
                best = d
    return best if best < float("inf") else NAN


def overpass_query(query: str, timeout_sec: int = 30) -> dict:
    mirrors = [
        "https://overpass-api.de/api/interpreter",
        "https://overpass.openstreetmap.fr/api/interpreter",
    ]
    for url in mirrors:
        try:
            r = requests.post(url, data={"data": query},
                              headers=OSM_HEADERS, timeout=timeout_sec)
            if r.status_code == 200:
                return r.json()
            time.sleep(3)
        except (requests.Timeout, requests.ConnectionError):
            time.sleep(3)
    return {"elements": []}


def parse_date(val) -> Optional[datetime]:
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return None
    if isinstance(val, datetime):
        return val
    if hasattr(val, "to_pydatetime"):
        return val.to_pydatetime()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(str(val).strip(), fmt)
        except ValueError:
            pass
    return None


def parse_ymd_range(year_val, month_val, day_val):
    """
    Parse dates from split Year / Month / Day columns.
    Day may be a single value (e.g. '24') or a range (e.g. '24-27' or '24–27').
    Returns (flood_date, end_date).
    """
    try:
        year  = int(year_val)
        month = int(month_val)
        day_str = str(day_val).strip().replace("–", "-")   # handle en-dash
        if "-" in day_str:
            parts = day_str.split("-")
            start_day = int(parts[0].strip())
            end_day   = int(parts[-1].strip())
        else:
            start_day = end_day = int(day_str)
        flood_date = datetime(year, month, start_day)
        end_date   = datetime(year, month, end_day)
        return flood_date, end_date
    except Exception:
        return None, None


# Cache geocoding results so repeated location names don't make duplicate calls
_geocode_cache: dict = {}

def geocode_location(name: str, country_hint: str = "Pakistan") -> tuple:
    """
    Geocode a place name to (lat, lon) using OSM Nominatim.
    Appends country_hint to the query to bias results (e.g. 'Raja Bazaar, Pakistan').
    Returns (lat, lon) floats or (None, None) if not found.
    """
    key = name.strip().lower()
    if key in _geocode_cache:
        return _geocode_cache[key]

    query = f"{name.strip()}, {country_hint}"
    try:
        r = requests.get(
            "https://nominatim.openstreetmap.org/search",
            params={"q": query, "format": "json", "limit": 1},
            headers=OSM_HEADERS,
            timeout=10,
        )
        r.raise_for_status()
        results = r.json()
        if results:
            lat = float(results[0]["lat"])
            lon = float(results[0]["lon"])
            print(f"    [Geocode] '{name}' → lat={lat:.5f}, lon={lon:.5f}")
            _geocode_cache[key] = (lat, lon)
            time.sleep(1.1)   # Nominatim rate limit: 1 req/sec
            return lat, lon
        else:
            print(f"    [Geocode] '{name}' not found — skipping row")
    except Exception as e:
        print(f"    [Geocode] ERROR for '{name}': {e}")

    _geocode_cache[key] = (None, None)
    return None, None


# ─────────────────────────────────────────────────────────────────────────────
# 1–5  Static Terrain (elevation, slope, aspect, curvature, TWI)
# ─────────────────────────────────────────────────────────────────────────────

def extract_terrain(lat: float, lon: float) -> dict:
    """
    Batch-query a 3×3 elevation grid (TERRAIN_CELL_M spacing) from OpenTopoData,
    then compute terrain derivatives at the centre cell using Horn's method.
    """
    cm = TERRAIN_CELL_M
    dlat = cm / 111_320.0
    dlon = cm / (111_320.0 * math.cos(math.radians(lat)) + 1e-9)

    # Offsets: (row_delta, col_delta); row positive = south, col positive = east
    offsets = [(-1, -1), (-1, 0), (-1, 1),
               ( 0, -1), ( 0, 0), ( 0, 1),
               ( 1, -1), ( 1, 0), ( 1, 1)]
    locations = "|".join(f"{lat + dr*dlat},{lon + dc*dlon}" for dr, dc in offsets)

    for attempt in range(3):
        try:
            r = requests.get(
                "https://api.opentopodata.org/v1/srtm30m",
                params={"locations": locations},
                timeout=30,
            )
            r.raise_for_status()
            results = r.json().get("results", [])
            if len(results) < 9:
                raise ValueError("Incomplete batch response")
            break
        except Exception as e:
            if attempt == 2:
                print(f"    [terrain] FAILED after 3 attempts: {e}")
                return {k: NAN for k in ("elevation", "slope", "aspect", "curvature", "twi")}
            time.sleep(2 ** attempt)

    elevs = [float(res["elevation"] or 0) for res in results]
    Z = np.array(elevs).reshape(3, 3)  # Z[row][col], row 0 = northernmost

    elev = float(Z[1, 1])

    # Horn's finite differences
    dz_dx = ((Z[0,2] + 2*Z[1,2] + Z[2,2]) - (Z[0,0] + 2*Z[1,0] + Z[2,0])) / (8 * cm)
    dz_dy = ((Z[0,0] + 2*Z[0,1] + Z[0,2]) - (Z[2,0] + 2*Z[2,1] + Z[2,2])) / (8 * cm)

    slope_rad = math.atan(math.sqrt(dz_dx**2 + dz_dy**2))
    slope_deg = math.degrees(slope_rad)

    aspect_deg = math.degrees(math.atan2(dz_dy, -dz_dx))
    if aspect_deg < 0:
        aspect_deg += 360.0

    # Laplacian curvature (positive = concave upward / bowl)
    d2x = (Z[1,0] - 2*Z[1,1] + Z[1,2]) / (cm ** 2)
    d2y = (Z[0,1] - 2*Z[1,1] + Z[2,1]) / (cm ** 2)
    curvature = -(d2x + d2y)

    # TWI = ln(A_s / tan(β)); approximate A_s = one cell area (local estimate)
    tan_beta = max(math.tan(slope_rad), 0.001)
    twi = min(math.log(cm / tan_beta), TWI_CAP)

    print(f"    [terrain] elev={elev:.1f}m  slope={slope_deg:.2f}°  "
          f"aspect={aspect_deg:.1f}°  curv={curvature:.5f}  twi={twi:.3f}")
    return {
        "elevation": round(elev, 2),
        "slope":     round(slope_deg, 4),
        "aspect":    round(aspect_deg, 3),
        "curvature": round(curvature, 7),
        "twi":       round(twi, 4),
    }


# ─────────────────────────────────────────────────────────────────────────────
# 1–5  Static Terrain — GEE version (Copernicus DEM GLO30 + HydroSHEDS TWI)
# ─────────────────────────────────────────────────────────────────────────────

def extract_terrain_gee(lat: float, lon: float) -> dict:
    """
    Query a 3×3 elevation grid from Copernicus DEM GLO30 via GEE.
    GLO30 is TanDEM-X based (±1 m vertical accuracy vs SRTM ±16 m).
    TWI uses HydroSHEDS 90 m flow accumulation for proper specific
    catchment area instead of the single-cell local approximation.
    Falls back to OpenTopoData SRTM if GEE is unavailable or fails.
    """
    if not _GEE_AVAILABLE:
        return extract_terrain(lat, lon)

    try:
        cm   = TERRAIN_CELL_M
        dlat = cm / 111_320.0
        dlon = cm / (111_320.0 * math.cos(math.radians(lat)) + 1e-9)

        offsets = [(-1,-1),(-1, 0),(-1, 1),
                   ( 0,-1),( 0, 0),( 0, 1),
                   ( 1,-1),( 1, 0),( 1, 1)]

        features = [
            ee.Feature(
                ee.Geometry.Point([lon + dc * dlon, lat + dr * dlat]),
                {"idx": (dr + 1) * 3 + (dc + 1)}
            )
            for dr, dc in offsets
        ]
        fc      = ee.FeatureCollection(features)
        dem     = ee.ImageCollection("COPERNICUS/DEM/GLO30").select("DEM").mosaic()
        sampled = dem.sampleRegions(collection=fc, scale=30, geometries=False)
        data    = sampled.getInfo()["features"]
        data.sort(key=lambda f: f["properties"]["idx"])
        elevs = [f["properties"]["DEM"] for f in data]

        if len(elevs) < 9:
            raise ValueError(f"Only {len(elevs)}/9 elevation samples")

        Z    = np.array(elevs, dtype=float).reshape(3, 3)
        elev = float(Z[1, 1])

        # Horn's finite differences
        dz_dx = ((Z[0,2]+2*Z[1,2]+Z[2,2]) - (Z[0,0]+2*Z[1,0]+Z[2,0])) / (8 * cm)
        dz_dy = ((Z[0,0]+2*Z[0,1]+Z[0,2]) - (Z[2,0]+2*Z[2,1]+Z[2,2])) / (8 * cm)

        slope_rad = math.atan(math.sqrt(dz_dx**2 + dz_dy**2))
        slope_deg = math.degrees(slope_rad)

        aspect_deg = math.degrees(math.atan2(dz_dy, -dz_dx))
        if aspect_deg < 0:
            aspect_deg += 360.0

        d2x = (Z[1,0] - 2*Z[1,1] + Z[1,2]) / cm**2
        d2y = (Z[0,1] - 2*Z[1,1] + Z[2,1]) / cm**2
        curvature = -(d2x + d2y)

        # TWI with HydroSHEDS flow accumulation
        # TWI = ln(SCA / tan β)  where SCA = flow_acc_cells × cell_area / cell_width
        tan_beta = max(math.tan(slope_rad), 0.001)
        try:
            fa_img    = ee.Image("WWF/HydroSHEDS/03ACC").select("b1")
            fa_result = fa_img.reduceRegion(
                reducer=ee.Reducer.mean(),
                geometry=ee.Geometry.Point([lon, lat]),
                scale=90,
                maxPixels=1,
            )
            fa_val = fa_result.get("b1").getInfo()
            if fa_val and float(fa_val) > 0:
                sca = float(fa_val) * HS_CELL_M   # cells × cell_size = upslope length (m)
                twi = min(math.log(sca / tan_beta), TWI_CAP)
            else:
                raise ValueError("zero/null flow accumulation")
        except Exception:
            twi = min(math.log(cm / tan_beta), TWI_CAP)

        print(f"    [terrain/GLO30] elev={elev:.1f}m  slope={slope_deg:.2f}°  "
              f"aspect={aspect_deg:.1f}°  curv={curvature:.5f}  twi={twi:.3f}")
        return {
            "elevation": round(elev, 2),
            "slope":     round(slope_deg, 4),
            "aspect":    round(aspect_deg, 3),
            "curvature": round(curvature, 7),
            "twi":       round(twi, 4),
        }

    except Exception as e:
        print(f"    [terrain/GLO30] ERROR: {e} — falling back to OpenTopoData SRTM")
        return extract_terrain(lat, lon)


# ─────────────────────────────────────────────────────────────────────────────
# 6  Distance to River
# ─────────────────────────────────────────────────────────────────────────────

def extract_dt_river(lat: float, lon: float, bbox: dict) -> float:
    s, w, n, e = bbox["south"], bbox["west"], bbox["north"], bbox["east"]
    query = (f'[out:json][timeout:25];'
             f'(way["waterway"~"river|stream"]({s},{w},{n},{e}););'
             f'out body geom;')
    elements = overpass_query(query).get("elements", [])
    dist = nearest_osm_dist(lat, lon, elements)
    print(f"    [DTRiver]    {len(elements)} features → {dist:.0f} m")
    return round(dist, 1)


# ─────────────────────────────────────────────────────────────────────────────
# 7  Distance to Drainage
# ─────────────────────────────────────────────────────────────────────────────

def extract_dt_drainage(lat: float, lon: float, bbox: dict) -> float:
    """
    Query OSM for man-made drainage features. Uses a broad tag set to maximise
    coverage in South Asia where OSM tagging is inconsistent:
      waterway = canal | drain | ditch | drainage | drain_line | floodway | spillway
      man_made  = drain | drainage_basin
    Falls back to any remaining waterway (excluding rivers/streams) if still empty.
    """
    s, w, n, e = bbox["south"], bbox["west"], bbox["north"], bbox["east"]

    # Primary: named drainage types
    query = (
        f'[out:json][timeout:30];'
        f'('
        f'way["waterway"~"canal|drain|ditch|drainage|floodway|spillway"]({s},{w},{n},{e});'
        f'way["man_made"~"drain|drainage"]({s},{w},{n},{e});'
        f');'
        f'out body geom;'
    )
    elements = overpass_query(query).get("elements", [])

    # Fallback: any waterway that is not a natural river/stream
    if not elements:
        query_fb = (
            f'[out:json][timeout:30];'
            f'(way["waterway"]({s},{w},{n},{e}););'
            f'out body geom;'
        )
        all_ww = overpass_query(query_fb).get("elements", [])
        # Exclude rivers and streams — those belong to DTRiver
        elements = [
            el for el in all_ww
            if el.get("tags", {}).get("waterway") not in ("river", "stream", "tidal_channel")
        ]

    dist = nearest_osm_dist(lat, lon, elements)
    status = f"{dist:.0f} m" if not math.isnan(dist) else "NaN (no drainage features in OSM)"
    print(f"    [DTDrainage] {len(elements)} features → {status}")
    return round(dist, 1) if not math.isnan(dist) else NAN


# ─────────────────────────────────────────────────────────────────────────────
# 8  Distance to Roads
# ─────────────────────────────────────────────────────────────────────────────

def extract_dt_roads(lat: float, lon: float, bbox: dict) -> float:
    s, w, n, e = bbox["south"], bbox["west"], bbox["north"], bbox["east"]
    query = (f'[out:json][timeout:25];'
             f'(way["highway"]({s},{w},{n},{e}););'
             f'out body geom;')
    elements = overpass_query(query).get("elements", [])
    dist = nearest_osm_dist(lat, lon, elements)
    print(f"    [DTRoads]    {len(elements)} features → {dist:.0f} m")
    return round(dist, 1)


# ─────────────────────────────────────────────────────────────────────────────
# 9–10  NDVI / NDBI  (Sentinel Hub Statistical API)
# ─────────────────────────────────────────────────────────────────────────────

# Token cache shared across rows
_sh_token: dict = {"value": None, "expires": 0.0}


def _get_sh_token() -> str:
    now = time.time()
    if _sh_token["value"] and now < _sh_token["expires"] - 60:
        return _sh_token["value"]
    r = requests.post(
        "https://identity.dataspace.copernicus.eu/auth/realms/CDSE/"
        "protocol/openid-connect/token",
        data={
            "client_id":     SENTINELHUB_CLIENT_ID,
            "client_secret": SENTINELHUB_CLIENT_SECRET,
            "grant_type":    "client_credentials",
        },
        timeout=15,
    )
    r.raise_for_status()
    data = r.json()
    _sh_token["value"]   = data["access_token"]
    _sh_token["expires"] = now + float(data.get("expires_in", 3600))
    return _sh_token["value"]


# Evalscript for Statistical API.
# - No `units` override: SCL is a classification band, not reflectance; mixing
#   units causes a 400 Bad Request from CDSE.
# - No `mosaicking` in setup: the Statistical API controls mosaicking via the
#   request payload, not the evalscript.
# - Cloud masking: SCL classes 3 (cloud shadow), 8 (cloud medium), 9 (cloud
#   high), 10 (thin cirrus) are excluded; everything else is kept.
_SH_EVALSCRIPT = """//VERSION=3
function setup() {
  return {
    input: [{ bands: ["B04","B08","B11","SCL"] }],
    output: [
      { id: "ndvi",     bands: 1, sampleType: "FLOAT32" },
      { id: "ndbi",     bands: 1, sampleType: "FLOAT32" },
      { id: "dataMask", bands: 1 }
    ]
  };
}
function evaluatePixel(s) {
  var cloud = [3,8,9,10];
  if (cloud.indexOf(Math.round(s.SCL)) >= 0) {
    return { ndvi:[NaN], ndbi:[NaN], dataMask:[0] };
  }
  var ndvi = (s.B08 - s.B04) / (s.B08 + s.B04 + 1e-5);
  var ndbi = (s.B11 - s.B08) / (s.B11 + s.B08 + 1e-5);
  return { ndvi:[ndvi], ndbi:[ndbi], dataMask:[1] };
}
"""


def extract_ndvi_ndbi(lat: float, lon: float, flood_date: datetime) -> dict:
    """
    Query Sentinel Hub Statistical API with daily aggregation over the
    NDVI_DAYS_BEFORE window ending on Flood_Date.  Collects mean NDVI/NDBI
    from every acquisition that has valid (non-NaN) pixels, then returns the
    mean of those acquisition-level means — effectively a cloud-masked median
    composite value pre-flood.
    """
    date_to   = flood_date.strftime("%Y-%m-%d")
    date_from = (flood_date - timedelta(days=NDVI_DAYS_BEFORE)).strftime("%Y-%m-%d")

    # Smaller bbox dedicated to SH — keeps pixel count within API limits.
    # resx/resy must be in CRS units (degrees for EPSG:4326), not metres.
    sh_bbox = bbox_from_point(lat, lon, SH_BBOX_BUFFER)

    payload = {
        "input": {
            "bounds": {
                "bbox": [sh_bbox["west"], sh_bbox["south"],
                         sh_bbox["east"], sh_bbox["north"]],
                "properties": {"crs": "http://www.opengis.net/def/crs/EPSG/0/4326"},
            },
            "data": [{
                "type": "sentinel-2-l2a",
                "dataFilter": {
                    "timeRange": {
                        "from": f"{date_from}T00:00:00Z",
                        "to":   f"{date_to}T23:59:59Z",
                    },
                    "maxCloudCoverage": MAX_CLOUD_COVER,
                },
            }],
        },
        "aggregation": {
            "timeRange": {
                "from": f"{date_from}T00:00:00Z",
                "to":   f"{date_to}T23:59:59Z",
            },
            "aggregationInterval": {"of": "P1D"},   # one entry per day with data
            "evalscript": _SH_EVALSCRIPT,
            "resx": SH_RES_DEG,
            "resy": SH_RES_DEG,
        },
        "calculations": {"default": {}},
    }

    try:
        token = _get_sh_token()
        r = requests.post(
            "https://sh.dataspace.copernicus.eu/api/v1/statistics",
            json=payload,
            headers={"Authorization": f"Bearer {token}"},
            timeout=90,
        )
        if not r.ok:
            print(f"    [NDVI/NDBI] ERROR {r.status_code}: {r.text[:300]}")
            return {"ndvi": NAN, "ndbi": NAN}
        resp = r.json()

        intervals = resp.get("data", [])
        if not intervals:
            print(f"    [NDVI/NDBI] No acquisitions in window {date_from} → {date_to}")
            return {"ndvi": NAN, "ndbi": NAN}

        # Collect per-day means, skip days where all pixels were cloudy (NaN)
        ndvi_vals, ndbi_vals = [], []
        for interval in intervals:
            outputs = interval.get("outputs", {})
            for key, store in (("ndvi", ndvi_vals), ("ndbi", ndbi_vals)):
                stats = (outputs.get(key, {})
                         .get("bands", {}).get("B0", {}).get("stats", {}))
                v = stats.get("mean")
                sample_count = stats.get("sampleCount", 0)
                no_data      = stats.get("noDataCount", 0)
                if v is not None and sample_count > no_data:
                    try:
                        fv = float(v)
                        if not math.isnan(fv):
                            store.append(fv)
                    except (TypeError, ValueError):
                        pass

        ndvi = float(np.mean(ndvi_vals)) if ndvi_vals else NAN
        ndbi = float(np.mean(ndbi_vals)) if ndbi_vals else NAN

        n_acq = len(ndvi_vals)
        print(f"    [NDVI] {ndvi:.4f}   [NDBI] {ndbi:.4f}   "
              f"({n_acq} valid acquisitions, window: {date_from} → {date_to})")
        return {
            "ndvi": round(ndvi, 5) if not math.isnan(ndvi) else NAN,
            "ndbi": round(ndbi, 5) if not math.isnan(ndbi) else NAN,
        }

    except Exception as e:
        print(f"    [NDVI/NDBI] ERROR: {e}")
        return {"ndvi": NAN, "ndbi": NAN}


# ─────────────────────────────────────────────────────────────────────────────
# 11  Impervious Surface
# ─────────────────────────────────────────────────────────────────────────────

# GHSL built-up epochs available in GEE
_GHSL_EPOCHS = [1975, 1980, 1985, 1990, 1995, 2000, 2005, 2010, 2015, 2020, 2025]


def extract_impervious(lat: float, lon: float, bbox: dict,
                       flood_year: int) -> float:
    """
    Priority order:
    1. Google Earth Engine — GHSL GHS_BUILT_S, epoch matched to flood year (global).
    2. Copernicus HRL WMS  — Europe only, 2018 reference.
    Returns built-up surface percentage (0–100) or NaN.
    """
    epoch = min(_GHSL_EPOCHS, key=lambda e: abs(e - flood_year))

    # ── Attempt 1: GEE GHSL (global, epoch-matched) ──────────────────────────
    if _GEE_AVAILABLE:
        try:
            point  = ee.Geometry.Point([lon, lat])  # type: ignore[name-defined]
            image  = ee.Image(f"JRC/GHSL/P2023A/GHS_BUILT_S/{epoch}")  # type: ignore[name-defined]
            # GHS_BUILT_S = m² of built-up surface per 100m×100m cell (max 10 000)
            result = image.reduceRegion(
                reducer=ee.Reducer.mean(),  # type: ignore[name-defined]
                geometry=point.buffer(500), # 500 m radius around point
                scale=100,
                maxPixels=1e6,
            ).getInfo()
            raw = result.get("built_surface")
            if raw is not None:
                pct = round(min(100.0, float(raw) / 100.0), 2)
                print(f"    [Impervious] GEE GHSL E{epoch} → {pct:.1f}%")
                return pct
            else:
                print(f"    [Impervious] GEE returned None for E{epoch} — trying fallback")
        except Exception as e:
            print(f"    [Impervious] GEE error: {e} — trying fallback")

    # ── Attempt 2: Copernicus HRL WMS (Europe only) ──────────────────────────
    wms = (
        "https://image.discomap.eea.europa.eu/arcgis/services/GioLand/"
        "HRL_Imperviousness_2018/ImageServer/WMSServer"
        f"?SERVICE=WMS&VERSION=1.3.0&REQUEST=GetMap&LAYERS=0"
        f"&BBOX={bbox['south']},{bbox['west']},{bbox['north']},{bbox['east']}"
        f"&CRS=EPSG:4326&WIDTH=32&HEIGHT=32&FORMAT=image%2Fpng&STYLES="
    )
    try:
        r = requests.get(wms, timeout=20)
        if r.status_code == 200 and r.headers.get("content-type", "").startswith("image"):
            try:
                from PIL import Image  # type: ignore[import-untyped]
                img = Image.open(io.BytesIO(r.content)).convert("L")
                arr = np.array(img, dtype=float)
                valid = arr[arr < 255]
                if len(valid) > 0:
                    val = round(float(np.mean(valid)), 2)
                    print(f"    [Impervious] HRL WMS (Europe) → {val:.1f}%")
                    return val
            except ImportError:
                pass
    except Exception:
        pass

    print("    [Impervious] All sources failed — returning NaN")
    return NAN


# ─────────────────────────────────────────────────────────────────────────────
# 12  Population Density
# ─────────────────────────────────────────────────────────────────────────────

def extract_population(lat: float, lon: float, flood_year: int) -> float:
    """
    Priority order:
    1. GEE — GHSL GHS_POP epoch matched to flood year (global, 100m, incl. Pakistan).
       population_count = people per 100m×100m cell → ×100 → people per km².
    2. WorldPop ArcGIS identify — global fallback.
    Returns people per km².
    """
    epoch = min(_GHSL_EPOCHS, key=lambda e: abs(e - flood_year))

    # ── Attempt 1: GEE GHSL population (global, year-matched) ────────────────
    if _GEE_AVAILABLE:
        try:
            point  = ee.Geometry.Point([lon, lat])  # type: ignore[name-defined]
            image  = ee.Image(f"JRC/GHSL/P2023A/GHS_POP/{epoch}")  # type: ignore[name-defined]
            result = image.reduceRegion(
                reducer=ee.Reducer.mean(),  # type: ignore[name-defined]
                geometry=point.buffer(500),
                scale=100,
                maxPixels=1e6,
            ).getInfo()
            raw = result.get("population_count")
            if raw is not None:
                # GHS_POP = people per 100m×100m cell; ×100 converts to per km²
                density = round(float(raw) * 100, 3)
                print(f"    [Population] GEE GHSL E{epoch} → {density:.2f} people/km²")
                return density
            else:
                print(f"    [Population] GEE returned None for E{epoch} — trying fallback")
        except Exception as e:
            print(f"    [Population] GEE error: {e} — trying fallback")

    # ── Attempt 2: WorldPop ArcGIS (global fallback) ─────────────────────────
    try:
        r = requests.get(
            "https://worldpop.arcgis.com/arcgis/rest/services/"
            "WorldPop_Population_Density_100m/ImageServer/identify",
            params={
                "geometry":     f"{lon},{lat}",
                "geometryType": "esriGeometryPoint",
                "sr":           "4326",
                "f":            "json",
            },
            timeout=20,
        )
        r.raise_for_status()
        val = r.json().get("value")
        if val and val not in ("NoData", "Null", None):
            pop = round(float(val), 3)
            print(f"    [Population] WorldPop ArcGIS → {pop:.2f} people/km²")
            return pop
    except Exception as e:
        print(f"    [Population] WorldPop fallback ERROR: {e}")

    return NAN


# ─────────────────────────────────────────────────────────────────────────────
# 13  Rainfall
# ─────────────────────────────────────────────────────────────────────────────

def extract_rainfall(lat: float, lon: float,
                     start_date: datetime, end_date: datetime) -> float:
    """
    Open-Meteo Archive API (ERA5) — cumulative precipitation over the flood
    event duration (start_date inclusive to end_date inclusive).
    Returns total mm.
    """
    # Start 1 day before the flood date to capture the triggering rainfall
    # that typically peaks overnight before the reported flood start.
    d_from = (start_date - timedelta(days=1)).strftime("%Y-%m-%d")
    d_to   = end_date.strftime("%Y-%m-%d")
    try:
        r = requests.get(
            "https://archive-api.open-meteo.com/v1/archive",
            params={
                "latitude":        lat,
                "longitude":       lon,
                "start_date":      d_from,
                "end_date":        d_to,
                "daily":           "precipitation_sum",
                "timezone":        "UTC",
            },
            timeout=25,
        )
        r.raise_for_status()
        precip = r.json().get("daily", {}).get("precipitation_sum", [])
        total  = sum(p for p in precip if p is not None)
        n_days = len(precip)
        print(f"    [Rainfall]   {total:.1f} mm over {n_days} days ({d_from} → {d_to})")
        return round(total, 2)
    except Exception as e:
        print(f"    [Rainfall] ERROR: {e}")
        return NAN


# ─────────────────────────────────────────────────────────────────────────────
# 13  Rainfall — GEE CHIRPS version (higher resolution than ERA5)
# ─────────────────────────────────────────────────────────────────────────────

def extract_rainfall_chirps(lat: float, lon: float,
                            start_date: datetime, end_date: datetime) -> float:
    """
    CHIRPS Daily via GEE — ~5.5 km resolution, 1981–present.
    Significantly finer grain than ERA5 (9 km), especially for
    localised monsoon events over Pakistan.
    Falls back to ERA5 via Open-Meteo if GEE unavailable or CHIRPS fails.
    """
    if not _GEE_AVAILABLE:
        return extract_rainfall(lat, lon, start_date, end_date)

    # Start 1 day before flood to capture overnight trigger rainfall
    d_from     = (start_date - timedelta(days=1)).strftime("%Y-%m-%d")
    d_to       = end_date.strftime("%Y-%m-%d")
    d_to_excl  = (end_date + timedelta(days=1)).strftime("%Y-%m-%d")

    try:
        chirps = (
            ee.ImageCollection("UCSB-CHG/CHIRPS/DAILY")
            .filterDate(d_from, d_to_excl)   # filterDate end is exclusive
            .select("precipitation")
            .sum()
        )
        result = chirps.reduceRegion(
            reducer=ee.Reducer.mean(),
            geometry=ee.Geometry.Point([lon, lat]),
            scale=5566,    # CHIRPS native pixel ~5.5 km
            maxPixels=1,
        )
        total = result.get("precipitation").getInfo()
        if total is not None:
            total = round(float(total), 2)
            print(f"    [Rainfall/CHIRPS] {total:.1f} mm ({d_from} → {d_to})")
            return total
        raise ValueError("CHIRPS returned None")
    except Exception as e:
        print(f"    [Rainfall/CHIRPS] {e} — falling back to ERA5")
        return extract_rainfall(lat, lon, start_date, end_date)


# ─────────────────────────────────────────────────────────────────────────────
# 14  Soil Texture
# ─────────────────────────────────────────────────────────────────────────────

_USDA_TEXTURE = {
    1: "Clay",  2: "Silty clay",  3: "Silty clay loam",
    4: "Sandy clay",  5: "Sandy clay loam",  6: "Clay loam",
    7: "Silt",  8: "Silt loam",  9: "Loam",
    10: "Sandy loam",  11: "Loamy sand",  12: "Sand",
}


def _usda_class_from_fractions(sand: float, silt: float, clay: float) -> str:
    """USDA texture triangle classification from sand/silt/clay percentages."""
    c, si, sa = clay, silt, sand
    if c >= 40 and si >= 40:   return "Silty clay"
    if c >= 40 and sa <= 45:   return "Clay"
    if c >= 35 and sa >= 45:   return "Sandy clay"
    if c >= 27 and sa <= 20:   return "Silty clay loam"
    if c >= 27 and sa <= 45:   return "Clay loam"
    if c >= 20 and sa >= 45:   return "Sandy clay loam"
    if si >= 80 and c < 12:    return "Silt"
    if si >= 50 and c < 27:    return "Silt loam"
    if c >= 7  and sa <= 52:   return "Loam"
    if sa >= 85:               return "Sand"
    if sa >= 70:               return "Loamy sand"
    return "Sandy loam"


def extract_soil_texture(lat: float, lon: float) -> str:
    """
    USDA soil texture class at 0–5 cm depth.
    Priority:
    1. GEE — OpenLandMap SOL_TEXTURE-CLASS_USDA-TT_M v02 (250 m global)
    2. SoilGrids v2 REST API (250 m) — derives USDA class from sand/silt/clay
    Returns USDA texture class name string, or '' on failure.
    """
    # ── Attempt 1: GEE OpenLandMap ──────────────────────────────────────────
    if _GEE_AVAILABLE:
        try:
            img = ee.Image("OpenLandMap/SOL/SOL_TEXTURE-CLASS_USDA-TT_M/v02").select("b0")
            result = img.reduceRegion(
                reducer=ee.Reducer.mode(),
                geometry=ee.Geometry.Point([lon, lat]).buffer(250),
                scale=250,
                maxPixels=100,
            ).getInfo()
            code = result.get("b0")
            if code is not None:
                label = _USDA_TEXTURE.get(int(round(float(code))),
                                          f"Class {int(round(float(code)))}")
                print(f"    [SoilTexture] OpenLandMap → {label} (code {int(round(float(code)))})")
                return label
        except Exception as e:
            print(f"    [SoilTexture] GEE error: {e} — trying SoilGrids REST")

    # ── Attempt 2: SoilGrids v2 REST API ───────────────────────────────────
    try:
        r = requests.get(
            "https://rest.isric.org/soilgrids/v2.0/properties/query",
            params=[
                ("lon", lon), ("lat", lat),
                ("property", "sand"), ("property", "silt"), ("property", "clay"),
                ("depth", "0-5cm"), ("value", "mean"),
            ],
            timeout=25,
        )
        r.raise_for_status()
        fracs: dict = {}
        for layer in r.json().get("properties", {}).get("layers", []):
            name = layer.get("name")
            val  = layer.get("depths", [{}])[0].get("values", {}).get("mean")
            if name in ("sand", "silt", "clay") and val is not None:
                fracs[name] = float(val) / 10.0   # g/kg → %
        if len(fracs) == 3:
            label = _usda_class_from_fractions(fracs["sand"], fracs["silt"], fracs["clay"])
            print(f"    [SoilTexture] SoilGrids → {label} "
                  f"(sand={fracs['sand']:.1f}% silt={fracs['silt']:.1f}% clay={fracs['clay']:.1f}%)")
            return label
    except Exception as e:
        print(f"    [SoilTexture] SoilGrids ERROR: {e}")

    return ""


# ─────────────────────────────────────────────────────────────────────────────
# 15  Land Use / Land Cover
# ─────────────────────────────────────────────────────────────────────────────

_ESA_WORLDCOVER = {
    10: "Tree cover",         20: "Shrubland",
    30: "Grassland",          40: "Cropland",
    50: "Built-up",           60: "Bare/sparse vegetation",
    70: "Snow and ice",       80: "Permanent water bodies",
    90: "Herbaceous wetland", 95: "Mangroves",
    100: "Moss and lichen",
}

_MODIS_IGBP = {
    1:  "Evergreen needleleaf forest",  2:  "Evergreen broadleaf forest",
    3:  "Deciduous needleleaf forest",  4:  "Deciduous broadleaf forest",
    5:  "Mixed forest",                 6:  "Closed shrubland",
    7:  "Open shrubland",               8:  "Woody savanna",
    9:  "Savanna",                      10: "Grassland",
    11: "Permanent wetland",            12: "Cropland",
    13: "Urban and built-up",           14: "Cropland/natural veg. mosaic",
    15: "Snow and ice",                 16: "Barren/sparsely vegetated",
    17: "Water",
}


def extract_lulc(lat: float, lon: float, flood_year: int) -> str:
    """
    Land Use / Land Cover classification.
    Priority:
    1. GEE — ESA WorldCover v200 2021 (10 m) — best spatial detail
    2. GEE — MODIS MCD12Q1 IGBP (500 m, year-matched 2001–2022) — temporal match
    Returns class name string, or '' on failure.
    """
    if not _GEE_AVAILABLE:
        return ""

    # ── Attempt 1: ESA WorldCover 2021 (10 m) ───────────────────────────────
    try:
        img = ee.Image("ESA/WorldCover/v200/2021").select("Map")
        result = img.reduceRegion(
            reducer=ee.Reducer.mode(),
            geometry=ee.Geometry.Point([lon, lat]).buffer(50),
            scale=10,
            maxPixels=1000,
        ).getInfo()
        code = result.get("Map")
        if code is not None:
            label = _ESA_WORLDCOVER.get(int(round(float(code))),
                                        f"Class {int(round(float(code)))}")
            print(f"    [LULC] ESA WorldCover 2021 → {label} (code {int(round(float(code)))})")
            return label
    except Exception as e:
        print(f"    [LULC] ESA WorldCover error: {e} — trying MODIS")

    # ── Attempt 2: MODIS MCD12Q1 IGBP year-matched (500 m) ──────────────────
    try:
        year = max(2001, min(flood_year, 2022))
        coll = (
            ee.ImageCollection("MODIS/061/MCD12Q1")
            .filter(ee.Filter.calendarRange(year, year, "year"))
            .select("LC_Type1")
            .first()
        )
        result = coll.reduceRegion(
            reducer=ee.Reducer.mode(),
            geometry=ee.Geometry.Point([lon, lat]).buffer(500),
            scale=500,
            maxPixels=100,
        ).getInfo()
        code = result.get("LC_Type1")
        if code is not None:
            label = _MODIS_IGBP.get(int(round(float(code))),
                                    f"Class {int(round(float(code)))}")
            print(f"    [LULC] MODIS MCD12Q1 {year} → {label} (code {int(round(float(code)))})")
            return label
    except Exception as e:
        print(f"    [LULC] MODIS error: {e}")

    return ""


# ─────────────────────────────────────────────────────────────────────────────
# Per-row orchestrator
# ─────────────────────────────────────────────────────────────────────────────

def extract_row(lat: float, lon: float,
                flood_date: Optional[datetime],
                end_date:   Optional[datetime]) -> dict:
    """Run all extractions for one row. Returns a flat dict of feature values."""
    result: dict = {}
    bbox = bbox_from_point(lat, lon)

    # ── Static terrain (5 features) ─────────────────────────────────────────
    # GEE Copernicus GLO30 preferred (±1m); falls back to OpenTopoData SRTM
    result.update(extract_terrain_gee(lat, lon))
    if not _GEE_AVAILABLE:
        time.sleep(1.2)   # OpenTopoData rate limit when GEE unavailable

    # ── Static soil + LULC (GEE; no date required) ───────────────────────────
    ref_year = flood_date.year if flood_date else 2021
    result["soil_texture"] = extract_soil_texture(lat, lon)
    result["lulc_class"]   = extract_lulc(lat, lon, ref_year)

    # ── Distances (3 features) ───────────────────────────────────────────────
    result["dt_river"]    = extract_dt_river(lat, lon, bbox)
    time.sleep(2.5)
    result["dt_drainage"] = extract_dt_drainage(lat, lon, bbox)
    time.sleep(2.5)
    result["dt_roads"]    = extract_dt_roads(lat, lon, bbox)
    time.sleep(2.5)

    # ── Date-sensitive (5 features) ──────────────────────────────────────────
    if flood_date is None:
        result.update({k: NAN for k in
                       ("ndvi", "ndbi", "impervious_pct",
                        "population_per_km2", "rainfall_mm")})
        result.setdefault("soil_texture", "")
        result.setdefault("lulc_class", "")
        return result

    flood_year = flood_date.year
    eff_end    = end_date if end_date else flood_date

    result.update(extract_ndvi_ndbi(lat, lon, flood_date))
    result["impervious_pct"]     = extract_impervious(lat, lon, bbox, flood_year)
    result["population_per_km2"] = extract_population(lat, lon, flood_year)
    # CHIRPS (5.5 km) preferred over ERA5 (9 km) for rainfall
    result["rainfall_mm"]        = extract_rainfall_chirps(lat, lon, flood_date, eff_end)

    return result


# ─────────────────────────────────────────────────────────────────────────────
# Core processor — usable from CLI and Streamlit
# ─────────────────────────────────────────────────────────────────────────────

def process_dataframe(
    df: pd.DataFrame,
    log_fn=print,
    progress_fn=None,
) -> tuple:
    """
    Extract all 15 geospatial features for every row in *df*.

    Parameters
    ----------
    df          : input DataFrame (modified in-place; a copy is made internally)
    log_fn      : callable(str) used for progress messages (default: print)
    progress_fn : optional callable(current_row, total_rows) for progress bars

    Returns
    -------
    (df_out, failed_rows)
      df_out      — DataFrame with original columns + labelled feature columns
      failed_rows — list of 0-based row indices that could not be processed
    """
    df = df.copy()

    log_fn(f"Loaded {len(df)} rows × {len(df.columns)} columns")
    log_fn(f"Columns: {list(df.columns)}")

    # Resolve coordinate / location columns
    col_lat = resolve_col(df, "lat")
    col_lon = resolve_col(df, "lon")
    col_loc = resolve_col(df, "location")

    use_geocode = (not col_lat or not col_lon) and col_loc
    if use_geocode:
        log_fn(f"  lat/lon not found — will geocode from '{col_loc}'")
        if "lat_geocoded" not in df.columns:
            df.insert(0, "lat_geocoded", NAN)
            df.insert(1, "lon_geocoded", NAN)
        col_lat, col_lon = "lat_geocoded", "lon_geocoded"
    elif not col_lat or not col_lon:
        raise ValueError(
            f"Cannot find lat/lon or location columns. "
            f"Found: {list(df.columns)}"
        )

    # Resolve date columns
    col_fd    = resolve_col(df, "flood_date")
    col_ed    = resolve_col(df, "end_date")
    col_year  = resolve_col(df, "year")
    col_month = resolve_col(df, "month")
    col_day   = resolve_col(df, "day")
    use_ymd   = (col_year and col_month and col_day) and not col_fd

    if use_ymd:
        log_fn(f"  Date mode: Year/Month/Day columns ({col_year}, {col_month}, {col_day})")
    elif col_fd:
        log_fn(f"  flood_date → '{col_fd}'  |  end_date → '{col_ed or 'NOT FOUND'}'")
    else:
        log_fn("  WARNING: No date columns found — date-sensitive features will be NaN")

    _STR_COLS = {"soil_texture", "lulc_class"}
    for fc in FEATURE_COLS:
        df[fc] = "" if fc in _STR_COLS else NAN

    failed_rows: list = []
    total = len(df)

    for idx, row in df.iterrows():
        row_num = idx + 1
        log_fn(f"\n{'═'*60}")
        log_fn(f"  Row {row_num} / {total}")

        # ── Resolve coordinates ──────────────────────────────────────────────
        try:
            raw_lat = row[col_lat] if col_lat else None
            raw_lon = row[col_lon] if col_lon else None
            has_coords = (
                raw_lat is not None and raw_lon is not None
                and str(raw_lat).strip() not in ("", "nan", "NaN", "None")
                and str(raw_lon).strip() not in ("", "nan", "NaN", "None")
            )
            if has_coords:
                lat = float(raw_lat)
                lon = float(raw_lon)
            elif col_loc:
                area_name = str(row[col_loc]).strip()
                log_fn(f"  No coordinates — geocoding '{area_name}' …")
                lat, lon = geocode_location(area_name)
                if lat is None or lon is None:
                    failed_rows.append(idx)
                    if progress_fn:
                        progress_fn(row_num, total)
                    continue
                df.at[idx, col_lat] = lat
                df.at[idx, col_lon] = lon
            else:
                raise ValueError("No lat/lon and no location column to geocode from")
        except (ValueError, TypeError) as e:
            log_fn(f"  SKIP: {e}")
            failed_rows.append(idx)
            if progress_fn:
                progress_fn(row_num, total)
            continue

        # ── Resolve dates ────────────────────────────────────────────────────
        if use_ymd:
            flood_date, end_date = parse_ymd_range(
                row[col_year], row[col_month], row[col_day]
            )
        else:
            flood_date = parse_date(row.get(col_fd)) if col_fd else None
            end_date   = parse_date(row.get(col_ed)) if col_ed else None

        log_fn(
            f"  lat={lat:.5f}  lon={lon:.5f}"
            + (f"  flood={flood_date.date()}" if flood_date else "")
            + (f"  end={end_date.date()}"     if end_date   else "")
        )

        # ── Extract features ─────────────────────────────────────────────────
        try:
            features = extract_row(lat, lon, flood_date, end_date)
            for fc in FEATURE_COLS:
                df.at[idx, fc] = features.get(fc, NAN)
        except Exception as e:
            log_fn(f"  ERROR on row {row_num}: {e}")
            failed_rows.append(idx)

        if progress_fn:
            progress_fn(row_num, total)

    df_out = df.rename(columns=FEATURE_COL_LABELS)
    log_fn(f"\n{'═'*60}")
    log_fn(f"  Complete. {total - len(failed_rows)}/{total} rows extracted.")
    if failed_rows:
        log_fn(f"  Failed rows (0-indexed): {failed_rows}")

    return df_out, failed_rows


# ─────────────────────────────────────────────────────────────────────────────
# CLI entry-point
# ─────────────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        print("Usage: python3 extract_geodata.py <input.xlsx>")
        sys.exit(1)

    input_path = sys.argv[1]
    if not os.path.exists(input_path):
        print(f"ERROR: File not found — {input_path}")
        sys.exit(1)

    print(f"\nReading {input_path} …")
    df = pd.read_excel(input_path)

    df_out, failed_rows = process_dataframe(df)

    stem = os.path.splitext(input_path)[0]
    output_path = f"{stem}_extracted.xlsx"
    df_out.to_excel(output_path, index=False)
    print(f"  Saved → {output_path}")
    if failed_rows:
        print(f"  Failed rows (0-indexed): {failed_rows}")
    print(f"{'═'*60}\n")


if __name__ == "__main__":
    main()






