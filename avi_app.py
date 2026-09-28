"""
map_editor.py — Interactive GIS Timber Damage Assessment (prototype)
=====================================================================

Replaces the QGIS step: upload a footprint, view it on satellite imagery,
split it into tree-stand chunks, click a chunk to fill in stand attributes,
and export a fully attributed shapefile + running volume totals.

Run standalone to test:
    pip install streamlit streamlit-folium folium geopandas shapely pandas openpyxl
    streamlit run map_editor.py

Integration with avi_app.py:
    This module is self-contained on purpose. Once you're happy with it,
    the cleanest integration is to make this the main page and import your
    fill_template() Word-export function for the final report step.

Expects (optional, same as your current app):
    BOREAL_TDA.xlsx / FOOTHILLS_TDA.xlsx in the same directory.
    If missing, the app still works — volumes just show as 0 with a warning.
"""

import hashlib
import io
import os
import datetime
import math
import re
import tempfile
import zipfile
from pathlib import Path

import folium
import geopandas as gpd
import pandas as pd
import streamlit as st
from folium.plugins import Draw, MeasureControl
from shapely.geometry import LineString, Point, shape
from shapely.ops import split as shapely_split

from streamlit_folium import st_folium

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

# ---------------------------------------------------------------------------
# Constants (same species logic as avi_app.py)
# ---------------------------------------------------------------------------

SPECIES_NAMES = {
    "Sw": "White spruce", "Sb": "Black spruce", "P": "Pine",
    "Fb": "Balsam fir", "Fd": "Douglas fir", "Lt": "Larch",
    "Aw": "Aspen", "Pb": "Balsam poplar", "Bw": "White birch",
}
SPECIES_CHOICES = [f"{c} ({SPECIES_NAMES[c]})" for c in sorted(SPECIES_NAMES)]
CONIFERS = {"Sw", "Sb", "P", "Fb", "Fd", "Lt"}
DECIDUOUS = {"Aw", "Pb", "Bw"}

EQUAL_AREA_EPSG = 3347   # Canada LCC — area calcs
DISPLAY_EPSG = 4326      # folium works in lon/lat
MIN_CHUNK_HA = 0.0005    # discard split slivers below this

ATTR_DEFAULTS = {
    "crown_density": 70, "avg_height": 0,
    "dom_sp": "Sw", "dom_pct": 70,
    "sec_sp": "", "sec_pct": 30,
    "region": "Boreal", "region_raw": "", "filled": False,
    "avi": "", "c_vol": 0.0, "d_vol": 0.0, "c_load": 0.0, "d_load": 0.0,
    "c_vol_ha": 0.0, "d_vol_ha": 0.0, "p3_code": "", "notes": "", "no_merch": False,
}


# ---------------------------------------------------------------------------
# TDA table + volume calculation (pure-function rewrite of your global version)
# ---------------------------------------------------------------------------

@st.cache_data
def load_tda(region: str):
    path = Path(__file__).resolve().parent / f"{region.upper()}_TDA.xlsx"
    if not path.exists():
        return None
    return pd.read_excel(path)


def _density_class(d):
    return "AB" if 6 <= d <= 50 else "CD"


def _height_bin(h):
    if h <= 4:
        return "0-4"
    if h <= 8:
        return "5-8"
    if h <= 10:
        return "9-10"
    if h <= 25:
        return str(h)
    if h <= 28:
        return "26-28"
    return "29+"


def _structure_group(dom_sp, dom_pct, sec_sp, sec_pct):
    t_dec = (dom_pct if dom_sp in DECIDUOUS else 0) + (sec_pct if sec_sp in DECIDUOUS else 0)
    t_con = (dom_pct if dom_sp in CONIFERS else 0) + (sec_pct if sec_sp in CONIFERS else 0)
    if t_dec >= 70:
        return "D"
    if t_con >= 70:
        if dom_sp == "Sw":
            return "C-Sw"
        if dom_sp == "P":
            return "C-P"
        if dom_sp == "Sb":
            return "C-Sb"
        return "C-Sw"  # Fb/Fd/Lt fall back to conifer-dominant table
    if t_con > 30 and t_dec < 70:
        return "MX-P" if dom_sp == "P" else "MX-Sx"
    return None


def build_avi_code(crown_density, avg_height, dom_sp, dom_pct, sec_sp, sec_pct, is_merch=True):
    code = "m" if is_merch else ""
    if 6 <= crown_density <= 30:
        code += "A"
    elif 31 <= crown_density <= 50:
        code += "B"
    elif 51 <= crown_density <= 70:
        code += "C"
    elif 71 <= crown_density <= 100:
        code += "D"
    code += str(avg_height) + dom_sp + str(dom_pct // 10)
    if dom_pct < 100 and sec_sp:
        code += sec_sp + str(sec_pct // 10)
    return code


def calc_volumes(crown_density, avg_height, dom_sp, dom_pct, sec_sp, sec_pct, area_ha, region):
    """Pure function version of calculate_avi_and_volumes(). Returns a dict."""
    out = {
        "avi": build_avi_code(crown_density, avg_height, dom_sp, dom_pct, sec_sp, sec_pct),
        "c_vol": 0.0, "d_vol": 0.0, "c_load": 0.0, "d_load": 0.0,
        "c_vol_ha": None, "d_vol_ha": 0.0, "group": None, "tda_total": 0,
        "warning": "",
    }
    df = load_tda(region)
    if df is None:
        out["warning"] = f"{region.upper()}_TDA.xlsx not found — volumes set to 0."
        return out
    try:
        key = f"{_height_bin(avg_height)} ({_density_class(crown_density)})"
        row = df[df["Height_and_Density"].str.strip() == key]
        group = _structure_group(dom_sp, dom_pct, sec_sp, sec_pct)
        valid = {"D", "MX-P", "MX-Sx", "C-Sw", "C-P", "C-Sb"}
        col = f"Total ({group})" if group in valid else "Total (D)"
        total = row[col].values[0] if (not row.empty and col in df.columns) else 0
        out["group"], out["tda_total"] = group, total

        if dom_pct == 100:
            c_vol_ha = total if dom_sp in CONIFERS else None
            d_vol_ha = total if dom_sp in DECIDUOUS else 0
        else:
            c_pct = (dom_pct if dom_sp in CONIFERS else 0) + (sec_pct if sec_sp in CONIFERS else 0)
            d_pct = (dom_pct if dom_sp in DECIDUOUS else 0) + (sec_pct if sec_sp in DECIDUOUS else 0)
            c_vol_ha = round((c_pct / 100) * total, 1) if c_pct > 0 else None
            d_vol_ha = round((d_pct / 100) * total, 1) if d_pct > 0 else 0

        out["c_vol_ha"], out["d_vol_ha"] = c_vol_ha, d_vol_ha
        out["c_vol"] = round(c_vol_ha * area_ha, 5) if c_vol_ha is not None else 0.0
        out["d_vol"] = round(d_vol_ha * area_ha, 5) if d_vol_ha is not None else 0.0
        out["c_load"] = round(out["c_vol"] / 40, 5)
        out["d_load"] = round(out["d_vol"] / 40, 5)
    except Exception as e:  # bad TDA table structure etc.
        out["warning"] = f"TDA table error: {e}"
    return out


# ---------------------------------------------------------------------------
# Geometry helpers
# ---------------------------------------------------------------------------

def _clean(gdf):
    gdf = gdf[gdf.geometry.notna()]
    gdf = gdf[~gdf.geometry.is_empty].copy()
    if gdf.empty:
        return gdf
    try:
        gdf["geometry"] = gdf.geometry.make_valid()
    except Exception:
        gdf["geometry"] = gdf.geometry.buffer(0)
    return gdf[~gdf.geometry.is_empty]


def _union(gs):
    try:
        return gs.union_all()
    except AttributeError:
        return gs.unary_union


def geom_area_ha(geom, crs=DISPLAY_EPSG):
    """Equal-area hectares for a single geometry."""
    s = gpd.GeoSeries([geom], crs=crs).to_crs(epsg=EQUAL_AREA_EPSG)
    return round(float(s.area.iloc[0]) / 10000, 4)


def load_footprint_from_zip(uploaded_file):
    """Extract a zipped shapefile, dissolve to one footprint, return in EPSG:4326."""
    tmp = Path(tempfile.mkdtemp(prefix="fp_"))
    zpath = tmp / uploaded_file.name
    zpath.write_bytes(uploaded_file.getbuffer())
    if not zipfile.is_zipfile(zpath):
        raise ValueError("Uploaded file is not a readable zip.")
    with zipfile.ZipFile(zpath) as z:
        z.extractall(tmp)
    shps = sorted(tmp.rglob("*.shp"))
    if not shps:
        raise ValueError("No .shp found inside the zip.")
    gdf = gpd.read_file(shps[0])
    if gdf.empty:
        raise ValueError("Shapefile is empty.")
    if gdf.crs is None:
        raise ValueError("Shapefile has no CRS (.prj missing).")
    gdf = _clean(gdf.explode(ignore_index=True))
    if gdf.empty:
        raise ValueError("No valid polygon geometry after cleaning.")
    footprint = _union(gdf.geometry)
    # Snap-merge: buffer out/in by 0.5 m (in equal-area CRS) so adjacent
    # polygons with tiny gaps dissolve into ONE clean shape — no zigzag
    # internal lines left over from the original feature boundaries.
    eq = gpd.GeoSeries([footprint], crs=gdf.crs).to_crs(epsg=EQUAL_AREA_EPSG)
    merged = eq.buffer(0.5).buffer(-0.5)
    merged = merged.make_valid() if hasattr(merged, "make_valid") else merged.buffer(0)
    return gpd.GeoDataFrame(geometry=merged, crs=f"EPSG:{EQUAL_AREA_EPSG}").to_crs(epsg=DISPLAY_EPSG)


def new_chunks_gdf(geoms):
    """Build a fresh chunks GeoDataFrame (partition of the footprint)."""
    rows = []
    for i, g in enumerate(geoms, start=1):
        row = {"chunk_id": i, "area_ha": geom_area_ha(g), **ATTR_DEFAULTS}
        rows.append(row)
    gdf = gpd.GeoDataFrame(rows, geometry=list(geoms), crs=DISPLAY_EPSG)
    return gdf


def _renumber(gdf):
    gdf = gdf.reset_index(drop=True)
    gdf["chunk_id"] = range(1, len(gdf) + 1)
    return gdf


def apply_cut_line(chunks, line):
    """Split every chunk crossed by the line. Attributes are inherited by pieces."""
    new_rows, new_geoms = [], []
    for _, row in chunks.iterrows():
        geom = row.geometry
        if not line.intersects(geom):
            new_rows.append(row.drop(labels="geometry").to_dict())
            new_geoms.append(geom)
            continue
        try:
            pieces = [g for g in shapely_split(geom, line).geoms
                      if g.geom_type in ("Polygon", "MultiPolygon")]
        except Exception:
            pieces = [geom]
        if len(pieces) <= 1:
            new_rows.append(row.drop(labels="geometry").to_dict())
            new_geoms.append(geom)
            continue
        for p in pieces:
            ha = geom_area_ha(p)
            if ha < MIN_CHUNK_HA:
                continue
            d = row.drop(labels="geometry").to_dict()
            d["area_ha"] = ha
            d["filled"] = False  # geometry changed → needs review
            new_rows.append(d)
            new_geoms.append(p)
    return _renumber(gpd.GeoDataFrame(new_rows, geometry=new_geoms, crs=DISPLAY_EPSG))


def apply_drawn_polygon(chunks, poly):
    """Carve a drawn polygon out of existing chunks: each intersected chunk
    becomes (chunk ∩ poly) + (chunk − poly). Keeps the partition intact."""
    new_rows, new_geoms = [], []
    for _, row in chunks.iterrows():
        geom = row.geometry
        inter = geom.intersection(poly)
        if inter.is_empty or geom_area_ha(inter) < MIN_CHUNK_HA:
            new_rows.append(row.drop(labels="geometry").to_dict())
            new_geoms.append(geom)
            continue
        remainder = geom.difference(poly)
        for part in (inter, remainder):
            if part.is_empty:
                continue
            ha = geom_area_ha(part)
            if ha < MIN_CHUNK_HA:
                continue
            d = row.drop(labels="geometry").to_dict()
            d["area_ha"] = ha
            d["filled"] = False
            new_rows.append(d)
            new_geoms.append(part)
    return _renumber(gpd.GeoDataFrame(new_rows, geometry=new_geoms, crs=DISPLAY_EPSG))


# ---------------------------------------------------------------------------
# ATS + Natural Regions layers (same files as avi_app.py)
#   ATS_QRT.zip  (contains .gpkg or .shp)  — app folder or ATS/ subfolder
#   Regions/     (Natural Regions shapefile)
# ---------------------------------------------------------------------------

APP_DIR = Path(__file__).resolve().parent


@st.cache_resource(show_spinner="Loading ATS layer (first run only, big file)…")
def load_ats_layer():
    candidates = [APP_DIR / "ATS_QRT.zip", APP_DIR / "ATS" / "ATS_QRT.zip"]
    zpath = next((c for c in candidates if c.exists()), None)
    if zpath is None or not zipfile.is_zipfile(zpath):
        return None
    ext = Path(tempfile.mkdtemp(prefix="ats_"))
    with zipfile.ZipFile(zpath) as z:
        z.extractall(ext)
    files = sorted(ext.rglob("*.gpkg")) or sorted(ext.rglob("*.shp"))
    if not files:
        return None
    try:
        return gpd.read_file(files[0])
    except Exception:
        return None


@st.cache_resource(show_spinner=False)
def load_regions_layer():
    folder = APP_DIR / "Regions"
    if not folder.exists():
        return None
    shps = sorted(folder.glob("*.shp"))
    if not shps:
        return None
    try:
        return gpd.read_file(shps[0])
    except Exception:
        return None


def _find_field(gdf, names):
    low = {c.lower(): c for c in gdf.columns}
    for n in names:
        if n.lower() in low:
            return low[n.lower()]
    return None


def _num(v, w):
    if v is None or str(v).strip() == "" or (isinstance(v, float) and pd.isna(v)):
        return ""
    m = re.search(r"\d+", str(v))
    return m.group(0).zfill(w) if m else str(v).strip()


def _ats_label(row, f):
    sec = _num(row.get(f["sec"], ""), 2) if f["sec"] else ""
    twp = _num(row.get(f["twp"], ""), 3) if f["twp"] else ""
    rge = _num(row.get(f["rge"], ""), 2) if f["rge"] else ""
    mer = str(row.get(f["m"], "") or "").strip().upper().replace(" ", "") if f["m"] else ""
    if mer and not mer.startswith("W"):
        mm = re.search(r"\d+", mer)
        mer = f"W{mm.group(0)}" if mm else mer
    qs = str(row.get(f["qs"], "") or "").strip().upper() if f["qs"] else ""
    if qs in {"NAN", "NONE", "NULL", "0", "-"}:
        qs = ""
    if sec and twp and rge and mer:
        return f"{qs + '-' if qs else ''}{sec}-{twp}-{rge}-{mer}"
    if f["label"] and str(row.get(f["label"], "") or "").strip():
        return str(row.get(f["label"])).strip()
    return ""


def ats_query(geom_4326, ats):
    """Returns (sorted list of ATS labels intersected, display subset in 4326)."""
    if ats is None or ats.crs is None:
        return [], None
    g = gpd.GeoSeries([geom_4326], crs=DISPLAY_EPSG).to_crs(ats.crs).iloc[0]
    try:
        idx = ats.sindex.query(g, predicate="intersects")
        cand = ats.iloc[idx].copy()
    except Exception:
        cand = ats[ats.geometry.intersects(g)].copy()
    if cand.empty:
        return [], None
    f = {
        "qs": _find_field(cand, ["QS", "QTR", "QUARTER", "QUARTERSEC"]),
        "sec": _find_field(cand, ["SEC", "SECTION"]),
        "twp": _find_field(cand, ["TWP", "TOWNSHIP"]),
        "rge": _find_field(cand, ["RGE", "RANGE"]),
        "m": _find_field(cand, ["M", "MER", "MERIDIAN"]),
        "label": _find_field(cand, ["Label", "LABEL", "ATS", "ATS_LABEL"]),
    }
    labels, keep = [], []
    for i, row in cand.iterrows():
        try:
            if not row.geometry.intersects(g):
                continue
        except Exception:
            continue
        lab = _ats_label(row, f)
        if lab:
            labels.append(lab)
            keep.append(i)
    if not labels:
        return [], None
    subset = cand.loc[keep].copy()
    subset["ats_label"] = labels
    subset = subset[["ats_label", "geometry"]].to_crs(epsg=DISPLAY_EPSG)
    return sorted(set(labels)), subset


def detect_region(geom_4326, regions):
    """Dominant natural region for ONE geometry (largest overlap wins when it
    spans more than one). Returns (tda_region '' if outside Boreal/Foothills,
    raw NRNAME)."""
    if regions is None or regions.crs is None:
        return "", ""
    nr = _find_field(regions, ["NRNAME", "Natural_Region", "NAT_REGION", "REGION"])
    if nr is None:
        return "", ""
    g = gpd.GeoSeries([geom_4326], crs=DISPLAY_EPSG).to_crs(regions.crs).iloc[0]
    cand = regions[regions.geometry.intersects(g)]
    if cand.empty:
        return "", ""
    g_eq = gpd.GeoSeries([geom_4326], crs=DISPLAY_EPSG).to_crs(epsg=EQUAL_AREA_EPSG).iloc[0]
    cand_eq = cand.to_crs(epsg=EQUAL_AREA_EPSG)
    areas = {}
    for i, row in cand_eq.iterrows():
        try:
            a = row.geometry.intersection(g_eq).area
        except Exception:
            a = 0
        name = str(cand.loc[i, nr]).strip()
        areas[name] = areas.get(name, 0) + a
    raw = max(areas, key=areas.get)
    low = raw.lower()
    tda = "Boreal" if "boreal" in low else ("Foothills" if "foothill" in low else "")
    return tda, raw


def assign_regions(chunks, regions=None):
    """Auto-detect the natural region of every stand; dropdown defaults follow."""
    if regions is None:
        regions = load_regions_layer()
    if regions is None or chunks is None or chunks.empty:
        return chunks
    for i, row in chunks.iterrows():
        tda, raw = detect_region(row.geometry, regions)
        chunks.loc[i, "region_raw"] = raw
        if tda:
            chunks.loc[i, "region"] = tda
    return chunks


# ---------------------------------------------------------------------------
# P3 map viewer (reads local "P3 Maps" folder; filenames must contain the
# 6-digit P3 code MRRTTT, e.g. 511048)
# ---------------------------------------------------------------------------

P3_FOLDER = APP_DIR / "P3 Maps"
P3_EXTS = {".pdf", ".png", ".jpg", ".jpeg", ".tif", ".tiff", ".gif", ".bmp"}


def ats_to_p3(value):
    """NE-20-48-11-W5 -> 511048. Also accepts P3:511048* or bare 511048."""
    value = str(value).strip()
    m = re.search(r"(?:P3\s*:\s*)?(\d{6})\*?", value, re.IGNORECASE)
    if m and "-" not in value:
        return m.group(1)
    mm = re.match(r"^(?:[A-Za-z]{2}-)?\d{1,2}-\d{1,3}-\d{1,2}-[Ww](\d)$",
                  value, re.IGNORECASE)
    if mm:
        parts = value.replace(" ", "-").split("-")
        rge = re.search(r"\d+", parts[-2]).group(0).zfill(2)
        twp = re.search(r"\d+", parts[-3]).group(0).zfill(3)
        return f"{mm.group(1)}{rge}{twp}"
    return None


def find_p3_files(code):
    if not P3_FOLDER.exists() or not code:
        return []
    return sorted(p for p in P3_FOLDER.rglob("*")
                  if p.is_file() and code in p.name and p.suffix.lower() in P3_EXTS)


def _grid_xy(sec):
    """Expected grid coords of a section centre: x in columns from west,
    y in rows from top (each 0-6)."""
    row = (sec - 1) // 6
    pos = (sec - 1) % 6
    col = pos if row % 2 == 1 else 5 - pos
    return col + 0.5, (5 - row) + 0.5


OCR_RENDER_WIDTH = 4800
P3_CALIB_DIRNAME = ".calibration"


@st.cache_resource(show_spinner=False)
def _ocr_engine():
    from rapidocr_onnxruntime import RapidOCR
    return RapidOCR()


def _robust_fit(cands):
    """RANSAC-style consensus fit of grid -> pixel from OCR'd section numbers.
    Bad reads get voted out because they don't agree with the grid geometry.
    cands: list of (sec, px, py, h). Returns (ax, bx, ay, by, found) or None."""
    import itertools
    import numpy as np
    if not cands:
        return None
    hmax = max(c[3] for c in cands)
    big = [c for c in cands if c[3] >= 0.55 * hmax]
    dd = []
    for c in big:  # dedupe near-identical detections (tile overlap)
        if not any(c[0] == d[0] and abs(c[1] - d[1]) < 30 and abs(c[2] - d[2]) < 30
                   for d in dd):
            dd.append(c)
    pts = [(v, _grid_xy(v), (x, y)) for v, x, y, h in dd]

    def ransac(axis):
        best = None
        for a, b in itertools.combinations(pts, 2):
            g1, g2 = a[1][axis], b[1][axis]
            if abs(g1 - g2) < 1:
                continue
            s = (a[2][axis] - b[2][axis]) / (g1 - g2)
            if s <= 0:
                continue
            off = a[2][axis] - s * g1
            inl = {p[0] for p in pts
                   if abs(s * p[1][axis] + off - p[2][axis]) < 0.2 * s}
            if best is None or len(inl) > len(best[1]):
                best = (s, inl)
        return best

    fx, fy = ransac(0), ransac(1)
    if not fx or not fy:
        return None
    common = [p for p in pts if p[0] in fx[1] and p[0] in fy[1]]
    if len({p[0] for p in common}) < 3:
        return None
    G = np.array([p[1] for p in common])
    P = np.array([p[2] for p in common])

    def lsq(g, p):
        A = np.vstack([g, np.ones_like(g)]).T
        (a, b), *_ = np.linalg.lstsq(A, p, rcond=None)
        return float(a), float(b)

    ax, bx = lsq(G[:, 0], P[:, 0])
    ay, by = lsq(G[:, 1], P[:, 1])
    if not (0.7 < ax / ay < 1.4):  # sections are square on the sheet
        return None
    labels = {p[0]: (float(p[2][0]), float(p[2][1])) for p in common}
    return ax, bx, ay, by, sorted(labels), labels


def _calibrate_bytes(data, suffix, page=0):
    """Hi-res tiled OCR of the sheet -> grid calibration dict, or None."""
    import numpy as np
    from PIL import Image
    Image.MAX_IMAGE_PIXELS = None
    if suffix == ".pdf":
        import fitz
        doc = fitz.open(stream=data, filetype="pdf")
        pg = doc[min(page, len(doc) - 1)]
        scale = OCR_RENDER_WIDTH / max(pg.rect.width, 1)
        pix = pg.get_pixmap(matrix=fitz.Matrix(scale, scale))
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    else:
        img = Image.open(io.BytesIO(data)).convert("RGB")
        if img.width < OCR_RENDER_WIDTH * 0.6:  # upscale small scans for OCR
            r = OCR_RENDER_WIDTH / img.width
            img = img.resize((int(img.width * r), int(img.height * r)))

    arr = np.array(img)
    H, W = arr.shape[:2]
    TILE, OV = 1300, 180
    tiles = [(ty, tx) for ty in range(0, H, TILE - OV)
             for tx in range(0, W, TILE - OV)]
    import random
    random.Random(0).shuffle(tiles)  # stratified-ish: cover the whole sheet early

    ocr = _ocr_engine()
    cands, fit, n = [], None, 0
    for ty, tx in tiles:
        tile = arr[ty:ty + TILE, tx:tx + TILE]
        if tile.size == 0 or tile.std() < 8:  # blank
            continue
        try:
            res, _ = ocr(tile)
        except Exception:
            continue
        n += 1
        for item in (res or []):
            box, text = item[0], str(item[1]).strip()
            if not re.fullmatch(r"\d{1,2}", text):
                continue
            v = int(text)
            if not (1 <= v <= 36):
                continue
            bx_ = [p[0] for p in box]
            by_ = [p[1] for p in box]
            cands.append((v, tx + (min(bx_) + max(bx_)) / 2,
                          ty + (min(by_) + max(by_)) / 2,
                          max(by_) - min(by_)))
        # stop as soon as the grid is confidently locked in
        if n % 3 == 0 and cands:
            fit = _robust_fit(cands)
            if fit and len(fit[4]) >= 4:
                break
    if fit is None:
        fit = _robust_fit(cands)
    if fit is None:
        return None
    ax, bx, ay, by, found, labels = fit
    offs = []
    for v, (px, py) in labels.items():
        gx, gy = _grid_xy(v)
        offs.append((px - (ax * gx + bx), py - (ay * gy + by)))
    n = max(len(offs), 1)
    label_offset = (sum(o[0] for o in offs) / n, sum(o[1] for o in offs) / n)
    return {"transform": (ax, bx, ay, by), "found": found,
            "labels": {str(k): v for k, v in labels.items()},
            "label_offset": label_offset, "img_size": (W, H)}


def _grid_lines_1d(profile, n=7):
    """Find n equally-spaced strong lines in an ink projection (6x6 township
    grid = 7 lines). Returns (first, last) positions or None."""
    try:
        from scipy.signal import find_peaks
    except ImportError:
        return None
    import numpy as np
    if profile.max() <= 0:
        return None
    pk, _ = find_peaks(profile, prominence=profile.max() * 0.12,
                       distance=max(1, int(len(profile) * 0.04)))
    if len(pk) < n:
        return None
    pk = np.array(pk)
    best = None
    for i in range(len(pk)):
        for j in range(i + 1, len(pk)):
            d = (pk[j] - pk[i]) / (n - 1)
            if d < len(profile) * 0.07:
                continue
            exp = pk[i] + d * np.arange(n)
            ok, err = True, 0
            for e in exp:
                near = pk[np.argmin(np.abs(pk - e))]
                if abs(near - e) > d * 0.25:
                    ok = False
                    break
                err += abs(near - e)
            if ok and (best is None or err < best[2]):
                best = (pk[i], pk[j], err)
    return None if best is None else (float(best[0]), float(best[1]))


@st.cache_data(show_spinner="Locating township grid on the sheet…")
def detect_grid_box(data, suffix, page=0):
    """Find the 6x6 township grid box directly from the map lines (no OCR).
    The regular grid is a unique signature no legend panel has, robust to panel
    position. Returns a calibration dict like detect_section_grid, or None."""
    import numpy as np
    from PIL import Image
    Image.MAX_IMAGE_PIXELS = None
    try:
        if suffix == ".pdf":
            import fitz
            pg = fitz.open(stream=data, filetype="pdf")[0]
            scale = 2400 / max(pg.rect.width, 1)
            pix = pg.get_pixmap(matrix=fitz.Matrix(scale, scale))
            img = Image.frombytes("RGB", (pix.width, pix.height),
                                  pix.samples).convert("L")
        else:
            img = Image.open(io.BytesIO(data)).convert("L")
            if img.width < 1600:
                r = 2400 / img.width
                img = img.resize((int(img.width * r), int(img.height * r)))
        W, H = img.size
        a = np.array(img).astype("float32")
        ink = (a < 128).astype("float32")
        gx = _grid_lines_1d(ink.sum(axis=0))
        gy = _grid_lines_1d(ink.sum(axis=1))
        if not gx or not gy:
            return None
        left, right = gx
        top, bottom = gy
        sx = (right - left) / 6.0
        sy = (bottom - top) / 6.0
        if sx <= 0 or sy <= 0 or not (0.85 < sx / sy < 1.18):
            return None  # not a square township grid
        # pixel = ax*gx + bx with gx from _grid_xy (0.5..5.5), gx=0 at west edge
        return {"transform": (sx, left, sy, top), "found": list(range(1, 37)),
                "labels": {}, "label_offset": (0.0, 0.0),
                "img_size": (W, H), "method": "gridbox"}
    except Exception:
        return None


@st.cache_data(show_spinner="Reading section numbers off the sheet (first time per map, ~1-2 min)…")
def detect_section_grid(path_str, page=0):
    """Calibration with a persistent sidecar file, so the OCR cost is paid once
    per map ever — restarts included. Sidecars live in 'P3 Maps/.calibration'."""
    import json
    path = Path(path_str)
    try:
        stat = path.stat()
        sig = [int(stat.st_size), int(stat.st_mtime), 3]  # v3: grid-box method
    except Exception:
        sig = None
    calib_dir = P3_FOLDER / P3_CALIB_DIRNAME
    calib_file = calib_dir / f"{path.name}.p{page}.json"
    if sig and calib_file.exists():
        try:
            j = json.loads(calib_file.read_text())
            if j.get("sig") == sig:
                return j["result"]
        except Exception:
            pass
    data = path.read_bytes()
    # Primary: detect the township grid box from map lines (no OCR, robust).
    result = detect_grid_box(data, path.suffix.lower(), page)
    # Fallback: OCR the printed section numbers (slower, noisier on dense sheets).
    if result is None:
        result = _calibrate_bytes(data, path.suffix.lower(), page)
    if sig:
        try:
            calib_dir.mkdir(exist_ok=True)
            calib_file.write_text(json.dumps({"sig": sig, "result": result}))
        except Exception:
            pass  # read-only folder etc. — in-memory cache still applies
    return result


AVI_CODE_RE = re.compile(r"[A-Za-z][A-Za-z0-9]{1,7}(?:-[A-Za-z])?")


def _codeish(txt):
    """Keep AVI-style stand codes (C2ASw-U, A3A-U, B2A-H); drop the site/density
    tokens like 92-G and junk. Requires >=2 letters and a leading letter."""
    t = txt.strip()
    if not re.fullmatch(r"[A-Za-z][A-Za-z0-9\-]{1,8}", t):
        return False
    if re.fullmatch(r"\d{1,2}-[A-Za-z]", t):  # site line e.g. 92-G
        return False
    return len(re.findall(r"[A-Za-z]", t)) >= 2


def quarter_bbox_frac(sec, qs, transform, img_size, pad=0.30):
    """Fractional bbox (0-1) on the sheet for a section's quarter (or whole
    section if qs blank), from the OCR grid calibration."""
    ax, bx, ay, by = transform
    oW, oH = img_size
    gx, gy = _grid_xy(sec)
    dx = 0.25 if qs in ("NE", "SE") else (-0.25 if qs in ("NW", "SW") else 0.0)
    dy = -0.25 if qs in ("NE", "NW") else (0.25 if qs in ("SE", "SW") else 0.0)
    fx = (ax * (gx + dx) + bx) / oW
    fy = (ay * (gy + dy) + by) / oH
    q = 0.25 if qs in ("NE", "NW", "SE", "SW") else 0.5
    hx = abs(ax) / oW * (q + pad)
    hy = abs(ay) / oH * (q + pad)
    return (fx - hx, fy - hy, fx + hx, fy + hy)


def _union_bbox(boxes):
    xs0 = min(b[0] for b in boxes); ys0 = min(b[1] for b in boxes)
    xs1 = max(b[2] for b in boxes); ys1 = max(b[3] for b in boxes)
    return (max(0, xs0), max(0, ys0), min(1, xs1), min(1, ys1))


@st.cache_data(show_spinner="Reading stand codes off the P3 map…")
def detect_codes_in_bbox(data, suffix, page, bbox_frac):
    """OCR AVI-style stand codes inside a fractional bbox of the sheet.
    Returns [(code, confidence)] sorted best-first. Renders only the clipped
    region at high DPI, so it's fast and memory-light."""
    import numpy as np
    from PIL import Image
    Image.MAX_IMAGE_PIXELS = None
    fx0, fy0, fx1, fy1 = bbox_frac
    if suffix == ".pdf":
        import fitz
        doc = fitz.open(stream=data, filetype="pdf")
        pg = doc[min(page, len(doc) - 1)]
        R = pg.rect
        clip = fitz.Rect(fx0 * R.width, fy0 * R.height, fx1 * R.width, fy1 * R.height)
        scale = 9600 / max(R.width, 1)
        pix = pg.get_pixmap(matrix=fitz.Matrix(scale, scale), clip=clip)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    else:
        full = Image.open(io.BytesIO(data)).convert("RGB")
        W, H = full.size
        img = full.crop((int(fx0 * W), int(fy0 * H), int(fx1 * W), int(fy1 * H)))
        if max(img.size) < 2000:
            r = 2400 / max(img.size)
            img = img.resize((int(img.width * r), int(img.height * r)))
    arr = np.array(img)
    ocr = _ocr_engine()
    TILE, OV = 1100, 140
    found = {}
    for ty in range(0, arr.shape[0], TILE - OV):
        for tx in range(0, arr.shape[1], TILE - OV):
            t = arr[ty:ty + TILE, tx:tx + TILE]
            if t.size == 0 or t.std() < 6:
                continue
            try:
                res, _ = ocr(t)
            except Exception:
                continue
            for it in (res or []):
                txt = str(it[1]).strip()
                cf = float(it[2]) if len(it) > 2 else 0.0
                if _codeish(txt) and (txt not in found or cf > found[txt]):
                    found[txt] = round(cf, 2)
    return sorted(found.items(), key=lambda x: -x[1])


def parse_ats_section(value):
    """Pull quarter + section out of an ATS string. Returns (qs, sec) or ('', None)."""
    mm = re.match(r"^(?:([A-Za-z]{2})-)?(\d{1,2})-(\d{1,3})-(\d{1,2})-[Ww]\d$",
                  str(value).strip(), re.IGNORECASE)
    if not mm:
        return "", None
    qs = (mm.group(1) or "").upper()
    sec = int(mm.group(2))
    if not (1 <= sec <= 36):
        return "", None
    return (qs if qs in {"NE", "NW", "SE", "SW"} else ""), sec


def section_center_frac(sec, qs=""):
    """Position of a section's centre on a township sheet (fractions of the grid,
    x from west, y from top). ATS grid: 6x6 serpentine, section 1 = SE corner."""
    row = (sec - 1) // 6          # 0 = south row
    pos = (sec - 1) % 6
    col = pos if row % 2 == 1 else 5 - pos   # column from west
    x = (col + 0.5) / 6
    y = (5 - row + 0.5) / 6
    q = 0.25 / 6                  # nudge toward the requested quarter
    if qs == "NE":
        x += q; y -= q
    elif qs == "NW":
        x -= q; y -= q
    elif qs == "SE":
        x += q; y += q
    elif qs == "SW":
        x -= q; y += q
    return x, y


@st.cache_data(show_spinner="Rendering P3 map…")
def render_p3_image(data, suffix, page=0):
    """Rasterize a P3 map (PDF page or image file) to a PIL image."""
    from PIL import Image
    Image.MAX_IMAGE_PIXELS = None  # big scanned maps
    if suffix == ".pdf":
        import fitz  # PyMuPDF
        doc = fitz.open(stream=data, filetype="pdf")
        pg = doc[min(page, len(doc) - 1)]
        scale = min(3.0, 2400 / max(pg.rect.width, 1))
        pix = pg.get_pixmap(matrix=fitz.Matrix(scale, scale))
        return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    return Image.open(io.BytesIO(data)).convert("RGB")


def crop_to_section(img, sec, qs, sections_across, margins):
    """Crop the sheet to a window centred on the section. margins = (l, r, t, b)
    fractions of the sheet occupied by border/legend outside the township grid."""
    ml, mr, mt, mb = margins
    W, H = img.size
    gx0, gy0 = W * ml, H * mt
    gw, gh = W * (1 - ml - mr), H * (1 - mt - mb)
    cx, cy = section_center_frac(sec, qs)
    px, py = gx0 + cx * gw, gy0 + cy * gh
    hw = (sections_across / 6) * gw / 2
    hh = (sections_across / 6) * gh / 2
    box = (max(0, px - hw), max(0, py - hh), min(W, px + hw), min(H, py + hh))
    return img.crop(tuple(int(v) for v in box))


@st.cache_data(show_spinner="Preparing P3 map…")
def p3_image_payload(data, suffix, page=0):
    """Rasterize a P3 map and return (base64 PNG data-uri, (width, height)).
    Rendered hi-res in grayscale: sharper when zoomed AND a smaller file
    than the old lower-res colour version (these sheets are line art)."""
    import base64
    from PIL import Image
    Image.MAX_IMAGE_PIXELS = None
    if suffix == ".pdf":
        import fitz
        doc = fitz.open(stream=data, filetype="pdf")
        pg = doc[min(page, len(doc) - 1)]
        scale = 4800 / max(pg.rect.width, 1)
        pix = pg.get_pixmap(matrix=fitz.Matrix(scale, scale))
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    else:
        img = Image.open(io.BytesIO(data))
    img = img.convert("L")
    maxdim = 5200
    if max(img.size) > maxdim:
        r = maxdim / max(img.size)
        img = img.resize((int(img.width * r), int(img.height * r)))
    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=True)
    uri = "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()
    return uri, img.size


def render_p3_section(ss):
    st.markdown("### 🗺️ P3 Map")
    if not P3_FOLDER.exists():
        st.caption('Folder "P3 Maps" not found next to map_editor.py — P3 viewer off.')
        return

    with st.expander("⚙️ Pre-calibrate all P3 maps (run once — then every map is instant)"):
        st.caption("Reads the section numbers off every map in the P3 Maps folder "
                   "and saves the calibration. New maps take ~1 min each; "
                   "already-calibrated maps are skipped instantly. You can leave "
                   "this running over lunch.")
        if st.button("Start pre-calibration", key="p3_precal"):
            files = sorted(p for p in P3_FOLDER.rglob("*")
                           if p.is_file() and p.suffix.lower() in P3_EXTS
                           and P3_CALIB_DIRNAME not in p.parts)
            if not files:
                st.warning("No map files found.")
            else:
                prog = st.progress(0.0)
                status = st.empty()
                ok = 0
                for i, f in enumerate(files):
                    status.write(f"Calibrating {f.name} … ({i + 1}/{len(files)})")
                    try:
                        if detect_section_grid(str(f), 0):
                            ok += 1
                    except Exception:
                        pass
                    prog.progress((i + 1) / len(files))
                status.write(f"Done — {ok}/{len(files)} maps calibrated. "
                             f"They'll all open instantly now.")

    first_ats = ss.ats_text.split(",")[0].strip() if ss.ats_text else ""
    if "p3_query" not in ss:
        ss.p3_query = first_ats  # auto: first ATS intersected
    query = st.text_input(
        "ATS or P3 code", key="p3_query",
        placeholder="NE-20-48-11-W5 or P3:511048*",
        help="Autofilled with the first ATS the footprint intersects. Paste any "
             "other ATS from the list above to view its map.")
    if not query.strip():
        return

    code = ats_to_p3(query)
    if not code:
        st.warning("Could not read that input. Try NE-20-48-11-W5 or P3:511048*.")
        return

    matches = find_p3_files(code)
    if not matches:
        st.warning(f"No file containing P3 code {code} found in the P3 Maps folder.")
        return

    if len(matches) > 1:
        pick = st.selectbox("Multiple maps match — choose one:",
                            [p.name for p in matches], key=f"p3_pick_{code}")
        path = next(p for p in matches if p.name == pick)
    else:
        path = matches[0]

    st.caption(f"P3 {code} · {path.name}")
    data = path.read_bytes()
    suffix = path.suffix.lower()
    qs, sec = parse_ats_section(query)

    try:
        page = 0
        if suffix == ".pdf":
            import fitz
            npages = len(fitz.open(stream=data, filetype="pdf"))
            if npages > 1:
                page = st.number_input("PDF page", 1, npages, 1,
                                       key=f"p3_page_{code}") - 1
        uri, (W, H) = p3_image_payload(data, suffix, page)

        # Where is the section? Prefer OCR calibration from the big section
        # numbers printed on the sheet; fall back to blind grid math.
        calibrated = False
        found = []
        if sec:
            grid = detect_section_grid(str(path), page)
            if grid:
                ax, bx, ay, by = grid["transform"]
                oW, oH = grid["img_size"]
                labels = grid.get("labels", {})
                if str(sec) in labels:
                    # circle goes exactly on the printed section number
                    px0, py0 = labels[str(sec)]
                else:
                    # estimate where the number would be: grid position plus
                    # the average number-offset learned from detected sections
                    gx, gy = _grid_xy(sec)
                    ox, oy = grid.get("label_offset", (0, 0))
                    px0 = ax * gx + bx + ox
                    py0 = ay * gy + by + oy
                px = px0 * (W / oW)
                py = py0 * (H / oH)
                if 0 <= px <= W and 0 <= py <= H:
                    center = [H - py, px]
                    calibrated = True
                    found = grid["found"]
        if sec and calibrated:
            zoom = int(math.floor(math.log2(700 / (W / 2))))   # ~3 sections across
        elif sec:
            cx, cy = section_center_frac(sec, qs)
            center = [H - cy * H, cx * W]                      # y flips in CRS.Simple
            zoom = int(math.floor(math.log2(700 / (W / 2))))
        else:
            center = [H / 2, W / 2]
            zoom = int(math.floor(math.log2(min(700 / W, 620 / H))))

        # overlay toggle (only meaningful once the grid is calibrated)
        show_grid = False
        if calibrated:
            show_grid = st.toggle("Show section / quarter grid", value=True,
                                  key=f"p3grid_{code}")

        pm = folium.Map(location=center, zoom_start=zoom, crs="Simple",
                        tiles=None, min_zoom=-6, max_zoom=5, zoom_control=True)
        folium.raster_layers.ImageOverlay(image=uri, bounds=[[0, 0], [H, W]]).add_to(pm)

        def gpt(gx, gy):
            """grid coord -> folium CRS.Simple [lat, lng] on the payload image."""
            ax, bx, ay, by = grid["transform"]
            oW, oH = grid["img_size"]
            px = (ax * gx + bx) * (W / oW)
            py = (ay * gy + by) * (H / oH)
            return [H - py, px]

        if show_grid and calibrated:
            for gx in range(7):  # section lines
                folium.PolyLine([gpt(gx, 0), gpt(gx, 6)], color="#00B0FF",
                                weight=2, opacity=0.7).add_to(pm)
            for gy in range(7):
                folium.PolyLine([gpt(0, gy), gpt(6, gy)], color="#00B0FF",
                                weight=2, opacity=0.7).add_to(pm)
            for h in [i + 0.5 for i in range(6)]:  # quarter lines (dashed)
                folium.PolyLine([gpt(h, 0), gpt(h, 6)], color="#00B0FF",
                                weight=1, opacity=0.4, dash_array="4 6").add_to(pm)
                folium.PolyLine([gpt(0, h), gpt(6, h)], color="#00B0FF",
                                weight=1, opacity=0.4, dash_array="4 6").add_to(pm)

        if sec:
            folium.CircleMarker(center, radius=14, color="#FF3D00", weight=3,
                                fill=False,
                                tooltip=f"Section {sec}{' ' + qs if qs else ''}"
                                ).add_to(pm)
        st_folium(pm, height=620, use_container_width=True,
                  key=f"p3view_{code}_{path.name}_{page}", returned_objects=[])

        if sec and calibrated:
            st.caption(f"Centred on section {sec}{' (' + qs + ')' if qs else ''} — "
                       f"position calibrated from section numbers read off the sheet "
                       f"({', '.join(map(str, found[:8]))}{'…' if len(found) > 8 else ''}). "
                       f"Drag to pan, scroll to zoom.")
        elif sec:
            st.caption(f"Centred on section {sec}{' (' + qs + ')' if qs else ''} — "
                       f"⚠️ couldn't read section numbers off this sheet, so the red "
                       f"circle is an estimate. Drag to pan, scroll to zoom.")
        else:
            st.caption("Drag to pan, scroll to zoom.")
    except ImportError:
        st.warning("PDF viewing needs PyMuPDF — run: pip install pymupdf")
    except Exception as e:
        st.warning(f"Could not display map ({e}).")
    st.download_button("⬇️ Download this P3 map", data, file_name=path.name,
                       key=f"p3_dl_{path.name}")


# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------

def init_state():
    ss = st.session_state
    ss.setdefault("chunks", None)            # GeoDataFrame partition of footprint
    ss.setdefault("footprint", None)         # 1-row GeoDataFrame, EPSG:4326
    ss.setdefault("selected_chunk", None)    # chunk_id or None
    ss.setdefault("processed_drawings", set())
    ss.setdefault("footprint_name", "")
    ss.setdefault("ats_text", "")
    ss.setdefault("ats_subset", None)  # GeoDataFrame of intersected ATS quarters (4326)


def drawing_hash(geojson_geom):
    return hashlib.md5(str(geojson_geom).encode()).hexdigest()


# ---------------------------------------------------------------------------
# Map construction
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def load_custom_layers():
    """Optional custom basemaps from layers.json next to map_editor.py.
    Format: [{"name": "...", "url": "...", "tms": false, "max_zoom": 20,
              "default": false}, ...]
    QGIS-style {-y} in the URL is converted automatically (implies TMS)."""
    import json
    f = APP_DIR / "layers.json"
    if not f.exists():
        return []
    try:
        raw = json.loads(f.read_text())
        out = []
        for item in raw:
            url = str(item.get("url", "")).strip()
            if not url:
                continue
            kind = str(item.get("type", "xyz")).lower()
            tms = bool(item.get("tms", False))
            if "{-y}" in url:
                url = url.replace("{-y}", "{y}")
                tms = True
            out.append({"name": str(item.get("name", "Custom layer")).strip(),
                        "url": url, "tms": tms, "type": kind,
                        "layers": str(item.get("layers", "")),
                        "fmt": str(item.get("format", "image/jpeg")),
                        "version": str(item.get("version", "1.1.1")),
                        "max_zoom": int(item.get("max_zoom", 20)),
                        "attr": str(item.get("attribution",
                                             item.get("name", "Custom"))),
                        "default": bool(item.get("default", False))})
        return out
    except Exception:
        return []


def chunk_style(row, selected_id):
    if row["chunk_id"] == selected_id:
        return {"color": "#FF9800", "weight": 3, "fillColor": "#FF9800", "fillOpacity": 0.35}
    if row["filled"]:
        return {"color": "#2e7d32", "weight": 2, "fillColor": "#4CAF50", "fillOpacity": 0.25}
    return {"color": "#c62828", "weight": 2, "fillColor": "#f44336", "fillOpacity": 0.20}


def build_map(footprint, chunks, mode, selected_id, ats_subset=None, p3_overlay=None, canopy_layers=None, hover_grid=None):
    bounds = footprint.total_bounds  # minx, miny, maxx, maxy
    center = [(bounds[1] + bounds[3]) / 2, (bounds[0] + bounds[2]) / 2]
    m = folium.Map(location=center, zoom_start=15, tiles=None)

    custom = load_custom_layers()

    def _add_custom(lyr):
        if lyr["type"] == "wms":
            folium.raster_layers.WmsTileLayer(
                url=lyr["url"], layers=lyr["layers"], fmt=lyr["fmt"],
                version=lyr["version"], transparent=False,
                name=lyr["name"], attr=lyr["attr"], overlay=False,
                control=True).add_to(m)
        else:
            folium.TileLayer(tiles=lyr["url"], name=lyr["name"],
                             attr=lyr["attr"], tms=lyr["tms"],
                             max_zoom=lyr["max_zoom"]).add_to(m)

    folium.TileLayer("OpenStreetMap", name="Streets").add_to(m)
    for lyr in [c for c in custom if not c["default"]]:
        _add_custom(lyr)
    folium.TileLayer(
        tiles="https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
        attr="Esri World Imagery",
        name="Satellite",
    ).add_to(m)
    # a custom layer flagged "default": true is added last -> it becomes the
    # default view on every rebuild; otherwise Esri Satellite stays default
    for lyr in [c for c in custom if c["default"]]:
        _add_custom(lyr)

    # ATS grid (only quarters intersecting the footprint)
    if ats_subset is not None and not ats_subset.empty:
        folium.GeoJson(
            ats_subset.__geo_interface__,
            name="ATS grid",
            style_function=lambda f: {"color": "#FFD54F", "weight": 1.5,
                                      "fill": False, "dashArray": "2 6"},
            tooltip=folium.GeoJsonTooltip(fields=["ats_label"], labels=False),
        ).add_to(m)

    # Footprint outline (always visible)
    folium.GeoJson(
        footprint.__geo_interface__,
        name="Footprint",
        style_function=lambda f: {"color": "#FFFFFF", "weight": 3, "fill": False, "dashArray": "6 4"},
    ).add_to(m)

    # Chunks
    if chunks is not None and not chunks.empty:
        for _, row in chunks.iterrows():
            style = chunk_style(row, selected_id)
            tip = (f"<b>Stand {row['chunk_id']}</b><br>"
                   f"Area: {row['area_ha']:.4f} ha<br>"
                   f"AVI: {row['avi'] or '—'}<br>"
                   f"{'✅ complete' if row['filled'] else '❌ not filled'}")
            folium.GeoJson(
                gpd.GeoSeries([row.geometry], crs=DISPLAY_EPSG).__geo_interface__,
                style_function=lambda f, s=style: s,
                tooltip=tip,
            ).add_to(m)

    # Draw controls depend on mode
    if mode == "Cut lines":
        Draw(draw_options={"polyline": {"shapeOptions": {"color": "#00E5FF", "weight": 3}},
                           "polygon": False, "rectangle": False, "circle": False,
                           "marker": False, "circlemarker": False},
             edit_options={"edit": False, "remove": False}).add_to(m)

    if p3_overlay:
        # rotated/skewed sheet placed on its geographic corners
        try:
            from folium.raster_layers import ImageOverlay
            corners = p3_overlay["corners_ll"]  # [TL,TR,BR,BL] as (lon,lat)
            lats = [c[1] for c in corners]; lons = [c[0] for c in corners]
            # ImageOverlay uses a bbox; for slight rotation this is close enough
            ImageOverlay(image=p3_overlay["uri"],
                         bounds=[[min(lats), min(lons)], [max(lats), max(lons)]],
                         opacity=p3_overlay.get("opacity", 0.5),
                         name="P3 map overlay").add_to(m)
        except Exception:
            pass

    if canopy_layers:
        try:
            from folium.raster_layers import ImageOverlay
            for name, ov in canopy_layers:
                if not ov:
                    continue
                uri, ov_bounds = ov[0], ov[1]
                ImageOverlay(image=uri, bounds=ov_bounds, opacity=0.6,
                             name=name, show=False).add_to(m)
        except Exception:
            pass

    if hover_grid:
        folium.GeoJson(
            hover_grid, name="Canopy height (hover for value)", show=True,
            style_function=lambda f: {"fillColor": "#ffffff", "fillOpacity": 0.0,
                                      "color": "#66ccff", "weight": 0.4,
                                      "opacity": 0.35},
            highlight_function=lambda f: {"fillOpacity": 0.25, "weight": 1.5,
                                          "color": "#ffffff", "fillColor": "#ffcc00"},
            tooltip=folium.GeoJsonTooltip(fields=["h"], aliases=["Height:"],
                                          sticky=True)).add_to(m)

    MeasureControl(primary_length_unit="meters",
                   secondary_length_unit="kilometers",
                   primary_area_unit="hectares").add_to(m)
    folium.LayerControl().add_to(m)
    m.fit_bounds([[bounds[1], bounds[0]], [bounds[3], bounds[2]]])
    return m


# ---------------------------------------------------------------------------
# Attribute form for the selected chunk
# ---------------------------------------------------------------------------

_P3_SYMS = sorted([("Sw", "Sw"), ("Sb", "Sb"), ("Fb", "Fb"), ("Fd", "Fd"),
                   ("Lt", "Lt"), ("Pb", "Pb"), ("Po", "Pb"), ("Bw", "Bw"),
                   ("Aw", "Aw"), ("Pl", "P"), ("Pj", "P"), ("A", "Aw"),
                   ("P", "P"), ("L", "Lt"), ("S", "Sw"), ("F", "Fb")],
                  key=lambda s: -len(s[0]))
_P3_DENS = {"A": 18, "B": 40, "C": 60, "D": 85}


def parse_p3_code(code):
    """Old P3/AVI code -> {density letter, species list}. e.g. C2ASw -> Aw+Sw."""
    t = re.sub(r"-[A-Za-z]$", "", str(code).strip())
    dens = t[0].upper() if t and t[0].upper() in _P3_DENS else ""
    body = t[1:] if dens else t
    body = re.sub(r"^[0-9IOol]+", "", body)  # strip height digit (+ OCR I/O/l)
    sp, i = [], 0
    while i < len(body):
        for sym, mapped in _P3_SYMS:
            if body[i:i + len(sym)].lower() == sym.lower():
                sp.append(mapped)
                i += len(sym)
                break
        else:
            i += 1
    seen = []
    for s in sp:
        if s not in seen:
            seen.append(s)
    return {"density": dens, "species": seen}


def georeference_p3(path, ats):
    """Fit a grid-space -> lat/lon affine for a P3 sheet, using the OCR'd
    section numbers matched to their real ATS section centroids. Returns dict
    with the sheet image (data-uri), its geographic corner bounds, and a
    rotation, or None. Loosely accurate: good for visual cross-reference."""
    grid = detect_section_grid(str(path), 0)
    if not grid or ats is None:
        return None
    code = ats_to_p3(path.name)
    m = re.search(r"(\d)(\d{2})(\d{3})", str(code) if code else "")
    if not m:
        return None
    mer, rge, twp = int(m.group(1)), int(m.group(2)), int(m.group(3))
    sec_f = _find_field(ats, ["SEC", "SECTION"])
    twp_f = _find_field(ats, ["TWP", "TOWNSHIP"])
    rge_f = _find_field(ats, ["RGE", "RANGE"])
    mer_f = _find_field(ats, ["M", "MER", "MERIDIAN"])
    if not all([sec_f, twp_f, rge_f]):
        return None
    import numpy as np

    def as_int(v):
        mm = re.search(r"\d+", str(v))
        return int(mm.group(0)) if mm else None

    sub = ats[(ats[twp_f].map(as_int) == twp) & (ats[rge_f].map(as_int) == rge)]
    if mer_f is not None:
        sub = sub[sub[mer_f].map(as_int) == mer]
    if sub.empty:
        return None
    cent = sub.to_crs(epsg=4326).geometry.centroid
    sub = sub.assign(_lon=cent.x, _lat=cent.y, _sec=sub[sec_f].map(as_int))
    # average duplicate quarter rows to one centroid per section
    persec = sub.groupby("_sec")[["_lon", "_lat"]].mean()

    ax, bx, ay, by = grid["transform"]
    oW, oH = grid["img_size"]
    G, LL = [], []
    for sec in grid["found"]:
        if sec in persec.index:
            gx, gy = _grid_xy(sec)
            G.append([gx, gy])
            LL.append([persec.loc[sec, "_lon"], persec.loc[sec, "_lat"]])
    if len(G) < 2:
        return None
    G = np.array(G)
    LL = np.array(LL)

    # fit lon = a*gx + b*gy + c ; lat = d*gx + e*gy + f (full affine)
    A = np.hstack([G, np.ones((len(G), 1))])
    (clon, *_), (clat, *_) = (np.linalg.lstsq(A, LL[:, 0], rcond=None),
                              np.linalg.lstsq(A, LL[:, 1], rcond=None))
    plon = np.linalg.lstsq(A, LL[:, 0], rcond=None)[0]
    plat = np.linalg.lstsq(A, LL[:, 1], rcond=None)[0]

    def g2ll(gx, gy):
        return (float(plon[0] * gx + plon[1] * gy + plon[2]),
                float(plat[0] * gx + plat[1] * gy + plat[2]))

    # sheet image corners in grid space come from the image extent:
    # invert grid->pixel to get grid coords at the four image corners.
    # pixel_x = ax*gx+bx over full width oW -> gx spans (0-bx)/ax .. (oW-bx)/ax
    gx0, gx1 = (0 - bx) / ax, (oW - bx) / ax
    gy0, gy1 = (0 - by) / ay, (oH - by) / ay
    corners_grid = [(gx0, gy0), (gx1, gy0), (gx1, gy1), (gx0, gy1)]  # TL,TR,BR,BL
    corners_ll = [g2ll(gx, gy) for gx, gy in corners_grid]

    # lightweight overlay image (~1500px) so the opacity slider stays snappy
    import base64
    from PIL import Image
    Image.MAX_IMAGE_PIXELS = None
    # readable overlay resolution; cached so the opacity slider stays snappy
    P3_OVERLAY_DIM = 3800
    data = path.read_bytes()
    if path.suffix.lower() == ".pdf":
        import fitz
        pgd = fitz.open(stream=data, filetype="pdf")[0]
        s = P3_OVERLAY_DIM / max(pgd.rect.width, 1)
        pix = pgd.get_pixmap(matrix=fitz.Matrix(s, s))
        oimg = Image.frombytes("RGB", (pix.width, pix.height), pix.samples).convert("L")
    else:
        oimg = Image.open(io.BytesIO(data)).convert("L")
        if max(oimg.size) > P3_OVERLAY_DIM:
            rr = P3_OVERLAY_DIM / max(oimg.size)
            oimg = oimg.resize((int(oimg.width * rr), int(oimg.height * rr)))
    ob = io.BytesIO(); oimg.save(ob, "PNG", optimize=True)
    uri = "data:image/png;base64," + base64.b64encode(ob.getvalue()).decode()
    return {"uri": uri, "corners_ll": corners_ll}


@st.cache_data(show_spinner="Aligning P3 sheet to the map…")
def georeference_p3_cached(path_str, sig):
    """Cached georeference so adjusting the overlay opacity doesn't recompute."""
    return georeference_p3(Path(path_str), load_ats_layer())


def detect_stand_profile(geom_4326):
    """Detect codes in the stand's intersected quarters, then aggregate into a
    species list, a frequency ranking (for dominant/2nd), and a modal crown-
    density class. Returns dict with codes, species_ranked, density_pct."""
    codes = _detect_stand_codes(geom_4326)
    from collections import Counter
    sp_count, dens_count = Counter(), Counter()
    parses = []
    for code, conf in codes:
        p = parse_p3_code(code)
        parses.append((code, conf, p["species"]))
        for s in p["species"]:
            sp_count[s] += 1
        if p["density"]:
            dens_count[p["density"]] += 1
    ranked = [s for s, _ in sp_count.most_common()]
    dens_letter = dens_count.most_common(1)[0][0] if dens_count else ""
    return {"codes": codes, "parses": parses, "species_ranked": ranked,
            "density_letter": dens_letter,
            "density_pct": _P3_DENS.get(dens_letter, None)}


def _detect_stand_codes(geom_4326):
    """For one stand geometry: find the P3 sheet + the quarter-sections it
    intersects, then OCR AVI codes only within those intersected portions."""
    ats = load_ats_layer()
    if ats is None:
        return []
    labels, _ = ats_query(geom_4326, ats)
    if not labels:
        return []
    # group labels by the P3 sheet (township) they belong to
    by_code = {}
    for lab in labels:
        c = ats_to_p3(lab)
        if c:
            by_code.setdefault(c, []).append(lab)
    all_found = {}
    for code, labs in by_code.items():
        files = find_p3_files(code)
        if not files:
            continue
        path = files[0]
        grid = detect_section_grid(str(path), 0)
        if not grid:
            continue
        boxes = []
        for lab in labs:
            qs, sec = parse_ats_section(lab)
            if sec is None:
                continue
            boxes.append(quarter_bbox_frac(sec, qs, grid["transform"],
                                           grid["img_size"]))
        if not boxes:
            continue
        bbox = _union_bbox(boxes)
        data = path.read_bytes()
        for cd, cf in detect_codes_in_bbox(data, path.suffix.lower(), 0, bbox):
            if cd not in all_found or cf > all_found[cd]:
                all_found[cd] = cf
    return sorted(all_found.items(), key=lambda x: -x[1])


@st.cache_data(show_spinner="Reading P3 species for this stand…")
def detect_stand_profile_cached(geom_wkt):
    """Cached per stand geometry so auto-detect doesn't re-run OCR on every click."""
    from shapely import wkt as _wkt
    return detect_stand_profile(_wkt.loads(geom_wkt))


def render_chunk_form(chunks, cid):
    idx = chunks.index[chunks["chunk_id"] == cid]
    if len(idx) == 0:
        st.session_state.selected_chunk = None
        return chunks
    i = idx[0]
    row = chunks.loc[i]

    # apply any queued suggestions (species/height/density) before widgets draw
    for pk, wk in [("dom_pending", "dom"), ("sec_pending", "sec"),
                   ("ht_pending", "ht"), ("cd_pending", "cd"), ("dp_pending", "dp")]:
        key = f"{pk}_{cid}"
        if key in st.session_state:
            st.session_state[wk + f"_{cid}"] = st.session_state.pop(key)

    st.markdown(f"### ✏️ Stand {cid} — {row['area_ha']:.4f} ha")

    k = f"_{cid}"  # per-chunk widget keys so switching chunks refreshes values

    no_merch = st.checkbox("🚫 No merchantable timber present",
                           value=bool(row.get("no_merch", False)), key="nomerch" + k,
                           help="Check for non-treed / cleared portions. Skips the "
                                "timber inputs — AVI is set to N/A and volumes to 0. "
                                "Area, date, project code and status still export.")

    NO_MERCH_NOTE = "No merchantable timber present within footprint."
    if no_merch:
        notes_val = str(row.get("notes", "") or "")
        if not notes_val.strip():
            notes_val = NO_MERCH_NOTE
        notes_nm = st.text_area("Notes", value=notes_val, key="nmnotes" + k)
        st.info("AVI will export as **N/A**; C/D volumes and loads as 0. "
                "Area_Ha, Add_Date, Project_Co and Status still populate.")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("💾 Save Stand", key="savenm" + k, type="primary",
                         use_container_width=True):
                idx = chunks.index[chunks["chunk_id"] == cid]
                if len(idx):
                    ii = idx[0]
                    chunks.loc[ii, ["no_merch", "avi", "c_vol", "d_vol", "c_load",
                                    "d_load", "c_vol_ha", "d_vol_ha", "notes",
                                    "filled"]] = [True, "N/A", 0.0, 0.0, 0.0, 0.0,
                                                  0.0, 0.0, notes_nm, True]
                    st.session_state.chunks = chunks
                    st.session_state.selected_chunk = None
                    st.rerun()
        with c2:
            if st.button("Close without saving", key="cancelnm" + k,
                         use_container_width=True):
                st.session_state.selected_chunk = None
                st.rerun()
        return chunks

    crown = st.slider("Crown Density (%)", 6, 100, int(row["crown_density"]), key="cd" + k)
    height = st.slider("Average Stand Tree Height (m)", 0, 40, int(row["avg_height"]), key="ht" + k)

    rp = find_canopy_raster()
    if rp is not None:
        try:
            sig = _raster_sig(rp)
        except Exception:
            sig = (0, 0)
        ch = canopy_height_for_bounds(row.geometry.wkt, str(rp), sig)
        if ch:
            _lbl = canopy_source_label(rp).lower()
            src_name = "GEDI" if "gedi" in _lbl else (
                "ICESat-2" if "icesat" in _lbl else
                "ETH 10m" if "eth" in _lbl else "canopy")
            hint_col1, hint_col2 = st.columns([3, 1])
            with hint_col1:
                st.caption(f"🛰️ Satellite canopy-height reference ({src_name}): "
                           f"mean **{ch['mean']:.0f} m** "
                           f"(range {ch['min']:.0f}–{ch['max']:.0f} m, "
                           f"{ch['n']} px). Coarse — cross-check only.")
            with hint_col2:
                if st.button(f"Use {ch['mean']:.0f} m", key="usech" + k):
                    st.session_state["ht" + k] = int(round(ch["mean"]))
                    st.rerun()
        else:
            st.caption("🛰️ No canopy-height data covering this stand "
                       "(outside mapped forest, or footprint too small).")

    dom_default = f"{row['dom_sp']} ({SPECIES_NAMES.get(row['dom_sp'], '')})"
    dom_sel = st.selectbox("Dominant Species", SPECIES_CHOICES,
                           index=SPECIES_CHOICES.index(dom_default) if dom_default in SPECIES_CHOICES else 0,
                           key="dom" + k)
    dom_sp = dom_sel.split(" ")[0]
    dom_pct = st.slider("Dominant Cover %", 0, 100, int(row["dom_pct"]), step=10, key="dp" + k)

    sec_opts = [""] + [c for c in SPECIES_CHOICES if c.split(" ")[0] != dom_sp]
    sec_default = f"{row['sec_sp']} ({SPECIES_NAMES.get(row['sec_sp'], '')})" if row["sec_sp"] else ""
    sec_sel = st.selectbox("2nd Species", sec_opts,
                           index=sec_opts.index(sec_default) if sec_default in sec_opts else 0,
                           key="sec" + k)
    sec_sp = sec_sel.split(" ")[0] if sec_sel else ""
    sec_pct = 100 - dom_pct if sec_sp else 0
    if sec_sp:
        st.caption(f"2nd Cover: {sec_pct}% (auto = 100 − dominant)")

    region = st.selectbox("Natural Region", ["Boreal", "Foothills"],
                          index=0 if row["region"] == "Boreal" else 1, key="rg" + k)
    pend_key = f"p3c_pending_{cid}"
    if pend_key in st.session_state:
        st.session_state["p3c" + k] = st.session_state.pop(pend_key)
    p3_code = st.text_input("P3 Code", value=str(row.get("p3_code", "") or ""),
                            key="p3c" + k, placeholder="e.g. C2ASw-U")
    notes = st.text_input("Notes", value=str(row.get("notes", "") or ""),
                          key="nt" + k)

    # --- Auto-detect P3 species for THIS stand (cached per geometry) ---
    with st.container():
        prof = None
        if P3_FOLDER.exists() and load_ats_layer() is not None:
            try:
                prof = detect_stand_profile_cached(row.geometry.wkt)
            except Exception:
                prof = None
        if prof is not None:
            ranked = prof["species_ranked"]
            if not ranked:
                st.caption("🌲 No P3 species auto-detected for this stand "
                           "(footprint may be off-sheet).")
            else:
                names = ", ".join(SPECIES_NAMES.get(s, s) for s in ranked)
                st.success(f"🌲 Species present (intersected quarters): {names}")

                # canopy-model height for the predicted AVI
                rp = find_canopy_raster()
                ch = None
                if rp is not None:
                    try:
                        sig = _raster_sig(rp)
                        ch = canopy_height_for_bounds(row.geometry.wkt, str(rp), sig)
                    except Exception:
                        ch = None
                # canopy map is a 2022 snapshot -> add growth since then.
                # rough species-blended rate, capped, using today's date.
                dom0 = ranked[0]
                fast = {"Pb", "Aw", "Bw"}      # deciduous, faster
                rate = 0.5 if dom0 in fast else 0.3  # m/yr, conservative
                years = max(0, (datetime.date.today() - datetime.date(2022, 7, 1)).days / 365.25)
                growth = round(rate * years, 1)
                base_h = ch["mean"] if ch else float(row["avg_height"])
                pred_h = int(round(base_h + (growth if ch else 0)))
                pred_h = max(0, min(40, pred_h))
                dom = ranked[0]
                sec = ranked[1] if len(ranked) > 1 else ""
                dpct = 70 if sec else 100
                dens_pct = prof["density_pct"] or int(row["crown_density"])
                pred = calc_volumes(dens_pct, pred_h, dom, dpct, sec,
                                    100 - dpct if sec else 0, row["area_ha"],
                                    row["region"])
                hs = (f"{pred_h} m (canopy 2022 {base_h:.0f}m +{growth}m growth)"
                      if ch else f"{pred_h} m (current)")
                st.info(f"**Predicted AVI:** `{pred['avi']}`  \n"
                        f"dominant **{SPECIES_NAMES.get(dom, dom)}**"
                        + (f", 2nd **{SPECIES_NAMES.get(sec, sec)}**" if sec else "")
                        + f"  ·  height {hs}  ·  crown ~{dens_pct}%"
                        + (f" (class {prof['density_letter']})" if prof["density_letter"] else ""))
                if st.button("✅ Apply prediction to this stand", key="applypred" + k):
                    st.session_state[f"dom_pending_{cid}"] = f"{dom} ({SPECIES_NAMES.get(dom, dom)})"
                    st.session_state[f"sec_pending_{cid}"] = (
                        f"{sec} ({SPECIES_NAMES.get(sec, sec)})" if sec else "")
                    st.session_state[f"ht_pending_{cid}"] = pred_h
                    st.session_state[f"dp_pending_{cid}"] = dpct
                    st.session_state[f"cd_pending_{cid}"] = dens_pct
                    st.rerun()

                with st.expander("Source codes (tap to fill P3 Code)"):
                    st.caption("OCR can misread letters (Sw↔SW, S↔5) — verify.")
                    cols = st.columns(3)
                    for j, (code, conf) in enumerate(prof["codes"]):
                        if cols[j % 3].button(code, key=f"pick_{cid}_{j}",
                                              help=f"confidence {conf}"):
                            st.session_state[f"p3c_pending_{cid}"] = code
                            st.rerun()

    raw = str(row.get("region_raw", "") or "")
    if raw:
        if "boreal" in raw.lower() or "foothill" in raw.lower():
            st.caption(f"📍 Auto-detected: {raw}")
        else:
            st.caption(f"⚠️ Auto-detected: **{raw}** — no TDA table for this "
                       f"region. Closest option is selected above; adjust if needed.")

    # Live preview
    res = calc_volumes(crown, height, dom_sp, dom_pct, sec_sp, sec_pct, row["area_ha"], region)
    if res["warning"]:
        st.warning(res["warning"])
    st.markdown(
        f"**AVI:** `{res['avi']}` &nbsp;|&nbsp; "
        f"**Con:** {res['c_vol']:.3f} m³ ({res['c_load']:.3f} loads) &nbsp;|&nbsp; "
        f"**Dec:** {res['d_vol']:.3f} m³ ({res['d_load']:.3f} loads)"
    )

    c1, c2 = st.columns(2)
    with c1:
        if st.button("💾 Save Stand", key="save" + k, type="primary", use_container_width=True):
            chunks.loc[i, ["crown_density", "avg_height", "dom_sp", "dom_pct",
                           "sec_sp", "sec_pct", "region", "avi",
                           "c_vol", "d_vol", "c_load", "d_load",
                           "c_vol_ha", "d_vol_ha", "p3_code", "notes", "filled"]] = [
                crown, height, dom_sp, dom_pct, sec_sp, sec_pct, region,
                res["avi"], res["c_vol"], res["d_vol"], res["c_load"], res["d_load"],
                res["c_vol_ha"] or 0.0, res["d_vol_ha"] or 0.0, p3_code, notes, True,
            ]
            st.session_state.chunks = chunks
            st.session_state.selected_chunk = None
            st.rerun()
    with c2:
        if st.button("Close without saving", key="cancel" + k, use_container_width=True):
            st.session_state.selected_chunk = None
            st.rerun()
    return chunks


# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

def export_shapefile_zip(chunks, project_code="", add_date=None):
    """Attributed shapefile as a zip — AiM Timber layer schema, exact labels:
    AVI, C_VOL_Ha, D_Vol_Ha, C_Vol, D_Vol, C_load, D_Load, Area_Ha, Status,
    Add_Date, Project_Co, P3_Code, Notes."""
    import datetime as _dt
    if add_date is None:
        add_date = _dt.date.today()
    out = gpd.GeoDataFrame(geometry=chunks.geometry, crs=chunks.crs)
    out["AVI"] = chunks["avi"]
    out["C_VOL_Ha"] = chunks["c_vol_ha"].astype(float)
    out["D_Vol_Ha"] = chunks["d_vol_ha"].astype(float)
    out["C_Vol"] = chunks["c_vol"].astype(float)
    out["D_Vol"] = chunks["d_vol"].astype(float)
    out["C_load"] = chunks["c_load"].astype(float)
    out["D_Load"] = chunks["d_load"].astype(float)
    out["Area_Ha"] = chunks["area_ha"].astype(float)
    out["Status"] = "1"
    out["Add_Date"] = add_date.strftime("%Y-%m-%d")
    out["Project_Co"] = str(project_code).strip()
    out["P3_Code"] = chunks["p3_code"].astype(str)
    out["Notes"] = chunks["notes"].astype(str)
    tmp = Path(tempfile.mkdtemp(prefix="export_"))
    shp = tmp / "tda_stands.shp"
    out.to_file(shp)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for f in tmp.iterdir():
            z.write(f, f.name)
    buf.seek(0)
    return buf


def export_summary_csv(chunks):
    cols = ["chunk_id", "area_ha", "region", "avi", "dom_sp", "dom_pct",
            "sec_sp", "sec_pct", "crown_density", "avg_height",
            "c_vol", "c_load", "d_vol", "d_load", "filled"]
    return chunks[cols].to_csv(index=False).encode()


# ---------------------------------------------------------------------------
# Canopy height reference (Sothe et al. 2022, GEDI/ICESat-2, 250 m, 2020)
#   Download the 490 MB zip from 4TU (link in the UI), unzip, and drop the
#   two GeoTIFFs into a "Canopy Height" folder next to map_editor.py.
#   Filenames contain "GEDI" and "ICESat" respectively.
# ---------------------------------------------------------------------------

CANOPY_FOLDER = APP_DIR / "Tree_Height_2022"


def _rio_source(path_or_url):
    """Wrap remote URLs for GDAL streaming (/vsicurl/), pass local paths through."""
    s = str(path_or_url)
    if s.startswith(("http://", "https://")):
        os.environ.setdefault("GDAL_DISABLE_READDIR_ON_OPEN", "EMPTY_DIR")
        os.environ.setdefault("CPL_VSIL_CURL_ALLOWED_EXTENSIONS", ".tif,.tiff")
        os.environ.setdefault("GDAL_HTTP_MULTIRANGE", "YES")
        os.environ.setdefault("VSI_CACHE", "TRUE")
        return "/vsicurl/" + s
    return s


def _raster_sig(path_or_url):
    """Cache signature. Local: (size, mtime). Remote: hash of the URL."""
    s = str(path_or_url)
    if s.startswith(("http://", "https://")):
        return (abs(hash(s)) % (10 ** 12), 0)
    try:
        st = Path(s).stat()
        return (int(st.st_size), int(st.st_mtime))
    except Exception:
        return (0, 0)


def _canopy_url_sources():
    """URLs listed in Tree_Height_2022/sources.txt (one per line, '#' comments).
    Streamed remotely via GDAL — no download needed."""
    f = CANOPY_FOLDER / "sources.txt"
    if not f.exists():
        return []
    out = []
    for line in f.read_text().splitlines():
        u = line.strip()
        if not u or u.startswith("#"):
            continue
        name = u.rsplit("/", 1)[-1] or u
        low = name.lower()
        if "eth" in low:
            label = f"ETH 10m ({name})"
        elif "meta" in low or "chm" in low:
            label = f"Meta 1m ({name})"
        else:
            label = name
        out.append((label + " [stream]", u))
    return out


def _product_key(name):
    """Group tiles of the same product (strip the NxxWyyy tile id + suffix)."""
    n = name.lower()
    if "eth" in n or "sentinel" in n:
        return "ETH 10m"
    if "ca_canopy" in n or "nfis" in n or "ntems" in n:
        return "NFIS 30m"
    if "gedi" in n:
        return "GEDI"
    if "icesat" in n:
        return "ICESat-2"
    # strip a trailing tile code like _N51W117 so siblings group together
    return re.sub(r"[_-]?[NS]\d{2}[EW]\d{3}.*$", "", name, flags=re.IGNORECASE) or name


def list_canopy_rasters():
    """Canopy sources as (label, path_or_grouptag). Tiles of the same product
    are collapsed into ONE seamless entry (e.g. all ETH tiles -> 'ETH 10m'),
    so you never pick a tile — the app reads whichever covers the footprint."""
    if not CANOPY_FOLDER.exists():
        return _canopy_url_sources()
    tifs = sorted(p for p in CANOPY_FOLDER.rglob("*")
                  if p.suffix.lower() in (".tif", ".tiff"))
    groups = {}
    for p in tifs:
        groups.setdefault(_product_key(p.name), []).append(p)
    out = []
    for key, paths in groups.items():
        if len(paths) == 1:
            out.append((f"{key} ({paths[0].stem})", str(paths[0])))
        else:
            # seamless multi-tile product: label carries a group:// tag
            out.append((f"{key} ({len(paths)} tiles, seamless)",
                        "group://" + key))
    out.extend(_canopy_url_sources())
    return out


def eth_tiles_for_geom(geom_4326):
    """Which ETH 3-degree tiles cover a footprint. Tiles are named by their
    lower-left corner in 3-degree steps. Returns list of tile names."""
    import math as _m
    minx, miny, maxx, maxy = geom_4326.bounds
    names = set()
    lat = _m.floor(miny / 3) * 3
    while lat <= maxy:
        lon = _m.floor(minx / 3) * 3
        while lon <= maxx:
            ns = "N" if lat >= 0 else "S"
            ew = "W" if lon < 0 else "E"
            names.add(f"{ns}{abs(lat):02d}{ew}{abs(lon):03d}")
            lon += 3
        lat += 3
    return sorted(names)


@st.cache_data(show_spinner=False)
def _covering_tiles(product_key, bounds):
    """Cached: which tiles of a product cover a footprint bbox (4326 bounds).
    Opening every tile to check bounds was the slow part on each refresh."""
    import rasterio
    from rasterio.warp import transform_bounds
    from shapely.geometry import box as _box
    box = _box(*bounds)
    tifs = sorted(p for p in CANOPY_FOLDER.rglob("*")
                  if p.suffix.lower() in (".tif", ".tiff")
                  and _product_key(p.name) == product_key)
    covering = []
    for p in tifs:
        try:
            with rasterio.open(_rio_source(p)) as ds:
                w, s2, e, n = transform_bounds(ds.crs, DISPLAY_EPSG, *ds.bounds)
            if box.intersects(_box(w, s2, e, n)):
                covering.append(str(p))
        except Exception:
            continue
    return covering


@st.cache_data(show_spinner="Merging canopy tiles…")
def _merge_tiles(paths):
    """Cached: merge tiles that a footprint spans into one temp GeoTIFF."""
    import rasterio
    from rasterio.merge import merge as rio_merge
    srcs = [rasterio.open(_rio_source(p)) for p in paths]
    mosaic, mtr = rio_merge(srcs)
    meta = srcs[0].meta.copy()
    meta.update(height=mosaic.shape[1], width=mosaic.shape[2], transform=mtr)
    outp = Path(tempfile.mkdtemp(prefix="canopy_merge_")) / "merged.tif"
    with rasterio.open(outp, "w", **meta) as dst:
        dst.write(mosaic)
    for sds in srcs:
        sds.close()
    return str(outp)


def _resolve_source_for_geom(source, geom_4326):
    """A canopy 'source' may be a real path/URL or a 'group://Product' tag.
    For a group, return the tile that covers geom (or a merged temp GeoTIFF if
    geom spans several). Tile-coverage lookup is cached for speed. Returns a
    path/URL string or None."""
    s = str(source)
    if not s.startswith("group://"):
        return s
    key = s[len("group://"):]
    bounds = tuple(round(v, 5) for v in geom_4326.bounds)
    covering = _covering_tiles(key, bounds)
    if not covering:
        return None
    if len(covering) == 1:
        return covering[0]
    try:
        return _merge_tiles(tuple(covering))
    except Exception:
        return covering[0]


def find_canopy_raster(preferred=None):
    """Default canopy source (path string or group:// tag) for readouts.
    NFIS/GEDI preferred, else first available. Resolve per-geometry with
    _resolve_source_for_geom() before opening."""
    rasters = list_canopy_rasters()
    if not rasters:
        return None
    if preferred:
        for _, srcv in rasters:
            if str(srcv) == str(preferred):
                return srcv
    for key in ("nfis", "ca_canopy", "gedi", "icesat"):
        for lb, srcv in rasters:
            if key in lb.lower() or key in str(srcv).lower():
                return srcv
    return rasters[0][1]


def canopy_source_label(source):
    """Short label for a source string, for captions."""
    s = str(source)
    if s.startswith("group://"):
        return s[len("group://"):]
    return Path(s).stem


@st.cache_data(show_spinner=False)
def canopy_height_for_bounds(_geom_wkt, raster_path_str, raster_sig):
    """Mean/min/max canopy height (m) over one stand geometry.
    raster_sig (size, mtime) is part of the cache key so edits invalidate it.
    Returns dict or None. Geometry passed as WKT so it's hashable for caching."""
    import numpy as np
    import rasterio
    from rasterio.mask import mask as rio_mask
    from shapely import wkt as _wkt
    from shapely.geometry import mapping
    geom = _wkt.loads(_geom_wkt)
    raster_path_str = _resolve_source_for_geom(raster_path_str, geom)
    if not raster_path_str:
        return None
    try:
        with rasterio.open(_rio_source(raster_path_str)) as ds:
            g = gpd.GeoSeries([geom], crs=DISPLAY_EPSG).to_crs(ds.crs).iloc[0]
            out, _ = rio_mask(ds, [mapping(g)], crop=True, filled=True)
            band = out[0].astype("float32")
            nod = ds.nodata
        vals = band.flatten()
        vals = vals[np.isfinite(vals) & (vals > -1e30)]
        # exclude 0 (model non-forest) and implausible values from the height mean
        vals = vals[(vals >= 0.5) & (vals < 80)]
        if vals.size == 0:
            return None
        return {"mean": float(np.mean(vals)), "min": float(np.min(vals)),
                "max": float(np.max(vals)), "n": int(vals.size)}
    except Exception:
        return None


@st.cache_data(show_spinner="Rendering canopy-height overlay…")
def canopy_overlay(footprint_wkt, raster_path_str, raster_sig, hmax=25.0):
    """Clip the canopy raster to the footprint's buffered box, reproject to
    EPSG:4326 (so it aligns with the basemap), and colour by height.
    Orange = 0/non-forest (model, not ground truth), red = low <2m,
    green ramp = taller, grey = no data. Returns (uri, bounds, coverage)."""
    import numpy as np
    import rasterio
    from rasterio.mask import mask as rio_mask
    from rasterio.warp import calculate_default_transform, reproject, Resampling
    from rasterio.transform import array_bounds
    from shapely import wkt as _wkt
    from shapely.geometry import box as _box, mapping
    from PIL import Image
    import base64
    geom = _wkt.loads(footprint_wkt)
    raster_path_str = _resolve_source_for_geom(raster_path_str, geom)
    if not raster_path_str:
        return None
    try:
        with rasterio.open(_rio_source(raster_path_str)) as ds:
            g = gpd.GeoSeries([geom], crs=DISPLAY_EPSG).to_crs(ds.crs).iloc[0]
            minx, miny, maxx, maxy = g.bounds
            bufm = max(300.0, 0.5 * max(maxx - minx, maxy - miny))
            win = _box(minx - bufm, miny - bufm, maxx + bufm, maxy + bufm)
            arr, tr = rio_mask(ds, [mapping(win)], crop=True, filled=True)
            src = arr[0].astype("float32")
            nod = ds.nodata if ds.nodata is not None else -3.4e38
            scrs = ds.crs
            H0, W0 = src.shape
            l, b, r, t = array_bounds(H0, W0, tr)        # reproject window to 4326 so the image lines up with satellite tiles.
        # cap the target grid so fine-res rasters (10m/1m) don't blow up memory
        # or exceed Streamlit's message size limit.
        dtr, dw, dh = calculate_default_transform(scrs, "EPSG:4326", W0, H0,
                                                  left=l, bottom=b, right=r, top=t)
        OV_MAXDIM = 1600
        if max(dw, dh) > OV_MAXDIM:
            sf = OV_MAXDIM / max(dw, dh)
            dw2, dh2 = max(1, int(dw * sf)), max(1, int(dh * sf))
            from rasterio.transform import Affine
            dtr = dtr * Affine.scale(dw / dw2, dh / dh2)
            dw, dh = dw2, dh2
        band = np.full((dh, dw), nod, dtype="float32")
        reproject(src, band, src_transform=tr, src_crs=scrs,
                  dst_transform=dtr, dst_crs="EPSG:4326",
                  src_nodata=nod, dst_nodata=nod, resampling=Resampling.nearest)
        w2, s2, e2, n2 = array_bounds(dh, dw, dtr)

        valid = np.isfinite(band) & (band > -1e30) & (band != nod)
        h = np.clip(np.where(valid, band, 0.0), 0, hmax)
        H, W = band.shape
        rgba = np.zeros((H, W, 4), dtype="uint8")
        nonforest = valid & (band < 0.5)              # 0 = model non-forest
        low = valid & (band >= 0.5) & (band < 2)      # genuine very low canopy
        tree = valid & (band >= 2)
        rgba[nonforest] = [255, 150, 0, 150]          # orange
        rgba[low] = [220, 40, 40, 170]                # red
        frac = np.clip((h - 2) / max(hmax - 2, 1), 0, 1)
        rgba[..., 0] = np.where(tree, (200 * (1 - frac)).astype("uint8"), rgba[..., 0])
        rgba[..., 1] = np.where(tree, (120 + 100 * frac).astype("uint8"), rgba[..., 1])
        rgba[..., 2] = np.where(tree, (60 * (1 - frac)).astype("uint8"), rgba[..., 2])
        rgba[..., 3] = np.where(tree, 170, rgba[..., 3])
        rgba[~valid] = [130, 130, 130, 80]            # grey = no data
        cover = float(valid.mean()) if valid.size else 0.0
        img = Image.fromarray(rgba, "RGBA")
        if max(img.size) < 600:
            rr = 600 / max(img.size)
            img = img.resize((int(W * rr), int(H * rr)), Image.NEAREST)
        buf = __import__("io").BytesIO()
        img.save(buf, "PNG")
        uri = "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()
        return uri, [[s2, w2], [n2, e2]], cover
    except Exception:
        return None


@st.cache_data(show_spinner=False)
def canopy_hover_grid(footprint_wkt, raster_path_str, raster_sig, max_cells=700):
    """Downsampled vector grid of canopy heights over the footprint's buffered
    box, for hover tooltips. Returns a GeoJSON FeatureCollection (props: h)."""
    import math as _m
    import numpy as np
    import rasterio
    from rasterio.mask import mask as rio_mask
    from shapely import wkt as _wkt
    from shapely.geometry import box as _box, mapping
    geom = _wkt.loads(footprint_wkt)
    raster_path_str = _resolve_source_for_geom(raster_path_str, geom)
    if not raster_path_str:
        return None
    try:
        with rasterio.open(_rio_source(raster_path_str)) as ds:
            g = gpd.GeoSeries([geom], crs=DISPLAY_EPSG).to_crs(ds.crs).iloc[0]
            minx, miny, maxx, maxy = g.bounds
            bufm = max(300.0, 0.5 * max(maxx - minx, maxy - miny))
            win = _box(minx - bufm, miny - bufm, maxx + bufm, maxy + bufm)
            arr, tr = rio_mask(ds, [mapping(win)], crop=True, filled=True)
            band = arr[0].astype("float32")
            nod = ds.nodata
            rcrs = ds.crs
        H, W = band.shape
        valid = np.isfinite(band)
        if nod is not None:
            valid &= (band != nod)
        blk = max(1, int(_m.ceil(_m.sqrt(H * W / max(max_cells, 1)))))
        geoms, hs = [], []
        for r0 in range(0, H, blk):
            for c0 in range(0, W, blk):
                sub = band[r0:r0 + blk, c0:c0 + blk]
                vsub = valid[r0:r0 + blk, c0:c0 + blk]
                if not vsub.any():
                    continue
                mh = float(sub[vsub].mean())
                x_left, y_top = tr * (c0, r0)
                x_right, y_bot = tr * (min(c0 + blk, W), min(r0 + blk, H))
                geoms.append(_box(min(x_left, x_right), min(y_top, y_bot),
                                  max(x_left, x_right), max(y_top, y_bot)))
                hs.append(round(mh, 1))
        if not geoms:
            return None
        gdf = gpd.GeoDataFrame({"h": hs}, geometry=geoms, crs=rcrs).to_crs(epsg=DISPLAY_EPSG)
        gdf["h"] = gdf["h"].map(lambda v: f"{v:.0f} m")
        return gdf.__geo_interface__
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Salvage report (port of avi_app.py fill_template — same Word layout)
# ---------------------------------------------------------------------------

VEG_TYPES = [
    "Native grassland", "Tame pasture", "Cropland", "Sparsely or non-vegetated",
    "Cutblock - planted", "Natural regeneration >2m", "Treed wetland",
    "Shrubby wetland", "Grass or grass-like wetland", "Native aspen parkland",
    "Other (specify)",
]
DEFAULT_WAIVER_JUSTIFICATION = ("Timber salvage is not considered economically viable, "
                                "given that the estimated volume is below 0.5 truckloads.")


def stand_percentages(done):
    """Con/dec + species split percentages from saved stands (same math as avi_app)."""
    def s(col_sp, col_pct, group):
        return sum(r[col_pct] for _, r in done.iterrows() if r[col_sp] in group)

    raw_con = s("dom_sp", "dom_pct", CONIFERS) + s("sec_sp", "sec_pct", CONIFERS)
    raw_dec = s("dom_sp", "dom_pct", DECIDUOUS) + s("sec_sp", "sec_pct", DECIDUOUS)
    pct_con = round(raw_con / (raw_con + raw_dec) * 100, 0) if (raw_con + raw_dec) > 0 else 0

    spruce = s("dom_sp", "dom_pct", {"Sw", "Sb"}) + s("sec_sp", "sec_pct", {"Sw", "Sb"})
    pine = s("dom_sp", "dom_pct", {"P"}) + s("sec_sp", "sec_pct", {"P"})
    if raw_con > 0:
        spruce_pct = int(round(spruce / raw_con * 100, 0))
        pine_pct = int(round(pine / raw_con * 100, 0))
        other_con_pct = int(round(100 - spruce_pct - pine_pct, 0))
    else:
        spruce_pct = pine_pct = other_con_pct = 0

    aspen = s("dom_sp", "dom_pct", {"Aw"}) + s("sec_sp", "sec_pct", {"Aw"})
    if raw_dec > 0:
        aspen_pct = int(round(aspen / raw_dec * 100, 0))
        other_dec_pct = int(round(100 - aspen_pct, 0))
    else:
        aspen_pct = other_dec_pct = 0

    return pct_con, spruce_pct, pine_pct, other_con_pct, aspen_pct, other_dec_pct


def _para(doc, before=0, after=0, align=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    if align is not None:
        p.alignment = align
    return p


def _run(p, text, size=10, bold=True, underline=False):
    r = p.add_run(text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.underline = underline
    return r


def build_salvage_docx(done, disposition, legal_loc, vegetation, other_specify,
                       disposition_fma, no_disposition_fma, ctlr_list,
                       salvage_waiver, justification):
    """Returns docx bytes. `done` = filled stands GeoDataFrame."""
    pct_con, spruce_pct, pine_pct, other_con_pct, aspen_pct, other_dec_pct = stand_percentages(done)

    total_c_vol = math.ceil(done["c_vol"].sum() * 10) / 10
    total_c_load = math.ceil(done["c_load"].sum() * 10) / 10
    total_d_vol = math.ceil(done["d_vol"].sum() * 10) / 10
    total_d_load = math.ceil(done["d_load"].sum() * 10) / 10

    def con_class_box(label):
        if label == "D" and pct_con < 30:
            return "☒"
        if label == "C" and pct_con > 70:
            return "☒"
        if label == "CD" and 50 <= pct_con <= 70:
            return "☒"
        if label == "DC" and 30 <= pct_con < 50:
            return "☒"
        return "☐"

    def box(l):
        return "☒" if l in vegetation else "☐"

    doc = Document()

    p = _para(doc, after=0, align=1)
    _run(p, "Vegetation and Timber Salvage Information", size=11, underline=True)

    p = _para(doc, after=0, align=1)
    _run(p, "Disposition: ")
    _run(p, disposition, bold=False)

    p = _para(doc, after=0, align=1)
    _run(p, "Legal Land Location: ")
    _run(p, legal_loc, bold=False)

    # Horizontal line
    p = doc.add_paragraph()
    p_pr = p._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "24")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "000000")
    bdr.append(bottom)
    p_pr.append(bdr)

    p = _para(doc, before=0, after=6)
    _run(p, "Vegetation and Timber Cover", size=12)

    p = _para(doc)
    _run(p, "Vegetation (check all that apply)")

    rows = [
        ("Native grassland", "Treed wetland"),
        ("Tame pasture", "Shrubby wetland"),
        ("Cropland", "Grass or grass-like wetland"),
        ("Sparsely or non-vegetated", "Native aspen parkland"),
        ("Cutblock - planted", "Other (specify)"),
    ]
    for left, right in rows:
        p = _para(doc)
        left_indent = ""  # all lefts in the pairs above are non-indented in original
        right_indent = (
            "\t\t" if left in ["Tame pasture", "Cropland"]
            else "\t" if left not in ["Sparsely or non-vegetated", "Tame pasture", "Cropland"]
            else ""
        )
        if right == "Treed wetland":
            _run(p, f"{left_indent}{box(left)} {left}{right_indent}\t{box(right)} {right}\t\t")
            _run(p, "Deciduous-dominant Forest:", underline=True)
        elif right == "Shrubby wetland":
            _run(p, f"{left_indent}{box(left)} {left}{right_indent}\t{box(right)} {right}"
                    f"\t\t{con_class_box('D')} D less than 30% coniferous")
        elif right == "Grass or grass-like wetland":
            _run(p, f"{left_indent}{box(left)} {left}{right_indent}\t{box(right)} {right}\t")
            _run(p, "Coniferous-dominant Forest:", underline=True)
        elif right == "Native aspen parkland":
            _run(p, f"{left_indent}{box(left)} {left}{right_indent}\t{box(right)} {right}"
                    f"\t\t{con_class_box('C')} C More than 70% coniferous")
        elif right == "Other (specify)":
            _run(p, f"{left_indent}{box(left)} {left}{right_indent}\t{box(right)} {right}")
            if "Other (specify)" in vegetation and other_specify:
                _run(p, ": ")
                _run(p, other_specify, bold=False, underline=True)
            _run(p, "\t\t")
            _run(p, "Mixedwood Forest:", underline=True)

    p = _para(doc)
    _run(p, f"{box('Natural regeneration >2m')} Natural regeneration >2m"
            f"\t\t\t\t\t{con_class_box('CD')} CD 70% to 50% coniferous")
    p = _para(doc)
    _run(p, f"\t\t\t\t\t\t\t\t{con_class_box('DC')} DC 50% to 30% coniferous")

    p = _para(doc, before=6)
    _run(p, "Timber Salvage:", underline=True)

    # Merchantable timber (always Yes in this workflow, same as avi_app)
    p = _para(doc, before=6)
    _run(p, "1.\tMerchantable timber present?   ☒ Yes    ☐ No")
    p = _para(doc)
    _run(p, "\tProvide a volume inventory as follows:")

    p = _para(doc)
    _run(p, "\tConiferous approx. volume: ")
    _run(p, f"{total_c_vol:.1f}", bold=False, underline=True)
    _run(p, " m³", bold=False, underline=True)
    _run(p, "  or  ")
    _run(p, f"{total_c_load:.1f}", bold=False, underline=True)
    _run(p, " loads", bold=False, underline=True)

    p = _para(doc)
    _run(p, "\tSpruce ")
    _run(p, f"{spruce_pct}%", bold=False, underline=True)
    _run(p, "    Pine ")
    _run(p, f"{pine_pct}%", bold=False, underline=True)
    _run(p, "    Other ")
    _run(p, f"{other_con_pct}%", bold=False, underline=True)

    p = _para(doc)
    _run(p, "\tDeciduous approx. volume: ")
    _run(p, f"{total_d_vol:.1f}", bold=False, underline=True)
    _run(p, " m³", bold=False, underline=True)
    _run(p, "  or  ")
    _run(p, f"{total_d_load:.1f}", bold=False, underline=True)
    _run(p, " loads", bold=False, underline=True)

    p = _para(doc)
    _run(p, "\tAspen ")
    _run(p, f"{aspen_pct}%", bold=False, underline=True)
    _run(p, "    Other ")
    _run(p, f"{other_dec_pct}%", bold=False, underline=True)

    p = _para(doc, before=6)
    _run(p, "2.\tSpecify the timber disposition or FMA(s) shown on LSAS:")
    p = _para(doc)
    _run(p, f"\t{'☒' if no_disposition_fma else '☐'} No disposition (Contact SRD field office)")
    p = _para(doc)
    _run(p, "\tDisposition number & Holder name of FMA: ")
    _run(p, disposition_fma, bold=False, underline=True)

    for ctlr in ctlr_list:
        if str(ctlr.get("type", "")).strip() or str(ctlr.get("number_holder", "")).strip():
            p = _para(doc)
            _run(p, f"\tDisposition number & Holder name of {ctlr['type']}: ")
            _run(p, ctlr["number_holder"], bold=False, underline=True)

    p = _para(doc, before=6)
    _run(p, "3.\tUtilization Standards:")
    p = _para(doc)
    _run(p, "\tConiferous ")
    _run(p, "15", bold=False, underline=True)
    _run(p, " cm stump diameter to a ")
    _run(p, "11", bold=False, underline=True)
    _run(p, " cm top diameter.")
    p = _para(doc)
    _run(p, "\tDeciduous ")
    _run(p, "15", bold=False, underline=True)
    _run(p, " cm stump diameter to a ")
    _run(p, "10", bold=False, underline=True)
    _run(p, " cm top diameter.")

    box_yes = "☒" if salvage_waiver == "Yes" else "☐"
    box_no = "☒" if salvage_waiver == "No" else "☐"
    p = _para(doc, before=6)
    _run(p, f"4.\tTimber salvage waiver requested?   {box_yes} Yes   {box_no} No")
    p = _para(doc)
    _run(p, "\tIf ‘Yes’, provide justification: ")
    if salvage_waiver == "Yes":
        _run(p, justification, bold=False, underline=True)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


def render_report_section(ss):
    st.divider()
    st.subheader("📄 Salvage Report")

    done = ss.chunks[ss.chunks["filled"]]
    if done.empty:
        st.info("Save at least one stand before generating the report.")
        return

    unfilled = len(ss.chunks) - len(done)
    if unfilled:
        st.warning(f"{unfilled} stand(s) not filled in yet — the report only "
                   f"includes the {len(done)} completed stand(s).")

    total_c_load = done["c_load"].sum()
    total_d_load = done["d_load"].sum()

    r1, r2 = st.columns(2)
    with r1:
        disposition = st.text_input(
            "Disposition", key="rep_disposition",
            help="Reference such as (RTF). OneStop applications include type and "
                 "number, e.g. RTF2525.")
        if "rep_legal" not in ss:
            ss.rep_legal = ss.ats_text  # autofilled from the ATS lookup
        legal_loc = st.text_input(
            "Legal Land Location", key="rep_legal",
            help="Autofilled from the ATS quarters intersected by the footprint — edit as needed.")
        vegetation = st.multiselect("Vegetation (check all that apply):",
                                    VEG_TYPES, key="rep_veg")
        other_specify = ""
        if "Other (specify)" in vegetation:
            other_specify = st.text_input("Other (specify):", key="rep_other")
    with r2:
        disposition_fma = st.text_input(
            "Disposition # of FMA & Holder Name:", key="rep_fma",
            help="Sketch Plan, PLSR (best source), EDP, Abadata, FMA/FMU maps, or "
                 "OneStop. If no FMA, contact the SRD field office.")
        no_disposition_fma = st.checkbox("None", key="rep_nofma")

        st.write("Coniferous/Deciduous Dispositions (Type–Number–Holder):")
        if "ctlr_list" not in ss:
            ss.ctlr_list = [{"type": "", "number_holder": ""}]
        for i in range(len(ss.ctlr_list)):
            c1, c2 = st.columns([1, 2])
            with c1:
                ss.ctlr_list[i]["type"] = st.text_input(
                    f"Type {i+1}", ss.ctlr_list[i]["type"], key=f"rep_ctlr_t{i}",
                    help="e.g. CTL, DTL, CTLR, CTLC, CTLD")
            with c2:
                ss.ctlr_list[i]["number_holder"] = st.text_input(
                    f"Number & Holder {i+1}", ss.ctlr_list[i]["number_holder"],
                    key=f"rep_ctlr_n{i}")
        if st.button("Add Another Disposition", key="rep_addctlr"):
            ss.ctlr_list.append({"type": "", "number_holder": ""})
            st.rerun()

    salvage_waiver = st.radio("Timber Salvage Waiver Requested?", ["Yes", "No"],
                              index=1, horizontal=True, key="rep_waiver",
                              help="Use when timber is uneconomic to salvage, e.g. "
                                   "less than 0.5 truckloads. Waiver rules vary by "
                                   "region and FMA.")
    st.caption(f"Total Coniferous Load: {total_c_load:.5f} · "
               f"Total Deciduous Load: {total_d_load:.5f}")

    justification = ""
    if salvage_waiver == "Yes":
        if not str(ss.get("rep_just", "")).strip():
            ss.rep_just = DEFAULT_WAIVER_JUSTIFICATION
        justification = st.text_area("Provide justification:", key="rep_just")

    if st.button("Done (Generate Report)", type="primary", key="rep_generate",
                 help="Builds the Timber form. Provide to AIM Lands staff for "
                      "submission to the FMA."):
        data = build_salvage_docx(done, disposition, legal_loc, vegetation,
                                  other_specify, disposition_fma, no_disposition_fma,
                                  ss.ctlr_list, salvage_waiver, justification)
        fname = f"Timber_Damage_Assessment_{disposition.strip() or 'Report'}.docx"
        st.success("Report generated!")
        st.download_button("📥 Download report", data, file_name=fname,
                           mime="application/vnd.openxmlformats-officedocument."
                                "wordprocessingml.document",
                           key="rep_download")


# ---------------------------------------------------------------------------
# Main app
# ---------------------------------------------------------------------------

def main():
    st.set_page_config(layout="wide", page_title="TDA Map Editor")
    init_state()
    ss = st.session_state

    st.header("🌲 TDA Map Editor — split, click, assess 🌲")

    # ---- Upload ----
    if ss.footprint is None:
        st.info("Upload a zipped shapefile of the project footprint to begin. "
                "The zip must include .shp, .shx, .dbf and .prj.")
        up = st.file_uploader("Footprint zip", type=["zip"])
        if up:
            try:
                fp = load_footprint_from_zip(up)
                ss.footprint = fp
                ss.footprint_name = Path(up.name).stem
                chunks = new_chunks_gdf([fp.geometry.iloc[0]])
                ss.chunks = assign_regions(chunks)
                labels, subset = ats_query(fp.geometry.iloc[0], load_ats_layer())
                ss.ats_text = ", ".join(labels)
                ss.ats_subset = subset
                st.rerun()
            except Exception as e:
                st.error(f"Could not load footprint: {e}")
        return

    fp_ha = geom_area_ha(ss.footprint.geometry.iloc[0])
    st.caption(f"**{ss.footprint_name}** · footprint {fp_ha:.4f} ha · "
               f"{len(ss.chunks)} stand(s) · "
               f"{int(ss.chunks['filled'].sum())} complete")
    if ss.ats_text:
        st.markdown("**Legal land locations (ATS intersected):**")
        st.code(ss.ats_text, language=None)
    elif load_ats_layer() is None:
        st.caption("ATS layer not found (ATS_QRT.zip) — legal locations unavailable.")
    if load_regions_layer() is None:
        st.caption("Regions layer not found (Regions/ folder) — natural region auto-detect off.")

    # ---- Mode + actions ----
    top1, topPC, top2, top3 = st.columns([2, 1, 1, 1])
    with top1:
        mode = st.radio("Map mode", ["Select & edit", "Cut lines"],
                        horizontal=True,
                        help="Select & edit: click a stand to fill in its attributes. "
                             "Cut lines: draw a line fully across the footprint to slice it.")
    with topPC:
        st.text_input("Project Code", key="project_code",
                      help="Written to every stand as Project_Co on export.")
    with top2:
        if st.button("↩️ Merge all back to one", use_container_width=True):
            ss.chunks = assign_regions(new_chunks_gdf([ss.footprint.geometry.iloc[0]]))
            ss.selected_chunk = None
            ss.processed_drawings = set()
            st.rerun()
    with top3:
        if st.button("🗑️ Start over (new upload)", use_container_width=True):
            for key in ["footprint", "chunks", "selected_chunk", "footprint_name", "ats_subset"]:
                ss[key] = None
            ss.ats_text = ""
            ss.processed_drawings = set()
            st.rerun()

    # ---- Layout: map left, form right ----
    map_col, form_col = st.columns([3, 2])

    with map_col:
        p3_overlay = None
        with st.expander("🗺️ Overlay P3 map on satellite (loose cross-reference)"):
            st.caption("Stretches the P3 sheet onto the map using the section-grid "
                       "calibration. Old scans have paper warp, so treat alignment "
                       "as approximate (±tens of metres).")
            if st.checkbox("Show P3 overlay", key="p3_overlay_on"):
                op = st.slider("Overlay opacity", 0.0, 1.0, 0.5, 0.05,
                               key="p3_overlay_op")
                try:
                    first = ss.ats_text.split(",")[0].strip() if ss.ats_text else ""
                    ocode = ats_to_p3(first) if first else None
                    files = find_p3_files(ocode) if ocode else []
                    if files:
                        ov = georeference_p3_cached(str(files[0]),
                                                    _raster_sig(files[0]))
                        if ov:
                            ov["opacity"] = op
                            p3_overlay = ov
                        else:
                            st.caption("Couldn't georeference this sheet "
                                       "(needs calibration + ATS layer).")
                    else:
                        st.caption("No P3 sheet found for the footprint's ATS.")
                except Exception as e:
                    st.caption(f"Overlay unavailable: {e}")
        # Canopy-height overlays are heavy, so only build them when asked —
        # keeps cutting/editing snappy. When on, they appear in the layer control.
        canopy_layers = []
        hover_grid = None
        show_canopy = st.checkbox(
            "🌲 Show canopy-height layers", value=False, key="show_canopy",
            help="Adds NFIS/ETH height overlays + a hover-to-read-height grid to "
                 "the map's layer control. Leave off while cutting for speed.")
        if show_canopy:
            _fp_wkt = ss.footprint.geometry.iloc[0].wkt
            _missing = []
            for _lab, _srcv in list_canopy_rasters():
                try:
                    _sigc = _raster_sig(_srcv)
                    _ov = canopy_overlay(_fp_wkt, str(_srcv), _sigc)
                    if _ov:
                        canopy_layers.append((_lab, _ov))
                        if hover_grid is None:
                            hover_grid = canopy_hover_grid(_fp_wkt, str(_srcv), _sigc)
                    elif str(_srcv).startswith("group://") or "ETH" in _lab:
                        _missing.append(_lab)
                except Exception:
                    pass
            if canopy_layers:
                st.caption("Hover any square to read its height. Colour layers "
                           "(🟧 non-forest 0m · 🟥 <2m · 🟩 taller) are in the "
                           "map's layer control, top-right.")
            if _missing:
                _needed = eth_tiles_for_geom(ss.footprint.geometry.iloc[0])
                _tnames = ", ".join(
                    f"ETH_GlobalCanopyHeight_10m_2020_{t}_Map.tif" for t in _needed)
                st.caption(f"No local tile covers this footprint for: "
                           f"{', '.join(_missing)}. Add to Tree_Height_2022: {_tnames}")
        m = build_map(ss.footprint, ss.chunks, mode, ss.selected_chunk,
                      ss.ats_subset, p3_overlay, canopy_layers, hover_grid)
        out = st_folium(
            m, height=620, use_container_width=True, key="tda_map",
            returned_objects=["last_active_drawing", "last_object_clicked", "last_clicked"],
        )

        # -- handle a new drawing (split) --
        drawing = (out or {}).get("last_active_drawing")
        if drawing and mode == "Cut lines":
            h = drawing_hash(drawing.get("geometry"))
            if h not in ss.processed_drawings:
                ss.processed_drawings.add(h)
                geom = shape(drawing["geometry"])
                if isinstance(geom, LineString):
                    before = len(ss.chunks)
                    ss.chunks = apply_cut_line(ss.chunks, geom)
                    if len(ss.chunks) == before:
                        st.toast("Cut line must fully cross a stand — nothing was split.", icon="⚠️")
                    else:
                        ss.chunks = assign_regions(ss.chunks)
                    st.rerun()

        # -- handle a click (select) --
        # Clicks on a chunk polygon arrive as last_object_clicked;
        # clicks on bare map arrive as last_clicked. Check both.
        clicked = (out or {}).get("last_object_clicked") or (out or {}).get("last_clicked")
        if clicked and mode == "Select & edit":
            pt = Point(clicked["lng"], clicked["lat"])
            hits = ss.chunks[ss.chunks.geometry.contains(pt)]
            if not hits.empty:
                cid = int(hits.iloc[0]["chunk_id"])
                if cid != ss.selected_chunk:
                    ss.selected_chunk = cid
                    st.rerun()

        # -- P3 map viewer (a second map) is heavy; render only when asked --
        if st.checkbox("🗺️ Show P3 map viewer", value=False, key="show_p3_viewer",
                       help="Opens the P3 sheet viewer below. Leave off while "
                            "cutting for speed."):
            render_p3_section(ss)

    with form_col:
        if ss.selected_chunk is not None and mode == "Select & edit":
            ss.chunks = render_chunk_form(ss.chunks, ss.selected_chunk)
        else:
            st.markdown("### Stands")
            view = ss.chunks[["chunk_id", "area_ha", "avi", "c_load", "d_load", "filled"]].rename(
                columns={"chunk_id": "Stand", "area_ha": "ha", "avi": "AVI",
                         "c_load": "C loads", "d_load": "D loads", "filled": "Done"})
            st.dataframe(view, use_container_width=True, hide_index=True, height=260)
            if mode != "Select & edit":
                st.info("Switch to **Select & edit** and click a stand on the map to fill it in.")

        # ---- Totals ----
        done = ss.chunks[ss.chunks["filled"]]
        tc_vol, td_vol = done["c_vol"].sum(), done["d_vol"].sum()
        tc_load, td_load = done["c_load"].sum(), done["d_load"].sum()
        st.markdown("### Running totals (completed stands)")
        t1, t2 = st.columns(2)
        t1.metric("Coniferous", f"{tc_vol:.2f} m³", f"{math.ceil(tc_load*10)/10:.1f} loads",
                  delta_color="off")
        t2.metric("Deciduous", f"{td_vol:.2f} m³", f"{math.ceil(td_load*10)/10:.1f} loads",
                  delta_color="off")

        # ---- Export ---- (Add_Date is stamped automatically with today)
        st.markdown("### Export")
        e1, e2 = st.columns(2)
        with e1:
            st.download_button("⬇️ Attributed shapefile (zip)",
                               export_shapefile_zip(ss.chunks,
                                                    ss.get("project_code", "")),
                               file_name=f"{ss.footprint_name}_tda_stands.zip",
                               mime="application/zip", use_container_width=True)
        with e2:
            st.download_button("⬇️ Stand summary (csv)",
                               export_summary_csv(ss.chunks),
                               file_name=f"{ss.footprint_name}_tda_summary.csv",
                               mime="text/csv", use_container_width=True)



    # ---- Salvage report (full width, below map + form) ----
    render_report_section(ss)


if __name__ == "__main__":
    main()
