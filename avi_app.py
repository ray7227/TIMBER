import os
import streamlit as st 
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from tempfile import NamedTemporaryFile
import base64
import re
import math
import geopandas as gpd
import zipfile
import shutil
from tempfile import mkdtemp
import tempfile
from pathlib import Path
import datetime

# --- Load TDA tables ---
@st.cache_data
def load_tda(region):
    path = f"{region.upper()}_TDA.xlsx"  # Looks for file in same directory as avi_app.py
    return pd.read_excel(path)

# --- Species mapping & choices ---
species_names = {
    "Sw": "White spruce",
    "Sb": "Black spruce",
    "P":  "Pine",
    "Fb": "Balsam fir",
    "Fd": "Douglas fir",
    "Lt": "Larch",
    "Aw": "Aspen",
    "Pb": "Balsam poplar",
    "Bw": "White birch",
}
species_codes = sorted(species_names)
species_choices = [f"{code} ({species_names[code]})" for code in species_codes]

# ← For TDA logic:
conifers = {"Sw", "Sb", "P", "Fb", "Fd", "Lt"}
deciduous = {"Aw", "Pb", "Bw"}

# --- Default values ---
default_values = {
    "is_merch": "Yes",
    "crown_density": 70,
    "avg_stand_height": 0,
    "dom_sel": species_choices[0],
    "dom_cover": 70,
    "sec_sel": "",
    "sec_cover": 30,
    "area": 1.0,
    "region": "Boreal",
    "disposition": "",
    "legal_loc": "",
    "vegetation": [],
    "other_specify_details": "",
    "disposition_fma": "",
    "no_disposition_fma": False,
    "disposition_ctlr": "",
    "salvage_waiver": "No",
    "justification": ""
}

# --- Session state initialization ---
if 'results_log' not in st.session_state:
    st.session_state.results_log = []
if 'current_entry_index' not in st.session_state:
    st.session_state.current_entry_index = -1  # -1 means new entry
if 'edit_mode' not in st.session_state:
    st.session_state.edit_mode = False
if 'show_salvage_form' not in st.session_state:
    st.session_state.show_salvage_form = False
if 'reset_trigger' not in st.session_state:
    st.session_state.reset_trigger = False
if 'dom_cover' not in st.session_state:
    st.session_state.dom_cover = default_values["dom_cover"]
if 'sec_cover' not in st.session_state:
    st.session_state.sec_cover = default_values["sec_cover"]
if 'dom_species' not in st.session_state:
    st.session_state.dom_species = species_choices[0].split(" ")[0]
if 'sec_species' not in st.session_state:
    st.session_state.sec_species = ""
if 'avg_stand_height' not in st.session_state:
    st.session_state.avg_stand_height = default_values["avg_stand_height"]
if 'is_merch' not in st.session_state:
    st.session_state.is_merch = default_values["is_merch"]
if 'crown_density' not in st.session_state:
    st.session_state.crown_density = default_values["crown_density"]
if 'dom_sel' not in st.session_state:
    st.session_state.dom_sel = default_values["dom_sel"]
if 'sec_sel' not in st.session_state:
    st.session_state.sec_sel = default_values["sec_sel"]
if 'area' not in st.session_state:
    st.session_state.area = default_values["area"]
if 'region' not in st.session_state:
    st.session_state.region = default_values["region"]

# --- Reset widget defaults if triggered ---
if st.session_state.reset_trigger:
    # Clear session state for non-widget keys
    st.session_state.results_log = []
    st.session_state.current_entry_index = -1
    st.session_state.edit_mode = False
    st.session_state.show_salvage_form = False
    st.session_state.dom_cover = default_values["dom_cover"]
    st.session_state.sec_cover = default_values["sec_cover"]
    st.session_state.dom_species = species_choices[0].split(" ")[0]
    st.session_state.sec_species = ""
    # Reset all widget-related session state keys
    st.session_state.is_merch = default_values["is_merch"]
    st.session_state.crown_density = default_values["crown_density"]
    st.session_state.avg_stand_height = default_values["avg_stand_height"]
    st.session_state.dom_sel = default_values["dom_sel"]
    st.session_state.sec_sel = default_values["sec_sel"]
    st.session_state.area = default_values["area"]
    st.session_state.region = default_values["region"]
    st.session_state.reset_trigger = False
    st.rerun()

# --- Page config ---
st.set_page_config(layout="wide")
st.header(
    "🌲 TIMBER: AVI/TDA/Report Generator",
    help="Before using this form:\n\n1. Open ArcMap and load the disturbed area .shp file into the Timber layer\n2. Use P3 satellite imagery to divide the footprint into tree stand sections and calculate the area of each polygon\n3. Identify the site LSD, locate the corresponding P3 map, and georeference it to the area using ground control points (GCPs) tied to township corners, then rectify the map\n\nFor this form:\n\nComplete the form step by step for each tree stand. Copy the values from the white boxes on the right into the ArcMap table, and enter \"Y\" if merchantable timber is present. After each stand, click “Save Entry” to save and move to a new entry. Once all stands are complete, click “Finish (Show Totals)” to calculate totals, then “Finish (Fill Salvage Draft)” to populate the final Timber form."
)

# --- Calculate AVI and volumes ---
def calculate_avi_and_volumes(is_merch, crown_density, avg_stand_height, dom_species, dom_cover, sec_species, sec_cover, area, region):
    global avi_code, c_vol, d_vol, c_load, d_load, c_vol_ha, d_vol_ha, group, total_val
    # Build AVI code
    avi_code = ""
    if is_merch.lower() == 'yes': avi_code += "m"
    if   6  <= crown_density <= 30: avi_code += "A"
    elif 31 <= crown_density <= 50: avi_code += "B"
    elif 51 <= crown_density <= 70: avi_code += "C"
    elif 71 <= crown_density <= 100: avi_code += "D"
    avi_code += str(avg_stand_height)
    avi_code += dom_species + str(dom_cover // 10)
    if dom_cover < 100 and sec_species:
        avi_code += sec_species + str(sec_cover // 10)

    # Helper functions
    def density_class(d):
        return "AB" if 6 <= d <= 50 else "CD"

    def height_bin(h):
        if h <= 4:   return "0-4"
        if h <= 8:   return "5-8"
        if h <= 10:  return "9-10"
        if h <= 25:  return str(h)  # Single values for 9-25
        if h <= 28:  return "26-28"
        return "29+"

    def get_structure_group(dom_sp, dom_pct, sec_sp, sec_pct):
        t_dec = (dom_pct if dom_sp in deciduous else 0) + (sec_pct if sec_sp in deciduous else 0)
        t_con = (dom_pct if dom_sp in conifers else 0) + (sec_pct if sec_sp in conifers else 0)
        if t_dec >= 70: return 'D'
        if t_con >= 70:
            if dom_sp == "Sw": return "C-Sw"
            if dom_sp == "P":  return "C-P"
            if dom_sp == "Sb": return "C-Sb"
            return "C-Sx"
        if t_con > 30 and t_dec < 70:
            if dom_sp == "P": return "MX-P"
            return "MX-Sx"
        return None

    try:
        df = load_tda(region)
        key = f"{height_bin(avg_stand_height)} ({density_class(crown_density)})"
        row = df[df["Height_and_Density"].str.strip() == key]
        group = get_structure_group(dom_species, dom_cover, sec_species, sec_cover)
        valid_groups = {"D", "MX-P", "MX-Sx", "C-Sw", "C-P", "C-Sb", "C-Sx"}
        total_col = f"Total ({group})" if group in valid_groups else "Total (D)"
        total_val = row[total_col].values[0] if not row.empty and total_col in df.columns else 0

        if dom_cover == 100:
            c_vol_ha = total_val if dom_species in conifers else None
            d_vol_ha = total_val if dom_species in deciduous else 0
        else:
            c_pct = (dom_cover if dom_species in conifers else 0) + (sec_cover if sec_species in conifers else 0)
            d_pct = (dom_cover if dom_species in deciduous else 0) + (sec_cover if sec_species in deciduous else 0)
            c_vol_ha = round((c_pct/100)*total_val, 1) if c_pct > 0 else None
            d_vol_ha = round((d_pct/100)*total_val, 1) if d_pct > 0 else 0

        c_vol = round(c_vol_ha * area, 5) if c_vol_ha is not None else 0
        d_vol = round(d_vol_ha * area, 5) if d_vol_ha is not None else 0
        c_load = round(c_vol / 30, 5) if c_vol is not None else 0
        d_load = round(d_vol / 30, 5) if d_vol is not None else 0
    except Exception as e:
        st.error(f"Error reading TDA table: {e}")
        c_vol = d_vol = c_load = d_load = 0
        c_vol_ha = d_vol_ha = None
        avi_code = ""
        group = None
        total_val = 0

# Initialize global variables with defaults
avi_code = ""
c_vol = d_vol = c_load = d_load = 0
c_vol_ha = d_vol_ha = None
group = None
total_val = 0

# --- Navigation bar ---
st.subheader(
    f"Entry {len(st.session_state.results_log) + 1 if st.session_state.current_entry_index == -1 else st.session_state.current_entry_index + 1} of {len(st.session_state.results_log)}"
    if st.session_state.results_log
    else "Add New Entry"
)
col_nav1, col_nav2, col_nav3 = st.columns([1, 1, 3])
with col_nav1:
    st.write("")  # Placeholder to maintain column layout
with col_nav2:
    if st.button("Save Entry"):
        # Recalculate derived species and volumes before saving
        dom_species = st.session_state.dom_sel.split(" ")[0] if st.session_state.dom_sel else ""
        sec_species = st.session_state.sec_sel.split(" ")[0] if st.session_state.sec_sel else ""
        # Update avg_stand_height in session state
        st.session_state.avg_stand_height = st.session_state.get("avg_stand_height", default_values["avg_stand_height"])
        # Calculate volumes and loads for the current entry
        calculate_avi_and_volumes(
            st.session_state.is_merch,
            st.session_state.crown_density,
            st.session_state.avg_stand_height,
            dom_species,
            st.session_state.dom_cover,
            sec_species,
            st.session_state.sec_cover,
            st.session_state.area,
            st.session_state.region
        )
        if st.session_state.current_entry_index >= 0 and st.session_state.edit_mode and st.session_state.results_log:
            # Save current entry if in edit mode
            entry_data = {
                "C_Vol": c_vol,
                "C_Load": c_load,
                "D_Vol": d_vol,
                "D_Load": d_load,
                "dom_sp": dom_species,
                "dom_pct": st.session_state.dom_cover,
                "sec_sp": sec_species,
                "sec_pct": st.session_state.sec_cover,
                "is_merch": st.session_state.is_merch == "Yes",
                "crown_density": st.session_state.crown_density,
                "avg_stand_height": st.session_state.avg_stand_height,
                "area": st.session_state.area,
                "region": st.session_state.region
            }
            st.session_state.results_log[st.session_state.current_entry_index] = entry_data
            st.success(f"Entry {st.session_state.current_entry_index + 1} saved!")
        elif st.session_state.current_entry_index == -1:
            # Save new entry
            entry_data = {
                "C_Vol": c_vol,
                "C_Load": c_load,
                "D_Vol": d_vol,
                "D_Load": d_load,
                "dom_sp": dom_species,
                "dom_pct": st.session_state.dom_cover,
                "sec_sp": sec_species,
                "sec_pct": st.session_state.sec_cover,
                "is_merch": st.session_state.is_merch == "Yes",
                "crown_density": st.session_state.crown_density,
                "avg_stand_height": st.session_state.avg_stand_height,
                "area": st.session_state.area,
                "region": st.session_state.region
            }
            st.session_state.results_log.append(entry_data)
            st.success("New entry saved!")

        # Transition to new entry without resetting input fields
        st.session_state.current_entry_index = -1
        st.session_state.edit_mode = False
        st.rerun()
with col_nav3:
    st.write(f"Entries Saved: {len(st.session_state.results_log)}")

# --- Inputs & AVI calculation ---
col1, col2 = st.columns(2)
with col1:
    # Load saved entry data if in edit mode and index is valid
    if (st.session_state.edit_mode and 
        st.session_state.current_entry_index >= 0 and 
        st.session_state.results_log and 
        st.session_state.current_entry_index < len(st.session_state.results_log)):
        entry = st.session_state.results_log[st.session_state.current_entry_index]
        st.session_state.is_merch = "Yes" if entry.get("is_merch", True) else "No"
        st.session_state.crown_density = entry.get("crown_density", default_values["crown_density"])
        st.session_state.avg_stand_height = entry.get("avg_stand_height", default_values["avg_stand_height"])
        st.session_state.dom_sel = f"{entry['dom_sp']} ({species_names[entry['dom_sp']]})"
        st.session_state.dom_cover = entry["dom_pct"]
        st.session_state.sec_cover = entry["sec_pct"]
        st.session_state.sec_sel = f"{entry['sec_sp']} ({species_names[entry['sec_sp']]})" if entry["sec_sp"] else ""
        st.session_state.area = entry.get("area", default_values["area"])
        st.session_state.region = entry.get("region", default_values["region"])

    is_merch = st.selectbox(
        "Is it merch?",
        ["Yes", "No"],
        key="is_merch",
        help="Enter \"No\" if the stand contains no trees or only very small or new growth trees, as these cannot be processed into merchantable timber (e.g., lumber or planks)."
    )
    crown_density = st.slider(
        "Crown Density (%)",
        6, 100,
        st.session_state.get("crown_density", default_values["crown_density"]),
        key="crown_density",
        help="Utilize recent satellite imagery to estimate crown density within the tree stand."
    )
    avg_stand_height = st.slider(
        "Average Stand Tree Height",
        0, 40,
        st.session_state.get("avg_stand_height", default_values["avg_stand_height"]),
        step=1,
        key="avg_stand_height",
        help="Use georeferenced P3 maps and satellite imagery to estimate tree height. The second value in old P3 AVI codes (e.g., C1SbLt) gives approximate height in meters (1=10m). Though outdated, this offers a general idea of past stand height—check map dates or cut blocks to help estimate current height. For older data, apply average growth rates: poplar 1–3 m/yr, aspen 0.5–1 m, birch 0.5–1.5 m, spruce 0.3–0.6 m, pine 0.5–1 m, fir 0.3–0.5 m, larch ~0.5 m, adjusting for local conditions. Google Earth shadow length can also be used with sun angle for trigonometric height estimates."
    )

    dom_sel = st.selectbox(
        "Dominant Species",
        species_choices,
        key="dom_sel",
        help="Enter the dominant species by percent cover within the stand"
    )
    dom_species = dom_sel.split(" ")[0]
    st.session_state.dom_species = dom_species
    
    # Dominant Cover % slider
    dom_cover = st.slider(
        "Dominant Cover %",
        0, 100,
        st.session_state.dom_cover,
        step=10,
        key="dom_cover_temp",
        on_change=lambda: st.session_state.update({
            'dom_cover': st.session_state.dom_cover_temp,
            'sec_cover': 100 - st.session_state.dom_cover_temp
        })
    )
    st.session_state.dom_cover = dom_cover

    sec_opts = [""] + [c for c in species_choices if c.split(" ")[0] != dom_species]
    sec_sel = st.selectbox(
        "2nd Species",
        sec_opts,
        key="sec_sel",
        help="Enter the second most dominant species by percent cover within the stand."
    )
    sec_species = sec_sel.split(" ")[0] if sec_sel else ""
    st.session_state.sec_species = sec_species
    
    # 2nd Cover % slider
    sec_cover = st.slider(
        "2nd Cover %",
        0, 100,
        st.session_state.sec_cover,
        step=10,
        key="sec_cover_temp",
        on_change=lambda: st.session_state.update({
            'sec_cover': st.session_state.sec_cover_temp,
            'dom_cover': 100 - st.session_state.sec_cover_temp
        })
    )
    st.session_state.sec_cover = sec_cover

    area = st.number_input(
        "Area (ha)",
        min_value=0.0,
        value=st.session_state.get("area", default_values["area"]),
        step=0.0001,
        format="%.4f",
        key="area",
        help="Enter the tree stand area (ha) as calculated in ArcMap."
    )
    region = st.selectbox(
        "Natural Region",
        ["Boreal", "Foothills"],
        key="region",
        help="Input the natural region using the ArcMap layer."
    )

# Calculate AVI and volumes after inputs are defined
calculate_avi_and_volumes(is_merch, crown_density, avg_stand_height, dom_species, dom_cover, sec_species, sec_cover, area, region)

# --- Styled outputs on the right (original colours) ---
with col2:
    st.markdown(f"""
    <div style='padding:1em; border:2px solid #4CAF50; border-radius:12px;
                background-color:#f9f9f9; color:#000;'>
        <h4 style='color:#4CAF50;'>Generated AVI Code</h4>
        <p style='font-size:24px; font-weight:bold;'>{avi_code}</p>
    </div>""", unsafe_allow_html=True)

    # Volume per Hectare (blue) with TDA values and Group
    con_vol_ha_str = "{:.5f}".format(c_vol_ha) if c_vol_ha is not None else "N/A"
    dec_vol_ha_str = "{:.5f}".format(d_vol_ha) if d_vol_ha > 0 else "0"
    st.markdown(f"""
    <div style='padding:1em; border:2px solid #2196F3; border-radius:12px;
                background-color:#f0f8ff; color:#000;'>
        <h4 style='color:#2196F3;'>Volume per Hectare</h4>
        <p><b>Con:</b> {con_vol_ha_str} m³/ha [TDA={total_val if c_vol_ha is not None else 'N/A'}, Group={group if c_vol_ha is not None else 'N/A'}]</p>
        <p><b>Dec:</b> {dec_vol_ha_str} m³/ha [TDA={total_val if d_vol_ha > 0 else 'N/A'}, Group={group if d_vol_ha > 0 else 'N/A'}]</p>
    </div>""", unsafe_allow_html=True)

    # Total Volume (orange)
    st.markdown(f"""
    <div style='padding:1em; border:2px solid #FF9800; border-radius:12px;
                background-color:#fff8e1; color:#000;'>
        <h4 style='color:#FF9800;'>Total Volume ({area} ha)</h4>
        <p><b>Con:</b> {c_vol:.5f} m³</p>
        <p><b>Dec:</b> {d_vol:.5f} m³</p>
    </div>""", unsafe_allow_html=True)

    # Load (purple)
    st.markdown(f"""
    <div style='padding:1em; border:2px solid #9C27B0; border-radius:12px;
                background-color:#f3e5f5; color:#000;'>
        <h4 style='color:#9C27B0;'>Load</h4>
        <p><b>Con:</b> {c_load:.5f}</p>
        <p><b>Dec:</b> {d_load:.5f}</p>
    </div>""", unsafe_allow_html=True)

    # --- P3 Map Search Converter ---
    # Function to convert LSD to P3 map search format
    def convert_lsd_to_p3(lsd):
        # Regular expression to match LSD format: e.g., NE-20-48-11-W5 or se-29-48-11-w5
        pattern = r'^(?:[A-Za-z]{2}-)?\d{1,2}-\d{1,3}-\d{1,2}-[Ww](\d)$'
        match = re.match(pattern, lsd.strip(), re.IGNORECASE)
        if match:
            meridian = match.group(1)  # Last group is the meridian number
            parts = lsd.strip().replace(" ", "-").split("-")
            # Extract range and township (last two parts before meridian)
            range_num = parts[-2]
            township = parts[-3]
            # Pad range to 2 digits and township to 3 digits
            range_padded = range_num.zfill(2)
            township_padded = township.zfill(3)
            return f"P3:{meridian}{range_padded}{township_padded}*"
        return None

    # Create two columns for heading and input
    col_p3_head, col_p3_input = st.columns([1, 2])
    with col_p3_head:
        st.markdown(
            "<h4 style='margin-bottom: 0; padding-bottom: 0;'>P3 Map Search Converter</h4>",
            unsafe_allow_html=True,
            help="Enter one or more LSDs (e.g., NE-20-48-11-W5) in the text area below, one per line. The output will show the SharePoint P3 map search format (P3:MRRTTT*)."
        )
    with col_p3_input:
        lsd_input = st.text_input(
            "",
            placeholder="NE-20-48-11-W5 SE-35-67-7-W6",
            key="lsd_input",
            label_visibility="collapsed"
        )

    # Process input and display output below Load box
    if lsd_input:
        # Split input by spaces or newlines
        lsds = [lsd.strip() for lsd in lsd_input.replace("\n", " ").split()]
        results = [convert_lsd_to_p3(lsd) for lsd in lsds if convert_lsd_to_p3(lsd)]
        if results:
            st.text("\n".join(results))

    # --- Tree Height Estimator ---
    st.markdown(
        "<h4 style='margin-bottom: 0; padding-bottom: 0;'>Tree Height Estimator</h4>",
        unsafe_allow_html=True,
        help="Tree growth rates assume ideal conditions and can vary with factors like species, soil, sunlight, and disturbance. Use judgment when estimating tree height for timber assessments."
    )

    # Use year and month only for P3 map update date
    p3_year = st.selectbox("P3 Map Update Year", list(range(1980, 2026)), index=2025-1980, key="p3_year")
    p3_month = st.selectbox("P3 Map Update Month", list(range(1, 13)), index=0, key="p3_month")
    p3_height = st.number_input("P3 Height (m)", min_value=0, value=10, step=1, key="p3_height")
    species_sel = st.selectbox("Species", species_choices, key="estimator_species")

    if species_sel:
        species_code = species_sel.split(" ")[0]
        growth_rates = {
            "Aw": 0.75,  # midpoint 0.5-1
            "Pb": 2.0,   # poplar 1-3, midpoint 2
            "Bw": 1.0,   # 0.5-1.5
            "Sw": 0.45,  # 0.3-0.6
            "Sb": 0.45,  # 0.3-0.6
            "P": 0.75,   # 0.5-1
            "Fb": 0.4,   # 0.3-0.5
            "Fd": 0.4,   # assume same as Fb
            "Lt": 0.5,   # 0.5
        }
        max_heights = {
            "Sw": 30,  # White spruce
            "Sb": 20,  # Black spruce
            "P": 30,   # Pine
            "Fb": 25,  # Balsam fir
            "Fd": 40,  # Douglas fir
            "Lt": 30,  # Larch
            "Aw": 25,  # Aspen
            "Pb": 30,  # Balsam poplar
            "Bw": 20,  # White birch
        }
        rate = growth_rates.get(species_code, 0.5)
        current_date = datetime.date(2025, 10, 1)
        p3_date = datetime.date(p3_year, p3_month, 1)
        days_passed = (current_date - p3_date).days
        years = max(0, days_passed / 365.25)  # Ensure non-negative
        added_growth = years * rate
        estimated_height = int(round(p3_height + added_growth))  # Round to whole number
        # Cap at species-specific maximum height
        estimated_height = min(estimated_height, max_heights.get(species_code, 40))

        st.markdown(f"""
        <div style='padding:1em; border:2px solid #607D8B; border-radius:12px;
                    background-color:#ECEFF1; color:#000;'>
          <h4 style='color:#607D8B;'>Estimated Current Height</h4>
          <p style='font-size:20px; font-weight:bold;'>{estimated_height} m</p>
        </div>""", unsafe_allow_html=True)

    # --- Tree Height Estimator from Shadows ---
    st.markdown(
        "<h4 style='margin-bottom: 0; padding-bottom: 0;'>Tree Height Estimator from Shadows</h4>",
        unsafe_allow_html=True,
        help="Estimate tree height using shadow length and approximate sun elevation for Northern Alberta (~55°N). Assumes flat ground and noon sun position. Use for formal estimation; adjust for terrain if needed."
    )

    sun_elevation_by_month = {
        "Jan": 12, "Feb": 20, "Mar": 32, "Apr": 45, "May": 55, "Jun": 60,
        "Jul": 55, "Aug": 45, "Sep": 32, "Oct": 20, "Nov": 12, "Dec": 8
    }
    shadow_month = st.selectbox("Image Month", list(sun_elevation_by_month.keys()), key="shadow_month")
    shadow_length = st.number_input("Shadow Length (m)", min_value=0.0, value=10.0, step=0.1, key="shadow_length")
    measured_height = st.number_input("Measured Height (m, optional)", min_value=0.0, value=0.0, step=0.1, key="measured_height")

    if shadow_month and shadow_length > 0:
        elev = sun_elevation_by_month[shadow_month]
        est_height = round(shadow_length * math.tan(math.radians(elev)), 2)
        display_text = f"<p><b>Month:</b> {shadow_month}</p>"
        display_text += f"<p><b>Sun Elevation:</b> {elev}°</p>"
        display_text += f"<p><b>Shadow Length:</b> {shadow_length} m</p>"
        display_text += f"<p style='font-size:20px; font-weight:bold;'>Estimated Height: {est_height:.2f} m</p>"
        if measured_height > 0:
            diff = measured_height - est_height
            display_text += f"<p><b>Measured Height:</b> {measured_height} m (Δ {diff:+.2f} m)</p>"

        st.markdown(f"""
        <div style='padding:1em; border:2px solid #607D8B; border-radius:12px;
                    background-color:#ECEFF1; color:#000;'>
          <h4 style='color:#607D8B;'>Estimated Height from Shadow</h4>
          {display_text}
        </div>""", unsafe_allow_html=True)

# --- Show totals ---
if st.button("Finish (Show Totals)", key="finish_totals"):
    total_c_vol = sum(e["C_Vol"] for e in st.session_state.results_log if e["C_Vol"] is not None)
    total_c_load = sum(e["C_Load"] for e in st.session_state.results_log if e["C_Load"] is not None)
    total_d_vol = sum(e["D_Vol"] for e in st.session_state.results_log if e["D_Vol"] is not None)
    total_d_load = sum(e["D_Load"] for e in st.session_state.results_log if e["D_Load"] is not None)
    raw_con = sum(e["dom_pct"] for e in st.session_state.results_log if e["dom_sp"] in conifers) + \
              sum(e["sec_pct"] for e in st.session_state.results_log if e["sec_sp"] in conifers)
    raw_dec = sum(e["dom_pct"] for e in st.session_state.results_log if e["dom_sp"] in deciduous) + \
              sum(e["sec_pct"] for e in st.session_state.results_log if e["sec_sp"] in deciduous)
    pct_con = round(raw_con/(raw_con+raw_dec)*100,0) if (raw_con+raw_dec)>0 else 0
    pct_dec = round(raw_dec/(raw_con+raw_dec)*100,0) if (raw_con+raw_dec)>0 else 0

    st.markdown(f"""
    <div style='padding:1em; border:2px solid #607D8B; border-radius:12px;
                background-color:#ECEFF1; color:#000;'>
      <h4 style='color:#607D8B;'>Final Tally</h4>
      <p><b>Total C_Vol:</b> {total_c_vol:.5f} m³</p>
      <p><b>Total C_Load:</b> {total_c_load:.5f}</p>
      <p><b>Total D_Vol:</b> {total_d_vol:.5f} m³</p>
      <p><b>Total D_Load:</b> {total_d_load:.5f}</p>
      <hr>
      <p><b>% Coniferous:</b> {pct_con}%</p>
      <p><b>% Deciduous:</b> {pct_dec}%</p>
    </div>""", unsafe_allow_html=True)

# --- Salvage form trigger ---
if st.button("Finish (Fill Salvage Draft)", key="finish_salvage"):
    st.session_state.show_salvage_form = True

# --- Salvage form & Word export ---
if st.session_state.show_salvage_form:
    st.subheader("Additional Information for Report Generation")

    disposition = st.text_input(
        "Disposition",
        key="disposition",
        help="The disposition will include a reference such as (RTF). If the application was submitted through a OneStop Application, it will include both the disposition type and the number, for example (RTF2525)."
    )
    legal_loc = st.text_input(
        "Legal Land Location",
        key="legal_loc",
        help="Start and end points of footprint or a single legal if it calls within one Quarter Section."
    )

    veg_types = [
        "Native grassland", "Tame pasture", "Cropland", "Sparsely or non-vegetated",
        "Cutblock - planted", "Natural regeneration >2m", "Treed wetland",
        "Shrubby wetland", "Grass or grass-like wetland", "Native aspen parkland",
        "Other (specify)"
    ]
    vegetation = st.multiselect(
        "Vegetation (check all that apply):",
        veg_types,
        key="vegetation",
        help="Broad description of what is on the project footprint. Most notable would be to use aerial imagery and field notes to determine if project is within regen/planted area. Plantations typically long straight rows of trees."
    )

    # Show text input for "Other (specify)" details if selected
    other_specify_details = ""
    if "Other (specify)" in vegetation:
        other_specify_details = st.text_input("Other (specify):", key="other_specify_details")

    disposition_fma = st.text_input(
        "Disposition # of FMA & Holder Name:",
        key="disposition_fma",
        help="Find FMA or Disposition Info on the Sketch Plan, PLSR (best source), EDP, Abadata (Terrain > Forest Management Areas), FMA/FMUMaps, or OneStop Application. If no FMA, contact the SRD field office."
    )
    no_disposition_fma = st.checkbox("None", key="no_disposition_fma")

    disposition_ctlr = st.text_input(
        "Disposition # of CTLR & Holder Name:",
        key="disposition_ctlr",
        help="Often CTLR, DTLR, etc. do not exist; if they do, look on the sketch plan, PLSR or EDP. Project plans will show the CTLR number and holder name. CTLR is divided into DTL (Deciduous Timber Licence) and CTL (Coniferous Timber Licence). Input all CTLs and DTLs in here if present."
    )

    salvage_waiver = st.radio(
        "Timber Salvage Waiver Requested?",
        ["Yes", "No"],
        horizontal=True,
        key="salvage_waiver",
        help="Use when timber is uneconomic to salvage, such as less than 0.5 truckloads. This allows legal on-site destruction of merchantable wood. Also applies to isolated decks on larger projects. Contact Alberta Forestry for guidance, as waiver rules vary by region and FMA."
    )
    if salvage_waiver == "Yes":
        justification = st.text_area("Provide justification:", key="justification")

    def fill_template():
        # --- calculate grouped percentages ---
        raw_con = sum(e["dom_pct"] for e in st.session_state.results_log if e["dom_sp"] in conifers) + \
                  sum(e["sec_pct"] for e in st.session_state.results_log if e["sec_sp"] in conifers)
        raw_dec = sum(e["dom_pct"] for e in st.session_state.results_log if e["dom_sp"] in deciduous) + \
                  sum(e["sec_pct"] for e in st.session_state.results_log if e["sec_sp"] in deciduous)
        pct_con = round(raw_con/(raw_con+raw_dec)*100,0) if (raw_con+raw_dec)>0 else 0

        # conifer splits
        spruce_raw = sum(
            e["dom_pct"] for e in st.session_state.results_log if e["dom_sp"] in {"Sw","Sb"}
        ) + sum(
            e["sec_pct"] for e in st.session_state.results_log if e["sec_sp"] in {"Sw","Sb"}
        )
        pine_raw = sum(
            e["dom_pct"] for e in st.session_state.results_log if e["dom_sp"]=="P"
        ) + sum(
            e["sec_pct"] for e in st.session_state.results_log if e["sec_sp"]=="P"
        )
        other_con = raw_con - spruce_raw - pine_raw
        if raw_con>0:
            spruce_pct = int(round(spruce_raw/raw_con*100,0))
            pine_pct = int(round(pine_raw/raw_con*100,0))
            other_con_pct = int(round(100 - spruce_pct - pine_pct,0))
        else:
            spruce_pct = pine_pct = other_con_pct = 0

        # deciduous splits
        aspen_raw = sum(
            e["dom_pct"] for e in st.session_state.results_log if e["dom_sp"]=="Aw"
        ) + sum(
            e["sec_pct"] for e in st.session_state.results_log if e["sec_sp"]=="Aw"
        )
        other_dec = raw_dec - aspen_raw
        if raw_dec>0:
            aspen_pct = int(round(aspen_raw/raw_dec*100,0))
            other_dec_pct = int(round(100 - aspen_pct,0))
        else:
            aspen_pct = other_dec_pct = 0

        # --- determine coniferous class checkbox ---
        def con_class_box(label):
            if label == "D" and pct_con < 30:
                return "☒"
            elif label == "C" and pct_con > 70:
                return "☒"
            elif label == "CD" and 50 <= pct_con <= 70:
                return "☒"
            elif label == "DC" and 30 <= pct_con < 50:
                return "☒"
            return "☐"

        # --- now build the Word file with adjusted formatting ---
        doc = Document()

        # Title
        p = doc.add_paragraph(); p.alignment = 1
        run = p.add_run("Vegetation and Timber Salvage Information")
        run.font.name = "Times New Roman"; run.font.size = Pt(11)
        run.font.bold = True; run.font.underline = True
        p.paragraph_format.space_after = Pt(0)

        # Disposition
        p = doc.add_paragraph(); p.alignment = 1
        r1 = p.add_run("Disposition: ");   r1.font.name = "Times New Roman"; r1.font.size = Pt(10); r1.font.bold = True
        r2 = p.add_run(disposition);       r2.font.name = "Times New Roman"; r2.font.size = Pt(10); r2.font.bold = False
        p.paragraph_format.space_after = Pt(0)

        # Legal Land Location
        p = doc.add_paragraph(); p.alignment = 1
        r1 = p.add_run("Legal Land Location: ");   r1.font.name = "Times New Roman"; r1.font.size = Pt(10); r1.font.bold = True
        r2 = p.add_run(legal_loc);                 r2.font.name = "Times New Roman"; r2.font.size = Pt(10); r2.font.bold = False
        p.paragraph_format.space_after = Pt(0)

        # Horizontal line
        p = doc.add_paragraph()
        p_pr = p._p.get_or_add_pPr()
        bdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single'); bottom.set(qn('w:sz'), '24'); bottom.set(qn('w:space'), '1'); bottom.set(qn('w:color'), '000000')
        bdr.append(bottom); p_pr.append(bdr)

        # Vegetation and Timber Cover header
        p = doc.add_paragraph()
        run = p.add_run("Vegetation and Timber Cover")
        run.font.name = "Times New Roman"; run.font.size = Pt(12); run.font.bold = True
        p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(6)

        # Vegetation subheader
        p = doc.add_paragraph()
        run = p.add_run("Vegetation (check all that apply)")
        run.font.name = "Times New Roman"; run.font.size = Pt(10); run.font.bold = True
        p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)

        # Paired checkboxes with formatting
        def box(l): return "☒" if l in vegetation else "☐"
        rows = [
            ("Native grassland", "Treed wetland"),
            ("Tame pasture", "Shrubby wetland"),
            ("Cropland", "Grass or grass-like wetland"),
            ("Sparsely or non-vegetated", "Native aspen parkland"),
            ("Cutblock - planted", "Other (specify)"),
        ]
        for left, right in rows:
            p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
            left_indent = "" if left in ["Native grassland", "Tame pasture", "Cropland", "Sparsely or non-vegetated", "Cutblock - planted"] else "\t"
            right_indent = "\t\t" if left in ["Tame pasture", "Cropland"] else "\t" if left not in ["Sparsely or non-vegetated", "Tame pasture", "Cropland"] else ""
            if right == "Treed wetland":
                extra_text = "\t\t"
                run = p.add_run(f"{left_indent}{box(left)} {left}{right_indent}\t{box(right)} {right}{extra_text}")
                run.font.name = "Times New Roman"; run.font.size = Pt(10); run.font.bold = True
                run = p.add_run("Deciduous-dominant Forest:")
                run.font.name = "Times New Roman"; run.font.size = Pt(10); run.font.bold = True; run.font.underline = True
            elif right == "Shrubby wetland":
                run = p.add_run(f"{left_indent}{box(left)} {left}{right_indent}\t{box(right)} {right}\t\t{con_class_box('D')} D less than 30% coniferous")
                run.font.name = "Times New Roman"; run.font.size = Pt(10); run.font.bold = True
            elif right == "Grass or grass-like wetland":
                extra_text = "\t"
                run = p.add_run(f"{left_indent}{box(left)} {left}{right_indent}\t{box(right)} {right}{extra_text}")
                run.font.name = "Times New Roman"; run.font.size = Pt(10); run.font.bold = True
                run = p.add_run("Coniferous-dominant Forest:")
                run.font.name = "Times New Roman"; run.font.size = Pt(10); run.font.bold = True; run.font.underline = True
            elif right == "Native aspen parkland":
                run = p.add_run(f"{left_indent}{box(left)} {left}{right_indent}\t{box(right)} {right}\t\t{con_class_box('C')} C More than 70% coniferous")
                run.font.name = "Times New Roman"; run.font.size = Pt(10); run.font.bold = True
            elif right == "Other (specify)":
                run = p.add_run(f"{left_indent}{box(left)} {left}{right_indent}\t{box(right)} {right}")
                run.font.name = "Times New Roman"; run.font.size = Pt(10); run.font.bold = True
                if "Other (specify)" in vegetation and other_specify_details:
                    run = p.add_run(f": ")
                    run.font.name = "Times New Roman"; run.font.size = Pt(10); run.font.bold = True
                    run = p.add_run(other_specify_details)
                    run.font.name = "Times New Roman"; run.font.size = Pt(10); run.font.bold = False; run.font.underline = True
                run = p.add_run("\t\t")
                run.font.name = "Times New Roman"; run.font.size = Pt(10); run.font.bold = True
                run = p.add_run("Mixedwood Forest:")
                run.font.name = "Times New Roman"; run.font.size = Pt(10); run.font.bold = True; run.font.underline = True
            else:
                run = p.add_run(f"{left_indent}{box(left)} {left}{right_indent}\t{box(right)} {right}")
                run.font.name = "Times New Roman"; run.font.size = Pt(10); run.font.bold = True
        p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
        run = p.add_run(f"{box('Natural regeneration >2m')} Natural regeneration >2m\t\t\t\t\t{con_class_box('CD')} CD 70% to 50% coniferous")
        run.font.name = "Times New Roman"; run.font.size = Pt(10); run.font.bold = True
        p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
        run = p.add_run(f"\t\t\t\t\t\t\t\t{con_class_box('DC')} DC 50% to 30% coniferous")
        run.font.name = "Times New Roman"; run.font.size = Pt(10); run.font.bold = True

        # Timber Salvage header
        p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(6); p.paragraph_format.space_after = Pt(0)
        run = p.add_run("Timber Salvage:")
        run.font.name = "Times New Roman"; run.font.size = Pt(10); run.font.bold = True; run.font.underline = True

        # 1. Merchantable timber...
        yes = "☒" if is_merch == "Yes" else "☐"; no = "☒" if is_merch == "No" else "☐"
        p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(6); p.paragraph_format.space_after = Pt(0)
        run = p.add_run(f"1.\tMerchantable timber present?   {yes} Yes    {no} No")
        run.font.name = "Times New Roman"; run.font.size = Pt(10); run.font.bold = True

        # Provide volume inventory
        p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
        run = p.add_run(f"\tProvide a volume inventory as follows:"); 
        run.font.name = "Times New Roman"; run.font.size = Pt(10); run.font.bold = True

        total_c_vol = sum(e["C_Vol"] for e in st.session_state.results_log if e["C_Vol"] is not None)
        total_c_load = sum(e["C_Load"] for e in st.session_state.results_log if e["C_Load"] is not None)
        total_d_vol = sum(e["D_Vol"] for e in st.session_state.results_log if e["D_Vol"] is not None)
        total_d_load = sum(e["D_Load"] for e in st.session_state.results_log if e["D_Load"] is not None)

        # Round volume and load to one decimal place using math.ceil
        total_c_vol = math.ceil(total_c_vol * 10) / 10
        total_c_load = math.ceil(total_c_load * 10) / 10
        total_d_vol = math.ceil(total_d_vol * 10) / 10
        total_d_load = math.ceil(total_d_load * 10) / 10

        # Coniferous volume
        p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
        run1 = p.add_run(f"\tConiferous approx. volume: "); 
        run1.font.name = "Times New Roman"; run1.font.size = Pt(10); run1.font.bold = True
        run2 = p.add_run(f"{total_c_vol:.1f}"); 
        run2.font.name = "Times New Roman"; run2.font.size = Pt(10); run2.font.bold = False; run2.font.underline = True
        run3 = p.add_run(f" m³"); 
        run3.font.name = "Times New Roman"; run3.font.size = Pt(10); run3.font.bold = False; run3.font.underline = True
        run4 = p.add_run(f"  or  "); 
        run4.font.name = "Times New Roman"; run4.font.size = Pt(10); run4.font.bold = True
        run5 = p.add_run(f"{total_c_load:.1f}"); 
        run5.font.name = "Times New Roman"; run5.font.size = Pt(10); run5.font.bold = False; run5.font.underline = True
        run6 = p.add_run(f" loads"); 
        run6.font.name = "Times New Roman"; run6.font.size = Pt(10); run6.font.bold = False; run6.font.underline = True

        # Coniferous species percentages
        p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
        run1 = p.add_run(f"\tSpruce "); 
        run1.font.name = "Times New Roman"; run1.font.size = Pt(10); run1.font.bold = True
        run2 = p.add_run(f"{spruce_pct}%"); 
        run2.font.name = "Times New Roman"; run2.font.size = Pt(10); run2.font.bold = False; run2.font.underline = True
        run3 = p.add_run(f"    Pine "); 
        run3.font.name = "Times New Roman"; run3.font.size = Pt(10); run3.font.bold = True
        run4 = p.add_run(f"{pine_pct}%"); 
        run4.font.name = "Times New Roman"; run4.font.size = Pt(10); run4.font.bold = False; run4.font.underline = True
        run5 = p.add_run(f"    Other "); 
        run5.font.name = "Times New Roman"; run5.font.size = Pt(10); run5.font.bold = True
        run6 = p.add_run(f"{other_con_pct}%"); 
        run6.font.name = "Times New Roman"; run6.font.size = Pt(10); run6.font.bold = False; run6.font.underline = True

        # Deciduous volume
        p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
        run1 = p.add_run(f"\tDeciduous approx. volume: "); 
        run1.font.name = "Times New Roman"; run1.font.size = Pt(10); run1.font.bold = True
        run2 = p.add_run(f"{total_d_vol:.1f}"); 
        run2.font.name = "Times New Roman"; run2.font.size = Pt(10); run2.font.bold = False; run2.font.underline = True
        run3 = p.add_run(f" m³"); 
        run3.font.name = "Times New Roman"; run3.font.size = Pt(10); run3.font.bold = False; run3.font.underline = True
        run4 = p.add_run(f"  or  "); 
        run4.font.name = "Times New Roman"; run4.font.size = Pt(10); run4.font.bold = True
        run5 = p.add_run(f"{total_d_load:.1f}"); 
        run5.font.name = "Times New Roman"; run5.font.size = Pt(10); run5.font.bold = False; run5.font.underline = True
        run6 = p.add_run(f" loads"); 
        run6.font.name = "Times New Roman"; run6.font.size = Pt(10); run6.font.bold = False; run6.font.underline = True

        # Deciduous species percentages
        p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
        run1 = p.add_run(f"\tAspen "); 
        run1.font.name = "Times New Roman"; run1.font.size = Pt(10); run1.font.bold = True
        run2 = p.add_run(f"{aspen_pct}%"); 
        run2.font.name = "Times New Roman"; run2.font.size = Pt(10); run2.font.bold = False; run2.font.underline = True
        run3 = p.add_run(f"    Other "); 
        run3.font.name = "Times New Roman"; run3.font.size = Pt(10); run3.font.bold = True
        run4 = p.add_run(f"{other_dec_pct}%"); 
        run4.font.name = "Times New Roman"; run4.font.size = Pt(10); run4.font.bold = False; run4.font.underline = True

        # Section 2: Timber disposition or FMA(s)
        p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(6); p.paragraph_format.space_after = Pt(0)
        run = p.add_run(f"2.\tSpecify the timber disposition or FMA(s) shown on LSAS:")
        run.font.name = "Times New Roman"; run.font.size = Pt(10); run.font.bold = True

        p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
        run = p.add_run(f"\t{'☒' if no_disposition_fma else '☐'} No disposition (Contact SRD field office)")
        run.font.name = "Times New Roman"; run.font.size = Pt(10); run.font.bold = True

        p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
        run = p.add_run(f"\tDisposition number & Holder name of FMA: ")
        run.font.name = "Times New Roman"; run.font.size = Pt(10); run.font.bold = True
        run = p.add_run(disposition_fma)
        run.font.name = "Times New Roman"; run.font.size = Pt(10); run.font.bold = False; run.font.underline = True

        p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
        run = p.add_run(f"\tDisposition number & Holder name of CTLR: ")
        run.font.name = "Times New Roman"; run.font.size = Pt(10); run.font.bold = True
        run = p.add_run(disposition_ctlr)
        run.font.name = "Times New Roman"; run.font.size = Pt(10); run.font.bold = False; run.font.underline = True

        # Section 3: Utilization Standards
        p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(6); p.paragraph_format.space_after = Pt(0)
        run = p.add_run(f"3.\tUtilization Standards:")
        run.font.name = "Times New Roman"; run.font.size = Pt(10); run.font.bold = True

        p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
        run1 = p.add_run("\tConiferous ")
        run1.font.name = "Times New Roman"; run1.font.size = Pt(10); run1.font.bold = True
        run2 = p.add_run("15")
        run2.font.name = "Times New Roman"; run2.font.size = Pt(10); run2.font.bold = False; run2.font.underline = True
        run3 = p.add_run(" cm stump diameter to a ")
        run3.font.name = "Times New Roman"; run3.font.size = Pt(10); run3.font.bold = True
        run4 = p.add_run("11")
        run4.font.name = "Times New Roman"; run4.font.size = Pt(10); run4.font.bold = False; run4.font.underline = True
        run5 = p.add_run(" cm top diameter.")
        run5.font.name = "Times New Roman"; run5.font.size = Pt(10); run5.font.bold = True

        p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
        run1 = p.add_run("\tDeciduous ")
        run1.font.name = "Times New Roman"; run1.font.size = Pt(10); run1.font.bold = True
        run2 = p.add_run("15")
        run2.font.name = "Times New Roman"; run2.font.size = Pt(10); run2.font.bold = False; run2.font.underline = True
        run3 = p.add_run(" cm stump diameter to a ")
        run3.font.name = "Times New Roman"; run3.font.size = Pt(10); run3.font.bold = True
        run4 = p.add_run("10")
        run4.font.name = "Times New Roman"; run4.font.size = Pt(10); run4.font.bold = False; run4.font.underline = True
        run5 = p.add_run(" cm top diameter.")
        run5.font.name = "Times New Roman"; run5.font.size = Pt(10); run5.font.bold = True

        # Section 4: Timber salvage waiver
        box_yes = "☒" if salvage_waiver == "Yes" else "☐"
        box_no = "☒" if salvage_waiver == "No" else "☐"

        p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(6); p.paragraph_format.space_after = Pt(0)
        run = p.add_run(f"4.\tTimber salvage waiver requested?   {box_yes} Yes   {box_no} No")
        run.font.name = "Times New Roman"; run.font.size = Pt(10); run.font.bold = True

        p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
        run = p.add_run(f"\tIf ‘Yes’, provide justification: ")
        run.font.name = "Times New Roman"; run.font.size = Pt(10); run.font.bold = True
        if salvage_waiver == "Yes":
            run = p.add_run(justification)
            run.font.name = "Times New Roman"; run.font.size = Pt(10); run.font.bold = False; run.font.underline = True

        tmp = NamedTemporaryFile(delete=False, suffix=".docx")
        doc.save(tmp.name)
        return tmp.name

    if st.button(
        "Done (Generate Report)",
        help="Save Timber form and convert to PDF. Provide to the form to AIM Lands staff and they can submit it to the FMA."
    ):
        out_path = fill_template()
        if out_path:
            st.success("Report generated!")
            with open(out_path, "rb") as f:
                b64 = base64.b64encode(f.read()).decode()
                st.markdown(
                    f'<a href="data:application/octet-stream;base64,{b64}" '
                    f'download="Filled_Salvage_Report.docx">📥 Download report</a>',
                    unsafe_allow_html=True
                )

# --- Reset ---
if st.button("Reset All Entries"):
    st.session_state.reset_trigger = True
    st.rerun()

# --- Shapefile Dissolver in Sidebar ---
st.sidebar.header("Shapefile Dissolver Tool")
st.sidebar.markdown("Drag and drop zip files containing shapefiles to dissolve polygons individually.")

# File uploader for drag-and-drop
uploaded_files = st.sidebar.file_uploader(
    "Upload .zip files", 
    type=["zip"], 
    accept_multiple_files=True, 
    help="Select or drag and drop .zip files containing shapefiles."
)

# Create a temporary directory for processing
temp_base_dir = Path(tempfile.mkdtemp())
output_dir = temp_base_dir / "dissolved_output"
output_dir.mkdir(parents=True, exist_ok=True)

# Log file for debugging
log_file = output_dir / "processing_log.txt"
with open(log_file, "w") as log:
    log.write("Processing started\n")

# Process uploaded files
if uploaded_files:
    for uploaded_file in uploaded_files:
        # Save uploaded .zip file to temporary directory
        zip_path = temp_base_dir / uploaded_file.name
        with open(zip_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        with open(log_file, "a") as log:
            log.write(f"\nProcessing {zip_path.name}...\n")
            st.sidebar.write(f"Processing {zip_path.name}...")

            # Create subfolder for this specific zip
            zip_subdir = output_dir / zip_path.stem
            zip_subdir.mkdir(exist_ok=True)

            # Temporary extraction folder
            temp_dir = temp_base_dir / f"temp_{zip_path.stem}"
            temp_dir.mkdir(exist_ok=True)

            try:
                # Extract all contents
                with zipfile.ZipFile(zip_path, "r") as z:
                    z.extractall(temp_dir)
                log.write(f"Extracted {zip_path.name} to {temp_dir}\n")

                # Find shapefiles
                shapefiles = list(temp_dir.glob("*.shp"))
                if not shapefiles:
                    log.write(f"No shapefiles found in {zip_path.name}, skipping.\n")
                    st.sidebar.warning(f"No shapefiles found in {zip_path.name}, skipping.")
                    continue

                if len(shapefiles) > 1:
                    log.write(f"Warning: Multiple shapefiles found in {zip_path.name}. Processing only: {shapefiles[0]}\n")
                    st.sidebar.warning(f"Warning: Multiple shapefiles found. Using: {shapefiles[0]}")

                # Load the shapefile
                try:
                    gdf = gpd.read_file(shapefiles[0])
                    log.write(f"Loaded shapefile: {shapefiles[0]}\n")
                except Exception as e:
                    log.write(f"Error reading {shapefiles[0]}: {str(e)}\n")
                    st.sidebar.error(f"Error reading {shapefiles[0]}: {str(e)}")
                    continue

                # Check for polygon geometries
                if not gdf.geometry.type.str.contains("Polygon|MultiPolygon").any():
                    log.write(f"Shapefile {shapefiles[0]} contains no polygons, skipping.\n")
                    st.sidebar.warning(f"Shapefile {shapefiles[0]} contains no polygons, skipping.")
                    continue

                # Fix invalid geometries
                gdf["geometry"] = gdf.geometry.buffer(0)  # Attempt to fix invalid geometries
                if gdf.geometry.is_valid.all():
                    log.write("All geometries are valid after buffering.\n")
                else:
                    log.write("Warning: Some geometries are still invalid after buffering.\n")
                    st.sidebar.warning("Warning: Some geometries are invalid.")

                # Dissolve all polygons into a single geometry
                try:
                    dissolved_gdf = gdf.dissolve()  # Dissolve all into one polygon
                    log.write("Dissolve operation successful.\n")
                except Exception as e:
                    log.write(f"Error during dissolve: {str(e)}\n")
                    st.sidebar.error(f"Error during dissolve: {str(e)}")
                    continue

                # Save the dissolved shapefile in the subfolder
                out_file = zip_subdir / f"{zip_path.stem}_dissolved.shp"
                try:
                    dissolved_gdf.to_file(out_file)
                    log.write(f"Saved dissolved shapefile: {out_file}\n")
                    st.sidebar.success(f"✅ Saved dissolved shapefile: {out_file}")
                except Exception as e:
                    log.write(f"Error saving {out_file}: {str(e)}\n")
                    st.sidebar.error(f"Error saving {out_file}: {str(e)}")

            finally:
                # Cleanup temporary folder
                if temp_dir.exists():
                    shutil.rmtree(temp_dir)
                    log.write(f"Cleaned up {temp_dir}\n")

    # Zip the entire output directory for download
    output_zip_path = temp_base_dir / "dissolved_shapefiles.zip"
    with zipfile.ZipFile(output_zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(output_dir):
            for file in files:
                zipf.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), output_dir))

    # Provide download link for the zipped outputs
    with open(output_zip_path, "rb") as f:
        st.sidebar.download_button(
            label="Download All Dissolved Shapefiles (Zip)",
            data=f,
            file_name="dissolved_shapefiles.zip",
            mime="application/zip"
        )

    # Provide download link for the log file
    with open(log_file, "rb") as f:
        st.sidebar.download_button(
            label="Download Processing Log",
            data=f,
            file_name=log_file.name,
            mime="text/plain"
        )
    
    st.sidebar.success("🎉 All zip files processed.")

# Cleanup temporary base directory when done
if temp_base_dir.exists():
    shutil.rmtree(temp_base_dir)
