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
    path = f"{region.upper()}_TDA.xlsx"
    return pd.read_excel(path)

# --- Species mapping & choices ---
species_names = {
    "Sw": "White spruce",
    "Sb": "Black spruce",
    "P": "Pine",
    "Fb": "Balsam fir",
    "Fd": "Douglas fir",
    "Lt": "Larch",
    "Aw": "Aspen",
    "Pb": "Balsam poplar",
    "Bw": "White birch",
}
species_codes = sorted(species_names)
species_choices = [f"{code} ({species_names[code]})" for code in species_codes]

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
    st.session_state.current_entry_index = -1
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
if 'ctlr_list' not in st.session_state:
    st.session_state.ctlr_list = [{"type": "", "number_holder": ""}]

# --- Reset logic ---
if st.session_state.reset_trigger:
    st.session_state.results_log = []
    st.session_state.current_entry_index = -1
    st.session_state.edit_mode = False
    st.session_state.show_salvage_form = False
    st.session_state.dom_cover = default_values["dom_cover"]
    st.session_state.sec_cover = default_values["sec_cover"]
    st.session_state.dom_species = species_choices[0].split(" ")[0]
    st.session_state.sec_species = ""
    st.session_state.is_merch = default_values["is_merch"]
    st.session_state.crown_density = default_values["crown_density"]
    st.session_state.avg_stand_height = default_values["avg_stand_height"]
    st.session_state.dom_sel = default_values["dom_sel"]
    st.session_state.sec_sel = default_values["sec_sel"]
    st.session_state.area = default_values["area"]
    st.session_state.region = default_values["region"]
    st.session_state.ctlr_list = [{"type": "", "number_holder": ""}]
    st.session_state.reset_trigger = False
    st.rerun()

# --- Page config ---
st.set_page_config(layout="wide")

st.header(
    "🌲 TIMBER: AVI/TDA/Report Generator",
    help="Before using this form:\n\n1. Open ArcMap and load the disturbed area .shp file into the Timber layer\n2. Use P3 satellite imagery to divide the footprint into tree stand sections and calculate the area of each polygon\n3. Identify the site LSD, locate the corresponding P3 map, and georeference it to the area using ground control points (GCPs) tied to township corners, then rectify the map\n\nFor this form:\n\nComplete the form step by step for each tree stand. Copy the values from the white boxes on the right into the ArcMap table, and enter \"Y\" if merchantable timber is present. After each stand, click “Save Entry” to save and move to a new entry. Once all stands are complete, click “Finish (Show Totals)” to calculate totals, then “Finish (Fill Salvage Draft)” to populate the final Timber form."
)

# --- AVI & volume calculation ---
def calculate_avi_and_volumes(is_merch, crown_density, avg_stand_height, dom_species, dom_cover, sec_species, sec_cover, area, region):
    global avi_code, c_vol, d_vol, c_load, d_load, c_vol_ha, d_vol_ha, group, total_val
    avi_code = ""
    if is_merch.lower() == 'yes': avi_code += "m"
    if 6 <= crown_density <= 30: avi_code += "A"
    elif 31 <= crown_density <= 50: avi_code += "B"
    elif 51 <= crown_density <= 70: avi_code += "C"
    elif 71 <= crown_density <= 100: avi_code += "D"
    avi_code += str(avg_stand_height)
    avi_code += dom_species + str(dom_cover // 10)
    if dom_cover < 100 and sec_species:
        avi_code += sec_species + str(sec_cover // 10)

    def density_class(d):
        return "AB" if 6 <= d <= 50 else "CD"

    def height_bin(h):
        if h <= 4: return "0-4"
        if h <= 8: return "5-8"
        if h <= 10: return "9-10"
        if h <= 25: return str(h)
        if h <= 28: return "26-28"
        return "29+"

    def get_structure_group(dom_sp, dom_pct, sec_sp, sec_pct):
        t_dec = (dom_pct if dom_sp in deciduous else 0) + (sec_pct if sec_sp in deciduous else 0)
        t_con = (dom_pct if dom_sp in conifers else 0) + (sec_pct if sec_sp in conifers else 0)
        if t_dec >= 70: return 'D'
        if t_con >= 70:
            if dom_sp == "Sw": return "C-Sw"
            if dom_sp == "P": return "C-P"
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

avi_code = ""
c_vol = d_vol = c_load = d_load = 0
c_vol_ha = d_vol_ha = None
group = None
total_val = 0

# --- Navigation ---
st.subheader(
    f"Entry {len(st.session_state.results_log) + 1 if st.session_state.current_entry_index == -1 else st.session_state.current_entry_index + 1} of {len(st.session_state.results_log)}"
    if st.session_state.results_log else "Add New Entry"
)

col_nav1, col_nav2, col_nav3 = st.columns([1, 1, 3])
with col_nav1:
    st.write("")
with col_nav2:
    if st.button("Save Entry"):
        dom_species = st.session_state.dom_sel.split(" ")[0] if st.session_state.dom_sel else ""
        sec_species = st.session_state.sec_sel.split(" ")[0] if st.session_state.sec_sel else ""
        st.session_state.avg_stand_height = st.session_state.get("avg_stand_height", default_values["avg_stand_height"])

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

        if st.session_state.current_entry_index >= 0 and st.session_state.edit_mode and st.session_state.results_log:
            st.session_state.results_log[st.session_state.current_entry_index] = entry_data
            st.success(f"Entry {st.session_state.current_entry_index + 1} saved!")
        else:
            st.session_state.results_log.append(entry_data)
            st.success("New entry saved!")

        st.session_state.current_entry_index = -1
        st.session_state.edit_mode = False
        st.rerun()
with col_nav3:
    st.write(f"Entries Saved: {len(st.session_state.results_log)}")

# --- Main inputs ---
col1, col2 = st.columns(2)
with col1:
    if (st.session_state.edit_mode and st.session_state.current_entry_index >= 0 and
        st.session_state.results_log and st.session_state.current_entry_index < len(st.session_state.results_log)):
        entry = st.session_state.results_log[st.session_state.current_entry_index]
        st.session_state.is_merch = "Yes" if entry.get("is_merch", True) else "No"
        st.session_state.crown_density = entry.get("crown_density", default_values["crown_density"])
        st.session_state.avg_stand_height = entry.get("avg_stand_height", default_values["avg_stand_height"])
        st.session_state.dom_sel = f"{entry['dom_sp']} ({species_names[entry['dom_sp']]})"
        st.session_state.dom_cover = entry["dom_pct"]
        st.session_state.sec_cover = entry["sec_pct"]
        st.session_state.sec_sel = f"{entry['sec_sp']} ({species_names[entry['sec_sp']]})" if entry.get("sec_sp") else ""
        st.session_state.area = entry.get("area", default_values["area"])
        st.session_state.region = entry.get("region", default_values["region"])

    is_merch = "Yes"
    st.session_state.is_merch = "Yes"

    crown_density = st.slider(
        "Crown Density (%)", 6, 100,
        st.session_state.get("crown_density", default_values["crown_density"]),
        key="crown_density"
    )

    avg_stand_height = st.slider(
        "Average Stand Tree Height", 0, 40,
        st.session_state.get("avg_stand_height", default_values["avg_stand_height"]),
        step=1, key="avg_stand_height"
    )

    dom_sel = st.selectbox("Dominant Species", species_choices, key="dom_sel")
    dom_species = dom_sel.split(" ")[0]
    st.session_state.dom_species = dom_species

    # ── Linked cover sliders ──
    if "dom_cover" not in st.session_state:
        st.session_state.dom_cover = 70
    if "sec_cover" not in st.session_state:
        st.session_state.sec_cover = 30

    def sync_covers():
        if st.session_state.dom_cover + st.session_state.sec_cover != 100:
            st.session_state.sec_cover = 100 - st.session_state.dom_cover

    sync_covers()

    col_dom, col_sec = st.columns(2)

    with col_dom:
        st.slider(
            "Dominant Cover %", 0, 100,
            value=st.session_state.dom_cover,
            step=10,
            key="dom_cover_widget",
            on_change=lambda: st.session_state.update({
                "dom_cover": st.session_state.dom_cover_widget,
                "sec_cover": 100 - st.session_state.dom_cover_widget
            })
        )
        st.session_state.dom_cover = st.session_state.dom_cover_widget

    with col_sec:
        st.slider(
            "2nd Cover %", 0, 100,
            value=st.session_state.sec_cover,
            step=10,
            key="sec_cover_widget",
            on_change=lambda: st.session_state.update({
                "sec_cover": st.session_state.sec_cover_widget,
                "dom_cover": 100 - st.session_state.sec_cover_widget
            })
        )
        st.session_state.sec_cover = st.session_state.sec_cover_widget

    # Final safety net
    if st.session_state.dom_cover + st.session_state.sec_cover != 100:
        st.session_state.sec_cover = 100 - st.session_state.dom_cover
        st.rerun()

    sec_opts = [""] + [c for c in species_choices if c.split(" ")[0] != dom_species]
    sec_sel = st.selectbox("2nd Species", sec_opts, key="sec_sel")
    sec_species = sec_sel.split(" ")[0] if sec_sel else ""
    st.session_state.sec_species = sec_species

    area = st.number_input("Area (ha)", min_value=0.0, value=st.session_state.get("area", 1.0),
                           step=0.0001, format="%.4f", key="area")

    region = st.selectbox("Natural Region", ["Boreal", "Foothills"], key="region")

calculate_avi_and_volumes(is_merch, crown_density, avg_stand_height, dom_species,
                          st.session_state.dom_cover, sec_species, st.session_state.sec_cover,
                          area, region)

# --- Right column outputs ---
with col2:
    st.markdown(f"""
    <div style='padding:1em; border:2px solid #4CAF50; border-radius:12px; background-color:#f9f9f9;'>
        <h4 style='color:#4CAF50;'>Generated AVI Code</h4>
        <p style='font-size:24px; font-weight:bold;'>{avi_code}</p>
    </div>""", unsafe_allow_html=True)

    con_vol_ha_str = f"{c_vol_ha:.5f}" if c_vol_ha is not None else "N/A"
    dec_vol_ha_str = f"{d_vol_ha:.5f}" if d_vol_ha > 0 else "0"

    st.markdown(f"""
    <div style='padding:1em; border:2px solid #2196F3; border-radius:12px; background-color:#f0f8ff;'>
        <h4 style='color:#2196F3;'>Volume per Hectare</h4>
        <p><b>Con:</b> {con_vol_ha_str} m³/ha [TDA={total_val if c_vol_ha is not None else 'N/A'}, Group={group if c_vol_ha is not None else 'N/A'}]</p>
        <p><b>Dec:</b> {dec_vol_ha_str} m³/ha [TDA={total_val if d_vol_ha > 0 else 'N/A'}, Group={group if d_vol_ha > 0 else 'N/A'}]</p>
    </div>""", unsafe_allow_html=True)

    st.markdown(f"""
    <div style='padding:1em; border:2px solid #FF9800; border-radius:12px; background-color:#fff8e1;'>
        <h4 style='color:#FF9800;'>Total Volume ({area} ha)</h4>
        <p><b>Con:</b> {c_vol:.5f} m³</p>
        <p><b>Dec:</b> {d_vol:.5f} m³</p>
    </div>""", unsafe_allow_html=True)

    st.markdown(f"""
    <div style='padding:1em; border:2px solid #9C27B0; border-radius:12px; background-color:#f3e5f5;'>
        <h4 style='color:#9C27B0;'>Load</h4>
        <p><b>Con:</b> {c_load:.5f}</p>
        <p><b>Dec:</b> {d_load:.5f}</p>
    </div>""", unsafe_allow_html=True)

    # P3 Map Search Converter
    def convert_lsd_to_p3(lsd):
        pattern = r'^(?:[A-Za-z]{2}-)?\d{1,2}-\d{1,3}-\d{1,2}-[Ww](\d)$'
        match = re.match(pattern, lsd.strip(), re.IGNORECASE)
        if match:
            meridian = match.group(1)
            parts = lsd.strip().replace(" ", "-").split("-")
            range_num = parts[-2].zfill(2)
            township = parts[-3].zfill(3)
            return f"P3:{meridian}{range_num}{township}*"
        return None

    st.subheader("P3 Map Search Converter", help="Enter LSDs (e.g. NE-20-48-11-W5), one per line or space-separated.")
    lsd_input = st.text_input("", placeholder="NE-20-48-11-W5 SE-35-67-7-W6", key="lsd_input", label_visibility="collapsed")
    if lsd_input:
        lsds = [lsd.strip() for lsd in lsd_input.replace("\n", " ").split() if lsd.strip()]
        results = [convert_lsd_to_p3(lsd) for lsd in lsds if convert_lsd_to_p3(lsd)]
        if results:
            st.text("\n".join(results))

# --- Totals display (without the subheader here) ---
if st.button("Finish (Show Totals)", key="finish_totals"):
    total_c_vol = sum(e["C_Vol"] for e in st.session_state.results_log if e["C_Vol"] is not None)
    total_c_load = sum(e["C_Load"] for e in st.session_state.results_log if e["C_Load"] is not None)
    total_d_vol = sum(e["D_Vol"] for e in st.session_state.results_log if e["D_Vol"] is not None)
    total_d_load = sum(e["D_Load"] for e in st.session_state.results_log if e["D_Load"] is not None)

    raw_con = sum(e["dom_pct"] for e in st.session_state.results_log if e["dom_sp"] in conifers) + \
              sum(e["sec_pct"] for e in st.session_state.results_log if e["sec_sp"] in conifers)
    raw_dec = sum(e["dom_pct"] for e in st.session_state.results_log if e["dom_sp"] in deciduous) + \
              sum(e["sec_pct"] for e in st.session_state.results_log if e["sec_sp"] in deciduous)

    pct_con = round(raw_con / (raw_con + raw_dec) * 100, 0) if (raw_con + raw_dec) > 0 else 0
    pct_dec = round(raw_dec / (raw_con + raw_dec) * 100, 0) if (raw_con + raw_dec) > 0 else 0

    st.markdown(f"""
    <div style='padding:1em; border:2px solid #607D8B; border-radius:12px; background-color:#ECEFF1;'>
      <h4 style='color:#607D8B;'>Final Tally</h4>
      <p><b>Total C_Vol:</b> {total_c_vol:.5f} m³</p>
      <p><b>Total C_Load:</b> {total_c_load:.5f}</p>
      <p><b>Total D_Vol:</b> {total_d_vol:.5f} m³</p>
      <p><b>Total D_Load:</b> {total_d_load:.5f}</p>
      <hr>
      <p><b>% Coniferous:</b> {pct_con}%</p>
      <p><b>% Deciduous:</b> {pct_dec}%</p>
    </div>""", unsafe_allow_html=True)

    col_load1, col_load2 = st.columns(2)
    with col_load1:
        st.markdown(
            f"""
            <div style="background-color: #f0f8ff; padding: 15px; border-radius: 8px; text-align: center; border: 1px solid #add8e6;">
                <strong>Total Coniferous Load</strong><br>
                <span style="font-size: 24px; color: #006400;">{total_c_load:.5f}</span>
            </div>
            """, unsafe_allow_html=True
        )
    with col_load2:
        st.markdown(
            f"""
            <div style="background-color: #fffaf0; padding: 15px; border-radius: 8px; text-align: center; border: 1px solid #ffdab9;">
                <strong>Total Deciduous Load</strong><br>
                <span style="font-size: 24px; color: #8b4513;">{total_d_load:.5f}</span>
            </div>
            """, unsafe_allow_html=True
        )

# --- Salvage form ---
if st.button("Finish (Fill Salvage Draft)", key="finish_salvage"):
    st.session_state.show_salvage_form = True

if st.session_state.show_salvage_form:
    st.subheader("Additional Information for Report Generation")

    disposition = st.text_input("Disposition", key="disposition")
    legal_loc = st.text_input("Legal Land Location", key="legal_loc")

    veg_types = [
        "Native grassland", "Tame pasture", "Cropland", "Sparsely or non-vegetated",
        "Cutblock - planted", "Natural regeneration >2m", "Treed wetland",
        "Shrubby wetland", "Grass or grass-like wetland", "Native aspen parkland",
        "Other (specify)"
    ]
    vegetation = st.multiselect("Vegetation (check all that apply):", veg_types, key="vegetation")

    other_specify_details = ""
    if "Other (specify)" in vegetation:
        other_specify_details = st.text_input("Other (specify):", key="other_specify_details")

    disposition_fma = st.text_input("Disposition # of FMA & Holder Name:", key="disposition_fma")
    no_disposition_fma = st.checkbox("None", key="no_disposition_fma")

    st.write("Coniferous/Deciduous Dispositions (Type–Number–Holder):")
    for i in range(len(st.session_state.ctlr_list)):
        col1, col2 = st.columns([1, 2])
        with col1:
            st.session_state.ctlr_list[i]["type"] = st.text_input(
                f"Type {i+1}", st.session_state.ctlr_list[i]["type"], key=f"ctlr_type_{i}"
            )
        with col2:
            st.session_state.ctlr_list[i]["number_holder"] = st.text_input(
                f"Number & Holder {i+1}", st.session_state.ctlr_list[i]["number_holder"], key=f"ctlr_number_holder_{i}"
            )

    if st.button("Add Another Disposition"):
        st.session_state.ctlr_list.append({"type": "", "number_holder": ""})
        st.rerun()

    # ── Moved here: heading right before the question ──
    st.subheader("Timber Salvage Waiver Requested?")

    salvage_waiver = st.radio(
        "",
        ["Yes", "No"],
        horizontal=True,
        key="salvage_waiver"
    )

    DEFAULT_WAIVER_JUSTIFICATION = "Timber salvage is not considered economically viable, given that the estimated volume is below 0.5 truckloads."
    if salvage_waiver == "Yes":
        if "justification" not in st.session_state or not str(st.session_state.justification).strip():
            st.session_state.justification = DEFAULT_WAIVER_JUSTIFICATION
        justification = st.text_area("Provide justification:", key="justification")

    def fill_template():
        raw_con = sum(e["dom_pct"] for e in st.session_state.results_log if e["dom_sp"] in conifers) + \
                  sum(e["sec_pct"] for e in st.session_state.results_log if e["sec_sp"] in conifers)
        raw_dec = sum(e["dom_pct"] for e in st.session_state.results_log if e["dom_sp"] in deciduous) + \
                  sum(e["sec_pct"] for e in st.session_state.results_log if e["sec_sp"] in deciduous)
        pct_con = round(raw_con / (raw_con + raw_dec) * 100, 0) if (raw_con + raw_dec) > 0 else 0

        spruce_raw = sum(e["dom_pct"] for e in st.session_state.results_log if e["dom_sp"] in {"Sw","Sb"}) + \
                     sum(e["sec_pct"] for e in st.session_state.results_log if e["sec_sp"] in {"Sw","Sb"})
        pine_raw = sum(e["dom_pct"] for e in st.session_state.results_log if e["dom_sp"] == "P") + \
                   sum(e["sec_pct"] for e in st.session_state.results_log if e["sec_sp"] == "P")
        other_con = raw_con - spruce_raw - pine_raw
        spruce_pct = pine_pct = other_con_pct = 0
        if raw_con > 0:
            spruce_pct = int(round(spruce_raw / raw_con * 100, 0))
            pine_pct = int(round(pine_raw / raw_con * 100, 0))
            other_con_pct = 100 - spruce_pct - pine_pct

        aspen_raw = sum(e["dom_pct"] for e in st.session_state.results_log if e["dom_sp"] == "Aw") + \
                    sum(e["sec_pct"] for e in st.session_state.results_log if e["sec_sp"] == "Aw")
        other_dec = raw_dec - aspen_raw
        aspen_pct = other_dec_pct = 0
        if raw_dec > 0:
            aspen_pct = int(round(aspen_raw / raw_dec * 100, 0))
            other_dec_pct = 100 - aspen_pct

        def con_class_box(label):
            if label == "D" and pct_con < 30: return "☒"
            if label == "C" and pct_con > 70: return "☒"
            if label == "CD" and 50 <= pct_con <= 70: return "☒"
            if label == "DC" and 30 <= pct_con < 50: return "☒"
            return "☐"

        doc = Document()
        # (the rest of the fill_template function remains unchanged — omitted here for brevity)
        # ... add paragraphs, formatting, volumes, etc. as in your original ...
        # At the end:
        filename = f"Timber_{disposition if disposition.strip() else 'Report'}.docx"
        tmp = NamedTemporaryFile(delete=False, suffix=".docx")
        doc.save(tmp.name)
        return tmp.name, filename

    if st.button("Done (Generate Report)"):
        out_path, filename = fill_template()
        if out_path:
            st.success("Report generated!")
            with open(out_path, "rb") as f:
                b64 = base64.b64encode(f.read()).decode()
                st.markdown(
                    f'<a href="data:application/octet-stream;base64,{b64}" download="{filename}">📥 Download report</a>',
                    unsafe_allow_html=True
                )

# --- Reset button ---
if st.button("Reset All Entries"):
    st.session_state.reset_trigger = True
    st.rerun()

# --- Sidebar Shapefile Dissolver Tool ---
# (unchanged — kept as is)
st.sidebar.header("Shapefile Dissolver Tool")
st.sidebar.markdown("Drag and drop zip files containing shapefiles to dissolve polygons individually.")
uploaded_files = st.sidebar.file_uploader("Upload .zip files", type=["zip"], accept_multiple_files=True)

temp_base_dir = Path(tempfile.mkdtemp())
output_dir = temp_base_dir / "dissolved_output"
output_dir.mkdir(parents=True, exist_ok=True)

log_file = output_dir / "processing_log.txt"
with open(log_file, "w") as log:
    log.write("Processing started\n")

# (the rest of the dissolver logic remains unchanged — processing zip files, dissolving, zipping output, etc.)
# ... omitted for brevity in this response ...

if temp_base_dir.exists():
    shutil.rmtree(temp_base_dir)
