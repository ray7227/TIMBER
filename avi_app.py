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
    path = f"{region.upper()}_TDA.xlsx"  # Looks for file in same directory as script
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
    "salvage_waiver": "No",
    "justification": ""
}

# --- Session state initialization ---
keys_to_init = {
    'results_log': [],
    'current_entry_index': -1,
    'edit_mode': False,
    'show_salvage_form': False,
    'reset_trigger': False,
    'dom_cover': default_values["dom_cover"],
    'sec_cover': default_values["sec_cover"],
    'dom_species': species_choices[0].split(" ")[0],
    'sec_species': "",
    'avg_stand_height': default_values["avg_stand_height"],
    'is_merch': default_values["is_merch"],
    'crown_density': default_values["crown_density"],
    'dom_sel': default_values["dom_sel"],
    'sec_sel': default_values["sec_sel"],
    'area': default_values["area"],
    'region': default_values["region"],
    'ctlr_list': [{"type": "", "number_holder": ""}],
}

for k, v in keys_to_init.items():
    if k not in st.session_state:
        st.session_state[k] = v

# --- Reset widget defaults if triggered ---
if st.session_state.reset_trigger:
    for k, v in default_values.items():
        if k in st.session_state:
            st.session_state[k] = v
    st.session_state.results_log = []
    st.session_state.current_entry_index = -1
    st.session_state.edit_mode = False
    st.session_state.show_salvage_form = False
    st.session_state.ctlr_list = [{"type": "", "number_holder": ""}]
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

# Initialize globals
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
    st.write("")  # Placeholder
with col_nav2:
    if st.button("Save Entry"):
        dom_species = st.session_state.dom_sel.split(" ")[0] if st.session_state.dom_sel else ""
        sec_species = st.session_state.sec_sel.split(" ")[0] if st.session_state.sec_sel else ""
        calculate_avi_and_volumes(
            "Yes",
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
            "is_merch": True,
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

# --- Inputs & AVI calculation ---
col1, col2 = st.columns(2)
with col1:
    # Load saved entry if editing
    if st.session_state.edit_mode and st.session_state.current_entry_index >= 0 and st.session_state.results_log:
        entry = st.session_state.results_log[st.session_state.current_entry_index]
        st.session_state.is_merch = "Yes" if entry.get("is_merch", True) else "No"
        st.session_state.crown_density = entry.get("crown_density", 70)
        st.session_state.avg_stand_height = entry.get("avg_stand_height", 0)
        st.session_state.dom_sel = f"{entry['dom_sp']} ({species_names[entry['dom_sp']]})" if entry.get('dom_sp') else species_choices[0]
        st.session_state.dom_cover = entry.get("dom_pct", 70)
        st.session_state.sec_cover = entry.get("sec_pct", 30)
        st.session_state.sec_sel = f"{entry['sec_sp']} ({species_names[entry['sec_sp']]})" if entry.get('sec_sp') else ""
        st.session_state.area = entry.get("area", 1.0)
        st.session_state.region = entry.get("region", "Boreal")

    crown_density = st.slider(
        "Crown Density (%)",
        6, 100,
        st.session_state.crown_density,
        key="crown_density"
    )

    avg_stand_height = st.slider(
        "Average Stand Tree Height",
        0, 40,
        st.session_state.avg_stand_height,
        step=1,
        key="avg_stand_height"
    )

    dom_sel = st.selectbox(
        "Dominant Species",
        species_choices,
        key="dom_sel"
    )
    dom_species = dom_sel.split(" ")[0]

    # Linked cover sliders
    col_dom, col_sec = st.columns(2)
    with col_dom:
        st.slider(
            "Dominant Cover %",
            0, 100,
            value=st.session_state.dom_cover,
            step=10,
            key="dom_cover_widget",
            on_change=lambda: st.session_state.update({
                'dom_cover': st.session_state.dom_cover_widget,
                'sec_cover': 100 - st.session_state.dom_cover_widget
            })
        )
        st.session_state.dom_cover = st.session_state.dom_cover_widget

    with col_sec:
        st.slider(
            "2nd Cover %",
            0, 100,
            value=st.session_state.sec_cover,
            step=10,
            key="sec_cover_widget",
            on_change=lambda: st.session_state.update({
                'sec_cover': st.session_state.sec_cover_widget,
                'dom_cover': 100 - st.session_state.sec_cover_widget
            })
        )
        st.session_state.sec_cover = st.session_state.sec_cover_widget

    # Final safety sync
    if st.session_state.dom_cover + st.session_state.sec_cover != 100:
        st.session_state.sec_cover = 100 - st.session_state.dom_cover
        st.rerun()

    sec_opts = [""] + [c for c in species_choices if c.split(" ")[0] != dom_species]
    sec_sel = st.selectbox(
        "2nd Species",
        sec_opts,
        key="sec_sel"
    )
    sec_species = sec_sel.split(" ")[0] if sec_sel else ""

    area = st.number_input(
        "Area (ha)",
        min_value=0.0,
        value=st.session_state.area,
        step=0.0001,
        format="%.4f",
        key="area"
    )

    region = st.selectbox(
        "Natural Region",
        ["Boreal", "Foothills"],
        key="region"
    )

calculate_avi_and_volumes("Yes", crown_density, avg_stand_height, dom_species,
                          st.session_state.dom_cover, sec_species, st.session_state.sec_cover,
                          area, region)

# --- Styled outputs ---
with col2:
    st.markdown(f"""
    <div style='padding:1em; border:2px solid #4CAF50; border-radius:12px; background-color:#f9f9f9; color:#000;'>
        <h4 style='color:#4CAF50;'>Generated AVI Code</h4>
        <p style='font-size:24px; font-weight:bold;'>{avi_code}</p>
    </div>""", unsafe_allow_html=True)

    con_vol_ha_str = "{:.5f}".format(c_vol_ha) if c_vol_ha is not None else "N/A"
    dec_vol_ha_str = "{:.5f}".format(d_vol_ha) if d_vol_ha > 0 else "0"
    st.markdown(f"""
    <div style='padding:1em; border:2px solid #2196F3; border-radius:12px; background-color:#f0f8ff; color:#000;'>
        <h4 style='color:#2196F3;'>Volume per Hectare</h4>
        <p><b>Con:</b> {con_vol_ha_str} m³/ha</p>
        <p><b>Dec:</b> {dec_vol_ha_str} m³/ha</p>
    </div>""", unsafe_allow_html=True)

    st.markdown(f"""
    <div style='padding:1em; border:2px solid #FF9800; border-radius:12px; background-color:#fff8e1; color:#000;'>
        <h4 style='color:#FF9800;'>Total Volume ({area} ha)</h4>
        <p><b>Con:</b> {c_vol:.5f} m³</p>
        <p><b>Dec:</b> {d_vol:.5f} m³</p>
    </div>""", unsafe_allow_html=True)

    st.markdown(f"""
    <div style='padding:1em; border:2px solid #9C27B0; border-radius:12px; background-color:#f3e5f5; color:#000;'>
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

    st.subheader("P3 Map Search Converter")
    lsd_input = st.text_input("", placeholder="NE-20-48-11-W5 SE-35-67-7-W6", label_visibility="collapsed")
    if lsd_input:
        lsds = [x.strip() for x in lsd_input.replace("\n", " ").split() if x.strip()]
        results = [convert_lsd_to_p3(x) for x in lsds if convert_lsd_to_p3(x)]
        if results:
            st.text("\n".join(results))

# --- Show totals ---
if st.button("Finish (Show Totals)", key="finish_totals"):
    total_c_vol = sum(e["C_Vol"] for e in st.session_state.results_log if e.get("C_Vol") is not None)
    total_c_load = sum(e["C_Load"] for e in st.session_state.results_log if e.get("C_Load") is not None)
    total_d_vol = sum(e["D_Vol"] for e in st.session_state.results_log if e.get("D_Vol") is not None)
    total_d_load = sum(e["D_Load"] for e in st.session_state.results_log if e.get("D_Load") is not None)

    raw_con = sum(e["dom_pct"] for e in st.session_state.results_log if e["dom_sp"] in conifers) + \
              sum(e["sec_pct"] for e in st.session_state.results_log if e["sec_sp"] in conifers)
    raw_dec = sum(e["dom_pct"] for e in st.session_state.results_log if e["dom_sp"] in deciduous) + \
              sum(e["sec_pct"] for e in st.session_state.results_log if e["sec_sp"] in deciduous)

    pct_con = round(raw_con / (raw_con + raw_dec) * 100, 0) if (raw_con + raw_dec) > 0 else 0
    pct_dec = 100 - pct_con

    st.markdown(f"""
    <div style='padding:1em; border:2px solid #607D8B; border-radius:12px; background-color:#ECEFF1; color:#000;'>
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
            """,
            unsafe_allow_html=True
        )
    with col_load2:
        st.markdown(
            f"""
            <div style="background-color: #fffaf0; padding: 15px; border-radius: 8px; text-align: center; border: 1px solid #ffdab9;">
                <strong>Total Deciduous Load</strong><br>
                <span style="font-size: 24px; color: #8b4513;">{total_d_load:.5f}</span>
            </div>
            """,
            unsafe_allow_html=True
        )

# --- Salvage form trigger ---
if st.button("Finish (Fill Salvage Draft)", key="finish_salvage"):
    st.session_state.show_salvage_form = True

# --- Salvage form & Word export ---
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
                f"Type {i+1}",
                st.session_state.ctlr_list[i]["type"],
                key=f"ctlr_type_{i}"
            )
        with col2:
            st.session_state.ctlr_list[i]["number_holder"] = st.text_input(
                f"Number & Holder {i+1}",
                st.session_state.ctlr_list[i]["number_holder"],
                key=f"ctlr_number_holder_{i}"
            )

    if st.button("Add Another Disposition"):
        st.session_state.ctlr_list.append({"type": "", "number_holder": ""})
        st.rerun()

    # The requested heading placement
    st.subheader("Timber Salvage Waiver Requested?")

    salvage_waiver = st.radio(
        "",
        ["Yes", "No"],
        horizontal=True,
        key="salvage_waiver"
    )

    if salvage_waiver == "Yes":
        if "justification" not in st.session_state or not str(st.session_state.justification).strip():
            st.session_state.justification = "Timber salvage is not considered economically viable, given that the estimated volume is below 0.5 truckloads."
        st.text_area("Provide justification:", key="justification")

    # fill_template function and report generation (add your original implementation here if needed)

# --- Reset ---
if st.button("Reset All Entries"):
    st.session_state.reset_trigger = True
    st.rerun()

# --- Shapefile Dissolver in Sidebar ---
st.sidebar.header("Shapefile Dissolver Tool")
st.sidebar.markdown("Drag and drop zip files containing shapefiles to dissolve polygons individually.")

uploaded_files = st.sidebar.file_uploader(
    "Upload .zip files",
    type=["zip"],
    accept_multiple_files=True,
    help="Select or drag and drop .zip files containing shapefiles."
)

temp_base_dir = Path(tempfile.mkdtemp())
output_dir = temp_base_dir / "dissolved_output"
output_dir.mkdir(parents=True, exist_ok=True)

log_file = output_dir / "processing_log.txt"
with open(log_file, "w") as log:
    log.write("Processing started\n")

if uploaded_files:
    for uploaded_file in uploaded_files:
        zip_path = temp_base_dir / uploaded_file.name
        with open(zip_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        with open(log_file, "a") as log:
            log.write(f"\nProcessing {zip_path.name}...\n")
            st.sidebar.write(f"Processing {zip_path.name}...")

        zip_subdir = output_dir / zip_path.stem
        zip_subdir.mkdir(exist_ok=True)

        temp_dir = temp_base_dir / f"temp_{zip_path.stem}"
        temp_dir.mkdir(exist_ok=True)

        try:
            with zipfile.ZipFile(zip_path, "r") as z:
                z.extractall(temp_dir)

            shapefiles = list(temp_dir.glob("*.shp"))
            if not shapefiles:
                st.sidebar.warning(f"No shapefiles found in {zip_path.name}")
                continue

            gdf = gpd.read_file(shapefiles[0])
            gdf["geometry"] = gdf.geometry.buffer(0)

            dissolved_gdf = gdf.dissolve()

            out_file = zip_subdir / f"{zip_path.stem}_dissolved.shp"
            dissolved_gdf.to_file(out_file)
            st.sidebar.success(f"Saved dissolved: {out_file.name}")

        except Exception as e:
            st.sidebar.error(f"Error processing {zip_path.name}: {e}")

        finally:
            if temp_dir.exists():
                shutil.rmtree(temp_dir)

    output_zip_path = temp_base_dir / "dissolved_shapefiles.zip"
    with zipfile.ZipFile(output_zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(output_dir):
            for file in files:
                zipf.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), output_dir))

    with open(output_zip_path, "rb") as f:
        st.sidebar.download_button(
            label="Download All Dissolved Shapefiles (Zip)",
            data=f,
            file_name="dissolved_shapefiles.zip",
            mime="application/zip"
        )

    with open(log_file, "rb") as f:
        st.sidebar.download_button(
            label="Download Processing Log",
            data=f,
            file_name="processing_log.txt",
            mime="text/plain"
        )

    st.sidebar.success("All zip files processed.")

# Cleanup
if temp_base_dir.exists():
    shutil.rmtree(temp_base_dir)
