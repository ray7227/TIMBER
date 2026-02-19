import streamlit as st
import pandas as pd

# --- Load TDA tables ---
@st.cache_data
def load_tda(region: str) -> pd.DataFrame:
    path = f"{region.upper()}_TDA.xlsx"
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

conifers = {"Sw", "Sb", "P", "Fb", "Fd", "Lt"}
deciduous = {"Aw", "Pb", "Bw"}

# --- Defaults ---
default_values = {
    "crown_density": 70,
    "avg_stand_height": 0,
    "dom_sel": species_choices[0],
    "dom_cover": 70,
    "sec_sel": "",
    "sec_cover": 30,
    "area": 1.0,
    "region": "Boreal",
    "salvage_waiver": "No",
}

# --- Session state init ---
if "results_log" not in st.session_state:
    st.session_state.results_log = []
if "show_salvage_form" not in st.session_state:
    st.session_state.show_salvage_form = False

# Inputs state
for k, v in default_values.items():
    if k not in st.session_state:
        st.session_state[k] = v

# --- Page ---
st.set_page_config(layout="wide")
st.header("🌲 TIMBER: AVI/TDA (MINIMAL TEST APP)")

# --- Core calc ---
def calculate_avi_and_volumes(is_merch, crown_density, avg_stand_height,
                              dom_species, dom_cover, sec_species, sec_cover,
                              area, region):
    avi_code = ""

    # Build AVI code
    if is_merch.lower() == "yes":
        avi_code += "m"
    if 6 <= crown_density <= 30:
        avi_code += "A"
    elif 31 <= crown_density <= 50:
        avi_code += "B"
    elif 51 <= crown_density <= 70:
        avi_code += "C"
    elif 71 <= crown_density <= 100:
        avi_code += "D"

    avi_code += str(avg_stand_height)
    avi_code += dom_species + str(dom_cover // 10)
    if dom_cover < 100 and sec_species:
        avi_code += sec_species + str(sec_cover // 10)

    def density_class(d):
        return "AB" if 6 <= d <= 50 else "CD"

    def height_bin(h):
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

    def get_structure_group(dom_sp, dom_pct, sec_sp, sec_pct):
        t_dec = (dom_pct if dom_sp in deciduous else 0) + (sec_pct if sec_sp in deciduous else 0)
        t_con = (dom_pct if dom_sp in conifers else 0) + (sec_pct if sec_sp in conifers else 0)
        if t_dec >= 70:
            return "D"
        if t_con >= 70:
            if dom_sp == "Sw":
                return "C-Sw"
            if dom_sp == "P":
                return "C-P"
            if dom_sp == "Sb":
                return "C-Sb"
            return "C-Sx"
        if t_con > 30 and t_dec < 70:
            if dom_sp == "P":
                return "MX-P"
            return "MX-Sx"
        return None

    # Defaults for safety
    c_vol = d_vol = c_load = d_load = 0.0
    c_vol_ha = d_vol_ha = None
    group = None
    total_val = 0

    try:
        df = load_tda(region)
        key = f"{height_bin(avg_stand_height)} ({density_class(crown_density)})"
        row = df[df["Height_and_Density"].astype(str).str.strip() == key]

        group = get_structure_group(dom_species, dom_cover, sec_species, sec_cover)
        valid_groups = {"D", "MX-P", "MX-Sx", "C-Sw", "C-P", "C-Sb", "C-Sx"}
        total_col = f"Total ({group})" if group in valid_groups else "Total (D)"
        total_val = row[total_col].values[0] if (not row.empty and total_col in df.columns) else 0

        if dom_cover == 100:
            c_vol_ha = total_val if dom_species in conifers else None
            d_vol_ha = total_val if dom_species in deciduous else 0
        else:
            c_pct = (dom_cover if dom_species in conifers else 0) + (sec_cover if sec_species in conifers else 0)
            d_pct = (dom_cover if dom_species in deciduous else 0) + (sec_cover if sec_species in deciduous else 0)
            c_vol_ha = round((c_pct / 100) * total_val, 1) if c_pct > 0 else None
            d_vol_ha = round((d_pct / 100) * total_val, 1) if d_pct > 0 else 0

        c_vol = round((c_vol_ha * area), 5) if c_vol_ha is not None else 0
        d_vol = round((d_vol_ha * area), 5) if d_vol_ha is not None else 0
        c_load = round((c_vol / 30), 5) if c_vol is not None else 0
        d_load = round((d_vol / 30), 5) if d_vol is not None else 0

    except Exception as e:
        st.error(f"Error reading TDA table: {e}")

    return avi_code, c_vol, d_vol, c_load, d_load, c_vol_ha, d_vol_ha, group, total_val

# --- Layout ---
col1, col2 = st.columns(2)

with col1:
    is_merch = "Yes"

    crown_density = st.slider(
        "Crown Density (%)",
        6, 100,
        int(st.session_state.crown_density),
        key="crown_density"
    )

    avg_stand_height = st.slider(
        "Average Stand Tree Height",
        0, 40,
        int(st.session_state.avg_stand_height),
        step=1,
        key="avg_stand_height"
    )

    dom_sel = st.selectbox(
        "Dominant Species",
        species_choices,
        index=species_choices.index(st.session_state.dom_sel) if st.session_state.dom_sel in species_choices else 0,
        key="dom_sel",
    )
    dom_species = dom_sel.split(" ")[0]

    # Simple covers (no session_state syncing logic at all)
    dom_cover = st.slider(
        "Dominant Cover %",
        0, 100,
        int(st.session_state.dom_cover),
        step=10,
        key="dom_cover"
    )

    sec_opts = [""] + [c for c in species_choices if c.split(" ")[0] != dom_species]
    sec_sel = st.selectbox(
        "2nd Species",
        sec_opts,
        index=sec_opts.index(st.session_state.sec_sel) if st.session_state.sec_sel in sec_opts else 0,
        key="sec_sel",
    )
    sec_species = sec_sel.split(" ")[0] if sec_sel else ""

    sec_cover = st.slider(
        "2nd Cover %",
        0, 100,
        int(st.session_state.sec_cover),
        step=10,
        key="sec_cover"
    )

    area = st.number_input(
        "Area (ha)",
        min_value=0.0,
        value=float(st.session_state.area),
        step=0.0001,
        format="%.4f",
        key="area",
    )

    region = st.selectbox(
        "Natural Region",
        ["Boreal", "Foothills"],
        index=["Boreal", "Foothills"].index(st.session_state.region) if st.session_state.region in ["Boreal", "Foothills"] else 0,
        key="region",
    )

    if st.button("Save Entry"):
        avi_code, c_vol, d_vol, c_load, d_load, *_ = calculate_avi_and_volumes(
            is_merch, crown_density, avg_stand_height,
            dom_species, dom_cover, sec_species, sec_cover,
            area, region
        )

        st.session_state.results_log.append({
            "C_Vol": c_vol,
            "C_Load": c_load,
            "D_Vol": d_vol,
            "D_Load": d_load,
            "dom_sp": dom_species,
            "dom_pct": dom_cover,
            "sec_sp": sec_species,
            "sec_pct": sec_cover,
            "is_merch": True,
            "crown_density": crown_density,
            "avg_stand_height": avg_stand_height,
            "area": area,
            "region": region,
        })
        st.success("New entry saved!")

with col2:
    avi_code, c_vol, d_vol, c_load, d_load, c_vol_ha, d_vol_ha, group, total_val = calculate_avi_and_volumes(
        is_merch, crown_density, avg_stand_height,
        dom_species, dom_cover, sec_species, sec_cover,
        area, region
    )

    st.markdown(f"""
    <div style='padding:1em; border:2px solid #4CAF50; border-radius:12px;
                background-color:#f9f9f9; color:#000;'>
        <h4 style='color:#4CAF50;'>Generated AVI Code</h4>
        <p style='font-size:24px; font-weight:bold;'>{avi_code}</p>
    </div>""", unsafe_allow_html=True)

    con_vol_ha_str = "{:.5f}".format(c_vol_ha) if c_vol_ha is not None else "N/A"
    dec_vol_ha_str = "{:.5f}".format(d_vol_ha) if (d_vol_ha is not None and d_vol_ha > 0) else "0"

    st.markdown(f"""
    <div style='padding:1em; border:2px solid #2196F3; border-radius:12px;
                background-color:#f0f8ff; color:#000;'>
        <h4 style='color:#2196F3;'>Volume per Hectare</h4>
        <p><b>Con:</b> {con_vol_ha_str} m³/ha [TDA={total_val if c_vol_ha is not None else 'N/A'}, Group={group if c_vol_ha is not None else 'N/A'}]</p>
        <p><b>Dec:</b> {dec_vol_ha_str} m³/ha [TDA={total_val if (d_vol_ha is not None and d_vol_ha > 0) else 'N/A'}, Group={group if (d_vol_ha is not None and d_vol_ha > 0) else 'N/A'}]</p>
    </div>""", unsafe_allow_html=True)

    st.markdown(f"""
    <div style='padding:1em; border:2px solid #FF9800; border-radius:12px;
                background-color:#fff8e1; color:#000;'>
        <h4 style='color:#FF9800;'>Total Volume ({area} ha)</h4>
        <p><b>Con:</b> {c_vol:.5f} m³</p>
        <p><b>Dec:</b> {d_vol:.5f} m³</p>
    </div>""", unsafe_allow_html=True)

    st.markdown(f"""
    <div style='padding:1em; border:2px solid #9C27B0; border-radius:12px;
                background-color:#f3e5f5; color:#000;'>
        <h4 style='color:#9C27B0;'>Load</h4>
        <p><b>Con:</b> {c_load:.5f}</p>
        <p><b>Dec:</b> {d_load:.5f}</p>
    </div>""", unsafe_allow_html=True)

st.divider()

# --- Show totals ---
if st.button("Finish (Show Totals)", key="finish_totals"):
    total_c_vol = sum(e["C_Vol"] for e in st.session_state.results_log if e.get("C_Vol") is not None)
    total_c_load = sum(e["C_Load"] for e in st.session_state.results_log if e.get("C_Load") is not None)
    total_d_vol = sum(e["D_Vol"] for e in st.session_state.results_log if e.get("D_Vol") is not None)
    total_d_load = sum(e["D_Load"] for e in st.session_state.results_log if e.get("D_Load") is not None)

    raw_con = sum(e["dom_pct"] for e in st.session_state.results_log if e.get("dom_sp") in conifers) + \
              sum(e["sec_pct"] for e in st.session_state.results_log if e.get("sec_sp") in conifers)
    raw_dec = sum(e["dom_pct"] for e in st.session_state.results_log if e.get("dom_sp") in deciduous) + \
              sum(e["sec_pct"] for e in st.session_state.results_log if e.get("sec_sp") in deciduous)

    pct_con = round(raw_con/(raw_con+raw_dec)*100, 0) if (raw_con+raw_dec) > 0 else 0
    pct_dec = round(raw_dec/(raw_con+raw_dec)*100, 0) if (raw_con+raw_dec) > 0 else 0

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

if st.session_state.show_salvage_form:
    st.subheader("Additional Information (Minimal)")

    salvage_waiver = st.radio(
        "Timber Salvage Waiver Requested?",
        ["Yes", "No"],
        horizontal=True,
        key="salvage_waiver"
    )

    # Boxes requested (under waiver)
    total_c_load_ui = sum(e["C_Load"] for e in st.session_state.results_log if e.get("C_Load") is not None)
    total_d_load_ui = sum(e["D_Load"] for e in st.session_state.results_log if e.get("D_Load") is not None)

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown(f"""
        <div style='padding:1em; border:2px solid #607D8B; border-radius:12px;
                    background-color:#ECEFF1; color:#000;'>
          <h4 style='color:#607D8B;'>Total Coniferous Load</h4>
          <p style='font-size:20px; font-weight:bold;'>{total_c_load_ui:.5f}</p>
        </div>""", unsafe_allow_html=True)
    with col_b:
        st.markdown(f"""
        <div style='padding:1em; border:2px solid #607D8B; border-radius:12px;
                    background-color:#ECEFF1; color:#000;'>
          <h4 style='color:#607D8B;'>Total Decidious Load</h4>
          <p style='font-size:20px; font-weight:bold;'>{total_d_load_ui:.5f}</p>
        </div>""", unsafe_allow_html=True)

st.divider()

if st.button("Reset All Entries"):
    st.session_state.results_log = []
    st.session_state.show_salvage_form = False
    st.rerun()
