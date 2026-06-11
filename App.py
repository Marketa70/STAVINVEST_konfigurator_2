import streamlit as st
import pandas as pd
import math
import io
import copy
import random
import os
from collections import defaultdict
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from openpyxl.drawing.image import Image as xlImage
from openpyxl.utils import get_column_letter

# --- NASTAVENÍ STRÁNKY ---
st.set_page_config(page_title="Konfigurátor Stavinvest", page_icon="✂️", layout="wide")

# ==========================================
# 🔒 PŘIHLAŠOVACÍ ÚDAJE
# ==========================================
UZIVATELE = {
    "admin@stavinvest.cz": "HlavniKlempir!",
    "test1@stavinvest.cz": "PlechovaStrecha1",
    "test2@stavinvest.cz": "Okapnice2026",
    "test3@stavinvest.cz": "TitanzinekRulez",
    "test4@stavinvest.cz": "FalcujemeDobre",
    "test5@stavinvest.cz": "OhybackaStroj",
    "test6@stavinvest.cz": "SvitekPlechu99",
    "test7@stavinvest.cz": "KlempiroveCZ",
    "test8@stavinvest.cz": "ZavetrnaLista#",
    "test9@stavinvest.cz": "StavinvestPro",
    "test10@stavinvest.cz": "NuzkyNaPlech123"
}

if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.current_user = ""

if not st.session_state.logged_in:
    st.markdown("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        st.markdown("<h2 style='text-align: center;'>🔒 Přihlášení do systému</h2>", unsafe_allow_html=True)
        with st.form("login_form"):
            email = st.text_input("E-mail")
            password = st.text_input("Heslo", type="password")
            submit = st.form_submit_button("Přihlásit se", use_container_width=True)

            if submit:
                if email in UZIVATELE and UZIVATELE[email] == password:
                    st.session_state.logged_in = True
                    st.session_state.current_user = email
                    st.rerun()
                else:
                    st.error("Chybný e-mail nebo heslo!")
    st.stop()

# --- BOČNÍ PANEL (Odhlášení) ---
st.sidebar.write(f"👤 Přihlášen(a): **{st.session_state.current_user}**")
if st.sidebar.button("🚪 Odhlásit se", use_container_width=True):
    st.session_state.logged_in = False
    st.session_state.current_user = ""
    st.rerun()

# ==========================================
# HLAVNÍ APLIKACE
# ==========================================
st.title("✂️ Konfigurátor Stavinvest")
st.info("💡 **Nová funkce:** Rozvinutou šíři (RŠ) nyní zadáváte ručně v milimetrech pro každý prvek zvlášť.")

# ==========================================
# MODULOVÝ PRUHOVÝ ALGORITMUS
# ==========================================
def pack_module_strips(items, coil_w, max_l, allow_rotation=False):
    best_modules = None
    best_len = float('inf')
    for iteration in range(200):
        test_items = copy.deepcopy(items)
        if iteration == 0: test_items.sort(key=lambda x: x['L'], reverse=True)
        elif iteration == 1: test_items.sort(key=lambda x: x['rš'], reverse=True)
        else: random.shuffle(test_items)
        for it in test_items:
            can_std = (it['L'] <= max_l and it['rš'] <= coil_w)
            can_rot = (allow_rotation and it['rš'] <= max_l and it['L'] <= coil_w)
            if iteration < 2:
                if can_std: it['dx'], it['dy'], it['rotated'] = it['L'], it['rš'], False
                elif can_rot: it['dx'], it['dy'], it['rotated'] = it['rš'], it['L'], True
                else: it['dx'], it['dy'], it['rotated'] = it['L'], it['rš'], False 
            else:
                if can_std and can_rot:
                    if random.random() > 0.5: it['dx'], it['dy'], it['rotated'] = it['rš'], it['L'], True
                    else: it['dx'], it['dy'], it['rotated'] = it['L'], it['rš'], False
                elif can_rot: it['dx'], it['dy'], it['rotated'] = it['rš'], it['L'], True
                else: it['dx'], it['dy'], it['rotated'] = it['L'], it['rš'], False
        groups = defaultdict(list)
        for it in test_items: groups[it['dy']].append(it)
        strips = []
        for dy, group_items in groups.items():
            if iteration % 2 == 0: group_items.sort(key=lambda x: x['dx'], reverse=True)
            else: random.shuffle(group_items)
            current_strips = []
            for it in group_items:
                placed = False
                for s in current_strips:
                    if s['l'] + it['dx'] <= max_l:
                        it['x'] = s['l']; s['items'].append(it); s['l'] += it['dx']; placed = True; break
                if not placed:
                    it['x'] = 0; current_strips.append({'w': dy, 'l': it['dx'], 'items': [it]})
            strips.extend(current_strips)
        strips.sort(key=lambda s: s['l'], reverse=True)
        modules = []
        for s in strips:
            placed = False
            for m in modules:
                if m['used_w'] + s['w'] <= coil_w:
                    s['y'] = m['used_w']
                    for it in s['items']: it['y'] = s['y']
                    m['strips'].append(s); m['used_w'] += s['w']; m['l'] = max(m['l'], s['l']); placed = True; break
            if not placed:
                s['y'] = 0
                for it in s['items']: it['y'] = 0
                modules.append({'used_w': s['w'], 'l': s['l'], 'strips': [s]})
        tot_len = sum(m['l'] for m in modules)
        if tot_len < best_len: best_len = tot_len; best_modules = modules
    formatted_bins = []
    if best_modules:
        for m in best_modules:
            placed = []
            for s in m['strips']:
                for it in s['items']: it['draw_w'] = it['dx']; it['draw_h'] = it['dy']; placed.append(it)
            formatted_bins.append({'w_coil': coil_w, 'odvinuto_mm': m['l'], 'placed': placed})
    return formatted_bins

# ==========================================
# DATA (Zjednodušeno pro příklad)
# ==========================================
if 'config' not in st.session_state: st.session_state.config = {"cena_ohyb": 12.0, "max_delka": 4000, "presah": 40}
if 'materialy_df' not in st.session_state: st.session_state.materialy_df = pd.DataFrame([{"Materiál": "svitek POZINK 0,55x1000mm", "Šířka (mm)": 1000, "Cena/m2": 200.0, "Max délka tabule (mm)": 50000}])
if 'prvky_df' not in st.session_state: st.session_state.prvky_df = pd.DataFrame([{"Typ prvku": "Závětrná lišta spodní", "Ohyby": 6}])
if 'zakazka' not in st.session_state: st.session_state.zakazka = []

mat_dict = {r["Materiál"]: r for _, r in st.session_state.materialy_df.iterrows()}
prv_dict = {r["Typ prvku"]: r for _, r in st.session_state.prvky_df.iterrows()}

def fmt_cz(value):
    parts = f"{value:,.2f}".split('.')
    return f"{parts[0].replace(',', ' ')},{parts[1]}"

tab_kalk, tab_nakres, tab_data, tab_nastaveni = st.tabs(["🧮 Kalkulátor", "📐 Nákres", "⚙️ Data", "🔧 Nastavení"])

with tab_kalk:
    st.header("1. Obecné údaje")
    col_t1, col_t2 = st.columns(2)
    with col_t1: v_mat = st.selectbox("Materiál", list(mat_dict.keys()))
    
    col_in, col_res = st.columns([1, 2])
    with col_in:
        st.header("2. Přidat položku")
        with st.form("pridat_polozku_form", clear_on_submit=True):
            v_prvek = st.selectbox("Prvek", list(prv_dict.keys()))
            v_rs = st.number_input("Rozvinutá šíře - RŠ (mm)", value=250)
            v_m = st.number_input("Délka (m)", value=2.5)
            v_ks = st.number_input("Kusů", value=1)
            if st.form_submit_button("➕ Přidat"):
                st.session_state.zakazka.append({"Prvek": v_prvek, "RŠ (mm)": v_rs, "Ohyby": 2, "Metrů": v_m, "Kusů": v_ks})
                st.rerun()

    with col_res:
        st.header("Výpočet a Optimalizace")
        if st.session_state.zakazka:
            if st.button("🚀 SPOČÍTAT ZAKÁZKU", type="primary"):
                # --- VÝPOČET CELKOVÉ PLOCHY ---
                celkova_plocha_m2 = 0
                items = []
                for p in st.session_state.zakazka:
                    plocha_ks = (p["RŠ (mm)"] / 1000.0) * p["Metrů"]
                    celkova_plocha_m2 += plocha_ks * p["Kusů"]
                    # (zde by pokračoval zbytek vaší logiky pro pack_module_strips)
                
                st.session_state.celkova_plocha_m2 = celkova_plocha_m2
                st.session_state.calc_done = True
                st.success("Výpočet dokončen!")

        if st.session_state.get('calc_done', False):
            st.write(f"### 📏 Celková vypočtená plocha prvků: {st.session_state.get('celkova_plocha_m2', 0):.2f} m²")
