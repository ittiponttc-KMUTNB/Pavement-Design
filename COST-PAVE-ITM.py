"""
ระบบวิเคราะห์ค่าก่อสร้างโครงสร้างชั้นทาง
Version 6.0 - Refactored
พัฒนาโดย: รศ.ดร.อิทธิพล มีผล — KMUTNB
- render_layer_editor() และ render_joint_editor() ใช้ st.data_editor แทน number_input loop
- ตัด Tab รูปภาพออก
- รวม get_default_*_layers() เป็น get_default_layers(ptype)
- รวม get_price_from_library() เป็นจุดเดียว
- ตรวจ syntax ด้วย ast.parse() ก่อน deploy
"""

import ast
import io
import json
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st

try:
    import plotly.graph_objects as go
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False

try:
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

# ── ตรวจ syntax ──────────────────────────────────────────────────────────────
_src = open(__file__).read()
try:
    ast.parse(_src)
except SyntaxError as _e:
    raise SyntaxError(f"[ast.parse] Syntax error in {__file__}: {_e}") from _e

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="วิเคราะห์ค่าก่อสร้างโครงสร้างชั้นทาง",
    page_icon="🛣️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Thai:wght@300;400;500;600&family=Space+Grotesk:wght@400;500;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans Thai', sans-serif;
}

/* ── Header ── */
.main-header {
    background: linear-gradient(135deg, #0f2942 0%, #1a4a7a 60%, #0d7377 100%);
    border-radius: 16px;
    padding: 2rem 2.5rem;
    margin-bottom: 1.5rem;
    position: relative;
    overflow: hidden;
}
.main-header::before {
    content: '';
    position: absolute;
    top: -40px; right: -40px;
    width: 200px; height: 200px;
    border-radius: 50%;
    background: rgba(255,255,255,0.04);
}
.main-header h1 {
    font-family: 'Space Grotesk', sans-serif;
    color: #ffffff;
    font-size: 1.75rem;
    font-weight: 700;
    margin: 0;
    letter-spacing: -0.5px;
}
.main-header p {
    color: rgba(255,255,255,0.65);
    font-size: 0.9rem;
    margin: 0.35rem 0 0 0;
}

/* ── Metric cards ── */
.metric-card {
    background: #ffffff;
    border: 1px solid #e8edf2;
    border-radius: 12px;
    padding: 1.1rem 1.3rem;
    box-shadow: 0 2px 8px rgba(0,0,0,0.05);
}
.metric-card .label {
    font-size: 0.78rem;
    color: #6b7a8d;
    font-weight: 500;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-bottom: 0.3rem;
}
.metric-card .value {
    font-family: 'Space Grotesk', sans-serif;
    font-size: 1.5rem;
    font-weight: 700;
    color: #0f2942;
}
.metric-card .unit {
    font-size: 0.8rem;
    color: #8a95a3;
    margin-left: 4px;
}

/* ── Section card ── */
.section-card {
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    border-left: 4px solid #1a4a7a;
    border-radius: 10px;
    padding: 1rem 1.2rem;
    margin-bottom: 1rem;
}

/* ── Selectbox green ── */
div[data-baseweb="select"] > div {
    background-color: #f0faf4 !important;
    border-color: #52b788 !important;
    border-radius: 8px !important;
}
div[data-baseweb="select"] span { color: #1b5e20 !important; font-weight: 500 !important; }
div[data-baseweb="menu"] { background-color: #f1f8e9 !important; }
div[data-baseweb="menu"] li { background-color: #f1f8e9 !important; color: #1b5e20 !important; }
div[data-baseweb="menu"] li:hover { background-color: #c8e6c9 !important; }

/* ── Checkbox accent ── */
label[data-baseweb="checkbox"] span { color: #0f2942 !important; }

/* ── data_editor clean ── */
[data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }

/* ── Tab styling ── */
button[data-baseweb="tab"] {
    font-family: 'Space Grotesk', sans-serif !important;
    font-weight: 600 !important;
    font-size: 0.9rem !important;
}

/* ── Footer ── */
.footer {
    text-align: center;
    color: #94a3b8;
    font-size: 0.8rem;
    padding: 1.5rem 0 0.5rem;
    border-top: 1px solid #e2e8f0;
    margin-top: 2rem;
}
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# DEFAULT PRICE TABLES
# ═══════════════════════════════════════════════════════════════════════════════

DEFAULT_AC_PRICES: dict = {
    'PMA Wearing Course':  {2.5:170, 3:203, 4:268, 5:333, 6:406, 7:471, 8:536, 9:601, 10:667},
    'AC Wearing Course':   {2.5:128, 3:152, 4:202, 5:250, 6:306, 7:355, 8:403, 9:452, 10:502},
    'AC Binder Course':    {2.5:129, 3:154, 4:202, 5:251, 6:308, 7:356, 8:405, 9:454, 10:503},
    'AC Base Course':      {2.5:129, 3:154, 4:202, 5:251, 6:308, 7:356, 8:405, 9:454, 10:503},
}

DEFAULT_CONCRETE_PRICES: dict = {
    'JPCP': {25:928,  28:1000, 30:1050, 32:1095, 35:1167},
    'JRCP': {25:924,  28:1002, 30:1050, 32:1106, 35:1184},
    'CRCP': {25:1245, 28:1358, 30:1430, 32:1509, 35:1622},
}

DEFAULT_BASE_PRICES: dict = {
    'Cement Treated Base (UCS 40 ksc)':                    1096,
    'Cement Modified Crushed Rock Base (UCS 24.5 ksc)':    864,
    'Crushed Rock Base Course':                             583,
    'Soil Cement Subbase (UCS 7 ksc)':                     854,
    'Soil Aggregate Subbase':                              375,
    'Selected Material A':                                 375,
    'Embankment':                                          200,
    'Sand Embankment':                                     220,
    'Prime Coat':                                          37.47,
    'Non Woven Geotextile':                                78,
    'Wire Mesh':                                           100,
    'Tack Coat':                                           20,
}

BASE_MATERIAL_LIST = [
    'Cement Treated Base (UCS 40 ksc)',
    'Cement Modified Crushed Rock Base (UCS 24.5 ksc)',
    'Crushed Rock Base Course',
    'Soil Cement Subbase (UCS 7 ksc)',
    'Soil Aggregate Subbase',
    'Selected Material A',
    'Embankment',
    'Sand Embankment',
]

WEARING_OPTIONS = ['AC Wearing Course', 'PMA Wearing Course']

# ═══════════════════════════════════════════════════════════════════════════════
# PRICE LIBRARY HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def get_price_library() -> dict:
    """ดึง price_library จาก session_state หรือใช้ default"""
    if 'price_library' in st.session_state:
        return st.session_state['price_library']
    return {
        'ac_prices':       DEFAULT_AC_PRICES,
        'concrete_prices': DEFAULT_CONCRETE_PRICES,
        'base_prices':     DEFAULT_BASE_PRICES,
    }


def lookup_price(name: str, thickness: float, ptype: str = 'AC') -> float:
    """ดึงราคา (บาท/ตร.ม.) จาก library ตามชื่อวัสดุและความหนา
    สำหรับ base materials คืนราคา บาท/ลบ.ม. (caller แปลงเองด้วย ×t/100)
    """
    lib = get_price_library()
    n = name.lower()

    def _nearest(d: dict, t: float) -> float:
        if not d:
            return 0.0
        if t in d:
            return float(d[t])
        return float(d[min(d.keys(), key=lambda x: abs(x - t))])

    # AC surface
    if 'pma' in n:
        return _nearest(lib['ac_prices'].get('PMA Wearing Course', {}), thickness)
    if 'wearing' in n:
        return _nearest(lib['ac_prices'].get('AC Wearing Course', {}), thickness)
    if 'binder' in n:
        return _nearest(lib['ac_prices'].get('AC Binder Course', {}), thickness)
    if ('asphalt' in n and 'base' in n) or 'ac base' in n or 'interlayer' in n:
        return _nearest(lib['ac_prices'].get('AC Base Course', {}), thickness)

    # Concrete
    for ct in ('jpcp', 'jrcp', 'crcp'):
        if ct in n:
            return _nearest(lib['concrete_prices'].get(ct.upper(), {}), thickness)
    # fallback by ptype
    if ptype in ('JPCP', 'JRCP', 'CRCP') and ('concrete' in n or 'ksc' in n or '350' in n):
        return _nearest(lib['concrete_prices'].get(ptype, {}), thickness)

    # Base materials — คืนราคา/ลบ.ม.
    if 'tack' in n:
        return float(lib['base_prices'].get('Tack Coat', 20))
    if 'prime' in n:
        return float(lib['base_prices'].get('Prime Coat', 37.47))
    if 'geotextile' in n:
        return float(lib['base_prices'].get('Non Woven Geotextile', 78))
    if 'wire' in n:
        return float(lib['base_prices'].get('Wire Mesh', 100))

    for key in BASE_MATERIAL_LIST:
        if key.lower() in n or n in key.lower():
            return float(lib['base_prices'].get(key, 0))

    return 0.0


# ═══════════════════════════════════════════════════════════════════════════════
# DEFAULT LAYERS / JOINTS
# ═══════════════════════════════════════════════════════════════════════════════

def get_default_layers(ptype: str, area_per_km: float = 22000) -> list:
    """คืน default layers list ตาม ptype: AC | JPCP | JRCP | CRCP"""
    lib = get_price_library()
    ac = lib['ac_prices']
    cp = lib['concrete_prices']
    bp = lib['base_prices']

    def ap(mat, t):  # ac price
        return ac.get(mat, {}).get(t, list(ac.get(mat, {0: 0}).values())[0] if ac.get(mat) else 0)

    if ptype == 'AC':
        return [
            {'name': 'AC Wearing Course',  'thickness': 7,  'unit': 'cm', 'quantity': area_per_km, 'qty_unit': 'sq.m', 'unit_cost': ap('AC Wearing Course', 7)},
            {'name': 'AC Binder Course',   'thickness': 7,  'unit': 'cm', 'quantity': area_per_km, 'qty_unit': 'sq.m', 'unit_cost': ap('AC Binder Course', 7)},
            {'name': 'AC Base Course',     'thickness': 10, 'unit': 'cm', 'quantity': area_per_km, 'qty_unit': 'sq.m', 'unit_cost': ap('AC Base Course', 10)},
            {'name': 'Tack Coat',          'thickness': 1,  'unit': 'Layer', 'quantity': area_per_km*2, 'qty_unit': 'sq.m', 'unit_cost': float(bp.get('Tack Coat', 20))},
        ]
    if ptype == 'JPCP':
        return [
            {'name': 'Concrete Slab (JPCP)', 'thickness': 28, 'unit': 'cm', 'quantity': area_per_km, 'qty_unit': 'sq.m', 'unit_cost': cp.get('JPCP', {}).get(28, 1000)},
            {'name': 'Non Woven Geotextile', 'thickness': 1,  'unit': 'ชั้น', 'quantity': area_per_km, 'qty_unit': 'sq.m', 'unit_cost': float(bp.get('Non Woven Geotextile', 78))},
        ]
    if ptype == 'JRCP':
        return [
            {'name': 'Concrete Slab (JRCP)', 'thickness': 28, 'unit': 'cm', 'quantity': area_per_km, 'qty_unit': 'sq.m', 'unit_cost': cp.get('JRCP', {}).get(28, 1002)},
            {'name': 'Wire Mesh',            'thickness': 1,  'unit': 'ชั้น', 'quantity': area_per_km, 'qty_unit': 'sq.m', 'unit_cost': float(bp.get('Wire Mesh', 100))},
            {'name': 'Non Woven Geotextile', 'thickness': 1,  'unit': 'ชั้น', 'quantity': area_per_km, 'qty_unit': 'sq.m', 'unit_cost': float(bp.get('Non Woven Geotextile', 78))},
        ]
    if ptype == 'CRCP':
        return [
            {'name': 'Concrete Slab (CRCP)', 'thickness': 25, 'unit': 'cm', 'quantity': area_per_km, 'qty_unit': 'sq.m', 'unit_cost': cp.get('CRCP', {}).get(25, 1245)},
            {'name': 'Wire Mesh',            'thickness': 1,  'unit': 'ชั้น', 'quantity': area_per_km, 'qty_unit': 'sq.m', 'unit_cost': float(bp.get('Wire Mesh', 100))},
            {'name': 'Non Woven Geotextile', 'thickness': 1,  'unit': 'ชั้น', 'quantity': area_per_km, 'qty_unit': 'sq.m', 'unit_cost': float(bp.get('Non Woven Geotextile', 78))},
        ]
    return []


def get_default_base_layers(ptype: str, area_per_km: float = 22000) -> list:
    """คืน default base layers list"""
    lib = get_price_library()
    bp = lib['base_prices']

    def b(mat, t):
        return float(bp.get(mat, 0)) * t / 100

    if ptype == 'AC':
        return [
            {'name': 'Crushed Rock Base Course', 'thickness': 20, 'unit': 'cm', 'quantity': area_per_km, 'qty_unit': 'sq.m',
             'unit_cost': b('Crushed Rock Base Course', 20), 'cost_cum': float(bp.get('Crushed Rock Base Course', 583))},
            {'name': 'Soil Aggregate Subbase',   'thickness': 30, 'unit': 'cm', 'quantity': area_per_km, 'qty_unit': 'sq.m',
             'unit_cost': b('Soil Aggregate Subbase', 30), 'cost_cum': float(bp.get('Soil Aggregate Subbase', 375))},
            {'name': 'Sand Embankment',           'thickness': 40, 'unit': 'cm', 'quantity': area_per_km, 'qty_unit': 'sq.m',
             'unit_cost': b('Sand Embankment', 40), 'cost_cum': float(bp.get('Sand Embankment', 220))},
        ]
    # Concrete types
    return [
        {'name': 'Soil Cement Subbase (UCS 7 ksc)', 'thickness': 20, 'unit': 'cm', 'quantity': area_per_km, 'qty_unit': 'sq.m',
         'unit_cost': b('Soil Cement Subbase (UCS 7 ksc)', 20), 'cost_cum': float(bp.get('Soil Cement Subbase (UCS 7 ksc)', 854))},
        {'name': 'Sand Embankment',                  'thickness': 50, 'unit': 'cm', 'quantity': area_per_km, 'qty_unit': 'sq.m',
         'unit_cost': b('Sand Embankment', 50), 'cost_cum': float(bp.get('Sand Embankment', 220))},
    ]


def get_default_joints(ptype: str, area_per_km: float = 22000, road_length: float = 1.0) -> list:
    """คืน default joints list ตาม ptype"""
    width_m = area_per_km / 1000
    if ptype == 'JPCP':
        trans_qty = (road_length * 1000 / 4) * width_m
        long_qty  = road_length * 1000
        return [
            {'name': 'Transverse Joint @4m', 'quantity': trans_qty, 'qty_unit': 'm', 'unit_cost': 430},
            {'name': 'Longitudinal Joint',   'quantity': long_qty,  'qty_unit': 'm', 'unit_cost': 120},
        ]
    if ptype == 'JRCP':
        trans_qty = (road_length * 1000 / 10) * width_m
        long_qty  = road_length * 1000
        return [
            {'name': 'Transverse Joint @10m', 'quantity': trans_qty, 'qty_unit': 'm', 'unit_cost': 430},
            {'name': 'Longitudinal Joint',    'quantity': long_qty,  'qty_unit': 'm', 'unit_cost': 120},
        ]
    if ptype == 'CRCP':
        long_qty = road_length * 1000
        return [
            {'name': 'Longitudinal Steel (CRCP)', 'quantity': long_qty, 'qty_unit': 'm', 'unit_cost': 200},
            {'name': 'Transverse Joint (End)',     'quantity': 0,        'qty_unit': 'm', 'unit_cost': 500},
        ]
    return []


# ═══════════════════════════════════════════════════════════════════════════════
# CALCULATE FUNCTIONS (ไม่แก้ logic เดิม)
# ═══════════════════════════════════════════════════════════════════════════════

BASE_KEYWORDS = ['crushed rock', 'soil aggregate', 'soil cement', 'cement modified',
                 'cement treated', 'selected material', 'embankment', 'sand embankment']


def calculate_layer_cost(layers: list, road_length_km: float = 1.0) -> tuple:
    """คำนวณค่าก่อสร้างจากชั้นโครงสร้าง — ไม่แก้ logic เดิม"""
    total = 0.0
    details = []
    for layer in layers:
        qty_raw   = float(layer['quantity'])
        unit_cost = float(layer['unit_cost'])
        cost      = qty_raw * unit_cost
        total    += cost

        name_lower = layer['name'].lower()
        is_base = any(kw in name_lower for kw in BASE_KEYWORDS)

        if is_base:
            thick_cm = float(layer.get('thickness', 1))
            u = layer.get('unit', 'cm').lower()
            if layer.get('cost_cum'):
                price_cum = float(layer['cost_cum'])
            elif thick_cm > 0 and u in ('cm', 'ซม.', 'ซ.ม.'):
                price_cum = unit_cost / (thick_cm / 100) if thick_cm > 0 else unit_cost
            else:
                price_cum = unit_cost
            qty_display = qty_raw * thick_cm / 100 if (thick_cm > 0 and u in ('cm', 'ซม.', 'ซ.ม.')) else qty_raw
            display_unit       = 'ลบ.ม.'
            display_price_str  = f"{price_cum:,.0f}"
            display_price_label = 'บาท/ลบ.ม.'
            qty_show = qty_display
        else:
            qty_show            = qty_raw
            display_unit        = 'ตร.ม.'
            display_price_str   = f"{unit_cost:,.0f}"
            display_price_label = 'บาท/ตร.ม.'

        details.append({
            'รายการ':              layer['name'],
            'ความหนา':             f"{layer['thickness']} {layer['unit']}",
            'ปริมาณ':              qty_show,
            'หน่วย':               display_unit,
            'ราคา/หน่วย':          unit_cost,
            'ราคา/หน่วย (แสดง)':   display_price_str,
            'หน่วยราคา':            display_price_label,
            'มูลค่า (บาท)':        cost,
        })
    return total, details


def calculate_joint_cost(joints: list, road_length_km: float = 1.0, include_joints: bool = True) -> tuple:
    """คำนวณค่ารอยต่อ — ไม่แก้ logic เดิม"""
    total = 0.0
    details = []
    for joint in joints:
        qty  = float(joint['quantity'])
        cost = qty * float(joint['unit_cost']) if include_joints else 0.0
        total += cost
        unit_th = 'ม.' if joint.get('qty_unit', 'm') == 'm' else joint.get('qty_unit', 'm')
        details.append({
            'รายการ':             joint['name'],
            'ความหนา':            '-',
            'ปริมาณ':             qty,
            'หน่วย':              unit_th,
            'ราคา/หน่วย':         float(joint['unit_cost']),
            'ราคา/หน่วย (แสดง)':  f"{float(joint['unit_cost']):,.0f}",
            'หน่วยราคา':           'บาท/ม.',
            'มูลค่า (บาท)':       cost,
        })
    return total, details


# ═══════════════════════════════════════════════════════════════════════════════
# RENDER LAYER EDITOR — ใช้ st.data_editor
# ═══════════════════════════════════════════════════════════════════════════════

def _auto_price_surface(name: str, thick: float, ptype: str) -> float:
    """คืนราคา บาท/ตร.ม. สำหรับ surface layer"""
    return lookup_price(name, thick, ptype)


def _auto_price_base(name: str, thick: float) -> tuple:
    """คืน (cost_per_sqm, cost_cum) สำหรับ base layer"""
    cost_cum = lookup_price(name, thick)  # บาท/ลบ.ม.
    cost_sqm = cost_cum * thick / 100 if thick > 0 else 0.0
    return cost_sqm, cost_cum


def render_layer_editor(
    ptype: str,
    key_prefix: str,
    total_width: float,
    road_length: float,
    v: int = 0,
) -> list:
    """
    แสดง editor โครงสร้างชั้นทางด้วย st.data_editor
    คืน updated_layers list สำหรับส่งต่อ calculate_layer_cost()
    """
    area_per_km = total_width * 1000
    proj_area   = area_per_km * road_length
    lib         = get_price_library()
    is_concrete = ptype in ('JPCP', 'JRCP', 'CRCP')

    # ── โหลด state ────────────────────────────────────────────────────────────
    sk_surf = f"{key_prefix}_surf_df_v{v}"
    sk_base = f"{key_prefix}_base_df_v{v}"

    if sk_surf not in st.session_state:
        defaults = get_default_layers(ptype, area_per_km)
        st.session_state[sk_surf] = pd.DataFrame(defaults)
    if sk_base not in st.session_state:
        defaults_b = get_default_base_layers(ptype, area_per_km)
        st.session_state[sk_base] = pd.DataFrame(defaults_b)

    updated_layers: list = []

    # ══════════════════════════════════════════════════════════════
    # SECTION A: ผิวทาง
    # ══════════════════════════════════════════════════════════════
    st.markdown('<div class="section-card"><b>🏗️ ผิวทาง</b> &nbsp;<span style="color:#6b7a8d;font-size:0.85rem">(บาท/ตร.ม.)</span></div>', unsafe_allow_html=True)

    surf_df = st.session_state[sk_surf].copy()

    # กำหนด columns ที่ user แก้ได้
    if is_concrete:
        surf_options = ['Concrete Slab (JPCP)', 'Concrete Slab (JRCP)', 'Concrete Slab (CRCP)',
                        'Non Woven Geotextile', 'Wire Mesh', 'Tack Coat', 'Prime Coat']
        editable_cols = {
            'name':      st.column_config.SelectboxColumn('รายการ', options=surf_options, required=True, width='large'),
            'thickness': st.column_config.NumberColumn('หนา (cm)', min_value=0.0, max_value=50.0, step=1.0, format='%.0f'),
            'unit_cost': st.column_config.NumberColumn('ราคา (บาท/ตร.ม.)', min_value=0.0, step=10.0, format='%.2f'),
        }
    else:
        surf_options = WEARING_OPTIONS + ['AC Binder Course', 'AC Base Course', 'Tack Coat', 'Prime Coat', 'Non Woven Geotextile']
        editable_cols = {
            'name':      st.column_config.SelectboxColumn('รายการ', options=surf_options, required=True, width='large'),
            'thickness': st.column_config.NumberColumn('หนา (cm)', min_value=0.0, max_value=30.0, step=0.5, format='%.1f'),
            'unit_cost': st.column_config.NumberColumn('ราคา (บาท/ตร.ม.)', min_value=0.0, step=10.0, format='%.2f'),
        }

    display_surf = surf_df[['name', 'thickness', 'unit_cost']].copy() if all(c in surf_df.columns for c in ['name', 'thickness', 'unit_cost']) else pd.DataFrame(columns=['name', 'thickness', 'unit_cost'])

    edited_surf = st.data_editor(
        display_surf,
        column_config=editable_cols,
        num_rows='dynamic',
        use_container_width=True,
        key=f"{key_prefix}_surf_editor_v{v}",
        hide_index=True,
    )

    # Auto-update ราคาจาก library เมื่อชื่อ/ความหนาเปลี่ยน
    for _, row in edited_surf.iterrows():
        name  = str(row.get('name', '') or '')
        thick = float(row.get('thickness', 0) or 0)
        if not name or thick == 0:
            continue
        # ราคาจาก library (ถ้าผู้ใช้ไม่แก้ หรือยังเป็น 0)
        lib_price = lookup_price(name, thick, ptype)
        unit_cost = float(row.get('unit_cost', 0) or 0)
        if unit_cost == 0 and lib_price > 0:
            unit_cost = lib_price

        unit = 'cm' if thick > 1 else 'ชั้น'
        updated_layers.append({
            'name':       name,
            'thickness':  thick,
            'unit':       unit,
            'quantity':   proj_area,
            'qty_unit':   'sq.m',
            'unit_cost':  unit_cost,
            'cost_per_sqm': unit_cost,
        })

    # บันทึก state กลับ
    st.session_state[sk_surf] = edited_surf.copy()

    # ══════════════════════════════════════════════════════════════
    # SECTION B: Checkboxes (AC Interlayer / Prime Coat / Geotextile)
    #            เฉพาะโครงสร้างคอนกรีต
    # ══════════════════════════════════════════════════════════════
    if is_concrete:
        st.markdown("---")
        col_cb1, col_cb2, col_cb3 = st.columns(3)

        with col_cb1:
            use_acil = st.checkbox(
                "AC Interlayer รองใต้แผ่น",
                value=True,
                key=f"{key_prefix}_use_acil_v{v}",
                help="ชั้น AC ที่รองใต้แผ่นคอนกรีต ทั่วไป 5 cm"
            )
        with col_cb2:
            use_pc = st.checkbox(
                "Prime Coat",
                value=True,
                key=f"{key_prefix}_use_pc_v{v}",
                help="ราดบน Base ก่อนปู AC Interlayer"
            )
        with col_cb3:
            use_geo = st.checkbox(
                "Non Woven Geotextile (แผ่น)",
                value=False,
                key=f"{key_prefix}_use_geo_v{v}",
                help="แผ่น Geotextile รองใต้แผ่นคอนกรีต (นอกเหนือจาก Editor ด้านบน)"
            )

        if use_acil:
            acil_thick = st.number_input(
                "ความหนา AC Interlayer (cm)",
                value=5.0, min_value=1.0, max_value=15.0, step=1.0,
                key=f"{key_prefix}_acil_thick_v{v}",
            )
            acil_price = lookup_price('AC Binder Course', acil_thick, ptype)
            if acil_price == 0:
                acil_price = 251.0
            st.caption(f"AC Interlayer {acil_thick:.0f} cm → **{acil_price:,.2f}** บาท/ตร.ม.")
            updated_layers.append({
                'name': f'AC Interlayer ({acil_thick:.0f} cm)',
                'thickness': acil_thick, 'unit': 'cm',
                'quantity': proj_area, 'qty_unit': 'sq.m',
                'unit_cost': acil_price, 'cost_per_sqm': acil_price,
            })

        if use_pc:
            pc_default = float(lib['base_prices'].get('Prime Coat', 37.47))
            pc_price = st.number_input(
                "ราคา Prime Coat (บาท/ตร.ม.)",
                value=pc_default, min_value=0.0, step=1.0,
                key=f"{key_prefix}_pc_price_v{v}",
            )
            updated_layers.append({
                'name': 'Prime Coat', 'thickness': 1, 'unit': 'Layer',
                'quantity': proj_area, 'qty_unit': 'sq.m',
                'unit_cost': pc_price, 'cost_per_sqm': pc_price,
            })

        if use_geo:
            geo_price = float(lib['base_prices'].get('Non Woven Geotextile', 78))
            st.caption(f"Non Woven Geotextile → **{geo_price:,.2f}** บาท/ตร.ม.")
            updated_layers.append({
                'name': 'Non Woven Geotextile', 'thickness': 1, 'unit': 'ชั้น',
                'quantity': proj_area, 'qty_unit': 'sq.m',
                'unit_cost': geo_price, 'cost_per_sqm': geo_price,
            })

    # ══════════════════════════════════════════════════════════════
    # SECTION C: พื้นทาง / รองพื้นทาง
    # ══════════════════════════════════════════════════════════════
    st.markdown("---")
    st.markdown('<div class="section-card"><b>🪨 พื้นทาง / รองพื้นทาง</b> &nbsp;<span style="color:#6b7a8d;font-size:0.85rem">(ราคาแสดงเป็น บาท/ตร.ม., คำนวณจาก บาท/ลบ.ม. × ความหนา)</span></div>', unsafe_allow_html=True)

    base_df = st.session_state[sk_base].copy()
    display_base = base_df[['name', 'thickness', 'unit_cost', 'cost_cum']].copy() if all(c in base_df.columns for c in ['name', 'thickness', 'unit_cost', 'cost_cum']) else pd.DataFrame(columns=['name', 'thickness', 'unit_cost', 'cost_cum'])

    edited_base = st.data_editor(
        display_base,
        column_config={
            'name':      st.column_config.SelectboxColumn('วัสดุ', options=BASE_MATERIAL_LIST, required=True, width='large'),
            'thickness': st.column_config.NumberColumn('หนา (cm)', min_value=0.0, max_value=150.0, step=5.0, format='%.0f'),
            'cost_cum':  st.column_config.NumberColumn('ราคา (บาท/ลบ.ม.)', min_value=0.0, step=10.0, format='%.0f'),
            'unit_cost': st.column_config.NumberColumn('ราคา (บาท/ตร.ม.)', min_value=0.0, step=1.0, format='%.2f', disabled=True),
        },
        num_rows='dynamic',
        use_container_width=True,
        key=f"{key_prefix}_base_editor_v{v}",
        hide_index=True,
    )

    st.caption("💡 แก้ **ราคา (บาท/ลบ.ม.)** ระบบจะคำนวณ บาท/ตร.ม. ให้อัตโนมัติ")

    for _, row in edited_base.iterrows():
        name  = str(row.get('name', '') or '')
        thick = float(row.get('thickness', 0) or 0)
        if not name or thick == 0:
            continue

        # ดึงราคา/ลบ.ม. จาก library ถ้า cost_cum ยังเป็น 0
        cost_cum = float(row.get('cost_cum', 0) or 0)
        if cost_cum == 0:
            cost_cum = lookup_price(name, thick)
        cost_sqm = cost_cum * thick / 100

        updated_layers.append({
            'name':       name,
            'thickness':  thick,
            'unit':       'cm',
            'quantity':   proj_area,
            'qty_unit':   'sq.m',
            'unit_cost':  cost_sqm,
            'cost_per_sqm': cost_sqm,
            'cost_cum':   cost_cum,
        })

    st.session_state[sk_base] = edited_base.copy()

    return updated_layers


# ═══════════════════════════════════════════════════════════════════════════════
# RENDER JOINT EDITOR — ใช้ st.data_editor
# ═══════════════════════════════════════════════════════════════════════════════

def render_joint_editor(
    ptype: str,
    key_prefix: str,
    area_per_km: float,
    road_length: float,
    v: int = 0,
) -> tuple:
    """
    แสดง editor รอยต่อด้วย st.data_editor
    คืน (updated_joints, include_joints)
    """
    width_m    = area_per_km / 1000
    lane_w     = st.session_state.get('project_info', {}).get('lane_width', 3.5)
    if not lane_w or lane_w <= 0:
        lane_w = 3.5
    total_area = area_per_km * road_length

    # spacing label
    if ptype == 'JPCP':
        joint_label = 'Transverse Joint @4m'
        spacing = 4
    else:
        joint_label = 'Transverse Joint @10m'
        spacing = 10

    sk_joint = f"{key_prefix}_joint_df_v{v}"

    if sk_joint not in st.session_state:
        defaults_j = get_default_joints(ptype, area_per_km, road_length)
        st.session_state[sk_joint] = pd.DataFrame(defaults_j)

    st.markdown("---")
    col_h1, col_h2 = st.columns([3, 1])
    with col_h1:
        if ptype == 'CRCP':
            st.markdown('<div class="section-card"><b>⛓️ Longitudinal Steel & Transverse Joint (CRCP)</b></div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="section-card"><b>🔗 รอยต่อ (Joints) — {ptype} ระยะ {spacing} ม.</b></div>', unsafe_allow_html=True)
    with col_h2:
        _cb_label = "รวมราคา Steel & Joints" if ptype == 'CRCP' else "รวมราคา Joints"
        include_joints = st.checkbox(_cb_label, value=True, key=f"{key_prefix}_include_joints_v{v}")

    joint_df = st.session_state[sk_joint].copy()
    display_joint = joint_df[['name', 'quantity', 'unit_cost']].copy() if all(c in joint_df.columns for c in ['name', 'quantity', 'unit_cost']) else pd.DataFrame(columns=['name', 'quantity', 'unit_cost'])

    # คำนวณ auto qty
    if ptype in ('JPCP', 'JRCP'):
        auto_trans = (road_length * 1000 / spacing) * width_m
        auto_long  = max(1, round(width_m / lane_w) - 1) * road_length * 1000
    else:
        auto_trans = 0.0
        auto_long  = max(1, round(width_m / lane_w) - 1) * road_length * 1000

    # ปรับ quantity ใน display ตาม ptype
    rows = []
    for _, row in display_joint.iterrows():
        name = str(row.get('name', ''))
        qty  = float(row.get('quantity', 0))
        if 'transverse' in name.lower():
            qty = auto_trans if ptype in ('JPCP', 'JRCP') else qty
        elif 'longitudinal' in name.lower() or 'steel' in name.lower():
            qty = auto_long
        rows.append({'name': name, 'quantity': qty, 'unit_cost': float(row.get('unit_cost', 0))})

    if rows:
        display_joint = pd.DataFrame(rows)

    edited_joint = st.data_editor(
        display_joint,
        column_config={
            'name':      st.column_config.TextColumn('รายการ', width='large'),
            'quantity':  st.column_config.NumberColumn('ปริมาณ (ม.)', min_value=0.0, step=100.0, format='%.0f'),
            'unit_cost': st.column_config.NumberColumn('ราคา/ม. (บาท)', min_value=0.0, step=10.0, format='%.0f'),
        },
        num_rows='dynamic',
        use_container_width=True,
        key=f"{key_prefix}_joint_editor_v{v}",
        hide_index=True,
    )

    # แสดงราคา/ตร.ม. ใต้ตาราง
    joint_total_cost = 0.0
    updated_joints = []
    for _, row in edited_joint.iterrows():
        name = str(row.get('name', '') or '')
        qty  = float(row.get('quantity', 0) or 0)
        uc   = float(row.get('unit_cost', 0) or 0)
        cpsqm = (qty * uc / total_area) if total_area > 0 else 0.0
        joint_total_cost += qty * uc
        updated_joints.append({
            'name': name, 'quantity': qty,
            'qty_unit': 'm', 'unit_cost': uc,
            'cost_per_sqm': cpsqm,
        })

    if total_area > 0:
        st.caption(f"รวม Joints = **{joint_total_cost/total_area:,.2f}** บาท/ตร.ม. | **{joint_total_cost/1e6:,.3f}** ล้านบาท/โครงการ")

    st.session_state[sk_joint] = edited_joint.copy()
    return updated_joints, include_joints


# ═══════════════════════════════════════════════════════════════════════════════
# WORD REPORT
# ═══════════════════════════════════════════════════════════════════════════════

def _set_run_font(run, size: int = 16, bold: bool = False, italic: bool = False):
    run.font.name = 'TH SarabunPSK'
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    for attr in ('w:eastAsia', 'w:ascii', 'w:hAnsi'):
        rFonts.set(qn(attr), 'TH SarabunPSK')


def generate_word_report(
    project_info: dict,
    all_details: dict,
    report_type: str = 'materials',
    chapter_num: str = '4',
    section_start: str = '4.7',
    intro_text: str = '',
) -> 'Document':
    """สร้างรายงาน Word — report_type: 'materials' | 'consultant'"""
    if not DOCX_AVAILABLE:
        raise ImportError("python-docx ไม่สามารถใช้งานได้")

    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'TH SarabunPSK'
    style.font.size = Pt(16)

    length = float(project_info.get('length', 1))

    def add_heading(text, size=16, bold=True, underline=False, space_before=6, space_after=3):
        para = doc.add_paragraph()
        para.paragraph_format.space_before = Pt(space_before)
        para.paragraph_format.space_after  = Pt(space_after)
        run = para.add_run(text)
        _set_run_font(run, size=size, bold=bold)
        run.underline = underline
        return para

    if report_type == 'consultant':
        add_heading(f"{section_start} รายงานวัสดุและราคาโครงสร้างชั้นทาง",
                    size=18, bold=True, underline=True, space_before=12, space_after=6)
        add_heading(f"{section_start}.1 ข้อมูลของถนน", size=16, bold=True, underline=True, space_before=8)
        if intro_text:
            p = doc.add_paragraph()
            p.paragraph_format.first_line_indent = Cm(1.0)
            _set_run_font(p.add_run(intro_text), size=16)
    else:
        doc.add_heading('รายงานวัสดุและราคาโครงสร้างชั้นทาง', 0)
        doc.add_heading('1. ข้อมูลโครงการ', level=1)

    fields = [
        ("ชื่อโครงการ",      project_info.get('name', '-')),
        ("ความยาวถนน",      f"{length:.2f} กม."),
        ("ความกว้างรวม",    f"{project_info.get('total_width', 0):.2f} ม."),
        ("จำนวนช่องจราจร",  f"{project_info.get('num_lanes', 4)} ช่อง"),
        ("อายุออกแบบ",      f"{project_info.get('design_life', 20)} ปี"),
    ]
    for label, value in fields:
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Cm(1.0)
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after  = Pt(2)
        r1 = p.add_run(f"{label}: ")
        _set_run_font(r1, size=16, bold=True)
        r2 = p.add_run(value)
        _set_run_font(r2, size=16)

    sec_detail = f"{section_start}.2" if report_type == 'consultant' else '2.'
    add_heading(f"{sec_detail} รายละเอียดวัสดุและราคา", size=16, bold=True, underline=True, space_before=10)

    summary_data = []
    for ptype, data in all_details.items():
        sname   = data.get('name', ptype)
        details = data.get('details', [])

        add_heading(f"ผิวทางประเภท {sname}", size=16, bold=True, space_before=6, space_after=2)
        if data.get('name_detail', '') and data['name_detail'] != sname:
            p_sub = doc.add_paragraph()
            r_sub = p_sub.add_run(data['name_detail'])
            _set_run_font(r_sub, size=14, italic=True)

        if details:
            table = doc.add_table(rows=len(details) + 2, cols=5)
            table.style = 'Table Grid'
            col_widths = [Cm(6.5), Cm(2.5), Cm(1.8), Cm(3.5), Cm(3.5)]
            for row in table.rows:
                for idx, cell in enumerate(row.cells):
                    cell.width = col_widths[idx]

            headers = ['รายการ', 'ปริมาณ', 'หน่วย', 'ราคา/หน่วย (บาท)', 'มูลค่า (บาท)']
            for j, h in enumerate(headers):
                p_h = table.rows[0].cells[j].paragraphs[0]
                p_h.alignment = WD_ALIGN_PARAGRAPH.CENTER
                _set_run_font(p_h.add_run(h), size=15, bold=True)

            subtotal = 0.0
            for i, d in enumerate(details):
                rc = table.rows[i + 1].cells
                _set_run_font(rc[0].paragraphs[0].add_run(str(d['รายการ'])), size=15)
                rc[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
                _set_run_font(rc[1].paragraphs[0].add_run(f"{d['ปริมาณ']:,.0f}"), size=15)
                rc[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                _set_run_font(rc[2].paragraphs[0].add_run(d.get('หน่วย', 'ตร.ม.')), size=15)
                rc[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
                _set_run_font(rc[3].paragraphs[0].add_run(d.get('ราคา/หน่วย (แสดง)', '')), size=15)
                _set_run_font(rc[3].paragraphs[0].add_run(f" ({d.get('หน่วยราคา', 'บาท/ตร.ม.')})"), size=12)
                rc[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
                _set_run_font(rc[4].paragraphs[0].add_run(f"{d['มูลค่า (บาท)']:,.0f}"), size=15)
                subtotal += d['มูลค่า (บาท)']

            # แถวรวม
            lr = table.rows[len(details) + 1]
            lr.cells[0].merge(lr.cells[3])
            lr.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
            _set_run_font(lr.cells[0].paragraphs[0].add_run(f"รวม {sname}"), size=15, bold=True)
            lr.cells[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
            _set_run_font(lr.cells[4].paragraphs[0].add_run(f"{subtotal:,.0f}"), size=15, bold=True)

            doc.add_paragraph()
            summary_data.append({
                'name':              sname,
                'total_value':       subtotal,
                'cost_per_km':       data.get('cost_per_km', 0),
                'cost_sqm':          data.get('cost_sqm', 0),
            })

    sec_sum = f"{section_start}.3" if report_type == 'consultant' else '3.'
    add_heading(f"{sec_sum} สรุปค่าใช้จ่าย", size=16, bold=True, underline=True, space_before=10)

    if summary_data:
        sum_table = doc.add_table(rows=len(summary_data) + 1, cols=4)
        sum_table.style = 'Table Grid'
        hdrs = ['ชนิดโครงสร้าง', 'มูลค่ารวม/กม. (บาท)', 'ราคา/กม. (ล้านบาท)', 'ราคา/ตร.ม. (บาท)']
        for j, h in enumerate(hdrs):
            ph = sum_table.rows[0].cells[j].paragraphs[0]
            ph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _set_run_font(ph.add_run(h), size=15, bold=True)
        for i, item in enumerate(summary_data):
            tpk = item['total_value'] / length if length > 0 else 0
            sum_table.rows[i+1].cells[0].text = item['name']
            sum_table.rows[i+1].cells[1].text = f"{tpk:,.0f}"
            sum_table.rows[i+1].cells[2].text = f"{item['cost_per_km']:.3f}"
            sum_table.rows[i+1].cells[3].text = f"{item['cost_sqm']:,.2f}"

    doc.add_paragraph()
    p_date = doc.add_paragraph()
    _set_run_font(p_date.add_run(f"รายงานสร้างเมื่อ: {datetime.now().strftime('%d/%m/%Y %H:%M')}"), size=14)
    p_by = doc.add_paragraph()
    _set_run_font(p_by.add_run("พัฒนาโดย รศ.ดร.อิทธิพล มีผล — KMUTNB"), size=14)
    return doc


# ═══════════════════════════════════════════════════════════════════════════════
# EXCEL TEMPLATE
# ═══════════════════════════════════════════════════════════════════════════════

def generate_excel_template() -> bytes:
    lib = get_price_library()
    ac_rows = []
    for mat, prices in lib['ac_prices'].items():
        row = {'Material': mat}
        for t, p in prices.items():
            row[f"{t}cm"] = p
        ac_rows.append(row)

    conc_rows = []
    for ct, prices in lib['concrete_prices'].items():
        row = {'Type': ct}
        for t, p in prices.items():
            row[f"{t}cm"] = p
        conc_rows.append(row)

    base_rows = [{'Material': k, 'Price (Baht/cu.m)': v}
                 for k, v in lib['base_prices'].items()]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame(ac_rows).to_excel(writer, sheet_name='AC_Prices', index=False)
        pd.DataFrame(conc_rows).to_excel(writer, sheet_name='Concrete_Prices', index=False)
        pd.DataFrame(base_rows).to_excel(writer, sheet_name='Base_Materials', index=False)
    output.seek(0)
    return output.getvalue()


def load_excel_price_library(uploaded_file) -> dict:
    """อ่าน Excel → dict price_library"""
    ac_df   = pd.read_excel(uploaded_file, sheet_name='AC_Prices')
    conc_df = pd.read_excel(uploaded_file, sheet_name='Concrete_Prices')
    base_df = pd.read_excel(uploaded_file, sheet_name='Base_Materials')

    ac_prices: dict = {}
    for _, row in ac_df.iterrows():
        mat = row['Material']
        prices = {}
        for col in ac_df.columns[1:]:
            try:
                t = float(col.replace('cm', '').strip())
                v = row[col]
                if pd.notna(v):
                    prices[t] = float(v)
            except (ValueError, TypeError):
                pass
        if prices:
            ac_prices[mat] = prices
    for mat, dp in DEFAULT_AC_PRICES.items():
        if mat not in ac_prices:
            ac_prices[mat] = dict(dp)
        else:
            for t, p in dp.items():
                ac_prices[mat].setdefault(t, p)

    conc_prices: dict = {}
    for _, row in conc_df.iterrows():
        ct = row['Type']
        prices = {}
        for col in conc_df.columns[1:]:
            try:
                t = int(float(col.replace('cm', '').strip()))
                v = row[col]
                if pd.notna(v):
                    prices[t] = float(v)
            except (ValueError, TypeError):
                pass
        if prices:
            conc_prices[ct] = prices
    for ct, dp in DEFAULT_CONCRETE_PRICES.items():
        if ct not in conc_prices:
            conc_prices[ct] = dict(dp)
        else:
            for t, p in dp.items():
                conc_prices[ct].setdefault(t, p)

    base_prices: dict = dict(DEFAULT_BASE_PRICES)
    for _, row in base_df.iterrows():
        try:
            mat = str(row['Material'])
            val = row['Price (Baht/cu.m)']
            if pd.notna(mat) and pd.notna(val):
                base_prices[mat] = float(val)
        except (ValueError, TypeError):
            pass

    return {'ac_prices': ac_prices, 'concrete_prices': conc_prices, 'base_prices': base_prices}


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN APP
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    # ── Header ────────────────────────────────────────────────────────────────
    st.markdown("""
    <div class="main-header">
        <h1>🛣️ ระบบวิเคราะห์ค่าก่อสร้างโครงสร้างชั้นทาง</h1>
        <p>ตามแนวทาง AASHTO 1993 &nbsp;|&nbsp; รองรับ AC · JPCP · JRCP · CRCP</p>
    </div>
    """, unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════════
    # SIDEBAR
    # ══════════════════════════════════════════════════════════════
    with st.sidebar:
        st.markdown("## 📋 ข้อมูลโครงการ")

        # ── Price Library (Excel) ──
        with st.expander("💰 Price Library (Excel)", expanded=False):
            uploaded_excel = st.file_uploader(
                "Upload Excel", type=['xlsx', 'xls'],
                key="sidebar_price_excel",
                help="ดาวน์โหลด Template จาก Tab 2 ก่อน แล้วแก้ราคา แล้ว Upload กลับ"
            )
            if uploaded_excel is not None:
                try:
                    lib = load_excel_price_library(uploaded_excel)
                    st.session_state['price_library'] = lib
                    st.success("✅ โหลด Price Library สำเร็จ")
                    st.caption(f"AC: {len(lib['ac_prices'])} ประเภท | Concrete: {len(lib['concrete_prices'])} ประเภท")
                except Exception as e:
                    st.error(f"❌ อ่านไฟล์ไม่สำเร็จ: {e}")

        st.divider()

        # ── Load Project (JSON) ──
        with st.expander("📂 โหลดโครงการ (JSON)", expanded=False):
            uploaded_json = st.file_uploader(
                "Upload JSON", type=['json'],
                key="sidebar_upload_json",
            )
            if uploaded_json is not None:
                try:
                    import hashlib
                    fb = uploaded_json.read()
                    fh = hashlib.md5(fb).hexdigest()
                    loaded = json.loads(fb.decode('utf-8'))
                    if 'project_info' in loaded:
                        st.info(f"📌 {loaded['project_info'].get('name', '-')}")
                        st.caption(f"บันทึกเมื่อ: {loaded.get('saved_at', '-')}")
                    if st.button("📥 นำเข้าข้อมูล", key="import_json_btn"):
                        if st.session_state.get('loaded_json_hash') != fh:
                            st.session_state['loaded_project'] = loaded
                            st.session_state['loaded_json_hash'] = fh
                            new_v = st.session_state.get('json_version', 0) + 1
                            st.session_state['json_version'] = new_v
                            # ล้าง data_editor state ทั้งหมด
                            keys_to_clear = [k for k in st.session_state
                                             if any(p in k for p in ['_surf_df_', '_base_df_', '_joint_df_',
                                                                       '_surf_editor_', '_base_editor_', '_joint_editor_'])]
                            for k in keys_to_clear:
                                del st.session_state[k]
                        st.rerun()
                except Exception as e:
                    st.error(f"❌ ไม่สามารถอ่านไฟล์: {e}")

        st.divider()

        # ── Project Info ──
        lp   = st.session_state.get('loaded_project', {})
        li   = lp.get('project_info', {})
        v_sb = st.session_state.get('json_version', 0)

        project_name = st.text_input(
            "ชื่อโครงการ",
            value=li.get('name', 'โครงการก่อสร้างทางหลวง'),
            key=f"sb_pname_v{v_sb}"
        )
        road_length = st.number_input(
            "ความยาวถนน (กม.)",
            value=float(li.get('length', 1.0)),
            min_value=0.1, step=0.1,
            key=f"sb_length_v{v_sb}"
        )
        design_life = st.number_input(
            "อายุออกแบบ (ปี)",
            value=int(li.get('design_life', 20)),
            min_value=1, max_value=50, step=1,
            key=f"sb_dlife_v{v_sb}"
        )

        st.divider()
        st.markdown("**📐 ขนาดถนน**")

        lane_width = st.number_input(
            "ความกว้างช่องจราจร (ม.)",
            value=float(li.get('lane_width', 3.5)),
            min_value=2.5, max_value=4.5, step=0.25,
            key=f"sb_lw_v{v_sb}"
        )
        lpo = [2, 3, 4]
        lp_def = li.get('num_lanes', 4) // 2
        lp_idx = lpo.index(lp_def) if lp_def in lpo else 0
        lanes_per_dir = st.selectbox(
            "ช่องจราจร/ทิศทาง",
            options=lpo, index=lp_idx,
            key=f"sb_lpd_v{v_sb}"
        )
        num_lanes = lanes_per_dir * 2

        shoulder_l = st.number_input(
            "ไหล่ทางซ้าย (ม.)",
            value=float(li.get('shoulder_left', 2.5)),
            min_value=0.0, max_value=4.0, step=0.25,
            key=f"sb_sl_v{v_sb}"
        )
        shoulder_r = st.number_input(
            "ไหล่ทางขวา (ม.)",
            value=float(li.get('shoulder_right', 1.5)),
            min_value=0.0, max_value=4.0, step=0.25,
            key=f"sb_sr_v{v_sb}"
        )

        road_surface_width = lane_width * num_lanes
        total_shoulders    = (shoulder_l + shoulder_r) * 2
        total_width        = road_surface_width + total_shoulders
        area_per_km        = total_width * 1000

        st.info(
            f"📏 ช่องรวม: **{num_lanes}** ช่อง\n"
            f"📏 ผิวจราจร: **{road_surface_width:.2f}** ม.\n"
            f"📏 ไหล่ทาง: **{total_shoulders:.2f}** ม.\n"
            f"📏 กว้างรวม: **{total_width:.2f}** ม."
        )

    project_info = {
        'name':         project_name,
        'length':       road_length,
        'design_life':  design_life,
        'lane_width':   lane_width,
        'num_lanes':    num_lanes,
        'shoulder_left':  shoulder_l,
        'shoulder_right': shoulder_r,
        'total_width':  total_width,
    }
    st.session_state['project_info'] = project_info

    v = st.session_state.get('json_version', 0)

    # ══════════════════════════════════════════════════════════════
    # TABS
    # ══════════════════════════════════════════════════════════════
    tab1, tab2, tab3 = st.tabs([
        "🏗️ กำหนดหน้าตัด",
        "💰 ราคาวัสดุ",
        "📊 สรุปต้นทุน & รายงาน",
    ])

    all_results: dict = {}   # เก็บผลคำนวณจาก Tab 1 ส่งต่อ Tab 3

    # ══════════════════════════════════════════════════════════════
    # TAB 1 — Section Setup
    # ══════════════════════════════════════════════════════════════
    with tab1:
        st.markdown(f"**โครงการ:** {project_name} &nbsp;|&nbsp; **ความยาว:** {road_length:.2f} กม. &nbsp;|&nbsp; **กว้างรวม:** {total_width:.2f} ม.")

        sub_ac, sub_jpcp, sub_jrcp, sub_crcp = st.tabs(["🛣️ AC", "🏗️ JPCP", "🏗️ JRCP", "🏗️ CRCP"])

        for sub_tab, ptype in [(sub_ac, 'AC'), (sub_jpcp, 'JPCP'), (sub_jrcp, 'JRCP'), (sub_crcp, 'CRCP')]:
            with sub_tab:
                kp = ptype.lower()
                layers = render_layer_editor(ptype, kp, total_width, road_length, v=v)

                joints = []
                include_joints = True
                if ptype in ('JPCP', 'JRCP', 'CRCP'):
                    joints, include_joints = render_joint_editor(ptype, kp, area_per_km, road_length, v=v)

                # คำนวณ
                layer_cost, layer_details = calculate_layer_cost(layers, road_length)
                joint_cost, joint_details = calculate_joint_cost(joints, road_length, include_joints)
                total_cost = layer_cost + joint_cost

                total_area = area_per_km * road_length
                cost_sqm   = total_cost / total_area if total_area > 0 else 0
                cost_per_km_m = total_cost / road_length / 1e6 if road_length > 0 else 0

                # Metric
                st.markdown("---")
                mc1, mc2, mc3 = st.columns(3)
                with mc1:
                    st.markdown(f"""<div class="metric-card">
                        <div class="label">💰 ราคา/ตร.ม.</div>
                        <div class="value">{cost_sqm:,.2f}<span class="unit">บาท</span></div>
                    </div>""", unsafe_allow_html=True)
                with mc2:
                    st.markdown(f"""<div class="metric-card">
                        <div class="label">📏 ราคา/กม.</div>
                        <div class="value">{cost_per_km_m:,.3f}<span class="unit">ล้านบาท</span></div>
                    </div>""", unsafe_allow_html=True)
                with mc3:
                    st.markdown(f"""<div class="metric-card">
                        <div class="label">⏱️ อายุออกแบบ</div>
                        <div class="value">{design_life}<span class="unit">ปี</span></div>
                    </div>""", unsafe_allow_html=True)

                # เก็บผลสำหรับ Tab 3
                all_results[ptype] = {
                    'name':        ptype,
                    'name_detail': f"{ptype} — {project_name}",
                    'layers':      layers,
                    'joints':      joints,
                    'details':     layer_details + joint_details,
                    'cost_total':  total_cost,
                    'cost_sqm':    cost_sqm,
                    'cost_per_km': cost_per_km_m,
                    'include_joints': include_joints,
                }

        # บันทึก all_results ไว้ใน session state สำหรับ Tab 3
        st.session_state['all_results'] = all_results

    # ══════════════════════════════════════════════════════════════
    # TAB 2 — Price Library
    # ══════════════════════════════════════════════════════════════
    with tab2:
        st.header("💰 ตารางราคาวัสดุ")
        st.info("💡 แก้ราคาได้โดยตรงในตาราง หรือ Upload Excel ใน Sidebar — การแก้ในตารางนี้จะมีผลทันที")

        lib = get_price_library()

        col_dl, col_up = st.columns([1, 2])
        with col_dl:
            tpl = generate_excel_template()
            st.download_button(
                "⬇️ Download Template Excel",
                data=tpl,
                file_name="price_library_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        st.subheader("🛣️ ราคา AC (บาท/ตร.ม.)")
        ac_rows = []
        for mat, prices in lib['ac_prices'].items():
            row = {'วัสดุ': mat}
            for t in [2.5, 3, 4, 5, 6, 7, 8, 9, 10]:
                row[f"{t}cm"] = prices.get(t, 0)
            ac_rows.append(row)
        ac_edited = st.data_editor(
            pd.DataFrame(ac_rows),
            use_container_width=True,
            hide_index=True,
            key="tab2_ac_editor",
        )

        st.subheader("🏗️ ราคาคอนกรีต (บาท/ตร.ม.)")
        cp_rows = []
        for ct, prices in lib['concrete_prices'].items():
            row = {'ประเภท': ct}
            for t in [25, 28, 30, 32, 35]:
                row[f"{t}cm"] = prices.get(t, 0)
            cp_rows.append(row)
        cp_edited = st.data_editor(
            pd.DataFrame(cp_rows),
            use_container_width=True,
            hide_index=True,
            key="tab2_cp_editor",
        )

        st.subheader("🪨 ราคาวัสดุพื้นทาง (บาท/ลบ.ม.)")
        bp_rows = [{'วัสดุ': k, 'ราคา (บาท/ลบ.ม.)': v} for k, v in lib['base_prices'].items()]
        bp_edited = st.data_editor(
            pd.DataFrame(bp_rows),
            use_container_width=True,
            hide_index=True,
            key="tab2_bp_editor",
        )

        if st.button("💾 บันทึกราคาที่แก้ไขลง Library", type="primary"):
            # อ่านจาก edited tables → update price_library
            new_ac: dict = {}
            for _, row in ac_edited.iterrows():
                mat = str(row['วัสดุ'])
                new_ac[mat] = {float(c.replace('cm', '')): float(row[c])
                               for c in ac_edited.columns if c.endswith('cm')}
            new_cp: dict = {}
            for _, row in cp_edited.iterrows():
                ct = str(row['ประเภท'])
                new_cp[ct] = {int(float(c.replace('cm', ''))): float(row[c])
                              for c in cp_edited.columns if c.endswith('cm')}
            new_bp: dict = {}
            for _, row in bp_edited.iterrows():
                new_bp[str(row['วัสดุ'])] = float(row['ราคา (บาท/ลบ.ม.)'])

            st.session_state['price_library'] = {
                'ac_prices': new_ac,
                'concrete_prices': new_cp,
                'base_prices': new_bp,
            }
            st.success("✅ อัพเดท Price Library สำเร็จ — กลับไป Tab 1 เพื่อดูผล")

    # ══════════════════════════════════════════════════════════════
    # TAB 3 — Cost Summary + Report
    # ══════════════════════════════════════════════════════════════
    with tab3:
        st.header("📊 สรุปต้นทุนและรายงาน")

        ar = st.session_state.get('all_results', all_results)

        if not ar:
            st.info("กรุณากำหนดหน้าตัดใน Tab 1 ก่อน")
        else:
            # ── Summary metrics ──
            ptypes_list = list(ar.keys())
            cols_m = st.columns(len(ptypes_list))
            for i, pt in enumerate(ptypes_list):
                r = ar[pt]
                with cols_m[i]:
                    st.markdown(f"""<div class="metric-card">
                        <div class="label">{pt}</div>
                        <div class="value">{r['cost_sqm']:,.0f}<span class="unit">บาท/ตร.ม.</span></div>
                        <div class="label" style="margin-top:6px">{r['cost_per_km']:.3f} ล้านบาท/กม.</div>
                    </div>""", unsafe_allow_html=True)

            st.markdown("---")

            # ── Plotly bar chart ──
            if PLOTLY_AVAILABLE:
                fig = go.Figure()
                fig.add_trace(go.Bar(
                    x=[ar[pt]['name'] for pt in ptypes_list],
                    y=[ar[pt]['cost_sqm'] for pt in ptypes_list],
                    marker_color=['#1a4a7a', '#0d7377', '#14a085', '#52b788'],
                    text=[f"{ar[pt]['cost_sqm']:,.0f}" for pt in ptypes_list],
                    textposition='outside',
                    name='ราคา (บาท/ตร.ม.)',
                ))
                fig.update_layout(
                    title='เปรียบเทียบราคา/ตร.ม. ตามประเภทโครงสร้าง',
                    yaxis_title='บาท/ตร.ม.',
                    plot_bgcolor='white',
                    paper_bgcolor='white',
                    font=dict(family='IBM Plex Sans Thai, sans-serif'),
                    height=380,
                )
                st.plotly_chart(fig, use_container_width=True)

            # ── ตารางสรุปรายละเอียด ──
            st.subheader("📋 รายละเอียดต้นทุนแต่ละประเภท")
            for pt in ptypes_list:
                r = ar[pt]
                with st.expander(f"🔍 {pt} — {r['cost_sqm']:,.2f} บาท/ตร.ม.", expanded=False):
                    if r['details']:
                        st.dataframe(
                            pd.DataFrame(r['details'])[['รายการ', 'ปริมาณ', 'หน่วย', 'ราคา/หน่วย (แสดง)', 'หน่วยราคา', 'มูลค่า (บาท)']],
                            use_container_width=True, hide_index=True
                        )

            st.markdown("---")
            st.subheader("💾 บันทึกและส่งออก")

            col_s1, col_s2, col_s3 = st.columns(3)

            # ── Save JSON ──
            with col_s1:
                construction_out: dict = {}
                for pt, r in ar.items():
                    construction_out[pt] = {
                        'layers': r.get('layers', []),
                        'joints': r.get('joints', []),
                        'include_joints': r.get('include_joints', True),
                    }
                save_data = {
                    'saved_at':    datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'project_info': project_info,
                    'construction': construction_out,
                }
                json_bytes = json.dumps(save_data, ensure_ascii=False, indent=2, default=str).encode('utf-8')
                fname_json = f"{project_name.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.json"
                st.download_button(
                    "📥 บันทึก JSON",
                    data=json_bytes,
                    file_name=fname_json,
                    mime="application/json",
                    use_container_width=True,
                )

            # ── Word Materials ──
            with col_s2:
                if DOCX_AVAILABLE:
                    if st.button("📄 Word แบบวัสดุ+ราคา", use_container_width=True):
                        try:
                            doc = generate_word_report(project_info, ar, report_type='materials')
                            buf = io.BytesIO()
                            doc.save(buf)
                            buf.seek(0)
                            fname_w = f"{project_name.replace(' ', '_')}_Materials_{datetime.now().strftime('%Y%m%d')}.docx"
                            st.download_button(
                                "⬇️ ดาวน์โหลด Word (วัสดุ)",
                                data=buf.getvalue(),
                                file_name=fname_w,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key="dl_word_mat",
                            )
                        except Exception as e:
                            st.error(f"สร้างรายงานไม่สำเร็จ: {e}")
                else:
                    st.warning("python-docx ไม่พร้อมใช้งาน")

            # ── Word Consultant ──
            with col_s3:
                if DOCX_AVAILABLE:
                    with st.expander("⚙️ ตั้งค่ารายงานที่ปรึกษา"):
                        sec_start = st.text_input("หมายเลขหัวข้อ", value="4.7", key="cons_sec")
                        intro_txt = st.text_area("บทเกริ่นนำ (ย่อหน้า)", value="", height=80, key="cons_intro")
                    if st.button("📑 Word แบบที่ปรึกษา", use_container_width=True):
                        try:
                            doc = generate_word_report(
                                project_info, ar,
                                report_type='consultant',
                                section_start=sec_start,
                                intro_text=intro_txt,
                            )
                            buf = io.BytesIO()
                            doc.save(buf)
                            buf.seek(0)
                            fname_c = f"{project_name.replace(' ', '_')}_Consultant_{datetime.now().strftime('%Y%m%d')}.docx"
                            st.download_button(
                                "⬇️ ดาวน์โหลด Word (ที่ปรึกษา)",
                                data=buf.getvalue(),
                                file_name=fname_c,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key="dl_word_cons",
                            )
                        except Exception as e:
                            st.error(f"สร้างรายงานไม่สำเร็จ: {e}")
                else:
                    st.warning("python-docx ไม่พร้อมใช้งาน")

    # ── Footer ──
    st.markdown("""
    <div class="footer">
        <b>รศ.ดร.อิทธิพล มีผล</b><br>
        ภาควิชาครุศาสตร์โยธา คณะครุศาสตร์อุตสาหกรรม มจพ.<br>
        <span style="color:#b0bec5">Pavement Structure Cost Analysis v6.0</span>
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
