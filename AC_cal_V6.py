"""
================================================================================
AASHTO 1993 Flexible Pavement Design - Streamlit Application (Version 6)
================================================================================
แอปพลิเคชันสำหรับออกแบบ Flexible Pavement ตามวิธี AASHTO 1993
ปรับปรุงตามมาตรฐานกรมทางหลวง (DOH Thailand)

[V6 Improvements — UX/UI Overhaul + ชื่อวัสดุใหม่]
- Quick Summary Banner (PASS/FAIL) ถาวรอยู่เหนือ Tabs
- ปุ่ม "คำนวณ" ชัดเจน อยู่ใต้ Tab 2 (ชั้นทาง)
- AC Sublayer UI เรียบง่ายขึ้น (ตารางเดียว radio + number_input)
- Sidebar ลดขนาด — ย้ายฐานข้อมูลวัสดุไปอยู่ใน Tab 2
- Export buttons จัดใหม่ 2 แถว + use_container_width
- Refactor _short_name เป็น function เดียว
- Sensitivity Analysis ใน Tab 3 พร้อม label ชัดเจน
- ชื่อวัสดุใหม่ตามมาตรฐาน ทล.:
    CTB → พื้นทางหินคลุกปรับปรุงคุณภาพด้วยปูนซีเมนต์ (Cement Treated Base)
    Wearing → ผิวทางแอสฟัลต์คอนกรีต (AC. Wearing Course)
    Binder → รองผิวทางแอสฟัลต์คอนกรีต (AC. Binder Course)
    Base Course → พื้นทางแอสฟัลต์คอนกรีต (AC. Base Course)

Author: รศ.ดร.อิทธิพล มีผล // ภาควิชาครุศาสตร์โยธา // มจพ.
Version: 6.0
================================================================================
"""

import streamlit as st
import numpy as np
import json
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.font_manager as fm
from io import BytesIO
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT


# ================================================================================
# CUSTOM ROOT-FINDING (แทน scipy.optimize.brentq)
# ================================================================================

def brentq(f, a, b, xtol=1e-12, maxiter=200):
    fa, fb = f(a), f(b)
    if fa * fb > 0:
        raise ValueError(f"f(a) and f(b) must have different signs")
    if abs(fa) < xtol:
        return a
    if abs(fb) < xtol:
        return b
    c, fc = a, fa
    d = e = b - a
    for _ in range(maxiter):
        if fb * fc > 0:
            c, fc = a, fa
            d = e = b - a
        if abs(fc) < abs(fb):
            a, b, c = b, c, b
            fa, fb, fc = fb, fc, fb
        tol1 = 2.0 * 2.2e-16 * abs(b) + 0.5 * xtol
        m = 0.5 * (c - b)
        if abs(m) <= tol1 or fb == 0.0:
            return b
        if abs(e) >= tol1 and abs(fa) > abs(fb):
            s = fb / fa
            if a == c:
                p = 2.0 * m * s
                q = 1.0 - s
            else:
                q = fa / fc
                r = fb / fc
                p = s * (2.0 * m * q * (q - r) - (b - a) * (r - 1.0))
                q = (q - 1.0) * (r - 1.0) * (s - 1.0)
            if p > 0:
                q = -q
            else:
                p = -p
            if 2.0 * p < min(3.0 * m * q - abs(tol1 * q), abs(e * q)):
                e = d
                d = p / q
            else:
                d = m
                e = m
        else:
            d = m
            e = m
        a, fa = b, fb
        if abs(d) > tol1:
            b += d
        else:
            b += tol1 if m > 0 else -tol1
        fb = f(b)
    return b


# ================================================================================
# PAGE CONFIGURATION
# ================================================================================

st.set_page_config(
    page_title="Flexible Pavement Design (AASHTO 1993) v6",
    page_icon="🛣️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ================================================================================
# MATERIAL DATABASE - ตามมาตรฐานกรมทางหลวง (DOH Thailand)
# ================================================================================

MATERIALS = {
    # ============ ชั้นผิวทาง (Surface) ============
    "ผิวทางลาดยาง AC": {
        "layer_coeff": 0.40, "drainage_coeff": 1.0,
        "mr_psi": 362500, "mr_mpa": 2500,
        "layer_type": "surface", "color": "#1C1C1C",
        "short_name": "AC", "english_name": "Asphalt Concrete"
    },
    "ผิวทางลาดยาง PMA": {
        "layer_coeff": 0.40, "drainage_coeff": 1.0,
        "mr_psi": 536500, "mr_mpa": 3700,
        "layer_type": "surface", "color": "#2C2C2C",
        "short_name": "PMA", "english_name": "Polymer Modified Asphalt"
    },

    # ============ ชั้นพื้นทาง (Base) ============
    "พื้นทางหินคลุกปรับปรุงคุณภาพด้วยปูนซีเมนต์ (Cement Treated Base)": {
        "layer_coeff": 0.18, "drainage_coeff": 1.0,
        "mr_psi": 174000, "mr_mpa": 1200,
        "layer_type": "base", "color": "#78909C",
        "short_name": "CTB", "english_name": "Cement Treated Base"
    },
    "พื้นทางหินคลุกผสมซีเมนต์ UCS 24.5 ksc.": {
        "layer_coeff": 0.15, "drainage_coeff": 1.0,
        "mr_psi": 123250, "mr_mpa": 850,
        "layer_type": "base", "color": "#607D8B",
        "short_name": "MOD.CRB", "english_name": "Mod.Crushed Rock Base"
    },
    "พื้นทางหินคลุก CBR 80%": {
        "layer_coeff": 0.13, "drainage_coeff": 1.0,
        "mr_psi": 50750, "mr_mpa": 350,
        "layer_type": "base", "color": "#795548",
        "short_name": "CAB", "english_name": "Crushed Rock Base"
    },
    "พื้นทางดินซีเมนต์ UCS 17.5 ksc.": {
        "layer_coeff": 0.13, "drainage_coeff": 1.0,
        "mr_psi": 50750, "mr_mpa": 350,
        "layer_type": "base", "color": "#8D6E63",
        "short_name": "SCB", "english_name": "Soil Cement Base"
    },
    "พื้นทางวัสดุหมุนเวียน (Recycling)": {
        "layer_coeff": 0.15, "drainage_coeff": 1.0,
        "mr_psi": 123250, "mr_mpa": 850,
        "layer_type": "base", "color": "#5D4037",
        "short_name": "RAP", "english_name": "Recycled Asphalt Pavement"
    },

    # ============ ชั้นรองพื้นทาง (Subbase) ============
    "รองพื้นทางวัสดุมวลรวม CBR 25%": {
        "layer_coeff": 0.10, "drainage_coeff": 1.0,
        "mr_psi": 21750, "mr_mpa": 150,
        "layer_type": "subbase", "color": "#FFB74D",
        "short_name": "GSB", "english_name": "Aggregate Subbase"
    },

    # ============ วัสดุคัดเลือก (Selected Material) ============
    "วัสดุคัดเลือก ก": {
        "layer_coeff": 0.08, "drainage_coeff": 1.0,
        "mr_psi": 14504, "mr_mpa": 100,
        "layer_type": "selected", "color": "#FFF176",
        "short_name": "SM-A", "english_name": "Selected Material"
    },

    # ============ ไม่ใช้วัสดุ (Skip layer) ============
    "ไม่ใช้วัสดุคัดเลือก (ใช้ดินทางทรพ)": {
        "layer_coeff": 0.00, "drainage_coeff": 1.0,
        "mr_psi": 0, "mr_mpa": 0,
        "layer_type": "none", "color": "#D7CCC8",
        "short_name": "NONE", "english_name": "None"
    }
}

# ================================================================================
# PRESET STRUCTURES - โครงสร้างมาตรฐาน ทล.
# ================================================================================

PRESET_STRUCTURES = {
    "--- เลือกโครงสร้างมาตรฐาน ---": None,
    "AC + CTB + GSB + SM (มาตรฐานหลัก)": {
        "description": "ผิวทาง AC / พื้นทาง CTB / รองพื้นทาง GSB / วัสดุคัดเลือก",
        "num_layers": 4,
        "layers": [
            {"material": "ผิวทางลาดยาง AC", "thickness_cm": 15.0},
            {"material": "พื้นทางหินคลุกปรับปรุงคุณภาพด้วยปูนซีเมนต์ (Cement Treated Base)", "thickness_cm": 15.0},
            {"material": "รองพื้นทางวัสดุมวลรวม CBR 25%", "thickness_cm": 15.0},
            {"material": "วัสดุคัดเลือก ก", "thickness_cm": 30.0},
        ]
    },
    "AC + MOD.CRB + GSB + SM": {
        "description": "ผิวทาง AC / หินคลุกผสมซีเมนต์ / รองพื้นทาง GSB / วัสดุคัดเลือก",
        "num_layers": 4,
        "layers": [
            {"material": "ผิวทางลาดยาง AC", "thickness_cm": 15.0},
            {"material": "พื้นทางหินคลุกผสมซีเมนต์ UCS 24.5 ksc.", "thickness_cm": 20.0},
            {"material": "รองพื้นทางวัสดุมวลรวม CBR 25%", "thickness_cm": 15.0},
            {"material": "วัสดุคัดเลือก ก", "thickness_cm": 30.0},
        ]
    },
    "AC + CAB + GSB + SM": {
        "description": "ผิวทาง AC / หินคลุก CBR 80% / รองพื้นทาง GSB / วัสดุคัดเลือก",
        "num_layers": 4,
        "layers": [
            {"material": "ผิวทางลาดยาง AC", "thickness_cm": 15.0},
            {"material": "พื้นทางหินคลุก CBR 80%", "thickness_cm": 20.0},
            {"material": "รองพื้นทางวัสดุมวลรวม CBR 25%", "thickness_cm": 15.0},
            {"material": "วัสดุคัดเลือก ก", "thickness_cm": 30.0},
        ]
    },
    "AC + SCB + GSB + SM": {
        "description": "ผิวทาง AC / ดินซีเมนต์ / รองพื้นทาง GSB / วัสดุคัดเลือก",
        "num_layers": 4,
        "layers": [
            {"material": "ผิวทางลาดยาง AC", "thickness_cm": 15.0},
            {"material": "พื้นทางดินซีเมนต์ UCS 17.5 ksc.", "thickness_cm": 20.0},
            {"material": "รองพื้นทางวัสดุมวลรวม CBR 25%", "thickness_cm": 15.0},
            {"material": "วัสดุคัดเลือก ก", "thickness_cm": 30.0},
        ]
    },
    "AC + CTB + GSB (ไม่ใช้ SM)": {
        "description": "ผิวทาง AC / พื้นทาง CTB / รองพื้นทาง GSB (ไม่มีวัสดุคัดเลือก)",
        "num_layers": 3,
        "layers": [
            {"material": "ผิวทางลาดยาง AC", "thickness_cm": 15.0},
            {"material": "พื้นทางหินคลุกปรับปรุงคุณภาพด้วยปูนซีเมนต์ (Cement Treated Base)", "thickness_cm": 20.0},
            {"material": "รองพื้นทางวัสดุมวลรวม CBR 25%", "thickness_cm": 20.0},
        ]
    },
    "PMA + CTB + GSB + SM": {
        "description": "ผิวทาง PMA / พื้นทาง CTB / รองพื้นทาง GSB / วัสดุคัดเลือก",
        "num_layers": 4,
        "layers": [
            {"material": "ผิวทางลาดยาง PMA", "thickness_cm": 15.0},
            {"material": "พื้นทางหินคลุกปรับปรุงคุณภาพด้วยปูนซีเมนต์ (Cement Treated Base)", "thickness_cm": 15.0},
            {"material": "รองพื้นทางวัสดุมวลรวม CBR 25%", "thickness_cm": 15.0},
            {"material": "วัสดุคัดเลือก ก", "thickness_cm": 30.0},
        ]
    },
    "AC + RAP + GSB + SM": {
        "description": "ผิวทาง AC / พื้นทาง Recycling / รองพื้นทาง GSB / วัสดุคัดเลือก",
        "num_layers": 4,
        "layers": [
            {"material": "ผิวทางลาดยาง AC", "thickness_cm": 15.0},
            {"material": "พื้นทางวัสดุหมุนเวียน (Recycling)", "thickness_cm": 20.0},
            {"material": "รองพื้นทางวัสดุมวลรวม CBR 25%", "thickness_cm": 15.0},
            {"material": "วัสดุคัดเลือก ก", "thickness_cm": 30.0},
        ]
    },
}

# ================================================================================
# RELIABILITY TABLE
# ================================================================================

RELIABILITY_ZR = {
    50: -0.000, 60: -0.253, 70: -0.524, 75: -0.674, 80: -0.841,
    85: -1.037, 90: -1.282, 91: -1.340, 92: -1.405, 93: -1.476,
    94: -1.555, 95: -1.645, 96: -1.751, 97: -1.881, 98: -2.054,
    99: -2.327, 99.9: -3.090
}

# ================================================================================
# DRAINAGE COEFFICIENT TABLE
# ================================================================================

DRAINAGE_TABLE = {
    "Excellent": {"description": "ระบายน้ำดีเยี่ยม (< 2 ชม.)",
                  "values": {"<1%": 1.40, "1-5%": 1.35, "5-25%": 1.30, ">25%": 1.20}},
    "Good":      {"description": "ระบายน้ำดี (1 วัน)",
                  "values": {"<1%": 1.35, "1-5%": 1.25, "5-25%": 1.15, ">25%": 1.00}},
    "Fair":      {"description": "ระบายน้ำพอใช้ (1 สัปดาห์)",
                  "values": {"<1%": 1.25, "1-5%": 1.15, "5-25%": 1.05, ">25%": 0.80}},
    "Poor":      {"description": "ระบายน้ำไม่ดี (1 เดือน)",
                  "values": {"<1%": 1.15, "1-5%": 1.05, "5-25%": 0.80, ">25%": 0.60}},
    "Very Poor": {"description": "ระบายน้ำไม่ดีมาก (ไม่ระบาย)",
                  "values": {"<1%": 1.05, "1-5%": 0.80, "5-25%": 0.60, ">25%": 0.40}},
}

# DOH AC Sublayer Thickness Standards (mm)
DOH_THICKNESS_STANDARDS = {
    "ผิวทางแอสฟัลต์คอนกรีต (AC. Wearing Course)":    [40, 45, 50, 55, 60, 65, 70],
    "รองผิวทางแอสฟัลต์คอนกรีต (AC. Binder Course)":  [40, 45, 50, 55, 60, 65, 70, 75, 80],
    "พื้นทางแอสฟัลต์คอนกรีต (AC. Base Course)":      [0, 70, 75, 80, 85, 90, 95, 100]
}

# ================================================================================
# HELPER: ชื่อวัสดุย่อ (Refactored — เดียวกันทั้งระบบ)
# ================================================================================

def short_material_name(mat_name: str) -> str:
    """แปลงชื่อวัสดุยาวเป็นชื่อย่อสำหรับแสดงในตาราง/รายงาน"""
    mapping = {
        "พื้นทางหินคลุกปรับปรุงคุณภาพด้วยปูนซีเมนต์ (Cement Treated Base)": "หินคลุกปรับปรุงคุณภาพด้วยปูนซีเมนต์ (CTB)",
        "พื้นทางหินคลุกผสมซีเมนต์ UCS 24.5 ksc.":   "หินคลุกผสมซีเมนต์ UCS ≥ 24.5 ksc",
        "พื้นทางหินคลุก CBR 80%":                    "หินคลุก CBR ≥ 80%",
        "พื้นทางดินซีเมนต์ UCS 17.5 ksc.":          "ดินซีเมนต์ UCS ≥ 17.5 ksc",
        "พื้นทางวัสดุหมุนเวียน (Recycling)":         "วัสดุหมุนเวียน (Recycling)",
        "รองพื้นทางวัสดุมวลรวม CBR 25%":             "รองพื้นทางวัสดุมวลรวม CBR ≥ 25%",
        "ผิวทางลาดยาง AC":                           "ผิวทางลาดยาง AC",
        "ผิวทางลาดยาง PMA":                          "ผิวทางลาดยาง PMA",
        "วัสดุคัดเลือก ก":                           "วัสดุคัดเลือก ก",
    }
    return mapping.get(mat_name, mat_name)


# ================================================================================
# CORE CALCULATION FUNCTIONS
# ================================================================================

def aashto_1993_equation(SN, W18, Zr, So, delta_psi, Mr):
    log_W18 = np.log10(W18)
    term1 = Zr * So
    term2 = 9.36 * np.log10(SN + 1) - 0.20
    numerator = np.log10(delta_psi / (4.2 - 1.5))
    denominator = 0.4 + (1094 / ((SN + 1) ** 5.19))
    term3 = numerator / denominator
    term4 = 2.32 * np.log10(Mr) - 8.07
    right_side = term1 + term2 + term3 + term4
    return right_side - log_W18


def calculate_sn_for_layer(W18, Zr, So, delta_psi, Mr):
    def f(SN):
        return aashto_1993_equation(SN, W18, Zr, So, delta_psi, Mr)
    try:
        return round(brentq(f, 0.01, 25.0, xtol=1e-6, maxiter=100), 2)
    except ValueError:
        return None


def calculate_w18_supported(SN, Zr, So, delta_psi, Mr):
    term1 = Zr * So
    term2 = 9.36 * np.log10(SN + 1) - 0.20
    numerator = np.log10(delta_psi / (4.2 - 1.5))
    denominator = 0.4 + (1094 / ((SN + 1) ** 5.19))
    term3 = numerator / denominator
    term4 = 2.32 * np.log10(Mr) - 8.07
    log_W18 = term1 + term2 + term3 + term4
    return 10 ** log_W18


def calculate_layer_thicknesses(W18, Zr, So, delta_psi, subgrade_mr, layers, ac_sublayers=None):
    results = {
        'layers': [], 'sn_values': [], 'subgrade_mr': subgrade_mr,
        'total_sn_required': None, 'total_sn_provided': 0,
        'ac_sublayers': ac_sublayers, 'warnings': []
    }

    active_layers = [l for l in layers if l['material'] != "ไม่ใช้วัสดุคัดเลือก (ใช้ดินทางทรพ)"]
    if not active_layers:
        results['warnings'].append("⚠️ ไม่มีชั้นทางที่ active")
        return results

    num_layers = len(active_layers)
    sn_values = []

    for i in range(num_layers - 1):
        mr_current = MATERIALS[active_layers[i]['material']]['mr_psi']
        mr_next = MATERIALS[active_layers[i + 1]['material']]['mr_psi']
        if mr_current < mr_next:
            results['warnings'].append(
                f"⚠️ ชั้นที่ {i+1} มีค่า Mr = {mr_current:,} psi ต่ำกว่าชั้นที่ {i+2} ที่มี Mr = {mr_next:,} psi "
                f"— ปกติชั้นบนควรมีค่า Mr สูงกว่าชั้นล่าง"
            )

    for i in range(num_layers):
        if i == num_layers - 1:
            mr_below = subgrade_mr
        else:
            mat_below = MATERIALS[active_layers[i + 1]['material']]
            mr_below = mat_below['mr_psi']
        sn_i = calculate_sn_for_layer(W18, Zr, So, delta_psi, mr_below)
        if sn_i is None:
            results['warnings'].append(
                f"⚠️ ไม่สามารถคำนวณ SN สำหรับชั้นที่ {i+1} ได้ — ค่า W18 อาจสูงเกินไป"
            )
        sn_values.append({'layer_index': i + 1, 'mr_below': mr_below, 'sn_required': sn_i})

    results['sn_values'] = sn_values
    results['total_sn_required'] = calculate_sn_for_layer(W18, Zr, So, delta_psi, subgrade_mr)

    if results['total_sn_required'] is None:
        results['warnings'].append("⚠️ ไม่สามารถคำนวณ SN_required ได้ — ลองปรับ W18, Reliability หรือ CBR")

    cumulative_sn = 0
    for i, layer in enumerate(active_layers):
        mat = MATERIALS[layer['material']]
        a_i = layer.get('layer_coeff', mat['layer_coeff'])
        m_i = layer.get('drainage_coeff', 1.0)
        sn_required_at_layer = sn_values[i]['sn_required'] if sn_values[i]['sn_required'] else 0

        if a_i > 0 and m_i > 0:
            remaining_sn = max(0, sn_required_at_layer - cumulative_sn)
            min_thickness_inch = remaining_sn / (a_i * m_i)
            min_thickness_cm = min_thickness_inch * 2.54
        else:
            min_thickness_inch = 0
            min_thickness_cm = 0

        design_thickness_cm = layer['thickness_cm']
        design_thickness_inch = design_thickness_cm / 2.54
        sn_contribution = a_i * design_thickness_inch * m_i
        cumulative_sn += sn_contribution
        is_ok = design_thickness_cm >= min_thickness_cm

        layer_ac_sublayers = ac_sublayers if i == 0 and ac_sublayers is not None else None

        results['layers'].append({
            'layer_no': i + 1,
            'material': layer['material'],
            'short_name': mat['short_name'],
            'english_name': mat.get('english_name', mat['short_name']),
            'mr_psi': mat['mr_psi'], 'mr_mpa': mat['mr_mpa'],
            'a_i': a_i, 'm_i': m_i,
            'sn_required_at_layer': sn_required_at_layer,
            'min_thickness_inch': round(min_thickness_inch, 2),
            'min_thickness_cm': round(min_thickness_cm, 1),
            'design_thickness_cm': design_thickness_cm,
            'design_thickness_inch': round(design_thickness_inch, 2),
            'sn_contribution': round(sn_contribution, 4),
            'cumulative_sn': round(cumulative_sn, 2),
            'is_ok': is_ok, 'color': mat['color'],
            'ac_sublayers': layer_ac_sublayers
        })

    results['total_sn_provided'] = round(cumulative_sn, 2)
    return results


def check_design(sn_required, sn_provided):
    if sn_required is None:
        return {'status': 'ERROR', 'passed': False,
                'message': 'ไม่สามารถคำนวณ SN_required ได้', 'safety_margin': 0}
    safety_margin = sn_provided - sn_required
    passed = sn_provided >= sn_required
    return {
        'status': 'OK' if passed else 'NG',
        'passed': passed,
        'safety_margin': round(safety_margin, 2),
        'message': f"SN_provided ({sn_provided:.2f}) {'≥' if passed else '<'} SN_required ({sn_required:.2f})"
    }


# ================================================================================
# SENSITIVITY ANALYSIS
# ================================================================================

def plot_sensitivity_cbr(W18, Zr, So, delta_psi, current_cbr):
    cbr_range = np.linspace(2, 20, 50)
    sn_values = []
    for cbr in cbr_range:
        mr = 1500 * cbr
        sn = calculate_sn_for_layer(W18, Zr, So, delta_psi, mr)
        sn_values.append(sn if sn else np.nan)

    fig, ax = plt.subplots(figsize=(7, 4))
    ax.plot(cbr_range, sn_values, 'b-', linewidth=2.5, label='SN required')

    current_mr  = 1500 * current_cbr
    current_sn  = calculate_sn_for_layer(W18, Zr, So, delta_psi, current_mr)
    if current_sn:
        ax.plot(current_cbr, current_sn, 'ro', markersize=12,
                label=f'Current: CBR={current_cbr:.1f}%, SN={current_sn:.2f}')
        # annotate
        ax.annotate(
            f'CBR={current_cbr:.1f}%\nSN={current_sn:.2f}',
            xy=(current_cbr, current_sn),
            xytext=(current_cbr + 1.5, current_sn + 0.3),
            fontsize=9,
            arrowprops=dict(arrowstyle='->', color='red', lw=1.2),
            color='red'
        )

    ax.set_xlabel('CBR (%)', fontsize=12)
    ax.set_ylabel('SN Required', fontsize=12)
    ax.set_title(
        'Sensitivity: SN Required vs CBR\n'
        '(Effect of Subgrade CBR on Required SN)',
        fontsize=11, fontweight='bold'
    )
    ax.legend(fontsize=10)
    ax.grid(True, alpha=0.3)
    ax.set_xlim(left=2)
    try:
        plt.tight_layout()
    except Exception:
        pass
    return fig


def plot_sensitivity_w18(Zr, So, delta_psi, Mr, current_w18):
    w18_range = np.logspace(5, 8.5, 50)
    sn_values = []
    for w18 in w18_range:
        sn = calculate_sn_for_layer(w18, Zr, So, delta_psi, Mr)
        sn_values.append(sn if sn else np.nan)

    fig, ax = plt.subplots(figsize=(7, 4))
    ax.semilogx(w18_range, sn_values, 'g-', linewidth=2.5, label='SN required')

    current_sn = calculate_sn_for_layer(current_w18, Zr, So, delta_psi, Mr)
    if current_sn:
        ax.semilogx(current_w18, current_sn, 'ro', markersize=12,
                    label=f'Current: W18={current_w18/1e6:.2f}M, SN={current_sn:.2f}')
        ax.annotate(
            f'W18={current_w18/1e6:.2f}M\nSN={current_sn:.2f}',
            xy=(current_w18, current_sn),
            xytext=(current_w18 * 0.15, current_sn + 0.4),
            fontsize=9,
            arrowprops=dict(arrowstyle='->', color='red', lw=1.2),
            color='red'
        )

    ax.set_xlabel('W18 (ESALs)', fontsize=12)
    ax.set_ylabel('SN Required', fontsize=12)
    ax.set_title(
        'Sensitivity: SN Required vs W18\n'
        '(Effect of Cumulative Traffic Load on Required SN)',
        fontsize=11, fontweight='bold'
    )
    ax.legend(fontsize=10)
    ax.grid(True, alpha=0.3)
    try:
        plt.tight_layout()
    except Exception:
        pass
    return fig


# ================================================================================
# VISUALIZATION FUNCTIONS
# ================================================================================

def _get_thai_fonts():
    import os
    sys_candidates = [
        ('/usr/share/fonts/truetype/tlwg/Garuda.ttf', '/usr/share/fonts/truetype/tlwg/Garuda-Bold.ttf'),
        ('/usr/share/fonts/opentype/tlwg/Garuda.otf', '/usr/share/fonts/opentype/tlwg/Garuda-Bold.otf'),
        ('/usr/share/fonts/truetype/tlwg/Loma.ttf', '/usr/share/fonts/truetype/tlwg/Loma-Bold.ttf'),
        ('/usr/share/fonts/opentype/tlwg/Loma.otf', '/usr/share/fonts/opentype/tlwg/Loma-Bold.otf'),
        ('/usr/share/fonts/truetype/noto/NotoSansThai-Regular.ttf', '/usr/share/fonts/truetype/noto/NotoSansThai-Bold.ttf'),
    ]
    for reg, bold in sys_candidates:
        if os.path.exists(reg):
            fp_r = fm.FontProperties(fname=reg)
            fp_b = fm.FontProperties(fname=bold) if os.path.exists(bold) else fm.FontProperties(fname=reg)
            return fp_r, fp_b, True
    return (fm.FontProperties(family='DejaVu Sans'),
            fm.FontProperties(family='DejaVu Sans', weight='bold'), False)


@st.cache_resource
def get_cached_thai_fonts():
    return _get_thai_fonts()


def plot_pavement_section(layers_result, subgrade_mr=None, subgrade_cbr=None, lang='en'):
    plt.rcParams['font.family'] = 'DejaVu Sans'
    thai_font = thai_font_bold = None
    has_thai = False
    if lang == 'th':
        thai_font, thai_font_bold, has_thai = get_cached_thai_fonts()
        if not has_thai:
            lang = 'en'

    def _fp(bold=False):
        if has_thai:
            return {'fontproperties': thai_font_bold if bold else thai_font}
        return {}

    if not layers_result:
        fig, ax = plt.subplots(figsize=(12, 8))
        ax.text(0.5, 0.5, 'No layers defined', ha='center', va='center', fontsize=14)
        ax.axis('off')
        return fig

    valid_layers = [l for l in layers_result if l.get('design_thickness_cm', 0) > 0]
    if not valid_layers:
        fig, ax = plt.subplots(figsize=(12, 8))
        ax.text(0.5, 0.5, 'No valid layers', ha='center', va='center', fontsize=14)
        ax.axis('off')
        return fig

    # Expand AC sublayers — ใช้ชื่อใหม่
    expanded_layers = []
    for layer in valid_layers:
        ac_sub = layer.get('ac_sublayers', None)
        if ac_sub is not None and layer['layer_no'] == 1:
            sub_info = [
                ('wearing', '#1C1C1C',
                 'ผิวทางแอสฟัลต์คอนกรีต (AC. Wearing Course)',
                 'AC. Wearing Course'),
                ('binder',  '#333333',
                 'รองผิวทางแอสฟัลต์คอนกรีต (AC. Binder Course)',
                 'AC. Binder Course'),
                ('base',    '#4A4A4A',
                 'พื้นทางแอสฟัลต์คอนกรีต (AC. Base Course)',
                 'AC. Base Course'),
            ]
            for key, color, th_name, en_name in sub_info:
                if ac_sub[key] > 0:
                    expanded_layers.append({
                        'design_thickness_cm': ac_sub[key],
                        'material': th_name if lang == 'th' else en_name,
                        'english_name': en_name, 'short_name': key[:2].upper() + 'C',
                        'color': color, 'mr_mpa': layer['mr_mpa'], 'is_sublayer': True
                    })
        else:
            expanded_layers.append(layer)

    draw_layers = expanded_layers
    total_thickness = sum(l['design_thickness_cm'] for l in draw_layers)

    fig, ax = plt.subplots(figsize=(12, 9))
    width = 3
    x_center = 7
    x_start = x_center - width / 2
    min_display_height = 6
    display_heights = [max(l['design_thickness_cm'], min_display_height) for l in draw_layers]
    total_display = sum(display_heights)
    dark_colors = ['#1C1C1C', '#2C2C2C', '#333333', '#4A4A4A', '#78909C', '#607D8B',
                   '#795548', '#8D6E63', '#5D4037', '#6D4C41', '#455A64']

    y_current = total_display
    for i, layer in enumerate(draw_layers):
        thickness = layer['design_thickness_cm']
        display_h = display_heights[i]
        color = layer.get('color', '#CCCCCC')
        e_mpa = layer.get('mr_mpa', 0)
        is_sublayer = layer.get('is_sublayer', False)

        if lang == 'th':
            name = layer.get('material', layer.get('short_name', f'Layer {i+1}'))
        else:
            name = layer.get('english_name', layer.get('short_name', f'Layer {i+1}'))

        y_bottom = y_current - display_h
        ls = '--' if is_sublayer else '-'
        lw = 1 if is_sublayer else 2
        rect = mpatches.Rectangle((x_start, y_bottom), width, display_h,
                                  linewidth=lw, linestyle=ls, edgecolor='black', facecolor=color)
        ax.add_patch(rect)
        yc = y_bottom + display_h / 2
        tc = 'white' if color in dark_colors else 'black'
        fs_center = 14 if is_sublayer else 16
        ax.text(x_center, yc, f'{thickness:.0f} cm',
                ha='center', va='center', fontsize=fs_center, fontweight='bold', color=tc)
        fs_name = 12 if is_sublayer else 14
        ax.text(x_start - 0.5, yc, name,
                ha='right', va='center', fontsize=fs_name, fontweight='bold', color='black', **_fp(True))
        if e_mpa and e_mpa > 0 and not is_sublayer:
            ax.text(x_start + width + 0.5, yc, f'E = {e_mpa:,.0f} MPa',
                    ha='left', va='center', fontsize=12, color='#0066CC')
        y_current = y_bottom

    # Subgrade
    sg_h = 6
    sg_yb = -sg_h
    ax.add_patch(mpatches.Rectangle((x_start, sg_yb), width, sg_h,
        linewidth=2, edgecolor='black', facecolor='#D7CCC8', hatch='///'))
    text_box_h = 3.5
    text_box_w = width * 0.85
    ax.add_patch(mpatches.FancyBboxPatch(
        (x_center - text_box_w / 2, sg_yb + (sg_h - text_box_h) / 2),
        text_box_w, text_box_h, boxstyle="round,pad=0.2",
        facecolor='#EFEBE9', edgecolor='#8D6E63', linewidth=1.5, alpha=0.95))

    sg_label = 'ดินเดิม (Subgrade)' if lang == 'th' else 'Subgrade'
    if subgrade_cbr:
        sg_label += f'\nCBR = {subgrade_cbr:.1f}%'
    ax.text(x_center, sg_yb + sg_h / 2, sg_label,
            ha='center', va='center', fontsize=12, fontweight='bold', color='#5D4037', **_fp(True))
    if subgrade_mr:
        ax.text(x_start + width + 0.5, sg_yb + sg_h / 2, f'Mr = {subgrade_mr:,} psi',
                ha='left', va='center', fontsize=12, color='#0066CC')

    # Total thickness arrow
    ax.annotate('', xy=(x_start + width + 3.5, total_display),
                xytext=(x_start + width + 3.5, 0),
                arrowprops=dict(arrowstyle='<->', color='red', lw=2))
    total_label = f'รวม\n{total_thickness:.0f} cm' if lang == 'th' else f'Total\n{total_thickness:.0f} cm'
    ax.text(x_start + width + 4, total_display / 2, total_label,
            ha='left', va='center', fontsize=14, color='red', fontweight='bold', **_fp(True))

    margin = 10
    ax.set_xlim(0, 15)
    ax.set_ylim(-sg_h - 4, total_display + margin)
    ax.axis('off')

    title_text = 'รูปตัดโครงสร้างชั้นทาง' if lang == 'th' else 'Pavement Structure'
    ax.set_title(title_text, fontsize=20, fontweight='bold', pad=20, **_fp(True))

    box_text = (f'ความหนารวมโครงสร้างชั้นทาง: {total_thickness:.0f} cm'
                if lang == 'th' else f'Total Pavement Thickness: {total_thickness:.0f} cm')
    ax.text(x_center, -sg_h - 2, box_text,
            ha='center', va='center', fontsize=15, fontweight='bold', **_fp(True),
            bbox=dict(boxstyle='round', facecolor='lightyellow', alpha=0.9, edgecolor='orange'))
    try:
        plt.tight_layout()
    except Exception:
        pass
    return fig


def get_figure_as_bytes(fig):
    buf = BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    buf.seek(0)
    return buf


# ================================================================================
# WORD EXPORT FUNCTIONS
# ================================================================================

def set_thai_font(run, size_pt=15, bold=False):
    run.font.name = 'TH SarabunPSK'
    run.font.size = Pt(size_pt)
    run.bold = bold
    try:
        run._element.rPr.rFonts.set(
            '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}cs', 'TH SarabunPSK')
    except Exception:
        pass


def set_equation_font(run, size_pt=11, bold=False, italic=True):
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic


def add_thai_paragraph(doc, text, size_pt=15, bold=False, alignment=None):
    para = doc.add_paragraph()
    if alignment:
        para.alignment = alignment
    run = para.add_run(text)
    set_thai_font(run, size_pt, bold)
    return para


def add_equation_paragraph(doc, text, size_pt=11, bold=False, italic=True):
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = para.add_run(text)
    set_equation_font(run, size_pt, bold, italic)
    return para


def add_table_header_shading(cell, fill_hex='BDD7EE'):
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    shading = OxmlElement('w:shd')
    shading.set(qn('w:val'), 'clear')
    shading.set(qn('w:color'), 'auto')
    shading.set(qn('w:fill'), fill_hex)
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_pr.append(shading)


def create_word_report(project_title, inputs, calc_results, design_check, fig, report_settings=None):
    """รายงาน Word แบบย่อ"""
    doc = Document()

    title = doc.add_heading('รายงานการออกแบบ Flexible Pavement', level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        set_thai_font(run, size_pt=24, bold=True)

    heading1 = doc.add_heading(f'โครงการ: {project_title}', level=1)
    for run in heading1.runs:
        set_thai_font(run, size_pt=18, bold=True)

    add_thai_paragraph(doc, f'วันที่ออกแบบ: {datetime.now().strftime("%d/%m/%Y %H:%M")}', size_pt=15)

    # Section 1
    h2 = doc.add_heading('1. วิธีการออกแบบ', level=2)
    for run in h2.runs:
        set_thai_font(run, size_pt=16, bold=True)
    add_thai_paragraph(doc,
        'การออกแบบโครงสร้างถนนใช้วิธี AASHTO 1993 Guide for Design of Pavement Structures '
        'ตามมาตรฐานกรมทางหลวง โดยใช้สมการหลักดังนี้:', size_pt=15)
    add_equation_paragraph(doc,
        'log₁₀(W₁₈) = Zᵣ·Sₒ + 9.36·log₁₀(SN+1) - 0.20 + '
        'log₁₀(ΔPSI/2.7) / [0.4 + 1094/(SN+1)⁵·¹⁹] + 2.32·log₁₀(Mᵣ) - 8.07',
        size_pt=11, italic=True)

    # Section 2: Inputs
    h2_2 = doc.add_heading('2. ข้อมูลนำเข้า (Design Inputs)', level=2)
    for run in h2_2.runs:
        set_thai_font(run, size_pt=16, bold=True)

    input_table = doc.add_table(rows=1, cols=3)
    input_table.style = 'Table Grid'
    for i, h in enumerate(['พารามิเตอร์', 'ค่า', 'หน่วย']):
        cell = input_table.rows[0].cells[i]
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        set_thai_font(r, size_pt=15, bold=True)
        add_table_header_shading(cell)

    for param, value, unit in [
        ('Design ESALs (W₁₈)', f'{inputs["W18"]:,.0f}', '18-kip ESAL'),
        ('Reliability (R)', f'{inputs["reliability"]}', '%'),
        ('Zᵣ', f'{inputs["Zr"]:.3f}', '-'),
        ('Sₒ', f'{inputs["So"]:.2f}', '-'),
        ('P₀', f'{inputs["P0"]:.1f}', '-'),
        ('Pₜ', f'{inputs["Pt"]:.1f}', '-'),
        ('ΔPSI', f'{inputs["delta_psi"]:.1f}', '-'),
        ('CBR ดินเดิม', f'{inputs.get("CBR", "-")}', '%'),
        ('Mᵣ = 1500×CBR', f'{inputs["Mr"]:,.0f}', 'psi'),
    ]:
        row = input_table.add_row()
        for j, val in enumerate([param, value, unit]):
            row.cells[j].text = ''
            p = row.cells[j].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j != 0 else WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(val)
            set_thai_font(r, size_pt=15)

    # Section 3: Material Properties
    h2_3 = doc.add_heading('3. คุณสมบัติวัสดุชั้นทาง', level=2)
    for run in h2_3.runs:
        set_thai_font(run, size_pt=16, bold=True)

    mat_table = doc.add_table(rows=1, cols=6)
    mat_table.style = 'Table Grid'
    for i, h in enumerate(['ชั้น', 'วัสดุ', 'aᵢ', 'mᵢ', 'Mᵣ (psi)', 'E (MPa)']):
        cell = mat_table.rows[0].cells[i]
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        set_thai_font(r, size_pt=15, bold=True)
        add_table_header_shading(cell)

    for layer in calc_results['layers']:
        row = mat_table.add_row()
        for j, val in enumerate([
            str(layer['layer_no']), short_material_name(layer['material']),
            f'{layer["a_i"]:.2f}', f'{layer["m_i"]:.2f}',
            f'{layer["mr_psi"]:,}', f'{layer["mr_mpa"]:,}'
        ]):
            row.cells[j].text = ''
            p = row.cells[j].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j != 1 else WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(val)
            set_thai_font(r, size_pt=15)

    # AC Sublayer breakdown
    ac_sub = calc_results.get('ac_sublayers', None)
    if ac_sub is not None:
        doc.add_paragraph()
        add_thai_paragraph(doc, 'รายละเอียดชั้นย่อยผิวทาง AC:', size_pt=15, bold=True)
        sub_table = doc.add_table(rows=1, cols=3)
        sub_table.style = 'Table Grid'
        for i, h in enumerate(['ชั้นย่อย', 'ความหนา (cm)', 'ความหนา (mm)']):
            cell = sub_table.rows[0].cells[i]
            cell.text = ''
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(h)
            set_thai_font(r, size_pt=15, bold=True)
            add_table_header_shading(cell)
        for name, thick_cm in [
            ('ผิวทางแอสฟัลต์คอนกรีต (AC. Wearing Course)', ac_sub['wearing']),
            ('รองผิวทางแอสฟัลต์คอนกรีต (AC. Binder Course)', ac_sub['binder']),
            ('พื้นทางแอสฟัลต์คอนกรีต (AC. Base Course)', ac_sub['base']),
            ('รวม', ac_sub['total']),
        ]:
            row = sub_table.add_row()
            for j, val in enumerate([name, f'{thick_cm:.1f}', f'{thick_cm*10:.0f}']):
                row.cells[j].text = ''
                p = row.cells[j].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j != 0 else WD_ALIGN_PARAGRAPH.LEFT
                r = p.add_run(val)
                set_thai_font(r, size_pt=15)

    # Section 4: Step-by-step
    h2_4 = doc.add_heading('4. การคำนวณความหนาแต่ละชั้น', level=2)
    for run in h2_4.runs:
        set_thai_font(run, size_pt=16, bold=True)

    for layer in calc_results['layers']:
        doc.add_paragraph()
        ln   = layer['layer_no']
        a_i  = layer['a_i']
        m_i  = layer['m_i']
        d_in = layer['design_thickness_inch']
        d_cm = layer['design_thickness_cm']
        sn_at     = layer['sn_required_at_layer']
        d_min_in  = layer['min_thickness_inch']
        d_min_cm  = layer['min_thickness_cm']
        sn_cont   = layer['sn_contribution']
        sn_cum    = layer['cumulative_sn']
        is_ok     = layer['is_ok']

        add_thai_paragraph(doc,
            f'ชั้นที่ {ln}: {short_material_name(layer["material"])}',
            size_pt=15, bold=True)
        add_thai_paragraph(doc,
            f'  • Mr = {layer["mr_psi"]:,} psi = {layer["mr_mpa"]:,} MPa\n'
            f'  • a_{ln} = {a_i:.2f}   m_{ln} = {m_i:.2f}',
            size_pt=15)

        # --- SN จากสมการ AASHTO (Times New Roman 11pt) ---
        add_thai_paragraph(doc, 'การคำนวณ SN:', size_pt=15, bold=True)
        add_equation_paragraph(doc,
            f'จากสมการ AASHTO 1993:   SN_{ln} = {sn_at:.2f}',
            size_pt=11, bold=True, italic=True)

        # --- ความหนาขั้นต่ำ (แสดงเฉพาะสมการทั่วไป) ---
        add_thai_paragraph(doc, 'การคำนวณความหนาขั้นต่ำ:', size_pt=15, bold=True)
        if ln == 1:
            add_equation_paragraph(doc,
                f'D_1 >= SN_1 / (a_1 × m_1)',
                size_pt=11, italic=True)
        else:
            add_equation_paragraph(doc,
                f'D_{ln} >= (SN_{ln} − SN_{ln-1}) / (a_{ln} × m_{ln})',
                size_pt=11, italic=True)

        # --- ความหนาที่เลือก ---
        add_thai_paragraph(doc, 'เลือกใช้ความหนา:', size_pt=15, bold=True)
        add_equation_paragraph(doc,
            f'D_{ln}(design)  =  {d_cm:.0f} cm  ({d_in:.2f} in)',
            size_pt=11, bold=True, italic=False)

        # --- SN contribution + แทนค่าตัวเลข ---
        add_thai_paragraph(doc, 'SN contribution:', size_pt=15, bold=True)
        add_equation_paragraph(doc,
            f'ΔSN_{ln} = a_{ln} × D_{ln} × m_{ln}'
            f'  =  {a_i:.2f} × {d_in:.2f} × {m_i:.2f}  =  {sn_cont:.3f}',
            size_pt=11, italic=True)
        add_equation_paragraph(doc,
            f'ΣSN  =  {sn_cum:.2f}',
            size_pt=11, bold=True, italic=False)

        status = '✓ OK — ความหนาเพียงพอ' if is_ok else f'✗ NG — ต้องเพิ่มอีก {d_min_cm - d_cm:.1f} cm'
        add_equation_paragraph(doc, f'สถานะ: {status}', size_pt=11, bold=True, italic=False)

    # Section 5: SN Summary Table
    h2_5 = doc.add_heading('5. ตารางสรุปการคำนวณ Structural Number', level=2)
    for run in h2_5.runs:
        set_thai_font(run, size_pt=16, bold=True)

    # caption ตารางสรุป SN
    if report_settings:
        tbl_no_sn   = report_settings.get('table_number_sn', '')
        tbl_cap_sn  = report_settings.get('table_caption_sn', 'สรุปผลการคำนวณ Structural Number ของโครงสร้างชั้นทาง')
        if tbl_no_sn:
            p_sn_cap = doc.add_paragraph()
            p_sn_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_sn_cap = p_sn_cap.add_run(f'ตารางที่ {tbl_no_sn}  {tbl_cap_sn}')
            set_thai_font(r_sn_cap, size_pt=15, bold=True)

    sn_table = doc.add_table(rows=1, cols=8)
    sn_table.style = 'Table Grid'
    for i, h in enumerate(['ชั้น', 'วัสดุ', 'aᵢ', 'mᵢ', 'Dᵢ (นิ้ว)', 'Dᵢ (ซม.)', 'ΔSNᵢ', 'ΣSN']):
        cell = sn_table.rows[0].cells[i]
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        set_thai_font(r, size_pt=15, bold=True)
        add_table_header_shading(cell)
    for layer in calc_results['layers']:
        row = sn_table.add_row()
        for j, val in enumerate([
            str(layer['layer_no']), short_material_name(layer['material']),
            f'{layer["a_i"]:.2f}', f'{layer["m_i"]:.2f}',
            f'{layer["design_thickness_inch"]:.2f}', f'{layer["design_thickness_cm"]:.0f}',
            f'{layer["sn_contribution"]:.3f}', f'{layer["cumulative_sn"]:.2f}'
        ]):
            row.cells[j].text = ''
            p = row.cells[j].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j != 1 else WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(val)
            set_thai_font(r, size_pt=15)

    # Section 6: Design Check
    h2_6 = doc.add_heading('6. ผลการตรวจสอบการออกแบบ', level=2)
    for run in h2_6.runs:
        set_thai_font(run, size_pt=16, bold=True)

    sn_req  = calc_results['total_sn_required']
    sn_prov = calc_results['total_sn_provided']
    passed  = design_check['passed']
    result_table = doc.add_table(rows=4, cols=2)
    result_table.style = 'Table Grid'
    result_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (param, value) in enumerate([
        ('SN Required (จากสมการ AASHTO)', f'{sn_req:.2f}'),
        ('SN Provided (จากชั้นทาง)',       f'{sn_prov:.2f}'),
        ('Safety Margin',                  f'{design_check["safety_margin"]:.2f}'),
        ('ผลการตรวจสอบ',                   'ผ่าน (OK)' if passed else 'ไม่ผ่าน (NG)'),
    ]):
        for j, val in enumerate([param, value]):
            result_table.rows[i].cells[j].text = ''
            p = result_table.rows[i].cells[j].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            bold_val = (i == 3)  # ผลการตรวจสอบ bold
            set_thai_font(r, size_pt=15, bold=bold_val)
            if i == 3:  # สีผลการตรวจสอบ
                r.font.color.rgb = RGBColor(0x00, 0x70, 0x00) if passed else RGBColor(0xC0, 0x00, 0x00)

    doc.add_paragraph()
    w18_sup = calculate_w18_supported(
        sn_prov, inputs['Zr'], inputs['So'], inputs['delta_psi'], inputs['Mr'])
    add_thai_paragraph(doc,
        f'W₁₈ ที่โครงสร้างรองรับได้ = {w18_sup/1e6:,.2f} ล้าน ESALs',
        size_pt=15, bold=True)

    doc.add_paragraph()
    if passed:
        summary_text = (f'สรุป: การออกแบบผ่านเกณฑ์ เนื่องจาก SN_provided ({sn_prov:.2f}) '
                        f'≥ SN_required ({sn_req:.2f})')
    else:
        summary_text = (f'สรุป: การออกแบบไม่ผ่านเกณฑ์ เนื่องจาก SN_provided ({sn_prov:.2f}) '
                        f'< SN_required ({sn_req:.2f}) กรุณาปรับเพิ่มความหนาชั้นทาง')
    p_sum = add_thai_paragraph(doc, summary_text, size_pt=15, bold=True)
    for run in p_sum.runs:
        run.font.color.rgb = RGBColor(0x00, 0x70, 0x00) if passed else RGBColor(0xC0, 0x00, 0x00)

    # Section 7: Figure
    h2_7 = doc.add_heading('7. ภาพตัดขวางโครงสร้างถนน', level=2)
    for run in h2_7.runs:
        set_thai_font(run, size_pt=16, bold=True)
    fig_bytes = get_figure_as_bytes(fig)
    doc.add_picture(fig_bytes, width=Inches(6))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Section 8: Summary table
    h2_8 = doc.add_heading('8. สรุปโครงสร้างชั้นทางที่ออกแบบ', level=2)
    for run in h2_8.runs:
        set_thai_font(run, size_pt=16, bold=True)

    _add_summary_layer_table(
        doc, calc_results, inputs, fig,
        set_thai_font_fn=set_thai_font,
        add_table_header_shading_fn=add_table_header_shading,
        caption_text=None,
    )

    doc.add_paragraph()
    add_thai_paragraph(doc,
        'รายงานนี้สร้างโดยแอปพลิเคชัน AASHTO 1993 Flexible Pavement Design v6.0\n'
        'พัฒนาโดย รศ.ดร.อิทธิพล มีผล // ภาควิชาครุศาสตร์โยธา // มจพ.',
        size_pt=12, alignment=WD_ALIGN_PARAGRAPH.CENTER)

    doc_bytes = BytesIO()
    doc.save(doc_bytes)
    doc_bytes.seek(0)
    return doc_bytes


def _add_summary_layer_table(doc, calc_results, inputs, fig,
                             set_thai_font_fn, add_table_header_shading_fn,
                             caption_text=None,
                             caption_bold=True, caption_underline=True,
                             use_run_fn=None):
    """
    ตารางสรุปโครงสร้างชั้นทาง 3 คอลัมน์ (สร้าง XML ระดับ w:tc โดยตรง):
      ซ้าย  (3800 DXA) : รูปตัดขวาง — vMerge ทอดทุกแถวข้อมูล
      กลาง  (1400 DXA) : ความหนา (ซม.)
      ขวา   (3872 DXA) : ชนิดวัสดุ
    Header: สีฟ้าอ่อน BDD7EE, bold, center
    """
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.parser import parse_xml
    from lxml import etree

    CBR_val = inputs.get('CBR', 3.0)
    COL_W   = [3800, 1400, 3872]
    NS      = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    # ---- build data rows ----
    structure_rows = []
    ac_sub  = calc_results.get('ac_sublayers', None)
    layers  = calc_results.get('layers', [])
    first   = layers[0] if layers else None

    if ac_sub is not None and first:
        for key, label in [
            ('wearing', 'ผิวทางแอสฟัลต์คอนกรีต\n(AC. Wearing Course)'),
            ('binder',  'รองผิวทางแอสฟัลต์คอนกรีต\n(AC. Binder Course)'),
            ('base',    'พื้นทางแอสฟัลต์คอนกรีต\n(AC. Base Course)'),
        ]:
            if ac_sub.get(key, 0) > 0:
                structure_rows.append((label, f"{ac_sub[key]:.0f}"))
        for layer in layers[1:]:
            structure_rows.append((short_material_name(layer['material']),
                                   f"{layer['design_thickness_cm']:.0f}"))
    else:
        for layer in layers:
            structure_rows.append((short_material_name(layer['material']),
                                   f"{layer['design_thickness_cm']:.0f}"))

    subgrade_mat   = f"Earth Embankment\nor Subgrade, CBR\u2265\n{CBR_val:.1f} %"
    subgrade_thick = "Existing"
    all_data_rows  = structure_rows + [(subgrade_mat, subgrade_thick)]

    # ---- helper: สร้าง w:tc element ใหม่จาก scratch ----
    def _make_tc_el(width_dxa, vmerge=None, valign='center'):
        """
        สร้าง <w:tc> element ใหม่ พร้อม tcPr (tcW + vMerge? + vAlign)
        vmerge: None | 'restart' | 'continue'
        """
        tc_el = OxmlElement('w:tc')
        tcPr  = OxmlElement('w:tcPr')

        tcW = OxmlElement('w:tcW')
        tcW.set(qn('w:w'),    str(width_dxa))
        tcW.set(qn('w:type'), 'dxa')
        tcPr.append(tcW)

        if vmerge is not None:
            vm = OxmlElement('w:vMerge')
            if vmerge == 'restart':
                vm.set(qn('w:val'), 'restart')
            # 'continue' => ไม่ set val attribute (Word ตีความเป็น continue)
            tcPr.append(vm)

        va = OxmlElement('w:vAlign')
        va.set(qn('w:val'), valign)
        tcPr.append(va)

        tc_el.append(tcPr)
        # empty paragraph (required by OOXML)
        tc_el.append(OxmlElement('w:p'))
        return tc_el

    # ---- helper: หา Cell object จาก tc element (เพื่อใส่ text/picture) ----
    def _wrap_tc(tc_el):
        """wrap raw tc element เป็น CT_Tc เพื่อใช้ python-docx API"""
        from docx.oxml.table import CT_Tc
        # CT_Tc inherit จาก lxml BaseElement แล้วก็เป็น tc_el นั่นเอง
        return tc_el

    # ---- helper: เพิ่ม paragraph ใน tc element ----
    def _tc_add_para(tc_el, text, bold=False, center=True, underline=False):
        # ลบ w:p ว่างที่สร้างไว้แล้ว
        for old_p in tc_el.findall(qn('w:p')):
            tc_el.remove(old_p)
        p_el = OxmlElement('w:p')
        pPr  = OxmlElement('w:pPr')
        jc   = OxmlElement('w:jc')
        jc.set(qn('w:val'), 'center' if center else 'left')
        pPr.append(jc)
        p_el.append(pPr)

        # split by newline → multiple runs with <w:br>
        parts = text.split('\n')
        for p_idx, part in enumerate(parts):
            if p_idx > 0:
                br_r = OxmlElement('w:r')
                br   = OxmlElement('w:br')
                br_r.append(br)
                p_el.append(br_r)
            r_el = OxmlElement('w:r')
            rPr  = OxmlElement('w:rPr')
            rFonts = OxmlElement('w:rFonts')
            rFonts.set(qn('w:ascii'),    'TH SarabunPSK')
            rFonts.set(qn('w:hAnsi'),    'TH SarabunPSK')
            rFonts.set(qn('w:cs'),       'TH SarabunPSK')
            rPr.append(rFonts)
            sz = OxmlElement('w:sz')
            sz.set(qn('w:val'), '30')   # 15pt = 30 half-points
            szCs = OxmlElement('w:szCs')
            szCs.set(qn('w:val'), '30')
            rPr.append(sz)
            rPr.append(szCs)
            if bold:
                rPr.append(OxmlElement('w:b'))
                rPr.append(OxmlElement('w:bCs'))
            if underline:
                u_el = OxmlElement('w:u')
                u_el.set(qn('w:val'), 'single')
                rPr.append(u_el)
            r_el.append(rPr)
            t_el = OxmlElement('w:t')
            t_el.text = part
            if part.startswith(' ') or part.endswith(' '):
                t_el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            r_el.append(t_el)
            p_el.append(r_el)

        tc_el.append(p_el)

    # ---- caption ----
    if caption_text:
        cap_p = doc.add_paragraph()
        cap_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if use_run_fn:
            use_run_fn(cap_p, caption_text, bold=caption_bold, underline=caption_underline)
        else:
            r = cap_p.add_run(caption_text)
            set_thai_font_fn(r, size_pt=15, bold=caption_bold)
            if caption_underline:
                r.underline = True

    # ---- build tbl element ----
    tbl_el = OxmlElement('w:tbl')

    tblPr = OxmlElement('w:tblPr')
    tblStyle = OxmlElement('w:tblStyle')
    tblStyle.set(qn('w:val'), 'TableGrid')
    tblPr.append(tblStyle)
    tblW_el = OxmlElement('w:tblW')
    tblW_el.set(qn('w:w'),    str(sum(COL_W)))
    tblW_el.set(qn('w:type'), 'dxa')
    tblPr.append(tblW_el)
    jc_tbl = OxmlElement('w:jc')
    jc_tbl.set(qn('w:val'), 'center')
    tblPr.append(jc_tbl)
    tbl_el.append(tblPr)

    tblGrid = OxmlElement('w:tblGrid')
    for w in COL_W:
        gc = OxmlElement('w:gridCol')
        gc.set(qn('w:w'), str(w))
        tblGrid.append(gc)
    tbl_el.append(tblGrid)

    # ---- Header row ----
    tr_hdr = OxmlElement('w:tr')
    hdr_labels = ['รายละเอียด', 'หนา (ซม.)', 'ชนิดวัสดุ']
    hdr_tc_els = []
    for j, (label, w) in enumerate(zip(hdr_labels, COL_W)):
        tc_el = _make_tc_el(w, vmerge=None, valign='center')
        _tc_add_para(tc_el, label, bold=True, center=True)
        tr_hdr.append(tc_el)
        hdr_tc_els.append(tc_el)
    tbl_el.append(tr_hdr)

    # ---- Figure bytes ----
    fig_buf = get_figure_as_bytes(fig)

    # ---- Data rows ----
    data_tc_els = []   # เก็บ tc element ที่สร้างเพื่อใส่รูปและ shading
    for i, (mat_name, thick) in enumerate(all_data_rows):
        tr = OxmlElement('w:tr')

        # col0: รูปตัดขวาง (vMerge)
        vm = 'restart' if i == 0 else 'continue'
        tc0 = _make_tc_el(COL_W[0], vmerge=vm, valign='center')
        if i == 0:
            # ลบ w:p ว่าง แล้วใส่ paragraph พร้อมรูป
            for old_p in tc0.findall(qn('w:p')):
                tc0.remove(old_p)
            p_el = OxmlElement('w:p')
            pPr  = OxmlElement('w:pPr')
            jc   = OxmlElement('w:jc')
            jc.set(qn('w:val'), 'center')
            pPr.append(jc)
            p_el.append(pPr)
            tc0.append(p_el)
            # จะใส่รูปภายหลังผ่าน python-docx run API
        tr.append(tc0)

        # col1: ความหนา
        tc1 = _make_tc_el(COL_W[1], vmerge=None, valign='center')
        _tc_add_para(tc1, thick, bold=False, center=True)
        tr.append(tc1)

        # col2: ชนิดวัสดุ
        tc2 = _make_tc_el(COL_W[2], vmerge=None, valign='center')
        _tc_add_para(tc2, mat_name, bold=False, center=True)
        tr.append(tc2)

        tbl_el.append(tr)
        data_tc_els.append((tc0, tc1, tc2))

    # ---- แทรก tbl ใน body ก่อน footer paragraph สุดท้าย ----
    body = doc.element.body
    body.append(tbl_el)

    # ---- ใส่รูปใน tc0 ของ row แรก ผ่าน python-docx (หลัง append เข้า body แล้ว) ----
    tc0_first = data_tc_els[0][0]
    # หา w:p ใน tc0_first
    p_in_tc = tc0_first.find(qn('w:p'))
    if p_in_tc is not None:
        from docx.oxml.ns import nsmap
        from docx.text.paragraph import Paragraph
        from docx.table import _Cell
        # สร้าง run ผ่าน lxml โดยตรงโดยใช้ InlineImage approach
        # ง่ายกว่า: สร้าง temporary doc เพื่อ render picture run แล้วย้าย drawing element
        tmp_doc = Document()
        fig_buf.seek(0)
        tmp_doc.add_picture(fig_buf, width=Inches(2.4))
        # drawing อยู่ใน paragraphs[-1].runs[-1]
        tmp_p    = tmp_doc.paragraphs[-1]._p
        tmp_runs = tmp_p.findall(qn('w:r'))
        drawing_r = None
        for r in tmp_runs:
            if r.find(qn('w:drawing')) is not None:
                drawing_r = r
                break
        if drawing_r is not None:
            from copy import deepcopy
            p_in_tc.append(deepcopy(drawing_r))

    # ---- ใส่ shading header ผ่าน add_table_header_shading_fn ----
    # ต้องหา Cell object จาก python-docx table
    # เนื่องจากตอนนี้ tbl_el อยู่ใน body แล้ว เราหาตารางล่าสุด
    real_tbl = doc.tables[-1]
    for j in range(3):
        add_table_header_shading_fn(real_tbl.rows[0].cells[j], fill_hex='BDD7EE')

    return real_tbl




def _build_structure_rows(calc_results, cbr_val):
    """สร้างรายการชั้นทางสำหรับตารางสรุป (ใช้ร่วมกันทั้ง 2 report)"""
    structure_rows = []
    row_num = 1
    ac_sub = calc_results.get('ac_sublayers', None)
    first_layer = calc_results['layers'][0] if calc_results.get('layers') else None

    if ac_sub is not None and first_layer:
        for key, label in [
            ('wearing', 'ผิวทางแอสฟัลต์คอนกรีต (AC. Wearing Course)'),
            ('binder',  'รองผิวทางแอสฟัลต์คอนกรีต (AC. Binder Course)'),
            ('base',    'พื้นทางแอสฟัลต์คอนกรีต (AC. Base Course)'),
        ]:
            if ac_sub.get(key, 0) > 0:
                structure_rows.append((row_num, label, f"{ac_sub[key]:.0f}"))
                row_num += 1
        for layer in calc_results['layers'][1:]:
            structure_rows.append((row_num, short_material_name(layer['material']),
                                   f"{layer['design_thickness_cm']:.0f}"))
            row_num += 1
    else:
        for layer in calc_results.get('layers', []):
            structure_rows.append((row_num, short_material_name(layer['material']),
                                   f"{layer['design_thickness_cm']:.0f}"))
            row_num += 1

    structure_rows.append((row_num, 'ดินคันทาง', f'CBR ≥ {cbr_val:.1f} %'))
    return structure_rows


def set_thai_distribute(para):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    pPr = para._element.get_or_add_pPr()
    jc = OxmlElement('w:jc')
    jc.set(qn('w:val'), 'thaiDistribute')
    pPr.append(jc)


def create_word_report_intro(project_title, inputs, calc_results, design_check, fig, report_settings):
    """รายงาน Word แบบที่ปรึกษา (รวมกับรายงานอื่น)"""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'TH SarabunPSK'
    style.font.size = Pt(15)
    try:
        style._element.rPr.rFonts.set(qn('w:eastAsia'), 'TH SarabunPSK')
    except Exception:
        pass

    sec_no   = report_settings.get('section_number', '4.4')
    tbl_inp  = report_settings.get('table_number_inputs', '4-8')
    tbl_mat  = report_settings.get('table_number_materials', '4-9')
    tbl_sn   = report_settings.get('table_number_sn', '4-10')
    fig_no   = report_settings.get('figure_number', '4-8')
    sec_title = report_settings.get('section_title', 'การออกแบบผิวทางลาดยาง (Flexible Pavement)')
    tbl_cap_inp  = report_settings.get('table_caption_inputs', 'ค่าพารามิเตอร์ที่ใช้ในการออกแบบผิวทางยืดหยุ่น')
    tbl_cap_mat  = report_settings.get('table_caption_materials', 'ค่าสัมประสิทธิ์และค่าโมดูลัสของวัสดุโครงสร้างชั้นทาง')
    tbl_cap_sn   = report_settings.get('table_caption_sn', 'สรุปผลการคำนวณ Structural Number ของโครงสร้างชั้นทาง')
    fig_cap  = report_settings.get('figure_caption', 'รูปตัดโครงสร้างชั้นทางที่ออกแบบ')

    RED   = RGBColor(0xCC, 0x00, 0x00)
    GREEN = RGBColor(0, 112, 0)
    BLUE  = RGBColor(0x00, 0x47, 0xAB)

    def _run(para, text, size=15, bold=False, italic=False, color=None, underline=False):
        r = para.add_run(text)
        r.font.name = 'TH SarabunPSK'
        r.font.size = Pt(size)
        r.bold = bold
        r.italic = italic
        r.underline = underline
        if color:
            r.font.color.rgb = color
        try:
            r._element.rPr.rFonts.set(qn('w:cs'), 'TH SarabunPSK')
        except Exception:
            pass
        return r

    def _eq(para, text, size=11, bold=False, italic=True):
        """Times New Roman 11pt สำหรับสมการ"""
        r = para.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(size)
        r.bold = bold
        r.italic = italic
        try:
            r._element.rPr.rFonts.set(qn('w:cs'), 'Times New Roman')
        except Exception:
            pass
        return r

    def _eq_para(text, indent_cm=2.0, bold=False, italic=True, align=WD_ALIGN_PARAGRAPH.LEFT):
        """Paragraph สมการ Times New Roman 11pt"""
        p = _para(indent_cm=indent_cm)
        p.alignment = align
        _eq(p, text, bold=bold, italic=italic)
        return p

    def _heading(text, level, size):
        h = doc.add_heading(text, level=level)
        for run in h.runs:
            run.font.name = 'TH SarabunPSK'
            run.font.size = Pt(size)
        return h

    def _para(indent_cm=0, space_before=0, space_after=4):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(indent_cm)
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        return p

    def _tbl_cell(cell, text, align=WD_ALIGN_PARAGRAPH.CENTER, size=15, bold=False, fill=None):
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = align
        r = p.add_run(text)
        set_thai_font(r, size_pt=size, bold=bold)
        if fill:
            add_table_header_shading(cell, fill)

    def _fig_caption(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(text)
        r.font.name = 'TH SarabunPSK'
        r.font.size = Pt(14)
        r.italic = True

    # ===== หัวข้อหลัก =====
    h_main = _heading(f'{sec_no}  {sec_title}', level=2, size=16)

    # ===== เกริ่นนำ =====
    W18 = inputs.get('W18', 0)
    reliability = inputs.get('reliability', 90)
    CBR = inputs.get('CBR', 5.0)
    Mr = inputs.get('Mr', 7500)
    sn_req  = calc_results.get('total_sn_required', 0)
    sn_prov = calc_results.get('total_sn_provided', 0)
    total_thick = sum(l['design_thickness_cm'] for l in calc_results.get('layers', []))
    num_layers  = len(calc_results.get('layers', []))
    passed_txt  = 'ผ่านเกณฑ์' if design_check.get('passed') else 'ไม่ผ่านเกณฑ์'

    p_intro = _para(indent_cm=0, space_before=6)
    p_intro.paragraph_format.first_line_indent = Cm(1.25)
    set_thai_distribute(p_intro)
    _run(p_intro, 'ถนนลาดยางซึ่งประกอบด้วยวัสดุงานทางหลายชนิด การออกแบบโครงสร้างถนนแบบยืดหยุ่น '
         '(Flexible Pavement) ใช้วิธี AASHTO 1993 Guide for Design of Pavement Structures '
         'โดยพิจารณาปัจจัยด้านปริมาณจราจรสะสม ESALs ความน่าเชื่อถือ และคุณสมบัติของดินรองรับ '
         'สำหรับโครงการนี้ที่ปรึกษาได้กำหนดค่าพารามิเตอร์หลักในการออกแบบ ได้แก่ '
         f'ปริมาณ W\u2081\u2088 = ')
    _run(p_intro, f'{W18:,.0f}', bold=True, color=BLUE)
    _run(p_intro, f' 18-kip ESALs ที่ระดับความน่าเชื่อถือ (Reliability) = ')
    _run(p_intro, f'{reliability}', bold=True, color=BLUE)
    _run(p_intro, f' % โดยมีดินเดิมค่า CBR = ')
    _run(p_intro, f'{CBR:.1f}', bold=True, color=BLUE)
    _run(p_intro, f' % (M\u1D63 = ')
    _run(p_intro, f'{Mr:,.0f}', bold=True, color=BLUE)
    _run(p_intro, f' psi) ผลการออกแบบได้โครงสร้างชั้นทาง ')
    _run(p_intro, f'{num_layers}', bold=True, color=BLUE)
    _run(p_intro, f' ชั้น ที่ SN\u200B_required = ')
    _run(p_intro, f'{sn_req:.2f}', bold=True, color=BLUE)
    _run(p_intro, ' และ SN\u200B_provided = ')
    _run(p_intro, f'{sn_prov:.2f}', bold=True, color=BLUE)
    _run(p_intro, f' ความหนารวม ')
    _run(p_intro, f'{total_thick:.0f}', bold=True, color=BLUE)
    _run(p_intro, f' ซม. การออกแบบ')
    _run(p_intro, passed_txt, bold=True, color=GREEN if design_check.get('passed') else RED)
    _run(p_intro, f' ดังแสดงผลการวิเคราะห์ในตารางที่ ')
    _run(p_intro, tbl_inp, bold=True)
    _run(p_intro, ' และตารางที่ ')
    _run(p_intro, tbl_mat, bold=True)
    _run(p_intro, ' และรูปที่ ')
    _run(p_intro, fig_no, bold=True)

    # ===== {sec_no}.1 วิธีการออกแบบ =====
    _heading(f'{sec_no}.1  วิธีการออกแบบ', level=3, size=15)
    p_meth = _para(indent_cm=0, space_before=4)
    p_meth.paragraph_format.first_line_indent = Cm(1.25)
    _run(p_meth, 'การออกแบบโครงสร้างถนนใช้วิธี AASHTO 1993 Guide for Design of Pavement Structures '
         'ตามมาตรฐานกรมทางหลวง โดยใช้สมการหลักดังนี้')

    _eq_para(
        'log10(W18) = Zr·So + 9.36·log10(SN+1) - 0.20\n'
        '                   + log10(ΔPSI/2.7) / [0.4 + 1094/(SN+1)^5.19]\n'
        '                   + 2.32·log10(Mr) - 8.07',
        indent_cm=2.5, italic=True
    )

    # ===== {sec_no}.2 ข้อมูลนำเข้า + ตาราง =====
    _heading(f'{sec_no}.2  ข้อมูลนำเข้า (Design Inputs)', level=3, size=15)
    p_intro2 = _para(indent_cm=0, space_before=4)
    p_intro2.paragraph_format.first_line_indent = Cm(1.25)
    set_thai_distribute(p_intro2)
    _run(p_intro2,
         'ในการออกแบบโครงสร้างถนนยืดหยุ่น การกำหนดค่าพารามิเตอร์นำเข้า (Design Inputs) ถือเป็น'
         'ขั้นตอนสำคัญที่มีผลโดยตรงต่อความถูกต้องและความน่าเชื่อถือของแบบโครงสร้างถนนที่ต้องการ '
         'เนื่องจากค่าพารามิเตอร์แต่ละตัวสะท้อนให้เห็นสภาพการใช้งานจริงของโครงสร้างถนน '
         'ได้แก่ ปริมาณการจราจรตลอดอายุการใช้งาน ระดับความน่าเชื่อถือที่ยอมรับได้ '
         'รวมถึงคุณสมบัติของดินรองรับในพื้นที่โครงการ สำหรับโครงการนี้ '
         'ที่ปรึกษาได้กำหนดค่าพารามิเตอร์หลักที่ใช้ในการออกแบบตามแนวทางของ AASHTO '
         'ซึ่งประกอบด้วยข้อมูลด้านความสามารถในการรับน้ำหนักของโครงสร้างชั้นทาง '
         'ปริมาณจราจรที่โครงสร้างถนนต้องรองรับตลอดอายุการใช้งาน '
         'รวมถึงคุณสมบัติของชั้นดินที่ต้องซ่อมบำรุงหรือปรับปรุงใหม่ '
         f'รวมถึงคุณสมบัติของดินชั้นรองรับ รายละเอียดของค่าพารามิเตอร์ทั้งหมดแสดงในตารางที่ {tbl_inp}')
    p_tbl1_cap = _para(indent_cm=0, space_before=4)
    p_tbl1_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(p_tbl1_cap, f'ตารางที่ {tbl_inp}  {tbl_cap_inp}', bold=True)

    inp_tbl = doc.add_table(rows=1, cols=3)
    inp_tbl.style = 'Table Grid'
    inp_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, h in enumerate(['พารามิเตอร์', 'ค่า', 'หน่วย']):
        _tbl_cell(inp_tbl.rows[0].cells[j], h, bold=True, fill='D9E2F3')
    for param, value, unit in [
        ('Design ESALs (W\u2081\u2088)', f'{W18:,.0f}', '18-kip ESAL'),
        ('Reliability (R)', f'{reliability}', '%'),
        ('Z\u1D63', f'{inputs.get("Zr", 0):.3f}', '-'),
        ('S\u2080', f'{inputs.get("So", 0):.2f}', '-'),
        ('P\u2080', f'{inputs.get("P0", 0):.1f}', '-'),
        ('P\u209C', f'{inputs.get("Pt", 0):.1f}', '-'),
        ('\u0394PSI', f'{inputs.get("delta_psi", 0):.1f}', '-'),
        ('CBR ดินเดิม', f'{CBR:.1f}', '%'),
        ('M\u1D63 = 1,500\u00D7CBR', f'{Mr:,.0f}', 'psi'),
    ]:
        row = inp_tbl.add_row()
        _tbl_cell(row.cells[0], param, align=WD_ALIGN_PARAGRAPH.LEFT)
        _tbl_cell(row.cells[1], value)
        _tbl_cell(row.cells[2], unit)

    # ===== {sec_no}.3 คุณสมบัติวัสดุ + ตาราง =====
    _heading(f'{sec_no}.3  คุณสมบัติวัสดุชั้นทาง', level=3, size=15)
    p_intro3 = _para(indent_cm=0, space_before=4)
    p_intro3.paragraph_format.first_line_indent = Cm(1.25)
    set_thai_distribute(p_intro3)
    _run(p_intro3,
         'วัสดุโครงสร้างชั้นทางแต่ละชนิดมีค่าสัมประสิทธิ์ชั้นทาง (Layer Coefficient) '
         'และค่าสัมประสิทธิ์การระบายน้ำ (Drainage Coefficient) '
         'โดยที่ปรึกษาเลือกใช้วัสดุและแสดงค่าสัมประสิทธิ์รวมถึงค่าโมดูลัสของวัสดุ'
         f'ต่าง ๆ ดังแสดงในตารางที่ {tbl_mat}')
    p_tbl2_cap = _para(indent_cm=0, space_before=4)
    p_tbl2_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(p_tbl2_cap, f'ตารางที่ {tbl_mat}  {tbl_cap_mat}', bold=True)

    mat_tbl = doc.add_table(rows=1, cols=6)
    mat_tbl.style = 'Table Grid'
    mat_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, h in enumerate(['ชั้น', 'วัสดุ', 'a\u1D62', 'm\u1D62', 'M\u1D63 (psi)', 'E (MPa)']):
        _tbl_cell(mat_tbl.rows[0].cells[j], h, bold=True, fill='D9E2F3')
    for layer in calc_results.get('layers', []):
        row = mat_tbl.add_row()
        _tbl_cell(row.cells[0], str(layer['layer_no']))
        _tbl_cell(row.cells[1], short_material_name(layer['material']), align=WD_ALIGN_PARAGRAPH.LEFT)
        _tbl_cell(row.cells[2], f'{layer["a_i"]:.2f}')
        _tbl_cell(row.cells[3], f'{layer["m_i"]:.2f}')
        _tbl_cell(row.cells[4], f'{layer["mr_psi"]:,}')
        _tbl_cell(row.cells[5], f'{layer["mr_mpa"]:,}')

    # AC Sublayer table (ถ้ามี)
    ac_sub = calc_results.get('ac_sublayers', None)
    if ac_sub:
        p_sub = _para(indent_cm=0, space_before=6)
        _run(p_sub, 'รายละเอียดชั้นย่อยผิวทางแอสฟัลต์คอนกรีต:', bold=True)
        sub_tbl = doc.add_table(rows=1, cols=3)
        sub_tbl.style = 'Table Grid'
        sub_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        for j, h in enumerate(['ชั้นย่อย', 'ความหนา (cm)', 'ความหนา (mm)']):
            _tbl_cell(sub_tbl.rows[0].cells[j], h, bold=True, fill='D9E2F3')
        for key, label in [
            ('wearing', 'ผิวทางแอสฟัลต์คอนกรีต (AC. Wearing Course)'),
            ('binder',  'รองผิวทางแอสฟัลต์คอนกรีต (AC. Binder Course)'),
            ('base',    'พื้นทางแอสฟัลต์คอนกรีต (AC. Base Course)'),
            ('total',   'รวม'),
        ]:
            val = ac_sub.get(key, 0)
            if key == 'total' or val > 0:
                row = sub_tbl.add_row()
                _tbl_cell(row.cells[0], label, align=WD_ALIGN_PARAGRAPH.LEFT)
                _tbl_cell(row.cells[1], f'{val:.1f}')
                _tbl_cell(row.cells[2], f'{val*10:.0f}' if key != 'total' else f'{val*10:.0f}')

    # ===== {sec_no}.4 ขั้นตอนการคำนวณ =====
    _heading(f'{sec_no}.4  ขั้นตอนการคำนวณความหนาชั้นทาง', level=3, size=15)
    p_intro4 = _para(indent_cm=0, space_before=4)
    p_intro4.paragraph_format.first_line_indent = Cm(1.25)
    set_thai_distribute(p_intro4)
    _run(p_intro4,
         'การคำนวณความหนาขั้นต่ำของแต่ละชั้น ใช้หลักการว่า Structural Number (SN) ที่จุดใด ๆ '
         'ต้องมากกว่าหรือเท่ากับ SN ที่ต้องการ โดยคำนวณจากค่า M\u1D63 ของชั้นถัดไป')

    for layer in calc_results.get('layers', []):
        sn_at    = layer['sn_required_at_layer']
        layer_no = layer['layer_no']
        a_i      = layer['a_i']
        m_i      = layer['m_i']
        d_in     = layer['design_thickness_inch']
        d_cm     = layer['design_thickness_cm']
        d_min_in = layer['min_thickness_inch']
        d_min_cm = layer['min_thickness_cm']
        sn_cont  = layer['sn_contribution']
        sn_cum   = layer['cumulative_sn']
        is_ok    = layer['is_ok']

        doc.add_paragraph()
        hdr_p = _para(indent_cm=1.0, space_before=6)
        _run(hdr_p, f'ชั้นที่ {layer_no}: {short_material_name(layer["material"])}',
             bold=True, underline=True)

        p_mat = _para(indent_cm=1.5)
        _run(p_mat, 'ข้อมูลวัสดุ:', bold=True)
        p_mat2 = _para(indent_cm=2.0)
        _run(p_mat2,
             f'\u2022 Mr = {layer["mr_psi"]:,} psi  =  {layer["mr_mpa"]:,} MPa\n'
             f'\u2022 Layer Coefficient (a{layer_no}) = {a_i:.2f}\n'
             f'\u2022 Drainage Coefficient (m{layer_no}) = {m_i:.2f}')

        # --- การคำนวณ SN ---
        p_sn = _para(indent_cm=1.5)
        _run(p_sn, 'การคำนวณ SN:', bold=True)
        # SN_N notation (ไม่ใช้ subscript 0)
        _eq_para(f'จากสมการ AASHTO 1993:   SN_{layer_no} = {sn_at:.2f}',
                 indent_cm=2.0, bold=True, italic=True)

        # --- ความหนาขั้นต่ำ (แสดงเฉพาะสมการทั่วไป) ---
        p_th = _para(indent_cm=1.5)
        _run(p_th, 'การคำนวณความหนาขั้นต่ำ:', bold=True)
        if layer_no == 1:
            _eq_para(
                'D_1 >= SN_1 / (a_1 \u00d7 m_1)',
                indent_cm=2.5, italic=True
            )
        else:
            _eq_para(
                f'D_{layer_no} >= (SN_{layer_no} \u2212 SN_{layer_no-1}) / (a_{layer_no} \u00d7 m_{layer_no})',
                indent_cm=2.5, italic=True
            )

        # --- ความหนาที่เลือกใช้ ---
        p_d = _para(indent_cm=1.5)
        _run(p_d, 'เลือกใช้ความหนา:', bold=True)
        _eq_para(
            f'D_{layer_no}(design)  =  {d_cm:.0f} cm  ({d_in:.2f} in)',
            indent_cm=2.5, bold=True, italic=False
        )

        # --- SN contribution พร้อมแทนค่า ---
        p_sn2 = _para(indent_cm=1.5)
        _run(p_sn2, 'SN contribution:', bold=True)
        _eq_para(
            f'\u0394SN_{layer_no} = a_{layer_no} \u00d7 D_{layer_no} \u00d7 m_{layer_no}'
            f'  =  {a_i:.2f} \u00d7 {d_in:.2f} \u00d7 {m_i:.2f}  =  {sn_cont:.3f}',
            indent_cm=2.5, italic=True
        )
        _eq_para(
            f'\u03a3SN  =  {sn_cum:.2f}',
            indent_cm=2.5, bold=True, italic=False
        )

        # --- สถานะ (ไม่แสดงวงเล็บตัวเลข) ---
        status_txt  = '\u2713 OK' if is_ok else '\u2717 NG'
        status_note = '\u0e04\u0e27\u0e32\u0e21\u0e2b\u0e19\u0e32\u0e40\u0e1e\u0e35\u0e22\u0e07\u0e1e\u0e2d' if is_ok else f'\u0e15\u0e49\u0e2d\u0e07\u0e40\u0e1e\u0e34\u0e48\u0e21\u0e04\u0e27\u0e32\u0e21\u0e2b\u0e19\u0e32\u0e2d\u0e35\u0e01 {d_min_cm - d_cm:.1f} cm'
        p_st = _para(indent_cm=2.0)
        _run(p_st, f'\u0e2a\u0e16\u0e32\u0e19\u0e30:  {status_txt}  \u2014  {status_note}',
             bold=True, color=GREEN if is_ok else RED)

    # ===== Section .5 ตารางสรุปการคำนวณ SN =====
    _heading(f'{sec_no}.5  สรุปการคำนวณ Structural Number', level=3, size=15)

    # เกริ่นนำ 4 บรรทัด
    p_sn_intro = _para(indent_cm=0, space_before=4)
    p_sn_intro.paragraph_format.first_line_indent = Cm(1.25)
    set_thai_distribute(p_sn_intro)
    _run(p_sn_intro,
         'เมื่อได้ทำการคำนวณความหนาของแต่ละชั้นทางตามขั้นตอนที่กล่าวมาข้างต้นแล้ว '
         'สามารถสรุปผลการคำนวณ Structural Number (SN) ของโครงสร้างชั้นทางได้ดังนี้ '
         'โดย Structural Number คือค่าที่แสดงถึงความสามารถในการรับน้ำหนักสะสมของโครงสร้างถนน '
         'ซึ่งคำนวณจากผลรวมของค่าสัมประสิทธิ์ชั้นทาง (aᵢ) ค่าสัมประสิทธิ์การระบายน้ำ (mᵢ) '
         'และความหนาของแต่ละชั้น (Dᵢ) ตามสมการ SN = Σ(aᵢ × Dᵢ × mᵢ) '
         f'สรุปผลการคำนวณแสดงในตารางที่ {tbl_sn}')

    # caption ตารางสรุป SN
    if tbl_sn:
        p_sn_cap = _para(indent_cm=0, space_before=4)
        p_sn_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(p_sn_cap, f'ตารางที่ {tbl_sn}  {tbl_cap_sn}', bold=True)

    sn_sum_tbl = doc.add_table(rows=1, cols=8)
    sn_sum_tbl.style = 'Table Grid'
    sn_sum_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, h in enumerate(['ชั้น', 'วัสดุ', 'aᵢ', 'mᵢ', 'Dᵢ (นิ้ว)', 'Dᵢ (ซม.)', 'ΔSNᵢ', 'ΣSN']):
        _tbl_cell(sn_sum_tbl.rows[0].cells[j], h, bold=True, fill='D9E2F3',
                  align=WD_ALIGN_PARAGRAPH.CENTER)
    for layer in calc_results.get('layers', []):
        row = sn_sum_tbl.add_row()
        for j, (val, align) in enumerate([
            (str(layer['layer_no']),                          WD_ALIGN_PARAGRAPH.CENTER),
            (short_material_name(layer['material']),          WD_ALIGN_PARAGRAPH.LEFT),
            (f'{layer["a_i"]:.2f}',                           WD_ALIGN_PARAGRAPH.CENTER),
            (f'{layer["m_i"]:.2f}',                           WD_ALIGN_PARAGRAPH.CENTER),
            (f'{layer["design_thickness_inch"]:.2f}',         WD_ALIGN_PARAGRAPH.CENTER),
            (f'{layer["design_thickness_cm"]:.0f}',           WD_ALIGN_PARAGRAPH.CENTER),
            (f'{layer["sn_contribution"]:.3f}',               WD_ALIGN_PARAGRAPH.CENTER),
            (f'{layer["cumulative_sn"]:.2f}',                 WD_ALIGN_PARAGRAPH.CENTER),
        ]):
            _tbl_cell(row.cells[j], val, align=align)

    doc.add_paragraph()

    # สรุปผล — ตารางผลการตรวจสอบ
    doc.add_paragraph()
    p_chk_cap = _para(indent_cm=0, space_before=6)
    p_chk_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(p_chk_cap, 'ผลการตรวจสอบการออกแบบ', bold=True)

    chk_tbl = doc.add_table(rows=4, cols=2)
    chk_tbl.style = 'Table Grid'
    chk_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    passed = design_check.get('passed', False)
    safety_margin = design_check.get('safety_margin', sn_prov - sn_req)
    for i, (param, value) in enumerate([
        ('SN Required (จากสมการ AASHTO)', f'{sn_req:.2f}'),
        ('SN Provided (จากชั้นทาง)',       f'{sn_prov:.2f}'),
        ('Safety Margin',                  f'{safety_margin:.2f}'),
        ('ผลการตรวจสอบ',                   'ผ่าน (OK)' if passed else 'ไม่ผ่าน (NG)'),
    ]):
        for j, val in enumerate([param, value]):
            chk_tbl.rows[i].cells[j].text = ''
            p = chk_tbl.rows[i].cells[j].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            bold_val = (i == 3)
            set_thai_font(r, size_pt=15, bold=bold_val)
            if i == 3:
                r.font.color.rgb = RGBColor(0x00, 0x70, 0x00) if passed else RGBColor(0xC0, 0x00, 0x00)

    doc.add_paragraph()
    if passed:
        summary_text = (f'สรุป: การออกแบบผ่านเกณฑ์ เนื่องจาก SN_provided ({sn_prov:.2f}) '
                        f'\u2265 SN_required ({sn_req:.2f})')
    else:
        summary_text = (f'สรุป: การออกแบบไม่ผ่านเกณฑ์ เนื่องจาก SN_provided ({sn_prov:.2f}) '
                        f'< SN_required ({sn_req:.2f}) กรุณาปรับเพิ่มความหนาชั้นทาง')
    p_sum_final = _para(indent_cm=0, space_before=4)
    _run(p_sum_final, summary_text, bold=True,
         color=RGBColor(0x00, 0x70, 0x00) if passed else RGBColor(0xC0, 0x00, 0x00))

    # ตารางสรุปโครงสร้างชั้นทาง (3 คอลัมน์ พร้อมรูปตัดขวาง)
    doc.add_paragraph()
    surf_name = short_material_name(calc_results['layers'][0]['material']) if calc_results.get('layers') else ''
    p_sf = _para(indent_cm=0, space_before=6)
    _run(p_sf, f'รูปแบบที่: {surf_name}', bold=True)

    def _run_for_caption(para, text, bold=False, underline=False):
        _run(para, text, bold=bold, underline=underline)

    _add_summary_layer_table(
        doc, calc_results, inputs, fig,
        set_thai_font_fn=set_thai_font,
        add_table_header_shading_fn=add_table_header_shading,
        caption_text=f'ตารางที่ {tbl_sn}  {tbl_cap_sn}',
        caption_bold=True, caption_underline=True,
        use_run_fn=_run_for_caption,
    )
    _fig_caption(f'รูปที่ {fig_no}  {fig_cap}')

    # Footer
    doc.add_paragraph()
    footer_p = doc.add_paragraph()
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(footer_p,
         'พัฒนาโดย รศ.ดร.อิทธิพล มีผล // ภาควิชาครุศาสตร์โยธา // มจพ.',
         size=12, italic=True)

    doc_bytes = BytesIO()
    doc.save(doc_bytes)
    doc_bytes.seek(0)
    return doc_bytes


# ================================================================================
# MAIN APP
# ================================================================================

def main():
    # ========================================
    # HEADER
    # ========================================
    st.title("🛣️  Flexible Pavement Design (AASHTO 1993) v6")
    st.markdown("**แอปพลิเคชันออกแบบโครงสร้างทางแบบยืดหยุ่น ตามวิธีการ AASHTO (1993) — มาตรฐานกรมทางหลวง**")

    # ========================================
    # SIDEBAR — กระชับ: Project / Preset / JSON / ภาษารูป
    # ========================================
    with st.sidebar:
        st.header("📋 ข้อมูลโครงการ")
        project_title = st.text_input(
            "ชื่อโครงการ",
            value=st.session_state.get('input_project_title', "โครงการออกแบบถนน"),
            key="project_title_input"
        )

        st.markdown("---")
        st.header("🏗️ Preset โครงสร้าง ทล.")
        preset_names = list(PRESET_STRUCTURES.keys())
        selected_preset = st.selectbox(
            "เลือกโครงสร้างมาตรฐาน", options=preset_names, index=0,
            help="เลือกเพื่อ Auto-fill วัสดุและความหนาเริ่มต้น (แก้ไขเพิ่มเติมได้)"
        )
        if selected_preset != "--- เลือกโครงสร้างมาตรฐาน ---":
            preset = PRESET_STRUCTURES[selected_preset]
            if preset:
                st.info(f"📋 {preset['description']}")
                if st.button("✅ ใช้โครงสร้างนี้", type="primary"):
                    st.session_state['input_num_layers'] = preset['num_layers']
                    for i, layer in enumerate(preset['layers']):
                        st.session_state[f'layer{i+1}_mat'] = layer['material']
                        st.session_state[f'layer{i+1}_thick'] = layer['thickness_cm']
                        mat = MATERIALS[layer['material']]
                        st.session_state[f'layer{i+1}_a'] = mat['layer_coeff']
                        st.session_state[f'layer{i+1}_m'] = mat['drainage_coeff']
                    st.session_state['use_ac_sublayers'] = False
                    st.session_state['ac_sublayers'] = None
                    st.rerun()

        st.markdown("---")
        st.header("💾 บันทึก/โหลดข้อมูล")
        uploaded_json = st.file_uploader(
            "📂 โหลดข้อมูลจากไฟล์ JSON", type=['json'],
            help="อัปโหลดไฟล์ JSON ที่บันทึกไว้ก่อนหน้า"
        )
        if uploaded_json is not None:
            try:
                loaded_data = json.load(uploaded_json)
                file_id = f"{uploaded_json.name}_{uploaded_json.size}"
                if st.session_state.get('last_uploaded_file') != file_id:
                    st.session_state['last_uploaded_file'] = file_id
                    st.session_state['input_W18']         = loaded_data.get('W18', 5000000)
                    st.session_state['input_reliability'] = loaded_data.get('reliability', 90)
                    st.session_state['input_So']          = loaded_data.get('So', 0.45)
                    st.session_state['input_P0']          = loaded_data.get('P0', 4.2)
                    st.session_state['input_Pt']          = loaded_data.get('Pt', 2.5)
                    st.session_state['input_CBR']         = loaded_data.get('CBR', 5.0)
                    st.session_state['input_num_layers']  = loaded_data.get('num_layers', 4)
                    st.session_state['input_project_title'] = loaded_data.get('project_title', 'โครงการออกแบบถนน')
                    rs = loaded_data.get('report_settings', {})
                    for key, default in [
                        ('section_number', '4.4'), ('table_number_inputs', '4-8'),
                        ('table_number_materials', '4-9'), ('figure_number', '4-8'),
                        ('section_title', 'การออกแบบผิวทางลาดยาง (Flexible Pavement)'),
                        ('table_caption_inputs', 'ค่าพารามิเตอร์ที่ใช้ในการออกแบบผิวทางยืดหยุ่น'),
                        ('table_caption_materials', 'ค่าสัมประสิทธิ์และค่าโมดูลัสของวัสดุโครงสร้างชั้นทาง'),
                        ('figure_caption', 'รูปตัดโครงสร้างชั้นทางที่ออกแบบ'),
                    ]:
                        if key in rs:
                            st.session_state[f'rs_{key}'] = rs[key]
                    layers = loaded_data.get('layers', [])
                    for i, layer in enumerate(layers):
                        mat_name = layer.get('material', '')
                        st.session_state[f'layer{i+1}_mat']   = mat_name
                        st.session_state[f'layer{i+1}_thick'] = layer.get('thickness_cm', 15.0)
                        st.session_state[f'layer{i+1}_m']     = layer.get('drainage_coeff', 1.0)
                        # Bug fix: restore a_i จาก JSON ถ้ามี ไม่อย่างนั้นใช้ค่า default ของวัสดุ
                        if 'layer_coeff' in layer:
                            st.session_state[f'layer{i+1}_a'] = layer['layer_coeff']
                        elif mat_name in MATERIALS:
                            st.session_state[f'layer{i+1}_a'] = MATERIALS[mat_name]['layer_coeff']
                    st.success("✅ โหลดข้อมูลสำเร็จ!")
                    st.rerun()
            except Exception as e:
                st.error(f"❌ ไม่สามารถอ่านไฟล์ได้: {e}")

        st.markdown("---")
        st.header("🖼️ ตั้งค่ารูปภาพ")
        figure_language = st.radio(
            "ภาษาในรูปภาพ", options=["English", "ภาษาไทย"], index=0,
            help="เลือกภาษาสำหรับแสดงในรูปตัดขวาง"
        )

    # ========================================
    # MAIN CONTENT — TABS
    # ========================================
    tab_input, tab_layers, tab_results, tab_report = st.tabs([
        "📝 ข้อมูลนำเข้า", "🏗️ ชั้นทาง", "📊 ผลลัพธ์", "📄 รายงาน"
    ])

    # ========================================
    # TAB 1: DESIGN INPUTS
    # ========================================
    with tab_input:
        st.header("📝 Design Inputs")
        col_t1, col_t2 = st.columns(2)

        with col_t1:
            st.subheader("1️⃣ Traffic & Reliability")
            W18 = st.number_input(
                "Design ESALs (W₁₈)",
                min_value=100000, max_value=250000000,
                value=st.session_state.get('input_W18', 5000000),
                step=100000, format="%d",
                help="จำนวน 18-kip ESAL ตลอดอายุการใช้งาน (สูงสุด 250 ล้าน)",
                key="input_W18"
            )
            esal_million = W18 / 1000000
            st.markdown(
                f'<p style="color:#1E90FF;font-size:20px;font-weight:bold;">'
                f'💡 W₁₈ = {esal_million:,.2f} ล้าน ESALs</p>',
                unsafe_allow_html=True)

            reliability_options = list(RELIABILITY_ZR.keys())
            current_reliability = st.session_state.get('input_reliability', 90)
            default_reliability_idx = (reliability_options.index(current_reliability)
                                       if current_reliability in reliability_options
                                       else reliability_options.index(90))
            reliability = st.selectbox(
                "Reliability Level (R)", options=reliability_options,
                index=default_reliability_idx, key="input_reliability"
            )
            Zr = RELIABILITY_ZR[reliability]
            st.info(f"Zᵣ = {Zr:.3f}")

            So = st.number_input(
                "Overall Standard Deviation (Sₒ)",
                min_value=0.30, max_value=0.60,
                value=st.session_state.get('input_So', 0.45),
                step=0.01, format="%.2f", key="input_So"
            )

        with col_t2:
            st.subheader("2️⃣ Serviceability")
            col_p1, col_p2 = st.columns(2)
            with col_p1:
                P0 = st.number_input("P₀ (Initial)", min_value=3.0, max_value=5.0,
                    value=st.session_state.get('input_P0', 4.2), step=0.1, key="input_P0")
            with col_p2:
                Pt = st.number_input("Pₜ (Terminal)", min_value=1.5, max_value=3.5,
                    value=st.session_state.get('input_Pt', 2.5), step=0.1, key="input_Pt")
            delta_psi = P0 - Pt
            if delta_psi <= 0:
                st.error(f"⚠️ **ΔPSI = {delta_psi:.1f}** — P₀ ต้องมากกว่า Pₜ! กรุณาตรวจสอบค่า")
            else:
                st.success(f"**ΔPSI = {delta_psi:.1f}**")

            st.subheader("3️⃣ Subgrade (ดินเดิม/ดินถม)")
            CBR = st.number_input("CBR (%)", min_value=1.0, max_value=30.0,
                value=st.session_state.get('input_CBR', 5.0), step=0.5,
                help="ค่า CBR ของดินเดิมหรือดินถมคันทาง", key="input_CBR")
            Mr = int(1500 * CBR)
            st.info(f"**Mᵣ = 1,500 × CBR = 1,500 × {CBR:.1f} = {Mr:,} psi**")

        with st.expander("📖 ตาราง Drainage Coefficient (AASHTO Table 2.4)"):
            st.markdown("**ค่าสัมประสิทธิ์การระบายน้ำ (mᵢ) — AASHTO 1993 Table 2.4**")
            st.markdown("ค่า default กรมทางหลวง = **1.0** (สภาพการระบายน้ำดี)")
            drain_data = []
            for quality, info in DRAINAGE_TABLE.items():
                row = {"คุณภาพการระบายน้ำ": f"{quality} — {info['description']}"}
                for pct, val in info['values'].items():
                    row[f"เวลาอิ่มตัว {pct}"] = f"{val:.2f}"
                drain_data.append(row)
            st.table(drain_data)

    # ========================================
    # TAB 2: LAYER CONFIGURATION
    # ========================================
    with tab_layers:
        st.header("🏗️ Layer Configuration")
        num_layers = st.slider(
            "จำนวนชั้นทาง", min_value=2, max_value=6,
            value=st.session_state.get('input_num_layers', 4),
            help="เลือกจำนวนชั้นทาง (2-6 ชั้น)", key="input_num_layers"
        )

        all_materials    = [m for m, p in MATERIALS.items() if p['layer_type'] != 'none']
        surface_materials = [m for m, p in MATERIALS.items() if p['layer_type'] == 'surface']

        layer_data = []
        status_placeholders = {}

        # ===== Global m panel =====
        # หลักการ: ใช้ st.session_state flag "apply_global_m" 
        # เพื่อแยก logic การ apply ออกจาก widget rendering
        # ไม่ใช้ rerun() เพื่อหลีกเลี่ยงการ reset widget values
        with st.container():
            st.markdown(
                '<div style="background:#EFF6FF;border:1.5px solid #3B82F6;border-radius:8px;'
                'padding:10px 16px 6px 16px;margin-bottom:12px;">'
                '<b style="color:#1D4ED8;">🔧 กำหนดค่า m (Drainage Coefficient) เดียวกันทุกชั้น</b>'
                '</div>', unsafe_allow_html=True
            )
            gcol1, gcol2 = st.columns([2, 1])
            with gcol1:
                global_m = st.number_input(
                    "ค่า m สำหรับทุกชั้นทาง",
                    min_value=0.40, max_value=1.50,
                    value=float(st.session_state.get('global_m_value', 1.00)),
                    step=0.05, format="%.2f",
                    help="กรอกค่าแล้วกด 'ใช้เหมือนกันทุกชั้น' — ยังแก้ไขแต่ละชั้นได้ภายหลัง"
                )
                # บันทึกค่าล่าสุดทุก rerun (ไม่ใช้ key เพื่อหลีกเลี่ยง widget state conflict)
                st.session_state['global_m_value'] = global_m
            with gcol2:
                st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
                if st.button("✅ ใช้เหมือนกันทุกชั้น", type="primary", use_container_width=True):
                    nl = st.session_state.get('input_num_layers', 4)
                    for idx in range(1, nl + 1):
                        st.session_state[f'layer{idx}_m'] = global_m
                    st.toast(f"✅ ตั้งค่า m = {global_m:.2f} ให้ทุกชั้นแล้ว", icon="✅")

        # ===== ชั้นที่ 1: ผิวทาง =====
        st.subheader("🔶 ชั้นที่ 1: ผิวทาง (Surface)")

        layer1_mat_default = st.session_state.get('layer1_mat', surface_materials[0])
        layer1_mat_idx = (surface_materials.index(layer1_mat_default)
                         if layer1_mat_default in surface_materials else 0)
        layer1_mat = st.selectbox(
            "เลือกวัสดุ", options=surface_materials,
            index=layer1_mat_idx, key="layer1_mat"
        )

        mat_props_1 = MATERIALS[layer1_mat]
        default_a1  = mat_props_1['layer_coeff']
        default_m1  = mat_props_1['drainage_coeff']

        # ===== AC Sublayer — UI เรียบง่าย =====
        use_sublayers = st.checkbox(
            "📐 แบ่งชั้นย่อยผิวทาง AC (Wearing / Binder / Base Course)",
            value=st.session_state.get('use_ac_sublayers', False),
            help="แบ่งชั้น AC ออกเป็น 3 ชั้นย่อย ตามมาตรฐานกรมทางหลวง",
            key="use_ac_sublayers"
        )

        if use_sublayers:
            st.info("📋 ความหนามาตรฐาน ทล.: Wearing 40-70 มม. / Binder 40-80 มม. / Base 70-100 มม.")

            # ตาราง 3 คอลัมน์ — แต่ละชั้นมีแค่ 1 input
            sub_labels = list(DOH_THICKNESS_STANDARDS.keys())
            sub_keys   = ['wearing', 'binder', 'base']
            sub_defaults = [5.0, 7.0, 10.0]
            sub_results  = {}

            cols = st.columns(3)
            for idx, (col, key, label, defaults_mm, default_cm) in enumerate(zip(
                cols, sub_keys, sub_labels,
                list(DOH_THICKNESS_STANDARDS.values()),
                sub_defaults
            )):
                with col:
                    st.markdown(f"**{label}**")
                    std_options = ["กำหนดเอง"] + (["ไม่ใช้"] if key == 'base' else []) + [f"{t} มม." for t in defaults_mm if t > 0]
                    std_sel = st.selectbox(
                        "มาตรฐาน ทล.", std_options, index=0, key=f"{key}_std_select"
                    )
                    if std_sel == "ไม่ใช้":
                        sub_results[key] = 0.0
                        st.metric("ความหนา", "0.0 cm")
                    elif std_sel != "กำหนดเอง":
                        sub_results[key] = int(std_sel.replace(" มม.", "")) / 10
                        st.metric("ความหนา", f"{sub_results[key]:.1f} cm")
                    else:
                        max_val = 15.0 if key != 'base' else 15.0
                        sub_results[key] = st.number_input(
                            "ความหนา (cm)", 0.0, max_val,
                            value=st.session_state.get(f'{key}_thick_val', default_cm),
                            step=0.5, key=f"{key}_thick"
                        )

            layer1_thick = sum(sub_results.values())
            st.markdown(
                f'<p style="color:#1E90FF;font-size:18px;font-weight:bold;">'
                f'📏 ความหนารวม AC = '
                f'{sub_results["wearing"]:.1f} + {sub_results["binder"]:.1f} + {sub_results["base"]:.1f}'
                f' = {layer1_thick:.1f} cm</p>',
                unsafe_allow_html=True)

            st.session_state['ac_sublayers'] = {
                'wearing': sub_results['wearing'],
                'binder':  sub_results['binder'],
                'base':    sub_results['base'],
                'total':   layer1_thick
            }

            col_am1, col_am2 = st.columns(2)
            with col_am1:
                st.markdown(f"a₁ <span style='color:#1E90FF;font-size:12px;'>(default={default_a1:.2f})</span>",
                            unsafe_allow_html=True)
                layer1_a = st.number_input("a1", 0.10, 0.50,
                    value=st.session_state.get('layer1_a', default_a1), step=0.01,
                    key="layer1_a", label_visibility="collapsed")
            with col_am2:
                st.markdown(f"m₁ <span style='color:#1E90FF;font-size:12px;'>(default={default_m1:.2f})</span>",
                            unsafe_allow_html=True)
                layer1_m = st.number_input("m1", 0.5, 1.5,
                    value=st.session_state.get('layer1_m', default_m1), step=0.05,
                    key="layer1_m", label_visibility="collapsed")
        else:
            st.session_state['ac_sublayers'] = None
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                layer1_thick = st.number_input("ความหนา (cm)", 1.0, 30.0,
                    value=st.session_state.get('layer1_thick', 5.0), step=1.0, key="layer1_thick")
            with col_b:
                st.markdown(f"a₁ <span style='color:#1E90FF;font-size:12px;'>(default={default_a1:.2f})</span>",
                            unsafe_allow_html=True)
                layer1_a = st.number_input("a1", 0.10, 0.50,
                    value=st.session_state.get('layer1_a', default_a1), step=0.01,
                    key="layer1_a", label_visibility="collapsed")
            with col_c:
                st.markdown(f"m₁ <span style='color:#1E90FF;font-size:12px;'>(default={default_m1:.2f})</span>",
                            unsafe_allow_html=True)
                layer1_m = st.number_input("m1", 0.5, 1.5,
                    value=st.session_state.get('layer1_m', default_m1), step=0.05,
                    key="layer1_m", label_visibility="collapsed")

        st.markdown(f'<p style="color:#1E90FF;font-size:14px;">E = {mat_props_1["mr_mpa"]:,} MPa</p>',
                    unsafe_allow_html=True)
        status_placeholders[1] = st.empty()

        layer_data.append({
            'material': layer1_mat,
            'thickness_cm': layer1_thick,
            'layer_coeff': layer1_a,
            'drainage_coeff': layer1_m
        })

        # ===== ชั้นที่ 2-6 =====
        default_materials = [
            "พื้นทางหินคลุกปรับปรุงคุณภาพด้วยปูนซีเมนต์ (Cement Treated Base)",
            "รองพื้นทางวัสดุมวลรวม CBR 25%",
            "วัสดุคัดเลือก ก",
            "วัสดุคัดเลือก ก",
            "วัสดุคัดเลือก ก"
        ]
        default_thickness = [15.0, 15.0, 30.0, 30.0, 30.0]
        layer_icons = ['🔷', '🔶', '🟢', '🟡', '🔴']

        for i in range(2, num_layers + 1):
            st.markdown("---")
            st.subheader(f"{layer_icons[i-2]} ชั้นที่ {i}")

            layer_i_mat_default = st.session_state.get(f'layer{i}_mat', default_materials[i-2])
            if layer_i_mat_default in all_materials:
                default_idx = all_materials.index(layer_i_mat_default)
            else:
                default_idx = (all_materials.index(default_materials[i-2])
                              if default_materials[i-2] in all_materials else 0)

            layer_mat = st.selectbox(
                f"เลือกวัสดุชั้นที่ {i}", options=all_materials,
                index=min(default_idx, len(all_materials)-1), key=f"layer{i}_mat"
            )

            mat_props = MATERIALS[layer_mat]
            default_a = mat_props['layer_coeff']
            default_m = mat_props['drainage_coeff']

            prev_mat_key = f'layer{i}_prev_mat'
            if prev_mat_key not in st.session_state:
                st.session_state[prev_mat_key] = layer_mat
            if st.session_state[prev_mat_key] != layer_mat:
                st.session_state[f'layer{i}_a'] = default_a
                st.session_state[f'layer{i}_m'] = default_m
                st.session_state[prev_mat_key] = layer_mat

            col_c, col_d, col_e = st.columns(3)
            with col_c:
                layer_thick = st.number_input("ความหนา (cm)", 1.0, 150.0,
                    value=st.session_state.get(f'layer{i}_thick', default_thickness[i-2]),
                    step=5.0, key=f"layer{i}_thick")
            with col_d:
                st.markdown(f"a{i} <span style='color:#1E90FF;font-size:12px;'>(default={default_a:.2f})</span>",
                            unsafe_allow_html=True)
                layer_a = st.number_input(f"a{i}", 0.01, 0.50,
                    value=st.session_state.get(f'layer{i}_a', default_a), step=0.01,
                    key=f"layer{i}_a", label_visibility="collapsed")
            with col_e:
                st.markdown(f"m{i} <span style='color:#1E90FF;font-size:12px;'>(default={default_m:.2f})</span>",
                            unsafe_allow_html=True)
                layer_m = st.number_input(f"m{i}", 0.5, 1.5,
                    value=st.session_state.get(f'layer{i}_m', default_m), step=0.05,
                    key=f"layer{i}_m", label_visibility="collapsed")

            st.markdown(f'<p style="color:#1E90FF;font-size:14px;">E = {mat_props["mr_mpa"]:,} MPa</p>',
                        unsafe_allow_html=True)
            status_placeholders[i] = st.empty()

            layer_data.append({
                'material': layer_mat,
                'thickness_cm': layer_thick,
                'layer_coeff': layer_a,
                'drainage_coeff': layer_m
            })

        # ===== ฐานข้อมูลวัสดุ (ย้ายมาอยู่ใน Tab 2) =====
        st.markdown("---")
        with st.expander("📚 ฐานข้อมูลวัสดุ (มาตรฐาน ทล.) — คลิกเพื่อดูค่า สปส. ทั้งหมด"):
            for mat_name, props in MATERIALS.items():
                if props['layer_coeff'] > 0:
                    st.markdown(f"**{mat_name}**")
                    st.markdown(f"- a = {props['layer_coeff']}, m = {props['drainage_coeff']}")
                    st.markdown(f"- MR = {props['mr_psi']:,} psi ({props['mr_mpa']:,} MPa)")
                    st.markdown("---")

    # ========================================
    # CALCULATION (ทำงานทุก rerun อัตโนมัติ)
    # ========================================
    inputs = {
        'W18': W18, 'reliability': reliability, 'Zr': Zr, 'So': So,
        'P0': P0, 'Pt': Pt, 'delta_psi': delta_psi, 'CBR': CBR, 'Mr': Mr
    }

    if delta_psi <= 0:
        st.error("❌ ไม่สามารถคำนวณได้: P₀ ต้องมากกว่า Pₜ (ΔPSI > 0) — กรุณาตรวจสอบค่าใน Tab ข้อมูลนำเข้า")
        st.stop()

    ac_sublayers = st.session_state.get('ac_sublayers', None)
    calc_results  = calculate_layer_thicknesses(W18, Zr, So, delta_psi, Mr, layer_data, ac_sublayers)
    design_check  = check_design(calc_results['total_sn_required'], calc_results['total_sn_provided'])

    # ===== QUICK SUMMARY BANNER — แสดงเหนือ tabs ทุกครั้ง =====
    if design_check['passed']:
        st.markdown(
            f"""<div style="background:#d4edda;border:2px solid #28a745;border-radius:10px;
            padding:14px 24px;text-align:center;margin:10px 0 18px 0;">
            <h3 style="color:#28a745;margin:0;">✅ PASS — การออกแบบผ่านเกณฑ์</h3>
            <p style="font-size:17px;margin:6px 0;">
            SN<sub>provided</sub> = <b>{calc_results['total_sn_provided']:.2f}</b> &nbsp;≥&nbsp;
            SN<sub>required</sub> = <b>{calc_results['total_sn_required']:.2f}</b>
            &nbsp;|&nbsp; Safety Margin = <b>{design_check['safety_margin']:.2f}</b>
            &nbsp;|&nbsp; โครงการ: <b>{project_title}</b>
            </p></div>""", unsafe_allow_html=True)
    else:
        st.markdown(
            f"""<div style="background:#f8d7da;border:2px solid #dc3545;border-radius:10px;
            padding:14px 24px;text-align:center;margin:10px 0 18px 0;">
            <h3 style="color:#dc3545;margin:0;">❌ FAIL — การออกแบบไม่ผ่าน</h3>
            <p style="font-size:17px;margin:6px 0;">
            SN<sub>provided</sub> = <b>{calc_results['total_sn_provided']:.2f}</b> &nbsp;&lt;&nbsp;
            SN<sub>required</sub> = <b>{calc_results['total_sn_required']:.2f}</b>
            &nbsp;|&nbsp; ขาดอีก = <b>{abs(design_check['safety_margin']):.2f}</b>
            &nbsp;|&nbsp; โครงการ: <b>{project_title}</b>
            </p></div>""", unsafe_allow_html=True)

    # Fill status placeholders in Tab 2
    for layer in calc_results['layers']:
        layer_no = layer['layer_no']
        if layer_no in status_placeholders:
            with status_placeholders[layer_no]:
                if layer['is_ok']:
                    st.success(f"✅ ผ่าน (ต้องการ ≥ {layer['min_thickness_cm']:.1f} cm)")
                else:
                    shortage = layer['min_thickness_cm'] - layer['design_thickness_cm']
                    st.error(f"❌ ไม่ผ่าน (ต้องเพิ่มอีก {shortage:.1f} cm)")

    # fig_lang — ใช้ร่วมกันทั้ง Tab 3 และ Tab 4
    fig_lang = 'th' if figure_language == "ภาษาไทย" else 'en'

    # ========================================
    # TAB 3: RESULTS
    # ========================================
    with tab_results:
        # ===== Warnings =====
        warnings = calc_results.get('warnings', [])
        if warnings:
            for w in warnings:
                st.warning(w)

        # ===== W18 Supported metrics =====
        w18_supported = calculate_w18_supported(
            calc_results['total_sn_provided'], Zr, So, delta_psi, Mr)
        w18_supported_million = w18_supported / 1_000_000
        w18_diff_percent = ((w18_supported - W18) / W18) * 100

        mc1, mc2, mc3, mc4 = st.columns(4)
        with mc1:
            st.metric("W₁₈ ออกแบบ", f"{W18/1e6:,.2f} ล้าน")
        with mc2:
            st.metric("W₁₈ รองรับได้", f"{w18_supported_million:,.2f} ล้าน",
                      delta=f"{w18_diff_percent:+.1f}%",
                      delta_color="normal" if w18_diff_percent >= 0 else "inverse")
        with mc3:
            st.metric("SN Required", f"{calc_results['total_sn_required']:.2f}")
        with mc4:
            st.metric("SN Provided", f"{calc_results['total_sn_provided']:.2f}",
                      delta=f"{design_check['safety_margin']:+.2f}")

        st.markdown("---")

        # ===== STEP-BY-STEP CALCULATION =====
        st.subheader("🔢 ขั้นตอนการคำนวณความหนาแต่ละชั้น")

        for layer in calc_results['layers']:
            layer_status = "✅" if layer['is_ok'] else "❌"
            st.markdown(f"### {layer_status} ชั้นที่ {layer['layer_no']}: {short_material_name(layer['material'])}")

            # AC sublayer info
            layer_ac_sub = layer.get('ac_sublayers', None)
            if layer_ac_sub is not None and layer['layer_no'] == 1:
                st.info(
                    f"**📐 แบ่งชั้นย่อย AC:** "
                    f"ผิวทาง (Wearing) = {layer_ac_sub['wearing']:.1f} cm  |  "
                    f"รองผิวทาง (Binder) = {layer_ac_sub['binder']:.1f} cm  |  "
                    f"พื้นทาง (Base) = {layer_ac_sub['base']:.1f} cm  |  "
                    f"**รวม = {layer_ac_sub['total']:.1f} cm**"
                )

            col_a, col_b = st.columns([1, 1])
            with col_a:
                st.markdown("**ข้อมูลวัสดุ:**")
                st.markdown(f"- E (MPa) = **{layer['mr_mpa']:,}**")
                st.markdown(f"- Mᵣ (psi) = **{layer['mr_psi']:,}**")
                st.markdown(f"- Layer Coefficient (a{layer['layer_no']}) = **{layer['a_i']:.2f}**")
                st.markdown(f"- Drain Coefficient (m{layer['layer_no']}) = **{layer['m_i']:.2f}**")

            with col_b:
                st.markdown("**จากสมการ AASHTO:**")
                sn_at_layer = layer['sn_required_at_layer']
                st.latex(f"SN_{{{layer['layer_no']}}} = {sn_at_layer:.2f}")

            st.markdown("**คำนวณความหนาผิวทาง:**")
            if layer['layer_no'] == 1:
                st.latex(
                    f"D_{{1}} \\geq \\frac{{SN_{{1}}}}{{a_{{1}} \\times m_{{1}}}} = "
                    f"\\frac{{{sn_at_layer:.2f}}}{{{layer['a_i']:.2f} \\times {layer['m_i']:.2f}}} = "
                    f"{layer['min_thickness_inch']:.2f} \\text{{ นิ้ว}} = {layer['min_thickness_cm']:.1f} \\text{{ ซม.}}")
            else:
                prev_sn = calc_results['layers'][layer['layer_no']-2]['cumulative_sn']
                st.latex(
                    f"D_{{{layer['layer_no']}}} \\geq "
                    f"\\frac{{SN_{{{layer['layer_no']}}} - SN_{{prev}}}}"
                    f"{{a_{{{layer['layer_no']}}} \\times m_{{{layer['layer_no']}}}}} = "
                    f"\\frac{{{sn_at_layer:.2f} - {prev_sn:.2f}}}"
                    f"{{{layer['a_i']:.2f} \\times {layer['m_i']:.2f}}} = "
                    f"{layer['min_thickness_inch']:.2f} \\text{{ นิ้ว}} = {layer['min_thickness_cm']:.1f} \\text{{ ซม.}}")

            result_cols = st.columns(4)
            with result_cols[0]:
                st.metric("ความหนาขั้นต่ำ", f"{layer['min_thickness_cm']:.1f} cm")
            with result_cols[1]:
                st.metric("ความหนาที่เลือก", f"{layer['design_thickness_cm']:.0f} cm",
                         delta=f"{layer['design_thickness_cm'] - layer['min_thickness_cm']:.1f} cm")
            with result_cols[2]:
                st.metric("SN contribution", f"{layer['sn_contribution']:.3f}")
            with result_cols[3]:
                st.metric("Cumulative SN", f"{layer['cumulative_sn']:.2f}")

            if layer['is_ok']:
                st.success(f"✅ **OK** — ความหนาเพียงพอ ({layer['design_thickness_cm']:.0f} ≥ {layer['min_thickness_cm']:.1f} cm)")
            else:
                st.error(f"❌ **NG** — ต้องเพิ่มความหนาอีก {layer['min_thickness_cm'] - layer['design_thickness_cm']:.1f} cm")
            st.markdown("---")

        # ===== SN Summary Table =====
        with st.expander("📋 ตารางสรุปการคำนวณ SN"):
            table_data = []
            for layer in calc_results['layers']:
                table_data.append({
                    'ชั้น': layer['layer_no'],
                    'วัสดุ': layer['short_name'],
                    'aᵢ': layer['a_i'],
                    'Dᵢ (cm)': layer['design_thickness_cm'],
                    'Dᵢ (in)': layer['design_thickness_inch'],
                    'mᵢ': layer['m_i'],
                    'E (MPa)': layer['mr_mpa'],
                    'SN contrib.': layer['sn_contribution'],
                    'SN cumul.': layer['cumulative_sn']
                })
            st.table(table_data)
            st.markdown(
                f"**สูตร:** $SN = \\sum_{{i=1}}^{{n}} a_i \\times D_i \\times m_i$  |  "
                f"**SN_provided = {calc_results['total_sn_provided']:.2f}**  |  "
                f"**SN_required = {calc_results['total_sn_required']:.2f}**")

        # ===== PAVEMENT SECTION FIGURE =====
        st.subheader("📐 ภาพตัดขวางโครงสร้างถนน")
        fig = plot_pavement_section(calc_results['layers'], Mr, CBR, lang=fig_lang)
        st.pyplot(fig)
        plt.close(fig)

        # ===== SENSITIVITY ANALYSIS =====
        st.markdown("---")
        st.subheader("📈 Sensitivity Analysis")
        st.caption("🔴 Red dot = current design value  |  กราฟแสดงผลกระทบของตัวแปรหลักต่อ SN_required")

        sens_col1, sens_col2 = st.columns(2)
        with sens_col1:
            fig_cbr = plot_sensitivity_cbr(W18, Zr, So, delta_psi, CBR)
            st.pyplot(fig_cbr)
            plt.close(fig_cbr)
        with sens_col2:
            fig_w18 = plot_sensitivity_w18(Zr, So, delta_psi, Mr, W18)
            st.pyplot(fig_w18)
            plt.close(fig_w18)

    # ========================================
    # TAB 4: REPORT & EXPORT
    # ========================================
    with tab_report:
        st.header("📄 ส่งออกรายงาน")

        # ===== Report Settings =====
        st.markdown("### 📝 ตั้งค่าหมายเลขหัวข้อและตารางสำหรับรายงาน Word")

        # --- กลุ่ม 1: เลขหัวข้อ / เลขรูป / ชื่อหัวข้อ ---
        st.markdown("**📌 หัวข้อและรูป**")
        col_g1a, col_g1b, col_g1c = st.columns([1, 1, 3])
        with col_g1a:
            rs_section_number = st.text_input(
                "เลขหัวข้อ",
                value=st.session_state.get('rs_section_number', '4.4'),
                key='rs_section_number'
            )
        with col_g1b:
            rs_figure_number = st.text_input(
                "เลขรูป",
                value=st.session_state.get('rs_figure_number', '4-8'),
                key='rs_figure_number'
            )
        with col_g1c:
            rs_section_title = st.text_input(
                "ชื่อหัวข้อ",
                value=st.session_state.get('rs_section_title',
                      'การออกแบบผิวทางลาดยาง (Flexible Pavement)'),
                key='rs_section_title'
            )

        # --- กลุ่ม 2: เลขตาราง 3 ตาราง ---
        st.markdown("**📋 เลขตาราง**")
        col_g2a, col_g2b, col_g2c = st.columns(3)
        with col_g2a:
            rs_table_number_inputs = st.text_input(
                "เลขตารางพารามิเตอร์",
                value=st.session_state.get('rs_table_number_inputs', '4-8'),
                key='rs_table_number_inputs'
            )
        with col_g2b:
            rs_table_number_materials = st.text_input(
                "เลขตารางวัสดุ",
                value=st.session_state.get('rs_table_number_materials', '4-9'),
                key='rs_table_number_materials'
            )
        with col_g2c:
            rs_table_number_sn = st.text_input(
                "เลขตารางสรุป SN",
                value=st.session_state.get('rs_table_number_sn', '4-10'),
                key='rs_table_number_sn'
            )

        # --- กลุ่ม 3: คำบรรยาย ---
        st.markdown("**🗒️ คำบรรยาย**")
        col_cap1, col_cap2 = st.columns(2)
        with col_cap1:
            rs_table_caption_inputs = st.text_input(
                "คำบรรยายตารางพารามิเตอร์",
                value=st.session_state.get('rs_table_caption_inputs',
                      'ค่าพารามิเตอร์ที่ใช้ในการออกแบบผิวทางยืดหยุ่น'),
                key='rs_table_caption_inputs'
            )
        with col_cap2:
            rs_table_caption_materials = st.text_input(
                "คำบรรยายตารางวัสดุ",
                value=st.session_state.get('rs_table_caption_materials',
                      'ค่าสัมประสิทธิ์และค่าโมดูลัสของวัสดุโครงสร้างชั้นทาง'),
                key='rs_table_caption_materials'
            )

        col_cap3, col_cap4 = st.columns(2)
        with col_cap3:
            rs_table_caption_sn = st.text_input(
                "คำบรรยายตารางสรุป SN",
                value=st.session_state.get('rs_table_caption_sn',
                      'สรุปผลการคำนวณ Structural Number ของโครงสร้างชั้นทาง'),
                key='rs_table_caption_sn'
            )
        with col_cap4:
            rs_figure_caption = st.text_input(
                "คำบรรยายรูป",
                value=st.session_state.get('rs_figure_caption', 'รูปตัดโครงสร้างชั้นทางที่ออกแบบ'),
                key='rs_figure_caption'
            )

        report_settings = {
            'section_number':          rs_section_number,
            'table_number_inputs':     rs_table_number_inputs,
            'table_number_materials':  rs_table_number_materials,
            'table_number_sn':         rs_table_number_sn,
            'figure_number':           rs_figure_number,
            'section_title':           rs_section_title,
            'table_caption_inputs':    rs_table_caption_inputs,
            'table_caption_materials': rs_table_caption_materials,
            'table_caption_sn':        rs_table_caption_sn,
            'figure_caption':          rs_figure_caption,
        }

        st.markdown("---")

        # ===== Preview บทเกริ่นนำ =====
        st.markdown("### 👁️ Preview บทเกริ่นนำ")

        total_thick_prev = sum(l['design_thickness_cm'] for l in calc_results['layers'])
        num_layers_prev  = len(calc_results['layers'])
        passed_prev      = 'ผ่านเกณฑ์' if design_check['passed'] else 'ไม่ผ่านเกณฑ์'

        def hl_purple(val):
            return f'<span style="background-color:#D8B4FE;padding:1px 4px;border-radius:3px;font-weight:bold;">{val}</span>'
        def hl_yellow(val):
            return f'<span style="background-color:#FDE68A;padding:1px 4px;border-radius:3px;font-weight:bold;">{val}</span>'

        intro_html = f"""
        <div style="background:#f9f9f9;padding:15px 20px;border-radius:8px;border:1px solid #ddd;
                    font-family:'TH SarabunPSK',Sarabun,sans-serif;font-size:16px;line-height:1.9;">
            <p style="font-weight:bold;margin-bottom:5px;">
                {hl_yellow(rs_section_number)}&nbsp;&nbsp;{hl_yellow(rs_section_title)}
            </p>
            <p style="text-indent:40px;text-align:justify;">
                ถนนลาดยางซึ่งประกอบด้วยวัสดุงานทางหลายชนิด การออกแบบโครงสร้างถนนแบบยืดหยุ่น (Flexible Pavement)
                ใช้วิธี AASHTO 1993 Guide for Design of Pavement Structures โดยพิจารณาปัจจัยด้านปริมาณจราจรสะสม ESALs
                ความน่าเชื่อถือ และคุณสมบัติของดินรองรับ
                สำหรับโครงการนี้ที่ปรึกษาได้กำหนดค่าพารามิเตอร์หลักในการออกแบบ ได้แก่
                ปริมาณ W&#8321;&#8328; = {hl_purple(f"{W18:,.0f}")} 18-kip ESALs
                ที่ระดับความน่าเชื่อถือ (Reliability) = {hl_purple(reliability)} %
                โดยมีดินเดิมค่า CBR = {hl_purple(f"{CBR:.1f}")} % (M&#7523; = {hl_purple(f"{Mr:,.0f}")} psi)
                ผลการออกแบบได้โครงสร้างชั้นทาง {hl_purple(num_layers_prev)} ชั้น
                ที่ SN&#8203;_required = {hl_purple(f"{calc_results['total_sn_required']:.2f}")}
                และ SN&#8203;_provided = {hl_purple(f"{calc_results['total_sn_provided']:.2f}")}
                ความหนารวม {hl_purple(f"{total_thick_prev:.0f}")} ซม.
                การออกแบบ{hl_purple(passed_prev)}
                ดังแสดงผลการวิเคราะห์ในตารางที่ <b>{hl_yellow(rs_table_number_inputs)}</b>
                และตารางที่ <b>{hl_yellow(rs_table_number_materials)}</b>
                และรูปที่ <b>{hl_yellow(rs_figure_number)}</b>
            </p>
        </div>
        """
        st.markdown(intro_html, unsafe_allow_html=True)
        st.caption("🟣 สีม่วง = ดึงจากผลคำนวณอัตโนมัติ | 🟡 สีเหลือง = ผู้ใช้กรอกเอง")
        st.markdown("---")

        # ===== EXPORT BUTTONS — จัดใหม่เป็น 2 แถว =====
        st.markdown("### 📥 ดาวน์โหลดรายงาน")

        # แถว 1: Word reports
        col_r1, col_r2 = st.columns(2)
        with col_r1:
            if st.button("📋 สร้างรายงานแบบที่ปรึกษา", type="primary",
                         use_container_width=True,
                         help="รายงานรูปแบบสำหรับรวมกับบทรายงานอื่น"):
                with st.spinner("กำลังสร้างรายงาน..."):
                    fig_intro = plot_pavement_section(calc_results['layers'], Mr, CBR, lang=fig_lang)
                    doc_intro_bytes = create_word_report_intro(
                        project_title, inputs, calc_results, design_check, fig_intro, report_settings)
                    plt.close(fig_intro)
                st.download_button(
                    label="⬇️ ดาวน์โหลดรายงานแบบที่ปรึกษา (.docx)",
                    data=doc_intro_bytes,
                    file_name=f"flexible_consult_{project_title}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

        with col_r2:
            if st.button("📝 สร้างรายงานแบบย่อ", use_container_width=True):
                with st.spinner("กำลังสร้างรายงาน..."):
                    fig_thai = plot_pavement_section(calc_results['layers'], Mr, CBR, lang=fig_lang)
                    doc_bytes = create_word_report(project_title, inputs, calc_results, design_check, fig_thai, report_settings)
                    plt.close(fig_thai)
                st.download_button(
                    label="⬇️ ดาวน์โหลดรายงานแบบย่อ (.docx)",
                    data=doc_bytes,
                    file_name=f"flexible_brief_{project_title}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

        # แถว 2: PNG + JSON
        col_r3, col_r4 = st.columns(2)
        with col_r3:
            fig_export = plot_pavement_section(calc_results['layers'], Mr, CBR, lang=fig_lang)
            fig_bytes  = get_figure_as_bytes(fig_export)
            plt.close(fig_export)
            st.download_button(
                label="📸 ดาวน์โหลดรูปตัดขวาง (.png)",
                data=fig_bytes,
                file_name=f"Pavement_Section_{datetime.now().strftime('%Y%m%d_%H%M')}.png",
                mime="image/png",
                use_container_width=True
            )

        with col_r4:
            export_data = {
                'project_title': project_title,
                'W18': W18, 'reliability': reliability, 'So': So,
                'P0': P0, 'Pt': Pt, 'CBR': CBR,
                'num_layers': num_layers,
                'layers': layer_data,
                'ac_sublayers': st.session_state.get('ac_sublayers', None),
                'report_settings': report_settings,
            }
            json_str = json.dumps(export_data, ensure_ascii=False, indent=2)
            st.download_button(
                label="💾 ดาวน์โหลดข้อมูล (.json)",
                data=json_str,
                file_name=f"Flexible_Input_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
                mime="application/json",
                use_container_width=True
            )

        st.markdown("---")

        # ===== Summary Table =====
        st.subheader("📊 สรุปผลการออกแบบ")
        summary_data = [
            ("ชื่อโครงการ", project_title),
            ("W₁₈ (Design ESALs)", f"{W18:,.0f} ({W18/1e6:,.2f} ล้าน)"),
            ("Reliability", f"{reliability}%"),
            ("CBR", f"{CBR:.1f}%"),
            ("Mᵣ (Subgrade)", f"{Mr:,} psi"),
            ("SN Required", f"{calc_results['total_sn_required']:.2f}"),
            ("SN Provided", f"{calc_results['total_sn_provided']:.2f}"),
            ("Safety Margin", f"{design_check['safety_margin']:.2f}"),
            ("ผลการตรวจสอบ", "✅ PASS" if design_check['passed'] else "❌ FAIL"),
        ]
        st.table(summary_data)

    # ===== FOOTER =====
    st.markdown("---")
    st.markdown("""
    <div style='text-align:center;color:gray;'>
    <p>AASHTO 1993 Flexible Pavement Design Application v6.0</p>
    <p>พัฒนาโดย รศ.ดร.อิทธิพล มีผล // ภาควิชาครุศาสตร์โยธา // มจพ.</p>
    </div>
    """, unsafe_allow_html=True)


# ================================================================================
# ENTRY POINT
# ================================================================================

if __name__ == "__main__":
    main()
