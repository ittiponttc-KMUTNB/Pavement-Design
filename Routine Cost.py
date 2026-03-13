import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="ระบบคำนวณงานบำรุงปกติ",
    page_icon="🛣️",
    layout="wide",
)

# ─────────────────────────────────────────────
# LOOKUP TABLES (จาก KA.xlsx, KC.xlsx, PDF)
# ─────────────────────────────────────────────

# ── KA: X1 ──────────────────────────────────
X1_MAP = {
    "High Type (AC/PM บนหินคลุก)": 0.00,
    "Intermediate Type (AC/PM บน Stabilized)": 0.50,
    "Low Type (ST บน Soil-Aggregate)": 1.00,
}

# ── KA: X2 ──────────────────────────────────
X2_BREAKS = [(0, 2, 1.00), (2.01, 3, 0.75), (3.01, 4, 0.50), (4.01, 5, 0.25), (5.01, 999, 0.00)]

# ── KA: X3 ──────────────────────────────────
X3_LOWER = [0,501,601,701,801,901,1001,1101,1201,1301,1401,1501,1601,1701,1801,1901,2001,2201,2401,2601,2801,3001,3301,3601,3901,4201,4501,4801,5101,5401,5701]
X3_UPPER = [500,600,700,800,900,1000,1100,1200,1300,1400,1500,1600,1700,1800,1900,2000,2200,2400,2600,2800,3000,3300,3600,3900,4200,4500,4600,5100,5400,5700,999999]
X3_VAL   = [0,0.04,0.08,0.12,0.16,0.20,0.24,0.29,0.33,0.37,0.41,0.45,0.49,0.53,0.57,0.61,0.69,0.78,0.86,0.94,1.02,1.14,1.27,1.37,1.51,1.64,1.76,1.88,2.00,2.13,2.25]

# X3 Dropdown options: label → factor value
X3_OPTIONS = {
    f"0 – 500       (X3 = 0.00)": 0.00,
    f"501 – 600     (X3 = 0.04)": 0.04,
    f"601 – 700     (X3 = 0.08)": 0.08,
    f"701 – 800     (X3 = 0.12)": 0.12,
    f"801 – 900     (X3 = 0.16)": 0.16,
    f"901 – 1,000   (X3 = 0.20)": 0.20,
    f"1,001 – 1,100 (X3 = 0.24)": 0.24,
    f"1,101 – 1,200 (X3 = 0.29)": 0.29,
    f"1,201 – 1,300 (X3 = 0.33)": 0.33,
    f"1,301 – 1,400 (X3 = 0.37)": 0.37,
    f"1,401 – 1,500 (X3 = 0.41)": 0.41,
    f"1,501 – 1,600 (X3 = 0.45)": 0.45,
    f"1,601 – 1,700 (X3 = 0.49)": 0.49,
    f"1,701 – 1,800 (X3 = 0.53)": 0.53,
    f"1,801 – 1,900 (X3 = 0.57)": 0.57,
    f"1,901 – 2,000 (X3 = 0.61)": 0.61,
    f"2,001 – 2,200 (X3 = 0.69)": 0.69,
    f"2,201 – 2,400 (X3 = 0.78)": 0.78,
    f"2,401 – 2,600 (X3 = 0.86)": 0.86,
    f"2,601 – 2,800 (X3 = 0.94)": 0.94,
    f"2,801 – 3,000 (X3 = 1.02)": 1.02,
    f"3,001 – 3,300 (X3 = 1.14)": 1.14,
    f"3,301 – 3,600 (X3 = 1.27)": 1.27,
    f"3,601 – 3,900 (X3 = 1.37)": 1.37,
    f"3,901 – 4,200 (X3 = 1.51)": 1.51,
    f"4,201 – 4,500 (X3 = 1.64)": 1.64,
    f"4,501 – 4,800 (X3 = 1.76)": 1.76,
    f"4,801 – 5,100 (X3 = 1.88)": 1.88,
    f"5,101 – 5,400 (X3 = 2.00)": 2.00,
    f"5,401 – 5,700 (X3 = 2.13)": 2.13,
    f"5,701+         (X3 = 2.25)": 2.25,
}

# ── KA: X4 ──────────────────────────────────
X4_BREAKS = [(0,3,0.00),(4,4,0.20),(5,5,0.40),(6,6,0.60),(7,7,0.80),(8,8,1.00),(9,9,1.20),(10,10,1.40),(11,11,1.60),(12,99999,1.80)]

# ── KA: X5 ──────────────────────────────────
X5_BREAKS = [(0,5.49,0.00),(5.50,5.99,0.02),(6.00,6.49,0.05),(6.50,6.99,0.10),(7.00,9999,0.19)]

# ── X6 / Y3 / Y4 / Y6 ── terrain types ──────
TERRAIN_MAP = {"ที่ราบ (0-3%)": "P", "ลูกเนิน (3-5%)": "R", "ลูกเนินสลับเขา (5-7%)": "RM", "เขา (>7%)": "S"}

X6_MAP = {"P": 0.00, "R": 0.02, "RM": 0.04, "S": 0.07}
Y3_MAP = {"P": 0.00, "R": 0.24, "RM": 0.36, "S": 0.48}
Y4_MAP = {"P": 0.00, "R": 0.24, "RM": 0.36, "S": 0.48}
Y6_MAP = {"P": 0.00, "R": 0.04, "RM": 0.08, "S": 0.12}

# ── Y1 ──────────────────────────────────────
Y1_BREAKS = [(0,20,0.00),(20.01,30,0.00),(30.01,40,0.00),(40.01,60,0.10),(60.01,80,0.20),(80.01,9999,0.30)]

# ── Y2 (ลาดยาง) ─────────────────────────────
Y2_BREAKS = [(0,0.50,0.00),(0.51,1.00,0.00),(1.01,1.50,0.00),(1.51,1.75,0.00),(1.76,2.00,0.10),(2.01,2.25,0.15),(2.26,9999,0.20)]

# ── Y5 ──────────────────────────────────────
Y5_BREAKS = [(0,20.99,0.00),(21,25,0.02),(25.01,30,0.04),(30.01,9999,0.06)]

# ── KC: Z1 ──────────────────────────────────
Z1_MAP = {1:0.00, 2:0.25, 3:0.50, 4:0.75, 5:1.00, 6:1.30, 7:1.60, 8:2.00}

# ── KC: Z2 ──────────────────────────────────
Z2_BREAKS = [(0,2,1.00),(2.01,3,0.75),(3.01,4,0.50),(4.01,5,0.25),(5.01,999,0.00)]

# ── KC: Z3 ──────────────────────────────────
Z3_LOWER = [0,1001,2001,3001,4001,5001,6001,7001,8001,9001,10001,15001]
Z3_UPPER = [1000,2000,3000,4000,5000,6000,7000,8000,9000,10000,15000,999999]
Z3_VAL   = [0,0.20,0.30,0.50,0.75,1.00,1.25,1.50,1.75,2.00,2.50,3.00]

# Z3 Dropdown options: label → factor value
Z3_OPTIONS = {
    "0 – 1,000        (Z3 = 0.00)": 0.00,
    "1,001 – 2,000    (Z3 = 0.20)": 0.20,
    "2,001 – 3,000    (Z3 = 0.30)": 0.30,
    "3,001 – 4,000    (Z3 = 0.50)": 0.50,
    "4,001 – 5,000    (Z3 = 0.75)": 0.75,
    "5,001 – 6,000    (Z3 = 1.00)": 1.00,
    "6,001 – 7,000    (Z3 = 1.25)": 1.25,
    "7,001 – 8,000    (Z3 = 1.50)": 1.50,
    "8,001 – 9,000    (Z3 = 1.75)": 1.75,
    "9,001 – 10,000   (Z3 = 2.00)": 2.00,
    "10,001 – 15,000  (Z3 = 2.50)": 2.50,
    "15,001+           (Z3 = 3.00)": 3.00,
}

# ── KC: Z4 ──────────────────────────────────
Z4_BREAKS = [(0,6.49,0.00),(6.50,6.99,0.08),(7.00,9999,0.17)]

# ── KB: A1 ──────────────────────────────────
A1_BREAKS = [(0,100,0.00),(101,150,0.13),(151,200,0.24),(201,250,0.36),(251,300,0.47),(301,350,0.59),(351,400,0.71),(401,9999,0.95)]

# ── KB: A3 ──────────────────────────────────
A3_BREAKS = [(0,6.49,0.00),(6.50,7.49,0.17),(7.50,8.49,0.33),(8.50,9.49,0.55),(9.50,10.49,0.67),(10.50,11.49,0.84),(11.50,9999,1.00)]

# ── KB: B1 ──────────────────────────────────
B1_BREAKS_KB = [(0,20,0.00),(20.01,30,0.08),(30.01,40,0.13),(40.01,50,0.21),(50.01,9999,0.24)]

# ── KB: B2 / B3 ──────────────────────────────
B2_MAP = {"P": 0.05, "R": 0.13, "RM": 0.22, "S": 0.32}
B3_MAP = {"P": 0.00, "R": 0.40, "RM": 0.60, "S": 0.80}

# ── KB: B4 ──────────────────────────────────
B4_BREAKS = [(0,20,0.02),(21,21,0.03),(22,22,0.10),(23,23,0.15),(24,24,0.20),(25,25,0.25),(26,26,0.30),(27,27,0.35),(28,28,0.40),(29,29,0.45),(30,9999,0.50)]

# ─────────────────────────────────────────────
# HELPER FUNCTIONS
# ─────────────────────────────────────────────

def lookup_range(value, breaks):
    for lo, hi, v in breaks:
        if lo <= value <= hi:
            return v
    return breaks[-1][2]

def lookup_list(value, lower, upper, vals):
    for i, (lo, hi) in enumerate(zip(lower, upper)):
        if lo <= value <= hi:
            return vals[i]
    return vals[-1]

def calc_Ka(x1, x2_cbr, x3_aadt, x4_age, x5_width, x6_terrain,
            y1_row, y2_shoulder, y3_terrain, y4_terrain, y5_bridge, y6_terrain):
    X1 = x1
    X2 = lookup_range(x2_cbr, X2_BREAKS)
    X3 = x3_aadt  # รับค่า factor โดยตรงจาก dropdown
    X4 = lookup_range(x4_age, X4_BREAKS)
    X5 = lookup_range(x5_width, X5_BREAKS)
    X6 = X6_MAP[x6_terrain]
    Y1 = lookup_range(y1_row, Y1_BREAKS)
    Y2 = lookup_range(y2_shoulder, Y2_BREAKS)
    Y3 = Y3_MAP[y3_terrain]
    Y4 = Y4_MAP[y4_terrain]
    Y5 = lookup_range(y5_bridge, Y5_BREAKS)
    Y6 = Y6_MAP[y6_terrain]
    Ka = 1 + 0.50 * (X1+X2+X3+X4+X5+X6+Y1+Y2+Y3+Y4+Y5+Y6)
    factors = {"X1":X1,"X2":X2,"X3":X3,"X4":X4,"X5":X5,"X6":X6,
               "Y1":Y1,"Y2":Y2,"Y3":Y3,"Y4":Y4,"Y5":Y5,"Y6":Y6}
    return Ka, factors

def calc_Kc(z1, z2_cbr, z3_aadt, z4_width,
            y1_row, y2_shoulder, y3_terrain, y4_terrain, y5_bridge, y6_terrain):
    Z1 = Z1_MAP.get(z1, 0)
    Z2 = lookup_range(z2_cbr, Z2_BREAKS)
    Z3 = z3_aadt  # รับค่า factor โดยตรงจาก dropdown
    Z4 = lookup_range(z4_width, Z4_BREAKS)
    Y1 = lookup_range(y1_row, Y1_BREAKS)
    Y2 = lookup_range(y2_shoulder, Y2_BREAKS)
    Y3 = Y3_MAP[y3_terrain]
    Y4 = Y4_MAP[y4_terrain]
    Y5 = lookup_range(y5_bridge, Y5_BREAKS)
    Y6 = Y6_MAP[y6_terrain]
    Kc = 1 + 0.50 * (Z1+Z2+Z3+Z4+Y1+Y2+Y3+Y4+Y5+Y6)
    factors = {"Z1":Z1,"Z2":Z2,"Z3":Z3,"Z4":Z4,
               "Y1":Y1,"Y2":Y2,"Y3":Y3,"Y4":Y4,"Y5":Y5,"Y6":Y6}
    return Kc, factors

def calc_Ks(a1_aadt, a3_width, b1_row, b2_terrain, b3_terrain, b4_bridge):
    A1 = lookup_range(a1_aadt, A1_BREAKS)
    A2 = 0.00  # ยังไม่มีข้อมูล (กรมทางหลวงกำลังศึกษา)
    A3 = lookup_range(a3_width, A3_BREAKS)
    B1 = lookup_range(b1_row, B1_BREAKS_KB)
    B2 = B2_MAP[b2_terrain]
    B3 = B3_MAP[b3_terrain]
    B4 = lookup_range(b4_bridge, B4_BREAKS)
    Ks = 1 + 0.70*(A1+A2+A3) + 0.30*(B1+B2+B3+B4)
    factors = {"A1":A1,"A2":A2,"A3":A3,"B1":B1,"B2":B2,"B3":B3,"B4":B4}
    return Ks, factors

def calc_K_prime(K, warranty_years):
    if warranty_years == 0:
        return K
    elif warranty_years == 1:
        return 0.5 * K
    else:
        return 0.25 * K

def calc_budget(dist_km, K, Km, N_std):
    raw = dist_km * K * Km * N_std
    return round(raw / 100) * 100  # ปรับเป็นหลักร้อย

def calc_workload(dist_equiv, K_prime):
    return round(dist_equiv * K_prime, 3)

# ─────────────────────────────────────────────
# SESSION STATE INIT
# ─────────────────────────────────────────────

def init_state():
    defaults = {
        "project_name": "โครงการบำรุงรักษาทางหลวง",
        "district": "",
        "year": "2568",
        "Na": 7000.0, "Ns": 6500.0, "Nc": 6000.0,
        "Km_a": 1.0, "Km_s": 1.0, "Km_c": 1.0,
        "rows_ac": [], "rows_cc": [], "rows_gr": [],
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

# ─────────────────────────────────────────────
# STYLES
# ─────────────────────────────────────────────

st.markdown("""
<style>
.main-title {font-size:1.5rem; font-weight:600; margin-bottom:0.2rem;}
.sub-title  {font-size:0.9rem; color:#666; margin-bottom:1.2rem;}
.k-value    {font-size:2rem; font-weight:700; color:#1a6b3c; text-align:center;}
.k-label    {font-size:0.8rem; color:#555; text-align:center; margin-top:-0.3rem;}
.budget-val {font-size:1.4rem; font-weight:700; color:#1565c0; text-align:center;}
.factor-item{display:flex; justify-content:space-between; padding:3px 0;
             border-bottom:1px solid #eee; font-size:0.82rem;}
.factor-name{color:#555;}
.factor-val {font-weight:600;}
.section-hd {font-weight:600; margin:1rem 0 0.4rem; color:#333;}
.note-box   {background:#fff8e1; border-left:4px solid #f9a825;
             padding:0.6rem 1rem; border-radius:4px; font-size:0.82rem; color:#5d4037;}
.success-box{background:#e8f5e9; border-left:4px solid #43a047;
             padding:0.6rem 1rem; border-radius:4px; font-size:0.82rem; color:#1b5e20;}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────

st.markdown('<div class="main-title">🛣️ ระบบคำนวณงานบำรุงปกติ</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">กองบำรุง กรมทางหลวง — คู่มือการคิดค่าปริมาณงานและงานบำรุงปกติ (2538)</div>', unsafe_allow_html=True)

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "⚙️ ตั้งค่าโครงการ",
    "🟫 ผิวแอสฟัลท์ (Ka)",
    "🟩 ผิวคอนกรีต (Kc)",
    "🟨 ผิวลูกรัง (Ks)",
    "📊 สรุปรวม & Export",
])

# ─────────────────────────────────────────────
# TAB 1: PROJECT SETTINGS
# ─────────────────────────────────────────────

with tab1:
    st.markdown("### ข้อมูลโครงการ")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.session_state["project_name"] = st.text_input("ชื่อโครงการ", st.session_state["project_name"])
    with c2:
        st.session_state["district"] = st.text_input("แขวงการทาง / สำนักงาน", st.session_state["district"])
    with c3:
        st.session_state["year"] = st.text_input("ปีงบประมาณ (พ.ศ.)", st.session_state["year"])

    st.markdown("---")
    st.markdown("### อัตราค่าบำรุงทางมาตรฐาน (N) และค่า Factor วัสดุ (Km)")
    st.markdown('<div class="note-box">💡 ค่าเริ่มต้นตามคู่มือกรมทางหลวง พ.ศ. 2538 — ผู้ใช้สามารถแก้ไขได้ตามปีงบประมาณปัจจุบัน</div>', unsafe_allow_html=True)
    st.markdown("")

    col_na, col_ns, col_nc = st.columns(3)
    with col_na:
        st.markdown("**ผิวแอสฟัลท์**")
        st.session_state["Na"] = st.number_input("Na มาตรฐาน (บาท/กม./ปี)", value=st.session_state["Na"], min_value=0.0, step=500.0, format="%.0f")
        st.session_state["Km_a"] = st.number_input("Km. วัสดุ (ลาดยาง)", value=st.session_state["Km_a"], min_value=0.01, step=0.01, format="%.3f")
    with col_ns:
        st.markdown("**ผิวลูกรัง**")
        st.session_state["Ns"] = st.number_input("Ns มาตรฐาน (บาท/กม./ปี)", value=st.session_state["Ns"], min_value=0.0, step=500.0, format="%.0f")
        st.session_state["Km_s"] = st.number_input("Km. วัสดุ (ลูกรัง)", value=st.session_state["Km_s"], min_value=0.01, step=0.01, format="%.3f")
    with col_nc:
        st.markdown("**ผิวคอนกรีต**")
        st.session_state["Nc"] = st.number_input("Nc มาตรฐาน (บาท/กม./ปี)", value=st.session_state["Nc"], min_value=0.0, step=500.0, format="%.0f")
        st.session_state["Km_c"] = st.number_input("Km. วัสดุ (คอนกรีต)", value=st.session_state["Km_c"], min_value=0.01, step=0.01, format="%.3f")

    st.markdown("---")
    st.markdown("### K' ปรับตามช่วงประกัน (Warranty Adjustment)")
    st.info("**K' = K** (ไม่มีประกัน)　|　**K' = 0.5 × K** (มีประกัน 1 ปี)　|　**K' = 0.25 × K** (มีประกัน > 1 ปี)\n\nWorkload = ระยะเทียบเท่า (กม.) × K'")

# ─────────────────────────────────────────────
# HELPER: Factor breakdown display
# ─────────────────────────────────────────────

def show_factor_breakdown(factors, K, K_prime, dist_km, lanes, budget, workload, surf_type):
    col_k, col_b, col_w = st.columns(3)
    with col_k:
        kp_label = f"K' = {K_prime:.4f}" if K != K_prime else ""
        st.markdown(f'<div class="k-value">{K:.4f}</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="k-label">K {surf_type} {kp_label}</div>', unsafe_allow_html=True)
    with col_b:
        st.markdown(f'<div class="budget-val">{budget:,.0f}</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="k-label">งบประมาณ (บาท/ปี)</div>', unsafe_allow_html=True)
    with col_w:
        st.markdown(f'<div class="budget-val" style="color:#6a1b9a">{workload:.3f}</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="k-label">Workload (หน่วย)</div>', unsafe_allow_html=True)

    with st.expander("📋 รายละเอียด Factor"):
        fc1, fc2 = st.columns(2)
        items = list(factors.items())
        half = len(items) // 2 + len(items) % 2
        with fc1:
            for k, v in items[:half]:
                st.markdown(f'<div class="factor-item"><span class="factor-name">{k}</span><span class="factor-val">{v:.4f}</span></div>', unsafe_allow_html=True)
        with fc2:
            for k, v in items[half:]:
                st.markdown(f'<div class="factor-item"><span class="factor-name">{k}</span><span class="factor-val">{v:.4f}</span></div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# HELPER: Common Y-factors form
# ─────────────────────────────────────────────

def y_factors_form(prefix="ac"):
    terrain_options = list(TERRAIN_MAP.keys())
    c1, c2, c3 = st.columns(3)
    with c1:
        y1 = st.number_input("Y1 ความกว้างเขตทาง (ม.)", value=40.0, min_value=0.0, step=5.0, key=f"{prefix}_y1")
        y2 = st.number_input("Y2 ไหล่ทางกว้างสุด 1 ข้าง (ม.)", value=2.50, min_value=0.0, step=0.25, key=f"{prefix}_y2")
    with c2:
        y3_lbl = st.selectbox("Y3 จราจรสงเคราะห์ (ภูมิประเทศ)", terrain_options, key=f"{prefix}_y3")
        y4_lbl = st.selectbox("Y4 ท่อระบายน้ำ (ภูมิประเทศ)", terrain_options, key=f"{prefix}_y4")
    with c3:
        y5 = st.number_input("Y5 สะพาน เฉลี่ย (ม./กม.)", value=0.0, min_value=0.0, step=1.0, key=f"{prefix}_y5")
        y6_lbl = st.selectbox("Y6 ทำความสะอาดระบาย (ภูมิประเทศ)", terrain_options, key=f"{prefix}_y6")
    return y1, y2, TERRAIN_MAP[y3_lbl], TERRAIN_MAP[y4_lbl], y5, TERRAIN_MAP[y6_lbl]

# ─────────────────────────────────────────────
# TAB 2: ASPHALT (Ka)
# ─────────────────────────────────────────────

with tab2:
    input_mode_ac = st.radio("วิธีป้อนข้อมูล", ["✏️ กรอกทีละสายทาง", "📂 Upload Excel (format Sheet C)"], horizontal=True, key="mode_ac")
    st.markdown("---")

    if input_mode_ac == "✏️ กรอกทีละสายทาง":
        st.markdown("#### ข้อมูลสายทาง")
        ca1, ca2, ca3, ca4 = st.columns(4)
        with ca1:
            route_id_ac = st.text_input("หมายเลข / ตอนควบคุม", key="ac_route_id")
            route_name_ac = st.text_input("ชื่อสายทาง", key="ac_route_name")
        with ca2:
            km_start_ac = st.number_input("กม. เริ่มต้น", value=0.000, step=0.001, format="%.3f", key="ac_km_s")
            km_end_ac = st.number_input("กม. สิ้นสุด", value=1.000, step=0.001, format="%.3f", key="ac_km_e")
        with ca3:
            lanes_ac = st.number_input("จำนวนช่องจราจร", value=2, min_value=1, step=1, key="ac_lanes")
            warranty_ac = st.number_input("มีประกันอีก (ปี) [0=ไม่มี]", value=0, min_value=0, step=1, key="ac_warranty")
        with ca4:
            dist_ac = km_end_ac - km_start_ac
            dist_equiv_ac = dist_ac * (lanes_ac / 2)
            st.metric("ระยะทาง (กม.)", f"{dist_ac:.3f}")
            st.metric("ระยะเทียบเท่า (กม.)", f"{dist_equiv_ac:.3f}")

        st.markdown("#### Factor X (ผิวทาง / ดิน / จราจร)")
        cx1, cx2, cx3 = st.columns(3)
        with cx1:
            x1_lbl = st.selectbox("X1 ลักษณะผิว+พื้นทาง", list(X1_MAP.keys()), key="ac_x1")
            x1_val = X1_MAP[x1_lbl]
            st.caption(f"X1 = {x1_val:.2f}")
        with cx2:
            x2_cbr = st.number_input("X2 CBR ดินเดิม (%)", value=5.0, min_value=0.0, step=0.5, key="ac_x2")
            x2_val = lookup_range(x2_cbr, X2_BREAKS)
            st.caption(f"X2 = {x2_val:.2f}")
        with cx3:
            x3_lbl = st.selectbox("X3 AADT ต่อ 2 ช่อง (คัน/วัน)", list(X3_OPTIONS.keys()), index=5, key="ac_x3")
            x3_val = X3_OPTIONS[x3_lbl]

        cx4, cx5, cx6 = st.columns(3)
        with cx4:
            x4_age = st.number_input("X4 อายุบริการ (ปี)", value=5, min_value=0, step=1, key="ac_x4")
            x4_val = lookup_range(x4_age, X4_BREAKS)
            st.caption(f"X4 = {x4_val:.2f}")
        with cx5:
            x5_width = st.number_input("X5 ความกว้างผิวทาง ต่อ 2 ช่อง (ม.)", value=7.0, min_value=0.0, step=0.5, key="ac_x5")
            x5_val = lookup_range(x5_width, X5_BREAKS)
            st.caption(f"X5 = {x5_val:.2f}")
        with cx6:
            x6_lbl = st.selectbox("X6 ภูมิประเทศ", list(TERRAIN_MAP.keys()), key="ac_x6")
            x6_val = X6_MAP[TERRAIN_MAP[x6_lbl]]
            st.caption(f"X6 = {x6_val:.2f}")

        st.markdown("#### Factor Y (เขตทาง / สิ่งก่อสร้าง)")
        y1, y2, y3t, y4t, y5, y6t = y_factors_form("ac")

        Ka, factors_ac = calc_Ka(x1_val, x2_cbr, x3_val, x4_age, x5_width,
                                  TERRAIN_MAP[x6_lbl], y1, y2, y3t, y4t, y5, y6t)
        Kap = calc_K_prime(Ka, warranty_ac)
        budget_ac = calc_budget(dist_ac, Ka, st.session_state["Km_a"], st.session_state["Na"])
        workload_ac = calc_workload(dist_equiv_ac, Kap)

        st.markdown("---")
        st.markdown("#### ผลการคำนวณ")
        show_factor_breakdown(factors_ac, Ka, Kap, dist_ac, lanes_ac, budget_ac, workload_ac, "แอสฟัลท์")

        if st.button("➕ เพิ่มสายทางนี้ลงตาราง", type="primary", key="add_ac"):
            if dist_ac <= 0:
                st.error("ระยะทางต้องมากกว่า 0")
            else:
                row = {
                    "ตอนควบคุม": route_id_ac, "ชื่อสายทาง": route_name_ac,
                    "กม.เริ่ม": km_start_ac, "กม.สิ้นสุด": km_end_ac,
                    "ระยะทาง(กม.)": round(dist_ac, 3), "ช่องจราจร": lanes_ac,
                    "ระยะเทียบเท่า(กม.)": round(dist_equiv_ac, 3),
                    **{k: f"{v:.4f}" for k, v in factors_ac.items()},
                    "K": round(Ka, 4), "ประกัน(ปี)": warranty_ac,
                    "K'": round(Kap, 4), "Workload(หน่วย)": workload_ac,
                    "งบประมาณ(บาท/ปี)": budget_ac,
                }
                st.session_state["rows_ac"].append(row)
                st.success(f"เพิ่มสายทาง '{route_name_ac}' แล้ว")

    else:
        st.markdown("#### Upload Excel (รูปแบบ Sheet C ของ KC.xlsx)")
        st.markdown('<div class="note-box">📌 คอลัมน์ที่ต้องการ: ตอนควบคุม, ชื่อสายทาง, กม.เริ่ม, กม.สิ้นสุด, ช่องจราจร, X1(รหัส h/i/l), X2(CBR), X3(AADT), X4(อายุ), X5(กว้างผิว), X6(P/R/RM/S), Y1-Y6, มีประกันอีก(ปี)</div>', unsafe_allow_html=True)
        uploaded_ac = st.file_uploader("เลือกไฟล์ Excel", type=["xlsx"], key="up_ac")
        if uploaded_ac:
            try:
                df_up = pd.read_excel(uploaded_ac, header=2)
                st.dataframe(df_up.head(10), use_container_width=True)
                st.info(f"พบ {len(df_up)} แถว — กดปุ่มด้านล่างเพื่อคำนวณ K อัตโนมัติ")
                if st.button("⚙️ คำนวณ K จากไฟล์", key="calc_up_ac"):
                    new_rows = []
                    for _, r in df_up.iterrows():
                        try:
                            dist = float(r.get("ระยะจริง\n(กม.)", 0) or (float(r.iloc[4]) - float(r.iloc[2])) / 1000)
                            ln = int(r.get("ช่อง\nจราจร", 2) or 2)
                            eq = dist * (ln / 2)
                            x1_code = str(r.get("X1", "h") or "h").lower()
                            x1_v = {"h": 0.0, "i": 0.5, "l": 1.0}.get(x1_code, 0.0)
                            cbr = float(r.get("X2", 5) or 5)
                            aadt = int(r.get("X3", 500) or 500)
                            age = int(r.get("X4", 5) or 5)
                            w = float(r.get("X5", 7) or 7)
                            x6t = str(r.get("X6", "P") or "P").upper()
                            y1v = float(r.get("Y1", 40) or 40)
                            y2v = float(r.get("Y2", 2.5) or 2.5)
                            y3t = str(r.get("Y3", "P") or "P").upper()
                            y4t = str(r.get("Y4", "P") or "P").upper()
                            y5v = float(r.get("Y5", 0) or 0)
                            y6t = str(r.get("Y6", "P") or "P").upper()
                            war = int(r.get("มีประกัน\nอีก(ปี)", 0) or 0)
                            Ka_, fac_ = calc_Ka(x1_v, cbr, aadt, age, w, x6t, y1v, y2v, y3t, y4t, y5v, y6t)
                            Kap_ = calc_K_prime(Ka_, war)
                            bud_ = calc_budget(dist, Ka_, st.session_state["Km_a"], st.session_state["Na"])
                            wl_ = calc_workload(eq, Kap_)
                            new_rows.append({
                                "ตอนควบคุม": r.iloc[0], "ชื่อสายทาง": r.iloc[1],
                                "ระยะทาง(กม.)": round(dist, 3), "ช่องจราจร": ln,
                                "ระยะเทียบเท่า(กม.)": round(eq, 3),
                                "K": round(Ka_, 4), "ประกัน(ปี)": war,
                                "K'": round(Kap_, 4), "Workload(หน่วย)": wl_,
                                "งบประมาณ(บาท/ปี)": bud_,
                            })
                        except Exception:
                            continue
                    st.session_state["rows_ac"].extend(new_rows)
                    st.success(f"เพิ่ม {len(new_rows)} สายทางแล้ว")
            except Exception as e:
                st.error(f"อ่านไฟล์ไม่ได้: {e}")

    # แสดงตาราง
    if st.session_state["rows_ac"]:
        st.markdown("---")
        st.markdown(f"#### 📋 ตารางสายทางผิวแอสฟัลท์ ({len(st.session_state['rows_ac'])} สายทาง)")
        df_ac = pd.DataFrame(st.session_state["rows_ac"])
        st.dataframe(df_ac[["ตอนควบคุม","ชื่อสายทาง","ระยะทาง(กม.)","K","K'","Workload(หน่วย)","งบประมาณ(บาท/ปี)"]], use_container_width=True)
        tot_bud = sum(r["งบประมาณ(บาท/ปี)"] for r in st.session_state["rows_ac"])
        tot_wl  = sum(r["Workload(หน่วย)"] for r in st.session_state["rows_ac"])
        m1, m2 = st.columns(2)
        m1.metric("งบประมาณรวม (บาท/ปี)", f"{tot_bud:,.0f}")
        m2.metric("Workload รวม (หน่วย)", f"{tot_wl:.3f}")
        if st.button("🗑️ ล้างตารางทั้งหมด", key="clr_ac"):
            st.session_state["rows_ac"] = []
            st.rerun()

# ─────────────────────────────────────────────
# TAB 3: CONCRETE (Kc)
# ─────────────────────────────────────────────

with tab3:
    input_mode_cc = st.radio("วิธีป้อนข้อมูล", ["✏️ กรอกทีละสายทาง", "📂 Upload Excel (format Sheet C)"], horizontal=True, key="mode_cc")
    st.markdown("---")

    if input_mode_cc == "✏️ กรอกทีละสายทาง":
        st.markdown("#### ข้อมูลสายทาง")
        cc1, cc2, cc3, cc4 = st.columns(4)
        with cc1:
            route_id_cc = st.text_input("หมายเลข / ตอนควบคุม", key="cc_route_id")
            route_name_cc = st.text_input("ชื่อสายทาง", key="cc_route_name")
        with cc2:
            km_start_cc = st.number_input("กม. เริ่มต้น", value=0.000, step=0.001, format="%.3f", key="cc_km_s")
            km_end_cc = st.number_input("กม. สิ้นสุด", value=1.000, step=0.001, format="%.3f", key="cc_km_e")
        with cc3:
            lanes_cc = st.number_input("จำนวนช่องจราจร", value=2, min_value=1, step=1, key="cc_lanes")
            warranty_cc = st.number_input("มีประกันอีก (ปี) [0=ไม่มี]", value=0, min_value=0, step=1, key="cc_warranty")
        with cc4:
            dist_cc = km_end_cc - km_start_cc
            dist_equiv_cc = dist_cc * (lanes_cc / 2)
            st.metric("ระยะทาง (กม.)", f"{dist_cc:.3f}")
            st.metric("ระยะเทียบเท่า (กม.)", f"{dist_equiv_cc:.3f}")

        st.markdown("#### Factor Z (สำหรับผิวคอนกรีต)")
        cz1, cz2, cz3, cz4 = st.columns(4)
        with cz1:
            z1_idx = st.selectbox("Z1 ดัชนีสภาพผิวทาง (1-8)", list(range(1, 9)), index=0, key="cc_z1")
            st.caption(f"Z1 = {Z1_MAP[z1_idx]:.2f}  (ความเสียหาย {z1_idx}%)")
        with cz2:
            z2_cbr = st.number_input("Z2 CBR ดินคันทาง (%)", value=5.0, min_value=0.0, step=0.5, key="cc_z2")
            z2_val = lookup_range(z2_cbr, Z2_BREAKS)
            st.caption(f"Z2 = {z2_val:.2f}")
        with cz3:
            z3_lbl = st.selectbox("Z3 AADT ต่อ 2 ช่อง (คัน/วัน)", list(Z3_OPTIONS.keys()), index=4, key="cc_z3")
            z3_val = Z3_OPTIONS[z3_lbl]
        with cz4:
            z4_width = st.number_input("Z4 ความกว้างผิวทาง ต่อ 2 ช่อง (ม.)", value=7.0, min_value=0.0, step=0.5, key="cc_z4")
            z4_val = lookup_range(z4_width, Z4_BREAKS)
            st.caption(f"Z4 = {z4_val:.2f}")

        st.markdown("#### Factor Y (เขตทาง / สิ่งก่อสร้าง)")
        y1c, y2c, y3ct, y4ct, y5c, y6ct = y_factors_form("cc")

        Kc, factors_cc = calc_Kc(z1_idx, z2_cbr, z3_val, z4_width,
                                   y1c, y2c, y3ct, y4ct, y5c, y6ct)
        Kcp = calc_K_prime(Kc, warranty_cc)
        budget_cc = calc_budget(dist_cc, Kc, st.session_state["Km_c"], st.session_state["Nc"])
        workload_cc = calc_workload(dist_equiv_cc, Kcp)

        st.markdown("---")
        st.markdown("#### ผลการคำนวณ")
        show_factor_breakdown(factors_cc, Kc, Kcp, dist_cc, lanes_cc, budget_cc, workload_cc, "คอนกรีต")

        if st.button("➕ เพิ่มสายทางนี้ลงตาราง", type="primary", key="add_cc"):
            if dist_cc <= 0:
                st.error("ระยะทางต้องมากกว่า 0")
            else:
                row = {
                    "ตอนควบคุม": route_id_cc, "ชื่อสายทาง": route_name_cc,
                    "กม.เริ่ม": km_start_cc, "กม.สิ้นสุด": km_end_cc,
                    "ระยะทาง(กม.)": round(dist_cc, 3), "ช่องจราจร": lanes_cc,
                    "ระยะเทียบเท่า(กม.)": round(dist_equiv_cc, 3),
                    **{k: f"{v:.4f}" for k, v in factors_cc.items()},
                    "K": round(Kc, 4), "ประกัน(ปี)": warranty_cc,
                    "K'": round(Kcp, 4), "Workload(หน่วย)": workload_cc,
                    "งบประมาณ(บาท/ปี)": budget_cc,
                }
                st.session_state["rows_cc"].append(row)
                st.success(f"เพิ่มสายทาง '{route_name_cc}' แล้ว")

    else:
        st.markdown("#### Upload Excel (รูปแบบ KC.xlsx Sheet C)")
        uploaded_cc = st.file_uploader("เลือกไฟล์ Excel", type=["xlsx"], key="up_cc")
        if uploaded_cc:
            try:
                df_upc = pd.read_excel(uploaded_cc, sheet_name="C", header=2)
                cols_needed = ["ตอน\nควบคุม","ชื่อสายทาง\n","กม.","กม.","ระยะจริง\n(กม.)","ช่อง\nจราจร","Z1","Z2","Z3","Z4","Y1","Y2","Y3","Y4","Y5","Y6","มีประกัน\nอีก(ปี)"]
                st.dataframe(df_upc.head(10), use_container_width=True)
                st.info(f"พบ {len(df_upc)} แถว")
                if st.button("⚙️ คำนวณ Kc จากไฟล์", key="calc_up_cc"):
                    new_rows_c = []
                    for _, r in df_upc.iterrows():
                        try:
                            dist = float(r.iloc[5] if pd.notna(r.iloc[5]) else 0)
                            if dist <= 0:
                                continue
                            ln = int(r.iloc[6] if pd.notna(r.iloc[6]) else 2)
                            eq = dist * (ln / 2)
                            z1i = int(r.iloc[8] if pd.notna(r.iloc[8]) else 1)
                            z2c = float(r.iloc[9] if pd.notna(r.iloc[9]) else 5)
                            z3a = int(r.iloc[10] if pd.notna(r.iloc[10]) else 1000)
                            z4w = float(r.iloc[11] if pd.notna(r.iloc[11]) else 7)
                            y1v = float(r.iloc[12] if pd.notna(r.iloc[12]) else 40)
                            y2v = float(r.iloc[13] if pd.notna(r.iloc[13]) else 2.5)
                            y3t = str(r.iloc[14] if pd.notna(r.iloc[14]) else "P")
                            y4t = str(r.iloc[15] if pd.notna(r.iloc[15]) else "P")
                            y5v = float(r.iloc[16] if pd.notna(r.iloc[16]) else 0)
                            y6t = str(r.iloc[17] if pd.notna(r.iloc[17]) else "P")
                            war = int(r.iloc[21] if pd.notna(r.iloc[21]) else 0)
                            Kc_, fac_ = calc_Kc(z1i, z2c, z3a, z4w, y1v, y2v, y3t, y4t, y5v, y6t)
                            Kcp_ = calc_K_prime(Kc_, war)
                            bud_ = calc_budget(dist, Kc_, st.session_state["Km_c"], st.session_state["Nc"])
                            wl_ = calc_workload(eq, Kcp_)
                            new_rows_c.append({
                                "ตอนควบคุม": r.iloc[0], "ชื่อสายทาง": r.iloc[1],
                                "ระยะทาง(กม.)": round(dist, 3), "ช่องจราจร": ln,
                                "ระยะเทียบเท่า(กม.)": round(eq, 3),
                                "K": round(Kc_, 4), "ประกัน(ปี)": war,
                                "K'": round(Kcp_, 4), "Workload(หน่วย)": wl_,
                                "งบประมาณ(บาท/ปี)": bud_,
                            })
                        except Exception:
                            continue
                    st.session_state["rows_cc"].extend(new_rows_c)
                    st.success(f"เพิ่ม {len(new_rows_c)} สายทางแล้ว")
            except Exception as e:
                st.error(f"อ่านไฟล์ไม่ได้: {e}")

    if st.session_state["rows_cc"]:
        st.markdown("---")
        st.markdown(f"#### 📋 ตารางสายทางผิวคอนกรีต ({len(st.session_state['rows_cc'])} สายทาง)")
        df_cc = pd.DataFrame(st.session_state["rows_cc"])
        st.dataframe(df_cc[["ตอนควบคุม","ชื่อสายทาง","ระยะทาง(กม.)","K","K'","Workload(หน่วย)","งบประมาณ(บาท/ปี)"]], use_container_width=True)
        tot_bud_c = sum(r["งบประมาณ(บาท/ปี)"] for r in st.session_state["rows_cc"])
        tot_wl_c  = sum(r["Workload(หน่วย)"] for r in st.session_state["rows_cc"])
        m1c, m2c = st.columns(2)
        m1c.metric("งบประมาณรวม (บาท/ปี)", f"{tot_bud_c:,.0f}")
        m2c.metric("Workload รวม (หน่วย)", f"{tot_wl_c:.3f}")
        if st.button("🗑️ ล้างตารางทั้งหมด", key="clr_cc"):
            st.session_state["rows_cc"] = []
            st.rerun()

# ─────────────────────────────────────────────
# TAB 4: GRAVEL (Ks)
# ─────────────────────────────────────────────

with tab4:
    st.markdown('<div class="note-box">⚠️ A2 (ลักษณะลมฟ้าอากาศ) — กรมทางหลวงกำลังศึกษาและเก็บสถิติ ยังไม่มีค่า Factor จึงใช้ A2 = 0.00</div>', unsafe_allow_html=True)
    st.markdown("---")

    st.markdown("#### ข้อมูลสายทาง")
    cg1, cg2, cg3, cg4 = st.columns(4)
    with cg1:
        route_id_gr = st.text_input("หมายเลข / ตอนควบคุม", key="gr_route_id")
        route_name_gr = st.text_input("ชื่อสายทาง", key="gr_route_name")
    with cg2:
        km_start_gr = st.number_input("กม. เริ่มต้น", value=0.000, step=0.001, format="%.3f", key="gr_km_s")
        km_end_gr = st.number_input("กม. สิ้นสุด", value=1.000, step=0.001, format="%.3f", key="gr_km_e")
    with cg3:
        lanes_gr = st.number_input("จำนวนช่องจราจร", value=2, min_value=1, step=1, key="gr_lanes")
        warranty_gr = st.number_input("มีประกันอีก (ปี) [0=ไม่มี]", value=0, min_value=0, step=1, key="gr_warranty")
    with cg4:
        dist_gr = km_end_gr - km_start_gr
        dist_equiv_gr = dist_gr * (lanes_gr / 2)
        st.metric("ระยะทาง (กม.)", f"{dist_gr:.3f}")
        st.metric("ระยะเทียบเท่า (กม.)", f"{dist_equiv_gr:.3f}")

    st.markdown("#### Factor A (ปริมาณจราจร / ลมฟ้าอากาศ / ความกว้าง)")
    cga1, cga2, cga3 = st.columns(3)
    with cga1:
        a1_aadt = st.number_input("A1 ADT (คัน/วัน)", value=300, min_value=0, step=10, key="gr_a1")
        a1_val = lookup_range(a1_aadt, A1_BREAKS)
        st.caption(f"A1 = {a1_val:.2f}")
    with cga2:
        st.metric("A2 ลมฟ้าอากาศ", "0.00")
        st.caption("ยังไม่มีข้อมูล — กรมทางหลวงกำลังศึกษา")
    with cga3:
        a3_width = st.number_input("A3 ความกว้างคันทาง (ผิว+ไหล่ทาง) (ม.)", value=7.0, min_value=0.0, step=0.5, key="gr_a3")
        a3_val = lookup_range(a3_width, A3_BREAKS)
        st.caption(f"A3 = {a3_val:.2f}")

    st.markdown("#### Factor B (เขตทาง / จราจรสงเคราะห์ / ระบายน้ำ / สะพาน)")
    terrain_options = list(TERRAIN_MAP.keys())
    cgb1, cgb2, cgb3, cgb4 = st.columns(4)
    with cgb1:
        b1_row = st.number_input("B1 ความกว้างเขตทาง (ม.)", value=30.0, min_value=0.0, step=5.0, key="gr_b1")
        b1_val = lookup_range(b1_row, B1_BREAKS_KB)
        st.caption(f"B1 = {b1_val:.2f}")
    with cgb2:
        b2_lbl = st.selectbox("B2 จราจรสงเคราะห์ (ภูมิประเทศ)", terrain_options, key="gr_b2")
        b2_val = B2_MAP[TERRAIN_MAP[b2_lbl]]
        st.caption(f"B2 = {b2_val:.2f}")
    with cgb3:
        b3_lbl = st.selectbox("B3 ระบายน้ำ (ภูมิประเทศ)", terrain_options, key="gr_b3")
        b3_val = B3_MAP[TERRAIN_MAP[b3_lbl]]
        st.caption(f"B3 = {b3_val:.2f}")
    with cgb4:
        b4_bridge = st.number_input("B4 ความยาวสะพาน (ม./กม.)", value=0.0, min_value=0.0, step=1.0, key="gr_b4")
        b4_val = lookup_range(b4_bridge, B4_BREAKS)
        st.caption(f"B4 = {b4_val:.2f}")

    Ks, factors_gr = calc_Ks(a1_aadt, a3_width, b1_row, TERRAIN_MAP[b2_lbl], TERRAIN_MAP[b3_lbl], b4_bridge)
    Ksp = calc_K_prime(Ks, warranty_gr)
    budget_gr = calc_budget(dist_gr, Ks, st.session_state["Km_s"], st.session_state["Ns"])
    workload_gr = calc_workload(dist_equiv_gr, Ksp)

    st.markdown("---")
    st.markdown("#### ผลการคำนวณ")
    show_factor_breakdown(factors_gr, Ks, Ksp, dist_gr, lanes_gr, budget_gr, workload_gr, "ลูกรัง")

    if st.button("➕ เพิ่มสายทางนี้ลงตาราง", type="primary", key="add_gr"):
        if dist_gr <= 0:
            st.error("ระยะทางต้องมากกว่า 0")
        else:
            row = {
                "ตอนควบคุม": route_id_gr, "ชื่อสายทาง": route_name_gr,
                "กม.เริ่ม": km_start_gr, "กม.สิ้นสุด": km_end_gr,
                "ระยะทาง(กม.)": round(dist_gr, 3), "ช่องจราจร": lanes_gr,
                "ระยะเทียบเท่า(กม.)": round(dist_equiv_gr, 3),
                **{k: f"{v:.4f}" for k, v in factors_gr.items()},
                "K": round(Ks, 4), "ประกัน(ปี)": warranty_gr,
                "K'": round(Ksp, 4), "Workload(หน่วย)": workload_gr,
                "งบประมาณ(บาท/ปี)": budget_gr,
            }
            st.session_state["rows_gr"].append(row)
            st.success(f"เพิ่มสายทาง '{route_name_gr}' แล้ว")

    if st.session_state["rows_gr"]:
        st.markdown("---")
        st.markdown(f"#### 📋 ตารางสายทางผิวลูกรัง ({len(st.session_state['rows_gr'])} สายทาง)")
        df_gr = pd.DataFrame(st.session_state["rows_gr"])
        st.dataframe(df_gr[["ตอนควบคุม","ชื่อสายทาง","ระยะทาง(กม.)","K","K'","Workload(หน่วย)","งบประมาณ(บาท/ปี)"]], use_container_width=True)
        tot_bud_g = sum(r["งบประมาณ(บาท/ปี)"] for r in st.session_state["rows_gr"])
        tot_wl_g  = sum(r["Workload(หน่วย)"] for r in st.session_state["rows_gr"])
        m1g, m2g = st.columns(2)
        m1g.metric("งบประมาณรวม (บาท/ปี)", f"{tot_bud_g:,.0f}")
        m2g.metric("Workload รวม (หน่วย)", f"{tot_wl_g:.3f}")
        if st.button("🗑️ ล้างตารางทั้งหมด", key="clr_gr"):
            st.session_state["rows_gr"] = []
            st.rerun()

# ─────────────────────────────────────────────
# TAB 5: SUMMARY & EXPORT
# ─────────────────────────────────────────────

with tab5:
    st.markdown(f"### 📊 สรุปรวม — {st.session_state['project_name']}")
    st.caption(f"{st.session_state['district']}  |  ปีงบประมาณ {st.session_state['year']}")

    n_ac = len(st.session_state["rows_ac"])
    n_cc = len(st.session_state["rows_cc"])
    n_gr = len(st.session_state["rows_gr"])

    bud_ac = sum(r["งบประมาณ(บาท/ปี)"] for r in st.session_state["rows_ac"])
    bud_cc = sum(r["งบประมาณ(บาท/ปี)"] for r in st.session_state["rows_cc"])
    bud_gr = sum(r["งบประมาณ(บาท/ปี)"] for r in st.session_state["rows_gr"])
    bud_total = bud_ac + bud_cc + bud_gr

    wl_ac = sum(r["Workload(หน่วย)"] for r in st.session_state["rows_ac"])
    wl_cc = sum(r["Workload(หน่วย)"] for r in st.session_state["rows_cc"])
    wl_gr = sum(r["Workload(หน่วย)"] for r in st.session_state["rows_gr"])
    wl_total = wl_ac + wl_cc + wl_gr

    dist_ac_tot = sum(r["ระยะทาง(กม.)"] for r in st.session_state["rows_ac"])
    dist_cc_tot = sum(r["ระยะทาง(กม.)"] for r in st.session_state["rows_cc"])
    dist_gr_tot = sum(r["ระยะทาง(กม.)"] for r in st.session_state["rows_gr"])
    dist_total  = dist_ac_tot + dist_cc_tot + dist_gr_tot

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("ระยะทางรวม (กม.)", f"{dist_total:.3f}")
    c2.metric("งบประมาณรวม (บาท/ปี)", f"{bud_total:,.0f}")
    c3.metric("Workload รวม (หน่วย)", f"{wl_total:.3f}")
    c4.metric("จำนวนสายทางทั้งหมด", f"{n_ac+n_cc+n_gr}")

    st.markdown("---")
    st.markdown("#### สรุปแยกตามประเภทผิวทาง")
    summary_data = {
        "ประเภทผิวทาง": ["🟫 แอสฟัลท์ (Ka)", "🟩 คอนกรีต (Kc)", "🟨 ลูกรัง (Ks)", "**รวม**"],
        "จำนวนสายทาง": [n_ac, n_cc, n_gr, n_ac+n_cc+n_gr],
        "ระยะทาง (กม.)": [round(dist_ac_tot,3), round(dist_cc_tot,3), round(dist_gr_tot,3), round(dist_total,3)],
        "Workload (หน่วย)": [round(wl_ac,3), round(wl_cc,3), round(wl_gr,3), round(wl_total,3)],
        "งบประมาณ (บาท/ปี)": [f"{bud_ac:,.0f}", f"{bud_cc:,.0f}", f"{bud_gr:,.0f}", f"{bud_total:,.0f}"],
    }
    st.dataframe(pd.DataFrame(summary_data), use_container_width=True, hide_index=True)

    st.markdown("---")

    # ── Export Excel ──────────────────────────
    def generate_excel():
        wb = openpyxl.Workbook()

        hdr_font  = Font(name="TH SarabunPSK", bold=True, size=13)
        data_font = Font(name="TH SarabunPSK", size=12)
        title_font= Font(name="TH SarabunPSK", bold=True, size=14)
        center    = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left_al   = Alignment(horizontal="left", vertical="center")
        fill_hdr  = PatternFill("solid", fgColor="1565C0")
        fill_sum  = PatternFill("solid", fgColor="E8F5E9")
        fill_tot  = PatternFill("solid", fgColor="FFF9C4")
        thin      = Side(style="thin", color="BBBBBB")
        border    = Border(left=thin, right=thin, top=thin, bottom=thin)

        def write_sheet(ws, rows, surf_type, factors_cols):
            ws.sheet_view.showGridLines = False
            # Title
            ws.merge_cells("A1:M1")
            ws["A1"] = f"งานบำรุงปกติ — ผิว{surf_type}  |  {st.session_state['project_name']}  |  {st.session_state['district']}  |  ปีงบประมาณ {st.session_state['year']}"
            ws["A1"].font = title_font
            ws["A1"].alignment = center

            headers = ["ตอนควบคุม","ชื่อสายทาง","กม.เริ่ม","กม.สิ้นสุด",
                       "ระยะทาง\n(กม.)","ช่องจราจร","ระยะเทียบเท่า\n(กม.)"] + \
                      factors_cols + ["K","ประกัน\n(ปี)","K'","Workload\n(หน่วย)","งบประมาณ\n(บาท/ปี)"]

            for ci, h in enumerate(headers, 1):
                cell = ws.cell(row=2, column=ci, value=h)
                cell.font = Font(name="TH SarabunPSK", bold=True, size=12, color="FFFFFF")
                cell.fill = fill_hdr
                cell.alignment = center
                cell.border = border

            for ri, row in enumerate(rows, 3):
                vals = [row.get("ตอนควบคุม",""), row.get("ชื่อสายทาง",""),
                        row.get("กม.เริ่ม",0), row.get("กม.สิ้นสุด",0),
                        row.get("ระยะทาง(กม.)",0), row.get("ช่องจราจร",2),
                        row.get("ระยะเทียบเท่า(กม.)",0)] + \
                       [float(row.get(fc,0)) for fc in factors_cols] + \
                       [row.get("K",0), row.get("ประกัน(ปี)",0),
                        row.get("K'",0), row.get("Workload(หน่วย)",0),
                        row.get("งบประมาณ(บาท/ปี)",0)]
                for ci, v in enumerate(vals, 1):
                    cell = ws.cell(row=ri, column=ci, value=v)
                    cell.font = data_font
                    cell.border = border
                    if isinstance(v, (int, float)):
                        cell.alignment = center
                    else:
                        cell.alignment = left_al

            # Total row
            if rows:
                tr = len(rows) + 3
                last_col = len(headers)
                for ci in range(1, last_col+1):
                    cell = ws.cell(row=tr, column=ci)
                    cell.fill = fill_tot
                    cell.border = border
                    cell.font = Font(name="TH SarabunPSK", bold=True, size=12)
                    cell.alignment = center
                ws.cell(row=tr, column=1, value="รวม")
                ws.cell(row=tr, column=5, value=sum(r.get("ระยะทาง(กม.)",0) for r in rows))
                ws.cell(row=tr, column=7, value=sum(r.get("ระยะเทียบเท่า(กม.)",0) for r in rows))
                wl_col = last_col - 1
                bud_col = last_col
                ws.cell(row=tr, column=wl_col, value=round(sum(r.get("Workload(หน่วย)",0) for r in rows),3))
                ws.cell(row=tr, column=bud_col, value=sum(r.get("งบประมาณ(บาท/ปี)",0) for r in rows))
                ws.cell(row=tr, column=bud_col).number_format = "#,##0"

            ws.column_dimensions["A"].width = 14
            ws.column_dimensions["B"].width = 24
            ws.column_dimensions["C"].width = 10
            ws.column_dimensions["D"].width = 10
            for i in range(5, last_col+1):
                ws.column_dimensions[get_column_letter(i)].width = 10
            ws.row_dimensions[1].height = 22
            ws.row_dimensions[2].height = 36

        # AC sheet
        ws_ac = wb.active
        ws_ac.title = "ผิวแอสฟัลท์"
        write_sheet(ws_ac, st.session_state["rows_ac"], "แอสฟัลท์",
                    ["X1","X2","X3","X4","X5","X6","Y1","Y2","Y3","Y4","Y5","Y6"])

        # Concrete sheet
        ws_cc = wb.create_sheet("ผิวคอนกรีต")
        write_sheet(ws_cc, st.session_state["rows_cc"], "คอนกรีต",
                    ["Z1","Z2","Z3","Z4","Y1","Y2","Y3","Y4","Y5","Y6"])

        # Gravel sheet
        ws_gr = wb.create_sheet("ผิวลูกรัง")
        write_sheet(ws_gr, st.session_state["rows_gr"], "ลูกรัง",
                    ["A1","A2","A3","B1","B2","B3","B4"])

        # Summary sheet
        ws_sum = wb.create_sheet("สรุปรวม")
        ws_sum.sheet_view.showGridLines = False
        ws_sum.merge_cells("A1:F1")
        ws_sum["A1"] = f"สรุปงานบำรุงปกติ  |  {st.session_state['project_name']}  |  {st.session_state['district']}  |  ปีงบประมาณ {st.session_state['year']}"
        ws_sum["A1"].font = title_font
        ws_sum["A1"].alignment = center

        sum_hdrs = ["ประเภทผิวทาง","จำนวนสายทาง","ระยะทาง (กม.)","Workload (หน่วย)","งบประมาณ (บาท/ปี)"]
        for ci, h in enumerate(sum_hdrs, 1):
            cell = ws_sum.cell(row=2, column=ci, value=h)
            cell.font = Font(name="TH SarabunPSK", bold=True, size=12, color="FFFFFF")
            cell.fill = fill_hdr
            cell.alignment = center
            cell.border = border

        sum_rows = [
            ("ผิวแอสฟัลท์ (Ka)", n_ac, round(dist_ac_tot,3), round(wl_ac,3), bud_ac),
            ("ผิวคอนกรีต (Kc)",  n_cc, round(dist_cc_tot,3), round(wl_cc,3), bud_cc),
            ("ผิวลูกรัง (Ks)",    n_gr, round(dist_gr_tot,3), round(wl_gr,3), bud_gr),
            ("รวม",               n_ac+n_cc+n_gr, round(dist_total,3), round(wl_total,3), bud_total),
        ]
        for ri, row_data in enumerate(sum_rows, 3):
            is_total = ri == len(sum_rows) + 2
            for ci, v in enumerate(row_data, 1):
                cell = ws_sum.cell(row=ri, column=ci, value=v)
                cell.font = Font(name="TH SarabunPSK", bold=is_total, size=12)
                cell.fill = fill_tot if is_total else fill_sum
                cell.alignment = center
                cell.border = border
                if ci == 5:
                    cell.number_format = "#,##0"

        for col, w in zip(["A","B","C","D","E"], [22,14,16,16,20]):
            ws_sum.column_dimensions[col].width = w

        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    col_exp1, col_exp2 = st.columns(2)
    with col_exp1:
        if bud_total > 0:
            excel_buf = generate_excel()
            fname = f"routine_maintenance_{st.session_state['year']}_{st.session_state['project_name'].replace(' ','_')}.xlsx"
            st.download_button(
                "📥 Export Excel (รายงานสรุป)",
                data=excel_buf,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
        else:
            st.info("ยังไม่มีข้อมูลสายทาง")

    with col_exp2:
        if bud_total > 0:
            if st.button("📤 ส่งค่า Routine Cost ให้โปรแกรม LCCA", type="secondary"):
                routine_cost_per_km = bud_total / dist_total if dist_total > 0 else 0
                st.session_state["routine_to_lcca"] = {
                    "total_budget_per_year": bud_total,
                    "total_distance_km": round(dist_total, 3),
                    "routine_cost_per_km_per_year": round(routine_cost_per_km, 2),
                    "workload_total": round(wl_total, 3),
                    "project": st.session_state["project_name"],
                    "year": st.session_state["year"],
                }
                st.success(f"✅ ส่งค่าแล้ว! Routine Cost = {routine_cost_per_km:,.2f} บาท/กม./ปี → session_state['routine_to_lcca']")
                st.json(st.session_state["routine_to_lcca"])
