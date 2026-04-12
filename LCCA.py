#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
โปรแกรมวิเคราะห์ต้นทุนตลอดอายุการใช้งานผิวทาง (LCCA) - Integrated v1.0
Life-Cycle Cost Analysis for Pavement Alternatives
================================================================================
พัฒนาโดย: รศ.ดร.อิทธิพล มีผล
ภาควิชาครุศาสตร์โยธา มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าพระนครเหนือ (KMUTNB)

โครงสร้างโปรแกรม:
  TAB 1: ข้อมูลโครงการ + ราคาก่อสร้าง (Upload Excel / กรอกมือ)
  TAB 2: Routine Cost (Ka สำหรับ AC | Kc สำหรับ Concrete) → Ka/Kc เฉลี่ย
  TAB 3: LCCA Analysis (กระแสเงินสด, NPV, EAC, Sensitivity)
  TAB 4: Word Report (รูปแบบที่ปรึกษา)
================================================================================
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from dataclasses import dataclass, field
from typing import List, Dict, Tuple
import json
import io
from datetime import datetime

# ── Optional imports ──────────────────────────────────────────────────────────
try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

try:
    from docx import Document as WordDocument
    from docx.shared import Inches, Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="LCCA Pavement Integrated v1.0",
    page_icon="🛣️",
    layout="wide",
)

# =============================================================================
# SECTION A: LOOKUP TABLES (Routine Cost)
# =============================================================================

X1_MAP = {
    "High Type (AC/PM บนหินคลุก)": 0.00,
    "Intermediate Type (AC/PM บน Stabilized)": 0.50,
    "Low Type (ST บน Soil-Aggregate)": 1.00,
}
X2_BREAKS = [(0,2,1.00),(2.01,3,0.75),(3.01,4,0.50),(4.01,5,0.25),(5.01,999,0.00)]
X3_OPTIONS = {
    "0 – 500       (X3=0.00)": 0.00,
    "501 – 600     (X3=0.04)": 0.04,
    "601 – 700     (X3=0.08)": 0.08,
    "701 – 800     (X3=0.12)": 0.12,
    "801 – 900     (X3=0.16)": 0.16,
    "901 – 1,000   (X3=0.20)": 0.20,
    "1,001 – 1,100 (X3=0.24)": 0.24,
    "1,101 – 1,200 (X3=0.29)": 0.29,
    "1,201 – 1,300 (X3=0.33)": 0.33,
    "1,301 – 1,400 (X3=0.37)": 0.37,
    "1,401 – 1,500 (X3=0.41)": 0.41,
    "1,501 – 1,600 (X3=0.45)": 0.45,
    "1,601 – 1,700 (X3=0.49)": 0.49,
    "1,701 – 1,800 (X3=0.53)": 0.53,
    "1,801 – 1,900 (X3=0.57)": 0.57,
    "1,901 – 2,000 (X3=0.61)": 0.61,
    "2,001 – 2,200 (X3=0.69)": 0.69,
    "2,201 – 2,400 (X3=0.78)": 0.78,
    "2,401 – 2,600 (X3=0.86)": 0.86,
    "2,601 – 2,800 (X3=0.94)": 0.94,
    "2,801 – 3,000 (X3=1.02)": 1.02,
    "3,001 – 3,300 (X3=1.14)": 1.14,
    "3,301 – 3,600 (X3=1.27)": 1.27,
    "3,601 – 3,900 (X3=1.37)": 1.37,
    "3,901 – 4,200 (X3=1.51)": 1.51,
    "4,201 – 4,500 (X3=1.64)": 1.64,
    "4,501 – 4,800 (X3=1.76)": 1.76,
    "4,801 – 5,100 (X3=1.88)": 1.88,
    "5,101 – 5,400 (X3=2.00)": 2.00,
    "5,401 – 5,700 (X3=2.13)": 2.13,
    "5,701+         (X3=2.25)": 2.25,
}
X4_BREAKS = [(0,3,0.00),(4,4,0.20),(5,5,0.40),(6,6,0.60),(7,7,0.80),
             (8,8,1.00),(9,9,1.20),(10,10,1.40),(11,11,1.60),(12,99999,1.80)]
X5_BREAKS = [(0,5.49,0.00),(5.50,5.99,0.02),(6.00,6.49,0.05),(6.50,6.99,0.10),(7.00,9999,0.19)]

TERRAIN_MAP  = {"ที่ราบ (0-3%)":"P","ลูกเนิน (3-5%)":"R","ลูกเนินสลับเขา (5-7%)":"RM","เขา (>7%)":"S"}
TERRAIN_KEYS = list(TERRAIN_MAP.keys())
X6_MAP  = {"P":0.00,"R":0.02,"RM":0.04,"S":0.07}
Y3_MAP  = {"P":0.00,"R":0.24,"RM":0.36,"S":0.48}
Y4_MAP  = {"P":0.00,"R":0.24,"RM":0.36,"S":0.48}
Y6_MAP  = {"P":0.00,"R":0.04,"RM":0.08,"S":0.12}
Y1_BREAKS = [(0,40,0.00),(40.01,60,0.10),(60.01,80,0.20),(80.01,9999,0.30)]
Y2_BREAKS = [(0,1.75,0.00),(1.76,2.00,0.10),(2.01,2.25,0.15),(2.26,9999,0.20)]
Y5_BREAKS = [(0,20.99,0.00),(21,25,0.02),(25.01,30,0.04),(30.01,9999,0.06)]

Z1_MAP   = {1:0.00,2:0.25,3:0.50,4:0.75,5:1.00,6:1.30,7:1.60,8:2.00}
Z2_BREAKS = [(0,2,1.00),(2.01,3,0.75),(3.01,4,0.50),(4.01,5,0.25),(5.01,999,0.00)]
Z3_OPTIONS = {
    "0 – 1,000        (Z3=0.00)": 0.00,
    "1,001 – 2,000    (Z3=0.20)": 0.20,
    "2,001 – 3,000    (Z3=0.30)": 0.30,
    "3,001 – 4,000    (Z3=0.50)": 0.50,
    "4,001 – 5,000    (Z3=0.75)": 0.75,
    "5,001 – 6,000    (Z3=1.00)": 1.00,
    "6,001 – 7,000    (Z3=1.25)": 1.25,
    "7,001 – 8,000    (Z3=1.50)": 1.50,
    "8,001 – 9,000    (Z3=1.75)": 1.75,
    "9,001 – 10,000   (Z3=2.00)": 2.00,
    "10,001 – 15,000  (Z3=2.50)": 2.50,
    "15,001+           (Z3=3.00)": 3.00,
}
Z4_BREAKS = [(0,6.49,0.00),(6.50,6.99,0.08),(7.00,9999,0.17)]

# =============================================================================
# SECTION B: CALCULATION FUNCTIONS (Routine Cost)
# =============================================================================

def lookup_range(value, breaks):
    for lo, hi, v in breaks:
        if lo <= value <= hi:
            return v
    return breaks[-1][2]

def calc_Ka_single(x1, x2_cbr, x3, x4_age, x5_width, x6_terrain,
                   y1_row, y2_shoulder, y3_terrain, y4_terrain, y5_bridge, y6_terrain):
    """คำนวณ Ka สำหรับ 1 ปี (ใช้ X4 จากอายุจริง)"""
    X1 = x1
    X2 = lookup_range(x2_cbr, X2_BREAKS)
    X3 = x3
    X4 = lookup_range(x4_age, X4_BREAKS)
    X5 = lookup_range(x5_width, X5_BREAKS)
    X6 = X6_MAP[x6_terrain]
    Y1 = lookup_range(y1_row, Y1_BREAKS)
    Y2 = lookup_range(y2_shoulder, Y2_BREAKS)
    Y3 = Y3_MAP[y3_terrain]
    Y4 = Y4_MAP[y4_terrain]
    Y5 = lookup_range(y5_bridge, Y5_BREAKS)
    Y6 = Y6_MAP[y6_terrain]
    Ka = 1 + 0.50*(X1+X2+X3+X4+X5+X6+Y1+Y2+Y3+Y4+Y5+Y6)
    factors = {"X1":X1,"X2":X2,"X3":X3,"X4":X4,"X5":X5,"X6":X6,
               "Y1":Y1,"Y2":Y2,"Y3":Y3,"Y4":Y4,"Y5":Y5,"Y6":Y6}
    return Ka, factors

def calc_Ka_average(x1, x2_cbr, x3, x4_start_age, x5_width, x6_terrain,
                    y1_row, y2_shoulder, y3_terrain, y4_terrain, y5_bridge, y6_terrain,
                    analysis_years):
    """คำนวณ Ka เฉลี่ยตลอด analysis_years โดย X4 เปลี่ยนตามอายุจริงแต่ละปี"""
    ka_list = []
    detail_rows = []
    for yr in range(1, analysis_years + 1):
        age = x4_start_age + (yr - 1)
        Ka, fac = calc_Ka_single(x1, x2_cbr, x3, age, x5_width, x6_terrain,
                                 y1_row, y2_shoulder, y3_terrain, y4_terrain, y5_bridge, y6_terrain)
        ka_list.append(Ka)
        detail_rows.append({"ปี": yr, "อายุ (ปี)": age, "X4": fac["X4"], "Ka": round(Ka, 4)})
    Ka_avg = np.mean(ka_list)
    return Ka_avg, pd.DataFrame(detail_rows)

def calc_Kc(z1_idx, z2_cbr, z3, z4_width,
            y1_row, y2_shoulder, y3_terrain, y4_terrain, y5_bridge, y6_terrain):
    Z1 = Z1_MAP.get(z1_idx, 0)
    Z2 = lookup_range(z2_cbr, Z2_BREAKS)
    Z3 = z3
    Z4 = lookup_range(z4_width, Z4_BREAKS)
    Y1 = lookup_range(y1_row, Y1_BREAKS)
    Y2 = lookup_range(y2_shoulder, Y2_BREAKS)
    Y3 = Y3_MAP[y3_terrain]
    Y4 = Y4_MAP[y4_terrain]
    Y5 = lookup_range(y5_bridge, Y5_BREAKS)
    Y6 = Y6_MAP[y6_terrain]
    Kc = 1 + 0.50*(Z1+Z2+Z3+Z4+Y1+Y2+Y3+Y4+Y5+Y6)
    factors = {"Z1":Z1,"Z2":Z2,"Z3":Z3,"Z4":Z4,
               "Y1":Y1,"Y2":Y2,"Y3":Y3,"Y4":Y4,"Y5":Y5,"Y6":Y6}
    return Kc, factors

# =============================================================================
# SECTION C: LCCA DATA STRUCTURES & FUNCTIONS
# =============================================================================

@dataclass
class MaintenanceActivity:
    name: str
    unit_cost: float   # บาท/ตร.ม.
    start_year: int
    frequency: int = 0  # 0 = one-time

@dataclass
class RehabActivity:
    name: str
    unit_cost: float
    year: int

@dataclass
class PavementAlternative:
    name: str
    pave_type: str
    construction_cost: float   # บาท/ตร.ม.
    area: float                # ตร.ม.
    maintenance: List[MaintenanceActivity]
    rehab: List[RehabActivity]
    salvage_pct: float = 20.0
    enabled: bool = True

def calc_pv(cost, year, discount_rate):
    if year < 0 or discount_rate < 0:
        return 0.0
    return cost * (1 + discount_rate) ** (-year)

def calc_eac(pw, discount_rate, n):
    if n <= 0 or discount_rate <= 0:
        return 0.0
    crf = discount_rate * (1 + discount_rate)**n / ((1 + discount_rate)**n - 1)
    return pw * crf

def build_cashflow(alt: PavementAlternative, n: int, dr: float, inc_salvage: bool) -> pd.DataFrame:
    rows = []
    area = alt.area
    rehab_years = sorted([r.year for r in alt.rehab if r.year <= n])
    rehab_set   = set(rehab_years)

    # ปีที่ 0: ก่อสร้าง
    rows.append({"ปี":0,"กิจกรรม":"ก่อสร้างเริ่มต้น","ประเภท":"ก่อสร้าง",
                 "ต้นทุน/หน่วย":alt.construction_cost,
                 "ต้นทุนตามปี":alt.construction_cost*area,
                 "PW_factor":1.0,
                 "มูลค่าปัจจุบัน":alt.construction_cost*area})

    # บำรุงรักษา (รีเซ็ตรอบหลังฟื้นฟู)
    for m in alt.maintenance:
        if m.frequency > 0:
            checkpoints = [0] + rehab_years
            for idx, cp in enumerate(checkpoints):
                end = checkpoints[idx+1] if idx+1 < len(checkpoints) else n+1
                yr = cp + m.frequency
                while yr < end and yr <= n:
                    if yr not in rehab_set:
                        cost = m.unit_cost * area
                        pwf  = (1+dr)**(-yr)
                        rows.append({"ปี":yr,"กิจกรรม":m.name,"ประเภท":"บำรุงรักษา",
                                     "ต้นทุน/หน่วย":m.unit_cost,"ต้นทุนตามปี":cost,
                                     "PW_factor":pwf,"มูลค่าปัจจุบัน":cost*pwf})
                    yr += m.frequency
        else:
            if m.start_year <= n and m.start_year not in rehab_set:
                cost = m.unit_cost * area
                pwf  = (1+dr)**(-m.start_year)
                rows.append({"ปี":m.start_year,"กิจกรรม":m.name,"ประเภท":"บำรุงรักษา",
                             "ต้นทุน/หน่วย":m.unit_cost,"ต้นทุนตามปี":cost,
                             "PW_factor":pwf,"มูลค่าปัจจุบัน":cost*pwf})

    # ฟื้นฟูสภาพ
    last_rehab_cost = alt.construction_cost * area
    last_rehab_year = 0
    for r in alt.rehab:
        if r.year <= n:
            cost = r.unit_cost * area
            pwf  = (1+dr)**(-r.year)
            rows.append({"ปี":r.year,"กิจกรรม":r.name,"ประเภท":"ฟื้นฟูสภาพ",
                         "ต้นทุน/หน่วย":r.unit_cost,"ต้นทุนตามปี":cost,
                         "PW_factor":pwf,"มูลค่าปัจจุบัน":cost*pwf})
            last_rehab_cost = cost
            last_rehab_year = r.year

    # มูลค่าซาก
    if inc_salvage:
        life_map = {"Flexible":15,"AC":15,"JPCP":20,"JRCP":20,"CRCP":25}
        exp_life = next((v for k,v in life_map.items() if k in alt.pave_type), 20)
        remaining = exp_life - (n - last_rehab_year)
        dep_per_yr = last_rehab_cost * (1 - alt.salvage_pct/100) / exp_life
        if remaining > 0:
            sv = last_rehab_cost - dep_per_yr*(n - last_rehab_year)
        else:
            sv = last_rehab_cost * alt.salvage_pct/100
        sv = max(sv, last_rehab_cost * alt.salvage_pct/100)
        pwf = (1+dr)**(-n)
        rows.append({"ปี":n,"กิจกรรม":"มูลค่าซาก","ประเภท":"มูลค่าซาก",
                     "ต้นทุน/หน่วย":-sv/area,"ต้นทุนตามปี":-sv,
                     "PW_factor":pwf,"มูลค่าปัจจุบัน":-sv*pwf})

    df = pd.DataFrame(rows).sort_values(["ปี","กิจกรรม"]).reset_index(drop=True)
    return df

def analyze_lcca(alternatives, n, dr, inc_salvage):
    summary_rows = []
    cf_dict = {}
    for alt in alternatives:
        if not alt.enabled:
            continue
        cf = build_cashflow(alt, n, dr, inc_salvage)
        cf_dict[alt.name] = cf
        pw   = cf["มูลค่าปัจจุบัน"].sum()
        eac  = calc_eac(pw, dr, n)
        pw_c = cf[cf["ประเภท"]=="ก่อสร้าง"]["มูลค่าปัจจุบัน"].sum()
        pw_m = cf[cf["ประเภท"]=="บำรุงรักษา"]["มูลค่าปัจจุบัน"].sum()
        pw_r = cf[cf["ประเภท"]=="ฟื้นฟูสภาพ"]["มูลค่าปัจจุบัน"].sum()
        pw_s = cf[cf["ประเภท"]=="มูลค่าซาก"]["มูลค่าปัจจุบัน"].sum()
        summary_rows.append({
            "ทางเลือก": alt.name,
            "ประเภทผิวทาง": alt.pave_type,
            "พื้นที่ (ตร.ม.)": alt.area,
            "ต้นทุนก่อสร้าง (บาท/ตร.ม.)": alt.construction_cost,
            "PW_ก่อสร้าง": pw_c,
            "PW_บำรุงรักษา": pw_m,
            "PW_ฟื้นฟูสภาพ": pw_r,
            "PW_มูลค่าซาก": pw_s,
            "มูลค่าปัจจุบันรวม (บาท)": pw,
            "EAC (บาท/ปี)": eac,
            "EAC (บาท/ตร.ม./ปี)": eac / alt.area if alt.area > 0 else 0,
        })
    df = pd.DataFrame(summary_rows)
    if len(df) > 0:
        df = df.sort_values("มูลค่าปัจจุบันรวม (บาท)").reset_index(drop=True)
        df.insert(0, "อันดับ", range(1, len(df)+1))
    return df, cf_dict

# =============================================================================
# SECTION D: SESSION STATE INIT
# =============================================================================

def init_state():
    defaults = {
        # TAB1
        "project_name": "โครงการก่อสร้างทางหลวง",
        "project_road":  "",
        "project_km":    "",
        "project_year":  str(datetime.now().year + 543),
        "cost_ac":   0.0,
        "cost_jpcp": 0.0,
        "cost_jrcp": 0.0,
        "cost_crcp": 0.0,
        "area_sqm":  10000.0,
        # TAB2 shared Y-factors
        "y_row":       40.0,
        "y_shoulder":  1.75,
        "y_terrain":   TERRAIN_KEYS[0],
        "y_bridge":    0.0,
        # TAB2 AC section
        "ac_x1_key":   list(X1_MAP.keys())[0],
        "ac_x2_cbr":   3.0,
        "ac_x3_key":   list(X3_OPTIONS.keys())[0],
        "ac_x4_age":   0,
        "ac_x5_width": 7.0,
        "ac_x6_terrain": TERRAIN_KEYS[0],
        "ac_Na":  35000.0,
        "ac_Km":  1.0,
        # TAB2 Concrete section
        "cc_z1_idx":   1,
        "cc_z2_cbr":   3.0,
        "cc_z3_key":   list(Z3_OPTIONS.keys())[0],
        "cc_z4_width": 7.0,
        "cc_z6_terrain": TERRAIN_KEYS[0],
        "cc_Nc":  35000.0,
        "cc_Km":  1.0,
        # Results from TAB2 → TAB3
        "ka_avg":  None,
        "kc_val":  None,
        "routine_ac_per_sqm":  None,
        "routine_cc_per_sqm":  None,
        # TAB3 LCCA settings
        "lcca_n":  20,
        "lcca_dr": 0.06,
        "lcca_salvage": True,
        "lcca_alternatives": None,  # List[PavementAlternative]
        # misc
        "show_gravel": False,
        "json_version": 0,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

# =============================================================================
# SECTION E: WORD REPORT HELPERS
# =============================================================================

def set_run_font(run, font="TH SarabunPSK", size=None, bold=False, italic=False):
    if size is None:
        size = Pt(14)
    run.font.name  = font
    run.font.size  = size
    run.font.bold  = bold
    run.font.italic = italic
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), font)
    rFonts.set(qn("w:hAnsi"), font)
    rFonts.set(qn("w:cs"),    font)

def add_thai_para(doc, text="", bold=False, first_indent=True, size=None):
    if size is None:
        size = Pt(14)
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    jc = OxmlElement("w:jc"); jc.set(qn("w:val"), "thaiDistribute"); pPr.append(jc)
    if first_indent:
        ind = OxmlElement("w:ind"); ind.set(qn("w:firstLine"), "720"); pPr.append(ind)
    if text:
        run = p.add_run(text)
        set_run_font(run, size=size, bold=bold)
    return p

def add_heading_word(doc, text, level=1):
    p = doc.add_heading(text, level=level)
    for run in p.runs:
        set_run_font(run, size=Pt(15 if level==1 else 14), bold=True)
    return p

def add_table_word(doc, headers, rows, col_widths=None):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    if col_widths:
        for i, w in enumerate(col_widths):
            table.columns[i].width = Cm(w)
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.paragraphs[0].clear()
        run = cell.paragraphs[0].add_run(str(h))
        set_run_font(run, size=Pt(14), bold=True)
    for row_data in rows:
        row = table.add_row()
        for i, val in enumerate(row_data):
            cell = row.cells[i]
            cell.paragraphs[0].clear()
            run = cell.paragraphs[0].add_run(str(val))
            set_run_font(run, size=Pt(14))
    return table

def generate_word_report(summary_df, cf_dict, n, dr, alternatives, base_sec="4"):
    doc = WordDocument()
    normal = doc.styles["Normal"]
    normal.font.name = "TH SarabunPSK"
    normal.font.size = Pt(14)
    for sec in doc.sections:
        sec.top_margin = Cm(2.5); sec.bottom_margin = Cm(2.5)
        sec.left_margin = Cm(3.0); sec.right_margin  = Cm(2.5)

    ss = st.session_state

    class SC:
        major = int(base_sec.split(".")[0]) if "." not in base_sec else int(base_sec.split(".")[0])
        minor_base = int(base_sec.split(".")[1]) if "." in base_sec else 0
        h1 = 0; h2 = 0

    def next_h1(title):
        SC.h1 += 1; SC.h2 = 0
        num = f"{SC.major}.{SC.minor_base + SC.h1}"
        add_heading_word(doc, f"{num}  {title}", level=1)

    def next_h2(title):
        SC.h2 += 1
        num = f"{SC.major}.{SC.minor_base + SC.h1}.{SC.h2}"
        add_heading_word(doc, f"{num}  {title}", level=2)

    # ── ปก ──────────────────────────────────────────────────────────────────
    tp = doc.add_paragraph(); tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = tp.add_run("รายการคำนวณการวิเคราะห์ต้นทุนตลอดอายุการใช้งานผิวทาง")
    set_run_font(run, size=Pt(18), bold=True)
    sp = doc.add_paragraph(); sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = sp.add_run("Life-Cycle Cost Analysis for Pavement Alternatives")
    set_run_font(run2, size=Pt(16), bold=True)
    doc.add_paragraph()
    pp = doc.add_paragraph(); pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run3 = pp.add_run(ss["project_name"])
    set_run_font(run3, size=Pt(15), bold=True)
    doc.add_paragraph()
    add_thai_para(doc, f"วันที่จัดทำ: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                  first_indent=False)
    doc.add_paragraph()

    # ── หัวข้อ 1: ข้อมูลโครงการ ─────────────────────────────────────────────
    next_h1("ข้อมูลโครงการ")
    add_table_word(doc,
        headers=["รายการ", "ข้อมูล"],
        rows=[
            ["ชื่อโครงการ",        ss["project_name"]],
            ["สายทาง / ตอนควบคุม", ss["project_road"]],
            ["กม. ที่",             ss["project_km"]],
            ["ปีงบประมาณ",         ss["project_year"]],
            ["พื้นที่วิเคราะห์",   f"{ss['area_sqm']:,.0f} ตร.ม."],
            ["ระยะวิเคราะห์",      f"{n} ปี"],
            ["อัตราคิดลด",         f"{dr*100:.1f}%"],
        ],
        col_widths=[5, 11],
    )
    doc.add_paragraph()

    # ── หัวข้อ 2: ต้นทุนก่อสร้าง ────────────────────────────────────────────
    next_h1("ต้นทุนก่อสร้างผิวทาง")
    add_table_word(doc,
        headers=["ผิวทาง", "ประเภท", "ต้นทุนก่อสร้าง (บาท/ตร.ม.)"],
        rows=[
            ["AC",   "ลาดยาง",    f"{ss['cost_ac']:,.2f}"],
            ["JPCP", "คอนกรีต",   f"{ss['cost_jpcp']:,.2f}"],
            ["JRCP", "คอนกรีต",   f"{ss['cost_jrcp']:,.2f}"],
            ["CRCP", "คอนกรีต",   f"{ss['cost_crcp']:,.2f}"],
        ],
        col_widths=[3, 4, 6],
    )
    doc.add_paragraph()

    # ── หัวข้อ 3: ค่าบำรุงรักษา ─────────────────────────────────────────────
    next_h1("ค่าบำรุงรักษาประจำปี (Routine Maintenance Cost)")
    add_thai_para(doc,
        "ค่าบำรุงรักษาประจำปีคำนวณตามคู่มือกรมทางหลวง พ.ศ. 2538 "
        "โดยใช้ค่าสัมประสิทธิ์ปรับแก้ Ka (ผิวแอสฟัลท์) และ Kc (ผิวคอนกรีต) "
        "สำหรับ Ka ได้คำนวณรายปีตลอดช่วงวิเคราะห์ (X4 เปลี่ยนตามอายุจริง) "
        "แล้วหาค่าเฉลี่ยเพื่อใช้ในการวิเคราะห์ LCCA")

    ka_avg = ss.get("ka_avg")
    kc_val = ss.get("kc_val")
    r_ac   = ss.get("routine_ac_per_sqm")
    r_cc   = ss.get("routine_cc_per_sqm")

    rows_maint = []
    if ka_avg is not None:
        rows_maint.append(["AC (ลาดยาง)", f"Ka = {ka_avg:.4f}",
                           f"Na = {ss['ac_Na']:,.0f} บาท/กม./ปี",
                           f"{r_ac:,.2f}" if r_ac else "-"])
    if kc_val is not None:
        rows_maint.append(["Concrete", f"Kc = {kc_val:.4f}",
                           f"Nc = {ss['cc_Nc']:,.0f} บาท/กม./ปี",
                           f"{r_cc:,.2f}" if r_cc else "-"])

    if rows_maint:
        add_table_word(doc,
            headers=["ผิวทาง", "K เฉลี่ย", "อัตรามาตรฐาน N", "ค่าบำรุงรักษา (บาท/ตร.ม./ปี)"],
            rows=rows_maint,
            col_widths=[3, 3.5, 5, 5.5],
        )
    doc.add_paragraph()

    # ── หัวข้อ 4: สรุปผล LCCA ───────────────────────────────────────────────
    next_h1("สรุปผลการวิเคราะห์ LCCA")
    add_thai_para(doc,
        f"การวิเคราะห์ต้นทุนตลอดอายุการใช้งาน (Life-Cycle Cost Analysis: LCCA) "
        f"ดำเนินการสำหรับระยะเวลาวิเคราะห์ {n} ปี ที่อัตราคิดลด {dr*100:.1f}% ต่อปี "
        f"ผลการวิเคราะห์แสดงในตารางด้านล่าง")

    if len(summary_df) > 0:
        tbl_rows = []
        for _, row in summary_df.iterrows():
            tbl_rows.append([
                str(int(row["อันดับ"])),
                row["ทางเลือก"],
                row["ประเภทผิวทาง"],
                f"{row['ต้นทุนก่อสร้าง (บาท/ตร.ม.)']:,.0f}",
                f"{row['มูลค่าปัจจุบันรวม (บาท)']:,.0f}",
                f"{row['EAC (บาท/ปี)']:,.0f}",
                f"{row['EAC (บาท/ตร.ม./ปี)']:,.2f}",
            ])
        add_table_word(doc,
            headers=["อันดับ","ทางเลือก","ประเภท","ต้นทุนก่อสร้าง\n(บาท/ตร.ม.)",
                     "มูลค่าปัจจุบันรวม\n(บาท)","EAC\n(บาท/ปี)","EAC\n(บาท/ตร.ม./ปี)"],
            rows=tbl_rows,
            col_widths=[1.5, 3, 2.5, 3, 4, 3.5, 3.5],
        )
        doc.add_paragraph()

        # สรุปคำแนะนำ
        best = summary_df.iloc[0]
        add_thai_para(doc,
            f"จากการวิเคราะห์พบว่า {best['ทางเลือก']} ({best['ประเภทผิวทาง']}) "
            f"มีมูลค่าปัจจุบันของต้นทุนตลอดอายุการใช้งานต่ำที่สุด "
            f"เท่ากับ {best['มูลค่าปัจจุบันรวม (บาท)']:,.0f} บาท "
            f"คิดเป็นต้นทุนเฉลี่ยรายปี {best['EAC (บาท/ปี)']:,.0f} บาทต่อปี "
            f"({best['EAC (บาท/ตร.ม./ปี)']:,.2f} บาทต่อตารางเมตรต่อปี) "
            f"จึงเป็นทางเลือกที่ประหยัดที่สุดในเชิงต้นทุนตลอดอายุการใช้งาน")

    # ── หัวข้อ 5: กระแสเงินสดรายทางเลือก ───────────────────────────────────
    next_h1("กระแสเงินสดรายทางเลือก")
    for alt_name, cf in cf_dict.items():
        next_h2(alt_name)
        pw_total = cf["มูลค่าปัจจุบัน"].sum()
        eac_val  = calc_eac(pw_total, dr, n)
        add_thai_para(doc,
            f"มูลค่าปัจจุบันรวม = {pw_total:,.0f} บาท  |  "
            f"EAC = {eac_val:,.0f} บาท/ปี",
            first_indent=False)
        cf_rows = []
        for _, r in cf.iterrows():
            cf_rows.append([
                str(int(r["ปี"])), r["กิจกรรม"], r["ประเภท"],
                f"{r['ต้นทุน/หน่วย']:,.2f}",
                f"{r['ต้นทุนตามปี']:,.0f}",
                f"{r['PW_factor']:.4f}",
                f"{r['มูลค่าปัจจุบัน']:,.0f}",
            ])
        add_table_word(doc,
            headers=["ปี","กิจกรรม","ประเภท","ต้นทุน/หน่วย","ต้นทุนตามปี","PW Factor","มูลค่าปัจจุบัน"],
            rows=cf_rows,
            col_widths=[1.2, 4.5, 2.5, 2.5, 2.8, 2.2, 2.8],
        )
        doc.add_paragraph()

    # Footer
    doc.add_paragraph()
    fp = doc.add_paragraph()
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fr = fp.add_run("รศ.ดร.อิทธิพล มีผล | ภาควิชาครุศาสตร์โยธา | KMUTNB")
    set_run_font(fr, size=Pt(12))

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# =============================================================================
# SECTION F: JSON HELPERS
# =============================================================================

def build_json():
    ss = st.session_state
    alts = ss.get("lcca_alternatives") or []
    alts_data = []
    for a in alts:
        alts_data.append({
            "name": a.name, "pave_type": a.pave_type,
            "construction_cost": a.construction_cost, "area": a.area,
            "salvage_pct": a.salvage_pct, "enabled": a.enabled,
            "maintenance": [{"name":m.name,"unit_cost":m.unit_cost,
                             "start_year":m.start_year,"frequency":m.frequency}
                            for m in a.maintenance],
            "rehab": [{"name":r.name,"unit_cost":r.unit_cost,"year":r.year}
                      for r in a.rehab],
        })
    return {
        "app": "LCCA_Integrated_v1",
        "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "project_name": ss["project_name"],
        "project_road": ss["project_road"],
        "project_km":   ss["project_km"],
        "project_year": ss["project_year"],
        "area_sqm":     ss["area_sqm"],
        "cost_ac": ss["cost_ac"], "cost_jpcp": ss["cost_jpcp"],
        "cost_jrcp": ss["cost_jrcp"], "cost_crcp": ss["cost_crcp"],
        "lcca_n": ss["lcca_n"], "lcca_dr": ss["lcca_dr"],
        "lcca_salvage": ss["lcca_salvage"],
        "ka_avg": ss.get("ka_avg"), "kc_val": ss.get("kc_val"),
        "routine_ac_per_sqm": ss.get("routine_ac_per_sqm"),
        "routine_cc_per_sqm": ss.get("routine_cc_per_sqm"),
        "alternatives": alts_data,
    }

def load_json(data: dict):
    ss = st.session_state
    for key in ["project_name","project_road","project_km","project_year",
                "area_sqm","cost_ac","cost_jpcp","cost_jrcp","cost_crcp",
                "lcca_n","lcca_dr","lcca_salvage","ka_avg","kc_val",
                "routine_ac_per_sqm","routine_cc_per_sqm"]:
        if key in data:
            ss[key] = data[key]
    alts = []
    for a in data.get("alternatives", []):
        alts.append(PavementAlternative(
            name=a["name"], pave_type=a["pave_type"],
            construction_cost=a["construction_cost"], area=a["area"],
            salvage_pct=a.get("salvage_pct",20.0), enabled=a.get("enabled",True),
            maintenance=[MaintenanceActivity(m["name"],m["unit_cost"],
                                            m["start_year"],m["frequency"])
                         for m in a.get("maintenance",[])],
            rehab=[RehabActivity(r["name"],r["unit_cost"],r["year"])
                   for r in a.get("rehab",[])],
        ))
    if alts:
        ss["lcca_alternatives"] = alts

# =============================================================================
# SECTION G: MAIN UI
# =============================================================================

st.title("🛣️ LCCA Pavement Integrated v1.0")
st.caption("Life-Cycle Cost Analysis for Pavement Alternatives | รศ.ดร.อิทธิพล มีผล | KMUTNB")

tab1, tab2, tab3, tab4 = st.tabs([
    "📋 TAB 1: ข้อมูลโครงการ & ราคาก่อสร้าง",
    "🔧 TAB 2: Routine Cost (Ka / Kc)",
    "📊 TAB 3: LCCA Analysis",
    "📄 TAB 4: Word Report",
])

# ─────────────────────────────────────────────────────────────────────────────
# TAB 1: ข้อมูลโครงการ + ราคาก่อสร้าง
# ─────────────────────────────────────────────────────────────────────────────
with tab1:
    st.header("📋 ข้อมูลโครงการและราคาก่อสร้าง")

    # JSON I/O
    col_j1, col_j2 = st.columns([1, 1])
    with col_j1:
        uploaded_json = st.file_uploader("📂 โหลดโครงการ (JSON)", type="json", key="json_load_t1")
        if uploaded_json:
            try:
                data = json.load(uploaded_json)
                load_json(data)
                st.success("✅ โหลดข้อมูลสำเร็จ")
                st.rerun()
            except Exception as e:
                st.error(f"โหลดไม่ได้: {e}")
    with col_j2:
        st.download_button("💾 บันทึก JSON",
            data=json.dumps(build_json(), ensure_ascii=False, indent=2).encode("utf-8"),
            file_name=f"LCCA_{st.session_state['project_name'].replace(' ','_')}.json",
            mime="application/json", key="json_dl_t1")

    st.divider()

    # ข้อมูลโครงการ
    st.subheader("🏗️ ข้อมูลโครงการ")
    c1, c2 = st.columns(2)
    with c1:
        st.session_state["project_name"] = st.text_input("ชื่อโครงการ", value=st.session_state["project_name"], key="pn_t1")
        st.session_state["project_road"] = st.text_input("สายทาง / ตอนควบคุม", value=st.session_state["project_road"], key="pr_t1")
    with c2:
        st.session_state["project_km"]   = st.text_input("กม. ที่ (เช่น กม.100+000 – กม.110+000)", value=st.session_state["project_km"], key="pk_t1")
        st.session_state["project_year"] = st.text_input("ปีงบประมาณ (พ.ศ.)", value=st.session_state["project_year"], key="py_t1")

    st.session_state["area_sqm"] = st.number_input(
        "พื้นที่วิเคราะห์ (ตร.ม.) — ใช้สำหรับคำนวณ LCCA ทุกทางเลือก",
        min_value=100.0, value=float(st.session_state["area_sqm"]),
        step=100.0, key="area_t1")

    st.divider()

    # ราคาก่อสร้าง
    st.subheader("💰 ราคาโครงสร้างชั้นทาง (บาท/ตร.ม.)")
    st.info("กรอกราคาก่อสร้างสุทธิต่อตารางเมตร หรือ Upload Excel template")

    # Upload Excel
    excel_file = st.file_uploader("📤 Upload Excel ราคาก่อสร้าง", type=["xlsx","xls"], key="excel_cost_t1")
    if excel_file and OPENPYXL_AVAILABLE:
        try:
            df_cost = pd.read_excel(excel_file, header=2)   # row 3 = header (0-indexed row 2)
            cost_map = {"AC": "cost_ac", "JPCP": "cost_jpcp", "JRCP": "cost_jrcp", "CRCP": "cost_crcp"}
            for _, row in df_cost.iterrows():
                key_str = str(row.iloc[0]).strip().upper()
                if key_str in cost_map:
                    val = row.iloc[2]  # column index 2 = ต้นทุนก่อสร้าง
                    if pd.notna(val):
                        st.session_state[cost_map[key_str]] = float(val)
            st.success("✅ โหลดราคาจาก Excel สำเร็จ")
        except Exception as e:
            st.error(f"อ่าน Excel ไม่ได้: {e}")

    # Template download
    if OPENPYXL_AVAILABLE:
        buf_tmpl = io.BytesIO()
        wb = openpyxl.Workbook()
        ws = wb.active
        from openpyxl.styles import Font, PatternFill, Alignment
        title_font = Font(name="TH SarabunPSK", bold=True, size=14)
        hdr_font   = Font(name="TH SarabunPSK", bold=True, size=13)
        body_font  = Font(name="TH SarabunPSK", size=13)
        blue_fill  = PatternFill("solid", fgColor="1F4E79")
        hdr_fill   = PatternFill("solid", fgColor="2E75B6")
        center     = Alignment(horizontal="center", vertical="center", wrap_text=True)

        ws.merge_cells("A1:C1")
        ws["A1"] = "ข้อมูลสำหรับวิเคราะห์ LCCA (Life-Cycle Cost Analysis)"
        ws["A1"].font = Font(name="TH SarabunPSK", bold=True, size=14, color="FFFFFF")
        ws["A1"].fill = blue_fill; ws["A1"].alignment = center
        ws["A2"] = "💡 คำแนะนำ: กรอกข้อมูลในแถวที่ 4-7 → บันทึกไฟล์ → อัปโหลดในโปรแกรม"
        ws["A2"].font = Font(name="TH SarabunPSK", size=12, italic=True)

        headers = ["ผิวทาง", "ประเภทผิวทาง", "ต้นทุนก่อสร้าง (บาท/ตร.ม.)"]
        for ci, h in enumerate(headers, 1):
            cell = ws.cell(row=3, column=ci, value=h)
            cell.font = hdr_font; cell.fill = hdr_fill
            cell.alignment = center
            cell.font = Font(name="TH SarabunPSK", bold=True, size=13, color="FFFFFF")

        template_data = [("AC","ลาดยาง",""),("JPCP","คอนกรีต",""),
                         ("JRCP","คอนกรีต",""),("CRCP","คอนกรีต","")]
        for ri, (a, b, c) in enumerate(template_data, 4):
            ws.cell(row=ri, column=1, value=a).font = Font(name="TH SarabunPSK", bold=True, size=13)
            ws.cell(row=ri, column=2, value=b).font = body_font
            ws.cell(row=ri, column=3, value=c).font = body_font
            for ci in range(1, 4):
                ws.cell(row=ri, column=ci).alignment = center

        ws.column_dimensions["A"].width = 10
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 28
        ws.row_dimensions[1].height = 24; ws.row_dimensions[3].height = 36
        wb.save(buf_tmpl); buf_tmpl.seek(0)
        st.download_button("📥 ดาวน์โหลด Excel Template", data=buf_tmpl,
            file_name="LCCA_cost_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="tmpl_dl_t1")

    # Manual input
    st.markdown("**หรือกรอกราคาด้านล่างได้เลย:**")
    col_ac, col_jp, col_jr, col_cr = st.columns(4)
    with col_ac:
        st.session_state["cost_ac"]   = st.number_input("AC (ลาดยาง)",   min_value=0.0, value=float(st.session_state["cost_ac"]),   step=10.0, format="%.2f", key="cac_t1")
    with col_jp:
        st.session_state["cost_jpcp"] = st.number_input("JPCP (คอนกรีต)", min_value=0.0, value=float(st.session_state["cost_jpcp"]), step=10.0, format="%.2f", key="cjp_t1")
    with col_jr:
        st.session_state["cost_jrcp"] = st.number_input("JRCP (คอนกรีต)", min_value=0.0, value=float(st.session_state["cost_jrcp"]), step=10.0, format="%.2f", key="cjr_t1")
    with col_cr:
        st.session_state["cost_crcp"] = st.number_input("CRCP (คอนกรีต)", min_value=0.0, value=float(st.session_state["cost_crcp"]), step=10.0, format="%.2f", key="ccr_t1")

    # Preview
    if any([st.session_state["cost_ac"], st.session_state["cost_jpcp"],
            st.session_state["cost_jrcp"], st.session_state["cost_crcp"]]):
        st.success("✅ ข้อมูลราคาพร้อมส่งไป TAB 3 แล้ว")
        st.dataframe(pd.DataFrame({
            "ผิวทาง": ["AC","JPCP","JRCP","CRCP"],
            "ประเภท": ["ลาดยาง","คอนกรีต","คอนกรีต","คอนกรีต"],
            "ต้นทุนก่อสร้าง (บาท/ตร.ม.)": [
                st.session_state["cost_ac"], st.session_state["cost_jpcp"],
                st.session_state["cost_jrcp"], st.session_state["cost_crcp"]],
        }), hide_index=True, use_container_width=True)

# ─────────────────────────────────────────────────────────────────────────────
# TAB 2: Routine Cost (Ka / Kc)
# ─────────────────────────────────────────────────────────────────────────────
with tab2:
    st.header("🔧 Routine Cost Calculator")
    st.info("คำนวณ Ka (ผิว AC) และ Kc (ผิวคอนกรีต) เพื่อหาค่าบำรุงรักษาประจำปี (บาท/ตร.ม./ปี)")

    # ── ตัวแปรร่วม Y ────────────────────────────────────────────────────────
    st.subheader("📌 ตัวแปรร่วม Y1–Y6 (ใช้ร่วมกันทั้ง AC และ Concrete)")
    yc1, yc2, yc3, yc4 = st.columns(4)
    with yc1:
        st.session_state["y_row"]      = st.number_input("Y1: ความกว้างเขตทาง (ม.)", min_value=0.0, value=float(st.session_state["y_row"]), step=1.0, key="yr_t2")
    with yc2:
        st.session_state["y_shoulder"] = st.number_input("Y2: ไหล่ทางกว้างสุด 1 ข้าง (ม.)", min_value=0.0, value=float(st.session_state["y_shoulder"]), step=0.25, key="ys_t2")
    with yc3:
        st.session_state["y_terrain"]  = st.selectbox("Y3/Y4/Y6: ภูมิประเทศ", TERRAIN_KEYS, index=TERRAIN_KEYS.index(st.session_state["y_terrain"]), key="yt_t2")
    with yc4:
        st.session_state["y_bridge"]   = st.number_input("Y5: ความยาวสะพาน (ม./กม.)", min_value=0.0, value=float(st.session_state["y_bridge"]), step=1.0, key="yb_t2")

    y_terrain_code = TERRAIN_MAP[st.session_state["y_terrain"]]
    Y1 = lookup_range(st.session_state["y_row"],      Y1_BREAKS)
    Y2 = lookup_range(st.session_state["y_shoulder"], Y2_BREAKS)
    Y3 = Y3_MAP[y_terrain_code]; Y4 = Y4_MAP[y_terrain_code]
    Y5 = lookup_range(st.session_state["y_bridge"],   Y5_BREAKS)
    Y6 = Y6_MAP[y_terrain_code]

    with st.expander("ดูค่า Y factors"):
        st.dataframe(pd.DataFrame({
            "Factor":["Y1","Y2","Y3","Y4","Y5","Y6"],
            "คำอธิบาย":["เขตทาง","ไหล่ทาง","จราจรสงเคราะห์","ท่อระบายน้ำ","สะพาน","ทำความสะอาดระบาย"],
            "ค่า":[Y1,Y2,Y3,Y4,Y5,Y6],
        }), hide_index=True)

    st.divider()

    # ── Section A: AC ────────────────────────────────────────────────────────
    col_ac_sec, col_cc_sec = st.columns(2)

    with col_ac_sec:
        st.subheader("🔵 Section A: ผิวแอสฟัลท์ (Ka)")

        ac_x1_key = st.selectbox("X1: ลักษณะผิวทาง", list(X1_MAP.keys()),
            index=list(X1_MAP.keys()).index(st.session_state["ac_x1_key"]), key="ax1_t2")
        st.session_state["ac_x1_key"] = ac_x1_key

        ac_x2 = st.number_input("X2: CBR ดินเดิม (%)", min_value=0.0, max_value=20.0,
            value=float(st.session_state["ac_x2_cbr"]), step=0.5, key="ax2_t2")
        st.session_state["ac_x2_cbr"] = ac_x2

        ac_x3_key = st.selectbox("X3: AADT (คัน/วัน)", list(X3_OPTIONS.keys()),
            index=list(X3_OPTIONS.keys()).index(st.session_state["ac_x3_key"]), key="ax3_t2")
        st.session_state["ac_x3_key"] = ac_x3_key

        ac_x4 = st.number_input("X4: อายุปัจจุบันของผิวทาง (ปี) — ปีเริ่มต้น",
            min_value=0, value=int(st.session_state["ac_x4_age"]), step=1, key="ax4_t2")
        st.session_state["ac_x4_age"] = ac_x4

        ac_x5 = st.number_input("X5: ความกว้างผิวทาง (ม.)", min_value=4.0,
            value=float(st.session_state["ac_x5_width"]), step=0.5, key="ax5_t2")
        st.session_state["ac_x5_width"] = ac_x5

        ac_x6_key = st.selectbox("X6: ภูมิประเทศ (AC)", TERRAIN_KEYS,
            index=TERRAIN_KEYS.index(st.session_state["ac_x6_terrain"]), key="ax6_t2")
        st.session_state["ac_x6_terrain"] = ac_x6_key

        ac_Na = st.number_input("Na: อัตราค่าบำรุงมาตรฐาน (บาท/กม./ปี)",
            min_value=1000.0, value=float(st.session_state["ac_Na"]), step=500.0, key="ana_t2")
        st.session_state["ac_Na"] = ac_Na
        ac_Km = st.number_input("Km: Factor วัสดุ (AC)", min_value=0.1,
            value=float(st.session_state["ac_Km"]), step=0.05, format="%.3f", key="akm_t2")
        st.session_state["ac_Km"] = ac_Km

    with col_cc_sec:
        st.subheader("🟠 Section B: ผิวคอนกรีต (Kc)")

        cc_z1 = st.selectbox("Z1: ดัชนีสภาพผิวทาง (1=ดีมาก … 8=แย่มาก)", list(range(1,9)),
            index=st.session_state["cc_z1_idx"]-1, key="cz1_t2")
        st.session_state["cc_z1_idx"] = cc_z1

        cc_z2 = st.number_input("Z2: CBR ดินคันทาง (%)", min_value=0.0, max_value=20.0,
            value=float(st.session_state["cc_z2_cbr"]), step=0.5, key="cz2_t2")
        st.session_state["cc_z2_cbr"] = cc_z2

        cc_z3_key = st.selectbox("Z3: AADT (คัน/วัน)", list(Z3_OPTIONS.keys()),
            index=list(Z3_OPTIONS.keys()).index(st.session_state["cc_z3_key"]), key="cz3_t2")
        st.session_state["cc_z3_key"] = cc_z3_key

        cc_z4 = st.number_input("Z4: ความกว้างผิวทาง (ม.)", min_value=4.0,
            value=float(st.session_state["cc_z4_width"]), step=0.5, key="cz4_t2")
        st.session_state["cc_z4_width"] = cc_z4

        cc_Nc = st.number_input("Nc: อัตราค่าบำรุงมาตรฐาน (บาท/กม./ปี)",
            min_value=1000.0, value=float(st.session_state["cc_Nc"]), step=500.0, key="cnc_t2")
        st.session_state["cc_Nc"] = cc_Nc
        cc_Km = st.number_input("Km: Factor วัสดุ (Concrete)", min_value=0.1,
            value=float(st.session_state["cc_Km"]), step=0.05, format="%.3f", key="ckm_t2")
        st.session_state["cc_Km"] = cc_Km

    st.divider()

    # ── คำนวณ ────────────────────────────────────────────────────────────────
    n_analysis = st.session_state.get("lcca_n", 20)
    st.info(f"ℹ️ Ka จะคำนวณรายปีตลอด {n_analysis} ปี (ตาม Analysis Period ใน TAB 3) แล้วหาค่าเฉลี่ย")

    if st.button("🔄 คำนวณ Ka และ Kc", type="primary", key="calc_t2"):
        # Ka รายปี → เฉลี่ย
        X1_val  = X1_MAP[ac_x1_key]
        X3_val  = X3_OPTIONS[ac_x3_key]
        X6_code = TERRAIN_MAP[ac_x6_key]

        ka_avg, ka_detail_df = calc_Ka_average(
            X1_val, ac_x2, X3_val, ac_x4, ac_x5, X6_code,
            st.session_state["y_row"], st.session_state["y_shoulder"],
            y_terrain_code, y_terrain_code,
            st.session_state["y_bridge"], y_terrain_code,
            n_analysis
        )

        # Kc (คงที่)
        Z3_val = Z3_OPTIONS[cc_z3_key]
        Kc, kc_factors = calc_Kc(
            cc_z1, cc_z2, Z3_val, cc_z4,
            st.session_state["y_row"], st.session_state["y_shoulder"],
            y_terrain_code, y_terrain_code,
            st.session_state["y_bridge"], y_terrain_code
        )

        # ค่าบำรุงรักษา บาท/ตร.ม./ปี
        # สมมติถนนกว้าง ac_x5 ม. ยาว 1 กม. = ac_x5*1000 ตร.ม.
        road_area_per_km = ac_x5 * 1000  # ตร.ม./กม.
        routine_ac  = (ac_Na * ka_avg * ac_Km) / road_area_per_km
        routine_cc  = (cc_Nc * Kc    * cc_Km) / (cc_z4 * 1000)

        st.session_state["ka_avg"]             = round(ka_avg, 4)
        st.session_state["kc_val"]             = round(Kc, 4)
        st.session_state["routine_ac_per_sqm"] = round(routine_ac, 4)
        st.session_state["routine_cc_per_sqm"] = round(routine_cc, 4)
        st.session_state["_ka_detail_df"]      = ka_detail_df
        st.session_state["_kc_factors"]        = kc_factors

        st.success("✅ คำนวณสำเร็จ — ข้อมูลส่งไป TAB 3 อัตโนมัติแล้ว")

    # ── แสดงผล ───────────────────────────────────────────────────────────────
    if st.session_state.get("ka_avg") is not None:
        ka_avg_show = st.session_state["ka_avg"]
        kc_show     = st.session_state["kc_val"]
        r_ac_show   = st.session_state["routine_ac_per_sqm"]
        r_cc_show   = st.session_state["routine_cc_per_sqm"]

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Ka เฉลี่ย", f"{ka_avg_show:.4f}")
        m2.metric("Kc", f"{kc_show:.4f}")
        m3.metric("ค่าบำรุง AC (บาท/ตร.ม./ปี)", f"{r_ac_show:.4f}")
        m4.metric("ค่าบำรุง Concrete (บาท/ตร.ม./ปี)", f"{r_cc_show:.4f}")

        col_d1, col_d2 = st.columns(2)
        with col_d1:
            st.markdown("**Ka รายปี (X4 เปลี่ยนตามอายุ)**")
            ka_df = st.session_state.get("_ka_detail_df")
            if ka_df is not None:
                st.dataframe(ka_df, hide_index=True, use_container_width=True, height=300)
        with col_d2:
            st.markdown("**Kc Factors**")
            kc_fac = st.session_state.get("_kc_factors", {})
            if kc_fac:
                st.dataframe(pd.DataFrame({
                    "Factor": list(kc_fac.keys()),
                    "ค่า":    [round(v,4) for v in kc_fac.values()],
                }), hide_index=True, use_container_width=True)

    # Legacy Gravel
    with st.expander("⚙️ ผิวลูกรัง (Legacy — ปิดใช้งานเป็นค่าเริ่มต้น)"):
        show_gr = st.checkbox("แสดงการคำนวณผิวลูกรัง (Ks)", value=st.session_state["show_gravel"], key="sg_t2")
        st.session_state["show_gravel"] = show_gr
        if show_gr:
            st.warning("⚠️ ผิวลูกรังไม่ใช้งานในโครงการปัจจุบัน (DOH เลิกใช้แล้ว) — สำรองไว้สำหรับงานวิจัย/อ้างอิงข้อมูลเก่าเท่านั้น")

# ─────────────────────────────────────────────────────────────────────────────
# TAB 3: LCCA Analysis
# ─────────────────────────────────────────────────────────────────────────────
with tab3:
    st.header("📊 LCCA Analysis")

    # พารามิเตอร์
    st.subheader("⚙️ พารามิเตอร์การวิเคราะห์")
    pc1, pc2, pc3 = st.columns(3)
    with pc1:
        st.session_state["lcca_n"]  = st.number_input("ระยะเวลาวิเคราะห์ (ปี)",
            min_value=5, max_value=50, value=int(st.session_state["lcca_n"]), step=1, key="ln_t3")
    with pc2:
        st.session_state["lcca_dr"] = st.number_input("อัตราคิดลด (%/ปี)",
            min_value=1.0, max_value=20.0, value=float(st.session_state["lcca_dr"])*100,
            step=0.5, key="ld_t3") / 100.0
    with pc3:
        st.session_state["lcca_salvage"] = st.checkbox("รวมมูลค่าซาก (Salvage Value)",
            value=st.session_state["lcca_salvage"], key="ls_t3")

    n  = st.session_state["lcca_n"]
    dr = st.session_state["lcca_dr"]

    st.divider()

    # ── สร้าง/จัดการ Alternatives ────────────────────────────────────────────
    st.subheader("🏗️ ทางเลือกผิวทาง (Alternatives)")

    # ดึงข้อมูลจาก TAB1 และ TAB2
    cost_map_t1 = {
        "AC":   st.session_state["cost_ac"],
        "JPCP": st.session_state["cost_jpcp"],
        "JRCP": st.session_state["cost_jrcp"],
        "CRCP": st.session_state["cost_crcp"],
    }
    r_ac_val = st.session_state.get("routine_ac_per_sqm") or 0.0
    r_cc_val = st.session_state.get("routine_cc_per_sqm") or 0.0
    area_val = st.session_state["area_sqm"]

    if st.button("🔄 สร้าง/รีเซ็ต Alternatives จากข้อมูล TAB 1 & 2", key="gen_alt_t3"):
        # Default rehab plans per type
        def make_alt(name, ptype, cost, maint_cost):
            if "AC" in ptype or "Flexible" in ptype:
                rehab_yr = max(10, n//2)
                maint_list = [
                    MaintenanceActivity("บำรุงรักษาประจำปี (Routine)", maint_cost, 1, 1),
                    MaintenanceActivity("Seal Coating", maint_cost*0.8, 3, 3),
                ]
                rehab_list = [RehabActivity(f"Overlay AC 50 มม.", cost*0.25, rehab_yr)]
            else:
                rehab_yr = max(15, int(n*0.75))
                maint_list = [
                    MaintenanceActivity("บำรุงรักษาประจำปี (Routine)", maint_cost, 1, 1),
                    MaintenanceActivity("Joint Maintenance", maint_cost*0.5, 5, 5),
                ]
                rehab_list = []
            return PavementAlternative(name=name, pave_type=ptype,
                construction_cost=cost, area=area_val,
                maintenance=maint_list, rehab=rehab_list,
                salvage_pct=20.0 if "AC" in ptype else 30.0)

        alts = []
        if cost_map_t1["AC"] > 0:
            alts.append(make_alt("ผิวทางยืดหยุ่น (AC)", "Flexible", cost_map_t1["AC"], r_ac_val))
        if cost_map_t1["JPCP"] > 0:
            alts.append(make_alt("JPCP", "JPCP", cost_map_t1["JPCP"], r_cc_val))
        if cost_map_t1["JRCP"] > 0:
            alts.append(make_alt("JRCP", "JRCP", cost_map_t1["JRCP"], r_cc_val))
        if cost_map_t1["CRCP"] > 0:
            alts.append(make_alt("CRCP", "CRCP", cost_map_t1["CRCP"], r_cc_val))
        if not alts:
            st.warning("⚠️ กรุณากรอกราคาก่อสร้างใน TAB 1 ก่อน")
        else:
            st.session_state["lcca_alternatives"] = alts
            st.success(f"✅ สร้าง {len(alts)} ทางเลือกสำเร็จ")

    # แก้ไข Alternatives
    alts = st.session_state.get("lcca_alternatives") or []
    if alts:
        st.markdown("**แก้ไขแผนบำรุงรักษาและฟื้นฟูสภาพ:**")
        for ai, alt in enumerate(alts):
            with st.expander(f"✏️ {alt.name} | ต้นทุน: {alt.construction_cost:,.0f} บาท/ตร.ม."):
                col_e1, col_e2 = st.columns(2)
                with col_e1:
                    new_cost = st.number_input("ต้นทุนก่อสร้าง (บาท/ตร.ม.)",
                        min_value=0.0, value=float(alt.construction_cost), step=10.0,
                        key=f"alt_cost_{ai}")
                    alts[ai].construction_cost = new_cost
                    new_sv = st.number_input("มูลค่าซาก (%)",
                        min_value=0.0, max_value=100.0, value=float(alt.salvage_pct), step=1.0,
                        key=f"alt_sv_{ai}")
                    alts[ai].salvage_pct = new_sv
                    alts[ai].enabled = st.checkbox("เปิดใช้งาน", value=alt.enabled, key=f"alt_en_{ai}")

                with col_e2:
                    st.markdown("**แผนบำรุงรักษา:**")
                    for mi, m in enumerate(alt.maintenance):
                        mc1, mc2 = st.columns([2,1])
                        with mc1:
                            new_mc = st.number_input(f"{m.name} (บาท/ตร.ม./ปี)",
                                min_value=0.0, value=float(m.unit_cost), step=1.0,
                                key=f"m_cost_{ai}_{mi}")
                            alts[ai].maintenance[mi].unit_cost = new_mc
                        with mc2:
                            new_mf = st.number_input("ความถี่ (ปี)",
                                min_value=0, value=int(m.frequency), step=1,
                                key=f"m_freq_{ai}_{mi}")
                            alts[ai].maintenance[mi].frequency = new_mf

                    st.markdown("**แผนฟื้นฟูสภาพ:**")
                    for ri2, r in enumerate(alt.rehab):
                        rc1, rc2 = st.columns([2,1])
                        with rc1:
                            new_rc = st.number_input(f"{r.name} (บาท/ตร.ม.)",
                                min_value=0.0, value=float(r.unit_cost), step=10.0,
                                key=f"r_cost_{ai}_{ri2}")
                            alts[ai].rehab[ri2].unit_cost = new_rc
                        with rc2:
                            new_ry = st.number_input("ปีที่ดำเนินการ",
                                min_value=1, max_value=n, value=int(r.year), step=1,
                                key=f"r_yr_{ai}_{ri2}")
                            alts[ai].rehab[ri2].year = new_ry

        st.session_state["lcca_alternatives"] = alts

        st.divider()

        # ── คำนวณ LCCA ──────────────────────────────────────────────────────
        if st.button("🚀 คำนวณ LCCA", type="primary", key="run_lcca_t3"):
            with st.spinner("กำลังคำนวณ..."):
                summary_df, cf_dict = analyze_lcca(
                    alts, n, dr, st.session_state["lcca_salvage"])
                st.session_state["_lcca_summary"] = summary_df
                st.session_state["_lcca_cf"]      = cf_dict

        # ── แสดงผล ──────────────────────────────────────────────────────────
        summary_df = st.session_state.get("_lcca_summary")
        cf_dict    = st.session_state.get("_lcca_cf", {})

        if summary_df is not None and len(summary_df) > 0:
            st.subheader("🏆 สรุปผล LCCA")

            # Metrics
            best = summary_df.iloc[0]
            mc1, mc2, mc3 = st.columns(3)
            mc1.metric("🥇 ทางเลือกที่ดีที่สุด", best["ทางเลือก"])
            mc2.metric("มูลค่าปัจจุบันรวมต่ำสุด", f"{best['มูลค่าปัจจุบันรวม (บาท)']:,.0f} บาท")
            mc3.metric("EAC ต่ำสุด", f"{best['EAC (บาท/ปี)']:,.0f} บาท/ปี")

            # ตารางสรุป
            disp_cols = ["อันดับ","ทางเลือก","ประเภทผิวทาง",
                         "ต้นทุนก่อสร้าง (บาท/ตร.ม.)",
                         "มูลค่าปัจจุบันรวม (บาท)","EAC (บาท/ปี)","EAC (บาท/ตร.ม./ปี)"]
            st.dataframe(summary_df[disp_cols].style.format({
                "ต้นทุนก่อสร้าง (บาท/ตร.ม.)": "{:,.0f}",
                "มูลค่าปัจจุบันรวม (บาท)": "{:,.0f}",
                "EAC (บาท/ปี)": "{:,.0f}",
                "EAC (บาท/ตร.ม./ปี)": "{:,.2f}",
            }), hide_index=True, use_container_width=True)

            # กราฟ Bar เปรียบเทียบ NPV
            st.subheader("📊 เปรียบเทียบมูลค่าปัจจุบันรวม (NPV)")
            fig_bar = go.Figure()
            colors = {"ก่อสร้าง":"#1f77b4","บำรุงรักษา":"#ff7f0e",
                      "ฟื้นฟูสภาพ":"#d62728","มูลค่าซาก":"#2ca02c"}
            for ptype, col_key in [("ก่อสร้าง","PW_ก่อสร้าง"),
                                   ("บำรุงรักษา","PW_บำรุงรักษา"),
                                   ("ฟื้นฟูสภาพ","PW_ฟื้นฟูสภาพ"),
                                   ("มูลค่าซาก","PW_มูลค่าซาก")]:
                fig_bar.add_trace(go.Bar(
                    name=ptype, x=summary_df["ทางเลือก"],
                    y=summary_df[col_key], marker_color=colors[ptype]))
            fig_bar.update_layout(barmode="relative", height=450,
                title="มูลค่าปัจจุบันรวม แยกตามประเภทต้นทุน",
                yaxis_title="บาท", xaxis_title="ทางเลือก")
            st.plotly_chart(fig_bar, use_container_width=True)

            # กราฟ EAC
            fig_eac = px.bar(summary_df, x="ทางเลือก", y="EAC (บาท/ปี)",
                color="ทางเลือก", title="ต้นทุนเฉลี่ยรายปี (EAC)",
                text_auto=".3s", height=400)
            fig_eac.update_traces(textposition="outside")
            st.plotly_chart(fig_eac, use_container_width=True)

            # Sensitivity Analysis
            st.subheader("📈 Sensitivity Analysis — อัตราคิดลด")
            sens_rows = []
            dr_range = np.linspace(max(dr-0.03, 0.01), dr+0.03, 7)
            for dr_i in dr_range:
                for alt in alts:
                    if not alt.enabled: continue
                    cf_i = build_cashflow(alt, n, dr_i, st.session_state["lcca_salvage"])
                    pw_i = cf_i["มูลค่าปัจจุบัน"].sum()
                    sens_rows.append({"อัตราคิดลด (%)": round(dr_i*100, 1),
                                      "ทางเลือก": alt.name, "NPV (บาท)": pw_i})
            sens_df = pd.DataFrame(sens_rows)
            fig_sens = px.line(sens_df, x="อัตราคิดลด (%)", y="NPV (บาท)",
                color="ทางเลือก", markers=True, height=400,
                title="Sensitivity Analysis — ผลกระทบของอัตราคิดลดต่อ NPV")
            st.plotly_chart(fig_sens, use_container_width=True)

            # ตารางกระแสเงินสด
            st.subheader("💰 ตารางกระแสเงินสดรายทางเลือก")
            alt_sel = st.selectbox("เลือกทางเลือก", list(cf_dict.keys()), key="cfsel_t3")
            if alt_sel in cf_dict:
                cf_show = cf_dict[alt_sel].copy()
                cf_show["ต้นทุน/หน่วย"]    = cf_show["ต้นทุน/หน่วย"].map(lambda x: f"{x:,.2f}")
                cf_show["ต้นทุนตามปี"]     = cf_show["ต้นทุนตามปี"].map(lambda x: f"{x:,.0f}")
                cf_show["PW_factor"]        = cf_show["PW_factor"].map(lambda x: f"{x:.4f}")
                cf_show["มูลค่าปัจจุบัน"]  = cf_show["มูลค่าปัจจุบัน"].map(lambda x: f"{x:,.0f}")
                st.dataframe(cf_show, hide_index=True, use_container_width=True, height=400)

# ─────────────────────────────────────────────────────────────────────────────
# TAB 4: Word Report
# ─────────────────────────────────────────────────────────────────────────────
with tab4:
    st.header("📄 Word Report — รูปแบบที่ปรึกษา")

    if not DOCX_AVAILABLE:
        st.error("❌ ต้องติดตั้ง python-docx: `pip install python-docx`")
    else:
        summary_df = st.session_state.get("_lcca_summary")
        cf_dict    = st.session_state.get("_lcca_cf", {})

        if summary_df is None or len(summary_df) == 0:
            st.warning("⚠️ กรุณาคำนวณ LCCA ใน TAB 3 ก่อน")
        else:
            # แสดงสรุปก่อน Generate
            st.subheader("✅ พร้อมสร้างรายงาน")
            st.dataframe(summary_df[["อันดับ","ทางเลือก","ประเภทผิวทาง",
                                     "มูลค่าปัจจุบันรวม (บาท)","EAC (บาท/ปี)"]]\
                         .style.format({"มูลค่าปัจจุบันรวม (บาท)":"{:,.0f}","EAC (บาท/ปี)":"{:,.0f}"}),
                         hide_index=True, use_container_width=True)

            # ตั้งค่า section numbering
            base_sec_input = st.text_input("Base Section (เช่น 4.1 หรือ 3.5)",
                value="4.1", key="base_sec_t4",
                help="เลข section หลักในรายงานที่ปรึกษา")

            if st.button("📋 สร้างรายงาน Word", type="primary", key="gen_word_t4"):
                with st.spinner("กำลังสร้างรายงาน..."):
                    try:
                        alts = st.session_state.get("lcca_alternatives") or []
                        word_buf = generate_word_report(
                            summary_df, cf_dict,
                            st.session_state["lcca_n"],
                            st.session_state["lcca_dr"],
                            alts,
                            base_sec=base_sec_input.strip() or "4.1",
                        )
                        proj = st.session_state["project_name"].replace(" ","_")
                        st.download_button(
                            "⬇️ ดาวน์โหลด Word Report",
                            data=word_buf,
                            file_name=f"LCCA_Report_{proj}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key="dl_word_t4",
                        )
                        st.success("✅ สร้างรายงานสำเร็จ — คลิกปุ่มด้านบนเพื่อดาวน์โหลด")
                    except Exception as e:
                        st.error(f"สร้างรายงานไม่ได้: {e}")
                        st.exception(e)

            st.divider()
            st.caption("รายงานประกอบด้วย: ข้อมูลโครงการ | ราคาก่อสร้าง | ค่าบำรุงรักษา | สรุปผล LCCA | กระแสเงินสดรายทางเลือก")
