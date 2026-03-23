#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
โปรแกรมวิเคราะห์ต้นทุนตลอดอายุการใช้งานผิวทาง (LCCA) - เวอร์ชัน 2.2
Pavement Life-Cycle Cost Analysis Program
================================================================================
พัฒนาโดย: รศ.ดร.อิทธิพล มีผล (Assoc. Prof. Dr. Ittipon Meepon)
ภาควิชาครุศาสตร์โยธา 
มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าพระนครเหนือ (KMUTNB)
Department of Civil Engineering Education
King Mongkut's University of Technology North Bangkok

พัฒนาสำหรับการเรียนการสอนและงานวิจัยด้านวิศวกรรมทาง

คุณสมบัติเวอร์ชัน 2.1:
- แก้ไขต้นทุนก่อสร้างได้เอง
- กำหนดพื้นที่โครงการได้เอง
- เพิ่มผิวทาง JRCP (Jointed Reinforced Concrete Pavement)
- แก้ไขแผนบำรุงรักษาและฟื้นฟูสภาพได้
- Upload Excel Template
- บันทึก/โหลดโครงการ

ประเภทผิวทาง:
1. Flexible Pavement (ผิวทางยืดหยุ่น/แอสฟัลต์)
2. JPCP - Jointed Plain Concrete Pavement (คอนกรีตไม่เสริมเหล็ก)
3. JRCP - Jointed Reinforced Concrete Pavement (คอนกรีตเสริมเหล็ก)
4. CRCP - Continuously Reinforced Concrete Pavement (คอนกรีตเสริมเหล็กต่อเนื่อง)
================================================================================
"""

import streamlit as st
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass, field
import json
import io
from datetime import datetime
import hashlib

# สำหรับ Excel
try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


# สำหรับส่งออก Word
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

# ตั้งค่าหน้าเว็บ
st.set_page_config(
    page_title="โปรแกรมวิเคราะห์ LCCA ผิวทาง v2.2",
    page_icon="🛣️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =============================================================================
# ส่วนที่ 1: โครงสร้างข้อมูล (Data Structures)
# =============================================================================

@dataclass
class กิจกรรมบำรุงรักษา:
    """โครงสร้างข้อมูลกิจกรรมบำรุงรักษา"""
    ชื่อกิจกรรม: str
    ต้นทุนต่อหน่วย: float  # บาท/ตร.ม.
    ปีเริ่มต้น: int
    ความถี่: int = 0  # 0 = ครั้งเดียว


@dataclass
class กิจกรรมฟื้นฟูสภาพ:
    """โครงสร้างข้อมูลกิจกรรมฟื้นฟูสภาพ"""
    ชื่อกิจกรรม: str
    ต้นทุนต่อหน่วย: float  # บาท/ตร.ม.
    ปีดำเนินการ: int


@dataclass
class ทางเลือกผิวทาง:
    """โครงสร้างข้อมูลทางเลือกผิวทาง"""
    ชื่อ: str
    ประเภท: str
    ต้นทุนก่อสร้าง: float  # บาท/ตร.ม.
    แผนบำรุงรักษา: List[กิจกรรมบำรุงรักษา]
    แผนฟื้นฟูสภาพ: List[กิจกรรมฟื้นฟูสภาพ]
    ร้อยละมูลค่าซาก: float = 20.0
    พื้นที่: float = 1000.0  # ตร.ม.
    ความหนา: float = 0.0  # ซม. (0 = ไม่ระบุ)
    เปิดใช้งาน: bool = True


# =============================================================================
# ส่วนที่ 2: ฟังก์ชันคำนวณหลัก (Core Calculation Functions)
# =============================================================================

def คำนวณมูลค่าปัจจุบัน(ต้นทุน: float, ปี: int, อัตราคิดลด: float) -> float:
    """
    คำนวณมูลค่าปัจจุบัน (Present Worth)
    สูตร: PW = FV × (1 + i)^(-n)
    """
    if ปี < 0 or อัตราคิดลด < 0:
        return 0.0
    pwf = (1 + อัตราคิดลด) ** (-ปี)
    return ต้นทุน * pwf


def คำนวณต้นทุนเฉลี่ยรายปี(pw: float, อัตราคิดลด: float, ระยะวิเคราะห์: int) -> float:
    """
    แปลงมูลค่าปัจจุบันเป็นต้นทุนเฉลี่ยรายปี (EAC)
    สูตร: EAC = PW × [i(1+i)^N] / [(1+i)^N - 1]
    """
    if ระยะวิเคราะห์ <= 0 or อัตราคิดลด <= 0:
        return 0.0
    ตัวเศษ = อัตราคิดลด * (1 + อัตราคิดลด) ** ระยะวิเคราะห์
    ตัวส่วน = (1 + อัตราคิดลด) ** ระยะวิเคราะห์ - 1
    crf = ตัวเศษ / ตัวส่วน
    return pw * crf


def คำนวณมูลค่าซาก(
    ต้นทุนฟื้นฟูครั้งสุดท้าย: float,
    ปีฟื้นฟูครั้งสุดท้าย: int,
    อายุใช้งานที่คาดหวัง: int,
    ระยะวิเคราะห์: int,
    ร้อยละมูลค่าซาก: float = 20.0
) -> float:
    """คำนวณมูลค่าซากโดยวิธี Straight-Line Depreciation"""
    อายุใช้งานที่เหลือ = อายุใช้งานที่คาดหวัง - (ระยะวิเคราะห์ - ปีฟื้นฟูครั้งสุดท้าย)
    
    if อายุใช้งานที่เหลือ <= 0:
        return ต้นทุนฟื้นฟูครั้งสุดท้าย * (ร้อยละมูลค่าซาก / 100.0)
    
    ค่าเสื่อมต่อปี = ต้นทุนฟื้นฟูครั้งสุดท้าย * (1 - ร้อยละมูลค่าซาก/100.0) / อายุใช้งานที่คาดหวัง
    มูลค่าซาก = ต้นทุนฟื้นฟูครั้งสุดท้าย - ค่าเสื่อมต่อปี * (ระยะวิเคราะห์ - ปีฟื้นฟูครั้งสุดท้าย)
    
    return max(มูลค่าซาก, ต้นทุนฟื้นฟูครั้งสุดท้าย * ร้อยละมูลค่าซาก / 100.0)


# =============================================================================
# ส่วนที่ 3: สร้างตารางกระแสเงินสด
# =============================================================================

def สร้างตารางกระแสเงินสด(
    ทางเลือก: ทางเลือกผิวทาง,
    ระยะวิเคราะห์: int,
    อัตราคิดลด: float,
    รวมมูลค่าซาก: bool = True
) -> pd.DataFrame:
    """
    สร้างตารางกระแสเงินสดรายปี
    
    Logic แบบ C - รีเซ็ตรอบบำรุงรักษาหลังฟื้นฟูสภาพ:
    - เมื่อทำงานฟื้นฟู (Rehabilitation) ผิวทางเหมือนใหม่
    - รอบบำรุงรักษาเริ่มนับใหม่จากปีที่ทำฟื้นฟู
    - ไม่ทำบำรุงรักษาในปีเดียวกับฟื้นฟู
    """
    รายการ = []
    พื้นที่ = ทางเลือก.พื้นที่
    
    # เรียงลำดับปีฟื้นฟูสภาพ
    ปีฟื้นฟูทั้งหมด = sorted([ฟ.ปีดำเนินการ for ฟ in ทางเลือก.แผนฟื้นฟูสภาพ if ฟ.ปีดำเนินการ <= ระยะวิเคราะห์])
    ปีฟื้นฟู_set = set(ปีฟื้นฟูทั้งหมด)
    
    # ปีที่ 0: ต้นทุนก่อสร้างเริ่มต้น
    ต้นทุนเริ่มต้น = ทางเลือก.ต้นทุนก่อสร้าง * พื้นที่
    รายการ.append({
        'ปี': 0,
        'กิจกรรม': 'ก่อสร้างเริ่มต้น',
        'ประเภท': 'ก่อสร้าง',
        'ต้นทุนต่อหน่วย': ทางเลือก.ต้นทุนก่อสร้าง,
        'ต้นทุนตามปี': ต้นทุนเริ่มต้น,
        'ตัวคูณ_PW': 1.0,
        'มูลค่าปัจจุบัน': ต้นทุนเริ่มต้น
    })
    
    # กิจกรรมบำรุงรักษา (รีเซ็ตรอบหลังฟื้นฟู)
    for บำรุง in ทางเลือก.แผนบำรุงรักษา:
        if บำรุง.ความถี่ > 0:
            # สร้างช่วงเวลา: [0, ปีฟื้นฟู1, ปีฟื้นฟู2, ..., ระยะวิเคราะห์]
            จุดเริ่มต้นช่วง = [0] + ปีฟื้นฟูทั้งหมด
            
            for idx, ปีเริ่มช่วง in enumerate(จุดเริ่มต้นช่วง):
                # หาจุดสิ้นสุดช่วง
                if idx + 1 < len(จุดเริ่มต้นช่วง):
                    ปีสิ้นสุดช่วง = จุดเริ่มต้นช่วง[idx + 1]
                else:
                    ปีสิ้นสุดช่วง = ระยะวิเคราะห์ + 1
                
                # คำนวณปีบำรุงรักษาในช่วงนี้ (เริ่มนับจาก ปีเริ่มช่วง + ความถี่)
                ปี = ปีเริ่มช่วง + บำรุง.ความถี่
                while ปี < ปีสิ้นสุดช่วง and ปี <= ระยะวิเคราะห์:
                    # ข้ามถ้าตรงกับปีฟื้นฟู
                    if ปี not in ปีฟื้นฟู_set:
                        ต้นทุน = บำรุง.ต้นทุนต่อหน่วย * พื้นที่
                        pwf = (1 + อัตราคิดลด) ** (-ปี)
                        รายการ.append({
                            'ปี': ปี,
                            'กิจกรรม': บำรุง.ชื่อกิจกรรม,
                            'ประเภท': 'บำรุงรักษา',
                            'ต้นทุนต่อหน่วย': บำรุง.ต้นทุนต่อหน่วย,
                            'ต้นทุนตามปี': ต้นทุน,
                            'ตัวคูณ_PW': pwf,
                            'มูลค่าปัจจุบัน': ต้นทุน * pwf
                        })
                    ปี += บำรุง.ความถี่
        else:
            # บำรุงรักษาครั้งเดียว (ไม่รีเซ็ต แต่ข้ามถ้าตรงกับปีฟื้นฟู)
            if บำรุง.ปีเริ่มต้น <= ระยะวิเคราะห์ and บำรุง.ปีเริ่มต้น not in ปีฟื้นฟู_set:
                ต้นทุน = บำรุง.ต้นทุนต่อหน่วย * พื้นที่
                pwf = (1 + อัตราคิดลด) ** (-บำรุง.ปีเริ่มต้น)
                รายการ.append({
                    'ปี': บำรุง.ปีเริ่มต้น,
                    'กิจกรรม': บำรุง.ชื่อกิจกรรม,
                    'ประเภท': 'บำรุงรักษา',
                    'ต้นทุนต่อหน่วย': บำรุง.ต้นทุนต่อหน่วย,
                    'ต้นทุนตามปี': ต้นทุน,
                    'ตัวคูณ_PW': pwf,
                    'มูลค่าปัจจุบัน': ต้นทุน * pwf
                })
    
    # กิจกรรมฟื้นฟูสภาพ
    ต้นทุนฟื้นฟูสุดท้าย = ทางเลือก.ต้นทุนก่อสร้าง * พื้นที่
    ปีฟื้นฟูสุดท้าย = 0
    
    for ฟื้นฟู in ทางเลือก.แผนฟื้นฟูสภาพ:
        if ฟื้นฟู.ปีดำเนินการ <= ระยะวิเคราะห์:
            ต้นทุน = ฟื้นฟู.ต้นทุนต่อหน่วย * พื้นที่
            pwf = (1 + อัตราคิดลด) ** (-ฟื้นฟู.ปีดำเนินการ)
            รายการ.append({
                'ปี': ฟื้นฟู.ปีดำเนินการ,
                'กิจกรรม': ฟื้นฟู.ชื่อกิจกรรม,
                'ประเภท': 'ฟื้นฟูสภาพ',
                'ต้นทุนต่อหน่วย': ฟื้นฟู.ต้นทุนต่อหน่วย,
                'ต้นทุนตามปี': ต้นทุน,
                'ตัวคูณ_PW': pwf,
                'มูลค่าปัจจุบัน': ต้นทุน * pwf
            })
            ต้นทุนฟื้นฟูสุดท้าย = ต้นทุน
            ปีฟื้นฟูสุดท้าย = ฟื้นฟู.ปีดำเนินการ
    
    # มูลค่าซาก
    if รวมมูลค่าซาก:
        # กำหนดอายุที่คาดหวังตามประเภทผิวทาง
        if 'Flexible' in ทางเลือก.ประเภท or 'ยืดหยุ่น' in ทางเลือก.ประเภท:
            อายุที่คาดหวัง = 15
        elif 'CRCP' in ทางเลือก.ประเภท:
            อายุที่คาดหวัง = 25
        else:  # JPCP, JRCP
            อายุที่คาดหวัง = 20
            
        sv = คำนวณมูลค่าซาก(
            ต้นทุนฟื้นฟูสุดท้าย, ปีฟื้นฟูสุดท้าย, อายุที่คาดหวัง,
            ระยะวิเคราะห์, ทางเลือก.ร้อยละมูลค่าซาก
        )
        pwf = (1 + อัตราคิดลด) ** (-ระยะวิเคราะห์)
        รายการ.append({
            'ปี': ระยะวิเคราะห์,
            'กิจกรรม': 'มูลค่าซาก',
            'ประเภท': 'มูลค่าซาก',
            'ต้นทุนต่อหน่วย': -sv / พื้นที่,
            'ต้นทุนตามปี': -sv,
            'ตัวคูณ_PW': pwf,
            'มูลค่าปัจจุบัน': -sv * pwf
        })
    
    df = pd.DataFrame(รายการ)
    df = df.sort_values(['ปี', 'กิจกรรม']).reset_index(drop=True)
    
    return df


# =============================================================================
# ส่วนที่ 4: เครื่องมือวิเคราะห์ LCCA
# =============================================================================

def วิเคราะห์_LCCA(
    ทางเลือกทั้งหมด: List[ทางเลือกผิวทาง],
    ระยะวิเคราะห์: int,
    อัตราคิดลด: float,
    รวมมูลค่าซาก: bool = True
) -> Tuple[pd.DataFrame, Dict[str, pd.DataFrame]]:
    """วิเคราะห์ LCCA สำหรับทางเลือกผิวทางหลายทางเลือก"""
    สรุป_รายการ = []
    กระแสเงินสด_dict = {}
    
    # กรองเฉพาะทางเลือกที่เปิดใช้งาน
    ทางเลือกที่ใช้ = [ท for ท in ทางเลือกทั้งหมด if ท.เปิดใช้งาน]
    
    for ทางเลือก in ทางเลือกที่ใช้:
        cf_table = สร้างตารางกระแสเงินสด(ทางเลือก, ระยะวิเคราะห์, อัตราคิดลด, รวมมูลค่าซาก)
        กระแสเงินสด_dict[ทางเลือก.ชื่อ] = cf_table
        
        # คำนวณผลรวม
        รวม_nominal = cf_table['ต้นทุนตามปี'].sum()
        รวม_pw = cf_table['มูลค่าปัจจุบัน'].sum()
        eac = คำนวณต้นทุนเฉลี่ยรายปี(รวม_pw, อัตราคิดลด, ระยะวิเคราะห์)
        
        ก่อสร้าง = cf_table[cf_table['ประเภท'] == 'ก่อสร้าง']['มูลค่าปัจจุบัน'].sum()
        บำรุงรักษา = cf_table[cf_table['ประเภท'] == 'บำรุงรักษา']['มูลค่าปัจจุบัน'].sum()
        ฟื้นฟู = cf_table[cf_table['ประเภท'] == 'ฟื้นฟูสภาพ']['มูลค่าปัจจุบัน'].sum()
        ซาก = cf_table[cf_table['ประเภท'] == 'มูลค่าซาก']['มูลค่าปัจจุบัน'].sum()
        
        # ตรวจสอบ attribute ความหนา
        ความหนา = getattr(ทางเลือก, 'ความหนา', 0.0)
        
        สรุป_รายการ.append({
            'ทางเลือก': ทางเลือก.ชื่อ,
            'ประเภทผิวทาง': ทางเลือก.ประเภท,
            'ความหนา_ซม': ความหนา,
            'พื้นที่_ตรม': ทางเลือก.พื้นที่,
            'ต้นทุนก่อสร้าง_ตรม': ทางเลือก.ต้นทุนก่อสร้าง,
            'PW_ก่อสร้าง': ก่อสร้าง,
            'PW_บำรุงรักษา': บำรุงรักษา,
            'PW_ฟื้นฟูสภาพ': ฟื้นฟู,
            'PW_มูลค่าซาก': ซาก,
            'ต้นทุนตามปีรวม': รวม_nominal,
            'มูลค่าปัจจุบันรวม': รวม_pw,
            'ต้นทุนเฉลี่ยรายปี': eac,
            'ต้นทุนต่อตรม_ต่อปี': eac / ทางเลือก.พื้นที่
        })
    
    สรุป_df = pd.DataFrame(สรุป_รายการ)
    if len(สรุป_df) > 0:
        สรุป_df['ลำดับ'] = สรุป_df['มูลค่าปัจจุบันรวม'].rank().astype(int)
        สรุป_df = สรุป_df.sort_values('มูลค่าปัจจุบันรวม').reset_index(drop=True)
    
    return สรุป_df, กระแสเงินสด_dict


# =============================================================================
# ส่วนที่ 5: การวิเคราะห์ความไว
# =============================================================================

def วิเคราะห์ความไว_อัตราคิดลด(
    ทางเลือกทั้งหมด: List[ทางเลือกผิวทาง],
    ระยะวิเคราะห์: int,
    อัตราฐาน: float,
    ช่วงการเปลี่ยนแปลง: float = 0.02,
    จำนวนจุด: int = 5,
    รวมมูลค่าซาก: bool = True
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """วิเคราะห์ความไวต่ออัตราคิดลด"""
    อัตราทดสอบ = np.linspace(max(อัตราฐาน - ช่วงการเปลี่ยนแปลง, 0.001), 
                              อัตราฐาน + ช่วงการเปลี่ยนแปลง, จำนวนจุด)
    ผลลัพธ์ = []
    
    ทางเลือกที่ใช้ = [ท for ท in ทางเลือกทั้งหมด if ท.เปิดใช้งาน]
    
    for อัตรา in อัตราทดสอบ:
        for ทางเลือก in ทางเลือกที่ใช้:
            cf = สร้างตารางกระแสเงินสด(ทางเลือก, ระยะวิเคราะห์, อัตรา, รวมมูลค่าซาก)
            pw = cf['มูลค่าปัจจุบัน'].sum()
            ผลลัพธ์.append({
                'อัตราคิดลด': อัตรา,
                'อัตราคิดลด_%': f"{อัตรา*100:.1f}%",
                'ทางเลือก': ทางเลือก.ชื่อ,
                'มูลค่าปัจจุบัน': pw,
                'ต้นทุนเฉลี่ยรายปี': คำนวณต้นทุนเฉลี่ยรายปี(pw, อัตรา, ระยะวิเคราะห์)
            })
    
    df = pd.DataFrame(ผลลัพธ์)
    pivot = df.pivot(index='อัตราคิดลด_%', columns='ทางเลือก', values='มูลค่าปัจจุบัน') if len(df) > 0 else pd.DataFrame()
    
    return df, pivot


# =============================================================================
# ส่วนที่ 6: ข้อมูลตัวอย่างเริ่มต้น (รวม JRCP)
# =============================================================================

def สร้างทางเลือกเริ่มต้น() -> List[ทางเลือกผิวทาง]:
    """สร้างทางเลือกผิวทางเริ่มต้น 4 ประเภท"""
    
    # 1. ผิวทางยืดหยุ่น (Flexible Pavement)
    flexible = ทางเลือกผิวทาง(
        ชื่อ="ผิวทางยืดหยุ่น (AC)",
        ประเภท="Flexible",
        ต้นทุนก่อสร้าง=1800.0,
        แผนบำรุงรักษา=[
            กิจกรรมบำรุงรักษา("บำรุงรักษาประจำปี", 25.0, ปีเริ่มต้น=1, ความถี่=1),
            กิจกรรมบำรุงรักษา("Seal Coating", 25.0, ปีเริ่มต้น=3, ความถี่=3),
            กิจกรรมบำรุงรักษา("ซ่อมเฉพาะจุด", 50.0, ปีเริ่มต้น=5, ความถี่=5),
        ],
        แผนฟื้นฟูสภาพ=[
            กิจกรรมฟื้นฟูสภาพ("Overlay AC 50 มม.", 450.0, ปีดำเนินการ=12),
            กิจกรรมฟื้นฟูสภาพ("ก่อสร้าง AC ใหม่", 1800.0, ปีดำเนินการ=20),
        ],
        ร้อยละมูลค่าซาก=20.0,
        พื้นที่=10000.0,
        ความหนา=15.0,
        เปิดใช้งาน=True
    )
    
    # 2. JPCP - Jointed Plain Concrete Pavement (คอนกรีตไม่เสริมเหล็ก)
    jpcp = ทางเลือกผิวทาง(
        ชื่อ="JPCP",
        ประเภท="JPCP",
        ต้นทุนก่อสร้าง=2800.0,
        แผนบำรุงรักษา=[
            กิจกรรมบำรุงรักษา("บำรุงรักษาประจำปี", 35.0, ปีเริ่มต้น=1, ความถี่=1),
            กิจกรรมบำรุงรักษา("Joint Maintenances", 35.0, ปีเริ่มต้น=5, ความถี่=5),
            กิจกรรมบำรุงรักษา("ซ่อมเฉพาะจุด", 40.0, ปีเริ่มต้น=10, ความถี่=10),
        ],
        แผนฟื้นฟูสภาพ=[
            กิจกรรมฟื้นฟูสภาพ("Diamond Grinding", 180.0, ปีดำเนินการ=20),
            กิจกรรมฟื้นฟูสภาพ("ก่อสร้าง JPCP ใหม่", 2800.0, ปีดำเนินการ=20),
        ],
        ร้อยละมูลค่าซาก=30.0,
        พื้นที่=10000.0,
        ความหนา=30.0,
        เปิดใช้งาน=True
    )
    
    # 3. JRCP - Jointed Reinforced Concrete Pavement (คอนกรีตเสริมเหล็ก)
    jrcp = ทางเลือกผิวทาง(
        ชื่อ="JRCP",
        ประเภท="JRCP",
        ต้นทุนก่อสร้าง=3000.0,
        แผนบำรุงรักษา=[
            กิจกรรมบำรุงรักษา("บำรุงรักษาประจำปี", 35.0, ปีเริ่มต้น=1, ความถี่=1),
            กิจกรรมบำรุงรักษา("Joint Maintenances", 35.0, ปีเริ่มต้น=6, ความถี่=6),
            กิจกรรมบำรุงรักษา("ซ่อมเฉพาะจุด", 35.0, ปีเริ่มต้น=12, ความถี่=12),
        ],
        แผนฟื้นฟูสภาพ=[
            กิจกรรมฟื้นฟูสภาพ("Diamond Grinding", 180.0, ปีดำเนินการ=20),
            กิจกรรมฟื้นฟูสภาพ("ก่อสร้าง JRCP ใหม่", 3000.0, ปีดำเนินการ=20),
        ],
        ร้อยละมูลค่าซาก=32.0,
        พื้นที่=10000.0,
        ความหนา=25.0,
        เปิดใช้งาน=True
    )
    
    # 4. CRCP - Continuously Reinforced Concrete Pavement
    crcp = ทางเลือกผิวทาง(
        ชื่อ="CRCP",
        ประเภท="CRCP",
        ต้นทุนก่อสร้าง=3500.0,
        แผนบำรุงรักษา=[
            กิจกรรมบำรุงรักษา("บำรุงรักษาประจำปี", 15.0, ปีเริ่มต้น=1, ความถี่=1),
            กิจกรรมบำรุงรักษา("Joint Maintenances", 15.0, ปีเริ่มต้น=8, ความถี่=8),
            กิจกรรมบำรุงรักษา("ซ่อมเฉพาะจุด", 30.0, ปีเริ่มต้น=12, ความถี่=12),
        ],
        แผนฟื้นฟูสภาพ=[
            กิจกรรมฟื้นฟูสภาพ("Diamond Grinding", 180.0, ปีดำเนินการ=20),
            กิจกรรมฟื้นฟูสภาพ("ก่อสร้าง CRCP ใหม่", 3500.0, ปีดำเนินการ=20),
        ],
        ร้อยละมูลค่าซาก=35.0,
        พื้นที่=10000.0,
        ความหนา=25.0,
        เปิดใช้งาน=True
    )
    
    return [flexible, jpcp, jrcp, crcp]


# =============================================================================
# ส่วนที่ 7: ฟังก์ชัน JSON Import/Export
# =============================================================================

def ทางเลือก_เป็น_dict(ทางเลือก: ทางเลือกผิวทาง) -> dict:
    """แปลงทางเลือกเป็น Dictionary"""
    return {
        'ชื่อ': ทางเลือก.ชื่อ,
        'ประเภท': ทางเลือก.ประเภท,
        'ต้นทุนก่อสร้าง': ทางเลือก.ต้นทุนก่อสร้าง,
        'ร้อยละมูลค่าซาก': ทางเลือก.ร้อยละมูลค่าซาก,
        'พื้นที่': ทางเลือก.พื้นที่,
        'ความหนา': getattr(ทางเลือก, 'ความหนา', 0.0),
        'เปิดใช้งาน': getattr(ทางเลือก, 'เปิดใช้งาน', True),
        'แผนบำรุงรักษา': [
            {
                'ชื่อกิจกรรม': ม.ชื่อกิจกรรม,
                'ต้นทุนต่อหน่วย': ม.ต้นทุนต่อหน่วย,
                'ปีเริ่มต้น': ม.ปีเริ่มต้น,
                'ความถี่': ม.ความถี่
            }
            for ม in ทางเลือก.แผนบำรุงรักษา
        ],
        'แผนฟื้นฟูสภาพ': [
            {
                'ชื่อกิจกรรม': ฟ.ชื่อกิจกรรม,
                'ต้นทุนต่อหน่วย': ฟ.ต้นทุนต่อหน่วย,
                'ปีดำเนินการ': ฟ.ปีดำเนินการ
            }
            for ฟ in ทางเลือก.แผนฟื้นฟูสภาพ
        ]
    }


def dict_เป็น_ทางเลือก(data: dict) -> ทางเลือกผิวทาง:
    """แปลง Dictionary เป็นทางเลือก"""
    แผนบำรุง = [
        กิจกรรมบำรุงรักษา(
            ชื่อกิจกรรม=ม['ชื่อกิจกรรม'],
            ต้นทุนต่อหน่วย=ม['ต้นทุนต่อหน่วย'],
            ปีเริ่มต้น=ม['ปีเริ่มต้น'],
            ความถี่=ม['ความถี่']
        )
        for ม in data['แผนบำรุงรักษา']
    ]
    
    แผนฟื้นฟู = [
        กิจกรรมฟื้นฟูสภาพ(
            ชื่อกิจกรรม=ฟ['ชื่อกิจกรรม'],
            ต้นทุนต่อหน่วย=ฟ['ต้นทุนต่อหน่วย'],
            ปีดำเนินการ=ฟ['ปีดำเนินการ']
        )
        for ฟ in data['แผนฟื้นฟูสภาพ']
    ]
    
    return ทางเลือกผิวทาง(
        ชื่อ=data['ชื่อ'],
        ประเภท=data['ประเภท'],
        ต้นทุนก่อสร้าง=data['ต้นทุนก่อสร้าง'],
        แผนบำรุงรักษา=แผนบำรุง,
        แผนฟื้นฟูสภาพ=แผนฟื้นฟู,
        ร้อยละมูลค่าซาก=data.get('ร้อยละมูลค่าซาก', 20.0),
        พื้นที่=data.get('พื้นที่', 10000.0),
        ความหนา=data.get('ความหนา', 0.0),
        เปิดใช้งาน=data.get('เปิดใช้งาน', True)
    )


# =============================================================================
# ส่วนที่ 8: ฟังก์ชันส่งออก Word
# =============================================================================

def สร้างรายงาน_Word(
    สรุป: pd.DataFrame,
    กระแสเงินสด: Dict[str, pd.DataFrame],
    ระยะวิเคราะห์: int,
    อัตราคิดลด: float,
    ทางเลือกทั้งหมด: List[ทางเลือกผิวทาง],
    ชื่อโครงการ: str = "โครงการก่อสร้างทางหลวง"
) -> io.BytesIO:
    """สร้างรายงาน LCCA ในรูปแบบ Word"""
    
    doc = WordDocument()
    
    # ตั้งค่าฟอนต์เริ่มต้น
    style = doc.styles['Normal']
    style.font.name = 'TH Sarabun New'
    style.font.size = Pt(14)
    
    # หัวข้อรายงาน
    title = doc.add_heading('รายงานการวิเคราะห์ต้นทุนตลอดอายุการใช้งานผิวทาง (LCCA)', level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # ชื่อโครงการ
    project_title = doc.add_heading(ชื่อโครงการ, level=1)
    project_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph()
    
    # ข้อมูลทั่วไป
    doc.add_heading('1. ข้อมูลทั่วไป', level=1)
    doc.add_paragraph(f'ชื่อโครงการ: {ชื่อโครงการ}')
    doc.add_paragraph(f'วันที่วิเคราะห์: {datetime.now().strftime("%d/%m/%Y %H:%M")}')
    doc.add_paragraph(f'ระยะเวลาวิเคราะห์: {ระยะวิเคราะห์} ปี')
    doc.add_paragraph(f'อัตราคิดลด: {อัตราคิดลด*100:.1f}%')
    doc.add_paragraph(f'จำนวนทางเลือก: {len(สรุป)} ทางเลือก')
    
    # ตารางทางเลือกที่วิเคราะห์
    doc.add_heading('2. ทางเลือกผิวทางที่วิเคราะห์', level=1)
    
    # สร้างตารางข้อมูลทางเลือก
    table1 = doc.add_table(rows=1, cols=5)
    table1.style = 'Table Grid'
    table1.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # หัวตาราง
    hdr_cells = table1.rows[0].cells
    headers1 = ['ทางเลือก', 'ประเภท', 'ความหนา (ซม.)', 'พื้นที่ (ตร.ม.)', 'ต้นทุนก่อสร้าง (บาท/ตร.ม.)']
    for i, header in enumerate(headers1):
        hdr_cells[i].text = header
        hdr_cells[i].paragraphs[0].runs[0].bold = True
    
    # ข้อมูลทางเลือก
    for _, row in สรุป.iterrows():
        row_cells = table1.add_row().cells
        row_cells[0].text = str(row['ทางเลือก'])
        row_cells[1].text = str(row['ประเภทผิวทาง'])
        row_cells[2].text = f"{row['ความหนา_ซม']:.1f}"
        row_cells[3].text = f"{row['พื้นที่_ตรม']:,.0f}"
        row_cells[4].text = f"{row['ต้นทุนก่อสร้าง_ตรม']:,.0f}"
    
    doc.add_paragraph()
    
    # ผลการวิเคราะห์
    doc.add_heading('3. ผลการวิเคราะห์ LCCA', level=1)
    
    # ตารางผลการวิเคราะห์
    table2 = doc.add_table(rows=1, cols=5)
    table2.style = 'Table Grid'
    table2.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    hdr_cells2 = table2.rows[0].cells
    headers2 = ['ลำดับ', 'ทางเลือก', 'มูลค่าปัจจุบันรวม (บาท)', 'EAC (บาท/ปี)', 'ต้นทุน (บาท/ตร.ม./ปี)']
    for i, header in enumerate(headers2):
        hdr_cells2[i].text = header
        hdr_cells2[i].paragraphs[0].runs[0].bold = True
    
    for _, row in สรุป.iterrows():
        row_cells = table2.add_row().cells
        row_cells[0].text = str(int(row['ลำดับ']))
        row_cells[1].text = str(row['ทางเลือก'])
        row_cells[2].text = f"{row['มูลค่าปัจจุบันรวม']:,.0f}"
        row_cells[3].text = f"{row['ต้นทุนเฉลี่ยรายปี']:,.0f}"
        row_cells[4].text = f"{row['ต้นทุนต่อตรม_ต่อปี']:,.2f}"
    
    doc.add_paragraph()
    
    # องค์ประกอบต้นทุน
    doc.add_heading('4. องค์ประกอบต้นทุน (มูลค่าปัจจุบัน)', level=1)
    
    table3 = doc.add_table(rows=1, cols=6)
    table3.style = 'Table Grid'
    table3.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    hdr_cells3 = table3.rows[0].cells
    headers3 = ['ทางเลือก', 'ก่อสร้าง (บาท)', 'บำรุงรักษา (บาท)', 'ฟื้นฟูสภาพ (บาท)', 'มูลค่าซาก (บาท)', 'รวม (บาท)']
    for i, header in enumerate(headers3):
        hdr_cells3[i].text = header
        hdr_cells3[i].paragraphs[0].runs[0].bold = True
    
    for _, row in สรุป.iterrows():
        row_cells = table3.add_row().cells
        row_cells[0].text = str(row['ทางเลือก'])
        row_cells[1].text = f"{row['PW_ก่อสร้าง']:,.0f}"
        row_cells[2].text = f"{row['PW_บำรุงรักษา']:,.0f}"
        row_cells[3].text = f"{row['PW_ฟื้นฟูสภาพ']:,.0f}"
        row_cells[4].text = f"{row['PW_มูลค่าซาก']:,.0f}"
        row_cells[5].text = f"{row['มูลค่าปัจจุบันรวม']:,.0f}"
    
    doc.add_paragraph()
    
    # สรุปผล
    doc.add_heading('5. สรุปผลการวิเคราะห์', level=1)
    
    ผู้ชนะ = สรุป.iloc[0]
    doc.add_paragraph(f'ทางเลือกที่ประหยัดที่สุด: {ผู้ชนะ["ทางเลือก"]}')
    doc.add_paragraph(f'มูลค่าปัจจุบันรวม: {ผู้ชนะ["มูลค่าปัจจุบันรวม"]:,.0f} บาท')
    doc.add_paragraph(f'ต้นทุนเฉลี่ยรายปี (EAC): {ผู้ชนะ["ต้นทุนเฉลี่ยรายปี"]:,.0f} บาท/ปี')
    
    if len(สรุป) > 1:
        doc.add_paragraph()
        doc.add_paragraph('การเปรียบเทียบกับทางเลือกอื่น:')
        for idx in range(1, len(สรุป)):
            อื่น = สรุป.iloc[idx]
            ส่วนต่าง = อื่น['มูลค่าปัจจุบันรวม'] - ผู้ชนะ['มูลค่าปัจจุบันรวม']
            ร้อยละ = (ส่วนต่าง / อื่น['มูลค่าปัจจุบันรวม']) * 100
            doc.add_paragraph(f'  - vs {อื่น["ทางเลือก"]}: ประหยัด {ส่วนต่าง:,.0f} บาท ({ร้อยละ:.1f}%)')
    
    # ===== เพิ่มส่วนทฤษฎีและสูตร =====
    doc.add_page_break()
    doc.add_heading('6. ทฤษฎีและสูตรการคำนวณ', level=1)
    
    # ทฤษฎี LCCA
    doc.add_heading('6.1 ทฤษฎี Life-Cycle Cost Analysis (LCCA)', level=2)
    doc.add_paragraph(
        'การวิเคราะห์ต้นทุนตลอดอายุการใช้งาน (LCCA) เป็นเครื่องมือทางเศรษฐศาสตร์วิศวกรรม '
        'ที่ใช้เปรียบเทียบทางเลือกการลงทุนต่างๆ โดยพิจารณาต้นทุนทั้งหมดตลอดอายุการใช้งานของโครงการ '
        'ประกอบด้วย ต้นทุนก่อสร้างเริ่มต้น ต้นทุนบำรุงรักษา ต้นทุนฟื้นฟูสภาพ และมูลค่าซากปลายทาง '
        'โดยแปลงต้นทุนทั้งหมดมาเป็นมูลค่าปัจจุบัน (Present Worth) เพื่อเปรียบเทียบในฐานเดียวกัน'
    )
    
    # สูตรมูลค่าปัจจุบัน
    doc.add_heading('6.2 สูตรมูลค่าปัจจุบัน (Present Worth)', level=2)
    doc.add_paragraph('สูตรแปลงต้นทุนในอนาคตมาเป็นมูลค่าปัจจุบัน:')
    doc.add_paragraph('    PW = FV × (1 + i)^(-n)', style='Normal')
    doc.add_paragraph('โดยที่:')
    doc.add_paragraph('    PW = มูลค่าปัจจุบัน (Present Worth)')
    doc.add_paragraph('    FV = มูลค่าอนาคต (Future Value)')
    doc.add_paragraph(f'    i = อัตราคิดลด (Discount Rate) = {อัตราคิดลด*100:.1f}%')
    doc.add_paragraph('    n = จำนวนปีนับจากปัจจุบัน')
    
    # สูตร EAC
    doc.add_heading('6.3 สูตรต้นทุนเฉลี่ยรายปี (EAC)', level=2)
    doc.add_paragraph('สูตรแปลงมูลค่าปัจจุบันรวมเป็นต้นทุนเฉลี่ยต่อปี:')
    doc.add_paragraph('    EAC = PW × [i × (1 + i)^n] / [(1 + i)^n - 1]', style='Normal')
    doc.add_paragraph('โดยที่:')
    doc.add_paragraph('    EAC = ต้นทุนเฉลี่ยรายปี (Equivalent Annual Cost)')
    doc.add_paragraph('    PW = มูลค่าปัจจุบันรวม')
    doc.add_paragraph(f'    i = อัตราคิดลด = {อัตราคิดลด*100:.1f}%')
    doc.add_paragraph(f'    n = ระยะเวลาวิเคราะห์ = {ระยะวิเคราะห์} ปี')
    
    # สูตรมูลค่าซาก
    doc.add_heading('6.4 สูตรมูลค่าซาก (Salvage Value)', level=2)
    doc.add_paragraph('สูตรคำนวณมูลค่าซากปลายทาง:')
    doc.add_paragraph('    SV = C × (RUL / L) × (r / 100)', style='Normal')
    doc.add_paragraph('โดยที่:')
    doc.add_paragraph('    SV = มูลค่าซาก (Salvage Value)')
    doc.add_paragraph('    C = ต้นทุนฟื้นฟูสภาพครั้งสุดท้าย')
    doc.add_paragraph('    RUL = อายุการใช้งานที่เหลือ (Remaining Useful Life)')
    doc.add_paragraph('    L = อายุที่คาดหวังของผิวทาง')
    doc.add_paragraph('    r = ร้อยละมูลค่าซาก (%)')
    
    # หมายเหตุ
    doc.add_paragraph()
    doc.add_paragraph('หมายเหตุ: รายงานนี้คำนวณตามมาตรฐาน AASHTO และ FHWA Life-Cycle Cost Analysis')
    
    # กระแสเงินสดรายละเอียด
    doc.add_page_break()
    doc.add_heading('7. รายละเอียดกระแสเงินสดแต่ละทางเลือก', level=1)
    
    for ชื่อทางเลือก, cf_table in กระแสเงินสด.items():
        doc.add_heading(f'6.{list(กระแสเงินสด.keys()).index(ชื่อทางเลือก)+1} {ชื่อทางเลือก}', level=2)
        
        # สร้างตารางกระแสเงินสด
        table_cf = doc.add_table(rows=1, cols=5)
        table_cf.style = 'Table Grid'
        
        hdr_cf = table_cf.rows[0].cells
        headers_cf = ['ปี', 'กิจกรรม', 'ต้นทุนตามปี (บาท)', 'ตัวคูณ PW', 'มูลค่าปัจจุบัน (บาท)']
        for i, header in enumerate(headers_cf):
            hdr_cf[i].text = header
            hdr_cf[i].paragraphs[0].runs[0].bold = True
        
        for _, row in cf_table.iterrows():
            row_cells = table_cf.add_row().cells
            row_cells[0].text = str(int(row['ปี']))
            row_cells[1].text = str(row['กิจกรรม'])
            row_cells[2].text = f"{row['ต้นทุนตามปี']:,.0f}"
            row_cells[3].text = f"{row['ตัวคูณ_PW']:.4f}"
            row_cells[4].text = f"{row['มูลค่าปัจจุบัน']:,.0f}"
        
        # รวม
        row_total = table_cf.add_row().cells
        row_total[0].text = ''
        row_total[1].text = 'รวม'
        row_total[1].paragraphs[0].runs[0].bold = True
        row_total[2].text = f"{cf_table['ต้นทุนตามปี'].sum():,.0f}"
        row_total[3].text = ''
        row_total[4].text = f"{cf_table['มูลค่าปัจจุบัน'].sum():,.0f}"
        row_total[4].paragraphs[0].runs[0].bold = True
        
        doc.add_paragraph()
    
    # Footer
    doc.add_paragraph()
    doc.add_paragraph('─' * 50)
    doc.add_paragraph(f'โครงการ: {ชื่อโครงการ}')
    doc.add_paragraph('รายงานนี้สร้างโดย: โปรแกรมวิเคราะห์ LCCA ผิวทาง v2.1')
    doc.add_paragraph('ภาควิชาครุศาสตร์โยธา มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าพระนครเหนือ')
    
    # บันทึกเป็น BytesIO
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    
    return file_stream


def สร้างรายงาน_Word_ที่ปรึกษา(
    สรุป: pd.DataFrame,
    กระแสเงินสด: Dict[str, pd.DataFrame],
    ระยะวิเคราะห์: int,
    อัตราคิดลด: float,
    ทางเลือกทั้งหมด: List[ทางเลือกผิวทาง],
    ชื่อโครงการ: str = "โครงการก่อสร้างทางหลวง",
    ข้อมูลโครงการ: dict = None,
    หมายเลขหัวข้อหลัก: str = "4.8"
) -> io.BytesIO:
    """
    สร้างรายงาน LCCA แบบที่ปรึกษา
    ใช้รูปแบบฟอนต์ TH SarabunPSK เหมือนรายงาน Word ต้นฉบับ
    หมายเลขหัวข้อหลักกำหนดได้ เช่น 4.8 → หัวข้อย่อยเริ่มที่ 4.8.1
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    FONT = 'TH SarabunPSK'
    SZ_H1      = Pt(15)   # Heading1  → 15pt
    SZ_H2      = Pt(15)   # Heading2  → 15pt
    SZ_BODY    = Pt(14)   # Normal    → 14pt
    SZ_FORMULA = Pt(11)   # สมการ Times New Roman 11pt

    if ข้อมูลโครงการ is None:
        ข้อมูลโครงการ = {}

    def sub(n: int) -> str:
        return f"{หมายเลขหัวข้อหลัก}.{n}"

    # ─── ฟังก์ชันช่วย set font ให้ run ───
    def set_run_font(run, size=SZ_BODY, bold=False, italic=False, color=None):
        run.font.name = FONT
        run.font.size = size
        run.font.bold = bold
        run.font.italic = italic
        if color:
            from docx.dml.color import RGBColor
            run.font.color.rgb = color
        # ตั้ง cs font (จำเป็นสำหรับภาษาไทย)
        rPr = run._r.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.insert(0, rFonts)
        rFonts.set(qn('w:ascii'), FONT)
        rFonts.set(qn('w:hAnsi'), FONT)
        rFonts.set(qn('w:cs'), FONT)

    def add_heading_para(text, level=1):
        """เพิ่ม heading ด้วย style Heading1/Heading2 และกำหนดฟอนต์ตรงๆ"""
        p = doc.add_heading(text, level=level)
        sz = SZ_H1 if level == 1 else SZ_H2
        for run in p.runs:
            set_run_font(run, size=sz, bold=True)
        return p

    def add_body_para(text='', bold=False, italic=False, first_line_indent=True):
        """
        เพิ่ม paragraph ธรรมดา ฟอนต์ TH SarabunPSK 14pt
        ใช้ thaiDistribute justification + firstLine indent เหมือน Word ต้นฉบับ
        """
        p = doc.add_paragraph()
        pPr = p._p.get_or_add_pPr()

        # Thai distribution justify
        jc = OxmlElement('w:jc')
        jc.set(qn('w:val'), 'thaiDistribute')
        pPr.append(jc)

        # First-line indent 720 twips (0.5 inch) — เหมือนต้นฉบับ
        if first_line_indent:
            ind = OxmlElement('w:ind')
            ind.set(qn('w:firstLine'), '720')
            pPr.append(ind)

        if text:
            run = p.add_run(text)
            set_run_font(run, size=SZ_BODY, bold=bold, italic=italic)
        return p

    def set_table_style(table):
        """กำหนด style ตารางให้เหมือน TableGrid และจัด center"""
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

    def add_table_header_row(table, headers: list):
        """ใส่ข้อมูลหัวตาราง bold"""
        hdr = table.rows[0].cells
        for i, h in enumerate(headers):
            hdr[i].paragraphs[0].clear()
            run = hdr[i].paragraphs[0].add_run(h)
            set_run_font(run, size=SZ_BODY, bold=True)

    def add_table_data_row(table, values: list):
        """เพิ่มแถวข้อมูลตาราง"""
        row_cells = table.add_row().cells
        for i, v in enumerate(values):
            row_cells[i].paragraphs[0].clear()
            run = row_cells[i].paragraphs[0].add_run(str(v))
            set_run_font(run, size=SZ_BODY)
        return row_cells

    # ─────────────────────────────────────────────────────────
    doc = WordDocument()

    # ตั้งค่าฟอนต์ default (Normal style)
    normal_style = doc.styles['Normal']
    normal_style.font.name = FONT
    normal_style.font.size = SZ_BODY

    # ═══════════════════════════════════════════════════════
    # หัวข้อหลัก  X.X  วิเคราะห์ต้นทุนตลอดอายุการใช้งานผิวทาง
    # ═══════════════════════════════════════════════════════
    add_heading_para(
        f"{หมายเลขหัวข้อหลัก}  วิเคราะห์ต้นทุนตลอดอายุการใช้งานผิวทาง"
        " (Life-Cycle Cost Analysis)",
        level=1
    )

    # ═══════════════════════════════════════════════════════
    # X.X.1  ทฤษฎี Life-Cycle Cost Analysis (LCCA)
    # ═══════════════════════════════════════════════════════
    add_heading_para(f"{sub(1)}  ทฤษฎี Life-Cycle Cost Analysis (LCCA)", level=1)

    add_body_para(
        'การวิเคราะห์ต้นทุนตลอดอายุการใช้งาน (LCCA) เป็นเครื่องมือทางเศรษฐศาสตร์วิศวกรรม'
        'ที่ใช้เปรียบเทียบทางเลือกการลงทุนต่างๆ'
        'โดยพิจารณาต้นทุนทั้งหมดตลอดอายุการใช้งานของโครงการ ประกอบด้วย ต้นทุนก่อสร้างเริ่มต้น'
        'ต้นทุนบำรุงรักษา ต้นทุนฟื้นฟูสภาพ และมูลค่าซากปลายทาง(ถ้ามี)'
        'โดยแปลงต้นทุนทั้งหมดมาเป็นมูลค่าปัจจุบัน (Present Worth) เพื่อเปรียบเทียบในฐานเดียวกัน'
    )

    # ═══════════════════════════════════════════════════════
    # X.X.2  สูตรมูลค่าปัจจุบัน (Present Worth)
    # ═══════════════════════════════════════════════════════
    add_heading_para(f"{sub(2)}  สูตรมูลค่าปัจจุบัน (Present Worth)", level=2)
    add_body_para('สูตรแปลงต้นทุนในอนาคตมาเป็นมูลค่าปัจจุบัน:')

    # สูตร — Times New Roman 11pt, ไม่มี indent
    p_formula = doc.add_paragraph()
    run_f = p_formula.add_run('PW = FV × (1 + i)^(-n)')
    run_f.font.name = 'Times New Roman'
    run_f.font.size = SZ_FORMULA
    run_f.font.italic = True

    add_body_para('โดยที่:')
    for line in [
        'PW = มูลค่าปัจจุบัน (Present Worth)',
        'FV = มูลค่าอนาคต (Future Value)',
        f'i  = อัตราคิดลด (Discount Rate) = {อัตราคิดลด*100:.1f}%',
        'n  = จำนวนปีนับจากปัจจุบัน',
    ]:
        add_body_para(line)

    # ═══════════════════════════════════════════════════════
    # X.X.3  สูตรต้นทุนเฉลี่ยรายปี (EAC)
    # ═══════════════════════════════════════════════════════
    add_heading_para(f"{sub(3)}  สูตรต้นทุนเฉลี่ยรายปี (EAC)", level=2)
    add_body_para('สูตรแปลงมูลค่าปัจจุบันรวมเป็นต้นทุนเฉลี่ยต่อปี:')

    # สูตร — Times New Roman 11pt, ไม่มี indent
    p_formula2 = doc.add_paragraph()
    run_f2 = p_formula2.add_run('EAC = PW × [i × (1 + i)^n] / [(1 + i)^n - 1]')
    run_f2.font.name = 'Times New Roman'
    run_f2.font.size = SZ_FORMULA
    run_f2.font.italic = True

    add_body_para('โดยที่:')
    for line in [
        'EAC = ต้นทุนเฉลี่ยรายปี (Equivalent Annual Cost)',
        'PW  = มูลค่าปัจจุบันรวม',
        f'i   = อัตราคิดลด = {อัตราคิดลด*100:.1f}%',
        f'n   = ระยะเวลาวิเคราะห์ = {ระยะวิเคราะห์} ปี',
    ]:
        add_body_para(line)

    # ═══════════════════════════════════════════════════════
    # X.X.4  ข้อมูลของโครงการสำหรับการคำนวณ
    # ═══════════════════════════════════════════════════════
    add_heading_para(f"{sub(4)}  ข้อมูลของโครงการสำหรับการคำนวณ", level=1)

    # แสดงเฉพาะค่าจำเป็น 3 รายการ
    ฟิลด์_proj = [
        ('ระยะเวลาวิเคราะห์', f'{ระยะวิเคราะห์} ปี'),
        ('อัตราคิดลด',          f'{อัตราคิดลด*100:.1f}%'),
        ('จำนวนทางเลือก',       f'{len(สรุป)} ทางเลือก'),
    ]

    for label, value in ฟิลด์_proj:
        p = doc.add_paragraph()
        run_label = p.add_run(f'{label}: ')
        set_run_font(run_label, size=SZ_BODY)
        run_value = p.add_run(str(value))
        set_run_font(run_value, size=SZ_BODY)

    # ═══════════════════════════════════════════════════════
    # X.X.5  ทางเลือกผิวทางที่วิเคราะห์
    # ═══════════════════════════════════════════════════════
    add_heading_para(f"{sub(5)}  ทางเลือกผิวทางที่วิเคราะห์", level=1)

    table1 = doc.add_table(rows=1, cols=5)
    set_table_style(table1)
    add_table_header_row(table1, ['ทางเลือก', 'ประเภท', 'ความหนา (ซม.)', 'พื้นที่ (ตร.ม.)', 'ต้นทุนก่อสร้าง (บาท/ตร.ม.)'])
    for _, row in สรุป.iterrows():
        add_table_data_row(table1, [
            row['ทางเลือก'],
            row['ประเภทผิวทาง'],
            f"{row['ความหนา_ซม']:.1f}",
            f"{row['พื้นที่_ตรม']:,.0f}",
            f"{row['ต้นทุนก่อสร้าง_ตรม']:,.0f}",
        ])

    doc.add_paragraph()

    # ═══════════════════════════════════════════════════════
    # X.X.6  ผลการวิเคราะห์ LCCA
    # ═══════════════════════════════════════════════════════
    add_heading_para(f"{sub(6)}  ผลการวิเคราะห์ LCCA", level=1)

    table2 = doc.add_table(rows=1, cols=5)
    set_table_style(table2)
    add_table_header_row(table2, ['ลำดับ', 'ทางเลือก', 'มูลค่าปัจจุบันรวม (บาท)', 'EAC (บาท/ปี)', 'ต้นทุน (บาท/ตร.ม./ปี)'])
    for _, row in สรุป.iterrows():
        add_table_data_row(table2, [
            int(row['ลำดับ']),
            row['ทางเลือก'],
            f"{row['มูลค่าปัจจุบันรวม']:,.0f}",
            f"{row['ต้นทุนเฉลี่ยรายปี']:,.0f}",
            f"{row['ต้นทุนต่อตรม_ต่อปี']:,.2f}",
        ])

    doc.add_paragraph()

    # ═══════════════════════════════════════════════════════
    # X.X.7  องค์ประกอบต้นทุน (มูลค่าปัจจุบัน)
    # ═══════════════════════════════════════════════════════
    add_heading_para(f"{sub(7)}  องค์ประกอบต้นทุน (มูลค่าปัจจุบัน)", level=1)

    table3 = doc.add_table(rows=1, cols=6)
    set_table_style(table3)
    add_table_header_row(table3, ['ทางเลือก', 'ก่อสร้าง (บาท)', 'บำรุงรักษา (บาท)', 'ฟื้นฟูสภาพ (บาท)', 'มูลค่าซาก (บาท)', 'รวม (บาท)'])
    for _, row in สรุป.iterrows():
        add_table_data_row(table3, [
            row['ทางเลือก'],
            f"{row['PW_ก่อสร้าง']:,.0f}",
            f"{row['PW_บำรุงรักษา']:,.0f}",
            f"{row['PW_ฟื้นฟูสภาพ']:,.0f}",
            f"{row['PW_มูลค่าซาก']:,.0f}",
            f"{row['มูลค่าปัจจุบันรวม']:,.0f}",
        ])

    doc.add_paragraph()

    # ═══════════════════════════════════════════════════════
    # X.X.8  รายละเอียดกระแสเงินสดแต่ละทางเลือก
    # ═══════════════════════════════════════════════════════
    add_heading_para(f"{sub(8)}  รายละเอียดกระแสเงินสดแต่ละทางเลือก", level=1)

    # สร้าง map ชื่อ → ประเภท เพื่อใช้ตั้งชื่อ sub-heading
    ประเภทแมพ = {ท.ชื่อ: ท.ประเภท for ท in ทางเลือกทั้งหมด}

    def ชื่อ_sub_heading(ชื่อ: str) -> str:
        """สร้างชื่อ sub-heading พร้อม prefix ตามประเภทผิวทาง"""
        ประเภท = ประเภทแมพ.get(ชื่อ, '')
        if 'Flexible' in ประเภท or 'ยืดหยุ่น' in ประเภท:
            return f"ผิวทางยืดหยุ่น (AC)"
        else:
            # JPCP, JRCP, CRCP → "ผิวทางคอนกรีตแบบ XXXX"
            return f"ผิวทางคอนกรีตแบบ {ชื่อ}"

    for idx_alt, (ชื่อทางเลือก, cf_table) in enumerate(กระแสเงินสด.items()):
        # Sub-heading พร้อม prefix ตามประเภทผิวทาง
        add_heading_para(ชื่อ_sub_heading(ชื่อทางเลือก), level=2)

        table_cf = doc.add_table(rows=1, cols=5)
        set_table_style(table_cf)
        add_table_header_row(table_cf, ['ปี', 'กิจกรรม', 'ต้นทุนตามปี (บาท)', 'ตัวคูณ PW', 'มูลค่าปัจจุบัน (บาท)'])

        for _, cf_row in cf_table.iterrows():
            add_table_data_row(table_cf, [
                int(cf_row['ปี']),
                cf_row['กิจกรรม'],
                f"{cf_row['ต้นทุนตามปี']:,.0f}",
                f"{cf_row['ตัวคูณ_PW']:.4f}",
                f"{cf_row['มูลค่าปัจจุบัน']:,.0f}",
            ])

        # แถวรวม — bold
        row_total = table_cf.add_row().cells
        row_total[0].paragraphs[0].clear()
        row_total[1].paragraphs[0].clear()
        run_total_label = row_total[1].paragraphs[0].add_run('รวม')
        set_run_font(run_total_label, size=SZ_BODY, bold=True)
        row_total[2].paragraphs[0].clear()
        run_total_sum = row_total[2].paragraphs[0].add_run(f"{cf_table['ต้นทุนตามปี'].sum():,.0f}")
        set_run_font(run_total_sum, size=SZ_BODY, bold=True)
        row_total[4].paragraphs[0].clear()
        run_total_pw = row_total[4].paragraphs[0].add_run(f"{cf_table['มูลค่าปัจจุบัน'].sum():,.0f}")
        set_run_font(run_total_pw, size=SZ_BODY, bold=True)

        doc.add_paragraph()

    # ═══════════════════════════════════════════════════════
    # X.X.9  สรุปผลการวิเคราะห์
    # ═══════════════════════════════════════════════════════
    add_heading_para(f"{sub(9)}  สรุปผลการวิเคราะห์", level=1)

    ผู้ชนะ = สรุป.iloc[0]

    # สร้างข้อความสรุป แบบเดียวกับ Word ต้นฉบับ
    p_summary = doc.add_paragraph()
    # Thai distribution + first-line indent
    pPr_s = p_summary._p.get_or_add_pPr()
    jc_s = OxmlElement('w:jc'); jc_s.set(qn('w:val'), 'thaiDistribute')
    pPr_s.append(jc_s)
    ind_s = OxmlElement('w:ind'); ind_s.set(qn('w:firstLine'), '720')
    pPr_s.append(ind_s)
    summary_text = (
        f'จากการวิเคราะห์ต้นทุนตลอดวงจรชีวิตของทางเลือกผิวทาง {len(สรุป)} ประเภท'
        f' พบว่า '
    )
    run_s1 = p_summary.add_run(summary_text)
    set_run_font(run_s1, size=SZ_BODY)

    run_winner = p_summary.add_run(f'{ผู้ชนะ["ทางเลือก"]}')
    set_run_font(run_winner, size=SZ_BODY, bold=True)

    run_s2 = p_summary.add_run(
        ' เป็นทางเลือกที่มีความคุ้มค่าทางเศรษฐศาสตร์สูงสุด'
        f' โดยมีมูลค่าปัจจุบันรวมเท่ากับ '
    )
    set_run_font(run_s2, size=SZ_BODY)

    run_npv = p_summary.add_run(f'{ผู้ชนะ["มูลค่าปัจจุบันรวม"]:,.0f} บาท')
    set_run_font(run_npv, size=SZ_BODY, bold=True)

    run_s3 = p_summary.add_run(' และต้นทุนเฉลี่ยรายปี (EAC) ที่ ')
    set_run_font(run_s3, size=SZ_BODY)

    run_eac = p_summary.add_run(f'{ผู้ชนะ["ต้นทุนเฉลี่ยรายปี"]:,.0f} บาท/ปี')
    set_run_font(run_eac, size=SZ_BODY, bold=True)

    if len(สรุป) > 1:
        compare_parts = []
        for idx in range(1, len(สรุป)):
            อื่น = สรุป.iloc[idx]
            ส่วนต่าง = อื่น['มูลค่าปัจจุบันรวม'] - ผู้ชนะ['มูลค่าปัจจุบันรวม']
            ร้อยละ = (ส่วนต่าง / อื่น['มูลค่าปัจจุบันรวม']) * 100
            compare_parts.append((อื่น['ทางเลือก'], ส่วนต่าง, ร้อยละ))

        run_s4 = p_summary.add_run(' เมื่อเปรียบเทียบกับทางเลือกอื่น ')
        set_run_font(run_s4, size=SZ_BODY)
        run_w2 = p_summary.add_run(f'{ผู้ชนะ["ทางเลือก"]}')
        set_run_font(run_w2, size=SZ_BODY, bold=True)

        for i, (ชื่ออื่น, diff, pct) in enumerate(compare_parts):
            sep = ' ประหยัดกว่า ' if i == 0 else ' ประหยัดกว่า '
            run_cmp = p_summary.add_run(f'{sep}{ชื่ออื่น} คิดเป็นร้อยละ {pct:.1f}')
            set_run_font(run_cmp, size=SZ_BODY)

    run_conclude = p_summary.add_run(
        f' ดังนั้น จึงมีข้อเสนอแนะให้เลือกใช้ '
    )
    set_run_font(run_conclude, size=SZ_BODY)
    run_w3 = p_summary.add_run(f'{ผู้ชนะ["ทางเลือก"]}')
    set_run_font(run_w3, size=SZ_BODY, bold=True)
    run_end = p_summary.add_run(
        ' เป็นทางเลือกหลักในการออกแบบ'
        ' เนื่องจากให้ต้นทุนรวมต่ำที่สุดตลอดอายุการใช้งานของโครงการ'
    )
    set_run_font(run_end, size=SZ_BODY)

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream


# =============================================================================
# ส่วนที่ 9: Streamlit Application
# =============================================================================


# =============================================================================
# Excel Upload Functions
# =============================================================================

def สร้าง_excel_template() -> io.BytesIO:
    """สร้าง Excel template สำหรับผู้ใช้กรอกข้อมูล (ปรับปรุงใหม่)"""
    
    if not OPENPYXL_AVAILABLE:
        return None
    
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils.dataframe import dataframe_to_rows
    from openpyxl.worksheet.datavalidation import DataValidation
    
    # สร้าง workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"  # ✅ เปลี่ยนเป็น Sheet1 เพื่อให้ตรงกับโค้ดอ่าน
    
    # === ส่วนที่ 1: Header หลัก ===
    ws['A1'] = 'ข้อมูลสำหรับวิเคราะห์ LCCA (Life-Cycle Cost Analysis)'
    ws['A1'].font = Font(size=14, bold=True, color="FFFFFF")
    ws['A1'].fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.merge_cells('A1:D1')
    
    # === ส่วนที่ 2: คำอธิบาย ===
    ws['A2'] = '💡 คำแนะนำ: กรอกข้อมูลในแถวที่ 4-7 → บันทึกไฟล์ → อัปโหลดในโปรแกรม'
    ws['A2'].font = Font(size=10, italic=True, color="0070C0")
    ws.merge_cells('A2:D2')
    
    # === ส่วนที่ 3: หัวตาราง ===
    headers = ['ผิวทาง', 'ประเภทผิวทาง', 'ความหนาผิวทาง (ซม.)', 'ต้นทุนก่อสร้าง (บาท/ตร.ม.)']
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_num)
        cell.value = header
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
    
    # === ส่วนที่ 4: ข้อมูลแถว ===
    data_rows = [
        ['AC', 'ลาดยาง', '', ''],
        ['JPCP', 'คอนกรีต', '', ''],
        ['JRCP', 'คอนกรีต', '', ''],
        ['CRCP', 'คอนกรีต', '', '']
    ]
    
    for row_idx, row_data in enumerate(data_rows, start=4):
        for col_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.value = value
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            # สีพื้นหลังสลับ
            if row_idx % 2 == 0:
                cell.fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
            
            # คอลัมน์ A ไม่ให้แก้ไข (ชื่อผิวทาง)
            if col_idx == 1:
                cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                cell.font = Font(bold=True)
    
    # === ส่วนที่ 5: Data Validation ===
    # Dropdown สำหรับประเภทผิวทาง (คอลัมน์ B)
    dv_type = DataValidation(type="list", formula1='"ลาดยาง,คอนกรีต"', allow_blank=False)
    dv_type.error = 'กรุณาเลือกจากรายการ'
    dv_type.errorTitle = 'ค่าไม่ถูกต้อง'
    ws.add_data_validation(dv_type)
    dv_type.add('B4:B7')
    
    # Validation ความหนา (คอลัมน์ C) - ต้องเป็นตัวเลข 10-50
    dv_thick = DataValidation(type="decimal", operator="between", formula1=10, formula2=50, allow_blank=True)
    dv_thick.error = 'กรุณากรอกตัวเลข 10-50 ซม.'
    dv_thick.errorTitle = 'ความหนาไม่ถูกต้อง'
    ws.add_data_validation(dv_thick)
    dv_thick.add('C4:C7')
    
    # Validation ต้นทุน (คอลัมน์ D) - ต้องเป็นตัวเลขบวก
    dv_cost = DataValidation(type="decimal", operator="greaterThan", formula1=0, allow_blank=False)
    dv_cost.error = 'กรุณากรอกตัวเลขมากกว่า 0'
    dv_cost.errorTitle = 'ต้นทุนไม่ถูกต้อง'
    ws.add_data_validation(dv_cost)
    dv_cost.add('D4:D7')
    
    # === ส่วนที่ 7: ปรับขนาดคอลัมน์ ===
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['C'].width = 25
    ws.column_dimensions['D'].width = 32
    
    # === ส่วนที่ 8: ปรับความสูงแถว ===
    ws.row_dimensions[1].height = 25
    ws.row_dimensions[3].height = 40
    
    # === ส่วนที่ 9: เพิ่ม Comments ===
    from openpyxl.comments import Comment
    
    ws['C4'].comment = Comment('กรอกความหนาผิวทาง เช่น 15, 25, 30 ซม.', 'LCCA System')
    ws['D4'].comment = Comment('กรอกต้นทุนก่อสร้าง เช่น 1800, 2500 บาท/ตร.ม.', 'LCCA System')
    
    # บันทึกไฟล์
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def อ่านข้อมูลจาก_excel(uploaded_file) -> Dict[str, Dict]:
    """
    อ่านข้อมูลจากไฟล์ Excel ที่ผู้ใช้ upload (ปรับปรุงใหม่ + Error Handling)
    
    Returns:
        Dict mapping ชื่อผิวทางเต็ม -> {ต้นทุน, ความหนา, ประเภท}
    """
    try:
        # === ตรวจสอบ 1: ไฟล์เป็น Excel หรือไม่ ===
        if not uploaded_file.name.endswith(('.xlsx', '.xls')):
            st.error("❌ ไฟล์ต้องเป็น Excel (.xlsx หรือ .xls)")
            return {}
        
        # === ตรวจสอบ 2: อ่านไฟล์ได้หรือไม่ ===
        try:
            df = pd.read_excel(uploaded_file, sheet_name='Sheet1', header=2)
        except Exception as e:
            st.error(f"❌ ไม่สามารถอ่านไฟล์ Excel: {str(e)}")
            st.info("💡 กรุณาใช้ Template ที่ดาวน์โหลดจากโปรแกรม")
            return {}
        
        # === ตรวจสอบ 3: มี columns ครบหรือไม่ ===
        required_cols = ['ผิวทาง', 'ประเภทผิวทาง', 'ความหนาผิวทาง (ซม.)', 'ต้นทุนก่อสร้าง (บาท/ตร.ม.)']
        missing_cols = [col for col in required_cols if col not in df.columns]
        
        if missing_cols:
            st.error(f"❌ ไฟล์ขาดคอลัมน์: {', '.join(missing_cols)}")
            st.info("💡 กรุณาดาวน์โหลด Template ใหม่และกรอกข้อมูล")
            return {}
        
        # === ตรวจสอบ 4: มีข้อมูลหรือไม่ ===
        if len(df) == 0:
            st.error("❌ ไฟล์ไม่มีข้อมูล กรุณากรอกข้อมูลในแถวที่ 4-7")
            return {}
        
        # แมพชื่อย่อกับชื่อเต็มในโปรแกรม
        ชื่อแมพ = {
            'AC': 'ผิวทางยืดหยุ่น (AC)',
            'JPCP': 'JPCP',
            'JRCP': 'JRCP',
            'CRCP': 'CRCP'
        }
        
        # แมพประเภทผิวทาง
        ประเภทแมพ = {
            'ลาดยาง': 'Flexible',
            'คอนกรีต': 'JPCP',
        }
        
        # สร้าง dictionary
        ข้อมูลทั้งหมด = {}
        errors = []
        warnings = []
        
        for idx, row in df.iterrows():
            # อ่านชื่อผิวทาง
            ชื่อย่อ = str(row['ผิวทาง']).strip()
            
            # ข้ามแถวที่ว่าง, NaN, หรือไม่ใช่ชื่อที่รู้จัก
            if ชื่อย่อ in ['', 'nan', 'NaN', 'None']:
                continue
                
            if ชื่อย่อ not in ชื่อแมพ:
                # แถวนี้ไม่ใช่ข้อมูลหลัก (อาจเป็นตัวอย่าง)
                continue
            
            # === อ่านต้นทุน (บังคับ) ===
            ต้นทุน_str = str(row['ต้นทุนก่อสร้าง (บาท/ตร.ม.)'])
            try:
                ต้นทุน = float(ต้นทุน_str.replace(',', ''))
                
                # ตรวจสอบค่าสมเหตุสมผล
                if ต้นทุน <= 0:
                    errors.append(f"{ชื่อย่อ}: ต้นทุนต้องมากกว่า 0 (ได้รับ: {ต้นทุน})")
                    continue
                elif ต้นทุน > 10000:
                    warnings.append(f"{ชื่อย่อ}: ต้นทุนสูงมาก ({ต้นทุน:,.0f} บาท/ตร.ม.)")
                
            except:
                errors.append(f"{ชื่อย่อ}: ต้นทุนไม่ถูกต้อง (ได้รับ: '{ต้นทุน_str}')")
                continue
            
            # === อ่านความหนา (ไม่บังคับ) ===
            ความหนา = None
            if 'ความหนาผิวทาง (ซม.)' in row:
                ความหนา_str = str(row['ความหนาผิวทาง (ซม.)'])
                try:
                    if ความหนา_str not in ['', 'nan', 'NaN', 'กรอกข้อมูล']:
                        ความหนา = float(ความหนา_str.replace(',', ''))
                        
                        # ตรวจสอบค่าสมเหตุสมผล
                        if ความหนา < 10 or ความหนา > 50:
                            warnings.append(f"{ชื่อย่อ}: ความหนาผิดปกติ ({ความหนา} ซม.)")
                except:
                    warnings.append(f"{ชื่อย่อ}: ความหนาไม่ถูกต้อง ('{ความหนา_str}') - ข้าม")
            
            # === อ่านประเภท (ไม่บังคับ) ===
            ประเภท = None
            if 'ประเภทผิวทาง' in row:
                ประเภท_str = str(row['ประเภทผิวทาง']).strip()
                if ประเภท_str in ประเภทแมพ:
                    ประเภท = ประเภทแมพ[ประเภท_str]
            
            # ใช้ชื่อเต็มจากแมพ
            ชื่อเต็ม = ชื่อแมพ.get(ชื่อย่อ, ชื่อย่อ)
            
            # เก็บข้อมูลทั้งหมด
            ข้อมูลทั้งหมด[ชื่อเต็ม] = {
                'ต้นทุน': ต้นทุน,
                'ความหนา': ความหนา,
                'ประเภท': ประเภท
            }
        
        # === แสดง Errors และ Warnings ===
        if errors:
            st.error("❌ พบข้อผิดพลาด:")
            for err in errors:
                st.write(f"  • {err}")
        
        if warnings:
            st.warning("⚠️ คำเตือน:")
            for warn in warnings:
                st.write(f"  • {warn}")
        
        # === ตรวจสอบผลลัพธ์สุดท้าย ===
        if len(ข้อมูลทั้งหมด) == 0:
            st.error("❌ ไม่พบข้อมูลที่ถูกต้อง กรุณาตรวจสอบไฟล์อีกครั้ง")
            st.info("💡 ต้นทุนเป็นข้อมูลบังคับ และต้องเป็นตัวเลขมากกว่า 0")
        
        return ข้อมูลทั้งหมด
    
    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาดไม่คาดคิด: {str(e)}")
        st.info("💡 กรุณาดาวน์โหลด Template ใหม่และลองอีกครั้ง")
        return {}


def main():
    """Main Streamlit Application"""
    
    # หัวข้อหลัก
    st.title("🛣️ โปรแกรมวิเคราะห์ต้นทุนตลอดอายุการใช้งานผิวทาง (LCCA)")
    st.markdown("""
    **Life-Cycle Cost Analysis for Pavement Alternatives - Version 2.2**
    
    **พัฒนาโดย:** รศ.ดร.อิทธิพล มีผล (Assoc. Prof. Dr. Ittipon Meepon)  
    ภาควิชาครุศาสตร์โยธา มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าพระนครเหนือ  
    Department of Civil Engineering Education, KMUTNB
    
    พัฒนาสำหรับการเรียนการสอนและงานวิจัยด้านวิศวกรรมทาง
    
    ✨ **v2.2:** แก้ไข JSON load ไม่อัปเดต, Reset มี Confirm, พื้นที่ Auto-apply, เพิ่ม Loading spinner
    """)
    
    st.divider()
    
    # ==========================================================================
    # Initialize Session State
    # ==========================================================================
    if 'ทางเลือกทั้งหมด' not in st.session_state:
        st.session_state.ทางเลือกทั้งหมด = สร้างทางเลือกเริ่มต้น()
    # json_version ใช้เป็น suffix ของ widget key ทุกตัวใน Tab 1
    # เมื่อ load JSON หรือ reset → เพิ่ม 1 → Streamlit ถือว่า widget ใหม่ → อ่านค่าจาก value= ใหม่
    if 'json_version' not in st.session_state:
        st.session_state['json_version'] = 0
    
    # ==========================================================================
    # Sidebar: พารามิเตอร์การวิเคราะห์
    # ==========================================================================
    with st.sidebar:
        st.header("⚙️ พารามิเตอร์การวิเคราะห์")
        
        # ชื่อโครงการ
        if 'ชื่อโครงการ' not in st.session_state:
            st.session_state.ชื่อโครงการ = "โครงการก่อสร้างทางหลวง"
        
        # ไม่ใช้ key= เพื่อหลีกเลี่ยง conflict เมื่อ JSON load เขียน session_state
        # ใช้ value=session_state แทน → Streamlit อ่านค่าใหม่ทุก rerun อัตโนมัติ
        ชื่อโครงการ = st.text_input(
            "🏗️ ชื่อโครงการ",
            value=st.session_state.ชื่อโครงการ,
            help="ระบุชื่อโครงการสำหรับแสดงในรายงาน"
        )
        st.session_state.ชื่อโครงการ = ชื่อโครงการ

        st.divider()
        
        ระยะวิเคราะห์ = st.slider(
            "ระยะเวลาวิเคราะห์ (ปี)",
            min_value=20, max_value=50, value=35, step=5,
            help="ระยะเวลาที่ใช้ในการวิเคราะห์เปรียบเทียบทางเลือก"
        )
        
        อัตราคิดลด = st.slider(
            "อัตราคิดลด (%)",
            min_value=2.0, max_value=10.0, value=4.0, step=0.5,
            help="Real Discount Rate (ไม่รวมอัตราเงินเฟ้อ)"
        ) / 100.0
        
        st.divider()
        
        # Toggle มูลค่าซาก
        st.subheader("💰 มูลค่าซาก (Salvage Value)")
        
        รวมมูลค่าซาก = st.toggle(
            "รวมมูลค่าซากในการคำนวณ",
            value=True,
            help="เปิดใช้งานเพื่อนำมูลค่าซากมาหักออกจากต้นทุนทั้งหมด"
        )
        
        if รวมมูลค่าซาก:
            st.info("✅ กำลังคำนวณมูลค่าซาก")
        else:
            st.warning("⚠️ ไม่คำนวณมูลค่าซาก")
        
        st.divider()
        
        # Excel Upload Section
        with st.expander("📤 อัปโหลดข้อมูลต้นทุน (Excel)", expanded=False):
            if OPENPYXL_AVAILABLE:
                # ดาวน์โหลด Template
                template_file = สร้าง_excel_template()
                if template_file:
                    st.download_button(
                        label="📥 ดาวน์โหลด Template Excel",
                        data=template_file,
                        file_name="LCCA_Template.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.document",
                        use_container_width=True
                    )
                
                st.caption("💡 ดาวน์โหลด template → กรอกข้อมูล → อัปโหลดกลับมา")
                
                # Upload File
                uploaded_file = st.file_uploader(
                    "เลือกไฟล์ Excel:",
                    type=['xlsx', 'xls'],
                    help="อัปโหลดไฟล์ Excel ที่กรอกข้อมูลแล้ว"
                )
                
                # Process uploaded file
                if uploaded_file is not None:
                    try:
                        ข้อมูลต้นทุน = อ่านข้อมูลจาก_excel(uploaded_file)
                        
                        if len(ข้อมูลต้นทุน) > 0:
                            file_hash = hashlib.md5(uploaded_file.getvalue()).hexdigest()
                            
                            # อัปเดตเฉพาะไฟล์ใหม่เท่านั้น
                            if st.session_state.get('loaded_excel_hash') != file_hash:
                                อัปเดตแล้ว = 0
                                for ท in st.session_state.ทางเลือกทั้งหมด:
                                    if ท.ชื่อ in ข้อมูลต้นทุน:
                                        info = ข้อมูลต้นทุน[ท.ชื่อ]
                                        if isinstance(info, dict):
                                            ท.ต้นทุนก่อสร้าง = info.get('ต้นทุน', ท.ต้นทุนก่อสร้าง)
                                            if info.get('ความหนา') is not None:
                                                ท.ความหนา = info['ความหนา']
                                            if info.get('ประเภท') is not None:
                                                ท.ประเภท = info['ประเภท']
                                        else:
                                            ท.ต้นทุนก่อสร้าง = float(info)
                                        อัปเดตแล้ว += 1
                                
                                st.session_state['loaded_excel_hash'] = file_hash
                                # force widget refresh เหมือน JSON load
                                st.session_state['json_version'] = st.session_state.get('json_version', 0) + 1
                                # เก็บ preview ไว้แสดงหลัง rerun
                                st.session_state['excel_preview'] = ข้อมูลต้นทุน
                                st.session_state['excel_upload_msg'] = f"✅ อัปเดตข้อมูลสำเร็จ! ({อัปเดตแล้ว} ทางเลือก)"
                                st.rerun()
                            else:
                                # ไฟล์เดิม — แสดง preview จาก session_state
                                st.success(st.session_state.get('excel_upload_msg', f"✅ โหลดข้อมูลแล้ว ({len(ข้อมูลต้นทุน)} รายการ)"))
                            
                            # แสดง preview เสมอถ้ามีข้อมูล
                            if st.session_state.get('excel_preview'):
                                with st.expander("👀 ดูข้อมูลที่อัปโหลด", expanded=True):
                                    for ชื่อ, ข้อมูล in st.session_state['excel_preview'].items():
                                        if isinstance(ข้อมูล, dict):
                                            ต้นทุน = ข้อมูล.get('ต้นทุน', 0)
                                            ความหนา = ข้อมูล.get('ความหนา')
                                            ประเภท = ข้อมูล.get('ประเภท')
                                            info_parts = [f"{ต้นทุน:,.0f} บาท/ตร.ม."]
                                            if ประเภท:
                                                info_parts.append(f"ประเภท: {ประเภท}")
                                            if ความหนา:
                                                info_parts.append(f"หนา: {ความหนา:.1f} ซม.")
                                            st.write(f"• **{ชื่อ}**: {', '.join(info_parts)}")
                                        else:
                                            st.write(f"• **{ชื่อ}**: {float(ข้อมูล):,.0f} บาท/ตร.ม.")
                        else:
                            st.error("❌ ไม่พบข้อมูลที่ถูกต้องในไฟล์")
                    
                    except Exception as e:
                        st.error(f"❌ เกิดข้อผิดพลาด: {str(e)}")
            else:
                st.warning("⚠️ ต้องติดตั้ง openpyxl: `pip install openpyxl`")
        
        st.divider()
        
        # พื้นที่โครงการ (ใช้ร่วมกันทุกทางเลือก)
        st.subheader("📐 พื้นที่โครงการ")
        
        # ดึงค่าพื้นที่ปัจจุบันจาก session_state
        if 'พื้นที่ร่วม' not in st.session_state:
            st.session_state['พื้นที่ร่วม'] = 10000.0
        
        พื้นที่ร่วม = st.number_input(
            "พื้นที่ (ตร.ม.)",
            min_value=100.0,
            max_value=1000000.0,
            value=st.session_state['พื้นที่ร่วม'],
            step=1000.0,
            help="เปลี่ยนค่าแล้ว Apply อัตโนมัติทุกทางเลือก"
        )
        
        # Auto-apply เมื่อค่าเปลี่ยน
        if พื้นที่ร่วม != st.session_state['พื้นที่ร่วม']:
            st.session_state['พื้นที่ร่วม'] = พื้นที่ร่วม
            for ท in st.session_state.ทางเลือกทั้งหมด:
                ท.พื้นที่ = พื้นที่ร่วม
            st.rerun()
        
        st.caption(f"✅ พื้นที่ปัจจุบัน: {พื้นที่ร่วม:,.0f} ตร.ม. (ใช้ร่วมกันทุกทางเลือก)")
        
        st.divider()
        
        st.subheader("📊 การวิเคราะห์ความไว")
        ช่วงอัตราคิดลด = st.slider(
            "ช่วงการเปลี่ยนแปลงอัตราคิดลด (±%)",
            min_value=1.0, max_value=4.0, value=2.0, step=0.5
        ) / 100.0
        
        st.divider()
        
        # === บันทึก/โหลดโครงการ ===
        with st.expander("💾 จัดการโครงการ (บันทึก/โหลด)", expanded=False):
            # Export JSON
            col_save1, col_save2 = st.columns(2)
            
            with col_save1:
                # สร้างข้อมูล JSON
                project_data = {
                    'ชื่อโครงการ': st.session_state.ชื่อโครงการ,
                    'ทางเลือก': [ทางเลือก_เป็น_dict(ท) for ท in st.session_state.ทางเลือกทั้งหมด],
                    'วันที่บันทึก': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
                
                json_str = json.dumps(project_data, ensure_ascii=False, indent=2)
                
                st.download_button(
                    label="💾 บันทึกโครงการ",
                    data=json_str,
                    file_name=f"LCCA_{st.session_state.ชื่อโครงการ}_{datetime.now().strftime('%Y%m%d')}.json",
                    mime="application/json",
                    use_container_width=True,
                    help="บันทึกข้อมูลโครงการทั้งหมดเป็นไฟล์ JSON"
                )
            
            with col_save2:
                # Reset button — ต้องกด 2 ครั้ง ป้องกันกดพลาด
                if not st.session_state.get('confirm_reset', False):
                    if st.button("🔄 รีเซ็ต", use_container_width=True, help="รีเซ็ตเป็นค่าเริ่มต้น"):
                        st.session_state['confirm_reset'] = True
                        st.rerun()
                else:
                    st.warning("⚠️ ยืนยันรีเซ็ต?")
                    col_yes, col_no = st.columns(2)
                    with col_yes:
                        if st.button("✅ ยืนยัน", use_container_width=True, type="primary"):
                            st.session_state.ทางเลือกทั้งหมด = สร้างทางเลือกเริ่มต้น()
                            st.session_state.ชื่อโครงการ = "โครงการก่อสร้างทางหลวง"
                            st.session_state['confirm_reset'] = False
                            st.session_state['json_version'] = st.session_state.get('json_version', 0) + 1
                            st.session_state.pop('loaded_json_hash', None)
                            st.session_state.pop('json_load_msg', None)
                            st.session_state.pop('uploaded_cost_data', None)
                            st.session_state.pop('loaded_excel_hash', None)
                            st.session_state.pop('excel_preview', None)
                            st.session_state.pop('excel_upload_msg', None)
                            st.rerun()
                    with col_no:
                        if st.button("❌ ยกเลิก", use_container_width=True):
                            st.session_state['confirm_reset'] = False
                            st.rerun()
            
            # Import JSON
            uploaded_json = st.file_uploader(
                "📂 โหลดโครงการที่บันทึกไว้",
                type=['json'],
                help="อัปโหลดไฟล์ .json ที่บันทึกไว้",
                key="json_uploader"
            )
            
            if uploaded_json is not None:
                try:
                    # ── Bug Fix: ใช้ hash ป้องกัน load ซ้ำทุก rerun ──
                    file_hash = hashlib.md5(uploaded_json.getvalue()).hexdigest()
                    
                    if st.session_state.get('loaded_json_hash') != file_hash:
                        data = json.loads(uploaded_json.getvalue().decode('utf-8'))
                        
                        # ตรวจสอบโครงสร้างข้อมูล
                        if 'ทางเลือก' not in data:
                            st.error("❌ ไฟล์ JSON ไม่ถูกต้อง: ไม่พบข้อมูล 'ทางเลือก'")
                        else:
                            st.session_state.ทางเลือกทั้งหมด = [dict_เป็น_ทางเลือก(d) for d in data['ทางเลือก']]
                            
                            if 'ชื่อโครงการ' in data:
                                st.session_state.ชื่อโครงการ = data['ชื่อโครงการ']
                                # widget ใช้ value=session_state → อ่านค่าใหม่อัตโนมัติหลัง rerun
                            
                            วันที่บันทึก = data.get('วันที่บันทึก', 'ไม่ทราบ')
                            # บันทึก hash ป้องกัน load ซ้ำ
                            st.session_state['loaded_json_hash'] = file_hash
                            st.session_state['json_load_msg'] = f"✅ โหลดโครงการสำเร็จ! (บันทึกเมื่อ {วันที่บันทึก})"
                            # เพิ่ม version → force widget ทุกตัวใน Tab 1 อ่านค่าใหม่
                            st.session_state['json_version'] = st.session_state.get('json_version', 0) + 1
                            st.rerun()
                    else:
                        # แสดงข้อความสำเร็จหลัง rerun
                        if st.session_state.get('json_load_msg'):
                            st.success(st.session_state['json_load_msg'])
                        
                except json.JSONDecodeError:
                    st.error("❌ ไฟล์ไม่ใช่ JSON ที่ถูกต้อง")
                except Exception as e:
                    st.error(f"❌ เกิดข้อผิดพลาด: {str(e)}")
                    st.info("💡 กรุณาตรวจสอบว่าเป็นไฟล์ที่บันทึกจากโปรแกรมนี้")
    
    # ==========================================================================
    # แท็บหลัก
    # ==========================================================================
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📝 ข้อมูลทางเลือก",
        "📊 ผลการวิเคราะห์",
        "💰 กระแสเงินสด",
        "📈 วิเคราะห์ความไว",
        "ℹ️ ทฤษฎี LCCA"
    ])
    
    # ==========================================================================
    # แท็บ 1: แก้ไขข้อมูลทางเลือก
    # ==========================================================================
    with tab1:
        st.header("📝 ข้อมูลทางเลือกผิวทาง")
        
        st.info("""
        💡 **คำแนะนำ:** 
        - กำหนดต้นทุนก่อสร้าง ความหนา และพื้นที่ได้ในแต่ละทางเลือก
        - เปิด/ปิดทางเลือกที่ต้องการเปรียบเทียบ
        - กำหนดแผนบำรุงรักษาและฟื้นฟูสภาพได้
        """)
        
        # แสดงข้อมูลแต่ละทางเลือก
        v = st.session_state.get('json_version', 0)  # version สำหรับ force refresh widget
        for i, ทางเลือก in enumerate(st.session_state.ทางเลือกทั้งหมด):
            with st.expander(f"{'✅' if ทางเลือก.เปิดใช้งาน else '❌'} ทางเลือกที่ {i+1}: {ทางเลือก.ชื่อ}", expanded=(i==0)):
                
                # แถวบนสุด: เปิด/ปิดใช้งาน
                col_enable = st.columns([3, 1])
                with col_enable[1]:
                    เปิดใช้ = st.checkbox(
                        "เปิดใช้งาน",
                        value=ทางเลือก.เปิดใช้งาน,
                        key=f"enable_{i}_v{v}"
                    )
                    ทางเลือก.เปิดใช้งาน = เปิดใช้
                
                st.markdown("---")
                
                # ข้อมูลหลัก
                st.subheader("🏗️ ข้อมูลหลัก")
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    ชื่อใหม่ = st.text_input(
                        "ชื่อทางเลือก",
                        value=ทางเลือก.ชื่อ,
                        key=f"name_{i}_v{v}"
                    )
                    ทางเลือก.ชื่อ = ชื่อใหม่
                
                with col2:
                    ประเภทใหม่ = st.selectbox(
                        "ประเภทผิวทาง",
                        options=["Flexible", "JPCP", "JRCP", "CRCP"],
                        index=["Flexible", "JPCP", "JRCP", "CRCP"].index(ทางเลือก.ประเภท) if ทางเลือก.ประเภท in ["Flexible", "JPCP", "JRCP", "CRCP"] else 0,
                        key=f"type_{i}_v{v}"
                    )
                    ทางเลือก.ประเภท = ประเภทใหม่
                
                with col3:
                    # อ่านจาก ทางเลือก โดยตรง (JSON load อัปเดตแล้ว)
                    ความหนาปัจจุบัน = getattr(ทางเลือก, 'ความหนา', 0.0)
                    if ความหนาปัจจุบัน == 0.0:
                        defaults = {'Flexible': 15.0, 'JPCP': 30.0, 'JRCP': 25.0, 'CRCP': 25.0}
                        ความหนาปัจจุบัน = defaults.get(ทางเลือก.ประเภท, 25.0)
                    
                    ความหนาใหม่ = st.number_input(
                        "ความหนา (ซม.)",
                        min_value=0.0,
                        max_value=100.0,
                        value=float(ความหนาปัจจุบัน),
                        step=1.0,
                        key=f"thickness_{i}_v{v}"
                    )
                    ทางเลือก.ความหนา = ความหนาใหม่
                
                col4, col5, col6 = st.columns(3)
                
                with col4:
                    # อ่านจาก ทางเลือก โดยตรง (JSON load อัปเดตแล้ว)
                    ต้นทุนใหม่ = st.number_input(
                        "ต้นทุนก่อสร้าง (บาท/ตร.ม.)",
                        min_value=0.0,
                        max_value=10000.0,
                        value=float(ทางเลือก.ต้นทุนก่อสร้าง),
                        step=100.0,
                        key=f"cost_{i}_v{v}"
                    )
                    ทางเลือก.ต้นทุนก่อสร้าง = ต้นทุนใหม่
                
                with col5:
                    # แสดงพื้นที่แบบ read-only (แก้ไขที่ Sidebar เท่านั้น)
                    st.metric(
                        label="พื้นที่ (ตร.ม.)",
                        value=f"{ทางเลือก.พื้นที่:,.0f}",
                        help="ปรับพื้นที่ได้ที่ Sidebar → พื้นที่โครงการ"
                    )
                
                with col6:
                    # แสดงช่องร้อยละมูลค่าซากเฉพาะเมื่อเปิด Toggle
                    if รวมมูลค่าซาก:
                        ซากใหม่ = st.number_input(
                            "ร้อยละมูลค่าซาก (%)",
                            min_value=0.0,
                            max_value=50.0,
                            value=float(ทางเลือก.ร้อยละมูลค่าซาก),
                            step=5.0,
                            key=f"salvage_{i}_v{v}"
                        )
                        ทางเลือก.ร้อยละมูลค่าซาก = ซากใหม่
                    else:
                        # เมื่อปิด Toggle → ไม่แสดงช่อง แต่เก็บค่าเดิมไว้
                        st.info("⚠️ ปิดการคำนวณมูลค่าซาก")
                
                # แสดงต้นทุนก่อสร้างรวม
                st.metric(
                    "💰 ต้นทุนก่อสร้างรวม",
                    f"{ทางเลือก.ต้นทุนก่อสร้าง * ทางเลือก.พื้นที่:,.0f} บาท"
                )
                
                st.markdown("---")
                
                # แผนบำรุงรักษา
                st.subheader("🔧 แผนบำรุงรักษา")
                
                # หัวข้อคอลัมน์
                col_h1, col_h2, col_h3, col_h4, col_h5 = st.columns([3, 2, 1, 1, 0.5])
                with col_h1:
                    st.markdown("**กิจกรรม**")
                with col_h2:
                    st.markdown("**ต้นทุน (บาท/ตร.ม.)**")
                with col_h3:
                    st.markdown("**ปีเริ่มต้น**")
                with col_h4:
                    st.markdown("**ทุกๆ (ปี)**")
                with col_h5:
                    st.markdown("**ลบ**")
                
                รายการที่จะลบ_maint = []
                for j, บำรุง in enumerate(ทางเลือก.แผนบำรุงรักษา):
                    col_m1, col_m2, col_m3, col_m4, col_m5 = st.columns([3, 2, 1, 1, 0.5])
                    
                    with col_m1:
                        บำรุง.ชื่อกิจกรรม = st.text_input(
                            "กิจกรรม",
                            value=บำรุง.ชื่อกิจกรรม,
                            key=f"maint_name_{i}_{j}_v{v}",
                            label_visibility="collapsed"
                        )
                    
                    with col_m2:
                        บำรุง.ต้นทุนต่อหน่วย = st.number_input(
                            "บาท/ตร.ม.",
                            min_value=0.0,
                            value=float(บำรุง.ต้นทุนต่อหน่วย),
                            step=5.0,
                            key=f"maint_cost_{i}_{j}_v{v}",
                            label_visibility="collapsed"
                        )
                    
                    with col_m3:
                        บำรุง.ปีเริ่มต้น = st.number_input(
                            "ปีเริ่ม",
                            min_value=1,
                            max_value=50,
                            value=int(บำรุง.ปีเริ่มต้น),
                            key=f"maint_year_{i}_{j}_v{v}",
                            label_visibility="collapsed"
                        )
                    
                    with col_m4:
                        บำรุง.ความถี่ = st.number_input(
                            "ทุกๆ (ปี)",
                            min_value=0,
                            max_value=20,
                            value=int(บำรุง.ความถี่),
                            key=f"maint_freq_{i}_{j}_v{v}",
                            label_visibility="collapsed"
                        )
                    
                    with col_m5:
                        if st.button("🗑️", key=f"del_maint_{i}_{j}_v{v}", help="ลบรายการนี้"):
                            รายการที่จะลบ_maint.append(j)
                
                # ลบรายการที่เลือก
                for idx in sorted(รายการที่จะลบ_maint, reverse=True):
                    if len(ทางเลือก.แผนบำรุงรักษา) > 1:
                        ทางเลือก.แผนบำรุงรักษา.pop(idx)
                        st.rerun()
                
                # ปุ่มเพิ่มรายการบำรุงรักษา
                if st.button(f"➕ เพิ่มกิจกรรมบำรุงรักษา", key=f"add_maint_{i}_v{v}"):
                    ทางเลือก.แผนบำรุงรักษา.append(
                        กิจกรรมบำรุงรักษา("บำรุงรักษาประจำปี", 50.0, ปีเริ่มต้น=1, ความถี่=1)
                    )
                    st.rerun()
                
                st.markdown("---")
                
                # แผนฟื้นฟูสภาพ
                st.subheader("🚀 แผนฟื้นฟูสภาพ")
                
                # หัวข้อคอลัมน์
                col_rh1, col_rh2, col_rh3, col_rh4 = st.columns([4, 2, 1, 0.5])
                with col_rh1:
                    st.markdown("**กิจกรรม**")
                with col_rh2:
                    st.markdown("**ต้นทุน (บาท/ตร.ม.)**")
                with col_rh3:
                    st.markdown("**ปีที่ดำเนินการ**")
                with col_rh4:
                    st.markdown("**ลบ**")
                
                # แสดงรายการฟื้นฟูสภาพ
                รายการที่จะลบ_rehab = []
                
                # สร้าง key version สำหรับ force update
                if f'rehab_update_version_{i}' not in st.session_state:
                    st.session_state[f'rehab_update_version_{i}'] = 0
                
                for k, ฟื้นฟู in enumerate(ทางเลือก.แผนฟื้นฟูสภาพ):
                    col_r1, col_r2, col_r3, col_r4 = st.columns([4, 2, 1, 0.5])
                    
                    with col_r1:
                        ชื่อกิจกรรมเก่า = ฟื้นฟู.ชื่อกิจกรรม
                        rv = st.session_state[f'rehab_update_version_{i}']
                        ชื่อกิจกรรมใหม่ = st.text_input(
                            "กิจกรรม",
                            value=ฟื้นฟู.ชื่อกิจกรรม,
                            key=f"rehab_name_{i}_{k}_v{v}_{rv}",
                            label_visibility="collapsed"
                        )
                        
                        if ชื่อกิจกรรมใหม่ != ชื่อกิจกรรมเก่า:
                            ฟื้นฟู.ชื่อกิจกรรม = ชื่อกิจกรรมใหม่
                            if "ก่อสร้างใหม่" in ชื่อกิจกรรมใหม่ or ("ก่อสร้าง" in ชื่อกิจกรรมใหม่ and "ใหม่" in ชื่อกิจกรรมใหม่):
                                ฟื้นฟู.ต้นทุนต่อหน่วย = float(ทางเลือก.ต้นทุนก่อสร้าง)
                                st.session_state[f'rehab_update_version_{i}'] += 1
                                st.rerun()
                        else:
                            ฟื้นฟู.ชื่อกิจกรรม = ชื่อกิจกรรมใหม่
                    
                    with col_r2:
                        ต้นทุนเริ่มต้น_ฟื้นฟู = float(ฟื้นฟู.ต้นทุนต่อหน่วย)
                        if "ก่อสร้างใหม่" in ฟื้นฟู.ชื่อกิจกรรม or ("ก่อสร้าง" in ฟื้นฟู.ชื่อกิจกรรม and "ใหม่" in ฟื้นฟู.ชื่อกิจกรรม):
                            ต้นทุนเริ่มต้น_ฟื้นฟู = float(ทางเลือก.ต้นทุนก่อสร้าง)
                        
                        ฟื้นฟู.ต้นทุนต่อหน่วย = st.number_input(
                            "บาท/ตร.ม.",
                            min_value=0.0,
                            value=ต้นทุนเริ่มต้น_ฟื้นฟู,
                            step=10.0,
                            key=f"rehab_cost_{i}_{k}_v{v}_{rv}",
                            label_visibility="collapsed",
                            help="💡 พิมพ์ 'ก่อสร้างใหม่' แล้วกด Enter → ดึงต้นทุนก่อสร้างอัตโนมัติ"
                        )
                    
                    with col_r3:
                        ฟื้นฟู.ปีดำเนินการ = st.number_input(
                            "ปีที่",
                            min_value=1,
                            max_value=50,
                            value=int(ฟื้นฟู.ปีดำเนินการ),
                            key=f"rehab_year_{i}_{k}_v{v}",
                            label_visibility="collapsed"
                        )
                    
                    with col_r4:
                        if st.button("🗑️", key=f"del_rehab_{i}_{k}_v{v}", help="ลบรายการนี้"):
                            รายการที่จะลบ_rehab.append(k)
                
                # ลบรายการที่เลือก
                for idx in sorted(รายการที่จะลบ_rehab, reverse=True):
                    if len(ทางเลือก.แผนฟื้นฟูสภาพ) > 1:
                        ทางเลือก.แผนฟื้นฟูสภาพ.pop(idx)
                        st.rerun()
                
                # ปุ่มเพิ่มรายการฟื้นฟูสภาพ
                if st.button(f"➕ เพิ่มกิจกรรมฟื้นฟูสภาพ", key=f"add_rehab_{i}_v{v}"):
                    ทางเลือก.แผนฟื้นฟูสภาพ.append(
                        กิจกรรมฟื้นฟูสภาพ("ก่อสร้างใหม่", float(ทางเลือก.ต้นทุนก่อสร้าง), ปีดำเนินการ=20)
                    )
                    st.rerun()
        
        st.divider()
        
        # สรุปทางเลือกที่เปิดใช้งาน
        ทางเลือกที่ใช้ = [ท for ท in st.session_state.ทางเลือกทั้งหมด if ท.เปิดใช้งาน]
        st.info(f"📋 ทางเลือกที่เปิดใช้งาน: **{len(ทางเลือกที่ใช้)}** จาก {len(st.session_state.ทางเลือกทั้งหมด)} ทางเลือก")
    
    # ==========================================================================
    # แท็บ 2: ผลการวิเคราะห์
    # ==========================================================================
    with tab2:
        st.header("📊 ผลการวิเคราะห์ LCCA")
        
        # แสดงชื่อโครงการ
        st.subheader(f"📋 โครงการ: {st.session_state.ชื่อโครงการ}")
        
        # แสดงพารามิเตอร์
        col1, col2, col3 = st.columns(3)
        with col1:
            st.info(f"📅 ระยะวิเคราะห์: **{ระยะวิเคราะห์} ปี**")
        with col2:
            st.info(f"📉 อัตราคิดลด: **{อัตราคิดลด*100:.1f}%**")
        with col3:
            ทางเลือกที่ใช้ = [ท for ท in st.session_state.ทางเลือกทั้งหมด if ท.เปิดใช้งาน]
            st.info(f"🛣️ ทางเลือกที่วิเคราะห์: **{len(ทางเลือกที่ใช้)}**")
        
        if len(ทางเลือกที่ใช้) == 0:
            st.warning("⚠️ กรุณาเปิดใช้งานอย่างน้อย 1 ทางเลือกในแท็บ 'แก้ไขข้อมูล'")
        else:
            # ดำเนินการวิเคราะห์
            with st.spinner("⏳ กำลังคำนวณ LCCA..."):
                สรุป, กระแสเงินสด = วิเคราะห์_LCCA(st.session_state.ทางเลือกทั้งหมด, ระยะวิเคราะห์, อัตราคิดลด, รวมมูลค่าซาก)
            
            st.subheader("🏆 ตารางเปรียบเทียบ")
            
            # จัดรูปแบบตาราง
            สรุป_display = สรุป[['ลำดับ', 'ทางเลือก', 'ประเภทผิวทาง', 'ความหนา_ซม', 'พื้นที่_ตรม', 
                                'ต้นทุนก่อสร้าง_ตรม', 'มูลค่าปัจจุบันรวม', 
                                'ต้นทุนเฉลี่ยรายปี', 'ต้นทุนต่อตรม_ต่อปี']].copy()
            สรุป_display.columns = ['ลำดับ', 'ทางเลือก', 'ประเภท', 'ความหนา (ซม.)', 'พื้นที่ (ตร.ม.)',
                                   'ก่อสร้าง (บาท/ตร.ม.)', 'มูลค่าปัจจุบัน (บาท)', 
                                   'EAC (บาท/ปี)', 'ต้นทุน (บาท/ตร.ม./ปี)']
            
            def highlight_best(row):
                if row['ลำดับ'] == 1:
                    return ['background-color: #90EE90'] * len(row)
                return [''] * len(row)
            
            st.dataframe(
                สรุป_display.style.apply(highlight_best, axis=1).format({
                    'ความหนา (ซม.)': '{:,.1f}',
                    'พื้นที่ (ตร.ม.)': '{:,.0f}',
                    'ก่อสร้าง (บาท/ตร.ม.)': '{:,.0f}',
                    'มูลค่าปัจจุบัน (บาท)': '{:,.0f}',
                    'EAC (บาท/ปี)': '{:,.0f}',
                    'ต้นทุน (บาท/ตร.ม./ปี)': '{:,.2f}'
                }),
                use_container_width=True,
                hide_index=True
            )
            
            # แสดงผู้ชนะ
            ผู้ชนะ = สรุป.iloc[0]
            st.success(f"""
            ### ⭐ ทางเลือกที่ประหยัดที่สุด: **{ผู้ชนะ['ทางเลือก']}**
            - มูลค่าปัจจุบันรวม: **{ผู้ชนะ['มูลค่าปัจจุบันรวม']:,.0f} บาท**
            - ต้นทุนเฉลี่ยรายปี: **{ผู้ชนะ['ต้นทุนเฉลี่ยรายปี']:,.0f} บาท/ปี**
            - ต้นทุนต่อตารางเมตรต่อปี: **{ผู้ชนะ['ต้นทุนต่อตรม_ต่อปี']:,.2f} บาท/ตร.ม./ปี**
            """)
            
            # เปรียบเทียบกับทางเลือกอื่น
            if len(สรุป) > 1:
                st.subheader("💡 การเปรียบเทียบกับทางเลือกอื่น")
                for idx in range(1, len(สรุป)):
                    อื่น = สรุป.iloc[idx]
                    ส่วนต่าง = อื่น['มูลค่าปัจจุบันรวม'] - ผู้ชนะ['มูลค่าปัจจุบันรวม']
                    ร้อยละ = (ส่วนต่าง / อื่น['มูลค่าปัจจุบันรวม']) * 100
                    st.info(f"📊 vs {อื่น['ทางเลือก']}: ประหยัด **{ส่วนต่าง:,.0f} บาท** ({ร้อยละ:.1f}%)")
            
            st.divider()
            
            # กราฟแท่งเปรียบเทียบ
            st.subheader("📊 กราฟเปรียบเทียบองค์ประกอบต้นทุน")
            
            fig_bar = go.Figure()
            
            สีองค์ประกอบ = {
                'PW_ก่อสร้าง': ('#1f77b4', 'ก่อสร้างเริ่มต้น'),
                'PW_บำรุงรักษา': ('#ff7f0e', 'บำรุงรักษา'),
                'PW_ฟื้นฟูสภาพ': ('#2ca02c', 'ฟื้นฟูสภาพ')
            }
            
            for องค์ประกอบ, (สี, ชื่อ) in สีองค์ประกอบ.items():
                fig_bar.add_trace(go.Bar(
                    name=ชื่อ,
                    x=สรุป['ทางเลือก'],
                    y=สรุป[องค์ประกอบ],
                    marker_color=สี
                ))
            
            fig_bar.update_layout(
                barmode='stack',
                title='องค์ประกอบต้นทุน (มูลค่าปัจจุบัน)',
                yaxis_title='มูลค่าปัจจุบัน (บาท)',
                xaxis_title='ทางเลือกผิวทาง',
                legend_title='องค์ประกอบ',
                height=500
            )
            
            st.plotly_chart(fig_bar, use_container_width=True)
            
            # กราฟวงกลม
            st.subheader("🥧 สัดส่วนต้นทุนแต่ละทางเลือก")
            
            cols = st.columns(min(len(สรุป), 4))
            for idx, (_, row) in enumerate(สรุป.iterrows()):
                with cols[idx % len(cols)]:
                    ข้อมูลวงกลม = {
                        'องค์ประกอบ': ['ก่อสร้าง', 'บำรุงรักษา', 'ฟื้นฟูสภาพ'],
                        'มูลค่า': [row['PW_ก่อสร้าง'], row['PW_บำรุงรักษา'], row['PW_ฟื้นฟูสภาพ']]
                    }
                    fig_pie = px.pie(
                        ข้อมูลวงกลม, 
                        values='มูลค่า', 
                        names='องค์ประกอบ',
                        title=row['ทางเลือก'],
                        color_discrete_sequence=px.colors.qualitative.Set2
                    )
                    fig_pie.update_layout(height=350)
                    st.plotly_chart(fig_pie, use_container_width=True)
            
            # ปุ่มส่งออกรายงาน
            st.divider()
            st.subheader("📄 ส่งออกรายงาน")
            
            col_export1, col_export2 = st.columns(2)
            
            with col_export1:
                if DOCX_AVAILABLE:
                    word_file = สร้างรายงาน_Word(
                        สรุป, กระแสเงินสด, ระยะวิเคราะห์, อัตราคิดลด,
                        st.session_state.ทางเลือกทั้งหมด,
                        st.session_state.ชื่อโครงการ
                    )
                    st.download_button(
                        label="📝 สร้างรายงาน Word แบบย่อ",
                        data=word_file,
                        file_name=f"LCCA_{st.session_state.ชื่อโครงการ}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                else:
                    st.warning("⚠️ ต้องติดตั้ง python-docx: `pip install python-docx`")
            
            with col_export2:
                csv_summary = สรุป.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="📊 ดาวน์โหลดสรุป CSV",
                    data=csv_summary,
                    file_name=f"LCCA_Summary_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )

            # ─── รายงานแบบที่ปรึกษา ───
            st.divider()
            st.subheader("📑 รายงานแบบที่ปรึกษา")
            st.caption("รายงานครบถ้วนพร้อมบทเกริ่นนำ ข้อมูลโครงการ และบทสรุป")

            if DOCX_AVAILABLE:
                # input หมายเลขหัวข้อ + ปุ่มอยู่ในแถวเดียวกัน
                col_num, col_btn = st.columns([1, 2])

                with col_num:
                    หมายเลข = st.text_input(
                        "🔢 หมายเลขหัวข้อหลัก",
                        value=st.session_state.get('หมายเลขหัวข้อ', '4.8'),
                        key="section_number",
                        help="เช่น 4.8 → หัวข้อย่อย 4.8.1, 4.8.2, ... 4.8.9"
                    )
                    st.session_state['หมายเลขหัวข้อ'] = หมายเลข
                    st.caption(f"ย่อย: {หมายเลข}.1 ~ {หมายเลข}.9")

                with col_btn:
                    word_file_pro = สร้างรายงาน_Word_ที่ปรึกษา(
                        สรุป, กระแสเงินสด, ระยะวิเคราะห์, อัตราคิดลด,
                        st.session_state.ทางเลือกทั้งหมด,
                        st.session_state.ชื่อโครงการ,
                        ข้อมูลโครงการ={},
                        หมายเลขหัวข้อหลัก=หมายเลข
                    )
                    st.write("")  # padding เพื่อให้ปุ่มอยู่ระดับเดียวกัน
                    st.download_button(
                        label="📄 สร้างรายงานแบบที่ปรึกษา",
                        data=word_file_pro,
                        file_name=f"LCCA_ที่ปรึกษา_{st.session_state.ชื่อโครงการ}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
            else:
                st.warning("⚠️ ต้องติดตั้ง python-docx: `pip install python-docx`")
    
    # ==========================================================================
    # แท็บ 3: กระแสเงินสด
    # ==========================================================================
    with tab3:
        st.header("💰 ตารางกระแสเงินสดรายปี")
        
        ทางเลือกที่ใช้ = [ท for ท in st.session_state.ทางเลือกทั้งหมด if ท.เปิดใช้งาน]
        
        if len(ทางเลือกที่ใช้) == 0:
            st.warning("⚠️ กรุณาเปิดใช้งานอย่างน้อย 1 ทางเลือก")
        else:
            with st.spinner("⏳ กำลังคำนวณกระแสเงินสด..."):
                สรุป, กระแสเงินสด = วิเคราะห์_LCCA(st.session_state.ทางเลือกทั้งหมด, ระยะวิเคราะห์, อัตราคิดลด, รวมมูลค่าซาก)
            
            ทางเลือกที่เลือก = st.selectbox(
                "เลือกทางเลือกที่ต้องการดูรายละเอียด:",
                options=[ท.ชื่อ for ท in ทางเลือกที่ใช้]
            )
            
            if ทางเลือกที่เลือก in กระแสเงินสด:
                cf_table = กระแสเงินสด[ทางเลือกที่เลือก].copy()
                
                # แสดงสรุป
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("ต้นทุนตามปีรวม", f"{cf_table['ต้นทุนตามปี'].sum():,.0f} บาท")
                with col2:
                    st.metric("มูลค่าปัจจุบันรวม", f"{cf_table['มูลค่าปัจจุบัน'].sum():,.0f} บาท")
                with col3:
                    eac = คำนวณต้นทุนเฉลี่ยรายปี(cf_table['มูลค่าปัจจุบัน'].sum(), อัตราคิดลด, ระยะวิเคราะห์)
                    st.metric("EAC", f"{eac:,.0f} บาท/ปี")
                
                st.divider()
                
                # จัดรูปแบบตาราง
                cf_display = cf_table.copy()
                cf_display['ต้นทุนต่อหน่วย'] = cf_display['ต้นทุนต่อหน่วย'].apply(lambda x: f"{x:,.2f}")
                cf_display['ต้นทุนตามปี'] = cf_display['ต้นทุนตามปี'].apply(lambda x: f"{x:,.0f}")
                cf_display['ตัวคูณ_PW'] = cf_display['ตัวคูณ_PW'].apply(lambda x: f"{x:.4f}")
                cf_display['มูลค่าปัจจุบัน'] = cf_display['มูลค่าปัจจุบัน'].apply(lambda x: f"{x:,.0f}")
                cf_display.columns = ['ปี', 'กิจกรรม', 'ประเภท', 'ต้นทุน/หน่วย', 'ต้นทุนตามปี (บาท)', 'ตัวคูณ PW', 'มูลค่าปัจจุบัน (บาท)']
                
                st.dataframe(cf_display, use_container_width=True, hide_index=True, height=500)
                
                # กราฟ Timeline
                st.subheader("📅 Timeline กระแสเงินสด")
                
                cf_plot = cf_table[cf_table['ต้นทุนตามปี'] > 0].copy()
                
                fig_timeline = px.scatter(
                    cf_plot,
                    x='ปี',
                    y='มูลค่าปัจจุบัน',
                    size='ต้นทุนตามปี',
                    color='ประเภท',
                    hover_name='กิจกรรม',
                    title=f'Timeline กระแสเงินสด - {ทางเลือกที่เลือก}',
                    labels={'ปี': 'ปี', 'มูลค่าปัจจุบัน': 'มูลค่าปัจจุบัน (บาท)'}
                )
                fig_timeline.update_layout(height=400)
                st.plotly_chart(fig_timeline, use_container_width=True)
                
                # ดาวน์โหลด CSV
                st.divider()
                csv = cf_table.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="⬇️ ดาวน์โหลดตารางกระแสเงินสด (CSV)",
                    data=csv,
                    file_name=f"cashflow_{ทางเลือกที่เลือก}.csv",
                    mime="text/csv"
                )
    
    # ==========================================================================
    # แท็บ 4: การวิเคราะห์ความไว
    # ==========================================================================
    with tab4:
        st.header("📈 การวิเคราะห์ความไว (Sensitivity Analysis)")
        
        ทางเลือกที่ใช้ = [ท for ท in st.session_state.ทางเลือกทั้งหมด if ท.เปิดใช้งาน]
        
        if len(ทางเลือกที่ใช้) == 0:
            st.warning("⚠️ กรุณาเปิดใช้งานอย่างน้อย 1 ทางเลือก")
        else:
            st.subheader("1️⃣ ความไวต่ออัตราคิดลด")
            
            with st.spinner("⏳ กำลังวิเคราะห์ความไว..."):
                ผลอัตราคิดลด, pivot_อัตราคิดลด = วิเคราะห์ความไว_อัตราคิดลด(
                    st.session_state.ทางเลือกทั้งหมด, ระยะวิเคราะห์, อัตราคิดลด, ช่วงอัตราคิดลด, รวมมูลค่าซาก
                )
            
            if len(ผลอัตราคิดลด) > 0:
                # กราฟเส้น
                fig_sens = px.line(
                    ผลอัตราคิดลด,
                    x='อัตราคิดลด',
                    y='มูลค่าปัจจุบัน',
                    color='ทางเลือก',
                    markers=True,
                    title='ผลกระทบของอัตราคิดลดต่อมูลค่าปัจจุบัน',
                    labels={'อัตราคิดลด': 'อัตราคิดลด', 'มูลค่าปัจจุบัน': 'มูลค่าปัจจุบัน (บาท)'}
                )
                fig_sens.update_layout(height=500)
                fig_sens.update_xaxes(tickformat='.1%')
                st.plotly_chart(fig_sens, use_container_width=True)
                
                # ตาราง Pivot
                st.markdown("**ตารางสรุปมูลค่าปัจจุบันตามอัตราคิดลด (บาท):**")
                pivot_display = pivot_อัตราคิดลด.copy()
                for col in pivot_display.columns:
                    pivot_display[col] = pivot_display[col].apply(lambda x: f"{x:,.0f}")
                st.dataframe(pivot_display, use_container_width=True)
                
                # วิเคราะห์
                อัตราต่ำสุด = ผลอัตราคิดลด['อัตราคิดลด'].min()
                อัตราสูงสุด = ผลอัตราคิดลด['อัตราคิดลด'].max()
                
                ผู้ชนะต่ำ = ผลอัตราคิดลด[ผลอัตราคิดลด['อัตราคิดลด'] == อัตราต่ำสุด].nsmallest(1, 'มูลค่าปัจจุบัน')['ทางเลือก'].values[0]
                ผู้ชนะสูง = ผลอัตราคิดลด[ผลอัตราคิดลด['อัตราคิดลด'] == อัตราสูงสุด].nsmallest(1, 'มูลค่าปัจจุบัน')['ทางเลือก'].values[0]
                
                if ผู้ชนะต่ำ == ผู้ชนะสูง:
                    st.success(f"✅ **{ผู้ชนะต่ำ}** เป็นทางเลือกที่ประหยัดที่สุดในทุกอัตราคิดลด (Robust Decision)")
                else:
                    st.warning(f"⚠️ ทางเลือกที่ดีที่สุดเปลี่ยนแปลง: {ผู้ชนะต่ำ} (อัตราต่ำ) vs {ผู้ชนะสูง} (อัตราสูง)")
    
    # ==========================================================================
    # แท็บ 5: ทฤษฎี LCCA
    # ==========================================================================
    with tab5:
        st.header("ℹ️ ทฤษฎี Life-Cycle Cost Analysis (LCCA)")
        
        st.markdown("""
        ## 1. ประเภทผิวทางคอนกรีต
        
        | ประเภท | ชื่อเต็ม | ลักษณะเด่น |
        |--------|---------|-----------|
        | **JPCP** | Jointed Plain Concrete Pavement | คอนกรีตไม่เสริมเหล็ก มีรอยต่อทุก 4-6 ม. |
        | **JRCP** | Jointed Reinforced Concrete Pavement | คอนกรีตเสริมเหล็ก รอยต่อห่าง 8-15 ม. |
        | **CRCP** | Continuously Reinforced Concrete Pavement | เสริมเหล็กต่อเนื่อง ไม่มีรอยต่อตามขวาง |
        
        ## 2. สูตรคำนวณหลัก
        
        ### 2.1 มูลค่าปัจจุบัน (Present Worth)
        """)
        
        st.latex(r"PW = FV \times (1 + i)^{-n}")
        
        st.markdown("""
        ### 2.2 ต้นทุนเฉลี่ยรายปี (Equivalent Annual Cost)
        """)
        
        st.latex(r"EAC = PW \times \frac{i(1+i)^N}{(1+i)^N - 1}")
        
        st.markdown("""
        ## 3. เปรียบเทียบผิวทางคอนกรีต
        
        | เกณฑ์ | JPCP | JRCP | CRCP |
        |------|------|------|------|
        | ต้นทุนก่อสร้าง | ต่ำ | ปานกลาง | สูง |
        | ระยะห่างรอยต่อ | 4-6 ม. | 8-15 ม. | ไม่มี |
        | เหล็กเสริม | ไม่มี | 0.1-0.25% | 0.6-0.7% |
        | ค่าบำรุงรักษา | สูง | ปานกลาง | ต่ำ |
        | อายุใช้งาน | 20-30 ปี | 25-35 ปี | 30-40 ปี |
        
        ## 4. เอกสารอ้างอิง
        
        - FHWA-SA-98-079: Life-Cycle Cost Analysis in Pavement Design
        - AASHTO Guide for Design of Pavement Structures
        - NCHRP Report 703: Guide for Pavement-Type Selection
        - มาตรฐานกรมทางหลวง
        """)
    
    # ==========================================================================
    # Footer
    # ==========================================================================
    st.divider()
    st.markdown("""
    ---
    **โปรแกรมวิเคราะห์ LCCA ผิวทาง v2.2** | พัฒนาสำหรับการเรียนการสอนด้านวิศวกรรมทาง  
    ภาควิชาครุศาสตร์โยธา มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าพระนครเหนือ
    """)


if __name__ == "__main__":
    main()
