"""
ESAL Calculator - AASHTO 1993 (Version 3.0)
โปรแกรมคำนวณปริมาณเพลาเดี่ยวมาตรฐานเทียบเท่า (Equivalent Single Axle Load)
สำหรับผิวทาง Rigid Pavement และ Flexible Pavement
ตามมาตรฐาน AASHTO Guide for Design of Pavement Structures (1993)

Features V3:
- รองรับ Export Excel และ Word ในรูปแบบมาตรฐาน
- Save/Load Project สำหรับแก้ไขภายหลัง
- คำนวณ ACC. ESAL (สะสม)
- Export Word รวม Flexible + Rigid ในไฟล์เดียว (ตามรูปแบบรายงานมาตรฐาน)
- ระบบหมายเลขหัวข้อ/ตาราง แบบ Auto-increment
- บทเกริ่นนำแยก Flexible / Rigid พร้อม Preview
- Font TH SarabunPSK 15pt

พัฒนาโดย: รศ.ดร.อิทธิพล มีผล ภาควิชาครุศาสตร์โยธา มจพ.
"""

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import json
import re
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

# ============================================================
# ข้อมูลรถบรรทุก 6 ชนิดตามกรมทางหลวงประเทศไทย
# ============================================================
TRUCKS = {
    'MB': {'desc': 'Medium Bus (รถโดยสารขนาดกลาง)', 'axles': [{'name': 'เพลาหน้า', 'load_ton': 4.0, 'type': 'Single'}, {'name': 'เพลาหลัง', 'load_ton': 11.0, 'type': 'Tandem'}]},
    'HB': {'desc': 'Heavy Bus (รถโดยสารขนาดใหญ่)', 'axles': [{'name': 'เพลาหน้า', 'load_ton': 5.0, 'type': 'Single'}, {'name': 'เพลาหลัง', 'load_ton': 20.0, 'type': 'Tandem'}]},
    'MT': {'desc': 'Medium Truck (รถบรรทุกขนาดกลาง)', 'axles': [{'name': 'เพลาหน้า', 'load_ton': 4.0, 'type': 'Single'}, {'name': 'เพลาหลัง', 'load_ton': 11.0, 'type': 'Single'}]},
    'HT': {'desc': 'Heavy Truck (รถบรรทุกขนาดใหญ่)', 'axles': [{'name': 'เพลาหน้า', 'load_ton': 5.0, 'type': 'Single'}, {'name': 'เพลาหลัง', 'load_ton': 20.0, 'type': 'Tandem'}]},
    'TR': {'desc': 'Full Trailer (รถพ่วง)', 'axles': [{'name': 'เพลาหน้า', 'load_ton': 5.0, 'type': 'Single'}, {'name': 'เพลาหลัง', 'load_ton': 20, 'type': 'Tandem'}, {'name': 'เพลาพ่วงหน้า', 'load_ton': 11, 'type': 'Single'}, {'name': 'เพลาพ่วงหลัง', 'load_ton': 11, 'type': 'Single'}]},
    'STR': {'desc': 'Semi-Trailer (รถกึ่งพ่วง)', 'axles': [{'name': 'เพลาหน้า', 'load_ton': 5.0, 'type': 'Single'}, {'name': 'เพลาหลัง', 'load_ton': 20.0, 'type': 'Tandem'}, {'name': 'เพลาพ่วงหลัง', 'load_ton': 20.0, 'type': 'Tandem'}]}
}

# ============================================================
# ตาราง Truck Factor คำนวณตาม AASHTO 1993
# ============================================================

# Rigid Pavement - pt = 2.0
TRUCK_FACTORS_RIGID_PT20 = {
    'MB':  {10: 3.718199, 11: 3.742581, 12: 3.754803, 13: 3.760977, 14: 3.764184, 15: 3.765855, 16: 3.766727},
    'HB':  {10: 6.125043, 11: 6.204343, 12: 6.247170, 13: 6.269632, 14: 6.281529, 15: 6.287867, 16: 6.291257},
    'MT':  {10: 3.718199, 11: 3.742581, 12: 3.754803, 13: 3.760977, 14: 3.764184, 15: 3.765855, 16: 3.766727},
    'HT':  {10: 6.125043, 11: 6.204343, 12: 6.247170, 13: 6.269632, 14: 6.281529, 15: 6.287867, 16: 6.291257},
    'TR':  {10: 13.466316, 11: 13.594592, 12: 13.661961, 13: 13.696817, 14: 13.715152, 15: 13.724934, 16: 13.730167},
    'STR': {10: 12.128867, 11: 12.287718, 12: 12.373488, 13: 12.418469, 14: 12.442292, 15: 12.454956, 16: 12.461738}
}

# Rigid Pavement - pt = 2.5
TRUCK_FACTORS_RIGID_PT25 = {
    'MB':  {10: 3.657799, 11: 3.711341, 12: 3.738346, 13: 3.752027, 14: 3.759145, 15: 3.762869, 16: 3.764817},
    'HB':  {10: 5.921064, 11: 6.092776, 12: 6.186668, 13: 6.236237, 14: 6.262582, 15: 6.276617, 16: 6.284134},
    'MT':  {10: 3.657799, 11: 3.711341, 12: 3.738346, 13: 3.752027, 14: 3.759145, 15: 3.762869, 16: 3.764817},
    'HT':  {10: 5.921064, 11: 6.092776, 12: 6.186668, 13: 6.236237, 14: 6.262582, 15: 6.276617, 16: 6.284134},
    'TR':  {10: 13.141034, 11: 13.420301, 12: 13.568419, 13: 13.645455, 14: 13.686091, 15: 13.707787, 16: 13.719438},
    'STR': {10: 11.720309, 11: 12.064293, 12: 12.252335, 13: 12.351598, 14: 12.404353, 15: 12.432524, 16: 12.447620}
}

# Rigid Pavement - pt = 3.0
TRUCK_FACTORS_RIGID_PT30 = {
    'MB':  {10: 3.581408, 11: 3.671458, 12: 3.717236, 13: 3.740520, 14: 3.752660, 15: 3.759033, 16: 3.762385},
    'HB':  {10: 5.668347, 11: 5.951971, 12: 6.109552, 13: 6.193451, 14: 6.238241, 15: 6.262146, 16: 6.274979},
    'MT':  {10: 3.581408, 11: 3.671458, 12: 3.717236, 13: 3.740520, 14: 3.752660, 15: 3.759033, 16: 3.762385},
    'HT':  {10: 5.668347, 11: 5.951971, 12: 6.109552, 13: 6.193451, 14: 6.238241, 15: 6.262146, 16: 6.274979},
    'TR':  {10: 12.734883, 11: 13.199416, 12: 13.448924, 13: 13.579571, 14: 13.648731, 15: 13.685766, 16: 13.705646},
    'STR': {10: 11.214096, 11: 11.782308, 12: 12.097912, 13: 12.265925, 14: 12.355613, 15: 12.403556, 16: 12.429280}
}

# Flexible Pavement - pt = 2.0
TRUCK_FACTORS_FLEX_PT20 = {
    'MB':  {4: 3.529011, 5: 3.598168, 6: 3.719257, 7: 3.810681, 8: 3.874256, 9: 3.916863},
    'HB':  {4: 3.332846, 5: 3.384895, 6: 3.458092, 7: 3.508785, 8: 3.541983, 9: 3.562854},
    'MT':  {4: 3.529011, 5: 3.598168, 6: 3.719257, 7: 3.810681, 8: 3.874256, 9: 3.916863},
    'HT':  {4: 3.332846, 5: 3.384895, 6: 3.458092, 7: 3.508785, 8: 3.541983, 9: 3.562854},
    'TR':  {4: 10.291092, 5: 10.488813, 6: 10.808050, 7: 11.043444, 8: 11.203523, 9: 11.310117},
    'STR': {4: 6.537851, 5: 6.649420, 6: 6.800056, 7: 6.903531, 8: 6.971366, 9: 7.014261}
}

# Flexible Pavement - pt = 2.5
TRUCK_FACTORS_FLEX_PT25 = {
    'MB':  {4: 3.069453, 5: 3.203842, 6: 3.451114, 7: 3.645241, 8: 3.779066, 9: 3.869188},
    'HB':  {4: 3.053625, 5: 3.157524, 6: 3.311765, 7: 3.421800, 8: 3.494667, 9: 3.541837},
    'MT':  {4: 3.069453, 5: 3.203842, 6: 3.451114, 7: 3.645241, 8: 3.779066, 9: 3.869188},
    'HT':  {4: 3.053625, 5: 3.157524, 6: 3.311765, 7: 3.421800, 8: 3.494667, 9: 3.541837},
    'TR':  {4: 9.069826, 5: 9.462000, 6: 10.120276, 7: 10.622935, 8: 10.967259, 9: 11.196528},
    'STR': {4: 5.955718, 5: 6.182789, 6: 6.501567, 7: 6.726542, 8: 6.874756, 9: 6.970223}
}

# Flexible Pavement - pt = 3.0
TRUCK_FACTORS_FLEX_PT30 = {
    'MB':  {4: 2.552540, 5: 2.742623, 6: 3.120508, 7: 3.433469, 8: 3.657896, 9: 3.812063},
    'HB':  {4: 2.728486, 5: 2.879854, 6: 3.125499, 7: 3.308196, 8: 3.432738, 9: 3.513580},
    'MT':  {4: 2.552540, 5: 2.742623, 6: 3.120508, 7: 3.433469, 8: 3.657896, 9: 3.812063},
    'HT':  {4: 2.728486, 5: 2.879854, 6: 3.125499, 7: 3.308196, 8: 3.432738, 9: 3.513580},
    'TR':  {4: 7.671306, 5: 8.245291, 6: 9.265343, 7: 10.082089, 8: 10.658949, 9: 11.046207},
    'STR': {4: 5.266321, 5: 5.609502, 6: 6.120685, 7: 6.495126, 8: 6.750547, 9: 6.915832}
}


def get_default_truck_factor(truck_code, pavement_type, pt, param):
    """ดึงค่า Truck Factor เริ่มต้นจากตาราง"""
    if pavement_type == 'rigid':
        if pt == 2.0:
            return TRUCK_FACTORS_RIGID_PT20[truck_code][param]
        elif pt == 2.5:
            return TRUCK_FACTORS_RIGID_PT25[truck_code][param]
        else:
            return TRUCK_FACTORS_RIGID_PT30[truck_code][param]
    else:
        if pt == 2.0:
            return TRUCK_FACTORS_FLEX_PT20[truck_code][param]
        elif pt == 2.5:
            return TRUCK_FACTORS_FLEX_PT25[truck_code][param]
        else:
            return TRUCK_FACTORS_FLEX_PT30[truck_code][param]


def calculate_esal_with_acc(traffic_df, truck_factors, lane_factor=0.5, direction_factor=1.0):
    """คำนวณ ESAL และ Accumulated ESAL จากข้อมูลปริมาณจราจร"""
    results = []
    acc_esal = 0
    
    for idx, row in traffic_df.iterrows():
        year = row.get('Year', idx + 1)
        year_data = {'Year': int(year)}
        
        total_aadt = 0
        for code in TRUCKS.keys():
            if code in traffic_df.columns:
                aadt = int(row[code])
                year_data[code] = aadt
                total_aadt += aadt
        
        year_data['AADT'] = total_aadt
        
        year_esal = 0
        for code in TRUCKS.keys():
            if code in traffic_df.columns:
                aadt = row[code]
                tf = truck_factors[code]
                esal = aadt * tf * lane_factor * direction_factor * 365
                year_esal += esal
        
        year_data['ESAL'] = int(round(year_esal))
        acc_esal += year_esal
        year_data['ACC_ESAL'] = int(round(acc_esal))
        
        results.append(year_data)
    
    return pd.DataFrame(results), int(round(acc_esal))


def create_template():
    """สร้าง Template Excel สำหรับอัพโหลดข้อมูล"""
    base = {'MB': 120, 'HB': 60, 'MT': 250, 'HT': 180, 'TR': 100, 'STR': 120}
    growth_rate = 1.045
    
    data = {'Year': list(range(1, 21))}
    for code in base.keys():
        data[code] = [int(base[code] * (growth_rate ** i)) for i in range(20)]
    
    return pd.DataFrame(data)


def create_excel_report(results_df, pavement_type, pt, param, lane_factor, direction_factor, 
                       total_esal, truck_factors, num_years):
    """สร้างรายงาน Excel ในรูปแบบมาตรฐาน"""
    wb = Workbook()
    ws = wb.active
    ws.title = "ESAL Report"
    
    # Styles
    header_font = Font(bold=True, size=14)
    title_font = Font(bold=True, size=16)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    header_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
    center_align = Alignment(horizontal='center', vertical='center')
    right_align = Alignment(horizontal='right', vertical='center')
    
    pavement_text = "Rigid Pavement" if pavement_type == 'rigid' else "Flexible Pavement"
    ws.merge_cells('A1:I1')
    ws['A1'] = f"ปริมาณเพลามาตรฐาน (ESALs) ระยะเวลาออกแบบ {num_years} ปี"
    ws['A1'].font = title_font
    ws['A1'].alignment = center_align
    
    ws.merge_cells('A2:I2')
    ws['A2'] = f"ผิวทางแบบ{'แข็ง' if pavement_type == 'rigid' else 'ยืดหยุ่น'} ({pavement_text})"
    ws['A2'].font = header_font
    ws['A2'].alignment = center_align
    
    param_label = f"D = {param}" if pavement_type == 'rigid' else f"SN = {param}"
    params_data = [
        ('รายการ', 'ค่า'),
        ('ประเภทผิวทาง', pavement_text),
        ('pt', str(pt)),
        ('พารามิเตอร์', param_label),
        ('Lane Factor', str(lane_factor)),
        ('Direction Factor', str(direction_factor)),
        ('ESAL รวม', f"{total_esal:,}"),
        ('จำนวนปี', str(num_years))
    ]
    
    for i, (label, value) in enumerate(params_data):
        row = 4 + i
        ws[f'A{row}'] = label
        ws[f'B{row}'] = value
        ws[f'A{row}'].border = border
        ws[f'B{row}'].border = border
        if i == 0:
            ws[f'A{row}'].fill = header_fill
            ws[f'B{row}'].fill = header_fill
            ws[f'A{row}'].font = Font(bold=True)
            ws[f'B{row}'].font = Font(bold=True)
    
    ws['D4'] = 'รหัส'
    ws['E4'] = 'ประเภท'
    ws['F4'] = 'Truck Factor'
    for col in ['D', 'E', 'F']:
        ws[f'{col}4'].fill = header_fill
        ws[f'{col}4'].font = Font(bold=True)
        ws[f'{col}4'].border = border
        ws[f'{col}4'].alignment = center_align
    
    for i, code in enumerate(TRUCKS.keys()):
        row = 5 + i
        ws[f'D{row}'] = code
        ws[f'E{row}'] = TRUCKS[code]['desc']
        ws[f'F{row}'] = f"{truck_factors[code]:.4f}"
        ws[f'D{row}'].border = border
        ws[f'E{row}'].border = border
        ws[f'F{row}'].border = border
        ws[f'D{row}'].alignment = center_align
        ws[f'F{row}'].alignment = right_align
    
    start_row = 14
    ws[f'I{start_row-1}'] = 'แสดงปริมาณสะสม'
    ws[f'I{start_row-1}'].font = Font(italic=True, size=9)
    
    headers = ['Year', 'MB', 'HB', 'MT', 'HT', 'TR', 'STR', 'AADT', 'ESAL', 'ACC. ESAL']
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=start_row, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = Font(bold=True)
        cell.border = border
        cell.alignment = center_align
    
    for row_idx, row_data in results_df.iterrows():
        excel_row = start_row + 1 + row_idx
        for col_idx, header in enumerate(headers, 1):
            if header == 'ACC. ESAL':
                value = row_data.get('ACC_ESAL', 0)
            else:
                value = row_data.get(header, 0)
            
            cell = ws.cell(row=excel_row, column=col_idx, value=value)
            cell.border = border
            
            if header in ['ESAL', 'ACC. ESAL', 'AADT']:
                cell.number_format = '#,##0'
                cell.alignment = right_align
            elif header == 'Year':
                cell.alignment = center_align
            else:
                cell.alignment = right_align
    
    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 3
    ws.column_dimensions['D'].width = 8
    ws.column_dimensions['E'].width = 35
    ws.column_dimensions['F'].width = 14
    for col in ['G', 'H', 'I', 'J']:
        ws.column_dimensions[col].width = 14
    
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ============================================================
# Auto-increment table number helper
# ============================================================
def increment_table_number(base_number, offset):
    """
    เพิ่มเลขตารางอัตโนมัติ เช่น "4-1" + offset=1 → "4-2", "4-1" + offset=2 → "4-3"
    รองรับรูปแบบ: "4-1", "4.1", "1"
    """
    # รูปแบบ "X-Y" เช่น "4-1"
    match = re.match(r'^(\d+)-(\d+)$', base_number.strip())
    if match:
        prefix = match.group(1)
        num = int(match.group(2))
        return f"{prefix}-{num + offset}"
    
    # รูปแบบ "X.Y" เช่น "4.1"
    match = re.match(r'^(\d+)\.(\d+)$', base_number.strip())
    if match:
        prefix = match.group(1)
        num = int(match.group(2))
        return f"{prefix}.{num + offset}"
    
    # รูปแบบตัวเลขเดียว เช่น "1"
    match = re.match(r'^(\d+)$', base_number.strip())
    if match:
        num = int(match.group(1))
        return str(num + offset)
    
    return f"{base_number}+{offset}"


# ============================================================
# Word Report Generation (python-docx)
# ============================================================
def create_word_report_single(results_df, pavement_type, pt, param, lane_factor, direction_factor,
                              total_esal, truck_factors, num_years, report_settings=None):
    """สร้างรายงาน Word สำหรับผิวทางประเภทเดียว (Flexible หรือ Rigid)"""
    try:
        from docx import Document
        from docx.shared import Pt, Cm, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.enum.table import WD_TABLE_ALIGNMENT
        from docx.oxml.ns import nsdecls, qn
        from docx.oxml import parse_xml, OxmlElement
    except ImportError:
        return None
    
    doc = Document()
    FONT_NAME = 'TH SarabunPSK'
    FONT_SIZE = 15
    TABLE_FONT_SIZE = 14
    
    # ตั้งค่า Normal style
    style = doc.styles['Normal']
    style.font.name = FONT_NAME
    style.font.size = Pt(FONT_SIZE)
    style._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_NAME)
    
    # ตั้งค่าหน้ากระดาษ A4
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    
    # กำหนดค่า default report settings
    if report_settings is None:
        report_settings = {}
    
    if pavement_type == 'flexible':
        section_num = report_settings.get('flex_section_number', '4.2.2')
        table_start = report_settings.get('flex_table_start', '4-1')
    else:
        section_num = report_settings.get('rigid_section_number', '4.2.3')
        table_start = report_settings.get('rigid_table_start', '4-4')
    
    tbl_param = table_start
    tbl_tf = increment_table_number(table_start, 1)
    tbl_esal = increment_table_number(table_start, 2)
    
    pavement_text = "Rigid Pavement" if pavement_type == 'rigid' else "Flexible Pavement"
    pavement_thai = "แบบแข็ง" if pavement_type == 'rigid' else "ยืดหยุ่น"
    param_label = f"D = {param} นิ้ว" if pavement_type == 'rigid' else f"SN = {param}"
    
    _build_section(doc, pavement_type, pavement_text, pavement_thai, section_num,
                   num_years, param_label, pt, lane_factor, direction_factor,
                   total_esal, truck_factors, results_df,
                   tbl_param, tbl_tf, tbl_esal,
                   FONT_NAME, FONT_SIZE, TABLE_FONT_SIZE)
    
    # Footer
    footer_para = doc.add_paragraph()
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer_para.add_run("พัฒนาเพื่อการเรียนการสอนโดย รศ.ดร.อิทธิพล มีผล ภาควิชาครุศาสตร์โยธา มจพ.")
    run.font.name = FONT_NAME
    run.font.size = Pt(14)
    run.italic = True
    run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_NAME)
    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output


def create_word_report_combined(traffic_df, flex_params, rigid_params, report_settings=None):
    """
    สร้างรายงาน Word รวม Flexible + Rigid ในไฟล์เดียว
    flex_params / rigid_params = dict with keys: pt, param, lane_factor, direction_factor, truck_factors
    """
    try:
        from docx import Document
        from docx.shared import Pt, Cm, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.enum.table import WD_TABLE_ALIGNMENT
        from docx.oxml.ns import nsdecls, qn
        from docx.oxml import parse_xml, OxmlElement
    except ImportError:
        return None
    
    doc = Document()
    FONT_NAME = 'TH SarabunPSK'
    FONT_SIZE = 15
    TABLE_FONT_SIZE = 14
    
    # ตั้งค่า Normal style
    style = doc.styles['Normal']
    style.font.name = FONT_NAME
    style.font.size = Pt(FONT_SIZE)
    style._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_NAME)
    
    # ตั้งค่าหน้ากระดาษ A4
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    
    if report_settings is None:
        report_settings = {}
    
    num_years = len(traffic_df)
    
    # ===== ส่วน Flexible Pavement =====
    fp = flex_params
    flex_results, flex_total = calculate_esal_with_acc(
        traffic_df, fp['truck_factors'], fp['lane_factor'], fp['direction_factor']
    )
    
    flex_section_num = report_settings.get('flex_section_number', '4.2.2')
    flex_table_start = report_settings.get('flex_table_start', '4-1')
    flex_tbl_param = flex_table_start
    flex_tbl_tf = increment_table_number(flex_table_start, 1)
    flex_tbl_esal = increment_table_number(flex_table_start, 2)
    
    flex_param_label = f"SN = {fp['param']}"
    
    _build_section(doc, 'flexible', 'Flexible Pavement', 'ยืดหยุ่น', flex_section_num,
                   num_years, flex_param_label, fp['pt'], fp['lane_factor'], fp['direction_factor'],
                   flex_total, fp['truck_factors'], flex_results,
                   flex_tbl_param, flex_tbl_tf, flex_tbl_esal,
                   FONT_NAME, FONT_SIZE, TABLE_FONT_SIZE)
    
    # Page break
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn as qn2
    p = doc.add_paragraph()
    run = p.add_run()
    br = OxmlElement('w:br')
    br.set(qn2('w:type'), 'page')
    run._element.append(br)
    
    # ===== ส่วน Rigid Pavement =====
    rp = rigid_params
    rigid_results, rigid_total = calculate_esal_with_acc(
        traffic_df, rp['truck_factors'], rp['lane_factor'], rp['direction_factor']
    )
    
    rigid_section_num = report_settings.get('rigid_section_number', '4.2.3')
    rigid_table_start = report_settings.get('rigid_table_start', '4-4')
    rigid_tbl_param = rigid_table_start
    rigid_tbl_tf = increment_table_number(rigid_table_start, 1)
    rigid_tbl_esal = increment_table_number(rigid_table_start, 2)
    
    rigid_param_label = f"D = {rp['param']} นิ้ว"
    
    _build_section(doc, 'rigid', 'Rigid Pavement', 'แบบแข็ง', rigid_section_num,
                   num_years, rigid_param_label, rp['pt'], rp['lane_factor'], rp['direction_factor'],
                   rigid_total, rp['truck_factors'], rigid_results,
                   rigid_tbl_param, rigid_tbl_tf, rigid_tbl_esal,
                   FONT_NAME, FONT_SIZE, TABLE_FONT_SIZE)
    
    # Footer
    footer_para = doc.add_paragraph()
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer_para.add_run("พัฒนาเพื่อการเรียนการสอนโดย รศ.ดร.อิทธิพล มีผล ภาควิชาครุศาสตร์โยธา มจพ.")
    run.font.name = FONT_NAME
    run.font.size = Pt(14)
    run.italic = True
    run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_NAME)
    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output


def _build_section(doc, pavement_type, pavement_text, pavement_thai, section_num,
                   num_years, param_label, pt, lane_factor, direction_factor,
                   total_esal, truck_factors, results_df,
                   tbl_param, tbl_tf, tbl_esal,
                   FONT_NAME, FONT_SIZE, TABLE_FONT_SIZE):
    """สร้างเนื้อหาหนึ่ง section (Flexible หรือ Rigid) ใน Word document"""
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn, nsdecls
    from docx.oxml import OxmlElement, parse_xml
    
    def set_run(run, font_name=FONT_NAME, font_size=FONT_SIZE, bold=False):
        run.font.name = font_name
        run.font.size = Pt(font_size)
        run.bold = bold
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    
    def set_cell_font(cell, font_name=FONT_NAME, font_size=TABLE_FONT_SIZE, bold=False, align=None):
        for paragraph in cell.paragraphs:
            if align:
                paragraph.alignment = align
            for run in paragraph.runs:
                run.font.name = font_name
                run.font.size = Pt(font_size)
                run.bold = bold
                run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    
    def set_cell_shading(cell, color="D9E2F3"):
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}"/>')
        cell._tc.get_or_add_tcPr().append(shading)
    
    def set_thai_distribute(paragraph):
        pPr = paragraph._element.get_or_add_pPr()
        jc = OxmlElement('w:jc')
        jc.set(qn('w:val'), 'thaiDistribute')
        pPr.append(jc)
    
    # ===== 1. หัวข้อ =====
    heading_para = doc.add_paragraph()
    heading_para.paragraph_format.space_after = Pt(6)
    
    # เลขหัวข้อ (tab) ชื่อหัวข้อ
    run = heading_para.add_run(f"{section_num}\t")
    set_run(run, font_size=FONT_SIZE, bold=True)
    
    run = heading_para.add_run(f"ปริมาณเพลามาตรฐาน (ESALs) ระยะเวลาออกแบบ {num_years} ปี ผิวทาง {pavement_text}")
    set_run(run, font_size=FONT_SIZE, bold=True)
    
    # ===== 2. บทเกริ่นนำ =====
    intro_para = doc.add_paragraph()
    intro_para.paragraph_format.first_line_indent = Cm(1.25)
    intro_para.paragraph_format.space_after = Pt(6)
    set_thai_distribute(intro_para)
    
    if pavement_type == 'flexible':
        intro_parts = [
            ("ในการคำนวณปริมาณเพลามาตรฐาน สำหรับผิวทางยืดหยุ่น ที่ปรึกษาได้กำหนดค่าพารามิเตอร์ต่าง ๆ และค่า ", False),
            ("Truck Factor ", False),
            ("ของรถบรรทุกหนัก ที่ใช้สำหรับการคำนวณ ดังแสดงในตารางที่ ", False),
            (f"{tbl_param}", False),
            (" และ ", False),
            (f"{tbl_tf}", False),
            (f" ดังนั้นค่าปริมาณเพลามาตรฐาน สำหรับผิวทางยืดหยุ่น ที่ระยะเวลาออกแบบ {num_years} ปี แสดงดังตารางที่ ", False),
            (f"{tbl_esal}", False),
        ]
    else:
        intro_parts = [
            ("ในการคำนวณปริมาณเพลามาตรฐานสำหรับผิวทางแบบแข็งหรือผิวทางคอนกรีต โดยที่ปรึกษาได้กำหนดค่าพารามิเตอร์ต่าง ๆ และค่า ", False),
            ("Truck Factor ", False),
            ("ของรถบรรทุกหนัก ที่ใช้สำหรับการคำนวณ ดังแสดงในตารางที่ ", False),
            (f"{tbl_param}", False),
            (" และ ", False),
            (f"{tbl_tf}", False),
            (f" ดังนั้นค่าปริมาณเพลามาตรฐาน สำหรับผิวทางแบบแข็ง ที่ระยะเวลาออกแบบ {num_years} ปี แสดงดังตารางที่ ", False),
            (f"{tbl_esal}", False),
        ]
    
    for text, is_bold in intro_parts:
        run = intro_para.add_run(text)
        set_run(run, font_size=FONT_SIZE, bold=is_bold)
    
    doc.add_paragraph()  # blank line
    
    # ===== 3. ตารางที่ X-1: ค่าพารามิเตอร์ =====
    cap1 = doc.add_paragraph()
    cap1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap1.paragraph_format.space_after = Pt(3)
    
    run = cap1.add_run(f"ตารางที่ {tbl_param} ")
    set_run(run, font_size=FONT_SIZE, bold=True)
    run = cap1.add_run("ค่าพารามิเตอร์ต่าง ๆ ที่ใช้สำหรับการคำนวณ")
    set_run(run, font_size=FONT_SIZE, bold=False)
    
    param_data = [
        ('รายการ', 'ค่า'),
        ('ประเภทผิวทาง', pavement_text),
        ('pt', str(pt)),
        ('พารามิเตอร์', param_label),
        ('Lane Factor', str(lane_factor)),
        ('Direction Factor', str(direction_factor)),
        ('ESAL รวม', f"{total_esal:,}"),
        ('จำนวนปี', str(num_years))
    ]
    
    param_table = doc.add_table(rows=len(param_data), cols=2)
    param_table.style = 'Table Grid'
    param_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    for i, (label, value) in enumerate(param_data):
        row = param_table.rows[i]
        row.cells[0].text = label
        row.cells[1].text = value
        is_header = (i == 0)
        set_cell_font(row.cells[0], font_size=TABLE_FONT_SIZE, bold=is_header)
        set_cell_font(row.cells[1], font_size=TABLE_FONT_SIZE, bold=is_header)
        if is_header:
            set_cell_shading(row.cells[0])
            set_cell_shading(row.cells[1])
    
    doc.add_paragraph()
    
    # ===== 4. ตารางที่ X-2: Truck Factor =====
    cap2 = doc.add_paragraph()
    cap2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap2.paragraph_format.space_after = Pt(3)
    
    run = cap2.add_run(f"ตารางที่ {tbl_tf} ")
    set_run(run, font_size=FONT_SIZE, bold=True)
    run = cap2.add_run("ค่า Truck Factor ของรถบรรทุกหนัก")
    set_run(run, font_size=FONT_SIZE, bold=False)
    
    tf_table = doc.add_table(rows=7, cols=3)
    tf_table.style = 'Table Grid'
    tf_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    hdr = tf_table.rows[0]
    for j, h in enumerate(['รหัส', 'ประเภท', 'Truck Factor']):
        hdr.cells[j].text = h
        set_cell_font(hdr.cells[j], font_size=TABLE_FONT_SIZE, bold=True,
                      align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_shading(hdr.cells[j])
    
    for i, code in enumerate(TRUCKS.keys()):
        row = tf_table.rows[i + 1]
        row.cells[0].text = code
        row.cells[1].text = TRUCKS[code]['desc']
        row.cells[2].text = f"{truck_factors[code]:.4f}"
        for cell in row.cells:
            set_cell_font(cell, font_size=TABLE_FONT_SIZE, bold=False)
        set_cell_font(row.cells[0], font_size=TABLE_FONT_SIZE, bold=False,
                      align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_font(row.cells[2], font_size=TABLE_FONT_SIZE, bold=False,
                      align=WD_ALIGN_PARAGRAPH.RIGHT)
    
    doc.add_paragraph()
    
    # ===== 5. ตารางที่ X-3: ESAL =====
    cap3 = doc.add_paragraph()
    cap3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap3.paragraph_format.space_after = Pt(3)
    
    run = cap3.add_run(f"ตารางที่ {tbl_esal} ")
    set_run(run, font_size=FONT_SIZE, bold=True)
    run = cap3.add_run(f"ค่าปริมาณเพลามาตรฐาน สำหรับผิวทาง{pavement_thai} ที่ระยะเวลาออกแบบ {num_years} ปี")
    set_run(run, font_size=FONT_SIZE, bold=False)
    
    headers = ['Year', 'MB', 'HB', 'MT', 'HT', 'TR', 'STR', 'AADT', 'ESAL', 'ACC. ESAL']
    esal_table = doc.add_table(rows=len(results_df) + 1, cols=len(headers))
    esal_table.style = 'Table Grid'
    esal_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    hdr = esal_table.rows[0]
    for j, header in enumerate(headers):
        hdr.cells[j].text = header
        for paragraph in hdr.cells[j].paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_font(hdr.cells[j], font_size=12, bold=True)
        set_cell_shading(hdr.cells[j])
    
    for i, row_data in results_df.iterrows():
        row = esal_table.rows[i + 1]
        for j, header in enumerate(headers):
            if header == 'ACC. ESAL':
                value = row_data.get('ACC_ESAL', 0)
            else:
                value = row_data.get(header, 0)
            
            if header in ['ESAL', 'ACC. ESAL', 'AADT']:
                row.cells[j].text = f"{int(value):,}"
            else:
                row.cells[j].text = str(int(value))
            
            for paragraph in row.cells[j].paragraphs:
                if header == 'Year':
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            set_cell_font(row.cells[j], font_size=12, bold=False)


def save_project(pavement_type, pt, param, lane_factor, direction_factor, truck_factors, traffic_df,
                 report_settings=None):
    """บันทึก Project เป็น JSON"""
    project = {
        'version': '3.0',
        'created': datetime.now().isoformat(),
        'pavement_type': pavement_type,
        'pt': pt,
        'param': param,
        'lane_factor': lane_factor,
        'direction_factor': direction_factor,
        'truck_factors': truck_factors,
        'traffic_data': traffic_df.to_dict('records'),
    }
    if report_settings:
        project['report_settings'] = report_settings
    return json.dumps(project, ensure_ascii=False, indent=2)


def load_project(json_data):
    """โหลด Project จาก JSON"""
    try:
        project = json.loads(json_data)
        return project
    except:
        return None


def get_all_truck_factors_table(pavement_type, pt):
    """สร้างตาราง Truck Factor ทั้งหมด"""
    data = []
    
    if pavement_type == 'rigid':
        params = [10, 11, 12, 13, 14, 15, 16]
        param_label = 'D'
        if pt == 2.0:
            tf_table = TRUCK_FACTORS_RIGID_PT20
        elif pt == 2.5:
            tf_table = TRUCK_FACTORS_RIGID_PT25
        else:
            tf_table = TRUCK_FACTORS_RIGID_PT30
    else:
        params = [4, 5, 6, 7, 8, 9]
        param_label = 'SN'
        if pt == 2.0:
            tf_table = TRUCK_FACTORS_FLEX_PT20
        elif pt == 2.5:
            tf_table = TRUCK_FACTORS_FLEX_PT25
        else:
            tf_table = TRUCK_FACTORS_FLEX_PT30
    
    for code in TRUCKS.keys():
        row = {'ประเภท': code, 'รายละเอียด': TRUCKS[code]['desc']}
        for p in params:
            col_name = f'{param_label}={p}"' if pavement_type == 'rigid' else f'{param_label}={p}'
            row[col_name] = f"{tf_table[code][p]:.4f}"
        data.append(row)
    
    return pd.DataFrame(data)


# ============================================================
# Preview HTML for intro paragraph
# ============================================================
def generate_intro_preview_html(pavement_type, num_years, tbl_param, tbl_tf, tbl_esal, section_num):
    """สร้าง HTML preview ของบทเกริ่นนำ พร้อม highlight สี"""
    
    PURPLE = "background-color: #D8B4FE; padding: 1px 4px; border-radius: 3px; font-weight: bold;"
    YELLOW = "background-color: #FDE68A; padding: 1px 4px; border-radius: 3px; font-weight: bold;"
    
    if pavement_type == 'flexible':
        pavement_thai = "ยืดหยุ่น"
        intro_html = (
            f'<span style="{YELLOW}">{section_num}</span>&nbsp;&nbsp;'
            f'<b>ปริมาณเพลามาตรฐาน (ESALs) ระยะเวลาออกแบบ '
            f'<span style="{PURPLE}">{num_years}</span> ปี ผิวทาง Flexible Pavement</b>'
            f'<br><br>'
            f'<p style="text-indent: 40px; text-align: justify; text-justify: inter-character; margin: 0;">'
            f'ในการคำนวณปริมาณเพลามาตรฐาน สำหรับผิวทาง{pavement_thai} '
            f'ที่ปรึกษาได้กำหนดค่าพารามิเตอร์ต่าง ๆ และค่า Truck Factor '
            f'ของรถบรรทุกหนัก ที่ใช้สำหรับการคำนวณ ดังแสดงในตารางที่ '
            f'<span style="{YELLOW}">{tbl_param}</span> และ '
            f'<span style="{YELLOW}">{tbl_tf}</span> '
            f'ดังนั้นค่าปริมาณเพลามาตรฐาน สำหรับผิวทาง{pavement_thai} '
            f'ที่ระยะเวลาออกแบบ <span style="{PURPLE}">{num_years}</span> ปี '
            f'แสดงดังตารางที่ <span style="{YELLOW}">{tbl_esal}</span></p>'
        )
    else:
        pavement_thai = "แบบแข็ง"
        intro_html = (
            f'<span style="{YELLOW}">{section_num}</span>&nbsp;&nbsp;'
            f'<b>ปริมาณเพลามาตรฐาน (ESALs) ระยะเวลาออกแบบ '
            f'<span style="{PURPLE}">{num_years}</span> ปี ผิวทาง Rigid Pavement</b>'
            f'<br><br>'
            f'<p style="text-indent: 40px; text-align: justify; text-justify: inter-character; margin: 0;">'
            f'ในการคำนวณปริมาณเพลามาตรฐานสำหรับผิวทางแบบแข็งหรือผิวทางคอนกรีต '
            f'โดยที่ปรึกษาได้กำหนดค่าพารามิเตอร์ต่าง ๆ และค่า Truck Factor '
            f'ของรถบรรทุกหนัก ที่ใช้สำหรับการคำนวณ ดังแสดงในตารางที่ '
            f'<span style="{YELLOW}">{tbl_param}</span> และ '
            f'<span style="{YELLOW}">{tbl_tf}</span> '
            f'ดังนั้นค่าปริมาณเพลามาตรฐาน สำหรับผิวทาง{pavement_thai} '
            f'ที่ระยะเวลาออกแบบ <span style="{PURPLE}">{num_years}</span> ปี '
            f'แสดงดังตารางที่ <span style="{YELLOW}">{tbl_esal}</span></p>'
        )
    
    return f'''
    <div style="background: #f9f9f9; padding: 15px; border-radius: 8px; border: 1px solid #ddd;
                font-family: 'TH SarabunPSK', sans-serif; font-size: 15px; line-height: 1.8;">
        {intro_html}
    </div>
    '''


# ============================================================
# Streamlit App
# ============================================================
def main():
    st.set_page_config(
        page_title="ESAL Calculator - AASHTO 1993 v3.0",
        page_icon="🛣️",
        layout="wide"
    )
    
    st.markdown("""
    <style>
    .main-header { font-size: 2.5rem; font-weight: bold; color: #1E3A5F; text-align: center; margin-bottom: 0.5rem; }
    .sub-header { font-size: 1.2rem; color: #4A6FA5; text-align: center; margin-bottom: 2rem; }
    .metric-box { background: linear-gradient(135deg, #1E3A5F 0%, #4A6FA5 100%); padding: 1.5rem; border-radius: 10px; color: white; text-align: center; margin: 0.5rem 0; }
    .metric-value { font-size: 2rem; font-weight: bold; }
    .metric-label { font-size: 0.9rem; opacity: 0.9; }
    </style>
    """, unsafe_allow_html=True)
    
    st.markdown('<p class="main-header">🛣️ ESAL Calculator v3.0</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">คำนวณปริมาณเพลาเดี่ยวมาตรฐานเทียบเท่า ตามมาตรฐาน AASHTO 1993</p>', unsafe_allow_html=True)
    
    # Initialize session state
    if 'traffic_df' not in st.session_state:
        st.session_state['traffic_df'] = None
    
    # Sidebar
    with st.sidebar:
        st.header("⚙️ พารามิเตอร์การคำนวณ")
        
        # Project Load/Save
        st.subheader("📁 Project")
        
        uploaded_project = st.file_uploader("📥 โหลด Project", type=['json'], key='load_project')
        if uploaded_project is not None:
            try:
                file_id = f"{uploaded_project.name}_{uploaded_project.size}"
                if st.session_state.get('last_uploaded_file') != file_id:
                    st.session_state['last_uploaded_file'] = file_id
                    
                    project = load_project(uploaded_project.read().decode('utf-8'))
                    if project:
                        st.session_state['input_pavement_type'] = project.get('pavement_type', 'rigid')
                        st.session_state['input_pt'] = project.get('pt', 2.5)
                        st.session_state['input_param'] = project.get('param', 12)
                        st.session_state['input_lane_factor'] = project.get('lane_factor', 0.5)
                        st.session_state['input_direction_factor'] = project.get('direction_factor', 0.9)
                        st.session_state['loaded_tf'] = project.get('truck_factors', {})
                        
                        # โหลด report_settings
                        rs = project.get('report_settings', {})
                        if rs:
                            for key, val in rs.items():
                                st.session_state[f'input_{key}'] = val
                        
                        loaded_traffic = project.get('traffic_data', None)
                        if loaded_traffic:
                            st.session_state['traffic_df'] = pd.DataFrame(loaded_traffic)
                        
                        st.success("✅ โหลด Project สำเร็จ!")
                        st.rerun()
                    else:
                        st.error("❌ ไม่สามารถอ่านไฟล์ได้")
            except Exception as e:
                st.error(f"❌ เกิดข้อผิดพลาด: {e}")
        
        default_pavement = st.session_state.get('input_pavement_type', 'rigid')
        default_pt = st.session_state.get('input_pt', 2.5)
        default_param = st.session_state.get('input_param', 12 if default_pavement == 'rigid' else 7)
        default_lane = st.session_state.get('input_lane_factor', 0.5)
        default_dir = st.session_state.get('input_direction_factor', 0.9)
        loaded_tf = st.session_state.get('loaded_tf', {})
        
        st.divider()
        
        pavement_options = ['rigid', 'flexible']
        pavement_idx = pavement_options.index(default_pavement) if default_pavement in pavement_options else 0
        pavement_type = st.selectbox(
            "ประเภทผิวทาง",
            options=pavement_options,
            index=pavement_idx,
            format_func=lambda x: '🧱 Rigid Pavement (คอนกรีต)' if x == 'rigid' else '🛤️ Flexible Pavement (ลาดยาง)',
            key="input_pavement_type"
        )
        
        pt_options = [2.0, 2.5, 3.0]
        pt_idx = pt_options.index(default_pt) if default_pt in pt_options else 1
        pt = st.selectbox(
            "Terminal Serviceability (pt)",
            options=pt_options,
            index=pt_idx,
            format_func=lambda x: f"pt = {x}",
            key="input_pt"
        )
        
        if pavement_type == 'rigid':
            param_options = [10, 11, 12, 13, 14, 15, 16]
            if default_param not in param_options:
                default_param = 12
            default_idx = param_options.index(default_param) if default_param in param_options else 2
            param = st.selectbox(
                "ความหนาพื้นคอนกรีต (D)",
                options=param_options,
                index=default_idx,
                format_func=lambda x: f"D = {x} นิ้ว",
                key="input_param_rigid"
            )
            param_label = f"D = {param} นิ้ว"
        else:
            param_options = [4, 5, 6, 7, 8, 9]
            if default_param not in param_options:
                default_param = 7
            default_idx = param_options.index(default_param) if default_param in param_options else 3
            param = st.selectbox(
                "Structural Number (SN)",
                options=param_options,
                index=default_idx,
                format_func=lambda x: f"SN = {x}",
                key="input_param_flex"
            )
            param_label = f"SN = {param}"
        
        st.divider()
        
        st.subheader("🚗 ค่าสัดส่วน")
        lane_factor = st.slider(
            "Lane Distribution Factor", 
            0.1, 1.0, 
            value=st.session_state.get('input_lane_factor', default_lane), 
            step=0.05,
            key="input_lane_factor"
        )
        direction_factor = st.slider(
            "Directional Factor", 
            0.5, 1.0, 
            value=st.session_state.get('input_direction_factor', default_dir), 
            step=0.1,
            key="input_direction_factor"
        )
        
        st.divider()
        
        st.subheader("🚛 ค่า Truck Factor")
        
        tf_key = f"tf_{pavement_type}_{pt}_{param}"
        
        if loaded_tf and tf_key not in st.session_state:
            st.session_state[tf_key] = {}
            for code in TRUCKS.keys():
                if code in loaded_tf:
                    st.session_state[tf_key][code] = loaded_tf[code]
                else:
                    st.session_state[tf_key][code] = get_default_truck_factor(code, pavement_type, pt, param)
        elif tf_key not in st.session_state:
            st.session_state[tf_key] = {}
            for code in TRUCKS.keys():
                st.session_state[tf_key][code] = get_default_truck_factor(code, pavement_type, pt, param)
        
        if st.button("🔄 Reset เป็นค่า Default", use_container_width=True):
            for code in TRUCKS.keys():
                st.session_state[tf_key][code] = get_default_truck_factor(code, pavement_type, pt, param)
            st.rerun()
        
        st.caption("กรอกค่า Truck Factor (แก้ไขได้)")
        
        truck_factors = {}
        for code in TRUCKS.keys():
            default_val = get_default_truck_factor(code, pavement_type, pt, param)
            current_val = st.session_state[tf_key].get(code, default_val)
            
            new_val = st.number_input(
                f"{code}",
                min_value=0.0,
                max_value=50.0,
                value=float(current_val),
                step=0.0001,
                format="%.4f",
                key=f"input_{tf_key}_{code}",
                help=f"{TRUCKS[code]['desc']} | Default: {default_val:.4f}"
            )
            
            st.session_state[tf_key][code] = new_val
            truck_factors[code] = new_val
        
        st.divider()
        
        st.subheader("📥 ดาวน์โหลด Template")
        template_df = create_template()
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            template_df.to_excel(writer, index=False, sheet_name='Traffic Data')
        st.download_button(
            label="📄 ดาวน์โหลด Template Excel",
            data=output.getvalue(),
            file_name="traffic_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    # Main Content
    tab1, tab2, tab3 = st.tabs(["📊 คำนวณ ESAL", "🚛 ข้อมูล Truck Factor", "📘 คู่มือ"])
    
    with tab1:
        col1, col2 = st.columns([1, 2])
        
        with col1:
            st.subheader("📤 อัพโหลดข้อมูลปริมาณจราจร")
            
            uploaded_file = st.file_uploader(
                "เลือกไฟล์ Excel",
                type=['xlsx', 'xls'],
                help="อัพโหลดไฟล์ Excel (หน่วย: คัน/วัน)"
            )
            
            if 'use_sample' not in st.session_state:
                st.session_state['use_sample'] = False
            
            if uploaded_file is not None:
                try:
                    traffic_df = pd.read_excel(uploaded_file)
                    st.session_state['traffic_df'] = traffic_df
                    st.success("✅ อัพโหลดสำเร็จ!")
                    st.session_state['use_sample'] = False
                except Exception as e:
                    st.error(f"❌ เกิดข้อผิดพลาด: {e}")
                    traffic_df = st.session_state.get('traffic_df', None)
            elif st.session_state.get('traffic_df') is not None:
                traffic_df = st.session_state['traffic_df']
            else:
                st.info("📌 อัพโหลดไฟล์ Excel หรือใช้ข้อมูลตัวอย่าง")
                
                if st.button("🔄 ใช้ข้อมูลตัวอย่าง", use_container_width=True):
                    st.session_state['use_sample'] = True
                    st.session_state['traffic_df'] = create_template()
                
                traffic_df = st.session_state.get('traffic_df', None)
            
            if traffic_df is not None:
                st.write("**ข้อมูลปริมาณจราจร (คัน/วัน):**")
                st.dataframe(traffic_df, use_container_width=True, height=350)
        
        with col2:
            st.subheader("📈 ผลการคำนวณ ESAL")
            
            if traffic_df is not None:
                results_df, total_esal = calculate_esal_with_acc(
                    traffic_df, truck_factors, lane_factor, direction_factor
                )
                
                # Metrics
                col_m1, col_m2, col_m3, col_m4 = st.columns(4)
                
                with col_m1:
                    st.markdown(f"""
                    <div class="metric-box">
                        <div class="metric-value">{total_esal:,}</div>
                        <div class="metric-label">ESAL รวมทั้งหมด</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col_m2:
                    st.markdown(f"""
                    <div class="metric-box">
                        <div class="metric-value">{len(traffic_df)}</div>
                        <div class="metric-label">จำนวนปี</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col_m3:
                    pavement_label = "Rigid" if pavement_type == 'rigid' else "Flexible"
                    st.markdown(f"""
                    <div class="metric-box">
                        <div class="metric-value">{pavement_label}</div>
                        <div class="metric-label">ประเภทผิวทาง</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col_m4:
                    st.markdown(f"""
                    <div class="metric-box">
                        <div class="metric-value">{param_label}</div>
                        <div class="metric-label">พารามิเตอร์</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                st.divider()
                
                # Truck Factor Table
                st.write("**🚛 ค่า Truck Factor ที่ใช้:**")
                tf_display = []
                for code, tf in truck_factors.items():
                    default_tf = get_default_truck_factor(code, pavement_type, pt, param)
                    status = "✅" if abs(tf - default_tf) < 0.0001 else "✏️ แก้ไข"
                    tf_display.append({
                        'รหัส': code, 
                        'ประเภท': TRUCKS[code]['desc'], 
                        'Truck Factor': f"{tf:.4f}",
                        'Default': f"{default_tf:.4f}",
                        'สถานะ': status
                    })
                st.dataframe(pd.DataFrame(tf_display), use_container_width=True, hide_index=True)
                
                st.divider()
                
                # ESAL Results Table
                st.write("**📊 ESAL รายปี:**")
                display_df = results_df.copy()
                display_df.columns = ['ปีที่', 'MB', 'HB', 'MT', 'HT', 'TR', 'STR', 'AADT', 'ESAL', 'ACC. ESAL']
                st.dataframe(display_df, use_container_width=True, height=400)
                
                st.divider()
                
                # ============================================================
                # ส่วนตั้งค่ารายงาน Word (Expander)
                # ============================================================
                with st.expander("📝 ตั้งค่ารายงาน Word", expanded=False):
                    st.markdown("#### หมายเลขหัวข้อและตาราง")
                    st.caption("กรอกเลขตารางเริ่มต้น ระบบจะเพิ่ม +1, +2 อัตโนมัติ (เช่น 4-1 → 4-2 → 4-3)")
                    
                    col_f1, col_f2, col_f3, col_f4 = st.columns(4)
                    
                    with col_f1:
                        flex_section_number = st.text_input(
                            "🛤️ เลขหัวข้อ Flexible",
                            value=st.session_state.get('input_flex_section_number', "4.2.2"),
                            key="input_flex_section_number"
                        )
                    with col_f2:
                        flex_table_start = st.text_input(
                            "🛤️ เลขตารางเริ่มต้น Flexible",
                            value=st.session_state.get('input_flex_table_start', "4-1"),
                            key="input_flex_table_start"
                        )
                    with col_f3:
                        rigid_section_number = st.text_input(
                            "🧱 เลขหัวข้อ Rigid",
                            value=st.session_state.get('input_rigid_section_number', "4.2.3"),
                            key="input_rigid_section_number"
                        )
                    with col_f4:
                        rigid_table_start = st.text_input(
                            "🧱 เลขตารางเริ่มต้น Rigid",
                            value=st.session_state.get('input_rigid_table_start', "4-4"),
                            key="input_rigid_table_start"
                        )
                    
                    # แสดงสรุปหมายเลขตาราง
                    col_sum1, col_sum2 = st.columns(2)
                    with col_sum1:
                        st.info(
                            f"**Flexible:** ตารางที่ {flex_table_start} (พารามิเตอร์), "
                            f"{increment_table_number(flex_table_start, 1)} (Truck Factor), "
                            f"{increment_table_number(flex_table_start, 2)} (ESAL)"
                        )
                    with col_sum2:
                        st.info(
                            f"**Rigid:** ตารางที่ {rigid_table_start} (พารามิเตอร์), "
                            f"{increment_table_number(rigid_table_start, 1)} (Truck Factor), "
                            f"{increment_table_number(rigid_table_start, 2)} (ESAL)"
                        )
                    
                    st.markdown("---")
                    
                    # ===== พารามิเตอร์สำหรับ Combined Report =====
                    st.markdown("#### ⚙️ พารามิเตอร์สำหรับรายงานรวม (Flexible + Rigid)")
                    st.caption("กำหนดพารามิเตอร์ของผิวทางอีกประเภทหนึ่ง สำหรับ export รายงานรวม")
                    
                    if pavement_type == 'rigid':
                        # ตั้งอยู่ Rigid → ต้องกรอก Flexible params
                        st.markdown("**🛤️ พารามิเตอร์ Flexible Pavement (สำหรับรายงานรวม)**")
                        col_p1, col_p2, col_p3, col_p4 = st.columns(4)
                        with col_p1:
                            comb_flex_pt = st.selectbox("pt (Flexible)", [2.0, 2.5, 3.0],
                                                        index=1, key="comb_flex_pt")
                        with col_p2:
                            comb_flex_sn = st.selectbox("SN", [4, 5, 6, 7, 8, 9],
                                                         index=3, key="comb_flex_sn")
                        with col_p3:
                            comb_flex_lane = st.number_input("Lane Factor", 0.1, 1.0,
                                                              value=lane_factor, step=0.05,
                                                              key="comb_flex_lane")
                        with col_p4:
                            comb_flex_dir = st.number_input("Direction Factor", 0.5, 1.0,
                                                             value=direction_factor, step=0.1,
                                                             key="comb_flex_dir")
                        
                        # Truck factors for flexible
                        comb_flex_tf = {}
                        for code in TRUCKS.keys():
                            comb_flex_tf[code] = get_default_truck_factor(code, 'flexible', comb_flex_pt, comb_flex_sn)
                    else:
                        # ตั้งอยู่ Flexible → ต้องกรอก Rigid params
                        st.markdown("**🧱 พารามิเตอร์ Rigid Pavement (สำหรับรายงานรวม)**")
                        col_p1, col_p2, col_p3, col_p4 = st.columns(4)
                        with col_p1:
                            comb_rigid_pt = st.selectbox("pt (Rigid)", [2.0, 2.5, 3.0],
                                                          index=1, key="comb_rigid_pt")
                        with col_p2:
                            comb_rigid_d = st.selectbox("D (นิ้ว)", [10, 11, 12, 13, 14, 15, 16],
                                                         index=3, key="comb_rigid_d")
                        with col_p3:
                            comb_rigid_lane = st.number_input("Lane Factor", 0.1, 1.0,
                                                               value=lane_factor, step=0.05,
                                                               key="comb_rigid_lane")
                        with col_p4:
                            comb_rigid_dir = st.number_input("Direction Factor", 0.5, 1.0,
                                                              value=direction_factor, step=0.1,
                                                              key="comb_rigid_dir")
                        
                        comb_rigid_tf = {}
                        for code in TRUCKS.keys():
                            comb_rigid_tf[code] = get_default_truck_factor(code, 'rigid', comb_rigid_pt, comb_rigid_d)
                    
                    st.markdown("---")
                    
                    # ===== Preview บทเกริ่นนำ =====
                    st.markdown("#### 👁️ ตัวอย่างบทเกริ่นนำ (Preview)")
                    
                    num_years = len(traffic_df)
                    
                    flex_tbl_param = flex_table_start
                    flex_tbl_tf = increment_table_number(flex_table_start, 1)
                    flex_tbl_esal = increment_table_number(flex_table_start, 2)
                    
                    rigid_tbl_param = rigid_table_start
                    rigid_tbl_tf = increment_table_number(rigid_table_start, 1)
                    rigid_tbl_esal = increment_table_number(rigid_table_start, 2)
                    
                    col_prev1, col_prev2 = st.columns(2)
                    
                    with col_prev1:
                        st.markdown("**🛤️ Flexible Pavement**")
                        html_flex = generate_intro_preview_html(
                            'flexible', num_years,
                            flex_tbl_param, flex_tbl_tf, flex_tbl_esal,
                            flex_section_number
                        )
                        st.markdown(html_flex, unsafe_allow_html=True)
                    
                    with col_prev2:
                        st.markdown("**🧱 Rigid Pavement**")
                        html_rigid = generate_intro_preview_html(
                            'rigid', num_years,
                            rigid_tbl_param, rigid_tbl_tf, rigid_tbl_esal,
                            rigid_section_number
                        )
                        st.markdown(html_rigid, unsafe_allow_html=True)
                    
                    st.caption("🟣 สีม่วง = ดึงจากข้อมูลอัตโนมัติ | 🟡 สีเหลือง = ผู้ใช้กรอกเอง")
                
                st.divider()
                
                # ============================================================
                # Download buttons
                # ============================================================
                st.write("**📥 ดาวน์โหลดรายงาน:**")
                
                # Collect report settings
                report_settings = {
                    'flex_section_number': st.session_state.get('input_flex_section_number', '4.2.2'),
                    'flex_table_start': st.session_state.get('input_flex_table_start', '4-1'),
                    'rigid_section_number': st.session_state.get('input_rigid_section_number', '4.2.3'),
                    'rigid_table_start': st.session_state.get('input_rigid_table_start', '4-4'),
                }
                
                col_dl1, col_dl2, col_dl3, col_dl4 = st.columns(4)
                
                with col_dl1:
                    excel_report = create_excel_report(
                        results_df, pavement_type, pt, param, lane_factor, direction_factor,
                        total_esal, truck_factors, len(traffic_df)
                    )
                    st.download_button(
                        label="📊 Excel (ปัจจุบัน)",
                        data=excel_report.getvalue(),
                        file_name=f"ESAL_Report_{pavement_type}_{param}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                with col_dl2:
                    word_report = create_word_report_single(
                        results_df, pavement_type, pt, param, lane_factor, direction_factor,
                        total_esal, truck_factors, len(traffic_df), report_settings
                    )
                    if word_report:
                        pv_label = "Rigid" if pavement_type == 'rigid' else "Flexible"
                        st.download_button(
                            label=f"📝 Word ({pv_label})",
                            data=word_report.getvalue(),
                            file_name=f"ESAL_Report_{pavement_type}_{param}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                    else:
                        st.warning("⚠️ กรุณาติดตั้ง python-docx")
                
                with col_dl3:
                    # Combined Word Report
                    try:
                        if pavement_type == 'rigid':
                            flex_params_comb = {
                                'pt': st.session_state.get('comb_flex_pt', 2.5),
                                'param': st.session_state.get('comb_flex_sn', 7),
                                'lane_factor': st.session_state.get('comb_flex_lane', lane_factor),
                                'direction_factor': st.session_state.get('comb_flex_dir', direction_factor),
                                'truck_factors': {code: get_default_truck_factor(
                                    code, 'flexible',
                                    st.session_state.get('comb_flex_pt', 2.5),
                                    st.session_state.get('comb_flex_sn', 7)
                                ) for code in TRUCKS.keys()}
                            }
                            rigid_params_comb = {
                                'pt': pt, 'param': param,
                                'lane_factor': lane_factor, 'direction_factor': direction_factor,
                                'truck_factors': truck_factors
                            }
                        else:
                            flex_params_comb = {
                                'pt': pt, 'param': param,
                                'lane_factor': lane_factor, 'direction_factor': direction_factor,
                                'truck_factors': truck_factors
                            }
                            rigid_params_comb = {
                                'pt': st.session_state.get('comb_rigid_pt', 2.5),
                                'param': st.session_state.get('comb_rigid_d', 13),
                                'lane_factor': st.session_state.get('comb_rigid_lane', lane_factor),
                                'direction_factor': st.session_state.get('comb_rigid_dir', direction_factor),
                                'truck_factors': {code: get_default_truck_factor(
                                    code, 'rigid',
                                    st.session_state.get('comb_rigid_pt', 2.5),
                                    st.session_state.get('comb_rigid_d', 13)
                                ) for code in TRUCKS.keys()}
                            }
                        
                        word_combined = create_word_report_combined(
                            traffic_df, flex_params_comb, rigid_params_comb, report_settings
                        )
                        if word_combined:
                            st.download_button(
                                label="📝 Word (รวม Flex+Rigid)",
                                data=word_combined.getvalue(),
                                file_name="ESAL_Report_Combined.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )
                        else:
                            st.warning("⚠️ กรุณาติดตั้ง python-docx")
                    except Exception as e:
                        st.error(f"❌ เกิดข้อผิดพลาด: {e}")
                
                with col_dl4:
                    project_json = save_project(
                        pavement_type, pt, param, lane_factor, direction_factor,
                        truck_factors, traffic_df, report_settings
                    )
                    st.download_button(
                        label="💾 บันทึก Project",
                        data=project_json,
                        file_name=f"ESAL_Project_{pavement_type}_{param}.json",
                        mime="application/json",
                        use_container_width=True
                    )
            else:
                st.warning("⚠️ กรุณาอัพโหลดข้อมูลหรือใช้ข้อมูลตัวอย่าง")
    
    with tab2:
        st.subheader("🚛 ข้อมูลรถบรรทุก 6 ประเภทตามกรมทางหลวง")
        
        truck_details = []
        for code, truck in TRUCKS.items():
            axle_info = []
            for axle in truck['axles']:
                axle_info.append(f"{axle['name']}: {axle['load_ton']} ตัน ({axle['type']})")
            truck_details.append({'รหัส': code, 'ประเภท': truck['desc'], 'ข้อมูลเพลา': ' | '.join(axle_info)})
        
        st.dataframe(pd.DataFrame(truck_details), use_container_width=True, hide_index=True)
        
        st.divider()
        st.subheader("📊 ตาราง Truck Factor (ค่า Default ตาม AASHTO 1993)")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.write("**🧱 Rigid Pavement (pt = 2.0)**")
            st.dataframe(get_all_truck_factors_table('rigid', 2.0), use_container_width=True, hide_index=True)
            
            st.write("**🧱 Rigid Pavement (pt = 2.5)**")
            st.dataframe(get_all_truck_factors_table('rigid', 2.5), use_container_width=True, hide_index=True)
            
            st.write("**🧱 Rigid Pavement (pt = 3.0)**")
            st.dataframe(get_all_truck_factors_table('rigid', 3.0), use_container_width=True, hide_index=True)
        
        with col2:
            st.write("**🛤️ Flexible Pavement (pt = 2.0)**")
            st.dataframe(get_all_truck_factors_table('flexible', 2.0), use_container_width=True, hide_index=True)
            
            st.write("**🛤️ Flexible Pavement (pt = 2.5)**")
            st.dataframe(get_all_truck_factors_table('flexible', 2.5), use_container_width=True, hide_index=True)
            
            st.write("**🛤️ Flexible Pavement (pt = 3.0)**")
            st.dataframe(get_all_truck_factors_table('flexible', 3.0), use_container_width=True, hide_index=True)
    
    with tab3:
        st.subheader("📘 คู่มือการใช้งาน")
        
        st.markdown("""
        ### 1️⃣ เตรียมไฟล์ Excel
        
        | คอลัมน์ | คำอธิบาย |
        |---------|----------|
        | `Year` | ปีที่ (1, 2, 3, ... n) |
        | `MB` | Medium Bus (คัน/วัน) |
        | `HB` | Heavy Bus (คัน/วัน) |
        | `MT` | Medium Truck (คัน/วัน) |
        | `HT` | Heavy Truck (คัน/วัน) |
        | `STR` | Semi-Trailer (คัน/วัน) |
        | `TR` | Full Trailer (คัน/วัน) |
        
        ### 2️⃣ ตั้งค่าพารามิเตอร์
        
        - **Rigid:** D = 10-16 นิ้ว
        - **Flexible:** SN = 4-9
        - **pt:** 2.0, 2.5 หรือ 3.0
        
        ### 3️⃣ ฟีเจอร์ในเวอร์ชัน 3.0
        
        - **ACC. ESAL:** แสดงค่า ESAL สะสม
        - **Export Excel:** รายงานในรูปแบบมาตรฐาน
        - **Export Word (แยก/รวม):** รายงานสำหรับเอกสาร
        - **ระบบหมายเลขตาราง:** Auto-increment กรอกเลขเริ่มต้น ระบบเพิ่มให้อัตโนมัติ
        - **บทเกริ่นนำ:** แยก Flexible / Rigid พร้อม Preview
        - **Save/Load Project:** บันทึกค่าทั้งหมดรวม report settings
        
        ### 4️⃣ สูตรคำนวณ ESAL
        """)
        
        st.latex(r'ESAL = \sum_{i=1}^{n} \sum_{j=1}^{6} (ADT_{ij} \times TF_j \times LF \times DF \times 365)')
        
        st.markdown("""
        ### 5️⃣ การ Export รายงาน Word
        
        | ปุ่ม | คำอธิบาย |
        |------|----------|
        | 📝 Word (Flexible/Rigid) | รายงานเฉพาะประเภทที่กำลังคำนวณ |
        | 📝 Word (รวม Flex+Rigid) | รายงานรวมทั้ง 2 ประเภทในไฟล์เดียว |
        
        กำหนดเลขหัวข้อและเลขตารางได้ที่ **📝 ตั้งค่ารายงาน Word** (Expander)
        
        ### 📚 อ้างอิง
        - AASHTO Guide for Design of Pavement Structures (1993)
        - กรมทางหลวง กระทรวงคมนาคม
        """)
    
    st.divider()
    st.markdown("""
    <div style="text-align: center; color: #888;">
        พัฒนาเพื่อการเรียนการสอนโดย รศ.ดร.อิทธิพล มีผล ภาควิชาครุศาสตร์โยธา มจพ. | ESAL Calculator v3.0
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
