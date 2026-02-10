# -*- coding: utf-8 -*-
"""
Pavement Design Report Merger – 10 Files
โครงสร้างรายงานตามที่กำหนด:

1) Truck Factor (ถ้ามี)
2) ESALs (หัวข้อใหญ่)
  2.1) ESALs (Flexible)
  2.2) ESALs (Rigid)
3) CBR Analysis
4) AC Design
5) ผิวทางคอนกรีต (หัวข้อใหญ่)
  5.1) JPCP/JRCP
  5.2) k-value (JPCP/JRCP)
  5.3) CRCP
  5.4) k-value (CRCP)
6) Cost Estimate (ถ้ามี)
"""

import streamlit as st
import os
import tempfile
from datetime import datetime
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docxcompose.composer import Composer
import io

# ═══════════════════════════════════════════════════════════════
# CONFIG: โครงสร้างรายงาน (10 ไฟล์อัปโหลด + หัวข้อใหญ่ 2, 5)
# ═══════════════════════════════════════════════════════════════

SECTION_CONFIG = [
    {
        "group_icon": "📊",
        "group_title": "1) Truck Factor (ถ้ามี)",
        "items": [
            {
                "key": "truck_factor",
                "title": "1) Truck Factor (ถ้ามี)",
                "label": "**1) Truck Factor (ถ้ามี)**",
                "uploader_label": "เลือกไฟล์ 1) Truck Factor",
                "help": "ไฟล์รายงานการคำนวณ Truck Factor (ถ้ามี)",
                "report_title": "1) Truck Factor",
            }
        ],
    },
    {
        "group_icon": "📈",
        "group_title": "2) ESALs",
        "group_caption": "การคำนวณ ESALs (Equivalent Single Axle Loads)",
        "columns": 2,
        "items": [
            {
                "key": "esals_flexible",
                "title": "2.1) ESALs (Flexible)",
                "label": "**2.1) ESALs (Flexible)**",
                "uploader_label": "เลือกไฟล์ 2.1) ESALs (Flexible)",
                "help": "ไฟล์รายงานการคำนวณ ESALs สำหรับผิวทางลาดยาง (Flexible Pavement)",
                "report_title": "2.1) ESALs (Flexible Pavement)",
            },
            {
                "key": "esals_rigid",
                "title": "2.2) ESALs (Rigid)",
                "label": "**2.2) ESALs (Rigid)**",
                "uploader_label": "เลือกไฟล์ 2.2) ESALs (Rigid)",
                "help": "ไฟล์รายงานการคำนวณ ESALs สำหรับผิวทางคอนกรีต (Rigid Pavement)",
                "report_title": "2.2) ESALs (Rigid Pavement)",
            },
        ],
    },
    {
        "group_icon": "🔬",
        "group_title": "3) CBR Analysis",
        "items": [
            {
                "key": "cbr_analysis",
                "title": "3) CBR Analysis",
                "label": "**3) CBR Analysis**",
                "uploader_label": "เลือกไฟล์ 3) CBR Analysis",
                "help": "ไฟล์รายงานการวิเคราะห์ค่า CBR",
                "report_title": "3) CBR Analysis",
            }
        ],
    },
    {
        "group_icon": "🛤️",
        "group_title": "4) AC Design",
        "items": [
            {
                "key": "ac_design",
                "title": "4) AC Design",
                "label": "**4) AC Design**",
                "uploader_label": "เลือกไฟล์ 4) AC Design",
                "help": "ไฟล์รายงานการออกแบบผิวทางลาดยาง (AC Design)",
                "report_title": "4) AC Design",
            }
        ],
    },
    {
        "group_icon": "🏗️",
        "group_title": "5) ผิวทางคอนกรีต",
        "group_caption": "การออกแบบผิวทางคอนกรีตและ k-value",
        "columns": 2,
        "items": [
            {
                "key": "jpcp_jrcp",
                "title": "5.1) JPCP/JRCP",
                "label": "**5.1) JPCP/JRCP**",
                "uploader_label": "เลือกไฟล์ 5.1) JPCP/JRCP",
                "help": "ไฟล์รายงานการออกแบบผิวทางคอนกรีต JPCP/JRCP",
                "report_title": "5.1) JPCP/JRCP",
            },
            {
                "key": "k_jpcp_jrcp",
                "title": "5.2) k-value (JPCP/JRCP)",
                "label": "**5.2) k-value (JPCP/JRCP)**",
                "uploader_label": "เลือกไฟล์ 5.2) k-value (JPCP/JRCP)",
                "help": "ไฟล์คำนวณ Corrected k-value สำหรับ JPCP/JRCP",
                "report_title": "5.2) k-value (JPCP/JRCP)",
            },
            {
                "key": "crcp",
                "title": "5.3) CRCP",
                "label": "**5.3) CRCP**",
                "uploader_label": "เลือกไฟล์ 5.3) CRCP",
                "help": "ไฟล์รายงานการออกแบบผิวทางคอนกรีต CRCP",
                "report_title": "5.3) CRCP",
            },
            {
                "key": "k_crcp",
                "title": "5.4) k-value (CRCP)",
                "label": "**5.4) k-value (CRCP)**",
                "uploader_label": "เลือกไฟล์ 5.4) k-value (CRCP)",
                "help": "ไฟล์คำนวณ Corrected k-value สำหรับ CRCP",
                "report_title": "5.4) k-value (CRCP)",
            },
        ],
    },
    {
        "group_icon": "💰",
        "group_title": "6) Cost Estimate (ถ้ามี)",
        "items": [
            {
                "key": "cost_estimate",
                "title": "6) Cost Estimate (ถ้ามี)",
                "label": "**6) Cost Estimate (ถ้ามี)**",
                "uploader_label": "เลือกไฟล์ 6) Cost Estimate",
                "help": "ไฟล์รายงานการประมาณราคาค่าก่อสร้าง (ถ้ามี)",
                "report_title": "6) Cost Estimate",
            }
        ],
    },
]

# ═══════════════════════════════════════════════════════════════
# PAGE CONFIG + CSS (สไตล์เดียวกับ v3.0)
# ═══════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="Pavement Design Report Merger – 10 Files",
    page_icon="🛣️",
    layout="wide"
)

st.markdown("""
<style>
    .main-header {
        font-size: 28px;
        font-weight: bold;
        color: #1E3A5F;
        text-align: center;
        padding: 20px;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border-radius: 10px;
        margin-bottom: 20px;
    }
    .sub-header {
        font-size: 18px;
        color: #4A5568;
        text-align: center;
        margin-bottom: 30px;
    }
    .file-section {
        background-color: #F7FAFC;
        padding: 15px;
        border-radius: 10px;
        margin-bottom: 10px;
        border-left: 4px solid #667eea;
    }
    .section-header {
        background-color: #C6F6D5;
        padding: 10px 15px;
        border-radius: 8px;
        margin: 15px 0 10px 0;
        font-weight: bold;
        color: #276749;
        border-left: 4px solid #38A169;
    }
    .section-caption {
        font-size: 14px;
        color: #4A5568;
        margin-bottom: 5px;
        margin-left: 5px;
    }
    .success-box {
        background-color: #C6F6D5;
        padding: 15px;
        border-radius: 10px;
        border-left: 4px solid #38A169;
    }
    .warning-box {
        background-color: #FEFCBF;
        padding: 15px;
        border-radius: 10px;
        border-left: 4px solid #D69E2E;
    }
    .stButton>button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        font-weight: bold;
        padding: 10px 30px;
        border-radius: 25px;
        border: none;
        font-size: 16px;
    }
    .stButton>button:hover {
        background: linear-gradient(135deg, #764ba2 0%, #667eea 100%);
    }
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════
# Utility
# ═══════════════════════════════════════════════════════════════

def set_thai_font(run, font_name="TH Sarabun New", font_size=15):
    run.font.name = font_name
    run.font.size = Pt(font_size)
    r = run._r
    rPr = r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:cs'), font_name)
    rFonts.set(qn('w:eastAsia'), font_name)


def set_page_margins(section):
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.orientation = WD_ORIENT.PORTRAIT
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.header_distance = Cm(1.25)
    section.footer_distance = Cm(1.25)


def validate_docx_file(file):
    try:
        file_bytes = file.read()
        file.seek(0)
        doc = Document(io.BytesIO(file_bytes))
        if len(doc.paragraphs) == 0 and len(doc.tables) == 0:
            return False, "ไฟล์ว่างเปล่า ไม่มีเนื้อหา"
        return True, ""
    except Exception as e:
        return False, f"ไฟล์เสียหายหรือไม่ใช่ไฟล์ .docx ที่ถูกต้อง ({str(e)})"


def get_all_items():
    items = []
    for group in SECTION_CONFIG:
        for item in group["items"]:
            items.append(item)
    return items

# ═══════════════════════════════════════════════════════════════
# สร้างปก + สารบัญ + รวมไฟล์
# ═══════════════════════════════════════════════════════════════

def create_cover_and_toc(uploaded_files, project_name, report_date):
    doc = Document()
    section = doc.sections[0]
    set_page_margins(section)

    # ปก
    spacer = doc.add_paragraph()
    spacer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    spacer.add_run("\n\n\n\n\n")

    main_title = doc.add_paragraph()
    main_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = main_title.add_run("รายงานการออกแบบโครงสร้างชั้นทาง")
    set_thai_font(run, font_size=24)
    run.font.bold = True

    if project_name:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(f"\n{project_name}")
        set_thai_font(run, font_size=20)
        run.font.bold = True

    date_p = doc.add_paragraph()
    date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = date_p.add_run(f"\n\n\n\n{report_date}")
    set_thai_font(run, font_size=16)

    doc.add_page_break()

    # สารบัญ
    toc_title = doc.add_paragraph()
    toc_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = toc_title.add_run("สารบัญ")
    set_thai_font(run, font_size=18)
    run.font.bold = True

    doc.add_paragraph()

    all_items = get_all_items()
    section_num = 1
    for item in all_items:
        if uploaded_files.get(item["key"]) is not None:
            toc_para = doc.add_paragraph()
            run = toc_para.add_run(f"{section_num}. {item['report_title']}")
            set_thai_font(run, font_size=15)
            section_num += 1

    doc.add_page_break()
    return doc


def create_section_header_doc(section_num, title):
    header_doc = Document()
    p = header_doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(f"{section_num}. {title}")
    set_thai_font(run, font_size=18)
    run.font.bold = True
    header_doc.add_paragraph()
    return header_doc


def merge_documents(uploaded_files, project_name, report_date, progress_callback=None):
    master_doc = create_cover_and_toc(uploaded_files, project_name, report_date)
    composer = Composer(master_doc)

    all_items = get_all_items()
    active_items = [(item, uploaded_files[item["key"]])
                    for item in all_items
                    if uploaded_files.get(item["key"]) is not None]
    total = len(active_items)

    for idx, (item, file) in enumerate(active_items):
        section_num = idx + 1

        header_doc = create_section_header_doc(section_num, item["report_title"])
        composer.append(header_doc)

        file_bytes = file.read()
        file.seek(0)
        source_doc = Document(io.BytesIO(file_bytes))
        composer.append(source_doc)

        if progress_callback:
            progress_callback((idx + 1) / total, f"กำลังรวม: {item['report_title']}")

    return composer.doc

# ═══════════════════════════════════════════════════════════════
# UI Rendering
# ═══════════════════════════════════════════════════════════════

def render_single_uploader(item):
    st.markdown('<div class="file-section">', unsafe_allow_html=True)
    st.markdown(item["label"])
    uploaded = st.file_uploader(
        item["uploader_label"],
        type=['docx'],
        key=item["key"],
        help=item["help"],
    )
    st.markdown('</div>', unsafe_allow_html=True)
    return uploaded


def render_upload_sections():
    uploaded_files = {}
    for group in SECTION_CONFIG:
        st.markdown(
            f'<div class="section-header">{group["group_icon"]} {group["group_title"]}</div>',
            unsafe_allow_html=True
        )
        if "group_caption" in group:
            st.markdown(f'<div class="section-caption">{group["group_caption"]}</div>', unsafe_allow_html=True)

        n_cols = group.get("columns", 1)
        items = group["items"]

        if n_cols > 1 and len(items) > 1:
            cols = st.columns(n_cols)
            for i, item in enumerate(items):
                with cols[i % n_cols]:
                    uploaded_files[item["key"]] = render_single_uploader(item)
        else:
            for item in items:
                uploaded_files[item["key"]] = render_single_uploader(item)

    return uploaded_files


def render_file_status(uploaded_files):
    st.markdown("### 📊 สถานะไฟล์ที่อัปโหลด")

    all_items = get_all_items()
    file_count = sum(1 for item in all_items if uploaded_files.get(item["key"]) is not None)

    cols = st.columns(3)
    for i, item in enumerate(all_items):
        with cols[i % 3]:
            if uploaded_files.get(item["key"]) is not None:
                st.success(f"{item['title']}: ✅ อัปโหลดแล้ว")
            else:
                st.warning(f"{item['title']}: ⬜ ยังไม่อัปโหลด")

    st.markdown(f"### 📈 อัปโหลดแล้ว: **{file_count}** จาก **{len(all_items)}** ไฟล์")
    return file_count

# ═══════════════════════════════════════════════════════════════
# Main
# ═══════════════════════════════════════════════════════════════

def main():
    st.markdown('<div class="main-header">🛣️ Pavement Design Report Merger – 10 Files</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">โปรแกรมรวมรายงานออกแบบโครงสร้างชั้นทาง ตามโครงสร้างมาตรฐานที่กำหนด</div>', unsafe_allow_html=True)

    st.markdown("### 📋 ข้อมูลโครงการ")
    col1, col2 = st.columns(2)
    with col1:
        project_name = st.text_input("ชื่อโครงการ", placeholder="กรอกชื่อโครงการ")
    with col2:
        report_date = st.date_input("วันที่รายงาน", datetime.now())
        report_date_str = report_date.strftime("%d/%m/%Y")

    st.markdown("---")

    st.markdown("### 📁 อัปโหลดไฟล์รายงาน (ไม่จำเป็นต้องครบทุกไฟล์)")
    st.info("ระบบจะสร้างปก + สารบัญ และรวมเฉพาะไฟล์ที่อัปโหลด เรียงตามลำดับมาตรฐาน")

    uploaded_files = render_upload_sections()

    st.markdown("---")

    file_count = render_file_status(uploaded_files)

    st.markdown("---")

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        merge_button = st.button("🔄 รวมไฟล์และสร้างรายงาน", use_container_width=True)

    if merge_button:
        if file_count == 0:
            st.error("❌ กรุณาอัปโหลดไฟล์อย่างน้อย 1 ไฟล์")
        else:
            validation_errors = []
            all_items = get_all_items()
            for item in all_items:
                file = uploaded_files.get(item["key"])
                if file is not None:
                    is_valid, error_msg = validate_docx_file(file)
                    if not is_valid:
                        validation_errors.append(f"❌ **{item['title']}**: {error_msg}")

            if validation_errors:
                st.error("พบไฟล์ที่มีปัญหา กรุณาตรวจสอบและอัปโหลดใหม่:")
                for err in validation_errors:
                    st.markdown(err)
            else:
                progress_bar = st.progress(0, text="เริ่มต้นรวมไฟล์...")

                def update_progress(fraction, text):
                    progress_bar.progress(fraction, text=text)

                try:
                    merged_doc = merge_documents(
                        uploaded_files,
                        project_name,
                        report_date_str,
                        progress_callback=update_progress
                    )

                    progress_bar.progress(1.0, text="✅ รวมไฟล์เรียบร้อยแล้ว!")

                    with tempfile.TemporaryDirectory() as temp_dir:
                        base_filename = "รายงานออกแบบโครงสร้างชั้นทาง_10ไฟล์"
                        if project_name:
                            base_filename = f"รายงานออกแบบ_{project_name.replace(' ', '_')}_10ไฟล์"

                        docx_path = os.path.join(temp_dir, f"{base_filename}.docx")
                        merged_doc.save(docx_path)

                        st.markdown('<div class="success-box">', unsafe_allow_html=True)
                        st.success(f"✅ รวมไฟล์เรียบร้อยแล้ว! ({file_count} ไฟล์)")
                        st.markdown('</div>', unsafe_allow_html=True)

                        st.markdown("### 📥 ดาวน์โหลดรายงาน")

                        with open(docx_path, 'rb') as f:
                            docx_data = f.read()
                        st.download_button(
                            label="📄 ดาวน์โหลดไฟล์ Word (.docx)",
                            data=docx_data,
                            file_name=f"{base_filename}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )

                except Exception as e:
                    st.error(f"❌ เกิดข้อผิดพลาด: {str(e)}")
                    st.exception(e)

    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #718096; font-size: 14px;">
        <p>Pavement Design Report Merger – 10 Files Edition</p>
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
