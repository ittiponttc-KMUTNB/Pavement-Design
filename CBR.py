import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')
import json
import base64
import io

# Check if python-docx is available
try:
    from docx import Document as DocxDocument
    from docx.shared import Inches, Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

st.set_page_config(
    page_title="CBR Percentile Analysis",
    page_icon="📊",
    layout="wide"
)

st.title("📊 การวิเคราะห์ค่า CBR ที่เปอร์เซ็นต์ไทล์")
st.markdown("### Subgrade CBR Analysis Tool")
st.markdown("---")

# Sample data (CBR values only)
sample_cbr = [14.8, 14.37, 5.31, 17.37, 5.48, 18.46, 4.85, 6.23,
              5.02, 10.78, 10.52, 14, 15.5, 8.7, 12.93, 8.19,
              8.1, 15.56, 16.88, 20.75, 20.3, 8, 7.84, 7.48,
              23.55, 8.92, 13.3, 13.5, 13.86, 7.18, 6.95, 5.8,
              6, 11.18, 9.69, 7.48]

# Sidebar for file upload
with st.sidebar:
    st.header("📁 อัปโหลดข้อมูล")
    
    # Upload JSON for settings
    st.markdown("#### 📂 โหลดการตั้งค่า")
    uploaded_json = st.file_uploader(
        "โหลดข้อมูลจากไฟล์ JSON",
        type=['json'],
        help="โหลดค่า Percentile และข้อมูล CBR จากไฟล์ JSON"
    )
    
    if uploaded_json is not None:
        try:
            loaded_data = json.load(uploaded_json)
            
            # ตรวจสอบว่าเป็นไฟล์ใหม่
            file_id = f"{uploaded_json.name}_{uploaded_json.size}"
            if st.session_state.get('last_uploaded_json') != file_id:
                st.session_state['last_uploaded_json'] = file_id
                
                # อัพเดท session_state
                if 'target_percentile' in loaded_data:
                    st.session_state['input_percentile'] = float(loaded_data['target_percentile'])
                if 'cbr_values' in loaded_data:
                    st.session_state['loaded_cbr_values'] = loaded_data['cbr_values']
                if 'use_sample' in loaded_data:
                    st.session_state['input_use_sample'] = loaded_data['use_sample']
                
                # โหลดค่า report settings ถ้ามี
                if 'report_settings' in loaded_data:
                    rs = loaded_data['report_settings']
                    if 'section_number' in rs:
                        st.session_state['input_section_number'] = rs['section_number']
                    if 'table_number' in rs:
                        st.session_state['input_table_number'] = rs['table_number']
                    if 'figure_number' in rs:
                        st.session_state['input_figure_number'] = rs['figure_number']
                    if 'section_title' in rs:
                        st.session_state['input_section_title'] = rs['section_title']
                    if 'table_caption' in rs:
                        st.session_state['input_table_caption'] = rs['table_caption']
                    if 'figure_caption' in rs:
                        st.session_state['input_figure_caption'] = rs['figure_caption']
                
                st.success("✅ โหลดการตั้งค่าสำเร็จ!")
                st.rerun()
                
        except Exception as e:
            st.error(f"❌ ไม่สามารถอ่านไฟล์ JSON ได้: {e}")
    
    st.markdown("---")
    
    # Upload Excel for CBR data
    st.markdown("#### 📊 อัปโหลดข้อมูล CBR")
    uploaded_file = st.file_uploader(
        "เลือกไฟล์ Excel (.xlsx)",
        type=['xlsx'],
        help="ไฟล์ควรมีคอลัมน์ CBR(%) เพียงคอลัมน์เดียว"
    )
    
    st.markdown("---")
    st.markdown("### 📋 รูปแบบข้อมูลที่ต้องการ")
    st.markdown("""
    | CBR(%) |
    |--------|
    | 14.8   |
    | 14.37  |
    | 5.31   |
    | ...    |
    """)
    st.info("ระบบจะคำนวณ Percentile ให้อัตโนมัติ")

# Process uploaded Excel file
if uploaded_file is not None:
    try:
        # Read Excel file
        df = pd.read_excel(uploaded_file)
        
        # Try to identify CBR column
        cbr_col = None
        
        for col in df.columns:
            col_lower = str(col).lower()
            if 'cbr' in col_lower:
                cbr_col = col
                break
        
        # If not found, use first column
        if cbr_col is None:
            cbr_col = df.columns[0]
        
        # Get CBR values
        cbr_values = pd.to_numeric(df[cbr_col], errors='coerce').dropna().tolist()
        
        st.success(f"✅ อ่านข้อมูลสำเร็จ: {len(cbr_values)} ตัวอย่าง")
        
    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาด: {str(e)}")
        st.info("กรุณาตรวจสอบรูปแบบไฟล์ Excel")
        cbr_values = None

elif 'loaded_cbr_values' in st.session_state and st.session_state['loaded_cbr_values']:
    # Use CBR values from loaded JSON
    cbr_values = st.session_state['loaded_cbr_values']
    st.info(f"📌 ใช้ข้อมูลจากไฟล์ JSON: {len(cbr_values)} ตัวอย่าง")

else:
    st.info("📌 กรุณาอัปโหลดไฟล์ Excel หรือใช้ข้อมูลตัวอย่าง")
    
    default_use_sample = st.session_state.get('input_use_sample', True)
    use_sample = st.checkbox(
        "ใช้ข้อมูลตัวอย่าง", 
        value=default_use_sample,
        key="input_use_sample"
    )
    
    if use_sample:
        cbr_values = sample_cbr
    else:
        cbr_values = None

if cbr_values is not None and len(cbr_values) > 0:
    
    # Sort CBR values
    cbr_sorted = np.sort(cbr_values)
    n = len(cbr_sorted)
    
    # Calculate cumulative percentile (percentage of values <= each CBR)
    cumulative_percentile = (np.arange(1, n + 1) / n) * 100
    
    # Create dataframe for display
    df_sorted = pd.DataFrame({
        'CBR': cbr_sorted,
        'Cumulative_Percentile': cumulative_percentile
    })
    
    # Create interpolation function using numpy
    def f_interp(x):
        return np.interp(x, cumulative_percentile, cbr_sorted)
    
    # Input percentile at the top
    st.markdown("### 🎯 กำหนดค่า Percentile")
    
    default_percentile = st.session_state.get('input_percentile', 90.0)
    target_percentile = st.number_input(
        "Percentile ที่ต้องการ (%)",
        min_value=0.0,
        max_value=100.0,
        value=default_percentile,
        step=1.0,
        help="ใส่ค่า Percentile ที่ต้องการหาค่า CBR",
        key="input_percentile"
    )
    
    # Calculate CBR at target percentile
    design_percentile = 100 - target_percentile
    
    if design_percentile >= cumulative_percentile.min() and \
       design_percentile <= cumulative_percentile.max():
        cbr_at_percentile = float(f_interp(design_percentile))
    else:
        cbr_at_percentile = float(f_interp(np.clip(design_percentile, 
                                                    cumulative_percentile.min(),
                                                    cumulative_percentile.max())))
    
    st.markdown("---")
    
    # =====================================================================
    # Report Numbering Settings (ผู้ใช้กรอกเอง - สีเหลือง)
    # =====================================================================
    st.markdown("### 📝 ตั้งค่าหัวข้อและเลขที่ สำหรับรายงาน Word")
    
    col_num1, col_num2, col_num3 = st.columns(3)
    
    with col_num1:
        section_number = st.text_input(
            "เลขหัวข้อ",
            value=st.session_state.get('input_section_number', "4.3"),
            key="input_section_number",
            help="เช่น 4.3, 5.1"
        )
    
    with col_num2:
        table_number = st.text_input(
            "เลขตาราง",
            value=st.session_state.get('input_table_number', "4-7"),
            key="input_table_number",
            help="เช่น 4-7, 5-1"
        )
    
    with col_num3:
        figure_number = st.text_input(
            "เลขรูป",
            value=st.session_state.get('input_figure_number', "4-7"),
            key="input_figure_number",
            help="เช่น 4-7, 5-1"
        )
    
    section_title = st.text_input(
        "ชื่อหัวข้อ",
        value=st.session_state.get('input_section_title', "ข้อมูลความแข็งแรงของดินฐานรากบริเวณพื้นที่โครงการ"),
        key="input_section_title",
        help="ชื่อหัวข้อในรายงาน"
    )
    
    col_cap1, col_cap2 = st.columns(2)
    
    with col_cap1:
        table_caption = st.text_input(
            "คำบรรยายตาราง",
            value=st.session_state.get('input_table_caption', "ค่าเปอร์เซ็นต์ไทล์ และค่า CBR ของตัวอย่างดินฐานรากตามแนวสายทาง"),
            key="input_table_caption",
            help="คำบรรยายใต้ตาราง"
        )
    
    with col_cap2:
        figure_caption = st.text_input(
            "คำบรรยายรูป",
            value=st.session_state.get('input_figure_caption', "กราฟแสดงความสัมพันธ์ระหว่าง Percentile และ CBR ของดินฐานรากตามแนวสายทาง"),
            key="input_figure_caption",
            help="คำบรรยายใต้รูป"
        )
    
    # =====================================================================
    # Preview introduction paragraph
    # =====================================================================
    st.markdown("---")
    st.markdown("### 👁️ ตัวอย่างข้อความเกริ่นนำในรายงาน")
    
    # Build preview - สีม่วงดึงจากข้อมูล, สีเหลืองผู้ใช้กรอก
    intro_preview = (
        f'<div style="font-family: TH SarabunPSK, Tahoma, sans-serif; font-size: 16px; line-height: 1.8; '
        f'background-color: #f9f9f9; padding: 15px; border-radius: 8px; border: 1px solid #ddd;">'
        f'<p style="margin-bottom: 5px;"><b>{section_number} &nbsp;&nbsp;&nbsp; {section_title}</b></p>'
        f'<p style="text-indent: 40px; text-align: justify;">'
        f'ความแข็งแรงของดินฐานรากบริเวณโดยรอบพื้นที่โครงการ หรือกำลังรับน้ำหนักของดินพื้นทางเดิม '
        f'หรือพื้นทางเดิมสามารถประเมินจากรายงานสำรวจภูมิประเทศของดิน ซึ่งสามารถทำการทดสอบได้หลากหลายวิธี เช่น Plate Bearing Test '
        f'CBR Test หรือ Modulus of Subgrade Reaction สำหรับการออกแบบถนนคอนกรีตนั้นใช้ค่า CBR ซึ่งนิยมใช้กันแพร่หลาย '
        f'เมื่อกำหนดกำลังรับน้ำหนักของดินพื้นทางเดิม โดยการเจาะสำรวจดินในสนามตามรายงานการสอบดินของ'
        f'ห้องปฏิบัติการ เพื่อหาค่า CBR ของดินพื้นทางเดินเพื่อเป็นข้อมูลในการออกแบบ ซึ่งผลการทดสอบค่า CBR ของ'
        f'ดินฐานรากตามแนวสายทาง จำนวน <span style="background-color: #D8B4FE; padding: 1px 4px; border-radius: 3px; font-weight: bold;">{n}</span> ตัวอย่าง '
        f'พบว่าที่เปอร์เซ็นต์ไทล์ ร้อยละ <span style="background-color: #D8B4FE; padding: 1px 4px; border-radius: 3px; font-weight: bold;">{target_percentile:.0f}</span> ของค่ากำลังที่พบเท่ากับ CBR '
        f'เท่ากับ <span style="background-color: #D8B4FE; padding: 1px 4px; border-radius: 3px; font-weight: bold;">{cbr_at_percentile:.1f}</span> % '
        f'ของจำนวนตัวอย่างทั้งหมด ที่มีกำลังรับแรงดันหรือค่า CBR เท่ากับ <span style="background-color: #D8B4FE; padding: 1px 4px; border-radius: 3px; font-weight: bold;">{int(cbr_at_percentile)}</span> % '
        f'นำไปใช้ในการออกแบบโครงสร้างถนน '
        f'ดังแสดงผลการวิเคราะห์ใน'
        f'<span style="background-color: #FDE68A; padding: 1px 4px; border-radius: 3px; font-weight: bold;">ตารางที่ {table_number}</span> '
        f'และ<span style="background-color: #FDE68A; padding: 1px 4px; border-radius: 3px; font-weight: bold;">รูปที่ {figure_number}</span></p>'
        f'</div>'
    )
    
    st.markdown(intro_preview, unsafe_allow_html=True)
    st.caption("🟣 สีม่วง = ดึงจากข้อมูลอัตโนมัติ | 🟡 สีเหลือง = ผู้ใช้กรอกเอง")
    
    st.markdown("---")
    
    # Graph section - full width
    st.markdown("### 📈 กราฟ Percentile vs CBR")
    
    # Create figure
    fig = go.Figure()
    
    # Calculate axis ranges
    x_max = max(cbr_sorted) * 1.1
    y_max = 100
    
    # Add main curve
    fig.add_trace(go.Scatter(
        x=cbr_sorted,
        y=100 - cumulative_percentile,  # Convert to "% >= value"
        mode='lines+markers',
        name='CBR Distribution',
        line=dict(color='blue', width=2),
        marker=dict(size=6, symbol='x', color='black')
    ))
    
    # Add horizontal red dashed line at target percentile
    fig.add_trace(go.Scatter(
        x=[0, cbr_at_percentile],
        y=[target_percentile, target_percentile],
        mode='lines',
        name=f'Percentile {target_percentile}%',
        line=dict(color='red', width=2, dash='dash')
    ))
    
    # Add vertical red dashed line at CBR value
    fig.add_trace(go.Scatter(
        x=[cbr_at_percentile, cbr_at_percentile],
        y=[0, target_percentile],
        mode='lines',
        name=f'CBR = {cbr_at_percentile:.2f}%',
        line=dict(color='red', width=2, dash='dash')
    ))
    
    # Add annotation for CBR value
    fig.add_annotation(
        x=cbr_at_percentile,
        y=0,
        text=f"<b>{cbr_at_percentile:.2f}</b>",
        showarrow=True,
        arrowhead=2,
        arrowsize=1,
        arrowwidth=2,
        arrowcolor='red',
        ax=0,
        ay=40,
        font=dict(size=16, color='red')
    )
    
    # Border line width (consistent for all 4 sides)
    border_width = 4
    
    # Update layout - remove axis lines, we'll draw border using shapes
    fig.update_layout(
        xaxis_title="CBR (%)",
        yaxis_title="Percentile (%)",
        xaxis=dict(
            range=[0, x_max],
            gridcolor='lightgray',
            showgrid=True,
            showline=False,  # Disable built-in axis line
            zeroline=False,
            ticks='outside',
            tickwidth=1,
            tickcolor='black',
            ticklen=5,
        ),
        yaxis=dict(
            range=[0, y_max],
            gridcolor='lightgray',
            showgrid=True,
            showline=False,  # Disable built-in axis line
            zeroline=False,
            ticks='outside',
            tickwidth=1,
            tickcolor='black',
            ticklen=5,
        ),
        plot_bgcolor='white',
        width=600,
        height=600,
        showlegend=True,
        legend=dict(
            yanchor="top",
            y=0.99,
            xanchor="right",
            x=0.99,
            bgcolor='rgba(255,255,255,0.8)',
            bordercolor='black',
            borderwidth=1
        ),
        title=dict(
            text=f"ค่าร้อยละ CBR ที่เปอร์เซ็นต์ไทล์ ร้อยละ {target_percentile:.0f}",
            x=0.5,
            xanchor='center'
        ),
        margin=dict(l=70, r=70, t=70, b=70)
    )
    
    # Draw complete border using a rectangle shape (ensures all 4 corners connect)
    fig.add_shape(
        type="rect",
        x0=0, y0=0,
        x1=x_max, y1=y_max,
        line=dict(color="black", width=border_width),
        xref="x", yref="y"
    )
    
    # Center the chart
    col_left, col_chart, col_right = st.columns([1, 2, 1])
    with col_chart:
        st.plotly_chart(fig, use_container_width=False)
    
    # Results section - below the graph
    st.markdown("---")
    
    col_result, col_stat = st.columns(2)
    
    with col_result:
        st.markdown("### 📊 ผลการวิเคราะห์")
        st.metric(
            label=f"CBR ที่ Percentile {target_percentile}%",
            value=f"{cbr_at_percentile:.2f} %"
        )
    
    with col_stat:
        st.markdown("### 📋 สถิติข้อมูล CBR")
        st.write(f"**จำนวนตัวอย่าง:** {n}")
        st.write(f"**ค่าต่ำสุด:** {np.min(cbr_values):.2f} %")
        st.write(f"**ค่าสูงสุด:** {np.max(cbr_values):.2f} %")
        st.write(f"**ค่าเฉลี่ย:** {np.mean(cbr_values):.2f} %")
        st.write(f"**ส่วนเบี่ยงเบนมาตรฐาน:** {np.std(cbr_values):.2f} %")
    
    # Export section
    st.markdown("---")
    st.markdown("### 💾 บันทึกข้อมูล")
    
    col_json, col_word = st.columns(2)
    
    with col_json:
        # Prepare export data for JSON
        export_data = {
            'target_percentile': target_percentile,
            'cbr_at_percentile': round(cbr_at_percentile, 2),
            'cbr_values': [float(v) for v in cbr_values],
            'statistics': {
                'n_samples': n,
                'min': round(float(np.min(cbr_values)), 2),
                'max': round(float(np.max(cbr_values)), 2),
                'mean': round(float(np.mean(cbr_values)), 2),
                'std': round(float(np.std(cbr_values)), 2)
            },
            'report_settings': {
                'section_number': section_number,
                'table_number': table_number,
                'figure_number': figure_number,
                'section_title': section_title,
                'table_caption': table_caption,
                'figure_caption': figure_caption,
            },
            'use_sample': st.session_state.get('input_use_sample', True)
        }
        
        json_str = json.dumps(export_data, ensure_ascii=False, indent=2)
        
        st.download_button(
            label="📥 Download JSON",
            data=json_str,
            file_name="cbr_percentile_data.json",
            mime="application/json",
            help="บันทึกข้อมูลและการตั้งค่าเป็นไฟล์ JSON"
        )
    
    with col_word:
        # Generate Word document using python-docx
        if DOCX_AVAILABLE:
            if st.button("📄 สร้างรายงาน Word", help="สร้างรายงานผลการวิเคราะห์เป็นไฟล์ Word"):
                try:
                    # Create Word document
                    doc = DocxDocument()
                    
                    # Set Thai font style
                    style = doc.styles['Normal']
                    style.font.name = 'TH SarabunPSK'
                    style.font.size = Pt(16)
                    style._element.rPr.rFonts.set(qn('w:eastAsia'), 'TH SarabunPSK')
                    
                    # Helper function to set cell background color
                    def set_cell_bg(cell, color_hex):
                        """Set background color for a cell using tcPr"""
                        tc = cell._tc
                        tcPr = tc.get_or_add_tcPr()
                        shading = OxmlElement('w:shd')
                        shading.set(qn('w:val'), 'clear')
                        shading.set(qn('w:color'), 'auto')
                        shading.set(qn('w:fill'), color_hex)
                        tcPr.append(shading)
                    
                    # =========================================================
                    # 1) Section heading: "4.3    ข้อมูลความแข็งแรง..."
                    # =========================================================
                    heading_para = doc.add_paragraph()
                    heading_run = heading_para.add_run(f'{section_number}\t{section_title}')
                    heading_run.font.name = 'TH SarabunPSK'
                    heading_run.font.size = Pt(16)
                    heading_run.font.bold = True
                    
                    # =========================================================
                    # 2) Introduction paragraph (เกริ่นนำ)
                    # =========================================================
                    intro_para = doc.add_paragraph()
                    intro_para.paragraph_format.first_line_indent = Cm(1.25)
                    
                    # Build intro text with mixed formatting
                    intro_parts = [
                        ('ความแข็งแรงของดินฐานรากบริเวณโดยรอบพื้นที่โครงการ หรือกำลังรับน้ำหนักของดินพื้นทางเดิม '
                         'หรือพื้นทางเดิมสามารถประเมินจากรายงานสำรวจภูมิประเทศของดิน ซึ่งสามารถทำการทดสอบได้หลากหลายวิธี เช่น Plate Bearing Test '
                         'CBR Test หรือ Modulus of Subgrade Reaction สำหรับการออกแบบถนนคอนกรีตนั้นใช้ค่า CBR ซึ่งนิยมใช้กันแพร่หลาย '
                         'เมื่อกำหนดกำลังรับน้ำหนักของดินพื้นทางเดิม โดยการเจาะสำรวจดินในสนามตามรายงานการสอบดินของ'
                         'ห้องปฏิบัติการ เพื่อหาค่า CBR ของดินพื้นทางเดินเพื่อเป็นข้อมูลในการออกแบบ ซึ่งผลการทดสอบค่า CBR ของ'
                         'ดินฐานรากตามแนวสายทาง จำนวน ', False),
                        (f'{n}', True),  # จำนวนตัวอย่าง - bold (ดึงจากข้อมูล)
                        (' ตัวอย่าง พบว่าที่เปอร์เซ็นต์ไทล์ ร้อยละ ', False),
                        (f'{target_percentile:.0f}', True),  # Percentile - bold (ดึงจากข้อมูล)
                        (' ของค่ากำลังที่พบเท่ากับ CBR เท่ากับ ', False),
                        (f'{cbr_at_percentile:.1f}', True),  # CBR value - bold (ดึงจากข้อมูล)
                        (' % ของจำนวนตัวอย่างทั้งหมด ที่มีกำลังรับแรงดันหรือค่า CBR เท่ากับ ', False),
                        (f'{int(cbr_at_percentile)}', True),  # CBR int - bold (ดึงจากข้อมูล)
                        (' % นำไปใช้ในการออกแบบโครงสร้างถนน ดังแสดงผลการวิเคราะห์ใน', False),
                        (f'ตารางที่ {table_number}', True),  # เลขตาราง - bold
                        (' และ', False),
                        (f'รูปที่ {figure_number}', True),  # เลขรูป - bold
                    ]
                    
                    for text, is_bold in intro_parts:
                        run = intro_para.add_run(text)
                        run.font.name = 'TH SarabunPSK'
                        run.font.size = Pt(16)
                        run.font.bold = is_bold
                    
                    doc.add_paragraph()  # spacing
                    
                    # =========================================================
                    # 3) TABLE (ตารางมาก่อนรูป)
                    # =========================================================
                    # Table caption above table
                    table_cap_para = doc.add_paragraph()
                    table_cap_run = table_cap_para.add_run(f'ตารางที่ {table_number} {table_caption}')
                    table_cap_run.font.name = 'TH SarabunPSK'
                    table_cap_run.font.size = Pt(16)
                    table_cap_run.font.bold = True
                    table_cap_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    # Calculate half point for splitting data
                    half_n = (n + 1) // 2
                    
                    # Create CBR data table with 6 columns (2 sets of 3 columns)
                    cbr_table = doc.add_table(rows=half_n+1, cols=6)
                    cbr_table.style = 'Table Grid'
                    cbr_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                    
                    # Header row
                    header_row = cbr_table.rows[0]
                    headers = ['ลำดับ', 'CBR (%)', 'Percentile (%)', 'ลำดับ', 'CBR (%)', 'Percentile (%)']
                    for j, header_text in enumerate(headers):
                        cell = header_row.cells[j]
                        run = cell.paragraphs[0].add_run(header_text)
                        run.font.name = 'TH SarabunPSK'
                        run.font.size = Pt(14)
                        run.font.bold = True
                        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        set_cell_bg(cell, 'D9E2F3')  # light blue header
                    
                    # Data rows - split into left and right halves
                    for i in range(half_n):
                        row = cbr_table.rows[i+1]
                        
                        # Left side data (first half)
                        left_idx = i
                        if left_idx < n:
                            cbr_val_left = cbr_sorted[left_idx]
                            pct_val_left = 100 - cumulative_percentile[left_idx]
                            left_data = [f'{left_idx+1}', f'{cbr_val_left:.2f}', f'{pct_val_left:.2f}']
                        else:
                            left_data = ['', '', '']
                        
                        # Right side data (second half)
                        right_idx = i + half_n
                        if right_idx < n:
                            cbr_val_right = cbr_sorted[right_idx]
                            pct_val_right = 100 - cumulative_percentile[right_idx]
                            right_data = [f'{right_idx+1}', f'{cbr_val_right:.2f}', f'{pct_val_right:.2f}']
                        else:
                            right_data = ['', '', '']
                        
                        # Fill left side (columns 0-2)
                        for j, val in enumerate(left_data):
                            cell = row.cells[j]
                            run = cell.paragraphs[0].add_run(val)
                            run.font.name = 'TH SarabunPSK'
                            run.font.size = Pt(14)
                            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        
                        # Fill right side (columns 3-5)
                        for j, val in enumerate(right_data):
                            cell = row.cells[j+3]
                            run = cell.paragraphs[0].add_run(val)
                            run.font.name = 'TH SarabunPSK'
                            run.font.size = Pt(14)
                            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    # Set column widths for CBR table
                    for row in cbr_table.rows:
                        row.cells[0].width = Cm(1.5)
                        row.cells[1].width = Cm(2.5)
                        row.cells[2].width = Cm(2.5)
                        row.cells[3].width = Cm(1.5)
                        row.cells[4].width = Cm(2.5)
                        row.cells[5].width = Cm(2.5)
                    
                    doc.add_paragraph()  # spacing
                    
                    # =========================================================
                    # 4) Statistics summary table
                    # =========================================================
                    h_stat = doc.add_paragraph()
                    h_stat_run = h_stat.add_run('ผลการวิเคราะห์')
                    h_stat_run.font.name = 'TH SarabunPSK'
                    h_stat_run.font.size = Pt(16)
                    h_stat_run.font.bold = True
                    
                    # Create statistics table
                    table = doc.add_table(rows=7, cols=2)
                    table.style = 'Table Grid'
                    table.alignment = WD_TABLE_ALIGNMENT.CENTER
                    
                    # Table data
                    table_data = [
                        ('รายการ', 'ค่า'),
                        ('จำนวนตัวอย่าง', f'{n}'),
                        ('ค่าต่ำสุด', f'{np.min(cbr_values):.2f} %'),
                        ('ค่าสูงสุด', f'{np.max(cbr_values):.2f} %'),
                        ('ค่าเฉลี่ย', f'{np.mean(cbr_values):.2f} %'),
                        ('ส่วนเบี่ยงเบนมาตรฐาน', f'{np.std(cbr_values):.2f} %'),
                        (f'CBR ที่ Percentile {target_percentile:.0f}%', f'{cbr_at_percentile:.2f} %')
                    ]
                    
                    for i, (col1, col2) in enumerate(table_data):
                        row = table.rows[i]
                        cell1 = row.cells[0]
                        cell2 = row.cells[1]
                        
                        run1 = cell1.paragraphs[0].add_run(col1)
                        run1.font.name = 'TH SarabunPSK'
                        run1.font.size = Pt(14)
                        if i == 0:
                            run1.font.bold = True
                            set_cell_bg(cell1, 'D9E2F3')
                            set_cell_bg(cell2, 'D9E2F3')
                        
                        run2 = cell2.paragraphs[0].add_run(col2)
                        run2.font.name = 'TH SarabunPSK'
                        run2.font.size = Pt(14)
                        if i == 0:
                            run2.font.bold = True
                        if i == 6:  # Last row - CBR result
                            run2.font.bold = True
                            run2.font.color.rgb = RGBColor(255, 0, 0)
                    
                    # Set column widths
                    for row in table.rows:
                        row.cells[0].width = Cm(6)
                        row.cells[1].width = Cm(4)
                    
                    doc.add_paragraph()  # spacing
                    
                    # =========================================================
                    # 5) FIGURE (รูปมาหลังตาราง)
                    # =========================================================
                    # Create chart using matplotlib
                    fig_mpl, ax = plt.subplots(figsize=(6, 6))
                    
                    # Plot main curve
                    y_plot = 100 - cumulative_percentile
                    ax.plot(cbr_sorted, y_plot, 'b-', linewidth=2, marker='x', 
                           markersize=6, markerfacecolor='black', markeredgecolor='black',
                           label='CBR Distribution')
                    
                    # Plot dashed lines
                    ax.plot([0, cbr_at_percentile], [target_percentile, target_percentile], 
                           'r--', linewidth=2, label=f'Percentile {target_percentile}%')
                    ax.plot([cbr_at_percentile, cbr_at_percentile], [0, target_percentile], 
                           'r--', linewidth=2, label=f'CBR = {cbr_at_percentile:.2f}%')
                    
                    # Annotation
                    ax.annotate(f'{cbr_at_percentile:.2f}', 
                               xy=(cbr_at_percentile, 0), 
                               xytext=(cbr_at_percentile, -8),
                               fontsize=12, color='red', fontweight='bold',
                               ha='center')
                    
                    ax.set_xlim(0, max(cbr_sorted) * 1.1)
                    ax.set_ylim(0, 100)
                    ax.set_xlabel('CBR (%)', fontsize=12)
                    ax.set_ylabel('Percentile (%)', fontsize=12)
                    ax.set_title(f'CBR at Percentile {target_percentile:.0f}%', fontsize=14)
                    ax.legend(loc='upper right', fontsize=10)
                    ax.grid(True, alpha=0.3)
                    
                    # Set border
                    for spine in ax.spines.values():
                        spine.set_linewidth(2)
                        spine.set_color('black')
                    
                    plt.tight_layout()
                    
                    # Save chart to buffer
                    chart_buffer = io.BytesIO()
                    fig_mpl.savefig(chart_buffer, format='png', dpi=150, 
                                   bbox_inches='tight', facecolor='white', edgecolor='none')
                    chart_buffer.seek(0)
                    plt.close(fig_mpl)
                    
                    # Add chart image to document
                    chart_para = doc.add_paragraph()
                    chart_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    chart_run = chart_para.add_run()
                    chart_run.add_picture(chart_buffer, width=Cm(12))
                    
                    # Add figure caption
                    caption = doc.add_paragraph()
                    caption_run = caption.add_run(f'รูปที่ {figure_number} {figure_caption}')
                    caption_run.font.name = 'TH SarabunPSK'
                    caption_run.font.size = Pt(16)
                    caption_run.font.bold = True
                    caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    # =========================================================
                    # 6) Footer
                    # =========================================================
                    doc.add_paragraph()
                    doc.add_paragraph()
                    footer1 = doc.add_paragraph()
                    footer1_run = footer1.add_run('พัฒนาโดย รศ.ดร.อิทธิพล มีผล')
                    footer1_run.font.name = 'TH SarabunPSK'
                    footer1_run.font.size = Pt(14)
                    footer1_run.font.italic = True
                    footer1.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    footer2 = doc.add_paragraph()
                    footer2_run = footer2.add_run('ภาควิชาครุศาสตร์โยธา คณะครุศาสตร์อุตสาหกรรม มจพ.')
                    footer2_run.font.name = 'TH SarabunPSK'
                    footer2_run.font.size = Pt(14)
                    footer2_run.font.italic = True
                    footer2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    # Save to buffer
                    buffer = io.BytesIO()
                    doc.save(buffer)
                    buffer.seek(0)
                    
                    st.download_button(
                        label="📥 Download Word",
                        data=buffer,
                        file_name="cbr_percentile_report.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    st.success("✅ สร้างรายงาน Word สำเร็จ!")
                    
                except Exception as e:
                    st.error(f"❌ ไม่สามารถสร้างรายงาน Word ได้: {e}")
        else:
            st.warning("⚠️ ต้องติดตั้ง python-docx เพื่อใช้งานฟีเจอร์นี้")
            st.code("pip install python-docx", language="bash")
    
    # Show data table
    st.markdown("---")
    st.markdown("### 📋 ตารางข้อมูล (เรียงตาม CBR)")
    
    # Create display table with calculated percentile
    df_display = pd.DataFrame({
        'ลำดับ': range(1, n + 1),
        'CBR (%)': cbr_sorted,
        'Percentile (%)': np.round(100 - cumulative_percentile, 2)
    })
    
    col_a, col_b = st.columns(2)
    
    with col_a:
        st.dataframe(
            df_display.head(len(df_display)//2 + 1),
            use_container_width=True,
            hide_index=True
        )
    
    with col_b:
        st.dataframe(
            df_display.tail(len(df_display)//2),
            use_container_width=True,
            hide_index=True
        )

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: gray;'>
    <p>สำหรับการวิเคราะห์ค่า CBR ดินฐานรากตามแนวสายทาง</p>
    <p>พัฒนาโดย รศ.ดร.อิทธิพล มีผล // ภาควิชาครุศาสตร์โยธา // คณะครุศาสตร์อุตสาหกรรม // มจพ.</p>
</div>
""", unsafe_allow_html=True)
