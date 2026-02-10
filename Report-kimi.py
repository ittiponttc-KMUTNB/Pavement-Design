import streamlit as st
from docx import Document
from io import BytesIO

st.set_page_config(
    page_title="Civil Engineering Word Merger",
    layout="wide"
)

# ---------------- CUSTOM CSS ----------------
st.markdown("""
<style>
    body {
        background-color: #f2f4f7;
    }
    .main-header {
        text-align: center;
        font-size: 40px;
        font-weight: 800;
        color: #002b5c;
        margin-bottom: -5px;
    }
    .sub-header {
        text-align: center;
        font-size: 20px;
        color: #555;
        margin-bottom: 25px;
    }
    .upload-card {
        background: white;
        padding: 18px;
        border-radius: 12px;
        border: 1px solid #d0d7de;
        box-shadow: 0px 2px 6px rgba(0,0,0,0.08);
        margin-bottom: 15px;
    }
    .upload-title {
        font-size: 17px;
        font-weight: 600;
        color: #003366;
        margin-bottom: 8px;
    }
    .merge-btn {
        background-color: #003366 !important;
        color: white !important;
        font-size: 18px !important;
        padding: 12px 20px !important;
        border-radius: 10px !important;
    }
</style>
""", unsafe_allow_html=True)

# ---------------- HEADER ----------------
st.markdown('<div class="main-header">Civil Engineering Word Merger</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">รวมไฟล์ Word ตามหมวดงานทางวิศวกรรมโยธา</div>', unsafe_allow_html=True)
st.write("---")

# ---------------- FILE ORDER ----------------
file_labels = [
    "1. Truck Factor",
    "2.1 ESALs (Flexible)",
    "2.2 ESALs (Rigid)",
    "3. CBR Analysis",
    "4. AC Design",
    "5.1 JPCP/JRCP",
    "6.1 k-value (JPCP/JRCP)",
    "5.2 CRCP",
    "6.2 k-value (CRCP)",
    "7. Cost Estimate"
]

uploaded_files = {}

# ---------------- LAYOUT 2 COLUMNS ----------------
col1, col2 = st.columns(2)

for i, label in enumerate(file_labels):
    target_col = col1 if i % 2 == 0 else col2
    with target_col:
        st.markdown(f'<div class="upload-card"><div class="upload-title">{label}</div>', unsafe_allow_html=True)
        uploaded_files[label] = st.file_uploader("", type=["docx"], key=label)
        st.markdown("</div>", unsafe_allow_html=True)

# ---------------- STATUS ----------------
uploaded_count = sum(1 for f in uploaded_files.values() if f is not None)

st.write("---")
st.markdown(f"### สถานะการอัปโหลด: **{uploaded_count} / 10 ไฟล์**")
st.progress(uploaded_count / 10)

if uploaded_count == 0:
    st.warning("กรุณาอัปโหลดอย่างน้อย 1 ไฟล์ก่อน")
elif uploaded_count < 10:
    st.info("จะรวมเฉพาะไฟล์ที่อัปโหลดเท่านั้น")

# ---------------- MERGE FUNCTION ----------------
def merge_word_files(files_dict):
    merged_doc = Document()
    first = True

    for label in file_labels:
        file = files_dict[label]
        if file is None:
            continue

        doc = Document(file)

        if not first:
            merged_doc.add_page_break()
        first = False

        for element in doc.element.body:
            merged_doc.element.body.append(element)

    output = BytesIO()
    merged_doc.save(output)
    output.seek(0)
    return output

# ---------------- MERGE BUTTON ----------------
st.write("---")
st.markdown("### รวมไฟล์ Word")

if uploaded_count > 0:
    if st.button("รวมไฟล์ทั้งหมดที่อัปโหลด", key="merge", use_container_width=True):
        merged_output = merge_word_files(uploaded_files)
        st.success("รวมไฟล์สำเร็จ พร้อมดาวน์โหลด")

        st.download_button(
            label="ดาวน์โหลดไฟล์ที่รวมแล้ว (.docx)",
            data=merged_output,
            file_name="merged_files.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
