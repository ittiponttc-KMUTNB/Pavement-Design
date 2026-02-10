import streamlit as st
from docx import Document
from io import BytesIO

# ---------------- PAGE CONFIG ----------------
st.set_page_config(
    page_title="Civil Engineering Word Merger",
    layout="wide",
    page_icon="📘"
)

# ---------------- CUSTOM CSS ----------------
st.markdown("""
<style>
    .upload-card {
        padding: 15px;
        border: 1px solid #d0d7de;
        border-radius: 10px;
        background-color: #f8f9fa;
        margin-bottom: 12px;
    }
    .header-title {
        text-align: center;
        font-size: 36px;
        font-weight: 700;
        color: #003366;
        margin-bottom: -10px;
    }
    .sub-header {
        text-align: center;
        font-size: 18px;
        color: #555;
        margin-bottom: 25px;
    }
    .merge-button {
        background-color: #003366 !important;
        color: white !important;
        font-size: 18px !important;
        padding: 10px 20px !important;
        border-radius: 8px !important;
    }
</style>
""", unsafe_allow_html=True)

# ---------------- HEADER ----------------
st.markdown('<div class="header-title">Civil Engineering Word Merger</div>', unsafe_allow_html=True)
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

# ---------------- UPLOAD AREA ----------------
st.markdown("### 📁 อัปโหลดไฟล์ (ไม่จำเป็นต้องครบ 10 ไฟล์)")

for label in file_labels:
    with st.container():
        st.markdown(f'<div class="upload-card"><b>📄 {label}</b>', unsafe_allow_html=True)
        uploaded_files[label] = st.file_uploader("", type=["docx"], key=label)
        st.markdown("</div>", unsafe_allow_html=True)

# ---------------- STATUS ----------------
uploaded_count = sum(1 for f in uploaded_files.values() if f is not None)

st.write("---")
st.markdown(f"### 📊 สถานะการอัปโหลด: **{uploaded_count} / 10 ไฟล์**")
st.progress(uploaded_count / 10)

if uploaded_count == 0:
    st.warning("⚠️ กรุณาอัปโหลดอย่างน้อย 1 ไฟล์ก่อน")
elif uploaded_count < 10:
    st.info("ℹ️ จะรวมเฉพาะไฟล์ที่อัปโหลดเท่านั้น")

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
st.markdown("### 🛠️ รวมไฟล์ Word")

if uploaded_count > 0:
    if st.button("📘 รวมไฟล์ทั้งหมดที่อัปโหลด", key="merge", help="รวมไฟล์ Word ที่อัปโหลด", use_container_width=True):
        merged_output = merge_word_files(uploaded_files)
        st.success("🎉 รวมไฟล์สำเร็จ! พร้อมดาวน์โหลด")

        st.download_button(
            label="⬇️ ดาวน์โหลดไฟล์ที่รวมแล้ว (.docx)",
            data=merged_output,
            file_name="merged_files.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
