import streamlit as st
from docx import Document
from io import BytesIO

st.set_page_config(
    page_title="Word Merger 10 Files",
    layout="centered",
    page_icon="📄"
)

# ---------------- HEADER ----------------
st.markdown("""
<div style="text-align:center;">
    <h1>📄 Word Merger – Civil Engineering</h1>
    <h3 style="color:#555;">รวมไฟล์ Word ตามหมวดงานทางวิศวกรรมโยธา</h3>
</div>
""", unsafe_allow_html=True)

st.write("---")

st.markdown("### 📁 อัปโหลดไฟล์ (ไม่จำเป็นต้องครบ 10 ไฟล์)")

# ---------------- FILE ORDER ----------------
file_labels = [
    "1. Truck Factor",
    "2.1 ESALs (Flexible)",
    "2.2 ESALs (Rigid)",
    "3. CBR Analysis",
    "4. AC Design",
    "5.1 JPCP/JRCP",
    "6.1 k-value (JPCP/JRCP)",   # moved here
    "5.2 CRCP",
    "6.2 k-value (CRCP)",        # moved here
    "7. Cost Estimate"
]

uploaded_files = {}

# ---------------- UPLOAD AREA ----------------
for label in file_labels:
    with st.container():
        st.markdown(f"**📄 {label}**")
        uploaded_files[label] = st.file_uploader("", type=["docx"], key=label)

# ---------------- COUNT ----------------
uploaded_count = sum(1 for f in uploaded_files.values() if f is not None)

st.write("---")
st.markdown(f"### 📊 สถานะการอัปโหลด: **{uploaded_count} / 10 ไฟล์**")

st.progress(uploaded_count / 10)

if uploaded_count == 0:
    st.warning("⚠️ กรุณาอัปโหลดอย่างน้อย 1 ไฟล์ก่อน")
elif uploaded_count < 10:
    st.info("ℹ️ จะรวมเฉพาะไฟล์ที่อัปโหลดเท่านั้น (ยังไม่ครบ 10 ไฟล์)")

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
    if st.button("📎 รวมไฟล์ทั้งหมดที่อัปโหลด"):
        merged_output = merge_word_files(uploaded_files)

        st.success("🎉 รวมไฟล์สำเร็จ! พร้อมดาวน์โหลด")

        st.download_button(
            label="⬇️ ดาวน์โหลดไฟล์ที่รวมแล้ว (.docx)",
            data=merged_output,
            file_name="merged_files.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
