import streamlit as st
from docx import Document
from io import BytesIO

st.set_page_config(page_title="Word Merger 10 Files", layout="centered")

st.title("📄 โปรแกรมรวมไฟล์ Word (10 ไฟล์ตามรายการ)")

st.write("อัปโหลดไฟล์ให้ครบทั้ง 10 ไฟล์ แล้วกดปุ่มรวมไฟล์")

# ลำดับไฟล์ใหม่ตามที่อาจารย์ต้องการ
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

# UI อัปโหลดไฟล์ทีละรายการ
for label in file_labels:
    uploaded_files[label] = st.file_uploader(f"{label}", type=["docx"])

# ฟังก์ชันรวมไฟล์
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

# ตรวจสอบจำนวนไฟล์ที่อัปโหลดแล้ว
uploaded_count = sum(1 for f in uploaded_files.values() if f is not None)
st.write(f"📌 อัปโหลดแล้ว: {uploaded_count} จาก 10 ไฟล์")

# ปุ่มรวมไฟล์
if uploaded_count == 10:
    if st.button("รวมไฟล์ Word ทั้ง 10 ไฟล์"):
        merged_output = merge_word_files(uploaded_files)

        st.download_button(
            label="⬇️ ดาวน์โหลดไฟล์ที่รวมแล้ว (.docx)",
            data=merged_output,
            file_name="merged_10_files.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
else:
    st.warning("กรุณาอัปโหลดไฟล์ให้ครบทั้ง 10 ไฟล์ก่อนจึงจะรวมได้")
