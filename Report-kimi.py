# ติดตั้ง: pip install docxcompose
from docxcompose.composer import Composer

def merge_with_composer(sections_list, files_dict, project, date_str):
    """ใช้ Composer รวมไฟล์ (วิธีที่เชื่อถือได้)"""
    # สร้างเอกสารหลัก
    master = Document()
    set_a4_margins(master.sections[0])
    create_cover_page(master, project, date_str)
    create_toc(master, sections_list)
    
    # บันทึกชั่วคราว
    master_io = io.BytesIO()
    master.save(master_io)
    master_io.seek(0)
    
    composer = Composer(Document(master_io))
    
    for i, section in enumerate(sections_list, 1):
        # เพิ่มหัวข้อ
        header_doc = Document()
        p = header_doc.add_paragraph()
        run = p.add_run(f"{i}. {section.title}")
        set_thai_font(run, size=20, bold=True)
        header_doc.add_paragraph()
        
        h_io = io.BytesIO()
        header_doc.save(h_io)
        h_io.seek(0)
        composer.append(Document(h_io))
        
        # เพิ่มเนื้อหา
        file_bytes = files_dict.get(section.id)
        if file_bytes:
            composer.append(Document(io.BytesIO(file_bytes)))
    
    return composer.doc
