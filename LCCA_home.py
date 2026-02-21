#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import streamlit as st

st.set_page_config(
    page_title="ระบบวิเคราะห์ทางหลวง | KMUTNB",
    page_icon="🛣️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
.hero-box {
    background: linear-gradient(135deg, #1E3A5F 0%, #2E6DAD 100%);
    padding: 2.5rem 2rem; border-radius: 16px;
    color: white; text-align: center; margin-bottom: 2rem;
}
.hero-box h1 { font-size: 2rem; margin: 0 0 0.4rem 0; }
.hero-box p  { font-size: 0.95rem; opacity: 0.85; margin: 0; }

.wf-card {
    background: white;
    border: 2px solid #2E6DAD;
    border-radius: 14px;
    padding: 1.4rem 1.2rem;
    text-align: center;
    box-shadow: 0 3px 12px rgba(46,109,173,0.12);
}
.wf-num {
    background: #1E3A5F; color: white;
    border-radius: 50%; width: 40px; height: 40px;
    display: inline-flex; align-items: center; justify-content: center;
    font-weight: bold; font-size: 1.2rem; margin-bottom: 0.6rem;
}
.wf-icon  { font-size: 2rem; display:block; margin-bottom: 0.3rem; }
.wf-title { font-size: 1rem; font-weight: 700; color: #1E3A5F; margin: 0 0 0.5rem 0; }
.wf-desc  { font-size: 0.82rem; color: #555; line-height: 1.5; margin: 0 0 0.7rem 0; }
.wf-out   { background: #EBF4FF; border-radius: 6px; padding: 0.35rem 0.6rem;
            font-size: 0.78rem; color: #1E3A5F; display:inline-block; }

.arrow-col { display:flex; align-items:center; justify-content:center;
             font-size: 2.2rem; color: #2E6DAD; padding: 0; }

.status-box { background:#f0f8e8; border-left:4px solid #4CAF50;
              border-radius:6px; padding:0.75rem 1rem; margin-top:0.4rem; font-size:0.9rem; }
.guide-step { display:flex; align-items:flex-start; gap:0.75rem;
              padding:0.6rem 0; border-bottom:1px solid #f0f0f0; }
.guide-num  { background:#1E3A5F; color:white; border-radius:50%;
              width:28px; height:28px; min-width:28px;
              display:flex; align-items:center; justify-content:center;
              font-size:0.85rem; font-weight:bold; }
</style>
""", unsafe_allow_html=True)

# ─── Hero ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero-box">
  <h1>🛣️ ระบบโปรแกรมวิเคราะห์ทางหลวง</h1>
  <p>Pavement Engineering Analysis System · KMUTNB</p>
  <p style="margin-top:0.4rem; font-size:0.85rem; opacity:0.75;">
    รศ.ดร.อิทธิพล มีผล &nbsp;|&nbsp; ภาควิชาครุศาสตร์โยธา &nbsp;|&nbsp;
    มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าพระนครเหนือ
  </p>
</div>
""", unsafe_allow_html=True)

# ─── Workflow Diagram ─────────────────────────────────────────────────────────
st.subheader("🔄 Workflow การใช้งาน")

c1, ca, c2, cb, c3 = st.columns([4, 0.8, 4, 0.8, 4])

with c1:
    st.markdown("""
    <div class="wf-card">
      <div class="wf-num">1</div>
      <span class="wf-icon">💰</span>
      <p class="wf-title">ราคาโครงสร้างชั้นทาง</p>
      <p class="wf-desc">
        คำนวณต้นทุนก่อสร้างจากวัสดุแต่ละชั้น<br>
        <b>AC / JPCP / JRCP / CRCP</b><br>
        อ้างอิงราคากรมทางหลวง
      </p>
      <span class="wf-out">📤 ส่ง: ต้นทุนก่อสร้าง (บาท/ตร.ม.)</span>
    </div>
    """, unsafe_allow_html=True)

with ca:
    st.markdown('<div class="arrow-col">⇒</div>', unsafe_allow_html=True)

with c2:
    st.markdown("""
    <div class="wf-card">
      <div class="wf-num">2</div>
      <span class="wf-icon">🔧</span>
      <p class="wf-title">ค่าบำรุงรักษาทางหลวง</p>
      <p class="wf-desc">
        คำนวณปริมาณงานและงบประมาณ<br>
        บำรุงรักษาปกติตามสูตร <b>กทช. แบบ A</b><br>
        แอสฟัลท์ / ลูกรัง / คอนกรีต
      </p>
      <span class="wf-out">📤 ส่ง: ค่าบำรุงรักษา (บาท/ตร.ม./ปี)</span>
    </div>
    """, unsafe_allow_html=True)

with cb:
    st.markdown('<div class="arrow-col">⇒</div>', unsafe_allow_html=True)

with c3:
    st.markdown("""
    <div class="wf-card">
      <div class="wf-num">3</div>
      <span class="wf-icon">📊</span>
      <p class="wf-title">LCCA วิเคราะห์ต้นทุน</p>
      <p class="wf-desc">
        วิเคราะห์ต้นทุนตลอดอายุการใช้งาน<br>
        <b>Present Worth / EAC</b><br>
        Sensitivity Analysis
      </p>
      <span class="wf-out">📥 รับ: ต้นทุน + ค่าบำรุงรักษา</span>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ─── สถานะข้อมูลที่รอส่ง ──────────────────────────────────────────────────────
has_cost  = 'cost_to_lcca'  in st.session_state
has_maint = 'maintenance_to_lcca' in st.session_state

if has_cost or has_maint:
    st.subheader("📬 ข้อมูลที่รอส่งไปยัง LCCA")
    if has_cost:
        n = len(st.session_state['cost_to_lcca'])
        st.markdown(f'<div class="status-box">✅ ต้นทุนก่อสร้าง — {n} ทางเลือก รอการยืนยัน</div>',
                    unsafe_allow_html=True)
    if has_maint:
        n = len(st.session_state['maintenance_to_lcca'])
        st.markdown(f'<div class="status-box">✅ ค่าบำรุงรักษา — {n} สายทาง รอการยืนยัน</div>',
                    unsafe_allow_html=True)
    st.info("💡 เปิดหน้า **3 · LCCA** (เมนูด้านซ้าย) แล้วระบบจะถามก่อนนำค่าไปใช้")
    st.markdown("<br>", unsafe_allow_html=True)

# ─── คู่มือการใช้งาน ──────────────────────────────────────────────────────────
st.divider()
with st.expander("📖 คู่มือการใช้งานแบบย่อ", expanded=False):
    st.markdown("""
    <div class="guide-step">
      <div class="guide-num">1</div>
      <div><b>เปิดโปรแกรม 1 (ราคาโครงสร้างชั้นทาง)</b><br>
           กรอกข้อมูลวัสดุแต่ละชั้น → กดคำนวณ → กด <b>"ส่งต้นทุนก่อสร้างไป LCCA"</b></div>
    </div>
    <div class="guide-step">
      <div class="guide-num">2</div>
      <div><b>เปิดโปรแกรม 2 (ค่าบำรุงรักษา)</b><br>
           กรอกข้อมูลสายทาง → กดคำนวณ → กด <b>"ส่งค่าบำรุงรักษาไป LCCA"</b></div>
    </div>
    <div class="guide-step">
      <div class="guide-num">3</div>
      <div><b>เปิดโปรแกรม 3 (LCCA)</b><br>
           ระบบจะแสดง <b>preview ข้อมูลที่จะรับ</b> พร้อมปุ่ม ✅ รับ / ❌ ไม่รับ</div>
    </div>
    <div class="guide-step" style="border:none;">
      <div class="guide-num">4</div>
      <div><b>กด ✅ รับค่าและนำไปใช้</b><br>
           LCCA จะอัปเดตต้นทุนก่อสร้างและแผนบำรุงรักษาอัตโนมัติ แล้ววิเคราะห์ได้เลย</div>
    </div>
    """, unsafe_allow_html=True)

# ─── ลิงก์เปิดโปรแกรม ────────────────────────────────────────────────────────
st.subheader("🚀 เปิดโปรแกรม")
b1, b2, b3 = st.columns(3)
with b1:
    st.page_link("pages/1_Cost_Structure.py",
                 label="1 · ราคาโครงสร้างชั้นทาง", icon="💰", use_container_width=True)
with b2:
    st.page_link("pages/2_Maintenance_Cost.py",
                 label="2 · ค่าบำรุงรักษาทางหลวง", icon="🔧", use_container_width=True)
with b3:
    st.page_link("pages/3_LCCA.py",
                 label="3 · LCCA วิเคราะห์ต้นทุน", icon="📊", use_container_width=True)

# ─── Footer ──────────────────────────────────────────────────────────────────
st.divider()
st.markdown("""
<div style='text-align:center; color:#aaa; font-size:0.82rem; padding:0.5rem 0 1.5rem;'>
  ระบบโปรแกรมวิเคราะห์ทางหลวง v1.0 &nbsp;|&nbsp;
  ภาควิชาครุศาสตร์โยธา คณะครุศาสตร์อุตสาหกรรม มจพ.<br>
  สงวนลิขสิทธิ์ © 2568 รศ.ดร.อิทธิพล มีผล
</div>
""", unsafe_allow_html=True)
