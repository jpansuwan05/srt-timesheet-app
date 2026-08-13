import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.worksheet.properties import PageSetupProperties
import io
import datetime
import re

st.set_page_config(page_title="SRT Timesheet App", layout="wide")
st.title("🚂 ระบบจัดการเวรและใบเบิกค่าตอบแทน (รฟท.)")

# ==========================================
# 0. ระบบสำรองข้อมูล (Backup & Restore)
# ==========================================
st.sidebar.subheader("💾 จัดการข้อมูลสำรอง")
if st.sidebar.button("📥 ดาวน์โหลดข้อมูลตารางเวร (.xlsx)"):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        st.session_state.roster_df.to_excel(writer, index=False, sheet_name='Roster')
    st.sidebar.download_button(label="คลิกเพื่อดาวน์โหลด", data=buffer, file_name="backup_roster.xlsx", mime="application/vnd.ms-excel")

uploaded_file = st.sidebar.file_uploader("📤 อัปโหลดไฟล์สำรอง (.xlsx) เพื่อทำงานต่อ", type=["xlsx"])
if uploaded_file:
    if st.sidebar.button("ยืนยันโหลดข้อมูล"):
        st.session_state.roster_df = pd.read_excel(uploaded_file)
        st.sidebar.success("โหลดข้อมูลสำเร็จ!")
        st.rerun()

# ==========================================
# 1. โหลดข้อมูลพนักงาน
# ==========================================
@st.cache_data
def load_employee_data():
    try:
        df = pd.read_excel("ข้อมูล.xlsx")
        emp_dict = {}
        for _, row in df.iterrows():
            name, pos = str(row["รายชื่อ"]).strip(), str(row["ตำแหน่ง"]).strip()
            p_clean = pos.replace(" ", "")
            role = pos
            if "นายสถานี" in p_clean: role = "นสน."
            elif "ช.นสน.ตช" in p_clean: role = "ช.นสน.1"
            elif "ช.นสน.ตค" in p_clean: role = "ช.นสน.2"
            elif "เสมียน" in p_clean: role = "เสมียน"
            elif "ประแจ" in p_clean: role = "ประแจ"
            elif "ฉิมพลี" in p_clean: role = "กั้นถนนฯฉิมพลี"
            elif "บางระมาด" in p_clean: role = "กั้นถนนฯบางระมาด"
            unique_key = f"{name}_{role}"
            emp_dict[unique_key] = {
                "ชื่อ-สกุล": name, "ตำแหน่ง": pos, "เลขประจำตัว": str(row["เลขประจำตัว"]),
                "เงินเดือน": str(row["เงินเดือน"]), "เรท": float(row["เรท 1 ชั่วโมง"]),
                "ประเภทบัญชี": str(row["ประเภทบัญชี"]), "รหัสบัญชี": str(row["รหัสบัญชี"]),
                "Role": role, "is_regular": True 
            }
        return emp_dict
    except: return None

if 'employees' not in st.session_state:
    st.session_state.employees = load_employee_data()
    data = []
    for i, (key, info) in enumerate(st.session_state.employees.items()):
        row = {"ขึ้นหน้าใหม่": False, "ลำดับ": i+1, "ชื่อ-สกุล": info['ชื่อ-สกุล'], "ตำแหน่งเบิก": info['ตำแหน่ง'], "Role (หน้าที่)": info['Role']}
        for d in range(1, 32): row[str(d)] = ""
        data.append(row)
    st.session_state.roster_df = pd.DataFrame(data)

# ==========================================
# 2. ข้อมูลส่วนกลาง และ ตาราง
# ==========================================
st.subheader("⚙️ 1. ตั้งค่าข้อมูลส่วนกลาง")
col_g1, col_g2 = st.columns(2)
global_data = {
    "val_13": col_g1.text_input("เดือนตัวเต็ม [13]", "สิงหาคม"),
    "val_7": col_g1.text_input("คำสั่งแขวง [7]", "5110/2520/2569"),
    "val_8": col_g2.text_input("วันที่ลงคำสั่ง [8]", "29 พ.ค. 69"),
    "val_14": col_g2.text_input("วันที่เซ็นเอกสารตัวย่อ [14]", "01 ก.ค. 69")
}

st.markdown("---")
st.subheader("🗓️ 2. จัดการตารางเวร")

# โค้ดส่วนการจัดการตาราง (Bulk Fill & Editor) เหมือนเดิม...
# (เพื่อความกระชับ โค้ดส่วนนี้ทำงานต่อเนื่องตามที่เคยตกลงกันไว้)
# [สรุปคือวางบล็อกตารางเวร และ Validation เดิมไว้ตรงนี้ครับ]

# ตรวจสอบความถูกต้อง
def validate_roster(roster_df):
    errors = []
    # (โค้ด Validate เดิม...)
    return errors

if st.button("🔍 ตรวจสอบความถูกต้อง", type="secondary"):
    errors = validate_roster(st.session_state.roster_df)
    if not errors: st.success("🎉 ตารางเวรถูกต้อง!")
    else:
        for e in errors: st.warning(e)

# ฟังก์ชัน Export 109 และ 177 เหมือนเดิม...
# [วางโค้ด generate_109 และ generate_177 และส่วน tab Export ไว้ท้ายสุดครับ]
