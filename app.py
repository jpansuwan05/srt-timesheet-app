import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.worksheet.properties import PageSetupProperties
import io
import datetime
import re
import os

st.set_page_config(page_title="SRT Timesheet App", layout="wide")
st.title("🚂 ระบบจัดการเวรและใบเบิกค่าตอบแทน (รฟท.)")
st.markdown("---")

# ==========================================
# 0. ระบบ Save/Load ข้อมูลอัตโนมัติ
# ==========================================
AUTOSAVE_FILE = "autosave_data.xlsx"

def save_data():
    with pd.ExcelWriter(AUTOSAVE_FILE) as writer:
        st.session_state.roster_df.to_excel(writer, sheet_name="roster", index=False)
        # เซฟ ind_data แยกชีต
        ind_df = pd.DataFrame.from_dict(st.session_state, orient='index').T
        ind_df.to_excel(writer, sheet_name="ind_data", index=False)

def load_saved_data():
    if os.path.exists(AUTOSAVE_FILE):
        try:
            st.session_state.roster_df = pd.read_excel(AUTOSAVE_FILE, sheet_name="roster")
            # โหลดสถานะ session กลับมาถ้าจำเป็น
            return True
        except: return False
    return False

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
    if not load_saved_data():
        # ถ้าไม่มีไฟล์เซฟ ให้โหลดค่าเริ่มต้น
        data = []
        for i, (key, info) in enumerate(st.session_state.employees.items()):
            row = {"ขึ้นหน้าใหม่": False, "ลำดับ": i+1, "ชื่อ-สกุล": info['ชื่อ-สกุล'], "ตำแหน่งเบิก": info['ตำแหน่ง'], "Role (หน้าที่)": info['Role']}
            for d in range(1, 32): row[str(d)] = ""
            data.append(row)
        st.session_state.roster_df = pd.DataFrame(data)

# ล้างค่าของเดือนใหม่
st.sidebar.warning("⚙️ หากต้องการเริ่มเดือนใหม่ ข้อมูลเดิมจะถูกล้าง")
if st.sidebar.button("🗑️ เริ่มเดือนใหม่ (ล้างข้อมูล)"):
    if os.path.exists(AUTOSAVE_FILE): os.remove(AUTOSAVE_FILE)
    st.rerun()

# ==========================================
# 2-3. หน้าจอ UI และตาราง (เรียกใช้ save_data ทุกครั้งที่อัปเดตตาราง)
# ==========================================
# ... (โค้ดการเพิ่มพนักงานและการกรอกเวรเหมือนเดิม)
# สำคัญ: ทุกครั้งที่มีการแก้ st.session_state.roster_df ให้สั่ง save_data() ทันที
def trigger_save():
    save_data()

# [เพิ่มส่วนจัดการ UI และ Validation เหมือนเดิม แต่ใส่ save_data() ไว้หลังตาราง]
# (เนื่องจากพื้นที่จำกัด ผมตัดส่วนซ้ำออก แต่คุณเอาไปวางแทนที่ได้เลย)
# ...
