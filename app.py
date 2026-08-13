import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.worksheet.properties import PageSetupProperties
import io
import datetime

st.set_page_config(page_title="SRT Timesheet App", layout="wide")
st.title("🚂 ระบบจัดการเวรและใบเบิกค่าตอบแทน (รฟท.)")
st.markdown("---")

# ==========================================
# 1. โหลดข้อมูลพนักงานจากไฟล์ "ข้อมูล.xlsx"
# ==========================================
@st.cache_data
def load_employee_data():
    try:
        df = pd.read_excel("ข้อมูล.xlsx")
        emp_dict = {}
        for _, row in df.iterrows():
            # ใช้ชื่อพนักงานเป็น Key ใน Dictionary
            emp_dict[str(row["รายชื่อ"])] = {
                "ตำแหน่ง": str(row["ตำแหน่ง"]) if pd.notna(row["ตำแหน่ง"]) else "-",
                "เลขประจำตัว": str(row["เลขประจำตัว"]) if pd.notna(row["เลขประจำตัว"]) else "-",
                "เงินเดือน": str(row["เงินเดือน"]) if pd.notna(row["เงินเดือน"]) else "-",
                "เรท": float(row["เรท 1 ชั่วโมง"]) if pd.notna(row["เรท 1 ชั่วโมง"]) else 0,
                "ประเภทบัญชี": str(row["ประเภทบัญชี"]) if pd.notna(row["ประเภทบัญชี"]) else "-",
                "รหัสบัญชี": str(row["รหัสบัญชี"]) if pd.notna(row["รหัสบัญชี"]) else "-"
            }
        return emp_dict
    except Exception as e:
        return None

employees = load_employee_data()
if not employees:
    st.error("❌ ไม่พบไฟล์ 'ข้อมูล.xlsx' หรือไฟล์มีปัญหา กรุณาอัปโหลดไว้ในโฟลเดอร์เดียวกับโค้ดครับ")
    st.stop()

# ข้อมูลรูปแบบเวร
shift_data = {
    "ว": {"text": "(06.00-18.00) น.", "hours": 4},
    "ค": {"text": "(00.00-06.00)(18.00-24.00) น.", "hours": 4},
    "ว/ค": {"text": "(06.00-12.00)(18.00-24.00) น.", "hours": 4},
    "ค/ว": {"text": "(00.00-06.00)(12.00-18.00) น.", "hours": 4},
    "0-12": {"text": "(00.00-12.00) น.", "hours": 4},
    "12-24": {"text": "(12.00-24.00) น.", "hours": 4},
    "(ว)": {"text": "(06.00-18.00) น.", "hours": 4},
    "(ค)": {"text": "(00.00-06.00)(18.00-24.00) น.", "hours": 4},
    "(ว/ค)": {"text": "(06.00-12.00)(18.00-24.00) น.", "hours": 4},
    "(ค/ว)": {"text": "(00.00-06.00)(12.00-18.00) น.", "hours": 4},
    "(0-12)": {"text": "(00.00-12.00) น.", "hours": 4},
    "(12-24)": {"text": "(12.00-24.00) น.", "hours": 4},
    "ย": {"text": "ย.", "hours": "-"},
    "พ": {"text": "พ.", "hours": "-"},
}

# ==========================================
# 2. ข้อมูลส่วนกลาง (ใช้ร่วมกันทั้งเดือน) [7],[8],[13],[14]
# ==========================================
st.subheader("⚙️ 1. ข้อมูลส่วนกลาง (กรอกครั้งเดียว ใช้กับทุกคน)")
col_g1, col_g2 = st.columns(2)
with col_g1:
    val_13 = st.text_input("เดือนตัวเต็ม [13] (เช่น มิถุนายน)", "มิถุนายน")
    val_7 = st.text_input("คำสั่งแขวง [7] (เช่น 5110/2520/2569)", "5110/2520/2569")
with col_g2:
    val_8 = st.text_input("วันที่ลงคำสั่ง [8] (เช่น 29 พ.ค. 69)", "29 พ.ค. 69")
    val_14 = st.text_input("วันที่เซ็นเอกสารตัวย่อ [14] (เช่น 01 ก.ค. 69)", "01 ก.ค. 69")

global_data = {"val_13": val_13, "val_7": val_7, "val_8": val_8, "val_14": val_14}

# ==========================================
# 3. ข้อมูลรายบุคคลและตารางเวร [4]-[6], [9]-[12]
# ==========================================
st.markdown("---")
st.subheader("📝 2. จัดการเวรและข้อมูลรายบุคคล")
selected_emp = st.selectbox("เลือกพนักงาน", list(employees.keys()))

# เก็บสถานะช่องกรอกข้อมูลส่วนตัวของแต่ละคนไม่ให้หายไปเวลากดเปลี่ยนชื่อ
if f'ind_data_{selected_emp}' not in st.session_state:
    st.session_state[f'ind_data_{selected_emp}'] = {
        "val_4": "-", "val_5": "-", "val_6": "-",
        "val_9": "-", "val_10": "-", "val_11": "-", "val_12": "-"
    }

st.markdown(f"**ข้อมูลเพิ่มเติมของ: {selected_emp}**")
c1, c2, c3, c4 = st.columns(4)
with c1:
    st.session_state[f'ind_data_{selected_emp}']['val_4'] = st.text_input("วันหยุด [4] (เช่น 22-24)", st.session_state[f'ind_data_{selected_emp}']['val_4'])
    st.session_state[f'ind_data_{selected_emp}']['val_9'] = st.text_input("หยุดพักผ่อน [9] (เช่น 22-24)", st.session_state[f'ind_data_{selected_emp}']['val_9'])
with c2:
    st.session_state[f'ind_data_{selected_emp}']['val_5'] = st.text_input("วันหยุดที่เบิกได้ [5] (เช่น 11-15)", st.session_state[f'ind_data_{selected_emp}']['val_5'])
    st.session_state[f'ind_data_{selected_emp}']['val_10'] = st.text_input("พักผ่อนตั้งแต่ [10] (เช่น 11)", st.session_state[f'ind_data_{selected_emp}']['val_10'])
with c3:
    st.session_state[f'ind_data_{selected_emp}']['val_6'] = st.text_input("เดือนตัวย่อ [6] (เช่น มิ.ย. 69)", st.session_state[f'ind_data_{selected_emp}']['val_6'])
    st.session_state[f'ind_data_{selected_emp}']['val_11'] = st.text_input("พักผ่อนถึง [11] (เช่น 15)", st.session_state[f'ind_data_{selected_emp}']['val_11'])
with c4:
    st.session_state[f'ind_data_{selected_emp}']['val_12'] = st.text_input("รวมพักผ่อน [12] (เช่น 08)", st.session_state[f'ind_data_{selected_emp}']['val_12'])

# ตารางกรอกเวร 1-31 วัน
st.markdown("**ตารางเวร 1-31 วัน** (กรอก ว, ค, ย, 0-12 ฯลฯ)")
if 'roster' not in st.session_state:
    st.session_state.roster = {emp: {str(day): "" for day in range(1, 32)} for emp in employees.keys()}

df_roster = pd.DataFrame([st.session_state.roster[selected_emp]])
edited_df = st.data_editor(df_roster, hide_index=True)
st.session_state.roster[selected_emp] = edited_df.iloc[0].to_dict()

# ==========================================
# 4. ฟังก์ชันเปิดไฟล์เทมเพลตและหยอดข้อมูล
# ==========================================
def generate_excel(emp_name, roster_data, global_vars, ind_vars):
    emp_info = employees[emp_name]
    rate = emp_info["เรท"]
    
    try:
        wb = openpyxl.load_workbook("ใบเบิก177 Update.xlsx")
    except FileNotFoundError:
        st.error("❌ ไม่พบเทมเพลต 'ใบเบิก177 Update.xlsx' กรุณาอัปโหลดเข้าสู่ GitHub ด้วยครับ")
        return None
        
    ws = wb.active

    # ดิกชันนารีสำหรับการแทนที่คำ (ไล่จากตัวเลขมากไปน้อยป้องกันการแทนที่ทับซ้อน)
    replacements = {
        "[NAME]": emp_name,
        "[16]": emp_info["รหัสบัญชี"],
        "[15]": emp_info["ประเภทบัญชี"],
        "[14]": global_vars["val_14"],
        "[13]": global_vars["val_13"],
        "[12]": ind_vars["val_12"],
        "[11]": ind_vars["val_11"],
        "[10]": ind_vars["val_10"],
        "[9]": ind_vars["val_9"],
        "[8]": global_vars["val_8"],
        "[7]": global_vars["val_7"],
        "[6]": ind_vars["val_6"],
        "[5]": ind_vars["val_5"],
        "[4]": ind_vars["val_4"],
        "[3]": f"{float(emp_info['เงินเดือน']):,.0f}" if emp_info['เงินเดือน'] != "-" else "-", # ใส่ลูกน้ำให้เงินเดือนสวยงาม
        "[2]": emp_info["เลขประจำตัว"],
        "[1]": emp_info["ตำแหน่ง"]
    }

    # สแกนทุกช่องในหน้ากระดาษเพื่อหาและแทนที่ตัวแปร [...]
    for r in range(1, 55):
        for c in range(1, 15):
            cell = ws.cell(row=r, column=c)
            # เช็คว่าเซลล์มีข้อมูลเป็นตัวหนังสือ และมีเครื่องหมาย [ ] ให้ทำการแทนที่
            if cell.value and isinstance(cell.value, str) and "[" in cell.value and "]" in cell.value:
                new_val = cell.value
                for key, val in replacements.items():
                    new_val = new_val.replace(key, str(val))
                cell.value = new_val

    # หยอดข้อมูลเวร 31 วัน (เริ่มตั้งแต่แถวที่ 7)
    start_row = 7
    for day in range(1, 32):
        row = start_row + day - 1
        shift = roster_data.get(str(day), "").strip()
        
        if shift in shift_data and shift_data[shift]["hours"] != "-":
            ws.cell(row=row, column=2, value=shift_data[shift]["text"])
            ws.cell(row=row, column=3, value=int(shift_data[shift]["hours"]))
            ws.cell(row=row, column=4, value=rate)
            ws.cell(row=row, column=5, value="00")
            ws.cell(row=row, column=6, value=int(shift_data[shift]["hours"]) * rate)
            ws.cell(row=row, column=7, value="00")
        else:
            if shift == "ย": ws.cell(row=row, column=2, value="ย.")
            elif shift == "พ": ws.cell(row=row, column=2, value="พ.")
            else: ws.cell(row=row, column=2, value="-")
                
            ws.cell(row=row, column=3, value="-")
            ws.cell(row=row, column=4, value="-")
            ws.cell(row=row, column=5, value="-")
            ws.cell(row=row, column=6, value="-")
            ws.cell(row=row, column=7, value="-")

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ==========================================
# 5. ปุ่มดาวน์โหลดเอกสาร
# ==========================================
st.markdown("---")
if st.button(f"ออกเอกสารใบเบิก ของ {selected_emp}", type="primary"):
    roster = st.session_state.roster[selected_emp]
    ind_data = st.session_state[f'ind_data_{selected_emp}']
    
    excel_file = generate_excel(selected_emp, roster, global_data, ind_data)
    
    if excel_file:
        st.success(f"เตรียมใบเบิกของ {selected_emp} เสร็จสิ้น! 🎉 (อ้างอิงเทมเพลต 'ใบเบิก177 Update')")
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel พร้อมพริ้นต์",
            data=excel_file,
            file_name=f"ใบเบิก_{selected_emp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
