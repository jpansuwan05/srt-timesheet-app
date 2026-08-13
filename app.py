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
    st.error("❌ ไม่พบไฟล์ 'ข้อมูล.xlsx' กรุณาอัปโหลดไว้ในโฟลเดอร์เดียวกับโค้ดครับ")
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
# 2. ข้อมูลส่วนกลาง (ใช้ร่วมกันทั้ง 109 และ 177)
# ==========================================
st.subheader("⚙️ 1. ตั้งค่าข้อมูลส่วนกลางประจำเดือน")
col_g1, col_g2 = st.columns(2)
with col_g1:
    val_13 = st.text_input("เดือนตัวเต็ม [13] (เช่น มิถุนายน)", "มิถุนายน")
    val_7 = st.text_input("คำสั่งแขวง [7] (เช่น 5110/2520/2569)", "5110/2520/2569")
with col_g2:
    val_8 = st.text_input("วันที่ลงคำสั่ง [8] (เช่น 29 พ.ค. 69)", "29 พ.ค. 69")
    val_14 = st.text_input("วันที่เซ็นเอกสารตัวย่อ [14] (เช่น 01 ก.ค. 69)", "01 ก.ค. 69")

global_data = {"val_13": val_13, "val_7": val_7, "val_8": val_8, "val_14": val_14}

# ==========================================
# 3. ตารางเวร 109 (กรอกข้อมูลทุกคนในจุดเดียว)
# ==========================================
st.markdown("---")
st.subheader("🗓️ 2. จัดการตารางเวร 1-31 วัน (ตารางหลัก)")
st.markdown("กรุณากรอกรหัสเวร (ว, ค, ย, 0-12 ฯลฯ) ของพนักงานทุกคนลงในตารางนี้ ระบบจะนำข้อมูลไปสร้างทั้งไฟล์ 109 และ 177 ให้อัตโนมัติ")

if 'roster_df' not in st.session_state:
    data = []
    for i, (name, info) in enumerate(employees.items()):
        row = {"ลำดับ": i+1, "ชื่อ-สกุล": name, "ตำแหน่ง": info['ตำแหน่ง']}
        for d in range(1, 32): row[str(d)] = ""
        data.append(row)
    df = pd.DataFrame(data)
    for d in range(1, 32): df[str(d)] = df[str(d)].astype(str)
    st.session_state.roster_df = df

# ตั้งค่าความกว้างและหน้าตาคอลัมน์ให้ออกมาเหมือนฟอร์ม 109
column_config = {
    "ลำดับ": st.column_config.NumberColumn("ลำดับ", width="small", disabled=True),
    "ชื่อ-สกุล": st.column_config.TextColumn("ชื่อ-สกุล", width="medium", disabled=True),
    "ตำแหน่ง": st.column_config.TextColumn("ตำแหน่ง", width="medium", disabled=True)
}
for d in range(1, 32):
    column_config[str(d)] = st.column_config.TextColumn(str(d), width="small")

edited_df = st.data_editor(st.session_state.roster_df, hide_index=True, use_container_width=True, column_config=column_config, key="roster_table")
st.session_state.roster_df = edited_df

# ==========================================
# 4. ฟังก์ชันสร้างไฟล์ 109
# ==========================================
def generate_109(month_full, roster_df):
    try:
        wb = openpyxl.load_workbook("109เปล่า.xlsx")
    except FileNotFoundError:
        st.error("❌ ไม่พบเทมเพลต '109เปล่า.xlsx' ในระบบ")
        return None
    ws = wb.active
    
    # หยอดเดือนตัวเต็ม
    for r in range(1, 15):
        for c in range(1, 15):
            val = ws.cell(row=r, column=c).value
            if val and isinstance(val, str) and "[13]" in val:
                ws.cell(row=r, column=c, value=val.replace("[13]", month_full))
                
    # หยอดรหัสเวรตามรายชื่อ
    for r in range(1, 100):
        name_val = ws.cell(row=r, column=2).value
        if name_val and isinstance(name_val, str):
            match = roster_df[roster_df['ชื่อ-สกุล'].str.strip() == name_val.strip()]
            if not match.empty:
                emp_row = match.iloc[0]
                for d in range(1, 32):
                    shift = emp_row[str(d)]
                    ws.cell(row=r, column=2+d, value=shift if pd.notna(shift) else "")
                    
    # ตั้งค่าหน้ากระดาษเป็นแนวนอน
    wsp = ws.sheet_properties
    if not wsp.pageSetUpPr:
        wsp.pageSetUpPr = PageSetupProperties()
    wsp.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = False # ปล่อยความสูงอิสระ
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ==========================================
# 5. ฟังก์ชันสร้างไฟล์ 177
# ==========================================
def generate_177(emp_name, roster_data, global_vars, ind_vars):
    emp_info = employees[emp_name]
    rate = emp_info["เรท"]
    try:
        wb = openpyxl.load_workbook("ใบเบิก177 Update.xlsx")
    except FileNotFoundError:
        st.error("❌ ไม่พบเทมเพลต 'ใบเบิก177 Update.xlsx' ในระบบ")
        return None
    ws = wb.active

    replacements = {
        "[NAME]": emp_name, "[16]": emp_info["รหัสบัญชี"], "[15]": emp_info["ประเภทบัญชี"],
        "[14]": global_vars["val_14"], "[13]": global_vars["val_13"], "[12]": ind_vars["val_12"],
        "[11]": ind_vars["val_11"], "[10]": ind_vars["val_10"], "[9]": ind_vars["val_9"],
        "[8]": global_vars["val_8"], "[7]": global_vars["val_7"], "[6]": ind_vars["val_6"],
        "[5]": ind_vars["val_5"], "[4]": ind_vars["val_4"],
        "[3]": f"{float(emp_info['เงินเดือน']):,.0f}" if emp_info['เงินเดือน'] != "-" else "-",
        "[2]": emp_info["เลขประจำตัว"], "[1]": emp_info["ตำแหน่ง"]
    }

    for r in range(1, 55):
        for c in range(1, 15):
            cell = ws.cell(row=r, column=c)
            if cell.value and isinstance(cell.value, str) and "[" in cell.value and "]" in cell.value:
                new_val = cell.value
                for key, val in replacements.items(): new_val = new_val.replace(key, str(val))
                cell.value = new_val

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
            ws.cell(row=row, column=3, value="-"); ws.cell(row=row, column=4, value="-")
            ws.cell(row=row, column=5, value="-"); ws.cell(row=row, column=6, value="-")
            ws.cell(row=row, column=7, value="-")
            
    # บังคับ Fit to Page (แนวตั้ง)
    wsp = ws.sheet_properties
    if not wsp.pageSetUpPr:
        wsp.pageSetUpPr = PageSetupProperties()
    wsp.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToHeight = 1
    ws.page_setup.fitToWidth = 1
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ==========================================
# 6. เมนูส่งออก (Export) แยกอิสระ 2 แท็บ
# ==========================================
st.markdown("---")
st.subheader("🖨️ 3. ส่งออกเอกสาร (Export)")
tab1, tab2 = st.tabs(["📄 ส่งออก 109 (ตารางเวรรวม)", "🧾 ส่งออก 177 (ใบเบิกรายบุคคล)"])

# แท็บส่งออก 109
with tab1:
    st.markdown("**ตารางเวร 109 ของทั้งสถานี**")
    if st.button("คลิกเพื่อสร้างไฟล์ 109", type="primary"):
        excel_109 = generate_109(global_data["val_13"], st.session_state.roster_df)
        if excel_109:
            st.success(f"สร้างไฟล์ 109 ประจำเดือน {global_data['val_13']} สำเร็จ!")
            st.download_button("📥 ดาวน์โหลดไฟล์ 109", data=excel_109, file_name=f"109_{global_data['val_13']}.xlsx")

# แท็บส่งออก 177
with tab2:
    st.markdown("**ข้อมูลเฉพาะบุคคลสำหรับใบเบิก 177**")
    selected_emp_177 = st.selectbox("เลือกพนักงานที่ต้องการเบิก 177", list(employees.keys()))
    
    if f'ind_data_{selected_emp_177}' not in st.session_state:
        st.session_state[f'ind_data_{selected_emp_177}'] = {
            "val_4": "-", "val_5": "-", "val_6": "-",
            "val_9": "-", "val_10": "-", "val_11": "-", "val_12": "-"
        }

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.session_state[f'ind_data_{selected_emp_177}']['val_4'] = st.text_input("วันหยุด [4]", st.session_state[f'ind_data_{selected_emp_177}']['val_4'])
        st.session_state[f'ind_data_{selected_emp_177}']['val_9'] = st.text_input("หยุดพักผ่อน [9]", st.session_state[f'ind_data_{selected_emp_177}']['val_9'])
    with c2:
        st.session_state[f'ind_data_{selected_emp_177}']['val_5'] = st.text_input("วันหยุดที่เบิกได้ [5]", st.session_state[f'ind_data_{selected_emp_177}']['val_5'])
        st.session_state[f'ind_data_{selected_emp_177}']['val_10'] = st.text_input("พักผ่อนตั้งแต่ [10]", st.session_state[f'ind_data_{selected_emp_177}']['val_10'])
    with c3:
        st.session_state[f'ind_data_{selected_emp_177}']['val_6'] = st.text_input("เดือนตัวย่อ [6]", st.session_state[f'ind_data_{selected_emp_177}']['val_6'])
        st.session_state[f'ind_data_{selected_emp_177}']['val_11'] = st.text_input("พักผ่อนถึง [11]", st.session_state[f'ind_data_{selected_emp_177}']['val_11'])
    with c4:
        st.session_state[f'ind_data_{selected_emp_177}']['val_12'] = st.text_input("รวมพักผ่อน [12]", st.session_state[f'ind_data_{selected_emp_177}']['val_12'])

    if st.button(f"ออกใบเบิก 177 ของ {selected_emp_177}", type="primary"):
        # ดึงเวรของคนนี้ 1-31 จากตารางใหญ่ด้านบน
        emp_row = st.session_state.roster_df[st.session_state.roster_df['ชื่อ-สกุล'] == selected_emp_177].iloc[0]
        roster_dict = {str(d): str(emp_row[str(d)]) if pd.notna(emp_row[str(d)]) else "" for d in range(1, 32)}
        
        ind_data = st.session_state[f'ind_data_{selected_emp_177}']
        excel_177 = generate_177(selected_emp_177, roster_dict, global_data, ind_data)
        
        if excel_177:
            st.success(f"สร้างใบเบิก 177 ของ {selected_emp_177} เสร็จสิ้น!")
            st.download_button("📥 ดาวน์โหลดไฟล์ 177", data=excel_177, file_name=f"ใบเบิก_{selected_emp_177}.xlsx")
