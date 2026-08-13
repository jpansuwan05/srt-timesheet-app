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
# 1. ฐานข้อมูลรหัสและพนักงาน
# ==========================================
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

employees = {
    "นายเจษฎากร ปานสุวรรณ": {"ตำแหน่ง": "พนักงานเดินรถ 6", "เลขประจำตัว": "2052576", "เงินเดือน": "20,640", "เรท": 86},
    "นายทองปิ่น จันทร์แปลง": {"ตำแหน่ง": "นายสถานีชุมทางตลิ่งชัน", "เลขประจำตัว": "1234567", "เงินเดือน": "56,150", "เรท": 234}
}

# ==========================================
# 2. UI: จัดการตารางเวร
# ==========================================
st.subheader("📝 1. บันทึกตารางเวร (109)")
col1, col2 = st.columns(2)
with col1: selected_emp = st.selectbox("เลือกพนักงาน", list(employees.keys()))
with col2: selected_month = st.selectbox("ประจำเดือน", ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"], index=5)

if 'roster' not in st.session_state:
    st.session_state.roster = {emp: {str(day): "" for day in range(1, 32)} for emp in employees.keys()}
    st.session_state.roster["นายเจษฎากร ปานสุวรรณ"] = {
        "1": "ว/ค", "2": "ค", "3": "ค", "4": "ค", "5": "ค/ว", "6": "ว", "7": "ว", "8": "ว", "9": "ว", "10": "ว",
        "11": "ว/ค", "12": "ว", "13": "ว", "14": "ว", "15": "ว", "16": "ว/ค", "17": "ค", "18": "ค", "19": "ค", "20": "ค/ว",
        "21": "ว", "22": "ย", "23": "ย", "24": "ย", "25": "ค", "26": "ค/ว", "27": "ว", "28": "ว", "29": "ว", "30": "ว/ค", "31": "-"
    }

df_roster = pd.DataFrame([st.session_state.roster[selected_emp]])
edited_df = st.data_editor(df_roster, hide_index=True)
st.session_state.roster[selected_emp] = edited_df.iloc[0].to_dict()

# ==========================================
# 3. ฟังก์ชัน หยอดข้อมูลและตั้งค่าหน้ากระดาษ
# ==========================================
def generate_excel(emp_name, month_name, roster_data):
    emp_info = employees[emp_name]
    rate = emp_info["เรท"]
    
    try:
        wb = openpyxl.load_workbook("ใบเบิก 177.xlsx")
    except FileNotFoundError:
        st.error("❌ ไม่พบไฟล์ 'ใบเบิก 177.xlsx' กรุณาอัปโหลดไฟล์นี้เข้า GitHub ด้วยครับ")
        return None
        
    ws = wb.active

    # แทนที่ข้อความหัวกระดาษ
    year = datetime.datetime.now().year + 543
    ws.cell(row=3, column=1, value=f"ประจำเดือน {month_name} พ.ศ. {year}")
    
    info_text = f"ชื่อ  {emp_name}      ตำแหน่ง  {emp_info['ตำแหน่ง']}      เลขประจำตัว  {emp_info['เลขประจำตัว']}      อัตราเงินเดือน  {emp_info['เงินเดือน']} บาท      ฝ่าย  ฝ่ายปฏิบัติการเดินรถ"
    ws.cell(row=4, column=1, value=info_text)

    # แทนที่ชื่อจุดเซ็น
    ws.cell(row=44, column=2, value=f"({emp_name})")

    # หยอดข้อมูลเวร
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

    # ----------------------------------------------------
    # การตั้งค่าหน้ากระดาษ (บังคับ Fit to 1 Page อัตโนมัติ)
    # ----------------------------------------------------
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
# 4. ปุ่มออกใบเบิก
# ==========================================
st.markdown("---")
st.subheader("🖨️ 2. สร้างใบเบิกค่าตอบแทน")

if st.button(f"ออกเอกสารใบเบิก ของ {selected_emp}"):
    roster = st.session_state.roster[selected_emp]
    excel_file = generate_excel(selected_emp, selected_month, roster)
    
    if excel_file:
        st.success(f"เตรียมฟอร์มของ {selected_emp} สำเร็จ! (ตั้งค่าพอดี A4 อัตโนมัติแล้ว)")
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel พร้อมพริ้นต์",
            data=excel_file,
            file_name=f"ใบเบิก_{selected_emp}_{selected_month}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
