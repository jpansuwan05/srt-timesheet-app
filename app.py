import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
import io
import datetime

st.set_page_config(page_title="SRT Timesheet App", layout="wide")
st.title("🚂 ระบบจัดการเวรและใบเบิกค่าตอบแทน (รฟท.)")
st.markdown("---")

# ==========================================
# 1. ฐานข้อมูลและตั้งค่ารหัส
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
# 2. UI: จัดการตารางเวร (109)
# ==========================================
st.subheader("📝 1. บันทึกตารางเวร (109)")
col1, col2 = st.columns(2)
with col1: selected_emp = st.selectbox("เลือกพนักงาน", list(employees.keys()))
with col2: selected_month = st.selectbox("ประจำเดือน", ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"], index=5)

if 'roster' not in st.session_state:
    st.session_state.roster = {emp: {str(day): "" for day in range(1, 32)} for emp in employees.keys()}
    # จำลองข้อมูลตั้งต้นให้ตรงกับภาพ 108 ชม.
    st.session_state.roster["นายเจษฎากร ปานสุวรรณ"] = {
        "1": "ว/ค", "2": "ค", "3": "ค", "4": "ค", "5": "ค/ว", "6": "ว", "7": "ว", "8": "ว", "9": "ว", "10": "ว",
        "11": "ว/ค", "12": "ว", "13": "ว", "14": "ว", "15": "ว", "16": "ว/ค", "17": "ค", "18": "ค", "19": "ค", "20": "ค/ว",
        "21": "ว", "22": "ย", "23": "ย", "24": "ย", "25": "ค", "26": "ค/ว", "27": "ว", "28": "ว", "29": "ว", "30": "ว/ค", "31": "-"
    }

df_roster = pd.DataFrame([st.session_state.roster[selected_emp]])
edited_df = st.data_editor(df_roster, hide_index=True)
st.session_state.roster[selected_emp] = edited_df.iloc[0].to_dict()


# ==========================================
# 3. ฟังก์ชันสร้างไฟล์ Excel ใบเบิก 100%
# ==========================================
def generate_excel(emp_name, month_name, roster_data):
    emp_info = employees[emp_name]
    rate = emp_info["เรท"]
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "แบบฟอร์มใบเบิก"

    # ตั้งความกว้างคอลัมน์เป๊ะตามแบบ
    widths = {'A': 5.5, 'B': 36, 'C': 7, 'D': 5, 'E': 4.5, 'F': 8, 'G': 4.5, 'H': 30}
    for col, w in widths.items(): ws.column_dimensions[col].width = w

    font_normal = Font(name="TH SarabunPSK", size=15)
    font_bold = Font(name="TH SarabunPSK", size=15, bold=True)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    
    border_solid = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    border_dotted = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='dotted'), bottom=Side(style='dotted'))

    # ส่วนหัว
    ws.merge_cells('A1:H1'); ws['A1'] = "การรถไฟแห่งประเทศไทย"
    ws['A1'].font = Font(name="TH SarabunPSK", size=18, bold=True); ws['A1'].alignment = align_center

    ws.merge_cells('A2:H2'); ws['A2'] = "รายการเบิกเงินค่าตอบแทนพิเศษ การทำงานเกินกำหนดเวลาทำงานปกติ"
    ws['A2'].font = font_bold; ws['A2'].alignment = align_center

    year = datetime.datetime.now().year + 543
    ws.merge_cells('A3:H3'); ws['A3'] = f"ประจำเดือน {month_name} พ.ศ. {year}"
    ws['A3'].font = font_bold; ws['A3'].alignment = align_center

    ws.merge_cells('A4:H4')
    ws['A4'] = f"ชื่อ  {emp_name}      ตำแหน่ง  {emp_info['ตำแหน่ง']}      เลขประจำตัว  {emp_info['เลขประจำตัว']}      อัตราเงินเดือน  {emp_info['เงินเดือน']} บาท      ฝ่าย  ฝ่ายปฏิบัติการเดินรถ"
    ws['A4'].font = font_normal; ws['A4'].alignment = align_left

    # หัวตาราง (แยก 2 บรรทัด)
    headers = [("A5:A6", "วันที่"), ("C5", "จำนวน"), ("D5:E5", "ชั่วโมงละ"), ("F5:G5", "จำนวนเงิน"), ("H5:H6", "หมายเหตุ")]
    for merge_range, text in headers:
        if ':' in merge_range: ws.merge_cells(merge_range)
        ws[merge_range.split(':')[0]] = text
        ws[merge_range.split(':')[0]].font = font_bold; ws[merge_range.split(':')[0]].alignment = align_center
        
    ws['B5'] = "รายการเบิก"; ws['B5'].font = font_bold; ws['B5'].alignment = align_center
    ws['B6'] = "(ทำจากเวลาใดถึงเวลาใด)"; ws['B6'].font = font_bold; ws['B6'].alignment = align_center
    ws['C6'] = "ชั่วโมง"; ws['C6'].font = font_bold; ws['C6'].alignment = align_center
    ws['D6'] = "บาท"; ws['E6'] = "สต."; ws['F6'] = "บาท"; ws['G6'] = "สต."
    
    for col in ['D', 'E', 'F', 'G']:
        ws[f'{col}6'].font = font_bold; ws[f'{col}6'].alignment = align_center

    # ตีเส้นกรอบหัวตารางทึบ
    for r in range(5, 7):
        for c in range(1, 9): ws.cell(row=r, column=c).border = border_solid

    # ข้อมูล 1-31 วัน (Python เป็นคนคำนวณเงินแล้วกรอกลง Excel)
    start_row = 7
    total_hours = 0
    total_money = 0
    
    for day in range(1, 32):
        row = start_row + day - 1
        ws.cell(row=row, column=1, value=day).alignment = align_center
        shift = roster_data.get(str(day), "").strip()
        
        text_val = "-"; hours_val = "-"; money_val = "-"
        
        if shift in shift_data and shift_data[shift]["hours"] != "-":
            text_val = shift_data[shift]["text"]
            hours_val = shift_data[shift]["hours"]
            money_val = int(hours_val) * rate
            total_hours += int(hours_val)
            total_money += money_val
        elif shift in ["ย", "พ"]:
            text_val = "ย." if shift == "ย" else "พ."
            
        ws.cell(row=row, column=2, value=text_val).alignment = align_center
        ws.cell(row=row, column=3, value=hours_val).alignment = align_center
        ws.cell(row=row, column=4, value=rate if hours_val != "-" else "-").alignment = align_center
        ws.cell(row=row, column=5, value="00" if hours_val != "-" else "-").alignment = align_center
        ws.cell(row=row, column=6, value=money_val).alignment = align_center
        ws.cell(row=row, column=7, value="00" if hours_val != "-" else "-").alignment = align_center
        
        for c in range(1, 9):
            ws.cell(row=row, column=c).font = font_normal
            ws.cell(row=row, column=c).border = border_dotted

    # แก้ไขขอบซ้ายขวาของข้อมูลรายวันให้เป็นเส้นทึบ
    for r in range(start_row, start_row + 31):
        ws.cell(row=r, column=1).border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='dotted'), bottom=Side(style='dotted'))
        ws.cell(row=r, column=8).border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='dotted'), bottom=Side(style='dotted'))

    # หมายเหตุด้านขวา
    ws['H7'] = " ในรอบเดือน"
    ws['H15'] = " วันหยุดประจำสัปดาห์ วันที่"
    ws['H26'] = " หมายเหตุ"
    ws['H27'] = " เบิกตามระเบียบการรถไฟฯ"
    for r in range(7, 38): ws[f'H{r}'].alignment = align_left
    for r in [7, 15, 26, 27]: ws[f'H{r}'].font = font_bold

    # แถวรวมยอด
    sum_row = 38
    ws.merge_cells(f'A{sum_row}:B{sum_row}')
    ws[f'A{sum_row}'] = "รวม"; ws[f'A{sum_row}'].alignment = align_center; ws[f'A{sum_row}'].font = font_bold
    ws[f'C{sum_row}'] = total_hours if total_hours > 0 else "-"
    ws[f'D{sum_row}'] = rate; ws[f'E{sum_row}'] = "00"
    ws[f'F{sum_row}'] = total_money if total_money > 0 else "-"; ws[f'G{sum_row}'] = "00"
    for col in ['A', 'C', 'D', 'E', 'F', 'G', 'H']:
        cell = ws.cell(row=sum_row, column=ws[col+'1'].column)
        cell.border = border_solid; cell.alignment = align_center; cell.font = font_bold

    # ลายเซ็นต์จัดตำแหน่งตามรูป
    f_row = 40
    ws.merge_cells(f'A{f_row}:C{f_row}'); ws[f'A{f_row}'] = "ข้าพเจ้าขอรับรองว่า รายการเบิกค่าตอบแทนการทำงาน"
    ws.merge_cells(f'D{f_row}:F{f_row}'); ws[f'D{f_row}'] = "ขอรับรองว่ารายการเบิกเงินในหลักฐาน"
    ws.merge_cells(f'G{f_row}:H{f_row}'); ws[f'G{f_row}'] = "ประเภทบัญชี"

    f_row = 44
    ws.merge_cells(f'A{f_row}:C{f_row}'); ws[f'A{f_row}'] = "......................................................................."
    ws.merge_cells(f'D{f_row}:F{f_row}'); ws[f'D{f_row}'] = f"({emp_name})"
    ws.merge_cells(f'G{f_row}:H{f_row}'); ws[f'G{f_row}'] = "......................................................................."

    for r in [40, 44]:
        ws[f'A{r}'].alignment = Alignment(horizontal="left"); ws[f'A{r}'].font = font_bold
        ws[f'D{r}'].alignment = Alignment(horizontal="center"); ws[f'D{r}'].font = font_bold
        ws[f'G{r}'].alignment = Alignment(horizontal="center"); ws[f'G{r}'].font = font_bold

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ==========================================
# 4. ปุ่มสร้างและดาวน์โหลดเอกสาร
# ==========================================
st.markdown("---")
st.subheader("🖨️ 2. สร้างใบเบิกค่าตอบแทน")

if st.button(f"ออกเอกสารใบเบิก ของ {selected_emp}"):
    roster = st.session_state.roster[selected_emp]
    excel_file = generate_excel(selected_emp, selected_month, roster)
    
    st.success(f"สร้างไฟล์ใบเบิกของ {selected_emp} ประจำเดือน {selected_month} สำเร็จ!")
    
    st.download_button(
        label="📥 ดาวน์โหลดไฟล์ Excel (แบบฟอร์ม 100%)",
        data=excel_file,
        file_name=f"ใบเบิก_{selected_emp}_{selected_month}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# 1. ตั้งค่าให้เป็นกระดาษ A4 แนวตั้ง
ws.page_setup.paperSize = ws.PAPERSIZE_A4
ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT

# 2. บังคับให้บีบตารางให้พอดี 1 หน้ากระดาษ (Fit to 1 page wide by 1 page tall)
ws.page_setup.fitToPage = True
ws.page_setup.fitToHeight = 1
ws.page_setup.fitToWidth = 1

# 3. ตั้งค่าขอบกระดาษ (Margin) ให้แคบลง เพื่อให้แบบฟอร์มดูใหญ่และไม่ถูกบีบจนเกินไป
from openpyxl.worksheet.page import PageMargins
ws.page_margins = PageMargins(left=0.25, right=0.25, top=0.5, bottom=0.5, header=0.3, footer=0.3)
