import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
import io
import datetime

st.set_page_config(page_title="SRT Timesheet App", layout="wide")
st.title("🚂 ระบบจัดการเวรและใบเบิกค่าตอบแทน (รฟท.)")
st.markdown("---")

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

st.subheader("📝 1. บันทึกตารางเวร (109)")
col1, col2 = st.columns(2)
with col1: selected_emp = st.selectbox("เลือกพนักงาน", list(employees.keys()))
with col2: selected_month = st.selectbox("ประจำเดือน", ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"], index=5)

if 'roster' not in st.session_state:
    st.session_state.roster = {emp: {str(day): "" for day in range(1, 32)} for emp in employees.keys()}
    st.session_state.roster["นายเจษฎากร ปานสุวรรณ"] = {
        "1": "ว/ค", "2": "ค", "3": "ค", "4": "ค", "5": "ค/ว", "6": "ว", "7": "ว", "8": "ว", "9": "ว", "10": "ว",
        "11": "ว/ค", "12": "ว", "13": "ว", "14": "ว", "15": "ว", "16": "ว/ค", "17": "ค", "18": "ค", "19": "ค", "20": "ค/ว",
        "21": "ว", "22": "ย", "23": "ย", "24": "ย", "25": "ค", "26": "ค/ว", "27": "ว", "28": "ว", "29": "ว", "30": "ว/ค", "31": ""
    }

df_roster = pd.DataFrame([st.session_state.roster[selected_emp]])
edited_df = st.data_editor(df_roster, hide_index=True)
st.session_state.roster[selected_emp] = edited_df.iloc[0].to_dict()

def generate_excel(emp_name, month_name, roster_data):
    emp_info = employees[emp_name]
    rate = emp_info["เรท"]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "แบบฟอร์มใบเบิก"
    widths = {'A': 5, 'B': 40, 'C': 12, 'D': 8, 'E': 5, 'F': 10, 'G': 5, 'H': 35}
    for col, w in widths.items(): ws.column_dimensions[col].width = w
    font_normal = Font(name="TH SarabunPSK", size=16)
    font_bold = Font(name="TH SarabunPSK", size=16, bold=True)
    font_head = Font(name="TH SarabunPSK", size=20, bold=True)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border_thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    border_dotted = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='dotted'), bottom=Side(style='dotted'))

    ws.merge_cells('A1:H1'); ws['A1'] = "การรถไฟแห่งประเทศไทย"
    ws['A1'].font = font_head; ws['A1'].alignment = align_center
    ws.merge_cells('A2:H2'); ws['A2'] = "รายการเบิกเงินค่าตอบแทนพิเศษ การทำงานเกินกำหนดเวลาทำงานปกติ"
    ws['A2'].font = font_bold; ws['A2'].alignment = align_center
    year = datetime.datetime.now().year + 543
    ws.merge_cells('A3:H3'); ws['A3'] = f"ประจำเดือน  {month_name}  พ.ศ. {year}"
    ws['A3'].font = font_bold; ws['A3'].alignment = align_center
    ws.merge_cells('A4:H4')
    ws['A4'] = f"ชื่อ  {emp_name}     ตำแหน่ง  {emp_info['ตำแหน่ง']}     เลขประจำตัว  {emp_info['เลขประจำตัว']}     อัตราเงินเดือน  {emp_info['เงินเดือน']} บาท     ฝ่าย  ฝ่ายปฏิบัติการเดินรถ"
    ws['A4'].font = font_bold; ws['A4'].alignment = align_center

    headers = [("A5:A6", "วันที่"), ("B5:B6", "รายการเบิก\n(ทำจากเวลาใดถึงเวลาใด)"), ("C5:C6", "จำนวน\nชั่วโมง"), 
               ("D5:E5", "ชั่วโมงละ"), ("F5:G5", "จำนวนเงิน"), ("H5:H6", "หมายเหตุ")]
    for merge_range, text in headers:
        ws.merge_cells(merge_range); ws[merge_range.split(':')[0]] = text
        ws[merge_range.split(':')[0]].font = font_bold; ws[merge_range.split(':')[0]].alignment = align_center
        
    ws['D6'] = "บาท"; ws['E6'] = "สต."; ws['F6'] = "บาท"; ws['G6'] = "สต."
    for r in range(5, 7):
        for c in range(1, 9): ws.cell(row=r, column=c).border = border_thin

    start_row = 7; total_hours = 0; total_money = 0
    for day in range(1, 32):
        row = start_row + day - 1
        ws.cell(row=row, column=1, value=day).alignment = align_center
        shift = roster_data.get(str(day), "").strip()
        text_val = "-"; hours_val = "-"; money_val = "-"
        if shift in shift_data:
            text_val = shift_data[shift]["text"]
            hours_val = shift_data[shift]["hours"]
            if hours_val != "-":
                money_val = int(hours_val) * rate
                total_hours += int(hours_val)
                total_money += money_val
                
        ws.cell(row=row, column=2, value=text_val).alignment = align_center
        ws.cell(row=row, column=3, value=hours_val).alignment = align_center
        ws.cell(row=row, column=4, value=rate if hours_val != "-" else "-").alignment = align_center
        ws.cell(row=row, column=5, value="00" if hours_val != "-" else "-").alignment = align_center
        ws.cell(row=row, column=6, value=money_val).alignment = align_center
        ws.cell(row=row, column=7, value="00" if hours_val != "-" else "-").alignment = align_center
        for c in range(1, 9):
            ws.cell(row=row, column=c).font = font_normal
            ws.cell(row=row, column=c).border = border_dotted

    ws['H7'] = " ในรอบเดือน"
    ws['H15'] = " วันหยุดประจำสัปดาห์ วันที่"
    ws['H26'] = " หมายเหตุ"
    ws['H27'] = " เบิกตามระเบียบการรถไฟฯ"
    for r in range(7, 38): ws[f'H{r}'].alignment = Alignment(horizontal="left", vertical="center")

    sum_row = start_row + 31
    ws.merge_cells(f'A{sum_row}:B{sum_row}')
    ws[f'A{sum_row}'] = "รวม"; ws[f'A{sum_row}'].alignment = align_center; ws[f'A{sum_row}'].font = font_bold
    ws[f'C{sum_row}'] = total_hours if total_hours > 0 else "-"
    ws[f'D{sum_row}'] = rate; ws[f'E{sum_row}'] = "00"; ws[f'F{sum_row}'] = total_money if total_money > 0 else "-"; ws[f'G{sum_row}'] = "00"
    for col in ['A', 'C', 'D', 'E', 'F', 'G', 'H']:
        cell = ws.cell(row=sum_row, column=ws[col+'1'].column)
        cell.border = border_thin; cell.alignment = align_center

    f_row = sum_row + 2
    ws.merge_cells(f'A{f_row}:B{f_row}'); ws[f'A{f_row}'] = "ข้าพเจ้าขอรับรองว่า รายการเบิกค่าตอบแทนการทำงานเกิน\nกำหนดเวลาทำงานปกติข้างต้นเป็นความจริง"
    ws.merge_cells(f'C{f_row}:F{f_row}'); ws[f'C{f_row}'] = "ขอรับรองว่ารายการเบิกเงินในหลักฐาน\nฉบับนี้ได้ตรวจสอบถูกต้องแล้ว"
    ws.merge_cells(f'G{f_row}:H{f_row}'); ws[f'G{f_row}'] = "ประเภทบัญชี\nได้ตรวจสอบถูกต้องแล้ว"
    f_row += 4
    ws.merge_cells(f'A{f_row}:B{f_row}'); ws[f'A{f_row}'] = ".......................................................................\nผู้เบิก"
    ws.merge_cells(f'C{f_row}:F{f_row}'); ws[f'C{f_row}'] = "(นายทองปิ่น จันทร์แปลง)\nตำแหน่ง นายสถานีชุมทางตลิ่งชัน"
    ws.merge_cells(f'G{f_row}:H{f_row}'); ws[f'G{f_row}'] = ".......................................................................\n(นายปฐม ชุมวงศ์)\nสตร.รบ. ปฏิบัติการแทน อตร."
    for r in [f_row - 4, f_row]:
        ws[f'A{r}'].alignment = align_center; ws[f'C{r}'].alignment = align_center; ws[f'G{r}'].alignment = align_center

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

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
