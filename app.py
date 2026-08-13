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
                "รหัสบัญชี": str(row["รหัสบัญชี"]) if pd.notna(row["รหัสบัญชี"]) else "-",
                "Role": "-" 
            }
        return emp_dict
    except Exception as e:
        return None

if 'employees' not in st.session_state:
    loaded_data = load_employee_data()
    if not loaded_data:
        st.error("❌ ไม่พบไฟล์ 'ข้อมูล.xlsx' กรุณาอัปโหลดไว้ในโฟลเดอร์เดียวกับโค้ดครับ")
        st.stop()
    st.session_state.employees = loaded_data

# เพิ่มรหัส ป, ก, 00-12, 00-24 เข้าไปในระบบ
shift_data = {
    "ว": {"text": "(06.00-18.00) น.", "hours": 4},
    "ค": {"text": "(00.00-06.00)(18.00-24.00) น.", "hours": 4},
    "ว/ค": {"text": "(06.00-12.00)(18.00-24.00) น.", "hours": 4},
    "ค/ว": {"text": "(00.00-06.00)(12.00-18.00) น.", "hours": 4},
    "0-12": {"text": "(00.00-12.00) น.", "hours": 4},
    "00-12": {"text": "(00.00-12.00) น.", "hours": 4},
    "12-24": {"text": "(12.00-24.00) น.", "hours": 4},
    "00-24": {"text": "(00.00-24.00) น.", "hours": 4},
    "(ว)": {"text": "(06.00-18.00) น.", "hours": 4},
    "(ค)": {"text": "(00.00-06.00)(18.00-24.00) น.", "hours": 4},
    "(ว/ค)": {"text": "(06.00-12.00)(18.00-24.00) น.", "hours": 4},
    "(ค/ว)": {"text": "(00.00-06.00)(12.00-18.00) น.", "hours": 4},
    "(0-12)": {"text": "(00.00-12.00) น.", "hours": 4},
    "(00-12)": {"text": "(00.00-12.00) น.", "hours": 4},
    "(12-24)": {"text": "(12.00-24.00) น.", "hours": 4},
    "ย": {"text": "ย.", "hours": "-"},
    "พ": {"text": "พ.", "hours": "-"},
    "ป": {"text": "ป.", "hours": "-"},
    "ก": {"text": "ก.", "hours": "-"},
}

# ==========================================
# 2. ข้อมูลส่วนกลาง
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
# 3. ตารางเวร 109 & เพิ่มพนักงานใหม่
# ==========================================
st.markdown("---")
st.subheader("🗓️ 2. จัดการตารางเวร 1-31 วัน (ตารางหลัก)")

if 'roster_df' not in st.session_state:
    data = []
    for i, (name, info) in enumerate(st.session_state.employees.items()):
        row = {"ลำดับ": i+1, "ชื่อ-สกุล": name, "ตำแหน่ง": info['ตำแหน่ง']}
        for d in range(1, 32): row[str(d)] = ""
        data.append(row)
    df = pd.DataFrame(data)
    for d in range(1, 32): df[str(d)] = df[str(d)].astype(str)
    st.session_state.roster_df = df

with st.expander("➕ เพิ่มพนักงานใหม่ / คนเข้าเวรแทน (เฉพาะเดือนนี้)"):
    with st.form("add_emp_form"):
        st.markdown("ข้อมูล Role จะใช้ควบคุมการจัดเวรภายในระบบเท่านั้น จะไม่ปรากฏบนเอกสารใบเบิก")
        c1, c2, c3, c4 = st.columns(4)
        new_name = c1.text_input("ชื่อ-สกุล (ห้ามซ้ำ)*")
        new_pos = c2.text_input("ตำแหน่งที่ใช้เบิก*")
        new_id = c3.text_input("เลขประจำตัว", "-")
        new_role = c4.selectbox("Role (หน้าที่เข้าเวร)", ["นสน.", "ช.นสน.1", "ช.นสน.2", "เสมียน", "ประแจ", "กั้นถนนฯฉิมพลี", "กั้นถนนฯบางระมาด"])
        
        c5, c6, c7, c8 = st.columns(4)
        new_salary = c5.text_input("เงินเดือน", "-")
        new_rate = c6.number_input("เรท 1 ชั่วโมง (บาท)", min_value=0.0, value=0.0, step=1.0)
        new_acctype = c7.text_input("ประเภทบัญชี", "-")
        new_acccode = c8.text_input("รหัสบัญชี", "-")
        
        submitted = st.form_submit_button("เพิ่มรายชื่อลงตารางเวร")
        if submitted:
            if new_name.strip() == "" or new_pos.strip() == "":
                st.error("กรุณากรอก ชื่อ-สกุล และ ตำแหน่งที่ใช้เบิก!")
            elif new_name in st.session_state.employees:
                st.error("มีชื่อพนักงานคนนี้ในตารางอยู่แล้ว!")
            else:
                st.session_state.employees[new_name] = {
                    "ตำแหน่ง": new_pos, "เลขประจำตัว": new_id, "เงินเดือน": new_salary,
                    "เรท": new_rate, "ประเภทบัญชี": new_acctype, "รหัสบัญชี": new_acccode,
                    "Role": new_role 
                }
                new_idx = len(st.session_state.roster_df) + 1
                new_row = {"ลำดับ": new_idx, "ชื่อ-สกุล": new_name, "ตำแหน่ง": new_pos}
                for d in range(1, 32): new_row[str(d)] = ""
                new_df = pd.DataFrame([new_row])
                st.session_state.roster_df = pd.concat([st.session_state.roster_df, new_df], ignore_index=True)
                st.success(f"เพิ่ม '{new_name}' (Role: {new_role}) ลงในตารางเรียบร้อยแล้ว!")
                st.rerun()

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
# 3.5 ระบบตรวจสอบเงื่อนไข (Validation)
# ==========================================
def get_shift_hours(shift_str):
    if not shift_str: return set()
    shift = str(shift_str).strip().replace('(', '').replace(')', '')
    mapping = {
        'ว': set(range(6, 18)),
        'ค': set(range(0, 6)) | set(range(18, 24)),
        'ว/ค': set(range(6, 12)) | set(range(18, 24)),
        'ค/ว': set(range(0, 6)) | set(range(12, 18)),
        '0-12': set(range(0, 12)),
        '00-12': set(range(0, 12)),
        '12-24': set(range(12, 24)),
        '00-24': set(range(0, 24))
    }
    return mapping.get(shift, set())

def validate_roster(roster_df):
    errors = []
    for d in range(1, 32):
        day = str(d)
        role_hours = {}
        
        nai_satanee_off = False
        nai_satanee_shift = ""
        nai_satanee_name = ""
        
        for idx, row in roster_df.iterrows():
            name = str(row['ชื่อ-สกุล']).strip()
            pos = str(row['ตำแหน่ง']).strip()
            shift = str(row[day]).strip()
            if not shift: continue
            
            emp_info = st.session_state.employees.get(name, {})
            role = emp_info.get("Role", "-")
            
            # ตีความ Logical Role จากตำแหน่งของพนักงานประจำ
            logical_role = role
            if role == "-":
                p_clean = pos.replace(" ", "")
                if "นายสถานี" in p_clean: logical_role = "นายสถานี"
                elif "ช.นสน.ตช" in p_clean: logical_role = "ช.นสน.1"
                elif "ช.นสน.ตค" in p_clean: logical_role = "ช.นสน.2"
                elif "เสมียน" in p_clean: logical_role = "เสมียน"
                elif "ประแจ" in p_clean: logical_role = "ประแจ"
                elif "ฉิมพลี" in p_clean: logical_role = "กั้นถนนฯฉิมพลี"
                elif "บางระมาด" in p_clean: logical_role = "กั้นถนนฯบางระมาด"
                else: logical_role = pos
            
            # เก็บข้อมูลการลาของนายสถานี (Rule 2)
            if logical_role == "นายสถานี":
                nai_satanee_shift = shift
                nai_satanee_name = name
                if shift in ['ย', 'พ', 'ป', 'ก', 'ย.', 'พ.', 'ป.', 'ก.']:
                    nai_satanee_off = True
                    
            h_set = get_shift_hours(shift)
            
            # จับกลุ่ม ช.นสน.1 และ ช.นสน.2 ไว้ด้วยกันเพื่อเช็คเวลาซ้อนทับ (Rule 3 & 4)
            group_key = logical_role
            if logical_role in ["ช.นสน.1", "ช.นสน.2"]:
                group_key = "ช.นสน.รวม"
                
            if group_key not in role_hours: role_hours[group_key] = []
            role_hours[group_key].append((name, shift, h_set, logical_role))
            
        # ตรวจสอบเงื่อนไข Rule 1 & 2
        for name, shift, h_set, l_role in role_hours.get("นสน.", []):
            if not nai_satanee_off and len(h_set) > 0:
                errors.append(f"🔴 วันที่ {d}: '{name}' (นสน.) ไม่สามารถเข้าเวรได้ เนื่องจาก '{nai_satanee_name or 'นายสถานีฯ'}' ไม่ได้ลาหยุด (ลงเวร {nai_satanee_shift})")
            
            allowed_ntn = ['ว', 'ค', 'ค/ว', 'ว/ค', '00-12', '0-12', '12-24', '00-24']
            clean_shift = shift.replace('(', '').replace(')', '')
            if len(h_set) > 0 and clean_shift not in allowed_ntn:
                errors.append(f"🔴 วันที่ {d}: '{name}' (นสน.) ลงเวร '{shift}' ผิดเงื่อนไข (อนุญาตเฉพาะ ว, ค, ค/ว, ว/ค, 00-12, 12-24, 00-24)")
        
        # ตรวจสอบการซ้อนทับกัน (Rule 4 & 5)
        for g_key, members in role_hours.items():
            if g_key == "นายสถานี": continue # ข้ามการเช็คซ้อนทับของนายสถานีเว้นแต่มีนายสถานีหลายคน
            
            for i in range(len(members)):
                for j in range(i+1, len(members)):
                    n1, s1, h1, lr1 = members[i]
                    n2, s2, h2, lr2 = members[j]
                    
                    if len(h1) > 0 and len(h2) > 0 and not h1.isdisjoint(h2): # ถ้าเวลาทำงานตัดกัน (ไม่ disjoint)
                        group_name = "ช.นสน." if g_key == "ช.นสน.รวม" else g_key
                        errors.append(f"🟠 วันที่ {d}: ตรวจพบเวลาซ้อนทับกันระหว่าง '{n1}' ({s1}) และ '{n2}' ({s2}) ในหน้าที่ {group_name}")
                        
    return errors

st.markdown("---")
st.subheader("✅ 3. ตรวจสอบเงื่อนไขตารางเวร")
if st.button("🔍 กดเพื่อตรวจสอบความถูกต้อง (Validate)", type="secondary"):
    errors = validate_roster(st.session_state.roster_df)
    if len(errors) == 0:
        st.success("🎉 สมบูรณ์แบบ! ตารางเวรถูกต้องตามเงื่อนไขทุกประการ ไม่มีเวลาซ้อนทับกันครับ")
    else:
        st.error(f"พบข้อผิดพลาด {len(errors)} รายการ กรุณาแก้ไขก่อนส่งออกเอกสาร:")
        for e in errors:
            st.warning(e)


# ==========================================
# 4. ฟังก์ชันสร้างไฟล์ 109
# ==========================================
def generate_109(global_vars, roster_df):
    try:
        wb = openpyxl.load_workbook("109เปล่า.xlsx")
    except FileNotFoundError:
        st.error("❌ ไม่พบเทมเพลต '109เปล่า.xlsx' ในระบบ")
        return None
    ws = wb.active
    
    replacements_109 = {
        "[14]": global_vars["val_14"], "[13]": global_vars["val_13"],
        "[8]": global_vars["val_8"], "[7]": global_vars["val_7"]
    }
    
    for r in range(1, 100):
        for c in range(1, 40):
            cell = ws.cell(row=r, column=c)
            if cell.value and isinstance(cell.value, str) and "[" in cell.value and "]" in cell.value:
                new_val = cell.value
                for key, val in replacements_109.items():
                    new_val = new_val.replace(key, str(val))
                cell.value = new_val
                
    for idx, row_data in roster_df.iterrows():
        emp_name = row_data['ชื่อ-สกุล'].strip()
        emp_pos = row_data['ตำแหน่ง'].strip()
        
        found_row = None
        for r in range(1, 100):
            name_val = ws.cell(row=r, column=2).value
            if name_val and isinstance(name_val, str) and name_val.strip() == emp_name:
                found_row = r
                break
                
        if not found_row:
            for r in range(7, 100): 
                if ws.cell(row=r, column=1).value is None and ws.cell(row=r, column=2).value is None:
                    ws.cell(row=r, column=1, value=idx+1)
                    ws.cell(row=r, column=2, value=emp_name)
                    ws.cell(row=r+1, column=2, value=emp_pos)
                    found_row = r
                    break
        
        if found_row:
            for d in range(1, 32):
                shift = row_data[str(d)]
                ws.cell(row=found_row, column=2+d, value=shift if pd.notna(shift) else "")
                    
    wsp = ws.sheet_properties
    if not wsp.pageSetUpPr:
        wsp.pageSetUpPr = PageSetupProperties()
    wsp.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = False
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
    emp_info = st.session_state.employees[emp_name]
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
        "[3]": f"{float(emp_info['เงินเดือน']):,.0f}" if (emp_info['เงินเดือน'] != "-" and str(emp_info['เงินเดือน']).isnumeric()) else emp_info['เงินเดือน'],
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
# 6. เมนูส่งออก (Export)
# ==========================================
st.markdown("---")
st.subheader("🖨️ 4. ส่งออกเอกสาร (Export)")
tab1, tab2 = st.tabs(["📄 ส่งออก 109 (ตารางเวรรวม)", "🧾 ส่งออก 177 (ใบเบิกรายบุคคล)"])

with tab1:
    st.markdown("**ตารางเวร 109 ของทั้งสถานี**")
    if st.button("คลิกเพื่อสร้างไฟล์ 109", type="primary"):
        excel_109 = generate_109(global_data, st.session_state.roster_df)
        if excel_109:
            st.success(f"สร้างไฟล์ 109 ประจำเดือน {global_data['val_13']} สำเร็จ!")
            st.download_button("📥 ดาวน์โหลดไฟล์ 109", data=excel_109, file_name=f"109_{global_data['val_13']}.xlsx")

with tab2:
    st.markdown("**ข้อมูลเฉพาะบุคคลสำหรับใบเบิก 177**")
    selected_emp_177 = st.selectbox("เลือกพนักงานที่ต้องการเบิก 177", list(st.session_state.employees.keys()))
    
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
        emp_row = st.session_state.roster_df[st.session_state.roster_df['ชื่อ-สกุล'] == selected_emp_177].iloc[0]
        roster_dict = {str(d): str(emp_row[str(d)]) if pd.notna(emp_row[str(d)]) else "" for d in range(1, 32)}
        
        ind_data = st.session_state[f'ind_data_{selected_emp_177}']
        excel_177 = generate_177(selected_emp_177, roster_dict, global_data, ind_data)
        
        if excel_177:
            st.success(f"สร้างใบเบิก 177 ของ {selected_emp_177} เสร็จสิ้น!")
            st.download_button("📥 ดาวน์โหลดไฟล์ 177", data=excel_177, file_name=f"ใบเบิก_{selected_emp_177}.xlsx")
