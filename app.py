import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.worksheet.properties import PageSetupProperties
import io
import datetime
import re
import json
import uuid
from streamlit_local_storage import LocalStorage
import streamlit.components.v1 as components

st.set_page_config(page_title="SRT Timesheet App", layout="wide")
st.title("🚂 ระบบจัดการเวรและใบเบิกค่าตอบแทน (รฟท.)")
st.markdown("---")

# ==========================================
# 0. ระบบดักจับการรีเฟรช และ Local Storage
# ==========================================
components.html("""
    <script>
        window.parent.addEventListener('beforeunload', function (e) {
            e.preventDefault();
            e.returnValue = 'คุณมีข้อมูลที่ยังไม่ได้บันทึก แน่ใจหรือไม่ว่าต้องการออกจากหน้านี้?';
        });
    </script>
""", height=0, width=0)

local_storage = LocalStorage()

def save_roster_to_local(df):
    if df is not None and not df.empty:
        json_str = df.to_json(orient='records')
        local_storage.setItem("srt_roster_data", json_str, key=f"ls_roster_{uuid.uuid4().hex}")

def load_roster_from_local():
    try:
        json_str = local_storage.getItem("srt_roster_data")
        if json_str:
            loaded_df = pd.read_json(io.StringIO(json_str), orient='records')
            for d in range(1, 32):
                if str(d) in loaded_df.columns:
                    loaded_df[str(d)] = loaded_df[str(d)].astype(str).replace('nan', '')
            return loaded_df
    except:
        return None
    return None

# ==========================================
# 1. เมนูแถบด้านข้าง (รีเซ็ตระบบ & สำรองข้อมูล)
# ==========================================
st.sidebar.subheader("🔄 เริ่มต้นเดือนใหม่")
if st.sidebar.button("🗑️ ล้างข้อมูล (เพื่ออัปโหลดรายชื่อใหม่)", type="primary"):
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    components.html("<script>localStorage.clear(); window.parent.location.reload();</script>", height=0)
    st.stop()

st.sidebar.markdown("---")
st.sidebar.subheader("💾 จัดการข้อมูลสำรอง (ตารางเวร)")
if 'roster_df' in st.session_state and st.session_state.roster_df is not None:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        st.session_state.roster_df.to_excel(writer, index=False, sheet_name='Roster')
    st.sidebar.download_button(
        label="📥 ดาวน์โหลดไฟล์สำรองตารางเวร (.xlsx)", 
        data=buffer, 
        file_name=f"backup_roster_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx", 
        mime="application/vnd.ms-excel"
    )

uploaded_backup = st.sidebar.file_uploader("📤 อัปโหลดไฟล์สำรอง (.xlsx) เพื่อทำงานต่อ", type=["xlsx"])
if uploaded_backup:
    if st.sidebar.button("ยืนยันโหลดตารางเวรเก่า"):
        try:
            loaded_df = pd.read_excel(uploaded_backup)
            for d in range(1, 32):
                if str(d) in loaded_df.columns:
                    loaded_df[str(d)] = loaded_df[str(d)].astype(str).replace('nan', '')
            st.session_state.roster_df = loaded_df
            save_roster_to_local(loaded_df)
            st.sidebar.success("โหลดข้อมูลสำเร็จ! 🎉")
            st.rerun()
        except Exception as e:
            st.sidebar.error(f"Error: {e}")

# ==========================================
# 2. 🛡️ ระบบโหลดข้อมูลพนักงานแบบปลอดภัย
# ==========================================
saved_emp_json = local_storage.getItem("srt_employees_data")

if saved_emp_json and 'employees' not in st.session_state:
    try: st.session_state.employees = json.loads(saved_emp_json)
    except: st.session_state.employees = None

if 'employees' not in st.session_state or not st.session_state.employees:
    if saved_emp_json:
        st.success("✅ **ตรวจพบข้อมูลตารางเวรที่คุณทำค้างไว้ในเครื่อง!**")
        if st.button("🔄 กู้คืนข้อมูลล่าสุดกลับมาทำงานต่อ", type="primary", use_container_width=True):
            st.session_state.employees = json.loads(saved_emp_json)
            st.rerun()
            
    st.warning("🔒 **ระบบความปลอดภัย:** ไม่พบข้อมูลพนักงานในระบบ กรุณาอัปโหลดไฟล์เพื่อเริ่มต้น")
    uploaded_emp_file = st.file_uploader("📂 อัปโหลดไฟล์ 'ข้อมูล.xlsx' ของสถานีคุณ", type=["xlsx"])
    
    if uploaded_emp_file is not None:
        try:
            df_emp = pd.read_excel(uploaded_emp_file)
            emp_dict = {}
            for _, row in df_emp.iterrows():
                name = str(row.get("รายชื่อ", "")).strip()
                if not name or name == "nan": continue
                pos = str(row.get("ตำแหน่ง", "-")).strip()
                p_clean = pos.replace(" ", "")
                
                role = pos
                if "นายสถานี" in p_clean: role = "นสน."
                elif "ช.นสน.ตช" in p_clean: role = "ช.นสน.1"
                elif "ช.นสน.ตค" in p_clean: role = "ช.นสน.2"
                elif "เสมียน" in p_clean: role = "เสมียน"
                elif "ประแจ" in p_clean: role = "ประแจ"
                elif "ฉิมพลี" in p_clean: role = "กั้นถนนฯฉิมพลี"
                elif "บางระมาด" in p_clean: role = "กั้นถนนฯบางระมาด"
                elif "บริการ" in p_clean or "ลูกจ้าง" in p_clean: role = "ลูกจ้าง"
                elif "กั้นถนน" in p_clean: role = "อื่นๆ"
                else: role = "อื่นๆ"
                
                unique_key = f"{name}_{role}"
                emp_dict[unique_key] = {
                    "ชื่อ-สกุล": name, "ตำแหน่ง": pos, 
                    "เลขประจำตัว": str(row.get("เลขประจำตัว", "-")) if pd.notna(row.get("เลขประจำตัว")) else "-",
                    "เงินเดือน": str(row.get("เงินเดือน", "-")) if pd.notna(row.get("เงินเดือน")) else "-", 
                    "เรท": float(row.get("เรท 1 ชั่วโมง", 0)) if pd.notna(row.get("เรท 1 ชั่วโมง")) else 0.0,
                    "ประเภทบัญชี": str(row.get("ประเภทบัญชี", "-")) if pd.notna(row.get("ประเภทบัญชี")) else "-", 
                    "รหัสบัญชี": str(row.get("รหัสบัญชี", "-")) if pd.notna(row.get("รหัสบัญชี")) else "-",
                    "Role": role, "is_regular": True 
                }
            st.session_state.employees = emp_dict
            local_storage.setItem("srt_employees_data", json.dumps(emp_dict), key=f"ls_emp_{uuid.uuid4().hex}")
            st.success("✅ โหลดข้อมูลพนักงานสำเร็จ! กำลังเข้าสู่ระบบ...")
            st.rerun()
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการอ่านไฟล์: {e}")
    st.stop()

# ==========================================
# 3. ข้อมูลตารางเวรตั้งต้น
# ==========================================
shift_data = {
    "ว": {"text": "(06.00-18.00) น.", "hours": 4}, "ค": {"text": "(00.00-06.00)(18.00-24.00) น.", "hours": 4},
    "ว/ค": {"text": "(06.00-12.00)(18.00-24.00) น.", "hours": 4}, "ค/ว": {"text": "(00.00-06.00)(12.00-18.00) น.", "hours": 4},
    "0-12": {"text": "(00.00-12.00) น.", "hours": 4}, "00-12": {"text": "(00.00-12.00) น.", "hours": 4},
    "12-24": {"text": "(12.00-24.00) น.", "hours": 4}, "00-24": {"text": "(00.00-24.00) น.", "hours": 4},
    "(ว)": {"text": "(06.00-18.00) น.", "hours": 4}, "(ค)": {"text": "(00.00-06.00)(18.00-24.00) น.", "hours": 4},
    "(ว/ค)": {"text": "(06.00-12.00)(18.00-24.00) น.", "hours": 4}, "(ค/ว)": {"text": "(00.00-06.00)(12.00-18.00) น.", "hours": 4},
    "(0-12)": {"text": "(00.00-12.00) น.", "hours": 4}, "(00-12)": {"text": "(00.00-12.00) น.", "hours": 4},
    "(12-24)": {"text": "(12.00-24.00) น.", "hours": 4},
    "ย": {"text": "ย.", "hours": "-"}, "ย.": {"text": "ย.", "hours": "-"},
    "พ": {"text": "พ.", "hours": "-"}, "พ.": {"text": "พ.", "hours": "-"},
    "ป": {"text": "ป.", "hours": "-"}, "ป.": {"text": "ป.", "hours": "-"},
    "ก": {"text": "ก.", "hours": "-"}, "ก.": {"text": "ก.", "hours": "-"},
    "น": {"text": "น.", "hours": "-"}, "น.": {"text": "น.", "hours": "-"},
    "ล": {"text": "ล.", "hours": "-"}, "ล.": {"text": "ล.", "hours": "-"}, "ลา": {"text": "ลา", "hours": "-"},
}
roles_list = ["นสน.", "ช.นสน.1", "ช.นสน.2", "เสมียน", "ประแจ", "กั้นถนนฯฉิมพลี", "กั้นถนนฯบางระมาด", "ลูกจ้าง", "อื่นๆ"]

def sort_roster_by_role(df, emp_dict):
    temp_df = df.copy()
    role_last_idx = {}
    for idx, row in temp_df.iterrows():
        name, role = str(row['ชื่อ-สกุล']).strip(), str(row['Role (หน้าที่)']).strip()
        info = emp_dict.get(f"{name}_{role}", {})
        if info.get('is_regular', False): role_last_idx[role] = idx
            
    def get_sort_key(row):
        name, role = str(row['ชื่อ-สกุล']).strip(), str(row['Role (หน้าที่)']).strip()
        info = emp_dict.get(f"{name}_{role}", {})
        if info.get('is_regular', False): return row.name * 1000 
        else:
            if role in role_last_idx: return role_last_idx[role] * 1000 + row.name + 1 
            else: return 999000 + row.name 
                
    temp_df['sort_key'] = temp_df.apply(get_sort_key, axis=1)
    temp_df = temp_df.sort_values('sort_key').reset_index(drop=True)
    temp_df['ลำดับ'] = range(1, len(temp_df) + 1)
    return temp_df.drop(columns=['sort_key'])

if 'roster_df' not in st.session_state:
    saved_df = load_roster_from_local()
    if saved_df is not None:
        st.session_state.roster_df = saved_df
        for _, row in saved_df.iterrows():
            name = str(row['ชื่อ-สกุล']).strip()
            role = str(row['Role (หน้าที่)']).strip()
            unique_key = f"{name}_{role}"
            if unique_key not in st.session_state.employees:
                 st.session_state.employees[unique_key] = {
                        "ชื่อ-สกุล": name, "ตำแหน่ง": str(row['ตำแหน่งเบิก']).strip(), "เลขประจำตัว": "-", 
                        "เงินเดือน": "-", "เรท": 0.0, "ประเภทบัญชี": "-", "รหัสบัญชี": "-", "Role": role, "is_regular": False
                 }
    else:
        data = []
        for i, (key, info) in enumerate(st.session_state.employees.items()):
            row = {"ขึ้นหน้าใหม่": False, "ลำดับ": i+1, "ชื่อ-สกุล": info['ชื่อ-สกุล'], "ตำแหน่งเบิก": info['ตำแหน่ง'], "Role (หน้าที่)": info['Role']}
            for d in range(1, 32): row[str(d)] = ""
            data.append(row)
        df = pd.DataFrame(data)
        for d in range(1, 32): df[str(d)] = df[str(d)].astype(str)
        st.session_state.roster_df = sort_roster_by_role(df, st.session_state.employees)
        save_roster_to_local(st.session_state.roster_df)

if 'ขึ้นหน้าใหม่' not in st.session_state.roster_df.columns:
    st.session_state.roster_df.insert(0, 'ขึ้นหน้าใหม่', False)

# ==========================================
# 4. ตั้งค่าส่วนกลาง & ตารางเวร
# ==========================================
st.subheader("⚙️ 1. ตั้งค่าข้อมูลส่วนกลางประจำเดือน")
col_g1, col_g2 = st.columns(2)

saved_global = local_storage.getItem("srt_global_data")
default_global = {"val_13": "สิงหาคม", "val_7": "5110/2520/2569", "val_8": "29 พ.ค. 69", "val_14": "01 ก.ค. 69"}
try:
    if saved_global: default_global.update(json.loads(saved_global))
except: pass

with col_g1:
    val_13 = st.text_input("เดือนตัวเต็ม [13]", default_global["val_13"])
    val_7 = st.text_input("คำสั่งแขวง [7]", default_global["val_7"])
with col_g2:
    val_8 = st.text_input("วันที่ลงคำสั่ง [8]", default_global["val_8"])
    val_14 = st.text_input("วันที่เซ็นเอกสารตัวย่อ [14]", default_global["val_14"])

global_data = {"val_13": val_13, "val_7": val_7, "val_8": val_8, "val_14": val_14}
local_storage.setItem("srt_global_data", json.dumps(global_data), key=f"ls_global_{uuid.uuid4().hex}")

st.markdown("---")
st.subheader("🗓️ 2. จัดการตารางเวร 1-31 วัน")

with st.expander("➕ เพิ่มพนักงานใหม่ / เข้าเวรแทน"):
    with st.form("add_emp_form"):
        c1, c2, c3, c4 = st.columns(4)
        new_name = c1.text_input("ชื่อ-สกุล*")
        new_pos = c2.text_input("ตำแหน่งเบิก*")
        new_role = c3.selectbox("Role (หน้าที่)", roles_list)
        new_rate = c4.number_input("เรท 1 ชม. (บาท)", min_value=0.0, value=0.0)
        
        if st.form_submit_button("เพิ่มพนักงาน"):
            if new_name.strip() == "" or new_pos.strip() == "": st.error("กรุณากรอก ชื่อ และ ตำแหน่ง!")
            else:
                unique_key = f"{new_name}_{new_role}"
                if unique_key not in st.session_state.employees:
                    st.session_state.employees[unique_key] = {
                        "ชื่อ-สกุล": new_name, "ตำแหน่ง": new_pos, "เลขประจำตัว": "-", 
                        "เงินเดือน": "-", "เรท": new_rate, "ประเภทบัญชี": "-", "รหัสบัญชี": "-", "Role": new_role, "is_regular": False
                    }
                    new_idx = len(st.session_state.roster_df) + 1
                    new_row = {"ขึ้นหน้าใหม่": False, "ลำดับ": new_idx, "ชื่อ-สกุล": new_name, "ตำแหน่งเบิก": new_pos, "Role (หน้าที่)": new_role}
                    for d in range(1, 32): new_row[str(d)] = ""
                    new_df = pd.DataFrame([new_row])
                    updated_df = pd.concat([st.session_state.roster_df, new_df], ignore_index=True)
                    st.session_state.roster_df = sort_roster_by_role(updated_df, st.session_state.employees)
                    save_roster_to_local(st.session_state.roster_df)
                    local_storage.setItem("srt_employees_data", json.dumps(st.session_state.employees), key=f"ls_emp_{uuid.uuid4().hex}")
                    st.rerun()

column_config = {
    "ขึ้นหน้าใหม่": st.column_config.CheckboxColumn("ขึ้นหน้าใหม่", width="small"),
    "ลำดับ": st.column_config.NumberColumn("ลำดับ", width="small", disabled=True),
    "ชื่อ-สกุล": st.column_config.TextColumn("ชื่อ-สกุล", width="medium"), 
    "ตำแหน่งเบิก": st.column_config.TextColumn("ตำแหน่งเบิก", width="medium"), 
    "Role (หน้าที่)": st.column_config.SelectboxColumn("Role", options=roles_list, width="small")
}
for d in range(1, 32): column_config[str(d)] = st.column_config.TextColumn(str(d), width="small")

edited_df = st.data_editor(st.session_state.roster_df, hide_index=True, use_container_width=True, column_config=column_config, key="roster_table")

if not edited_df.equals(st.session_state.roster_df):
    st.session_state.roster_df = edited_df
    save_roster_to_local(edited_df)

for _, row in edited_df.iterrows():
    name = str(row.get('ชื่อ-สกุล', '')).strip()
    role = str(row.get('Role (หน้าที่)', '')).strip()
    if not name: continue
    key = f"{name}_{role}"
    if key in st.session_state.employees:
        st.session_state.employees[key]['Role'] = role
        st.session_state.employees[key]['ชื่อ-สกุล'] = name
        st.session_state.employees[key]['ตำแหน่ง'] = row.get('ตำแหน่งเบิก', '')

# ==========================================
# 5. ฟังก์ชันสร้างไฟล์ Excel (109, 177, และรายงานปฏิบัติงาน)
# ==========================================
def generate_109(global_vars, roster_df):
    try: wb = openpyxl.load_workbook("109เปล่า.xlsx")
    except: return None, 0
    pages, current_page_rows = [], []
    for idx, row in roster_df.iterrows():
        if len(current_page_rows) >= 15 or (row.get("ขึ้นหน้าใหม่", False) and len(current_page_rows) > 0):
            pages.append(current_page_rows)
            current_page_rows = []
        current_page_rows.append(row)
    if current_page_rows: pages.append(current_page_rows)
    total_pages = max(len(pages), 1)
    if len(pages) == 0: pages = [[]]
    template_ws = wb.active
    template_ws.title = "หน้าที่ 1"
    worksheets = [template_ws]
    for p in range(2, total_pages + 1):
        new_ws = wb.copy_worksheet(template_ws)
        new_ws.title = f"หน้าที่ {p}"
        worksheets.append(new_ws)
    replacements_109 = {"[14]": global_vars["val_14"], "[13]": global_vars["val_13"], "[8]": global_vars["val_8"], "[7]": global_vars["val_7"]}
    
    for page_idx, ws in enumerate(worksheets):
        page_num = page_idx + 1
        page_data = pages[page_idx]
        for r in range(1, 15):
            for c in range(1, 40):
                c_cell = ws.cell(row=r, column=c)
                val = c_cell.value
                if val and isinstance(val, str):
                    new_val = val
                    for k, v in replacements_109.items(): new_val = new_val.replace(k, str(v))
                    if "หน้า" in new_val and "/" in new_val: new_val = re.sub(r'หน้า\s*\d+\s*/\s*\d+', f'หน้า {page_num}/{total_pages}', new_val)
                    if type(c_cell).__name__ != 'MergedCell': c_cell.value = new_val
        for r in range(39, 55):
            for c in range(1, 40):
                c_cell = ws.cell(row=r, column=c)
                val = c_cell.value
                if val and isinstance(val, str) and "[" in val:
                    new_val = val
                    for k, v in replacements_109.items(): new_val = new_val.replace(k, str(v))
                    if type(c_cell).__name__ != 'MergedCell': c_cell.value = new_val
        for r in range(8, 37, 2):
            ws.cell(row=r, column=1).value = ""
            ws.cell(row=r, column=2).value = ""
            ws.cell(row=r+1, column=2).value = ""
            for d in range(1, 32): ws.cell(row=r, column=2+d).value = ""
        current_excel_row = 8
        for row_data in page_data:
            ws.cell(row=current_excel_row, column=1).value = row_data['ลำดับ']
            ws.cell(row=current_excel_row, column=2).value = str(row_data['ชื่อ-สกุล']).strip()
            ws.cell(row=current_excel_row+1, column=2).value = str(row_data['ตำแหน่งเบิก']).strip()
            for d in range(1, 32):
                shift = row_data.get(str(d), "")
                ws.cell(row=current_excel_row, column=2+d).value = str(shift).strip() if pd.notna(shift) else ""
            current_excel_row += 2
        if not ws.sheet_properties.pageSetUpPr: ws.sheet_properties.pageSetUpPr = PageSetupProperties()
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0 
        ws.page_setup.paperSize = ws.PAPERSIZE_A4; ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    # 📌 แก้ไขบรรทัดนี้แล้วครับ ส่งค่า output และ total_pages กลับไปทั้งคู่
    return output, total_pages 

def generate_177(unique_key, roster_data, global_vars, ind_vars):
    emp_info = st.session_state.employees.get(unique_key)
    if not emp_info: return None
    rate = emp_info["เรท"]
    try: wb = openpyxl.load_workbook("ใบเบิก177 Update.xlsx")
    except: return None
    ws = wb.active
    replacements = {
        "[NAME]": emp_info["ชื่อ-สกุล"], "[16]": emp_info["รหัสบัญชี"], "[15]": emp_info["ประเภทบัญชี"],
        "[14]": global_vars["val_14"], "[13]": global_vars["val_13"], "[12]": ind_vars["val_12"],
        "[11]": ind_vars["val_11"], "[10]": ind_vars["val_10"], "[9]": ind_vars["val_9"],
        "[8]": global_vars["val_8"], "[7]": global_vars["val_7"], "[6]": ind_vars["val_6"],
        "[5]": ind_vars["val_5"], "[4]": ind_vars["val_4"],
        "[3]": f"{float(emp_info['เงินเดือน']):,.0f}" if (emp_info['เงินเดือน'] != "-" and str(emp_info['เงินเดือน']).isnumeric()) else emp_info['เงินเดือน'],
        "[2]": emp_info["เลขประจำตัว"], "[1]": emp_info["ตำแหน่ง"]
    }
    for r in range(1, 55):
        for c in range(1, 25): 
            c_cell = ws.cell(row=r, column=c)
            val = c_cell.value
            if val and isinstance(val, str) and "[" in val:
                new_val = val
                for k, v in replacements.items(): new_val = new_val.replace(k, str(v))
                if type(c_cell).__name__ != 'MergedCell': c_cell.value = new_val
    start_row = 7
    for day in range(1, 32):
        row = start_row + day - 1
        shift = str(roster_data.get(str(day), "")).strip()
        sData = shift_data.get(shift)
        if sData and sData["hours"] != "-":
            ws.cell(row=row, column=2, value=sData["text"])
            ws.cell(row=row, column=3, value=int(sData["hours"]))
            ws.cell(row=row, column=4, value=rate)
            ws.cell(row=row, column=5, value="00")
            ws.cell(row=row, column=6, value=int(sData["hours"]) * rate)
            ws.cell(row=row, column=7, value="00")
        else:
            val = sData["text"] if sData else (shift if shift else "-")
            ws.cell(row=row, column=2, value=val)
            for col in range(3, 8): 
                if type(ws.cell(row=row, column=col)).__name__ != 'MergedCell':
                    ws.cell(row=row, column=col, value="-")
    if not ws.sheet_properties.pageSetUpPr: ws.sheet_properties.pageSetUpPr = PageSetupProperties()
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToHeight = 1; ws.page_setup.fitToWidth = 1
    ws.page_setup.paperSize = ws.PAPERSIZE_A4; ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def generate_report_work(unique_key, roster_data, global_vars):
    emp_info = st.session_state.employees.get(unique_key)
    if not emp_info: return None
    
    try: wb = openpyxl.load_workbook("รายงานปฏิบัติงาน.xlsx")
    except: return None
    ws = wb.active

    replacements = {
        "[NAME]": emp_info["ชื่อ-สกุล"],
        "[1]": emp_info["ตำแหน่ง"],
        "[2]": emp_info["เลขประจำตัว"],
        "[3]": f"{float(emp_info['เงินเดือน']):,.0f}" if (emp_info['เงินเดือน'] != "-" and str(emp_info['เงินเดือน']).isnumeric()) else emp_info['เงินเดือน'],
        "[14]": global_vars["val_14"],
        "[13]": global_vars["val_13"],
        "[8]": global_vars["val_8"], 
        "[7]": global_vars["val_7"],
    }

    for r in range(1, 55):
        for c in range(1, 25):
            c_cell = ws.cell(row=r, column=c)
            val = c_cell.value
            if val and isinstance(val, str):
                new_val = val
                for k, v in replacements.items(): new_val = new_val.replace(k, str(v))
                if type(c_cell).__name__ != 'MergedCell': 
                    c_cell.value = new_val

    if type(ws.cell(row=2, column=7)).__name__ != 'MergedCell':
        ws.cell(row=2, column=7).value = global_vars["val_13"]
    if type(ws.cell(row=44, column=2)).__name__ != 'MergedCell':
        ws.cell(row=44, column=2).value = emp_info["ชื่อ-สกุล"]

    start_row = 8
    
    ranges_to_unmerge = []
    for merged_range in ws.merged_cells.ranges:
        if start_row <= merged_range.min_row <= start_row + 31:
            if merged_range.min_col <= 7 and merged_range.max_col >= 2:
                ranges_to_unmerge.append(merged_range.coord)
    
    for r_coord in ranges_to_unmerge:
        ws.unmerge_cells(r_coord)

    leave_types = ["ย", "ย.", "พ", "พ.", "ป", "ป.", "ก", "ก.", "น", "น.", "ล", "ล.", "ลา"]
    
    for day in range(1, 32):
        row = start_row + day - 1
        shift = str(roster_data.get(str(day), "")).strip()
        
        start_time, end_time = "-", "-"
        is_holiday = False
        
        if shift:
            s_clean = shift.replace("(", "").replace(")", "")
            if s_clean == "ว": start_time, end_time = "06.00", "18.00"
            elif s_clean == "ค": start_time, end_time = "00.00-06.00", "18.00-24.00"
            elif s_clean == "ว/ค": start_time, end_time = "06.00-12.00", "18.00-24.00"
            elif s_clean == "ค/ว": start_time, end_time = "00.00-06.00", "12.00-18.00"
            elif s_clean in ["0-12", "00-12"]: start_time, end_time = "00.00", "12.00"
            elif s_clean == "12-24": start_time, end_time = "12.00", "24.00"
            elif s_clean == "00-24": start_time, end_time = "00.00", "24.00"
            elif s_clean in leave_types: 
                start_time = s_clean
                end_time = ""
                is_holiday = True 
            else: start_time, end_time = shift, ""
            
        for c in range(2, 8): ws.cell(row=row, column=c).value = None

        if is_holiday:
            ws.cell(row=row, column=2).value = start_time
            ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=7)
            ws.cell(row=row, column=2).alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
        else:
            ws.cell(row=row, column=2).value = start_time
            ws.cell(row=row, column=5).value = end_time

        c_role = ws.cell(row=row, column=8)
        if type(c_role).__name__ != 'MergedCell':
            if start_time not in ["-", ""] and start_time not in leave_types:
                c_role.value = emp_info["ตำแหน่ง"]
            else:
                c_role.value = ""
            
    if not ws.sheet_properties.pageSetUpPr: ws.sheet_properties.pageSetUpPr = PageSetupProperties()
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToHeight = 1; ws.page_setup.fitToWidth = 1
    ws.page_setup.paperSize = ws.PAPERSIZE_A4; ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ==========================================
# 6. เมนูส่งออก (Export)
# ==========================================
st.markdown("---")
st.subheader("🖨️ 4. ส่งออกเอกสาร Excel สำเร็จรูป")
tab1, tab2 = st.tabs(["📄 ส่งออก 109 (ตารางเวรรวม)", "🧾 ส่งออกเอกสารรายบุคคล (177 และ รายงานปฏิบัติงาน)"])

with tab1:
    st.markdown("**ตารางเวร 109 ของทั้งสถานี** (จัดหน้ากระดาษอัตโนมัติ)")
    if st.button("คลิกเพื่อสร้างไฟล์ 109", type="primary"):
        excel_109, total_pages = generate_109(global_data, st.session_state.roster_df)
        if excel_109:
            st.success(f"สร้างไฟล์ 109 เสร็จสิ้น! (รวม {total_pages} หน้า)")
            st.download_button("📥 ดาวน์โหลดไฟล์ 109 (.xlsx)", data=excel_109, file_name=f"109_{global_data['val_13']}.xlsx")

with tab2:
    st.markdown("**ข้อมูลเฉพาะบุคคลสำหรับใบเบิก**")
    active_emp_options = [f"{r['ชื่อ-สกุล']} ({r['Role (หน้าที่)']})" for _, r in st.session_state.roster_df.iterrows()]
    selected_key_177_display = st.selectbox("เลือกพนักงานที่ต้องการสร้างเอกสาร", active_emp_options)
    
    if selected_key_177_display:
        sel_name = selected_key_177_display.split(" (")[0]
        sel_role = selected_key_177_display.split(" (")[1].replace(")", "")
        selected_key_177 = f"{sel_name}_{sel_role}"
        
        saved_ind = local_storage.getItem(f"srt_ind_{selected_key_177}")
        default_ind = {"val_4": "-", "val_5": "-", "val_6": "-", "val_9": "-", "val_10": "-", "val_11": "-", "val_12": "-"}
        try:
            if saved_ind: default_ind.update(json.loads(saved_ind))
        except: pass

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            default_ind['val_4'] = st.text_input("วันหยุด [4]", default_ind['val_4'])
            default_ind['val_9'] = st.text_input("หยุดพักผ่อน [9]", default_ind['val_9'])
        with c2:
            default_ind['val_5'] = st.text_input("วันหยุดที่เบิกได้ [5]", default_ind['val_5'])
            default_ind['val_10'] = st.text_input("พักผ่อนตั้งแต่ [10]", default_ind['val_10'])
        with c3:
            default_ind['val_6'] = st.text_input("เดือนตัวย่อ [6]", default_ind['val_6'])
            default_ind['val_11'] = st.text_input("พักผ่อนถึง [11]", default_ind['val_11'])
        with c4:
            default_ind['val_12'] = st.text_input("รวมพักผ่อน [12]", default_ind['val_12'])

        local_storage.setItem(f"srt_ind_{selected_key_177}", json.dumps(default_ind), key=f"ls_ind_{uuid.uuid4().hex}")

        st.markdown("---")
        col_btn1, col_btn2 = st.columns(2)
        
        with col_btn1:
            if st.button(f"🧾 ออกใบเบิก 177", type="primary", use_container_width=True):
                emp_row = st.session_state.roster_df[
                    (st.session_state.roster_df['ชื่อ-สกุล'] == sel_name) & 
                    (st.session_state.roster_df['Role (หน้าที่)'] == sel_role)
                ].iloc[0]
                roster_dict = {str(d): str(emp_row[str(d)]) if pd.notna(emp_row[str(d)]) else "" for d in range(1, 32)}
                
                excel_177 = generate_177(selected_key_177, roster_dict, global_data, default_ind)
                if excel_177:
                    st.success(f"สร้างใบเบิก 177 เสร็จสิ้น!")
                    st.download_button("📥 ดาวน์โหลดไฟล์ 177 (.xlsx)", data=excel_177, file_name=f"177_{sel_name}.xlsx")
                    
        with col_btn2:
            if st.button(f"🕒 ออกรายงานปฏิบัติงาน", type="primary", use_container_width=True):
                emp_row = st.session_state.roster_df[
                    (st.session_state.roster_df['ชื่อ-สกุล'] == sel_name) & 
                    (st.session_state.roster_df['Role (หน้าที่)'] == sel_role)
                ].iloc[0]
                roster_dict = {str(d): str(emp_row[str(d)]) if pd.notna(emp_row[str(d)]) else "" for d in range(1, 32)}
                
                excel_work = generate_report_work(selected_key_177, roster_dict, global_data)
                if excel_work:
                    st.success(f"สร้างรายงานปฏิบัติงาน เสร็จสิ้น!")
                    st.download_button("📥 ดาวน์โหลดรายงานปฏิบัติงาน (.xlsx)", data=excel_work, file_name=f"รายงานปฏิบัติงาน_{sel_name}.xlsx")
