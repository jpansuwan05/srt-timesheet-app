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
            name = str(row["รายชื่อ"]).strip()
            pos = str(row["ตำแหน่ง"]).strip() if pd.notna(row["ตำแหน่ง"]) else "-"
            
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
                "ชื่อ-สกุล": name,
                "ตำแหน่ง": pos,
                "เลขประจำตัว": str(row["เลขประจำตัว"]) if pd.notna(row["เลขประจำตัว"]) else "-",
                "เงินเดือน": str(row["เงินเดือน"]) if pd.notna(row["เงินเดือน"]) else "-",
                "เรท": float(row["เรท 1 ชั่วโมง"]) if pd.notna(row["เรท 1 ชั่วโมง"]) else 0,
                "ประเภทบัญชี": str(row["ประเภทบัญชี"]) if pd.notna(row["ประเภทบัญชี"]) else "-",
                "รหัสบัญชี": str(row["รหัสบัญชี"]) if pd.notna(row["รหัสบัญชี"]) else "-",
                "Role": role,
                "is_regular": True # แท็กว่าเป็นพนักงานประจำ
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

shift_data = {
    "ว": {"text": "(06.00-18.00) น.", "hours": 4}, "ค": {"text": "(00.00-06.00)(18.00-24.00) น.", "hours": 4},
    "ว/ค": {"text": "(06.00-12.00)(18.00-24.00) น.", "hours": 4}, "ค/ว": {"text": "(00.00-06.00)(12.00-18.00) น.", "hours": 4},
    "0-12": {"text": "(00.00-12.00) น.", "hours": 4}, "00-12": {"text": "(00.00-12.00) น.", "hours": 4},
    "12-24": {"text": "(12.00-24.00) น.", "hours": 4}, "00-24": {"text": "(00.00-24.00) น.", "hours": 4},
    "(ว)": {"text": "(06.00-18.00) น.", "hours": 4}, "(ค)": {"text": "(00.00-06.00)(18.00-24.00) น.", "hours": 4},
    "(ว/ค)": {"text": "(06.00-12.00)(18.00-24.00) น.", "hours": 4}, "(ค/ว)": {"text": "(00.00-06.00)(12.00-18.00) น.", "hours": 4},
    "(0-12)": {"text": "(00.00-12.00) น.", "hours": 4}, "(00-12)": {"text": "(00.00-12.00) น.", "hours": 4},
    "(12-24)": {"text": "(12.00-24.00) น.", "hours": 4},
    "ย": {"text": "ย.", "hours": "-"}, "พ": {"text": "พ.", "hours": "-"},
    "ป": {"text": "ป.", "hours": "-"}, "ก": {"text": "ก.", "hours": "-"},
}

roles_list = ["นสน.", "ช.นสน.1", "ช.นสน.2", "เสมียน", "ประแจ", "กั้นถนนฯฉิมพลี", "กั้นถนนฯบางระมาด", "ลูกจ้าง", "อื่นๆ"]

# --- ฟังก์ชันช่วยจัดเรียงชื่อคนมาแทนให้อยู่ต่อจากคนประจำ ---
def sort_roster_by_role(df, emp_dict):
    temp_df = df.copy()
    role_last_idx = {}
    
    # หาบรรทัดสุดท้ายของพนักงานประจำในแต่ละ Role
    for idx, row in temp_df.iterrows():
        name = str(row['ชื่อ-สกุล']).strip()
        role = str(row['Role (หน้าที่)']).strip()
        info = emp_dict.get(f"{name}_{role}", {})
        if info.get('is_regular', False):
            role_last_idx[role] = idx
            
    def get_sort_key(row):
        name = str(row['ชื่อ-สกุล']).strip()
        role = str(row['Role (หน้าที่)']).strip()
        info = emp_dict.get(f"{name}_{role}", {})
        
        if info.get('is_regular', False):
            return row.name * 1000 # คนประจำให้อยู่ที่เดิม
        else:
            if role in role_last_idx:
                return role_last_idx[role] * 1000 + row.name + 1 # คนมาแทน ให้ต่อท้ายคนประจำ
            else:
                return 999000 + row.name # ถ้าไม่มีคนประจำ Role นี้เลย ให้ไปต่อท้ายสุด
                
    temp_df['sort_key'] = temp_df.apply(get_sort_key, axis=1)
    temp_df = temp_df.sort_values('sort_key').reset_index(drop=True)
    temp_df['ลำดับ'] = range(1, len(temp_df) + 1)
    return temp_df.drop(columns=['sort_key'])

# ==========================================
# 2. ข้อมูลส่วนกลาง
# ==========================================
st.subheader("⚙️ 1. ตั้งค่าข้อมูลส่วนกลางประจำเดือน")
col_g1, col_g2 = st.columns(2)
with col_g1:
    val_13 = st.text_input("เดือนตัวเต็ม [13]", "มิถุนายน")
    val_7 = st.text_input("คำสั่งแขวง [7]", "5110/2520/2569")
with col_g2:
    val_8 = st.text_input("วันที่ลงคำสั่ง [8]", "29 พ.ค. 69")
    val_14 = st.text_input("วันที่เซ็นเอกสารตัวย่อ [14]", "01 ก.ค. 69")

global_data = {"val_13": val_13, "val_7": val_7, "val_8": val_8, "val_14": val_14}

# ==========================================
# 3. ตารางเวร 109 & เพิ่มพนักงานใหม่
# ==========================================
st.markdown("---")
st.subheader("🗓️ 2. จัดการตารางเวร 1-31 วัน (ตารางหลัก)")

if 'roster_df' not in st.session_state:
    data = []
    for i, (key, info) in enumerate(st.session_state.employees.items()):
        row = {"ลำดับ": i+1, "ชื่อ-สกุล": info['ชื่อ-สกุล'], "ตำแหน่งเบิก": info['ตำแหน่ง'], "Role (หน้าที่)": info['Role']}
        for d in range(1, 32): row[str(d)] = ""
        data.append(row)
    df = pd.DataFrame(data)
    for d in range(1, 32): df[str(d)] = df[str(d)].astype(str)
    # จัดเรียงตั้งแต่ตอนเริ่มต้น
    st.session_state.roster_df = sort_roster_by_role(df, st.session_state.employees)

with st.expander("➕ เพิ่มพนักงานใหม่ / เข้าเวรแทน (เสียบชื่อต่อจากคนประจำอัตโนมัติ)"):
    with st.form("add_emp_form"):
        c1, c2, c3, c4 = st.columns(4)
        new_name = c1.text_input("ชื่อ-สกุล*")
        new_pos = c2.text_input("ตำแหน่งที่ใช้เบิก*")
        new_id = c3.text_input("เลขประจำตัว", "-")
        new_role = c4.selectbox("Role (หน้าที่เข้าเวร)", roles_list)
        
        c5, c6, c7, c8 = st.columns(4)
        new_salary = c5.text_input("เงินเดือน", "-")
        new_rate = c6.number_input("เรท 1 ชั่วโมง (บาท)", min_value=0.0, value=0.0, step=1.0)
        new_acctype = c7.text_input("ประเภทบัญชี", "-")
        new_acccode = c8.text_input("รหัสบัญชี", "-")
        
        submitted = st.form_submit_button("เพิ่มลงตารางเวร")
        if submitted:
            if new_name.strip() == "" or new_pos.strip() == "":
                st.error("กรุณากรอก ชื่อ-สกุล และ ตำแหน่งที่ใช้เบิก!")
            else:
                unique_key = f"{new_name}_{new_role}"
                if unique_key in st.session_state.employees:
                    st.error(f"มีรายชื่อ '{new_name}' ใน Role '{new_role}' อยู่ในระบบแล้ว")
                else:
                    st.session_state.employees[unique_key] = {
                        "ชื่อ-สกุล": new_name, "ตำแหน่ง": new_pos, "เลขประจำตัว": new_id, 
                        "เงินเดือน": new_salary, "เรท": new_rate, "ประเภทบัญชี": new_acctype, 
                        "รหัสบัญชี": new_acccode, "Role": new_role,
                        "is_regular": False # แท็กว่าเป็นคนมาแทน
                    }
                    new_idx = len(st.session_state.roster_df) + 1
                    new_row = {"ลำดับ": new_idx, "ชื่อ-สกุล": new_name, "ตำแหน่งเบิก": new_pos, "Role (หน้าที่)": new_role}
                    for d in range(1, 32): new_row[str(d)] = ""
                    new_df = pd.DataFrame([new_row])
                    
                    # เพิ่มเข้า DataFrame แล้วสั่งจัดเรียงใหม่ทันที
                    updated_df = pd.concat([st.session_state.roster_df, new_df], ignore_index=True)
                    st.session_state.roster_df = sort_roster_by_role(updated_df, st.session_state.employees)
                    
                    st.success(f"เพิ่ม '{new_name}' เรียบร้อย! ระบบจัดเรียงให้อยู่หมวดหมู่ {new_role} แล้ว")
                    st.rerun()

# เครื่องมือช่วยกรอกเวร
with st.expander("⚡ เครื่องมือช่วยกรอกเวรแบบด่วน (เติมรหัสเวรหลายวันรวดเดียว)"):
    with st.form("bulk_fill_form"):
        c1, c2, c3, c4 = st.columns([2, 1, 1, 1])
        emp_options = [f"{r['ชื่อ-สกุล']} ({r['Role (หน้าที่)']})" for _, r in st.session_state.roster_df.iterrows()]
        target_emp = c1.selectbox("1. เลือกพนักงาน", emp_options)
        target_shift = c2.text_input("2. กรอกรหัสเวร (เช่น ว, ค, ย)")
        start_day = c3.number_input("3. ใส่วันที่เริ่มต้น", min_value=1, max_value=31, value=1)
        end_day = c4.number_input("4. ใส่วันที่สิ้นสุด", min_value=1, max_value=31, value=31)
        
        if st.form_submit_button("เติมข้อมูลเวรทันที!"):
            if target_shift.strip() == "": st.warning("กรุณากรอกรหัสเวรก่อนครับ")
            elif start_day > end_day: st.warning("วันที่เริ่มต้น ต้องน้อยกว่าหรือเท่ากับ วันที่สิ้นสุด")
            else:
                raw_name = target_emp.split(" (")[0]
                raw_role = target_emp.split(" (")[1].replace(")", "")
                match_idx = st.session_state.roster_df[
                    (st.session_state.roster_df['ชื่อ-สกุล'] == raw_name) & 
                    (st.session_state.roster_df['Role (หน้าที่)'] == raw_role)
                ].index
                if not match_idx.empty:
                    for d in range(start_day, end_day + 1):
                        st.session_state.roster_df.at[match_idx[0], str(d)] = target_shift.strip()
                    st.success(f"เติมเวรสำเร็จ!")
                    st.rerun()

column_config = {
    "ลำดับ": st.column_config.NumberColumn("ลำดับ", width="small", disabled=True),
    "ชื่อ-สกุล": st.column_config.TextColumn("ชื่อ-สกุล", width="medium"), 
    "ตำแหน่งเบิก": st.column_config.TextColumn("ตำแหน่งเบิก", width="medium"), 
    "Role (หน้าที่)": st.column_config.SelectboxColumn("Role (หน้าที่)", options=roles_list, width="small")
}
for d in range(1, 32): column_config[str(d)] = st.column_config.TextColumn(str(d), width="small")

edited_df = st.data_editor(st.session_state.roster_df, hide_index=True, use_container_width=True, column_config=column_config, key="roster_table")
st.session_state.roster_df = edited_df

# อัปเดตข้อมูลหากมีการแก้ Role ในตาราง
for _, row in edited_df.iterrows():
    name, role = str(row['ชื่อ-สกุล']).strip(), str(row['Role (หน้าที่)']).strip()
    key = f"{name}_{role}"
    if key in st.session_state.employees:
        st.session_state.employees[key]['Role'] = role
        st.session_state.employees[key]['ชื่อ-สกุล'] = name
        st.session_state.employees[key]['ตำแหน่ง'] = row['ตำแหน่งเบิก']

# ==========================================
# 3.5 ระบบตรวจสอบเงื่อนไข (Validation)
# ==========================================
def get_shift_hours(shift_str):
    if not shift_str: return set()
    shift = str(shift_str).strip().replace('(', '').replace(')', '')
    mapping = {
        'ว': set(range(6, 18)), 'ค': set(range(0, 6)) | set(range(18, 24)),
        'ว/ค': set(range(6, 12)) | set(range(18, 24)), 'ค/ว': set(range(0, 6)) | set(range(12, 18)),
        '0-12': set(range(0, 12)), '00-12': set(range(0, 12)),
        '12-24': set(range(12, 24)), '00-24': set(range(0, 24))
    }
    return mapping.get(shift, set())

def validate_roster(roster_df):
    errors = []
    for d in range(1, 32):
        day, role_hours, nai_satanee_off, nai_satanee_shift = str(d), {}, False, ""
        
        for _, row in roster_df.iterrows():
            name, role, shift = str(row['ชื่อ-สกุล']).strip(), str(row['Role (หน้าที่)']).strip(), str(row[day]).strip()
            if not shift: continue
            
            if role == "นสน.":
                nai_satanee_shift = shift
                if shift in ['ย', 'พ', 'ป', 'ก', 'ย.', 'พ.', 'ป.', 'ก.']: nai_satanee_off = True
                    
            h_set = get_shift_hours(shift)
            group_key = "ช.นสน.รวม" if role in ["ช.นสน.1", "ช.นสน.2"] else role
            
            if group_key not in role_hours: role_hours[group_key] = []
            role_hours[group_key].append((name, shift, h_set, role))
            
        for name, shift, h_set, l_role in role_hours.get("นสน.", []):
            if not nai_satanee_off and len(h_set) > 0:
                errors.append(f"🔴 วันที่ {d}: '{name}' เข้าเวร นสน. ไม่ได้ เนื่องจาก นสน.ประจำ ไม่ได้ลา")
            
            allowed = ['ว', 'ค', 'ค/ว', 'ว/ค', '00-12', '0-12', '12-24', '00-24']
            clean_shift = shift.replace('(', '').replace(')', '')
            if len(h_set) > 0 and clean_shift not in allowed:
                errors.append(f"🔴 วันที่ {d}: '{name}' ลงเวร '{shift}' ผิดเงื่อนไขของ นสน.")
        
        for g_key, members in role_hours.items():
            if g_key == "นสน.": continue 
            for i in range(len(members)):
                for j in range(i+1, len(members)):
                    n1, s1, h1, lr1 = members[i]
                    n2, s2, h2, lr2 = members[j]
                    
                    if len(h1) > 0 and len(h2) > 0 and not h1.isdisjoint(h2): 
                        if n1 == n2:
                            errors.append(f"🟠 วันที่ {d}: '{n1}' ลงเวลาซ้อนทับกันเองใน Role '{lr1}' และ '{lr2}'")
                        else:
                            group_name = "ช.นสน." if g_key == "ช.นสน.รวม" else g_key
                            errors.append(f"🟠 วันที่ {d}: เวลาซ้อนทับกันระหว่าง '{n1}' ({s1}) และ '{n2}' ({s2}) ในหน้าที่ {group_name}")
    return errors

st.markdown("---")
st.subheader("✅ 3. ตรวจสอบเงื่อนไขตารางเวร")
if st.button("🔍 กดเพื่อตรวจสอบความถูกต้อง (Validate)", type="secondary"):
    errors = validate_roster(st.session_state.roster_df)
    if len(errors) == 0: st.success("🎉 ตารางเวรถูกต้องตามเงื่อนไขทุกประการ!")
    else:
        st.error(f"พบข้อผิดพลาด {len(errors)} รายการ:")
        for e in errors: st.warning(e)


# ==========================================
# 4. ฟังก์ชันสร้างไฟล์ 109 (แบบปลอดภัย ไม่ทำไฟล์พัง)
# ==========================================
def generate_109(global_vars, roster_df):
    try: wb = openpyxl.load_workbook("109เปล่า.xlsx")
    except: return None
    ws = wb.active
    
    # 1. หยอดตัวแปรส่วนกลาง (วันที่, คำสั่ง)
    replacements_109 = {"[14]": global_vars["val_14"], "[13]": global_vars["val_13"], "[8]": global_vars["val_8"], "[7]": global_vars["val_7"]}
    
    for r in range(1, 100):
        for c in range(1, 40):
            cell = ws.cell(row=r, column=c)
            # เช็คค่าอย่างปลอดภัย ป้องกัน error
            val = cell.value
            if val and isinstance(val, str) and "[" in val:
                new_val = val
                for key, val_rep in replacements_109.items(): 
                    new_val = new_val.replace(key, str(val_rep))
                
                # เขียนค่ากลับอย่างระมัดระวัง (ถ้าเป็นช่อง Merged ระบบอาจจะไม่ยอมให้แก้ตรงๆ ถ้าชี้ไปผิดช่องย่อย)
                # เราใช้ try-except ครอบไว้เผื่อเกิด Error กับช่อง MergedCell
                try:
                    cell.value = new_val
                except AttributeError:
                    pass 
                
    # 2. หยอดรหัสเวร โดยอิงจากการ "ค้นหาชื่อ" ที่มีอยู่ในตาราง 109 เปล่าอยู่แล้ว
    # (ระบบจะไม่พยายามสร้างบรรทัดใหม่ เพื่อป้องกัน MergedCell Error)
    
    # กรุ๊ปข้อมูลคนเดียวกันเข้าด้วยกัน
    grouped_df = roster_df.groupby('ชื่อ-สกุล').agg(lambda x: ' '.join(set([str(i) for i in x if str(i).strip()]))).reset_index()
    
    for _, row_data in grouped_df.iterrows():
        emp_name = row_data['ชื่อ-สกุล'].strip()
        
        # วิ่งหาบรรทัดที่มีชื่อตรงกันในคอลัมน์ที่ 2 (B)
        found_row = None
        for r in range(1, 100):
            try:
                cell_val = str(ws.cell(row=r, column=2).value).strip()
                if cell_val == emp_name:
                    found_row = r
                    break
            except AttributeError:
                pass # ข้ามถ้าเป็น MergedCell ที่อ่านค่าไม่ได้
                
        # ถ้าหาชื่อเจอ หยอดรหัสเวรได้เลย (คอลัมน์เริ่มที่ 3 เป็นต้นไป)
        if found_row:
            for d in range(1, 32):
                shift = row_data[str(d)]
                if pd.notna(shift) and str(shift).strip() != "":
                    try:
                        # สมมติช่องเวรเริ่มที่คอลัมน์ C(3) ในวันที่ 1
                        ws.cell(row=found_row, column=2+d, value=str(shift).strip())
                    except AttributeError:
                        pass
                    
    # ตั้งค่าหน้ากระดาษ
    wsp = ws.sheet_properties
    if not wsp.pageSetUpPr: wsp.pageSetUpPr = PageSetupProperties()
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
def generate_177(unique_key, roster_data, global_vars, ind_vars):
    emp_info = st.session_state.employees[unique_key]
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
        for c in range(1, 15):
            cell = ws.cell(row=r, column=c)
            if cell.value and isinstance(cell.value, str) and "[" in cell.value:
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
    if not wsp.pageSetUpPr: wsp.pageSetUpPr = PageSetupProperties()
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
    
    active_emp_options = [f"{r['ชื่อ-สกุล']} ({r['Role (หน้าที่)']})" for _, r in st.session_state.roster_df.iterrows()]
    selected_key_177_display = st.selectbox("เลือกพนักงานและ Role ที่ต้องการเบิก 177", active_emp_options)
    
    if selected_key_177_display:
        sel_name = selected_key_177_display.split(" (")[0]
        sel_role = selected_key_177_display.split(" (")[1].replace(")", "")
        selected_key_177 = f"{sel_name}_{sel_role}"
        
        st.write(f"กำลังออกใบเบิกให้: **{sel_name}** (ในหน้าที่ **{sel_role}**)")

        if f'ind_data_{selected_key_177}' not in st.session_state:
            st.session_state[f'ind_data_{selected_key_177}'] = {
                "val_4": "-", "val_5": "-", "val_6": "-",
                "val_9": "-", "val_10": "-", "val_11": "-", "val_12": "-"
            }

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.session_state[f'ind_data_{selected_key_177}']['val_4'] = st.text_input("วันหยุด [4]", st.session_state[f'ind_data_{selected_key_177}']['val_4'])
            st.session_state[f'ind_data_{selected_key_177}']['val_9'] = st.text_input("หยุดพักผ่อน [9]", st.session_state[f'ind_data_{selected_key_177}']['val_9'])
        with c2:
            st.session_state[f'ind_data_{selected_key_177}']['val_5'] = st.text_input("วันหยุดที่เบิกได้ [5]", st.session_state[f'ind_data_{selected_key_177}']['val_5'])
            st.session_state[f'ind_data_{selected_key_177}']['val_10'] = st.text_input("พักผ่อนตั้งแต่ [10]", st.session_state[f'ind_data_{selected_key_177}']['val_10'])
        with c3:
            st.session_state[f'ind_data_{selected_key_177}']['val_6'] = st.text_input("เดือนตัวย่อ [6]", st.session_state[f'ind_data_{selected_key_177}']['val_6'])
            st.session_state[f'ind_data_{selected_key_177}']['val_11'] = st.text_input("พักผ่อนถึง [11]", st.session_state[f'ind_data_{selected_key_177}']['val_11'])
        with c4:
            st.session_state[f'ind_data_{selected_key_177}']['val_12'] = st.text_input("รวมพักผ่อน [12]", st.session_state[f'ind_data_{selected_key_177}']['val_12'])

        if st.button(f"ออกใบเบิก 177 ของ {sel_name}", type="primary"):
            emp_row = st.session_state.roster_df[
                (st.session_state.roster_df['ชื่อ-สกุล'] == sel_name) & 
                (st.session_state.roster_df['Role (หน้าที่)'] == sel_role)
            ].iloc[0]
            
            roster_dict = {str(d): str(emp_row[str(d)]) if pd.notna(emp_row[str(d)]) else "" for d in range(1, 32)}
            ind_data = st.session_state[f'ind_data_{selected_key_177}']
            excel_177 = generate_177(selected_key_177, roster_dict, global_data, ind_data)
            
            if excel_177:
                st.success(f"สร้างใบเบิก 177 ของ {sel_name} เสร็จสิ้น!")
                st.download_button("📥 ดาวน์โหลดไฟล์ 177", data=excel_177, file_name=f"ใบเบิก_{sel_name}_{sel_role}.xlsx")
