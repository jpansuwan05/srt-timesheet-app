import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.styles import Border, Side, Alignment, Font
import io
import datetime
import calendar
import re
import json
import uuid
import zipfile  
from copy import copy
from streamlit_local_storage import LocalStorage
import streamlit.components.v1 as components

st.set_page_config(page_title="SRT Timesheet App", page_icon="🚂", layout="wide")

# ==========================================
# ✨ เวทมนตร์ CSS แต่งหน้าตา Web App ให้ดู Modern
# ==========================================
st.markdown("""
    <style>
        /* 1. แต่งหน้าตาของการ์ด Dashboard (Metric) ให้มีเงาและขอบมน */
        div[data-testid="metric-container"] {
            background-color: #f8f9fa;
            border-radius: 15px;
            padding: 15px 20px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            border-left: 6px solid #FF4B4B;
            transition: transform 0.2s ease;
        }
        div[data-testid="metric-container"]:hover {
            transform: scale(1.02);
        }
        
        /* 2. แต่งปุ่มกดให้โค้งมนและเด้งได้ตอนเอาเมาส์ชี้ */
        div.stButton > button {
            border-radius: 10px;
            font-weight: 600;
            transition: all 0.3s ease;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
        }
        div.stButton > button:hover {
            transform: translateY(-3px);
            box-shadow: 0 6px 12px rgba(0, 0, 0, 0.15);
        }
        
        /* 3. ปรับขอบโค้งมนให้กับ Container และ Expander */
        div[data-testid="stVerticalBlock"] > div[style*="border"] {
            border-radius: 15px !important;
            box-shadow: 0 4px 15px rgba(0,0,0,0.05) !important;
            border: 1px solid #f0f2f6 !important;
            padding: 20px !important;
            background-color: #ffffff;
        }
        div[data-testid="stExpander"] {
            border-radius: 10px !important;
            border: 1px solid #e9ecef !important;
            box-shadow: 0 2px 5px rgba(0,0,0,0.02) !important;
        }

        /* 4. รองรับโหมดมืด (Dark Mode) */
        @media (prefers-color-scheme: dark) {
            div[data-testid="metric-container"] {
                background-color: #262730;
                box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
            }
            div[data-testid="stVerticalBlock"] > div[style*="border"] {
                background-color: #0e1117;
                border: 1px solid #333 !important;
            }
            div[data-testid="stExpander"] {
                border: 1px solid #333 !important;
            }
        }
    </style>
""", unsafe_allow_html=True)

st.title("🚂 ระบบจัดการเวรและใบเบิกค่าตอบแทน (รฟท.)")

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
# 1. เมนูแถบด้านข้าง (Sidebar)
# ==========================================
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/th/thumb/7/77/State_Railway_of_Thailand_logo.png/200px-State_Railway_of_Thailand_logo.png", width=120)
    st.markdown("### 🔄 เริ่มต้นเดือนใหม่")
    if st.button("🗑️ ล้างข้อมูล (อัปโหลดรายชื่อใหม่)", type="primary", use_container_width=True):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        components.html("<script>localStorage.clear(); window.parent.location.reload();</script>", height=0)
        st.stop()

    st.markdown("---")
    st.markdown("### 💾 สำรองข้อมูลตารางเวร")
    if 'roster_df' in st.session_state and st.session_state.roster_df is not None:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            st.session_state.roster_df.to_excel(writer, index=False, sheet_name='Roster')
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์สำรอง (.xlsx)", 
            data=buffer, 
            file_name=f"backup_roster_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx", 
            mime="application/vnd.ms-excel",
            use_container_width=True
        )

    st.markdown("<br>", unsafe_allow_html=True)
    uploaded_backup = st.file_uploader("📤 โหลดไฟล์สำรองเดิมมาทำต่อ", type=["xlsx"])
    if uploaded_backup:
        if st.button("ยืนยันโหลดตารางเวรเก่า", use_container_width=True):
            try:
                loaded_df = pd.read_excel(uploaded_backup)
                for d in range(1, 32):
                    if str(d) in loaded_df.columns:
                        loaded_df[str(d)] = loaded_df[str(d)].astype(str).replace('nan', '')
                st.session_state.roster_df = loaded_df
                save_roster_to_local(loaded_df)
                st.success("โหลดข้อมูลสำเร็จ! 🎉")
                st.rerun()
            except Exception as e:
                st.error(f"Error: {e}")

# ==========================================
# 2. 🛡️ ระบบโหลดข้อมูลพนักงานแบบปลอดภัย
# ==========================================
saved_emp_json = local_storage.getItem("srt_employees_data")

if saved_emp_json and 'employees' not in st.session_state:
    try: st.session_state.employees = json.loads(saved_emp_json)
    except: st.session_state.employees = None

if 'employees' not in st.session_state or not st.session_state.employees:
    if saved_emp_json:
        st.info("💡 **ตรวจพบข้อมูลตารางเวรที่คุณทำค้างไว้ในเครื่อง!**")
        if st.button("🔄 กู้คืนข้อมูลล่าสุดกลับมาทำงานต่อ", type="primary"):
            st.session_state.employees = json.loads(saved_emp_json)
            st.rerun()
            
    with st.container(border=True):
        st.warning("🔒 **ยังไม่มีข้อมูลพนักงานในระบบ กรุณาอัปโหลดไฟล์เพื่อเริ่มต้น**")
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
                    if "นายสถานี" in p_clean or ("นสน" in p_clean and "ช.นสน" not in p_clean): role = "นสน."
                    elif "ช.นสน.ตช" in p_clean or "ช.นสน.1" in p_clean: role = "ช.นสน.1"
                    elif "ช.นสน.ตค" in p_clean or "ช.นสน.2" in p_clean: role = "ช.นสน.2"
                    elif "เสมียน" in p_clean: role = "เสมียน"
                    elif "ประแจ" in p_clean: role = "ประแจ"
                    elif "ฉิมพลี" in p_clean: role = "กั้นถนนฯฉิมพลี"
                    elif "บางระมาด" in p_clean: role = "กั้นถนนฯบางระมาด"
                    elif "บริการ" in p_clean or "ลูกจ้าง" in p_clean: role = "ลูกจ้าง"
                    elif "กั้นถนน" in p_clean: role = "อื่นๆ"
                    else: role = "อื่นๆ"
                    
                    emp_id_raw = str(row.get("เลขประจำตัว", "-")) if pd.notna(row.get("เลขประจำตัว")) else "-"
                    if emp_id_raw.endswith(".0"): emp_id_raw = emp_id_raw[:-2]
                    
                    acc1_raw = str(row.get("รหัสบัญชี", "-")) if pd.notna(row.get("รหัสบัญชี")) else "-"
                    if acc1_raw.endswith(".0"): acc1_raw = acc1_raw[:-2]
                    
                    acc2_raw = str(row.get("รหัสบัญชี2", "-")) if pd.notna(row.get("รหัสบัญชี2")) else "-"
                    if acc2_raw.endswith(".0"): acc2_raw = acc2_raw[:-2]
                    
                    unique_key = f"{name}_{role}"
                    emp_dict[unique_key] = {
                        "ชื่อ-สกุล": name, "ตำแหน่ง": pos, 
                        "เลขประจำตัว": emp_id_raw,
                        "เงินเดือน": str(row.get("เงินเดือน", "-")) if pd.notna(row.get("เงินเดือน")) else "-", 
                        "เรท": float(row.get("เรท 1 ชั่วโมง", 0)) if pd.notna(row.get("เรท 1 ชั่วโมง")) else 0.0,
                        "ประเภทบัญชี": str(row.get("ประเภทบัญชี", "-")) if pd.notna(row.get("ประเภทบัญชี")) else "-", 
                        "รหัสบัญชี": acc1_raw,
                        "รหัสบัญชี2": acc2_raw,
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
leave_types = ["ย", "ย.", "พ", "พ.", "ป", "ป.", "ก", "ก.", "น", "น.", "ล", "ล.", "ลา"]
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
                        "เงินเดือน": "-", "เรท": 0.0, "ประเภทบัญชี": "-", "รหัสบัญชี": "-", "รหัสบัญชี2": "-", "Role": role, "is_regular": False
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

saved_global = local_storage.getItem("srt_global_data")
default_global = {"val_13": "สิงหาคม", "year_be": 2569, "val_7": "5110/2520/2569", "val_8": "29 พ.ค. 69", "val_14": "01 ก.ค. 69", "public_holidays_dict": {}}
try:
    if saved_global: 
        parsed = json.loads(saved_global)
        if "public_holidays" in parsed and isinstance(parsed["public_holidays"], list):
            parsed["public_holidays_dict"] = {str(k): "" for k in parsed["public_holidays"]}
        default_global.update(parsed)
except: pass

months_list = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
default_m = default_global.get("val_13", "สิงหาคม")
if default_m not in months_list: default_m = "สิงหาคม"
default_year = int(default_global.get("year_be", 2569))

month_idx = months_list.index(default_m) + 1
year_ce = default_year - 543
first_weekday, num_days = calendar.monthrange(year_ce, month_idx)

# ==========================================
# 4. ตั้งค่าส่วนกลาง & ตารางเวร
# ==========================================
with st.container(border=True):
    st.subheader("⚙️ 1. ตั้งค่าข้อมูลส่วนกลางประจำเดือน")
    col_g1, col_g2 = st.columns(2)
    with col_g1:
        val_13 = st.selectbox("เดือน [13]", months_list, index=months_list.index(default_m))
        val_7 = st.text_input("คำสั่งแขวง [7]", default_global.get("val_7", ""))
    with col_g2:
        year_be = st.number_input("ปี พ.ศ. (สำหรับการจัดปฏิทิน)", min_value=2500, max_value=2600, value=default_year)
        val_8 = st.text_input("วันที่ลงคำสั่ง [8]", default_global.get("val_8", ""))
        val_14 = st.text_input("วันที่เซ็นเอกสารตัวย่อ [14]", default_global.get("val_14", ""))

    month_idx = months_list.index(val_13) + 1
    year_ce = year_be - 543
    first_weekday, num_days = calendar.monthrange(year_ce, month_idx)

    st.markdown("---")
    st.markdown("##### 🎉 ตั้งค่าวันหยุดนักขัตฤกษ์ประจำเดือน")
    
    old_ph_dict = default_global.get("public_holidays_dict", {})
    public_holidays_list = st.multiselect(
        "1) เลือกวันที่ตรงกับ 'วันหยุดนักขัตฤกษ์' (ใบ 178 จะถือเป็นวันหยุดพิเศษ)", 
        list(range(1, num_days + 1)),
        default=[int(k) for k in old_ph_dict.keys() if int(k) <= num_days]
    )

    ph_dict = {}
    if public_holidays_list:
        st.markdown("2) โปรดระบุชื่อวันหยุดนักขัตฤกษ์ (เพื่อแสดงในใบ 109 และ 178)")
        cols = st.columns(min(len(public_holidays_list), 4))
        for i, d in enumerate(sorted(public_holidays_list)):
            with cols[i % 4]:
                old_name = old_ph_dict.get(str(d), "")
                ph_name = st.text_input(f"ชื่อวันหยุด (วันที่ {d})", value=old_name, key=f"ph_{d}")
                ph_dict[str(d)] = ph_name

    global_data = {
        "val_13": val_13, "year_be": year_be, "val_7": val_7, 
        "val_8": val_8, "val_14": val_14, "public_holidays_dict": ph_dict
    }
    local_storage.setItem("srt_global_data", json.dumps(global_data), key=f"ls_global_{uuid.uuid4().hex}")

# ==========================================
# 📈 Dashboard สรุปข้อมูล
# ==========================================
st.markdown("### 📊 ภาพรวมข้อมูลสถานี")
m1, m2, m3, m4 = st.columns(4)
m1.metric("👥 จำนวนพนักงาน", f"{len(st.session_state.roster_df)} คน")
m2.metric("🗓️ ประจำเดือน", f"{val_13} {year_be}")
m3.metric("📅 จำนวนวันในเดือนนี้", f"{num_days} วัน")
m4.metric("🎉 วันหยุดนักขัตฤกษ์", f"{len(ph_dict)} วัน")
st.markdown("---")

with st.container(border=True):
    st.subheader(f"🗓️ 2. จัดการตารางเวร 1-{num_days} วัน")

    with st.expander("➕ เพิ่มพนักงานใหม่ / เข้าเวรแทน (คลิกเพื่อเปิด)"):
        with st.form("add_emp_form"):
            c1, c2, c3, c4 = st.columns(4)
            new_name = c1.text_input("ชื่อ-สกุล*")
            new_pos = c2.text_input("ตำแหน่งเบิก*")
            new_role = c3.selectbox("Role (หน้าที่)", roles_list)
            new_rate = c4.number_input("เรท 1 ชม. (บาท)", min_value=0.0, value=0.0)
            
            if st.form_submit_button("เพิ่ม / อัปเดตพนักงาน", type="primary"):
                if new_name.strip() == "" or new_pos.strip() == "": st.error("กรุณากรอก ชื่อ และ ตำแหน่ง!")
                else:
                    unique_key = f"{new_name}_{new_role}"
                    old_data = st.session_state.employees.get(unique_key, {})
                    old_salary = old_data.get("เงินเดือน", "-")
                    old_rate = float(old_data.get("เรท", 0.0))
                    
                    st.session_state.employees[unique_key] = {
                        "ชื่อ-สกุล": new_name, "ตำแหน่ง": new_pos, 
                        "เลขประจำตัว": old_data.get("เลขประจำตัว", "-"), 
                        "เงินเดือน": old_salary, 
                        "เรท": new_rate if new_rate > 0 else old_rate, 
                        "ประเภทบัญชี": old_data.get("ประเภทบัญชี", "-"), 
                        "รหัสบัญชี": old_data.get("รหัสบัญชี", "-"), 
                        "รหัสบัญชี2": old_data.get("รหัสบัญชี2", "-"), 
                        "Role": new_role, "is_regular": old_data.get("is_regular", False)
                    }
                    
                    exists = any((r['ชื่อ-สกุล'] == new_name and r['Role (หน้าที่)'] == new_role) for _, r in st.session_state.roster_df.iterrows())
                    if not exists:
                        new_idx = len(st.session_state.roster_df) + 1
                        new_row = {"ขึ้นหน้าใหม่": False, "ลำดับ": new_idx, "ชื่อ-สกุล": new_name, "ตำแหน่งเบิก": new_pos, "Role (หน้าที่)": new_role}
                        for d in range(1, 32): new_row[str(d)] = ""
                        new_df = pd.DataFrame([new_row])
                        updated_df = pd.concat([st.session_state.roster_df, new_df], ignore_index=True)
                        st.session_state.roster_df = sort_roster_by_role(updated_df, st.session_state.employees)
                    save_roster_to_local(st.session_state.roster_df)
                    local_storage.setItem("srt_employees_data", json.dumps(st.session_state.employees), key=f"ls_emp_{uuid.uuid4().hex}")
                    st.success("✅ อัปเดตข้อมูลพนักงานเรียบร้อย!")
                    st.rerun()

    column_config = {
        "ขึ้นหน้าใหม่": st.column_config.CheckboxColumn("ขึ้นหน้าใหม่", width="small"),
        "ลำดับ": st.column_config.NumberColumn("ลำดับ", width="small", disabled=True),
        "ชื่อ-สกุล": st.column_config.TextColumn("ชื่อ-สกุล", width="medium"), 
        "ตำแหน่งเบิก": st.column_config.TextColumn("ตำแหน่งเบิก", width="medium"), 
        "Role (หน้าที่)": st.column_config.SelectboxColumn("Role", options=roles_list, width="small")
    }
    for d in range(1, 32): 
        if d <= num_days:
            column_config[str(d)] = st.column_config.TextColumn(str(d), width="small")
        else:
            column_config[str(d)] = st.column_config.TextColumn(f"{d} (ไม่มี)", width="small", disabled=True)

    edited_df = st.data_editor(st.session_state.roster_df, hide_index=True, use_container_width=True, column_config=column_config, key="roster_table")

    for d in range(1, 32):
        d_str = str(d)
        if d_str in edited_df.columns:
            edited_df[d_str] = edited_df[d_str].fillna("").astype(str).replace(["None", "nan", "<NA>"], "")

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
# 🤖 🛡️ ระบบตรวจสอบความถูกต้องของตารางเวร (AI Validator)
# ==========================================
st.markdown("---")
st.markdown("#### 🤖 ระบบ AI ตรวจสอบข้อผิดพลาดของตารางเวร (ตามเงื่อนไข รฟท.)")
validator_errors = []

def get_shift_clean_for_val(row_data, day_num):
    if str(day_num) not in row_data: return ""
    val = row_data[str(day_num)]
    if pd.isna(val) or str(val).strip() in ["None", "nan", "<NA>", ""]: return ""
    s = str(val).strip().replace("(", "").replace(")", "")
    return s

for d in range(1, num_days + 1):
    day_info = {r: {'reg':[], 'sub':[]} for r in roles_list}
    
    for _, r_data in edited_df.iterrows():
        r_name = str(r_data.get('ชื่อ-สกุล', '')).strip()
        r_role = str(r_data.get('Role (หน้าที่)', '')).strip()
        shift_c = get_shift_clean_for_val(r_data, d)
        if not r_name or not shift_c: continue
        
        k = f"{r_name}_{r_role}"
        is_reg = st.session_state.employees.get(k, {}).get('is_regular', False)
        if r_role in day_info:
            if is_reg: day_info[r_role]['reg'].append({'name': r_name, 'shift': shift_c})
            else: day_info[r_role]['sub'].append({'name': r_name, 'shift': shift_c})

    nsn_allowed = ["ว", "ค", "ค/ว", "ว/ค", "00-12", "12-24", "00-24", "0-12"]
    for p in day_info['นสน.']['reg'] + day_info['นสน.']['sub']:
        if p['shift'] not in leave_types and p['shift'] not in nsn_allowed:
            validator_errors.append(f"วันที่ {d}: {p['name']} (นสน.) ลงรหัส '{p['shift']}' ไม่ถูกต้อง (ต้องเป็น ว, ค, ค/ว, ว/ค, 00-12, 12-24, 00-24)")
            
    reg_nsn_shift = day_info['นสน.']['reg'][0]['shift'] if day_info['นสน.']['reg'] else ""
    for p in day_info['นสน.']['sub']:
        if p['shift'] not in leave_types:
            if reg_nsn_shift not in ["ย", "ย.", "พ", "พ.", "ป", "ป.", "ก", "ก.", "ล", "ลา", "น"]:
                validator_errors.append(f"วันที่ {d}: {p['name']} (มาแทน นสน.) เข้าเวรไม่ได้ เพราะนายสถานีตัวจริงไม่ได้ลา (ตัวจริงลง '{reg_nsn_shift}')")
                
    def get_active(role_str):
        return [x for x in (day_info[role_str]['reg'] + day_info[role_str]['sub']) if x['shift'] not in leave_types]
        
    active_ch1 = get_active('ช.นสน.1')
    active_ch2 = get_active('ช.นสน.2')
    
    for p1 in active_ch1:
        for p2 in active_ch2:
            s1, s2 = p1['shift'], p2['shift']
            if s1 == s2:
                validator_errors.append(f"วันที่ {d}: {p1['name']} ({s1}) และ {p2['name']} ({s2}) เข้าเวรซ้ำกัน")
            elif s1 == "ว" and s2 != "ค":
                validator_errors.append(f"วันที่ {d}: {p1['name']} ลง 'ว' แต่ {p2['name']} ลง '{s2}' (ต้องคู่กับ 'ค')")
            elif s1 == "ค" and s2 != "ว":
                validator_errors.append(f"วันที่ {d}: {p1['name']} ลง 'ค' แต่ {p2['name']} ลง '{s2}' (ต้องคู่กับ 'ว')")
            elif s1 == "ว/ค" and s2 != "ค/ว":
                validator_errors.append(f"วันที่ {d}: {p1['name']} ลง 'ว/ค' แต่ {p2['name']} ลง '{s2}' (ต้องคู่กับ 'ค/ว')")
            elif s1 == "ค/ว" and s2 != "ว/ค":
                validator_errors.append(f"วันที่ {d}: {p1['name']} ลง 'ค/ว' แต่ {p2['name']} ลง '{s2}' (ต้องคู่กับ 'ว/ค')")

    for role in roles_list:
        if role == 'นสน.': continue
        for r in day_info[role]['reg']:
            for s in day_info[role]['sub']:
                if r['shift'] not in leave_types and s['shift'] not in leave_types:
                    if r['shift'] == s['shift']:
                        validator_errors.append(f"วันที่ {d}: {s['name']} ลงเวร '{s['shift']}' ซ้ำกับตัวจริง {r['name']} ในตำแหน่ง {role}")
                        
if validator_errors:
    st.error(f"พบข้อผิดพลาดในตารางเวรทั้งหมด {len(validator_errors)} จุด (โปรดตรวจสอบและแก้ไข)")
    for err in validator_errors:
        st.warning(err, icon="⚠️")
else:
    st.success("✅ ตารางเวรถูกต้องตามเงื่อนไขของ รฟท. (ไม่พบการลงเวลาทับซ้อนหรือผิดคู่)")

# ==========================================
# 5. ฟังก์ชันสร้างไฟล์ Excel
# ==========================================

def extract_employee_stats(roster_row):
    weekly_days = []
    leave_days_vacation = []
    
    is_in_period = False
    for d in range(1, 32):
        val = str(roster_row.get(str(d), "")).strip()
        if "(" in val: is_in_period = True
        clean_val = val.replace("(", "").replace(")", "")
        
        if clean_val in ['ย', 'ย.']:
            weekly_days.append(d)
        elif is_in_period and clean_val and clean_val not in ["พ", "พ.", "ป", "ป.", "ก", "ก.", "น", "น.", "ล", "ล.", "ลา", "-"]:
            weekly_days.append(d)
            
        if clean_val in ['พ', 'พ.']:
            leave_days_vacation.append(d)
            
        if ")" in val: is_in_period = False

    def to_ranges(days_list):
        if not days_list: return []
        ranges = []
        start = days_list[0]
        prev = days_list[0]
        for d in days_list[1:]:
            if d == prev + 1: prev = d
            else:
                ranges.append(f"{start}-{prev}" if start != prev else f"{start}")
                start = d
                prev = d
        ranges.append(f"{start}-{prev}" if start != prev else f"{start}")
        return ranges
        
    w_ranges = to_ranges(weekly_days)
    val_4 = w_ranges[0] if len(w_ranges) > 0 else "-"
    val_5 = ",".join(w_ranges[1:]) if len(w_ranges) > 1 else "-"
    val_17 = f"{len(weekly_days):02d}"
    
    v_ranges = to_ranges(leave_days_vacation)
    val_9 = ",".join(v_ranges) if v_ranges else "-"
    val_10 = str(leave_days_vacation[0]) if leave_days_vacation else "-"
    val_11 = str(leave_days_vacation[-1]) if leave_days_vacation else "-"
    val_12 = f"{len(leave_days_vacation):02d}"
    
    return val_4, val_5, val_17, val_9, val_10, val_11, val_12

def generate_109(global_vars, roster_df, num_days, first_weekday):
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
    days_th_abbr = ["จ.", "อ.", "พ.", "พฤ.", "ศ.", "ส.", "อา."]
    
    ph_dict_local = global_vars.get("public_holidays_dict", {})
    ph_append_str = ""
    if ph_dict_local:
        ph_texts = []
        for d_str, name in sorted(ph_dict_local.items(), key=lambda x: int(x[0])):
            ph_texts.append(f"วันที่ {d_str} {name}" if name else f"วันที่ {d_str}")
        ph_append_str = " (" + ", ".join(ph_texts) + ")"

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
                    
        for r in range(35, 55):
            for c in range(1, 40):
                c_cell = ws.cell(row=r, column=c)
                val = c_cell.value
                if val and isinstance(val, str):
                    new_val = val
                    for k, v in replacements_109.items(): 
                        if k in new_val:
                            new_val = new_val.replace(k, str(v))
                    
                    if "วันหยุดนักขัตฤกษ์" in new_val and ph_append_str:
                        if ph_append_str not in new_val:
                            new_val = new_val.replace("วันหยุดนักขัตฤกษ์", f"วันหยุดนักขัตฤกษ์{ph_append_str}")
                            
                    if type(c_cell).__name__ != 'MergedCell': 
                        c_cell.value = new_val
        
        date_row = None
        for r in range(4, 10):
            if str(ws.cell(row=r, column=3).value).strip() == "1" and str(ws.cell(row=r, column=4).value).strip() == "2":
                date_row = r; break
        if not date_row: date_row = 7
        day_row = date_row - 1
        
        for r in range(8, 37, 2):
            ws.cell(row=r, column=1).value = ""
            ws.cell(row=r, column=2).value = ""
            ws.cell(row=r+1, column=2).value = ""
            for d in range(1, 32): ws.cell(row=r, column=2+d).value = ""
            
        for d in range(1, 32):
            col = 2 + d
            if d <= num_days:
                wd = (first_weekday + d - 1) % 7
                ws.cell(row=day_row, column=col).value = days_th_abbr[wd]
            else:
                ws.cell(row=day_row, column=col).value = ""
                ws.cell(row=date_row, column=col).value = ""
        
        current_excel_row = 8
        for row_data in page_data:
            ws.cell(row=current_excel_row, column=1).value = row_data['ลำดับ']
            ws.cell(row=current_excel_row, column=2).value = str(row_data['ชื่อ-สกุล']).strip()
            ws.cell(row=current_excel_row+1, column=2).value = str(row_data['ตำแหน่งเบิก']).strip()
            role_val = str(row_data.get('Role (หน้าที่)', '')).strip()
            
            for d in range(1, 32):
                col = 2 + d
                c_cell = ws.cell(row=current_excel_row, column=col)
                if d <= num_days:
                    shift = row_data.get(str(d), "")
                    c_cell.value = str(shift).strip() if pd.notna(shift) else ""
                    if c_cell.font: new_font = copy(c_cell.font)
                    else: new_font = Font()
                    
                    if shift:
                        s_clean = str(shift).strip().replace("(", "").replace(")", "")
                        if s_clean in leave_types: new_font.color = "FF0000" 
                        elif s_clean in ["00-12", "0-12", "12-24"]:
                            new_font.color = "008000" 
                            if new_font.size: new_font.size = new_font.size - 2 
                            else: new_font.size = 12
                        elif role_val == "กั้นถนนฯฉิมพลี" and s_clean in ["ว", "ค", "ว/ค", "ค/ว"]:
                            new_font.color = "0000FF" 
                        else: new_font.color = "000000" 
                    else: new_font.color = "000000"
                    c_cell.font = new_font
                else:
                    c_cell.value = ""
                    if c_cell.font:
                        nf = copy(c_cell.font)
                        nf.color = "000000"
                        c_cell.font = nf
                
            current_excel_row += 2
            
        if not ws.sheet_properties.pageSetUpPr: ws.sheet_properties.pageSetUpPr = PageSetupProperties()
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0 
        ws.page_setup.paperSize = ws.PAPERSIZE_A4; ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output, total_pages

def generate_177(unique_key, roster_data, global_vars, ind_vars, num_days):
    emp_info = st.session_state.employees.get(unique_key)
    if not emp_info: return None
    
    emp_id_str = str(emp_info.get("เลขประจำตัว", "-"))
    if emp_id_str.endswith(".0"): emp_id_str = emp_id_str[:-2]
    
    raw_salary = str(emp_info.get('เงินเดือน', '-'))
    if raw_salary != "-" and raw_salary.replace('.', '', 1).isdigit():
        salary_str = f"{float(raw_salary):,.0f}"
    else:
        salary_str = raw_salary

    try: wb = openpyxl.load_workbook("ใบ177 Update.xlsx")
    except: return None
    ws = wb.active
    replacements = {
        "[NAME]": emp_info["ชื่อ-สกุล"], "[16]": emp_info.get("รหัสบัญชี", "-"), "[15]": emp_info["ประเภทบัญชี"],
        "[14]": global_vars["val_14"], "[13]": global_vars["val_13"], 
        "[17]": ind_vars.get("val_17", ""), "[12]": ind_vars.get("val_12", ""),
        "[11]": ind_vars.get("val_11", ""), "[10]": ind_vars.get("val_10", ""), "[9]": ind_vars.get("val_9", ""),
        "[8]": global_vars["val_8"], "[7]": global_vars["val_7"], "[6]": ind_vars.get("val_6", ""),
        "[5]": ind_vars.get("val_5", ""), "[4]": ind_vars.get("val_4", ""),
        "[3]": salary_str,  
        "[2]": emp_id_str, "[1]": emp_info["ตำแหน่ง"]
    }
    for r in range(1, 55):
        for c in range(1, 40): 
            c_cell = ws.cell(row=r, column=c)
            val = c_cell.value
            if val and isinstance(val, str) and "[" in val:
                new_val = val
                for k, v in replacements.items(): new_val = new_val.replace(k, str(v))
                if type(c_cell).__name__ != 'MergedCell': c_cell.value = new_val
                
    start_row = 7
    rate_val = float(emp_info["เรท"]) if emp_info["เรท"] else 0.0
    rate_baht = int(rate_val)
    rate_satang = int(round((rate_val - rate_baht) * 100))
    
    def set_cell_val_color(r, c, val, color_hex="000000"):
        cell = ws.cell(row=r, column=c)
        if type(cell).__name__ != 'MergedCell':
            cell.value = val
            if cell.font:
                nf = copy(cell.font)
                nf.color = color_hex
                cell.font = nf
            else:
                cell.font = Font(color=color_hex)
    
    for day in range(1, 32):
        row = start_row + day - 1
        if day > num_days:
            set_cell_val_color(row, 2, "", "000000")
            for col in range(3, 8): set_cell_val_color(row, col, "", "000000")
            continue
            
        shift_raw = str(roster_data.get(str(day), "")).strip()
        shift_clean = shift_raw.replace("(", "").replace(")", "")
        sData = shift_data.get(shift_clean)
        
        is_holiday = shift_clean in leave_types
        font_color = "FF0000" if is_holiday else "000000"
        
        if sData and sData["hours"] != "-":
            hours_val = int(sData["hours"])
            total_money = hours_val * rate_val
            total_baht = int(total_money)
            total_satang = int(round((total_money - total_baht) * 100))

            set_cell_val_color(row, 2, sData["text"], font_color)
            set_cell_val_color(row, 3, hours_val, font_color)
            set_cell_val_color(row, 4, rate_baht if rate_baht > 0 else 0, font_color)
            set_cell_val_color(row, 5, f"{rate_satang:02d}", font_color)
            set_cell_val_color(row, 6, total_baht if total_baht > 0 else 0, font_color)
            set_cell_val_color(row, 7, f"{total_satang:02d}", font_color)
        else:
            val = sData["text"] if sData else (shift_raw if shift_raw else "-")
            set_cell_val_color(row, 2, val, font_color)
            for col in range(3, 8): 
                set_cell_val_color(row, col, "-", font_color)
                    
    for r in range(37, 45):
        cell_v1 = str(ws.cell(row=r, column=1).value).strip()
        cell_v2 = str(ws.cell(row=r, column=2).value).strip()
        if "รวม" in cell_v1 or "รวม" in cell_v2:
            set_cell_val_color(r, 4, rate_baht if rate_baht > 0 else 0, "000000")
            set_cell_val_color(r, 5, f"{rate_satang:02d}", "000000")
            break
            
    if not ws.sheet_properties.pageSetUpPr: ws.sheet_properties.pageSetUpPr = PageSetupProperties()
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToHeight = 1; ws.page_setup.fitToWidth = 1
    ws.page_setup.paperSize = ws.PAPERSIZE_A4; ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def generate_178(unique_key, roster_data, global_vars, ind_vars, num_days):
    emp_info = st.session_state.employees.get(unique_key)
    if not emp_info: return None
    
    emp_id_str = str(emp_info.get("เลขประจำตัว", "-"))
    if emp_id_str.endswith(".0"): emp_id_str = emp_id_str[:-2]
    
    raw_salary = str(emp_info.get('เงินเดือน', '-'))
    if raw_salary != "-" and raw_salary.replace('.', '', 1).isdigit():
        salary_str = f"{float(raw_salary):,.0f}"
    else:
        salary_str = raw_salary

    try: wb = openpyxl.load_workbook("178 อัพเดท.xlsx")
    except: return None
    ws = wb.active
    
    ph_dict_local = global_vars.get("public_holidays_dict", {})
    
    replacements = {
        "[NAME]": emp_info["ชื่อ-สกุล"], "[16]": emp_info.get("รหัสบัญชี2", "-"), "[15]": emp_info["ประเภทบัญชี"],
        "[14]": global_vars["val_14"], "[13]": global_vars["val_13"],
        "[8]": global_vars["val_8"], "[7]": global_vars["val_7"],
        "[17]": ind_vars.get("val_17", ""), "[12]": ind_vars.get("val_12", ""), 
        "[11]": ind_vars.get("val_11", ""), "[10]": ind_vars.get("val_10", ""), "[9]": ind_vars.get("val_9", ""),
        "[6]": ind_vars.get("val_6", ""), "[5]": ind_vars.get("val_5", ""), "[4]": ind_vars.get("val_4", ""),
        "[3]": salary_str,  
        "[2]": emp_id_str, "[1]": emp_info["ตำแหน่ง"]
    }
    
    for r in range(1, 55):
        for c in range(1, 40): 
            c_cell = ws.cell(row=r, column=c)
            val = c_cell.value
            if val and isinstance(val, str):
                new_val = val
                for k, v in replacements.items(): 
                    if k in new_val:
                        new_val = new_val.replace(k, str(v))
                if type(c_cell).__name__ != 'MergedCell': c_cell.value = new_val

    rate_val = float(emp_info["เรท"]) if emp_info["เรท"] else 0.0
    daily_rate = rate_val * 8
    
    if type(ws.cell(row=3, column=12)).__name__ != 'MergedCell': ws.cell(row=3, column=12).value = rate_val
    if type(ws.cell(row=5, column=12)).__name__ != 'MergedCell': ws.cell(row=5, column=12).value = daily_rate
        
    start_row = 7
    weekly_holiday_count = 0
    public_holiday_count = 0
    is_in_weekly_period = False
    
    for day in range(1, 32):
        row = start_row + day
        ws.cell(row=row, column=1).value = str(day) 
        
        for col in range(2, 11):
            if type(ws.cell(row=row, column=col)).__name__ != 'MergedCell':
                ws.cell(row=row, column=col).value = None

        if day > num_days:
            ws.cell(row=row, column=1).value = ""
            continue
            
        shift_raw = str(roster_data.get(str(day), "")).strip()
        if "(" in shift_raw: is_in_weekly_period = True
        shift_clean = shift_raw.replace("(", "").replace(")", "")
        
        if shift_clean and shift_clean not in leave_types and shift_clean != "-":
            is_public = str(day) in ph_dict_local
            is_weekly = False
            
            if is_in_weekly_period and not is_public:
                is_weekly = True
                
            if is_public or is_weekly:
                t1_start, t1_end, t2_start, t2_end = None, None, None, None
                if shift_clean == "ว": t1_start, t1_end = "06.00", "18.00"
                elif shift_clean == "ค": t1_start, t1_end, t2_start, t2_end = "00.00", "06.00", "18.00", "24.00"
                elif shift_clean == "ว/ค": t1_start, t1_end, t2_start, t2_end = "06.00", "12.00", "18.00", "24.00"
                elif shift_clean == "ค/ว": t1_start, t1_end, t2_start, t2_end = "00.00", "06.00", "12.00", "18.00"
                elif shift_clean in ["0-12", "00-12"]: t1_start, t1_end = "00.00", "12.00"
                elif shift_clean == "12-24": t1_start, t1_end = "12.00", "24.00"
                elif shift_clean == "00-24": t1_start, t1_end = "00.00", "24.00"
                else: t1_start, t1_end = shift_clean, ""
                
                ws.cell(row=row, column=3).value = emp_info["ตำแหน่ง"]
                if t1_start: ws.cell(row=row, column=4).value = t1_start
                if t1_end: ws.cell(row=row, column=5).value = t1_end
                if t2_start: ws.cell(row=row, column=6).value = t2_start
                if t2_end: ws.cell(row=row, column=7).value = t2_end
                ws.cell(row=row, column=8).value = 1 
                ws.cell(row=row, column=10).value = daily_rate 
                
                if is_public:
                    h_name = ph_dict_local.get(str(day), "วันหยุดนักขัตฤกษ์")
                    if not h_name.strip(): h_name = "วันหยุดนักขัตฤกษ์"
                    ws.cell(row=row, column=2).value = f"({h_name})"
                    public_holiday_count += 1
                elif is_weekly:
                    ws.cell(row=row, column=2).value = "(วันหยุดประจำสัปดาห์)"
                    weekly_holiday_count += 1
                    
        if ")" in shift_raw: is_in_weekly_period = False
                    
    if type(ws.cell(row=39, column=8)).__name__ != 'MergedCell': ws.cell(row=39, column=8).value = None
    if type(ws.cell(row=40, column=8)).__name__ != 'MergedCell': ws.cell(row=40, column=8).value = None
    
    total_days_final = weekly_holiday_count + public_holiday_count
    ws.cell(row=39, column=8).value = total_days_final
    
    if type(ws.cell(row=42, column=8)).__name__ != 'MergedCell':
        ws.cell(row=42, column=8).value = f"=SUM(H39:H41)"
    
    if not ws.sheet_properties.pageSetUpPr: ws.sheet_properties.pageSetUpPr = PageSetupProperties()
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToHeight = 1; ws.page_setup.fitToWidth = 1
    ws.page_setup.paperSize = ws.PAPERSIZE_A4; ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def generate_report_work(unique_key, roster_data, global_vars, num_days):
    emp_info = st.session_state.employees.get(unique_key)
    if not emp_info: return None
    
    emp_id_str = str(emp_info.get("เลขประจำตัว", "-"))
    if emp_id_str.endswith(".0"): emp_id_str = emp_id_str[:-2]
    
    raw_salary = str(emp_info.get('เงินเดือน', '-'))
    if raw_salary != "-" and raw_salary.replace('.', '', 1).isdigit():
        salary_str = f"{float(raw_salary):,.0f}"
    else:
        salary_str = raw_salary
        
    try: wb = openpyxl.load_workbook("รายงานปฏิบัติงาน.xlsx")
    except: return None
    ws = wb.active

    replacements = {
        "[NAME]": emp_info["ชื่อ-สกุล"],
        "[1]": emp_info["ตำแหน่ง"],
        "[2]": emp_id_str,
        "[3]": salary_str,
        "[14]": global_vars["val_14"],
        "[13]": global_vars["val_13"],
        "[8]": global_vars["val_8"], 
        "[7]": global_vars["val_7"],
        "[17]": ind_vars.get("val_17", ""),
    }

    for r in range(1, 55):
        for c in range(1, 40):
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
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    center_align = Alignment(horizontal='center', vertical='center')
    
    ranges_to_unmerge = []
    for merged_range in list(ws.merged_cells.ranges):
        if start_row <= merged_range.min_row <= start_row + 31 and start_row <= merged_range.max_row <= start_row + 31:
            if merged_range.min_col >= 2 and merged_range.max_col <= 7:
                ranges_to_unmerge.append(merged_range.coord)
    
    for r_coord in ranges_to_unmerge:
        ws.unmerge_cells(r_coord)
        
    def apply_style(r, c, val, color_hex="000000"):
        cell = ws.cell(row=r, column=c)
        if type(cell).__name__ != 'MergedCell':
            cell.value = val
            cell.border = thin_border
            if cell.font:
                nf = copy(cell.font)
                nf.color = color_hex
                cell.font = nf
            else:
                cell.font = Font(color=color_hex)
            return cell
        return None

    for day in range(1, 32):
        row = start_row + day - 1
        
        if day > num_days:
            for c in range(2, 8): 
                cell = ws.cell(row=row, column=c)
                cell.value = None
                cell.border = thin_border
            apply_style(row, 8, "", "000000")
            continue
            
        shift_raw = str(roster_data.get(str(day), "")).strip()
        start_time, end_time = "-", "-"
        is_holiday = False
        
        if shift_raw:
            s_clean = shift_raw.replace("(", "").replace(")", "")
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
            else: start_time, end_time = shift_raw, "" 
            
        font_color = "FF0000" if is_holiday else "000000"

        for c in range(2, 8): 
            cell = ws.cell(row=row, column=c)
            cell.value = None
            cell.border = thin_border

        if is_holiday:
            cell = apply_style(row, 2, start_time, font_color)
            ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=7)
            if cell: cell.alignment = center_align
        else:
            cell1 = apply_style(row, 2, start_time, font_color)
            cell2 = apply_style(row, 5, end_time, font_color)
            ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
            ws.merge_cells(start_row=row, start_column=5, end_row=row, end_column=7)
            if cell1: cell1.alignment = center_align
            if cell2: cell2.alignment = center_align

        apply_style(row, 8, emp_info["ตำแหน่ง"] if start_time not in ["-", ""] and not is_holiday else "", "000000")
            
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
with st.container(border=True):
    st.subheader("🖨️ 3. ส่งออกเอกสาร Excel สำเร็จรูป")
    tab1, tab2, tab3 = st.tabs(["📄 ส่งออก 109", "👤 ส่งออกรายบุคคล", "📦 ส่งออกแบบกลุ่ม (หลายคนพร้อมกัน)"])

    with tab1:
        st.markdown("**ตารางเวร 109 ของทั้งสถานี** (จัดหน้ากระดาษอัตโนมัติ)")
        if st.button("คลิกเพื่อสร้างไฟล์ 109", type="primary"):
            excel_109, total_pages = generate_109(global_data, st.session_state.roster_df, num_days, first_weekday)
            if excel_109:
                st.success(f"สร้างไฟล์ 109 เสร็จสิ้น! (รวม {total_pages} หน้า)")
                st.download_button("📥 ดาวน์โหลดไฟล์ 109 (.xlsx)", data=excel_109, file_name=f"109_{global_data['val_13']}.xlsx")

    with tab2:
        st.markdown("**ข้อมูลเฉพาะบุคคลสำหรับใบเบิก**")
        st.info("💡 **ตารางเวรคือ Master Data:** หากคุณต้องการเปลี่ยนช่วงวันหยุด หรือวันลาพักผ่อน กรุณาแก้ไขเครื่องหมาย `(` `)` `ย` หรือ `พ` ในตารางเวรด้านบน ระบบจะอัปเดตกล่องข้อความด้านล่าง และซิงค์ทุกเอกสารให้อัตโนมัติ 100% ครับ")
        
        active_emp_options = [f"{r['ชื่อ-สกุล']} ({r['Role (หน้าที่)']})" for _, r in st.session_state.roster_df.iterrows()]
        selected_key_177_display = st.selectbox("เลือกพนักงานที่ต้องการสร้างเอกสาร", active_emp_options, key="single_select")
        
        if selected_key_177_display:
            sel_name = selected_key_177_display.split(" (")[0]
            sel_role = selected_key_177_display.split(" (")[1].replace(")", "")
            selected_key_177 = f"{sel_name}_{sel_role}"
            
            saved_ind = local_storage.getItem(f"srt_ind_{selected_key_177}")
            default_ind = {"val_6": "-"}
            try:
                if saved_ind: default_ind.update(json.loads(saved_ind))
            except: pass
            
            emp_row = st.session_state.roster_df[(st.session_state.roster_df['ชื่อ-สกุล'] == sel_name) & (st.session_state.roster_df['Role (หน้าที่)'] == sel_role)].iloc[0]
            roster_dict = {str(d): str(emp_row[str(d)]) if pd.notna(emp_row[str(d)]) else "" for d in range(1, 32)}
            
            val_4_auto, val_5_auto, val_17_auto, val_9_auto, val_10_auto, val_11_auto, val_12_auto = extract_employee_stats(roster_dict)

            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.text_input("วันหยุด [4]", value=val_4_auto, disabled=True)
                st.text_input("หยุดพักผ่อน [9]", value=val_9_auto, disabled=True)
            with c2:
                st.text_input("วันหยุด [5]", value=val_5_auto, disabled=True)
                st.text_input("พักผ่อนตั้งแต่ [10]", value=val_10_auto, disabled=True)
            with c3:
                st.text_input("รวมวันหยุดประจำสัปดาห์ [17]", value=val_17_auto, disabled=True)
                st.text_input("พักผ่อนถึง [11]", value=val_11_auto, disabled=True)
            with c4:
                val_6_manual = st.text_input("เดือนตัวย่อ [6]", value=default_ind.get('val_6', 'ส.ค.69'))
                st.text_input("รวมพักผ่อน [12]", value=val_12_auto, disabled=True)

            export_ind = {
                "val_4": val_4_auto, "val_5": val_5_auto, "val_17": val_17_auto,
                "val_9": val_9_auto, "val_10": val_10_auto, "val_11": val_11_auto, "val_12": val_12_auto,
                "val_6": val_6_manual
            }

            local_storage.setItem(f"srt_ind_{selected_key_177}", json.dumps({"val_6": val_6_manual}), key=f"ls_ind_{uuid.uuid4().hex}")

            st.markdown("<br>", unsafe_allow_html=True)
            col_btn1, col_btn2, col_btn3 = st.columns(3)
            
            with col_btn1:
                if st.button(f"🧾 ออกใบเบิก 177 (ทำล่วงเวลา)", use_container_width=True):
                    excel_177 = generate_177(selected_key_177, roster_dict, global_data, export_ind, num_days)
                    if excel_177:
                        st.success(f"สร้างใบเบิก 177 เสร็จสิ้น!")
                        st.download_button("📥 ดาวน์โหลดไฟล์ 177", data=excel_177, file_name=f"177_{sel_name}.xlsx", use_container_width=True)
                    else:
                        st.error("ไม่พบไฟล์ 'ใบ177 Update.xlsx' ในระบบ (กรุณาอัปโหลดก่อนครับ)")
                        
            with col_btn2:
                if st.button(f"🎉 ออกใบเบิก 178 (วันหยุด)", use_container_width=True):
                    excel_178 = generate_178(selected_key_177, roster_dict, global_data, export_ind, num_days)
                    if excel_178:
                        st.success(f"สร้างใบเบิก 178 เสร็จสิ้น!")
                        st.download_button("📥 ดาวน์โหลดไฟล์ 178", data=excel_178, file_name=f"178_{sel_name}.xlsx", use_container_width=True)
                    else:
                        st.error("ไม่พบไฟล์ '178 อัพเดท.xlsx' ในระบบ (กรุณาอัปโหลดก่อนครับ)")

            with col_btn3:
                if st.button(f"🕒 ออกรายงานปฏิบัติงาน", use_container_width=True):
                    excel_work = generate_report_work(selected_key_177, roster_dict, global_data, num_days)
                    if excel_work:
                        st.success(f"สร้างรายงานปฏิบัติงาน เสร็จสิ้น!")
                        st.download_button("📥 ดาวน์โหลดรายงานฯ", data=excel_work, file_name=f"รายงานปฏิบัติงาน_{sel_name}.xlsx", use_container_width=True)

    with tab3:
        st.markdown("**ดาวน์โหลดเอกสารของพนักงานหลายคนพร้อมกันในคลิกเดียว ระบบจะแยกไฟล์ใส่โฟลเดอร์ให้เป็นระเบียบ**")
        
        saved_ind_batch = local_storage.getItem(f"srt_ind_batch")
        default_val_6 = "ส.ค.69"
        try:
            if saved_ind_batch: default_val_6 = json.loads(saved_ind_batch).get("val_6", "ส.ค.69")
        except: pass
        
        batch_val_6 = st.text_input("เดือนตัวย่อ [6] (ใช้กับทุกคน)", value=default_val_6, key="batch_v6")
        local_storage.setItem(f"srt_ind_batch", json.dumps({"val_6": batch_val_6}), key=f"ls_batch_{uuid.uuid4().hex}")
        
        all_emp_options = [f"{r['ชื่อ-สกุล']} ({r['Role (หน้าที่)']})" for _, r in st.session_state.roster_df.iterrows()]
        
        select_all = st.checkbox("เลือกพนักงานทั้งหมด", value=True)
        if select_all:
            selected_batch_emps = st.multiselect("เลือกพนักงานที่ต้องการส่งออก", all_emp_options, default=all_emp_options)
        else:
            selected_batch_emps = st.multiselect("เลือกพนักงานที่ต้องการส่งออก", all_emp_options)
            
        if st.button("📦 สร้างไฟล์ ZIP ส่งออกแบบกลุ่ม", type="primary", use_container_width=True):
            if not selected_batch_emps:
                st.warning("กรุณาเลือกพนักงานอย่างน้อย 1 คนครับ")
            else:
                with st.spinner(f"กำลังประมวลผลเอกสาร {len(selected_batch_emps)} รายการ..."):
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        for emp_display in selected_batch_emps:
                            sel_name = emp_display.split(" (")[0]
                            sel_role = emp_display.split(" (")[1].replace(")", "")
                            unique_key = f"{sel_name}_{sel_role}"
                            
                            emp_row = st.session_state.roster_df[(st.session_state.roster_df['ชื่อ-สกุล'] == sel_name) & (st.session_state.roster_df['Role (หน้าที่)'] == sel_role)].iloc[0]
                            roster_dict = {str(d): str(emp_row[str(d)]) if pd.notna(emp_row[str(d)]) else "" for d in range(1, 32)}
                            
                            val_4, val_5, val_17, val_9, val_10, val_11, val_12 = extract_employee_stats(roster_dict)
                            
                            export_ind = {
                                "val_4": val_4, "val_5": val_5, "val_17": val_17,
                                "val_9": val_9, "val_10": val_10, "val_11": val_11, "val_12": val_12,
                                "val_6": batch_val_6
                            }
                            
                            excel_177 = generate_177(unique_key, roster_dict, global_data, export_ind, num_days)
                            if excel_177: zip_file.writestr(f"ใบเบิก_177/177_{sel_name}.xlsx", excel_177.getvalue())
                            
                            excel_178 = generate_178(unique_key, roster_dict, global_data, export_ind, num_days)
                            if excel_178: zip_file.writestr(f"ใบเบิก_178/178_{sel_name}.xlsx", excel_178.getvalue())
                            
                            excel_work = generate_report_work(unique_key, roster_dict, global_data, num_days)
                            if excel_work: zip_file.writestr(f"รายงานปฏิบัติงาน/รายงาน_{sel_name}.xlsx", excel_work.getvalue())
                            
                    st.session_state["batch_zip_export"] = zip_buffer.getvalue()
                    
        if "batch_zip_export" in st.session_state:
            st.success("✅ แพ็กไฟล์สำเร็จแล้ว! กดดาวน์โหลดด้านล่างได้เลยครับ")
            st.download_button(
                "📥 ดาวน์โหลดไฟล์ ZIP (รวมทุกเอกสาร แยกโฟลเดอร์ให้แล้ว)",
                data=st.session_state["batch_zip_export"],
                file_name=f"เอกสารเบิกเงิน_รฟท_{global_data['val_13']}.zip",
                mime="application/zip",
                use_container_width=True
            )
