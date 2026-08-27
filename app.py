import streamlit as st
import pandas as pd
import re
from datetime import datetime
from io import StringIO, BytesIO

# 1. ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="ระบบติดตามวินัยนักเรียน", page_icon="📊", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600&display=swap');
    html, body, [class*="css"] { font-family: 'Sarabun', sans-serif; }
    .stApp { background-color: #FAFAFA; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 ระบบติดตามวินัยนักเรียน (Enterprise Edition)")

# 📌 สร้างระบบแยก 2 หน้าต่าง (Tabs)
tab1, tab2 = st.tabs(["📝 นำเข้าและจัดการข้อมูล", "📊 แดชบอร์ดผู้บริหาร"])

# ตัวแปรกลางสำหรับส่งข้อมูลไปหน้า Dashboard
dashboard_df = pd.DataFrame()

# ตัวแปร Global สำหรับงานทะเบียน
SCHOOL_TOTAL_STUDENTS = 1745
SCHOOL_TOTAL_MALE = 577
SCHOOL_TOTAL_FEMALE = 1168
REG_UPDATED_DATE = "ยังไม่ได้อัปโหลดไฟล์งานทะเบียน"
registry_students_df = pd.DataFrame()
resigned_ids = set()
ALL_SCHOOL_ROOMS = set()

# =========================================================================
# 🚀 โหลดไฟล์อัตโนมัติจาก GitHub (Registry & Master) ทันทีที่เปิดเว็บ
# =========================================================================

# 1. โหลดไฟล์ทะเบียนอัตโนมัติ
default_reg_path = "registry_database.xlsx"
found_exact = False
try:
    xls_default = pd.ExcelFile(default_reg_path)
    for sheet in xls_default.sheet_names:
        sheet_name_str = str(sheet)
        if "ลาออก" in sheet_name_str or "สละสิทธิ์" in sheet_name_str:
            df_resigned = pd.read_excel(xls_default, sheet_name=sheet, header=None, skiprows=3)
            for col in df_resigned.columns:
                ids = df_resigned[col].astype(str).str.extract(r'^(\d{5})$').dropna()[0].tolist()
                resigned_ids.update(ids)
                
        if "สรุป" in sheet_name_str:
            df_sum = pd.read_excel(xls_default, sheet_name=sheet, header=None)
            for i, row in df_sum.iterrows():
                for val in row:
                    if "สำรวจเมื่อ" in str(val):
                        match = re.search(r'สำรวจเมื่อ\s*วันที่\s*(.*)', str(val))
                        if match:
                            REG_UPDATED_DATE = match.group(1).strip()
                if row.astype(str).str.contains('ม.ปลาย').any():
                    nums = pd.to_numeric(row, errors='coerce').dropna().tolist()
                    if len(nums) >= 3:
                        SCHOOL_TOTAL_MALE = int(nums[-3])
                        SCHOOL_TOTAL_FEMALE = int(nums[-2])
                        SCHOOL_TOTAL_STUDENTS = int(nums[-1])
                        found_exact = True
                        
    students_list = []
    current_room = ""
    for sheet in ['M4', 'M5', 'M6']:
        if sheet in xls_default.sheet_names:
            df_sheet = pd.read_excel(xls_default, sheet_name=sheet)
            for i, row in df_sheet.iterrows():
                val = str(row.iloc[1])
                match = re.search(r'ม\.?\s*(\d)[\./](\d+)', val)
                if match:
                    current_room = f"ม. {match.group(1)}/{match.group(2)}"
                    ALL_SCHOOL_ROOMS.add(current_room)
                    
                student_id = str(row.get('Unnamed: 2', '')).strip()
                if re.match(r'^\d{5}$', student_id):
                    prefix = str(row.get('Unnamed: 3', '')).strip() if pd.notna(row.get('Unnamed: 3')) else ''
                    fname = str(row.get('Unnamed: 4', '')).strip() if pd.notna(row.get('Unnamed: 4')) else ''
                    lname = str(row.get('Unnamed: 5', '')).strip() if pd.notna(row.get('Unnamed: 5')) else ''
                    name = f"{prefix}{fname} {lname}".strip()
                    
                    no_val = str(row.iloc[1]).strip()
                    std_no = int(float(no_val)) if no_val.replace('.', '', 1).isdigit() else ""
                    
                    students_list.append({
                        'เลขที่': std_no,
                        'รหัสนักเรียน': student_id, 
                        'ชื่อ-สกุล': name, 
                        'ห้องเรียน': current_room
                    })
    
    if students_list:
        temp_df = pd.DataFrame(students_list)
        registry_students_df = temp_df[~temp_df['รหัสนักเรียน'].isin(resigned_ids)]
        
    if found_exact:
        REG_UPDATED_DATE = f"{REG_UPDATED_DATE} (จากระบบอัตโนมัติ)"
except Exception:
    pass

# 2. โหลดไฟล์ Master Database อัตโนมัติ
existing_df = pd.DataFrame()
default_master_path = "master_database.xlsx"
try:
    existing_df = pd.read_excel(default_master_path, sheet_name='Database')
    if 'รหัสนักเรียน' in existing_df.columns:
        existing_df['รหัสนักเรียน'] = existing_df['รหัสนักเรียน'].astype(str).str.replace(r'\.0$', '', regex=True)
        existing_df['ลำดับ'] = existing_df['ลำดับ'].astype(int, errors='ignore')
        if "การพัฒนา (สรุปผล)" in existing_df.columns:
            existing_df = existing_df.drop(columns=["การพัฒนา (สรุปผล)"])
except Exception:
    pass


# ==========================================
# 📍 TAB 1: หน้าจัดการข้อมูล
# ==========================================
with tab1:
    with st.sidebar:
        st.header("📂 1. อัปโหลดไฟล์ประชากรนักเรียน (ฝ่ายทะเบียน)")
        st.info("💡 โหลดข้อมูลจาก GitHub อัตโนมัติ สามารถอัปโหลดไฟล์ใหม่เพื่อทับได้")
        reg_file = st.file_uploader("อัปโหลดไฟล์ประชากรนักเรียน (อัปเดตใหม่)", type=['xls', 'xlsx'])
        
        if reg_file:
            try:
                xls = pd.ExcelFile(reg_file)
                found_exact = False
                resigned_ids = set()
                ALL_SCHOOL_ROOMS = set()
                
                for sheet in xls.sheet_names:
                    sheet_name_str = str(sheet)
                    if "ลาออก" in sheet_name_str or "สละสิทธิ์" in sheet_name_str:
                        df_resigned = pd.read_excel(xls, sheet_name=sheet, header=None, skiprows=3)
                        for col in df_resigned.columns:
                            ids = df_resigned[col].astype(str).str.extract(r'^(\d{5})$').dropna()[0].tolist()
                            resigned_ids.update(ids)
                            
                    if "สรุป" in sheet_name_str:
                        df_sum = pd.read_excel(xls, sheet_name=sheet, header=None)
                        for i, row in df_sum.iterrows():
                            for val in row:
                                if "สำรวจเมื่อ" in str(val):
                                    match = re.search(r'สำรวจเมื่อ\s*วันที่\s*(.*)', str(val))
                                    if match:
                                        REG_UPDATED_DATE = match.group(1).strip()
                            if row.astype(str).str.contains('ม.ปลาย').any():
                                nums = pd.to_numeric(row, errors='coerce').dropna().tolist()
                                if len(nums) >= 3:
                                    SCHOOL_TOTAL_MALE = int(nums[-3])
                                    SCHOOL_TOTAL_FEMALE = int(nums[-2])
                                    SCHOOL_TOTAL_STUDENTS = int(nums[-1])
                                    found_exact = True
                                    
                students_list = []
                current_room = ""
                for sheet in ['M4', 'M5', 'M6']:
                    if sheet in xls.sheet_names:
                        df_sheet = pd.read_excel(xls, sheet_name=sheet)
                        for i, row in df_sheet.iterrows():
                            val = str(row.iloc[1])
                            match = re.search(r'ม\.?\s*(\d)[\./](\d+)', val)
                            if match:
                                current_room = f"ม. {match.group(1)}/{match.group(2)}"
                                ALL_SCHOOL_ROOMS.add(current_room)
                                
                            student_id = str(row.get('Unnamed: 2', '')).strip()
                            if re.match(r'^\d{5}$', student_id):
                                prefix = str(row.get('Unnamed: 3', '')).strip() if pd.notna(row.get('Unnamed: 3')) else ''
                                fname = str(row.get('Unnamed: 4', '')).strip() if pd.notna(row.get('Unnamed: 4')) else ''
                                lname = str(row.get('Unnamed: 5', '')).strip() if pd.notna(row.get('Unnamed: 5')) else ''
                                name = f"{prefix}{fname} {lname}".strip()
                                
                                no_val = str(row.iloc[1]).strip()
                                std_no = int(float(no_val)) if no_val.replace('.', '', 1).isdigit() else ""
                                
                                students_list.append({
                                    'เลขที่': std_no,
                                    'รหัสนักเรียน': student_id, 
                                    'ชื่อ-สกุล': name, 
                                    'ห้องเรียน': current_room
                                })
                
                if students_list:
                    temp_df = pd.DataFrame(students_list)
                    registry_students_df = temp_df[~temp_df['รหัสนักเรียน'].isin(resigned_ids)]

                if found_exact:
                    st.success(f"✅ ดึงยอดปัจจุบันสำเร็จ: {SCHOOL_TOTAL_STUDENTS} คน (อัปเดต: {REG_UPDATED_DATE})")
                else:
                    st.warning("ดึงยอดสรุปไม่ได้ ใช้ค่าตั้งต้นแทนค่ะ")
            except Exception as e:
                st.error(f"⚠️ อ่านไฟล์ทะเบียนไม่สำเร็จ: {e}")
                
        SCHOOL_TOTAL_ROOMS = st.number_input("🏫 จำนวนห้องเรียนทั้งหมดในระบบ", min_value=1, max_value=100, value=len(ALL_SCHOOL_ROOMS) if ALL_SCHOOL_ROOMS else 45)
        
        st.markdown("---")
        st.header("📂 2. อัปโหลดฐานข้อมูลแม่ (Master)")
        master_file = st.file_uploader("อัปโหลดไฟล์ Master Database (อัปเดตใหม่)", type=['xlsx'])
        
        st.markdown("---")
        st.header("📅 3. ตั้งค่ารอบการตรวจใหม่")
        num_weeks = st.number_input("จำนวนสัปดาห์ที่ต้องการตั้งค่า", min_value=1, max_value=10, value=1)
        
        def update_col_name(idx):
            d = st.session_state[f"sel_date_{idx}"]
            st.session_state[f"name_{idx}"] = f"สัปดาห์ที่ {d.strftime('%d/%m/')}{d.year + 543}"
        
        weeks_config = []
        for i in range(num_weeks):
            st.markdown(f"**📌 สัปดาห์ที่ {i+1}**")
            sel_date = st.date_input(f"จิ้มเลือกวันที่ตรวจ", key=f"sel_date_{i}", on_change=update_col_name, args=(i,))
            if f"name_{i}" not in st.session_state:
                st.session_state[f"name_{i}"] = f"สัปดาห์ที่ {sel_date.strftime('%d/%m/')}{sel_date.year + 543}"
            week_name = st.text_input(f"ชื่อคอลัมน์ (แก้ไขได้)", key=f"name_{i}")
            date_rng = st.date_input(f"เลือกช่วงวันที่ครอบคลุม", [], key=f"rng_{i}")
            weeks_config.append({'name': week_name, 'range': date_rng})
            st.markdown("---")

    st.header("📥 4. นำเข้าข้อมูลการตรวจรอบใหม่")
    st.info("นำไฟล์ Excel จากระบบมาอัปโหลดที่นี่ (ลากวาง 45 ไฟล์พร้อมกันเลยค่ะ)")
    uploaded_files = st.file_uploader("ลากไฟล์การตรวจรอบใหม่มาวางที่นี่", type=['xls', 'xlsx'], accept_multiple_files=True)

    def extract_info(html_text):
        date_match = re.search(r'วันที่\s*(\d{2}/\d{2}/\d{4})', html_text)
        return date_match.group(1) if date_match else None

    def sort_rooms(room_str):
        try:
            nums = re.findall(r'\d+', str(room_str))
            if len(nums) >= 2:
                return int(nums[0]) * 100 + int(nums[1])
        except:
            pass
        return 9999

    if master_file:
        try:
            existing_df = pd.read_excel(master_file, sheet_name='Database')
            if 'รหัสนักเรียน' in existing_df.columns:
                existing_df['รหัสนักเรียน'] = existing_df['รหัสนักเรียน'].astype(str).str.replace(r'\.0$', '', regex=True)
                existing_df['ลำดับ'] = existing_df['ลำดับ'].astype(int, errors='ignore')
                if "การพัฒนา (สรุปผล)" in existing_df.columns:
                    existing_df = existing_df.drop(columns=["การพัฒนา (สรุปผล)"])
                st.success(f"✅ โหลดฐานข้อมูลแม่สำเร็จ! (พบประวัตินักเรียน {len(existing_df)} คน)")
        except Exception as e:
            st.error(f"⚠️ โหลดไฟล์ฐานข้อมูลแม่ไม่สำเร็จ: {e}")

    final_df = None

    if uploaded_files:
        all_ranges_valid = all(len(w['range']) == 2 for w in weeks_config)
        if not all_ranges_valid:
            st.warning("⚠️ กรุณาเลือกช่วงวันที่ให้ครบ 2 วันในแถบตั้งค่าด้านซ้ายนะคะ")
        else:
            all_students = []
            for file in uploaded_files:
                try:
                    content = file.getvalue().decode('utf-8', errors='ignore')
                    file_date_str = extract_info(content)
                    matched_week_name = None
                    skip_this_file = False
                    
                    if file_date_str:
                        file_date = datetime.strptime(file_date_str, "%d/%m/%Y").date()
                        check_date = file_date.replace(year=file_date.year - 543) if file_date.year > 2400 else file_date
                        
                        for w in weeks_config:
                            start_date, end_date = w['range'][0], w['range'][1]
                            if start_date <= check_date <= end_date:
                                matched_week_name = w['name']
                                break
                        
                        if not matched_week_name:
                            user_choice = st.radio(
                                f"ต้องการดำเนินการอย่างไรกับไฟล์ {file.name}?",
                                ["❌ ยกเลิก (ไม่ใช้ไฟล์นี้)", "✅ ดำเนินการต่อ (ใช้ไฟล์นี้เป็นข้อมูล)"],
                                key=f"choice_{file.name}", horizontal=True
                            )
                            if user_choice == "❌ ยกเลิก (ไม่ใช้ไฟล์นี้)": skip_this_file = True
                            else: matched_week_name = weeks_config[0]['name']
                            
                    if skip_this_file: continue
                    if not matched_week_name: matched_week_name = weeks_config[0]['name']
                    
                    html_io = StringIO(content)
                    dfs = pd.read_html(html_io)
                    
                    for df in dfs:
                        if 'ชื่อนักเรียน' in df.to_string():
                            df.columns = [col[-1] for col in df.columns]
                            for idx, row in df.iterrows():
                                if str(row.get('ลำดับ', '')).strip().isdigit():
                                    raw_room = str(row.get('ห้องเรียน', '')).strip()
                                    match = re.search(r'ม\.?\s*(\d)[\./](\d+)', raw_room)
                                    room = f"ม. {match.group(1)}/{match.group(2)}" if match else raw_room
                                    ALL_SCHOOL_ROOMS.add(room)
                                    
                                    passed = str(row.get('ผ่าน', '')).strip() == '/'
                                    remarks = str(row.get('หมายเหตุ', '')).strip()
                                    failed_reasons = []
                                    cols = df.columns.tolist()
                                    start_idx = cols.index('ผ่าน') + 1
                                    end_idx = cols.index('หมายเหตุ')
                                    for col_idx in range(start_idx, end_idx):
                                        if str(row.iloc[col_idx]).strip() == '/':
                                            failed_reasons.append(cols[col_idx])
                                    
                                    if passed: status = "ผ่าน"
                                    elif failed_reasons: status = f"ไม่ผ่าน ({failed_reasons[0]})"
                                    elif remarks and remarks != 'nan': status = f"ไม่ได้ตรวจ ({remarks})"
                                    else: status = "ไม่ได้ตรวจ"
                                    
                                    all_students.append({
                                        "ลำดับ": int(row['ลำดับ']),
                                        "รหัสนักเรียน": str(row.get('รหัสนักเรียน', '')).strip(),
                                        "ชื่อนักเรียน": str(row['ชื่อนักเรียน']).strip(),
                                        "ห้องเรียน": room,
                                        matched_week_name: status
                                    })
                except Exception as e:
                    st.error(f"เกิดข้อผิดพลาดในการอ่านไฟล์ {file.name}: {e}")

            if all_students:
                new_df = pd.DataFrame(all_students)
                groupby_cols = ["ลำดับ", "รหัสนักเรียน", "ชื่อนักเรียน", "ห้องเรียน"]
                new_df = new_df.groupby(groupby_cols, as_index=False).last()
                
                if not existing_df.empty:
                    existing_df_idx = existing_df.set_index("รหัสนักเรียน")
                    new_df_idx = new_df.set_index("รหัสนักเรียน")
                    merged_idx = new_df_idx.combine_first(existing_df_idx)
                    final_df = merged_idx.reset_index()
                else:
                    final_df = new_df
    else:
        if not existing_df.empty:
            final_df = existing_df.copy()

    if final_df is not None and not final_df.empty:
        final_df['room_sort'] = final_df['ห้องเรียน'].apply(sort_rooms)
        final_df = final_df.sort_values(by=['room_sort', 'ลำดับ']).drop(columns=['room_sort'])
        
        fixed_cols = ["ลำดับ", "รหัสนักเรียน", "ชื่อนักเรียน", "ห้องเรียน"]
        dynamic_cols = [c for c in final_df.columns if c not in fixed_cols and c != "การพัฒนา (สรุปผล)"]
        final_df = final_df[fixed_cols + dynamic_cols]
        
        if resigned_ids:
            for w in dynamic_cols:
                mask = final_df['รหัสนักเรียน'].astype(str).isin(resigned_ids) & \
                       final_df[w].astype(str).str.contains(r"ไม่ได้ตรวจ|nan|None", na=True)
                final_df.loc[mask, w] = "⚪ ลาออก"
        
        def eval_trend(row):
            statuses = []
            for c in dynamic_cols:
                val = str(row[c]).strip()
                if val != 'nan' and not val.startswith("ไม่ได้ตรวจ") and val != "None" and val != "⚪ ลาออก":
                    statuses.append(val)
                    
            if str(row[dynamic_cols[-1]]).strip() == "⚪ ลาออก": return "⚫ พ้นสภาพ/ลาออก"
            if not statuses: return "⚪ รอประเมิน"
            
            latest_stat = statuses[-1] 
            if "ไม่ผ่าน" in latest_stat: return "🔴 ต้องปรับปรุง"
            if "ผ่าน" == latest_stat:
                if len(statuses) >= 3 and all(s == "ผ่าน" for s in statuses[-3:]): return "⭐⭐⭐ ดีเยี่ยม"
                elif all(s == "ผ่าน" for s in statuses): return "⭐⭐⭐ ดีเยี่ยม"
                elif len(statuses) >= 2 and statuses[-2] == "ผ่าน": return "⭐⭐ ดี"
                else: return "🟢 ดีขึ้น"
            return "⚪ รอประเมิน"
            
        final_df["การพัฒนา (สรุปผล)"] = final_df.apply(eval_trend, axis=1)
        dashboard_df = final_df.copy()

        st.success("✨ ประมวลผลและเตรียมข้อมูล Master Database เรียบร้อยแล้ว!")
        st.dataframe(final_df, use_container_width=True, hide_index=True)
        st.markdown("---")
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            final_df.to_excel(writer, index=False, sheet_name='Database')
            
        check_date_str = st.session_state.get("sel_date_0", datetime.now()).strftime('%Y%m%d')
        current_date_str = datetime.now().strftime('%Y%m%d')
        
        st.download_button(
            label="📥 ดาวน์โหลดฐานข้อมูล (Master Database) ไปเก็บไว้",
            data=output.getvalue(),
            file_name=f"Master_Database_{check_date_str}_{current_date_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

# ==========================================
# 📍 TAB 2: หน้าแดชบอร์ดผู้บริหาร
# ==========================================
with tab2:
    st.info(f"📅 **ข้อมูลประชากรนักเรียนอัปเดตล่าสุด:** {REG_UPDATED_DATE}")
    st.header("📈 แดชบอร์ดผู้บริหาร: สรุปผลการติดตามวินัยนักเรียน")
    
    def get_gender(name):
        name_str = str(name).strip()
        if name_str.startswith(('นาย', 'ด.ช.', 'เด็กชาย')): return 'ชาย'
        if name_str.startswith(('นางสาว', 'ด.ญ.', 'น.ส.', 'เด็กหญิง')): return 'หญิง'
        return 'ไม่ระบุ'
        
    def render_student_table(df_to_render):
        if not df_to_render.empty:
            df_show = df_to_render.copy()
            if "ลำดับ" in df_show.columns:
                df_show = df_show.rename(columns={"ลำดับ": "เลขที่", "ชื่อนักเรียน": "ชื่อ-สกุล"})
            cols = ["เลขที่", "รหัสนักเรียน", "ชื่อ-สกุล", "ห้องเรียน"]
            cols = [c for c in cols if c in df_show.columns]
            st.dataframe(df_show[cols], use_container_width=True, hide_index=True)
        else:
            st.info("ไม่มีข้อมูลในหมวดหมู่นี้ค่ะ")

    if not dashboard_df.empty:
        dashboard_df['เพศ'] = dashboard_df['ชื่อนักเรียน'].apply(get_gender)
        if not registry_students_df.empty:
            registry_students_df['เพศ'] = registry_students_df['ชื่อ-สกุล'].apply(get_gender)
            
        week_cols = [c for c in dashboard_df.columns if "สัปดาห์ที่" in c]
        
        if week_cols:
            options = ["🌟 สรุปภาพรวมทั้งหมด"] + week_cols
            selected_option = st.selectbox("📅 เลือกคอลัมน์สัปดาห์ที่ต้องการดูสรุป", options, index=0)
            st.markdown("---")
            
            all_present_rooms = set(dashboard_df['ห้องเรียน'].unique())
            formatted_registry_rooms = {r.replace("ม.", "ม. ").replace("  ", " ") for r in ALL_SCHOOL_ROOMS}
            final_all_rooms = sorted(list(all_present_rooms.union(formatted_registry_rooms)), key=sort_rooms)
            
            if selected_option == "🌟 สรุปภาพรวมทั้งหมด":
                st.markdown(f"### 👩‍🏫 สถิติการส่งผลประเมินของครูประจำชั้น (สะสมทั้งหมด)")
                
                submitted_perfect = []
                missing_latest = []
                room_missing_stats = []
                latest_week = week_cols[-1]
                
                for room in final_all_rooms:
                    room_data = dashboard_df[dashboard_df['ห้องเรียน'] == room]
                    missed_weeks = []
                    
                    if room_data.empty:
                        missed_weeks = week_cols
                    else:
                        for w in week_cols:
                            has_evaluated = room_data[w].astype(str).str.contains(r"ผ่าน|ไม่ผ่าน").any()
                            if not has_evaluated:
                                missed_weeks.append(w)
                                
                    if not missed_weeks:
                        submitted_perfect.append(room)
                    else:
                        room_missing_stats.append({
                            "ห้องเรียน": room,
                            "ขาดส่ง (ครั้ง)": len(missed_weeks),
                            "สัปดาห์ที่ขาด": ", ".join(missed_weeks)
                        })
                        
                    if latest_week in missed_weeks:
                        missing_latest.append(room)
                
                col1, col2, col3 = st.columns(3, gap="large")
                with col1:
                    st.metric("🏫 จำนวนห้องเรียนทั้งหมด", f"{SCHOOL_TOTAL_ROOMS} ห้อง")
                    with st.expander("👉 รายชื่อนักเรียนทั้งหมด (จากทะเบียน)"):
                        if not registry_students_df.empty: render_student_table(registry_students_df)
                        else: st.info("อัปโหลดไฟล์จากฝ่ายทะเบียนเพื่อดูรายชื่อค่ะ")
                with col2:
                    st.metric("🏆 ส่งครบ 100% ทุกรอบ", f"{len(submitted_perfect)} ห้อง")
                    with st.expander("👉 ดูห้องที่ส่งครบ 100%"):
                        st.write(", ".join(submitted_perfect) if submitted_perfect else "-")
                with col3:
                    st.metric("🚨 ขาดส่ง (สัปดาห์ล่าสุด)", f"{len(missing_latest)} ห้อง")
                    with st.expander("👉 ประวัติขาดส่งสะสม (ทุกห้อง)"):
                        if room_missing_stats:
                            df_miss = pd.DataFrame(room_missing_stats).sort_values(by="ขาดส่ง (ครั้ง)", ascending=False)
                            st.dataframe(df_miss, use_container_width=True, hide_index=True)
                        else:
                            st.success("เยี่ยมมาก! ทุกห้องส่งครบ 100% ไม่มีประวัติขาดส่งเลยค่ะ")
                
                st.markdown("---")
                st.markdown(f"### 📊 ภาพรวมสถิตินักเรียนแบบสะสม (เกณฑ์การพัฒนา)")
                
                def has_been_checked(row):
                    for w in week_cols:
                        if re.search(r"ผ่าน|ไม่ผ่าน", str(row[w])): return True
                    return False
                    
                checked_mask = dashboard_df.apply(has_been_checked, axis=1)
                checked_overall_df = dashboard_df[checked_mask]
                
                total_checked = len(checked_overall_df)
                male_checked = len(checked_overall_df[checked_overall_df['เพศ'] == 'ชาย'])
                female_checked = len(checked_overall_df[checked_overall_df['เพศ'] == 'หญิง'])
                
                df_reg_male = registry_students_df[registry_students_df['เพศ'] == 'ชาย'] if not registry_students_df.empty else pd.DataFrame()
                df_reg_female = registry_students_df[registry_students_df['เพศ'] == 'หญิง'] if not registry_students_df.empty else pd.DataFrame()
                
                c1, c2, c3 = st.columns(3, gap="large")
                with c1:
                    st.metric("👥 นักเรียนสถานะปัจจุบัน (หักลาออก)", f"{SCHOOL_TOTAL_STUDENTS} คน")
                    with st.expander("👉 ดูรายชื่อทั้งหมด"): render_student_table(registry_students_df)
                with c2:
                    st.metric("👦 ชาย", f"{SCHOOL_TOTAL_MALE} คน")
                    with st.expander("👉 ดูรายชื่อ (ชาย)"): render_student_table(df_reg_male)
                with c3:
                    st.metric("👧 หญิง", f"{SCHOOL_TOTAL_FEMALE} คน")
                    with st.expander("👉 ดูรายชื่อ (หญิง)"): render_student_table(df_reg_female)
                
                st.markdown("##### 📌 จำนวนนักเรียนที่ได้รับการประเมินแล้ว (อย่างน้อย 1 ครั้ง)")
                sc1, sc2, sc3 = st.columns(3, gap="large")
                with sc1:
                    st.metric("✅ รวมได้รับการประเมิน", f"{total_checked} คน")
                    with st.expander("👉 ดูรายชื่อ"): render_student_table(checked_overall_df)
                with sc2:
                    st.metric("👦 ชายที่ได้รับการประเมิน", f"{male_checked} คน")
                    with st.expander("👉 ดูรายชื่อ (ชาย)"): render_student_table(checked_overall_df[checked_overall_df['เพศ'] == 'ชาย'])
                with sc3:
                    st.metric("👧 หญิงที่ได้รับการประเมิน", f"{female_checked} คน")
                    with st.expander("👉 ดูรายชื่อ (หญิง)"): render_student_table(checked_overall_df[checked_overall_df['เพศ'] == 'หญิง'])

                # ❌ เพิ่มส่วนนี้: กลุ่มนักเรียนที่ไม่ได้รับการประเมิน (พร้อม Dropdown ครบชุด)
                not_checked_overall_df = dashboard_df[~checked_mask]
                total_not_checked = len(not_checked_overall_df)
                male_not_checked = len(not_checked_overall_df[not_checked_overall_df['เพศ'] == 'ชาย'])
                female_not_checked = len(not_checked_overall_df[not_checked_overall_df['เพศ'] == 'หญิง'])

                st.markdown("##### ❌ จำนวนนักเรียนที่ไม่ได้รับการประเมิน (ยังไม่เคยถูกตรวจเลย)")
                nc1, nc2, nc3 = st.columns(3, gap="large")
                with nc1:
                    st.metric("❌ รวมไม่ได้รับการประเมิน", f"{total_not_checked} คน")
                    with st.expander("👉 ดูรายชื่อ"): render_student_table(not_checked_overall_df)
                with nc2:
                    st.metric("❌ ชายที่ไม่ได้รับการประเมิน", f"{male_not_checked} คน")
                    with st.expander("👉 ดูรายชื่อ (ชาย)"): render_student_table(not_checked_overall_df[not_checked_overall_df['เพศ'] == 'ชาย'])
                with nc3:
                    st.metric("❌ หญิงที่ไม่ได้รับการประเมิน", f"{female_not_checked} คน")
                    with st.expander("👉 ดูรายชื่อ (หญิง)"): render_student_table(not_checked_overall_df[not_checked_overall_df['เพศ'] == 'หญิง'])
                
                st.markdown("##### 📌 สรุปเกณฑ์การพัฒนาล่าสุดของนักเรียน")
                df_ex = dashboard_df[dashboard_df["การพัฒนา (สรุปผล)"] == "⭐⭐⭐ ดีเยี่ยม"]
                df_gd = dashboard_df[dashboard_df["การพัฒนา (สรุปผล)"] == "⭐⭐ ดี"]
                df_im = dashboard_df[dashboard_df["การพัฒนา (สรุปผล)"] == "🟢 ดีขึ้น"]
                df_nd = dashboard_df[dashboard_df["การพัฒนา (สรุปผล)"] == "🔴 ต้องปรับปรุง"]
                df_re = dashboard_df[dashboard_df["การพัฒนา (สรุปผล)"] == "⚫ พ้นสภาพ/ลาออก"]
                
                rc1, rc2, rc3, rc4, rc5 = st.columns(5, gap="medium")
                with rc1:
                    st.metric("⭐⭐⭐ ดีเยี่ยม", f"{len(df_ex)} คน")
                    with st.expander("👉 ดูรายชื่อ"): render_student_table(df_ex)
                with rc2:
                    st.metric("⭐⭐ ดี", f"{len(df_gd)} คน")
                    with st.expander("👉 ดูรายชื่อ"): render_student_table(df_gd)
                with rc3:
                    st.metric("🟢 ดีขึ้น", f"{len(df_im)} คน")
                    with st.expander("👉 ดูรายชื่อ"): render_student_table(df_im)
                with rc4:
                    st.metric("🔴 ต้องปรับปรุง", f"{len(df_nd)} คน")
                    with st.expander("👉 ดูรายชื่อ"): render_student_table(df_nd)
                with rc5:
                    st.metric("⚫ พ้นสภาพ/ลาออก", f"{len(df_re)} คน")
                    with st.expander("👉 ดูรายชื่อ"): render_student_table(df_re)
                
            else:
                selected_week = selected_option
                st.markdown(f"### 👩‍🏫 การปฏิบัติงานของครูประจำชั้น (ประจำ{selected_week})")
                submitted_rooms = []
                missing_rooms = []
                
                for room in final_all_rooms:
                    room_data = dashboard_df[dashboard_df['ห้องเรียน'] == room]
                    if room_data.empty:
                        missing_rooms.append(room)
                        continue
                        
                    has_evaluated = room_data[selected_week].astype(str).str.contains(r"ผ่าน|ไม่ผ่าน").any()
                    if not has_evaluated:
                        missing_rooms.append(room)
                    else:
                        submitted_rooms.append(room)
                
                col1, col2, col3 = st.columns(3, gap="large")
                with col1:
                    st.metric("🏫 จำนวนห้องเรียนทั้งหมด", f"{SCHOOL_TOTAL_ROOMS} ห้อง")
                    with st.expander("👉 รายชื่อทั้งหมด (จากทะเบียน)"):
                        if not registry_students_df.empty: render_student_table(registry_students_df)
                        else: st.info("อัปโหลดไฟล์จากฝ่ายทะเบียนเพื่อดูรายชื่อค่ะ")
                with col2:
                    st.metric("✅ ตรวจและส่งผลแล้ว", f"{len(submitted_rooms)} ห้อง")
                    with st.expander("👉 ดูห้องที่ส่งผล"): st.write(", ".join(submitted_rooms) if submitted_rooms else "-")
                with col3:
                    st.metric("❌ ยังไม่ส่งผลตรวจ", f"{len(missing_rooms)} ห้อง")
                    with st.expander("👉 ดูห้องที่ยังไม่ส่ง"): st.write(", ".join(missing_rooms) if missing_rooms else "-")
                    
                st.markdown("---")
                
                st.markdown(f"### 📌 สถิติการตรวจประชากรประจำ{selected_week}")
                
                checked_df = dashboard_df[dashboard_df[selected_week].astype(str).str.contains(r"ผ่าน|ไม่ผ่าน")]
                total_checked = len(checked_df)
                male_checked = len(checked_df[checked_df['เพศ'] == 'ชาย'])
                female_checked = len(checked_df[checked_df['เพศ'] == 'หญิง'])
                
                w_c1, w_c2, w_c3 = st.columns(3, gap="large")
                with w_c1:
                    st.metric("👥 นักเรียนที่ได้รับการตรวจ", f"{total_checked} คน")
                    with st.expander("👉 ดูรายชื่อ (ที่ตรวจแล้ว)"): render_student_table(checked_df)
                with w_c2:
                    st.metric("👦 ชายที่ได้รับการตรวจ", f"{male_checked} คน")
                    with st.expander("👉 ดูรายชื่อ (ชายที่ตรวจแล้ว)"): render_student_table(checked_df[checked_df['เพศ'] == 'ชาย'])
                with w_c3:
                    st.metric("👧 หญิงที่ได้รับการตรวจ", f"{female_checked} คน")
                    with st.expander("👉 ดูรายชื่อ (หญิงที่ตรวจแล้ว)"): render_student_table(checked_df[checked_df['เพศ'] == 'หญิง'])
                
                st.markdown("---")
                st.markdown(f"### 📊 ภาพรวมสถิตินักเรียน (ประจำ{selected_week})")
                
                df_pass = checked_df[checked_df[selected_week] == "ผ่าน"]
                df_fail = checked_df[checked_df[selected_week].astype(str).str.contains("ไม่ผ่าน")]
                df_resigned = dashboard_df[dashboard_df[selected_week] == "⚪ ลาออก"]
                df_missing = dashboard_df[~dashboard_df.index.isin(checked_df.index) & ~dashboard_df.index.isin(df_resigned.index)]
                
                missing_students_count = SCHOOL_TOTAL_STUDENTS - total_checked
                
                df_reg_male = registry_students_df[registry_students_df['เพศ'] == 'ชาย'] if not registry_students_df.empty else pd.DataFrame()
                df_reg_female = registry_students_df[registry_students_df['เพศ'] == 'หญิง'] if not registry_students_df.empty else pd.DataFrame()
                
                c1, c2, c3 = st.columns(3, gap="large")
                with c1:
                    st.metric("👥 นักเรียนสถานะปัจจุบัน", f"{SCHOOL_TOTAL_STUDENTS} คน")
                    with st.expander("👉 ดูรายชื่อทั้งหมด"): render_student_table(registry_students_df)
                with c2:
                    st.metric("👦 ชาย", f"{SCHOOL_TOTAL_MALE} คน")
                    with st.expander("👉 ดูรายชื่อ (ชาย)"): render_student_table(df_reg_male)
                with c3:
                    st.metric("👧 หญิง", f"{SCHOOL_TOTAL_FEMALE} คน")
                    with st.expander("👉 ดูรายชื่อ (หญิง)"): render_student_table(df_reg_female)
                
                st.markdown("##### 📌 ผลการตรวจระเบียบรายสัปดาห์")
                sc1, sc2, sc3, sc4 = st.columns(4, gap="medium")
                with sc1:
                    st.metric("🟢 ผ่านระเบียบ", f"{len(df_pass)} คน")
                    with st.expander("👉 ดูรายชื่อ"): render_student_table(df_pass)
                with sc2:
                    st.metric("🔴 ไม่ผ่านระเบียบ", f"{len(df_fail)} คน")
                    with st.expander("👉 duodenum"): render_student_table(df_fail) # ปรับแสดงผลปกติ
                with sc3:
                    st.metric("⚫ ลาออก", f"{len(df_resigned)} คน")
                    with st.expander("👉 ดูรายชื่อ"): render_student_table(df_resigned)
                with sc4:
                    st.metric("⚪ ขาด/ยังไม่ได้ประเมิน", f"{missing_students_count} คน")
                    with st.expander("👉 ดูรายชื่อ"): render_student_table(df_missing)

        else:
            st.info("ยังไม่มีคอลัมน์ข้อมูลสัปดาห์การตรวจในระบบค่ะ")
    else:
        st.warning("👈 กรุณานำเข้าและประมวลผลข้อมูลในแท็บ 'นำเข้าและจัดการข้อมูล' ก่อนนะคะ")
