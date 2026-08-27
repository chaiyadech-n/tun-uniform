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

# ==========================================
# 📍 TAB 1: หน้าจัดการข้อมูล
# ==========================================
with tab1:
    with st.sidebar:
        # --- 📌 ส่วนที่ 1: ข้อมูลประชากรจากฝ่ายทะเบียน ---
        st.header("📂 1. อัปโหลดไฟล์ประชากรนักเรียน (ฝ่ายทะเบียน)")
        st.info("💡 อัปโหลดไฟล์เพื่ออัปเดตยอดประชากรจริง และซิงค์รายชื่อนักเรียนอัตโนมัติ")
        reg_file = st.file_uploader("อัปโหลดไฟล์ประชากรนักเรียน", type=['xls', 'xlsx'])
        
        if reg_file:
            try:
                xls = pd.ExcelFile(reg_file)
                found_exact = False
                
                # 1. หาเด็กลาออก และหาหน้าสรุปยอด
                for sheet in xls.sheet_names:
                    sheet_name_str = str(sheet) # หุ้มเกราะให้ชื่อ Sheet
                    
                    if "ลาออก" in sheet_name_str or "สละสิทธิ์" in sheet_name_str:
                        df_resigned = pd.read_excel(xls, sheet_name=sheet, header=None, skiprows=3)
                        for col in df_resigned.columns:
                            ids = df_resigned[col].astype(str).str.extract(r'^(\d{5})$').dropna()[0].tolist()
                            resigned_ids.update(ids)
                            
                    if "สรุป" in sheet_name_str:
                        df_sum = pd.read_excel(xls, sheet_name=sheet, header=None)
                        for i, row in df_sum.iterrows():
                            for val in row:
                                # 📌 จุดแก้บั๊ก: หุ้มเกราะให้ val เป็น str(val) เสมอ ป้องกัน float (ช่องว่าง)
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
                                    
                # 2. ดึงรายชื่อเด็กทั้งหมดมาเก็บไว้ในลิ้นชัก
                students_list = []
                for sheet in ['M4', 'M5', 'M6']:
                    if sheet in xls.sheet_names:
                        df_sheet = pd.read_excel(xls, sheet_name=sheet, skiprows=3)
                        for i, row in df_sheet.iterrows():
                            student_id = str(row['Unnamed: 2']).strip()
                            if re.match(r'^\d{5}$', student_id):
                                prefix = str(row['Unnamed: 3']).strip() if pd.notna(row['Unnamed: 3']) else ''
                                fname = str(row['Unnamed: 4']).strip() if pd.notna(row['Unnamed: 4']) else ''
                                lname = str(row['Unnamed: 5']).strip() if pd.notna(row['Unnamed: 5']) else ''
                                name = f"{prefix}{fname} {lname}".strip()
                                students_list.append({'รหัสนักเรียน': student_id, 'ชื่อ-สกุล': name, 'ระดับชั้น': sheet})
                
                if students_list:
                    temp_df = pd.DataFrame(students_list)
                    registry_students_df = temp_df[~temp_df['รหัสนักเรียน'].isin(resigned_ids)]

                if found_exact:
                    st.success(f"✅ ดึงยอดปัจจุบันสำเร็จ: {SCHOOL_TOTAL_STUDENTS} คน (อัปเดต: {REG_UPDATED_DATE})")
                else:
                    st.warning("ดึงยอดสรุปไม่ได้ ใช้ค่าตั้งต้นแทนค่ะ")
            except Exception as e:
                st.error(f"⚠️ อ่านไฟล์ทะเบียนไม่สำเร็จ: {e}")
                
        SCHOOL_TOTAL_ROOMS = st.number_input("🏫 จำนวนห้องเรียนทั้งหมดในระบบ (ปรับแก้ได้)", min_value=1, max_value=100, value=45)
        
        st.markdown("---")

        # --- 📌 ส่วนที่ 2: อัปโหลด Master Database ---
        st.header("📂 2. อัปโหลดฐานข้อมูลแม่ (Master)")
        master_file = st.file_uploader("อัปโหลดไฟล์ Master Database", type=['xlsx'])
        
        st.markdown("---")
        
        # --- 📌 ส่วนที่ 3: ตั้งค่าสัปดาห์ ---
        st.header("📅 3. ตั้งค่ารอบการตรวจใหม่")
        num_weeks = st.number_input("จำนวนสัปดาห์ที่ต้องการตั้งค่า", min_value=1, max_value=10, value=1)
        
        def update_col_name(idx):
            d = st.session_state[f"sel_date_{idx}"]
            st.session_state[f"name_{idx}"] = f"สัปดาห์ที่ {d.strftime('%d/%m/')}{d.year + 543}"
        
        weeks_config = []
        for i in range(num_weeks):
            st.markdown(f"**📌 สัปดาห์ที่ {i+1}**")
            sel_date = st.date_input(
                f"จิ้มเลือกวันที่ตรวจ", 
                key=f"sel_date_{i}", 
                on_change=update_col_name, 
                args=(i,)
            )
            if f"name_{i}" not in st.session_state:
                st.session_state[f"name_{i}"] = f"สัปดาห์ที่ {sel_date.strftime('%d/%m/')}{sel_date.year + 543}"
                
            week_name = st.text_input(f"ชื่อคอลัมน์ (แก้ไขได้)", key=f"name_{i}")
            date_rng = st.date_input(f"เลือกช่วงวันที่ครอบคลุม", [], key=f"rng_{i}")
            
            weeks_config.append({'name': week_name, 'range': date_rng})
            st.markdown("---")

    # --- 📌 พื้นที่หลัก: อัปโหลดไฟล์การตรวจ ---
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

    existing_df = pd.DataFrame()

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
                            st.warning(f"⚠️ ตรวจพบไฟล์ '{file.name}' (วันที่ {file_date_str}) ไม่อยู่ในช่วงเวลาที่กำหนดค่ะ")
                            user_choice = st.radio(
                                f"ต้องการดำเนินการอย่างไรกับไฟล์ {file.name}?",
                                ["❌ ยกเลิก (ไม่นำเข้าไฟล์นี้)", "✅ ดำเนินการต่อ (บันทึกลงคอลัมน์สัปดาห์ที่ตั้งค่าไว้เพื่อแทนที่ข้อมูลเดิม)"],
                                key=f"choice_{file.name}",
                                horizontal=True
                            )
                            if user_choice == "❌ ยกเลิก (ไม่นำเข้าไฟล์นี้)":
                                skip_this_file = True
                            else:
                                matched_week_name = weeks_config[0]['name']
                                
                    if skip_this_file:
                        continue
                    if not matched_week_name:
                        matched_week_name = weeks_config[0]['name']
                    
                    html_io = StringIO(content)
                    dfs = pd.read_html(html_io)
                    
                    for df in dfs:
                        if 'ชื่อนักเรียน' in df.to_string():
                            df.columns = [col[-1] for col in df.columns]
                            for idx, row in df.iterrows():
                                if str(row.get('ลำดับ', '')).strip().isdigit():
                                    room = str(row.get('ห้องเรียน', '')).strip()
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
        
        # จัดการเด็กลาออกอัตโนมัติ
        if resigned_ids:
            for w in dynamic_cols:
                mask = final_df['รหัสนักเรียน'].astype(str).isin(resigned_ids) & \
                       final_df[w].astype(str).str.contains(r"ไม่ได้ตรวจ|nan|None", na=True)
                final_df.loc[mask, w] = "⚪ ลาออก"
        
        # เกณฑ์ให้ดาว
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
            
        if "sel_date_0" in st.session_state:
            check_date_str = st.session_state["sel_date_0"].strftime('%Y%m%d')
        else:
            check_date_str = datetime.now().strftime('%Y%m%d')
            
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
    # 📌 ส่วนแสดงวันที่อัปเดตข้อมูลให้โดดเด่น
    st.info(f"📅 **ข้อมูลประชากรนักเรียนอัปเดตล่าสุด:** {REG_UPDATED_DATE}")
    st.header("📈 แดชบอร์ดผู้บริหาร: สรุปผลการติดตามวินัยนักเรียน")
    
    # 📌 ครอบคลุมทุกคำนำหน้าที่เป็นไปได้
    def get_gender(name):
        name_str = str(name).strip()
        if name_str.startswith(('นาย', 'ด.ช.', 'เด็กชาย')): return 'ชาย'
        if name_str.startswith(('นางสาว', 'ด.ญ.', 'น.ส.', 'เด็กหญิง')): return 'หญิง'
        return 'ไม่ระบุ'
        
    if not dashboard_df.empty:
        dashboard_df['เพศ'] = dashboard_df['ชื่อนักเรียน'].apply(get_gender)
        
        week_cols = [c for c in dashboard_df.columns if "สัปดาห์ที่" in c]
        if week_cols:
            options = ["🌟 สรุปภาพรวมทั้งหมด"] + week_cols
            selected_option = st.selectbox("📅 เลือกคอลัมน์สัปดาห์ที่ต้องการดูสรุป", options, index=0)
            st.markdown("---")
            
            if selected_option == "🌟 สรุปภาพรวมทั้งหมด":
                st.markdown(f"### 👩‍🏫 การปฏิบัติงานของครูประจำชั้น (สรุปภาพรวมทุกสัปดาห์)")
                all_present_rooms = dashboard_df['ห้องเรียน'].unique()
                submitted_all_weeks_rooms = []
                missing_any_week_rooms = []
                
                for room in all_present_rooms:
                    room_data = dashboard_df[dashboard_df['ห้องเรียน'] == room]
                    missed_weeks = []
                    for w in week_cols:
                        has_evaluated = room_data[w].astype(str).str.contains(r"ผ่าน|ไม่ผ่าน").any()
                        if not has_evaluated:
                            missed_weeks.append(w)
                    if missed_weeks:
                        missing_any_week_rooms.append(room)
                    else:
                        submitted_all_weeks_rooms.append(room)
                        
                missing_count = SCHOOL_TOTAL_ROOMS - len(submitted_all_weeks_rooms)
                
                col1, col2, col3 = st.columns(3)
                col1.metric("🏫 จำนวนห้องเรียนทั้งหมด", f"{SCHOOL_TOTAL_ROOMS} ห้อง")
                col2.metric("✅ ตรวจและส่งครบทุกรอบ", f"{len(submitted_all_weeks_rooms)} ห้อง")
                col3.metric("❌ ส่งผลตรวจไม่ครบ/ขาดส่ง", f"{missing_count} ห้อง")
                
                # 📌 กล่องจิ้มดูรายละเอียดที่เจ้านายขอค่ะ
                c_exp1, c_exp2, c_exp3 = st.columns(3)
                with c_exp1:
                    with st.expander("👉 รายชื่อนักเรียน (จากไฟล์ทะเบียน)"):
                        if not registry_students_df.empty:
                            st.dataframe(registry_students_df, use_container_width=True, hide_index=True)
                        else:
                            st.info("อัปโหลดไฟล์จากฝ่ายทะเบียนเพื่อดูรายชื่อค่ะ")
                with c_exp2:
                    with st.expander("👉 ดูห้องที่ส่งครบทุกรอบ"):
                        st.write(", ".join(submitted_all_weeks_rooms) if submitted_all_weeks_rooms else "-")
                with c_exp3:
                    with st.expander("👉 ดูห้องที่ขาดส่ง"):
                        st.write(", ".join(missing_any_week_rooms) if missing_any_week_rooms else "-")
                
                st.markdown("---")
                st.markdown(f"### 📊 ภาพรวมสถิตินักเรียนแบบสะสม (เกณฑ์การพัฒนา)")
                
                def has_been_checked(row):
                    for w in week_cols:
                        if re.search(r"ผ่าน|ไม่ผ่าน", str(row[w])):
                            return True
                    return False
                    
                checked_mask = dashboard_df.apply(has_been_checked, axis=1)
                checked_overall_df = dashboard_df[checked_mask]
                
                total_checked = len(checked_overall_df)
                male_checked = len(checked_overall_df[checked_overall_df['เพศ'] == 'ชาย'])
                female_checked = len(checked_overall_df[checked_overall_df['เพศ'] == 'หญิง'])
                
                c1, c2, c3 = st.columns(3)
                c1.metric("👥 นักเรียนสถานะปัจจุบัน (หักลาออก)", f"{SCHOOL_TOTAL_STUDENTS} คน")
                c2.metric("👦 ชาย", f"{SCHOOL_TOTAL_MALE} คน")
                c3.metric("👧 หญิง", f"{SCHOOL_TOTAL_FEMALE} คน")
                
                st.markdown("##### 📌 จำนวนนักเรียนที่ได้รับการประเมินแล้ว (อย่างน้อย 1 ครั้ง)")
                sc1, sc2, sc3 = st.columns(3)
                sc1.metric("✅ รวมได้รับการประเมิน", f"{total_checked} คน")
                sc2.metric("👦 ชายที่ได้รับการประเมิน", f"{male_checked} คน")
                sc3.metric("👧 หญิงที่ได้รับการประเมิน", f"{female_checked} คน")
                
                st.markdown("##### 📌 สรุปเกณฑ์การพัฒนาล่าสุดของนักเรียน")
                excellent = len(dashboard_df[dashboard_df["การพัฒนา (สรุปผล)"] == "⭐⭐⭐ ดีเยี่ยม"])
                good = len(dashboard_df[dashboard_df["การพัฒนา (สรุปผล)"] == "⭐⭐ ดี"])
                improved = len(dashboard_df[dashboard_df["การพัฒนา (สรุปผล)"] == "🟢 ดีขึ้น"])
                needs_improvement = len(dashboard_df[dashboard_df["การพัฒนา (สรุปผล)"] == "🔴 ต้องปรับปรุง"])
                resigned_count = len(dashboard_df[dashboard_df["การพัฒนา (สรุปผล)"] == "⚫ พ้นสภาพ/ลาออก"])
                
                rc1, rc2, rc3, rc4, rc5 = st.columns(5)
                rc1.metric("⭐⭐⭐ ดีเยี่ยม", f"{excellent} คน")
                rc2.metric("⭐⭐ ดี", f"{good} คน")
                rc3.metric("🟢 ดีขึ้น", f"{improved} คน")
                rc4.metric("🔴 ต้องปรับปรุง", f"{needs_improvement} คน")
                rc5.metric("⚫ พ้นสภาพ/ลาออก", f"{resigned_count} คน")
                
            else:
                selected_week = selected_option
                st.markdown(f"### 👩‍🏫 การปฏิบัติงานของครูประจำชั้น (ประจำ{selected_week})")
                all_present_rooms = dashboard_df['ห้องเรียน'].unique()
                submitted_rooms = []
                missing_rooms = []
                
                for room in all_present_rooms:
                    room_data = dashboard_df[dashboard_df['ห้องเรียน'] == room]
                    has_evaluated = room_data[selected_week].astype(str).str.contains(r"ผ่าน|ไม่ผ่าน").any()
                    if not has_evaluated:
                        missing_rooms.append(room)
                    else:
                        submitted_rooms.append(room)
                
                missing_count = SCHOOL_TOTAL_ROOMS - len(submitted_rooms)
                
                col1, col2, col3 = st.columns(3)
                col1.metric("🏫 จำนวนห้องเรียนทั้งหมด", f"{SCHOOL_TOTAL_ROOMS} ห้อง")
                col2.metric("✅ ตรวจและส่งผลแล้ว", f"{len(submitted_rooms)} ห้อง")
                col3.metric("❌ ยังไม่ส่งผลตรวจ", f"{missing_count} ห้อง")
                
                # 📌 กล่องจิ้มดูรายละเอียด (รายสัปดาห์)
                c_exp1, c_exp2, c_exp3 = st.columns(3)
                with c_exp1:
                    with st.expander("👉 รายชื่อนักเรียน (จากไฟล์ทะเบียน)"):
                        if not registry_students_df.empty:
                            st.dataframe(registry_students_df, use_container_width=True, hide_index=True)
                        else:
                            st.info("อัปโหลดไฟล์จากฝ่ายทะเบียนเพื่อดูรายชื่อค่ะ")
                with c_exp2:
                    with st.expander("👉 ดูห้องที่ส่งผล"):
                        st.write(", ".join(submitted_rooms) if submitted_rooms else "-")
                with c_exp3:
                    with st.expander("👉 ดูห้องที่ยังไม่ส่ง"):
                        st.write(", ".join(missing_rooms) if missing_rooms else "-")
                    
                st.markdown("---")
                st.markdown(f"### 📊 ภาพรวมสถิตินักเรียน (ประจำ{selected_week})")
                
                checked_df = dashboard_df[dashboard_df[selected_week].astype(str).str.contains(r"ผ่าน|ไม่ผ่าน")]
                total_checked = len(checked_df)
                male_checked = len(checked_df[checked_df['เพศ'] == 'ชาย'])
                female_checked = len(checked_df[checked_df['เพศ'] == 'หญิง'])
                
                passed = len(checked_df[checked_df[selected_week] == "ผ่าน"])
                failed = len(checked_df[checked_df[selected_week].astype(str).str.contains("ไม่ผ่าน")])
                resigned = len(dashboard_df[dashboard_df[selected_week] == "⚪ ลาออก"])
                missing_students = SCHOOL_TOTAL_STUDENTS - total_checked
                
                c1, c2, c3 = st.columns(3)
                c1.metric("👥 นักเรียนสถานะปัจจุบัน", f"{SCHOOL_TOTAL_STUDENTS} คน")
                c2.metric("👦 ชาย", f"{SCHOOL_TOTAL_MALE} คน")
                c3.metric("👧 หญิง", f"{SCHOOL_TOTAL_FEMALE} คน")
                
                st.markdown("##### 📌 ผลการตรวจระเบียบรายสัปดาห์")
                sc1, sc2, sc3, sc4 = st.columns(4)
                sc1.metric("🟢 ผ่านระเบียบ", f"{passed} คน")
                sc2.metric("🔴 ไม่ผ่านระเบียบ", f"{failed} คน")
                sc3.metric("⚫ ลาออก", f"{resigned} คน")
                sc4.metric("⚪ ขาด/ยังไม่ได้ประเมิน", f"{missing_students} คน")

        else:
            st.info("ยังไม่มีคอลัมน์ข้อมูลสัปดาห์การตรวจในระบบค่ะ")
    else:
        st.warning("👈 กรุณานำเข้าและประมวลผลข้อมูลในแท็บ 'นำเข้าและจัดการข้อมูล' ก่อนนะคะ")
