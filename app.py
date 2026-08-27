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

# ==========================================
# 📍 TAB 1: หน้าจัดการข้อมูล
# ==========================================
with tab1:
    with st.sidebar:
        st.header("📂 1. อัปโหลดฐานข้อมูลแม่ (Master)")
        st.info("💡 นำไฟล์ Excel ฐานข้อมูลล่าสุด (ที่มีประวัติเดิม) มาอัปโหลดที่นี่ เพื่อสะสมคอลัมน์ต่อค่ะ")
        master_file = st.file_uploader("อัปโหลดไฟล์ Master Database", type=['xlsx'])
        
        st.markdown("---")
        
        st.header("📅 2. ตั้งค่ารอบการตรวจใหม่")
        num_weeks = st.number_input("จำนวนสัปดาห์ที่ต้องการตั้งค่า", min_value=1, max_value=10, value=1)
        
        def update_col_name(idx):
            d = st.session_state[f"sel_date_{idx}"]
            st.session_state[f"name_{idx}"] = f"สัปดาห์ที่ {d.strftime('%d/%m/')}{d.year + 543}"
        
        weeks_config = []
        for i in range(num_weeks):
            st.markdown(f"**📌 สัปดาห์ที่ {i+1}**")
            sel_date = st.date_input(
                f"1. จิ้มเลือกวันที่ตรวจ", 
                key=f"sel_date_{i}", 
                on_change=update_col_name, 
                args=(i,)
            )
            if f"name_{i}" not in st.session_state:
                st.session_state[f"name_{i}"] = f"สัปดาห์ที่ {sel_date.strftime('%d/%m/')}{sel_date.year + 543}"
                
            week_name = st.text_input(f"2. ชื่อคอลัมน์ (แก้ไขได้)", key=f"name_{i}")
            date_rng = st.date_input(f"3. เลือกช่วงวันที่ครอบคลุม", [], key=f"rng_{i}")
            
            weeks_config.append({'name': week_name, 'range': date_rng})
            st.markdown("---")

    st.header("📥 3. นำเข้าข้อมูลการตรวจรอบใหม่")
    st.info("นำไฟล์ Excel จากระบบมาอัปโหลดที่นี่ (ลากวางได้ 45 ไฟล์พร้อมกันเลยค่ะ)")
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
                st.success(f"✅ โหลดฐานข้อมูลแม่สำเร็จ! (พบนักเรียน {len(existing_df)} คน)")
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
                                # บังคับยัดลงคอลัมน์แรกที่ตั้งค่าไว้ เพื่อรักษาโครงสร้างตาราง
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
                # ใช้ .last() เพื่อให้ข้อมูลที่ถูกประมวลผลทีหลัง ทับข้อมูลก่อนหน้า (ตอบโจทย์ยึดผลล่าสุด)
                new_df = new_df.groupby(groupby_cols, as_index=False).last()
                
                if not existing_df.empty:
                    existing_df_idx = existing_df.set_index("รหัสนักเรียน")
                    new_df_idx = new_df.set_index("รหัสนักเรียน")
                    # combine_first จะเอาข้อมูลใหม่ไปทับข้อมูลเก่าใน Master เสมอ
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
        
        def eval_trend(row):
            statuses = []
            for c in dynamic_cols:
                val = str(row[c]).strip()
                if val != 'nan' and not val.startswith("ไม่ได้ตรวจ") and val != "None":
                    statuses.append(val)
                    
            if not statuses: return "⚪ รอประเมิน"
            latest_stat = statuses[-1] 
            if "ไม่ผ่าน" in latest_stat: return "🔴 ต้องปรับปรุง"
            if "ผ่าน" == latest_stat:
                if all(s == "ผ่าน" for s in statuses): return "⭐⭐⭐ ดีเยี่ยม"
                if len(statuses) >= 2 and statuses[-2] == "ผ่าน": return "⭐⭐ ดี"
                return "🟢 ดีขึ้น"
            return "⚪ รอประเมิน"
            
        final_df["การพัฒนา (สรุปผล)"] = final_df.apply(eval_trend, axis=1)
        
        dashboard_df = final_df.copy()

        st.success("✨ ประมวลผลและเตรียมข้อมูล Master Database เรียบร้อยแล้ว!")
        st.dataframe(final_df, use_container_width=True, hide_index=True)
        st.markdown("---")
        
        # --- สร้างปุ่มดาวน์โหลดไฟล์ Excel Master พร้อมชื่อสุดสมาร์ท ---
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            final_df.to_excel(writer, index=False, sheet_name='Database')
            
        # 📌 ดึงวันที่จาก "1. จิ้มเลือกวันที่ตรวจ" ของสัปดาห์ที่ 1 (sel_date_0)
        if "sel_date_0" in st.session_state:
            check_date_str = st.session_state["sel_date_0"].strftime('%Y%m%d')
        else:
            check_date_str = datetime.now().strftime('%Y%m%d') # สำรองเผื่อระบบหาไม่เจอ
            
        # 📌 วันที่ปัจจุบันที่กดดาวน์โหลด
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
    st.header("📈 แดชบอร์ดผู้บริหาร: สรุปผลการติดตามวินัยนักเรียน")
    
    if not dashboard_df.empty:
        week_cols = [c for c in dashboard_df.columns if "สัปดาห์ที่" in c]
        if week_cols:
            selected_week = st.selectbox("📅 เลือกคอลัมน์สัปดาห์ที่ต้องการดูสรุป", week_cols, index=len(week_cols)-1)
            
            all_rooms = dashboard_df['ห้องเรียน'].unique()
            submitted_rooms = []
            missing_rooms = []
            
            for room in all_rooms:
                room_data = dashboard_df[dashboard_df['ห้องเรียน'] == room]
                is_missing = room_data[selected_week].astype(str).str.contains(r"ไม่ได้ตรวจ|nan|None").all()
                if is_missing:
                    missing_rooms.append(room)
                else:
                    submitted_rooms.append(room)
            
            st.markdown("---")
            st.markdown(f"### 👩‍🏫 การปฏิบัติงานของครูประจำชั้น (ประจำ{selected_week})")
            
            col1, col2, col3 = st.columns(3)
            col1.metric("🏫 จำนวนห้องเรียนทั้งหมด", f"{len(all_rooms)} ห้อง")
            col2.metric("✅ ตรวจและส่งผลแล้ว", f"{len(submitted_rooms)} ห้อง")
            col3.metric("❌ ยังไม่ส่งผลตรวจ", f"{len(missing_rooms)} ห้อง")
            
            if missing_rooms:
                st.error(f"🚨 **รายชื่อห้องที่ยังไม่พบข้อมูลการตรวจ:**\n\n {', '.join(missing_rooms)}")
            else:
                st.success("🎉 ยอดเยี่ยมมากค่ะ! ครูประจำชั้นทุกห้องดำเนินการตรวจและส่งผลครบถ้วน 100%")
                
            st.markdown("---")
            st.markdown(f"### 📊 ภาพรวมสถิตินักเรียน (ประจำ{selected_week})")
            
            total_students = len(dashboard_df)
            passed = len(dashboard_df[dashboard_df[selected_week] == "ผ่าน"])
            failed = len(dashboard_df[dashboard_df[selected_week].astype(str).str.contains("ไม่ผ่าน")])
            not_checked = len(dashboard_df[dashboard_df[selected_week].astype(str).str.contains(r"ไม่ได้ตรวจ|nan|None")])
            
            s_col1, s_col2, s_col3, s_col4 = st.columns(4)
            s_col1.metric("👥 นักเรียนทั้งหมด", f"{total_students} คน")
            s_col2.metric("🟢 ผ่านระเบียบ", f"{passed} คน")
            s_col3.metric("🔴 ไม่ผ่านระเบียบ", f"{failed} คน")
            s_col4.metric("⚪ ไม่ได้ประเมิน/ลา", f"{not_checked} คน")

        else:
            st.info("ยังไม่มีคอลัมน์ข้อมูลสัปดาห์การตรวจในระบบค่ะ")
    else:
        st.warning("👈 กรุณานำเข้าและประมวลผลข้อมูลในแท็บ 'นำเข้าและจัดการข้อมูล' ก่อนนะคะ")
