import streamlit as st
import pandas as pd
import re
from datetime import datetime
from io import StringIO

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

# 2. เมนูด้านข้างสำหรับตั้งค่าและนำเข้าข้อมูล
with st.sidebar:
    st.header("📅 1. ตั้งค่ารอบการตรวจ")
    
    num_weeks = st.number_input("จำนวนสัปดาห์ที่ต้องการตั้งค่า", min_value=1, max_value=10, value=1)
    
    weeks_config = []
    for i in range(num_weeks):
        st.markdown(f"**📌 สัปดาห์ที่ {i+1}**")
        
        # คืนชีพปฏิทินจิ้มเลือกวันที่ เพื่อสร้างชื่ออัตโนมัติ
        sel_date = st.date_input(f"1. จิ้มเลือกวันที่ตรวจ", key=f"sel_date_{i}")
        default_name = f"สัปดาห์ที่ {sel_date.strftime('%d/%m/')}{sel_date.year + 543}"
        
        # นำชื่อมาใส่กล่อง เผื่อเจ้านายอยากพิมพ์แก้ไข
        week_name = st.text_input(f"2. ชื่อคอลัมน์ (แก้ไขได้)", value=default_name, key=f"name_{i}")
        
        # เลือกช่วงวันที่
        date_rng = st.date_input(f"3. เลือกช่วงวันที่ครอบคลุม", [], key=f"rng_{i}")
        
        weeks_config.append({'name': week_name, 'range': date_rng})
        st.markdown("---")
    
    st.header("📥 2. นำเข้าข้อมูล")
    st.info("นำไฟล์ Excel จากระบบดิจิทัลติดตามนักเรียนมาอัปโหลดที่นี่")
    uploaded_files = st.file_uploader("ลากไฟล์ Excel ทุกห้องมาวางพร้อมกัน", type=['xls', 'xlsx'], accept_multiple_files=True)

# 3. ฟังก์ชันดึงวันที่และจัดเรียงห้อง
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

# 4. ประมวลผลและสร้างตาราง
if uploaded_files:
    all_ranges_valid = all(len(w['range']) == 2 for w in weeks_config)
    
    if not all_ranges_valid:
        st.warning("⚠️ กรุณาเลือกช่วงวันที่ให้ครบ 2 วัน (วันเริ่มต้น - วันสิ้นสุด) ในทุกๆ สัปดาห์ที่ตั้งค่าไว้นะคะ")
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
                        st.warning(f"⚠️ ตรวจพบไฟล์ '{file.name}' (วันที่ {file_date_str}) ไม่อยู่ในช่วงเวลาที่กำหนดไว้เลยค่ะ")
                        
                        user_choice = st.radio(
                            f"ต้องการดำเนินการอย่างไรกับไฟล์ {file.name}?",
                            ["❌ ยกเลิก (ไม่นำเข้าไฟล์นี้)", "✅ ดำเนินการต่อ (สร้างคอลัมน์ใหม่ตามวันที่ในไฟล์)"],
                            key=f"choice_{file.name}",
                            horizontal=True
                        )
                        
                        if user_choice == "❌ ยกเลิก (ไม่นำเข้าไฟล์นี้)":
                            skip_this_file = True
                        else:
                            matched_week_name = f"สัปดาห์ที่ {file_date.strftime('%d/%m/')}{file_date.year}"
                            
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
            final_df = pd.DataFrame(all_students)
            
            groupby_cols = ["ลำดับ", "รหัสนักเรียน", "ชื่อนักเรียน", "ห้องเรียน"]
            final_df = final_df.groupby(groupby_cols, as_index=False).last()
            
            final_df['room_sort'] = final_df['ห้องเรียน'].apply(sort_rooms)
            final_df = final_df.sort_values(by=['room_sort', 'ลำดับ']).drop(columns=['room_sort'])
            
            week_cols = [c for c in final_df.columns if c not in groupby_cols and c != "การพัฒนา (สรุปผล)"]
            def eval_trend(row):
                statuses = [str(row[c]) for c in week_cols if pd.notna(row[c])]
                if not statuses: return "🟢 ทรงตัว"
                latest_stat = statuses[-1] 
                if "ไม่ผ่าน" in latest_stat: return "🔴 ต้องปรับปรุง"
                if all("ผ่าน" == s for s in statuses if s != "ไม่ได้ตรวจ"): return "🌟 ดีเยี่ยม"
                return "🟢 ทรงตัว"
                
            final_df["การพัฒนา (สรุปผล)"] = final_df.apply(eval_trend, axis=1)

            st.success("✨ ระบบจัดการเรียงข้อมูลและสรุปผลเรียบร้อยแล้ว")
            st.dataframe(final_df, use_container_width=True, hide_index=True)
else:
    st.info("👈 กรุณาตั้งค่ารอบการตรวจ และอัปโหลดไฟล์ Excel ที่เมนูด้านซ้ายมือ")
