import streamlit as st
import pandas as pd
import re
from datetime import datetime

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
    period_name = st.text_input("ชื่อสัปดาห์ (เช่น สัปดาห์ที่ 26 ส.ค. 2569)", "สัปดาห์ที่ 1")
    date_range = st.date_input("เลือกช่วงวันที่ตรวจ", [])
    
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
    if len(date_range) < 2:
        st.warning("⚠️ กรุณาเลือกช่วงวันที่ในแถบด้านข้างให้ครบ 2 วัน (วันเริ่มต้น - วันสิ้นสุด)")
    else:
        start_date, end_date = date_range[0], date_range[1]
        all_students = []
        
        for file in uploaded_files:
            try:
                content = file.getvalue().decode('utf-8', errors='ignore')
                file_date_str = extract_info(content)
                
                if file_date_str:
                    file_date = datetime.strptime(file_date_str, "%d/%m/%Y").date()
                    if not (start_date <= file_date <= end_date):
                        st.error(f"⚠️ ไฟล์ {file.name} มีวันที่ตรวจ ({file_date_str}) ไม่อยู่ในช่วงที่ตั้งไว้ ระบบจะทำการข้ามไฟล์นี้")
                        continue
                        
                dfs = pd.read_html(content)
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
                                    period_name: status
                                })
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาดในการอ่านไฟล์ {file.name}: {e}")

        if all_students:
            final_df = pd.DataFrame(all_students)
            final_df['room_sort'] = final_df['ห้องเรียน'].apply(sort_rooms)
            final_df = final_df.sort_values(by=['room_sort', 'ลำดับ']).drop(columns=['room_sort'])
            
            def eval_trend(stat):
                if "ผ่าน" == stat: return "🌟 ดีเยี่ยม"
                if "ไม่ผ่าน" in stat: return "🔴 ต้องปรับปรุง"
                return "🟢 ทรงตัว"
            final_df["การพัฒนา (สรุปผล)"] = final_df[period_name].apply(eval_trend)

            st.success("✨ ระบบจัดการเรียงข้อมูลและสรุปผลเรียบร้อยแล้ว")
            st.dataframe(final_df, use_container_width=True, hide_index=True)
else:
    st.info("👈 กรุณาตั้งค่ารอบการตรวจ และอัปโหลดไฟล์ Excel ที่เมนูด้านซ้ายมือ")
