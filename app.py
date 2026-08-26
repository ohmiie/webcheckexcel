import streamlit as st
import openpyxl
import pandas as pd

# ---------------------------------------------------------
# การตั้งค่าหน้าเว็บ
# ---------------------------------------------------------
st.set_page_config(page_title="ระบบประเมินทักษะ Excel", layout="wide", page_icon="📊")

# สร้างตัวแปรเก็บข้อมูลตารางคะแนนรวม (ถ้ายังไม่มีให้สร้างใหม่)
if 'results_df' not in st.session_state:
    st.session_state['results_df'] = pd.DataFrame(columns=[
        "ลำดับ", "โรงเรียน", "สังกัด", "จังหวัด", "ทีม", "ชื่อ-สกุล ผู้เข้าแข่งขัน",
        "การจัดเตรียมแผ่นงานและบันทึกข้อมูล\n(1 คะแนน)",
        'การคำนวณและประมวลผลข้อมูลใน Sheet "Student_Scores" (24 คะแนน)',
        'การเชื่อมโยงและจัดเรียงข้อมูล\nใน Sheet "Student_Scores2"\n(7 คะแนน)',
        'การรายงานตารางสรุปข้อมูลใน Sheet "Report_Table"\n(12 คะแนน)',
        'การสร้างแผ่นงานประมวลผลเกรด Sheet "Grade_Summary"\n(2 คะแนน)',
        'การคำนวณเกรดและสถานะประเมินใน Sheet "Grade_Summary"\n(18 คะแนน)',
        'การเพิ่มแผ่นงานแดชบอร์ด Sheet "Report_Dashboard"\n(2 คะแนน)',
        'การสร้างแดชบอร์ดสรุปผลและแผนภูมิใน Sheet "Report_Dashboard"\n(24 คะแนน)',
        'การตั้งค่าหน้ากระดาษและการพิมพ์รูปแบบ PDF\n(10 คะแนน)',
        "รวม\n(100 คะแนน)"
    ])

st.title("📊 ระบบประเมินและให้คะแนนทักษะวิชาการ Excel")

# ==========================================
# ส่วนฟอร์มกรอกข้อมูลผู้เข้าแข่งขัน
# ==========================================
with st.expander("📝 1. กรอกข้อมูลผู้เข้าแข่งขัน", expanded=True):
    col_i1, col_i2, col_i3 = st.columns(3)
    order_no = col_i1.text_input("ลำดับ", placeholder="เช่น 1")
    school = col_i2.text_input("โรงเรียน", placeholder="เช่น วิทยาลัยอาชีวศึกษา...")
    affiliation = col_i3.text_input("สังกัด", placeholder="เช่น เทศบาลนครนครปฐม")
    
    col_i4, col_i5, col_i6 = st.columns(3)
    province = col_i4.text_input("จังหวัด", placeholder="เช่น นครปฐม")
    team = col_i5.text_input("ทีม", placeholder="เช่น ทีม A")
    student_name = col_i6.text_input("ชื่อ-สกุล ผู้เข้าแข่งขัน", placeholder="เช่น นางสาวเม ทา ซิน ปุย")

# ==========================================
# อัปโหลดไฟล์
# ==========================================
uploaded_file = st.file_uploader("📂 2. อัปโหลดไฟล์ Excel ของผู้เข้าแข่งขัน (.xlsx)", type=["xlsx"])

# ตัวแปรเก็บคะแนนแต่ละหมวด
c1, c2, c3, c4, c5, c6, c7, c8, c9 = 0, 0, 0, 0, 0, 0, 0, 0, 0

def get_safe_formula(cell):
    if cell and cell.value:
        return str(cell.value).upper()
    return ""

if uploaded_file is not None:
    st.info("🔄 กำลังประมวลผลไฟล์ กรุณารอสักครู่...")
    
    try:
        wb_formula = openpyxl.load_workbook(uploaded_file, data_only=False)
        wb_data = openpyxl.load_workbook(uploaded_file, data_only=True)
        sheet_names = wb_formula.sheetnames
        
        col_left, col_right = st.columns([1.5, 1])
        
        with col_left:
            st.header("🤖 ส่วนตรวจอัตโนมัติ (Auto-Grading)")
            
            # --- หมวด 1: เตรียมแผ่นงาน (1 คะแนน) ---
            if "Student_Scores" in sheet_names:
                c1 += 1
                st.write("✅ 1.1 เปลี่ยนชื่อ Sheet เป็น 'Student_Scores'")
            
            # --- หมวด 2: คำนวณ Student_Scores (สูตร 18 + Manual 6 = 24 คะแนน) ---
            if "Student_Scores" in sheet_names:
                ws_form = wb_formula["Student_Scores"]
                ws_dat = wb_data["Student_Scores"]
                
                if ws_dat.max_row >= 50: c2 += 2
                if ws_dat.max_column >= 10: c2 += 2
                
                if any("SUM(" in get_safe_formula(ws_form[f'I{r}']) for r in range(3, 10)): c2 += 3
                if any("IF(" in get_safe_formula(ws_form[f'K{r}']) for r in range(3, 10)): c2 += 4
                if any("-" in get_safe_formula(ws_form[f'L{r}']) for r in range(3, 10)): c2 += 2
                
                if "Net_Score_Data" in wb_formula.defined_names: c2 += 2
                
                if any("RANK" in get_safe_formula(ws_form[f'M{r}']) for r in range(3, 10)): c2 += 3
                
                st.success(f"คำนวณ Student_Scores อัตโนมัติได้: {c2}/18 คะแนน (รอ Manual อีก 6)")

            # --- หมวด 3: Student_Scores2 (สูตร 4 + Manual 3 = 7 คะแนน) ---
            if "Student_Scores2" in sheet_names:
                c3 += 1
                ws_form2 = wb_formula["Student_Scores2"]
                if any("STUDENT_SCORES!" in get_safe_formula(ws_form2[f'B{r}']) for r in range(2, 10)): c3 += 3
                st.success(f"เชื่อมโยง Student_Scores2 อัตโนมัติได้: {c3}/4 คะแนน (รอ Manual อีก 3)")

            # --- หมวด 5: สร้าง Grade_Summary (สูตร 1 + Manual 1 = 2 คะแนน) ---
            if "Grade_Summary" in sheet_names:
                c5 += 1

            # --- หมวด 6: คำนวณ Grade_Summary (สูตร 18 = 18 คะแนน) ---
            if "Grade_Summary" in sheet_names:
                ws_grade = wb_formula["Grade_Summary"]
                if any("STUDENT_SCORES!" in get_safe_formula(ws_grade[f'B{r}']) for r in range(3, 10)): c6 += 4
                if any(get_safe_formula(ws_grade[f'F{r}']).count("IF(") >= 4 for r in range(3, 10)): c6 += 8
                if any("IF(" in get_safe_formula(ws_grade[f'G{r}']) for r in range(3, 10)): c6 += 6
                st.success(f"คำนวณเกรด Grade_Summary ได้: {c6}/18 คะแนน")

            # --- หมวด 7: แดชบอร์ด (การเพิ่ม Sheet 2 คะแนน) ---
            if "Report_Dashboard" in sheet_names:
                c7 += 2
            
            # --- หมวด 8: แดชบอร์ดสรุปผล (สูตร 12 + Manual 12 = 24 คะแนน) ---
            if "Report_Dashboard" in sheet_names:
                ws_dash = wb_formula["Report_Dashboard"]
                dash_formulas = " ".join([get_safe_formula(cell) for row in ws_dash.iter_rows() for cell in row if cell.data_type == 'f'])
                
                if "COUNT(" in dash_formulas or "COUNTA(" in dash_formulas: c8 += 2
                if dash_formulas.count("COUNTIF(") >= 2: c8 += 10
                st.success(f"คำนวณ Dashboard อัตโนมัติได้: {c8}/12 คะแนน (รอ Manual อีก 12)")

        with col_right:
            st.header("✍️ ส่วนตรวจด้วยสายตา (Manual)")
            
            with st.expander("คะแนนส่วนกราฟและรูปแบบ", expanded=True):
                if st.checkbox("2.7 แทรกกราฟ Line Sparklines (3 คะแนน)"): c2 += 3
                if st.checkbox("2.8 คำนวณ Average, Max, Min ท้ายตาราง (3 คะแนน)"): c2 += 3
                if st.checkbox("3.3 จัดเรียงข้อมูล (Sort) 1-50 (3 คะแนน)"): c3 += 3
                if st.checkbox("5.2 โครงสร้างหัวตาราง Grade_Summary (1 คะแนน)"): c5 += 1
                if st.checkbox("7.2 โครงสร้าง Dashboard สวยงาม (3 คะแนน)"): c8 += 3
            
            with st.expander("ข้อ 4: Report_Table (12 คะแนน)", expanded=True):
                if st.checkbox("4.1 มีชีต 'Report_Table' (1 คะแนน)"): c4 += 1
                if st.checkbox("4.2 สร้าง PivotTable สรุปถูกต้อง (5 คะแนน)"): c4 += 5
                if st.checkbox("4.3 สร้าง PivotChart รูปแบบ Column (3 คะแนน)"): c4 += 3
                if st.checkbox("4.4 มีเครื่องมือ Slicer 'กลุ่มเรียน' (3 คะแนน)"): c4 += 3
                
            with st.expander("ข้อ 7: Report_Dashboard", expanded=True):
                if st.checkbox("7.7 สร้าง PivotTable สรุปเกรด (4 คะแนน)"): c8 += 4
                if st.checkbox("7.8 สร้าง PivotChart รูปแบบ Column (3 คะแนน)"): c8 += 3
                if st.checkbox("7.9 มีเครื่องมือ Slicer 'กลุ่มเรียน' (2 คะแนน)"): c8 += 2

            with st.expander("ข้อ 8: ไฟล์ PDF และการพิมพ์ (10 คะแนน)", expanded=True):
                if st.checkbox("8.1 Grade_Summary แนวตั้ง A4, Scale 90% (2 คะแนน)"): c9 += 2
                if st.checkbox("8.1.1 Margins (บน/ล่าง 1, ซ้าย/ขวา 0.5, H/F 1) (2 คะแนน)"): c9 += 2
                if st.checkbox("8.1.2 มี Header ขวา / Footer กลาง และส่งเป็น PDF (2 คะแนน)"): c9 += 2
                if st.checkbox("8.2 กราฟ SEC04 แนวนอน, A4, Margins (2 คะแนน)"): c9 += 2
                if st.checkbox("8.3 ส่งเป็น PDF ชื่อไฟล์ 'SEC04' (2 คะแนน)"): c9 += 2

        # รวมคะแนนทั้งหมด
        total_score = c1 + c2 + c3 + c4 + c5 + c6 + c7 + c8 + c9
        
        st.divider()
        st.markdown(f"<h1 style='text-align: center; color: #1f77b4;'>🏆 คะแนนรวมสุทธิ: {total_score} / 100</h1>", unsafe_allow_html=True)
        
        # ==========================================
        # ปุ่มออกคะแนน (แสดงผลลงตารางบนเว็บ)
        # ==========================================
        if st.button("🎯 ออกคะแนน", use_container_width=True):
            if student_name == "":
                st.warning("⚠️ กรุณากรอกชื่อผู้เข้าแข่งขันก่อนออกคะแนน")
            else:
                new_row = {
                    "ลำดับ": order_no, 
                    "โรงเรียน": school, 
                    "สังกัด": affiliation, 
                    "จังหวัด": province, 
                    "ทีม": team, 
                    "ชื่อ-สกุล ผู้เข้าแข่งขัน": student_name,
                    "การจัดเตรียมแผ่นงานและบันทึกข้อมูล\n(1 คะแนน)": c1,
                    'การคำนวณและประมวลผลข้อมูลใน Sheet "Student_Scores" (24 คะแนน)': c2,
                    'การเชื่อมโยงและจัดเรียงข้อมูล\nใน Sheet "Student_Scores2"\n(7 คะแนน)': c3,
                    'การรายงานตารางสรุปข้อมูลใน Sheet "Report_Table"\n(12 คะแนน)': c4,
                    'การสร้างแผ่นงานประมวลผลเกรด Sheet "Grade_Summary"\n(2 คะแนน)': c5,
                    'การคำนวณเกรดและสถานะประเมินใน Sheet "Grade_Summary"\n(18 คะแนน)': c6,
                    'การเพิ่มแผ่นงานแดชบอร์ด Sheet "Report_Dashboard"\n(2 คะแนน)': c7,
                    'การสร้างแดชบอร์ดสรุปผลและแผนภูมิใน Sheet "Report_Dashboard"\n(24 คะแนน)': c8,
                    'การตั้งค่าหน้ากระดาษและการพิมพ์รูปแบบ PDF\n(10 คะแนน)': c9,
                    "รวม\n(100 คะแนน)": total_score
                }
                st.session_state['results_df'] = pd.concat([st.session_state['results_df'], pd.DataFrame([new_row])], ignore_index=True)
                st.success(f"✅ เพิ่มคะแนนของ {student_name} ลงตารางแสดงผลด้านล่างเรียบร้อยแล้ว!")

    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาดในการโหลดไฟล์ รายละเอียด: {e}")

# ==========================================
# แสดงตารางสรุปผลบนหน้าเว็บ
# ==========================================
if not st.session_state['results_df'].empty:
    st.divider()
    st.header("📑 ตารางสรุปคะแนน (แสดงผลบนเว็บ)")
    st.dataframe(st.session_state['results_df'], use_container_width=True)
    
    if st.button("🗑️ ล้างข้อมูลตารางทั้งหมด"):
        st.session_state['results_df'] = st.session_state['results_df'].iloc[0:0] 
        st.rerun()
