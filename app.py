import streamlit as st
import openpyxl

# ---------------------------------------------------------
# การตั้งค่าหน้าเว็บ
# ---------------------------------------------------------
st.set_page_config(page_title="ระบบประเมินทักษะ Excel", layout="wide", page_icon="📊")

st.title("📊 ระบบประเมินและให้คะแนนทักษะวิชาการ Excel")
st.markdown("อัปโหลดไฟล์เพื่อตรวจคะแนนทันที ระบบจะไม่บันทึกหรือจดจำข้อมูลค้างไว้เพื่อความรวดเร็ว")

# ==========================================
# อัปโหลดไฟล์ (แบบเพียวๆ ไม่มีดึงชื่อ)
# ==========================================
uploaded_file = st.file_uploader("📂 อัปโหลดไฟล์ Excel (.xlsx)", type=["xlsx"])

# ตัวแปรเก็บคะแนนแต่ละหมวด
c1, c2, c3, c4, c5, c6, c7, c8, c9 = 0, 0, 0, 0, 0, 0, 0, 0, 0

def get_safe_formula(cell):
    if cell and cell.value:
        return str(cell.value).upper()
    return ""

if uploaded_file is not None:
    st.info("🔄 กำลังประมวลผลไฟล์ กรุณารอสักครู่...")
    
    try:
        # โหลดไฟล์
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
                
                st.success(f"คำนวณ Student_Scores อัตโนมัติได้: {c2}/18 คะแนน")

            # --- หมวด 3: Student_Scores2 (สูตร 4 + Manual 3 = 7 คะแนน) ---
            if "Student_Scores2" in sheet_names:
                c3 += 1
                ws_form2 = wb_formula["Student_Scores2"]
                if any("STUDENT_SCORES!" in get_safe_formula(ws_form2[f'B{r}']) for r in range(2, 10)): c3 += 3
                st.success(f"เชื่อมโยง Student_Scores2 อัตโนมัติได้: {c3}/4 คะแนน")

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
                st.success(f"คำนวณ Dashboard อัตโนมัติได้: {c8}/12 คะแนน")

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
        st.markdown(f"<h3 style='text-align: center;'>หมวด 1: {c1} | หมวด 2: {c2} | หมวด 3: {c3} | หมวด 4: {c4} | หมวด 5: {c5}</h3>", unsafe_allow_html=True)
        st.markdown(f"<h3 style='text-align: center;'>หมวด 6: {c6} | หมวด 7: {c7} | หมวด 8: {c8} | หมวด 9 (PDF): {c9}</h3>", unsafe_allow_html=True)

    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาดในการโหลดไฟล์ รายละเอียด: {e}")
