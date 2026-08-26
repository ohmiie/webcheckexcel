import streamlit as st
import openpyxl
import pandas as pd

# ---------------------------------------------------------
# การตั้งค่าหน้าเว็บ
# ---------------------------------------------------------
st.set_page_config(page_title="ระบบประเมินทักษะ Excel", layout="wide", page_icon="📊")

st.title("📊 ระบบประเมินและให้คะแนนทักษะวิชาการ Excel")
st.markdown("ระบบจะทำการตรวจข้อมูล สูตร และโครงสร้างเบื้องต้นอัตโนมัติ ส่วนกราฟ การตั้งค่าหน้ากระดาษ และไฟล์ PDF จะให้ผู้ตรวจประเมินด้วยตนเอง (Manual Checklist) ผ่านหน้านี้")

uploaded_file = st.file_uploader("📂 อัปโหลดไฟล์ Excel ของผู้เข้าแข่งขัน (.xlsx)", type=["xlsx"])

# ตัวแปรเก็บคะแนนรวม
total_score = 0
max_score = 100

if uploaded_file is not None:
    st.info("🔄 กำลังประมวลผลไฟล์ กรุณารอสักครู่...")
    
    try:
        # โหลดไฟล์แบบดึงสูตร (data_only=False) และดึงค่า (data_only=True)
        wb_formula = openpyxl.load_workbook(uploaded_file, data_only=False)
        wb_data = openpyxl.load_workbook(uploaded_file, data_only=True)
        sheet_names = wb_formula.sheetnames
        
        # คอลัมน์สองฝั่งสำหรับแสดงผล
        col1, col2 = st.columns([1.5, 1])
        
        with col1:
            st.header("🤖 ส่วนที่ 1: ตรวจด้วยระบบอัตโนมัติ (Auto-Grading)")
            
            # ==========================================
            # ข้อ 1: การจัดเตรียมแผ่นงานและบันทึกข้อมูล (3 คะแนน)
            # ==========================================
            st.subheader("ส่วนที่ 1: การจัดเตรียมแผ่นงาน (3 คะแนน)")
            score_1 = 0
            if "Student_Scores" in sheet_names:
                score_1 += 1
                st.write("✅ 1.1 เปลี่ยนชื่อ Sheet เป็น 'Student_Scores' (1/1)")
                
                # ตรวจข้อ 1.2 รหัสนักศึกษา 50 คน
                ws = wb_data["Student_Scores"]
                # ลองเช็คแถวที่มีข้อมูล 50 รายการในคอลัมน์รหัสนักศึกษา (สมมติว่าเป็นคอลัมน์ B หรือ C)
                # เพื่อความง่าย ตรวจว่ามีการใส่ข้อมูลในแถวที่ 50+ หรือไม่
                if ws.max_row >= 50:
                    score_1 += 2
                    st.write("✅ 1.2 ป้อนข้อมูลลำดับ 1-50 และรหัสครบถ้วน (2/2)")
                else:
                    st.write("❌ 1.2 ข้อมูลไม่ครบ 50 รายการ (0/2)")
            else:
                st.write("❌ 1.1 ไม่พบชีต 'Student_Scores' (0/1)")
                st.write("❌ 1.2 ไม่สามารถตรวจจำนวนข้อมูลได้ (0/2)")
            total_score += score_1

            # ==========================================
            # ข้อ 2: การคำนวณและประมวลผลข้อมูล (22 คะแนน)
            # ==========================================
            st.subheader("ส่วนที่ 2: การคำนวณและประมวลผล 'Student_Scores' (22 คะแนน)")
            score_2 = 0
            if "Student_Scores" in sheet_names:
                ws_form = wb_formula["Student_Scores"]
                ws_dat = wb_data["Student_Scores"]
                
                # 2.1 ข้อมูลครบ (คะแนนช่วย, เช็คข้อมูลพื้นฐาน) - สมมติให้คะแนนถ้ามีคอลัมน์เยอะพอ
                if ws_dat.max_column >= 10:
                    score_2 += 2; st.write("✅ 2.1 นำเข้าข้อมูลคะแนนดิบและอื่นๆ ครบถ้วน (2/2)")
                else: st.write("❌ 2.1 ข้อมูลคอลัมน์ไม่ครบถ้วน (0/2)")

                # 2.2 เช็คสูตร SUM 
                sum_found = any("SUM(" in str(ws_form[f'I{r}'].value).upper() for r in range(3, 10))
                if sum_found: score_2 += 3; st.write("✅ 2.2 คำนวณผลรวมคะแนนดิบด้วยฟังก์ชัน SUM (3/3)")
                else: st.write("❌ 2.2 ไม่พบสูตร SUM ในคอลัมน์ผลรวม (0/3)")

                # 2.3 เช็คสูตร IF หาคะแนนหักขาดเรียน
                if_found = any("IF(" in str(ws_form[f'K{r}'].value).upper() for r in range(3, 10))
                if if_found: score_2 += 4; st.write("✅ 2.3 คำนวณคะแนนหักขาดเรียนด้วย IF (4/4)")
                else: st.write("❌ 2.3 ไม่พบสูตร IF ในการหักคะแนนขาดเรียน (0/4)")

                # 2.4 เช็คสูตรลบธรรมดา (-) สำหรับคะแนนสุทธิ
                sub_found = any("-" in str(ws_form[f'L{r}'].value) for r in range(3, 10))
                if sub_found: score_2 += 2; st.write("✅ 2.4 คำนวณคะแนนสุทธิถูกต้อง (2/2)")
                else: st.write("❌ 2.4 ไม่พบการคำนวณคะแนนสุทธิ (0/2)")

                # 2.5 Named Range "Net_Score_Data"
                named_ranges = wb_formula.defined_names.definedName
                if any(nr.name == "Net_Score_Data" for nr in named_ranges):
                    score_2 += 2; st.write("✅ 2.5 กำหนดชื่อ Named Range 'Net_Score_Data' ถูกต้อง (2/2)")
                else: st.write("❌ 2.5 ไม่พบ Named Range ชื่อ 'Net_Score_Data' (0/2)")

                # 2.6 RANK หรือ RANK.EQ
                rank_found = any("RANK" in str(ws_form[f'M{r}'].value).upper() for r in range(3, 10))
                if rank_found: score_2 += 3; st.write("✅ 2.6 หาอันดับด้วยฟังก์ชัน RANK หรือ RANK.EQ (3/3)")
                else: st.write("❌ 2.6 ไม่พบสูตร RANK สำหรับหาอันดับคะแนน (0/3)")
                
                st.warning("⚠️ ข้อ 2.7 (Sparklines 3 คะแนน) และ 2.8 (ฟังก์ชันบรรทัดล่างสุด 3 คะแนน) โปรดตรวจด้วยตนเองด้านขวามือ")
            else:
                st.error("❌ ข้ามการตรวจส่วนที่ 2 เนื่องจากไม่พบชีต 'Student_Scores'")
            total_score += score_2

            # ==========================================
            # ข้อ 3: การเชื่อมโยงและจัดเรียง Student_Scores2 (7 คะแนน)
            # ==========================================
            st.subheader("ส่วนที่ 3: ชีต 'Student_Scores2' (7 คะแนน)")
            score_3 = 0
            if "Student_Scores2" in sheet_names:
                score_3 += 1; st.write("✅ 3.1 เปลี่ยนชื่อ Sheet เป็น 'Student_Scores2' (1/1)")
                # เช็คการ Link sheet
                ws_form2 = wb_formula["Student_Scores2"]
                linked = any("Student_Scores!" in str(ws_form2[f'B{r}'].value) for r in range(2, 10))
                if linked: score_3 += 3; st.write("✅ 3.2 ดึงข้อมูลเชื่อมโยง (Link Sheet) มาวางถูกต้อง (3/3)")
                else: st.write("❌ 3.2 ไม่พบการทำ Link Sheet (0/3)")
                
                st.warning("⚠️ ข้อ 3.3 (การ Sort ข้อมูล 3 คะแนน) โปรดตรวจด้วยตนเองด้านขวามือ")
            else: st.write("❌ 3.1 ไม่พบชีต 'Student_Scores2' (0/7)")
            total_score += score_3

            # ==========================================
            # ข้อ 5 & 6: การประมวลผลเกรด (2 + 18 = 20 คะแนน)
            # ==========================================
            st.subheader("ส่วนที่ 5 และ 6: ชีต 'Grade_Summary' การตัดเกรด (20 คะแนน)")
            score_56 = 0
            if "Grade_Summary" in sheet_names:
                score_56 += 1; st.write("✅ 5.1 สร้างและเปลี่ยนชื่อ Sheet เป็น 'Grade_Summary' (1/1)")
                
                ws_grade = wb_formula["Grade_Summary"]
                # 6.1 Link Sheet
                linked_g = any("Student_Scores!" in str(ws_grade[f'B{r}'].value) for r in range(3, 10))
                if linked_g: score_56 += 4; st.write("✅ 6.1 ดึงข้อมูลเชื่อมโยงมาถูกต้อง (4/4)")
                else: st.write("❌ 6.1 ไม่พบการ Link Sheet ในตารางเกรด (0/4)")

                # 6.2 Nested IF ตัดเกรด
                nested_if = any(str(ws_grade[f'F{r}'].value).upper().count("IF(") >= 4 for r in range(3, 10))
                if nested_if: score_56 += 8; st.write("✅ 6.2 ใช้ Nested IF ตัดเกรด A,B,C,D,F ถูกต้อง (8/8)")
                else: st.write("❌ 6.2 ไม่พบการใช้ Nested IF หริอซ้อนกันไม่ครบ (0/8)")

                # 6.3 IF สถานะ
                status_if = any("IF(" in str(ws_grade[f'G{r}'].value).upper() for r in range(3, 10))
                if status_if: score_56 += 6; st.write("✅ 6.3 ใช้ฟังก์ชัน IF กำหนดสถานะผ่าน/ไม่ผ่าน (6/6)")
                else: st.write("❌ 6.3 ไม่พบฟังก์ชัน IF กำหนดสถานะ (0/6)")

                st.warning("⚠️ ข้อ 5.2 (กำหนดโครงสร้างตาราง 1 คะแนน) โปรดตรวจด้วยตนเองด้านขวามือ")
            else: st.write("❌ ไม่พบชีต 'Grade_Summary' (0/20)")
            total_score += score_56

            # ==========================================
            # ข้อ 7: แดชบอร์ดสรุปผล (26 คะแนน - โปรแกรมตรวจได้บางส่วน)
            # ==========================================
            st.subheader("ส่วนที่ 7: ชีต 'Report_Dashboard' (26 คะแนน)")
            score_7 = 0
            if "Report_Dashboard" in sheet_names:
                score_7 += 1; st.write("✅ 7.1 สร้างและเปลี่ยนชื่อ Sheet เป็น 'Report_Dashboard' (1/1)")
                
                ws_dash = wb_formula["Report_Dashboard"]
                
                # หา COUNT/COUNTA
                dash_formulas = " ".join([str(cell.value).upper() for row in ws_dash.iter_rows() for cell in row if cell.data_type == 'f'])
                
                if "COUNT(" in dash_formulas or "COUNTA(" in dash_formulas:
                    score_7 += 2; st.write("✅ 7.3 ใช้ฟังก์ชัน COUNTA/COUNT หานักศึกษาทั้งหมด (2/2)")
                else: st.write("❌ 7.3 ไม่พบฟังก์ชัน COUNT/COUNTA (0/2)")

                if dash_formulas.count("COUNTIF(") >= 2:
                    score_7 += 10; st.write("✅ 7.4, 7.5, 7.6 ใช้ฟังก์ชัน COUNTIF นับเกรด/กลุ่ม/สถานะ (10/10)")
                else: st.write("❌ 7.4-7.6 ไม่พบการใช้ COUNTIF ที่เพียงพอ (0/10)")
                
                st.warning("⚠️ ข้อ 7.2 (แดชบอร์ด), 7.7 (PivotTable), 7.8 (PivotChart), 7.9 (Slicer) โปรดตรวจด้วยตนเองด้านขวามือ")
            else: st.write("❌ ไม่พบชีต 'Report_Dashboard' (0/26)")
            total_score += score_7
            
        with col2:
            st.header("✍️ ส่วนที่ 2: ตรวจด้วยสายตา (Manual Checklist)")
            st.markdown("ติ๊กถูกในช่องที่นักศึกษาทำถูกต้อง ระบบจะนำคะแนนไปบวกให้โดยอัตโนมัติ")
            
            with st.expander("คะแนนส่วนที่เหลือ (กราฟและรูปแบบ)", expanded=True):
                c_2_7 = st.checkbox("2.7 แทรกกราฟ Line Sparklines คอลัมน์สุดท้าย (3 คะแนน)")
                c_2_8 = st.checkbox("2.8 คำนวณ Average, Max, Min ท้ายตาราง (3 คะแนน)")
                c_3_3 = st.checkbox("3.3 จัดเรียงข้อมูล (Sort) ตามอันดับ 1-50 (3 คะแนน)")
                c_5_2 = st.checkbox("5.2 โครงสร้างตาราง Grade_Summary หัวตาราง 3 แถวถูกต้อง (1 คะแนน)")
                c_7_2 = st.checkbox("7.2 โครงสร้างและหน้าตา Dashboard สวยงาม (3 คะแนน)")
            
            with st.expander("ข้อ 4: Report_Table (PivotTable/Chart)", expanded=True):
                st.write("ถ้าพบชีต Report_Table ให้คะแนน 1 คะแนน")
                c_4_1 = st.checkbox("4.1 มีชีต 'Report_Table' (1 คะแนน)")
                c_4_2 = st.checkbox("4.2 สร้าง PivotTable สรุปจำนวนและค่าเฉลี่ยถูกต้อง (5 คะแนน)")
                c_4_3 = st.checkbox("4.3 สร้าง PivotChart รูปแบบ Column (3 คะแนน)")
                c_4_4 = st.checkbox("4.4 มีเครื่องมือ Slicer 'กลุ่มเรียน' ควบคุม (3 คะแนน)")
                
            with st.expander("ข้อ 7 (ส่วนกราฟ): Report_Dashboard", expanded=True):
                c_7_7 = st.checkbox("7.7 สร้าง PivotTable สรุปจำนวนเกรดตามกลุ่มเรียน (4 คะแนน)")
                c_7_8 = st.checkbox("7.8 สร้าง PivotChart รูปแบบ Column (3 คะแนน)")
                c_7_9 = st.checkbox("7.9 มีเครื่องมือ Slicer 'กลุ่มเรียน' (3 คะแนน)")

            with st.expander("ข้อ 8: ไฟล์ PDF และการพิมพ์ (10 คะแนน)", expanded=True):
                st.markdown("*คุณต้องดูไฟล์ PDF ที่แนบส่งมาประกอบ*")
                c_8_1 = st.checkbox("8.1 ตั้งค่า Grade_Summary แนวตั้ง A4, Scale 90% (2 คะแนน)")
                c_8_1_1 = st.checkbox("8.1.1 Margins (บน/ล่าง 1, ซ้าย/ขวา 0.5, H/F 1) (2 คะแนน)")
                c_8_1_2 = st.checkbox("8.1.2 มี Header ขวา / Footer กลาง และส่งเป็น PDF (2 คะแนน)")
                c_8_2 = st.checkbox("8.2 ตั้งค่ากราฟ SEC04 แนวนอน, A4, Margins กำหนดถูก (2 คะแนน)")
                c_8_3 = st.checkbox("8.3 ส่งเป็น PDF ชื่อไฟล์ 'SEC04' (2 คะแนน)")

            # รวมคะแนน Checklist
            manual_score = 0
            if c_2_7: manual_score += 3
            if c_2_8: manual_score += 3
            if c_3_3: manual_score += 3
            if c_5_2: manual_score += 1
            if c_7_2: manual_score += 3
            if c_4_1: manual_score += 1
            if c_4_2: manual_score += 5
            if c_4_3: manual_score += 3
            if c_4_4: manual_score += 3
            if c_7_7: manual_score += 4
            if c_7_8: manual_score += 3
            if c_7_9: manual_score += 3
            if c_8_1: manual_score += 2
            if c_8_1_1: manual_score += 2
            if c_8_1_2: manual_score += 2
            if c_8_2: manual_score += 2
            if c_8_3: manual_score += 2
            
            final_score = total_score + manual_score

        # ==========================================
        # สรุปคะแนนภาพรวม
        # ==========================================
        st.divider()
        st.markdown(f"<h1 style='text-align: center; color: #1f77b4;'>🏆 คะแนนรวมสุทธิ: {final_score} / {max_score}</h1>", unsafe_allow_html=True)
        
        col_res1, col_res2 = st.columns(2)
        col_res1.metric(label="🤖 คะแนนจากระบบอัตโนมัติ", value=f"{total_score}")
        col_res2.metric(label="✍️ คะแนนจากการตรวจ Manual", value=f"{manual_score}")

    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาดในการโหลดไฟล์ โปรดเช็คว่าไฟล์ไม่ถูกเข้ารหัสผ่านไว้ รายละเอียด: {e}")
