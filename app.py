import streamlit as st
import openpyxl

st.set_page_config(page_title="ระบบตรวจข้อสอบ Excel", layout="centered")

st.title("📊 ระบบตรวจข้อสอบ Excel อัตโนมัติ")
st.write("ให้นักศึกษาอัปโหลดไฟล์นามสกุล `.xlsx` เพื่อตรวจให้คะแนนตามเกณฑ์")

uploaded_file = st.file_uploader("อัปโหลดไฟล์ข้อสอบ Excel ของคุณที่นี่", type=["xlsx"])

if uploaded_file is not None:
    st.info("กำลังประมวลผลไฟล์...")
    
    try:
        # โหลดไฟล์ Excel เพื่ออ่านสูตรและการตั้งค่า
        wb = openpyxl.load_workbook(uploaded_file, data_only=False)
        sheet_names = wb.sheetnames
        
        total_score = 0
        max_score = 100 # สมมติคะแนนเต็มตามเกณฑ์
        
        st.subheader("📝 ผลการตรวจประเมิน")
        
        # ---------------------------------------------------------
        # ตรวจข้อ 1: การเปลี่ยนชื่อ Sheet (ตัวอย่างข้อ 1.1)
        # ---------------------------------------------------------
        if "Student_Scores" in sheet_names:
            total_score += 1
            st.success("✅ 1.1 ผ่าน: มีแผ่นงานชื่อ 'Student_Scores' (ได้ 1 คะแนน)")
        else:
            st.error("❌ 1.1 ไม่ผ่าน: ไม่พบแผ่นงานชื่อ 'Student_Scores' (ได้ 0 คะแนน)")
            
        # ---------------------------------------------------------
        # ตรวจข้อ 6.2: การใช้ฟังก์ชัน Nested IF ตัดเกรด
        # ---------------------------------------------------------
        if "Grade_Summary" in sheet_names:
            ws_grade = wb["Grade_Summary"]
            # สมมติให้นักศึกษาทำสูตรที่เซลล์ F4
            cell_f4 = ws_grade['F4']
            
            if cell_f4.data_type == 'f' and str(cell_f4.value).upper().count('IF(') >= 4:
                total_score += 8
                st.success("✅ 6.2 ผ่าน: ใช้ฟังก์ชัน Nested IF ตัดเกรดได้ถูกต้อง (ได้ 8 คะแนน)")
            else:
                st.error("❌ 6.2 ไม่ผ่าน: ไม่พบการใช้ฟังก์ชัน IF ซ้อนกันตามเกณฑ์ (ได้ 0 คะแนน)")
        else:
             st.warning("⚠️ ข้ามการตรวจข้อ 6.2 เนื่องจากไม่พบแผ่นงาน 'Grade_Summary'")

        # ---------------------------------------------------------
        # ตรวจข้อ 8.1.1: กำหนด Margins หน้ากระดาษ
        # ---------------------------------------------------------
        if "Grade_Summary" in sheet_names:
            ws_grade = wb["Grade_Summary"]
            margins = ws_grade.page_margins
            
            # เช็คว่า บน/ล่าง = 1 นิ้ว, ซ้าย/ขวา = 0.5 นิ้ว
            if margins.top == 1.0 and margins.bottom == 1.0 and margins.left == 0.5 and margins.right == 0.5:
                total_score += 2
                st.success("✅ 8.1.1 ผ่าน: ตั้งค่าระยะขอบ (Margins) ถูกต้อง (ได้ 2 คะแนน)")
            else:
                st.error("❌ 8.1.1 ไม่ผ่าน: ตั้งค่าระยะขอบ (Margins) ไม่ถูกต้อง (ได้ 0 คะแนน)")
                
        # ---------------------------------------------------------
        # สรุปคะแนน
        # ---------------------------------------------------------
        st.divider()
        st.metric(label="คะแนนรวมที่ได้", value=f"{total_score} / {max_score}")
        
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการอ่านไฟล์: {e} โปรดตรวจสอบว่าไฟล์ไม่ได้ถูกเข้ารหัสผ่านไว้")
