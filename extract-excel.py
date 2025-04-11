import os
import base64
import pandas as pd
import streamlit as st
import requests
import json
from PIL import Image
import io

# ฟังก์ชันสำหรับแปลงรูปภาพเป็น base64
def encode_image(image_file):
    return base64.b64encode(image_file.getvalue()).decode('utf-8')

# ฟังก์ชันสำหรับสร้าง prompt จากไฟล์ Excel
def generate_prompt_from_excel(excel_file):
    # อ่านไฟล์ Excel
    df = pd.read_excel(excel_file)
    
    # แสดงคอลัมน์ที่มีในไฟล์ Excel
    st.write("คอลัมน์ที่พบในไฟล์ Excel:", list(df.columns))
    
    # ตรวจสอบคอลัมน์และใช้คอลัมน์ที่เหมาะสม
    attribute_col = None
    unit_col = None
    
    # ตรวจสอบคอลัมน์ attribute name (มีหลายชื่อที่เป็นไปได้)
    possible_attribute_cols = ['attribute_name', 'attribute', 'name', 'attributes', 'Attribute', 'ATTRIBUTE', 'field', 'Field', 'FIELD']
    for col in possible_attribute_cols:
        if col in df.columns:
            attribute_col = col
            break
    
    # ถ้าไม่พบคอลัมน์ attribute ให้ใช้คอลัมน์แรก
    if attribute_col is None:
        attribute_col = df.columns[0]
        st.warning(f"ไม่พบคอลัมน์ชื่อแอตทริบิวต์ที่รู้จัก จะใช้คอลัมน์แรก '{attribute_col}' แทน")
    
    # ตรวจสอบคอลัมน์หน่วยวัด (มีหลายชื่อที่เป็นไปได้)
    possible_unit_cols = ['unit_of_measure', 'unit', 'Unit', 'UNIT', 'uom', 'UOM', 'unit of measure', 'Unit of Measure']
    for col in possible_unit_cols:
        if col in df.columns:
            unit_col = col
            break
    
    if unit_col is None:
        st.info("ไม่พบคอลัมน์หน่วยวัด การสกัดข้อมูลจะไม่มีข้อมูลหน่วยวัด")
    
    # สร้าง prompt text จากข้อมูลใน Excel
    prompt_parts = ["กรุณาสกัดข้อมูลทั้งหมดจากรูปภาพนี้และแสดงผลในรูปแบบ JSON ที่มีโครงสร้างชัดเจน โดยใช้ key เป็นภาษาอังกฤษและ value เป็นข้อมูลที่พบ โดยเอาเฉพาะ attributes ดังต่อไปนี้\n"]
    
    for i, row in df.iterrows():
        attribute = row[attribute_col]
        
        # ข้ามแถวที่ไม่มีข้อมูล
        if pd.isna(attribute) or str(attribute).strip() == '':
            continue
            
        if unit_col is not None and unit_col in df.columns:
            unit = row.get(unit_col, '')
            if pd.notna(unit) and str(unit).strip() != '':
                prompt_parts.append(f"{i+1}: {attribute} [{unit}]")
            else:
                prompt_parts.append(f"{i+1}: {attribute}")
        else:
            prompt_parts.append(f"{i+1}: {attribute}")
    
    prompt_parts.append("\nหากไม่พบข้อมูลสำหรับ attribute ใด ให้ใส่ค่า - แทน ไม่ต้องเดาค่า และให้รวมหน่วยวัดไว้ในค่าที่ส่งกลับด้วย ระวังเรื่องการอ่านค่าซ้ำซ้อนกันด้วย")
    
    return "\n".join(prompt_parts)

# ฟังก์ชันสำหรับเรียกใช้ Gemini API
def extract_data_from_image(api_key, image_data, prompt_text):
    url = f"https://generativelanguage.googleapis.com/v1/models/gemini-2.0-flash:generateContent?key={api_key}"
    
    payload = {
        "contents": [
            {
                "parts": [
                    {"text": prompt_text},
                    {
                        "inline_data": {
                            "mime_type": "image/jpeg", 
                            "data": image_data
                        }
                    }
                ]
            }
        ],
        "generation_config": {
            "temperature": 0.2,
            "top_p": 0.85,
            "max_output_tokens": 8000  # เพิ่มค่าให้สูงขึ้นเพื่อรองรับข้อมูลจำนวนมาก
        }
    }
    
    headers = {
        "Content-Type": "application/json"
    }
    
    response = requests.post(url, headers=headers, data=json.dumps(payload))
    
    if response.status_code == 200:
        response_data = response.json()
        # ดึงข้อมูลจาก response
        if 'candidates' in response_data and len(response_data['candidates']) > 0:
            text_content = response_data['candidates'][0]['content']['parts'][0]['text']
            return text_content
        else:
            return "ไม่พบข้อมูลในการตอบกลับ"
    else:
        return f"เกิดข้อผิดพลาด: {response.status_code} - {response.text}"

# สร้าง UI ด้วย Streamlit
st.title("ระบบสกัดข้อมูลจากรูปภาพ")

# รับ API key
api_key = value="AIzaSyDb8iBV1EWqLvjheG_44gh3vQHfpmYGOCI"

# สร้างแท็บสำหรับเลือกวิธีการกำหนด attributes
tab1, tab2 = st.tabs(["ใช้ไฟล์ Excel", "ใช้ attributes ที่กำหนดไว้แล้ว"])

with tab1:
    # อัปโหลดไฟล์ Excel ที่มีรายการ attributes
    st.subheader("อัปโหลดไฟล์ Excel ที่มีรายการ attributes")
    excel_file = st.file_uploader("เลือกไฟล์ Excel ที่มี attributes", type=["xlsx", "xls"], key="excel_uploader")
    
    if excel_file:
        try:
            df_preview = pd.read_excel(excel_file)
            st.write("ตัวอย่างข้อมูลจากไฟล์ Excel:")
            st.dataframe(df_preview.head())
            
            # เก็บไฟล์ Excel ไว้ใช้งานต่อ
            excel_file.seek(0)
            
            # อธิบายวิธีการใช้งานไฟล์ Excel
            st.info("""
            วิธีการเตรียมไฟล์ Excel:
            1. ไฟล์ควรมีคอลัมน์ที่เก็บชื่อแอตทริบิวต์ (เช่น 'attribute_name', 'attribute', 'name', หรือคอลัมน์แรก)
            2. ถ้าต้องการระบุหน่วยวัด ให้มีคอลัมน์หน่วยวัด (เช่น 'unit_of_measure', 'unit', 'uom')
            3. ระบบจะใช้คอลัมน์แรกเป็นชื่อแอตทริบิวต์หากไม่พบคอลัมน์ที่ตั้งชื่อตามที่คาดหวัง
            """)
            
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการอ่านไฟล์ Excel: {e}")
            excel_file = None

with tab2:
    # กำหนดค่า prompt โดยตรง
    st.subheader("ใช้ attributes ที่กำหนดไว้แล้ว")
    use_default_attributes = st.checkbox("ใช้ attributes ที่กำหนดไว้แล้ว", value=True)
    
    if use_default_attributes:
        st.info("จะใช้รายการ attributes ที่กำหนดไว้แล้วในโปรแกรม")
        # กำหนดค่า prompt โดยตรง - ขยายเป็น attributes ตามที่ต้องการ
        default_prompt = """กรุณาสกัดข้อมูลทั้งหมดจากรูปภาพนี้และแสดงผลในรูปแบบ JSON ที่มีโครงสร้างชัดเจน โดยใช้ key เป็นภาษาอังกฤษและ value เป็นข้อมูลที่พบ โดยเอาเฉพาะ field ดังต่อไปนี้

1: MANUFACTURER
2: MODEL
3: SERIAL_NO
4: STANDARD
5: CAPACITY [kVA]
6: HIGH_SIDE_RATED_VOLTAGE [kV]
7: LOW_SIDE_RATED_VOLTAGE [kV]
8: HIGH_SIDE_RATED_CURRENT [A]
9: LOW_SIDE_RATED_CURRENT [A]
10: IMPEDANCE_VOLTAGE [%]
11: VECTOR_GROUP

หากไม่พบข้อมูลสำหรับ attribute ใด ให้เว้นว่างหรือใส่ค่า null ไม่ต้องเดาค่า และให้รวมหน่วยวัดไว้ในค่าที่ส่งกลับด้วย"""

# อัปโหลดหลายไฟล์
st.subheader("อัปโหลดรูปภาพที่ต้องการสกัดข้อมูล")
uploaded_files = st.file_uploader("อัปโหลดรูปภาพ (เลือกได้หลายรูป โดยการกด ctrl + คลิกทีละรูป)", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="image_uploader")

# เมื่อผู้ใช้กดปุ่มประมวลผล
if st.button("ประมวลผล") and api_key and uploaded_files:
    # กำหนด prompt ตามวิธีการที่เลือก
    if 'excel_file' in locals() and excel_file is not None:
        try:
            prompt_text = generate_prompt_from_excel(excel_file)
            st.info("ใช้ attributes จากไฟล์ Excel")
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการสร้าง prompt จากไฟล์ Excel: {e}")
            if use_default_attributes:
                prompt_text = default_prompt
                st.warning("เกิดข้อผิดพลาด จะใช้ attributes ที่กำหนดไว้แล้วแทน")
            else:
                st.stop()
    elif use_default_attributes:
        prompt_text = default_prompt
        st.info("ใช้ attributes ที่กำหนดไว้แล้ว")
    else:
        st.error("กรุณาเลือกวิธีการกำหนด attributes")
        st.stop()
    
    # แสดง prompt ที่จะใช้
    with st.expander("ดู Prompt ที่จะใช้"):
        st.text(prompt_text)
    
    # แสดง progress bar
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    results = []
    
    for i, file in enumerate(uploaded_files):
        # อัปเดต progress
        progress = (i + 1) / len(uploaded_files)
        progress_bar.progress(progress)
        
        # แสดงสถานะการทำงาน
        status_text.text(f"กำลังประมวลผล {i+1}/{len(uploaded_files)}: {file.name}")
        
        # แสดงรูปภาพที่กำลังประมวลผล
        st.image(file, caption=f"กำลังประมวลผล: {file.name}", width=300)
        
        try:
            # แปลงรูปภาพเป็น base64
            image_data = encode_image(file)
            
            # เรียกใช้ Gemini API
            response = extract_data_from_image(api_key, image_data, prompt_text)
            
            # พยายามแปลง response เป็น JSON
            extracted_data = {}
            try:
                # ค้นหาส่วนที่เป็น JSON ในข้อความตอบกลับ
                json_start = response.find('{')
                json_end = response.rfind('}') + 1
                
                if json_start >= 0 and json_end > json_start:
                    json_str = response[json_start:json_end]
                    extracted_data = json.loads(json_str)
                else:
                    extracted_data = {"raw_text": response}
            except json.JSONDecodeError:
                extracted_data = {"raw_text": response}
            
            # เพิ่มข้อมูลลงในผลลัพธ์
            result_entry = {
                "file_name": file.name,
                "extracted_data": extracted_data if isinstance(extracted_data, dict) else {"raw_text": response}
            }
            
            results.append(result_entry)
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการประมวลผลไฟล์ {file.name}: {e}")
            results.append({
                "file_name": file.name,
                "extracted_data": {"error": str(e)}
            })
    
    # เมื่อประมวลผลเสร็จสิ้น
    status_text.text("ประมวลผลเสร็จสิ้น!")
    
    # จัดเตรียมข้อมูลสำหรับ Excel
    excel_data = []
    
    for result in results:
        if isinstance(result["extracted_data"], dict) and "raw_text" not in result["extracted_data"] and "error" not in result["extracted_data"]:
            # นำข้อมูลจาก JSON มาขยายออก
            row_data = {"file_name": result["file_name"]}
            row_data.update(result["extracted_data"])
            excel_data.append(row_data)
        else:
            # ใช้ข้อความทั้งหมด
            error_text = result["extracted_data"].get("error", "")
            raw_text = result["extracted_data"].get("raw_text", "")
            
            excel_data.append({
                "file_name": result["file_name"],
                "extracted_text": raw_text if not error_text else error_text
            })
    
    if excel_data:
        # สร้าง DataFrame
        df = pd.DataFrame(excel_data)
        
        # แสดงตัวอย่างข้อมูล
        st.subheader("ตัวอย่างข้อมูลที่สกัดได้")
        st.dataframe(df)
        
        # สร้างไฟล์ Excel
        output_file = "extracted_data.xlsx"
        df.to_excel(output_file, index=False)
        
        # ให้ผู้ใช้ดาวน์โหลดไฟล์
        with open(output_file, "rb") as file:
            st.download_button(
                label="ดาวน์โหลดไฟล์ Excel",
                data=file,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("ไม่มีข้อมูลที่สกัดได้")