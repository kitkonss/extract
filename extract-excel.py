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
# ฟังก์ชันสำหรับสร้าง prompt จากไฟล์ Excel
def generate_prompt_from_excel(excel_file):
    # ลองอ่านไฟล์ Excel แบบมีและไม่มีหัวคอลัมน์
    try:
        # พยายามอ่านแบบมีหัวคอลัมน์ก่อน
        df = pd.read_excel(excel_file)
        
        # ตรวจสอบว่าคอลัมน์แรกมีค่าเป็นตัวเลขหรือไม่
        first_col = df.columns[0]
        is_numeric_header = isinstance(first_col, (int, float))
        
        if is_numeric_header:
            # ถ้าหัวคอลัมน์เป็นตัวเลข แสดงว่าน่าจะอ่านผิด
            # อ่านใหม่แบบไม่มีหัวคอลัมน์
            excel_file.seek(0)  # ย้อนกลับไปที่จุดเริ่มต้นของไฟล์
            df = pd.read_excel(excel_file, header=None)
            df.columns = ['attribute_name'] + [f'col_{i}' for i in range(1, len(df.columns))]
            st.info("ตรวจพบรูปแบบไฟล์ Excel แบบไม่มีหัวคอลัมน์ กำลังปรับรูปแบบให้อ่านได้")
    except Exception as e:
        # หากเกิดข้อผิดพลาด ลองอ่านแบบไม่มีหัวคอลัมน์
        excel_file.seek(0)
        df = pd.read_excel(excel_file, header=None)
        df.columns = ['attribute_name'] + [f'col_{i}' for i in range(1, len(df.columns))]
        st.warning(f"เกิดข้อผิดพลาดในการอ่านไฟล์แบบมีหัวคอลัมน์: {e} \nกำลังอ่านแบบไม่มีหัวคอลัมน์แทน")
    
    # แสดงคอลัมน์ที่มีในไฟล์ Excel
    st.write("คอลัมน์ที่พบในไฟล์ Excel:", list(df.columns))
    
    # กำหนดชื่อคอลัมน์สำหรับ attribute
    attribute_col = 'attribute_name'
    if attribute_col not in df.columns:
        # ตรวจสอบคอลัมน์ attribute name (มีหลายชื่อที่เป็นไปได้)
        possible_attribute_cols = ['attribute_name', 'attribute', 'name', 'attributes', 'Attribute', 'ATTRIBUTE', 'field', 'Field', 'FIELD']
        for col in possible_attribute_cols:
            if col in df.columns:
                attribute_col = col
                break
        
        # ถ้าไม่พบคอลัมน์ attribute ให้ใช้คอลัมน์แรก
        if attribute_col not in df.columns:
            attribute_col = df.columns[0]
            st.warning(f"ไม่พบคอลัมน์ชื่อแอตทริบิวต์ที่รู้จัก จะใช้คอลัมน์แรก '{attribute_col}' แทน")
    
    # ตรวจสอบคอลัมน์หน่วยวัด (มีหลายชื่อที่เป็นไปได้)
    unit_col = None
    possible_unit_cols = ['unit_of_measure', 'unit', 'Unit', 'UNIT', 'uom', 'UOM', 'unit of measure', 'Unit of Measure']
    for col in possible_unit_cols:
        if col in df.columns:
            unit_col = col
            break
    
    # หากไม่พบคอลัมน์หน่วยวัด ตรวจสอบคอลัมน์ที่ 2 (ในกรณีที่เป็นไฟล์แบบไม่มีหัว)
    if unit_col is None and len(df.columns) > 1:
        potential_unit_col = df.columns[1]
        # ดูว่าคอลัมน์ที่ 2 มีข้อมูลหน่วยวัดหรือไม่ โดยตรวจสอบคำว่า "kg", "V", "A" ฯลฯ
        sample_values = df[potential_unit_col].dropna().astype(str).tolist()[:10]  # สุ่มดู 10 ค่าแรก
        unit_keywords = ['kg', 'V', 'A', 'kV', 'kVA', 'C', '°C', 'mm', 'cm', 'm', '%']
        has_unit = False
        for value in sample_values:
            if any(keyword in value for keyword in unit_keywords):
                has_unit = True
                break
        
        if has_unit:
            unit_col = potential_unit_col
            st.info(f"ตรวจพบข้อมูลที่อาจเป็นหน่วยวัดในคอลัมน์ '{potential_unit_col}'")
    
    if unit_col is None:
        st.info("ไม่พบคอลัมน์หน่วยวัด การสกัดข้อมูลจะไม่มีข้อมูลหน่วยวัด")
    
    # สร้าง prompt text จากข้อมูลใน Excel
    # สร้าง prompt text แบบใหม่ - ส่วนนี้แก้ไข
    prompt_parts = ["""กรุณาสกัดข้อมูลทั้งหมดจากรูปภาพนี้และแสดงผลในรูปแบบ JSON ที่มีโครงสร้างชัดเจน โดยใช้ key เป็นภาษาอังกฤษและ value เป็นข้อมูลที่พบ
                    ให้ return ค่า attributes กลับด้วยค่า attribute เท่านั้นห้าม return เป็น index เด็ดขาดและไม่ต้องเอาค่า index มาด้วย 
                    โดยเอาเฉพาะ attributes ดังต่อไปนี้ 
                    \n
                    """]
    
    # ใช้เลขกำกับตัวเลือกแต่ดึงข้อความจริงจาก attribute column
    for i, row in df.iterrows():
        attribute = str(row[attribute_col]).strip()
        
        # ข้ามแถวที่ไม่มีข้อมูล
        if pd.isna(attribute) or attribute == '':
            continue
            
        # แปลงชื่อ attribute เป็นรูปแบบเหมาะสม (ตัวพิมพ์ใหญ่แบบ snake_case)
        attribute_name = attribute.upper().replace(' ', '_')
        
        if unit_col is not None and unit_col in df.columns:
            unit = row.get(unit_col, '')
            if pd.notna(unit) and str(unit).strip() != '':
                prompt_parts.append(f"{i+1}: {attribute} [{unit}]")
            else:
                prompt_parts.append(f"{i+1}: {attribute}")
        else:
            prompt_parts.append(f"{i+1}: {attribute}")
    
    prompt_parts.append("\nหากไม่พบข้อมูลสำหรับ attribute ใด ให้ใส่ค่า - แทน ไม่ต้องเดาค่า และให้รวม attribute และหน่วยวัดไว้ในค่าที่ส่งกลับด้วย ")
    
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
            "max_output_tokens": 9000  
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
    
# ฟังก์ชันสำหรับสร้างรหัส POWTR-
def generate_powtr_code(extracted_data):
    """
    สร้างรหัส POWTR- ตามเงื่อนไขที่กำหนด:
    - ตัวแรก: 3, 1 หรือ 0 ขึ้นอยู่กับ phase
    - ตัวที่สอง: E, H, M หรือ L ขึ้นอยู่กับ voltage level
    - ตัวที่สาม: O หรือ D ขึ้นอยู่กับ type (oil-immersed หรือ dry-type)
    - ตัวที่สี่: O, F หรือ N ขึ้นอยู่กับ usage tap changer
    """
    try:
        # หาข้อมูลเกี่ยวกับ phase (default เป็น 3)
        phase_char = "3"  # default เป็น 3 phase
        
        # ตรวจสอบจาก Vector Group หรือชื่อ model ถ้ามี
        vector_group = str(extracted_data.get("VECTOR_GROUP", "")).upper()
        model = str(extracted_data.get("MODEL", "")).upper()
        
        # ตรวจสอบว่าเป็น single phase หรือไม่
        if "1PH" in model or "SINGLE" in model or "1-PHASE" in model:
            phase_char = "1"
        elif "1PH" in vector_group or "SINGLE" in vector_group:
            phase_char = "1"
        
        # ตรวจสอบ voltage level จาก HIGH_SIDE_RATED_VOLTAGE
        voltage_char = "-"  # default เป็น "-"
        high_voltage = extracted_data.get("HIGH SIDE NOMINAL SYSTEM VOLTAGE (KV) [kV]", "")
        high_side = extracted_data.get("HIGH SIDE NOMINAL SYSTEM VOLTAGE (KV) [kV]", "")
        
        if isinstance(high_voltage, str):
            import re
            voltage_value = re.findall(r'[\d.]+', high_voltage)
            if voltage_value:
                voltage = float(voltage_value[0])
                
                if "kV" in high_voltage or "kV" in high_side:
                    pass
                elif "V" in high_voltage or "V" in high_side:
                    voltage = voltage / 1000
                
                if voltage >= 345:
                    voltage_char = "E"  # Extra high voltage (345-765 kV)
                elif voltage >= 100:
                    voltage_char = "H"  # High voltage (100-345 kV)
                elif voltage >= 1:
                    voltage_char = "M"  # Medium voltage (1-100 kV)
                else:
                    voltage_char = "L"  # Low voltage (50-1000V)
        
        # ตรวจสอบ type (oil-immersed หรือ dry-type)
        type_char = "-"  # default เป็น "-"
        model = str(extracted_data.get("MODEL", "")).upper()
        standard = str(extracted_data.get("STANDARD", "")).upper()
        oil_onsulation = str(extracted_data.get("OIL INSULATION TYPE", "")).upper()
        
        if "DRY" in model or "DRY" in standard:
            type_char = "D"
        elif "OIL" in model or "OIL" in standard or "OIL" in oil_onsulation:
            type_char = "O"
        
        # ตรวจสอบ tap changer
        tap_char = "-"  # default เป็น "-"
        all_data = ' '.join([str(v) for v in extracted_data.values()]).upper()
        
        if "ON-LOAD TAP-CHANGER" in all_data or "OLTC" in all_data or "WITH ON-LOAD TAP-CHANGER" in all_data:
            tap_char = "O"  # On load tap change
        elif "OFF-LOAD TAP-CHANGER" in all_data or "OCTC" in all_data or "OFFLOAD TAP" in all_data or "FLTC" in all_data:
            tap_char = "F"  # Off Load tap change
        
        # สร้างรหัส POWTR-
        powtr_code = f"POWTR-{phase_char}{voltage_char}{type_char}{tap_char}"
        
        return powtr_code
    except Exception as e:
        return "ไม่สามารถระบุได้"  # กรณีเกิดข้อผิดพลาด ให้ใช้ค่าผลลัพธ์ว่า "ไม่สามารถระบุได้"



# แก้ไขส่วนการเตรียมข้อมูลสำหรับ Excel ในบรรทัดหลังจาก "for result in results:"
# โดยเพิ่มส่วนการคำนวณรหัส POWTR- และเพิ่มลงในข้อมูล

# นำเข้า Jinja2 สำหรับแสดงวิธีการได้มาของรหัส POWTR- (optional)
def calculate_and_add_powtr_codes(results):
    """
    คำนวณและเพิ่มรหัส POWTR- สำหรับแต่ละไฟล์ที่ประมวลผล
    """
    for result in results:
        if isinstance(result["extracted_data"], dict) and "raw_text" not in result["extracted_data"] and "error" not in result["extracted_data"]:
            # สร้างรหัส POWTR- จากข้อมูลที่สกัดได้
            powtr_code = generate_powtr_code(result["extracted_data"])
            
            # เพิ่มรหัส POWTR- ลงในข้อมูลที่สกัดได้
            result["extracted_data"]["POWTR_CODE"] = powtr_code
    
    return results

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


    # เพิ่มรหัส POWTR- ให้กับแต่ละไฟล์
    results = calculate_and_add_powtr_codes(results)
    
    # แสดงรหัส POWTR- ที่สร้างขึ้น
    st.subheader("POWTR-CODE ที่สร้างขึ้น")
    for result in results:
        if isinstance(result["extracted_data"], dict) and "POWTR_CODE" in result["extracted_data"]:
            st.write(f"{result['extracted_data']['POWTR_CODE']}")
    
        
# ---------------------------------------------------------------------------
    excel_data = []

# เพิ่มข้อมูล POWTR_CODE เป็นคอลัมน์แรก
    for result in results:
        if isinstance(result["extracted_data"], dict) and "raw_text" not in result["extracted_data"] and "error" not in result["extracted_data"]:
            # เริ่มด้วย POWTR_CODE ก่อน
            powtr_code = result["extracted_data"].get("POWTR_CODE", "")
        
        # สร้างแถวข้อมูลใหม่ที่ POWTR_CODE มาอยู่ก่อน ATTRIBUTE
            for key, value in result["extracted_data"].items():
                if key != "POWTR_CODE":  # ข้ามข้อมูล POWTR_CODE
                    row_data = {"POWTR_CODE": powtr_code, "ATTRIBUTE": key, "VALUE": value}
                    excel_data.append(row_data)
        else:
        # กรณีมีข้อผิดพลาด
            error_text = result["extracted_data"].get("error", "")
            raw_text = result["extracted_data"].get("raw_text", "")
            excel_data.append({"POWTR_CODE": "", "ATTRIBUTE": "Error", "VALUE": error_text or raw_text})

# สร้าง DataFrame ใหม่โดยเริ่มจากการแสดง POWTR_CODE ก่อน
    df = pd.DataFrame(excel_data)

# กำหนดให้ไม่แสดง `file_name` และ `POWTR_CODE` แรกในแต่ละแถว
    df = df[df['ATTRIBUTE'] != 'file_name']  # ลบแถวที่มี 'file_name'
    df = df[df['ATTRIBUTE'] != 'POWTR_CODE']  # ลบแถวที่มี 'POWTR_CODE' ถ้าต้องการ

# แสดงข้อมูล
    st.subheader("ตัวอย่างข้อมูลที่สกัดได้ (แบบแถว)")
    st.dataframe(df)

# สร้างไฟล์ Excel ที่มีข้อมูลใหม่
    output_file = "extracted_data_sorted.xlsx"
    df.to_excel(output_file, index=False)

# ให้ผู้ใช้ดาวน์โหลดไฟล์
    with open(output_file, "rb") as file:
        st.download_button(
            label="ดาวน์โหลดไฟล์ Excel (แบบแถว)",
            data=file,
            file_name=output_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    # # เพิ่มตัวเลือกในการดาวน์โหลดข้อมูลแบบเดิม (แบบคอลัมน์) ถ้าต้องการ
    # # เตรียมข้อมูลแบบ column-based
    # column_data = []
    # for result in results:
    #     if isinstance(result["extracted_data"], dict) and "raw_text" not in result["extracted_data"] and "error" not in result["extracted_data"]:
    #         row_data = {"file_name": result["file_name"]}
            
    #         # เพิ่มรหัส POWTR_CODE เป็นคอลัมน์แรกหลังจาก file_name
    #         if "POWTR_CODE" in result["extracted_data"]:
    #             row_data["POWTR_CODE"] = result["extracted_data"]["POWTR_CODE"]
            
    #         # เพิ่มข้อมูลอื่นๆ
    #         for key, value in result["extracted_data"].items():
    #             if key != "POWTR_CODE":  # ข้ามข้อมูล POWTR_CODE ที่เพิ่มไปแล้ว
    #                 row_data[key] = value
            
    #         column_data.append(row_data)
    #     else:
    #         error_text = result["extracted_data"].get("error", "")
    #         raw_text = result["extracted_data"].get("raw_text", "")
        
    #         column_data.append({
    #             "file_name": result["file_name"],
    #             "extracted_text": raw_text if not error_text else error_text
    #         })
    
# เหลือแค่หาวิธีการอ่าน tap changer ก็เสร็จละม้าง