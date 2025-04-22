# ---------------------  extract-excel.py  (FULL)  ---------------------
import os, base64, json, re, io, requests, pandas as pd, streamlit as st
from PIL import Image

# -------------------------------------------------------------------- #
# 1)  Utilities                                                        #
# -------------------------------------------------------------------- #
def encode_image(image_file):
    """Return Base‑64 from an uploaded image file (Streamlit)."""
    return base64.b64encode(image_file.getvalue()).decode('utf-8')


def _kv_from_text(txt: str) -> float | None:
    """
    Return the highest *kV* value found in a string, or None.
    Accept only numbers that are explicitly followed by kV to avoid BIL 900, 1050 kV etc.
    """
    kv_matches = re.findall(r'(\d+\.?\d*)\s*(?:k[Vv])', txt)
    if kv_matches:
        return max(float(v) for v in kv_matches)
    return None


# -------------------------------------------------------------------- #
# 2)  Prompt generator (เหมือนเดิม)                                    #
# -------------------------------------------------------------------- #
def generate_prompt_from_excel(excel_file):
    """
    Read an Excel list of attributes + (optionally) units, then build a Thai prompt
    telling Gemini to extract those exact fields in JSON.
    """
    # ----- read Excel whether it has a header row or not -----
    try:
        df = pd.read_excel(excel_file)
        first_col = df.columns[0]
        is_numeric_header = isinstance(first_col, (int, float))
        if is_numeric_header:
            excel_file.seek(0)
            df = pd.read_excel(excel_file, header=None)
            df.columns = ['attribute_name'] + [f'col_{i}' for i in range(1, len(df.columns))]
            st.info("ตรวจพบไฟล์ไม่มีหัวคอลัมน์ – กำลังปรับให้อ่านได้")
    except Exception as e:
        excel_file.seek(0)
        df = pd.read_excel(excel_file, header=None)
        df.columns = ['attribute_name'] + [f'col_{i}' for i in range(1, len(df.columns))]
        st.warning(f"อ่านไฟล์แบบมีหัวคอลัมน์ไม่ได้: {e}  → ใช้โหมดไม่มีหัว")

    st.write("คอลัมน์ที่พบ:", list(df.columns))

    attribute_col = 'attribute_name'
    if attribute_col not in df.columns:
        for c in ['attribute_name', 'attribute', 'name', 'attributes',
                  'Attribute', 'ATTRIBUTE', 'field', 'Field', 'FIELD']:
            if c in df.columns:
                attribute_col = c; break
        if attribute_col not in df.columns:
            attribute_col = df.columns[0]
            st.warning(f"ไม่พบคอลัมน์ชื่อ attribute ที่รู้จัก – ใช้คอลัมน์ '{attribute_col}' แทน")

    unit_col = None
    for c in ['unit_of_measure', 'unit', 'Unit', 'UNIT', 'uom', 'UOM',
              'unit of measure', 'Unit of Measure']:
        if c in df.columns:
            unit_col = c; break

    if unit_col is None and len(df.columns) > 1:
        potential = df.columns[1]
        sample = df[potential].dropna().astype(str).tolist()[:10]
        if any(any(k in v for k in ['kg', 'V', 'A', 'kV', 'kVA', 'C', '°C',
                                    'mm', 'cm', 'm', '%']) for v in sample):
            unit_col = potential
            st.info(f"ตรวจพบคอลัมน์ '{potential}' อาจเป็นหน่วยวัด")

    prompt_parts = ["""กรุณาสกัดข้อมูลทั้งหมดจากรูปภาพนี้และแสดงผลในรูปแบบ JSON ที่มีโครงสร้างชัดเจน โดยใช้ key เป็นภาษาอังกฤษและ value เป็นข้อมูลที่พบ
ให้ return ค่า attributes กลับด้วยค่า attribute เท่านั้นห้าม return เป็น index เด็ดขาดและไม่ต้องเอาค่า index มาด้วย
โดยเอาเฉพาะ attributes ดังต่อไปนี้\n"""]

    for i, row in df.iterrows():
        attr = str(row[attribute_col]).strip()
        if pd.isna(attr) or attr == '':
            continue
        if unit_col and unit_col in df.columns and pd.notna(row[unit_col]) and str(row[unit_col]).strip():
            prompt_parts.append(f"{i+1}: {attr} [{row[unit_col]}]")
        else:
            prompt_parts.append(f"{i+1}: {attr}")

    prompt_parts.append("\nหากไม่พบข้อมูลสำหรับ attribute ใด ให้ใส่ค่า - แทน ไม่ต้องเดาค่า และให้รวม attribute และหน่วยวัดไว้ในค่าที่ส่งกลับด้วย")
    return "\n".join(prompt_parts)


# -------------------------------------------------------------------- #
# 3)  Gemini API                                                       #
# -------------------------------------------------------------------- #
def extract_data_from_image(api_key, image_b64, prompt_text):
    url = f"https://generativelanguage.googleapis.com/v1/models/gemini-2.0-flash:generateContent?key={api_key}"
    payload = {
        "contents": [{
            "parts": [
                {"text": prompt_text},
                {"inline_data": {"mime_type": "image/jpeg", "data": image_b64}}
            ]
        }],
        "generation_config": {"temperature": 0.2, "top_p": 0.85, "max_output_tokens": 9000}
    }
    r = requests.post(url, headers={"Content-Type": "application/json"}, data=json.dumps(payload))
    if r.status_code == 200:
        data = r.json()
        if data.get('candidates'):
            return data['candidates'][0]['content']['parts'][0]['text']
        return "ไม่พบข้อมูลในการตอบกลับ"
    return f"เกิดข้อผิดพลาด: {r.status_code} - {r.text}"


# -------------------------------------------------------------------- #
# 4)  POWTR‑CODE generator                                             #
# -------------------------------------------------------------------- #
def generate_powtr_code(extracted: dict) -> str:
    """Return POWTR‑CODE string from extracted dict."""
    try:
        # ----- phase ----------------------------------------------------
        phase = '3'
        for v in extracted.values():
            up = str(v).upper()
            if any(t in up for t in ('1PH', '1-PH', 'SINGLE')):
                phase = '1'; break

        # ----- voltage --------------------------------------------------
        high_kv = None
        for k, v in extracted.items():
            if 'VOLT' in k.upper():          # covers VOLTAGE / VOLATGE
                kv = _kv_from_text(str(v))
                if kv is not None:
                    high_kv = kv if high_kv is None else max(high_kv, kv)

        if high_kv is None:
            voltage_char = '-'
        elif high_kv > 765:
            return 'POWTR-3-OO'              # over‑limit rule
        elif high_kv >= 345:
            voltage_char = 'E'
        elif high_kv >= 100:
            voltage_char = 'H'
        elif high_kv >= 1:
            voltage_char = 'M'
        else:
            voltage_char = 'L'

        # ----- type (oil / dry) ----------------------------------------
        type_char = '-'
        for v in extracted.values():
            up = str(v).upper()
            if 'DRY' in up:
                type_char = 'D'; break
            if 'OIL' in up:
                type_char = 'O'

        # ----- tap changer ---------------------------------------------
        tap_char = 'F'   # default Off‑load
        for v in extracted.values():
            up = str(v).upper()
            if any(t in up for t in ('ON‑LOAD', 'ON-LOAD', 'OLTC')):
                tap_char = 'O'; break
            if any(t in up for t in ('OFF‑LOAD', 'OFF-LOAD', 'FLTC', 'OCTC')):
                tap_char = 'F'

        return f'POWTR-{phase}{voltage_char}{type_char}{tap_char}'
    except Exception:
        return 'ไม่สามารถระบุได้'


def calculate_and_add_powtr_codes(results: list[dict]) -> list[dict]:
    for r in results:
        if isinstance(r.get('extracted_data'), dict) and not any(
                k in r['extracted_data'] for k in ('raw_text', 'error')):
            r['extracted_data']['POWTR_CODE'] = generate_powtr_code(r['extracted_data'])
    return results


# -------------------------------------------------------------------- #
# 5)  Streamlit UI                                                     #
# -------------------------------------------------------------------- #
st.title("ระบบสกัดข้อมูลจากรูปภาพ")

# NOTE: ตามคำขอ – เก็บ API KEY ไว้ในโค้ด
api_key = "AIzaSyDb8iBV1EWqLvjheG_44gh3vQHfpmYGOCI"

tab1, tab2 = st.tabs(["ใช้ไฟล์ Excel", "ใช้ attributes ที่กำหนดไว้แล้ว"])
with tab1:
    excel_file = st.file_uploader("เลือกไฟล์ Excel ที่มี attributes", ["xlsx", "xls"], key="excel_up")
    if excel_file:
        st.dataframe(pd.read_excel(excel_file).head())

with tab2:
    use_default = st.checkbox("ใช้ attributes ที่กำหนดไว้แล้ว", True)
    if use_default:
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

uploaded_files = st.file_uploader("อัปโหลดรูปภาพ (หลายรูปได้)", ["jpg", "png", "jpeg"],
                                  accept_multiple_files=True, key="img_up")

if st.button("ประมวลผล") and api_key and uploaded_files:
    prompt = default_prompt
    if 'excel_file' in locals() and excel_file:
        prompt = generate_prompt_from_excel(excel_file)

    st.expander("Prompt ที่ใช้").write(prompt)

    results, bar, status = [], st.progress(0), st.empty()
    for i, f in enumerate(uploaded_files, 1):
        bar.progress(i / len(uploaded_files))
        status.write(f"กำลังประมวลผล {i}/{len(uploaded_files)} – {f.name}")
        img_b64 = encode_image(f)
        resp = extract_data_from_image(api_key, img_b64, prompt)

        try:
            js = json.loads(resp[resp.find('{'):resp.rfind('}') + 1])
        except Exception:
            js = {"raw_text": resp}

        results.append({"file_name": f.name, "extracted_data": js})

    results = calculate_and_add_powtr_codes(results)

    # -------- show codes ----------
    st.subheader("POWTR‑CODE ที่สร้างได้")
    for r in results:
        st.write(r['extracted_data'].get('POWTR_CODE', '—'))

    # -------- table ---------------
    rows = []
    for r in results:
        data = r['extracted_data']
        code = data.get('POWTR_CODE', '')
        if 'raw_text' in data or 'error' in data:
            rows.append({"POWTR_CODE": code, "ATTRIBUTE": "Error",
                         "VALUE": data.get('error', data.get('raw_text', ''))})
        else:
            for k, v in data.items():
                if k != 'POWTR_CODE':
                    rows.append({"POWTR_CODE": code, "ATTRIBUTE": k, "VALUE": v})
    df = pd.DataFrame(rows)
    st.subheader("ตัวอย่างข้อมูลที่สกัดได้ (แบบแถว)")
    st.dataframe(df)

    # -------- download ------------
    buff = io.BytesIO()
    with pd.ExcelWriter(buff, engine='openpyxl') as w:
        df.to_excel(w, index=False)
    buff.seek(0)
    st.download_button("ดาวน์โหลดไฟล์ Excel", buff,
                       "extracted_data_sorted.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
# -------------------------------------------------------------------- #
