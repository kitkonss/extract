# --------------------  extract-excel.py  (FULL FILE – 27 Apr 2025)  --------------------
import os, base64, json, re, io, imghdr, requests, pandas as pd, streamlit as st
from PIL import Image

# --------------------------------------------------------------------------- #
# 1)  Utilities                                                               #
# --------------------------------------------------------------------------- #
def encode_image(file) -> tuple[str, str]:
    """Convert an uploaded image file to (base64‑string, mime‑type) for Gemini."""
    raw = file.getvalue()
    kind = imghdr.what(None, raw) or 'jpeg'
    mime = f"image/{kind}"
    return base64.b64encode(raw).decode('utf-8'), mime


def _kv_from_text(txt: str) -> float | None:
    """
    Return the highest **system voltage** (kV) found in *txt*.

    • Accept “… 525000 V”, “34.5 kV”, “220 kV”, etc.  
    • Ignore values that are clearly power/current ratings (kVA, A, kA, VA).  
    • Skip numbers near “BIL” or “IMPULSE”.  
    • Discard absurdly large values > 1500 kV.
    """
    txt_u = txt.upper()
    best = None

    # split on '/', ',', ';' to isolate “33/30000/309000 kV” cases
    for chunk in re.split(r'[\/,;]', txt_u):
        chunk = chunk.strip()

        # skip chunks that mention KVA / VA / KA / A
        if re.search(r'\bK?VA\b|\bKA\b|\bAMP|\bA\b', chunk):
            continue

        # skip if near BIL / IMPULSE
        if 'BIL' in chunk or 'IMPULSE' in chunk:
            continue

        for m in re.finditer(r'(\d+(?:\.\d+)?)\s*([K]?V)(?![A-Z])', chunk):
            n = float(m.group(1))
            unit = m.group(2).upper()
            kv = n if unit == 'KV' else n / 1000
            if kv > 1500:          # absurdly high → ignore
                continue
            best = kv if best is None else max(best, kv)

    return best


# --------------------------------------------------------------------------- #
# 2)  Prompt generator                    #
# --------------------------------------------------------------------------- #
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
ให้ return ค่า attributes กลับด้วยค่า attribute เท่านั้นห้าม return เป็น index เด็ดขาดและไม่ต้องเอาค่า index มาด้วย ให้ระวังเรื่อง voltage high side หน่วยต้องเป็น V หรือ kV เท่านั้น
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



# --------------------------------------------------------------------------- #
# 3)  Gemini API call                                                         #
# --------------------------------------------------------------------------- #
def extract_data_from_image(api_key: str, img_b64: str, mime: str, prompt: str) -> str:
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-04-17:generateContent?key={api_key}"
    payload = {
        "contents": [{
            "parts": [
                {"text": prompt},
                {"inlineData": {"mimeType": mime, "data": img_b64}}
            ]
        }],
        "generationConfig": {"temperature": 0.2, "topP": 0.85, "maxOutputTokens": 9000}
    }
    r = requests.post(url, headers={"Content-Type": "application/json"}, data=json.dumps(payload))
    if r.ok and r.json().get('candidates'):
        return r.json()['candidates'][0]['content']['parts'][0]['text']
    return f"API ERROR {r.status_code}: {r.text}"


# --------------------------------------------------------------------------- #
# 4)  POWTR‑CODE generator                                                    #
# --------------------------------------------------------------------------- #
def generate_powtr_code(extracted: dict) -> str:
    try:
        # … (ขั้นตอน 1 และ 2 เหมือนเดิม) …

        # 3) Type → default = '-'  (เมื่อตรวจไม่เจอทั้ง DRY และ OIL)
        t_char = '-'
        for v in extracted.values():
            u = str(v).upper()
            if 'DRY' in u:
                t_char = 'D'
                break
            # ตรวจหา oil-based cooling class (OIL, ONAN, OFAF, ...)
            if any(kw in u for kw in ('OIL', 'ONAN', 'OFAF')):
                t_char = 'O'
                break

        # 4) Tap‑changer (เหมือนเดิม)
        tap_char = 'F'
        for v in extracted.values():
            u = str(v).upper()
            if any(x in u for x in ('ON‑LOAD', 'ON-LOAD', 'OLTC')):
                tap_char = 'O'
                break
            if any(x in u for x in ('OFF‑LOAD', 'OFF-LOAD', 'FLTC', 'OCTC')):
                tap_char = 'F'

        return f'POWTR-{phase}{v_char}{t_char}{tap_char}'
    except Exception:
        return 'ไม่สามารถระบุได้'



def add_powtr_codes(results):
    for r in results:
        d = r.get('extracted_data', {})
        if isinstance(d, dict) and not any(k in d for k in ('error', 'raw_text')):
            d['POWTR_CODE'] = generate_powtr_code(d)
    return results


# --------------------------------------------------------------------------- #
# 5)  Streamlit UI                                                            #
# --------------------------------------------------------------------------- #
st.title("ระบบสกัดข้อมูลหม้อแปลง + POWTR‑CODE ")

API_KEY = "AIzaSyDb8iBV1EWqLvjheG_44gh3vQHfpmYGOCI"

tab1, tab2 = st.tabs(["ใช้ไฟล์ Excel", "ใช้ attributes ที่กำหนดไว้แล้ว"])
with tab1:
    excel_f = st.file_uploader("เลือกไฟล์ Excel attributes", ["xlsx", "xls"])
    if excel_f:
        st.dataframe(pd.read_excel(excel_f).head())
with tab2:
    use_def = st.checkbox("ใช้ attributes ที่กำหนดไว้แล้ว", True)
    if use_def:
        default_prompt = """กรุณาสกัดข้อมูลทั้งหมดจากรูปภาพนี้และแสดงผลในรูปแบบ JSON ที่มีโครงสร้างชัดเจน ..."""

imgs = st.file_uploader("อัปโหลดรูปภาพ (หลายไฟล์)", ["jpg", "png", "jpeg"],
                        accept_multiple_files=True)

if st.button("ประมวลผล") and API_KEY and imgs:
    prompt = default_prompt
    if 'excel_f' in locals() and excel_f:
        prompt = generate_prompt_from_excel(excel_f)
    st.expander("Prompt").write(prompt)

    results, bar, status = [], st.progress(0), st.empty()
    for i, f in enumerate(imgs, 1):
        bar.progress(i / len(imgs))
        status.write(f"กำลังประมวลผล {i}/{len(imgs)} – {f.name}")
        b64, mime = encode_image(f)
        resp = extract_data_from_image(API_KEY, b64, mime, prompt)

        try:
            js = json.loads(resp[resp.find('{'):resp.rfind('}') + 1])
        except Exception:
            js = {"error": resp}

        results.append({"file_name": f.name, "extracted_data": js})

    results = add_powtr_codes(results)

    st.subheader("POWTR‑CODE ที่สร้างได้")
    for r in results:
        st.write(r['extracted_data'].get('POWTR_CODE', '—'))

    rows = []
    for r in results:
        d = r['extracted_data']; code = d.get('POWTR_CODE', '')
        if 'error' in d or 'raw_text' in d:
            rows.append({"POWTR_CODE": code, "ATTRIBUTE": "Error",
                         "VALUE": d.get('error', d.get('raw_text', ''))})
        else:
            for k, v in d.items():
                if k != 'POWTR_CODE':
                    rows.append({"POWTR_CODE": code, "ATTRIBUTE": k, "VALUE": v})
    df = pd.DataFrame(rows)
    st.subheader("ตัวอย่างข้อมูลที่สกัดได้ (แบบแถว)")
    st.dataframe(df)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as w:
        df.to_excel(w, index=False)
    buf.seek(0)
    st.download_button("ดาวน์โหลดไฟล์ Excel", buf,
                       "extracted_data_sorted.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
# --------------------------------------------------------------------------- #
