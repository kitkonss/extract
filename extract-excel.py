# --------------------  extract-excel.py  (FULL FILE – 28 Apr 2025)  --------------------
import os, base64, json, re, io, imghdr, requests, pandas as pd, streamlit as st
from PIL import Image

# --------------------------------------------------------------------------- #
# 1)  Utilities                                                               #
# --------------------------------------------------------------------------- #
def encode_image(file) -> tuple[str, str]:
    raw = file.getvalue()
    kind = imghdr.what(None, raw) or 'jpeg'
    mime = f"image/{'jpg' if kind == 'jpeg' else kind}"
    return base64.b64encode(raw).decode('utf-8'), mime


def _kv_from_text(txt: str) -> float | None:
    """
    Pick the largest *system‑voltage* (kV) from txt.
    Reject chunks with kVA / VA / kA / A, BIL / IMPULSE, or >1500 kV.
    """
    best = None
    txt_u = txt.upper()
    for chunk in re.split(r'[\/,;\n]', txt_u):
        chunk = chunk.strip()
        if re.search(r'\bK?VA\b|\bKA\b|\b[A-Z]?AMP\b|\b[A-Z]?A\b', chunk):
            continue
        if 'BIL' in chunk or 'IMPULSE' in chunk:
            continue
        for m in re.finditer(r'(\d+(?:\.\d+)?)\s*([K]?V)(?![A-Z])', chunk):
            val = float(m.group(1))
            kv = val if m.group(2).upper() == 'KV' else val / 1000
            if kv > 1500:
                continue
            best = kv if best is None else max(best, kv)
    return best


def clean_voltage_fields(data: dict) -> dict:
    """
    Ensure voltage‑related attributes contain only lines with V / kV.
    If value contains kVA or no ‘V’, try to replace with the first V/kV line,
    else set to '-'.
    """
    v_keys = [k for k in data.keys() if re.search(r'VOLT|HV|LV|RATED', k, re.I)]
    for k in v_keys:
        txt = str(data[k])
        # already OK?
        if re.search(r'\b\d', txt) and re.search(r'\b[Vv]\b|KV', txt) and 'KVA' not in txt.upper():
            continue
        # attempt to find proper V/kV substring
        found = re.search(r'(\d[\d\.\s]*[K]?V\b)', txt.upper())
        if found and 'KVA' not in found.group(1):
            data[k] = found.group(1).replace(' ', '')
        else:
            data[k] = '-'
    return data


# --------------------------------------------------------------------------- #
# 2)  Prompt generator                                                        #
# --------------------------------------------------------------------------- #
def generate_prompt_from_excel(excel_file):
    """
    Build a Gemini prompt from an attribute list.
    Extra instruction: **For all voltage‑related fields, capture ONLY the line
    that contains a value followed by V/kV (ignore kVA/kA/A).**
    """
    try:
        df = pd.read_excel(excel_file)
        if isinstance(df.columns[0], (int, float)):
            excel_file.seek(0)
            df = pd.read_excel(excel_file, header=None)
            df.columns = ['attribute_name'] + [f'col_{i}' for i in range(1, len(df.columns))]
            st.info("ไฟล์ไม่มี header – ปรับเรียบร้อย")
    except Exception:
        excel_file.seek(0)
        df = pd.read_excel(excel_file, header=None)
        df.columns = ['attribute_name'] + [f'col_{i}' for i in range(1, len(df.columns))]
        st.warning("อ่าน header ไม่ได้ – ใช้โหมดไม่มี header")

    attr_col = 'attribute_name' if 'attribute_name' in df.columns else df.columns[0]
    unit_col = None
    for c in df.columns:
        if re.fullmatch(r'(unit(_of_measure)?|uom)', str(c), re.I):
            unit_col = c; break

    prompt = ["""กรุณาสกัดข้อมูลทั้งหมดจากรูปภาพนี้และแสดงผลเป็น JSON (key ภาษาอังกฤษ, value ตามที่พบ)
- **ห้าม** ใส่ index
- **Voltage‑related fields** (HV/LV Rated Voltage, Voltage Level ฯลฯ)  
  → ให้เลือกเฉพาะบรรทัดที่มีหน่วย *V* หรือ *kV* (ห้าม kVA / kA / A)
- หากไม่พบข้อมูล ให้ใส่ "-" 
รายการ attributes มีดังนี้:\n"""]

    for i, row in df.iterrows():
        a = str(row[attr_col]).strip()
        if not a or pd.isna(a):
            continue
        if unit_col and pd.notna(row.get(unit_col, '')):
            prompt.append(f"{i+1}: {a} [{row[unit_col]}]")
        else:
            prompt.append(f"{i+1}: {a}")
    prompt.append("\nจงตอบกลับเฉพาะ JSON object")
    return '\n'.join(prompt)


# --------------------------------------------------------------------------- #
# 3)  Gemini call                                                             #
# --------------------------------------------------------------------------- #
def extract_data_from_image(api_key, img_b64, mime, prompt):
    url = f"https://generativelanguage.googleapis.com/v1/models/gemini-2.0-flash:generateContent?key={api_key}"
    body = {
        "contents": [{
            "parts": [
                {"text": prompt},
                {"inlineData": {"mimeType": mime, "data": img_b64}}
            ]
        }],
        "generationConfig": {"temperature": 0.2, "topP": 0.85, "maxOutputTokens": 9000}
    }
    r = requests.post(url, headers={"Content-Type": "application/json"}, data=json.dumps(body))
    if r.ok and r.json().get('candidates'):
        return r.json()['candidates'][0]['content']['parts'][0]['text']
    return f"API ERROR {r.status_code}: {r.text}"


# --------------------------------------------------------------------------- #
# 4)  POWTR‑CODE generator                                                    #
# --------------------------------------------------------------------------- #
def generate_powtr_code(extracted: dict) -> str:
    try:
        phase = '1' if any(re.search(r'1PH|1-PH|SINGLE', str(v), re.I)
                           for v in extracted.values()) else '3'

        high_kv = None
        for k, v in extracted.items():
            if re.search(r'VOLT|HV|LV|RATED', k, re.I):
                kv = _kv_from_text(str(v))
                if kv is not None:
                    high_kv = kv if high_kv is None else max(high_kv, kv)

        if high_kv is None:
            v_char = '-'
        elif high_kv > 765:
            return 'POWTR-3-OO'
        elif high_kv >= 345:
            v_char = 'E'
        elif high_kv >= 100:
            v_char = 'H'
        elif high_kv >= 1:
            v_char = 'M'
        else:
            v_char = 'L'

        t_char = '-'
        for v in extracted.values():
            s = str(v).upper()
            if 'DRY' in s:
                t_char = 'D'; break
            if 'OIL' in s:
                t_char = 'O'

        tap = 'O' if any(re.search(r'ON[-\s]?LOAD|OLTC', str(v), re.I)
                         for v in extracted.values()) else 'F'
        return f'POWTR-{phase}{v_char}{t_char}{tap}'
    except Exception:
        return 'ไม่สามารถระบุได้'


# --------------------------------------------------------------------------- #
# 5)  Streamlit UI                                                            #
# --------------------------------------------------------------------------- #
st.title("Extractor & POWTR‑CODE (v 28 Apr 2025)")

API_KEY = "AIzaSyDb8iBV1EWqLvjheG_44gh3vQHfpmYGOCI"

tab1, tab2 = st.tabs(["ใช้ไฟล์ Excel", "ใช้ attributes ที่กำหนดไว้แล้ว"])
with tab1:
    excel_up = st.file_uploader("Excel attributes", ["xlsx", "xls"])
    if excel_up:
        st.dataframe(pd.read_excel(excel_up).head())
with tab2:
    use_default = st.checkbox("ใช้ attributes ที่กำหนดไว้แล้ว", True)
    if use_default:
        default_prompt = """กรุณาสกัดข้อมูลทั้งหมดจากรูปภาพนี้และแสดงผลเป็น JSON ..."""

images = st.file_uploader("อัปโหลดรูปภาพ", ["jpg", "png", "jpeg"],
                          accept_multiple_files=True)

if st.button("ประมวลผล") and API_KEY and images:
    prompt = default_prompt
    if excel_up:
        prompt = generate_prompt_from_excel(excel_up)
    st.expander("Prompt").write(prompt)

    res, bar, stat = [], st.progress(0), st.empty()
    for i, f in enumerate(images, 1):
        bar.progress(i/len(images))
        stat.write(f"กำลังประมวลผล {i}/{len(images)} – {f.name}")
        b64, mime = encode_image(f)
        raw = extract_data_from_image(API_KEY, b64, mime, prompt)

        try:
            data = json.loads(raw[raw.find('{'): raw.rfind('}')+1])
        except Exception:
            data = {"error": raw}

        if 'error' not in data and 'raw_text' not in data:
            data = clean_voltage_fields(data)
        res.append({"file": f.name, "extracted_data": data})

    # add POWTR codes
    for r in res:
        if isinstance(r['extracted_data'], dict) and 'error' not in r['extracted_data']:
            r['extracted_data']['POWTR_CODE'] = generate_powtr_code(r['extracted_data'])

    st.subheader("POWTR‑CODE")
    for r in res:
        st.write(r['extracted_data'].get('POWTR_CODE', '—'))

    rows = []
    for r in res:
        d = r['extracted_data']
        code = d.get('POWTR_CODE', '')
        if 'error' in d:
            rows.append({"POWTR_CODE": code, "ATTRIBUTE": "Error", "VALUE": d['error']})
        else:
            for k, v in d.items():
                if k != 'POWTR_CODE':
                    rows.append({"POWTR_CODE": code, "ATTRIBUTE": k, "VALUE": v})
    df = pd.DataFrame(rows)
    st.subheader("ผลลัพธ์แบบแถว")
    st.dataframe(df)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as w:
        df.to_excel(w, index=False)
    buf.seek(0)
    st.download_button("ดาวน์โหลด Excel", buf, "extracted_data_sorted.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
# --------------------------------------------------------------------------- #
