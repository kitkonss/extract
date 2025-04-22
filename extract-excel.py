# ---------------------  extract-excel.py  (full file)  ---------------------
import os, base64, json, re, io, requests, pandas as pd, streamlit as st
from PIL import Image

# ---------- 1. utilities ----------
def encode_image(image_file):
    return base64.b64encode(image_file.getvalue()).decode('utf-8')

def _kv_from_text(txt: str) -> float | None:
    """
    Return highest kV value found in a string, or None.
    Only numbers *followed by* kV/kV. are accepted to avoid BIL 900 etc.
    """
    kv_matches = re.findall(r'(\d+\.?\d*)\s*(?:k[Vv])', txt)
    if kv_matches:
        return max(float(v) for v in kv_matches)
    return None

# ---------- 2. prompt helper (เหมือนเดิม ย่อไว้) ----------
def generate_prompt_from_excel(excel_file):
    # ... (ฟังก์ชันเดิมไม่เปลี่ยน ใส่ไว้เหมือนในเวอร์ชันก่อน) ...
    # [ตัดทอนเพื่อความกระชับ – คงโค้ดเดิมไว้ทั้งหมด]
    pass  # ← คัดลอกฟังก์ชันเดิมมาตรงนี้

# ---------- 3. Gemini call ----------
def extract_data_from_image(api_key, image_data, prompt_text):
    url = f"https://generativelanguage.googleapis.com/v1/models/gemini-2.0-flash:generateContent?key={api_key}"
    payload = {
        "contents": [{
            "parts": [
                {"text": prompt_text},
                {"inline_data": {"mime_type": "image/jpeg", "data": image_data}}
            ]
        }],
        "generation_config": {"temperature": 0.2, "top_p": 0.85, "max_output_tokens": 9000}
    }
    r = requests.post(url, headers={"Content-Type": "application/json"}, data=json.dumps(payload))
    if r.status_code == 200:
        data = r.json()
        return data['candidates'][0]['content']['parts'][0]['text'] \
               if data.get('candidates') else "ไม่พบข้อมูลในการตอบกลับ"
    return f"เกิดข้อผิดพลาด: {r.status_code} - {r.text}"

# ---------- 4. POWTR‑code ----------
def generate_powtr_code(extracted: dict) -> str:
    """
    POWTR-{phase}{voltage}{type}{tap}
    phase : 3 / 1 / 0
    voltage: E H M L  (‘-’ unknown, over‑limit => special code)
    type   : O / D / -
    tap    : O / F / -
    """
    try:
        # (a) phase ---------------------------------------------------------
        phase = '3'
        for v in extracted.values():
            up = str(v).upper()
            if any(x in up for x in ('1PH', '1-PH', 'SINGLE')):
                phase = '1'; break

        # (b) voltage -------------------------------------------------------
        high_kv = None
        for k, v in extracted.items():
            if 'VOLT' in k.upper():                       # ครอบคลุม VOLTAGE / VOLATGE
                kv = _kv_from_text(str(v))
                if kv is not None:
                    high_kv = kv if high_kv is None else max(high_kv, kv)
        if high_kv is None:
            voltage = '-'
        elif high_kv > 765:
            return 'POWTR-3-OO'                           # rule for > 765 kV
        elif high_kv >= 345:
            voltage = 'E'
        elif high_kv >= 100:
            voltage = 'H'
        elif high_kv >= 1:
            voltage = 'M'
        else:
            voltage = 'L'

        # (c) type ----------------------------------------------------------
        type_char = '-'
        for v in extracted.values():
            up = str(v).upper()
            if 'DRY' in up:
                type_char = 'D'; break
            if 'OIL' in up:
                type_char = 'O'
        # (d) tap changer ---------------------------------------------------
        tap = 'F'    # default Off‑load
        for v in extracted.values():
            up = str(v).upper()
            if any(x in up for x in ('ON‑LOAD', 'ON-LOAD', 'OLTC')):
                tap = 'O'; break
            if any(x in up for x in ('OFF‑LOAD', 'OFF-LOAD', 'FLTC', 'OCTC')):
                tap = 'F'
        return f'POWTR-{phase}{voltage}{type_char}{tap}'
    except Exception:
        return 'ไม่สามารถระบุได้'

def calculate_and_add_powtr_codes(results: list[dict]) -> list[dict]:
    for r in results:
        if isinstance(r.get('extracted_data'), dict) and \
           not any(k in r['extracted_data'] for k in ('raw_text', 'error')):
            r['extracted_data']['POWTR_CODE'] = generate_powtr_code(r['extracted_data'])
    return results

# ---------- 5. Streamlit UI (ส่วนอื่นเหมือนเดิม) ----------
st.title("ระบบสกัดข้อมูลจากรูปภาพ")

api_key = "YOUR_GEMINI_API_KEY"      # <-- ใส่คีย์ของคุณ

tab1, tab2 = st.tabs(["ใช้ไฟล์ Excel", "ใช้ attributes ที่กำหนดไว้แล้ว"])
with tab1:
    excel_file = st.file_uploader("เลือกไฟล์ Excel ที่มี attributes",
                                  type=["xlsx", "xls"], key="excel_up")
    if excel_file:
        st.dataframe(pd.read_excel(excel_file).head())
with tab2:
    use_default = st.checkbox("ใช้ attributes ที่กำหนดไว้แล้ว", True)
    if use_default:
        default_prompt = """กรุณาสกัดข้อมูล... (prompt เดิม)"""
uploaded_files = st.file_uploader("อัปโหลดรูปภาพ (หลายรูปได้)",
                                  ["jpg", "png", "jpeg"], True, key="img_up")

if st.button("ประมวลผล") and api_key and uploaded_files:
    prompt = default_prompt
    if 'excel_file' in locals() and excel_file:
        prompt = generate_prompt_from_excel(excel_file)
    st.expander("Prompt").write(prompt)

    results, bar, status = [], st.progress(0), st.empty()
    for i, f in enumerate(uploaded_files, 1):
        bar.progress(i/len(uploaded_files))
        status.write(f"กำลังประมวลผล {i}/{len(uploaded_files)} – {f.name}")
        img64 = encode_image(f)
        resp = extract_data_from_image(api_key, img64, prompt)

        # parse JSON (เหมือนเดิม)
        try:
            js = json.loads(resp[resp.find('{'):resp.rfind('}')+1])
        except Exception:
            js = {"raw_text": resp}
        results.append({"file_name": f.name, "extracted_data": js})

    results = calculate_and_add_powtr_codes(results)

    # --- แสดง/บันทึกผล ------------------------------
    st.subheader("POWTR‑CODE ที่สร้างได้")
    for r in results:
        st.write(r['extracted_data'].get('POWTR_CODE', '—'))

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

    # download
    buff = io.BytesIO()
    with pd.ExcelWriter(buff, engine='openpyxl') as w:
        df.to_excel(w, index=False)
    buff.seek(0)
    st.download_button("ดาวน์โหลดไฟล์ Excel", buff,
                       "extracted_data_sorted.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
# ---------------------------------------------------------------------------
