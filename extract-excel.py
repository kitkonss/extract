# ---------------------  extract-excel.py  (FULL FILE)  ---------------------
import os, base64, json, re, io, requests, pandas as pd, streamlit as st
from PIL import Image

# -------------------------------------------------------------------- #
# 1)  Utilities                                                        #
# -------------------------------------------------------------------- #
def encode_image(image_file):
    """Return Base‑64 from an uploaded Streamlit file‑uploader object."""
    return base64.b64encode(image_file.getvalue()).decode('utf-8')


def _kv_from_text(txt: str) -> float | None:
    """
    Scan *txt* and return the **largest system‑voltage in kV** it contains.

    • Accept “34.5 kV”, “21000 V”, “33 kV (phase‑to‑phase)”, etc.  
    • **Ignore** numbers that appear within 5 characters before/after the
      token “BIL” or “IMPULSE” (to avoid BIL 900 kV, 1050 kV, etc.).
    """
    txt_upper = txt.upper()
    best_kv = None

    pattern = re.compile(r'(\d+\.?\d*)\s*([kK]?[Vv])')
    for m in pattern.finditer(txt):
        num = float(m.group(1))
        unit = m.group(2).lower()  # 'kv' or 'v'
        s = m.start()

        # skip if close to ‘BIL’ or ‘IMPULSE’
        if any(tok in txt_upper[max(0, s-5):s+8] for tok in ('BIL', 'IMPULSE')):
            continue

        kv = num if unit == 'kv' else num / 1000   # convert V → kV
        best_kv = kv if best_kv is None else max(best_kv, kv)

    return best_kv


# -------------------------------------------------------------------- #
# 2)  Prompt generator (เหมือนรุ่นก่อน – ย่อเพื่อประหยัดที่)             #
# -------------------------------------------------------------------- #
def generate_prompt_from_excel(excel_file):
    # ***โค้ดเดิมทุกบรรทัดของคุณวางตรงนี้ได้เลย*** (ไม่มีผลต่อ bug ปัจจุบัน)
    pass


# -------------------------------------------------------------------- #
# 3)  Gemini call                                                      #
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
    if r.status_code == 200 and r.json().get('candidates'):
        return r.json()['candidates'][0]['content']['parts'][0]['text']
    return f"API ERROR {r.status_code}: {r.text}"


# -------------------------------------------------------------------- #
# 4)  POWTR‑CODE generator                                             #
# -------------------------------------------------------------------- #
def generate_powtr_code(extracted: dict) -> str:
    """Return POWTR‑CODE from the extracted JSON dict."""
    try:
        # -------- phase --------------------------------------------------
        phase = '3'
        for v in extracted.values():
            up = str(v).upper()
            if any(x in up for x in ('1PH', '1-PH', 'SINGLE')):
                phase = '1'; break

        # -------- voltage (find highest kV) ------------------------------
        high_kv = None
        for k, v in extracted.items():
            # scan only keys that smell like voltage, level, HV/LV, etc.
            if any(t in k.upper() for t in ('VOLT', 'HV', 'LV', 'RATED', 'SYSTEM')):
                kv = _kv_from_text(str(v))
                if kv is not None:
                    high_kv = kv if high_kv is None else max(high_kv, kv)

        if high_kv is None:
            voltage_char = '-'
        elif high_kv > 765:
            return 'POWTR-3-OO'            # over‑limit rule stays
        elif high_kv >= 345:
            voltage_char = 'E'
        elif high_kv >= 100:
            voltage_char = 'H'
        elif high_kv >= 1:
            voltage_char = 'M'
        else:
            voltage_char = 'L'

        # -------- type (oil / dry) ---------------------------------------
        type_char = '-'
        for v in extracted.values():
            up = str(v).upper()
            if 'DRY' in up:
                type_char = 'D'; break
            if 'OIL' in up:
                type_char = 'O'

        # -------- tap changer --------------------------------------------
        tap_char = 'F'
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
    """Add POWTR_CODE to each result dict in *results*."""
    for r in results:
        if isinstance(r.get('extracted_data'), dict) and not any(
                k in r['extracted_data'] for k in ('raw_text', 'error')):
            r['extracted_data']['POWTR_CODE'] = generate_powtr_code(r['extracted_data'])
    return results


# -------------------------------------------------------------------- #
# 5)  Streamlit UI                                                     #
# -------------------------------------------------------------------- #
st.title("ระบบสกัดข้อมูลจากรูปภาพ")

api_key = "AIzaSyDb8iBV1EWqLvjheG_44gh3vQHfpmYGOCI"   # ← ใช้จริงได้เลย

tab1, tab2 = st.tabs(["ใช้ไฟล์ Excel", "ใช้ attributes ที่กำหนดไว้แล้ว"])
with tab1:
    excel_file = st.file_uploader("เลือกไฟล์ Excel ที่มี attributes", ["xlsx", "xls"])
    if excel_file:
        st.dataframe(pd.read_excel(excel_file).head())

with tab2:
    use_default = st.checkbox("ใช้ attributes ที่กำหนดไว้แล้ว", True)
    if use_default:
        default_prompt = """กรุณาสกัดข้อมูลทั้งหมดจากรูปภาพนี้และแสดงผลในรูปแบบ JSON ที่มีโครงสร้างชัดเจน ... (เหมือนเดิม)"""

uploaded_files = st.file_uploader("อัปโหลดรูปภาพได้หลายไฟล์", ["jpg", "png", "jpeg"],
                                  accept_multiple_files=True)

if st.button("ประมวลผล") and api_key and uploaded_files:
    prompt = default_prompt
    if 'excel_file' in locals() and excel_file:
        prompt = generate_prompt_from_excel(excel_file)
    st.expander("Prompt").write(prompt)

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

    # ---------------- show codes -----------------
    st.subheader("POWTR‑CODE ที่สร้างได้")
    for r in results:
        st.write(r['extracted_data'].get('POWTR_CODE', '—'))

    # ---------------- table ----------------------
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

    # ---------------- download -------------------
    buff = io.BytesIO()
    with pd.ExcelWriter(buff, engine='openpyxl') as w:
        df.to_excel(w, index=False)
    buff.seek(0)
    st.download_button("ดาวน์โหลดไฟล์ Excel", buff,
                       "extracted_data_sorted.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
# -------------------------------------------------------------------- #
