# --------------------  extract-excel.py  (FULL FILE – 26 Apr 2025)  --------------------
import os, base64, json, re, io, imghdr, requests, pandas as pd, streamlit as st
from PIL import Image

# --------------------------------------------------------------------------- #
# 1)  Utilities                                                               #
# --------------------------------------------------------------------------- #
def encode_image(file) -> tuple[str, str]:
    """
    Convert an uploaded image file to (base64‑string, mime‑type).
    Detects PNG/JPEG so we can set the correct mimeType for Gemini.
    """
    raw = file.getvalue()
    kind = imghdr.what(None, raw) or 'jpeg'          # default jpeg
    mime = f"image/{'jpg' if kind == 'jpeg' else kind}"
    return base64.b64encode(raw).decode('utf-8'), mime


def _kv_from_text(txt: str) -> float | None:
    """
    Return the highest *system‑voltage* in kV that appears in *txt*.

    • Accept “… 525000 V”, “34.5 kV”, “220kV”, etc.  
    • **Ignore**:
        – Strings that are in kVA, VA, kA, A, … (negative look‑ahead)  
        – Any number within 5 chars of “BIL” or “IMPULSE”.
    """
    t = txt.upper()
    best = None
    #     number          k?V          <--- no letter allowed after V
    pat = re.compile(r'(\d+(?:\.\d+)?)(?:\s*)([K]?V)(?![A-Z])', re.I)
    for m in pat.finditer(t):
        s = m.start()
        # skip BIL / IMPULSE vicinity
        if any(tok in t[max(0, s-5):s+8] for tok in ('BIL', 'IMPULSE')):
            continue
        num = float(m.group(1))
        unit = m.group(2).upper()        # 'V' or 'KV'
        kv = num if unit == 'KV' else num / 1000
        best = kv if best is None else max(best, kv)
    return best


# --------------------------------------------------------------------------- #
# 2)  Prompt generator  (โค้ดเดิมของคุณ วางตรงนี้ได้เลย – ไม่แก้)           #
# --------------------------------------------------------------------------- #
def generate_prompt_from_excel(excel_file):
    # … (ใช้เวอร์ชันเดิมของคุณ ไม่มีการเปลี่ยนแปลง) …
    pass


# --------------------------------------------------------------------------- #
# 3)  Gemini API call                                                         #
# --------------------------------------------------------------------------- #
def extract_data_from_image(api_key: str, img_b64: str, mime: str, prompt: str) -> str:
    url = f"https://generativelanguage.googleapis.com/v1/models/gemini-2.0-flash:generateContent?key={api_key}"
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
        # ---- phase ------------------------------------------------------
        phase = '3'
        if any(any(t in str(v).upper() for t in ('1PH', '1-PH', 'SINGLE'))
               for v in extracted.values()):
            phase = '1'

        # ---- voltage (max kV) ------------------------------------------
        high_kv = None
        for k, v in extracted.items():
            if any(t in k.upper() for t in ('VOLT', 'HV', 'LV', 'RATED', 'SYSTEM')):
                kv = _kv_from_text(str(v))
                if kv is not None:
                    high_kv = kv if high_kv is None else max(high_kv, kv)

        if high_kv is None:
            v_char = '-'
        elif high_kv > 765:
            return 'POWTR-3-OO'                       # over‑limit rule
        elif high_kv >= 345:
            v_char = 'E'
        elif high_kv >= 100:
            v_char = 'H'
        elif high_kv >= 1:
            v_char = 'M'
        else:
            v_char = 'L'

        # ---- type (oil / dry) ------------------------------------------
        t_char = '-'
        for v in extracted.values():
            u = str(v).upper()
            if 'DRY' in u:
                t_char = 'D'; break
            if 'OIL' in u:
                t_char = 'O'

        # ---- tap changer -----------------------------------------------
        tap = 'F'
        for v in extracted.values():
            u = str(v).upper()
            if any(x in u for x in ('ON‑LOAD', 'ON-LOAD', 'OLTC')):
                tap = 'O'; break
            if any(x in u for x in ('OFF‑LOAD', 'OFF-LOAD', 'FLTC', 'OCTC')):
                tap = 'F'

        return f'POWTR-{phase}{v_char}{t_char}{tap}'
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
st.title("ระบบสกัดข้อมูลหม้อแปลง + POWTR‑CODE (ปรับไม่จับ kVA)")

API_KEY = "AIzaSyDb8iBV1EWqLvjheG_44gh3vQHfpmYGOCI"

tab1, tab2 = st.tabs(["ใช้ไฟล์ Excel", "ใช้ attributes ที่กำหนดไว้แล้ว"])
with tab1:
    excel_f = st.file_uploader("เลือกไฟล์ Excel attributes", ["xlsx", "xls"])
    if excel_f:
        st.dataframe(pd.read_excel(excel_f).head())
with tab2:
    use_default = st.checkbox("ใช้ attributes ที่กำหนดไว้แล้ว", True)
    if use_default:
        default_prompt = """กรุณาสกัดข้อมูลทั้งหมดจากรูปภาพนี้และแสดงผลในรูปแบบ JSON ที่มีโครงสร้างชัดเจน ..."""

imgs = st.file_uploader("อัปโหลดรูปภาพ (หลายไฟล์)", ["jpg", "png", "jpeg"],
                        accept_multiple_files=True)

if st.button("ประมวลผล") and API_KEY and imgs:
    prompt = default_prompt
    if 'excel_f' in locals() and excel_f:
        prompt = generate_prompt_from_excel(excel_f)
    st.expander("Prompt ที่ใช้").write(prompt)

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

    # --- show codes -----------------------------------------------------
    st.subheader("POWTR‑CODE ที่สร้างได้")
    for r in results:
        st.write(r['extracted_data'].get('POWTR_CODE', '—'))

    # --- flatten to rows -------------------------------------------------
    rows = []
    for r in results:
        data = r['extracted_data']
        c = data.get('POWTR_CODE', '')
        if 'error' in data or 'raw_text' in data:
            rows.append({"POWTR_CODE": c, "ATTRIBUTE": "Error",
                         "VALUE": data.get('error', data.get('raw_text', ''))})
        else:
            for k, v in data.items():
                if k != 'POWTR_CODE':
                    rows.append({"POWTR_CODE": c, "ATTRIBUTE": k, "VALUE": v})
    df = pd.DataFrame(rows)
    st.subheader("ตัวอย่างข้อมูลที่สกัดได้ (แบบแถว)")
    st.dataframe(df)

    # --- download -------------------------------------------------------
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as w:
        df.to_excel(w, index=False)
    buf.seek(0)
    st.download_button("ดาวน์โหลดไฟล์ Excel", buf,
                       "extracted_data_sorted.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
# --------------------------------------------------------------------------- #
