
import os, base64, json, re, io, requests, pandas as pd, streamlit as st
from datetime import datetime
from typing import Dict, List, Tuple
from PIL import Image
from openpyxl import load_workbook

# ----------------------------------------------------------------------------
# 0)  CONFIGURATION
# ----------------------------------------------------------------------------
API_KEY = "AIzaSyDb8iBV1EWqLvjheG_44gh3vQHfpmYGOCI"

GITHUB_ATTR_URL = (
    "https://raw.githubusercontent.com/kitkonss/extract/main/ATTRIBUTE.xlsx"
)

# ----------------------------------------------------------------------------
# 1)  UTILITY – download ATTRIBUTE.xlsx from GitHub (mandatory)
# ----------------------------------------------------------------------------

def fetch_attributes_from_github(url: str = GITHUB_ATTR_URL) -> io.BytesIO:
    """Download the ATTRIBUTE Excel from GitHub and return as BytesIO."""
    r = requests.get(url, timeout=15)
    r.raise_for_status()
    buf = io.BytesIO(r.content)
    buf.seek(0)
    return buf

# ----------------------------------------------------------------------------
# 2)  LOW‑LEVEL HELPERS  (encoding, regex utils, splitting value‑unit, etc.)
# ----------------------------------------------------------------------------

def encode_image(file) -> Tuple[str, str]:
    raw = file.getvalue()
    kind = "jpeg"  # treat as JPEG in prompt; Gemini not fussy
    mime = f"image/{kind}"
    return base64.b64encode(raw).decode("utf-8"), mime


def split_value_unit(raw: str) -> Tuple[str, str]:
    if pd.isna(raw):
        return "-", ""
    txt = str(raw).strip()
    m = re.match(r"^([\d\.\-]+)\s*([A-Za-z%°µΩ\\/]*)$", txt)
    if m:
        return m.group(1), m.group(2)
    return txt, ""


# ----------------------------------------------------------------------------
# 3)  PROMPT GENERATION  (from ATTRIBUTE.xlsx)
# ----------------------------------------------------------------------------

def generate_prompt_from_excel(excel_file: io.BytesIO) -> str:
    """Build Thai‑language prompt listing all attributes + units."""
    df = pd.read_excel(excel_file)
    # find attribute column
    attribute_col = None
    for c in df.columns:
        if str(c).lower() in (
            "attribute_name",
            "attribute",
            "name",
            "attributes",
            "field",
        ):
            attribute_col = c
            break
    if attribute_col is None:
        attribute_col = df.columns[0]
    # find optional unit column
    unit_col = None
    for c in df.columns:
        if "unit" in str(c).lower() or "uom" in str(c).lower():
            unit_col = c
            break

    lines = [
        """กรุณาสกัดข้อมูลทั้งหมดจากรูปภาพนี้และแสดงผลในรูปแบบ JSON ที่มีโครงสร้างชัดเจน โดยใช้ key เป็นภาษาอังกฤษและ value เป็นข้อมูลที่พบ
ให้ return ค่า attributes กลับด้วยค่า attribute เท่านั้นห้าม return เป็น index เด็ดขาดและไม่ต้องเอาค่า index มาด้วย ให้ระวังเรื่อง voltage high side หน่วยต้องเป็น V หรือ kV เท่านั้น
โดยเอาเฉพาะ attributes ดังต่อไปนี้"""
    ]
    for i, row in df.iterrows():
        attr = str(row[attribute_col]).strip()
        if not attr or pd.isna(attr):
            continue
        if unit_col and pd.notna(row[unit_col]) and str(row[unit_col]).strip():
            lines.append(f"{i+1}: {attr} [{row[unit_col]}]")
        else:
            lines.append(f"{i+1}: {attr}")
    lines.append(
        "\nหากไม่พบข้อมูลสำหรับ attribute ใด ให้ใส่ค่า - แทน ไม่ต้องเดาค่า และให้รวม attribute และหน่วยวัดไว้ในค่าที่ส่งกลับด้วย"
    )
    return "\n".join(lines)

# ----------------------------------------------------------------------------
# 4)  GEMINI VISION  (call)
# ----------------------------------------------------------------------------

def extract_data_from_image(api_key: str, img_b64: str, mime: str, prompt: str) -> Dict:
    """Return JSON dict (or {'error': ..}) from Gemini Vision API."""
    endpoint = (
        "https://generativelanguage.googleapis.com/v1beta/models/" "gemini-2.5-flash-preview-05-20:generateContent"
    )
    url = f"{endpoint}?key={api_key}"
    payload = {
        "contents": [
            {
                "parts": [
                    {"text": prompt},
                    {"inlineData": {"mimeType": mime, "data": img_b64}},
                ]
            }
        ],
        "generationConfig": {"temperature": 0.2, "topP": 0.80, "maxOutputTokens": 9000},
    }
    r = requests.post(url, headers={"Content-Type": "application/json"}, data=json.dumps(payload))
    if r.ok and r.json().get("candidates"):
        raw = r.json()["candidates"][0]["content"]["parts"][0]["text"]
        try:
            return json.loads(raw[raw.find("{") : raw.rfind("}") + 1])
        except Exception:
            return {"error": raw}
    return {"error": f"API ERROR {r.status_code}: {r.text}"}

# ----------------------------------------------------------------------------
# 5)  POWTR‑CODE LOGIC
# ----------------------------------------------------------------------------

PHASE_DIGIT = "3"  # assumption – three‑phase only

def kv_numbers(text: str):
    """
    คืน list ค่าแรงดันเป็นหน่วย kV จากข้อความ
    - จับ   230 kV   / 230kV
    - จับ   230000 V / 23 000V  แล้ว /1000 → 230 kV
    """
    out = []
    for m in re.finditer(r"(\d+(?:[.,]\d+)?)(\s*[kK]?[vV])", text):
        num = float(m.group(1).replace(",", "."))
        unit = m.group(2).strip().lower()
        if unit == "v":            # เป็นโวลต์ → แปลงเป็น kV
            num /= 1000.0
        out.append(num)
    return out
# -----------------------------------------------------------------------

def voltage_letter(kv_high: float | None) -> str:
    """
    E : 345–765 kV
    H : 100–<345 kV
    M :   1–<100 kV
    L : 0.05–<1 kV
    - : ไม่พบค่า
    """
    if kv_high is None:
        return "-"
    kv = float(kv_high)
    if 345 <= kv <= 765:
        return "E"
    if 100 <= kv < 345:
        return "H"
    if   1 <= kv < 100:
        return "M"
    if 0.05 <= kv < 1:
        return "L"
    return "-"


def type_letter(cooling: str) -> str:
    cooling = str(cooling).upper()
    # crude rule: anything mentioning "DRY" → D else default O
    return "D" if "DRY" in cooling else "O"


def has_oltc(attributes: Dict) -> bool:
    """Detect OLTC from attributes dict."""
    keys = [k for k in attributes.keys() if re.search(r"TAP", k, re.I)]
    for k in keys:
        val = str(attributes[k]).upper()
        if any(x in val for x in ("OLTC", "ON‑LOAD", "DETC", "LOAD TAP")):
            return True
    return False

def _classify_tap(text: str) -> str | None:
    """คืน 'O' - On-load, 'F' - Off-load / Off-circuit, 'N' - No tap, หรือ None ถ้าไม่เจอ"""
    txt = text.upper()

    if re.search(r"\b(OLTC|ON[\s-]?LOAD)\b", txt):
        return "O"

    if re.search(r"\b(FLTC|OFF[\s-]*(LOAD|CIRCUIT))\b", txt):
        return "F"

    if re.search(r"\b(NTC|NO[\s-]?TAP)\b", txt):
        return "N"

    return None


def tap_letter(attrs: dict) -> str:
    """
    เลือกตัวอักษร Tap-changer ตามลำดับความเชื่อมั่นของแหล่งข้อมูล
        1) ช่อง 'USAGE TAP CHANGER'
        2) ช่อง 'TYPE OF TRANSFORMER'
        3) คำที่เหลือทั้งหมด
    """
    # 1️⃣  ดูคีย์ที่ชัด ๆ ก่อน
    for k, v in attrs.items():
        if "USAGE" in k.upper() and "TAP" in k.upper():
            res = _classify_tap(str(v))
            if res:
                return res

    # 2️⃣  รองลงมา-ช่องชนิดหม้อแปลง
    for k, v in attrs.items():
        if "TYPE" in k.upper() and "TRANSFORMER" in k.upper():
            res = _classify_tap(str(v))
            if res:
                return res

    # 3️⃣  สุดท้าย-กวาดข้อความทุกช่อง
    all_text = " ".join(f"{k} {v}" for k, v in attrs.items())
    res = _classify_tap(all_text)
    if res:
        return res

    # 4️⃣  ไม่พบคำบ่งชี้ ⇒ N
    return "N"


def generate_powtr_code(attributes: Dict) -> str:
    part1 = PHASE_DIGIT               # (1) phase digit

    # ---------- 2) หาแรงดันข้างสูง (kV) ----------
    kv_high = None

    # helper ─ดึงตัวเลขที่ผูกกับหน่วย V/kV
    def kv_numbers(text: str) -> list[float]:
        out = []
        for m in re.finditer(r"(\d+(?:[.,]\d+)?)(\s*[kK]?[vV])", text):
            num = float(m.group(1).replace(",", "."))
            unit = m.group(2).strip().lower()
            if unit == "v":           # เป็นโวลต์ → แปลงเป็น kV
                num /= 1000.0
            out.append(num)
        return out

    raw_txts = [f"{k} {v}" for k, v in attributes.items()]          # รวม key+value

    # 2-A : จับรูป “High: … kV / V”
    for txt in raw_txts:
        m = re.search(r"[Hh]igh[^0-9]{0,10}(\d+(?:[.,]\d+)?)(\s*[kK]?[vV])", txt)
        if m:
            num = float(m.group(1).replace(",", "."))
            if m.group(2).strip().lower() == "v":
                num /= 1000.0
            kv_high = num
            break

    # 2-B : ถ้ายังไม่ได้ ดูรูป “HV: … kV / V”
    if kv_high is None:
        for txt in raw_txts:
            m = re.search(r"\bH[Vv][^\d]{0,10}(\d+(?:[.,]\d+)?)(\s*[kK]?[vV])", txt)
            if m:
                num = float(m.group(1).replace(",", "."))
                if m.group(2).strip().lower() == "v":
                    num /= 1000.0
                kv_high = num
                break

    # 2-C : สแกนทุก field เอาเฉพาะค่าระหว่าง 0.05–765 kV แล้วเลือกมากสุด
    if kv_high is None:
        kvs = []
        for txt in raw_txts:
            kvs += kv_numbers(txt)
        kvs = [x for x in kvs if 0.05 <= x <= 765]
        if kvs:
            kv_high = max(kvs)
    # -----------------------------------------------


    part2 = voltage_letter(kv_high)   # ได้ '-' ถ้าไม่พบ

    # (3) cooling / insulation ⇒ type letter
    cooling = attributes.get("TYPE", "") or attributes.get("INSULATION", "")
    part3 = type_letter(cooling)
    
    part4 = tap_letter(attributes)

    return f"POWTR-{part1}{part2}{part3}{part4}"
# -----------------------------------------------------------------------
# ------------------------------------------------------------------------


def add_powtr_codes(results: List[Dict]) -> List[Dict]:
    """Mutate each result dict to include POWTR_CODE if missing."""
    for r in results:
        data = r["extracted_data"]
        if "POWTR_CODE" not in data or not data["POWTR_CODE"]:
            data["POWTR_CODE"] = generate_powtr_code(data)
    return results


# ----------------- VALIDATION HELPERS --------------------------------------

RE_CODE = re.compile(r"^POWTR-(\d)([EHML])([OD])([OF])$")


def validate_powtr_code(code: str, attrs: Dict) -> bool:
    """Return True if code matches regenerated version."""
    if not code or not RE_CODE.match(code):
        return False
    return code == generate_powtr_code(attrs)


def process_excel(df: pd.DataFrame) -> pd.DataFrame:
    """Validate POWTR_CODE column in an uploaded spreadsheet."""
    attr_cols = [c for c in df.columns if c not in ("POWTR_CODE", "Is_Correct")]
    results = []
    for _, row in df.iterrows():
        attrs = {c: row[c] for c in attr_cols if pd.notna(row[c])}
        given = row.get("POWTR_CODE", "")
        ok = validate_powtr_code(given, attrs)
        row_out = row.to_dict()
        row_out["Is_Correct"] = ok
        if not ok:
            row_out["Suggested"] = generate_powtr_code(attrs)
        results.append(row_out)
    return pd.DataFrame(results)

# ----------------------------------------------------------------------------
# 6)  STREAMLIT UI
# ----------------------------------------------------------------------------

st.set_page_config(page_title="Transformer Extractor")
st.header("🔎 สกัดข้อมูลจากรูปภาพ")


tab1, tab2, tab3 = st.tabs([
    "สกัดจากรูปภาพ", "ตรวจสอบ POWTR-CODE", "ประมวลผลจาก validated"
])

# ---- TAB 1 -----------------------------------------------------------------
with tab1:
    images = st.file_uploader(
        "1. อัปโหลดรูปภาพ (รองรับหลายไฟล์)", ["jpg", "jpeg", "png"], True, key="img_upl"
    )
    if st.button("ประมวลผลภาพ") and images:
        attr_buf = fetch_attributes_from_github()
        prompt = generate_prompt_from_excel(attr_buf)
        # st.expander("Prompt ที่ใช้กับ Gemini").write(prompt)

        results = []
        prog = st.progress(0.0)
        for i, img in enumerate(images, 1):
            prog.progress(i / len(images))
            b64, mime = encode_image(img)
            data = extract_data_from_image(API_KEY, b64, mime, prompt)
            results.append({"file_name": img.name, "extracted_data": data})
        prog.empty()

        results = add_powtr_codes(results)

        # Flatten to long data frame
        rows = []
        for r in results:
            d = r["extracted_data"]
            common = {
                "FILE": r["file_name"],
                "POWTR_CODE": d.get("POWTR_CODE", ""),
            }
            if "error" in d:
                rows.append({**common, "ATTRIBUTE": "error", "VALUE": d["error"]})
            else:
                for k, v in d.items():
                    if k == "POWTR_CODE":
                        continue
                    rows.append({**common, "ATTRIBUTE": k, "VALUE": v})
        df_long = pd.DataFrame(rows)
        st.dataframe(df_long, use_container_width=True)

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            df_long.to_excel(w, index=False)
        buf.seek(0)
        st.download_button("ดาวน์โหลด Excel", buf, "extracted_long.xlsx")

# ---- TAB 2 -----------------------------------------------------------------
with tab3:
    st.subheader("ประมวลผลไฟล์ validated"); 
    validated_file = st.file_uploader("เลือกไฟล์ validated_powtr_codes.xlsx", ["xlsx"], key="val_upl")
    if st.button("ประมวลผล validated") and validated_file:
        df_val = pd.read_excel(validated_file)
        df_val = df_val[df_val.get("Is_Correct", True) == True]
        attr_buf = fetch_attributes_from_github()
        canonical = pd.read_excel(attr_buf).iloc[:, 0].dropna().astype(str).tolist()

        rows: List[Dict] = []
        for _, row in df_val.iterrows():
            asset = row.get("Location", "")
            powtr = row.get("Correct_POWTR_CODE", "")
            plant = row.get("Plant", "")
            siteid = (plant[:3] + "0") if plant else ""
            for attr in canonical:
                raw = row.get(attr, "-")
                val, unit = split_value_unit(raw)
                rows.append(
                    {
                        "ASSETNUM": asset,
                        "SITEID": siteid,
                        "POWTR_CODE": powtr,
                        "ATTRIBUTE": attr,
                        "VALUE": val,
                        "MEASUREUNIT": unit,
                    }
                )
        df_out = pd.DataFrame(rows)
        st.dataframe(df_out)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            df_out.to_excel(w, index=False)
        buf.seek(0)
        st.download_button("ดาวน์โหลด", buf, "extracted_long_from_validated.xlsx")

# ---- TAB 3 -----------------------------------------------------------------
with tab2:
    st.subheader("POWTR‑CODE Validator")
    upl = st.file_uploader("อัปโหลด Excel เพื่อตรวจสอบ", ["xlsx", "xls"], key="chk_upl")
    if upl:
        df_in = pd.read_excel(upl)
        df_out = process_excel(df_in)
        st.dataframe(df_out)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            df_out.to_excel(w, index=False)
        buf.seek(0)
        st.download_button(
            "ดาวน์โหลดรายงาน", buf, "validated_powtr_codes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.write(
            f"**รวม** {len(df_out)} แถว | **ถูกต้อง** {df_out['Is_Correct'].sum()} | **ผิด** {(~df_out['Is_Correct']).sum()}"
        )

# ---------------------------------------------------------------------------
# END OF FILE
