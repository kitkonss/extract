# -------------------  extract-excel.py  (FULL – 29 Apr 2025)  -------------------
import os, base64, json, re, io, imghdr, requests, pandas as pd, streamlit as st
from PIL import Image

# -------------------------------------------------------------------- #
# 1)  Utilities                                                        #
# -------------------------------------------------------------------- #
def encode_image(f):                     # -> (b64, mime)
    raw = f.getvalue();  kind = imghdr.what(None, raw) or 'jpeg'
    return base64.b64encode(raw).decode(), f"image/{'jpg' if kind=='jpeg' else kind}"

def _kv_from_text(txt: str) -> float | None:
    """Return highest kV in text (skip kVA/VA/kA/A, BIL, >1 500 kV)."""
    best = None
    for chunk in re.split(r'[\/,;\n]', txt.upper()):
        if re.search(r'K?VA\b|KA\b|AMP\b|\b[A-Z]?A\b', chunk):   # skip kVA, A
            continue
        if 'BIL' in chunk or 'IMPULSE' in chunk:
            continue
        for m in re.finditer(r'(\d+(?:\.\d+)?)\s*([K]?V)(?![A-Z])', chunk):
            kv = float(m.group(1)) if m.group(2) == 'KV' else float(m.group(1))/1000
            if kv > 1500:  continue
            best = kv if best is None else max(best, kv)
    return best

def fallback_voltage_scan(raw_text: str) -> str | None:
    """
    Scan Gemini raw response (string) for the first sensible HV voltage line.
    Accept patterns like 'HV 21000 V', 'PRI. VOLT 22000 V'.
    """
    for line in raw_text.upper().splitlines():
        if re.search(r'\b(HV|PRI|PRIMARY|HIGH)\b.*\d', line):
            m = re.search(r'(\d+(?:\.\d+)?)\s*([K]?V)(?![A-Z])', line)
            if m and 'KVA' not in line:
                return f"{m.group(1)} {m.group(2)}"
    return None

def clean_voltage_fields(data, raw_text):
    """Replace '-' voltage values by fallback scan (if available)."""
    for k in list(data.keys()):
        if re.search(r'VOLT|HV|LV|PRI|SEC', k, re.I) and (data[k] in {'-', '', None}):
            fb = fallback_voltage_scan(raw_text)
            if fb:  data[k] = fb
    return data

# -------------------------------------------------------------------- #
# 2)  Prompt generator                                                 #
# -------------------------------------------------------------------- #
def generate_prompt_from_excel(excel_file):
    df = pd.read_excel(excel_file)
    if isinstance(df.columns[0], (int,float)):
        excel_file.seek(0); df = pd.read_excel(excel_file, header=None)
        df.columns = ['attribute_name'] + [f'col_{i}' for i in range(1,len(df.columns))]
    attr_col = 'attribute_name' if 'attribute_name' in df.columns else df.columns[0]
    unit_col = next((c for c in df.columns if re.fullmatch(r'(unit|uom)', str(c), re.I)), None)

    prompt = ["""Extract **JSON** with these attributes (English keys only).  
For *voltage‑related* fields (HV/LV Rated Voltage, Primary Voltage, Voltage Level, etc.)  
→ pick the line that shows a number followed by **V** or **kV** — ignore kVA/kA/A.  
If not found, return "-"."""]
    for i,row in df.iterrows():
        a=str(row[attr_col]).strip();  u=row.get(unit_col,'')
        if a and not pd.isna(a):
            prompt.append(f"{i+1}: {a}" + (f" [{u}]" if u and not pd.isna(u) else ""))
    prompt.append("Respond **only** with a JSON object.")
    return '\n'.join(prompt)

# -------------------------------------------------------------------- #
# 3)  Gemini API                                                       #
# -------------------------------------------------------------------- #
def gemini(api_key,b64,mime,prompt):
    url=f"https://generativelanguage.googleapis.com/v1/models/gemini-2.0-flash:generateContent?key={api_key}"
    body={"contents":[{"parts":[{"text":prompt},
                                {"inlineData":{"mimeType":mime,"data":b64}}]}],
          "generationConfig":{"temperature":0.2,"topP":0.85,"maxOutputTokens":9000}}
    r=requests.post(url,headers={"Content-Type":"application/json"},data=json.dumps(body))
    if r.ok and r.json().get('candidates'):
        return r.json()['candidates'][0]['content']['parts'][0]['text']
    return f"API ERROR {r.status_code}: {r.text}"

# -------------------------------------------------------------------- #
# 4)  POWTR‑CODE                                                       #
# -------------------------------------------------------------------- #
def make_powtr(d):
    phase='1' if any(re.search(r'1PH|1-PH|SINGLE',str(v),re.I) for v in d.values()) else '3'
    v_max=max((_kv_from_text(str(v)) or -1) for k,v in d.items()
              if re.search(r'VOLT|HV|LV|PRI|SEC',k,re.I))
    if v_max==-1: v_char='-'
    elif v_max>765: return 'POWTR-3-OO'
    elif v_max>=345: v_char='E'
    elif v_max>=100: v_char='H'
    elif v_max>=1  : v_char='M'
    else:            v_char='L'
    t_char='D' if any('DRY' in str(v).upper() for v in d.values()) else \
            ('O' if any('OIL' in str(v).upper() for v in d.values()) else '-')
    tap='O' if any(re.search(r'ON[- ]?LOAD|OLTC',str(v),re.I) for v in d.values()) else 'F'
    return f"POWTR-{phase}{v_char}{t_char}{tap}"

# -------------------------------------------------------------------- #
# 5)  Streamlit UI                                                     #
# -------------------------------------------------------------------- #
st.title("Transformer Nameplate Extractor + POWTR‑Code (29 Apr 2025)")

API="AIzaSyDb8iBV1EWqLvjheG_44gh3vQHfpmYGOCI"
tab1,tab2=st.tabs(["ใช้ไฟล์ Excel","ใช้ attribute สำเร็จรูป"])
with tab1:
    excel=st.file_uploader("Excel attributes",["xlsx","xls"])
    if excel: st.dataframe(pd.read_excel(excel).head())
with tab2:
    prompt_default="""Extract JSON with keys: MANUFACTURER, MODEL, SERIAL_NO, STANDARD,
CAPACITY [kVA], HV_RATED_VOLTAGE [V], LV_RATED_VOLTAGE [V],
HV_RATED_CURRENT [A], LV_RATED_CURRENT [A], IMPEDANCE_VOLTAGE [%],
VECTOR_GROUP"""

imgs=st.file_uploader("อัปโหลดรูปภาพ",["jpg","png","jpeg"],accept_multiple_files=True)

if st.button("ประมวลผล") and imgs:
    prompt = prompt_default if not excel else generate_prompt_from_excel(excel)
    st.expander("Prompt").write(prompt)
    out=[]
    bar,stat=st.progress(0),st.empty()

    for i,f in enumerate(imgs,1):
        bar.progress(i/len(imgs));  stat.write(f"กำลังประมวลผล {f.name}")
        b64,mime=encode_image(f);   raw=gemini(API,b64,mime,prompt)

        try: data=json.loads(raw[raw.find('{'):raw.rfind('}')+1])
        except Exception: data={"raw_text":raw}

        if 'raw_text' in data:
            data=clean_voltage_fields(data, data['raw_text'])
        else:
            data=clean_voltage_fields(data, raw)

        if 'POWTR_CODE' not in data:
            data['POWTR_CODE']=make_powtr(data)
        out.append({"file":f.name,"data":data})

    st.subheader("POWTR‑CODE")
    for r in out: st.write(r['data']['POWTR_CODE'])

    rows=[]
    for r in out:
        d=r['data']; code=d['POWTR_CODE']
        for k,v in d.items():
            if k!='POWTR_CODE': rows.append({"POWTR_CODE":code,"ATTRIBUTE":k,"VALUE":v})
    df=pd.DataFrame(rows); st.dataframe(df)

    buf=io.BytesIO()
    with pd.ExcelWriter(buf,engine='openpyxl') as w: df.to_excel(w,index=False)
    buf.seek(0)
    st.download_button("ดาวน์โหลด Excel",buf,"extracted_data_sorted.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
# -------------------------------------------------------------------- #
