import streamlit as st, pandas as pd, requests, json, base64, io, re, os
from openpyxl import load_workbook
from pathlib import Path

# ────────────────────────── 0. CONFIG & CONSTANTS ───────────────────────────
TEMPLATE_PATH = Path('Template-MxLoader-Classification POW-TR.xlsm')  # commit ไฟล์นี้ใน repo
ATTRIBUTE_LIST = [
    'TYPE OF TRANSFORMER', 'STANDARD', 'MVA RATING', 'VOLTAGE LEVEL',
    'VECTOR GROUP', 'TYPE OF COOLING', 'OFF‑CIRCUIT TAP CHANGER',
    'SERIAL NUMBER', 'PERCENT IMPEDANCE', 'CONNECTION SYMBOL'
]

oil_kw = {'OIL','ONAN','ONAF','OFAF','OFWF','OA','OF','ON','ONO','OFA'}
dry_kw = {'DRY','CAST','RESIN'}

# ────────────────────────── 1. LOAD TEMPLATE ───────────────────────────────
if not TEMPLATE_PATH.exists():
    st.error('⚠️ ไม่พบไฟล์ template .xlsm ใน repo – โปรดเพิ่มไฟล์แล้ว deploy ใหม่')
    st.stop()
wb_tpl = load_workbook(TEMPLATE_PATH, keep_vba=True)

# ────────────────────────── 2. UI INPUTS ───────────────────────────────────
st.title('⚡ Transformer OCR → POWTR‑CODE → MxLoader (.xlsm)')

pam_xls   = st.file_uploader('📒 PAM.xlsx', ['xlsx','xls'])
imgs      = st.file_uploader('🖼️ Nameplate images', ['jpg','jpeg','png'], accept_multiple_files=True)

# API key จาก Secrets หรือกรอกมือ
api_key = os.getenv('GEMINI_API_KEY') or st.text_input('🔑 Gemini API key', type='password')

if pam_xls is not None:
    pam_df  = pd.read_excel(pam_xls)
    loc_col = st.selectbox('คอลัมน์ Location/AssetNUM ใน PAM', pam_df.columns,
                           index=list(pam_df.columns).index('Location') if 'Location' in pam_df.columns else 0)
else:
    pam_df  = pd.DataFrame(); loc_col=''

loc_map={}
if imgs and not pam_df.empty:
    st.markdown('**กรอก Location/AssetNUM ให้แต่ละรูป**')
    for im in imgs:
        loc_map[im.name] = st.text_input(im.name, key=im.name)

site_default = st.text_input('SITEID (default)', 'SBK0')

# ────────────────────────── 3. OCR / PROMPT ───────────────────────────────
def encode_img(f): return base64.b64encode(f.getvalue()).decode('utf-8')

def build_prompt_and_map():
    idx_map = {str(i+1): a for i, a in enumerate(ATTRIBUTE_LIST)}
    prompt  = ["กรุณาสกัดข้อมูลทั้งหมดจากรูปภาพเป็น JSON (ไม่พบให้ใส่ '-')\n"]
    prompt += [f"{i+1}: {a}" for i, a in enumerate(ATTRIBUTE_LIST)]
    return '\n'.join(prompt), idx_map

def gemini_ocr(api_key, img_b64, prompt):
    url=f"https://generativelanguage.googleapis.com/v1/models/gemini-2.0-flash:generateContent?key={api_key}"
    body={"contents":[{"parts":[{"text":prompt},
                                {"inline_data":{"mime_type":"image/jpeg","data":img_b64}}]}],
          "generation_config":{"temperature":0.2,"max_output_tokens":4096}}
    r=requests.post(url,json=body)
    if r.status_code!=200:
        return {"error":f"{r.status_code}:{r.text}"}
    txt=r.json()['candidates'][0]['content']['parts'][0]['text']
    try:
        return json.loads(txt[txt.find('{'):txt.rfind('}')+1])
    except:
        return {"raw_text":txt}

# ────────────────────────── 4. POWTR‑CODE LOGIC ───────────────────────────
def detect_voltage_kv(dic):
    """เลือกค่า kV ที่เหมาะสม (หลีกเลี่ยง BIL 900 kV)"""
    kvs_low, kvs_high = [], []
    pat = re.compile(r'(\d+(?:[.,]\d+)?)\s*(kV|KV|kv)?')
    for txt in map(str, dic.values()):
        if 'BIL' in txt.upper():     # ข้ามข้อความที่มี BIL
            continue
        for num, unit in pat.findall(txt):
            try: val=float(num.replace(',','.'))
            except: continue
            if unit:   # มี kV แนบมาด้วย
                kvs_low.append(val)
            else:
                # ไม่มีหน่วย – วิเคราะห์จากขนาดตัวเลข
                kv = val/1000 if val > 1000 else val
                kvs_low.append(kv)
    if kvs_low: return max(kvs_low)
    # fallback: อาจเหลือค่าที่มาจาก BIL (900 ฯลฯ) → ใช้ตัวต่ำสุดของ high list
    for txt in map(str, dic.values()):
        for num, unit in pat.findall(txt):
            try: val=float(num.replace(',','.'))
            except: continue
            if unit and val > 345: kvs_high.append(val)
    if kvs_high: return min(kvs_high)
    return None

def gen_powtr(data):
    txt=' '.join(map(str, data.values())).upper()
    phase = '1' if any(k in txt for k in ('1PH','1‑PHASE','SINGLE')) else '3'
    kv = detect_voltage_kv(data)
    if kv is None:
        v_char='-'
    else:
        v_char = 'E' if kv>=345 else 'H' if kv>=100 else 'M' if kv>=1 else 'L'
    type_char = 'O' if any(k in txt for k in oil_kw) else 'D'
    tap_char  = 'O' if any(k in txt for k in ('OLTC','ON‑LOAD')) else 'F'
    return f'POWTR-{phase}{v_char}{type_char}{tap_char}'

# ────────────────────────── 5. BUILD MxLoader ROWS ───────────────────────
ASHEET,HROW,DSTART = 'AssetAttr',2,3
header=[c.value for c in wb_tpl[ASHEET][HROW]]
col={h:i for i,h in enumerate(header) if h}

def idx2attr(k, mapping): return mapping.get(str(k).strip(), k)

def make_rows(asset, site, powtr, ocr, mapping):
    rows=[]
    hier=f"POWTR \\ {powtr}"
    for k,v in ocr.items():
        if k in ('error','raw_text'): continue
        attr=idx2attr(k, mapping)
        m=re.search(r'\((.*?)\)\s*$', attr)
        unit=m.group(1).strip() if m else ''
        r=['']*len(header)
        r[col['ASSETNUM']]=asset
        r[col['SITEID']]=site
        r[col['HIERARCHYPATH']]=hier
        r[col['ASSETSPEC.\nASSETATTRID']]=attr
        r[col['ASSETSPEC.ALNVALUE']]=v
        r[col['ASSETSPEC.MEASUREUNITID']]=unit
        rows.append(r)
    return rows

# ────────────────────────── 6. RUN ───────────────────────────────────────
if st.button('🚀 Run') and api_key and imgs and not pam_df.empty:
    prompt, mapping = build_prompt_and_map()
    ws = wb_tpl[ASHEET]
    if ws.max_row>=DSTART:
        ws.delete_rows(DSTART, ws.max_row-DSTART+1)

    summary=[]; prog=st.progress(0.)
    for i,im in enumerate(imgs,1):
        prog.progress(i/len(imgs))
        loc = loc_map.get(im.name,'').strip()
        if not loc:
            st.warning(f'{im.name} ไม่มี Location – ข้าม'); continue
        ocr = gemini_ocr(api_key, encode_img(im), prompt)
        powtr = gen_powtr(ocr) if isinstance(ocr, dict) else '-'
        pam_cls = pam_df.loc[pam_df[loc_col]==loc, 'Classification'].iat[0] \
                  if loc in pam_df[loc_col].values and 'Classification' in pam_df.columns else ''
        summary.append({'Image':im.name,'Asset':loc,
                        'POWTR(OCR)':powtr,'Classification(PAM)':pam_cls,
                        'Match?': powtr==pam_cls})
        for r in make_rows(loc, site_default, powtr, ocr, mapping):
            ws.append(r)

    st.subheader('ผลการตรวจ')
    st.dataframe(pd.DataFrame(summary))

    buf=io.BytesIO(); wb_tpl.save(buf); buf.seek(0)
    st.download_button('⬇️ Download MxLoader file', buf,
                       'MxLoader_POWTR_Result.xlsm',
                       'application/vnd.ms-excel.sheet.macroEnabled.12')
