# app.py  —  Transformer OCR  ➜  POWTR‑CODE  ➜  MxLoader (.xlsm)
import streamlit as st, pandas as pd, requests, json, base64, io, re, os
from openpyxl import load_workbook
from pathlib import Path

# ─────────────── 0. Load fixed files (template & attribute) ────────────────
TEMPLATE_PATH  = Path('Template-MxLoader-Classification POW-TR.xlsm')
ATTRIBUTE_PATH = Path('ATTRIBUTE.xlsx')

if not TEMPLATE_PATH.exists():
    TEMPLATE_PATH = st.file_uploader('📂 Template‑MxLoader‑Classification POW‑TR.xlsm', ['xlsm'])
    if TEMPLATE_PATH is None:
        st.error('ต้องมีไฟล์ template .xlsm ใน repo หรืออัปโหลด'); st.stop()
wb_tpl = load_workbook(TEMPLATE_PATH, keep_vba=True)

if not ATTRIBUTE_PATH.exists():
    ATTRIBUTE_PATH = st.file_uploader('📑 ATTRIBUTE.xlsx', ['xlsx','xls'])
    if ATTRIBUTE_PATH is None:
        st.error('ต้องมีไฟล์ ATTRIBUTE.xlsx ใน repo หรืออัปโหลด'); st.stop()

# ─────────────── 1. UI Inputs ───────────────────────────────────────────────
st.title('⚡ Transformer OCR → POWTR‑CODE → MxLoader (.xlsm)')

pam_xls = st.file_uploader('📒 PAM.xlsx', ['xlsx','xls'])
imgs    = st.file_uploader('🖼️ Nameplate images', ['jpg','jpeg','png'],
                           accept_multiple_files=True)
api_key = value="AIzaSyDb8iBV1EWqLvjheG_44gh3vQHfpmYGOCI"

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

# ─────────────── 2. Build prompt from ATTRIBUTE.xlsx ───────────────────────
def build_prompt_and_map(attr_path):
    df = pd.read_excel(attr_path, header=None)
    df.columns = ['attr'] + [f'c{i}' for i in range(1,len(df.columns))]
    attrs = [str(a).strip() for a in df['attr'] if str(a).strip()]
    idx_map = {str(i+1): a for i, a in enumerate(attrs)}

    prompt = """
โปรดอ่านแผ่นป้ายหม้อแปลงและคืนค่าเป็น JSON ตามรูปแบบนี้
(หากไม่พบให้คืนเป็นค่าว่าง ""):

{
  "HIGH_SIDE_VOLTAGE_KV": <ตัวเลข kV ของแรงดันขาเข้า>,
  "PHASE": <1 หรือ 3>,
  "COOLING_TYPE": "<ONAN/ONAF/OFAF/DRY/…>",
  "TAP_CHANGER": "ON‑LOAD" | "OFF‑CIRCUIT" | "—",
  "VECTOR_GROUP": "<ตัวอย่าง: YNd11>"
}

ห้ามใส่ค่า BIL หรือ AC withstand ลงใน HIGH_SIDE_VOLTAGE_KV
"""
    # เพิ่มรายการ attribute ปกติไว้ท้าย prompt (ช่วยให้โมเดลหาได้ครบ)
    prompt += "\n\nเพิ่มเติมกรุณาดึงค่าต่อไปนี้:\n"
    for i,a in enumerate(attrs,1):
        prompt += f"{i}: {a}\n"
    prompt += "\nหากไม่พบค่าให้ใส่ '-'"

    return prompt, idx_map

def encode_img(file): return base64.b64encode(file.getvalue()).decode('utf-8')

# ─────────────── 3. Gemini OCR ─────────────────────────────────────────────
def gemini_ocr(key, img_b64, prompt):
    url=f"https://generativelanguage.googleapis.com/v1/models/gemini-2.0-flash:generateContent?key={key}"
    body={"contents":[{"parts":[{"text":prompt},
                                {"inline_data":{"mime_type":"image/jpeg","data":img_b64}}]}],
          "generation_config":{"temperature":0.2,"max_output_tokens":4096}}
    r=requests.post(url,json=body)
    if r.status_code!=200:
        return {"error":f"{r.status_code}:{r.text}"}
    txt=r.json()['candidates'][0]['content']['parts'][0]['text']
    try:
        return json.loads(txt[txt.find('{'):txt.rfind('}')+1])
    except Exception:
        return {"raw_text": txt}

# ─────────────── 4. POWTR‑CODE logic ──────────────────────────────────────
oil_kw = {'OIL','ONAN','ONAF','OFAF','OFWF','OA','OF','ON','ONO','OFA'}
dry_kw = {'DRY','CAST','RESIN'}

def detect_voltage_kv(dic):
    """หา kV สูงสุด (หลีกเลี่ยง BIL / AC Test)"""
    kv=[]
    pat=re.compile(r'(\d+(?:[.,]\d+)?)\s*(kV|KV|kv|V|v)?')
    for txt in map(str,dic.values()):
        upper=txt.upper()
        if 'BIL' in upper or '/ AC' in upper or 'AC ' in upper:
            continue
        for num,unit in pat.findall(txt):
            try: val=float(num.replace(',','.'))
            except: continue
            if unit and unit.lower().startswith('k'):
                kv.append(val)
            elif unit and unit.lower().startswith('v'):
                kv.append(val/1000)
            else:  # ไม่มีหน่วย
                kv.append(val/1000 if val>1000 else val)
    return max(kv) if kv else None

def gen_powtr(data):
    """สร้าง POWTR‑CODE จาก key ตรง ถ้าไม่ครบ fallback regex"""
    all_txt = ' '.join(map(str, data.values())).upper()
    phase = str(data.get('PHASE','3')).replace('.0','')
    if phase not in {'1','3'}: phase='3'

    hv_key = data.get('HIGH_SIDE_VOLTAGE_KV')
    kv = None
    if hv_key:
        try: kv=float(hv_key)
        except: kv=None
    if kv is None:
        kv = detect_voltage_kv(data)

    v_char = '-' if kv is None else (
        'E' if kv>=345 else 'H' if kv>=100 else 'M' if kv>=1 else 'L')

    cooling = str(data.get('COOLING_TYPE','')).upper()
    t_char = 'O' if any(k in cooling for k in oil_kw) else 'D'

    tap_field = str(data.get('TAP_CHANGER','')).upper()
    tap_char = 'O' if 'ON' in tap_field else 'F'

    return f'POWTR-{phase}{v_char}{t_char}{tap_char}', kv

# ─────────────── 5. MxLoader row builder ─────────────────────────────────
ASHEET,HROW,DSTART='AssetAttr',2,3
hdr=[c.value for c in wb_tpl[ASHEET][HROW]]
col={h:i for i,h in enumerate(hdr) if h}

def idx2attr(k,idx_map): return idx_map.get(str(k).strip(), k)

def make_rows(asset,site,powtr,ocr,idx_map):
    hier=f"POWTR \\ {powtr}"; rows=[]
    for k,v in ocr.items():
        if k in ('error','raw_text'): continue
        attr=idx2attr(k, idx_map)
        mu=re.search(r'\((.*?)\)\s*$', attr); unit=mu.group(1).strip() if mu else ''
        r=['']*len(hdr)
        r[col['ASSETNUM']],r[col['SITEID']],r[col['HIERARCHYPATH']]=asset,site,hier
        r[col['ASSETSPEC.\nASSETATTRID']]=attr
        r[col['ASSETSPEC.ALNVALUE']]=v
        r[col['ASSETSPEC.MEASUREUNITID']]=unit
        rows.append(r)
    return rows

def show_debug(i, ocr, kv):
    with st.expander(f'OCR & Debug – image #{i+1}'):
        st.json(ocr, expanded=False)
        st.write(f'**Voltage decision:** ใช้ {kv if kv else "N/A"} kV '
                 '(หลังกรอง BIL/AC)')

# ─────────────── 6. RUN ────────────────────────────────────────────────
if st.button('🚀 Run') and api_key and imgs and not pam_df.empty:
    prompt, idx_map = build_prompt_and_map(ATTRIBUTE_PATH)

    ws = wb_tpl[ASHEET]
    if ws.max_row>=DSTART:
        ws.delete_rows(DSTART, ws.max_row-DSTART+1)

    results=[]; prog=st.progress(0.)
    for i,im in enumerate(imgs,1):
        prog.progress(i/len(imgs))
        loc=loc_map.get(im.name,'').strip()
        if not loc:
            st.warning(f'{im.name} ไม่มี Location — ข้าม'); continue

        ocr = gemini_ocr(api_key, encode_img(im), prompt)
        powtr, kv_used = gen_powtr(ocr if isinstance(ocr,dict) else {})

        show_debug(i, ocr, kv_used)

        pam_cls = pam_df.loc[pam_df[loc_col]==loc, 'Classification'].iat[0] \
                  if loc in pam_df[loc_col].values and 'Classification' in pam_df.columns else ''
        results.append({'Image':im.name,'Asset':loc,
                        'POWTR(OCR)':powtr,'Classification(PAM)':pam_cls,
                        'Match?':powtr==pam_cls})
        for r in make_rows(loc, site_default, powtr, ocr, idx_map):
            ws.append(r)

    st.subheader('ผลการตรวจ')
    st.dataframe(pd.DataFrame(results))

    buf=io.BytesIO(); wb_tpl.save(buf); buf.seek(0)
    st.download_button('⬇️ Download MxLoader file', buf,
                       'MxLoader_POWTR_Result.xlsm',
                       'application/vnd.ms-excel.sheet.macroEnabled.12')
