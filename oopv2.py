# app.py – deploy‑ready (ATTRIBUTE.xlsx & template ใน repo)
import streamlit as st, pandas as pd, requests, json, base64, io, re, os
from openpyxl import load_workbook
from pathlib import Path

# ───────────── 0. เตรียมไฟล์คงที่ใน repo ────────────────────────────
TEMPLATE_PATH  = Path('Template-MxLoader-Classification POW-TR.xlsm')
ATTRIBUTE_PATH = Path('ATTRIBUTE.xlsx')

if not TEMPLATE_PATH.exists():
    TEMPLATE_PATH = st.file_uploader('📂 Template‑MxLoader‑Classification POW‑TR.xlsm', ['xlsm'])
    if TEMPLATE_PATH is None:
        st.error('กรุณาเพิ่มไฟล์ template .xlsm'); st.stop()
wb_tpl = load_workbook(TEMPLATE_PATH, keep_vba=True)

if not ATTRIBUTE_PATH.exists():
    ATTRIBUTE_PATH = st.file_uploader('📑 ATTRIBUTE.xlsx', ['xlsx','xls'])
    if ATTRIBUTE_PATH is None:
        st.error('กรุณาเพิ่มไฟล์ ATTRIBUTE.xlsx'); st.stop()

# ───────────── 1. ค่าคงที่อื่น ๆ ────────────────────────────────────
oil_kw = {'OIL','ONAN','ONAF','OFAF','OFWF','OA','OF','ON','ONO','OFA'}
dry_kw = {'DRY','CAST','RESIN'}

# ───────────── 2. UI Input ───────────────────────────────────────────
st.title('⚡ Transformer OCR → POWTR‑CODE → MxLoader (.xlsm)')

pam_xls = st.file_uploader('📒 PAM.xlsx', ['xlsx','xls'])
imgs    = st.file_uploader('🖼️ Nameplate images', ['jpg','jpeg','png'], accept_multiple_files=True)

api_key = os.getenv('GEMINI_API_KEY') or st.text_input('🔑 Gemini API key', type='password')

if pam_xls is not None:
    pam_df  = pd.read_excel(pam_xls)
    loc_col = st.selectbox('คอลัมน์ Location/AssetNUM ใน PAM', pam_df.columns,
                           index=list(pam_df.columns).index('Location') if 'Location' in pam_df.columns else 0)
else:
    pam_df=pd.DataFrame(); loc_col=''

loc_map={}
if imgs and not pam_df.empty:
    st.markdown('**กรอก Location/AssetNUM ให้แต่ละรูป**')
    for im in imgs:
        loc_map[im.name] = st.text_input(im.name, key=im.name)

site_default = st.text_input('SITEID (default)', 'SBK0')

# ───────────── 3. สร้าง prompt จาก ATTRIBUTE.xlsx ───────────────────
def build_prompt_and_map(attr_path):
    df = pd.read_excel(attr_path, header=None)
    df.columns=['attr']+[f'c{i}' for i in range(1,len(df.columns))]
    attrs=[str(a).strip() for a in df['attr'] if str(a).strip()]
    idx_map={str(i+1):a for i,a in enumerate(attrs)}
    prompt = ["กรุณาสกัดข้อมูลทั้งหมดจากรูปเป็น JSON (ไม่พบใส่ '-')\n"]
    prompt+= [f"{i+1}: {a}" for i,a in enumerate(attrs)]
    return '\n'.join(prompt), idx_map

def encode_img(f): return base64.b64encode(f.getvalue()).decode('utf-8')

def gemini_ocr(key,b64,prompt):
    url=f"https://generativelanguage.googleapis.com/v1/models/gemini-2.0-flash:generateContent?key={key}"
    body={"contents":[{"parts":[{"text":prompt},
                                {"inline_data":{"mime_type":"image/jpeg","data":b64}}]}],
          "generation_config":{"temperature":0.2,"max_output_tokens":4096}}
    r=requests.post(url,json=body)
    if r.status_code!=200: return {"error":f"{r.status_code}:{r.text}"}
    txt=r.json()['candidates'][0]['content']['parts'][0]['text']
    try: return json.loads(txt[txt.find('{'):txt.rfind('}')+1])
    except: return {"raw_text":txt}

# ───────────── 4. Logic POWTR‑CODE ───────────────────────────────────
def detect_voltage_kv(dic):
    """
    Return highest kV but:
    • ข้ามข้อความที่มี BIL หรือ '/ AC'
    • แปลง 'V' (volts) เป็น kV เมื่อ >1 000 V
    """
    kv=[]
    pat=re.compile(r'(\d+(?:[.,]\d+)?)\s*(kV|KV|kv|V|v|volt|VOLTS?)?')
    for txt in map(str,dic.values()):
        up=txt.upper()
        if 'BIL' in up:          # ข้ามทั้งสตริง
            continue
        for num,unit in pat.findall(txt):
            try: val=float(num.replace(',','.'))
            except: continue
            ctx=up[up.find(num): up.find(num)+20]  # เนื้อหาใกล้ๆ
            if '/ AC' in ctx or ' AC ' in ctx:     # insulation AC withstand
                continue
            if unit and unit.lower().startswith('k'):
                kv.append(val)
            elif unit and unit.lower().startswith('v'):
                kv.append(val/1000)
            else:
                kv.append(val/1000 if val>1000 else val)
    return max(kv) if kv else None

def gen_powtr(d):
    txt=' '.join(map(str,d.values())).upper()
    phase='1' if any(k in txt for k in ('1PH','1‑PHASE','SINGLE')) else '3'
    kv=detect_voltage_kv(d)
    v_char='-' if kv is None else ('E' if kv>=345 else 'H' if kv>=100 else 'M' if kv>=1 else 'L')
    type_char='O' if any(k in txt for k in oil_kw) else 'D'
    tap_char='O' if any(k in txt for k in ('OLTC','ON‑LOAD')) else 'F'
    return f'POWTR-{phase}{v_char}{type_char}{tap_char}'

# ───────────── 5. สร้างแถว AssetAttr ────────────────────────────────
ASHEET,HROW,DSTART='AssetAttr',2,3
hdr=[c.value for c in wb_tpl[ASHEET][HROW]]
col={h:i for i,h in enumerate(hdr) if h}

def idx2attr(k,map_idx): return map_idx.get(str(k).strip(), k)
def make_rows(asset,site,powtr,ocr,map_idx):
    hier=f"POWTR \\ {powtr}"; rows=[]
    for k,v in ocr.items():
        if k in ('error','raw_text'): continue
        attr=idx2attr(k,map_idx)
        m=re.search(r'\((.*?)\)\s*$',attr); unit=m.group(1).strip() if m else ''
        r=['']*len(hdr)
        r[col['ASSETNUM']]=asset
        r[col['SITEID']]=site
        r[col['HIERARCHYPATH']]=hier
        r[col['ASSETSPEC.\nASSETATTRID']]=attr
        r[col['ASSETSPEC.ALNVALUE']]=v
        r[col['ASSETSPEC.MEASUREUNITID']]=unit
        rows.append(r)
    return rows

# ───────────── 6. RUN ───────────────────────────────────────────────
if st.button('🚀 Run') and api_key and imgs and not pam_df.empty:
    prompt, idx_map = build_prompt_and_map(ATTRIBUTE_PATH)
    # ล้าง sheet เก่า
    ws=wb_tpl[ASHEET]
    if ws.max_row>=DSTART: ws.delete_rows(DSTART, ws.max_row-DSTART+1)

    out=[]; prog=st.progress(0.)
    for i,im in enumerate(imgs,1):
        prog.progress(i/len(imgs))
        loc=loc_map.get(im.name,'').strip()
        if not loc: st.warning(f'{im.name} ไม่มี Location — ข้าม'); continue
        ocr=gemini_ocr(api_key, encode_img(im), prompt)
        powtr=gen_powtr(ocr) if isinstance(ocr,dict) else '-'
        pam_cls = pam_df.loc[pam_df[loc_col]==loc,'Classification'].iat[0] \
                  if loc in pam_df[loc_col].values and 'Classification' in pam_df.columns else ''
        out.append({'Image':im.name,'Asset':loc,
                    'POWTR(OCR)':powtr,'Classification(PAM)':pam_cls,
                    'Match?':powtr==pam_cls})
        for r in make_rows(loc, site_default, powtr, ocr, idx_map): ws.append(r)

    st.subheader('ผลการตรวจ')
    st.dataframe(pd.DataFrame(out))

    buf=io.BytesIO(); wb_tpl.save(buf); buf.seek(0)
    st.download_button('⬇️ Download MxLoader file', buf,
                       'MxLoader_POWTR_Result.xlsm',
                       'application/vnd.ms-excel.sheet.macroEnabled.12')
