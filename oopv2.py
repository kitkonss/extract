# app.py  –  OCR ➜ POWTR‑CODE ➜ MxLoader (.xlsm)
import streamlit as st, pandas as pd, requests, json, base64, io, re, os
from openpyxl import load_workbook
from pathlib import Path

# ─── 0. fixed files ──────────────────────────────────────────────────
TPL  = Path('Template-MxLoader-Classification POW-TR.xlsm')
ATTR = Path('ATTRIBUTE.xlsx')
if not TPL.exists():
    TPL = st.file_uploader('📂 template .xlsm', ['xlsm']);  st.stop() if TPL is None else None
wb = load_workbook(TPL, keep_vba=True)
if not ATTR.exists():
    ATTR = st.file_uploader('📑 ATTRIBUTE.xlsx', ['xlsx','xls']);  st.stop() if ATTR is None else None

# ─── 1. attribute list & index map ───────────────────────────────────
df_attr = pd.read_excel(ATTR, header=None)
ATTR_LIST = [str(a).strip() for a in df_attr[0] if str(a).strip()]
IDX_MAP   = {str(i+1): a for i, a in enumerate(ATTR_LIST)}   # "1"→attr

def build_prompt():
    p = """
คืน JSON เช่น
{ "HIGH_SIDE_VOLTAGE_KV": 230, "PHASE": 3,
  "COOLING_TYPE": "ONAN / ONAF",
  "TAP_CHANGER": "OFF‑CIRCUIT",
  "VECTOR_GROUP": "YNd1" }

หากไม่พบให้ใส่ "" และ **อย่าใช้ค่า BIL / AC withstand**

พร้อมคืนค่าต่อไปนี้ (key เป็นเลข):\n"""
    p += '\n'.join(f"{i}: {a}" for i, a in enumerate(ATTR_LIST, 1))
    return p

PROMPT = build_prompt()

# ─── 2. UI ───────────────────────────────────────────────────────────
st.title('⚡ Transformer OCR → POWTR‑CODE → MxLoader')
pam_xls = st.file_uploader('📒 PAM.xlsx', ['xlsx','xls'])
imgs    = st.file_uploader('🖼️ Images', ['jpg','jpeg','png'], accept_multiple_files=True)
api_key = os.getenv('GEMINI_API_KEY') or st.text_input('API key', type='password')

if pam_xls is not None:
    pam_df  = pd.read_excel(pam_xls)
    loc_col = st.selectbox('Location column', pam_df.columns,
                           index=list(pam_df.columns).index('Location') if 'Location' in pam_df else 0)
else:
    pam_df = pd.DataFrame(); loc_col=''

loc_map={}
if imgs and not pam_df.empty:
    st.markdown('**AssetNUM / Location ต่อไฟล์**')
    for im in imgs:
        loc_map[im.name]=st.text_input(im.name,key=im.name)
site_default=st.text_input('SITEID (default)','SBK0')

# ─── 3. OCR helper ──────────────────────────────────────────────────
def b64(file): return base64.b64encode(file.getvalue()).decode()
def gemini(api,k,p):
    r=requests.post(
        f"https://generativelanguage.googleapis.com/v1/models/gemini-2.0-flash:generateContent?key={api}",
        json={"contents":[{"parts":[{"text":p},
                                    {"inline_data":{"mime_type":"image/jpeg","data":k}}]}],
              "generation_config":{"temperature":0.2,"max_output_tokens":4096}})
    if r.status_code!=200: return {"error":r.text}
    t=r.json()['candidates'][0]['content']['parts'][0]['text']
    try:return json.loads(t[t.find('{'):t.rfind('}')+1])
    except: return {"raw_text":t}

# ─── 4. POWTR logic ─────────────────────────────────────────────────
oil_kw={'OIL','ONAN','ONAF','OFAF','OFWF','OA','OF','ON','ONO','OFA'}
def detect_kv(dic):
    patt=re.compile(r'(\d{2,7}(?:[ ,]\d{3})*(?:[.,]\d+)?)\s*(kV|KV|kv|V|v)?')
    kvs=[]
    for v in dic.values():
        if isinstance(v,str) and ('BIL' in v.upper() or '/ AC' in v.upper()): continue
        for raw,unit in patt.findall(str(v)):
            num=raw.replace(' ','').replace(',','')
            try: val=float(num.replace(',','.'))
            except: continue
            kv = val/1000 if (unit.lower().startswith('v') if unit else val>1000) else val
            if kv<=765: kvs.append(kv)
    return (max(kvs) if kvs else None, kvs)

def powtr(d):
    ph=str(d.get('PHASE','3')).replace('.0','')
    kv, cand=detect_kv(d)
    v='-' if kv is None else ('E' if kv>=345 else 'H' if kv>=100 else 'M' if kv>=1 else 'L')
    cool=str(d.get('COOLING_TYPE','')).upper()
    t='O' if any(k in cool for k in oil_kw) else 'D'
    tap='O' if 'ON' in str(d.get('TAP_CHANGER','')).upper() else 'F'
    return f'POWTR-{ph}{v}{t}{tap}', kv, cand

# ─── 5. Sheet helpers ───────────────────────────────────────────────
ASHEET,HROW,DSTART='AssetAttr',2,3
hdr=[c.value for c in wb[ASHEET][HROW]]
col={h:i for i,h in enumerate(hdr) if h}
def full_dict(ocr):                     # เติมค่าว่างให้ครบ
    d={a:'' for a in ATTR_LIST}; d.update(ocr); return d
def blank(v,a): return '' if v in {'-','',None} or str(v).strip().upper()==a.upper() else v
def rows(asset,site,code,fd):
    hier=f"POWTR \\ {code}"; out=[]
    for a in ATTR_LIST:
        v=blank(fd.get(a,''),a)
        u=re.search(r'\((.*?)\)\s*$',a); unit=u.group(1).strip() if u else ''
        r=['']*len(hdr)
        r[col['ASSETNUM']],r[col['SITEID']],r[col['HIERARCHYPATH']]=asset,site,hier
        r[col['ASSETSPEC.\nASSETATTRID']]=a
        r[col['ASSETSPEC.ALNVALUE']]=v
        r[col['ASSETSPEC.MEASUREUNITID']]=unit
        out.append(r)
    return out
def dbg(i,d,kv,cand):
    with st.expander(f'Debug #{i+1}'):
        st.json(d); st.write('kV cand →',cand,'| chosen →',kv)

# ─── 6. RUN ─────────────────────────────────────────────────────────
if st.button('🚀 Run') and api_key and imgs and not pam_df.empty:
    ws=wb[ASHEET]
    if ws.max_row>=DSTART: ws.delete_rows(DSTART, ws.max_row-DSTART+1)
    res=[]; prog=st.progress(0.)
    for i,im in enumerate(imgs,1):
        prog.progress(i/len(imgs))
        loc=loc_map.get(im.name,'').strip()
        if not loc: st.warning(f'{im.name} ไม่มี Location'); continue
        ocr_raw=gemini(api_key, b64(im), PROMPT)
        ocr_full=full_dict(ocr_raw if isinstance(ocr_raw,dict) else {})
        code, kv, cand=powtr(ocr_full)
        dbg(i, ocr_full, kv, cand)
        pam_cls=pam_df.loc[pam_df[loc_col]==loc,'Classification'].iat[0] \
                if loc in pam_df[loc_col].values and 'Classification' in pam_df.columns else ''
        res.append({'Image':im.name,'Asset':loc,'POWTR(OCR)':code,
                    'Classification(PAM)':pam_cls,'Match?':code==pam_cls})
        for r in rows(loc,site_default,code,ocr_full): ws.append(r)

    st.dataframe(pd.DataFrame(res))
    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    st.download_button('⬇️ Download', buf, 'MxLoader_POWTR_Result.xlsm',
                       'application/vnd.ms-excel.sheet.macroEnabled.12')
