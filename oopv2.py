# app.py  –  single‑file deploy
import streamlit as st, pandas as pd, requests, json, base64, io, re, os
from openpyxl import load_workbook
from pathlib import Path

# ───────────── 0. fixed files ─────────────
TPL = Path('Template-MxLoader-Classification POW-TR.xlsm')
ATTR = Path('ATTRIBUTE.xlsx')
if not TPL.exists():
    TPL = st.file_uploader('📂 template .xlsm', ['xlsm'])
    if TPL is None: st.stop()
wb = load_workbook(TPL, keep_vba=True)
if not ATTR.exists():
    ATTR = st.file_uploader('📑 ATTRIBUTE.xlsx', ['xlsx','xls'])
    if ATTR is None: st.stop()

# ───────────── 1. UI ─────────────
st.title('⚡ Transformer OCR ▸ POWTR‑CODE ▸ MxLoader')
pam_xls = st.file_uploader('📒 PAM.xlsx', ['xlsx','xls'])
imgs    = st.file_uploader('🖼️ Images', ['jpg','jpeg','png'], accept_multiple_files=True)
api_key = value="AIzaSyDb8iBV1EWqLvjheG_44gh3vQHfpmYGOCI"

if pam_xls is not None:
    pam_df = pd.read_excel(pam_xls)
    loc_col = st.selectbox('Location column', pam_df.columns,
                           index=list(pam_df.columns).index('Location') if 'Location' in pam_df else 0)
else:
    pam_df = pd.DataFrame(); loc_col=''

loc_map={}
if imgs and not pam_df.empty:
    st.markdown('**Map Location → file**')
    for im in imgs:
        loc_map[im.name] = st.text_input(im.name, key=im.name)

# ───────────── 2. prompt ─────────────
def build_prompt(attr):
    df=pd.read_excel(attr, header=None); df.columns=['a']+df.columns[1:].tolist()
    lines=["คืน JSON ดังนี้ (ไม่พบใส่ \"\"):\n"
           "{ \"HIGH_SIDE_VOLTAGE_KV\":…, \"PHASE\":…, \"COOLING_TYPE\":…, "
           "\"TAP_CHANGER\":\"ON‑LOAD|OFF‑CIRCUIT\", \"VECTOR_GROUP\":… }"]
    for i,a in enumerate(df['a'],1):
        if str(a).strip(): lines.append(f"{i}: {a}")
    idx={str(i+1):str(a).strip() for i,a in enumerate(df['a']) if str(a).strip()}
    return '\n'.join(lines), idx

prompt, idx_map = build_prompt(ATTR)

# ───────────── 3. OCR ─────────────
def enc(f): return base64.b64encode(f.getvalue()).decode()
def ocr(api, b64, pr):
    r=requests.post(
        f"https://generativelanguage.googleapis.com/v1/models/gemini-2.0-flash:generateContent?key={api}",
        json={"contents":[{"parts":[{"text":pr},
                                    {"inline_data":{"mime_type":"image/jpeg","data":b64}}]}],
              "generation_config":{"temperature":0.2,"max_output_tokens":4096}})
    if r.status_code!=200: return {"error":r.text}
    t=r.json()['candidates'][0]['content']['parts'][0]['text']
    try:return json.loads(t[t.find('{'):t.rfind('}')+1])
    except: return {"raw_text":t}

# ───────────── 4. POWTR logic ─────────────
oil_kw={'OIL','ONAN','ONAF','OFAF','OFWF','OA','OF','ON','ONO','OFA'}

def kV_detector(d):
    pat=re.compile(r'(\d{2,7}(?:[.,]\d+)?)\s*(kV|KV|kv|V|v)?')
    good=[]
    for txt in map(str,d.values()):
        u=txt.upper()
        if any(x in u for x in ('BIL','IMP','LIGHTNING','/ AC',' AC ')): continue
        for num,unit in pat.findall(txt):
            try: val=float(num.replace(',','.'))
            except: continue
            kv = val/1000 if (unit and unit.lower().startswith('v') or (not unit and val>1000)) else val
            if kv<=765: good.append(kv)
    return (max(good), good) if good else (None, [])

def powtr(d):
    ph=str(d.get('PHASE','3')).replace('.0','')
    kv, cand=kV_detector(d)
    v='-' if kv is None else ('E' if kv>=345 else 'H' if kv>=100 else 'M' if kv>=1 else 'L')
    cooling=(str(d.get('COOLING_TYPE',''))+' '+str(d.get('TYPE OF COOLING',''))).upper()
    t='O' if any(k in cooling for k in oil_kw) else 'D'
    tap='O' if 'ON' in str(d.get('TAP_CHANGER','')).upper() else 'F'
    return f'POWTR-{ph}{v}{t}{tap}', kv, cand

def debug(i, ocr_d, kv, cand):
    with st.expander(f'Debug #{i+1}'):
        st.json(ocr_d)
        st.write('kV candidates →', cand)
        st.write('chosen →', kv)

# ───────────── 5. MxLoader sheet prep ─────────────
ASHEET,HROW,DSTART='AssetAttr',2,3
hdr=[c.value for c in wb[ASHEET][HROW]]
col={h:i for i,h in enumerate(hdr) if h}
def idx2a(k): return idx_map.get(str(k).strip(), k)
def rows(asset,site,code,ocr_d):
    hier=f"POWTR \\ {code}"
    lst=[]
    for k,v in ocr_d.items():
        if k in ('error','raw_text'): continue
        attr=idx2a(k)
        m=re.search(r'\((.*?)\)\s*$',attr); unit=m.group(1).strip() if m else ''
        r=['']*len(hdr); r[col['ASSETNUM']]=asset; r[col['SITEID']]=site
        r[col['HIERARCHYPATH']]=hier; r[col['ASSETSPEC.\nASSETATTRID']]=attr
        r[col['ASSETSPEC.ALNVALUE']]=v; r[col['ASSETSPEC.MEASUREUNITID']]=unit
        lst.append(r)
    return lst

# ───────────── 6. RUN ─────────────
if st.button('🚀 Run') and api_key and imgs and not pam_df.empty:
    ws=wb[ASHEET]
    if ws.max_row>=DSTART: ws.delete_rows(DSTART, ws.max_row-DSTART+1)
    res=[]; prog=st.progress(0.)
    for i,im in enumerate(imgs,1):
        prog.progress(i/len(imgs))
        loc=loc_map.get(im.name,'').strip()
        if not loc: st.warning(f'{im.name} ไม่มี Location'); continue
        oc=ocr(api_key, enc(im), prompt)
        code, kv, cand=powtr(oc if isinstance(oc,dict) else {})
        debug(i, oc, kv, cand)
        pam_cls=pam_df.loc[pam_df[loc_col]==loc,'Classification'].iat[0] \
                if loc in pam_df[loc_col].values and 'Classification' in pam_df else ''
        res.append({'Image':im.name,'Asset':loc,'POWTR(OCR)':code,
                    'Classification(PAM)':pam_cls,'Match?':code==pam_cls})
        for r in rows(loc, 'SBK0', code, oc): ws.append(r)
    st.dataframe(pd.DataFrame(res))
    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    st.download_button('⬇️ download', buf, 'MxLoader_POWTR_Result.xlsm',
                       'application/vnd.ms-excel.sheet.macroEnabled.12')
