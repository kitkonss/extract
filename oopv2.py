# app.py – OCR ▸ POWTR‑CODE ▸ MxLoader (.xlsm)  2024‑06‑22
import streamlit as st, pandas as pd, requests, json, base64, io, re, os
from openpyxl import load_workbook
from pathlib import Path

# ────────── fixed files ──────────
TPL  = Path('Template-MxLoader-Classification POW-TR.xlsm')
ATTR = Path('ATTRIBUTE.xlsx')
if not TPL.exists(): TPL = st.file_uploader('template .xlsm', ['xlsm']); st.stop() if TPL is None else None
wb = load_workbook(TPL, keep_vba=True)
if not ATTR.exists(): ATTR = st.file_uploader('ATTRIBUTE.xlsx', ['xlsx','xls']); st.stop() if ATTR is None else None

# ────────── attribute list & prompt ──────────
attr_df  = pd.read_excel(ATTR, header=None)
ATTRS    = [str(a).strip() for a in attr_df[0] if str(a).strip()]
IDX_MAP  = {str(i+1):a for i,a in enumerate(ATTRS,1)}

PROMPT = (
"คืน JSON เช่น\n"
"{ \"HIGH_SIDE_VOLTAGE_KV\": 230, \"PHASE\": 3,\n"
"  \"COOLING_TYPE\": \"ONAN / ONAF\", \"TAP_CHANGER\": \"OFF‑CIRCUIT\",\n"
"  \"VECTOR_GROUP\": \"YNd1\" }\n"
"ค่าไม่พบให้ใส่ \"\" (**อย่าใช้ BIL / AC withstand**)\n\n"
"และคืนค่าต่อไปนี้ (ใช้ key เป็นเลข index):\n" +
'\n'.join(f"{i}: {a}" for i,a in enumerate(ATTRS,1))
)

# ────────── UI ──────────
st.title('⚡ Transformer OCR → POWTR‑CODE → MxLoader (.xlsm)')
pam_xl = st.file_uploader('PAM.xlsx', ['xlsx','xls'])
imgs   = st.file_uploader('Nameplate images', ['jpg','jpeg','png'], accept_multiple_files=True)
api    = os.getenv('GEMINI_API_KEY') or st.text_input('Gemini API‑key', type='password')

if pam_xl is not None:
    pam = pd.read_excel(pam_xl)
    loc_col = st.selectbox('Location column in PAM', pam.columns,
                           index=list(pam.columns).index('Location') if 'Location' in pam else 0)
else:
    pam = pd.DataFrame(); loc_col=''

loc_map={}
if imgs and not pam.empty:
    st.markdown('**ใส่ AssetNUM / Location ให้แต่ละรูป**')
    for im in imgs: loc_map[im.name] = st.text_input(im.name, key=im.name)

site_default = st.text_input('SITEID', 'SBK0')

# ────────── OCR helper ──────────
def b64(f): return base64.b64encode(f.getvalue()).decode()
def gemini(img64):
    r=requests.post(
        f"https://generativelanguage.googleapis.com/v1/models/gemini-2.0-flash:generateContent?key={api}",
        json={"contents":[{"parts":[{"text":PROMPT},
                                    {"inline_data":{"mime_type":"image/jpeg","data":img64}}]}],
              "generation_config":{"temperature":0.2,"max_output_tokens":4096}})
    if r.status_code!=200: return {"error":r.text}
    t=r.json()['candidates'][0]['content']['parts'][0]['text']
    try:return json.loads(t[t.find('{'):t.rfind('}')+1])
    except: return {"raw_text":t}

# ────────── POWTR logic ──────────
oil_kw={'OIL','ONAN','ONAF','OFAF','OFWF','OA','OF','ON','ONO','OFA'}
def detect_kv(d):
    kvs=[]; pat=re.compile(r'(\d{2,7}(?:[ ,]\d{3})*(?:[.,]\d+)?)\s*(kV|KV|kv|V|v)?')
    for txt in map(str,d.values()):
        for part in str(txt).split('|'):               # 1) แยก "241 500|235 750"
            up=part.upper()
            if any(x in up for x in ('BIL','/ AC',' AC ','IMPULSE','LIGHTNING')): continue
            for raw,unit in pat.findall(part):
                num=raw.replace(' ','').replace(',','')
                try: val=float(num.replace(',','.'))
                except: continue
                if unit and unit.lower().startswith('k'):
                    kv=val
                elif unit and unit.lower().startswith('v'):
                    kv=val/1000
                else:                                   # ไม่มีหน่วย
                    if val>1000: kv=val/1000           # assume Volt
                    elif val>=2: kv=val                # อาจเป็น kV
                    else: continue                     # ทิ้งเลขเล็ก
                if 2<=kv<=765: kvs.append(kv)
    return (max(kvs) if kvs else None, kvs)

def powtr(d):
    ph=str(d.get('PHASE','3')).replace('.0','')
    kv, cand = detect_kv(d)
    v='-' if kv is None else ('E' if kv>=345 else 'H' if kv>=100 else 'M' if kv>=1 else 'L')
    cool=str(d.get('COOLING_TYPE','')).upper()
    t='O' if any(k in cool for k in oil_kw) else 'D'
    tap='O' if 'ON' in str(d.get('TAP_CHANGER','')).upper() else 'F'
    return f'POWTR-{ph}{v}{t}{tap}', kv, cand

# ────────── MxLoader helpers ──────────
ASHEET,HROW,DSTART='AssetAttr',2,3
hdr=[c.value for c in wb[ASHEET][HROW]]
col={h:i for i,h in enumerate(hdr) if h}
def full_dict(ocr): d={a:'' for a in ATTRS}; d.update(ocr); return d
def blank(v,a): return '' if v in {'','-'} or str(v).upper()==a.upper() else v
def rows(asset,site,code,d):
    hier=f"POWTR \\ {code}"; out=[]
    for a in ATTRS:
        v=blank(d.get(a,''),a)
        u=re.search(r'\((.*?)\)\s*$',a); unit=u.group(1).strip() if u else ''
        r=['']*len(hdr)
        r[col['ASSETNUM']],r[col['SITEID']],r[col['HIERARCHYPATH']]=asset,site,hier
        r[col['ASSETSPEC.\nASSETATTRID']]=a; r[col['ASSETSPEC.ALNVALUE']]=v
        r[col['ASSETSPEC.MEASUREUNITID']]=unit
        out.append(r)
    return out
def debug(i,d,kv,cand):
    with st.expander(f'Debug #{i+1}'):
        st.json(d); st.write('kV cand →',cand); st.write('chosen →',kv)

# ────────── RUN ──────────
if st.button('🚀 Run') and api and imgs and not pam.empty:
    ws=wb[ASHEET];  ws.delete_rows(DSTART, ws.max_row-DSTART+1) if ws.max_row>=DSTART else None
    result=[]; prog=st.progress(0.)
    for i,im in enumerate(imgs,1):
        prog.progress(i/len(imgs))
        loc=loc_map.get(im.name,'').strip()
        if not loc: st.warning(f'{im.name} ไม่มี Location'); continue
        raw   = gemini(b64(im))
        data  = full_dict(raw if isinstance(raw,dict) else {})
        code, kv, cand = powtr(data)
        debug(i,data,kv,cand)
        pam_cls = pam.loc[pam[loc_col]==loc,'Classification'].iat[0] if loc in pam[loc_col].values else ''
        result.append({'Image':im.name,'Asset':loc,'POWTR(OCR)':code,
                       'Classification(PAM)':pam_cls,'Match?':code==pam_cls})
        for r in rows(loc,site_default,code,data): ws.append(r)

    st.dataframe(pd.DataFrame(result))
    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    st.download_button('Download MxLoader file', buf,
                       'MxLoader_POWTR_Result.xlsm',
                       'application/vnd.ms-excel.sheet.macroEnabled.12')
