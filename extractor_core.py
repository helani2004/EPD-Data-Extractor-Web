"""Local EPD extraction app. Run python app.py. No PDFs are stored or deleted."""
import hashlib,json,re,sqlite3,threading,queue,math,os,shutil,subprocess,tempfile
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Alignment,Font
import fitz
from engine import inspect_pdf,clean,numeric,positioned_lines
from ai_review import review_pdf
BASE=Path(__file__).resolve().parent
TEMPLATE=BASE/'EPDinputFormat.xlsx'
GROUPS=['A1-A3','C1','C2','C3','C4','D']
CODES=['GWP-total','GWP-fossil','GWP-biogenic','GWP-luluc','ODP','AP','EP-freshwater','EP-marine','EP-terrestrial','POCP','ADP-minerals','ADP-fossil','WDP']
STARTS=[28,46,59,72,85,98]

def headers():
 w=load_workbook(TEMPLATE);h=[w.active.cell(2,c).value for c in range(1,112)];w.close();return h
HEADERS=headers()
FIELDS=list(dict.fromkeys(h for h in HEADERS if h is not None))
IMPACT_FIELDS=HEADERS[27:40]
SKIP={'url','Material Category','EPD type(Product Specific/Industry-Average)','Ꟙ-% for service life calculation',
 'P (fraction to calculate maintenance period)','Ꝋ=repair factor of ith material (%)','Fu-Reuse % in raw material',
 'Fr-Recycle % in raw material','Fb-Bio-based /Renewable ,Non renewable % in raw material','ꟕ-fraction of waste produced during recycling'}

def code(label):
 s=clean(label).lower().replace('–','-').replace('−','-');compact=re.sub(r'\s+','',s)
 for c in CODES[:4]:
  if c.lower() in compact:return c
 if 'climatechange-total' in compact:return 'GWP-total'
 if 'ozonedepletion' in compact or re.search(r'\bodp\d*\b',s):return 'ODP'
 if 'acidification' in compact or re.search(r'\bap\d*\b',s):return 'AP'
 for c in ['EP-freshwater','EP-marine','EP-terrestrial']:
  if c.lower() in compact:return c
 if 'pocp' in compact or 'photochemicalozone' in compact:return 'POCP'
 if re.fullmatch(r'adpe\d*',compact) or ('adp' in compact or 'abiotic' in compact) and ('mineral' in compact or 'metal' in compact):return 'ADP-minerals'
 if re.fullmatch(r'adpf\d*',compact) or ('adp' in compact or 'abiotic' in compact) and ('fossil' in compact):return 'ADP-fossil'
 if 'wateruse' in compact or re.search(r'\bwdp\d*\b',s):return 'WDP'
 return None

def num(text):
 s=clean(text).replace('−','-')
 # No unqualified comma interpretation: decimal commas require a review.
 if ',' in s:return None
 m=re.fullmatch(r'([-+]?\d+(?:\.\d+)?(?:[Ee][-+]?\d+)?)\s*(kg|g|t|tonnes?|years?|%)?',s,re.I)
 return float(m.group(1)) if m else None

def ocr_executable():
 candidates=[os.environ.get('EPD_TESSERACT_EXE'),shutil.which('tesseract'),r'C:\Program Files\Tesseract-OCR\tesseract.exe',r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe']
 return next((p for p in candidates if p and Path(p).is_file()),None)

def scanned_pages(path):
 pages=[]
 with fitz.open(path) as doc:
  for i,page in enumerate(doc):
   words=page.get_text('words');area=page.rect.get_area()
   image_regions=[fitz.Rect(info['bbox']) for info in page.get_image_info()]
   image_scan=any(box.get_area()>.5*area and sum(1 for w in words if box.contains(fitz.Point((w[0]+w[2])/2,(w[1]+w[3])/2)))<8 for box in image_regions)
   if image_scan or (not words and image_regions):pages.append(i)
 return pages

def process(path):
 """OCR image-only pages into a temporary searchable PDF, then use the reader."""
 path=Path(path);pages=scanned_pages(path)
 if not pages:return process_readable(path)
 exe=ocr_executable()
 if not exe:
  r=process_readable(path);r['warnings'].insert(0,'Scanned pages detected, but Tesseract was not found. Install it in C:\\Program Files\\Tesseract-OCR or set EPD_TESSERACT_EXE.');return r
 notes=[];completed=[]
 with tempfile.TemporaryDirectory(prefix='epd_ocr_') as folder:
  folder=Path(folder)
  with fitz.open(path) as source,fitz.open() as output:
   for i,page in enumerate(source):
    if i not in pages:output.insert_pdf(source,from_page=i,to_page=i);continue
    image_path=folder/f'page_{i+1}.png';stem=folder/f'ocr_{i+1}'
    try:
     page.get_pixmap(dpi=300,alpha=False).save(image_path)
     command=[str(exe),str(image_path),str(stem),'-l','eng','--dpi','300','pdf']
     run=subprocess.run(command,capture_output=True,text=True,timeout=180,creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
     if run.returncode or not stem.with_suffix('.pdf').is_file():raise RuntimeError((run.stderr or 'No searchable PDF generated')[-1200:])
     with fitz.open(stem.with_suffix('.pdf')) as recognised:
      output.insert_pdf(recognised);completed.append(i+1)
    except Exception as error:
     output.insert_pdf(source,from_page=i,to_page=i);notes.append(f'Page {i+1}: OCR failed; original page retained. {error}')
   readable=folder/'searchable.pdf';output.save(readable)
  r=process_readable(readable)
  r['file']=path.name
  for value in r['values']:value['file']=path.name
  r['ocr_pages']=completed
  r['warnings']=notes+([f'OCR used on pages {completed}. Verify decimal points, signs, exponents and table alignment. Cell coordinates refer to the temporary searchable copy.'] if completed else [])+r['warnings']
  return r

def process_readable(path):
 r=inspect_pdf(path);data={h:'' for h in FIELDS};data.update({f'__{m}__{f}':'' for m in GROUPS for f in IMPACT_FIELDS});evidence=[];issues=list(r['warnings'])
 def put(field,value,page,raw,method,status='extracted'):
  evidence.append({'field':field,'value':value,'page':page,'raw':raw,'method':method,'status':status})
  if status.startswith('review required'):
   issues.append(field+': '+status+'; candidate retained in evidence only.');return
  if value is not None and value!='':
   old=data.get(field,'')
   if old!='' and old!=value:
    if field=='Product name':
     values=old.split(' | ')
     if value not in values:data[field]=old+' | '+str(value)
    else:issues.append(f'{field}: conflicting candidates; retained first value. Review page {page}.')
   else:data[field]=value
 for v in r['details']:
  if v['field']=='Version date':
   evidence.append({**v,'method':'explicit version date; not substituted for publication date'});continue
  if v['field'] in data:put(v['field'],v['value'],v['page'],v['raw'],v['method'],v.get('status','extracted'))
 # Map indicators only AFTER module columns and values were recovered by geometry.
 mapped={};conflicted=set()
 for v in r['values']:
  c=code(v['indicator']);m=v['module']
  if c and m in GROUPS:
   u=re.sub(r'[^a-z0-9]','',v['unit'].lower())
   legacy=v.get('basis','').startswith('legacy') or (c=='AP' and 'so2' in u) or (c=='POCP' and ('c2h4' in u or 'ethylene' in u))
   if legacy:
    issues.append(f"Page {v['page']}: {c} uses legacy units ({v['unit']}); retained in Impact audit, excluded from the EN 15804+A2 template.");continue
   field=IMPACT_FIELDS[CODES.index(c)];key=f'__{m}__{field}'
   if key in conflicted:continue
   if key in mapped and (mapped[key]['value']!=v['value'] or mapped[key]['status']!=v['status']):
    issues.append(f'{m} {c}: conflicting impact tables; choose appropriate product/scenario before use.')
    data[key]='';conflicted.add(key);continue
   mapped[key]=v
   data[key]=v['value'] if v['status']=='numeric' else ''
   evidence.append({'field':key,'value':v['value'],'page':v['page'],'raw':v['raw'],'method':f"table {v['table']}, cell {v['bbox']}",'status':v['status']})
 total=f'__A1-A3__{IMPACT_FIELDS[0]}'
 if not isinstance(data.get(total),(int,float)):
  issues.append('A1-A3 GWP Total unresolved: recovered sections are shown for review. Check source evidence and enter the correct value before adding this record.')
 # Non-impact sections use actual table cells. Read through fitz to avoid hidden overlapping text.
 cell_tables=[(info['page'],t.get('visible_rows',t['rows'])) for info in r['pages'] for t in info['tables']]
 with fitz.open(path) as doc:
  extras=[]
  labels={'declared unit':'Functional unit/Declared Unit','declared unit and reference flow':'Functional unit/Declared Unit',
          'mass per declared unit':'Mass per declared unit (tonne/Declared unit)',
          'reference service life':'TL- Technical Life (years)','technical life':'TL- Technical Life (years)',
          'bulk density':'Bulk desnsity','cpc code':'CPC code','un cpc code':'CPC code',
          'product category rules (pcr)':'PCR','product category rules':'PCR',
          'construction wastage':'Construction material Wastage %','installation wastage':'Construction material Wastage %'}
  composition={};ratios={};tech=[];mass_conversion=[];reuse_candidates=[]
  for page,rows in cell_tables:
   for ri,row in enumerate(rows):
    cells=[clean(x) for x in row];non=[(i,x) for i,x in enumerate(cells) if x]
    if not non:continue
    first=non[0][1];low=first.lower().rstrip(':')
    if low in labels:
     val=' '.join(x for _,x in non[1:])
     if not val and ri+1<len(rows):val=clean(' '.join(x or '' for x in rows[ri+1]))
     if val and val.lower() not in {'-','n/a','declared unit'}:extras.append((labels[low],val,page,'table label/value'))
    if low=='conversion factor to mass' and len(non)>1:mass_conversion.append((non[1][1],page))
    if 'average electricity used during installation' in low and len(non)>1:
     v=num(non[-1][1]);
     if v is not None:put(HEADERS[110],v,page,non[-1][1],'installation scenario table')
    if 'kg for re-use' in ' '.join(cells).lower() or low=='recovery system, for re-use':
     vals=[num(x) for _,x in non[1:] if num(x) is not None]
     if vals:reuse_candidates.append((vals[-1],page,first))
   for hi,row in enumerate(rows):
    h=[clean(x).lower() for x in row]
    if 'level' in h and 'name of criterion' in h:
     vi=h.index('level')
     for rr in rows[hi+1:]:
      if vi<len(rr) and clean(rr[vi]):tech.append((clean(rr[vi]),page))
    elif 'property' in h and 'value' in h and 'regulation' in h:
     vi=h.index('value');ni=h.index('property');pending=''
     for rr in rows[hi+1:]:
      name=clean(rr[ni]);value=clean(rr[vi]);
      if name:pending=name
      if value and pending:tech.append((pending+': '+value,page));pending=''
    if 'product material' in h and any('weight' in x and 'kg' in x for x in h):
     wi=next(i for i,x in enumerate(h) if 'weight' in x and 'kg' in x)
     for rr in rows[hi+1:]:
      name=clean(rr[0]);v=num(clean(rr[wi]))
      if name and v is not None:composition[name.lower()]=(v,page)
   for row in rows:
    non=[clean(x) for x in row if clean(x)]
    if len(non)>=2 and re.fullmatch(r'\d+(?:\.\d+)?:\d+(?:\.\d+)?',non[-1]):
     a,b=map(float,non[-1].split(':'))
     if abs(a+b-100)<1e-6:ratios[non[-2].lower()]=(a/100,b/100,page)
  # Exact section labels in borderless layouts. Use same-baseline text and bounded
  # vertical regions, not a search over the entire document for a nearby number.
  for pi,pg in enumerate(doc):
   lines=positioned_lines(pg)
   for x in lines:
    txt=x['text'];low=txt.lower();field=None;raw=''
    for lab,f in labels.items():
     if low==lab or low==lab+':':field=f;break
     if low.startswith(lab+':'):field=f;raw=txt.split(':',1)[1].strip();break
    if low.startswith('conversion factor to mass:'):
     raw=txt.split(':',1)[1].strip()
     if not raw:
      cand=[v for v in lines if 0<v['bbox'][1]-x['bbox'][1]<25 and abs(v['bbox'][0]-x['bbox'][0])<15]
      if cand:raw=min(cand,key=lambda v:v['bbox'][1])['text']
     if raw:mass_conversion.append((raw,pi+1))
    if field and not raw:
     b=x['bbox'];right=[v for v in lines if 0<v['bbox'][0]-b[2]<180 and abs(v['bbox'][1]-b[1])<3]
     if right:raw=min(right,key=lambda v:v['bbox'][0])['text']
     else:
      below=[v for v in lines if 0<v['bbox'][1]-b[1]<(40 if field=='Functional unit/Declared Unit' else 25) and abs(v['bbox'][0]-b[0])<15]
      if below:raw=min(below,key=lambda v:v['bbox'][1])['text']
    if field=='CPC code' and raw:
     bottom=x['bbox'][3]
     for continuation in lines:
      cb=continuation['bbox']
      if cb[1]>bottom and cb[1]-bottom<5 and abs(cb[0]-x['bbox'][0])<3:
       raw+=' '+continuation['text'];bottom=cb[3]
    if field and raw:extras.append((field,raw,pi+1,'positioned section label/value'))
   # Product standards are a bounded heading region in the same column.
   for x in lines:
    if x['text'].upper().rstrip(':')=='PRODUCT STANDARDS':
     b=x['bbox'];following=[v for v in lines if b[1]<v['bbox'][1]<b[1]+100 and abs(v['bbox'][0]-b[0])<15]
     selected=[]
     for v in following:
      if v['text'].upper().startswith(('TECHNICAL SPEC','PRODUCT ','PACKAGING')):break
      selected.append(v['text'])
     if selected:put('Product Standards',' '.join(selected),pi+1,' '.join(selected),'bounded standards section')
   # Evidence for installation waste must occur inside the installation section.
   for block in pg.get_text('blocks'):
    text=clean(block[4])
    if ('installation' in text.lower() or 'construction site' in text.lower()) and ('wastage' in text.lower() or 'waste rate' in text.lower()):
     issues.append(f'Page {pi+1}: installation/waste narrative needs review; no percentage inferred from paragraph.')
  # PCR tables may have the heading and actual PCR declaration in separate rows.
  for page,rows in cell_tables:
   if any('product category rules' in clean(x).lower() for row in rows for x in row):
    for row in rows:
     for raw in row:
      raw=clean(raw)
      if re.match(r'^(?:PCR\s+\d{4}:|c-PCR-\d+)',raw,re.I):put('PCR',raw,page,raw,'PCR declaration cell')
  # For text-only technical standards use explicit EN / ASTM records in product sections.
  if not data.get('Product Standards'):
   for pi,pg in enumerate(doc):
    lines=positioned_lines(pg)
    if any(x['text'].upper()=='PRODUCT INFORMATION' for x in lines):
     vv=[x['text'] for x in lines if re.match(r'^(?:EN[- ]\d|ASTM\s+[A-Z]?\d)',x['text'])]
     if vv:put('Product Standards','; '.join(vv),pi+1,'; '.join(vv),'explicit standards in product section')
  # Persist the recovered non-impact tables as audit data; no PDF retention required.
  r['section_tables']=[{'page':p,'rows':rows} for p,rows in cell_tables]
 for field,raw,page,method in extras:
  if field=='Functional unit/Declared Unit':
   if re.search(r'(?:\b\d+(?:[.,]\d+)?\s*(?:metric\s+)?(?:tonnes?\b|kg\b|g\b|t\b|m[²³23]\b|pieces?\b|units?\b))|Mass\s*\(kg\)',raw,re.I):put(field,raw,page,raw,method)
   else:evidence.append({'field':field,'value':None,'page':page,'raw':raw,'method':method,'status':'unresolved candidate; not a declared measurement'})
  elif field=='Mass per declared unit (tonne/Declared unit)':
   m=re.fullmatch(r'([0-9.]+)\s*(kg|g|t|tonne|tonnes)',raw,re.I)
   if m:put(field,float(m[1])*{'kg':.001,'g':.000001,'t':1,'tonne':1,'tonnes':1}[m[2].lower()],page,raw,'unit conversion to tonnes')
  elif field=='TL- Technical Life (years)':
   m=re.fullmatch(r'([0-9.]+)\s*years?',raw,re.I)
   if m:put(field,float(m[1]),page,raw,method)
  elif field=='Bulk desnsity':
   m=re.fullmatch(r'([0-9.]+)\s*(kg/m[³3]|g/cm[³3])',raw,re.I)
   if m:put(field,float(m[1]),page,raw,method);put('Unit of bulk density',m[2],page,raw,method)
  elif field=='PCR':
   if re.search(r'(?:PCR\s+\d{4}:|c-PCR-\d+)',raw,re.I):
    if data.get(field) and raw not in data[field]:data[field]+='; '+raw
    else:put(field,raw,page,raw,method)
  elif field=='Construction material Wastage %':
   m=re.fullmatch(r'([0-9.]+)\s*%',raw)
   if m:put(field,float(m[1]),page,raw,method)
  else:put(field,raw,page,raw,method)
 unit=data.get('Functional unit/Declared Unit','')
 if not data.get('Mass per declared unit (tonne/Declared unit)'):
  explicit=re.search(r'\b(\d+(?:\.\d+)?)\s*(metric\s+tonnes?|tonnes?|kg|g|t)\b(?!\s*/)',unit,re.I)
  if explicit:
   amount=float(explicit[1]);u=explicit[2].lower();factor=.001 if u=='kg' else .000001 if u=='g' else 1
   src=next((e for e in evidence if e['field']=='Functional unit/Declared Unit' and e.get('value')==unit),{})
   put('Mass per declared unit (tonne/Declared unit)',amount*factor,src.get('page'),unit,'explicit declared mass converted to tonnes')
 if not data.get('Mass per declared unit (tonne/Declared unit)'):
  m=re.search(r'Mass\s*\(kg\),\s*([0-9.]+)',unit,re.I)
  if m:
   source=next((e for e in evidence if e['field']=='Functional unit/Declared Unit' and e['value']==unit),{})
   put('Mass per declared unit (tonne/Declared unit)',float(m[1])*.001,source.get('page'),unit,'declared reference flow kg to tonnes')
  else:
   for raw,page in mass_conversion:
    m=re.fullmatch(r'([0-9.]+)\s*kg/m\s*[²2]',raw)
    if m and re.match(r'^1\s*m[²2]',unit):put('Mass per declared unit (tonne/Declared unit)',float(m[1])*.001,page,raw,'kg/m² × explicit 1 m² declared unit, converted to tonnes')
 if tech:put('Technical specification',', '.join(dict.fromkeys(v for v,p in tech)),tech[0][1],' | '.join(v for v,p in tech),'technical table criteria/values; test methods excluded')
 mass=data.get('Mass per declared unit (tonne/Declared unit)');masskg=mass*1000 if isinstance(mass,(float,int)) else None
 if reuse_candidates and masskg:
  values={v for v,p,l in reuse_candidates}
  if len(values)==1 and 0<=next(iter(values))<=masskg:
   v,page,lab=reuse_candidates[0];put('Cu-the fraction of material collected for reuse',v/masskg,page,str(v),'explicit kg for reuse / declared kg')
  else:issues.append('Reuse scenario table contains conflicting values; fraction left blank. Review default and additional scenarios.')
 if composition and ratios:
  covered=0;recycle=0;landfill=0;unmatched=[]
  for name,(weight,page) in composition.items():
   normalized=name.replace(' profiles','')
   ratio=ratios.get(normalized)
   if ratio:covered+=weight;recycle+=weight*ratio[0];landfill+=weight*ratio[1]
   else:unmatched.append(name)
  if unmatched:
   issues.append('End-of-life fractions: no explicit scenario for '+', '.join(unmatched)+'. Overall landfill/recycling fractions left blank; partial material totals are not treated as complete.')
  elif masskg and abs(covered-masskg)/masskg<.01:
   put('Landfill fraction',landfill/covered,page,str(ratios),'mass-weighted explicit material scenarios')
   put('Cr-the fractions of material collected for recycle',recycle/covered,page,str(ratios),'mass-weighted explicit material scenarios')
 if len(set(v['value'] for v in r['details'] if v['field']=='Product name'))>1:
  issues.append('Multiple products share this declaration. Product names retained together; impact values are not duplicated or assigned to individual products without an explicit table association.')
 r['data']=data;r['evidence']=evidence
 for field in FIELDS:
  if field in IMPACT_FIELDS or field in SKIP:continue
  if data.get(field,'')=='':issues.append(field+': not recovered; review required or not stated in source.')
 for m in GROUPS:
  for field in IMPACT_FIELDS:
   key=f'__{m}__{field}'
   if key not in mapped:issues.append(m+' '+field+': no matching indicator cell recovered.')
 issues.append('Extraction uses layout plus explicit labels. Scanned PDFs and unfamiliar layouts are not guaranteed. All values remain reviewable against the source.')
 r['warnings']=list(dict.fromkeys(issues));return r

def save_excel(records,destination):
 w=load_workbook(TEMPLATE);ws=w.active
 if ws.max_row>=3:ws.delete_rows(3,ws.max_row-2)
 for rec in records:
  d=rec['result']['data'];row=[d.get(h,'') for h in HEADERS]
  for m,start in zip(GROUPS,STARTS):
   for i,field in enumerate(IMPACT_FIELDS):row[start-1+i]=d.get(f'__{m}__{field}','')
  ws.append(row)
 index=w.create_sheet('Record index');index.append(['ID','Template row','Source filename','EPD number','Created','Reviewed'])
 audit=w.create_sheet('Extraction review');audit.append(['ID','Issue'])
 ev=w.create_sheet('Source evidence');ev.append(['ID','Field','Value','Page','Raw','Method','Status'])
 impacts=w.create_sheet('Impact audit');impacts.append(['ID','Page','Indicator','Unit','Module','Raw','Value','Status'])
 tables=w.create_sheet('Section tables');tables.append(['ID','Page','Table row (raw JSON)'])
 for i,rec in enumerate(records,3):
  r=rec['result'];index.append([f"{rec['id']:05d}",i,r['file'],r['data'].get('EPD Number'),rec.get('created',''),bool(r.get('reviewed'))])
  for warn in r['warnings']:audit.append([f"{rec['id']:05d}",warn])
  for x in r['evidence']:ev.append([f"{rec['id']:05d}",x.get('field'),x.get('value'),x.get('page'),x.get('raw'),x.get('method'),x.get('status')])
  for v in r['values']:impacts.append([f"{rec['id']:05d}",v['page'],v['indicator'],v['unit'],v['module'],v['raw'],v['value'],v['status']])
  for t in r.get('section_tables',[]):
   for row in t['rows']:tables.append([f"{rec['id']:05d}",t['page'],json.dumps(row,ensure_ascii=False)])
 ai=w.create_sheet('AI review');ai.append(['ID','Field','Code value','AI proposal','Decision','Page','Raw evidence','Unit','Context','Calculation','Product','Scenario','Declared unit','Basis'])
 for rec in records:
  for a in rec['result'].get('ai_review',[]):ai.append([str(rec['id'])]+[a.get(k,'') for k in ['field','old','proposed','decision','page','raw','unit','context','calculation','product','scenario','declared_unit','basis']])
 for sheet in [index,audit,ev,impacts,tables,ai]:
  sheet.freeze_panes='A2';sheet.auto_filter.ref=sheet.dimensions
  for c in sheet[1]:c.font=Font(bold=True)
  for col in sheet.columns:sheet.column_dimensions[col[0].column_letter].width=40
 audit.column_dimensions['B'].width=110
 w.save(destination)
