"""First section: inspect PDF geometry and extract lifecycle impact tables.
Install: python -m pip install pymupdf pdfplumber openpyxl
Run: python EPD_Section_Inspector.py
CLI: python EPD_Section_Inspector.py file.pdf --output results.xlsx
Original PDFs are never deleted. Includes impact tables and document details.
Document labels are matched exactly and values associated using coordinates.
Missing details are not invented; version dates remain separate from publication dates.
Multiple products remain separate candidates. No database integration yet.
"""
import argparse
import json
import logging
import re
from pathlib import Path
import fitz
import pdfplumber
from openpyxl import Workbook
logging.getLogger('pdfminer').setLevel(logging.ERROR)

MODULES = ['A1-A3','A1','A2','A3','A4','A5','B1','B2','B3','B4','B5','B6','B7','C1','C2','C3','C4','D']

def clean(value):
    return re.sub(r'\s+', ' ', str(value or '')).strip()

def module(value):
    s=clean(value).upper().replace('\x00','-').replace('–','-').replace('−','-').replace('—','-')
    s=re.sub(r'\bMODULES?\b','',s)
    s=re.sub(r'\(?\s*(?:TOTAL|TOT\.)\s*\)?','',s)
    s=re.sub(r'[\s()]','',s)
    if s in {'A1+A2+A3','A1/A2/A3'}:s='A1-A3'
    return s if s in MODULES else None

def numeric(raw):
    s=clean(raw).replace('−','-').replace(' ','')
    # A single decimal comma in scientific notation is explicit decimal syntax.
    # Mixed separators and standalone comma numbers remain unresolved.
    if re.fullmatch(r'[-+]?\d+,\d+[Ee][-+]?\d+',s):s=s.replace(',','.')
    if re.fullmatch(r'[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][-+]?\d+)?',s):
        return float(s), 'numeric'
    if s.upper() in {'ND','MND','MNR','NA','N/A','-','NOTDECLARED','NOTRELEVANT'}:
        return None, 'not declared/relevant'
    return None, 'blank' if not s else 'review required'


DETAIL_LABELS = {
    'EPD Program Operator': ['programme operator','program operator'],
    'Regional Hub': ['regional hub'],
    'EPD Number': ['epd registration number','registration number','epd number'],
    'Publication date': ['publication date','date of publication'],
    'Version date': ['version date'],
    'Valid until': ['valid until','validity date'],
    'Manufacturer/EPD owner': ['epd owner','owner of the epd','manufacturer'],
    'Product name': ['product name'],
    'Geographical scope': ['geographical scope','geographical scope(s)'],
}

def positioned_lines(page):
    # Join words on the same baseline, but keep distant columns separate.
    words=sorted(page.get_text('words'),key=lambda w:(round(w[1],1),w[0]))
    bands=[]
    for w in words:
        band=next((b for b in reversed(bands[-5:]) if abs(b[0]-w[1])<2),None)
        if band is None: bands.append([w[1],[w]])
        else:band[1].append(w)
    output=[]
    for _,items in bands:
        items.sort(key=lambda w:w[0]);groups=[]
        for w in items:
            if not groups or w[0]-groups[-1][-1][2]>22:groups.append([w])
            else:groups[-1].append(w)
        for g in groups:
            output.append({'text':clean(' '.join(w[4] for w in g)),
                           'bbox':[min(w[0] for w in g),min(w[1] for w in g),max(w[2] for w in g),max(w[3] for w in g)]})
    return sorted(output,key=lambda x:(x['bbox'][1],x['bbox'][0]))

def extract_details(path):
    records=[]
    aliases={a:field for field,aa in DETAIL_LABELS.items() for a in aa}
    def identify(text):
        text=clean(text).lower().rstrip(':').strip()
        if text in aliases:return aliases[text],''
        for a in sorted(aliases,key=len,reverse=True):
            if text.startswith(a+':'):
                return aliases[a],clean(text[len(a)+1:])
        return None,''
    with fitz.open(path) as doc:
        for pi,page in enumerate(doc):
            if pi >= 8: break  # Document details section; avoid back-cover contact labels.
            lines=positioned_lines(page)
            for label in lines:
                field,inline=identify(label['text'])
                if not field:continue
                box=label['bbox']; chosen=[];method=''
                if inline:
                    # Use original capitalization, not identify()'s normalized copy.
                    chosen=[{'text':clean(label['text'].split(':',1)[1]),'bbox':box}];method='same-line label/value'
                else:
                    right=[x for x in lines if 2 <= x['bbox'][0]-box[2] <= 180 and abs(x['bbox'][1]-box[1])<3]
                    if right:
                        chosen=[min(right,key=lambda x:x['bbox'][0])];method='aligned value column'
                    else:
                        below=[x for x in lines if 0<x['bbox'][1]-box[1]<65 and
                               (abs(x['bbox'][0]-box[0])<12 or abs(x['bbox'][2]-box[2])<12)]
                        below.sort(key=lambda x:x['bbox'][1])
                        if below and not identify(below[0]['text'])[0]:chosen=[below[0]];method='value below label'
                if not chosen:continue
                if field=='Manufacturer/EPD owner' and method=='aligned value column':
                    b=chosen[0]['bbox']
                    cont=[x for x in lines if 0<x['bbox'][1]-b[1]<18 and abs(x['bbox'][0]-b[0])<5]
                    if cont:
                        nxt=min(cont,key=lambda x:x['bbox'][1])
                        # Stop at the next label row in the key column.
                        new_label=any(abs(x['bbox'][1]-nxt['bbox'][1])<3 and x['bbox'][0]<b[0]-10 and (identify(x['text'])[0] or clean(x['text']).lower() in {'address','contact details','website'}) for x in lines)
                        if not new_label:chosen.append(nxt)
                raw=' '.join(x['text'] for x in chosen);value=raw;status='extracted; review source'
                if field in {'Publication date','Version date','Valid until'}:
                    compact=re.sub(r'\s+','',raw)
                    try:
                        from datetime import date
                        value=date.fromisoformat(compact).isoformat()
                    except ValueError:status='review required: date format';value=raw
                if field=='EPD Number':value=re.sub(r'\s*-\s*','-',raw)
                records.append({'field':field,'value':value,'raw':raw,'page':pi+1,
                                'label':label['text'],'label_bbox':box,
                                'value_bbox':[x['bbox'] for x in chosen],'method':method,'status':status})
        # Cover-page "from" is an explicit ownership relationship, not a paragraph keyword.
        if True:  # Preserve explicit cover ownership evidence as an additional candidate.
            lines=positioned_lines(doc[0])
            for x in lines:
                if x['text'].lower()=='from':
                    b=x['bbox'];v=[v for v in lines if 0<v['bbox'][1]-b[1]<40 and abs(v['bbox'][0]-b[0])<12]
                    if v:
                        y=min(v,key=lambda v:v['bbox'][1]);records.append({'field':'Manufacturer/EPD owner','value':y['text'],'raw':y['text'],'page':1,'label':'from','label_bbox':b,'value_bbox':[y['bbox']],'method':'cover ownership block','status':'extracted; review source'})
                    break
    # Keep all occurrences, including multiple product names, with separate page evidence.
    missing=[f for f in DETAIL_LABELS if not any(r['field']==f for r in records)]
    return records,missing


def layout_fallback(page,page_number,filename):
    """Read borderless tables from aligned text objects, preserving cell rectangles."""
    bands=[]
    for w in sorted(page.get_text('words'),key=lambda w:(w[1],w[0])):
        band=next((b for b in reversed(bands[-8:]) if abs(b[0]-w[1])<3),None)
        if band is None:bands.append([w[1],[w]])
        else:band[1].append(w)
    for band in bands:band[1].sort(key=lambda w:w[0])
    headers=[]
    for i,(y,words) in enumerate(bands):
        units=[w for w in words if w[4].lower().rstrip(':') in {'unit','units'}]
        anchors={module(w[4]):(w[0]+w[2])/2 for w in words if module(w[4])}
        if units and anchors and all(x>units[0][0] for x in anchors.values()):headers.append((i,units[0],anchors))
    output=[]
    for ti,(hi,unitword,anchors) in enumerate(headers,1):
        end=headers[ti][0] if ti<len(headers) else len(bands)
        prior=clean(page.get_textbox(fitz.Rect(0,max(0,unitword[1]-220),page.rect.width,unitword[3])))
        legacy_context=bool(re.search(r'EN15804(?::2012)?\+A1',prior.upper().replace(' ','')))
        ordered=sorted(anchors.items(),key=lambda item:item[1]);ux=unitword[0]
        first=(unitword[0]+unitword[2]+2*ordered[0][1])/4
        for ri in range(hi+1,end):
            words=bands[ri][1]
            label=clean(' '.join(w[4] for w in words if w[2]<ux))
            unit=clean(' '.join(w[4] for w in words if ux<=((w[0]+w[2])/2)<first))
            if not label or not unit:continue
            cells=[]
            for ci,(m,x) in enumerate(ordered):
                left=first if ci==0 else (ordered[ci-1][1]+x)/2
                right=page.rect.width if ci==len(ordered)-1 else (x+ordered[ci+1][1])/2
                found=[w for w in words if left<=(w[0]+w[2])/2<right]
                raw=clean(' '.join(w[4] for w in found));v,status=numeric(raw)
                box=[min(w[0] for w in found),min(w[1] for w in found),max(w[2] for w in found),max(w[3] for w in found)] if found else None
                cells.append((m,raw,v,status,box))
            if not any(c[3]=='numeric' for c in cells):continue
            basis='legacy EN15804+A1' if legacy_context or 'SO2' in unit else 'unclassified/current'
            for m,raw,v,status,box in cells:
                output.append({'file':filename,'page':page_number,'table':f'layout-{ti}',
                    'indicator':label,'unit':unit,'module':m,'raw':raw,'value':v,'status':status,'bbox':box,'basis':basis,'method':'aligned text geometry fallback'})
    return output


def inspect_pdf(path):
    """Find tables by geometry, locate module headers, read cells by position.
    Table selection requires lifecycle headers plus a unit column, not a page heading.
    Labels are preserved as evidence rather than forced into an indicator dictionary.
    """
    path=Path(path)
    details, missing_details=extract_details(path)
    result={'file':path.name,'pages':[], 'values':[], 'warnings':[], 'details':details, 'missing_details':missing_details}
    with fitz.open(path) as doc, pdfplumber.open(path) as pdf:
        for page_index, p in enumerate(pdf.pages):
            fp=doc[page_index]
            info={'page':page_index+1,'width':p.width,'height':p.height,
                  'blocks':fp.get_text('blocks'), 'headings':[], 'tables':[]}
            for block in fp.get_text('dict')['blocks']:
                for line in block.get('lines',[]):
                    for span in line['spans']:
                        if span['size']>=12 or span['flags'] & 16:
                            info['headings'].append({'text':span['text'],'bbox':span['bbox'],'font':span['font'],'size':span['size']})
            for table_index,t in enumerate(p.find_tables()):
                rows=t.extract()
                visible_rows=[list(row) for row in rows]
                repairs=[]
                # A second text engine resolves broken font mappings within the SAME cell.
                # Never infer a missing sign or substitute a guessed exponent.
                for ri, row in enumerate(rows):
                    for ci, raw in enumerate(row):
                        cell=t.rows[ri].cells[ci]
                        if cell and not fp.rotation and abs(fp.rect.width-p.width)<1 and abs(fp.rect.height-p.height)<1:
                            direct=clean(fp.get_textbox(fitz.Rect(cell)))
                            if direct and '\x00' not in direct and '\ufffd' not in direct:visible_rows[ri][ci]=direct
                        if cell and raw and '\x00' in raw:
                            if fp.rotation or abs(fp.rect.width-p.width)>1 or abs(fp.rect.height-p.height)>1:
                                result['warnings'].append(f'Page {page_index+1}: incompatible coordinates; damaged cell requires review.')
                                continue
                            rect=fitz.Rect(cell)
                            recovered=clean(fp.get_textbox(rect))
                            if recovered and '\x00' not in recovered and '\ufffd' not in recovered:
                                repairs.append({'row':ri,'column':ci,'original':raw,'recovered':recovered})
                                row[ci]=recovered
                info['tables'].append({'index':table_index+1,'bbox':t.bbox,'rows':rows,'visible_rows':visible_rows,'repairs':repairs})
                best=None
                for ri,row in enumerate(rows):
                    cols={ci:module(v) for ci,v in enumerate(row) if module(v)}
                    if cols and (best is None or len(cols)>len(best[1])):best=(ri,cols)
                if best is None:continue
                hi,cols=best
                unit_col=None
                for ui in range(max(0,hi-1),min(len(rows),hi+3)):
                    for ci,v in enumerate(rows[ui]):
                        if re.fullmatch(r'units?(?:\(s\))?',clean(v).lower()):unit_col=(ui,ci)
                if unit_col is None:continue
                urow,uci=unit_col
                ubox=t.rows[urow].cells[uci]
                if not ubox:continue
                ux=(ubox[0]+ubox[2])/2
                # Anchor columns by their physical locations, not logical indices:
                # merged rows can have different cell indexes in the same table.
                anchors={}
                for ci,m in cols.items():
                    box=t.rows[hi].cells[ci]
                    if box:anchors[m]=(box[0]+box[2])/2
                for hri in range(hi+1,min(len(rows),hi+3)):
                    if any(numeric(v)[1]=='numeric' for v in visible_rows[hri] if clean(v)):break
                    for ci,v in enumerate(rows[hri]):
                        m=module(v);box=t.rows[hri].cells[ci]
                        if m and box and m not in anchors:anchors[m]=(box[0]+box[2])/2
                if len(set(cols.values()))!=len(cols):
                    result['warnings'].append(f'Page {page_index+1}, table {table_index+1}: duplicate or missing module header geometry; review required.')
                    continue
                table_text=' '.join(clean(v) for row in visible_rows for v in row)
                basis='legacy EN15804+A1' if re.search(r'kg\s+SO',table_text,re.I) and not re.search(r'GWP[- –]total',table_text,re.I) else 'unclassified/current'
                def at(ri,x):
                    candidates=[(ci,box) for ci,box in enumerate(t.rows[ri].cells) if box and box[0]<=x<=box[2]]
                    if not candidates:return '',None
                    ci,box=min(candidates,key=lambda item:item[1][2]-item[1][0])
                    return clean(visible_rows[ri][ci]),box
                for ri in range(hi+1,len(rows)):
                    unit,unit_box=at(ri,ux)
                    if not unit or unit.lower() in {'unit','units'}:continue
                    labels=[]
                    for ci,box in enumerate(t.rows[ri].cells):
                        if box and box[2]<=ubox[0]+1:
                            value=clean(visible_rows[ri][ci])
                            if value and value not in labels:labels.append(value)
                    label=' | '.join(labels)
                    if not label:continue
                    recovered=[(m,*at(ri,x)) for m,x in anchors.items()]
                    if not any(numeric(raw)[1] in {'numeric','not declared/relevant'} for m,raw,box in recovered):continue
                    for m,raw,box in recovered:
                        val,status=numeric(raw)
                        result['values'].append({'file':path.name,'page':page_index+1,'table':table_index+1,
                            'indicator':label,'unit':unit,'module':m,'raw':raw,'value':val,'status':status,'bbox':box,'basis':basis})
            if not any(v['page']==page_index+1 for v in result['values']):
                fallback=layout_fallback(fp,page_index+1,path.name)
                if fallback:
                    result['values'].extend(fallback)
                    result['warnings'].append(f'Page {page_index+1}: used aligned text geometry because ruled tables were not recovered. Review cell evidence.')
            result['pages'].append(info)
    if sum(len(clean(block[4])) for info in result['pages'] for block in info['blocks'])<50:
        result['warnings'].append('Image-only or text-empty PDF: OCR is required before reliable extraction. No values were inferred from images.')
    if not result['values']:
        result['warnings'].append('No lifecycle impact tables recovered. Borderless, scanned or unfamiliar layouts need further work; this is not a successful extraction.')
    return result

def export(result,destination):
    wb=Workbook(); ws=wb.active; ws.title='Impact values'
    columns=['file','page','table','indicator','unit','module','raw','value','status','bbox']
    ws.append(columns)
    for v in result['values']:
        ws.append([json.dumps(v[c]) if c=='bbox' else v[c] for c in columns])
    ws.freeze_panes='A2'; ws.auto_filter.ref=ws.dimensions
    for col,width in [('A',45),('D',60),('E',25),('F',15),('G',20),('H',20),('I',25),('J',45)]: ws.column_dimensions[col].width=width
    details=wb.create_sheet('Document details')
    dc=['field','value','raw','page','label','method','status','label_bbox','value_bbox']
    details.append(dc)
    for row in result.get('details',[]):
        details.append([json.dumps(row[c],ensure_ascii=False) if c.endswith('bbox') else row[c] for c in dc])
    details.freeze_panes='A2';details.auto_filter.ref=details.dimensions
    for c in ['A','B','C','E','F','G']:details.column_dimensions[c].width=45
    audit=wb.create_sheet('Warnings');audit.append(['Warning'])
    for f in result.get('missing_details',[]):audit.append([f+' — no explicit label/value recovered; not inferred.'])
    audit.append(['Document details scan the first eight pages. Extracted candidates require source review; multiple product names are retained separately.'])
    for x in result['warnings']: audit.append([x])
    audit.column_dimensions['A'].width=110
    wb.save(destination)

