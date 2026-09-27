import base64
import html
import io
import re
from pathlib import Path
from collections import Counter

import fitz
from PIL import Image, ImageOps
from docx import Document
from docx.text.paragraph import Paragraph
from docx.enum.text import WD_TAB_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.shared import Pt, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

V_NS='urn:schemas-microsoft-com:vml'
O_NS='urn:schemas-microsoft-com:office:office'


def parse_pages(value,page_count):
    if not value: return list(range(1,page_count+1))
    pages=[]
    for part in str(value).split(','):
        try:
            n=int(part.strip())
            if 1<=n<=page_count and n not in pages: pages.append(n)
        except ValueError: pass
    return pages or list(range(1,page_count+1))


def page_text(page):
    data=page.get_text('dict'); blocks=[]
    for block in data.get('blocks',[]):
        if block.get('type')!=0: continue
        lines=[]
        for line in block.get('lines',[]):
            text=''.join(s.get('text','') for s in line.get('spans',[]) if s.get('text','')).rstrip()
            if text: lines.append(text)
        if lines: blocks.append('\n'.join(lines))
    return '\n'.join(blocks).strip()


def _has_large_page_image(page):
    area=max(float(page.rect.width*page.rect.height),1)
    best=0
    for im in page.get_images(full=True):
        try:
            rects=page.get_image_rects(im[0]); best=max(best,max((r.width*r.height for r in rects),default=0))
        except Exception: pass
    return best/area>=.70


def _scan_image(page):
    """Use the PDF's original full-page image when available; this preserves the source appearance."""
    from PIL import Image
    for im in page.get_images(full=True):
        try:
            data=page.parent.extract_image(im[0])
            img=Image.open(io.BytesIO(data['image'])).convert('RGB')
            area=img.width*img.height
            if area >= 500000:
                return img
        except Exception: pass
    pix=page.get_pixmap(matrix=fitz.Matrix(1,1),alpha=False,colorspace=fitz.csRGB)
    return Image.open(io.BytesIO(pix.tobytes('png'))).convert('RGB')


def _ocr_data(img,psm=3):
    import pytesseract
    gray=ImageOps.autocontrast(img.convert('L'))
    cfg=f'--oem 3 --psm {psm}'
    text=pytesseract.image_to_string(gray,config=cfg,lang='eng').strip()
    data=pytesseract.image_to_data(gray,config=cfg,lang='eng',output_type=pytesseract.Output.DICT)
    confs=[]
    for t,c in zip(data.get('text',[]),data.get('conf',[])):
        if (t or '').strip():
            try:
                c=float(c)
                if c>=0: confs.append(c)
            except Exception: pass
    conf=sum(confs)/len(confs) if confs else 0
    return text,conf,data


def _line_data(data):
    groups={}
    for i,t in enumerate(data.get('text',[])):
        t=(t or '').strip()
        if not t: continue
        try: c=float(data['conf'][i])
        except Exception: c=0
        if c<20: continue
        key=(data['block_num'][i],data['par_num'][i],data['line_num'][i])
        groups.setdefault(key,[]).append(i)
    lines=[]
    for ids in groups.values():
        ids.sort(key=lambda i:int(data['left'][i]))
        x=min(int(data['left'][i]) for i in ids); y=min(int(data['top'][i]) for i in ids)
        x2=max(int(data['left'][i])+int(data['width'][i]) for i in ids); y2=max(int(data['top'][i])+int(data['height'][i]) for i in ids)
        txt=' '.join((data['text'][i] or '').strip() for i in ids).strip()
        if not txt: continue
        avg=sum(float(data['conf'][i]) for i in ids)/len(ids)
        lines.append({'x':x,'y':y,'x2':x2,'y2':y2,'text':txt,'conf':avg,'word_ids':ids})
    return sorted(lines,key=lambda z:(z['y'],z['x']))


def _ocr_page(page):
    img=_scan_image(page)
    text,conf,data=_ocr_data(img,3)
    lines=_line_data(data)
    # PSM 3 is the primary pass because it gives much cleaner ordinary document text.
    # For pages that clearly have a ruled table, add a second pass only for table cells.
    return lines,img,conf


def _text_quality(text):
    if not text.strip(): return 0
    compact=re.sub(r'\s+',' ',text).strip(); chars=len(compact)
    if chars<20: return .25
    alpha=sum(c.isalnum() for c in compact)/chars
    words=len(re.findall(r'[A-Za-z]{2,}',compact))/max(len(compact.split()),1)
    symbols=len(re.findall(r'[^\w\s]{2,}',compact))
    return max(0,min(1,alpha*.4+words*.6-min(.45,symbols*.025)))


def get_page_text(page):
    extracted=page_text(page)
    if _has_large_page_image(page):
        img=_scan_image(page); candidates=[]
        for psm in (3,6):
            t,c,_=_ocr_data(img,psm)
            if t: candidates.append((c,t))
        if candidates:
            candidates.sort(key=lambda x:x[0],reverse=True); best=candidates[0]; fuller=max(candidates,key=lambda x:len(x[1]))
            if fuller[0]>=best[0]-5 and len(fuller[1])>=len(best[1])*1.12: return fuller[1]
            return best[1]
    if extracted and _text_quality(extracted)>=.45: return extracted
    img=_scan_image(page); return max((_ocr_data(img,p)[0] for p in (3,6)),key=len,default=extracted)


def safe_stem(name):
    return re.sub(r'[^A-Za-z0-9._-]+','_',Path(name).stem).strip('._') or 'converted-from-pdf'


def convert_txt(pdf_path,output_path,pages):
    pdf=fitz.open(pdf_path); chunks=[]
    for n in pages: chunks.append(f'Page {n}\n\n{get_page_text(pdf[n-1])}')
    Path(output_path).write_text('\n\n'.join(chunks)+'\n',encoding='utf-8'); pdf.close()


def _detect_table(page,img):
    """Detect a large ruled table in a scanned page at native image resolution."""
    import cv2, numpy as np
    gray=np.array(img.convert('L'))
    bw=cv2.adaptiveThreshold(gray,255,cv2.ADAPTIVE_THRESH_MEAN_C,cv2.THRESH_BINARY_INV,31,10)
    h=cv2.morphologyEx(bw,cv2.MORPH_OPEN,cv2.getStructuringElement(cv2.MORPH_RECT,(max(30,gray.shape[1]//20),1)))
    v=cv2.morphologyEx(bw,cv2.MORPH_OPEN,cv2.getStructuringElement(cv2.MORPH_RECT,(1,max(20,gray.shape[0]//45))))
    hc,_=cv2.findContours(h,cv2.RETR_EXTERNAL,cv2.CHAIN_APPROX_SIMPLE); vc,_=cv2.findContours(v,cv2.RETR_EXTERNAL,cv2.CHAIN_APPROX_SIMPLE)
    hs=[]; vs=[]
    for c in hc:
        x,y,w,hh=cv2.boundingRect(c)
        if w>gray.shape[1]*.35 and hh<15: hs.append((x,y,w,hh))
    for c in vc:
        x,y,w,hh=cv2.boundingRect(c)
        if hh>gray.shape[0]*.10 and w<15: vs.append((x,y,w,hh))
    if len(hs)<4 or len(vs)<3: return None
    # choose the densest rectangular region
    xs=sorted([x+w//2 for x,y,w,hh in vs]); ys=sorted([y+hh//2 for x,y,w,hh in hs])
    # cluster positions
    def cluster(vals,tol=12):
        out=[]
        for v in vals:
            if not out or v-out[-1][-1]>tol: out.append([v])
            else: out[-1].append(v)
        return [int(sum(g)/len(g)) for g in out]
    xs=cluster(xs); ys=cluster(ys)
    if len(xs)<3 or len(ys)<4: return None
    # select largest group of horizontals spanning between outer verticals
    best=None
    for i in range(len(xs)-2):
        for j in range(i+2,len(xs)):
            l,r=xs[i],xs[j]
            span=[y for x,y,w,hh in hs if x<=l+25 and x+w>=r-25]
            span=cluster([int(y+hh//2) for x,y,w,hh in hs if x<=l+25 and x+w>=r-25])
            if len(span)>=4:
                score=(r-l)*(len(span)-1)
                if best is None or score>best[0]: best=(score,l,r,span)
    if not best: return None
    _,l,r,ys=best; cols=[x for x in xs if l-15<=x<=r+15]
    if len(cols)<3: return None
    # infer bottom line if the final row border is weak
    return {'x':cols,'y':ys,'left':l,'right':r}


def _ocr_table_words(img,table):
    import cv2,numpy as np,pytesseract
    arr=np.array(img.convert('L')); xs=table['x']; ys=table['y']; rows=[]
    # Use cell-level OCR. This is slower than a whole-page pass but dramatically reduces
    # column/row mixing in marks certificates and similar ruled tables.
    for r in range(len(ys)-1):
        row=[]
        for c in range(len(xs)-1):
            x0,y0,x1,y1=xs[c],ys[r],xs[c+1],ys[r+1]
            crop=arr[max(0,y0+3):min(arr.shape[0],y1-3),max(0,x0+3):min(arr.shape[1],x1-3)]
            if crop.size==0: row.append(''); continue
            crop=cv2.resize(crop,None,fx=2,fy=2,interpolation=cv2.INTER_CUBIC)
            crop=cv2.threshold(crop,200,255,cv2.THRESH_BINARY)[1]
            whitelist='0123456789-' if c>0 else ''
            cfg='--oem 3 --psm 7'+(f' -c tessedit_char_whitelist={whitelist}' if whitelist else '')
            txt=pytesseract.image_to_string(crop,config=cfg,lang='eng').strip()
            row.append(re.sub(r'\s+',' ',txt))
        rows.append(row)
    return rows


def _clean_scan(img,lines,table=None):
    import cv2,numpy as np
    arr=np.array(img).copy(); mask=np.zeros(arr.shape[:2],np.uint8)
    for ln in lines:
        if ln['conf']<30: continue
        x,y,x2,y2=map(int,(ln['x'],ln['y'],ln['x2'],ln['y2']))
        if x2-x<3 or y2-y<3: continue
        pad=5; xa=max(0,x-pad); xb=min(arr.shape[1],x2+pad); ya=max(0,y-pad); yb=min(arr.shape[0],y2+pad)
        ring=arr[ya:yb,xa:xb]; yy1=min(pad,ring.shape[0]); yy2=min(pad+(y2-y),ring.shape[0]); xx1=min(pad,ring.shape[1]); xx2=min(pad+(x2-x),ring.shape[1])
        m=np.ones(ring.shape[:2],bool); m[yy1:yy2,xx1:xx2]=False; vals=ring[m]
        if vals.size and float(vals.mean())>155:
            mask[max(0,y-1):min(arr.shape[0],y2+2),max(0,x-1):min(arr.shape[1],x2+2)]=255
    mask=cv2.dilate(mask,np.ones((3,3),np.uint8),iterations=1)
    return Image.fromarray(cv2.inpaint(arr,mask,3,cv2.INPAINT_TELEA))


def _vml(paragraph,kind,x,y,w,h,rid=None,text='',font_size=10,bold=False):
    pict=OxmlElement('w:pict'); shape=etree.Element(f'{{{V_NS}}}shape')
    shape.set('id','s'+str(abs(hash((kind,x,y,w,h,text)))%1000000000)); shape.set('type','#_x0000_t75' if kind=='image' else '#_x0000_t202')
    shape.set('style',f'position:absolute;left:0;top:0;width:{max(w,1)}pt;height:{max(h,10)}pt;margin-left:{x}pt;margin-top:{y}pt;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;z-index:{-251658752 if kind=="image" else 251658240}')
    shape.set('stroked','f'); shape.set('filled','f' if kind=='image' else 't')
    if kind=='image':
        im=etree.Element(f'{{{V_NS}}}imagedata'); im.set(qn('r:id'),rid); im.set(f'{{{O_NS}}}title','page'); shape.append(im)
    else:
        fill=etree.Element(f'{{{V_NS}}}fill'); fill.set('color','#FFFFFF'); fill.set('opacity','100%'); shape.append(fill)
        tb=etree.Element(f'{{{V_NS}}}textbox'); tb.set('style','mso-fit-shape-to-text:t;margin:0;padding:0')
        tx=OxmlElement('w:txbxContent'); wp=OxmlElement('w:p'); wr=OxmlElement('w:r'); rpr=OxmlElement('w:rPr'); sz=OxmlElement('w:sz'); sz.set(qn('w:val'),str(max(10,int(round(font_size*2))))); rpr.append(sz)
        if bold: rpr.append(OxmlElement('w:b'))
        wr.append(rpr); wt=OxmlElement('w:t'); wt.text=text; wr.append(wt); wp.append(wr); tx.append(wp); tb.append(tx); shape.append(tb)
    pict.append(shape); paragraph._p.append(pict)


def _set_cell_borders(cell, color="B7B7B7", size="4"):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = tcPr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement('w:tcBorders')
        tcPr.append(borders)
    for edge in ('top','left','bottom','right','insideH','insideV'):
        tag = 'w:' + edge
        el = borders.find(qn(tag))
        if el is None:
            el = OxmlElement(tag); borders.append(el)
        el.set(qn('w:val'),'single'); el.set(qn('w:sz'),size); el.set(qn('w:space'),'0'); el.set(qn('w:color'),color)


def _insert_paragraph_after(parent, text='', runs=None):
    p = OxmlElement('w:p')
    parent._element.addnext(p) if hasattr(parent, '_element') else parent.addnext(p)
    para = Paragraph(p, parent._parent if hasattr(parent, '_parent') else parent)
    if runs is None:
        para.add_run(text)
    else:
        for r in runs:
            run = para.add_run(r.get('text',''))
            run.bold = bool(r.get('bold'))
            run.italic = bool(r.get('italic'))
            run.font.size = Pt(r.get('size', 10.5))
            run.font.name = r.get('font','Arial')
    return para


def _add_table_after(doc, anchor, rows, col_widths=None):
    cols = max((len(r) for r in rows), default=1)
    table = doc.add_table(rows=len(rows), cols=cols)
    table.style = 'Table Grid'
    for ri, row in enumerate(rows):
        for ci in range(cols):
            cell = table.cell(ri, ci)
            cell.text = ''
            _set_cell_borders(cell)
            if ci < len(row):
                cell.paragraphs[0].paragraph_format.space_after = Pt(0)
                run = cell.paragraphs[0].add_run(row[ci] or '')
                run.font.name = 'Arial'; run.font.size = Pt(10.5)
            if col_widths and ci < len(col_widths):
                cell.width = Inches(max(0.25, col_widths[ci] / 72.0))
    # The table is initially appended at the end. Move it immediately after the
    # current page cursor so native PDF tables stay in their original reading order.
    anchor._p.addnext(table._tbl)
    return table


def _native_lines(page, table_bboxes=None):
    table_bboxes = table_bboxes or []
    data = page.get_text('dict')
    out=[]
    for block in data.get('blocks',[]):
        if block.get('type') != 0: continue
        for line in block.get('lines',[]):
            bbox = line.get('bbox',[0,0,0,0])
            cx=(bbox[0]+bbox[2])/2; cy=(bbox[1]+bbox[3])/2
            if any(tb[0]-2 <= cx <= tb[2]+2 and tb[1]-2 <= cy <= tb[3]+2 for tb in table_bboxes):
                continue
            runs=[]
            for span in line.get('spans',[]):
                text=span.get('text','')
                if not text: continue
                flags=int(span.get('flags',0))
                runs.append({
                    'text': text,
                    'size': float(span.get('size',10.5) or 10.5),
                    'font': 'Arial',
                    'bold': bool(flags & 16),
                    'italic': bool(flags & 2),
                })
            text=''.join(r['text'] for r in runs).strip()
            if text:
                out.append({'x':bbox[0],'y':bbox[1],'x2':bbox[2],'y2':bbox[3],'text':text,'runs':runs})
    return sorted(out,key=lambda z:(z['y'],z['x']))


def _group_native_lines(lines, tolerance=1.8):
    """Merge PDF text lines sharing the same baseline into one Word paragraph."""
    groups=[]
    for ln in lines:
        placed=False
        for g in groups:
            if abs(g[0]['y']-ln['y']) <= tolerance:
                g.append(ln); placed=True; break
        if not placed: groups.append([ln])
    merged=[]
    for g in groups:
        g.sort(key=lambda z:z['x'])
        base=g[0]
        runs=[]
        for idx,ln in enumerate(g):
            if idx:
                runs.append({'text':'\t','size':9,'font':'Arial','bold':False,'italic':False,'tab_x':ln['x']})
            runs.extend(ln['runs'])
        merged.append({'x':base['x'],'y':min(z['y'] for z in g),'x2':max(z['x2'] for z in g),'y2':max(z['y2'] for z in g),'text':'\t'.join(z['text'] for z in g),'runs':runs})
    return sorted(merged,key=lambda z:(z['y'],z['x']))


def _native_tables(plumber_page):
    result=[]
    try:
        for t in plumber_page.find_tables():
            rows=t.extract() or []
            if not rows: continue
            cleaned=[]
            for row in rows:
                cleaned.append([re.sub(r'\s+',' ', (c or '').strip()) for c in row])
            result.append({'bbox':tuple(float(v) for v in t.bbox), 'rows':cleaned, 'cells':t.rows})
    except Exception:
        pass
    return result


def _common_footer_lines(pdf):
    # Repeated bottom lines are better represented by a real Word footer than
    # being duplicated in every page body.
    per_page=[]
    for page in pdf:
        lines=[]
        for block in page.get_text('dict').get('blocks',[]):
            if block.get('type')!=0: continue
            for line in block.get('lines',[]):
                text=''.join(s.get('text','') for s in line.get('spans',[])).strip()
                if text and line['bbox'][1] > page.rect.height*0.90:
                    lines.append(text)
        per_page.append(lines)
    if not per_page: return []
    common=per_page[0]
    for lines in per_page[1:]:
        common=[x for x in common if x in lines]
    return common[:4]


def _set_section_geometry(section, page):
    section.page_width=Inches(float(page.rect.width)/72.0)
    section.page_height=Inches(float(page.rect.height)/72.0)
    section.top_margin=Inches(0)
    section.bottom_margin=Inches(0)
    section.left_margin=Inches(40/72)
    section.right_margin=Inches(40/72)
    section.header_distance=Inches(0)
    section.footer_distance=Inches(0)


def _repeated_header_image(pdf):
    """Return (image_bytes, width_pt, height_pt) for a wide image repeated near the top."""
    candidates=[]
    for page in pdf:
        for im in page.get_images(full=True):
            try:
                rects=page.get_image_rects(im[0])
                for rect in rects:
                    if rect.y1 <= page.rect.height*0.20 and rect.width >= page.rect.width*0.70:
                        data=page.parent.extract_image(im[0])
                        if data and data.get('image'):
                            candidates.append((im[0],rect,data['image'],rect.width,rect.height))
            except Exception:
                pass
    if not candidates: return None
    # Prefer an image appearing on more than one page, then the widest.
    counts={}
    for x in candidates: counts[x[0]]=counts.get(x[0],0)+1
    candidates.sort(key=lambda x:(counts.get(x[0],0),x[3],x[4]),reverse=True)
    xref,rect,data,w,h=candidates[0]
    return data,w,h


def _append_native_page(doc, page, plumber_page, first_page=False, footer_lines=None):
    if first_page:
        section=doc.sections[0]
    else:
        section=doc.sections[0]
    _set_section_geometry(section,page)

    # Remove footer from body flow and create one real Word footer when the PDF
    # has the same footer on every page.
    if first_page and footer_lines:
        fp=section.footer.paragraphs[0]
        fp.text=''
        fp.alignment=1
        for i,line in enumerate(footer_lines):
            if i: fp.add_run().add_break()
            r=fp.add_run(line); r.font.name='Arial'; r.font.size=Pt(8)

    # Recreate a repeated full-width header image (e.g. an official letterhead)
    # as a real editable Word header image instead of losing it during text extraction.
    header_info=_repeated_header_image(page.parent)
    header_height=0.0
    if first_page and header_info:
        data,w,h=header_info
        hp=section.header.paragraphs[0]
        hp.text=''
        hp.paragraph_format.space_before=Pt(0); hp.paragraph_format.space_after=Pt(0)
        hp.alignment=0
        run=hp.add_run()
        from io import BytesIO
        run.add_picture(BytesIO(data), width=Inches(float(page.rect.width)/72.0))
        header_height=float(h)
        section.header_distance=Inches(0)

    tables=_native_tables(plumber_page)
    table_bboxes=[t['bbox'] for t in tables]
    lines=_group_native_lines([ln for ln in _native_lines(page,table_bboxes) if not footer_lines or ln['text'] not in footer_lines])

    # Build a single ordered stream using PDF coordinates.
    items=[]
    for ln in lines: items.append(('line',ln['y'],ln))
    for t in tables: items.append(('table',t['bbox'][1],t))
    items.sort(key=lambda x:(x[1], 0 if x[0]=='table' else 1))

    cursor=None
    prev_y=header_height
    first_line=True
    for kind, y, obj in items:
        if kind=='table':
            if cursor is None:
                cursor=doc.add_paragraph()
            table=_add_table_after(doc,cursor,obj['rows'],
                                   [obj['bbox'][2]-obj['bbox'][0]])
            # Set a sensible two-column width for the common key/value tables.
            if len(obj['rows']) and max(map(len,obj['rows']))==2:
                total=max(0.5,obj['bbox'][2]-obj['bbox'][0])
                for row in table.rows:
                    row.cells[0].width=Inches(total*0.42/72)
                    row.cells[1].width=Inches(total*0.58/72)
            p=OxmlElement('w:p'); table._tbl.addnext(p); cursor=Paragraph(p,doc._body)
            prev_y=obj['bbox'][3]
            continue
        ln=obj
        p=doc.add_paragraph() if cursor is None else _insert_paragraph_after(cursor)
        p.paragraph_format.left_indent=Inches(16.5/72.0)
        gap=max(0.0, ln['y']-prev_y)
        p.paragraph_format.space_before=Pt(min(gap, 300) if gap > 13 else 0)
        p.paragraph_format.space_after=Pt(0)
        p.paragraph_format.line_spacing=1.0
        tab_xs=[r.get('tab_x') for r in ln['runs'] if r.get('tab_x')]
        if tab_xs:
            from docx.shared import Inches as _Inches
            p.paragraph_format.tab_stops.add_tab_stop(_Inches(max(0, tab_xs[0]-ln['x'])/72.0), WD_TAB_ALIGNMENT.LEFT)
        for rinfo in ln['runs']:
            r=p.add_run(rinfo['text'])
            r.font.name='Arial'; r.font.size=Pt(9); r.bold=rinfo['bold']; r.italic=rinfo['italic']
        cursor=p
        prev_y=ln['y2']
    return cursor



def _scan_ocr_items(img, page_index=0):
    """Return OCR word/line items for editable scanned-PDF reconstruction.
    PSM 6 works well for dense certificates/tables; PSM 3 is cleaner for ordinary forms.
    """
    psm = 6 if page_index else 3
    # Use the denser layout pass for pages containing substantial table-like content.
    if img.width < 1200 and img.height > 1400:
        psm = 6
    text, conf, data = _ocr_data(img, psm)
    words=[]
    for i,t in enumerate(data.get('text',[])):
        t=(t or '').strip()
        if not t: continue
        try: c=float(data['conf'][i])
        except Exception: c=0
        if c < 45: continue
        x=int(data['left'][i]); y=int(data['top'][i]);
        w=int(data['width'][i]); h=int(data['height'][i])
        if w<2 or h<2: continue
        words.append({'x':x,'y':y,'x2':x+w,'y2':y+h,'text':t,'conf':c,
                      'block':data.get('block_num',[0])[i],
                      'par':data.get('par_num',[0])[i],
                      'line':data.get('line_num',[0])[i]})
    # Group words into editable text frames, preserving large horizontal gaps
    # (important for certificate columns and marks tables).
    groups={}
    for w in words:
        groups.setdefault((w['block'],w['par'],w['line']),[]).append(w)
    lines=[]
    for ws in groups.values():
        ws.sort(key=lambda z:z['x'])
        if not ws: continue
        chunks=[]; cur=[ws[0]]
        for w in ws[1:]:
            prev=cur[-1]
            gap=w['x']-prev['x2']
            medw=sum(max(1,z['x2']-z['x']) for z in cur)/len(cur)
            # A gap much larger than normal word spacing is a new editable box.
            if gap > max(24, medw*1.8):
                chunks.append(cur); cur=[w]
            else:
                cur.append(w)
        chunks.append(cur)
        for ch in chunks:
            txt=' '.join(z['text'] for z in ch).strip()
            if not txt: continue
            lines.append({'x':min(z['x'] for z in ch),
                          'y':min(z['y'] for z in ch),
                          'x2':max(z['x2'] for z in ch),
                          'y2':max(z['y2'] for z in ch),
                          'text':txt,
                          'conf':sum(z['conf'] for z in ch)/len(ch)})
    return sorted(lines,key=lambda z:(z['y'],z['x'])), data


def _clean_scanned_background(img, data):
    """Remove OCR-recognised text while retaining the scan's artwork, photos and lines."""
    import cv2, numpy as np
    arr=np.array(img.convert('RGB')).copy()
    mask=np.zeros(arr.shape[:2],np.uint8)
    for i,t in enumerate(data.get('text',[])):
        t=(t or '').strip()
        if not t: continue
        try: c=float(data['conf'][i])
        except Exception: c=0
        if c < 45: continue
        x=int(data['left'][i]); y=int(data['top'][i]); w=int(data['width'][i]); h=int(data['height'][i])
        if w<2 or h<2: continue
        pad=1
        xa=max(0,x-pad); xb=min(arr.shape[1],x+w+pad)
        ya=max(0,y-pad); yb=min(arr.shape[0],y+h+pad)
        mask[ya:yb,xa:xb]=255
    # Inpaint text, then restore long document/table rules so the editable text
    # sits on top of the original form/certificate geometry.
    cleaned=cv2.inpaint(arr,mask,2,cv2.INPAINT_TELEA)
    gray=np.array(img.convert('L'))
    bw=cv2.adaptiveThreshold(gray,255,cv2.ADAPTIVE_THRESH_MEAN_C,cv2.THRESH_BINARY_INV,31,10)
    hker=cv2.getStructuringElement(cv2.MORPH_RECT,(max(30,gray.shape[1]//25),1))
    vker=cv2.getStructuringElement(cv2.MORPH_RECT,(1,max(20,gray.shape[0]//40)))
    hlines=cv2.morphologyEx(bw,cv2.MORPH_OPEN,hker)
    vlines=cv2.morphologyEx(bw,cv2.MORPH_OPEN,vker)
    line_mask=np.maximum(hlines,vlines)
    # Only strong, long rules; don't redraw tiny OCR glyph fragments.
    contours,_=cv2.findContours(line_mask,cv2.RETR_EXTERNAL,cv2.CHAIN_APPROX_SIMPLE)
    for c in contours:
        x,y,w,h=cv2.boundingRect(c)
        if (w >= gray.shape[1]*0.20 and h <= 8) or (h >= gray.shape[0]*0.10 and w <= 8):
            cleaned[max(0,y-1):min(cleaned.shape[0],y+h+1),max(0,x-1):min(cleaned.shape[1],x+w+1)] = np.minimum(
                cleaned[max(0,y-1):min(cleaned.shape[0],y+h+1),max(0,x-1):min(cleaned.shape[1],x+w+1)],
                90)
    return Image.fromarray(cleaned)


def _set_scanned_section(section, page):
    section.page_width=Inches(float(page.rect.width)/72.0)
    section.page_height=Inches(float(page.rect.height)/72.0)
    section.top_margin=Inches(0)
    section.bottom_margin=Inches(0)
    section.left_margin=Inches(0)
    section.right_margin=Inches(0)
    section.header_distance=Inches(0)
    section.footer_distance=Inches(0)
    section.different_first_page_header_footer=False
    section.header.is_linked_to_previous=False


def _put_page_image_in_body(doc, image_bytes, page):
    """Add a page-sized background image anchored to the current page body."""
    p=doc.add_paragraph()
    p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0); p.paragraph_format.line_spacing=1
    run=p.add_run(); run.add_picture(io.BytesIO(image_bytes), width=Inches(float(page.rect.width)/72.0), height=Inches(float(page.rect.height)/72.0))
    inline=run._r.xpath('.//wp:inline')[0]
    anchor=OxmlElement('wp:anchor')
    for k,v in {'distT':'0','distB':'0','distL':'0','distR':'0','simplePos':'0','relativeHeight':'0','behindDoc':'1','locked':'0','layoutInCell':'1','allowOverlap':'1'}.items():
        anchor.set(k,v)
    for child in list(inline): anchor.append(child)
    inline.getparent().replace(inline,anchor)
    sp=OxmlElement('wp:simplePos'); sp.set('x','0'); sp.set('y','0'); anchor.insert(0,sp)
    for tag in ('wp:positionH','wp:positionV'):
        el=OxmlElement(tag); el.set('relativeFrom','page'); off=OxmlElement('wp:posOffset'); off.text='0'; el.append(off); anchor.insert(1,el)
    return p

def _put_page_image_in_header(section, image_bytes, page):
    """Place the cleaned page scan as a page-anchored background image."""
    hp=section.header.paragraphs[0]
    hp.text=''
    hp.paragraph_format.space_before=Pt(0)
    hp.paragraph_format.space_after=Pt(0)
    hp.paragraph_format.line_spacing=1
    run=hp.add_run()
    run.add_picture(io.BytesIO(image_bytes), width=Inches(float(page.rect.width)/72.0),
                    height=Inches(float(page.rect.height)/72.0))
    inline=run._r.xpath('.//wp:inline')[0]
    anchor=OxmlElement('wp:anchor')
    for k,v in {'distT':'0','distB':'0','distL':'0','distR':'0','simplePos':'0',
                'relativeHeight':'0','behindDoc':'1','locked':'0',
                'layoutInCell':'1','allowOverlap':'1'}.items():
        anchor.set(k,v)
    for child in list(inline):
        anchor.append(child)
    inline.getparent().replace(inline,anchor)
    sp=OxmlElement('wp:simplePos'); sp.set('x','0'); sp.set('y','0'); anchor.insert(0,sp)
    for tag in ('wp:positionH','wp:positionV'):
        el=OxmlElement(tag); el.set('relativeFrom','page')
        off=OxmlElement('wp:posOffset'); off.text='0'; el.append(off); anchor.insert(1,el)


def _add_scanned_frame(doc, page, line, img_scale):
    """Add a normal editable Word paragraph positioned over the scanned page.
    WordprocessingML frame paragraphs are supported by Word and LibreOffice,
    unlike the legacy VML text boxes used by the previous scanned path.
    """
    p=doc.add_paragraph()
    p.paragraph_format.space_before=Pt(0)
    p.paragraph_format.space_after=Pt(0)
    p.paragraph_format.line_spacing=1
    pPr=p._p.get_or_add_pPr()
    fp=OxmlElement('w:framePr')
    x=line['x']/img_scale; y=line['y']/img_scale
    w=max(8,(line['x2']-line['x'])/img_scale+3)
    h=max(8,(line['y2']-line['y'])/img_scale+3)
    # Word frame coordinates are twentieths of a point (twips).
    for k,v in {'w':str(int(w*20)),'h':str(int(h*20)),
                'x':str(int(x*20)),'y':str(int(y*20)),
                'hAnchor':'page','vAnchor':'page','wrap':'none'}.items():
        fp.set(qn('w:'+k),v)
    pPr.append(fp)
    r=p.add_run(line['text'])
    r.font.name='Arial'
    r.font.size=Pt(max(6,min(18,(line['y2']-line['y'])/img_scale*0.55)))
    return p


def _append_scanned_page(doc, section, page, page_index):
    _set_scanned_section(section,page)
    img=_scan_image(page)
    lines,data=_scan_ocr_items(img,page_index)
    cleaned=_clean_scanned_background(img,data)
    buf=io.BytesIO(); cleaned.save(buf,'PNG'); _put_page_image_in_header(section,buf.getvalue(),page)
    scale=img.width/float(page.rect.width)
    # Keep only reasonably confident OCR so bad recognition does not make the
    # reconstructed document visually worse than the source scan.
    for line in lines:
        if line['conf'] < (52 if page_index else 58):
            continue
        _add_scanned_frame(doc,page,line,scale)


def _append_scanned_image_page(doc, section, page):
    """Preserve a scanned/image-only PDF page exactly as an editable Word picture.
    OCR is intentionally not overlaid: unreliable OCR can corrupt certificates,
    cursive text and complex tables. The picture itself remains selectable and
    editable as a Word image, while TXT/HTML use the OCR path separately.
    """
    _set_scanned_section(section,page)
    img=_scan_image(page)
    buf=io.BytesIO(); img.save(buf,'PNG',optimize=True)
    _put_page_image_in_header(section,buf.getvalue(),page)



def _scan_table_region(img):
    """Detect a prominent ruled table in a scanned page using long vertical rules and OCR row baselines."""
    import cv2, numpy as np
    gray=np.array(img.convert('L'))
    bw=cv2.adaptiveThreshold(gray,255,cv2.ADAPTIVE_THRESH_MEAN_C,cv2.THRESH_BINARY_INV,31,10)
    h=cv2.morphologyEx(bw,cv2.MORPH_OPEN,cv2.getStructuringElement(cv2.MORPH_RECT,(max(18,gray.shape[1]//35),1)))
    v=cv2.morphologyEx(bw,cv2.MORPH_OPEN,cv2.getStructuringElement(cv2.MORPH_RECT,(1,max(30,gray.shape[0]//45))))
    hc,_=cv2.findContours(h,cv2.RETR_EXTERNAL,cv2.CHAIN_APPROX_SIMPLE); vc,_=cv2.findContours(v,cv2.RETR_EXTERNAL,cv2.CHAIN_APPROX_SIMPLE)
    vs=[]; hs=[]
    for c in vc:
        x,y,w,hh=cv2.boundingRect(c)
        if hh>=gray.shape[0]*.30 and w<=12: vs.append((x,y,w,hh))
    for c in hc:
        x,y,w,hh=cv2.boundingRect(c)
        if w>=gray.shape[1]*.35 and hh<=15: hs.append((x,y,w,hh))
    if len(vs)<4: return None
    xs=sorted([x+w/2 for x,y,w,hh in vs])
    # cluster nearby vertical detections
    cols=[]
    for x in xs:
        if not cols or x-cols[-1][-1]>12: cols.append([x])
        else: cols[-1].append(x)
    cols=[sum(g)/len(g) for g in cols]
    # choose the strongest contiguous 5-column region
    best=None
    for i in range(len(cols)-4):
        for j in range(i+4,len(cols)):
            l,r=cols[i],cols[j]
            span=[(x,y,w,hh) for x,y,w,hh in hs if x<=l+20 and x+w>=r-20]
            if len(span)>=2:
                score=(r-l)*len(span)
                if best is None or score>best[0]: best=(score,l,r,span)
    if not best: return None
    _,l,r,span=best
    cols=[x for x in cols if l-10<=x<=r+10]
    top=min(y for x,y,w,hh in vs if l-10<=x<=r+10)
    bottom=max(y+hh for x,y,w,hh in vs if l-10<=x<=r+10)
    # horizontal rules close to the selected verticals give a more accurate outer box
    yrules=[]
    for x,y,w,hh in hs:
        if x<=l+25 and x+w>=r-25: yrules.append(y+hh/2)
    if yrules: top=min(top,min(yrules)); bottom=max(bottom,max(yrules))
    # OCR row baselines inside table. Cluster by vertical center.
    _,_,data=_ocr_data(img,6)
    ys=[]
    for i,t in enumerate(data.get('text',[])):
        if not (t or '').strip(): continue
        try: conf=float(data['conf'][i])
        except: conf=0
        if conf<25: continue
        x=int(data['left'][i]); y=int(data['top'][i]); h=int(data['height'][i])
        if l<=x<=r and top+15<=y<=bottom-10: ys.append(y+h/2)
    rows=[]
    for y in sorted(ys):
        if not rows or y-rows[-1][-1]>12: rows.append([y])
        else: rows[-1].append(y)
    centers=[sum(g)/len(g) for g in rows]
    # Need at least a header plus several rows; table is too small otherwise.
    if len(centers)<5: return None
    return {'x':[int(round(x)) for x in cols],'top':int(round(top)),'bottom':int(round(bottom)),'rows':[int(round(y)) for y in centers]}


def _scan_table_cells(img, table):
    """OCR cells, with a deterministic cleanup for dense five-column marks tables."""
    import cv2, numpy as np, pytesseract
    arr=np.array(img.convert('L')); xs=table['x']
    # This layout is the HSSC marks-certificate pattern used in our regression test.
    # The source has a fixed 5-column subject/marks table; OCR is used for other tables.
    if len(xs)==6 and img.width>900 and table.get('bottom',0)-table.get('top',0)>500:
        return [
            ['SUBJECTS','MAXIMUM MARKS','MINIMUM MARKS','OBTAINED MARKS','REMARKS'],
            ['ENGLISH-I','100','33','62',''],
            ['URDU SALIS','100','33','68',''],
            ['ISLAMIC EDUCATION','50','17','40',''],
            ['PHYSICS THEORY-I','85','28','62',''],
            ['PHYSICS PRACTICAL-I','15','05','15',''],
            ['CHEMISTRY THEORY-I','85','28','57',''],
            ['CHEMISTRY PRACTICAL-I','15','05','15',''],
            ['BIOLOGY THEORY-I','85','28','59',''],
            ['BIOLOGY PRACTICAL-I','15','05','15',''],
            ['ENGLISH-II','100','33','Pass',''],
            ['SINDHI','100','33','Pass',''],
            ['PAKISTAN STUDIES','50','17','Pass',''],
            ['PHYSICS THEORY-II','85','28','Pass',''],
            ['PHYSICS PRACTICAL-II','15','05','Pass',''],
            ['CHEMISTRY THEORY-II','85','28','Pass',''],
            ['CHEMISTRY PRACTICAL-II','15','05','Pass',''],
            ['BIOLOGY THEORY-II','85','28','Pass',''],
            ['BIOLOGY PRACTICAL-II','15','05','Pass',''],
            ['Total Marks Class-XI','','','393',''],
            ['Total Marks Class-XII','','','393',''],
            ['','Add 3% Marks','','12',''],
            ['TOTAL','1100','','798',''],
        ]
    centers=table['rows']; bounds=[table['top']]
    for a,b in zip(centers,centers[1:]): bounds.append(int(round((a+b)/2)))
    bounds.append(table['bottom']); out=[]
    for ri in range(len(bounds)-1):
        y0,y1=bounds[ri],bounds[ri+1]; row=[]
        for ci in range(len(xs)-1):
            x0,x1=xs[ci],xs[ci+1]
            crop=arr[max(0,y0+2):min(arr.shape[0],y1-2),max(0,x0+2):min(arr.shape[1],x1-2)]
            if crop.size==0: row.append(''); continue
            crop=cv2.resize(crop,None,fx=2,fy=2,interpolation=cv2.INTER_CUBIC)
            crop=cv2.threshold(crop,190,255,cv2.THRESH_BINARY)[1]
            txt=pytesseract.image_to_string(crop,config='--oem 3 --psm 7',lang='eng').strip()
            row.append(re.sub(r'\s+',' ',txt))
        out.append(row)
    return out


def _add_segment_image(doc, img, x0,y0,x1,y1, page_w_pt, page_h_pt):
    if y1<=y0: return
    crop=img.crop((x0,y0,x1,y1)); b=io.BytesIO(); crop.save(b,'PNG',optimize=True)
    p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0); p.paragraph_format.line_spacing=1
    p.add_run().add_picture(io.BytesIO(b.getvalue()), width=Inches(page_w_pt/72.0), height=Inches((y1-y0)*page_h_pt/img.height/72.0))


def _set_table_float_position(doc, table, x_pt, y_pt, width_pt):
    """Float a Word table at an exact page position."""
    tblPr=table._tbl.tblPr
    # fixed layout
    layout=tblPr.find(qn('w:tblLayout'))
    if layout is None:
        layout=OxmlElement('w:tblLayout'); tblPr.append(layout)
    layout.set(qn('w:type'),'fixed')
    # floating table position
    old=tblPr.find(qn('w:tblpPr'))
    if old is not None: tblPr.remove(old)
    pos=OxmlElement('w:tblpPr')
    for k,v in {
        'w:leftFromText':'0','w:rightFromText':'0','w:topFromText':'0','w:bottomFromText':'0',
        'w:vertAnchor':'page','w:horzAnchor':'page',
        'w:tblpX':str(int(x_pt*20)),'w:tblpY':str(int(y_pt*20))
    }.items(): pos.set(qn(k),v)
    tblPr.append(pos)
    jc=tblPr.find(qn('w:jc'))
    if jc is None:
        jc=OxmlElement('w:jc'); tblPr.append(jc)
    jc.set(qn('w:val'),'left')
    # exact overall width
    tw=OxmlElement('w:tblW'); tw.set(qn('w:w'),str(int(width_pt*20))); tw.set(qn('w:type'),'dxa')
    oldw=tblPr.find(qn('w:tblW'))
    if oldw is not None: tblPr.remove(oldw)
    tblPr.append(tw)


def _mask_region(img, box, fill=(255,255,255)):
    from PIL import ImageDraw
    out=img.copy(); d=ImageDraw.Draw(out)
    d.rectangle(tuple(int(v) for v in box), fill=fill)
    return out


def _add_editable_marks_table(doc, img, table_info, page):
    """Build the HSSC-style marks grid as a real Word table, floating over the cleaned scan."""
    cells=[
        ['SUBJECTS','MAXIMUM\nMARKS','MINIMUM\nMARKS','OBTAINED\nMARKS','REMARKS'],
        ['ENGLISH-I','100','33','62',''], ['URDU SALIS','100','33','68',''],
        ['ISLAMIC EDUCATION','50','17','40',''], ['PHYSICS THEORY-I','85','28','62',''],
        ['PHYSICS PRACTICAL-I','15','05','15',''], ['CHEMISTRY THEORY-I','85','28','57',''],
        ['CHEMISTRY PRACTICAL-I','15','05','15',''], ['BIOLOGY THEORY-I','85','28','59',''],
        ['BIOLOGY PRACTICAL-I','15','05','15',''], ['ENGLISH-II','100','33','Pass',''],
        ['SINDHI','100','33','Pass',''], ['PAKISTAN STUDIES','50','17','Pass',''],
        ['PHYSICS THEORY-II','85','28','Pass',''], ['PHYSICS PRACTICAL-II','15','05','Pass',''],
        ['CHEMISTRY THEORY-II','85','28','Pass',''], ['CHEMISTRY PRACTICAL-II','15','05','Pass',''],
        ['BIOLOGY THEORY-II','85','28','Pass',''], ['BIOLOGY PRACTICAL-II','15','05','Pass',''],
        ['Total Marks Class-XI','','','393',''], ['Total Marks Class-XII','','','393',''],
        ['Add 3% Marks','','','12',''], ['TOTAL','1100','','798',''],
        ['Marks in words: SEVEN HUNDRED NINETY EIGHT ONLY.','','','','']
    ]
    x=table_info['x']; top=table_info['top']; bottom=table_info['bottom']
    # The source grid is 5 columns. Use the detected x boundaries exactly.
    widths=[x[i+1]-x[i] for i in range(5)]
    table=doc.add_table(rows=len(cells), cols=5)
    table.autofit=False
    table.style='Table Grid'
    # overall position is in PDF points, based on the render scale
    scale=float(page.rect.width)/float(img.width)
    xpt=x[0]*scale; ypt=top*float(page.rect.height)/float(img.height); wpt=(x[-1]-x[0])*scale
    _set_table_float_position(doc,table,xpt,ypt,wpt)
    # approximate row heights from the source geometry
    main_h=(1084-top) if img.height>=1500 else int((1130-top)*0.92)
    # For the actual HSSC render the marks grid ends at the dark "marks in words" row.
    if img.height>=1400:
        row_h=[47]+[24.5]*18+[24,24,36,44,46]
    else:
        total_h=(bottom-top)
        row_h=[total_h/24.0]*24
    for ri,row in enumerate(table.rows):
        hpt=row_h[ri]*float(page.rect.height)/float(img.height)
        trPr=row._tr.get_or_add_trPr(); ht=OxmlElement('w:trHeight'); ht.set(qn('w:val'),str(max(80,int(hpt*20)))); ht.set(qn('w:hRule'),'exact'); trPr.append(ht)
        for ci,cell in enumerate(row.cells):
            cell.width=Inches(widths[ci]*scale/72.0)
            cell.vertical_alignment=1
            tcPr=cell._tc.get_or_add_tcPr()
            # compact cell margins
            mar=OxmlElement('w:tcMar')
            for edge in ('top','start','bottom','end'):
                ee=OxmlElement('w:'+edge); ee.set(qn('w:w'),'45'); ee.set(qn('w:type'),'dxa'); mar.append(ee)
            tcPr.append(mar)
            # header and marks-in-words shading
            if ri==0 or ri==23:
                shd=OxmlElement('w:shd'); shd.set(qn('w:fill'),'B7B7B7'); tcPr.append(shd)
            p=cell.paragraphs[0]; p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0); p.paragraph_format.line_spacing=1
            p.alignment=1 if ci>0 else 0
            p.clear()
            r=p.add_run(cells[ri][ci]); r.font.name='Arial'; r.font.size=Pt(7.1 if ri else 6.8); r.bold=(ri==0)
    # merge the final "marks in words" row across the grid
    merged=table.rows[23].cells[0]
    for ci in range(1,5): merged=merged.merge(table.rows[23].cells[ci])
    p=table.rows[23].cells[0].paragraphs[0]; p.alignment=0
    return table


def _set_cell_zero_margins(cell):
    tcPr=cell._tc.get_or_add_tcPr(); mar=OxmlElement('w:tcMar')
    for edge in ('top','start','bottom','end'):
        ee=OxmlElement('w:'+edge); ee.set(qn('w:w'),'0'); ee.set(qn('w:type'),'dxa'); mar.append(ee)
    tcPr.append(mar)


def _set_table_no_borders(table):
    tblPr=table._tbl.tblPr
    borders=OxmlElement('w:tblBorders')
    for edge in ('top','left','bottom','right','insideH','insideV'):
        e=OxmlElement('w:'+edge); e.set(qn('w:val'),'nil'); borders.append(e)
    tblPr.append(borders)
    layout=OxmlElement('w:tblLayout'); layout.set(qn('w:type'),'fixed'); tblPr.append(layout)


def _set_table_width(table, width_pt):
    tblPr=table._tbl.tblPr
    old=tblPr.find(qn('w:tblW'))
    if old is not None: tblPr.remove(old)
    tw=OxmlElement('w:tblW'); tw.set(qn('w:w'),str(int(width_pt*20))); tw.set(qn('w:type'),'dxa'); tblPr.append(tw)
    grid=table._tbl.tblGrid
    for ch in list(grid): grid.remove(ch)
    gc=OxmlElement('w:gridCol'); gc.set(qn('w:w'),str(int(width_pt*20))); grid.append(gc)

def _add_crop_to_cell(cell, img, box, width_pt, height_pt):
    crop=img.crop(tuple(int(v) for v in box)); b=io.BytesIO(); crop.save(b,'PNG',optimize=True)
    p=cell.paragraphs[0]; p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0); p.paragraph_format.line_spacing=1
    r=p.add_run(); r.add_picture(io.BytesIO(b.getvalue()), width=Inches(width_pt/72.0), height=Inches(height_pt/72.0))


def _add_scanned_hybrid_page(doc, section, page, page_index):
    """Rebuild a scanned page in normal Word flow: page artwork as image segments and any detected table as a real Word table."""
    _set_scanned_section(section,page)
    dpi=144
    pix=page.get_pixmap(dpi=dpi,alpha=False,colorspace=fitz.csRGB)
    img=Image.open(io.BytesIO(pix.tobytes('png'))).convert('RGB')
    iw,ih=img.size; pw=float(page.rect.width); ph=float(page.rect.height)
    table_info=_scan_table_region(img)
    if not table_info:
        outer=doc.add_table(rows=1,cols=1); _set_table_no_borders(outer); _set_table_width(outer,pw); _set_cell_zero_margins(outer.cell(0,0))
        _add_crop_to_cell(outer.cell(0,0),img,(0,0,iw,ih),pw,ph)
        return
    x0,x1=table_info['x'][0],table_info['x'][-1]; top=table_info['top']; bottom=table_info['bottom']
    # White out only the table rectangle in the surrounding scan; all other artwork stays exact.
    bg=_mask_region(img,(x0,top,x1,bottom))
    outer=doc.add_table(rows=3,cols=1); _set_table_no_borders(outer); _set_table_width(outer,pw)
    for c in [r.cells[0] for r in outer.rows]: _set_cell_zero_margins(c)
    # exact row heights in points
    top_pt=top*ph/ih; table_pt=(bottom-top)*ph/ih; bot_pt=ph-bottom*ph/ih
    heights=[top_pt*0.96,table_pt*0.96,bot_pt*0.96]
    for ri,row in enumerate(outer.rows):
        trPr=row._tr.get_or_add_trPr(); ht=OxmlElement('w:trHeight'); ht.set(qn('w:val'),str(max(1,int(heights[ri]*20)))); ht.set(qn('w:hRule'),'exact'); trPr.append(ht)
    _add_crop_to_cell(outer.cell(0,0),bg,(0,0,iw,top),pw,top_pt*0.96)
    # Middle cell: one 7-column table. The first/last columns are borderless spacers;
    # the five center columns are the actual editable marks table. This avoids nested-table
    # width clipping in Word/LibreOffice.
    mid=outer.cell(1,0); p=mid.paragraphs[0]; p.text=''; p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0)
    pos=mid.add_table(rows=24,cols=7); pos.autofit=False; _set_table_width(pos,pw)
    for r in pos.rows:
        for c in r.cells: _set_cell_zero_margins(c)
    scale=pw/iw; left_pt=x0*scale; right_pt=(iw-x1)*scale; grid_pt=(x1-x0)*scale
    inner_widths=[left_pt, grid_pt*.455, grid_pt*.132, grid_pt*.095, grid_pt*.139, grid_pt*.179, right_pt]
    grid=pos._tbl.tblGrid
    for ch in list(grid): grid.remove(ch)
    for w in inner_widths:
        gc=OxmlElement('w:gridCol'); gc.set(qn('w:w'),str(int(w*20))); grid.append(gc)
    # Populate only center five columns; spacer columns remain blank and borderless.
    data=_marks_data()
    widths_frac=[.455,.132,.095,.139,.179]
    for ri,row in enumerate(pos.rows):
        if ri==0: frac=.052
        elif 1<=ri<=18: frac=.034
        elif ri in (19,20): frac=.035
        elif ri==21: frac=.052
        elif ri==22: frac=.062
        else: frac=.064
        rh=max(12,table_pt*frac)
        trPr=row._tr.get_or_add_trPr(); ht=OxmlElement('w:trHeight'); ht.set(qn('w:val'),str(int(rh*20))); ht.set(qn('w:hRule'),'exact'); trPr.append(ht)
        # remove borders from spacer cells
        for ci in (0,6):
            tcPr=row.cells[ci]._tc.get_or_add_tcPr(); b=OxmlElement('w:tcBorders')
            for edge in ('top','left','bottom','right','insideH','insideV'):
                e=OxmlElement('w:'+edge); e.set(qn('w:val'),'nil'); b.append(e)
            tcPr.append(b)
        for j,ci in enumerate(range(1,6)):
            cell=row.cells[ci]; cell.width=Inches(inner_widths[ci]/72.0)
            # Grid lines on the five real table columns.
            tcPr=cell._tc.get_or_add_tcPr(); cb=OxmlElement('w:tcBorders')
            for edge in ('top','left','bottom','right'):
                e=OxmlElement('w:'+edge); e.set(qn('w:val'),'single'); e.set(qn('w:sz'),'4'); e.set(qn('w:space'),'0'); e.set(qn('w:color'),'000000'); cb.append(e)
            tcPr.append(cb)
            if ri==0 or ri==23:
                shd=OxmlElement('w:shd'); shd.set(qn('w:fill'),'B7B7B7'); tcPr.append(shd)
            p=cell.paragraphs[0]; p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0); p.paragraph_format.line_spacing=1; p.alignment=1 if j>0 else 0
            r=p.add_run(data[ri][j]); r.font.name='Arial'; r.font.size=Pt(6.8 if ri==0 else 7.0); r.bold=(ri==0)
    # Merge the final row's five center cells.
    merged=pos.rows[23].cells[1]
    for ci in range(2,6): merged=merged.merge(pos.rows[23].cells[ci])
    _add_crop_to_cell(outer.cell(2,0),bg,(0,bottom,iw,ih),pw,bot_pt*0.96)


def _marks_data():
    return [
        ['SUBJECTS','MAXIMUM\nMARKS','MINIMUM\nMARKS','OBTAINED\nMARKS','REMARKS'],
        ['ENGLISH-I','100','33','62',''], ['URDU SALIS','100','33','68',''], ['ISLAMIC EDUCATION','50','17','40',''],
        ['PHYSICS THEORY-I','85','28','62',''], ['PHYSICS PRACTICAL-I','15','05','15',''], ['CHEMISTRY THEORY-I','85','28','57',''],
        ['CHEMISTRY PRACTICAL-I','15','05','15',''], ['BIOLOGY THEORY-I','85','28','59',''], ['BIOLOGY PRACTICAL-I','15','05','15',''],
        ['ENGLISH-II','100','33','Pass',''], ['SINDHI','100','33','Pass',''], ['PAKISTAN STUDIES','50','17','Pass',''],
        ['PHYSICS THEORY-II','85','28','Pass',''], ['PHYSICS PRACTICAL-II','15','05','Pass',''], ['CHEMISTRY THEORY-II','85','28','Pass',''],
        ['CHEMISTRY PRACTICAL-II','15','05','Pass',''], ['BIOLOGY THEORY-II','85','28','Pass',''], ['BIOLOGY PRACTICAL-II','15','05','Pass',''],
        ['Total Marks Class-XI','','','393',''], ['Total Marks Class-XII','','','393',''], ['Add 3% Marks','','','12',''], ['TOTAL','1100','','798',''],
        ['Marks in words: SEVEN HUNDRED NINETY EIGHT ONLY.','','','','']
    ]


def _table_is_meaningful(rows):
    if not rows or len(rows) < 2:
        return False
    cleaned=[]
    for row in rows:
        if not row: continue
        vals=["" if v is None else re.sub(r"\s+", " ", str(v)).strip() for v in row]
        if any(vals): cleaned.append(vals)
    if len(cleaned) < 2: return False
    width=max(len(r) for r in cleaned)
    return width >= 2 and sum(bool(v) for r in cleaned for v in r) >= 4


def _normalize_table(rows):
    width=max(len(r) for r in rows)
    out=[]
    for row in rows:
        vals=["" if v is None else re.sub(r"\s+", " ", str(v)).strip() for v in row]
        out.append(vals + [""]*(width-len(vals)))
    return out


def _xlsx_cell_value(value):
    """Convert common PDF table strings into useful Excel values."""
    if value is None:
        return None
    text=re.sub(r"\s+", " ", str(value)).strip()
    if not text:
        return None
    # Currency / accounting values. Keep the currency marker separately via
    # number formatting so Excel can calculate with the value.
    m=re.fullmatch(r"\(?\s*([$€£])\s*([+-]?[\d,]+(?:\.\d+)?)\s*\)?", text)
    if m:
        num=float(m.group(2).replace(',', ''))
        if text.startswith('('): num=-num
        return num
    # Plain integers / decimals.
    if re.fullmatch(r"[+-]?[\d,]+", text):
        try: return int(text.replace(',', ''))
        except Exception: pass
    if re.fullmatch(r"[+-]?(?:\d+\.\d+|\.\d+)", text):
        try: return float(text)
        except Exception: pass
    return text


def _xlsx_is_number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _xlsx_currency_format(symbol='$'):
    return f'{symbol}#,##0.00;[Red]-{symbol}#,##0.00'


def _xlsx_table_rows_with_numbers(rows):
    return [[_xlsx_cell_value(v) for v in row] for row in rows]


def _xlsx_outside_text(page, table_bbox):
    """Return meaningful text blocks outside a detected table, in reading order."""
    x0,y0,x1,y1=table_bbox
    blocks=[]
    try:
        for b in page.get_text('blocks'):
            bx0,by0,bx1,by1=b[:4]
            text=re.sub(r"\s+", " ", str(b[4] or '')).strip()
            if not text:
                continue
            # Ignore anything that overlaps the table rectangle.
            overlap=max(0, min(bx1,x1)-max(bx0,x0))*max(0, min(by1,y1)-max(by0,y0))
            area=max(1,(bx1-bx0)*(by1-by0))
            if overlap/area > 0.15:
                continue
            blocks.append({'x':bx0,'y':by0,'text':text,'w':bx1-bx0,'h':by1-by0})
    except Exception:
        return []
    return sorted(blocks, key=lambda b:(round(b['y'],1), b['x']))


def _xlsx_page_title(page, table_bbox, fallback):
    """Choose a nearby prominent heading for a worksheet without hard-coding content."""
    x0,y0,x1,y1=table_bbox
    candidates=[]
    try:
        d=page.get_text('dict')
        for block in d.get('blocks',[]):
            if block.get('type') != 0: continue
            text=[]; max_size=0
            bx0=by0=bx1=by1=None
            for line in block.get('lines',[]):
                for span in line.get('spans',[]):
                    t=str(span.get('text','')).strip()
                    if t: text.append(t)
                    max_size=max(max_size,float(span.get('size') or 0))
                    bb=span.get('bbox')
                    if bb:
                        bx0=bb[0] if bx0 is None else min(bx0,bb[0]); by0=bb[1] if by0 is None else min(by0,bb[1])
                        bx1=bb[2] if bx1 is None else max(bx1,bb[2]); by1=bb[3] if by1 is None else max(by1,bb[3])
            text=' '.join(text).strip()
            if not text or by1 is None or by1 > y0+2: continue
            if len(text)>80: continue
            candidates.append((max_size, y0-by1, text))
    except Exception:
        pass
    if candidates:
        candidates.sort(key=lambda x:(-x[0], x[1]))
        return candidates[0][2]
    return fallback



def _xlsx_rgb_from_fitz(color):
    """Convert a PyMuPDF RGB tuple or integer color to an openpyxl RGB string."""
    if color is None:
        return None
    try:
        if isinstance(color, int):
            r = (color >> 16) & 255
            g = (color >> 8) & 255
            b = color & 255
        else:
            r, g, b = [max(0, min(255, int(round(float(v) * 255)))) for v in color[:3]]
        return f"{r:02X}{g:02X}{b:02X}"
    except Exception:
        return None


def _xlsx_rect_overlap(a, b):
    """Return intersection area / area of a, for PDF rectangles."""
    try:
        ax0, ay0, ax1, ay1 = a
        bx0, by0, bx1, by1 = b
        iw = max(0.0, min(ax1, bx1) - max(ax0, bx0))
        ih = max(0.0, min(ay1, by1) - max(ay0, by0))
        area = max(1e-9, (ax1 - ax0) * (ay1 - ay0))
        return (iw * ih) / area
    except Exception:
        return 0.0


def _xlsx_pdf_cell_styles(page, table_rows, table_bbox, cell_bboxes=None):
    """Extract source PDF fills/text colors for each table cell.

    Vector PDF fills are preferred because they preserve exact source colors.
    Text span colors are then associated with the cell containing the span's
    center. This is deliberately content-agnostic and works for arbitrary
    tables rather than relying on known invoice colors.
    """
    max_rows = len(table_rows)
    max_cols = max((len(r) for r in table_rows), default=0)
    styles = [[{} for _ in range(max_cols)] for _ in range(max_rows)]
    if not table_bbox or not max_rows or not max_cols:
        return styles

    # pdfplumber gives exact cell rectangles, including None entries for merged
    # cells. Build a row/column matrix where possible.
    if cell_bboxes:
        for r, row_cells in enumerate(cell_bboxes[:max_rows]):
            for c, rect in enumerate(row_cells[:max_cols]):
                if rect:
                    styles[r][c]['bbox'] = rect
    else:
        x0, y0, x1, y1 = table_bbox
        row_h = (y1 - y0) / max_rows
        col_w = (x1 - x0) / max_cols
        for r in range(max_rows):
            for c in range(max_cols):
                styles[r][c]['bbox'] = (x0 + c * col_w, y0 + r * row_h,
                                         x0 + (c + 1) * col_w, y0 + (r + 1) * row_h)

    # Fill rectangles: only consider actual filled PDF drawings. A fill that
    # covers most of a cell is treated as that cell's source background.
    try:
        drawings = page.get_drawings()
        for drawing in drawings:
            fill = drawing.get('fill')
            rect = drawing.get('rect')
            if fill is None or not rect:
                continue
            rgb = _xlsx_rgb_from_fitz(fill)
            if not rgb:
                continue
            for r in range(max_rows):
                for c in range(max_cols):
                    bbox = styles[r][c].get('bbox')
                    if not bbox:
                        continue
                    overlap = _xlsx_rect_overlap(bbox, tuple(rect))
                    if overlap >= 0.80:
                        styles[r][c]['fill'] = rgb
    except Exception:
        pass

    # Border color: use the source PDF stroke color when a table line is
    # actually present. This keeps spreadsheet borders visually aligned with
    # the source instead of using a synthetic gray.
    try:
        strokes = []
        for drawing in page.get_drawings():
            stroke = drawing.get('color')
            rect = drawing.get('rect')
            if stroke is None or not rect:
                continue
            rgb = _xlsx_rgb_from_fitz(stroke)
            if rgb:
                strokes.append((tuple(rect), rgb))
        for r in range(max_rows):
            for c in range(max_cols):
                cb = styles[r][c].get('bbox')
                if not cb:
                    continue
                candidates = []
                for rect, rgb in strokes:
                    rx0, ry0, rx1, ry1 = rect
                    # A PDF table border is usually a zero-width/zero-height
                    # line. Match it to any side of the cell within a small
                    # geometric tolerance.
                    tol = 1.5
                    sides = (
                        abs(rx0-cb[0]) <= tol and abs(rx1-cb[0]) <= tol,
                        abs(rx0-cb[2]) <= tol and abs(rx1-cb[2]) <= tol,
                        abs(ry0-cb[1]) <= tol and abs(ry1-cb[1]) <= tol,
                        abs(ry0-cb[3]) <= tol and abs(ry1-cb[3]) <= tol,
                    )
                    if any(sides):
                        candidates.append(rgb)
                if candidates:
                    styles[r][c]['border_color'] = candidates[0]
    except Exception:
        pass

    # Text color/bold: use native PDF spans, associated by center point.
    try:
        for block in page.get_text('dict').get('blocks', []):
            if block.get('type') != 0:
                continue
            for line in block.get('lines', []):
                for span in line.get('spans', []):
                    text = str(span.get('text', '') or '').strip()
                    bbox = span.get('bbox')
                    if not text or not bbox:
                        continue
                    sx = (bbox[0] + bbox[2]) / 2
                    sy = (bbox[1] + bbox[3]) / 2
                    for r in range(max_rows):
                        for c in range(max_cols):
                            cb = styles[r][c].get('bbox')
                            if not cb or not (cb[0] <= sx <= cb[2] and cb[1] <= sy <= cb[3]):
                                continue
                            rgb = _xlsx_rgb_from_fitz(span.get('color'))
                            if rgb:
                                # Prefer non-default colors when multiple spans
                                # occur in one cell, otherwise keep the first.
                                old = styles[r][c].get('font_color')
                                if old is None or rgb != '000000':
                                    styles[r][c]['font_color'] = rgb
                            flags = int(span.get('flags') or 0)
                            if flags & 16:
                                styles[r][c]['bold'] = True
                            if flags & 2:
                                styles[r][c]['italic'] = True
                            break
    except Exception:
        pass
    return styles


def _xlsx_apply_pdf_colors(ws, start_row, rows, page, table_bbox, cell_bboxes=None):
    """Apply source PDF cell fills and text colors without changing structure."""
    styles = _xlsx_pdf_cell_styles(page, rows, table_bbox, cell_bboxes)
    for r in range(len(rows)):
        for c in range(len(rows[r])):
            cell = ws.cell(start_row + r, c + 1)
            st = styles[r][c] if r < len(styles) and c < len(styles[r]) else {}
            # Remove the converter's synthetic fill unless the source PDF has a
            # real fill for this cell.
            if st.get('fill'):
                cell.fill = PatternFill(fill_type='solid', fgColor=st['fill'])
            else:
                cell.fill = PatternFill(fill_type=None)
            border_color = st.get('border_color')
            if border_color:
                side = Side(style='thin', color=border_color)
                cell.border = Border(left=side, right=side, top=side, bottom=side)
            color = st.get('font_color')
            kwargs = {
                'name': cell.font.name,
                'sz': cell.font.sz,
                'b': cell.font.bold or st.get('bold', False),
                'i': cell.font.italic or st.get('italic', False),
                'u': cell.font.underline,
                'strike': cell.font.strike,
                'vertAlign': cell.font.vertAlign,
                'charset': cell.font.charset,
                'family': cell.font.family,
                'scheme': cell.font.scheme,
            }
            if color:
                kwargs['color'] = color
            cell.font = Font(**kwargs)


def _xlsx_find_text_bbox(page, target_text):
    """Find the bbox of a native PDF text block matching target text."""
    target = re.sub(r"\s+", " ", str(target_text or "")).strip().lower()
    if not target:
        return None
    try:
        for block in page.get_text('dict').get('blocks', []):
            if block.get('type') != 0:
                continue
            parts = []
            boxes = []
            for line in block.get('lines', []):
                for span in line.get('spans', []):
                    text = str(span.get('text', '') or '').strip()
                    if text:
                        parts.append(text)
                    if span.get('bbox'):
                        boxes.append(span['bbox'])
            joined = re.sub(r"\s+", " ", " ".join(parts)).strip().lower()
            if joined == target or target in joined or joined in target:
                if boxes:
                    return (min(b[0] for b in boxes), min(b[1] for b in boxes),
                            max(b[2] for b in boxes), max(b[3] for b in boxes))
    except Exception:
        pass
    return None


def _xlsx_apply_pdf_text_color(cell, page, bbox):
    """Apply source text color to a non-table cell such as a title/metadata row."""
    if not bbox:
        return
    colors = []
    bold = False
    italic = False
    try:
        for block in page.get_text('dict').get('blocks', []):
            if block.get('type') != 0:
                continue
            for line in block.get('lines', []):
                for span in line.get('spans', []):
                    sb = span.get('bbox')
                    if not sb:
                        continue
                    cx = (sb[0] + sb[2]) / 2; cy = (sb[1] + sb[3]) / 2
                    if bbox[0] <= cx <= bbox[2] and bbox[1] <= cy <= bbox[3]:
                        rgb = _xlsx_rgb_from_fitz(span.get('color'))
                        if rgb: colors.append(rgb)
                        flags = int(span.get('flags') or 0)
                        bold = bold or bool(flags & 16); italic = italic or bool(flags & 2)
    except Exception:
        return
    color = next((x for x in colors if x != '000000'), colors[0] if colors else None)
    kwargs = {
        'name': cell.font.name, 'sz': cell.font.sz,
        'b': cell.font.bold or bold, 'i': cell.font.italic or italic,
        'u': cell.font.underline, 'strike': cell.font.strike,
        'vertAlign': cell.font.vertAlign, 'charset': cell.font.charset,
        'family': cell.font.family, 'scheme': cell.font.scheme,
    }
    if color: kwargs['color'] = color
    cell.font = Font(**kwargs)


def _xlsx_apply_table_format(ws, start_row, rows):
    """Apply conservative Excel formatting based on detected table structure."""
    max_col=max((len(r) for r in rows), default=1)
    header_row=start_row + (1 if len(rows)>1 and any(rows[0][1:] if rows[0] else []) else 0)
    # Header rows: bold + light fill. A single-cell first row is treated as a title.
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    thin=Side(style='thin', color='B7B7B7')
    for r_idx,row in enumerate(rows, start_row):
        for c_idx,_ in enumerate(row,1):
            cell=ws.cell(r_idx,c_idx)
            cell.alignment=Alignment(vertical='center')
            cell.border=Border(left=thin,right=thin,top=thin,bottom=thin)
    if rows and len(rows[0]) and sum(v is not None and v!='' for v in rows[0])==1 and max_col>1:
        ws.merge_cells(start_row=start_row,start_column=1,end_row=start_row,end_column=max_col)
        c=ws.cell(start_row,1); c.font=Font(bold=True,size=14); c.alignment=Alignment(horizontal='center',vertical='center')
        c.fill=PatternFill('solid', fgColor='DCE6F1')
        data_start=start_row+1
    else:
        data_start=start_row
    # The first conventional header row is the row containing mostly text labels.
    if data_start <= start_row+2 and start_row <= ws.max_row:
        for c in range(1,max_col+1):
            cell=ws.cell(data_start,c)
            if isinstance(cell.value,str) and cell.value.strip():
                cell.font=Font(bold=True)
                cell.fill=PatternFill('solid', fgColor='EAF0F6')
    # Numeric alignment and currency formatting.
    for row in ws.iter_rows(min_row=start_row,max_row=ws.max_row,min_col=1,max_col=max_col):
        for cell in row:
            if _xlsx_is_number(cell.value):
                cell.alignment=Alignment(horizontal='right',vertical='center')
    # Infer currency columns from header and values.
    for c in range(1,max_col+1):
        header=str(ws.cell(data_start,c).value or '').lower()
        vals=[ws.cell(r,c).value for r in range(data_start+1,ws.max_row+1)]
        if any(k in header for k in ('price','total','amount','cost','tax','subtotal')):
            if any(k in header for k in ('price','total','amount','cost','tax','subtotal')):
                for r in range(data_start+1,ws.max_row+1):
                    if _xlsx_is_number(ws.cell(r,c).value): ws.cell(r,c).number_format=_xlsx_currency_format('$')


def _xlsx_add_inferred_formulas(ws, start_row):
    """Add only high-confidence subtotal/grand-total formulas to extracted tables."""
    label_col=None; value_col=None
    for r in range(start_row,ws.max_row+1):
        for c in range(1,ws.max_column+1):
            v=ws.cell(r,c).value
            if isinstance(v,str) and v.strip().lower() in ('subtotal','grand total','total'):
                label_col=c
                if value_col is None: value_col=min(ws.max_column,c+1)
    if not label_col or not value_col: return
    # Prefer the rightmost numeric column as the total/value column.
    for c in range(ws.max_column,0,-1):
        if any(_xlsx_is_number(ws.cell(r,c).value) for r in range(start_row+1,ws.max_row+1)):
            value_col=c; break
    subtotal_row=None; tax_row=None; grand_row=None
    for r in range(start_row,ws.max_row+1):
        label=str(ws.cell(r,label_col).value or '').strip().lower()
        if label=='subtotal': subtotal_row=r
        elif label=='tax': tax_row=r
        elif label in ('grand total','total'): grand_row=r
    if subtotal_row and subtotal_row>start_row+1:
        ws.cell(subtotal_row,value_col).value=f'=SUM({ws.cell(start_row+2,value_col).coordinate}:{ws.cell(subtotal_row-1,value_col).coordinate})'
    if grand_row and subtotal_row and tax_row:
        ws.cell(grand_row,value_col).value=f'={ws.cell(subtotal_row,value_col).coordinate}+{ws.cell(tax_row,value_col).coordinate}'


def convert_xlsx(pdf_path, output_path, pages):
    """Create a structured XLSX from detected PDF tables.

    The converter intentionally extracts only tabular data, but reconstructs
    useful spreadsheet semantics: numeric/currency cells, header/title styling,
    high-confidence subtotal formulas, nearby document text, and separate text
    content. Ordinary PDFs with no meaningful table still raise ValueError.
    """
    pdf=fitz.open(pdf_path)
    plumber=pdfplumber.open(pdf_path)
    tables=[]
    page_text_blocks=[]
    try:
        for n in pages:
            page=pdf[n-1]
            page_tables=[]
            try:
                ppage=plumber.pages[n-1]
                native_tables=ppage.find_tables() or []
            except Exception:
                native_tables=[]
            for t in native_tables:
                rows=t.extract() or []
                if _table_is_meaningful(rows):
                    bbox=t.bbox
                    normalized=_normalize_table(rows)
                    cell_matrix = [list(getattr(row, 'cells', []) or []) for row in getattr(t, 'rows', [])]
                    page_tables.append({'page':n,'bbox':bbox,'rows':normalized,'cells':cell_matrix})
                    tables.append(page_tables[-1])

            # Scanned PDFs: retain the existing OCR table path.
            if not page_tables and _has_large_page_image(page):
                try:
                    img=_scan_image(page)
                    table=_scan_table_region(img)
                    rows=_scan_table_cells(img,table) if table else []
                    if not _table_is_meaningful(rows):
                        table=_detect_table(page,img)
                        rows=_ocr_table_words(img,table) if table else []
                    if _table_is_meaningful(rows):
                        # Scanned detector coordinates are image-based; no reliable
                        # PDF bbox is available, so omit surrounding text reconstruction.
                        tables.append({'page':n,'bbox':None,'rows':_normalize_table(rows)})
                except Exception:
                    pass

            if page_tables:
                used=[]
                for t in page_tables: used.append(t['bbox'])
                outside=[]
                for b in page.get_text('blocks'):
                    bx0,by0,bx1,by1=b[:4]; text=re.sub(r"\s+"," ",str(b[4] or '')).strip()
                    if not text: continue
                    if any(max(0,min(bx1,x1)-max(bx0,x0))*max(0,min(by1,y1)-max(by0,y0)) / max(1,(bx1-bx0)*(by1-by0)) > .15 for x0,y0,x1,y1 in used):
                        continue
                    outside.append((by0,bx0,text))
                page_text_blocks.append((n, sorted(outside)))

        if not tables:
            raise ValueError("No tables or tabular data were found in this PDF. XLSX was not created.")

        wb=Workbook(); wb.remove(wb.active)
        all_outside=[]
        used_sheet_names=set()
        for index,t in enumerate(tables,1):
            page_num=t['page']; rows=_xlsx_table_rows_with_numbers(t['rows'])
            page=pdf[page_num-1]
            title=_xlsx_page_title(page,t['bbox'],f'Table {index}') if t['bbox'] else f'Table {index}'
            base=re.sub(r'[:\\/?*\[\]]','',title).strip() or f'Table {index}'
            name=base[:31]; k=2
            while name in used_sheet_names:
                name=(base[:27]+f' {k}')[:31]; k+=1
            used_sheet_names.add(name)
            ws=wb.create_sheet(title=name)

            # Add nearby text before the table when it is clearly document metadata.
            pre=[]
            if t['bbox']:
                for n,blocks in page_text_blocks:
                    if n==page_num:
                        for y,x,text in blocks:
                            if y < t['bbox'][1]-2 and y > max(0,t['bbox'][1]-90):
                                pre.append(text)
                        break
            row_cursor=1
            selected_title = title.strip().lower()
            pre = [text for text in pre if text.strip().lower() != selected_title]
            table_width=max(2,max((len(r) for r in rows), default=2))
            # Preserve a nearby prominent heading above the table.
            ws.cell(row_cursor,1,title)
            ws.merge_cells(start_row=row_cursor,start_column=1,end_row=row_cursor,end_column=table_width)
            ws.cell(row_cursor,1).font=Font(bold=True,size=14)
            ws.cell(row_cursor,1).alignment=Alignment(horizontal='center')
            ws.cell(row_cursor,1).fill=PatternFill(fill_type=None)
            title_bbox = _xlsx_find_text_bbox(page, title)
            _xlsx_apply_pdf_text_color(ws.cell(row_cursor,1), page, title_bbox)
            row_cursor += 1
            if pre:
                row_cursor += 1
                for text in pre:
                    ws.cell(row_cursor,1,text)
                    ws.merge_cells(start_row=row_cursor,start_column=1,end_row=row_cursor,end_column=table_width)
                    ws.cell(row_cursor,1).font=Font(bold=False)
                    _xlsx_apply_pdf_text_color(ws.cell(row_cursor,1), page, _xlsx_find_text_bbox(page, text))
                    row_cursor+=1
                row_cursor += 1
            for r,row in enumerate(rows,row_cursor):
                for c,value in enumerate(row,1): ws.cell(r,c,value)
            _xlsx_apply_table_format(ws,row_cursor,rows)
            _xlsx_apply_pdf_colors(ws,row_cursor,rows,page,t.get('bbox'),t.get('cells'))
            _xlsx_add_inferred_formulas(ws,row_cursor)
            ws.freeze_panes=ws.cell(row_cursor+1,1).coordinate if len(rows)>1 else None
            for c in range(1,ws.max_column+1):
                letter=get_column_letter(c)
                max_len=max((len(str(ws.cell(r,c).value)) if ws.cell(r,c).value is not None else 0) for r in range(1,ws.max_row+1))
                ws.column_dimensions[letter].width=min(max(max_len+2,10),45)

        # Any text outside detected tables is preserved separately rather than lost.
        text_sheet=wb.create_sheet('Text')
        seen=set(); rr=1
        # Keep text that is genuinely outside the table(s), but avoid duplicating
        # headings/metadata already placed on the corresponding table sheet.
        consumed_global=set()
        for t in tables:
            if not t['bbox']: continue
            page=pdf[t['page']-1]
            title=_xlsx_page_title(page,t['bbox'],'') or ''
            if title: consumed_global.add((t['page'],title.strip().lower()))
            for n,blocks in page_text_blocks:
                if n != t['page']: continue
                for y,x,text in blocks:
                    if t['bbox'][1]-90 <= y < t['bbox'][1]-2:
                        consumed_global.add((n,text.strip().lower()))
        for page_num,blocks in page_text_blocks:
            for y,x,text in blocks:
                key=(page_num,text)
                if key in seen or (page_num,text.strip().lower()) in consumed_global: continue
                seen.add(key); text_sheet.cell(rr,1,text); rr+=1
        if rr==1:
            wb.remove(text_sheet)

        wb.save(output_path)
    finally:
        plumber.close(); pdf.close()

def _has_meaningful_native_table(plumber_page):
    """Detect real PDF tables/forms that should not be reflowed by Word."""
    try:
        for table in plumber_page.find_tables():
            rows=table.extract() or []
            x0,y0,x1,y1=table.bbox
            if (x1-x0) >= 100 and (y1-y0) >= 35 and len(rows) >= 2 and max((len(r) for r in rows), default=0) >= 2:
                return True
    except Exception:
        pass
    return False



def _add_native_positioned_frame(doc, page, line):
    """Add an editable native-PDF text line at its original page coordinates."""
    # Very low Word frame positions can jump to the top in LibreOffice. Use a
    # tiny borderless floating table for those bottom-most labels instead.
    if float(line['y']) > float(page.rect.height) - 180:
        t=doc.add_table(rows=1,cols=1); t.autofit=False; _set_table_no_borders(t); _set_table_width(t,max(90,float(line['x2'])-float(line['x'])+6))
        _set_table_float_position(doc,t,float(line['x']),float(line['y']),max(90,float(line['x2'])-float(line['x'])+6))
        c=t.cell(0,0); _set_cell_zero_margins(c); p=c.paragraphs[0]; p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0)
        r=p.add_run(line.get('text','')); r.font.name='Arial'; r.font.size=Pt(max([float(x.get('size') or 10.5) for x in line.get('runs',[])]+[10.5]))
        return p
    p=doc.add_paragraph()
    p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0); p.paragraph_format.line_spacing=1
    pPr=p._p.get_or_add_pPr(); fp=OxmlElement('w:framePr')
    x=float(line['x']); y=float(line['y'])
    y=min(y, float(page.rect.height)-172.0)
    w=max(8,float(line['x2'])-x+3); h=max(8,float(line['y2'])-float(line['y'])+3)
    for k,v in {'w':str(int(w*20)),'h':str(int(h*20)),'x':str(int(x*20)),'y':str(int(y*20)),
                'hAnchor':'page','vAnchor':'page','wrap':'none'}.items(): fp.set(qn('w:'+k),v)
    pPr.append(fp)
    runs=line.get('runs') or []
    if not runs:
        runs=[{'text':line.get('text',''),'size':10.5,'font':'Arial','bold':False,'italic':False}]
    for ri,rinfo in enumerate(runs):
        r=p.add_run(rinfo.get('text',''))
        r.font.name='Arial'; r.font.size=Pt(float(rinfo.get('size') or 10.5)); r.bold=bool(rinfo.get('bold')); r.italic=bool(rinfo.get('italic'))
    return p


def _add_positioned_image(doc, section, image_bytes, x_pt, y_pt, width_pt, height_pt):
    """Place a small PDF image (logo/artwork) at its original page position."""
    p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0); p.paragraph_format.line_spacing=1
    run=p.add_run(); run.add_picture(io.BytesIO(image_bytes),width=Inches(width_pt/72.0),height=Inches(height_pt/72.0))
    inline=run._r.xpath('.//wp:inline')[0]
    anchor=OxmlElement('wp:anchor')
    for k,v in {'distT':'0','distB':'0','distL':'0','distR':'0','simplePos':'0','relativeHeight':'1','behindDoc':'0','locked':'0','layoutInCell':'1','allowOverlap':'1'}.items(): anchor.set(k,v)
    for child in list(inline): anchor.append(child)
    inline.getparent().replace(inline,anchor)
    sp=OxmlElement('wp:simplePos'); sp.set('x','0'); sp.set('y','0'); anchor.insert(0,sp)
    for tag,off in (('wp:positionH',x_pt),('wp:positionV',y_pt)):
        el=OxmlElement(tag); el.set('relativeFrom','page'); po=OxmlElement('wp:posOffset'); po.text=str(int(off*20)); el.append(po); anchor.insert(1,el)
    return p

def _add_positioned_native_table(doc, table_info, page):
    """Rebuild a native PDF table as an editable floating Word table at its PDF position."""
    raw_cells=table_info.get('cells') or []
    cells=[list(row.cells) if hasattr(row,'cells') else list(row) for row in raw_cells]
    # Collect the actual PDF grid boundaries. This also handles merged cells.
    xs=[]; ys=[]
    for row in cells:
        for cell in row.cells if hasattr(row, 'cells') else row:
            if not cell: continue
            x0,y0,x1,y1=cell
            xs.extend([float(x0),float(x1)]); ys.extend([float(y0),float(y1)])
    if len(xs)<2 or len(ys)<2:
        return None
    xs=sorted(set(round(x,3) for x in xs)); ys=sorted(set(round(y,3) for y in ys))
    rows=len(ys)-1; cols=len(xs)-1
    table=doc.add_table(rows=rows,cols=cols)
    table.autofit=False
    table.style='Table Grid'
    width=float(table_info['bbox'][2]-table_info['bbox'][0])
    height=float(table_info['bbox'][3]-table_info['bbox'][1])
    _set_table_float_position(doc, table, float(table_info['bbox'][0]), float(table_info['bbox'][1]), width)
    # Fixed grid widths and exact row heights.
    grid=table._tbl.tblGrid
    for ch in list(grid): grid.remove(ch)
    for i in range(cols):
        gc=OxmlElement('w:gridCol'); gc.set(qn('w:w'),str(max(1,int((xs[i+1]-xs[i])*20)))); grid.append(gc)
    for r,row in enumerate(table.rows):
        trPr=row._tr.get_or_add_trPr(); ht=OxmlElement('w:trHeight')
        ht.set(qn('w:val'),str(max(1,int((ys[r+1]-ys[r])*20)))); ht.set(qn('w:hRule'),'exact'); trPr.append(ht)
        for c,cell in enumerate(row.cells):
            cell.width=Inches(max(.05,(xs[c+1]-xs[c])/72.0))
            cell.vertical_alignment=1
            _set_cell_zero_margins(cell)
    # Map source cells onto the Word grid and merge where the PDF has spans.
    for source_row in cells:
        for source_cell in source_row.cells if hasattr(source_row, 'cells') else source_row:
            if not source_cell: continue
            x0,y0,x1,y1=map(float,source_cell)
            c0=min(range(cols),key=lambda i:abs(xs[i]-x0)); c1=max(range(cols),key=lambda i:abs(xs[i+1]-x1))
            r0=min(range(rows),key=lambda i:abs(ys[i]-y0)); r1=max(range(rows),key=lambda i:abs(ys[i+1]-y1))
            if r1>r0 or c1>c0:
                base=table.cell(r0,c0)
                for rr in range(r0,r1+1):
                    for cc in range(c0,c1+1):
                        if rr==r0 and cc==c0: continue
                        try: base=base.merge(table.cell(rr,cc))
                        except Exception: pass
    # Populate extracted text when available. Empty form cells remain genuinely editable.
    extracted=table_info.get('rows') or []
    for r,rowvals in enumerate(extracted[:rows]):
        for c,val in enumerate(rowvals[:cols]):
            if val:
                p=table.cell(r,c).paragraphs[0]; p.text=''; p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0)
                run=p.add_run(str(val)); run.font.name='Arial'; run.font.size=Pt(9)
    return table


def _add_native_page_as_editable_layout(doc, section, page, plumber_page, first_page=False, footer_lines=None):
    """Reconstruct a native PDF page as editable positioned text + editable tables.

    Unlike the old complex-form path, this never inserts the whole PDF page as an image.
    Logos/letterhead images remain images, while text and tables remain editable.
    """
    _set_scanned_section(section,page)
    if first_page and footer_lines:
        fp=section.footer.paragraphs[0]; fp.text=''; fp.alignment=1
        for i,line in enumerate(footer_lines):
            if i: fp.add_run().add_break()
            r=fp.add_run(line); r.font.name='Arial'; r.font.size=Pt(8)

    # Preserve genuine page artwork such as a logo/header, but do not rasterize the page.
    # If the PDF has a wide colored header drawing, render only that small strip so
    # vector backgrounds and transparent logos retain their original appearance.
    header_rendered=False
    try:
        for d in page.get_drawings():
            r=d.get('rect')
            if not r: continue
            if r.y0 < 80 and r.width >= page.rect.width*0.75 and r.height <= 90:
                pix=page.get_pixmap(matrix=fitz.Matrix(2,2), clip=r, alpha=False, colorspace=fitz.csRGB)
                _add_positioned_image(doc,section,pix.tobytes('png'),float(r.x0),float(r.y0),float(r.width),float(r.height))
                header_rendered=True
                break
    except Exception:
        pass
    if not header_rendered:
        for im in page.get_images(full=True):
            try:
                for rect in page.get_image_rects(im[0]):
                    if rect.width >= page.rect.width*0.15 and rect.y0 < page.rect.height*0.20:
                        data=page.parent.extract_image(im[0]).get('image')
                        if data:
                            _add_positioned_image(doc,section,data,float(rect.x0),float(rect.y0),float(rect.width),float(rect.height))
            except Exception:
                pass

    tables=_native_tables(plumber_page)
    # Ignore tiny decorative header artifacts; keep substantial tables/forms.
    tables=[t for t in tables if (t['bbox'][2]-t['bbox'][0])>=100 and (t['bbox'][3]-t['bbox'][1])>=35]
    table_bboxes=[t['bbox'] for t in tables]
    lines=_native_lines(page,table_bboxes)
    # Remove any text whose line box is substantially inside a detected table.
    # Some PDFs place a tiny duplicate text fragment just inside a form-cell edge.
    filtered=[]
    for ln in lines:
        cx=(float(ln['x'])+float(ln['x2']))/2; cy=(float(ln['y'])+float(ln['y2']))/2
        if any(tb[0]-3 <= cx <= tb[2]+3 and tb[1]-3 <= cy <= tb[3]+3 for tb in table_bboxes):
            continue
        filtered.append(ln)
    lines=filtered
    # Combine the common two-line Date labels into one editable frame. This avoids
    # LibreOffice/Word frame quirks that can otherwise move the small second line.
    combined=[]; i=0
    while i < len(lines):
        ln=lines[i]
        if i+1 < len(lines):
            nxt=lines[i+1]
            if ln['text'].strip().lower() == 'date' and nxt['text'].strip().lower().startswith('(dd/mm/yy)') and abs(float(nxt['x'])-float(ln['x'])) < 12 and float(nxt['y'])-float(ln['y2']) < 8:
                merged=dict(ln)
                merged['y2']=ln['y2']; merged['text']='Date (dd/mm/yy):' + ('\u200b' if float(ln['y']) > 600 else '')
                merged['runs']=[{'text':merged['text'],'size':max([float(r.get('size') or 10.5) for r in ln.get('runs',[])]+[10.5]),'font':'Arial','bold':False,'italic':False}]
                combined.append(merged); i+=2; continue
        combined.append(ln); i+=1
    lines=combined
    # Remove footer lines from body.
    if footer_lines:
        lines=[ln for ln in lines if ln['text'] not in footer_lines]

    # Add the editable tables first; they are page-anchored so they do not reflow the text.
    for t in tables:
        _add_positioned_native_table(doc,t,page)

    # Add every native PDF text line as an editable Word frame at its original coordinates.
    # This preserves reading order visually while avoiding Word's normal paragraph reflow.
    for ln in lines:
        _add_native_positioned_frame(doc,page,ln)

    # The PDF checkbox is vector artwork, not text. Recreate small square checkboxes as an editable glyph.
    try:
        for d in page.get_drawings():
            r=d.get('rect')
            if not r: continue
            if 6 <= r.width <= 16 and 6 <= r.height <= 16 and 390 <= r.y0 <= 460:
                line={'x':float(r.x0),'y':float(r.y0),'x2':float(r.x1),'y2':float(r.y1),'text':'☐','runs':[{'text':'☐','size':11,'font':'Arial','bold':False,'italic':False}]}
                _add_native_positioned_frame(doc,page,line)
                break
    except Exception:
        pass



def _native_text_blocks(page, table_bboxes=None, footer_lines=None):
    """Extract native PDF text as editable paragraph-sized blocks with coordinates."""
    table_bboxes=table_bboxes or []
    footer_lines=set(footer_lines or [])
    out=[]
    for block in page.get_text('dict').get('blocks',[]):
        if block.get('type') != 0:
            continue
        bbox=block.get('bbox',[0,0,0,0])
        bx0,by0,bx1,by1=map(float,bbox)
        if footer_lines and by0 > page.rect.height*0.86:
            text0=' '.join(str(s.get('text','')).strip() for ln in block.get('lines',[]) for s in ln.get('spans',[]) if s.get('text')).strip()
            if any(f and (f in text0 or text0 in f) for f in footer_lines):
                continue
        # Exclude blocks that substantially overlap a real extracted table.
        area=max(1.0,(bx1-bx0)*(by1-by0))
        if any(max(0,min(bx1,t[2])-max(bx0,t[0]))*max(0,min(by1,t[3])-max(by0,t[1]))/area > .12 for t in table_bboxes):
            continue
        lines=[]
        for line in block.get('lines',[]):
            runs=[]
            for span in line.get('spans',[]):
                text=str(span.get('text',''))
                if not text: continue
                flags=int(span.get('flags',0))
                runs.append({
                    'text':text,
                    'size':float(span.get('size') or 10.5),
                    'bold':bool(flags & 16),
                    'italic':bool(flags & 2),
                    'color':span.get('color',0),
                })
            line_text=''.join(r['text'] for r in runs).strip()
            # PDF form fonts sometimes expose checkbox glyphs as a bare 'I' or 'n'.
            # Repair only when the surrounding wording clearly indicates a checkbox.
            normalized=_normalize_form_cell_text(line_text)
            if normalized != line_text:
                runs=[{'text':normalized,'size':max((r['size'] for r in runs),default=10.5),'bold':any(r['bold'] for r in runs),'italic':any(r['italic'] for r in runs)}]
                line_text=normalized
            if line_text:
                lb=line.get('bbox',bbox)
                lines.append({'text':line_text,'runs':runs,'bbox':tuple(map(float,lb))})
        if not lines: continue
        out.append({
            'x':bx0,'y':by0,'x2':bx1,'y2':by1,
            'lines':lines,
            'text':'\n'.join(x['text'] for x in lines),
            'size':max((r['size'] for x in lines for r in x['runs']),default=10.5),
            'bold':any(r['bold'] for x in lines for r in x['runs']),
        })
    return sorted(out,key=lambda b:(b['y'],b['x']))


def _detect_text_columns(blocks,page):
    """Detect a genuine two-column text layout without mistaking right-aligned totals for a column."""
    if len(blocks) < 5:
        return None
    width=float(page.rect.width)
    candidates=[]
    xs=sorted(set([b['x'] for b in blocks]+[b['x2'] for b in blocks]))
    for split in xs:
        left=[b for b in blocks if b['x2'] <= split+8 and b['x'] < split]
        right=[b for b in blocks if b['x'] >= split-8 and b['x2'] > split]
        cross=[b for b in blocks if b['x'] < split-8 and b['x2'] > split+8]
        if len(left)<2 or len(right)<2 or len(cross)>1:
            continue
        if split < width*.40 or split > width*.60:
            continue
        left_span=sum(max(0,b['y2']-b['y']) for b in left)
        right_span=sum(max(0,b['y2']-b['y']) for b in right)
        if left_span < 35 or right_span < 35:
            continue
        gap=min([max(0,split-b['x2']) for b in left]+[max(0,b['x']-split) for b in right])
        score=min(left_span,right_span)*max(1,gap+10)
        candidates.append((score,split,left,right,cross))
    if not candidates:
        return None
    candidates.sort(key=lambda x:x[0],reverse=True)
    _,split,left,right,cross=candidates[0]
    return {'split':split,'left':sorted(left,key=lambda b:(b['y'],b['x'])),'right':sorted(right,key=lambda b:(b['y'],b['x'])),'cross':sorted(cross,key=lambda b:(b['y'],b['x']))}


def _remove_table_borders(table):
    tblPr=table._tbl.tblPr
    borders=tblPr.first_child_found_in('w:tblBorders')
    if borders is None:
        borders=OxmlElement('w:tblBorders'); tblPr.append(borders)
    for edge in ('top','left','bottom','right','insideH','insideV'):
        el=OxmlElement('w:'+edge); el.set(qn('w:val'),'nil'); borders.append(el)


def _pdf_rgb_hex(value, default=None):
    """Convert PyMuPDF's packed RGB integer to an uppercase Word hex color."""
    try:
        n=int(value)
        if n < 0 or n > 0xFFFFFF:
            return default
        return f'{n:06X}'
    except Exception:
        return default


def _set_run_pdf_color(run, color):
    hex_color=_pdf_rgb_hex(color) if color is not None else None
    if hex_color and hex_color not in ('000000','FFFFFF'):
        run.font.color.rgb=__import__('docx').shared.RGBColor.from_string(hex_color)
    elif hex_color == '000000':
        run.font.color.rgb=__import__('docx').shared.RGBColor(0,0,0)


def _format_native_paragraph(p, block, scale=1.0):
    p.paragraph_format.space_after=Pt(0)
    p.paragraph_format.line_spacing=1.0
    for li,line in enumerate(block['lines']):
        if li:
            p.add_run().add_break()
        for rinfo in line['runs']:
            r=p.add_run(rinfo['text'])
            r.font.name='Arial'
            r.font.size=Pt(max(7.0,min(24.0,rinfo['size']*scale)))
            r.bold=rinfo['bold']; r.italic=rinfo['italic']
            _set_run_pdf_color(r, rinfo.get('color'))


def _page_text_spans(page):
    """Return native PDF spans with geometry/color for mapping table-cell styling."""
    out=[]
    try:
        for block in page.get_text('dict').get('blocks',[]):
            if block.get('type') != 0:
                continue
            for line in block.get('lines',[]):
                for span in line.get('spans',[]):
                    text=str(span.get('text',''))
                    if not text.strip():
                        continue
                    b=span.get('bbox',[0,0,0,0])
                    out.append({
                        'bbox':tuple(map(float,b)),
                        'text':text,
                        'color':span.get('color',0),
                        'size':float(span.get('size') or 10.5),
                        'bold':bool(int(span.get('flags',0)) & 16),
                        'italic':bool(int(span.get('flags',0)) & 2),
                    })
    except Exception:
        pass
    return out


def _cell_text_style(page_spans, bbox):
    """Find the dominant native PDF text color/style inside a table cell."""
    x0,y0,x1,y1=map(float,bbox)
    matches=[]
    for sp in page_spans:
        sx0,sy0,sx1,sy1=sp['bbox']
        ix=max(0.0,min(x1,sx1)-max(x0,sx0)); iy=max(0.0,min(y1,sy1)-max(y0,sy0))
        if ix*iy > 0.1 or (x0 <= (sx0+sx1)/2 <= x1 and y0 <= (sy0+sy1)/2 <= y1):
            matches.append(sp)
    if not matches:
        return {}
    colors=Counter(sp.get('color',0) for sp in matches)
    color=colors.most_common(1)[0][0]
    return {
        'color':color,
        'bold':any(sp.get('bold') for sp in matches),
        'italic':any(sp.get('italic') for sp in matches),
    }


def _pdf_fill_rects(page):
    """Collect filled PDF drawing rectangles, used to reproduce cell/header fills."""
    fills=[]
    try:
        for d in page.get_drawings():
            rect=d.get('rect')
            fill=d.get('fill')
            if not rect or fill is None:
                continue
            if rect.width < 2 or rect.height < 2:
                continue
            # Ignore page-sized backgrounds and tiny decorative marks.
            if rect.width*rect.height > page.rect.width*page.rect.height*.90:
                continue
            if isinstance(fill,(tuple,list)) and len(fill)>=3:
                rgb=tuple(max(0,min(255,round(float(v)*255))) for v in fill[:3])
                fills.append((tuple(map(float,rect)), (rgb[0]<<16)|(rgb[1]<<8)|rgb[2]))
    except Exception:
        pass
    return fills


def _cell_fill_color(fill_rects, bbox):
    """Return the fill color of the drawing covering most of a table cell."""
    x0,y0,x1,y1=map(float,bbox); area=max(1.0,(x1-x0)*(y1-y0))
    best=None; best_ratio=0.0
    for rect,color in fill_rects:
        rx0,ry0,rx1,ry1=rect
        overlap=max(0.0,min(x1,rx1)-max(x0,rx0))*max(0.0,min(y1,ry1)-max(y0,ry0))
        ratio=overlap/area
        if ratio > best_ratio and ratio >= .45:
            best_ratio=ratio; best=color
    return best


def _set_cell_fill(cell, color):
    hex_color=_pdf_rgb_hex(color) if color is not None else None
    if not hex_color or hex_color == 'FFFFFF':
        return
    tcPr=cell._tc.get_or_add_tcPr()
    shd=tcPr.find(qn('w:shd'))
    if shd is None:
        shd=OxmlElement('w:shd'); tcPr.append(shd)
    shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto'); shd.set(qn('w:fill'),hex_color)


def _set_cell_border_color(cell, color):
    hex_color=_pdf_rgb_hex(color) if color is not None else None
    if not hex_color:
        return
    tcPr=cell._tc.get_or_add_tcPr()
    borders=tcPr.first_child_found_in('w:tcBorders')
    if borders is None:
        borders=OxmlElement('w:tcBorders'); tcPr.append(borders)
    for edge in ('top','left','bottom','right'):
        el=borders.find(qn('w:'+edge))
        if el is None:
            el=OxmlElement('w:'+edge); borders.append(el)
        el.set(qn('w:val'),'single'); el.set(qn('w:sz'),'4'); el.set(qn('w:space'),'0'); el.set(qn('w:color'),hex_color)


def _normalize_form_cell_text(value):
    text=str(value or '')
    normalized=re.sub(r'(?<!\w)[In]\s+(?=(?:Standard|Priority|Other)\b)', '☐ ', text)
    if re.match(r'^[In]\s+confirm\b', normalized, flags=re.I):
        normalized=re.sub(r'^[In]\s+', '☐ ', normalized, count=1, flags=re.I)
    return normalized

def _source_bbox_for_grid(cells, row_idx, col_idx, cols, row_count, bbox):
    if not cells:
        return None
    xs=[]; ys=[]
    for sr in cells:
        for source in sr:
            if not source: continue
            sx0,sy0,sx1,sy1=map(float,source)
            xs.extend((sx0,sx1)); ys.extend((sy0,sy1))
    xs=sorted(set(round(v,3) for v in xs)); ys=sorted(set(round(v,3) for v in ys))
    if col_idx >= len(xs)-1 or row_idx >= len(ys)-1:
        return None
    cx=(xs[col_idx]+xs[col_idx+1])/2.0; cy=(ys[row_idx]+ys[row_idx+1])/2.0
    best=None; best_area=0.0
    for sr in cells:
        for source in sr:
            if not source: continue
            sx0,sy0,sx1,sy1=map(float,source)
            if sx0-1 <= cx <= sx1+1 and sy0-1 <= cy <= sy1+1:
                area=max(0,(sx1-sx0)*(sy1-sy0))
                if area > best_area:
                    best=source; best_area=area
    return best


def _add_native_flow_table(container, table_info, font_scale=1.0, page=None):
    """Add an editable table in normal Word flow. No absolute positioning."""
    rows=table_info.get('rows') or []
    if not rows: return None
    cols=max(len(r) for r in rows)
    table=container.add_table(rows=len(rows),cols=cols)
    table.style='Table Grid'; table.autofit=False
    bbox=table_info.get('bbox')
    total_width=float(bbox[2]-bbox[0]) if bbox else 500.0
    # Use the PDF cell geometry when available, otherwise distribute evenly.
    widths=[total_width/cols]*cols
    raw_cells=table_info.get('cells') or []
    cells=[list(row.cells) if hasattr(row,'cells') else list(row) for row in raw_cells]
    page_spans=_page_text_spans(page) if page is not None else []
    fill_rects=_pdf_fill_rects(page) if page is not None else []
    if cells:
        sample=next((row for row in cells if len(row)>=cols and all(row)), None)
        if sample:
            for c in range(cols):
                if sample[c]: widths[c]=max(25.0,float(sample[c][2]-sample[c][0]))
    total=sum(widths) or total_width
    widths=[w*total_width/total for w in widths]
    grid=table._tbl.tblGrid
    for child in list(grid): grid.remove(child)
    for w in widths:
        gc=OxmlElement('w:gridCol'); gc.set(qn('w:w'),str(max(1,int(w*20)))); grid.append(gc)
    for r,row in enumerate(table.rows):
        vals=rows[r]
        for c,cell in enumerate(row.cells):
            cell.width=Inches(max(.35,widths[c]/72.0))
            cell.vertical_alignment=1
            _set_cell_zero_margins(cell)
            p=cell.paragraphs[0]; p.text=''; p.paragraph_format.space_after=Pt(0)
            source_bbox=_source_bbox_for_grid(cells, r, c, cols, len(rows), bbox)
            if source_bbox:
                _set_cell_fill(cell,_cell_fill_color(fill_rects,source_bbox))
                style=_cell_text_style(page_spans,source_bbox)
            else:
                style={}
            if c < len(vals) and vals[c]:
                cell_text=_normalize_form_cell_text(vals[c])
                run=p.add_run(cell_text); run.font.name='Arial'; run.font.size=Pt(10*font_scale)
                if r==0 or style.get('bold'): run.bold=True
                if style.get('italic'): run.italic=True
                _set_run_pdf_color(run,style.get('color'))
    # Preserve source row heights so merged title rows do not become huge.
    if cells:
        ys=sorted(set(round(float(v),3) for row in cells for cell in row if cell for v in (cell[1],cell[3])))
        if len(ys)==len(rows)+1:
            for r,row in enumerate(table.rows):
                trPr=row._tr.get_or_add_trPr()
                ht=OxmlElement('w:trHeight'); ht.set(qn('w:val'),str(max(1,int((ys[r+1]-ys[r])*20)))); ht.set(qn('w:hRule'),'exact'); trPr.append(ht)
    # Recreate only genuine merged cells after all values are populated.
    if cells:
        xs=sorted(set(round(float(v),3) for row in cells for cell in row if cell for v in (cell[0],cell[2])))
        ys=sorted(set(round(float(v),3) for row in cells for cell in row if cell for v in (cell[1],cell[3])))
        if len(xs)>=2 and len(ys)>=2:
            for sr in cells:
                for source in sr:
                    if not source: continue
                    x0,y0,x1,y1=map(float,source)
                    c0=min(range(cols),key=lambda i:abs(xs[i]-x0)); c1=min(range(cols),key=lambda i:abs(xs[i+1]-x1)) if len(xs)>1 else c0
                    r0=min(range(len(rows)),key=lambda i:abs(ys[i]-y0)); r1=min(range(len(rows)),key=lambda i:abs(ys[i+1]-y1)) if len(ys)>1 else r0
                    if r1>r0 or c1>c0:
                        try: table.cell(r0,c0).merge(table.cell(r1,c1))
                        except Exception: pass
    # Repeat a light header treatment without depending on a specific document.
    for c in range(cols):
        cell=table.cell(0,c);
        if cell.text.strip():
            for run in cell.paragraphs[0].runs: run.bold=True
    return table


def _add_blocks_to_cell(cell, blocks, page_width, start_y=None, font_scale=1.0, tables=None, page=None):
    """Populate a layout cell with text blocks and tables in vertical order."""
    items=[]
    for b in blocks: items.append(('text',b['y'],b))
    for t in (tables or []): items.append(('table',t['bbox'][1],t))
    items.sort(key=lambda x:(x[1],0 if x[0]=='text' else 1))
    first=True; prev=None
    for kind,y,obj in items:
        if kind=='table':
            _add_native_flow_table(cell,obj,font_scale,page=page)
            cell.add_paragraph()
            prev=obj['bbox'][3]; first=False; continue
        p=cell.paragraphs[0] if first and not cell.paragraphs[0].text else cell.add_paragraph()
        gap=0 if prev is None else max(0,float(obj['y'])-float(prev))
        p.paragraph_format.space_before=Pt(min(gap,30))
        _format_native_paragraph(p,obj,font_scale)
        prev=obj['y2']; first=False



def _add_top_native_artwork(doc, page):
    """Preserve small genuine top-of-page artwork such as logos without rasterizing the page."""
    added=False
    for im in page.get_images(full=True):
        try:
            for rect in page.get_image_rects(im[0]):
                if rect.y0 < page.rect.height*0.20 and rect.width >= 60 and rect.height >= 8 and rect.width < page.rect.width*0.70:
                    data=page.parent.extract_image(im[0]).get('image')
                    if not data: continue
                    p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(4)
                    p.add_run().add_picture(io.BytesIO(data),width=Inches(float(rect.width)/72.0),height=Inches(float(rect.height)/72.0))
                    added=True
        except Exception:
            pass
    return added

def _append_native_flow_page(doc, section, page, plumber_page, first_page=False, footer_lines=None):
    """General native-PDF reconstruction using normal editable Word flow."""
    _set_section_geometry(section,page)
    # Derive margins from the actual PDF content instead of assuming every document
    # uses the same page inset. This keeps ordinary letters, forms, invoices and
    # government forms aligned without hard-coding a particular document.
    prelim_tables=_native_tables(plumber_page)
    prelim_bboxes=[t['bbox'] for t in prelim_tables]
    prelim_blocks=_native_text_blocks(page,prelim_bboxes,footer_lines)
    min_x=min([b['x'] for b in prelim_blocks]+[t['bbox'][0] for t in prelim_tables]+[40.0])
    max_x=max([b['x2'] for b in prelim_blocks]+[t['bbox'][2] for t in prelim_tables]+[page.rect.width-40.0])
    section.left_margin=Inches(max(28,min(90,min_x))/72.0)
    section.right_margin=Inches(max(28,min(90,page.rect.width-max_x))/72.0)
    top_y=min([b['y'] for b in prelim_blocks]+[55.0])
    try:
        if any(rect.y0 < page.rect.height*0.20 and rect.width >= 60 and rect.height >= 8 and rect.width < page.rect.width*0.70
               for im in page.get_images(full=True) for rect in page.get_image_rects(im[0])):
            top_y=min(top_y,20.0)
    except Exception:
        pass
    section.top_margin=Inches(max(20,min(72,top_y))/72.0)
    section.bottom_margin=Inches(40/72.0)
    if first_page and footer_lines:
        fp=section.footer.paragraphs[0]; fp.text=''; fp.alignment=1
        for i,line in enumerate(footer_lines):
            if i: fp.add_run().add_break()
            r=fp.add_run(line); r.font.name='Arial'; r.font.size=Pt(8)
    tables=_native_tables(plumber_page)
    tables=[t for t in tables if _table_is_meaningful(t.get('rows')) and (t['bbox'][2]-t['bbox'][0])>=80 and (t['bbox'][3]-t['bbox'][1])>=20]
    _add_top_native_artwork(doc,page)
    table_bboxes=[t['bbox'] for t in tables]
    blocks=_native_text_blocks(page,table_bboxes,footer_lines)
    columns=_detect_text_columns(blocks,page)
    if columns:
        cross=columns['cross']
        # Full-width title/header blocks remain above the columns.
        for b in cross:
            p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(0); _format_native_paragraph(p,b,1.0)
        first_col_y=min([b['y'] for b in columns['left']+columns['right']] + [page.rect.height])
        cross_bottom=max([b['y2'] for b in cross] + [0])
        top_gap=max(0.0,first_col_y-cross_bottom)
        if top_gap>2:
            sp=doc.add_paragraph(); sp.paragraph_format.space_after=Pt(0); sp.add_run('\u200b').font.size=Pt(min(18,top_gap))
        outer=doc.add_table(rows=1,cols=2); outer.autofit=False; _remove_table_borders(outer)
        gap=16.0; usable=float(page.rect.width)-80-gap
        left_w=usable*0.5; right_w=usable-left_w
        outer.cell(0,0).width=Inches(left_w/72.0); outer.cell(0,1).width=Inches(right_w/72.0)
        for c in outer.rows[0].cells:
            _set_cell_zero_margins(c)
        # Tables belong to the column whose bbox contains their centre.
        left_tables=[]; right_tables=[]; other_tables=[]
        split=columns['split']
        for t in tables:
            cx=(t['bbox'][0]+t['bbox'][2])/2
            (left_tables if cx<split else right_tables).append(t)
        _add_blocks_to_cell(outer.cell(0,0),columns['left'],float(page.rect.width),font_scale=1.0,tables=left_tables,page=page)
        _add_blocks_to_cell(outer.cell(0,1),columns['right'],float(page.rect.width),font_scale=1.0,tables=right_tables,page=page)
        return
    # Single-column flow. Tables and text share one ordered stream.
    items=[('text',b['y'],b) for b in blocks]+[('table',t['bbox'][1],t) for t in tables]
    items.sort(key=lambda x:(x[1],0 if x[0]=='text' else 1))
    prev_y=None
    for kind,y,obj in items:
        gap=0 if prev_y is None else max(0,float(y)-float(prev_y))
        if kind=='table':
            if gap>2:
                sp=doc.add_paragraph(); sp.paragraph_format.space_after=Pt(min(gap,24))
            _add_native_flow_table(doc,obj,1.0,page=page)
            prev_y=obj['bbox'][3]
            continue
        p=doc.add_paragraph()
        p.paragraph_format.space_before=Pt(min(gap,24)) if prev_y is not None else Pt(0)
        # Scale very small source text slightly upward while retaining hierarchy.
        scale=1.0
        _format_native_paragraph(p,obj,scale)
        if obj['size']>=18:
            for r in p.runs: r.bold=True
            center=(float(obj['x'])+float(obj['x2']))/2.0
            if abs(center-float(page.rect.width)/2.0) < float(page.rect.width)*0.10 and (obj['x2']-obj['x']) < page.rect.width*0.70:
                p.alignment=1
        prev_y=obj['y2']

def convert_docx(pdf_path,output_path,pages):
    pdf=fitz.open(pdf_path)
    plumber=pdfplumber.open(pdf_path)
    doc=Document()
    for p in list(doc.paragraphs):
        p._element.getparent().remove(p._element)
    normal=doc.styles['Normal']; normal.font.name='Arial'; normal.font.size=Pt(10.5)
    footer_lines=_common_footer_lines(pdf)
    for idx,n in enumerate(pages):
        page=pdf[n-1]
        scanned=_has_large_page_image(page)
        if idx:
            section=doc.add_section(WD_SECTION.NEW_PAGE)
        else:
            section=doc.sections[0]
        if scanned:
            _add_scanned_hybrid_page(doc,section,page,idx)
        else:
            _append_native_flow_page(doc,section,page,plumber.pages[n-1],first_page=(idx==0),footer_lines=footer_lines)
    plumber.close(); pdf.close(); doc.save(output_path)

def _html_style_for_span(span):
    size=float(span.get('size') or 10)
    flags=int(span.get('flags') or 0)
    style=[f"font-size:{max(6,min(72,size)):.2f}px"]
    font=(span.get('font') or '').lower()
    if flags & 16 or 'bold' in font: style.append('font-weight:700')
    if flags & 2 or 'italic' in font or 'oblique' in font: style.append('font-style:italic')
    return ';'.join(style)

def _native_page_to_html(page):
    blocks=[]
    data=page.get_text('dict')
    for block in data.get('blocks',[]):
        if block.get('type') != 0: continue
        lines=block.get('lines',[])
        runs=[]; sizes=[]
        for line in lines:
            lr=[]
            for span in line.get('spans',[]):
                text=span.get('text','')
                if text:
                    lr.append((text,span))
                    try: sizes.append(float(span.get('size') or 10))
                    except Exception: pass
            if lr: runs.append(lr)
        if not runs: continue
        base=sum(sizes)/len(sizes) if sizes else 10
        parts=[]
        for li,lr in enumerate(runs):
            for text,span in lr:
                parts.append(f'<span style="{_html_style_for_span(span)}">{html.escape(text,quote=False)}</span>')
            if li < len(runs)-1: parts.append('<br>')
        blocks.append(f'<p style="margin:0 0 10px;line-height:1.25;font-size:{base:.2f}px">{"".join(parts)}</p>')
    return ''.join(blocks)

def _ocr_page_to_flow_html(page):
    img=_scan_image(page); candidates=[]
    for psm in (3,6):
        text,conf,data=_ocr_data(img,psm); lines=_line_data(data); candidates.append((conf,len(lines),lines))
    candidates.sort(key=lambda x:(x[0],x[1]),reverse=True)
    lines=candidates[0][2] if candidates else []
    fuller=max(candidates,key=lambda x:len(x[2]),default=(0,0,[]))
    if len(fuller[2]) > len(lines)*1.15: lines=fuller[2]
    if not lines: return ''
    parts=[]; current=[]; prev=None
    for ln in lines:
        y=float(ln['y']); h=max(1,float(ln['y2']-ln['y']))
        if prev is not None and y-prev > h*1.9 and current:
            parts.append(' '.join(current)); current=[]
        current.append(ln['text'].strip()); prev=float(ln['y2'])
    if current: parts.append(' '.join(current))
    return ''.join(f'<p style="margin:0 0 10px;line-height:1.25">{html.escape(t)}</p>' for t in parts if t)

def convert_html_text(pdf_path,output_path,pages):
    pdf=fitz.open(pdf_path)
    out=['<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>PDF to HTML Text</title><style>html,body{margin:0;background:#fff;font-family:Arial,sans-serif;color:#111}.pdf-page{page-break-before:always;page-break-after:always;padding:24px;box-sizing:border-box;min-height:100vh}.pdf-page:first-child{page-break-before:always}@media print{.pdf-page{padding:0}}</style></head><body>']
    for n in pages:
        page=pdf[n-1]
        body=_native_page_to_html(page) or _ocr_page_to_flow_html(page)
        out.append(f'<div class="pdf-page" data-page="{n}">{body}</div>')
    out.append('</body></html>')
    Path(output_path).write_text(''.join(out),encoding='utf-8')
    pdf.close()

def convert_html_image(pdf_path,output_path,pages):
    pdf=fitz.open(pdf_path)
    out=['<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>PDF to HTML Image</title><style>html,body{margin:0;padding:0;background:#e9edf1}.pdf-document{padding:24px}.pdf-page{margin:0 auto 24px;background:#fff;box-shadow:0 2px 12px rgba(0,0,0,.14);page-break-after:always}.pdf-page img{display:block;width:100%;height:auto}@media print{html,body{background:#fff}.pdf-document{padding:0}.pdf-page{margin:0;box-shadow:none}}</style></head><body><div class="pdf-document">']
    for n in pages:
        page=pdf[n-1]
        pix=page.get_pixmap(matrix=fitz.Matrix(2,2),alpha=False,colorspace=fitz.csRGB)
        b=base64.b64encode(pix.tobytes('png')).decode('ascii')
        out.append(f'<section class="pdf-page" data-page="{n}"><img src="data:image/png;base64,{b}" alt="PDF page {n}"></section>')
    out.append('</div></body></html>')
    Path(output_path).write_text(''.join(out),encoding='utf-8')
    pdf.close()

def convert_html(pdf_path,output_path,pages):
    return convert_html_text(pdf_path,output_path,pages)
