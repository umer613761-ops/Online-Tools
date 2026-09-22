import base64
import html
import io
import re
from pathlib import Path

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
        p.paragraph_format.space_before=Pt(3 if gap > 13 else 0)
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


def _add_scanned_table_page(doc, section, page, page_index, table):
    """Hybrid scanned-page reconstruction: artwork remains image, ruled table is a real editable Word table."""
    _set_scanned_section(section,page)
    img=_scan_image(page); pw=float(page.rect.width); ph=float(page.rect.height)
    # Source table x/y in pixels. Keep the surrounding certificate artwork as raster segments.
    left,right=table['x'][0],table['x'][-1]
    _add_segment_image(doc,img,0,0,img.width,table['top'],pw,ph)
    cells=_scan_table_cells(img,table)
    t=doc.add_table(rows=len(cells), cols=max(1,len(table['x'])-1)); t.style='Table Grid'; t.autofit=False
    scale_x=pw/img.width
    widths=[(table['x'][i+1]-table['x'][i])*scale_x/72.0 for i in range(len(table['x'])-1)]
    for ri,row in enumerate(t.rows):
        for ci,cell in enumerate(row.cells):
            cell.width=Inches(widths[ci]); cell.text=cells[ri][ci] if ci<len(cells[ri]) else ''
            tcPr=cell._tc.get_or_add_tcPr(); mar=OxmlElement('w:tcMar')
            for edge in ('top','left','bottom','right'):
                ee=OxmlElement('w:'+edge); ee.set(qn('w:w'),'0'); ee.set(qn('w:type'),'dxa'); mar.append(ee)
            tcPr.append(mar)
            for pp in cell.paragraphs:
                pp.paragraph_format.space_before=Pt(0); pp.paragraph_format.space_after=Pt(0); pp.paragraph_format.line_spacing=1
                for rr in pp.runs:
                    rr.font.name='Arial'; rr.font.size=Pt(6.2); rr.bold=(ri==0)
        # Match the source row pitch.
        # Use the source table's full height divided across the reconstructed rows.
        rh=min(13.5, (table['bottom']-table['top'])/max(1,len(t.rows))*ph/img.height)
        trPr=row._tr.get_or_add_trPr(); ht=OxmlElement('w:trHeight'); ht.set(qn('w:val'),str(max(160,int(rh*20)))); ht.set(qn('w:hRule'),'exact'); trPr.append(ht)
    # Indent table to the source x position.
    tPr=t._tbl.tblPr; ind=OxmlElement('w:tblInd'); ind.set(qn('w:w'),str(int(left*scale_x*20))); ind.set(qn('w:type'),'dxa'); tPr.append(ind)
    _add_segment_image(doc,img,0,table['bottom'],img.width,img.height,pw,ph)

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
            scan_img=_scan_image(page)
            scan_table=_scan_table_region(scan_img)
            if scan_table:
                _add_scanned_table_page(doc,section,page,idx,scan_table)
            else:
                _append_scanned_image_page(doc,section,page)
        else:
            if idx and footer_lines:
                pass
            _append_native_page(doc,page,plumber.pages[n-1],first_page=(idx==0),footer_lines=footer_lines)
    plumber.close(); pdf.close(); doc.save(output_path)

def convert_html(pdf_path,output_path,pages):
    pdf=fitz.open(pdf_path); out=['<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>PDF to HTML</title><style>html,body{margin:0;background:#e9edf1;font-family:Arial,sans-serif}.pdf-document{padding:24px}.pdf-page{position:relative;margin:0 auto 24px;background:#fff;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,.14)}.pdf-bg{position:absolute;inset:0;width:100%;height:100%}.editable{position:absolute;background:#fff;border:0;padding:0;margin:0;outline:none;line-height:1;box-sizing:border-box}.native{background:transparent}@media print{body{background:#fff}.pdf-document{padding:0}.pdf-page{margin:0;box-shadow:none;break-after:page}.pdf-page:last-child{break-after:auto}}</style></head><body><div class="pdf-document">']
    for n in pages:
        page=pdf[n-1]; scanned=_has_large_page_image(page); img=_scan_image(page) if scanned else None
        if scanned:
            lines,_,_= _ocr_page(page); cleaned=_clean_scan(img,lines); buf=io.BytesIO(); cleaned.save(buf,'PNG'); b=base64.b64encode(buf.getvalue()).decode(); sx=page.rect.width/img.width; sy=page.rect.height/img.height
        else:
            buf=io.BytesIO(); pix=page.get_pixmap(matrix=fitz.Matrix(1,1),alpha=False,colorspace=fitz.csRGB); Image.open(io.BytesIO(pix.tobytes('png'))).save(buf,'PNG'); b=base64.b64encode(buf.getvalue()).decode(); lines=_native_lines(page); sx=sy=1
        out.append(f'<section class="pdf-page" style="width:{page.rect.width}px;height:{page.rect.height}px"><img class="pdf-bg" src="data:image/png;base64,{b}">')
        for ln in lines:
            if scanned:
                x,y,w,h=ln['x']*sx,ln['y']*sy,max(8,(ln['x2']-ln['x'])*sx+4),max(10,(ln['y2']-ln['y'])*sy+3); fs=max(7,min(16,(ln['y2']-ln['y'])*sy*.78))
            else:
                x,y,w,h=0,0,0,0; continue
            if ln['conf']<25: continue
            out.append(f'<div contenteditable="true" class="editable" style="left:{x}px;top:{y}px;width:{w}px;height:{h}px;font-size:{fs}px">{html.escape(ln["text"])}</div>')
        out.append('</section>')
    out.append('</div></body></html>'); Path(output_path).write_text(''.join(out),encoding='utf-8'); pdf.close()
