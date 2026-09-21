import base64
import html
import io
import re
from pathlib import Path

import fitz
from PIL import Image, ImageOps
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

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


def convert_docx(pdf_path,output_path,pages):
    pdf=fitz.open(pdf_path); doc=Document()
    for i,n in enumerate(pages):
        sec=doc.sections[0] if i==0 else doc.add_section(1)
        page=pdf[n-1]; pw=page.rect.width; ph=page.rect.height
        sec.page_width=Inches(pw/72); sec.page_height=Inches(ph/72)
        sec.top_margin=sec.bottom_margin=sec.left_margin=sec.right_margin=Inches(0); sec.header_distance=sec.footer_distance=Inches(0); sec.header.is_linked_to_previous=False
        header=sec.header; p=header.paragraphs[0]; p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0)
        scanned=_has_large_page_image(page)
        if scanned:
            img=_scan_image(page); lines,_,_= _ocr_page(page); table=_detect_table(page,img)
            cleaned=_clean_scan(img,lines,table)
            tmp=Path(output_path).with_name(f'.toolnest_{i+1}.png'); cleaned.save(tmp)
            rid,_=doc.part.get_or_add_image(str(tmp)); _vml(p,'image',0,0,pw,ph,rid=rid)
            sx=pw/img.width; sy=ph/img.height
            # Replace table text with targeted cell OCR when a ruled table is found.
            table_boxes=[]
            if table:
                cells=_ocr_table_words(img,table)
                for r,row in enumerate(cells):
                    for c,text in enumerate(row):
                        if not text: continue
                        x=table['x'][c]*sx; y=table['y'][r]*sy; w=(table['x'][c+1]-table['x'][c])*sx; h=(table['y'][r+1]-table['y'][r])*sy
                        # table text is centered/left depending on column; a white box hides original printed text.
                        fs=max(6,min(13,(table['y'][r+1]-table['y'][r])*.42))
                        _vml(p,'text',x+2,y+2,max(8,w-4),max(10,h-4),text=text,font_size=fs)
                        table_boxes.append((table['x'][c],table['y'][r],table['x'][c+1],table['y'][r+1]))
            for ln in lines:
                if not ln['text'] or ln['conf']<25: continue
                # Skip OCR lines inside detected table; cells above handle them.
                if table and table['left']<=ln['x']<=table['right'] and table['y'][0]<=ln['y']<=table['y'][-1]:
                    continue
                x,y,x2,y2=ln['x']*sx,ln['y']*sy,ln['x2']*sx,ln['y2']*sy
                import numpy as np
                arr=np.array(img); xa=max(0,int(ln['x']-3)); xb=min(arr.shape[1],int(ln['x2']+3)); ya=max(0,int(ln['y']-3)); yb=min(arr.shape[0],int(ln['y2']+3)); sample=arr[ya:yb,xa:xb].mean() if xb>xa and yb>ya else 255
                if sample<145: continue
                fs=max(6,min(15,(y2-y)*.78)); _vml(p,'text',x,y,max(8,x2-x+4),max(10,y2-y+3),text=ln['text'],font_size=fs)
            tmp.unlink(missing_ok=True)
        else:
            # Native text PDFs: keep the existing editable text path.
            for ln in _native_lines(page):
                para=doc.add_paragraph(); para.paragraph_format.space_after=Pt(4); para.add_run(ln['text'])
        doc.add_paragraph('')
    pdf.close(); doc.save(output_path)


def _native_lines(page):
    data=page.get_text('dict'); out=[]
    for block in data.get('blocks',[]):
        if block.get('type')!=0: continue
        for line in block.get('lines',[]):
            text=''.join(s.get('text','') for s in line.get('spans',[])).strip()
            if text: out.append({'text':text})
    return out


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
