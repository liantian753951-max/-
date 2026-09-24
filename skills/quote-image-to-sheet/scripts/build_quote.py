"""Deterministic quote builder. No credentials/network; only Pillow + stdlib.

Input JSON: project, supplier, items:[{name,box:[left,top,right,bottom]}].
Boxes refer to EXIF-corrected original pixels. Outputs embedded-image XLSX,
contact sheet and a machine-readable verification report in one invocation.
"""
import argparse
import copy
import hashlib
import json
import re
import time
import zipfile
from datetime import date
from pathlib import Path
from xml.etree import ElementTree as E
from PIL import Image, ImageOps, ImageDraw

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/package/2006/relationships'
D = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
C = 'http://schemas.openxmlformats.org/package/2006/content-types'
E.register_namespace('', S)
E.register_namespace('r', R)
def tag(ns, name): return '{'+ns+'}'+name
def xml(node):
    # OPC readers require unprefixed package relationship/content-type roots.
    E.register_namespace('',node.tag.split('}')[0][1:])
    return E.tostring(node, encoding='utf-8', xml_declaration=True)
def cell(row, col):
    ref = col+row.get('r')
    for c in row.findall(tag(S,'c')):
        if c.get('r') == ref: return c
    return E.SubElement(row, tag(S,'c'), {'r':ref})
def put(c, value=None):
    for child in list(c): c.remove(child)
    c.attrib.pop('t', None)
    if value is None: return
    if isinstance(value,(int,float)):
        E.SubElement(c,tag(S,'v')).text=str(value)
    else:
        c.set('t','inlineStr')
        E.SubElement(E.SubElement(c,tag(S,'is')),tag(S,'t')).text=str(value)
def renumber(row,n):
    row.set('r',str(n))
    for c in row.findall(tag(S,'c')): c.set('r',re.sub(r'\d+$',str(n),c.get('r')))
    return row

def formula(c, expression):
    put(c)
    c.set('t','str')
    E.SubElement(c,tag(S,'f')).text=expression
    E.SubElement(c,tag(S,'v')).text=''

def date_style(parts):
    styles=E.fromstring(parts['xl/styles.xml'])
    formats=styles.find(tag(S,'numFmts'))
    if formats is None:
        formats=E.Element(tag(S,'numFmts')); styles.insert(0,formats)
    fmt=max([163]+[int(x.get('numFmtId')) for x in formats])+1
    E.SubElement(formats,tag(S,'numFmt'),{'numFmtId':str(fmt),'formatCode':'yyyy-mm-dd'})
    formats.set('count',str(len(formats)))
    xfs=styles.find(tag(S,'cellXfs'))
    sheet=E.fromstring(parts['xl/worksheets/sheet1.xml'])
    original=sheet.find(".//"+tag(S,'c')+"[@r='I5']")
    xf=copy.deepcopy(xfs[int(original.get('s','0'))]); xf.set('numFmtId',str(fmt)); xf.set('applyNumberFormat','1')
    index=len(xfs); xfs.append(xf); xfs.set('count',str(len(xfs)))
    parts['xl/styles.xml']=xml(styles)
    return str(index)

def clean_template(source, destination):
    """Keep format and daily rate, remove previous project data and all images."""
    with zipfile.ZipFile(source) as z: parts={n:z.read(n) for n in z.namelist()}
    root=E.fromstring(parts['xl/worksheets/sheet1.xml'])
    strings=E.fromstring(parts['xl/sharedStrings.xml']) if 'xl/sharedStrings.xml' in parts else []
    texts=[''.join(x.itertext()) for x in strings]
    for c in root.iter(tag(S,'c')):
        if c.get('t')=='s': put(c,texts[int(c.find(tag(S,'v')).text)])
    rows={int(r.get('r')):r for r in root.find(tag(S,'sheetData'))}
    for n in range(5,12):
        for c in list(rows[n]):
            if not c.get('r','').startswith('F'): put(c)
    put(cell(rows[2],'C'))
    put(cell(rows[3],'C'))
    put(cell(rows[12],'H'))
    for node in list(root):
        if node.tag in (tag(S,'drawing'),tag(S,'legacyDrawing')): root.remove(node)
    for name in list(parts):
        if name.startswith(('xl/media/','xl/drawings/')) or name in ('xl/sharedStrings.xml','xl/worksheets/_rels/sheet1.xml.rels'):
            del parts[name]
    for name in ['[Content_Types].xml','xl/_rels/workbook.xml.rels']:
        r=E.fromstring(parts[name])
        for node in list(r):
            if any(v in str(node.attrib) for v in ('sharedStrings','drawing','/media/')): r.remove(node)
        parts[name]=xml(r)
    parts['xl/worksheets/sheet1.xml']=xml(root)
    Path(destination).parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(destination,'w',zipfile.ZIP_DEFLATED) as z:
        for n,b in parts.items(): z.writestr(n,b)

def build(source, manifest, output, template):
    start=time.perf_counter()
    spec=json.loads(Path(manifest).read_text(encoding='utf-8-sig'))
    generated=date.today()
    deadline=date.fromisoformat(spec['deadline']) if spec.get('deadline') else None
    items=spec['items']; project=spec['project'].strip(); supplier=spec['supplier'].strip()
    if not project or not supplier or not 1<=len(items)<=100: raise ValueError('Missing project/supplier or invalid item count')
    if spec.get('expected_count',len(items))!=len(items): raise ValueError('Item count mismatch')
    im=ImageOps.exif_transpose(Image.open(source)).convert('RGB')
    crops=[]
    for item in items:
        if not item['name'].strip(): raise ValueError('Empty item name')
        x1,y1,x2,y2=item['box']
        if not all(isinstance(v,int) for v in item['box']) or not (0<=x1<x2<=im.width and 0<=y1<y2<=im.height):
            raise ValueError('Invalid crop box: '+str(item['box']))
        crops.append(im.crop((x1,y1,x2,y2)))
    out=Path(output); out.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(template) as z: parts={n:z.read(n) for n in z.namelist()}
    ds=date_style(parts)
    book=E.fromstring(parts['xl/workbook.xml'])
    calc=book.find(tag(S,'calcPr'))
    if calc is None: calc=E.SubElement(book,tag(S,'calcPr'))
    calc.set('calcMode','auto'); calc.set('fullCalcOnLoad','1')
    parts['xl/workbook.xml']=xml(book)
    root=E.fromstring(parts['xl/worksheets/sheet1.xml']); data=root.find(tag(S,'sheetData'))
    old={int(r.get('r')):r for r in data}
    if root.find(tag(S,'drawing')) is not None: raise ValueError('Use a sanitized template')
    exemplar=copy.deepcopy(old[5]); total=copy.deepcopy(old[12])
    for r in list(data):
        if int(r.get('r'))>=5: data.remove(r)
    put(cell(old[2],'C'),project+'报价表'); put(cell(old[3],'C'),'供应商名称：'+supplier)
    drawing=E.Element(tag(D,'wsDr')); rels=E.Element(tag(P,'Relationships'))
    contact=Image.new('RGB',(900,340*((len(items)+2)//3)),'white'); pen=ImageDraw.Draw(contact)
    for i,(item,crop) in enumerate(zip(items,crops)):
        row=renumber(copy.deepcopy(exemplar),i+5)
        for c in row:
            if not c.get('r','').startswith('F'): put(c)
        put(cell(row,'B'),i+1); put(cell(row,'E'),item['name']); data.append(row)
        n=i+5
        formula(cell(row,'H'),f'IF(COUNT(F{n}:G{n})<2,"",F{n}*G{n})')
        put(cell(row,'I'),(generated-date(1899,12,30)).days); cell(row,'I').set('s',ds)
        cell(row,'J').set('s',ds)
        if deadline: put(cell(row,'J'),(deadline-date(1899,12,30)).days)
        name=f'item-{i+1:02}.png'; crop.save(out/name)
        parts[f'xl/media/{name}']=(out/name).read_bytes()
        thumb=ImageOps.contain(crop,(280,300)); x=(i%3)*300; y=(i//3)*340
        contact.paste(thumb,(x+(300-thumb.width)//2,y+25)); pen.text((x+12,y+7),str(i+1),fill='black')
        height=float(row.get('ht','264.75'))*4/3; width=495
        factor=min((width-20)/crop.width,(height-20)/crop.height)
        w,h=round(crop.width*factor),round(crop.height*factor)
        anchor=E.SubElement(drawing,tag(D,'oneCellAnchor'))
        fr=E.SubElement(anchor,tag(D,'from'))
        for key,val in [('col',3),('colOff',round((width-w)/2*9525)),('row',i+4),('rowOff',round((height-h)/2*9525))]:
            E.SubElement(fr,tag(D,key)).text=str(val)
        E.SubElement(anchor,tag(D,'ext'),{'cx':str(w*9525),'cy':str(h*9525)})
        pic=E.SubElement(anchor,tag(D,'pic')); nv=E.SubElement(pic,tag(D,'nvPicPr'))
        E.SubElement(nv,tag(D,'cNvPr'),{'id':str(i+1),'name':item['name']})
        E.SubElement(nv,tag(D,'cNvPicPr'))
        fill=E.SubElement(pic,tag(D,'blipFill'))
        E.SubElement(fill,tag(A,'blip'),{tag(R,'embed'):f'rId{i+1}'})
        E.SubElement(E.SubElement(fill,tag(A,'stretch')),tag(A,'fillRect'))
        sp=E.SubElement(pic,tag(D,'spPr')); geo=E.SubElement(sp,tag(A,'prstGeom'),{'prst':'rect'}); E.SubElement(geo,tag(A,'avLst'))
        E.SubElement(anchor,tag(D,'clientData'))
        E.SubElement(rels,tag(P,'Relationship'),{'Id':f'rId{i+1}','Type':R+'/image','Target':'../media/'+name})
    last=len(items)+5; renumber(total,last)
    for c in total: put(c)
    put(cell(total,'G'),'总价'); data.append(total)
    formula(cell(total,'H'),f'IF(COUNT(G5:G{last-1})=0,"",SUM(H5:H{last-1}))')
    dim=root.find(tag(S,'dimension'))
    if dim is not None: dim.set('ref',f'A1:L{last}')
    E.SubElement(root,tag(S,'drawing'),{tag(R,'id'):'rIdQuote'})
    sr=E.Element(tag(P,'Relationships')); E.SubElement(sr,tag(P,'Relationship'),{'Id':'rIdQuote','Type':R+'/drawing','Target':'../drawings/quote.xml'})
    parts['xl/worksheets/_rels/sheet1.xml.rels']=xml(sr)
    parts['xl/drawings/quote.xml']=xml(drawing); parts['xl/drawings/_rels/quote.xml.rels']=xml(rels)
    parts['xl/worksheets/sheet1.xml']=xml(root)
    ct=E.fromstring(parts['[Content_Types].xml'])
    if not any(n.get('Extension')=='png' for n in ct): E.SubElement(ct,tag(C,'Default'),{'Extension':'png','ContentType':'image/png'})
    E.SubElement(ct,tag(C,'Override'),{'PartName':'/xl/drawings/quote.xml','ContentType':'application/vnd.openxmlformats-officedocument.drawing+xml'})
    parts['[Content_Types].xml']=xml(ct)
    result=out/'quote.xlsx'; temp=out/'quote.tmp'
    with zipfile.ZipFile(temp,'w',zipfile.ZIP_DEFLATED) as z:
        for n,b in parts.items(): z.writestr(n,b)
    temp.replace(result); contact.save(out/'contact-sheet.png')
    # Inspect serialized bytes, not in-memory values: catches stale formula export bugs.
    with zipfile.ZipFile(result) as z:
        saved=E.fromstring(z.read('xl/worksheets/sheet1.xml'))
        assert len([n for n in z.namelist() if n.startswith('xl/media/')])==len(items)
        assert len(list(saved.iter(tag(S,'f'))))==len(items)+1, 'Missing calculation formula'
        for c in saved.iter(tag(S,'c')):
            m=re.fullmatch(r'([A-Z]+)(\d+)',c.get('r','')); col,n=m[1],int(m[2])
            if (5<=n<last and col in ('G','K','L')) or (5<=n<last and col=='J' and deadline is None):
                assert len(c)==0, 'Expected blank: '+c.get('r')
    report={'title':project+'报价表','count':len(items),'embedded_images':len(items),'verified':True,'days_blank':True,
            'start_date':generated.isoformat(),'deadline':deadline.isoformat() if deadline else None,
            'xlsx':str(result.resolve()),'contact_sheet':str((out/'contact-sheet.png').resolve()),
            'sha256':hashlib.sha256(result.read_bytes()).hexdigest(),'elapsed_seconds':round(time.perf_counter()-start,3)}
    (out/'result.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    return report

if __name__=='__main__':
    p=argparse.ArgumentParser(); sub=p.add_subparsers(dest='mode',required=True)
    clean=sub.add_parser('cache-template'); clean.add_argument('source'); clean.add_argument('destination')
    run=sub.add_parser('build'); run.add_argument('source'); run.add_argument('manifest'); run.add_argument('output')
    run.add_argument('--template',default=str(Path(__file__).resolve().parents[1]/'assets/template.xlsx'))
    a=p.parse_args()
    if a.mode=='cache-template': clean_template(a.source,a.destination)
    else: print(json.dumps(build(a.source,a.manifest,a.output,a.template),ensure_ascii=False))
