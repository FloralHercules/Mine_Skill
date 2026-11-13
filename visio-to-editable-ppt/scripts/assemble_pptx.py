"""Assemble native diagrams, separate images, and unchanged OLE binaries.

Uses a blank PowerPoint-created template. No metafile ungrouping or raster
substitution of editable geometry is performed.
"""
from pathlib import Path
import argparse, copy, hashlib, json, math, re, zipfile
from lxml import etree as E

P='http://schemas.openxmlformats.org/presentationml/2006/main'
A='http://schemas.openxmlformats.org/drawingml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
REL='http://schemas.openxmlformats.org/package/2006/relationships'
CT='http://schemas.openxmlformats.org/package/2006/content-types'
NS={'p':P,'a':A,'r':R}
EMU=12700
def q(ns,name):return '{'+ns+'}'+name
def sub(el,ns,tag_name,**kw):return E.SubElement(el,q(ns,tag_name),**{k:str(v) for k,v in kw.items()})
def xml(el):return E.tostring(el,encoding='UTF-8',xml_declaration=True,standalone=True)
def get_map(path):return {int(o.get('source_id')):[copy.deepcopy(x) for x in o] for o in E.parse(str(path)).getroot()}
def xfrm(el):
    if el.tag==q(P,'grpSp'):return el.find('p:grpSpPr/a:xfrm',NS)
    if el.tag==q(P,'graphicFrame'):return el.find('p:xfrm',NS)
    return el.find('p:spPr/a:xfrm',NS)
def bounds(el):
    xf=xfrm(el);o=xf.find('a:off',NS);e=xf.find('a:ext',NS)
    x,y,cx,cy=map(int,[o.get('x'),o.get('y'),e.get('cx'),e.get('cy')])
    t=math.radians(float(xf.get('rot','0'))/60000);co,si=abs(math.cos(t)),abs(math.sin(t));w=cx*co+cy*si;h=cx*si+cy*co
    return x+cx/2-w/2,y+cy/2-h/2,w,h
def combine(elements,sid):
    if not elements:raise ValueError(f'No native content for source shape {sid}; inspect unsupported geometry')
    if len(elements)==1:
        el=elements[0];el.find('.//p:cNvPr',NS).set('name',f'Visio_{sid}');return el
    boxes=[bounds(x) for x in elements];x=min(b[0] for b in boxes);y=min(b[1] for b in boxes)
    w=max(b[0]+b[2] for b in boxes)-x;h=max(b[1]+b[3] for b in boxes)-y
    g=E.Element(q(P,'grpSp'));nv=sub(g,P,'nvGrpSpPr');sub(nv,P,'cNvPr',id=1,name=f'Visio_{sid}');sub(nv,P,'cNvGrpSpPr');sub(nv,P,'nvPr')
    xf=sub(sub(g,P,'grpSpPr'),A,'xfrm')
    for tag,attrs in [('off',{'x':round(x),'y':round(y)}),('ext',{'cx':max(1,round(w)),'cy':max(1,round(h))}),('chOff',{'x':round(x),'y':round(y)}),('chExt',{'cx':max(1,round(w)),'cy':max(1,round(h))})]:sub(xf,A,tag,**attrs)
    g.extend(elements);return g
def picture(obj,rid):
    p=obj['placement'];sp=E.Element(q(P,'pic'));nv=sub(sp,P,'nvPicPr');sub(nv,P,'cNvPr',id=1,name=f"Visio_{obj['source_id']}");sub(nv,P,'cNvPicPr');sub(nv,P,'nvPr')
    fill=sub(sp,P,'blipFill');b=sub(fill,A,'blip');b.set(q(R,'embed'),rid);sub(sub(fill,A,'stretch'),A,'fillRect')
    pr=sub(sp,P,'spPr');xf=sub(pr,A,'xfrm');sub(xf,A,'off',x=round(p['x']*EMU),y=round(p['y']*EMU));sub(xf,A,'ext',cx=max(1,round(p['w']*EMU)),cy=max(1,round(p['h']*EMU)))
    sub(sub(pr,A,'prstGeom',prst='rect'),A,'avLst');sub(sub(pr,A,'ln'),A,'noFill')
    return sp
def ole_frame(pic,rid,prog_id):
    if not prog_id:raise ValueError('OLE program identifier is unknown; inspect the original object before embedding')
    frame=E.Element(q(P,'graphicFrame'));nv=sub(frame,P,'nvGraphicFramePr');nv.append(copy.deepcopy(pic.find('p:nvPicPr/p:cNvPr',NS)));sub(nv,P,'cNvGraphicFramePr');sub(nv,P,'nvPr')
    xf=copy.deepcopy(xfrm(pic));xf.tag=q(P,'xfrm');frame.append(xf)
    data=sub(sub(frame,A,'graphic'),A,'graphicData',uri=P.rsplit('/',1)[0]+'/ole')
    ext=xf.find('a:ext',NS);ole=sub(data,P,'oleObj',showAsIcon=0,name=prog_id,progId=prog_id,imgW=ext.get('cx'),imgH=ext.get('cy'));ole.set(q(R,'id'),rid);sub(ole,P,'embed');ole.append(pic)
    return frame
def build(work):
    m=json.loads((work/'manifest.json').read_text(encoding='utf-8-sig'));geo=get_map(work/'geometry-map.xml');text=get_map(work/'text-map.xml')
    with zipfile.ZipFile(work/'base.pptx') as z:parts={i.filename:z.read(i.filename) for i in z.infolist()}
    slide=E.fromstring(parts['ppt/slides/slide1.xml']);tree=slide.find('p:cSld/p:spTree',NS)
    for e in list(tree):
        if e.tag not in (q(P,'nvGrpSpPr'),q(P,'grpSpPr')):tree.remove(e)
    relname='ppt/slides/_rels/slide1.xml.rels'
    rels=E.fromstring(parts[relname]) if relname in parts else E.Element(q(REL,'Relationships'),nsmap={None:REL})
    ctypes=E.fromstring(parts['[Content_Types].xml'])
    if not any(t.get('Extension')=='png' for t in ctypes):sub(ctypes,CT,'Default',Extension='png',ContentType='image/png')
    next_rid=max([int(x.get('Id')[3:]) for x in rels if re.fullmatch(r'rId\d+',x.get('Id',''))]+[0])+1
    def relationship(kind,target):
        nonlocal next_rid
        rid=f'rId{next_rid}';next_rid+=1;sub(rels,REL,'Relationship',Id=rid,Type=R+'/'+kind,Target=target);return rid
    binaries=[];label_count=0
    with zipfile.ZipFile(work/'flattened.vsdx') as source:
        for obj in m['objects']:
            sid=obj['source_id']
            if obj['raster']:
                media=f'visio-{sid}.png';parts['ppt/media/'+media]=Path(obj['asset']).read_bytes();shape=picture(obj,relationship('image','../media/'+media))
                if obj.get('ole_binary'):
                    original=obj.get('source_part') or 'visio/embeddings/'+obj['ole_binary']
                    data=source.read(original);part=f'ppt/embeddings/visio-{sid}.bin';parts[part]=data
                    sub(ctypes,CT,'Override',PartName='/'+part,ContentType='application/vnd.openxmlformats-officedocument.oleObject')
                    shape=ole_frame(shape,relationship('oleObject','../embeddings/'+Path(part).name),obj.get('prog_id'))
                    binaries.append({'source_id':sid,'sha256':hashlib.sha256(data).hexdigest(),'part':part})
            else:shape=combine(geo.get(sid,[])+text.get(sid,[]),sid)
            if obj.get('text'):
                expect=re.sub(r'\s+','',obj['text']);actual=re.sub(r'\s+','',''.join(shape.xpath('.//a:t/text()',namespaces=NS)))
                if expect!=actual:raise ValueError(f'Label content mismatch on shape {sid}')
                label_count+=1
            tree.append(shape)
    for i,e in enumerate(slide.findall('.//p:cNvPr',NS),1):e.set('id',str(i))
    parts['ppt/slides/slide1.xml']=xml(slide);parts[relname]=xml(rels);parts['[Content_Types].xml']=xml(ctypes)
    dest=work/'candidate.pptx'
    with zipfile.ZipFile(dest,'w',zipfile.ZIP_DEFLATED) as z:
        for name,data in parts.items():z.writestr(name,data)
    report={'labels':label_count,'ole_objects':len(binaries),'ole_binaries':binaries,'native_shapes':len(slide.findall('.//p:sp',NS)),'groups':len(slide.findall('.//p:grpSp',NS)),'path_commands':sum(len(slide.findall('.//a:'+op,NS)) for op in ('moveTo','lnTo','cubicBezTo','quadBezTo','arcTo')),'slide_xml_bytes':len(parts['ppt/slides/slide1.xml'])}
    (work/'assembly-report.json').write_text(json.dumps(report,indent=2),encoding='utf8');print(json.dumps({k:v for k,v in report.items() if k!='ole_binaries'}))

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--work-dir',type=Path,required=True);a=p.parse_args();build(a.work_dir.resolve())
