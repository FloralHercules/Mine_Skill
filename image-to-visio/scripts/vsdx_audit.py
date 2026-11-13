"""Read native shapes and per-page MathType embedding relationships without Office."""
import hashlib
import posixpath
import zipfile
import xml.etree.ElementTree as E
from pathlib import Path

NS = {'v':'http://schemas.microsoft.com/office/visio/2012/main',
      'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}

def sha(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda:f.read(1024*1024), b''):
            h.update(chunk)
    return h.hexdigest()

def rels(z, part):
    rp = posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part)+'.rels')
    if rp not in z.namelist():
        return {}
    return {r.get('Id'):posixpath.normpath(posixpath.join(posixpath.dirname(part),r.get('Target')))
            for r in E.fromstring(z.read(rp)) if r.get('TargetMode') != 'External'}

def audit(path):
    with zipfile.ZipFile(path) as z:
        assert z.testzip() is None, 'Corrupt ZIP member'
        index = E.fromstring(z.read('visio/pages/pages.xml'))
        links = rels(z, 'visio/pages/pages.xml')
        pages = []
        for pg in index:
            rel = pg.find('v:Rel',NS)
            part = links[rel.get('{'+NS['r']+'}id')]
            root = E.fromstring(z.read(part))
            pr = rels(z,part)
            shapes = root.findall('.//v:Shape',NS)
            math = []
            for s in shapes:
                foreign = s.find('v:ForeignData',NS)
                if foreign is None:
                    continue
                r = foreign.find('v:Rel',NS)
                target = pr.get(r.get('{'+NS['r']+'}id')) if r is not None else None
                if target and b'Equation.DSMT4' in z.read(target):
                    math.append(s.get('NameU') or s.get('Name'))
            texts = [''.join(t.itertext()) for t in root.findall('.//v:Text',NS)]
            # Compare content after round-tripping; XML bytes legitimately change on save.
            pages.append({'id':pg.get('ID'),'name':pg.get('Name'),'name_u':pg.get('NameU'),
                          'part':part,'shape_count':len(shapes),'math_count':len(math),
                          'math_shapes':math,'texts':texts,
                          'shape_names':[s.get('NameU') or s.get('Name') for s in shapes]})
        doc = E.fromstring(z.read('visio/document.xml'))
        return {'path':str(Path(path).resolve()),'pages':pages,
                'fonts':[f.get('NameU') for f in doc.findall('.//v:FaceName',NS)],
                'math_count':sum(p['math_count'] for p in pages)}

def same_content(before, after):
    return (before['shape_count']==after['shape_count'] and
            before['math_count']==after['math_count'] and
            sorted(before['texts'])==sorted(after['texts']))
