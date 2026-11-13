"""Native Visio renderer and isolated, resumable MathType batch worker."""
import json
import re
import time
from pathlib import Path
from runtime import office, bootstrap

KINDS = {'box','text','oval','poly','line','photo','math'}

def read_scene(path):
    path = Path(path).resolve()
    s = json.loads(path.read_text(encoding='utf-8'))
    assert s['width']>0 and s['height']>0
    ids=set()
    for o in s['elements']:
        assert o['kind'] in KINDS, o
        assert re.fullmatch(r'[A-Za-z][A-Za-z0-9_-]*',o['id']), o['id']
        assert o['id'] not in ids, 'Duplicate id '+o['id']
        ids.add(o['id'])
        if o['kind'] in {'box','text','oval','photo','math'}:
            assert o['w']>0 and o['h']>0, o['id']
        if o['kind']=='math': assert o.get('tex','').strip(), o['id']
    return s

def cell(sh,n,f): sh.CellsU(n).FormulaU=str(f)
def rgb(c): return 'RGB('+','.join(str(int(c[i:i+2],16)) for i in (1,3,5))+')'

def place(sh,o,H,scale,fit=False):
    w,h=o['w']/scale,o['h']/scale
    if fit:
        ow,oh=sh.CellsU('Width').ResultIU,sh.CellsU('Height').ResultIU
        fac=min(w/ow,h/oh);w,h=ow*fac,oh*fac
    for key,value in [('Width',w),('Height',h),('PinX',(o['x']+o['w']/2)/scale),
                      ('PinY',(H-o['y']-o['h']/2)/scale)]:cell(sh,key,value)

def image_file(o,scene,scene_path,outdir):
    bootstrap()
    from PIL import Image
    if o.get('file'):
        return (scene_path.parent/o['file']).resolve()
    src=(scene_path.parent/scene['source_image']).resolve()
    im=Image.open(src)
    # Crop coordinates are in scene coordinates, not the original image's pixel size.
    c=[round(v*(im.width/scene['width'] if i%2==0 else im.height/scene['height'])) for i,v in enumerate(o['crop'])]
    outdir.mkdir(parents=True,exist_ok=True)
    output=outdir/(o['id']+'.png');im.crop(c).save(output)
    return output

def native(scene_path,output):
    scene_path=Path(scene_path).resolve();output=Path(output).resolve()
    scene=read_scene(scene_path);H=scene['height'];scale=scene.get('units_per_inch',100)
    com,pc=office();v=com.DispatchEx('Visio.InvisibleApp');d=None
    try:
        v.AlertResponse=7;v.ScreenUpdating=False;v.EventsEnabled=False;v.DeferRecalc=True
        d=v.Documents.Add('');p=d.Pages.Item(1);p.Name=scene.get('title','Diagram')[:60]
        cell(p.PageSheet,'PageWidth',scene['width']/scale);cell(p.PageSheet,'PageHeight',H/scale)
        bg=p.DrawRectangle(0,0,scene['width']/scale,H/scale);bg.NameU='Page_background'
        cell(bg,'FillForegnd','RGB(255,255,255)');cell(bg,'LinePattern',0)
        for o in scene['elements']:
            kind=o['kind']
            if kind=='math':continue
            if kind in {'box','text','oval'}:
                method=p.DrawOval if kind=='oval' else p.DrawRectangle
                sh=method(o['x']/scale,(H-o['y']-o['h'])/scale,(o['x']+o['w'])/scale,(H-o['y'])/scale)
                cell(sh,'FillForegnd',rgb(o.get('fill','#FFFFFF')))
                cell(sh,'LineColor',rgb(o.get('line','#9DC9E5')))
                cell(sh,'LineWeight',f"{o.get('line_weight',.65)} pt")
                if kind=='box':cell(sh,'Rounding',o.get('round',6)/scale)
                if o.get('transparent'):cell(sh,'FillPattern',0)
                if kind=='text':
                    sh.Text=o['text'];cell(sh,'FillPattern',0);cell(sh,'LinePattern',0)
                    cell(sh,'Char.Font','FONT("Times New Roman")');cell(sh,'Char.AsianFont','FONT("宋体")')
                    cell(sh,'Char.Size',f"{o.get('size',18)*72/scale} pt")
                    cell(sh,'Char.Style',int(o.get('bold',False)));cell(sh,'Char.Color',rgb(o.get('color','#12253D')))
                    cell(sh,'Para.HorzAlign',o.get('align',0));cell(sh,'VerticalAlign',o.get('valign',1))
                    for n in ['LeftMargin','RightMargin','TopMargin','BottomMargin']:cell(sh,n,0)
                    for m in re.finditer(r'[\u2e80-\u9fff\uf900-\ufaff\uff00-\uffef\u3000-\u303f]+',o['text']):
                        ch=sh.Characters
                        # Visio indexes UTF-16 code units.
                        ch.Begin=len(o['text'][:m.start()].encode('utf-16-le'))//2
                        ch.End=len(o['text'][:m.end()].encode('utf-16-le'))//2
                        ch._oleobj_.Invoke(ch._oleobj_.GetIDsOfNames('CharProps'),0,pc.DISPATCH_PROPERTYPUT,0,0,d.Fonts.Item('宋体').ID)
            elif kind=='line':
                sh=p.DrawLine(o['x1']/scale,(H-o['y1'])/scale,o['x2']/scale,(H-o['y2'])/scale)
                cell(sh,'LineColor',rgb(o.get('color','#447CAC')));cell(sh,'LineWeight',f"{o.get('width',1.5)*72/scale} pt")
                if o.get('arrow'):cell(sh,'EndArrow',4)
            elif kind=='poly':
                pts=[v for x,y in o['points'] for v in (x/scale,(H-y)/scale)]
                sh=p.DrawPolyline(pts,0);cell(sh,'FillForegnd',rgb(o.get('fill','#EAF4FF')));cell(sh,'LineColor',rgb(o.get('line','#6BA3C6')))
            elif kind=='photo':
                sh=p.Import(str(image_file(o,scene,scene_path,output.parent/'crops')));place(sh,o,H,scale)
            sh.NameU=o['id']
        v.DeferRecalc=False;v.EventsEnabled=True;d.SaveAs(str(output))
    finally:
        if d is not None:d.Close()
        v.Quit()

def equations(scene_path,output,limit=6):
    # One fresh Visio AND Word process per small batch; no process-name-based cleanup.
    from vsdx_audit import audit
    scene=read_scene(scene_path);H=scene['height'];scale=scene.get('units_per_inch',100)
    saved=set(audit(output)['pages'][0]['math_shapes'])
    todo=[o for o in scene['elements'] if o['kind']=='math' and o['id'] not in saved][:limit]
    if not todo:return
    com,pc=office();v=com.DispatchEx('Visio.InvisibleApp');d=w=wd=None
    try:
        v.AlertResponse=7;d=v.Documents.OpenEx(str(Path(output).resolve()),64+128+8);p=d.Pages.Item(1)
        w=com.DispatchEx('Word.Application');w.Visible=False;w.DisplayAlerts=0
        for o in todo:
            print('MathType '+o['id'],flush=True)
            wd=w.Documents.Add();r=wd.Range(0,0);r.Text='$'+o['tex']+'$';r.Select()
            w.Run('MTCommand_TeXToggle')
            assert wd.InlineShapes.Count==1, 'MathType conversion did not produce one object'
            eq=wd.InlineShapes.Item(1);assert eq.OLEFormat.ProgID=='Equation.DSMT4'
            for attempt in range(3):
                try:
                    eq.Range.Copy();pc.PumpWaitingMessages();time.sleep(.6);p.PasteSpecial(65536);break
                except Exception:
                    if attempt==2:raise
                    time.sleep(.7)
            sh=p.Shapes.Item(p.Shapes.Count);sh.NameU=o['id'];place(sh,o,H,scale,True);cell(sh,'LinePattern',0)
            d.Save();eq=None;r=None;wd.Close(0);wd=None
        d.Save()
    finally:
        # Close only instances created here; never terminate another Office process.
        if wd is not None:
            try:wd.Close(0)
            except Exception:pass
        if w is not None:
            try:w.Quit(0)
            except Exception:pass
        if d is not None:d.Close()
        v.Quit()

def export(path,preview,page=1):
    com,pc=office();v=com.DispatchEx('Visio.InvisibleApp');d=None
    try:
        d=v.Documents.OpenEx(str(Path(path).resolve()),2+64+128+8)
        d.Pages.Item(page).Export(str(Path(preview).resolve()))
    finally:
        if d is not None:d.Close()
        v.Quit()
