"""Project-scoped image-to-Visio CLI. All user-visible drawings share one master."""
import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
import time
import uuid
from pathlib import Path
from vsdx_audit import audit, sha, same_content
from scene_renderer import read_scene

if hasattr(sys.stdout,'reconfigure'):sys.stdout.reconfigure(encoding='utf-8')

CONFIG='visio-project.json'

def write_json(path,data):
    path=Path(path);path.parent.mkdir(parents=True,exist_ok=True)
    tmp=path.with_name(path.name+'.'+uuid.uuid4().hex+'.tmp')
    tmp.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8');os.replace(tmp,path)

def config(project):
    root=Path(project).resolve();cp=root/CONFIG
    c=json.loads(cp.read_text(encoding='utf-8'))
    master=(root/c['master']).resolve()
    assert master.is_relative_to(root), 'Master must stay inside this project'
    return root,c,master

def init(project,master,register=None):
    root=Path(project).resolve();root.mkdir(parents=True,exist_ok=True)
    mp=(root/master).resolve();assert mp.is_relative_to(root)
    cp=root/CONFIG
    if cp.exists():
        _,c,existing=config(root)
        if existing!=mp:raise RuntimeError('Project already has a different master; choose explicitly before changing it.')
        return c
    c={'schema_version':1,'master':mp.relative_to(root).as_posix(),
       'fonts':{'latin':'Times New Roman','chinese':'SimSun'},'equations':'MathType OLE',
       'policy':'one-master-new-page-per-figure','figures':{}}
    if register:
        assert mp.exists()
        r=audit(mp);assert len(r['pages'])==1,'Registration requires an explicit page for multi-page files'
        c['figures'][register]=[{'page_name':r['pages'][0]['name'],'source_sha256':sha(mp),'registered_existing':True}]
    write_json(cp,c);return c

def signature(scene_path):
    path=Path(scene_path).resolve();s=read_scene(path)
    h=hashlib.sha256(path.read_bytes())
    files=set()
    if s.get('source_image'):files.add((path.parent/s['source_image']).resolve())
    for o in s['elements']:
        if o['kind']=='photo' and o.get('file'):files.add((path.parent/o['file']).resolve())
    for f in sorted(files):h.update(sha(f).encode())
    return h.hexdigest()

def worker(args,timeout=180):
    cmd=[sys.executable,str(Path(__file__).resolve()),*args]
    # Only this owned Python worker is terminated on a timeout. Never kill Office by name/PID.
    try:
        subprocess.run(cmd,check=True,timeout=timeout)
    except subprocess.TimeoutExpired as e:
        raise RuntimeError('Office worker timed out. Checkpoint retained; project master untouched. Inspect owned Office state before resuming; do not blindly terminate Office processes.') from e

def validate_page(path,scene_path):
    scene=read_scene(scene_path);r=audit(path)
    assert len(r['pages'])==1
    pg=r['pages'][0];expected={o['id'] for o in scene['elements']}
    assert expected<=set(pg['shape_names']), 'Missing named elements'
    math={o['id'] for o in scene['elements'] if o['kind']=='math'}
    assert math==set(pg['math_shapes']), 'Missing or unexpected MathType OLE objects'
    if any(o['kind']=='text' for o in scene['elements']):
        assert 'Times New Roman' in r['fonts'], 'Missing Latin font'
    if any(o['kind']=='text' and any('\u4e00'<=ch<='\u9fff' for ch in o['text']) for o in scene['elements']):
        assert 'SimSun' in r['fonts'] or '宋体' in r['fonts'], 'Missing Chinese font'
    return r

def render(scene_path,job):
    scene_path=Path(scene_path).resolve();job=Path(job).resolve();job.mkdir(parents=True,exist_ok=True)
    state=job/'build-state.json';out=job/'page.vsdx';sig=signature(scene_path)
    if state.exists():
        previous=json.loads(state.read_text(encoding='utf-8'))
        if previous['signature']!=sig:raise RuntimeError('Scene/assets changed; use a new build directory so the prior checkpoint remains intact.')
    else:
        if out.exists():raise RuntimeError('Untracked output exists; use a fresh job directory.')
        write_json(state,{'signature':sig,'scene':str(scene_path),'status':'building'})
    if not out.exists():worker(['_native','--scene',str(scene_path),'--output',str(out)])
    s=read_scene(scene_path);need=sum(o['kind']=='math' for o in s['elements'])
    while audit(out)['math_count']<need:
        before=audit(out)['math_count']
        worker(['_math','--scene',str(scene_path),'--output',str(out)],timeout=240)
        if audit(out)['math_count']<=before:raise RuntimeError('No equation progress; checkpoint retained.')
    report=validate_page(out,scene_path);write_json(job/'validation.json',report)
    worker(['_export','--input',str(out),'--output',str(job/'preview.png')])
    equations=[o for o in s['elements'] if o['kind']=='math']
    (job/'equations.md').write_text('# Formula source\n\n'+'\n\n'.join('## '+o['id']+'\n```latex\n'+o['tex']+'\n```' for o in equations),encoding='utf-8')
    write_json(state,{'signature':sig,'scene':str(scene_path),'status':'validated','output':str(out)})
    return {'output':str(out),'preview':str(job/'preview.png'),'math_count':need}

def merge_worker(source,candidate,page_name):
    from runtime import office
    com,pc=office();v=com.DispatchEx('Visio.InvisibleApp');dest=src=None
    try:
        v.AlertResponse=7;v.ScreenUpdating=False
        src=v.Documents.OpenEx(str(Path(source).resolve()),2+64+128+8)
        sp=src.Pages.Item(1)
        if Path(candidate).exists():
            dest=v.Documents.OpenEx(str(Path(candidate).resolve()),64+128+8);tp=dest.Pages.Add()
        else:
            dest=v.Documents.Add('');tp=dest.Pages.Item(1)
        tp.Name=page_name
        for name in ['PageWidth','PageHeight','PageScale','DrawingScale']:
            tp.PageSheet.CellsU(name).FormulaU=sp.PageSheet.CellsU(name).FormulaU
        selection=sp.CreateSelection(1,0x100)
        assert selection.Count==sp.Shapes.Count, 'Not all top-level source elements selected'
        for attempt in range(3):
            try:
                selection.Copy(1);pc.PumpWaitingMessages();time.sleep(.8+attempt);tp.Paste(1);break
            except Exception:
                if attempt==2:raise
                # A partial failed paste can only affect this new, uncommitted page.
                if tp.Shapes.Count:
                    tp.Delete(1);tp=dest.Pages.Add();tp.Name=page_name
                    for name in ['PageWidth','PageHeight','PageScale','DrawingScale']:
                        tp.PageSheet.CellsU(name).FormulaU=sp.PageSheet.CellsU(name).FormulaU
                pc.PumpWaitingMessages();time.sleep(1)
        assert tp.Shapes.Count==sp.Shapes.Count,'Copy/paste changed native top-level shape count'
        dest.SaveAs(str(Path(candidate).resolve()))
    finally:
        if src is not None:src.Close()
        if dest is not None:dest.Close()
        v.Quit()

def append(project,source,figure_id,page_name):
    root,c,master=config(project);source=Path(source).resolve()
    assert source!=master,'Source is already the project master; register it instead of duplicating it.'
    sr=audit(source);assert len(sr['pages'])==1,'Append one completed page at a time'
    digest=sha(source)
    history=c['figures'].get(figure_id,[])
    if master.exists():
        current=audit(master)
        for item in history:
            if item['source_sha256']==digest and any(p['name']==item['page_name'] for p in current['pages']):
                return {'status':'already-present','master':str(master),'page_name':item['page_name']}
    else:current={'pages':[]}
    name=page_name[:55];taken={p['name'] for p in current['pages']};version=2
    while name in taken:
        name=page_name[:48]+' - v'+str(version);version+=1
    work=root/'visio-assets'/'.staging';work.mkdir(parents=True,exist_ok=True)
    lock=root/'visio-assets'/'.project.lock'
    try:fd=os.open(lock,os.O_CREAT|os.O_EXCL|os.O_WRONLY)
    except FileExistsError:raise RuntimeError('Another project update or interrupted transaction owns the lock; inspect it before retrying.')
    os.write(fd,str(os.getpid()).encode());os.close(fd)
    candidate=work/(uuid.uuid4().hex+'.vsdx');before=sha(master) if master.exists() else None
    try:
        if master.exists():shutil.copy2(master,candidate)
        worker(['_merge','--input',str(source),'--output',str(candidate),'--page-name',name],timeout=240)
        after=audit(candidate)
        assert len(after['pages'])==len(current['pages'])+1,'Wrong project page count'
        for old in current['pages']:
            retained=next(p for p in after['pages'] if p['name']==old['name'])
            assert same_content(old,retained),'Existing page content changed: '+old['name']
        new=next(p for p in after['pages'] if p['name']==name)
        assert same_content(sr['pages'][0],new),'Native text or MathType lost during merge'
        if (sha(master) if master.exists() else None)!=before:
            raise RuntimeError('Master changed during generation. Candidate retained; no overwrite performed.')
        master.parent.mkdir(parents=True,exist_ok=True)
        if master.exists():
            backup=root/'visio-assets'/'.history'/(time.strftime('%Y%m%d-%H%M%S')+'-'+uuid.uuid4().hex[:8]+'.vsdx')
            backup.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(master,backup)
        # A Windows file lock causes a safe failure here; do not force-close the user's Visio.
        os.replace(candidate,master)
        c['figures'].setdefault(figure_id,[]).append({'page_name':name,'source_sha256':digest,'source':str(source)})
        write_json(root/CONFIG,c);write_json(root/'visio-assets'/'project-validation.json',after)
        return {'status':'appended','master':str(master),'page_name':name,'page_count':len(after['pages'])}
    finally:
        lock.unlink(missing_ok=True)

def main():
    p=argparse.ArgumentParser(description=__doc__);sub=p.add_subparsers(dest='command',required=True)
    q=sub.add_parser('init');q.add_argument('--project',required=True);q.add_argument('--master',default='项目绘图.vsdx');q.add_argument('--register-existing')
    q=sub.add_parser('status');q.add_argument('--project',required=True)
    q=sub.add_parser('render');q.add_argument('--scene',required=True);q.add_argument('--job',required=True)
    q=sub.add_parser('append');q.add_argument('--project',required=True);q.add_argument('--input',required=True);q.add_argument('--figure-id',required=True);q.add_argument('--page-name',required=True)
    q=sub.add_parser('audit');q.add_argument('--input',required=True)
    q=sub.add_parser('validate-scene');q.add_argument('--scene',required=True)
    for name in ['_native','_math']:
        q=sub.add_parser(name);q.add_argument('--scene',required=True);q.add_argument('--output',required=True)
    q=sub.add_parser('_export');q.add_argument('--input',required=True);q.add_argument('--output',required=True)
    q=sub.add_parser('_merge');q.add_argument('--input',required=True);q.add_argument('--output',required=True);q.add_argument('--page-name',required=True)
    a=p.parse_args();result=None
    if a.command=='init':result=init(a.project,a.master,a.register_existing)
    elif a.command=='status':
        root,c,m=config(a.project);result={'configuration':c,'master':str(m),'audit':audit(m) if m.exists() else None}
    elif a.command=='render':result=render(a.scene,a.job)
    elif a.command=='append':result=append(a.project,a.input,a.figure_id,a.page_name)
    elif a.command=='audit':result=audit(a.input)
    elif a.command=='validate-scene':result={'elements':len(read_scene(a.scene)['elements'])}
    elif a.command=='_merge':merge_worker(a.input,a.output,a.page_name)
    else:
        from scene_renderer import native,equations,export
        if a.command=='_native':native(a.scene,a.output)
        elif a.command=='_math':equations(a.scene,a.output)
        elif a.command=='_export':export(a.input,a.output)
    if result is not None:print(json.dumps(result,ensure_ascii=False,indent=2))

if __name__=='__main__':main()
