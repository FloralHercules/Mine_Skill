"""Orchestrate an editable Visio page conversion in an isolated work folder."""
from pathlib import Path
import argparse, hashlib, importlib.util, json, shutil, subprocess, sys
SCRIPTS=Path(__file__).resolve().parent

def quote_ps(value):return "'"+str(value).replace("'","''")+"'"
def run_py(script,*args):subprocess.run([sys.executable,str(SCRIPTS/script),*map(str,args)],check=True)
def run_ps(script,**params):
    shell=shutil.which('pwsh') or shutil.which('powershell')
    if not shell:raise RuntimeError('PowerShell with desktop Visio and PowerPoint is required')
    # PS5 reads BOM-less .ps1 as ANSI under -File. Read explicitly as UTF-8.
    command='& ([scriptblock]::Create([IO.File]::ReadAllText('+quote_ps(SCRIPTS/script)+',[Text.Encoding]::UTF8))) '
    command+=' '.join('-'+key if value is True else '-'+key+' '+quote_ps(value) for key,value in params.items() if value is not None and value is not False)
    try:subprocess.run([shell,'-NoProfile','-Command',command],check=True,timeout=240)
    except subprocess.TimeoutExpired as err:
        raise RuntimeError('Office conversion exceeded 240 seconds. Inspect the created Office instance and work logs before retrying; do not terminate unrelated Office processes.') from err

def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--source',type=Path,required=True);p.add_argument('--page',required=True)
    p.add_argument('--work-dir',type=Path,required=True);p.add_argument('--output',type=Path,required=True)
    p.add_argument('--slide-width-in',type=float);p.add_argument('--geometry-overrides',type=Path)
    p.add_argument('--resume',action='store_true',help='Reuse exports from an identical source/page in this work folder')
    p.add_argument('--overwrite-output',action='store_true',help='Replace only the exact output file explicitly selected')
    a=p.parse_args();source=a.source.resolve();work=a.work_dir.resolve();output=a.output.resolve()
    missing=[name for name in ('lxml','PIL') if importlib.util.find_spec(name) is None]
    if missing:p.error('Python dependencies missing: '+', '.join(missing)+'. Use an environment with lxml and Pillow before starting Office.')
    if not source.is_file():p.error('Source VSDX not found')
    if source.suffix.lower()!='.vsdx':p.error('Expected a .vsdx source')
    if output.suffix.lower()!='.pptx':p.error('Output must end with .pptx')
    if output.exists() and not a.overwrite_output:p.error('Output exists; choose a new filename or explicitly use --overwrite-output')
    if work==source.parent:p.error('Use a separate private work directory, not the source folder')
    signature={'source':str(source),'source_sha256':hashlib.sha256(source.read_bytes()).hexdigest(),'page':a.page,'slide_width_in':a.slide_width_in}
    marker=work/'conversion-context.json'
    if a.resume:
        if not marker.is_file() or json.loads(marker.read_text(encoding='utf8'))!=signature:p.error('Resume source/page/settings do not match the original work folder')
    else:
        if work.exists() and any(work.iterdir()):p.error('Work directory must be empty unless using --resume')
        work.mkdir(parents=True,exist_ok=True);marker.write_text(json.dumps(signature,ensure_ascii=False,indent=2),encoding='utf8')
        print('Exporting original and flattened Visio objects',flush=True)
        run_ps('export_visio.ps1',Source=source,PageName=a.page,WorkDir=work,ExportNative3DDiagnostics=True)
        args=['--source',work/'flattened.vsdx','--page-name',a.page,'--output',work/'manifest.json']
        if a.slide_width_in:args+=['--slide-width-in',a.slide_width_in]
        run_py('inspect_vsdx.py',*args)
    if a.geometry_overrides:
        override=a.geometry_overrides.resolve()
        if override!=work/'geometry-overrides.json':shutil.copyfile(override,work/'geometry-overrides.json')
    manifest=json.loads((work/'manifest.json').read_text(encoding='utf-8-sig'))
    if manifest['counts']['unresolved_groups']:
        p.error('Unresolved Visio groups remain; review export-report.json before assembling editable objects')
    if not all(0<x<=4032 for x in (manifest['slide_width_pt'],manifest['slide_height_pt'])):
        p.error('PowerPoint page dimensions must be within 56 inches; set --slide-width-in to scale proportionally')
    run_py('crop_rasters.py','--work-dir',work)
    run_py('native_geometry.py','--work-dir',work)
    run_py('native_text.py','--work-dir',work)
    run_ps('powerpoint.ps1',Mode='template',WorkDir=work)
    run_py('assemble_pptx.py','--work-dir',work)
    run_ps('powerpoint.ps1',Mode='verify',WorkDir=work)
    run_py('validate_output.py','--work-dir',work)
    output.parent.mkdir(parents=True,exist_ok=True)
    shutil.copyfile(work/'verified.pptx',output)
    print(json.dumps({'output':str(output),'preview':str(work/'preview.png'),'review_required':'Compare preview.png against reference.png and flattened-reference.png before final delivery.'},ensure_ascii=True))

if __name__=='__main__':main()
