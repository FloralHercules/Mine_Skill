"""Crop isolated Visio bitmap exports and retain slide placement."""
from pathlib import Path
import argparse, json
from PIL import Image, ImageDraw

def crop(work):
    manifest=json.loads((work/'manifest.json').read_text(encoding='utf-8-sig'))
    for obj in manifest['objects']:
        if not obj['raster']:continue
        file=work/'assets'/f"shape-{obj['source_id']}.png"
        image=Image.open(file).convert('RGBA')
        alpha=image.getchannel('A')
        # The visible calibration frame is at the page edge, never content.
        # Clear its perimeter only: legitimate magenta inside a picture stays.
        ImageDraw.Draw(alpha).rectangle((0,0,image.width-1,image.height-1),outline=0,width=3)
        image.putalpha(alpha)
        bbox=alpha.point(lambda x:255 if x>8 else 0).getbbox()
        if bbox is None:raise ValueError(f"Empty isolated export for shape {obj['source_id']}")
        left,top,right,bottom=bbox
        dest=work/'assets'/f"cropped-{obj['source_id']}.png"
        image.crop(bbox).save(dest)
        obj['asset']=str(dest.resolve())
        obj['placement']={'x':left/image.width*manifest['slide_width_pt'],'y':top/image.height*manifest['slide_height_pt'],'w':(right-left)/image.width*manifest['slide_width_pt'],'h':(bottom-top)/image.height*manifest['slide_height_pt']}
    (work/'manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding='utf8')
    print(json.dumps({'cropped':sum(bool(o['raster']) for o in manifest['objects'])}))

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--work-dir',type=Path,required=True);a=p.parse_args();crop(a.work_dir.resolve())
