import json
from pathlib import Path
import tempfile
import unittest
from PIL import Image, ImageDraw
from crop_rasters import crop


class CropTests(unittest.TestCase):
    def fixture(self, folder, content=True):
        folder=Path(folder);(folder/'assets').mkdir()
        image=Image.new('RGBA',(300,200),(0,0,0,0))
        draw=ImageDraw.Draw(image)
        draw.rectangle((0,0,299,199),outline=(255,0,255,255),width=1)
        if content:draw.rectangle((100,50,119,69),fill=(255,0,255,255))
        image.save(folder/'assets/shape-7.png')
        manifest={'slide_width_pt':600,'slide_height_pt':400,'objects':[{'source_id':7,'raster':True}]}
        (folder/'manifest.json').write_text(json.dumps(manifest),encoding='utf8')
        return folder

    def test_preserves_picture_colors_and_position(self):
        with tempfile.TemporaryDirectory() as folder:
            folder=self.fixture(folder);crop(folder)
            result=json.loads((folder/'manifest.json').read_text(encoding='utf8'))['objects'][0]
            self.assertEqual(result['placement'],{'x':200,'y':100,'w':40,'h':40})
            with Image.open(result['asset']) as image:
                self.assertEqual(image.size,(20,20))
                self.assertEqual(image.getpixel((10,10)),(255,0,255,255))

    def test_empty_export_is_rejected(self):
        with tempfile.TemporaryDirectory() as folder:
            folder=self.fixture(folder,False)
            with self.assertRaisesRegex(ValueError,'Empty isolated export'):
                crop(folder)


if __name__=='__main__':unittest.main()
