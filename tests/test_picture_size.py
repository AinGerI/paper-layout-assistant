import importlib.util
from pathlib import Path
import tempfile
import unittest
import zipfile
import json
from unittest.mock import patch
from lxml import etree
from PIL import Image
ROOT=Path(__file__).resolve().parents[1]
spec=importlib.util.spec_from_file_location('picture_converter',ROOT/'scripts/markdown_to_word.py')
c=importlib.util.module_from_spec(spec);spec.loader.exec_module(c)
class PictureSizeTests(unittest.TestCase):
    def measure(self,pixels,dpi=None,metadata=None,page_limit=False):
        with tempfile.TemporaryDirectory() as tmp:
            d=Path(tmp);options={} if dpi is None else {'dpi':dpi}
            Image.new('RGB',pixels,'white').save(d/'figure.png',**options)
            (d/'s.md').write_text('# 标题\n\n![图](figure.png)',encoding='utf8')
            profile=ROOT/'profiles/chinese-paper.json'
            if page_limit:
                settings=json.loads(profile.read_text());settings['page'].update(width_cm=8,left_cm=2,right_cm=2);settings['figure'].update(max_width_cm=100,max_height_cm=100)
                profile=d/'p.json';profile.write_text(json.dumps(settings))
            original_open=c.Image.open
            def open_image(*args,**kwargs):
                im=original_open(*args,**kwargs)
                if metadata is not None:im.info['dpi']=metadata
                return im
            with patch.object(c.Image,'open',side_effect=open_image):c.convert(d/'s.md',d/'o.docx',d/'report',profile)
            self.picture_report=json.loads((d/'report/coverage.json').read_text())['pictures'][0]
            with zipfile.ZipFile(d/'o.docx') as z: root=etree.fromstring(z.read('word/document.xml'))
            e=root.find('.//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent')
            return int(e.get('cx'))/360000,int(e.get('cy'))/360000
    def test_small_image_not_enlarged(self):
        w,h=self.measure((100,100));self.assertAlmostEqual(w,100/96*2.54,places=3);self.assertAlmostEqual(h,w,places=3)
    def test_declared_dpi(self):
        w,h=self.measure((300,150),(300,300));self.assertAlmostEqual(w,2.54,places=3);self.assertAlmostEqual(h,1.27,places=3)
    def test_large_image_shrinks_proportionally(self):
        w,h=self.measure((2400,1200),(96,96));self.assertAlmostEqual(w,15.5,places=3);self.assertAlmostEqual(h,7.75,places=3)
    def test_unequal_dpi_preserves_pixel_ratio(self):
        w,h=self.measure((100,200),metadata=(100,200));self.assertAlmostEqual(w,1.27,places=3);self.assertAlmostEqual(h,2.54,places=3)
    def test_invalid_dpi_fallback_recorded(self):
        w,h=self.measure((100,100),metadata=(-1,0));self.assertAlmostEqual(w,100/96*2.54,places=3);self.assertEqual(self.picture_report['sizing_basis'],'estimated_96dpi')
    def test_tall_image_height_bound(self):
        w,h=self.measure((100,2000));self.assertAlmostEqual(h,18,places=3);self.assertAlmostEqual(w,.9,places=3)
    def test_page_bound_wins_over_profile(self):
        w,h=self.measure((2400,1200),page_limit=True);self.assertAlmostEqual(w,4,places=3);self.assertAlmostEqual(h,2,places=3)
if __name__=='__main__':unittest.main()
