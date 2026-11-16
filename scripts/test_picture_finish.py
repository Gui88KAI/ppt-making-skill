import hashlib
import tempfile
import unittest
from pathlib import Path
from PIL import Image
from pptx import Presentation
from pptx.util import Inches,Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.oxml.ns import qn
from picture_finish import soft_edge,remove_picture_frames
from render_preview import render


class PictureFinishTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.root=Path(self.temp.name)
        Image.new('RGB',(100,100),'red').save(self.root/'photo.png')
        self.prs=Presentation()
        self.slide=self.prs.slides.add_slide(self.prs.slide_layouts[6])

    def tearDown(self):
        self.temp.cleanup()

    def test_soft_edge_is_nondestructive_and_updates_instead_of_duplicating(self):
        pic=self.slide.shapes.add_picture(str(self.root/'photo.png'),0,0)
        digest=hashlib.sha256(pic.image.blob).hexdigest()
        soft_edge(pic,12);soft_edge(pic,18)
        self.prs.save(self.root/'deck.pptx')
        loaded=Presentation(self.root/'deck.pptx').slides[0].shapes[0]
        self.assertEqual(hashlib.sha256(loaded.image.blob).hexdigest(),digest)
        effects=loaded._element.findall('.//'+qn('a:softEdge'))
        self.assertEqual(len(effects),1)
        self.assertEqual(int(effects[0].get('rad')),int(Pt(18)))
        soft_edge(loaded,0)
        self.assertEqual(loaded._element.findall('.//'+qn('a:softEdge')),[])

    def test_frame_removal_preserves_content_panels(self):
        frame=self.slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE,100000,100000,440000,440000)
        frame.fill.solid();frame.fill.fore_color.rgb=RGBColor.from_string('F2E0C4')
        pic=self.slide.shapes.add_picture(str(self.root/'photo.png'),120000,120000,400000,400000)
        panel=self.slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE,800000,100000,500000,500000)
        panel.fill.solid();panel.fill.fore_color.rgb=RGBColor.from_string('F2E0C4')
        panel.text='内容保留'
        self.assertEqual(remove_picture_frames(self.slide),[frame.shape_id])
        ids=[s.shape_id for s in self.slide.shapes]
        self.assertIn(pic.shape_id,ids);self.assertIn(panel.shape_id,ids)

    def test_preview_respects_transparency_and_native_soft_edge(self):
        self.slide.background.fill.solid()
        self.slide.background.fill.fore_color.rgb=RGBColor(0,0,255)
        pic=self.slide.shapes.add_picture(str(self.root/'photo.png'),0,0,Inches(3),Inches(3))
        soft_edge(pic,24)
        self.prs.save(self.root/'deck.pptx')
        render(self.root/'deck.pptx',self.root/'preview',width=400)
        with Image.open(self.root/'preview'/'p01.png') as preview:
            self.assertEqual(preview.getpixel((0,0)),(0,0,255))
            self.assertEqual(preview.getpixel((60,60)),(255,0,0))
        # Real transparent PNG should also preserve slide color without soft edges.
        transparent=Image.new('RGBA',(20,20),(255,0,0,0))
        transparent.save(self.root/'transparent.png')
        p=Presentation();s=p.slides.add_slide(p.slide_layouts[6])
        s.background.fill.solid();s.background.fill.fore_color.rgb=RGBColor(0,0,255)
        s.shapes.add_picture(str(self.root/'transparent.png'),0,0,Inches(2),Inches(2))
        p.save(self.root/'alpha.pptx')
        render(self.root/'alpha.pptx',self.root/'alpha_preview',width=400)
        with Image.open(self.root/'alpha_preview'/'p01.png') as preview:
            self.assertEqual(preview.getpixel((10,10)),(0,0,255))


if __name__=='__main__':
    unittest.main()
