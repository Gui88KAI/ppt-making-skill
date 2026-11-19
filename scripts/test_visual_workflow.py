import tempfile
import unittest
from pathlib import Path
from PIL import Image, PngImagePlugin
from pptx import Presentation
from pptx.util import Inches
from pptx.oxml.ns import qn
from entry_motion import click_fades
from image_usage_check import inspect


class VisualWorkflowTests(unittest.TestCase):
    def test_reencoded_identical_picture_is_duplicate(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder)
            im=Image.new('RGB',(80,60),'#cc8844')
            im.save(root/'one.png')
            info=PngImagePlugin.PngInfo();info.add_text('different','metadata')
            im.save(root/'renamed.png',pnginfo=info)
            self.assertNotEqual((root/'one.png').read_bytes(),(root/'renamed.png').read_bytes())
            prs=Presentation()
            for name in ('one.png','renamed.png'):
                prs.slides.add_slide(prs.slide_layouts[6]).shapes.add_picture(str(root/name),0,0)
            prs.save(root/'deck.pptx')
            result=inspect(root/'deck.pptx')
            self.assertEqual(result['status'],'failed')
            self.assertEqual(len(result['duplicates']),1)
            self.assertEqual(len(result['duplicates'][0]),2)

    def test_baseline_reuse_is_reported(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder);old=root/'old';old.mkdir()
            Image.new('RGB',(60,40),'navy').save(old/'old.png')
            prs=Presentation();s=prs.slides.add_slide(prs.slide_layouts[6])
            s.shapes.add_picture(str(old/'old.png'),0,0)
            prs.save(root/'deck.pptx')
            result=inspect(root/'deck.pptx',baseline=old)
            self.assertEqual(len(result['baseline_matches']),1)
            self.assertEqual(result['status'],'failed')

    def test_grouped_pictures_are_checked(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder)
            Image.new('RGB',(50,50),'teal').save(root/'pic.png')
            prs=Presentation();s=prs.slides.add_slide(prs.slide_layouts[6])
            group=s.shapes.add_group_shape()
            group.shapes.add_picture(str(root/'pic.png'),0,0)
            s.shapes.add_picture(str(root/'pic.png'),Inches(2),0)
            prs.save(root/'deck.pptx')
            result=inspect(root/'deck.pptx')
            self.assertEqual(result['checked_pictures'],2)
            self.assertEqual(len(result['duplicates']),1)

    def test_click_groups_resolve_actual_shapes_and_unique_time_ids(self):
        prs=Presentation();s=prs.slides.add_slide(prs.slide_layouts[6])
        shapes=[s.shapes.add_textbox(Inches(i),0,Inches(1),Inches(1)) for i in range(3)]
        groups=[[shapes[0].shape_id,shapes[1].shape_id],[shapes[2].shape_id]]
        click_fades(s,groups)
        timing=s._element.find(qn('p:timing'))
        nodes=timing.findall('.//'+qn('p:cTn'))
        ids=[n.get('id') for n in nodes]
        self.assertEqual(len(ids),len(set(ids)))
        self.assertEqual(sum(n.get('nodeType')=='clickEffect' for n in nodes),2)
        self.assertEqual({int(n.get('spid')) for n in timing.findall('.//'+qn('p:spTgt'))},set(sum(groups,[])))
        self.assertEqual(len(timing.findall('.//'+qn('p:animEffect'))),3)

    def test_invalid_animation_target_does_not_mutate_slide(self):
        prs=Presentation();s=prs.slides.add_slide(prs.slide_layouts[6])
        with self.assertRaises(ValueError):
            click_fades(s,[[9999]])
        self.assertIsNone(s._element.find(qn('p:timing')))



if __name__=='__main__':
    unittest.main()
