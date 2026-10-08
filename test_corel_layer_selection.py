import unittest
from types import SimpleNamespace
from corel_export.vector_tools import CorelSession


class LayerSelectionTests(unittest.TestCase):
    def make_session(self):
        shapes=[SimpleNamespace(Layer=SimpleNamespace(Name=name,Visible=True),
                                Type=3,PowerClip=None,StaticID=index)
                for index,name in enumerate(('D4','CUT_OUT','INFO','_DXF_CHECK_12345678'),1)]
        session=CorelSession.__new__(CorelSession)
        session.current=lambda:None
        session.page=SimpleNamespace(Shapes=shapes)
        session.app=SimpleNamespace(ActiveSelectionRange=shapes[:2])
        session.shape_contours=lambda shape,progress=None:[dict(id=shape.StaticID,layer=shape.Layer.Name)]
        return session

    def test_unchecked_layers_not_read(self):
        session=self.make_session()
        contours,skipped=session.snapshot(False,layers={'D4'})
        self.assertEqual([c['layer'] for c in contours],['D4'])
        self.assertEqual(skipped,[])
        self.assertEqual(list(session.refs),[1])

    def test_selection_intersects_checked_layers(self):
        session=self.make_session()
        contours,_=session.snapshot(True,layers={'CUT_OUT'})
        self.assertEqual([c['layer'] for c in contours],['CUT_OUT'])

    def test_no_checked_layers_returns_no_contours(self):
        self.assertEqual(self.make_session().snapshot(False,layers=set())[0],[])

    def test_info_and_markers_still_excluded(self):
        contours,_=self.make_session().snapshot(False,layers={'INFO','_DXF_CHECK_12345678'})
        self.assertEqual(contours,[])
