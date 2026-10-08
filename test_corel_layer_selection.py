import unittest
from types import SimpleNamespace
from corel_export.vector_tools import CorelSession, excluded_layer, ISSUE_STYLES, draw_issue_symbol


class LayerSelectionTests(unittest.TestCase):
    def test_issue_colors_are_distinct_and_readable_on_white(self):
        colors=[style[2] for style in ISSUE_STYLES.values()]
        self.assertEqual(len(set(colors)),len(colors))
        for color in colors:
            channels=[int(color[i:i+2],16)/255 for i in (1,3,5)]
            linear=[v/12.92 if v<=.04045 else ((v+.055)/1.055)**2.4 for v in channels]
            luminance=sum(v*w for v,w in zip(linear,(.2126,.7152,.0722)))
            self.assertGreaterEqual(1.05/(luminance+.05),4.5,color)

    def test_complete_check_replaces_marks_even_if_empty(self):
        from unittest.mock import Mock
        session=CorelSession.__new__(CorelSession);session.mark=Mock()
        self.assertTrue(session.refresh_marks(dict(complete=True,issues=[])))
        session.mark.assert_called_once_with([],None)

    def test_incomplete_check_preserves_marks(self):
        from unittest.mock import Mock
        session=CorelSession.__new__(CorelSession);session.mark=Mock()
        self.assertFalse(session.refresh_marks(dict(complete=False,issues=[])))
        session.mark.assert_not_called()

    def test_cancelled_check_preserves_marks(self):
        from unittest.mock import Mock
        session=CorelSession.__new__(CorelSession);session.mark=Mock()
        with self.assertRaises(InterruptedError):
            session.refresh_marks(dict(complete=True,issues=[]),lambda:True)
        session.mark.assert_not_called()

    def test_technical_layer_filter(self):
        for name in ['Guides',' INFO ','Desktop','Document Grid','Направляющие','_DXF_CHECK_abc']:
            self.assertTrue(excluded_layer(name),name)
        for name in ['D4','CUT_OUT','P4_1,5mm','Layer 1','Guide design']:
            self.assertFalse(excluded_layer(name),name)

    def test_technical_geometry_is_not_read(self):
        session=self.make_session()
        session.page.Shapes[0].Layer.Name='Guides'
        contours,_=session.snapshot(False)
        self.assertEqual([c['layer'] for c in contours],['CUT_OUT'])

    def test_legend_draws_all_kinds(self):
        from unittest.mock import Mock
        for kind,style in ISSUE_STYLES.items():
            canvas=Mock()
            draw_issue_symbol(canvas,kind,style[2])
            self.assertTrue(canvas.mock_calls)

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
