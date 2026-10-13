import copy
import unittest
from unittest.mock import Mock
from corel_export.vector_tools import CorelSession


class FocusTests(unittest.TestCase):
    def setUp(self):
        self.session=CorelSession.__new__(CorelSession)
        self.session.current=Mock()
        self.session.unchanged=Mock(side_effect=AssertionError('Full-document check during navigation'))
        self.session.app=Mock();self.session.doc=Mock();self.session.scale=1
        self.session.refs={n:Mock() for n in range(1,101)}
        for shape in self.session.refs.values():shape.Layer.Visible=True
        self.session.contours=[dict(id=n,subpath=1,layer='CUT',closed=True,segments=[n]) for n in range(1,101)]
        self.session.shape_contours=Mock(side_effect=lambda shape:[copy.deepcopy(self.session.contours[next(n for n,s in self.session.refs.items() if s is shape)-1])])
        self.issue=dict(ids=[2,1],point=[10,20])

    def test_only_issue_objects_are_read_and_selected(self):
        self.session.focus(self.issue)
        self.assertEqual(self.session.shape_contours.call_count,2)
        self.session.unchanged.assert_not_called()
        self.assertEqual(self.session.app.CreateShapeRange.return_value.Add.call_count,2)
        self.session.doc.ActiveWindow.ActiveView.SetViewArea.assert_called_once()

    def test_changed_or_deleted_or_hidden_object_does_not_zoom(self):
        for change in ('geometry','deleted','hidden'):
            with self.subTest(change=change):
                self.setUp()
                if change=='geometry':self.session.shape_contours.return_value=[];self.session.shape_contours.side_effect=None
                if change=='deleted':del self.session.refs[1]
                if change=='hidden':self.session.refs[1].Layer.Visible=False
                with self.assertRaisesRegex(ValueError,'Повторите проверку'):
                    self.session.focus(self.issue)
                self.session.doc.ActiveWindow.ActiveView.SetViewArea.assert_not_called()

    def test_unrelated_object_is_not_read(self):
        original=self.session.shape_contours.side_effect
        def read(shape):
            if shape is self.session.refs[100]:raise RuntimeError('Unrelated object unavailable')
            return original(shape)
        self.session.shape_contours.side_effect=read
        self.session.focus(self.issue)
