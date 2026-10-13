import unittest
from unittest.mock import Mock
from corel_export.vector_tools import CorelSession,excluded_layer


class JoinProgressTests(unittest.TestCase):
    def session(self):
        s=CorelSession.__new__(CorelSession)
        s.app=Mock();s.doc=Mock();s.current=Mock()
        s.refs={1:Mock(),2:Mock()}
        s.join_outputs=[(Mock(),Mock()) for _ in range(3)]
        s.join_delete_refs=s.refs.copy();s.marker_layers=Mock(return_value=[])
        return s

    def test_create_and_replace_progress(self):
        s=self.session();events=[]
        s._apply_join(lambda *event:events.append(event))
        self.assertIn(('Создание векторов',0,3),events)
        self.assertIn(('Создание векторов',3,3),events)
        self.assertIn(('Подготовка замены исходных объектов',2,2),events)
        self.assertLess(events.index(('Создание векторов',3,3)),events.index(('Замена исходных объектов',0,0)))
        s.doc.EndCommandGroup.assert_called_once();s.doc.Undo.assert_not_called()

    def test_failure_reports_rollback(self):
        s=self.session();events=[]
        s.join_outputs[1][0].Layer.CreateCurve.side_effect=RuntimeError('test')
        with self.assertRaises(RuntimeError):s._apply_join(lambda *event:events.append(event))
        self.assertEqual(events[-1][0],'Отмена незавершённых изменений')
        s.doc.Undo.assert_called_once()

    def test_validation_reports_objects(self):
        s=self.session();events=[]
        s.shape_contours=Mock(return_value=[]);s.signature=s.digest([])
        s.unchanged(lambda *event:events.append(event))
        self.assertEqual([e[1:] for e in events],[(0,2),(1,2),(2,2)])

    def test_service_layers_ignored_in_snapshot_and_join(self):
        s=self.session();shapes=[]
        for index,name in enumerate(('INFO','#','CUT')):
            shape=Mock();shape.Layer.Name=name;shape.Layer.Visible=True;shape.Layer.Editable=True
            shape.Type=3;shape.StaticID=index;shape.PowerClip=None;shape.Fill.Type=0
            shapes.append(shape)
        s.app.ActiveSelectionRange=shapes
        s.shape_contours=Mock(return_value=[])
        roots,leaves=s.join_selection()
        self.assertEqual(set(leaves),{2})
        s.snapshot(True)
        s.shape_contours.assert_called_once_with(shapes[2],None)
        self.assertEqual(set(s.refs),{2})
        self.assertTrue(excluded_layer(' info '));self.assertTrue(excluded_layer(' # '))
        self.assertFalse(excluded_layer('CUT'))


if __name__=='__main__':unittest.main()
