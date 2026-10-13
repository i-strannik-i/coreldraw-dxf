import unittest
import copy
from unittest.mock import Mock
from corel_export.vector_tools import CorelSession,layer_check_summary
from vector_audit import audit


def line(sid,layer,a,b):
    return dict(id=sid,layer=layer,closed=False,segments=[('L',[a,a,b,b])])


class LayerWorkflows(unittest.TestCase):
    def test_completed_layers_survive_cancellation(self):
        source=[line(1,'CUT_A',(0,0),(1,0)),line(2,'CUT_B',(0,0),(1,0))]
        reports=[]
        with self.assertRaises(InterruptedError):
            audit(source,cancelled=lambda:bool(reports),layer_done=reports.append)
        self.assertEqual([r['layer'] for r in reports],['CUT_A'])
        self.assertTrue(reports[0]['complete'])
        self.assertIn('Нулевые: 0',layer_check_summary(reports[0]))

    def test_separate_layers_never_join_even_at_identical_coordinates(self):
        s=CorelSession.__new__(CorelSession);s.current=Mock();s.app=Mock();s.doc=Mock();s.scale=1
        s.contours=[line(1,'A',(0,0),(1,0)),line(2,'A',(1,0),(2,0)),
                    line(3,'B',(0,0),(1,0)),line(4,'B',(1,0),(2,0))]
        s.refs={}
        for c in s.contours:
            shape=Mock();shape.StaticID=c['id'];shape.Layer.Name=c['layer'];shape.Type=3;shape.Fill.Type=0
            shape.Outline.Type=1;shape.Outline.Width=.1
            color=shape.Outline.Color.GetCopy.return_value
            color.RGBRed=0;color.RGBGreen=0;color.RGBBlue=0
            s.refs[c['id']]=shape
        s.join_selection=Mock(return_value=(list(s.refs.values()),s.refs))
        source=list(s.contours)
        s.shape_contours=Mock(side_effect=lambda shape,progress=None:[c for c in source if c['id']==shape.StaticID])
        before,after=s.prepare_join(.1,optimize_path=True,selection_only=False)
        self.assertEqual(before,(0,4));self.assertEqual(after,(0,2))
        self.assertEqual([source.Layer.Name for source,route in s.join_outputs],['A','B'])
        self.assertEqual(len(s.join_layer_reports),2)
        self.assertEqual(set(s.join_delete_refs),{1,2,3,4})
        s.app.CreateCurve.assert_not_called()
        self.assertFalse(s.join_selection_only)
        self.assertEqual(s.shape_contours.call_count,4)
        self.assertFalse(s.join_selection.call_args.args[1])
        # Failed quarantine blocks all optimization, even otherwise valid layers.
        duplicate=copy.deepcopy(s.contours[2]);duplicate['id']=5;s.contours.append(duplicate)
        shape=copy.deepcopy(s.refs[3]);shape.StaticID=5;s.refs[5]=shape;s.join_open_refs[5]=shape
        s.move_duplicates=Mock(side_effect=ValueError('Дубли недоступны'))
        for optimized in (False,True):
            with self.assertRaisesRegex(ValueError,'Дубли недоступны'):
                s.prepare_join(.1,refresh=False,optimize_path=optimized,selection_only=False)
            self.assertFalse(s.join_changed)

    def test_duplicates_added_after_preview_block_apply(self):
        s=CorelSession.__new__(CorelSession);s.join_changed=True;s.join_selection_only=False
        shapes=[Mock(StaticID=i) for i in (1,2)]
        s.join_selection=Mock(return_value=(shapes,dict(enumerate(shapes,1))))
        s.shape_contours=Mock(side_effect=lambda shape:[line(shape.StaticID,'A',(0,0),(1,0))])
        s._apply_join=Mock()
        with self.assertRaisesRegex(ValueError,'Появились дубли'):s.apply_join()
        s._apply_join.assert_not_called()

    def test_each_layer_reports_zero_issues(self):
        reports=[]
        result=audit([line(1,'A',(0,0),(1,0)),line(2,'B',(0,0),(1,0))],layer_done=reports.append)
        self.assertEqual(len(reports),2)
        self.assertTrue(result['complete']);self.assertFalse(result['issues'])
        self.assertTrue(all('проверен' in layer_check_summary(r) for r in reports))

if __name__=='__main__':unittest.main()
