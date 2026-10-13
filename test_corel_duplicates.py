import copy
import random
import unittest
from unittest.mock import Mock
from corel_export.vector_tools import CorelSession
from corel_export.vector_tools import excluded_layer
from duplicate_geometry import duplicate_groups, minimum_rotation
from vector_audit import audit
from marker_layout import place_label, overlaps


def square(sid=1,layer='CUT_OUT'):
    points=[(0,0),(10,0),(10,10),(0,10),(0,0)]
    return dict(id=sid,layer=layer,closed=True,segments=[('L',[a,b]) for a,b in zip(points,points[1:])])


class DuplicateTests(unittest.TestCase):
    def test_quarantine_preserves_scope_and_removes_only_copies(self):
        session=CorelSession.__new__(CorelSession)
        a,b,c=square(),square(2),square(3)
        c['segments']=[(k,[(x+30,y) for x,y in p]) for k,p in c['segments']]
        session.contours=[a,b,c]
        session.refs={1:object(),2:object(),3:object()}
        refs=dict(session.refs)
        def move(groups, preserve_selection=False,progress=None):
            self.assertTrue(preserve_selection)
            self.assertEqual(groups[0]['copies'],[2])
            session.contours=[];session.refs={}
            return 1
        session.move_duplicates=Mock(side_effect=move)
        self.assertEqual(session.quarantine_duplicates(),1)
        self.assertEqual(session.contours,[a,c])
        self.assertEqual(session.refs,{1:refs[1],3:refs[3]})
        self.assertEqual(session.signature,session.digest([a,c]))
        self.assertFalse(audit(session.contours)['issues'])

    def test_cancel_before_transfer_does_not_move(self):
        session=CorelSession.__new__(CorelSession)
        session.contours=[square(),square(2)]
        session.move_duplicates=Mock()
        with self.assertRaises(InterruptedError):
            session.quarantine_duplicates(lambda:True)
        session.move_duplicates.assert_not_called()

    def test_transfer_failure_stops_before_further_checks(self):
        session=CorelSession.__new__(CorelSession)
        session.contours=[square(),square(2)]
        session.refs={1:object(),2:object()}
        session.move_duplicates=Mock(side_effect=ValueError('locked'))
        with self.assertRaisesRegex(ValueError,'locked'):
            session.quarantine_duplicates()
        self.assertEqual(len(session.contours),2)

    def test_reverse_and_rotated_start(self):
        a=square();b=square(2)
        b['segments']=[(k,list(reversed(p))) for k,p in reversed(b['segments'])]
        b['segments']=b['segments'][2:]+b['segments'][:2]
        groups=duplicate_groups([a,b,square(3)])
        self.assertEqual(groups[0]['keep'],1)
        self.assertEqual(groups[0]['copies'],[2,3])
        self.assertEqual([i['kind'] for i in audit([a,b])['issues']],['duplicate'])

    def test_cross_layer_and_nearby_not_exact(self):
        a=square();b=square(2,'D4');c=square(3)
        c['segments']=[(k,[(x+.01,y) for x,y in p]) for k,p in c['segments']]
        self.assertEqual(duplicate_groups([a,b,c]),[])

    def test_partial_subpath_is_not_moved_as_whole_object(self):
        a=square();b=square(2);extra=square(2)
        extra['segments']=[(k,[(x+20,y) for x,y in p]) for k,p in extra['segments']]
        groups=duplicate_groups([a,b,extra])
        self.assertEqual(len(groups),1)
        self.assertEqual(groups[0]['mode'],'subpaths')
        self.assertEqual(groups[0]['members'],[{'id':1,'subpath':0},{'id':2,'subpath':0}])
        self.assertEqual([i['kind'] for i in audit([a,b,extra])['issues']],['duplicate'])

    def test_repeated_subpath_inside_combined_object(self):
        a=square();b=copy.deepcopy(a)
        groups=duplicate_groups([a,b])
        self.assertEqual(groups[0]['members'],[{'id':1,'subpath':0},{'id':1,'subpath':1}])
        self.assertEqual([i['kind'] for i in audit([a,b])['issues']],['duplicate'])

    def test_partial_duplicates_never_cross_layers(self):
        self.assertFalse(duplicate_groups([square(1,'D4'),square(2,'BACK_D4')]))

    def test_subpath_order_and_bezier_direction(self):
        a=dict(id=1,layer='CUT_ON',closed=False,segments=[('B',[(0,0),(1,2),(3,4),(5,6)])])
        b=copy.deepcopy(a);b['id']=2;b['segments'][0][1].reverse()
        self.assertEqual(len(duplicate_groups([a,b])),1)
        b['closed']=True
        self.assertEqual(duplicate_groups([a,b]),[])

    def test_rotation_matches_exhaustive_reference(self):
        rng=random.Random(4)
        for size in range(1,60):
            values=tuple(rng.randrange(4) for _ in range(size))
            self.assertEqual(minimum_rotation(values),min(values[i:]+values[:i] for i in range(size)))

    def test_service_layers_excluded(self):
        self.assertTrue(excluded_layer('_DXF_DUPLICATES__CUT_OUT__abc12345'))

    def test_dense_labels_do_not_overlap(self):
        occupied=[];markers=[(-1,-1,1,1)]
        for kind in ['zero','overlap','intersection','open','duplicate']*8:
            box=place_label((0,0),kind,3,3,occupied,markers)
            self.assertFalse(any(overlaps(box,b) for b in occupied))
            if kind=='zero':self.assertGreater(box[1],0)
            if kind=='overlap':self.assertLess(box[2],0)
            if kind=='intersection':self.assertGreater(box[0],0)
            if kind=='open':self.assertLess(box[3],0)
            occupied.append(box)
