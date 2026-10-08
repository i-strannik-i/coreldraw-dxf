import copy
import random
import unittest
from corel_export.vector_tools import excluded_layer
from duplicate_geometry import duplicate_groups, minimum_rotation
from vector_audit import audit
from marker_layout import place_label, overlaps


def square(sid=1,layer='CUT_OUT'):
    points=[(0,0),(10,0),(10,10),(0,10),(0,0)]
    return dict(id=sid,layer=layer,closed=True,segments=[('L',[a,b]) for a,b in zip(points,points[1:])])


class DuplicateTests(unittest.TestCase):
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

    def test_partial_subpath_not_duplicate_object(self):
        a=square();b=square(2);extra=square(2)
        extra['segments']=[(k,[(x+20,y) for x,y in p]) for k,p in extra['segments']]
        self.assertEqual(duplicate_groups([a,b,extra]),[])

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
