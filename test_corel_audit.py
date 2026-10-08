import unittest
from pathlib import Path
import sys
sys.path.insert(0,str(Path(__file__).parent/'engine'))
from vector_audit import audit,ambiguous_endpoints


def contour(points,sid=1,closed=False,layer='CUT_ON'):
    return dict(id=sid,layer=layer,closed=closed,segments=[('L',[a,a,b,b]) for a,b in zip(points,points[1:])])


class AuditTests(unittest.TestCase):
    def test_rectangle_no_false_intersections(self):
        result=audit([contour([(0,0),(10,0),(10,10),(0,10),(0,0)],closed=True)])
        self.assertEqual(result['issues'],[])

    def test_crossing_location(self):
        r=audit([contour([(0,0),(10,10)]),contour([(0,10),(10,0)],2)])
        self.assertEqual(r['issues'][0]['point'],(5,5))
        self.assertEqual(r['issues'][0]['ids'],[1,2])

    def test_backtrack_detected_even_adjacent(self):
        r=audit([contour([(0,0),(10,0),(5,0)])])
        self.assertEqual(r['issues'][0]['kind'],'overlap')

    def test_bowtie_self_intersection(self):
        r=audit([contour([(0,0),(10,10),(0,10),(10,0),(0,0)],closed=True)])
        self.assertEqual(len(r['issues']),1)

    def test_layers_separate_unless_requested(self):
        c=[contour([(0,0),(10,10)]),contour([(0,10),(10,0)],2,layer='V_3')]
        self.assertFalse(audit(c)['issues'])
        self.assertTrue(audit(c,True)['issues'])

    def test_zero_and_open(self):
        r=audit([contour([(0,0),(0,0),(10,0)],layer='CUT_OUT')])
        self.assertEqual({i['kind'] for i in r['issues']},{'zero','open'})

    def test_ambiguous_join_is_blocked(self):
        c=[contour([(0,0),(1,0)]),contour([(1,0),(2,0)],2),contour([(1,0),(1,1)],3)]
        self.assertTrue(ambiguous_endpoints(c,.1))

    def test_unique_chain_allowed(self):
        self.assertFalse(ambiguous_endpoints([contour([(0,0),(1,0)]),contour([(1.05,0),(2,0)],2)],.1))
