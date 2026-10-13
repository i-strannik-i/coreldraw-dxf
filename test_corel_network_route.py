import math
import unittest
from engine.network_route import closed_network,value


def line(a,b):return ('L',[a,a,b,b])
def contour(points,closed=False):
    return dict(layer='CUT',closed=closed,segments=[line(a,b) for a,b in zip(points,points[1:])])


class NetworkRouteTests(unittest.TestCase):
    def verify(self,plan):
        occurrences=[i for route in plan['routes'] for i in route['edge_ids']]
        for i,edge in enumerate(plan['edges']):
            if edge['source'] is not None:self.assertIn(i,occurrences)
            else:self.assertLessEqual(edge['weight'],plan['tolerance']+1e-10)
        for route in plan['routes']:
            for a,b in zip(route['segments'],route['segments'][1:]+route['segments'][:1]):
                self.assertLess(math.dist(a[1][-1],b[1][0]),1e-6)

    def test_crossing_dead_ends_require_return(self):
        plan=closed_network([contour([(-1,0),(1,0)]),contour([(0,-1),(0,1)])],.01)
        self.verify(plan)
        self.assertEqual(len(plan['routes']),1)
        self.assertEqual(plan['graph_intersections'],1)
        self.assertAlmostEqual(plan['source_length'],4)
        self.assertAlmostEqual(plan['routes'][0]['length'],8)

    def test_diagonal_has_minimum_repeat(self):
        plan=closed_network([contour([(0,0),(10,0),(10,10),(0,10),(0,0)],True),contour([(0,0),(10,10)])],.1)
        self.verify(plan)
        self.assertAlmostEqual(plan['routes'][0]['extra_length'],math.sqrt(200))

    def test_gap_respects_tolerance(self):
        source=[contour([(0,0),(10,0),(10,10),(0,10),(0,.06)])]
        joined=closed_network(source,.1);strict=closed_network(source,.01)
        self.verify(joined);self.verify(strict)
        self.assertAlmostEqual(joined['routes'][0]['length'],40,places=6)
        self.assertAlmostEqual(strict['routes'][0]['length'],79.88,places=6)
        self.assertFalse(any(e['source'] is None for e in strict['edges']))

    def test_disconnected_components_stay_separate(self):
        plan=closed_network([contour([(0,0),(1,0)]),contour([(5,0),(6,0)])],.1)
        self.verify(plan);self.assertEqual(len(plan['routes']),2)

    def test_endpoint_to_middle_within_tolerance(self):
        plan=closed_network([contour([(-2,0),(2,0)]),contour([(0,.05),(0,2)])],.1)
        self.verify(plan);self.assertEqual(len(plan['routes']),1)
        gaps=[e for e in plan['edges'] if e['source'] is None]
        self.assertTrue(gaps);self.assertLessEqual(max(e['weight'] for e in gaps),.1)

    def test_bezier_split_preserves_geometry(self):
        bezier=('B',[(-2,0),(-1,1),(1,-1),(2,0)])
        plan=closed_network([dict(layer='CUT',closed=False,segments=[bezier]),contour([(0,-2),(0,2)])],.01)
        self.verify(plan);self.assertEqual(len(plan['routes']),1)
        pieces=[e for e in plan['edges'] if e['source'] is not None and e['source'][0]==0]
        self.assertGreaterEqual(len(pieces),2)
        for edge in pieces:
            _,a,b=edge['source']
            for t in (0,.2,.5,.8,1):
                self.assertLess(math.dist(value(edge['segment'],t),value(bezier,a+(b-a)*t)),1e-6)

    def test_rejects_layers_and_overlap(self):
        a=contour([(0,0),(2,0)]);b=contour([(1,0),(3,0)])
        with self.assertRaisesRegex(ValueError,'наложения'):closed_network([a,b],.1)
        b['layer']='OTHER'
        with self.assertRaisesRegex(ValueError,'отдельно'):closed_network([a,b],.1)

    def test_endpoint_on_bezier_interior(self):
        b=('B',[(0,0),(0,10),(10,10),(10,0)])
        p=value(b,.37)
        plan=closed_network([dict(layer='CUT',closed=False,segments=[b]),contour([p,(p[0],15)])],.01)
        self.verify(plan);self.assertEqual(len(plan['routes']),1)


if __name__=='__main__':unittest.main()
