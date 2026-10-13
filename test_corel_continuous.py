import math
import unittest
from engine.continuous_vectors import continuous_vectors,continuous_summary,corel_segments,same_vectors
from engine.network_route import value
from test_corel_network_route import contour


class ContinuousTests(unittest.TestCase):
    def verify(self,source,tolerance=.1):
        plan=continuous_vectors(source,tolerance)
        visits=[i for r in plan['routes'] for i in r['edge_ids']]
        self.assertEqual(len(visits),len(set(visits)))
        for i,e in enumerate(plan['edges']):
            if e['source'] is not None:
                self.assertIn(i,visits)
                si,a,b=e['source'];original=[s for c in source for s in c['segments']][si]
                for t in (0,.25,.5,.75,1):
                    self.assertLess(math.dist(value(original,a+(b-a)*t),value(e['segment'],t)),2e-6)
            elif i in visits:self.assertLessEqual(e['weight'],tolerance)
        for r in plan['routes']:
            for a,b in zip(r['segments'],r['segments'][1:]):self.assertLess(math.dist(a[1][-1],b[1][0]),1e-6)
        return plan

    def test_progress_preserves_geometry_and_can_cancel(self):
        source=[contour([(0,0),(10,0),(10,10),(0,.05)])]
        events=[]
        actual=continuous_vectors(source,.1,progress=lambda *e:events.append(e))
        self.assertEqual(actual,continuous_vectors(source,.1))
        self.assertTrue(any(e[0]=='Поиск пересечений' for e in events))
        for target in {e[0] for e in events}:
            def cancel(stage,done,total):
                if stage==target:raise InterruptedError('cancel')
            with self.assertRaises(InterruptedError):continuous_vectors(source,.1,progress=cancel)

    def test_unchanged_vectors_ignore_direction_and_order(self):
        source=[contour([(0,0),(10,0)]),contour([(0,2),(10,2)])]
        reverse=[contour([(10,2),(0,2)]),contour([(10,0),(0,0)])]
        self.assertTrue(same_vectors(source,reverse,1))
        self.assertFalse(same_vectors(source,reverse[:1],1))
        self.assertFalse(same_vectors(source,[contour([(0,0),(10,0),(0,2),(10,2)])],1))

    def test_unchanged_closed_start_point(self):
        a=[contour([(0,0),(10,0),(10,10),(0,0)],True)]
        b=[contour([(10,0),(10,10),(0,0),(10,0)],True)]
        self.assertTrue(same_vectors(a,b,1))
        b[0]['closed']=False
        self.assertFalse(same_vectors(a,b,1))

    def test_tail_first_one_open(self):
        p=self.verify([contour([(0,0),(10,0),(10,10),(0,10),(0,0)],True),contour([(0,0),(-3,0)])])
        self.assertEqual(len(p['routes']),1);self.assertFalse(p['routes'][0]['closed'])
        self.assertEqual(p['routes'][0]['segments'][0][1][0],(-3,0))

    def test_four_ends_need_two_vectors(self):
        p=self.verify([contour([(-1,0),(1,0)]),contour([(0,-1),(0,1)])])
        self.assertEqual(len(p['routes']),2)

    def test_eight_ends_need_four(self):
        p=self.verify([contour([(0,0),(math.cos(i*math.pi/4)*10,math.sin(i*math.pi/4)*10)]) for i in range(8)])
        self.assertEqual(len(p['routes']),4)

    def test_gap_tolerance(self):
        cs=[contour([(0,0),(10,0),(10,10),(0,10),(0,.06)])]
        p=self.verify(cs,.1);q=self.verify(cs,.01)
        # Closing a lone open vector is desirable even if vector count stays one.
        self.assertTrue(p['routes'][0]['closed']);self.assertFalse(q['routes'][0]['closed'])

    def test_disconnected(self):
        p=self.verify([contour([(0,0),(1,0)]),contour([(3,0),(4,0)])])
        self.assertEqual(len(p['routes']),2)

    def test_two_close_parts_join(self):
        p=self.verify([contour([(0,0),(1,0)]),contour([(1.05,0),(2,0)])])
        self.assertEqual(len(p['routes']),1);self.assertEqual(p['connectors'],1)

    def test_bezier_crossing(self):
        p=self.verify([dict(layer='CUT',closed=False,segments=[('B',[(-2,0),(-1,1),(1,-1),(2,0)])]),contour([(0,-2),(0,2)])])
        self.assertEqual(len(p['routes']),2)

    def test_no_cross_layer(self):
        a=contour([(0,0),(1,0)]);b=contour([(1,0),(2,0)]);b['layer']='OTHER'
        with self.assertRaises(ValueError):continuous_vectors([a,b],.1)

    def test_summary(self):
        s=continuous_summary((0,4),(1,1),dict(connectors=2,max_connector=.05))
        self.assertIn('4 → 2',s);self.assertNotIn('Расчётный путь',s)

    def test_duplicate_bezier_rejected(self):
        c=dict(layer='CUT',closed=False,segments=[('B',[(0,0),(1,2),(3,2),(4,0)])])
        with self.assertRaisesRegex(ValueError,'совпадающие'):continuous_vectors([c,c],.1)

    def test_single_bezier_near_closed(self):
        c=dict(layer='CUT',closed=False,segments=[('B',[(0,0),(5,8),(-5,8),(.05,0)])])
        p=self.verify([c]);self.assertTrue(p['routes'][0]['closed'])

    def test_many_odd_nodes_no_size_limit(self):
        p=self.verify([contour([(0,0),(10*math.cos(i*.2),10*math.sin(i*.2))]) for i in range(30)])
        self.assertEqual(len(p['routes']),15)

    def test_corel_precision_no_new_zero_edges(self):
        segments=contour([(0,0),(1,0),(1.000003,0),(2,0)])['segments']
        result=corel_segments(segments,1)
        self.assertEqual(len(result),2)
        self.assertEqual(result[0][1][-1],result[1][1][0])
        self.assertEqual(result[-1][1][-1],(2,0))

    def test_connector_rounded_outside_tolerance_is_rejected(self):
        # Diagonal rounding can make an originally permitted bridge too long.
        cs=[contour([(-1,0),(.000049,.000049)]),contour([(.060051,.08003),(1,1)])]
        p=continuous_vectors(cs,.1,1)
        self.assertEqual(p['connectors'],0)


if __name__=='__main__':unittest.main()
