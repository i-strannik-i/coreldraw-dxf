import itertools
import math
import unittest
from engine.closed_route import shortest_closed_route,oriented_segments


def contour(a,b):return dict(closed=False,segments=[('B',[a,(a[0]+1,a[1]+2),(b[0]-1,b[1]-2),b])])


class ClosedRouteTests(unittest.TestCase):
    def test_matches_exhaustive_search(self):
        curves=[contour((0,1),(2,5)),contour((8,3),(9,9)),contour((4,-6),(10,-5)),
                contour((-8,3),(-4,8)),contour((-6,-5),(-9,-1)),contour((12,12),(3,14))]
        optimum=math.inf
        for permutation in itertools.permutations(range(1,len(curves))):
            for reverse in itertools.product((False,True),repeat=len(curves)-1):
                sequence=[oriented_segments(curves[0])]+[oriented_segments(curves[i],r) for i,r in zip(permutation,reverse)]
                distance=sum(math.dist(sequence[i][-1][1][-1],sequence[(i+1)%len(sequence)][0][1][0]) for i in range(len(sequence)))
                optimum=min(optimum,distance)
        plan=shortest_closed_route(curves)
        self.assertAlmostEqual(plan['added_length'],optimum)
        self.assertEqual(sorted(i for i,r in plan['order']),list(range(len(curves))))
        for current,following in zip(plan['segments'],plan['segments'][1:]+plan['segments'][:1]):
            self.assertEqual(current[1][-1],following[1][0])

    def test_reversal_preserves_bezier(self):
        c=contour((0,0),(5,5))
        self.assertEqual(oriented_segments(c,True)[0][1],list(reversed(c['segments'][0][1])))

    def test_one_open_curve_gets_closing_segment(self):
        plan=shortest_closed_route([contour((0,0),(3,4))])
        self.assertEqual(plan['added_length'],5)
        self.assertEqual(len(plan['segments']),2)


if __name__=='__main__':unittest.main()
