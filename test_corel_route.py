import unittest
from engine.route_order import optimize_route,route_length,join_summary


def line(a,b):
    return dict(closed=False,segments=[('L',[a,a,b,b])])


class RouteTests(unittest.TestCase):
    def test_shorter_order_preserves_direction(self):
        curves=[line((0,0),(1,0)),line((100,0),(101,0)),line((2,0),(3,0))]
        order,before,after=optimize_route(curves)
        self.assertEqual(order,[0,2,1])
        self.assertEqual((before,after),(198,98))

    def test_inner_before_outer(self):
        outer=dict(closed=True,segments=[('L',[(0,0),(0,0),(10,10),(10,10)]),('L',[(10,10),(10,10),(0,0),(0,0)])])
        curves=[outer,line((2,2),(3,3))]
        order,_,_=optimize_route(curves)
        self.assertEqual(order,[1,0])

    def test_zero_one_and_equal_bounds(self):
        self.assertEqual(optimize_route([]),([],0,0))
        curves=[line((0,0),(1,0))]*3
        order,before,after=optimize_route(curves)
        self.assertEqual(sorted(order),[0,1,2])
        self.assertLessEqual(after,before)
        self.assertEqual(after,route_length(curves,order))

    def test_join_is_not_closure(self):
        text=join_summary((0,2),(0,1),.1)
        self.assertIn('Соединений: 1. Замыканий: 0.',text)
        self.assertIn('открытых — 1',text)

    def test_working_drawing_distance_summary(self):
        text=join_summary((351,554),(351,554),.1,(257972.7335,134393.2418))
        self.assertIn('258,0 → 134,4 м',text)
        self.assertIn('Сокращение: 48%',text)
        self.assertIn('Сокращения пути нет',join_summary((0,1),(0,1),.1,(0,0)))


if __name__=='__main__':unittest.main()
