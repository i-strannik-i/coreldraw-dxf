import math
from pathlib import Path
import sys
import unittest
import ezdxf
sys.path.insert(0, str(Path(__file__).parent / 'engine'))
from contour_dxf import polyline_vertices, requires_closed


class ContourTests(unittest.TestCase):
    def test_closed_rectangle_is_one_closed_entity(self):
        p = [(0,0),(10,0),(10,20),(0,20),(0,0)]
        segs = [('L',(a,a,b,b)) for a,b in zip(p,p[1:])]
        v = polyline_vertices(segs, True, .05)
        d = ezdxf.new()
        e = d.modelspace().add_lwpolyline(v, format='xyb', close=True)
        self.assertTrue(e.closed)
        self.assertEqual(len(e),4)
        self.assertTrue(all(x.dxftype()=='LINE' for x in e.virtual_entities()))

    def test_arc_bulges_keep_direction(self):
        k = 4*(math.sqrt(2)-1)/3
        points = ((10,0),(10,10*k),(10*k,10),(0,10))
        v = polyline_vertices([('B',points)],False,.05)
        r = polyline_vertices([('B',tuple(reversed(points)))],False,.05)
        self.assertAlmostEqual(v[0][2],math.tan(math.pi/8),places=3)
        self.assertAlmostEqual(r[0][2],-v[0][2],places=8)
        self.assertEqual(v[-1][:2], points[-1])

    def test_open_stays_open_and_preserves_endpoints(self):
        p=((0,0),(10,20),(-10,20),(1,0))
        v=polyline_vertices([('B',p)],False,.05)
        self.assertEqual(v[0][:2],p[0])
        self.assertEqual(v[-1][:2],p[-1])
        self.assertGreater(len(v),2)

    def test_gap_is_not_automatically_closed(self):
        with self.assertRaises(ValueError):
            polyline_vertices([('L',((0,0),(0,0),(1,0),(1,0)))],True,.05)

    def test_process_layers(self):
        for name in ('CUT_OUT','CUT_IN','D4','BACK_CUT_OUT','P_2.5'):
            self.assertTrue(requires_closed(name))
        for name in ('CUT_ON','V_3','INFO','#','B'):
            self.assertFalse(requires_closed(name))
