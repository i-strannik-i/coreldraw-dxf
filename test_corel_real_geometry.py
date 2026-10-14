"""Regression geometry extracted from the user's failing test drawing."""
import json
from pathlib import Path
import unittest
import sys
sys.path.insert(0,str(Path(__file__).parent/'engine'))
from duplicate_geometry import duplicate_groups, duplicate_removals
from vector_audit import audit


class RealGeometryTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.data=json.loads((Path(__file__).parent/'test_data/corel_geometry_regression.json').read_text(encoding='utf-8'))

    def test_repeated_hole_inside_compound_is_duplicate_not_overlap(self):
        contours=self.data['compound']
        result=audit(contours)
        self.assertEqual([i['kind'] for i in result['issues']],['duplicate'])
        group=result['issues'][0]
        self.assertEqual(group['mode'],'subpaths')
        self.assertEqual(len(group['copies']),1)
        whole,partial=duplicate_removals([group])
        self.assertFalse(whole)
        self.assertEqual(partial,{3682:{1}})
        self.assertFalse(audit([contours[0],contours[2]])['issues'])

    def test_identical_vectors_on_independent_layers_are_not_duplicates(self):
        contours=self.data['independent_layers']
        self.assertEqual({c['layer'] for c in contours},{'BACK_D4','OUT'})
        self.assertFalse(duplicate_groups(contours))
        self.assertFalse(audit(contours)['issues'])

    def test_actual_same_layer_crossing_is_not_hidden(self):
        result=audit(self.data['actual_crossing'])
        self.assertTrue(any(i['kind']=='intersection' and
                            abs(i['point'][0]+9491.67553259)<.1 and
                            abs(i['point'][1]+9365.99516269)<.1
                            for i in result['issues']))
