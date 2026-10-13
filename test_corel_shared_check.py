import copy
import unittest
from unittest.mock import patch
from corel_export.vector_tools import CorelSession
from corel_export.check_pipeline import SharedCheck


class Session:
    def __init__(self):
        self.source=[dict(id=1,layer='CUT',subpath=1,closed=False,segments=[[(0,0),(1,0)]])]
        self.skipped=[]
        self.moved=0

    def snapshot(self, selection, reading=None, layers=None):
        self.contours=copy.deepcopy(self.source)
        self.signature=CorelSession.digest(self.contours)
        return self.contours,self.skipped

    def quarantine_duplicates(self, cancelled,progress=None):
        return self.moved


class SharedCheckTests(unittest.TestCase):
    def test_export_and_inspection_reuse_same_audit(self):
        session=Session();pipeline=SharedCheck(session)
        with patch('corel_export.check_pipeline.audit',return_value={'complete':True,'issues':[]}) as audit:
            first=pipeline.run(.1,False,{'CUT'},lambda:False)
            second=pipeline.run(.1,False,None,lambda:False)
        self.assertFalse(first['reused']);self.assertTrue(second['reused'])
        audit.assert_called_once()

    def test_geometry_addition_parameters_scope_and_skips_invalidate(self):
        for change in ('geometry','addition','tolerance','selection','layers','skipped'):
            with self.subTest(change=change):
                session=Session();pipeline=SharedCheck(session)
                with patch('corel_export.check_pipeline.audit',return_value={'complete':True,'issues':[]}) as audit:
                    pipeline.run(.1,False,{'CUT'},lambda:False)
                    if change=='geometry':session.source[0]['closed']=True
                    if change=='addition':session.source.append(dict(session.source[0],id=2))
                    if change=='skipped':session.skipped=[3]
                    result=pipeline.run(.2 if change=='tolerance' else .1,change=='selection',
                                        {'OTHER'} if change=='layers' else {'CUT'},lambda:False)
                self.assertFalse(result['reused']);self.assertEqual(audit.call_count,2)

    def test_incomplete_and_cancelled_checks_are_not_reused(self):
        session=Session();pipeline=SharedCheck(session)
        with patch('corel_export.check_pipeline.audit',return_value={'complete':False,'issues':[]}) as audit:
            for _ in range(2):pipeline.run(.1,False,None,lambda:False)
            self.assertEqual(audit.call_count,2)
            with self.assertRaises(InterruptedError):pipeline.run(.1,False,None,lambda:True)
