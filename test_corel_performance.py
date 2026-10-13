import unittest
from unittest.mock import Mock
from corel_export.corel_performance import paused_redraw


class RedrawTests(unittest.TestCase):
    def test_redraw_is_restored_on_success_and_error(self):
        for previous in (False,True):
            for fail in (False,True):
                with self.subTest(previous=previous,fail=fail):
                    app=Mock(Optimization=previous)
                    try:
                        with paused_redraw(app):
                            self.assertTrue(app.Optimization)
                            if fail:raise InterruptedError()
                    except InterruptedError:pass
                    self.assertEqual(app.Optimization,previous)
                    self.assertEqual(app.Refresh.call_count,int(not previous))
