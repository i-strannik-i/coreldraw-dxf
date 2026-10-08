import unittest
from corel_export.ui_shell import THEMES, clean_log


class BetaUiTests(unittest.TestCase):
    def test_logs_remove_sensitive_lines(self):
        value='Готово\nC:\\Users\\Designer\\secret.cdr\nPassword=abc\nuser@example.com\nрезка деталь.dxf\nСлой D4: 12'
        result=clean_log(value)
        for secret in ('Designer','abc','example.com','деталь.dxf'):
            self.assertNotIn(secret,result)
        self.assertIn('Слой D4: 12',result)

    def test_issue_colors_readable_in_both_themes(self):
        def luminance(color):
            rgb=[int(color[i:i+2],16)/255 for i in (1,3,5)]
            return sum((v/12.92 if v<=.04045 else ((v+.055)/1.055)**2.4)*w for v,w in zip(rgb,(.2126,.7152,.0722)))
        for p in THEMES.values():
            colors=[p[k] for k in ('zero','overlap','intersection','open')]
            self.assertEqual(len(set(colors)),4)
            for color in colors:
                a,b=sorted((luminance(p['panel']),luminance(color)))
                self.assertGreaterEqual((b+.05)/(a+.05),4.5)

