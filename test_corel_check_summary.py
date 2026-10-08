import unittest
from corel_export.vector_tools import check_summary


class CheckSummaryTests(unittest.TestCase):
    def summary(self, complete=True, count=11, skipped=(), issues=()):
        return check_summary({'complete':complete,'issues':issues},count,skipped,.1,True,{'OUT'})

    def test_success_includes_scope(self):
        text=self.summary()
        for value in ('Проверка завершена','Ошибок и предупреждений не найдено',
                      'Только выделенные','слои: OUT','Контуров: 11','0.1 мм'):
            self.assertIn(value,text)

    def test_incomplete_empty_or_skipped_never_claim_success(self):
        for kwargs in ({'complete':False},{'count':0},{'skipped':['unsupported']}):
            self.assertNotIn('Ошибок и предупреждений не найдено',self.summary(**kwargs))

    def test_issues_reported(self):
        self.assertIn('Найдено проблем: 1',self.summary(issues=[{'kind':'open'}]))
