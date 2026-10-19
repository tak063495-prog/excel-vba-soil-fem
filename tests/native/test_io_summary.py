"""Reject altered numerical/log evidence and unverified performance claims."""
import unittest
from summarize_io_benchmark import compare
from summarize_io_regression import compare as regression


class EvidenceChecks(unittest.TestCase):
    def setUp(self):
        self.reference = dict(Case='Case1', Variant='off', Mode='AUTO', Precision='AUTO',
            StateCopy=1, GmresSimd=1, IoFast=0, NativeDllSha256='abc', ExclusiveExcel=True,
            Fos='bracket', Bracket=[2.2, 2.2125, 2.20625, .0125, True, 17], Disp=[0., 1.],
            SolverEventsSha256='def', SolverEventsBytes=100, WallSeconds=75,
            Run='75|PASS|True|1e-13|7|8|9|10|statecopy=45600;gmres_simd=830;fallback=0')
        self.candidate = dict(self.reference, IoFast=1, WallSeconds=55)

    def test_exact_comparison_and_rejections(self):
        self.assertEqual(compare(self.reference, self.candidate)['max_abs_displacement_delta'], 0)
        for key, value in [('IoFast', 0), ('StateCopy', 0), ('GmresSimd', 0), ('ExclusiveExcel', False),
            ('NativeDllSha256', 'wrong'), ('Fos', 'wrong'), ('Disp', [0., 1.000000000001]),
            ('Bracket', [2.2, 2.3, 2.25, .1, True, 17]), ('SolverEventsSha256', 'wrong'),
            ('SolverEventsBytes', 101), ('Run', self.candidate['Run'].replace('1e-13', '2e-13')),
            ('Run', self.candidate['Run'].replace('PASS', 'FAIL')), ('WallSeconds', 76),
            ('WallSeconds', float('nan'))]:
            with self.subTest(key=key, value=value):
                changed = dict(self.candidate, **{key: value})
                with self.assertRaises(AssertionError):
                    compare(self.reference, changed)
        for key, value in [('Bracket', [2.2]), ('Disp', [float('nan')]),
                           ('SolverEventsSha256', ''), ('SolverEventsBytes', 0),
                           ('Run', self.reference['Run'].replace('fallback=0', 'fallback=1'))]:
            r = dict(self.reference, **{key: value}); c = dict(self.candidate, **{key: value})
            with self.subTest(key=key, value=value), self.assertRaises(AssertionError):
                compare(r, c)

    def test_event_scope(self):
        r = dict(self.reference, IoFast=1, EventFast=0)
        c = dict(self.candidate, IoFast=1, EventFast=1)
        self.assertEqual(compare(r, c, 'events')['max_abs_displacement_delta'], 0)
        self.assertTrue(regression([r], [c], 'events')[0]['logs_exact'])
        for key, value in [('IoFast', 0), ('EventFast', 0), ('Disp', [0., 1.1]), ('SolverEventsSha256', 'wrong')]:
            bad = dict(c, **{key: value})
            with self.subTest(key=key), self.assertRaises(AssertionError):
                compare(r, bad, 'events')
        with self.assertRaises(AssertionError):
            compare(r, c)

    def test_regression_accepts_noisy_time_but_rejects_changed_result(self):
        c = dict(self.candidate, WallSeconds=80)
        self.assertTrue(regression([self.reference], [c])[0]['logs_exact'])
        c['Disp'] = [0., 1.1]
        with self.assertRaises(AssertionError):
            regression([self.reference], [c])


if __name__ == '__main__':
    unittest.main()
