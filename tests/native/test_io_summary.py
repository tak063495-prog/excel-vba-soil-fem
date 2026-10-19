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

    def test_assembly_scope(self):
        r = dict(self.reference, IoFast=1, EventFast=1, AssemblyFast=0, Assembly=[0, 1024, .01],Flow='INCONSISTENT')
        c = dict(self.candidate, IoFast=1, EventFast=1, AssemblyFast=1, Assembly=[7804, 1024, .01],Flow='INCONSISTENT')
        self.assertEqual(compare(r, c, 'assembly')['max_abs_displacement_delta'], 0)
        self.assertTrue(regression([r], [c], 'assembly')[0]['logs_exact'])
        for key,value in [('IoFast',0),('EventFast',0),('AssemblyFast',0),
                          ('Assembly',[0,1024,.01]),('Assembly',[7804,1025,.01]),
                          ('Assembly',[7804,1024,.02]),('Disp',[0.,1.1])]:
            with self.subTest(key=key,value=value), self.assertRaises(AssertionError):
                compare(r,dict(c,**{key:value}),'assembly')
        self.assertTrue(regression([r],[dict(c,WallSeconds=90)],'assembly')[0]['logs_exact'])
        with self.assertRaises(AssertionError):
            regression([r],[dict(c,Assembly=[0,1024,.01])],'assembly')
        self.assertTrue(regression([dict(r,Flow='DAVIS')],[dict(c,Flow='DAVIS',Assembly=[0,1024,.01])],'assembly'))

    def test_regression_accepts_noisy_time_but_rejects_changed_result(self):
        c = dict(self.candidate, WallSeconds=80)
        self.assertTrue(regression([self.reference], [c])[0]['logs_exact'])
        c['Disp'] = [0., 1.1]
        with self.assertRaises(AssertionError):
            regression([self.reference], [c])

    def test_gmres_scope_and_real_path_gate(self):
        r=dict(self.reference,IoFast=1,EventFast=1,AssemblyFast=1,Assembly=[7804,5120,.01],
               Flow='INCONSISTENT',GmresTuned=0,TunedCalls=0,DecisionEventsSha256='decision',NonGmresSolverEventsRows=100)
        c=dict(r,GmresTuned=1,TunedCalls=830,WallSeconds=55)
        self.assertEqual(compare(r,c,'gmres')['max_abs_displacement_delta'],0)
        self.assertTrue(regression([r],[c],'gmres')[0]['logs_exact'])
        for key,value in [('GmresTuned',0),('TunedCalls',0),('AssemblyFast',0),('Assembly',[7805,5120,.01]),('Disp',[0.,1.1]),('DecisionEventsSha256','changed'),('NonGmresSolverEventsRows',99)]:
            with self.subTest(key=key),self.assertRaises(AssertionError):
                compare(r,dict(c,**{key:value}),'gmres')
        with self.assertRaises(AssertionError):
            regression([r],[dict(c,TunedCalls=0)],'gmres')
        r=dict(r,Run=r['Run'].replace('gmres_simd=830','gmres_simd=0'))
        c=dict(c,Run=r['Run'],TunedCalls=0)
        self.assertTrue(regression([r],[c],'gmres')[0]['logs_exact'])
        self.assertTrue(regression([dict(r,NonGmresSolverEventsRows=0)],[dict(c,NonGmresSolverEventsRows=0)],'gmres')[0]['logs_exact'])
        with self.assertRaises(AssertionError):
            regression([dict(r,NonGmresSolverEventsRows=-1)],[dict(c,NonGmresSolverEventsRows=-1)],'gmres')
        changed=dict(c,SolverEventsSha256='different',SolverEventsBytes=110)
        self.assertFalse(regression([r],[changed],'gmres')[0]['logs_exact'])


if __name__ == '__main__':
    unittest.main()
