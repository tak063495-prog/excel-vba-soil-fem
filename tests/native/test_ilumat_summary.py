import copy,unittest
from summarize_ilumat import compare

class Checks(unittest.TestCase):
    def setUp(self):
        self.r=dict(Case='Case1',Variant='off',Mode='AUTO',Precision='AUTO',Flow='INCONSISTENT',NativeDllSha256='native',MaterialDllSha256='old',
            Assembly=[7804,5120,.01],Fos='same',Bracket=[2.2,2.2125,2.20625,.0125,True,17],Disp=[0.,1.],ExclusiveExcel=True,
            StateCopy=1,GmresSimd=1,IoFast=1,EventFast=1,AssemblyFast=1,GmresTuned=1,GmresPrepared=0,PreparedCalls=0,TunedCalls=830,
            Run='45|PASS|True|1e-13|7|6713|4078|76|gmres_simd=830;fallback=0||||||material_same',
            SolverEventsSha256='oldlog',SolverEventsBytes=13000000,DecisionEventsSha256='same',NonGmresSolverEventsRows=10869,WallSeconds=45.)
        self.c=dict(self.r,MaterialDllSha256='new',GmresPrepared=1,PreparedCalls=830,SolverEventsSha256='newlog',WallSeconds=43.)
    def test_primary_and_path_guards(self):
        self.assertEqual(compare(self.r,self.c)['max_abs_displacement_delta'],0.)
        for key,value in [('Disp',[0.,1.000000000001]),('PreparedCalls',0),('GmresPrepared',0),('MaterialDllSha256','old'),
            ('NativeDllSha256','different'),('DecisionEventsSha256','different'),('Assembly',[7805,5120,.01]),('ExclusiveExcel',False),
            ('Run',self.c['Run'].replace('1e-13','2e-13')),('Run',self.c['Run'].replace('material_same','material_changed'))]:
            with self.subTest(key=key),self.assertRaises(AssertionError):compare(self.r,dict(self.c,**{key:value}))
    def test_time_is_reported_even_if_a_pair_is_slower(self):
        self.assertLess(compare(self.r,dict(self.c,WallSeconds=46.))['reduction_percent'],0)
        with self.assertRaises(AssertionError):compare(self.r,dict(self.c,WallSeconds=float('nan')))
    def test_fixed_fs_requires_all_event_bytes(self):
        r=dict(self.r,Run=self.r['Run'].replace('gmres_simd=830','gmres_simd=0'),TunedCalls=0,NonGmresSolverEventsRows=0)
        c=dict(self.c,Run=r['Run'],TunedCalls=0,PreparedCalls=0,NonGmresSolverEventsRows=0,SolverEventsSha256=r['SolverEventsSha256'])
        self.assertTrue(compare(r,c,False)['logs_exact'])
        with self.assertRaises(AssertionError):compare(r,dict(c,SolverEventsSha256='other'),False)

if __name__=='__main__':unittest.main()
