import unittest
from summarize_element import compare
from test_ilumat_summary import Checks as PreviousChecks

class Checks(unittest.TestCase):
    def setUp(self):
        old=PreviousChecks();old.setUp()
        self.r=dict(old.r,GmresPrepared=1,PreparedCalls=830,GmresEventRows=98851)
        self.c=dict(self.r,MaterialDllSha256='new',WallSeconds=43.)

    def test_exact_output_and_metadata_guards(self):
        self.assertTrue(compare(self.r,self.c)['logs_exact'])
        for key,value in [('Disp',[0.,1.000000000001]),('SolverEventsSha256','changed'),
                          ('SolverEventsBytes',0),('GmresEventRows',0),('PreparedCalls',0),
                          ('GmresPrepared',0),('MaterialDllSha256','old'),('ExclusiveExcel',False),
                          ('WallSeconds',float('nan')),('Run',self.c['Run'].replace('1e-13','2e-13'))]:
            with self.subTest(key=key),self.assertRaises(AssertionError):
                compare(self.r,dict(self.c,**{key:value}))

    def test_slower_results_are_retained(self):
        self.assertLess(compare(self.r,dict(self.c,WallSeconds=46.))['reduction_percent'],0)

if __name__=='__main__':unittest.main()
