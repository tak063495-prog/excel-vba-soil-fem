"""Ensure performance evidence cannot silently accept mislabeled/invalid runs."""
import copy
import unittest
from summarize_inco_benchmark import compare

class EvidenceChecks(unittest.TestCase):
    def setUp(self):
        self.reference=dict(Case='Case1',Variant='off',Mode='AUTO',Precision='AUTO',NativeDllSha256='abc',
            StateCopy=0,GmresSimd=0,Fos='bracket',Bracket=[2.2,2.2125,2.20625,0.0125,True,17],ExclusiveExcel=True,
            Run='150|PASS|True|1e-13|7|8|9|10|statecopy=0;gmres_simd=0;fallback=0',Disp=[0.0,1.0],WallSeconds=150)
        self.candidate=copy.deepcopy(self.reference)
        self.candidate.update(StateCopy=1,GmresSimd=1,WallSeconds=70,
            Run=self.reference['Run'].replace('statecopy=0;gmres_simd=0','statecopy=45600;gmres_simd=830'))

    def test_valid_and_rejection_gates(self):
        self.assertTrue(compare(self.reference,self.candidate)['displacements_exact'])
        for key,value in [('Fos','wrong'),('ExclusiveExcel',False),('StateCopy',0),('GmresSimd',0),
            ('Disp',[float('nan'),1.0]),('Disp',[0.0,1.1]),('WallSeconds',160),('NativeDllSha256','wrong'),
            ('Run',self.candidate['Run'].replace('PASS','FAIL')),('Run',self.candidate['Run'].replace('1e-13','1e-3')),
            ('Bracket',[2.2,2.3,2.25,0.1,True,17])]:
            with self.subTest(key=key,value=value):
                changed=copy.deepcopy(self.candidate); changed[key]=value
                with self.assertRaises(AssertionError):compare(self.reference,changed)
        changed=copy.deepcopy(self.reference);changed['StateCopy']=1
        with self.assertRaises(AssertionError):compare(changed,self.candidate)

    def test_missing_bracket_requires_explicit_legacy_opt_in(self):
        candidate=copy.deepcopy(self.candidate);candidate.pop('Bracket')
        with self.assertRaises(AssertionError):compare(self.reference,candidate,True)
        reference=copy.deepcopy(self.reference);reference.pop('Bracket')
        with self.assertRaises(AssertionError):compare(reference,candidate)
        self.assertIsNone(compare(reference,candidate,True)['raw_bracket'])

if __name__=='__main__':unittest.main()
