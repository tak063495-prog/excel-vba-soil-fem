"""Issue 64: validators must reject every non-finite sample before reduction."""
import json
from pathlib import Path
import tempfile
import unittest

from check_fixed_stress_history import check as history_check
from check_fixed_stress_water_ramp import check as ramp_check


class NonfiniteRegression(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.actual = self.root/'actual'; self.actual.mkdir()
        self.expected = self.root/'expected'; self.expected.mkdir()
        self.record = {'status':'PASS|SYNTHETIC', 'nodes':[[1,.001,.002]],
                       'pressure':[10]*4, 'stress':[0]*17}
        self.write_history()

    def write_history(self, field=None, value=None):
        (self.actual/'continued.json').write_text(json.dumps(self.record), encoding='utf-8')
        (self.root/'result.json').write_text(json.dumps({'stress':[0]*17}), encoding='utf-8')
        (self.expected/'node.csv').write_text(f'node,ux,uy\n1,{value if field=="ux" else .001},.002\n', encoding='utf-8')
        (self.expected/'water.csv').write_text('element,gauss,pressure_kPa\n'+''.join(
            f'1,{i+1},{value if field==i else 10}\n' for i in range(4)), encoding='utf-8')

    def history(self): return history_check(self.actual, self.expected, self.root/'result.json')

    def test_history_finite(self):
        self.assertEqual(self.history()['status'], 'PASS')
        self.write_history(0, 11)
        with self.assertRaises(ValueError): self.history()

    def test_history_nonfinite_csv(self):
        for value in ('NaN', 'Infinity', '-Infinity'):
            for field in ('ux',0,1,3):
                with self.subTest(value=value, field=field):
                    self.write_history(field, value)
                    with self.assertRaisesRegex(ValueError, 'non-finite'): self.history()

    def test_history_nonfinite_json(self):
        for key in ('pressure','stress'):
            for value in (float('nan'),float('inf'),-float('inf')):
                self.record[key][-1] = value
                self.write_history()
                with self.assertRaisesRegex(ValueError, 'non-finite'): self.history()
            self.record[key][-1] = 0 if key=='stress' else 10

    def ramp(self, value=None, frame=2):
        run=self.root/'run'; run.mkdir(exist_ok=True)
        for i in range(1,5):
            folder=run/f'C{i:06d}'; folder.mkdir(exist_ok=True)
            (folder/'fixed_stress.csv').write_text(f'STATE,{i*10},0,0,0,0\nPHASE,3,{i*10},40,10,1\n',encoding='utf-8')
            load=-(40+9.81*(2+i/4)) if value is None or i!=frame else value
            (folder/'dof.csv').write_text(f'dof,locked_sw\n0,0\n1,{load}\n',encoding='utf-8')
        (run/'result.json').write_text(json.dumps({'metrics':[0,0,69.43]}),encoding='utf-8')
        return ramp_check(run)

    def test_ramp(self):
        self.assertEqual(self.ramp()['status'],'PASS')
        with self.assertRaises(ValueError): self.ramp(1)
        for value in ('NaN','Infinity','-Infinity'):
            for frame in (1,2,4):
                with self.subTest(value=value, frame=frame):
                    with self.assertRaisesRegex(ValueError,'non-finite'): self.ramp(value,frame)


if __name__ == '__main__': unittest.main()
