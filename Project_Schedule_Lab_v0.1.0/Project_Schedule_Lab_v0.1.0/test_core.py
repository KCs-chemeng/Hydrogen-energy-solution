import unittest,tempfile
from pathlib import Path
import numpy as np
from engine import *
from excel_io import read_excel,write_excel
class Tests(unittest.TestCase):
 def base(self):
  p=demo();p['tasks']=p['tasks'][:2]
  for t in p['tasks']:t['o']=t['m']=t['p']=2
  return p
 def test_serial_parallel(self):
  p=self.base();self.assertEqual(percentile(simulate(p,100)['total'],80),day(p['start'])+4);p['tasks'][1]['pred']=[];self.assertEqual(percentile(simulate(p,100)['total'],80),day(p['start'])+2)
 def test_exclusion(self):
  p=self.base();p['tasks'][0]['state']='제외';self.assertEqual(percentile(simulate(p,100)['total'],80),day(p['start'])+2)
 def test_cycle_missing(self):
  p=self.base();p['tasks'][0]['pred']=['A02'];self.assertTrue(validate(p)[0]);p['tasks'][0]['pred']=['unknown'];self.assertTrue(validate(p)[0])
 def test_actual_remaining(self):
  p=self.base();p['asof']='2026-10-05';p['tasks'][0].update(state='완료',actualStart='2026-10-01',actualEnd='2026-10-03');p['tasks'][1].update(state='진행',actualStart='2026-10-03',ro=1,rm=1,rp=1);self.assertEqual(percentile(simulate(p,100)['total'],80),day('2026-10-06'))
 def test_workday_holiday(self):
  p=self.base();p['tasks']=p['tasks'][:1];p['start']=p['asof']='2026-10-02';p['holidays']=['2026-10-05'];p['tasks'][0].update(o=1,m=1,p=1,calendar='근무일');self.assertEqual(percentile(simulate(p,100)['total'],80),day('2026-10-06'));p['tasks'][0].update(o=0,m=0,p=0);p['start']=p['asof']='2026-10-03';self.assertEqual(percentile(simulate(p,100)['total'],80),day('2026-10-03'))
 def test_repeatability(self):
  p=demo();a=simulate(p,1000,42);b=simulate(p,1000,42);np.testing.assert_equal(a['total'],b['total']);self.assertGreater(percentile(a['total'],90),percentile(a['total'],50))
 def test_common_roundtrip(self):
  p=read_excel(Path(__file__).with_name('common_Rev1.xlsx'));self.assertEqual(len(p['tasks']),243);self.assertEqual(len(p['modules']),29);p['start']=p['asof']='2026-10-01';self.assertTrue(validate(p)[0]);p['modules']={m:'적용' for m in p['modules']};r=simulate(p,5000)
  with tempfile.TemporaryDirectory() as td:
   path=Path(td)/'roundtrip.xlsx';write_excel(path,p,r);self.assertEqual(p,read_excel(path))
 def test_store(self):
  from app import Store
  with tempfile.TemporaryDirectory() as td:
   s=Store(Path(td)/'test.sqlite');p=demo();s.add(p,'분석','test',summary(simulate(p,10),p));self.assertEqual(s.get(s.rows()[0][0])['project'],p);s.db.close()
if __name__=='__main__':unittest.main(verbosity=2)
