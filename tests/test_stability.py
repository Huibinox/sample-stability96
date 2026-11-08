import sys
import tempfile
import threading
import unittest
from pathlib import Path
from datetime import datetime, timedelta
from openpyxl import Workbook
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from stability96 import SharedStore, read_excel, replay, category, collection, deadline

def workbook(path, rows):
    wb=Workbook(); ws=wb.active
    ws.append(['样品编号','BAS课题号','Collection Date and Time','Collection Date','Collection Time'])
    for row in rows: ws.append(row)
    wb.save(path); wb.close()

class CoreTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory(); self.root=Path(self.temp.name)
        self.store=SharedStore(self.root)
    def tearDown(self): self.temp.cleanup()
    def sample(self): return replay(self.store.read_events())[0]['260770.00274']
    def seed(self):
        path=self.root/'samples.xlsx'
        workbook(path,[['260770.00274','P001','2026-11-07 09:59',None,None]])
        self.store.import_file(path,'Alice'); return path

    def test_combined_split_and_missing(self):
        self.assertEqual(collection(None,'2026-11-07','09:59')[0],'2026-11-07 09:59:00')
        self.assertEqual(collection('2026-11-07 09:59',None,None)[0],'2026-11-07 09:59:00')
        self.assertEqual(collection('2026-11-07',None,None)[0],'')
        self.assertTrue(collection('2026-11-07 09:59','2026-11-07','10:01')[2])
        self.assertEqual(collection(None,'2026-11-07',0)[0],'2026-11-07 00:00:00')

    def test_excel_date_only_does_not_become_midnight(self):
        path=self.root/'date.xlsx'; wb=Workbook(); ws=wb.active
        ws.append(['样品编号','课题号','Collection Date and Time'])
        ws.append(['A','P',datetime(2026,11,7)])
        ws['C2'].number_format='yyyy-mm-dd'; wb.save(path); wb.close()
        rows,_=read_excel(path); self.assertEqual(rows[0]['collected'],'')
        self.assertEqual(rows[0]['date'],'2026-11-07')

    def test_96h_boundary_and_done(self):
        self.seed(); s=self.sample(); end=datetime(2026,11,11,9,59)
        self.assertEqual(deadline(s),end)
        self.assertEqual(category(s,end-timedelta(seconds=1)),'active')
        self.assertEqual(category(s,end),'expired')
        self.store.action('complete',s['bas'],'Bob','')
        self.assertEqual(category(self.sample(),end),'done')

    def test_dedup_completion_and_undo(self):
        path=self.seed(); self.assertIsNone(self.store.import_file(path,'Bob'))
        s=self.sample(); self.store.action('complete',s['bas'],'Alice','')
        path2=self.root/'new.xlsx'; workbook(path2,[['260770.00274','P001','2026-11-07 09:59',None,None],['B','P001',None,'2026-11-08',None]])
        self.store.import_file(path2,'Bob')
        self.assertTrue(self.sample()['done']); self.assertEqual(self.sample()['completed_by'],'Alice')
        self.store.action('undo',s['bas'],'Bob',self.sample()['status_rev'])
        self.assertFalse(self.sample()['done'])
        self.assertEqual(len(replay(self.store.read_events())[0]),2)

    def test_conflict_and_manual_fix_survive_import(self):
        self.seed(); path=self.root/'conflict.xlsx'
        workbook(path,[['260770.00274','P002','2026-11-08 08:00',None,None]])
        self.store.import_file(path,'Bob'); self.assertTrue(self.sample()['conflicts'])
        self.store.action('edit','260770.00274','Alice','',dict(project='P003',collected='2026-11-08 09:00:00'))
        self.assertFalse(self.sample()['conflicts'])
        other=self.root/'other.xlsx'
        workbook(other,[['260770.00274','P002',None,'2026-11-09','08:00']])
        self.store.import_file(other,'Bob'); s=self.sample()
        self.assertEqual(s['project'],'P003'); self.assertEqual(s['collected'],'2026-11-08 09:00:00')
        self.assertFalse(s['conflicts'])

    def test_concurrent_completion_and_stale_undo(self):
        self.seed(); outcomes=[]; barrier=threading.Barrier(2)
        def complete(actor):
            barrier.wait()
            try: SharedStore(self.root).action('complete','260770.00274',actor,''); outcomes.append('ok')
            except ValueError: outcomes.append('stale')
        ts=[threading.Thread(target=complete,args=(name,)) for name in ('Alice','Bob')]
        for t in ts:t.start()
        for t in ts:t.join()
        self.assertCountEqual(outcomes,['ok','stale'])
        first=self.sample()['status_rev']
        self.store.action('undo','260770.00274','Alice',first)
        self.store.action('complete','260770.00274','Bob',self.sample()['status_rev'])
        with self.assertRaises(ValueError): self.store.action('undo','260770.00274','Alice',first)

    def test_two_readers_concurrent_import(self):
        path=self.seed(); before=len(self.store.read_events())
        threads=[threading.Thread(target=lambda:SharedStore(self.root).import_file(path,'Bob')) for _ in range(2)]
        for t in threads:t.start()
        for t in threads:t.join()
        self.assertEqual(len(self.store.read_events()),before)

    def test_bad_file_no_event_and_missing_share(self):
        path=self.root/'bad.xlsx'; path.write_text('bad')
        with self.assertRaises(Exception):self.store.import_file(path,'Alice')
        self.assertEqual(self.store.read_events(),[])
        with self.assertRaises(OSError):SharedStore(self.root/'absent').read_events()

if __name__=='__main__':unittest.main()
