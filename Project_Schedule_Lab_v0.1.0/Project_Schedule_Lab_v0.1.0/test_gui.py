"""Real Tk widgets; run under a valid DISPLAY, using a temporary database."""
import os
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch
import tkinter as tk
from tkinter import ttk
from excel_io import write_excel
from importing import prepare


@unittest.skipUnless(os.environ.get('DISPLAY'),'GUI needs DISPLAY (e.g. Xvfb)')
class GuiTests(unittest.TestCase):
    def setUp(self):
        import app as module
        self.tmp=tempfile.TemporaryDirectory()
        self.patch=patch.object(module,'DATA',Path(self.tmp.name));self.patch.start()
        self.app=module.App();self.errors=[]
        self.app.report_callback_exception=lambda *args:self.errors.append(args)
        self.app.update()
    def tearDown(self):
        for win in self.app.winfo_children():
            if isinstance(win,tk.Toplevel): win.destroy()
        self.app.close();self.patch.stop();self.tmp.cleanup()
        self.assertFalse(self.errors,self.errors)
    def descendants(self,root):
        result=[]
        for c in root.winfo_children(): result.append(c);result.extend(self.descendants(c))
        return result
    def button(self,root,text):
        return next(w for w in self.descendants(root) if isinstance(w,ttk.Button) and w.cget('text')==text)
    def test_recalculate_simulate_baseline_reopen(self):
        app=self.app;app.recalculate();self.assertEqual(app.result['n'],1)
        self.assertTrue(app.canvas.find_all());self.assertIn('총 여유',app.resulttext.get('1.0','end'))
        with patch('ui_v2.simpledialog.askstring',return_value='승인 테스트'): app.baseline()
        self.assertEqual(len(app.store.baselines(app.p['uid'])),1)
        app.runbtn.invoke();deadline=time.monotonic()+20
        while app.busy and time.monotonic()<deadline: app.update();time.sleep(.01)
        self.assertFalse(app.busy);self.assertEqual(app.result['n'],5000)
        self.assertEqual(len(app.store.baselines(app.p['uid'])),1)
        uid=app.p['uid'];app.store.db.close()
        from storage import Store
        app.store=Store(Path(self.tmp.name)/'history.sqlite3')
        self.assertEqual(app.store.load(uid),app.p)
        self.assertTrue(app.store.db.execute('SELECT count(*) FROM forecasts').fetchone()[0])

    def test_import_preview_apply_and_ui_editors(self):
        app=self.app;path=Path(self.tmp.name)/'import.xlsx'
        incoming=__import__('copy').deepcopy(app.p);incoming['tasks'][0]['name']='미리보기 변경'
        write_excel(path,incoming);plan=prepare(path,app.p,app.store)
        app.import_preview(plan);app.update()
        win=next(w for w in app.winfo_children() if isinstance(w,tk.Toplevel))
        self.button(win,'선택한 변경 적용').invoke();app.update()
        self.assertEqual(app.p['tasks'][0]['name'],'미리보기 변경')
        app.edit_task(0);app.update()
        edit=next(w for w in app.winfo_children() if isinstance(w,tk.Toplevel));self.assertTrue(self.button(edit,'저장'));edit.destroy()
        for method in (app.relationships_dialog,app.calendars_dialog,app.project_dialog,app.open_project):
            method();app.update()
            win=next(w for w in app.winfo_children() if isinstance(w,tk.Toplevel));win.destroy()
        app.view_baselines();app.update()
        for w in list(app.winfo_children()):
            if isinstance(w,tk.Toplevel): w.destroy()

    def test_probability_selection_moves_successors_and_relation_anchors(self):
        import numpy as np
        from engine import simulate,percentile
        from test_v2 import project,edge
        app=self.app
        p=project((5,2,2,0),[edge('A','B'),edge('A','C','SS',1),edge('A','D','FF',2)])
        p['tasks'][0].update(o=1,m=5,p=20)
        app.replace_project(p,'확률 표시 회귀 시험');app.result=simulate(app.p,5000,42)
        r=app.result;original={k:v.copy() for k,v in r['starts'].items()}
        np.testing.assert_array_equal(r['ends']['A'],r['starts']['B'])
        def center(tag):
            box=app.canvas.coords(tag);return (box[0]+box[2])/2
        def expected_x(value):
            from engine import day
            lo=min(day(app.p['start']),min(percentile(v,10) for v in r['starts'].values()))
            hi=max(lo+30,percentile(r['total'],95),day(app.p.get('target')) or lo)
            return 430+(value-lo)/(hi-lo)*(max(1250,app.canvas.winfo_width())-40-430)
        positions={}
        for q in (75,95):
            app.q.set(q);app.draw();app.update()
            positions[q]=center('start:B')
            self.assertAlmostEqual(center('end:A'),center('start:B'))
            for k in ('A','B','C','D'):
                self.assertAlmostEqual(center('start:'+k),expected_x(percentile(r['starts'][k],q)))
                self.assertAlmostEqual(center('end:'+k),expected_x(percentile(r['ends'][k],q)))
                self.assertIn('P'+str(q),app.canvas.itemcget('dates:'+k,'text'))
            for pred,succ,typ in [('A','B','FS'),('A','C','SS'),('A','D','FF')]:
                coords=app.canvas.coords(f'relation:{pred}:{succ}:{typ}')
                self.assertAlmostEqual(coords[0],center(('start:' if typ=='SS' else 'end:')+pred))
                self.assertAlmostEqual(coords[-2],center(('end:' if typ=='FF' else 'start:')+succ))
            self.assertTrue(app.canvas.find_withtag('milestone:D'))
        self.assertGreater(positions[95],positions[75])
        for k,v in original.items(): np.testing.assert_array_equal(r['starts'][k],v)

    def test_import_file_apply_stays_visible_and_switches_large_project(self):
        from model import canonical
        app=self.app;old_uid=app.p['uid'];path=Path(self.tmp.name)/'433_tasks.xlsx'
        p=canonical(dict(id='IMPORT-433',name='2027 신규 프로젝트',start='2027-01-01',asof='2027-01-01',
                         modules={'시험':'적용'},tasks=[dict(id=f'WBS-{i:03d}',name=f'작업 {i}',module='시험',o=1,m=2,p=4,pred=[]) for i in range(433)]))
        write_excel(path,p);app.search.set('기존 프로젝트 검색어')
        with patch('ui_v2.filedialog.askopenfilename',return_value=str(path)),patch('ui_v2.messagebox.askyesno',side_effect=AssertionError('Unexpected extra confirmation')):
            app.import_file()
        app.update();win=next(w for w in app.winfo_children() if isinstance(w,tk.Toplevel))
        button=self.button(win,'선택한 변경 적용')
        for geometry in ('1150x760','1000x650','680x460'):
            win.geometry(geometry);app.update()
            self.assertTrue(button.winfo_ismapped(),geometry)
            top=button.winfo_rooty()-win.winfo_rooty()
            self.assertGreaterEqual(top,0);self.assertLess(top+button.winfo_height(),win.winfo_height())
        self.assertEqual(app.p['uid'],old_uid)  # Reading/reviewing never applies early.
        button.invoke();app.update()
        self.assertEqual(app.p['id'],'IMPORT-433');self.assertEqual(len(app.p['tasks']),433)
        self.assertEqual(len(app.tree.get_children()),433);self.assertEqual(app.search.get(),'')
        self.assertEqual(app.tabs.select(),str(app.tasktab));self.assertIn('가져오기 완료',app.status.cget('text'))
        self.assertIn('433',app.caption.cget('text'))
        app.count.set('1000');app.runbtn.invoke();deadline=time.monotonic()+20
        while app.busy and time.monotonic()<deadline: app.update();time.sleep(.01)
        self.assertFalse(app.busy);self.assertEqual(len(app.result['starts']),433)
        self.assertEqual(app.result['project_uid'],app.p['uid']);self.assertEqual(app.result['asof'],'2027-01-01')

    def test_import_cancel_keeps_current_project_and_errors_disable_apply(self):
        app=self.app;old_uid=app.p['uid'];path=Path(self.tmp.name)/'test.xlsx'
        write_excel(path,app.p);plan=prepare(path,app.p,app.store)
        win=app.import_preview(plan);app.update();self.button(win,'취소').invoke();app.update()
        self.assertEqual(app.p['uid'],old_uid);self.assertIn('취소',app.status.cget('text'))
        from excel_io import Issue
        plan.issues.append(Issue('오류',str(path),'작업목록',2,'A','A','시험 오류','수정하세요'))
        win=app.import_preview(plan);app.update()
        button=self.button(win,'선택한 변경 적용');self.assertTrue(button.winfo_ismapped());self.assertTrue(button.instate(['disabled']))
        win.destroy()

    def test_project_clone_and_backup_restore_ui(self):
        app=self.app;old=app.p['uid'];backup=Path(self.tmp.name)/'backup.sqlite3'
        with patch('ui_v2.simpledialog.askstring',return_value='복제'): app.duplicate_project()
        self.assertNotEqual(app.p['uid'],old);self.assertEqual(len(app.store.projects()),2)
        with patch('ui_v2.filedialog.asksaveasfilename',return_value=str(backup)): app.backup_db()
        with patch('ui_v2.simpledialog.askstring',return_value='신규'): app.new_project()
        self.assertEqual(len(app.store.projects()),3)
        with patch('ui_v2.filedialog.askopenfilename',return_value=str(backup)),patch('ui_v2.messagebox.askyesno',return_value=True): app.restore_db()
        self.assertEqual(len(app.store.projects()),2)


if __name__=='__main__': unittest.main(verbosity=2)
