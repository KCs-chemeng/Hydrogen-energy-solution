"""Small workflow additions to the existing Tkinter screens."""
import json
import threading
import queue
from copy import deepcopy
from datetime import date
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from model import canonical, clone_project, uid, sync_pred, TASK_DEFAULTS
from engine import deterministic, inspect, exclusion_impact, textday, summary, day
from excel_io import read_document, read_excel, COLS, DATES, NUMBERS, REL_COLS
from importing import prepare


def result_record(r,p):
    out=summary(r,p)
    from engine import percentile
    out.update(starts={k:textday(percentile(v,50)) for k,v in r['starts'].items()},
               ends={k:textday(percentile(v,50)) for k,v in r['ends'].items()},
               warnings=r.get('warnings',[]),float=r.get('float',{}),critical=r.get('critical',[]),
               target_delta=r.get('target_delta'),project_uid=p['uid'])
    return out


class Workflows:
    def build_v2(self):
        menu=tk.Menu(self)
        project=tk.Menu(menu,tearoff=False)
        for label,cmd in [('새 프로젝트',self.new_project),('프로젝트 열기',self.open_project),('저장',self.save_project),('프로젝트 복제',self.duplicate_project),('DB 백업',self.backup_db),('DB 복원',self.restore_db)]:
            project.add_command(label=label,command=lambda c=cmd:self.guarded(c))
        menu.add_cascade(label='프로젝트',menu=project)
        tools=tk.Menu(menu,tearoff=False)
        for label,cmd in [('모듈 추가 · 적용',self.modules_dialog),('관계 유형 · 시차',self.relationships_dialog),('달력 · 추가 근무일',self.calendars_dialog),('입력 오류 · 충돌 확인',self.check_input),('결정론적 재계산 (M)',self.recalculate),('기준 일정 버전 조회',self.view_baselines)]:
            tools.add_command(label=label,command=lambda c=cmd:self.guarded(c))
        menu.add_cascade(label='일정 · 데이터',menu=tools);self.configure(menu=menu)
        self.resulttab=ttk.Frame(self.tabs);self.tabs.add(self.resulttab,text='예측 · 임계 · 오류')
        bar=ttk.Frame(self.resulttab);bar.pack(fill='x')
        for text,cmd in [('M 기준 재계산',self.recalculate),('오류·경고 확인',self.check_input),('관계 편집',self.relationships_dialog),('달력 편집',self.calendars_dialog),('기준 일정 버전',self.view_baselines)]:
            ttk.Button(bar,text=text,command=lambda c=cmd:self.guarded(c)).pack(side='left',padx=4,pady=6)
        self.resulttext=tk.Text(self.resulttab,wrap='none',font=('TkFixedFont',10));self.resulttext.pack(fill='both',expand=True)
        self.bind_all('<Control-s>',lambda e:self.guarded(self.save_project))

    def report_callback_exception(self,exc_type,exc,tb):
        import traceback
        if hasattr(self,'store'):
            (self.store.path.parent/'error.log').write_text(''.join(traceback.format_exception(exc_type,exc,tb)),encoding='utf-8')
        messagebox.showerror('처리 실패',str(exc),parent=self)

    def guarded(self,cmd):
        if self.busy: messagebox.showinfo('계산 중','계산 완료 또는 중지 후 변경하세요.',parent=self);return
        try: return cmd()
        except Exception as exc: messagebox.showerror('처리 오류',str(exc),parent=self)

    def show_text(self,title,text):
        win=tk.Toplevel(self);win.title(title);win.geometry('1000x650')
        t=tk.Text(win,wrap='word',padx=12,pady=12);t.pack(fill='both',expand=True);t.insert('1.0',text);t.configure(state='disabled')
        return win

    def changed(self,kind,reason):
        try:
            self.store.add(self.p,kind,reason)
        except Exception:
            old=self.store.load(self.p.get('uid'))
            if old: self.p=old
            self.result=None
            self.after_idle(self.sync)
            raise
        self.result=None;self.metrics_update();self.fill_history();self.draw()
        if hasattr(self,'resulttext'): self.resulttext.delete('1.0','end')
        self.note('저장 완료 · 예측은 재계산하세요. 기준 일정 버전은 보존됩니다.')

    def save_project(self):
        self.apply_dates();self.store.add(self.p,'저장','사용자 저장');self.note('프로젝트 저장 완료 (변경 없는 저장은 이력 추가 없음)')

    def fill_history(self):
        self.hist.delete(*self.hist.get_children())
        for row in self.store.rows(self.p.get('uid')): self.hist.insert('','end',iid=str(row[0]),values=row)

    def new_project(self):
        name=simpledialog.askstring('새 프로젝트','프로젝트명:',parent=self)
        if not name: return
        today=date.today().isoformat()
        p=canonical(dict(id='PJ-'+uid()[:8],name=name,start=today,asof=today,target='',modules={'공통':'적용'},tasks=[]))
        self.replace_project(p,'새 프로젝트')

    def open_project(self):
        rows=self.store.projects()
        win=tk.Toplevel(self);win.title('프로젝트 열기');win.geometry('720x400')
        tree=ttk.Treeview(win,columns=('id','name','version'),show='headings')
        for k,label in [('id','프로젝트 ID'),('name','프로젝트명'),('version','데이터 버전')]: tree.heading(k,text=label)
        tree.pack(fill='both',expand=True)
        for k,pid,name,v in rows: tree.insert('','end',iid=k,values=(pid,name,v))
        def choose():
            if self.busy: return
            s=tree.selection()
            if s: self.replace_project(self.store.load(s[0]),'저장 프로젝트 열기');win.destroy()
        ttk.Button(win,text='열기',command=choose).pack();tree.bind('<Double-1>',lambda e:choose())

    def duplicate_project(self):
        name=simpledialog.askstring('프로젝트 복제','복제 프로젝트명:',parent=self)
        if name:
            self.save_project();p=clone_project(self.p);p['name']=name
            self.replace_project(p,'프로젝트 복제 (내부·작업 ID 재생성)')

    def replace_project(self,p,reason):
        if self.busy: return
        self.store.add(self.p,'저장','프로젝트 전환 전 저장')
        new=canonical(p);self.store.add(new,'프로젝트 열기',reason)
        self.p=new;self.result=None;self.source_path=None;self.sync();self.resulttext.delete('1.0','end');self.note(reason)

    def load_demo(self):
        from engine import demo
        if not self.busy:
            p=clone_project(demo());self.replace_project(p,'새 예제 프로젝트 (설명용 가정 기간)')

    def common(self):
        if self.busy: return
        path=Path(__file__).with_name('common_Rev1.xlsx')
        doc=read_document(path)
        # Open a separate draft rather than importing blank dates over an existing project.
        p=canonical(doc.project)
        p['uid']=uid();p['id']='PJ-'+uid()[:8]
        self.replace_project(p,'공통 WBS를 별도 초안으로 열었습니다. 날짜와 모듈 적용성을 입력하세요.')
        self.source_path=path.resolve();self.check_input()

    def import_file(self):
        if self.busy: return
        path=filedialog.askopenfilename(filetypes=[('Excel','*.xlsx')])
        if not path: return
        try:
            self.save_project()
            doc=read_document(path)
            p=doc.project
            target=self.p
            if (p.get('uid') and p['uid']!=self.p['uid']) or (p.get('id') and p['id']!=self.p['id']):
                existing=self.store.load(p.get('uid',''))
                if existing: target=existing
                elif messagebox.askyesno('다른 프로젝트','별도 신규 프로젝트로 비교할까요? 현재 프로젝트는 보존합니다.'):
                    target=None
                else: return
            plan=prepare(path,target,self.store)
            self.import_preview(plan)
        except Exception as exc: messagebox.showerror('가져오기 오류',str(exc),parent=self)

    def import_preview(self,plan):
        self.store.record_preview(plan)
        win=tk.Toplevel(self);win.title('엑셀 사전 검토 — 더블클릭으로 적용 항목 선택');win.geometry('1200x800');win.transient(self);win.grab_set()
        ttk.Label(win,text='누락 행은 보존 · 빈 셀은 보존 · __CLEAR__는 삭제 · 충돌은 기본 미선택 · 관계 시트 우선').pack(anchor='w',padx=10,pady=8)
        tree=ttk.Treeview(win,columns=('select','kind','task','field','old','new'),show='headings',height=16)
        for k,label,width in [('select','적용',55),('kind','분류',80),('task','작업',110),('field','필드',110),('old','변경 전',330),('new','변경 후',330)]: tree.heading(k,text=label);tree.column(k,width=width)
        tree.pack(fill='both',expand=True)
        selected={c.key for c in plan.changes if not c.conflict}
        def render():
            tree.delete(*tree.get_children())
            for i,c in enumerate(plan.changes):
                tree.insert('','end',iid=str(i),values=('✓' if c.key in selected else '', '충돌' if c.conflict else c.kind,c.task,c.field,str(c.before),str(c.after)))
        def toggle(event):
            row=tree.identify_row(event.y)
            if row:
                key=plan.changes[int(row)].key
                if key in selected: selected.remove(key)
                else: selected.add(key)
                render()
        tree.bind('<Double-1>',toggle);render()
        report=tk.Text(win,height=12,wrap='word');report.pack(fill='both',expand=True)
        report.insert('1.0',json.dumps(plan.report(),ensure_ascii=False,indent=2)+'\n\n'+'\n'.join(map(str,plan.issues)));report.configure(state='disabled')
        ack=tk.BooleanVar();ttk.Checkbutton(win,text='오류·경고·누락 및 선택한 충돌 덮어쓰기를 확인했습니다.',variable=ack).pack(anchor='w')
        reason=tk.StringVar(value='엑셀 검토 후 적용');ttk.Label(win,text='변경 사유').pack(anchor='w');ttk.Entry(win,textvariable=reason).pack(fill='x')
        def apply():
            if self.busy: return
            try:
                if any(c.conflict and c.key in selected for c in plan.changes) and not ack.get(): raise ValueError('선택한 충돌 덮어쓰기를 확인하세요.')
                new=self.store.apply_import(plan,selected,reason.get(),ack.get())
                self.p=new;self.result=None;self.source_path=Path(plan.filename).resolve();self.sync();self.resulttext.delete('1.0','end');win.destroy();self.note('선택한 변경을 하나의 트랜잭션으로 적용했습니다.')
            except Exception as exc: messagebox.showerror('적용 차단',str(exc),parent=win)
        button=ttk.Button(win,text='선택한 변경 적용',command=apply);button.pack(pady=8)
        if any(i.severity=='오류' for i in plan.issues): button.state(['disabled'])

    def export_file(self):
        if self.busy: return
        from excel_io import write_excel
        self.save_project()
        path=filedialog.asksaveasfilename(defaultextension='.xlsx',initialfile=self.p['id']+'_schedule.xlsx',filetypes=[('Excel','*.xlsx')])
        if not path: return
        if self.source_path and Path(path).resolve()==self.source_path: messagebox.showwarning('원본 보존','다른 파일명으로 저장하세요.');return
        try: write_excel(path,self.p,self.result,self.store.baselines(self.p['uid']));self.note('표준 엑셀 내보내기 완료')
        except Exception as exc: messagebox.showerror('내보내기 오류',str(exc))

    def edit_task(self,index,copy=None):
        if self.busy: return
        t=deepcopy(self.p['tasks'][index]) if index is not None else dict(deepcopy(TASK_DEFAULTS),**(deepcopy(copy) if copy else {}))
        if index is None:
            t.update(id='T-'+uid()[:12],uid=uid(),name=(copy['name']+' 복사') if copy else '',module=copy['module'] if copy else next(iter(self.p['modules']),'공통'),state='계획',actualStart='',actualEnd='',progress=None)
        win=tk.Toplevel(self);win.title('작업 상세 · 실적 · 제약조건');win.geometry('930x850');win.transient(self);win.grab_set()
        canvas=tk.Canvas(win);scroll=ttk.Scrollbar(win,orient='vertical',command=canvas.yview);canvas.configure(yscrollcommand=scroll.set);scroll.pack(side='right',fill='y');canvas.pack(side='left',fill='both',expand=True)
        f=ttk.Frame(canvas,padding=16);canvas.create_window((0,0),window=f,anchor='nw');f.bind('<Configure>',lambda e:canvas.configure(scrollregion=canvas.bbox('all')))
        entries={}
        fields={k:v for k,v in COLS.items() if k not in ('pred','uid')}
        for n,(k,label) in enumerate(fields.items()):
            row,col=divmod(n,2);col*=2
            ttk.Label(f,text=label).grid(row=row,column=col,sticky='w',padx=5,pady=7)
            v=tk.StringVar(value='' if t.get(k) is None else t.get(k,''));entries[k]=v
            vals={'state':['계획','진행','완료','보류','제외'],'calendar':['근무일','달력일'],'calendarId':['default']+list(self.p['calendars']), 'module':list(self.p['modules'])}.get(k)
            e=ttk.Combobox(f,textvariable=v,values=vals,width=25,state='readonly') if vals else ttk.Entry(f,textvariable=v,width=27)
            e.grid(row=row,column=col+1,pady=7,padx=5)
            if k=='id' and index is not None: e.configure(state='readonly')
        bottom=(len(fields)+1)//2
        ttk.Label(f,text='선후행 관계는 [일정 · 데이터 → 관계 유형 · 시차]에서 편집합니다.\n기준 시작/완료는 구버전 참고값. 승인 일정은 별도 버전으로 보존합니다.').grid(row=bottom,column=0,columnspan=4,sticky='w')
        reason=tk.StringVar();ttk.Label(f,text='변경 사유').grid(row=bottom+1,column=0);ttk.Entry(f,textvariable=reason,width=65).grid(row=bottom+1,column=1,columnspan=3)
        def commit():
            try:
                if self.busy: return
                if not reason.get().strip(): raise ValueError('변경 사유 필요')
                new=deepcopy(t)
                for k,v in entries.items(): new[k]=v.get().strip()
                for k in NUMBERS: new[k]=float(new[k]) if new[k] else None
                for k in DATES: day(new[k])
                if not new['id'] or not new['name']: raise ValueError('ID와 작업명 필요')
                if index is None and any(x['id']==new['id'] for x in self.p['tasks']): raise ValueError('중복 ID')
                p=deepcopy(self.p)
                if index is None: p['tasks'].append(new)
                else: p['tasks'][index]=new
                errors,_,_=inspect(p,scheduling=False)
                if errors: raise ValueError('\n'.join(errors))
                self.p=p;self.changed('작업 편집',reason.get());self.sync();win.destroy()
            except Exception as exc: messagebox.showerror('입력 확인',str(exc),parent=win)
        ttk.Button(f,text='저장',command=commit).grid(row=bottom+2,column=3,pady=15)

    def set_module(self,value):
        if self.busy: return
        selected=self.mods.selection()
        if not selected: return
        p=deepcopy(self.p)
        for s in selected: p['modules'][self.mods.item(s,'values')[0]]=value
        impact=exclusion_impact(p)
        if value=='제외' and impact:
            if not messagebox.askyesno('제외 영향 검토','다음 관계는 보존되며 계산이 차단됩니다. 원복 또는 명시적인 관계 수정이 필요합니다.\n\n'+'\n'.join(f"{r['pred']} → {r['succ']} {r['type']} / {r.get('deliverable','')}" for r in impact)+'\n\n제외 상태를 저장할까요?'): return
        reason=simpledialog.askstring('모듈 변경','변경 사유:',parent=self)
        if reason: self.p=p;self.changed('모듈 변경',reason);self.fill_modules()

    def exclude_task(self):
        if self.busy: return
        i=self.select_task()
        if i is None: return
        p=deepcopy(self.p);p['tasks'][i]['state']='제외'
        impact=exclusion_impact(p)
        if impact and not messagebox.askyesno('제외 영향','관계가 유지되어 계산이 차단됩니다.\n'+'\n'.join(f"{r['pred']} → {r['succ']}" for r in impact)+'\n제외 상태를 저장할까요?'): return
        reason=simpledialog.askstring('작업 제외','변경 사유:',parent=self)
        if reason: self.p=p;self.changed('작업 제외',reason);self.fill_tasks()

    def modules_dialog(self):
        rows=[dict(name=k,state=v) for k,v in self.p['modules'].items()]
        def save(rows,reason):
            modules={}
            for row in rows:
                if not row['name'] or row['name'] in modules: raise ValueError('모듈 이름 누락/중복')
                if row['state'] not in ('적용','제외','검토 필요'): raise ValueError('모듈 적용 상태 오류')
                modules[row['name']]=row['state']
            p=deepcopy(self.p);p['modules']=modules
            errors,_,_=inspect(p,scheduling=False)
            if errors: raise ValueError('\n'.join(errors))
            impact=exclusion_impact(p)
            if impact and not messagebox.askyesno('연결 영향 확인','제외된 모듈의 연결 관계는 보존되며 계산이 차단됩니다. 저장할까요?'): return
            self.p=p;self.changed('모듈 변경',reason);self.sync()
        self.table_editor('모듈 추가 · 적용',dict(name='모듈명',state='적용 상태'),rows,save,dict(state=['적용','제외','검토 필요']),
                          '사용 중인 모듈은 삭제할 수 없습니다. 제외 후 관계를 검토하거나 적용으로 되돌리세요.')

    def relationships_dialog(self):
        self.table_editor('관계 유형 · 시차',REL_COLS,deepcopy(self.p['relations']),self.save_relations,
                          dict(type=['FS','SS','FF'],lagUnit=['달력일','근무일'],lagCalendar=['default']+list(self.p['calendars']),
                               pred=[t['id'] for t in self.p['tasks']],succ=[t['id'] for t in self.p['tasks']]),
                          '근거·산출물을 확인하세요. 제외 작업 관계도 보존됩니다. 삭제는 사유와 함께 이력에 남습니다.')

    def table_editor(self,title,columns,rows,save,choices=None,help_text=''):
        win=tk.Toplevel(self);win.title(title);win.geometry('1100x780');win.transient(self);win.grab_set()
        ttk.Label(win,text=help_text,wraplength=1050).pack(anchor='w',padx=10,pady=8)
        tree=ttk.Treeview(win,columns=list(columns),show='headings',height=10)
        for k,label in columns.items(): tree.heading(k,text=label);tree.column(k,width=110)
        tree.pack(fill='both',expand=True)
        entries={};f=ttk.Frame(win);f.pack(fill='x')
        for i,(k,label) in enumerate(columns.items()):
            row,col=divmod(i,3);col*=2;ttk.Label(f,text=label).grid(row=row,column=col,sticky='w',padx=5,pady=5)
            v=tk.StringVar();entries[k]=v
            e=ttk.Combobox(f,textvariable=v,values=choices[k],width=22) if choices and k in choices else ttk.Entry(f,textvariable=v,width=24)
            e.grid(row=row,column=col+1,padx=5,pady=5)
        def render():
            tree.delete(*tree.get_children())
            for i,r in enumerate(rows): tree.insert('','end',iid=str(i),values=[r.get(k,'') for k in columns])
        def selected(e=None):
            s=tree.selection()
            if s:
                for k,v in entries.items(): v.set(rows[int(s[0])].get(k,''))
        tree.bind('<<TreeviewSelect>>',selected)
        def put(new=False):
            r={k:v.get().strip() for k,v in entries.items()}
            s=tree.selection()
            if new or not s: rows.append(r)
            else: rows[int(s[0])]=r
            render()
        def delete():
            s=tree.selection()
            if s: rows.pop(int(s[0]));render()
        bar=ttk.Frame(win);bar.pack(fill='x')
        for label,cmd in [('새 행 추가',lambda:put(True)),('선택 행 수정',put),('선택 행 삭제',delete)]: ttk.Button(bar,text=label,command=cmd).pack(side='left',padx=8,pady=8)
        reason=tk.StringVar();ttk.Label(win,text='변경 사유 (행 추가/수정 후 전체 저장)').pack(anchor='w');ttk.Entry(win,textvariable=reason).pack(fill='x')
        def commit():
            try:
                if self.busy: return
                if not reason.get().strip(): raise ValueError('변경 사유 필요')
                save(rows,reason.get());win.destroy()
            except Exception as exc: messagebox.showerror('저장 차단',str(exc),parent=win)
        ttk.Button(win,text='전체 저장',command=commit).pack(pady=8);render()

    def save_relations(self,rows,reason):
        p=deepcopy(self.p)
        for r in rows: r['lag']=float(r.get('lag') or 0)
        p['relations']=rows;sync_pred(p)
        errors,_,_=inspect(p,scheduling=False)
        # Existing incomplete draft project dates should be fixed via settings first.
        if errors: raise ValueError('\n'.join(errors))
        self.p=p;self.changed('관계 변경',reason);self.sync()

    def calendars_dialog(self):
        rows=[dict(id='default',weekmask='0,1,2,3,4',holidays=';'.join(self.p['holidays']),workdays=';'.join(self.p['workdays']))]
        for k,cfg in self.p['calendars'].items():
            if k=='default': rows[0]=dict(id=k,weekmask=','.join(map(str,cfg['weekmask'])),holidays=';'.join(cfg['holidays']),workdays=';'.join(cfg['workdays']))
            else: rows.append(dict(id=k,weekmask=','.join(map(str,cfg.get('weekmask',[0,1,2,3,4]))),holidays=';'.join(cfg.get('holidays',[])),workdays=';'.join(cfg.get('workdays',[]))))
        def save(rows,reason):
            calendars={}
            for r in rows:
                if not r['id'] or r['id'] in calendars: raise ValueError('달력 ID 누락/중복')
                cfg=dict(weekmask=[int(x.strip()) for x in r['weekmask'].split(',')],
                         holidays=[x.strip() for x in r['holidays'].split(';') if x.strip()],workdays=[x.strip() for x in r['workdays'].split(';') if x.strip()])
                for v in cfg['holidays']+cfg['workdays']: day(v)
                calendars[r['id']]=cfg
            if 'default' not in calendars: raise ValueError('default 달력 유지 필요')
            p=deepcopy(self.p);p['calendars']=calendars;p['holidays']=calendars['default']['holidays'];p['workdays']=calendars['default']['workdays']
            errors,_,_=inspect(p,scheduling=False)
            if errors: raise ValueError('\n'.join(errors))
            self.p=p;self.changed('달력 변경',reason)
        self.table_editor('프로젝트 달력',dict(id='달력 ID',weekmask='요일(월=0~일=6, 쉼표)',holidays='휴일(; 구분)',workdays='추가근무일(; 구분)'),rows,save,
                          help_text='날짜 YYYY-MM-DD, 세미콜론 구분. 추가 근무일이 휴일보다 우선. 국가 공휴일은 자동 등록하지 않습니다.')

    def project_dialog(self):
        if self.busy: return
        win=tk.Toplevel(self);win.title('프로젝트 설정');win.geometry('650x520');win.transient(self);win.grab_set()
        entries={};f=ttk.Frame(win,padding=20);f.pack(fill='both',expand=True)
        for row,(k,label) in enumerate([('id','프로젝트 ID (식별자는 고정)'),('name','이름'),('start','착수일'),('asof','기준일 (당일 시작 시점)'),('target','관리 목표 완료일')]):
            ttk.Label(f,text=label).grid(row=row,column=0,sticky='w',pady=12)
            v=tk.StringVar(value=self.p[k]);entries[k]=v;e=ttk.Entry(f,textvariable=v,width=35);e.grid(row=row,column=1,pady=12)
            if k=='id': e.configure(state='readonly')
        ttk.Label(f,text='휴일·추가 근무일·작업별 달력은 [일정 · 데이터 → 달력]에서 편집합니다.\n목표일은 강제 날짜가 아닙니다. 고정 날짜는 작업 상세에서 지정합니다.',wraplength=570).grid(row=5,column=0,columnspan=2,pady=10)
        reason=tk.StringVar();ttk.Label(f,text='변경 사유').grid(row=6,column=0);ttk.Entry(f,textvariable=reason).grid(row=6,column=1)
        def commit():
            try:
                if not reason.get().strip(): raise ValueError('변경 사유 필요')
                p=deepcopy(self.p);p.update({k:v.get().strip() for k,v in entries.items()})
                if not p['name']: raise ValueError('프로젝트명 필요')
                for k in ('start','asof','target'):
                    if day(p[k]) is None and k!='target': raise ValueError('착수일/기준일 필요')
                if day(p['asof'])<day(p['start']): raise ValueError('기준일은 착수일 이후여야 합니다.')
                self.p=p;self.changed('설정 변경',reason.get());self.sync();win.destroy()
            except Exception as exc: messagebox.showerror('설정 오류',str(exc),parent=win)
        ttk.Button(f,text='설정 저장',command=commit).grid(row=7,column=1,pady=20)

    def check_input(self):
        self.apply_dates();errors,warnings,_=inspect(self.p)
        text='오류 (계산 차단)\n'+('\n'.join(errors) or '없음')+'\n\n경고\n'+('\n'.join(warnings) or '없음')
        if not errors:
            try:
                r=deterministic(self.p);text+='\n\n계산 경고\n'+('\n'.join(r['warnings']) or '없음')
            except Exception as exc: text+='\n\n일정 충돌\n'+str(exc)
        self.show_text('입력 오류 · 일정 충돌',text)
        return errors,warnings

    def recalculate(self):
        self.save_project()
        r=deterministic(self.p)
        self.result=r;self.store.add(self.p,'결정론적 계산','최빈 M 기준 예측',result_record(r,self.p));self.fill_history();self.metrics_update();self.draw();self.display_result(r)
        self.note('결정론적 예측 완료 · 기준 일정 보존')

    def display_result(self,r):
        text=f"엔진 {r['engine']} / 데이터 v{r['data_version']} / 기준일 {r['asof']} / 반복 {r['n']} / Seed {r['seed']}\n"
        text+='\n'.join(r.get('warnings',[]))+'\n'
        if r.get('target_delta') is not None: text+=f"목표일까지 차이(달력일): {r['target_delta']} (음수=지연)\n"
        text+='\nID | 예측 시작(P50) | 완료(P50) | 총 여유(달력일) | 임계\n'
        from engine import percentile
        for t in self.p['tasks']:
            k=t['id'];floating=r.get('float',{}).get(k)
            text+=f"{k} | {textday(percentile(r['starts'][k],50))} | {textday(percentile(r['ends'][k],50))} | {floating if floating is not None else '미정의/확률결과'} | {'임계' if k in r.get('critical',[]) else ''}\n"
        self.resulttext.delete('1.0','end');self.resulttext.insert('1.0',text);self.tabs.select(self.resulttab)

    def baseline(self):
        if self.busy: return
        try:
            self.save_project();r=deterministic(self.p)
            name=simpledialog.askstring('기준 일정 승인','버전 이름 / 승인 근거:',parent=self)
            if name: self.store.baseline(self.p,name,r);self.note('승인 기준 일정 새 버전 저장 완료')
        except Exception as exc: messagebox.showerror('기준 일정 저장 차단',str(exc))

    def view_baselines(self):
        lines=[]
        for b in self.store.baselines(self.p['uid']):
            lines.append(f"#{b['id']} {b['name']} / 데이터 v{b['version']} / {b['created']} / {b['actor']}")
            if b['legacy']: lines.append('구버전 입력 스냅샷 (계산된 승인 날짜 아님)')
            else:
                for k,s in b['payload']['starts'].items(): lines.append(f"  {k}: {textday(s)} ~ {textday(b['payload']['ends'][k])}")
            lines.append('')
        self.show_text('승인 기준 일정 버전 (읽기 전용)','\n'.join(lines) or '저장된 버전이 없습니다.')

    def history_detail(self):
        s=self.hist.selection()
        payload=self.store.get(int(s[0])) if s else {}
        self.show_text('변경 이력 · 필드별 전후값',json.dumps(dict(event=payload,changes=self.store.history(self.p['uid'])),ensure_ascii=False,indent=2))

    def restore(self):
        if self.busy: return
        s=self.hist.selection()
        if s and messagebox.askyesno('입력 복원','선택한 입력을 새 데이터 버전으로 복원할까요? 기준 일정 버전은 유지됩니다.'):
            p=canonical(self.store.get(int(s[0]))['project'])
            if p['uid']!=self.p['uid']: raise ValueError('다른 프로젝트 이력')
            p['data_version']=self.p['data_version'];self.p=p;self.changed('입력 복원','이력 #'+s[0]);self.sync()

    def backup_db(self):
        self.save_project()
        path=filedialog.asksaveasfilename(defaultextension='.sqlite3',filetypes=[('SQLite backup','*.sqlite3')])
        if path: self.store.backup(path);self.note('일관된 DB 백업 완료: '+path)

    def restore_db(self):
        path=filedialog.askopenfilename(filetypes=[('SQLite backup','*.sqlite3')])
        if path and messagebox.askyesno('DB 복원','모든 프로젝트 DB를 이 백업으로 복원합니다. 현재 DB는 자동 백업합니다. 계속할까요?'):
            self.store.restore_backup(path);projects=self.store.projects()
            if projects: self.p=self.store.load(projects[0][0])
            else:
                from engine import demo
                self.p=canonical(demo());self.store.add(self.p,'새 프로젝트','빈 백업 복원 후 예제 생성')
            self.result=None;self.source_path=None;self.sync();self.resulttext.delete('1.0','end');self.note('DB 복원 완료')

    def poll(self):
        try:
            while True:
                kind,value=self.events.get_nowait()
                if kind=='progress': self.progress['value']=value
                else:
                    self.busy=False;self.runbtn.state(['!disabled']);self.cancelbtn.state(['disabled'])
                    for b in self.buttons+self.date_entries: b.state(['!disabled'])
                    if kind=='done':
                        try:
                            self.store.add(self.p,'몬테카를로','분석 입력·예측 결과',result_record(value,self.p))
                            self.result=value;self.fill_history();self.metrics_update();self.draw();self.display_result(value);self.note(f"{value['n']:,}회 계산 완료")
                        except Exception as exc: self.result=None;messagebox.showerror('결과 저장 실패',str(exc))
                    else: messagebox.showerror('분석 차단',value);self.note('분석 중단 · 제약과 입력을 확인하세요.')
        except queue.Empty: pass
        self.poll_id=self.after(100,self.poll)

    def close(self):
        if self.busy:
            if not messagebox.askyesno('종료','계산을 중단하고 종료할까요?'): return
            self.cancel.set()
        try:
            if not self.busy: self.save_project()
        except Exception as exc: messagebox.showerror('저장 실패 — 종료 취소',str(exc));return
        if getattr(self,'poll_id',None): self.after_cancel(self.poll_id)
        self.store.db.close();self.destroy()
