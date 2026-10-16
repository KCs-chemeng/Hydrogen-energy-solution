"""Project Schedule Lab desktop prototype. Python 3.11+ / Windows."""
import tkinter as tk
from tkinter import ttk,filedialog,messagebox,simpledialog
import json,sqlite3,threading,queue,os,sys,traceback
from pathlib import Path
from datetime import datetime,date,timedelta
from collections import defaultdict
import numpy as np
from engine import demo,simulate,validate,summary,percentile,textday,day,VERSION
from excel_io import read_excel,write_excel,COLS,DATES
BG='#f1f5f9';NAVY='#15304c';BLUE='#1261cb';GREEN='#05976d';MUTED='#64748b'
BASE=Path(__file__).resolve().parent
DATA=Path(os.environ.get('LOCALAPPDATA',str(Path.home())))/'ProjectScheduleLab';DATA.mkdir(parents=True,exist_ok=True)
class Store:
    def __init__(self,path):
        self.db=sqlite3.connect(path)
        self.db.execute('CREATE TABLE IF NOT EXISTS events (seq INTEGER PRIMARY KEY, created TEXT, pj TEXT, kind TEXT, reason TEXT, payload TEXT)')
        self.db.commit()
    def add(self,p,kind,reason,result=None):
        payload={'project':p,'summary':result};self.db.execute('INSERT INTO events(created,pj,kind,reason,payload) VALUES(?,?,?,?,?)',(datetime.now().isoformat(timespec='seconds'),p['id'],kind,reason,json.dumps(payload,ensure_ascii=False)));self.db.commit()
    def rows(self):return self.db.execute('SELECT seq,created,pj,kind,reason FROM events ORDER BY seq DESC LIMIT 300').fetchall()
    def get(self,seq):return json.loads(self.db.execute('SELECT payload FROM events WHERE seq=?',(seq,)).fetchone()[0])
class App(tk.Tk):
    def __init__(self):
        super().__init__();self.title('공정표 | Project Schedule Lab v'+VERSION);self.geometry('1500x900');self.minsize(1100,680);self.configure(bg=BG)
        self.option_add('*Font',('맑은 고딕',10));self.style=ttk.Style(self);self.style.theme_use('clam')
        self.style.configure('.',font=('맑은 고딕',10));self.style.configure('TFrame',background=BG);self.style.configure('TLabel',background=BG,foreground=NAVY);self.style.configure('TButton',padding=(12,7),background='#ffffff',foreground=NAVY);self.style.map('TButton',background=[('active','#dbeafe')]);self.style.configure('Primary.TButton',background=BLUE,foreground='white');self.style.configure('Treeview',rowheight=32,background='white',fieldbackground='white',borderwidth=0);self.style.configure('Treeview.Heading',background='#e8eef5',foreground=NAVY,padding=9);self.style.map('Treeview',background=[('selected','#dbeafe')],foreground=[('selected',NAVY)]);self.style.configure('TNotebook.Tab',padding=(24,12))
        self.store=Store(DATA/'history.sqlite3');self.p=demo();self.result=None;self.busy=False;self.events=queue.Queue();self.cancel=threading.Event();self.collapsed=set();self.rowids={};self.source_path=None
        previous=self.store.rows()
        if previous:
            try:self.p=self.store.get(previous[0][0])['project']
            except Exception:pass
        self.protocol('WM_DELETE_WINDOW',self.close)
        self.build();self.sync();self.after(100,self.poll)
    def build(self):
        header=tk.Frame(self,bg=NAVY,height=70);header.pack(fill='x');header.pack_propagate(False)
        tk.Label(header,text='▥  공정표',font=('맑은 고딕',21,'bold'),bg=NAVY,fg='white').pack(side='left',padx=24)
        tk.Label(header,text='PROJECT SCHEDULE LAB   /   PYTHON DESKTOP',font=('Segoe UI',10),bg=NAVY,fg='#b6c9dd').pack(side='left')
        tk.Label(header,text='v0.1.0  초안',bg=NAVY,fg='#98e1c6').pack(side='right',padx=25)
        self.toolbar=ttk.Frame(self,padding=(18,12));self.toolbar.pack(fill='x');self.buttons=[]
        for title,cmd in [('엑셀 가져오기',self.import_file),('입력·결과 엑셀 내보내기',self.export_file),('공통 243개 WBS',self.common),('예제 프로젝트',self.load_demo),('설정 · 휴일',self.project_dialog),('기준선 저장',self.baseline)]:
            b=ttk.Button(self.toolbar,text=title,command=cmd);b.pack(side='left',padx=(0,8));self.buttons.append(b)
        self.caption=ttk.Label(self,padding=(24,0,24,8),font=('맑은 고딕',13,'bold'));self.caption.pack(anchor='w')
        cfg=ttk.Frame(self,padding=(24,8));cfg.pack(fill='x');self.config={};self.date_entries=[]
        for k,label in [('start','시작일'),('asof','분석 기준일'),('target','목표 완료일')]:
            ttk.Label(cfg,text=label).pack(side='left',padx=(0,7));v=tk.StringVar();self.config[k]=v;e=ttk.Entry(cfg,textvariable=v,width=12);e.pack(side='left',padx=(0,14));e.bind('<FocusOut>',self.apply_dates);self.date_entries.append(e)
        ttk.Label(cfg,text='반복').pack(side='left');self.count=tk.StringVar(value='5000');ttk.Combobox(cfg,textvariable=self.count,values=['1000','5000','10000'],width=7,state='readonly').pack(side='left',padx=7)
        ttk.Label(cfg,text='Seed').pack(side='left');self.seed=tk.StringVar(value='42');ttk.Entry(cfg,textvariable=self.seed,width=7).pack(side='left',padx=7)
        self.runbtn=ttk.Button(cfg,text='▶  시뮬레이션 실행',style='Primary.TButton',command=self.run);self.runbtn.pack(side='right')
        self.cancelbtn=ttk.Button(cfg,text='중지',command=self.cancel.set);self.cancelbtn.pack(side='right',padx=7);self.cancelbtn.state(['disabled'])
        stats=ttk.Frame(self,padding=(24,10));stats.pack(fill='x');self.metrics=[]
        for name in ['P50 완료일','P80 완료일','P90 완료일','목표일 달성 확률']:
            frame=tk.Frame(stats,bg='white',highlightbackground='#dae3ed',highlightthickness=1);frame.pack(side='left',fill='x',expand=True,padx=(0,10));tk.Label(frame,text=name,bg='white',fg=MUTED,font=('맑은 고딕',10)).pack(anchor='w',padx=15,pady=(10,3));v=tk.Label(frame,text='—',bg='white',fg=NAVY,font=('Segoe UI',21,'bold'));v.pack(anchor='w',padx=15,pady=(0,10));self.metrics.append(v)
        self.tabs=ttk.Notebook(self);self.tabs.pack(fill='both',expand=True,padx=24,pady=10)
        self.roadtab=ttk.Frame(self.tabs);self.disttab=ttk.Frame(self.tabs);self.tasktab=ttk.Frame(self.tabs);self.modtab=ttk.Frame(self.tabs);self.histtab=ttk.Frame(self.tabs);self.guidetab=ttk.Frame(self.tabs)
        for tab,title in [(self.roadtab,'확률 로드맵'),(self.disttab,'완료일 분포'),(self.tasktab,'작업 편집'),(self.modtab,'모듈 적용'),(self.histtab,'변경 · 분석 이력'),(self.guidetab,'사용 안내')]:self.tabs.add(tab,text=title)
        self.distribution=tk.Canvas(self.disttab,bg='white',highlightthickness=0);self.distribution.pack(fill='both',expand=True);self.distribution.bind('<Configure>',lambda e:self.draw_distribution())
        bar=ttk.Frame(self.roadtab,padding=10);bar.pack(fill='x');self.q=tk.IntVar(value=80);self.qlabel=ttk.Label(bar,text='확률 기준  P80',font=('맑은 고딕',11,'bold'));self.qlabel.pack(side='left',padx=10)
        tk.Scale(bar,from_=50,to=95,resolution=5,orient='horizontal',variable=self.q,command=lambda _:self.draw(),length=270,bg=BG,highlightthickness=0).pack(side='left')
        ttk.Label(bar,text='모듈').pack(side='left',padx=(25,5));self.filter=tk.StringVar(value='전체');self.filterbox=ttk.Combobox(bar,textvariable=self.filter,state='readonly',width=22);self.filterbox.pack(side='left');self.filterbox.bind('<<ComboboxSelected>>',lambda e:self.draw())
        self.links=tk.BooleanVar(value=True);ttk.Checkbutton(bar,text='선행 연결선',variable=self.links,command=self.draw).pack(side='left',padx=15)
        ttk.Button(bar,text='전체 펼치기',command=self.expand).pack(side='right')
        ttk.Label(self.roadtab,text='  파랑: 시작일 P50~선택 P   초록: 완료일 P50~선택 P   얇은선: 시작 P10~완료 선택 P   연결선: P50 기준 관계 (단일 실행 일정 아님)',foreground=MUTED,padding=(10,4)).pack(fill='x')
        area=ttk.Frame(self.roadtab);area.pack(fill='both',expand=True);area.rowconfigure(0,weight=1);area.columnconfigure(0,weight=1)
        self.canvas=tk.Canvas(area,bg='white',highlightthickness=0);self.canvas.grid(row=0,column=0,sticky='nsew');ys=ttk.Scrollbar(area,orient='vertical',command=self.canvas.yview);ys.grid(row=0,column=1,sticky='ns');xs=ttk.Scrollbar(area,orient='horizontal',command=self.canvas.xview);xs.grid(row=1,column=0,sticky='ew');self.canvas.configure(yscrollcommand=ys.set,xscrollcommand=xs.set);self.canvas.bind('<Configure>',lambda e:self.draw());self.canvas.bind('<Double-1>',self.canvas_edit);self.canvas.bind('<Button-1>',self.canvas_click);self.canvas.bind('<MouseWheel>',lambda e:self.canvas.yview_scroll(int(-e.delta/120),'units'))
        tb=ttk.Frame(self.tasktab,padding=10);tb.pack(fill='x');ttk.Label(tb,text='작업명 또는 ID 검색').pack(side='left');self.search=tk.StringVar();self.search.trace_add('write',lambda *_:self.fill_tasks());ttk.Entry(tb,textvariable=self.search,width=30).pack(side='left',padx=10)
        for text,cmd in [('작업 추가',lambda:self.edit_task(None)),('선택 편집',self.edit_selected),('선택 복사',self.copy_task),('선택 제외',self.exclude_task)]:ttk.Button(tb,text=text,command=cmd).pack(side='left',padx=4)
        columns=['id','module','name','state','calendar','o','m','p','pred'];self.tree=ttk.Treeview(self.tasktab,columns=columns,show='headings',selectmode='browse')
        for k in columns:self.tree.heading(k,text=COLS[k]);self.tree.column(k,width=280 if k=='name' else 150 if k in ['module','pred'] else 75,stretch=k in ['name','module','pred'])
        self.tree.pack(side='left',fill='both',expand=True);scroll=ttk.Scrollbar(self.tasktab,command=self.tree.yview);scroll.pack(side='right',fill='y');self.tree.configure(yscrollcommand=scroll.set);self.tree.bind('<Double-1>',lambda e:self.edit_selected())
        ttk.Label(self.modtab,text='선택한 모듈의 적용 여부를 바꾸세요. 검토 필요가 남아 있으면 계산을 차단합니다.',padding=12).pack(anchor='w')
        mb=ttk.Frame(self.modtab);mb.pack(fill='x')
        for v in ['적용','제외','검토 필요']:ttk.Button(mb,text='선택 모듈 → '+v,command=lambda x=v:self.set_module(x)).pack(side='left',padx=8,pady=6)
        self.mods=ttk.Treeview(self.modtab,columns=['name','value','count'],show='headings');self.mods.pack(fill='both',expand=True)
        for k,t in [('name','모듈'),('value','적용 여부'),('count','작업 수')]:self.mods.heading(k,text=t)
        hb=ttk.Frame(self.histtab,padding=10);hb.pack(fill='x');ttk.Button(hb,text='선택 시점 입력 복원',command=self.restore).pack(side='left');ttk.Button(hb,text='선택 변경 상세',command=self.history_detail).pack(side='left',padx=8);ttk.Label(hb,text='최근 300개 표시 · 전체 기록은 로컬 DB에 보관',foreground=MUTED).pack(side='right')
        self.hist=ttk.Treeview(self.histtab,columns=['seq','created','pj','kind','reason'],show='headings');self.hist.pack(fill='both',expand=True)
        for k,t in [('seq','번호'),('created','시각'),('pj','PJ ID'),('kind','구분'),('reason','변경 사유')]:self.hist.heading(k,text=t);self.hist.column(k,width=60 if k=='seq' else 160)
        guide=tk.Text(self.guidetab,wrap='word',relief='flat',bg='white',fg=NAVY,padx=25,pady=20,font=('맑은 고딕',11));guide.pack(fill='both',expand=True);guide.insert('1.0',GUIDE);guide.configure(state='disabled')
        footer=ttk.Frame(self,padding=(24,5,24,12));footer.pack(fill='x');self.status=ttk.Label(footer,text='준비');self.status.pack(side='left');self.progress=ttk.Progressbar(footer,length=200,maximum=100);self.progress.pack(side='right')
    def note(self,s):self.status.configure(text=s)
    def sync(self):
        for k,v in self.config.items():v.set(self.p.get(k,''))
        self.caption.configure(text=f'{self.p["id"]}   /   {self.p["name"]}    ·    {len(self.p["tasks"])}개 작업')
        self.filterbox['values']=['전체']+list(self.p['modules']);self.filter.set('전체');self.collapsed.clear();self.fill_tasks();self.fill_modules();self.fill_history();self.metrics_update();self.draw()
    def metrics_update(self):
        values=['—']*4
        if self.result:
            s=summary(self.result,self.p);values=[s['P50'],s['P80'],s['P90'],f'{s["probability"]*100:.1f}%' if s['probability'] is not None else '목표 미설정']
        for l,v in zip(self.metrics,values):l.configure(text=v)
    def changed(self,kind,reason):
        self.result=None;self.store.add(self.p,kind,reason);self.metrics_update();self.fill_history();self.draw();self.note('입력 변경됨 · 분석 결과를 갱신하려면 다시 실행하세요.')
    def apply_dates(self,event=None):
        if self.busy:return
        changes=[]
        for k,v in self.config.items():
            s=v.get().strip()
            if s!=self.p.get(k,''):changes.append(f'{k}: {self.p.get(k,"")} → {s}')
        if changes:
            for k,v in self.config.items():self.p[k]=v.get().strip()
            self.changed('날짜 변경','; '.join(changes))
    def fill_tasks(self):
        self.tree.delete(*self.tree.get_children());q=self.search.get().lower()
        for i,t in enumerate(self.p['tasks']):
            if q and q not in (t['id']+' '+t['name']).lower():continue
            self.tree.insert('', 'end',iid=str(i),values=[', '.join(t[k]) if k=='pred' else t.get(k,'') if t.get(k) is not None else '' for k in ['id','module','name','state','calendar','o','m','p','pred']])
    def fill_modules(self):
        self.mods.delete(*self.mods.get_children())
        for i,(m,v) in enumerate(self.p['modules'].items()):self.mods.insert('','end',iid=str(i),values=[m,v,sum(t['module']==m for t in self.p['tasks'])])
    def fill_history(self):
        self.hist.delete(*self.hist.get_children())
        for row in self.store.rows():self.hist.insert('','end',iid=str(row[0]),values=row)
    def select_task(self):
        s=self.tree.selection();return int(s[0]) if s else None
    def edit_selected(self):
        i=self.select_task()
        if i is not None:self.edit_task(i)
    def copy_task(self):
        i=self.select_task()
        if i is None or self.busy:return
        self.edit_task(None,self.p['tasks'][i])
    def exclude_task(self):
        i=self.select_task()
        if i is None or self.busy:return
        reason=simpledialog.askstring('작업 제외','제외 사유:',parent=self)
        if reason:
            self.p['tasks'][i]['state']='제외';self.changed('작업 제외',self.p['tasks'][i]['id']+': '+reason);self.fill_tasks()
    def edit_task(self,index,copy=None):
        if self.busy:return
        t=dict(self.p['tasks'][index]) if index is not None else dict(copy or {},id='TASK-'+datetime.now().strftime('%H%M%S'),name=(copy['name']+' 복사') if copy else '',module=copy['module'] if copy else next(iter(self.p['modules']),'공통'),state='계획')
        win=tk.Toplevel(self);win.title('작업 상세 편집');win.geometry('860x730');win.transient(self);win.grab_set();frame=ttk.Frame(win,padding=20);frame.pack(fill='both',expand=True);frame.columnconfigure(1,weight=1);frame.columnconfigure(3,weight=1);entries={}
        for n,(k,label) in enumerate(COLS.items()):
            row,col=divmod(n,2);col*=2;ttk.Label(frame,text=label).grid(row=row,column=col,sticky='w',padx=6,pady=8)
            value=t.get(k,'');value=', '.join(value) if k=='pred' and isinstance(value,list) else '' if value is None else value
            v=tk.StringVar(value=value);entries[k]=v
            vals={'state':['계획','진행','완료','보류','제외'],'calendar':['달력일','근무일'],'module':list(self.p['modules'])}.get(k)
            e=ttk.Combobox(frame,textvariable=v,values=vals,state='readonly' if k in ['state','calendar'] else 'normal') if vals else ttk.Entry(frame,textvariable=v)
            e.grid(row=row,column=col+1,sticky='ew',padx=6,pady=8)
            if k=='id' and index is not None:e.configure(state='readonly')
            if k=='calendar' and not value:v.set('달력일')
        ttk.Label(frame,text='날짜: YYYY-MM-DD / 선행 ID: 쉼표 구분 / 기간: 0 ≤ O ≤ M ≤ P / 진행은 잔여 기간 입력',foreground=MUTED).grid(row=11,column=0,columnspan=4,pady=10,sticky='w')
        ttk.Label(frame,text='변경 사유').grid(row=12,column=0);reason=tk.StringVar();ttk.Entry(frame,textvariable=reason).grid(row=12,column=1,columnspan=3,sticky='ew')
        def commit():
            try:
                new={k:v.get().strip() for k,v in entries.items()};new['pred']=[x.strip() for x in new['pred'].replace(';',',').split(',') if x.strip()]
                for k in ['o','m','p','ro','rm','rp']:new[k]=float(new[k]) if new[k] else None
                if not new['id'] or not new['name']:raise ValueError('ID와 작업명을 입력하세요.')
                if index is None and any(x['id']==new['id'] for x in self.p['tasks']):raise ValueError('중복 ID입니다.')
                for k in DATES:
                    if new[k]:day(new[k])
                if not reason.get().strip():raise ValueError('변경 사유를 입력하세요.')
                diffs='; '.join(f'{COLS[k]}: {t.get(k)} → {new[k]}' for k in COLS if t.get(k)!=new[k])
                if index is None:self.p['tasks'].append(new)
                else:self.p['tasks'][index]=new
                if new['module'] not in self.p['modules']:self.p['modules'][new['module']]='검토 필요'
                self.changed('작업 편집',new['id']+' | '+reason.get()+' | '+diffs);self.sync();win.destroy()
            except Exception as e:messagebox.showerror('입력 확인',str(e),parent=win)
        ttk.Button(frame,text='변경 저장',style='Primary.TButton',command=commit).grid(row=13,column=3,pady=20,sticky='e')
    def set_module(self,value):
        if self.busy:return
        selected=self.mods.selection()
        if not selected:return
        reason=simpledialog.askstring('모듈 변경','변경 사유:',parent=self)
        if not reason:return
        names=[]
        for s in selected:
            m=self.mods.item(s,'values')[0];names.append(m);self.p['modules'][m]=value
        self.changed('모듈 변경',', '.join(names)+' → '+value+' | '+reason);self.fill_modules()
    def replace_project(self,p,reason):
        self.p=p;self.result=None;self.store.add(p,'프로젝트 불러오기',reason);self.sync();self.note(reason+' · 적용 모듈과 날짜를 확인하세요.')
    def import_file(self):
        if self.busy:return
        path=filedialog.askopenfilename(filetypes=[('Excel workbook','*.xlsx')])
        if not path:return
        try:
            p=read_excel(path)
            if messagebox.askyesno('가져오기',f'{len(p["tasks"])}개 작업을 가져옵니다. 현재 입력은 이력에 보관하고 교체할까요?'):
                self.store.add(self.p,'교체 전 저장','엑셀 가져오기 전');self.source_path=Path(path).resolve();self.replace_project(p,Path(path).name)
        except Exception as e:messagebox.showerror('가져오기 오류',str(e))
    def export_file(self):
        if self.busy:return
        self.apply_dates();path=filedialog.asksaveasfilename(defaultextension='.xlsx',initialfile='PJ_schedule.xlsx',filetypes=[('Excel workbook','*.xlsx')])
        if not path:return
        if self.source_path and Path(path).resolve()==self.source_path:messagebox.showwarning('다른 이름으로 저장','원본 보존을 위해 다른 파일명을 지정하세요.');return
        try:write_excel(path,self.p,self.result);self.note('엑셀 내보내기 완료: '+path)
        except Exception as e:messagebox.showerror('저장 오류',str(e))
    def common(self):
        if self.busy:return
        if messagebox.askyesno('공통 템플릿','현재 입력은 이력에 보관하고 공통 243개 작업을 불러올까요?'):
            try:self.store.add(self.p,'교체 전 저장','공통 템플릿 전환');self.source_path=BASE/'common_Rev1.xlsx';self.replace_project(read_excel(self.source_path),'공통 템플릿')
            except Exception as e:messagebox.showerror('오류',str(e))
    def load_demo(self):
        if not self.busy and messagebox.askyesno('예제','현재 입력은 이력에 보관하고 예제로 전환할까요?'):self.store.add(self.p,'교체 전 저장','예제 전환');self.source_path=None;self.replace_project(demo(),'예제 (기간은 설명용 가정값)')
    def project_dialog(self):
        if self.busy:return
        win=tk.Toplevel(self);win.title('프로젝트 설정');win.geometry('560x510');win.transient(self);win.grab_set();f=ttk.Frame(win,padding=24);f.pack(fill='both',expand=True);vs={}
        for k,label in [('id','PJ ID'),('name','프로젝트명')]:
            ttk.Label(f,text=label).pack(anchor='w');v=tk.StringVar(value=self.p[k]);vs[k]=v;ttk.Entry(f,textvariable=v).pack(fill='x',pady=(3,12))
        ttk.Label(f,text='휴일 목록 · 한 줄에 YYYY-MM-DD 하나\n근무일 작업에만 적용. 국가 공휴일은 자동 등록하지 않습니다.').pack(anchor='w');txt=tk.Text(f,height=9);txt.pack(fill='both',expand=True,pady=10);txt.insert('1.0','\n'.join(self.p.get('holidays',[])))
        def commit():
            try:
                holidays=[x.strip() for x in txt.get('1.0','end').splitlines() if x.strip()]
                for h in holidays:day(h)
                if not all(v.get().strip() for v in vs.values()):raise ValueError('ID와 프로젝트명을 입력하세요.')
                self.p.update({k:v.get().strip() for k,v in vs.items()});self.p['holidays']=sorted(set(holidays));self.changed('설정 변경','프로젝트 정보 / 휴일 목록 변경');self.sync();win.destroy()
            except Exception as e:messagebox.showerror('오류',str(e),parent=win)
        ttk.Button(f,text='설정 저장',command=commit).pack(anchor='e')
    def baseline(self):
        if self.busy:return
        self.apply_dates();name=simpledialog.askstring('기준선','기준선 이름 / 승인 근거:',parent=self)
        if name:self.store.add(self.p,'기준선',name,summary(self.result,self.p) if self.result else None);self.fill_history();self.note('기준선 스냅샷 저장 완료')
    def restore(self):
        if self.busy:return
        s=self.hist.selection()
        if s and messagebox.askyesno('복원','선택 시점의 입력으로 복원할까요? 현재 입력도 이력에 남깁니다.'):
            p=self.store.get(int(s[0]))['project'];self.store.add(self.p,'복원 전 저장','복원 전 상태');self.source_path=None;self.replace_project(p,'이력 #'+s[0]+' 복원')
    def history_detail(self):
        s=self.hist.selection()
        if not s:return
        row=self.hist.item(s[0],'values');payload=self.store.get(int(s[0]));win=tk.Toplevel(self);win.title('이력 상세');win.geometry('780x500');t=tk.Text(win,wrap='word',padx=20,pady=20);t.pack(fill='both',expand=True);t.insert('1.0',row[4]+'\n\n'+json.dumps(payload.get('summary'),ensure_ascii=False,indent=2));t.configure(state='disabled')
    def run(self):
        if self.busy:return
        self.apply_dates();errors,_=validate(self.p)
        if errors:messagebox.showerror('입력 검증', '\n'.join(errors[:30])+('\n… 추가 오류 생략' if len(errors)>30 else ''));return
        try:n=int(self.count.get());seed=int(self.seed.get());assert 0<=seed<=2**32-1
        except Exception:messagebox.showerror('입력 검증','반복 횟수와 Seed(0~4294967295)를 확인하세요.');return
        self.busy=True;self.cancel.clear();self.runbtn.state(['disabled']);self.cancelbtn.state(['!disabled']);self.result=None;self.metrics_update();self.draw();self.note('시뮬레이션 계산 중…');snapshot=json.loads(json.dumps(self.p))
        for b in self.buttons+self.date_entries:b.state(['disabled'])
        def worker():
            try:self.events.put(('done',simulate(snapshot,n,seed,lambda x:self.events.put(('progress',x)),self.cancel)))
            except Exception as e:self.events.put(('error',str(e)))
        threading.Thread(target=worker,daemon=True).start()
    def poll(self):
        try:
            while True:
                kind,value=self.events.get_nowait()
                if kind=='progress':self.progress['value']=value
                else:
                    self.busy=False;self.runbtn.state(['!disabled']);self.cancelbtn.state(['disabled'])
                    for b in self.buttons+self.date_entries:b.state(['!disabled'])
                    if kind=='done':self.result=value;self.store.add(self.p,'분석 실행','몬테카를로 입력·결과',summary(value,self.p));self.fill_history();self.metrics_update();self.draw();self.note(f'{value["n"]:,}회 분석 완료 · 입력 가정에 따른 예측입니다.');self.tabs.select(self.roadtab)
                    else:messagebox.showerror('분석 중단',value);self.note('분석 중단 · 입력 확인 필요')
        except queue.Empty:pass
        self.after(100,self.poll)
    def expand(self):self.collapsed.clear();self.draw()
    def draw(self):
        if not hasattr(self,'canvas'):return
        self.draw_distribution()
        c=self.canvas;c.delete('all');self.rowids={};self.group_y={};q=self.q.get();self.qlabel.configure(text=f'확률 기준  P{q}')
        width=max(1250,c.winfo_width());left=430;right=width-40
        c.create_rectangle(0,0,width,44,fill='#e8eef5',outline='');c.create_text(20,22,text='작업 / 모듈',anchor='w',font=('맑은 고딕',11,'bold'),fill=NAVY)
        if not self.result:
            c.create_text(width/2,170,text='시뮬레이션을 실행하면 확률 로드맵이 표시됩니다.',font=('맑은 고딕',16),fill=MUTED);c.create_text(width/2,210,text='엑셀 가져오기 → 적용 모듈 검토 → 시작일·기준일 입력 → 실행',font=('맑은 고딕',11),fill=MUTED);c.configure(scrollregion=(0,0,width,300));return
        r=self.result;visible=[t for t in self.p['tasks'] if self.filter.get() in ['전체',t['module']] and self.p['modules'][t['module']]=='적용' and t['state']!='제외'];groups=defaultdict(list)
        for t in visible:groups[t['module']].append(t)
        lo=min(day(self.p['start']),min(percentile(v,10) for v in r['starts'].values()));hi=max(lo+30,percentile(r['total'],95),day(self.p.get('target')) or lo);span=hi-lo;x=lambda d:left+(d-lo)/span*(right-left)
        ys={};y=92;group_rows=[]
        for m,ts in groups.items():
            group_rows.append((y,m,len(ts)));self.group_y[y]=m;y+=36
            if m not in self.collapsed:
                for t in ts:ys[t['id']]=y+20;self.rowids[y+20]=t['id'];y+=40
        height=max(y+30,c.winfo_height());step=max(7,int(span/9))
        for d in range(lo,hi+1,step):
            px=x(d);c.create_line(px,44,px,height,fill='#e2e8f0');c.create_text(px+3,22,text=textday(d)[5:],anchor='w',fill=MUTED,font=('Segoe UI',10))
        c.create_line(left-15,0,left-15,height,fill='#d9e2ed')
        for gy,m,n in group_rows:
            c.create_rectangle(0,gy,width,gy+36,fill='#eef2f7',outline='');c.create_text(18,gy+18,text=('＋' if m in self.collapsed else '－')+'  '+m,anchor='w',fill=NAVY,font=('맑은 고딕',11,'bold'));c.create_text(right,gy+18,text=f'{n}개 작업',anchor='e',fill=MUTED)
        if self.links.get():
            for t in visible:
                if t['id'] not in ys:continue
                for pred in t['pred']:
                    if pred in ys:
                        a=x(percentile(r['ends'][pred],50));b=x(percentile(r['starts'][t['id']],50));c.create_line(a,ys[pred],a+14,ys[pred]+8,b-14,ys[t['id']]-8,b,ys[t['id']],smooth=True,fill='#bdcbdc',arrow=tk.LAST,width=1)
        for t in visible:
            k=t['id']
            if k not in ys:continue
            yy=ys[k];c.create_line(0,yy+20,width,yy+20,fill='#f0f3f7');label=t['name'];label=label if len(label)<34 else label[:33]+'…';c.create_text(24,yy-7,text=label,anchor='w',fill=NAVY,font=('맑은 고딕',10));c.create_text(24,yy+9,text=f'{k}   ·   {t["state"]}   ·   {t.get("calendar","달력일")}',anchor='w',fill=MUTED,font=('맑은 고딕',8))
            s10,s50,sq=[x(percentile(r['starts'][k],v)) for v in (10,50,q)];f50,fq=[x(percentile(r['ends'][k],v)) for v in (50,q)]
            c.create_line(s10,yy,fq,yy,fill='#5cddb0',width=4);c.create_rectangle(s50,yy-8,max(s50+3,sq),yy+8,fill=BLUE,outline='');c.create_rectangle(f50,yy-8,max(f50+3,fq),yy+8,fill=GREEN,outline='');c.create_line(s50,yy,f50,yy,fill='#3387c6',width=2)
        total=percentile(r['total'],q);c.create_rectangle(0,44,width,88,fill='#f8fafc',outline='');c.create_text(24,66,text=f'PROJECT FINISH   ·   P{q}  {textday(total)}',anchor='w',font=('Segoe UI',11,'bold'),fill=NAVY);c.create_line(x(percentile(r['total'],10)),66,x(total),66,fill='#44d7a5',width=5);c.create_oval(x(total)-5,61,x(total)+5,71,fill=GREEN,outline='')
        if self.p.get('target'):
            tx=x(day(self.p['target']));c.create_line(tx,44,tx,height,fill='#d59b38',dash=(5,4));c.create_text(tx,81,text='목표',anchor='n',fill='#b07a1d',font=('맑은 고딕',9))
        c.configure(scrollregion=(0,0,width,height))
    def draw_distribution(self):
        if not hasattr(self,'distribution'):return
        c=self.distribution;c.delete('all');w=max(c.winfo_width(),700);h=max(c.winfo_height(),350)
        c.create_text(35,32,text='프로젝트 완료일 분포',anchor='w',font=('맑은 고딕',17,'bold'),fill=NAVY)
        if not self.result:
            c.create_text(w/2,h/2,text='분석 실행 후 완료일 빈도 분포를 표시합니다.',fill=MUTED);return
        a=self.result['total'];bins=min(30,max(1,int(a.max()-a.min()+1)));counts,edges=np.histogram(a,bins=bins);lo=edges[0];hi=max(lo+1,edges[-1]);bottom=h-65;top=80;left=65;right=w-45;x=lambda d:left+(d-lo)/(hi-lo)*(right-left);peak=max(counts)
        for fraction in [0,.25,.5,.75,1]:
            y=bottom-fraction*(bottom-top);c.create_line(left,y,right,y,fill='#e2e8f0');c.create_text(left-8,y,text=str(int(peak*fraction)),anchor='e',fill=MUTED)
        for i,v in enumerate(counts):c.create_rectangle(x(edges[i])+1,bottom-v/peak*(bottom-top),x(edges[i+1])-1,bottom,fill='#508bea',outline='')
        for q,color in [(50,BLUE),(80,GREEN),(90,'#7656b5')]:
            d=percentile(a,q);px=x(d);c.create_line(px,top,px,bottom,fill=color,dash=(4,3),width=2);c.create_text(px,top-15-(q-50)/40*15,text=f'P{q} {textday(d)}',fill=color,font=('Segoe UI',10,'bold'))
        for i in range(6):
            d=lo+(hi-lo)*i/5;c.create_text(x(d),bottom+20,text=textday(d),fill=MUTED,font=('Segoe UI',9))
        c.create_text(left,h-20,text=f'{len(a):,}회 / 독립 삼각분포 / 빈도는 해당 구간에 완료된 시뮬레이션 수',anchor='w',fill=MUTED)
    def canvas_click(self,event):
        y=self.canvas.canvasy(event.y)
        for gy,m in self.group_y.items():
            if gy<=y<gy+36:
                if m in self.collapsed:self.collapsed.remove(m)
                else:self.collapsed.add(m)
                self.draw();break
    def canvas_edit(self,event):
        y=self.canvas.canvasy(event.y)
        for yy,k in self.rowids.items():
            if abs(yy-y)<18:
                i=next(i for i,t in enumerate(self.p['tasks']) if t['id']==k);self.edit_task(i);break
    def close(self):
        if self.busy:
            if not messagebox.askyesno('종료','계산을 중단하고 종료할까요?'):return
            self.cancel.set()
        else:self.apply_dates()
        self.store.add(self.p,'종료 저장','프로그램 종료');self.store.db.close();self.destroy()
GUIDE='''사용 순서\n\n1. 예제 프로젝트에서 [시뮬레이션 실행]을 누릅니다.\n2. 확률 기준을 P50~P95로 바꾸고 일정 범위 변화를 확인합니다.\n3. 공통 WBS 또는 기존 공통 Rev1 엑셀을 가져옵니다.\n4. 모듈 적용 탭에서 검토 필요 항목을 적용/제외로 결정합니다.\n5. 작업 편집에서 작업을 더블클릭해 담당자, 선행 ID, O/M/P를 수정합니다.\n6. 입력·결과 엑셀 내보내기 → Excel에서 편집 → 다시 가져오기를 반복합니다.\n\n계산 규칙\n\n• O=낙관 최소기간, M=최빈기간, P=비관 최대기간. 독립 삼각분포에서 추출 후 올림합니다.\n• FS + 0 선행관계만 지원합니다. 자원·Velocity·Tracks·SS/FF·Lag는 이번 초안에 없습니다.\n• 계획: 선행 완료일 / 분석 기준일 / 착수 제한일 중 가장 늦은 날부터 시작합니다.\n• 진행: 실제 시작일은 고정하고 기준일 이후 잔여 O/M/P를 사용합니다.\n• 완료: 실제 시작·완료일에 고정. 보류 및 검토 필요 모듈은 계산을 막습니다.\n• 제외: 기간 0으로 처리하며 선행관계를 유지합니다. 승인 등 마일스톤은 O=M=P=0입니다.\n• 경과일 기준입니다. 10/1 시작 1달력일 작업은 10/2에 완료합니다.\n• 근무일은 월~금 및 수동 등록한 휴일 목록을 반영합니다. 주말 착수는 다음 근무일로 이동합니다.\n• 초기에 제공하는 기간은 설명용·미보정 가정값입니다. 실제 사업 실적으로 검토하세요.\n\n화면 읽기\n\n얇은선: 시작 P10부터 완료 선택 P까지. 파랑: 시작 P50~선택 P, 초록: 완료 P50~선택 P.\n각 막대는 개별 작업의 주변분포를 요약하므로 막대 전체가 하나의 실행 가능한 일정은 아닙니다.\n모듈명을 클릭하면 접고 펼칠 수 있고, 작업행을 더블클릭하면 상세 편집창을 엽니다.\nP80은 입력 가정 아래 시뮬레이션 약 80%가 해당 날짜까지 완료됨을 의미합니다.\n\n엑셀 교환\n\n공통 Rev1과 프로그램 전용 교환 XLSX를 지원합니다. 임의의 다른 서식 자동 매핑은 미지원입니다.\n전용 양식: 프로젝트 / 모듈 / 작업목록 / 휴일 / 사용안내 (+ 분석요약 / 작업결과).\n열 이름은 유지하고 행 추가/정렬은 가능합니다. 기간 칸은 수식 대신 숫자로 입력합니다.\n내보내기는 원본 전체 서식을 복제하지 않습니다. 원본의 문서대장 등 부가 시트는 원본에서 관리하세요.\n결과 시트는 마지막 분석 시점 스냅샷이며 엑셀에서 입력을 바꿔도 자동 재계산되지 않습니다.\n\n저장 · 이력\n\n작업 수정은 변경 사유와 입력 스냅샷을 SQLite에 보관합니다. 분석 시 입력/요약/Seed/버전을 저장합니다.\n기준선 저장은 승인 시점 입력을 별도 이벤트로 남깁니다. 이력 탭에서 입력을 복원할 수 있습니다.\n로컬 DB 위치: %LOCALAPPDATA%\\ProjectScheduleLab\\history.sqlite3\n백업은 프로그램 종료 후 이 DB를 복사하세요. 팀 동시 편집과 서버 동기화는 지원하지 않습니다.\n\n참조\n\n시각 구성은 Bitovi jira-auto-scheduler의 확률 로드맵 화면을 참고했습니다.\nhttps://github.com/bitovi/jira-auto-scheduler\n원본 코드와 로고는 포함하지 않았으며, Python O/M/P 엔진으로 독립 구현했습니다.\n'''
if __name__=='__main__':
    try:App().mainloop()
    except Exception:
        (DATA/'error.log').write_text(traceback.format_exc(),encoding='utf-8');raise
