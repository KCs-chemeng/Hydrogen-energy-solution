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
from storage import Store
from model import canonical
from ui_v2 import Workflows
class App(Workflows,tk.Tk):
    def __init__(self):
        super().__init__();self.title('공정표 | Project Schedule Lab v'+VERSION);self.geometry('1500x900');self.minsize(1100,680);self.configure(bg=BG)
        self.option_add('*Font',('맑은 고딕',10));self.style=ttk.Style(self);self.style.theme_use('clam')
        self.style.configure('.',font=('맑은 고딕',10));self.style.configure('TFrame',background=BG);self.style.configure('TLabel',background=BG,foreground=NAVY);self.style.configure('TButton',padding=(12,7),background='#ffffff',foreground=NAVY);self.style.map('TButton',background=[('active','#dbeafe')]);self.style.configure('Primary.TButton',background=BLUE,foreground='white');self.style.configure('Treeview',rowheight=32,background='white',fieldbackground='white',borderwidth=0);self.style.configure('Treeview.Heading',background='#e8eef5',foreground=NAVY,padding=9);self.style.map('Treeview',background=[('selected','#dbeafe')],foreground=[('selected',NAVY)]);self.style.configure('TNotebook.Tab',padding=(24,12))
        self.store=Store(DATA/'history.sqlite3');self.p=demo();self.result=None;self.busy=False;self.events=queue.Queue();self.cancel=threading.Event();self.collapsed=set();self.rowids={};self.source_path=None
        previous=self.store.projects()
        self.p=self.store.load(previous[0][0]) if previous else canonical(self.p)
        if not previous:self.store.add(self.p,'프로젝트 생성','초기 예제')
        self.protocol('WM_DELETE_WINDOW',self.close)
        self.build();self.build_v2();self.sync();self.poll_id=self.after(100,self.poll)
    def build(self):
        header=tk.Frame(self,bg=NAVY,height=70);header.pack(fill='x');header.pack_propagate(False)
        tk.Label(header,text='▥  공정표',font=('맑은 고딕',21,'bold'),bg=NAVY,fg='white').pack(side='left',padx=24)
        tk.Label(header,text='PROJECT SCHEDULE LAB   /   PYTHON DESKTOP',font=('Segoe UI',10),bg=NAVY,fg='#b6c9dd').pack(side='left')
        tk.Label(header,text='v'+VERSION,bg=NAVY,fg='#98e1c6').pack(side='right',padx=25)
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
        ttk.Label(self.roadtab,text='  파랑: 선택 P 시작   초록: 선택 P 완료   실선·연결선: 선택 P 기준   회색: 시작 P10~완료 선택 P 참고 범위 (개별 분위수 요약)',foreground=MUTED,padding=(10,4)).pack(fill='x')
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
    def select_task(self):
        s=self.tree.selection();return int(s[0]) if s else None
    def edit_selected(self):
        i=self.select_task()
        if i is not None:self.edit_task(i)
    def copy_task(self):
        i=self.select_task()
        if i is None or self.busy:return
        self.edit_task(None,self.p['tasks'][i])
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
                for edge in self.p['relations']:
                    if edge['succ']!=t['id']:continue
                    pred=edge['pred']
                    if pred in ys:
                        a=x(percentile(r['starts' if edge['type']=='SS' else 'ends'][pred],q));b=x(percentile(r['ends' if edge['type']=='FF' else 'starts'][t['id']],q));c.create_line(a,ys[pred],a+14,ys[pred]+8,b-14,ys[t['id']]-8,b,ys[t['id']],smooth=True,fill='#bdcbdc',arrow=tk.LAST,width=1,tags=('relation',f"relation:{pred}:{t['id']}:{edge['type']}"))
        for t in visible:
            k=t['id']
            if k not in ys:continue
            yy=ys[k];c.create_line(0,yy+20,width,yy+20,fill='#f0f3f7');label=t['name'];label=label if len(label)<34 else label[:33]+'…';c.create_text(24,yy-7,text=label,anchor='w',fill=NAVY,font=('맑은 고딕',10));c.create_text(24,yy+9,text=f'{k}   ·   {t["state"]}   ·   {t.get("calendar","달력일")}',anchor='w',fill=MUTED,font=('맑은 고딕',8))
            start_day=percentile(r['starts'][k],q);finish_day=percentile(r['ends'][k],q)
            sq,fq=x(start_day),x(finish_day);s10=x(percentile(r['starts'][k],10))
            c.create_line(s10,yy,fq,yy,fill='#d9e2ed',width=3,tags=('uncertainty',f'uncertainty:{k}'))
            c.create_line(sq,yy,fq,yy,fill='#3387c6',width=6,tags=('schedule',f'schedule:{k}'))
            c.create_rectangle(sq-3,yy-8,sq+3,yy+8,fill=BLUE,outline='',tags=('task-start',f'start:{k}'))
            c.create_rectangle(fq-3,yy-8,fq+3,yy+8,fill=GREEN,outline='',tags=('task-end',f'end:{k}'))
            if start_day==finish_day:
                c.create_polygon(sq,yy-8,sq+7,yy,sq,yy+8,sq-7,yy,fill=GREEN,outline=BLUE,tags=('milestone',f'milestone:{k}'))
            c.create_text((sq+fq)/2,yy-13,text=f'P{q} {textday(start_day)[5:]} → {textday(finish_day)[5:]}',fill=MUTED,font=('Segoe UI',8),tags=('schedule-date',f'dates:{k}'))
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
GUIDE='로드맵의 시작·완료 표식과 연결선은 선택한 확률 기준을 함께 사용합니다. 개별 작업 분위수 요약이며 단일 실행 일정은 아닙니다.\n\n프로젝트 메뉴: 생성·열기·저장·복제·DB 백업·복원\n일정 · 데이터 메뉴: 관계·달력·오류 검토·M 기준 재계산·기준 일정 조회\n\n엑셀은 비교 후 선택 적용합니다. 빈 셀은 보존, __CLEAR__는 삭제, 누락 행은 보존합니다. 충돌은 기본 미선택입니다. 관계 시트가 선행 ID 열보다 우선합니다.\n\nFS/SS/FF와 0 이상 정수 시차 지원. 시작부터 완료까지 경과일 방식이며 1일 작업은 다음 날 완료합니다. 기본 주5일, 추가 근무일이 휴일보다 우선합니다. 공휴일은 직접 관리합니다.\n\n기준일은 당일 시작 시점입니다. 진행 잔여기간은 기준일부터 계산하며 실제 시작은 고정합니다. 실적 관계 위반은 경고하고 미완료 FS는 잔여 작업을 지연합니다. 보류는 재개일 없으면 차단합니다.\n\n모듈/작업 제외 시 관계는 보존하며 연결 검토 전 계산을 차단합니다. 무조건 0일로 재연결하지 않습니다. 고정 날짜 충돌은 계산 실패입니다. 목표일은 강제하지 않습니다.\n\n기준 일정은 M 기준 계산의 승인 버전입니다. 재계산·Import·입력 복원으로 승인 버전이 바뀌지 않습니다. 총 여유는 목표일까지 차이와 별개이며 달력일 단위입니다. 실적 고정 작업은 여유 미정의입니다.\n\n표준 Export는 원본 임의 시트·서식·그림을 복제하지 않습니다. 전체 이력은 DB에 있습니다. 자세한 이전/검수 절차는 README_KO.md를 확인하세요.\n'
if __name__=='__main__':
    try:App().mainloop()
    except Exception:
        (DATA/'error.log').write_text(traceback.format_exc(),encoding='utf-8');raise
