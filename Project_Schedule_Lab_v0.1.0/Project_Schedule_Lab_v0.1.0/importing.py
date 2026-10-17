"""Read/compare/apply pipeline. Never deletes missing tasks or silently resolves conflicts."""
from dataclasses import dataclass, field
from copy import deepcopy
from pathlib import Path
from model import canonical, TASK_DEFAULTS, uid, sync_pred
from excel_io import read_document, Issue, CLEAR
from engine import inspect


@dataclass
class Change:
    key: str
    task: str
    field: str
    before: object
    after: object
    kind: str = '수정'
    conflict: bool = False


@dataclass
class Plan:
    current: dict
    filename: str
    changes: list = field(default_factory=list)
    issues: list = field(default_factory=list)
    missing: list = field(default_factory=list)
    same: list = field(default_factory=list)
    incoming: dict = field(default_factory=dict)

    def report(self):
        return dict(new=sum(c.kind=='신규' for c in self.changes),changed=len(self.changes),
                    unchanged=self.same,missing=self.missing,conflicts=[c.key for c in self.changes if c.conflict],
                    issues=[str(i) for i in self.issues])


def empty_for_document(doc):
    p=doc.project
    return canonical(dict(uid=p.get('uid') or uid(),id=p.get('id') or 'PJ-'+uid()[:8],
                          name=p.get('name') or Path(doc.filename).stem,start='',asof='',target='',tasks=[],modules={}))


def prepare(path,current=None,store=None):
    doc=read_document(path)
    current=canonical(current) if current is not None else empty_for_document(doc)
    p=doc.project
    plan=Plan(deepcopy(current),str(path),issues=list(doc.issues),incoming=p)
    def issue(message,task='',field='id',severity='오류',fix='수정 후 다시 비교하세요'):
        s,row,col,value=doc.cells.get((task,field),('프로젝트',1,'A',task))
        plan.issues.append(Issue(severity,str(path),s,row,col,value,message,fix))
    if not p: return plan
    if p.get('uid') and p['uid'] != current['uid']: issue('다른 프로젝트 UID입니다. 신규 프로젝트로 가져오세요.')
    if p.get('id') and p['id'] != current['id']: issue('프로젝트 ID 불일치. 신규로 가져오거나 올바른 파일을 선택하세요.')
    version=doc.export_version
    base=None
    if version is not None:
        try:
            version=int(version)
            base=store.load(current['uid'],version) if store else None
            if version != current['data_version'] and not base:
                issue('Export 기준 버전이 없습니다. 다른 값은 모두 충돌로 표시합니다.',severity='경고')
        except (ValueError,TypeError): issue('잘못된 Export 데이터 버전')
    elif current['tasks']:
        issue('Export 버전 없는 외부 파일. 기존 값 변경은 명시적 선택이 필요합니다.',severity='경고')
    def add(task,key,a,b):
        if a==b: return
        conflict=False
        if base is not None:
            if task=='$project': old=base.get(key)
            else: old=next((t.get(key) for t in base['tasks'] if t['id']==task),None)
            if b==old: return  # Excel unchanged: preserve later local edits.
            conflict=a!=old
        elif current['tasks'] and (version is None or version != current['data_version']): conflict=True
        plan.changes.append(Change(f'{task}:{key}',task,key,deepcopy(a),deepcopy(b),conflict=conflict))
    old={t['id']:t for t in current['tasks']}
    incoming_ids=set()
    for i,raw in enumerate(p.get('tasks',[])):
        raw=deepcopy(raw)
        k=raw.get('id')
        if not k:
            if current['tasks']: issue('Activity ID 없는 행은 기존 작업에 매핑할 수 없습니다. ID를 지정하거나 신규 프로젝트로 등록하세요.');continue
            k='T-'+uid()[:12];raw['id']=k
            issue('ID 없는 행을 신규 작업으로 등록합니다: '+k,severity='경고')
        if k in incoming_ids: issue('중복 Activity ID',k);continue
        incoming_ids.add(k)
        if k in old:
            if raw.get('uid') and raw['uid']!=old[k]['uid']: issue('기존 작업 UID 불일치. ID 자동 병합 금지',k,'uid');continue
            count=len(plan.changes)
            for key,v in raw.items():
                if key in ('id','uid') or (key=='pred' and doc.explicit_relations): continue
                if v in (None,'') or v==[]: continue
                after=(None if key in ('o','m','p','ro','rm','rp','progress') else [] if key=='pred' else '') if v==CLEAR else v
                add(k,key,old[k].get(key),after)
            if len(plan.changes)==count: plan.same.append(k)
        else:
            if raw.get('uid') and any(t['uid']==raw['uid'] for t in old.values()): issue('기존 UID의 Activity ID 변경은 지원하지 않습니다.',k);continue
            t=deepcopy(TASK_DEFAULTS);t.update(id=k,uid=raw.get('uid') or uid())
            for key,v in raw.items():
                if v not in (None,'',CLEAR): t[key]=v
            if doc.explicit_relations: t['pred']=[]
            plan.changes.append(Change(k+':$new',k,'$new',None,t,kind='신규'))
    plan.missing=sorted(set(old)-incoming_ids)
    if plan.missing: issue('파일에서 누락된 업무는 보존합니다: '+', '.join(plan.missing),severity='경고')
    for key in ('name','start','asof','target','modules','holidays','workdays','calendars'):
        value=p.get(key)
        if key in doc.metadata.get('clear_sections',[]): value=[]
        elif value in (None,'',[],{}): continue
        add('$project',key,current.get(key),'' if value==CLEAR else value)
    if doc.explicit_relations:
        # Removing a relation requires explicit selection of this visible change, unlike missing tasks.
        candidate_rel=deepcopy(p['relations'])
        for i,r in enumerate(candidate_rel):
            matches=[old_r for old_r in current['relations'] if old_r['pred']==r.get('pred') and old_r['succ']==r.get('succ') and (old_r['type']==r.get('type') or not r.get('type'))]
            if len(matches)==1:
                for key in r:
                    raw_cell=doc.cells.get(('$relation'+str(i),key),('',0,'',None))[3]
                    if raw_cell in (None,''): r[key]=deepcopy(matches[0].get(key,r[key]))
                    elif raw_cell==CLEAR: r[key]=''
            for key,v in dict(type='FS',lag=0,lagUnit='달력일',lagCalendar='default',note='',deliverable='').items():
                if r.get(key) in (None,''): r[key]=v
        add('$project','relations',current['relations'],candidate_rel)
    else:
        # Compatibility column replaces only the incoming task's FS+0 edges, with an explicit review row.
        rel=deepcopy(current['relations'])
        for raw in p.get('tasks',[]):
            k=raw.get('id')
            if raw.get('pred') in (None,[], ''): continue
            rel=[r for r in rel if r['succ']!=k]
            for a in ([] if raw['pred']==CLEAR else raw['pred']):
                rel.append(dict(pred=a,succ=k,type='FS',lag=0,lagUnit='달력일',lagCalendar='default',note='',deliverable=''))
        # New tasks' pred values are also represented in this shared authoritative relation list.
        if rel!=current['relations']: add('$project','relations',current['relations'],rel)
        plan.changes=[c for c in plan.changes if c.field!='pred']
    # Validate the whole proposed file, even when it will not be applied. Unselected subsets revalidated later.
    candidate=_apply(plan,{c.key for c in plan.changes})
    errors,warnings,_=inspect(candidate,scheduling=False)
    for msg in errors:
        task=next((k for k in incoming_ids if msg.startswith(k+':') or msg.startswith(k+'.')), '')
        key=next((k for k in TASK_DEFAULTS if msg.startswith(task+'.'+k+':')), 'id')
        if msg.startswith('프로젝트.'):
            task='$project';key=msg.split('.',1)[1].split(':',1)[0]
        elif msg.startswith('관계 ') or ' → ' in msg:
            idx=next((i for i,r in enumerate(p.get('relations',[])) if str(r.get('pred'))+' → '+str(r.get('succ')) in msg),None)
            if idx is not None:
                task='$relation'+str(idx);key='lag' if '시차' in msg else 'type' if '관계 ' in msg and '미지원' in msg else 'pred'
        elif '상태' in msg and task: key='state'
        elif '진척률' in msg and task: key='progress'
        elif '달력' in msg and task: key='calendarId'
        elif '실제 시작' in msg and task: key='actualStart'
        elif '실제 완료' in msg and task: key='actualEnd'
        elif 'O ≤ M' in msg: key='ro' if any(t['id']==task and t.get('state')=='진행' for t in candidate['tasks']) else 'o'
        issue(msg,task,key)
    for msg in warnings:
        module=next((m for m in candidate['modules'] if msg.startswith(m+':')),None)
        issue(msg,'$module' if module else '',module or 'id',severity='경고',fix='모듈을 적용/제외로 결정한 후 계산하세요' if module else '경고 내용을 검토하고 확인 후 진행하세요')
    return plan


def _apply(plan,selected):
    p=deepcopy(plan.current)
    tasks={t['id']:t for t in p['tasks']}
    for c in plan.changes:
        if c.key not in selected: continue
        if c.field=='$new':
            t=deepcopy(c.after);p['tasks'].append(t);tasks[t['id']]=t
        elif c.task=='$project': p[c.field]=deepcopy(c.after)
        else: tasks[c.task][c.field]=deepcopy(c.after)
    sync_pred(p)
    return canonical(p)


def materialize(plan,selected=None,acknowledge=False):
    if any(i.severity=='오류' for i in plan.issues): raise ValueError('\n'.join(str(i) for i in plan.issues if i.severity=='오류'))
    if any(i.severity=='경고' for i in plan.issues) and not acknowledge: raise ValueError('경고 확인 후 적용하세요.')
    if selected is None:
        if any(c.conflict for c in plan.changes): raise ValueError('수정 충돌: 적용할 필드를 명시적으로 선택하세요.')
        selected={c.key for c in plan.changes}
    selected=set(selected)
    if not selected <= {c.key for c in plan.changes}: raise ValueError('알 수 없는 변경 선택')
    p=_apply(plan,selected)
    errors,_,_=inspect(p,scheduling=False)
    if errors: raise ValueError('\n'.join(errors))
    return p
