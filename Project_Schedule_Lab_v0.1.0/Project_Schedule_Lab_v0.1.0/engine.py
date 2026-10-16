"""Independent O/M/P Monte Carlo engine. No Jira or Bitovi code dependency."""
from datetime import date, timedelta
import math
import numpy as np
VERSION='0.1.0'
def day(value):
    if not value:return None
    return date.fromisoformat(str(value)[:10]).toordinal()
def textday(value):return date.fromordinal(int(value)).isoformat()
def validate(p):
    errors=[];tasks=p.get('tasks',[]);ids={};order=[];state={}
    for k in ('start','asof'):
        try:
            if day(p.get(k)) is None:raise ValueError()
        except Exception:errors.append(f'{k}: YYYY-MM-DD 날짜를 입력하세요.')
    try:
        if p.get('target'):day(p['target'])
    except Exception:errors.append('목표일: YYYY-MM-DD 날짜를 입력하세요.')
    try:
        if day(p.get('asof'))<day(p.get('start')):errors.append('분석 기준일이 프로젝트 시작일보다 빠릅니다.')
    except Exception:pass
    if not tasks:errors.append('작업이 없습니다.')
    if len(tasks)>2000:errors.append('최대 2,000개 작업을 지원합니다.')
    for t in tasks:
        k=t.get('id','')
        if not k or k in ids:errors.append(f'작업 ID 누락/중복: {k}')
        ids[k]=t
        if not isinstance(t.get('pred'),list):errors.append(f'{k}: 선행 ID 목록 오류');continue
        if t.get('state') not in ['계획','진행','완료','보류','제외']:errors.append(f'{k}: 상태 오류')
        mod=p.get('modules',{}).get(t.get('module'))
        if mod not in ['적용','제외']:errors.append(f'{t.get("module")}: 적용 여부 검토 필요')
        if mod=='제외' or t.get('state')=='제외':continue
        if t.get('state')=='보류':errors.append(f'{k}: 보류 해제 후 계산하세요.')
        if t.get('calendar','달력일') not in ['달력일','근무일']:errors.append(f'{k}: 달력 단위 오류')
        try:
            if t.get('constraint'):day(t['constraint'])
            if t['state'] in ['진행','완료']:
                s=day(t.get('actualStart'))
                if not s or s>day(p['asof']):raise ValueError()
            if t['state']=='완료':
                f=day(t.get('actualEnd'))
                if not f or f<s or f>day(p['asof']):raise ValueError()
        except Exception:errors.append(f'{k}: 실제 날짜/착수 제한일 오류')
        if t.get('state')!='완료':
            vals=[t.get(k) for k in (('ro','rm','rp') if t.get('state')=='진행' else ('o','m','p'))]
            if any(not isinstance(v,(int,float)) or not math.isfinite(v) or v<0 for v in vals) or not vals[0]<=vals[1]<=vals[2]:errors.append(f'{k}: 0 ≤ O ≤ M ≤ P 입력 필요 (진행은 잔여 기간)')
    # Iterative topological sort, preserving original order; excluded nodes keep edges.
    indeg={k:0 for k in ids};children={k:[] for k in ids}
    for k,t in ids.items():
        for pred in set(t.get('pred',[]) if isinstance(t.get('pred'),list) else []):
            if pred not in ids:errors.append(f'{k}: 없는 선행 ID {pred}')
            else:indeg[k]+=1;children[pred].append(k)
    queue=[k for k in ids if indeg[k]==0]
    for k in queue:
        order.append(ids[k])
        for c in children[k]:
            indeg[c]-=1
            if indeg[c]==0:queue.append(c)
    if len(order)!=len(ids):errors.append('선행관계에 순환이 있습니다.')
    for h in p.get('holidays',[]):
        try:day(h)
        except Exception:errors.append(f'휴일 날짜 오류: {h}')
    return list(dict.fromkeys(errors)),order

def simulate(p,n=5000,seed=42,progress=None,cancel=None):
    errors,order=validate(p)
    if errors:raise ValueError('\n'.join(errors))
    if not 1<=n<=20000:raise ValueError('반복 횟수는 1~20,000 범위입니다.')
    rng=np.random.default_rng(seed);starts={};ends={};base=day(p['start']);asof=day(p['asof'])
    epoch=date(1970,1,1).toordinal();holidays=np.array(p.get('holidays',[]),dtype='datetime64[D]')
    for idx,t in enumerate(order):
        if cancel and cancel.is_set():raise InterruptedError('계산이 취소되었습니다.')
        pre=np.full(n,base,dtype=np.int64)
        for key in t['pred']:pre=np.maximum(pre,ends[key])
        excluded=p['modules'][t['module']]=='제외' or t['state']=='제외'
        if excluded:s=pre;f=pre.copy()
        elif t['state']=='완료':
            s=np.full(n,day(t['actualStart']),dtype=np.int64);f=np.full(n,day(t['actualEnd']),dtype=np.int64)
            # Completed actuals may predate analysis start; verify predecessor actuals only.
            for pred in t['pred']:
                if np.any(ends[pred]>s):raise ValueError(t['id']+': 실제 착수일과 선행 완료일 충돌 (분석 시작일도 확인)')
        else:
            ongoing=t['state']=='진행';s=np.full(n,day(t['actualStart']),dtype=np.int64) if ongoing else np.maximum(pre,max(asof,day(t.get('constraint')) or base))
            if ongoing and np.any(pre>s):raise ValueError(t['id']+': 진행 작업의 선행작업이 실제 착수일까지 완료되지 않았습니다.')
            origin=np.full(n,asof,dtype=np.int64) if ongoing else s.copy()
            a,m,b=[t[k] for k in (('ro','rm','rp') if ongoing else ('o','m','p'))]
            duration=np.full(n,math.ceil(a),dtype=np.int64) if a==b else np.ceil(rng.triangular(a,m,b,n)).astype(np.int64)
            if t.get('calendar','달력일')=='근무일':
                dates=(origin-epoch).astype('datetime64[D]');rolled=np.busday_offset(dates,0,roll='forward',holidays=holidays)
                f=np.busday_offset(rolled,duration,roll='forward',holidays=holidays).astype(np.int64)+epoch
                f=np.where(duration==0,origin,f)
                if not ongoing:s=np.where(duration==0,s,rolled.astype(np.int64)+epoch)
            else:f=origin+duration
        starts[t['id']]=s;ends[t['id']]=f
        if progress:progress(int((idx+1)/len(order)*100))
    total=np.maximum.reduce(list(ends.values()))
    return {'starts':starts,'ends':ends,'total':total,'n':n,'seed':seed,'engine':VERSION}
def percentile(a,q):return int(np.quantile(a,q/100,method='higher'))
def summary(r,p):return {'P50':textday(percentile(r['total'],50)),'P80':textday(percentile(r['total'],80)),'P90':textday(percentile(r['total'],90)),'probability':float(np.mean(r['total']<=day(p['target']))) if p.get('target') else None,'iterations':r['n'],'seed':r['seed'],'engine':VERSION}
def demo():
    rows=[('A01','01 기획','사업범위 · 설계기준 확정',3,5,8,[]),('A02','01 기획','부지 조건 및 현황 확인',5,8,15,['A01']),('B01','02 설계','기본배치 · 공정설계',8,12,22,['A02']),('B02','02 설계','분야별 기본설계 검토',5,10,18,['B01']),('B03','02 설계','기본설계 승인',0,0,0,['B02']),('C01','03 인허가','신청도서 작성',7,12,20,['B03']),('C02','03 인허가','기관 검토 · 보완 대응',15,30,55,['C01']),('C03','03 인허가','승인 마일스톤',0,0,0,['C02']),('D01','04 구매','기자재 사양 · 발주',7,12,18,['B03']),('D02','04 구매','제작 · 검사 · 현장반입',25,40,65,['D01']),('E01','05 시공','시공도서 발행',5,10,15,['C03']),('E02','05 시공','설비 설치 및 배관공사',15,25,40,['E01','D02']),('F01','06 시운전','검사 · 시운전',7,12,20,['E02']),('F02','06 시운전','인수인계 완료',0,0,0,['F01'])]
    return {'id':'DEMO-001','name':'생산시설 구축 · 예제','start':'2026-10-01','asof':'2026-10-01','target':'2027-05-01','holidays':[],'modules':{r[1]:'적용' for r in rows},'tasks':[dict(zip(['id','module','name','o','m','p','pred'],r),state='계획',calendar='달력일',role='',deliverable='',ro=None,rm=None,rp=None,actualStart='',actualEnd='',constraint='',baselineStart='',baselineEnd='',note='') for r in rows]}
