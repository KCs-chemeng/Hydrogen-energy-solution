"""Shared vectorized FS/SS/FF scheduler for deterministic and Monte Carlo calculations."""
from datetime import date, datetime
from functools import lru_cache
import math
import numpy as np
from model import canonical

VERSION = '0.2.2'
MIN_DAY = date(1900, 1, 1).toordinal()
MAX_DAY = date(2200, 1, 1).toordinal()


def day(value):
    if value in (None, ''):
        return None
    if isinstance(value, datetime):
        value = value.date()
    result = (value if isinstance(value, date) else date.fromisoformat(str(value))).toordinal()
    if not MIN_DAY <= result < MAX_DAY:
        raise ValueError('지원 날짜 범위: 1900-01-01 ~ 2199-12-31')
    return result


def textday(value):
    return date.fromordinal(int(value)).isoformat()


@lru_cache(maxsize=64)
def _working(holidays, extra, weekmask):
    dates = np.arange(MIN_DAY, MAX_DAY, dtype=np.int64)
    mask = np.isin((dates - 1) % 7, weekmask)
    mask &= ~np.isin(dates, holidays)
    mask |= np.isin(dates, extra)
    return dates[mask]


class Calendar:
    def __init__(self, p, unit='달력일', key='default'):
        self.unit = unit
        cfg = p.get('calendars', {}).get(key, {})
        holidays = cfg.get('holidays', p.get('holidays', []) if key == 'default' else [])
        extra = cfg.get('workdays', p.get('workdays', []) if key == 'default' else [])
        self.working = _working(tuple(sorted(day(v) for v in holidays)),
                                tuple(sorted(day(v) for v in extra)), tuple(cfg.get('weekmask', [0, 1, 2, 3, 4])))

    def add(self, origin, duration):
        origin, duration = np.asarray(origin, dtype=np.int64), np.asarray(duration, dtype=np.int64)
        if self.unit == '달력일':
            result = origin + duration
        else:
            idx = np.searchsorted(self.working, origin, side='left') + duration
            if np.any((idx < 0) | (idx >= len(self.working))):
                raise ValueError('계산이 달력 지원 범위를 벗어났습니다.')
            result = np.where(duration == 0, origin, self.working[idx])
        if np.any((result < MIN_DAY) | (result >= MAX_DAY)):
            raise ValueError('계산 날짜 범위를 벗어났습니다.')
        return result

    def roll(self, origin, duration):
        if self.unit == '달력일':
            return np.asarray(origin)
        idx = np.searchsorted(self.working, origin)
        if np.any(idx >= len(self.working)):
            raise ValueError('달력 범위 초과')
        return np.where(np.asarray(duration) == 0, origin, self.working[idx])

    def earliest(self, finish_bound, duration):
        """Earliest valid start whose finish meets a lower bound (FF)."""
        if self.unit == '달력일':
            return np.asarray(finish_bound) - duration
        idx = np.searchsorted(self.working, finish_bound, side='left') - duration
        idx = np.maximum(idx,0)
        if np.any(idx >= len(self.working)):
            raise ValueError('달력 범위 초과')
        return np.where(np.asarray(duration) == 0, finish_bound, self.working[idx])

    def latest(self, finish_bound, duration):
        """Latest valid start whose finish is at most an upper bound (backward pass)."""
        if self.unit == '달력일':
            return np.asarray(finish_bound) - duration
        idx = np.searchsorted(self.working, finish_bound, side='right') - 1 - duration
        if np.any((idx < 0) | (idx >= len(self.working))):
            raise ValueError('달력 범위 초과')
        return np.where(np.asarray(duration) == 0, finish_bound, self.working[idx])


def relations(p):
    if 'relations' in p:
        return p['relations']
    return [dict(pred=a, succ=t['id'], type='FS', lag=0, lagUnit='달력일', lagCalendar='default')
            for t in p['tasks'] for a in t.get('pred', [])]


def exclusion_impact(p):
    excluded = {t['id'] for t in p['tasks'] if t.get('state') == '제외' or p['modules'].get(t['module']) == '제외'}
    return [r for r in relations(p) if r['pred'] in excluded or r['succ'] in excluded]


def inspect(p, scheduling=True):
    """Return errors, warnings, topological order. Drafts may contain unresolved schedule inputs."""
    errors, warnings = [], []
    tasks = p.get('tasks', [])
    ids = {}
    def error(message): errors.append(message)
    for k in ('start', 'asof', 'target'):
        try:
            if day(p.get(k)) is None and k != 'target': raise ValueError()
        except (ValueError, TypeError): error(f'프로젝트.{k}: YYYY-MM-DD 날짜 필요')
    try:
        if day(p['asof']) < day(p['start']): error('기준일이 프로젝트 시작일보다 빠릅니다.')
    except (ValueError, TypeError, KeyError): pass
    if not tasks and scheduling: error('작업이 없습니다.')
    if len(tasks) > 2000: error('최대 2,000개 작업 지원')
    task_uids = set()
    calendar_ids = {'default'} | set(p.get('calendars', {}))
    for key in calendar_ids:
        cfg = p.get('calendars', {}).get(key, {})
        mask = cfg.get('weekmask', [0, 1, 2, 3, 4])
        if not mask or any(type(v) is not int or v not in range(7) for v in mask): error(f'달력 {key}: weekmask는 0(월)~6(일) 목록')
        for field in ('holidays', 'workdays'):
            for v in cfg.get(field, p.get(field, []) if key == 'default' else []):
                try:
                    if day(v) is None: raise ValueError()
                except (ValueError, TypeError): error(f'달력 {key}.{field}: 날짜 오류 {v}')
    for t in tasks:
        k = t.get('id', '')
        if not k or k in ids: error(f'작업 ID 누락/중복: {k}')
        ids[k] = t
        if t.get('uid'):
            if t['uid'] in task_uids: error(f'{k}: 내부 UID 중복')
            task_uids.add(t['uid'])
        if not t.get('name'): error(f'{k}: 작업명 필요')
        state = {'미착수':'계획', '진행 중':'진행'}.get(t.get('state'), t.get('state'))
        if state not in ('계획', '진행', '완료', '보류', '제외'): error(f'{k}: 상태 오류')
        module = p.get('modules', {}).get(t.get('module'))
        if module not in ('적용', '제외', '검토 필요'): error(f'{k}: 모듈 정의 오류')
        if module == '검토 필요':
            if scheduling: error(f'{t.get("module")}: 적용 여부 검토 필요')
            else: warnings.append(f'{t.get("module")}: 적용 여부 검토 필요; 계산 전 결정하세요.')
        if t.get('calendar', '달력일') not in ('달력일', '근무일'): error(f'{k}: 기간단위 오류')
        if t.get('calendarId', 'default') not in calendar_ids: error(f'{k}: 없는 달력')
        for f in ('actualStart','actualEnd','constraint','baselineStart','baselineEnd','fixedStart','fixedEnd','resume'):
            try: day(t.get(f))
            except (ValueError, TypeError): error(f'{k}.{f}: 날짜 오류 {t.get(f)}')
        try:
            s, f, a = day(t.get('actualStart')), day(t.get('actualEnd')), day(p.get('asof'))
            if state in ('진행', '완료') and (s is None or (a and s > a)): error(f'{k}: 기준일 이전 실제 시작 필요')
            if state == '완료' and (f is None or (s and f < s) or (a and f > a)): error(f'{k}: 실제 완료일 모순')
            if state in ('계획', '보류') and (s or f): error(f'{k}: 미착수/보류 상태와 실제 날짜 모순')
            if state == '진행' and f: error(f'{k}: 진행 중 실제 완료일 입력됨')
        except (ValueError, TypeError): pass
        progress = t.get('progress')
        if progress is not None:
            if not isinstance(progress, (int,float)) or not math.isfinite(progress) or not 0 <= progress <= 100: error(f'{k}: 진척률 0~100 필요')
            elif (state == '완료' and progress != 100) or (state in ('계획','보류') and progress != 0) or (state == '진행' and progress == 100): error(f'{k}: 상태/진척률 모순')
        excluded = module == '제외' or state == '제외'
        if not excluded and scheduling and state == '보류' and not t.get('resume'): error(f'{k}: 보류 재개 가능일 필요 (미확정 일정)')
        if not excluded and state != '완료':
            keys = ('ro','rm','rp') if state == '진행' else ('o','m','p')
            vals = [t.get(f) for f in keys]
            if not scheduling and all(v is None for v in vals):
                warnings.append(f'{k}: 기간 미입력, 계산 전 확인 필요')
            elif any(type(v) not in (int,float) or not math.isfinite(v) or v < 0 for v in vals) or not vals[0] <= vals[1] <= vals[2]:
                error(f'{k}: 0 ≤ O ≤ M ≤ P 필요 (진행은 잔여 O/M/P)')
    edges = relations(p)
    indeg = {k:0 for k in ids}
    children = {k:[] for k in ids}
    seen = set()
    for r in edges:
        a, b = r.get('pred'), r.get('succ')
        if a not in ids or b not in ids: error(f'관계 {a} → {b}: 존재하지 않거나 다른 프로젝트 작업 참조'); continue
        if r.get('project_uid') and r['project_uid'] != p.get('uid'): error(f'{a} → {b}: 다른 프로젝트 관계')
        if a == b: error(f'{a}: 자기 참조')
        typ = r.get('type','FS')
        if typ not in ('FS','SS','FF'): error(f'{a} → {b}: 미지원 관계 {typ}')
        lag = r.get('lag',0)
        if type(lag) not in (int,float) or not math.isfinite(lag) or lag < 0 or int(lag) != lag: error(f'{a} → {b}: 시차는 0 이상의 정수')
        if r.get('lagUnit','달력일') not in ('근무일','달력일') or r.get('lagCalendar','default') not in calendar_ids: error(f'{a} → {b}: 시차 달력 오류')
        key = (a,b,typ)
        if key in seen: error(f'{a} → {b}: {typ} 관계 중복')
        seen.add(key)
        indeg[b] += 1
        children[a].append(b)
    queue = [k for k in ids if indeg[k] == 0]
    order = []
    for k in queue:
        order.append(ids[k])
        for c in children[k]:
            indeg[c] -= 1
            if indeg[c] == 0: queue.append(c)
    if len(order) != len(ids):
        # Iterative DFS returns a concrete cycle, without recursion limits on large WBS.
        active, done, path = set(), set(), []
        cycle = None
        for root in ids:
            if root in done: continue
            stack = [(root, iter(children[root]))]; active.add(root); path.append(root)
            while stack and cycle is None:
                node, it = stack[-1]
                nxt = next(it, None)
                if nxt is None:
                    stack.pop(); active.remove(node); done.add(node); path.pop()
                elif nxt in active:
                    cycle = path[path.index(nxt):] + [nxt]
                elif nxt not in done:
                    stack.append((nxt, iter(children[nxt]))); active.add(nxt); path.append(nxt)
            if cycle: break
        error('순환: ' + ' → '.join(cycle or []))
    impact = exclusion_impact(p)
    if impact and scheduling:
        error('제외 작업 연결 검토 필요: ' + '; '.join(f"{r['pred']} → {r['succ']} ({r.get('type','FS')})" for r in impact))
    elif impact:
        warnings.append('제외 작업과 연결된 관계가 보존되어 있습니다. 명시적으로 관계를 검토/수정하기 전 계산 차단')
    return list(dict.fromkeys(errors)), list(dict.fromkeys(warnings)), order


def validate(p):
    errors, _, order = inspect(p)
    return errors, order


def _run(p, n=1, seed=42, stochastic=False, progress=None, cancel=None):
    if type(seed) is not int or not 0 <= seed <= 2**32-1: raise ValueError('Seed는 0~4294967295 정수')
    p = canonical(p, identities=False)
    errors, warnings, order = inspect(p)
    if errors: raise ValueError('\n'.join(errors))
    rng = np.random.default_rng(seed)
    starts, ends, durations = {}, {}, {}
    active = []
    base, asof = day(p['start']), day(p['asof'])
    incoming = {t['id']:[] for t in order}
    for r in relations(p): incoming[r['succ']].append(r)
    for idx, t in enumerate(order):
        if cancel and cancel.is_set(): raise InterruptedError('계산 취소')
        k, state = t['id'], t['state']
        cal = Calendar(p, t['calendar'], t.get('calendarId','default'))
        if p['modules'][t['module']] == '제외' or state == '제외':
            starts[k] = ends[k] = np.full(n, base, dtype=np.int64)
            durations[k] = np.zeros(n, dtype=np.int64)
            continue
        active.append(k)
        ongoing = state == '진행'
        duration = np.zeros(n, dtype=np.int64)
        if state != '완료':
            a,m,b = [t[f] for f in (('ro','rm','rp') if ongoing else ('o','m','p'))]
            duration = np.full(n, math.ceil(m), dtype=np.int64) if not stochastic or a == b else np.ceil(rng.triangular(a,m,b,n)).astype(np.int64)
        durations[k] = duration
        lower_start = np.full(n, max(base,asof,day(t.get('constraint')) or base,day(t.get('resume')) or base),dtype=np.int64)
        lower_finish = np.full(n,base,dtype=np.int64)
        historical = []
        for r in incoming[k]:
            typ = r['type']
            anchor = ends[r['pred']] if typ in ('FS','FF') else starts[r['pred']]
            bound = Calendar(p,r['lagUnit'],r['lagCalendar']).add(anchor,int(r['lag']))
            if typ == 'FF': lower_finish = np.maximum(lower_finish,bound)
            else: lower_start = np.maximum(lower_start,bound)
            historical.append((r,bound))
        if state == '완료':
            s = np.full(n,day(t['actualStart']),dtype=np.int64)
            f = np.full(n,day(t['actualEnd']),dtype=np.int64)
        elif ongoing:
            s = np.full(n,day(t['actualStart']),dtype=np.int64)
            origin = np.full(n,asof,dtype=np.int64)
            for r,bound in historical:
                if r['type'] == 'FS': origin = np.maximum(origin,bound)
            origin = np.maximum(origin,cal.earliest(lower_finish,duration))
            f = cal.add(origin,duration)
        else:
            s = cal.roll(np.maximum(lower_start,cal.earliest(lower_finish,duration)),duration)
            if t.get('fixedStart'):
                fixed = day(t['fixedStart'])
                if np.any(s > fixed) or np.any(cal.roll(fixed,duration) != fixed): raise ValueError(f'{k}: 고정 시작일과 관계/달력 충돌')
                s = np.full(n,fixed,dtype=np.int64)
            if t.get('fixedEnd'):
                fixed = day(t['fixedEnd'])
                candidate = cal.latest(fixed,duration)
                if np.any(candidate < s) or np.any(cal.add(candidate,duration) != fixed): raise ValueError(f'{k}: 고정 완료일과 관계/달력 충돌')
                if t.get('fixedStart') and np.any(candidate != s): raise ValueError(f'{k}: 고정 시작/완료와 기간 충돌')
                s = candidate
            f = cal.add(s,duration)
        if state in ('진행','완료'):
            for r,bound in historical:
                actual_anchor = f if r['type'] == 'FF' else s
                if np.any(bound > actual_anchor):
                    warnings.append(f"{k}: 실적 관계 위반 {r['pred']} {r['type']} (실제 날짜 보존; 미완료 FS는 잔여 작업 착수에 적용)")
        for field, value in (('fixedStart',s),('fixedEnd',f)):
            if t.get(field) and np.any(value != day(t[field])): raise ValueError(f'{k}: {field}와 실제/예측 날짜 충돌')
        starts[k], ends[k] = s,f
        if progress: progress(int((idx+1)/len(order)*100))
    if not active: raise ValueError('적용 작업이 없습니다.')
    total = np.maximum.reduce([ends[k] for k in active])
    return dict(starts=starts,ends=ends,total=total,n=n,seed=seed,engine=VERSION,
                asof=p['asof'],data_version=p.get('data_version',0),project_uid=p.get('uid'),
                warnings=list(dict.fromkeys(warnings)),_durations=durations)


def simulate(p,n=5000,seed=42,progress=None,cancel=None):
    if type(n) is not int or not 1 <= n <= 20000: raise ValueError('반복 횟수: 1~20,000')
    return _run(p,n,seed,True,progress,cancel)


def deterministic(p):
    """Backward pass over generalized precedence. Float is in elapsed calendar days.

    Actual starts/completions are immutable; their float is undefined (None).
    Baselines and target dates never set the network finish boundary.
    """
    p = canonical(p,identities=False)
    r = _run(p)
    _,_,order = inspect(p)
    latest_s, latest_f, floats = {},{},{}
    finish = int(r['total'][0])
    outgoing = {t['id']:[] for t in order}
    for edge in relations(p): outgoing[edge['pred']].append(edge)
    taskmap = {t['id']:t for t in order}
    for t in reversed(order):
        k = t['id']
        cal = Calendar(p,t['calendar'],t['calendarId'])
        es,ef = int(r['starts'][k][0]),int(r['ends'][k][0])
        d = int(r['_durations'][k][0])
        finish_cap, start_cap = finish, finish
        for edge in outgoing[k]:
            succ = edge['succ']
            successor = taskmap[succ]
            if successor['state']=='완료' or (successor['state']=='진행' and edge['type']=='SS'):
                continue  # Historical relation violations warn; they cannot move real dates.
            if successor['state']=='진행' and edge['type']=='FS':
                succcal = Calendar(p,successor['calendar'],successor['calendarId'])
                bound = int(succcal.latest(latest_f[succ],int(r['_durations'][succ][0])))
            else:
                bound = latest_f[succ] if edge['type']=='FF' else latest_s[succ]
            lagcal = Calendar(p,edge['lagUnit'],edge['lagCalendar'])
            cap = int(lagcal.latest(bound,int(edge['lag'])))
            if edge['type'] == 'SS': start_cap = min(start_cap,cap)
            else: finish_cap = min(finish_cap,cap)
        if t['state']=='완료':
            latest_s[k],latest_f[k],floats[k] = es,ef,None
        elif t['state']=='진행':
            lf = max(ef,finish_cap)
            if t.get('fixedEnd'): lf = ef
            latest_s[k],latest_f[k],floats[k] = es,lf,None
        else:
            ls = min(int(cal.latest(finish_cap,d)),start_cap)
            if t.get('fixedStart') or t.get('fixedEnd'): ls = es
            # A start cap on a weekend must roll backward, never forward.
            # latest(ls,0) preserves a milestone; normal work starts must be a workday.
            if d and t['calendar']=='근무일':
                ix = np.searchsorted(cal.working,min(int(cal.latest(finish_cap,d)),start_cap),side='right')-1
                ls = int(cal.working[ix])
                if t.get('fixedStart') or t.get('fixedEnd'): ls = es
            latest_s[k],latest_f[k] = ls,int(cal.add(ls,d))
            floats[k] = ls-es
    r['float'] = floats
    r['critical'] = [k for k,v in floats.items() if v == 0 and p['modules'][taskmap[k]['module']]!='제외' and taskmap[k]['state']!='제외']
    r['latest_starts'],r['latest_ends'] = latest_s,latest_f
    r['target_delta'] = day(p['target'])-finish if p.get('target') else None
    r['milestones'] = {t['id']:textday(r['ends'][t['id']][0]) for t in order if t.get('m') == 0 and t['state']!='제외'}
    return r


def percentile(a,q):
    return int(np.quantile(a,q/100,method='higher'))


def summary(r,p):
    return dict(P50=textday(percentile(r['total'],50)),P80=textday(percentile(r['total'],80)),
                P90=textday(percentile(r['total'],90)),
                probability=float(np.mean(r['total']<=day(p['target']))) if p.get('target') else None,
                iterations=r['n'],seed=r['seed'],engine=r['engine'],data_version=r.get('data_version',0),
                asof=r.get('asof',p['asof']))


def demo():
    rows=[('A01','01 기획','사업범위 · 설계기준 확정',3,5,8,[]),('A02','01 기획','부지 조건 및 현황 확인',5,8,15,['A01']),('B01','02 설계','기본배치 · 공정설계',8,12,22,['A02']),('B02','02 설계','분야별 기본설계 검토',5,10,18,['B01']),('B03','02 설계','기본설계 승인',0,0,0,['B02']),('C01','03 인허가','신청도서 작성',7,12,20,['B03']),('C02','03 인허가','기관 검토 · 보완 대응',15,30,55,['C01']),('C03','03 인허가','승인 마일스톤',0,0,0,['C02']),('D01','04 구매','기자재 사양 · 발주',7,12,18,['B03']),('D02','04 구매','제작 · 검사 · 현장반입',25,40,65,['D01']),('E01','05 시공','시공도서 발행',5,10,15,['C03']),('E02','05 시공','설비 설치 및 배관공사',15,25,40,['E01','D02']),('F01','06 시운전','검사 · 시운전',7,12,20,['E02']),('F02','06 시운전','인수인계 완료',0,0,0,['F01'])]
    return {'id':'DEMO-001','name':'생산시설 구축 · 예제','start':'2026-10-01','asof':'2026-10-01','target':'2027-05-01','holidays':[],'modules':{r[1]:'적용' for r in rows},'tasks':[dict(zip(['id','module','name','o','m','p','pred'],r),state='계획',calendar='달력일',role='',deliverable='',ro=None,rm=None,rp=None,actualStart='',actualEnd='',constraint='',baselineStart='',baselineEnd='',note='') for r in rows]}
