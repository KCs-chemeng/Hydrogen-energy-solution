"""Versioned canonical input. Public Activity IDs are project-scoped; UIDs are immutable."""
from copy import deepcopy
from uuid import uuid4
import json

SCHEMA_VERSION = 2
TASK_DEFAULTS = dict(module='공통', name='', state='계획', calendar='근무일', calendarId='default',
                     role='', deliverable='', wbs='', o=None, m=None, p=None, ro=None, rm=None,
                     rp=None, pred=[], actualStart='', actualEnd='', constraint='', baselineStart='',
                     baselineEnd='', fixedStart='', fixedEnd='', resume='', progress=None, note='')
STATE_ALIASES = {'미착수':'계획', '진행 중':'진행'}


def uid():
    return str(uuid4())


def canonical(p, identities=True):
    p = deepcopy(p)
    p.setdefault('schema_version', SCHEMA_VERSION)
    if identities:
        p.setdefault('uid', uid())
    p.setdefault('data_version', 0)
    p.setdefault('id', 'PJ-' + uid()[:8])
    p.setdefault('name', '새 프로젝트')
    p.setdefault('target', '')
    p.setdefault('holidays', [])
    p.setdefault('workdays', [])
    p.setdefault('calendars', {})
    if 'default' in p['calendars']:
        p['calendars']['default']['holidays'] = deepcopy(p['holidays'])
        p['calendars']['default']['workdays'] = deepcopy(p['workdays'])
    p.setdefault('modules', {})
    for t in p.setdefault('tasks', []):
        for k, v in TASK_DEFAULTS.items():
            t.setdefault(k, deepcopy(v))
        t['state'] = STATE_ALIASES.get(t['state'], t['state'])
        if identities:
            t.setdefault('uid', uid())
        p['modules'].setdefault(t['module'], '검토 필요')
    if 'relations' not in p:
        p['relations'] = [dict(pred=a, succ=t['id'], type='FS', lag=0, lagUnit='달력일',
                               lagCalendar='default', note='', deliverable='')
                          for t in p['tasks'] for a in t.get('pred', [])]
    for r in p['relations']:
        for k, v in dict(type='FS', lag=0, lagUnit='달력일', lagCalendar='default', note='', deliverable='').items():
            r.setdefault(k, v)
    sync_pred(p)
    return p


def sync_pred(p):
    for t in p['tasks']:
        t['pred'] = list(dict.fromkeys(r['pred'] for r in p.get('relations', []) if r['succ'] == t['id']))


def clone_project(p, public_id=None):
    p = canonical(p)
    mapping = {t['id']: 'T-' + uid()[:12] for t in p['tasks']}
    p['uid'] = uid()
    p['id'] = public_id or 'PJ-' + uid()[:8]
    p['data_version'] = 0
    for t in p['tasks']:
        t['uid'] = uid()
        t['id'] = mapping[t['id']]
    for r in p['relations']:
        r['pred'], r['succ'] = mapping[r['pred']], mapping[r['succ']]
        r.pop('project_uid', None)
    sync_pred(p)
    return p


def content(p):
    """Ignore row order and derived metadata, but never ignore identity or source values."""
    p = deepcopy(p)
    for k in ('data_version', 'schema_version', '_source', '_export_version', '_schema', '_cells', '_issues'):
        p.pop(k, None)
    p['tasks'] = sorted(p.get('tasks', []), key=lambda t: t['id'])
    p['relations'] = sorted(p.get('relations', []), key=lambda r: json.dumps(r, sort_keys=True, ensure_ascii=False))
    p['holidays'] = sorted(p.get('holidays', []))
    p['workdays'] = sorted(p.get('workdays', []))
    return p


def dumps(v):
    return json.dumps(v, ensure_ascii=False, sort_keys=True, allow_nan=False)
