"""Transactional local SQLite repository, optimistic versions and non-destructive v1 migration."""
from pathlib import Path
from copy import deepcopy
from datetime import datetime, timezone
import getpass
import json
import os
import sqlite3
import tempfile
from model import canonical, content, dumps

DB_VERSION = 2


def now(): return datetime.now(timezone.utc).isoformat(timespec='microseconds')


def diff(before, after):
    changes = []
    a, b = content(before or {}), content(after)
    for field in sorted(set(a) | set(b)):
        if field in ('tasks', 'relations'):
            if field == 'relations':
                if a.get(field) != b.get(field): changes.append(('',field,a.get(field),b.get(field)))
                continue
            old = {t['id']:t for t in a.get(field,[])}
            new = {t['id']:t for t in b.get(field,[])}
            for task in sorted(set(old)|set(new)):
                for key in sorted(set(old.get(task,{}))|set(new.get(task,{}))):
                    v1,v2 = old.get(task,{}).get(key),new.get(task,{}).get(key)
                    if v1 != v2: changes.append((task,key,v1,v2))
        elif a.get(field) != b.get(field): changes.append(('',field,a.get(field),b.get(field)))
    return changes


class Store:
    def __init__(self,path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True,exist_ok=True)
        self.db = sqlite3.connect(self.path)
        self.db.execute('PRAGMA foreign_keys=ON')
        version = self.db.execute('PRAGMA user_version').fetchone()[0]
        if version > DB_VERSION:
            self.db.close(); raise ValueError('더 최신 DB입니다. 최신 프로그램으로 여세요.')
        legacy = self.db.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='events'").fetchone()
        if version < DB_VERSION and legacy:
            self.backup(self.path.with_name(self.path.name+'.pre-v2-'+datetime.now().strftime('%Y%m%d%H%M%S%f')+'.sqlite3'))
        if version < DB_VERSION:
            self._migrate(bool(legacy))

    def _migrate(self, legacy):
        # SQLite DDL and legacy conversion share one transaction; original events are retained verbatim.
        ddl = [
            'CREATE TABLE IF NOT EXISTS events (seq INTEGER PRIMARY KEY, created TEXT, pj TEXT, kind TEXT, reason TEXT, payload TEXT)',
            'CREATE TABLE projects (uid TEXT PRIMARY KEY, public_id TEXT NOT NULL, name TEXT NOT NULL, version INTEGER NOT NULL, payload TEXT NOT NULL, updated TEXT NOT NULL)',
            'CREATE TABLE versions (project_uid TEXT REFERENCES projects(uid), version INTEGER, payload TEXT NOT NULL, PRIMARY KEY(project_uid,version))',
            'CREATE TABLE tasks (project_uid TEXT REFERENCES projects(uid), uid TEXT, activity_id TEXT, payload TEXT, PRIMARY KEY(project_uid,uid), UNIQUE(project_uid,activity_id))',
            'CREATE TABLE relations (project_uid TEXT REFERENCES projects(uid), seq INTEGER, payload TEXT, PRIMARY KEY(project_uid,seq))',
            'CREATE TABLE modules (project_uid TEXT REFERENCES projects(uid), name TEXT, state TEXT, PRIMARY KEY(project_uid,name))',
            'CREATE TABLE calendars (project_uid TEXT REFERENCES projects(uid), name TEXT, payload TEXT, PRIMARY KEY(project_uid,name))',
            'CREATE TABLE changes (seq INTEGER PRIMARY KEY, project_uid TEXT, task_id TEXT, field TEXT, before_value TEXT, after_value TEXT, created TEXT, actor TEXT, reason TEXT, source TEXT, version INTEGER)',
            'CREATE TABLE event_scope (event_id INTEGER PRIMARY KEY REFERENCES events(seq), project_uid TEXT)',
            'CREATE TABLE baselines (id INTEGER PRIMARY KEY, project_uid TEXT REFERENCES projects(uid), name TEXT, created TEXT, actor TEXT, version INTEGER, payload TEXT, legacy INTEGER DEFAULT 0)',
            'CREATE TABLE forecasts (id INTEGER PRIMARY KEY, project_uid TEXT REFERENCES projects(uid), version INTEGER, created TEXT, payload TEXT)',
            'CREATE TABLE imports (id INTEGER PRIMARY KEY, project_uid TEXT, created TEXT, filename TEXT, status TEXT, report TEXT)',
        ]
        try:
            with self.db:
                self.db.execute('BEGIN IMMEDIATE')
                for stmt in ddl: self.db.execute(stmt)
                if legacy:
                    identities = {}
                    for seq,pj,kind,reason,raw in self.db.execute('SELECT seq,pj,kind,reason,payload FROM events ORDER BY seq').fetchall():
                        payload = json.loads(raw)
                        p = canonical(payload['project'])
                        # v1 had public project IDs only. Preserve all events and choose the last snapshot per ID.
                        identities.setdefault(pj, (p['uid'],{}))
                        p['uid'], task_ids = identities[pj]
                        for t in p['tasks']:
                            task_ids.setdefault(t['id'],t['uid']); t['uid']=task_ids[t['id']]
                        old = self.load(p['uid'])
                        p['data_version'] = (old['data_version'] if old else 0)+1
                        self._write(p)
                        for task,field,a,b in diff(old,p):
                            self.db.execute('INSERT INTO changes(project_uid,task_id,field,before_value,after_value,created,actor,reason,source,version) VALUES (?,?,?,?,?,?,?,?,?,?)',
                                            (p['uid'],task,field,dumps(a),dumps(b),now(),'legacy (unknown)',reason or 'v1 이전','마이그레이션',p['data_version']))
                        self.db.execute('INSERT INTO event_scope VALUES (?,?)',(seq,p['uid']))
                        if kind == '기준선':
                            self.db.execute('INSERT INTO baselines(project_uid,name,created,actor,version,payload,legacy) VALUES (?,?,?,?,?,?,1)',
                                            (p['uid'],reason,now(),'migration',p['data_version'],dumps({'input':p,'note':'v1 입력 스냅샷; 승인 계산 일정 아님'})))
                self.db.execute(f'PRAGMA user_version={DB_VERSION}')
        except Exception:
            self.db.close(); raise

    def _write(self,p):
        payload = dumps(p)
        self.db.execute('INSERT INTO projects VALUES (?,?,?,?,?,?) ON CONFLICT(uid) DO UPDATE SET public_id=excluded.public_id,name=excluded.name,version=excluded.version,payload=excluded.payload,updated=excluded.updated',
                        (p['uid'],p['id'],p['name'],p['data_version'],payload,now()))
        self.db.execute('INSERT INTO versions VALUES (?,?,?)',(p['uid'],p['data_version'],payload))
        for table in ('tasks','relations','modules','calendars'):
            self.db.execute(f'DELETE FROM {table} WHERE project_uid=?',(p['uid'],))
        self.db.executemany('INSERT INTO tasks VALUES (?,?,?,?)',[(p['uid'],t['uid'],t['id'],dumps(t)) for t in p['tasks']])
        self.db.executemany('INSERT INTO relations VALUES (?,?,?)',[(p['uid'],i,dumps(r)) for i,r in enumerate(p['relations'])])
        self.db.executemany('INSERT INTO modules VALUES (?,?,?)',[(p['uid'],m,s) for m,s in p['modules'].items()])
        calendars = deepcopy(p['calendars'])
        calendars.setdefault('default',dict(holidays=p['holidays'],workdays=p['workdays'],weekmask=[0,1,2,3,4]))
        self.db.executemany('INSERT INTO calendars VALUES (?,?,?)',[(p['uid'],k,dumps(v)) for k,v in calendars.items()])

    def load(self,project_uid,version=None):
        if version is None: row=self.db.execute('SELECT payload FROM projects WHERE uid=?',(project_uid,)).fetchone()
        else: row=self.db.execute('SELECT payload FROM versions WHERE project_uid=? AND version=?',(project_uid,version)).fetchone()
        return json.loads(row[0]) if row else None

    def projects(self):
        return self.db.execute('SELECT uid,public_id,name,version FROM projects ORDER BY updated DESC').fetchall()

    def _save(self,p,kind,reason,source,actor,result=None):
        new = canonical(p)
        old = self.load(new['uid'])
        if old and new['data_version'] != old['data_version']:
            raise ValueError('저장 충돌: 다른 수정이 있습니다. 프로젝트를 다시 열고 비교하세요.')
        changes = diff(old,new)
        if changes:
            if not reason.strip(): raise ValueError('변경 사유가 필요합니다.')
            if len({t['id'] for t in new['tasks']}) != len(new['tasks']): raise ValueError('작업 ID 중복')
            if old:
                old_uids={t['id']:t['uid'] for t in old['tasks']}
                for t in new['tasks']:
                    if t['id'] in old_uids and t['uid'] != old_uids[t['id']]: raise ValueError('기존 작업 내부 UID 변경 금지')
            new['data_version'] = (old['data_version'] if old else 0)+1
            self._write(new)
            for task,field,a,b in changes:
                self.db.execute('INSERT INTO changes(project_uid,task_id,field,before_value,after_value,created,actor,reason,source,version) VALUES (?,?,?,?,?,?,?,?,?,?)',
                                (new['uid'],task,field,dumps(a),dumps(b),now(),actor,reason,source,new['data_version']))
        if changes or result is not None:
            cur=self.db.execute('INSERT INTO events(created,pj,kind,reason,payload) VALUES (?,?,?,?,?)',
                                (now(),new['id'],kind,reason,dumps(dict(project=new,summary=result))))
            self.db.execute('INSERT INTO event_scope VALUES (?,?)',(cur.lastrowid,new['uid']))
        if result is not None:
            self.db.execute('INSERT INTO forecasts(project_uid,version,created,payload) VALUES (?,?,?,?)',
                            (new['uid'],new['data_version'],now(),dumps(result)))
        return new

    def add(self,p,kind,reason,result=None,source='화면 편집',actor=None):
        with self.db:
            self.db.execute('BEGIN IMMEDIATE')
            new=self._save(p,kind,reason,source,actor or getpass.getuser(),result)
        p.clear();p.update(new)
        return new

    def record_preview(self,plan):
        with self.db:
            self.db.execute('INSERT INTO imports(project_uid,created,filename,status,report) VALUES (?,?,?,?,?)',
                            (plan.current.get('uid'),now(),plan.filename,'reviewed',dumps(plan.report())))

    def apply_import(self,plan,selected=None,reason='엑셀 Import',acknowledge=False):
        from importing import materialize
        try:
            candidate = materialize(plan,selected,acknowledge)
            current = self.load(candidate['uid'])
            if current and current['data_version'] != plan.current['data_version']: raise ValueError('미리보기 이후 데이터가 바뀌었습니다. 다시 비교하세요.')
            self.backup(self.path.with_name(self.path.name+'.pre-import-'+datetime.now().strftime('%Y%m%d%H%M%S%f')+'.sqlite3'))
            with self.db:
                self.db.execute('BEGIN IMMEDIATE')
                new=self._save(candidate,'엑셀 Import',reason,'엑셀 Import',getpass.getuser())
                self.db.execute('INSERT INTO imports(project_uid,created,filename,status,report) VALUES (?,?,?,?,?)',
                                (new['uid'],now(),plan.filename,'applied',dumps(plan.report())))
            return new
        except Exception as exc:
            with self.db:
                self.db.execute('INSERT INTO imports(project_uid,created,filename,status,report) VALUES (?,?,?,?,?)',
                                (plan.current.get('uid'),now(),plan.filename,'failed',str(exc)))
            raise

    def rows(self,project_uid=None):
        if project_uid:
            return self.db.execute('SELECT e.seq,e.created,e.pj,e.kind,e.reason FROM events e JOIN event_scope s ON s.event_id=e.seq WHERE s.project_uid=? ORDER BY e.seq DESC LIMIT 300',(project_uid,)).fetchall()
        return self.db.execute('SELECT seq,created,pj,kind,reason FROM events ORDER BY seq DESC LIMIT 300').fetchall()

    def get(self,seq):
        row=self.db.execute('SELECT payload FROM events WHERE seq=?',(seq,)).fetchone()
        if not row: raise KeyError(seq)
        payload=json.loads(row[0])
        if not payload['project'].get('uid'):
            scope=self.db.execute('SELECT project_uid FROM event_scope WHERE event_id=?',(seq,)).fetchone()
            if scope:
                p=canonical(payload['project']);p['uid']=scope[0]
                identities={}
                for version_payload, in self.db.execute('SELECT payload FROM versions WHERE project_uid=? ORDER BY version',(scope[0],)):
                    for t in json.loads(version_payload)['tasks']: identities[t['id']]=t['uid']
                for t in p['tasks']:
                    if t['id'] in identities: t['uid']=identities[t['id']]
                payload['project']=p
        return payload

    def history(self,project_uid):
        names=('task','field','before','after','created','actor','reason','source','version')
        return [dict(zip(names,r)) for r in self.db.execute('SELECT task_id,field,before_value,after_value,created,actor,reason,source,version FROM changes WHERE project_uid=? ORDER BY seq DESC',(project_uid,))]

    def baseline(self,p,name,result):
        if not name.strip(): raise ValueError('기준 일정 이름/승인 근거 필요')
        saved=self.load(p['uid'])
        if not saved or content(saved)!=content(p) or saved['data_version']!=p['data_version']: raise ValueError('먼저 프로젝트를 저장하세요.')
        if result.get('data_version') != p['data_version']: raise ValueError('계산 버전 불일치')
        payload=dict(starts={k:int(v[0]) for k,v in result['starts'].items()},ends={k:int(v[0]) for k,v in result['ends'].items()},engine=result['engine'],asof=result['asof'])
        with self.db:
            self.db.execute('INSERT INTO baselines(project_uid,name,created,actor,version,payload) VALUES (?,?,?,?,?,?)',
                            (p['uid'],name,now(),getpass.getuser(),p['data_version'],dumps(payload)))

    def baselines(self,project_uid):
        return [dict(id=r[0],name=r[1],created=r[2],actor=r[3],version=r[4],payload=json.loads(r[5]),legacy=bool(r[6])) for r in self.db.execute('SELECT id,name,created,actor,version,payload,legacy FROM baselines WHERE project_uid=? ORDER BY id',(project_uid,))]

    def backup(self,dest):
        dest=Path(dest)
        if dest.resolve()==self.path.resolve(): raise ValueError('원본 DB를 덮어쓸 수 없습니다.')
        if dest.exists(): raise ValueError('백업 대상 파일이 이미 있습니다. 새 이름을 사용하세요.')
        dest.parent.mkdir(parents=True,exist_ok=True)
        out=sqlite3.connect(dest)
        try: self.db.backup(out)
        finally: out.close()
        return dest

    def restore_backup(self,source):
        source=Path(source)
        if source.resolve()==self.path.resolve(): raise ValueError('현재 DB는 복원 원본으로 사용할 수 없습니다.')
        check=sqlite3.connect(f'file:{source.resolve().as_posix()}?mode=ro',uri=True)
        try:
            if check.execute('PRAGMA integrity_check').fetchone()[0]!='ok': raise ValueError('백업 무결성 오류')
            version=check.execute('PRAGMA user_version').fetchone()[0]
            names={r[0] for r in check.execute("SELECT name FROM sqlite_master WHERE type='table'")}
            if version!=DB_VERSION or not {'projects','versions','events','changes','baselines','imports','tasks','event_scope'} <= names: raise ValueError('v2 프로젝트 백업만 복원 가능합니다. v1은 먼저 별도 폴더에서 앱으로 이전하세요.')
            if check.execute('PRAGMA foreign_key_check').fetchall(): raise ValueError('백업 관계 무결성 오류')
            for raw, in check.execute('SELECT payload FROM projects'):
                p=json.loads(raw)
                if not isinstance(p.get('tasks'),list) or not p.get('uid'): raise ValueError('잘못된 프로젝트 백업')
            fd,tmp=tempfile.mkstemp(dir=self.path.parent,suffix='.sqlite3');os.close(fd)
            out=sqlite3.connect(tmp)
            try: check.backup(out)
            finally: out.close()
        finally: check.close()
        self.backup(self.path.with_name(self.path.name+'.pre-restore-'+datetime.now().strftime('%Y%m%d%H%M%S%f')+'.sqlite3'))
        self.db.close()
        try: os.replace(tmp,self.path)
        finally:
            if os.path.exists(tmp): os.unlink(tmp)
            self.db=sqlite3.connect(self.path); self.db.execute('PRAGMA foreign_keys=ON')
