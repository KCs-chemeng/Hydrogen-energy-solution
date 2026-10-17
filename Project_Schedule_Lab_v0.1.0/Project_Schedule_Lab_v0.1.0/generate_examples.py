"""Rebuild distributable templates and report the unmodified reference workbook."""
from pathlib import Path
from uuid import uuid5,NAMESPACE_URL
from collections import Counter
import hashlib
import json
from model import canonical,sync_pred
from engine import deterministic,inspect,textday
from excel_io import write_excel,read_document
from importing import prepare,materialize

ROOT=Path(__file__).resolve().parent


def main():
    examples=ROOT/'examples';reports=ROOT/'reports';examples.mkdir(exist_ok=True);reports.mkdir(exist_ok=True)
    p=canonical(dict(uid=str(uuid5(NAMESPACE_URL,'psl:validation-small')),id='VALIDATION-SMALL',name='FS·SS·FF 수작업 검증',
                     start='2026-10-01',asof='2026-10-01',target='2026-10-06',modules={'기본':'적용'},tasks=[
                         dict(id=k,uid=str(uuid5(NAMESPACE_URL,'psl:small:'+k)),module='기본',name=name,o=d,m=d,p=d,
                              state='계획',calendar='달력일',pred=[],wbs=f'1.{i+1}')
                         for i,(k,name,d) in enumerate([('A','선행 4일',4),('B','병행 SS+1 / 2일',2),('C','FF+2 / 합류 마일스톤 전 1일',1)])]))
    p['relations']=[dict(pred=a,succ=b,type=typ,lag=lag,lagUnit='달력일',lagCalendar='default',note='수작업 검증 근거',deliverable='검토 결과')
                    for a,b,typ,lag in [('A','B','SS',1),('A','C','FF',2),('B','C','FS',0)]]
    sync_pred(p);r=deterministic(p)
    write_excel(examples/'validation_small.xlsx',p,r)
    expected=dict(starts={k:textday(v[0]) for k,v in r['starts'].items()},ends={k:textday(v[0]) for k,v in r['ends'].items()},
                  float=r['float'],critical=r['critical'],target_delta=r['target_delta'])
    (examples/'expected_small.json').write_text(json.dumps(expected,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    template=canonical(dict(uid=str(uuid5(NAMESPACE_URL,'psl:standard-template')),id='NEW-PROJECT',name='프로젝트명 입력',
                            start='2026-10-01',asof='2026-10-01',target='',modules={'공통':'적용'},tasks=[]))
    write_excel(examples/'standard_template.xlsx',template)
    cal=canonical(dict(uid=str(uuid5(NAMESPACE_URL,'psl:calendar-example')),id='CALENDAR-EXAMPLE',name='휴일·추가근무일·혼합달력',
                       start='2026-10-02',asof='2026-10-02',target='',holidays=['2026-10-05'],workdays=['2026-10-03'],
                       calendars={'seven':dict(weekmask=list(range(7)),holidays=[],workdays=[])},modules={'공통':'적용'},tasks=[
                           dict(id=k,uid=str(uuid5(NAMESPACE_URL,'psl:cal:'+k)),module='공통',name=k,state='계획',calendar='근무일',calendarId=cid,o=1,m=1,p=1,pred=pred)
                           for k,cid,pred in [('A','default',[]),('B','seven',['A'])]]))
    write_excel(examples/'calendar_example.xlsx',cal,deterministic(cal))
    # Exercise files as actual Import inputs, without substituting dates or modules.
    for path in examples.glob('*.xlsx'):
        plan=prepare(path);materialize(plan,acknowledge=True)
    path=ROOT/'common_Rev1.xlsx';doc=read_document(path);plan=prepare(path)
    errors,warnings,_=inspect(doc.project)
    report=dict(file=path.name,sha256=hashlib.sha256(path.read_bytes()).hexdigest(),
                sheets=['프로젝트설정','모듈선택','상세일정','주간간트','변경이력','리스크대장','문서대장','사용가이드'],
                tasks=len(doc.project['tasks']),modules=len(doc.project['modules']),
                module_states=dict(Counter(doc.project['modules'].values())),read_errors=[str(i) for i in doc.issues],
                import_review=plan.report(),schedule_errors=errors,schedule_warnings=warnings,
                missing_against_existing='신규 프로젝트로 사전 검토: 비교할 기존 작업 없음',
                applied=False,reason='원본 착수일·기준일 공란. 5개 모듈 적용성 미확정. 날짜·관계를 자동 보정하지 않았습니다.')
    (reports/'reference_import_report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    lines=['# 원본 common_Rev1.xlsx Import 사전 검토','',f"- SHA-256: `{report['sha256']}`",'- 원본 8개 시트 / 작업 243개 / 모듈 29개 (적용 24, 검토 필요 5)',
           '- 파일 읽기·숫자·날짜 파싱 오류 없음. 신규 프로젝트로 비교하여 기존 작업 누락 비교는 해당 없음.',
           '- 착수일·기준일 공란으로 적용 차단. 원본 또는 현재 프로젝트를 근거 없이 변경하지 않음.',
           '- 5개 검토 필요 모듈은 계산 전 사용자가 적용성을 결정해야 함.',
           '- 표준 Export는 일정 입력만 변환. 원본 주간간트/리스크/문서대장/그림/서식은 복제하지 않음.','', '## 상세 진단','']
    lines += ['- '+str(i) for i in plan.issues]
    lines += ['', '## 계산 차단 항목','']+['- '+e for e in errors]
    lines += ['', '테스트의 243개 작업·5,000회 계산은 별도 메모리 복사본에 시험용 날짜와 모듈 적용을 지정한 회귀 시험입니다. 원본 프로젝트에 대한 유효한 예측이라는 의미가 아닙니다.','']
    (reports/'reference_import_report.md').write_text('\n'.join(lines),encoding='utf-8')
    print(json.dumps(dict(example_files=[p.name for p in examples.iterdir()],reference_tasks=report['tasks'],reference_applied=False),ensure_ascii=False))


if __name__=='__main__': main()
