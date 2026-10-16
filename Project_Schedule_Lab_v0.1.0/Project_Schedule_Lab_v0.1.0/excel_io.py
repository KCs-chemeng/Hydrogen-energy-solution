"""Python application's user-requested Excel interchange; never overwrites the input file."""
from datetime import date,datetime
from pathlib import Path
import os,tempfile
from openpyxl import load_workbook,Workbook
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from engine import summary,percentile,textday
COLS={'id':'Activity ID','module':'모듈','name':'작업명','state':'상태','role':'담당','deliverable':'산출물','calendar':'기간단위','o':'O','m':'M','p':'P','pred':'선행 ID','ro':'잔여 O','rm':'잔여 M','rp':'잔여 P','actualStart':'실제 시작','actualEnd':'실제 완료','constraint':'착수 제한일','baselineStart':'기준 시작','baselineEnd':'기준 완료','note':'비고'}
DATES=['actualStart','actualEnd','constraint','baselineStart','baselineEnd']
def dt(v):
    if isinstance(v,(date,datetime)):return v.strftime('%Y-%m-%d')
    return str(v or '').strip()
def number(v):
    if v in (None,''):return None
    if isinstance(v,str) and v.startswith('='):raise ValueError('기간 입력칸에 수식이 있습니다. 숫자 값으로 붙여넣으세요.')
    return float(v)
def read_excel(path):
    w=load_workbook(path,data_only=False)
    if '작업목록' in w.sheetnames:
        cfg={r[0]:r[1] for r in w['프로젝트'].iter_rows(min_row=2,values_only=True) if r[0]};mods={str(r[0]):str(r[1] or '검토 필요') for r in w['모듈'].iter_rows(min_row=2,values_only=True) if r[0]}
        p={'id':str(cfg.get('PJ_ID') or 'PJ-NEW'),'name':str(cfg.get('프로젝트명') or Path(path).stem),'start':dt(cfg.get('시작일')),'asof':dt(cfg.get('기준일')),'target':dt(cfg.get('목표일')),'holidays':[],'modules':mods,'tasks':[]}
        if '휴일' in w.sheetnames:p['holidays']=[dt(r[0]) for r in w['휴일'].iter_rows(min_row=2,values_only=True) if r[0]]
        rows=list(w['작업목록'].values);headers=list(rows[0]);missing=set(COLS.values())-set(headers)
        if missing:raise ValueError('필수 열 누락: '+', '.join(sorted(missing)))
        for row in rows[1:]:
            if not any(v is not None for v in row):continue
            t={k:row[headers.index(v)] if headers.index(v)<len(row) else None for k,v in COLS.items()}
            for k in ['id','module','name','state','role','deliverable','calendar','note']:t[k]=str(t[k] or '').strip()
            t['pred']=[s.strip() for s in str(t['pred'] or '').replace(';',',').split(',') if s.strip()]
            for k in ['o','m','p','ro','rm','rp']:t[k]=number(t[k])
            for k in DATES:t[k]=dt(t[k])
            t['state']=t['state'] or '계획';t['calendar']=t['calendar'] or '달력일';p['tasks'].append(t)
    elif all(s in w.sheetnames for s in ['상세일정','모듈선택','프로젝트설정']):
        c=w['프로젝트설정'];p={'id':str(c['B6'].value or 'PJ-NEW'),'name':str(c['B9'].value or Path(path).stem),'start':dt(c['B7'].value),'asof':dt(c['B8'].value),'target':dt(c['B11'].value),'holidays':[],'modules':{},'tasks':[]}
        for r in w['모듈선택'].iter_rows(min_row=6,values_only=True):
            if r[0]:p['modules'][str(r[0])]=str(r[1] or '검토 필요')
        for r in w['상세일정'].iter_rows(min_row=6,values_only=True):
            if not r[0]:continue
            t={k:'' for k in COLS};t.update(id=str(r[0]),module=str(r[1]),name=str(r[2]),state=r[4] or '계획',role=r[5] or '',deliverable=r[6] or '',calendar='달력일',pred=[str(x) for x in r[23:29] if x])
            for k,v in zip(['o','m','p','ro','rm','rp'],r[12:18]):t[k]=number(v)
            for k,v in zip(['baselineStart','baselineEnd','actualStart','actualEnd','constraint'],r[7:12]):t[k]=dt(v)
            p['tasks'].append(t)
    else:raise ValueError('지원 형식: 공통 Rev1 또는 이 프로그램에서 내보낸 교환용 XLSX입니다.')
    w.close()
    if len(p['tasks'])>2000:raise ValueError('최대 2,000개 작업을 지원합니다.')
    return p

def write_excel(path,p,r=None):
    w=Workbook();w.remove(w.active)
    s=w.create_sheet('프로젝트');s.append(['항목','값'])
    for k,v in [('형식버전','PSL-0.1'),('PJ_ID',p['id']),('프로젝트명',p['name']),('시작일',p['start']),('기준일',p['asof']),('목표일',p.get('target',''))]:s.append([k,date.fromisoformat(v) if k in ['시작일','기준일','목표일'] and v else v])
    s=w.create_sheet('모듈');s.append(['모듈','적용여부'])
    for k,v in p['modules'].items():s.append([k,v])
    dv=DataValidation(type='list',formula1='"적용,제외,검토 필요"');s.add_data_validation(dv);dv.add('B2:B2001')
    s=w.create_sheet('작업목록');s.append(list(COLS.values()))
    for t in p['tasks']:
        row=[]
        for k in COLS:
            v=t.get(k,'')
            if k=='pred':v=', '.join(v)
            elif k in DATES and v:v=date.fromisoformat(v)
            row.append(v)
        s.append(row)
    for col,values in [('D','계획,진행,완료,보류,제외'),('G','달력일,근무일')]:
        dv=DataValidation(type='list',formula1='"'+values+'"');s.add_data_validation(dv);dv.add(f'{col}2:{col}2001')
    s=w.create_sheet('휴일');s.append(['휴일 날짜'])
    for h in p.get('holidays',[]):s.append([date.fromisoformat(h)])
    s=w.create_sheet('사용안내');s.append(['항목','설명'])
    for row in [('입력','작업목록의 ID는 고유해야 합니다. 선행 ID는 쉼표로 구분합니다.'),('전체/잔여','계획은 O/M/P, 진행은 잔여 O/M/P와 실제 시작, 완료는 실제 시작/완료를 입력합니다.'),('일수','경과일 방식: 10/1 시작 1달력일 작업은 10/2 완료. 근무일은 월~금 및 휴일 시트 반영.'),('확률','독립 삼각분포, 기간 올림. 자원 가용량·지연 상관관계는 미반영.'),('검토','초기 템플릿 기간은 미보정 가정값입니다. 적용 모듈과 날짜를 검토하세요.'),('내보내기','원본 서식·기타 시트 복제가 아닌 프로그램 교환 양식입니다. 결과는 마지막 실행의 스냅샷입니다.'),('이력','수정사유·분석이력은 로컬 SQLite에 저장됩니다. 이 파일에는 현재 입력과 선택적으로 결과만 포함됩니다.')]:s.append(row)
    if r is not None:
        s=w.create_sheet('분석요약');s.append(['항목','결과'])
        for k,v in summary(r,p).items():s.append([k,v])
        s=w.create_sheet('작업결과');s.append(['Activity ID','작업명','시작 P50','완료 P50','완료 P80','완료 P90'])
        for t in p['tasks']:
            k=t['id'];s.append([k,t['name']]+[date.fromisoformat(textday(percentile(a,q))) for a,q in [(r['starts'][k],50),(r['ends'][k],50),(r['ends'][k],80),(r['ends'][k],90)]])
    for s in w:
        s.freeze_panes='A2';s.auto_filter.ref=s.dimensions;s.sheet_view.showGridLines=False
        for c in s[1]:c.fill=PatternFill('solid',fgColor='183653');c.font=Font(color='FFFFFF',bold=True);c.alignment=Alignment(vertical='center')
        s.row_dimensions[1].height=27
        for row in s.iter_rows(min_row=2):
            for c in row:
                if isinstance(c.value,str):c.data_type='s' # prevent formula execution from free-text input
                if isinstance(c.value,(date,datetime)):c.number_format='yyyy-mm-dd'
                c.font=Font(name='맑은 고딕',size=10,color='16365D');c.fill=PatternFill('solid',fgColor='FFFBE6' if s.title in ['작업목록','프로젝트','모듈','휴일'] else 'F1F5FA');c.alignment=Alignment(vertical='top',wrap_text=True)
        for column in s.columns:
            key=column[0].column_letter;s.column_dimensions[key].width=22
        if s.title=='작업목록':s.column_dimensions['C'].width=48;s.column_dimensions['K'].width=32
        if s.title=='사용안내':s.column_dimensions['B'].width=100
    dest=Path(path);fd,tmp=tempfile.mkstemp(suffix='.xlsx',dir=dest.parent);os.close(fd)
    try:w.save(tmp);os.replace(tmp,dest)
    finally:
        if os.path.exists(tmp):os.unlink(tmp)
