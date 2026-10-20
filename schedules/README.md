# 2027년 1월 착수 일정

[**최신 R2 — 100개 업무·Skid 발주형 엑셀 다운로드**](https://github.com/KCs-chemeng/Hydrogen-energy-solution/raw/refs/heads/main/schedules/LNG_CO2_2027_01_100Tasks_Import.xlsx)

[변경 브리핑](LNG_CO2_2027_100Tasks_BRIEFING.md)

- LNG 기반 수소생산·탄소포집 프로젝트. 시작·기준일 2027-01-01.
- 100개 업무 / 16개 모듈 / 9개 Skid·발주 Package의 조달 45개 업무.
- 레이아웃 → 건축허가용 도서 선행 → 상세설계 후속, 허가 심사와 병행.
- Package별 발주·Vendor 자료·제작승인·제작·FAT/입고 및 독립 설치.
- Import 오류·경고 0건, 데이터 시험 9개·실제 GUI Import·5,000회 계산 검증.
- 담당자 개인명·원본 참고/실적 자료 제외. 납기·Package 경계·허가 적용성은 검토 전 계산용 가정.

앱 v0.2.2에서 **엑셀 가져오기 → 상단 선택한 변경 적용** 후 작업 수 100개와 프로젝트 ID `LNG-CC-202701-R2-100`을 확인하세요. 기존 433개 프로젝트와 별도로 생성됩니다.

구 버전은 [기존 433개 일정](LNG_CO2_2027_01_Schedule_Import.xlsx)으로 보존했습니다. 기존 데이터나 승인 이력을 지우지 않습니다.

재생성: `python schedules/build_100_task_schedule.py` (앱의 numpy/openpyxl 환경 사용).
검증: `python -m unittest discover -s schedules -p test_100_task_schedule.py -v`.
