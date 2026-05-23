# SAP ABAP Job List

SAP ABAP, SAP ERP 관련 신입 채용 공고를 잡코리아와 사람인에서 수집해 엑셀 파일로 정리하는 스크립트입니다.
Codex로 생성된 스크립트입니다.

## 구조

```text
.
├── code/
│   ├── update_jobs.py      # 채용 공고 수집 및 엑셀 생성
│   └── run_daily_jobs.py   # 하루 1회 실행 제어용 래퍼
├── .gitignore
└── README.md
```

## 실행

```bash
python3 code/update_jobs.py
```

실행하면 저장소 루트에 `sap_abap_jobs.xlsx`가 생성됩니다.

엑셀 파일은 다음 두 시트로 구성됩니다.

- `모집중`
- `마감`

각 공고에는 공고명, 링크, 회사, 조건, 모집 기간, 사이트, 최초 수집일, 최근 확인일이 포함됩니다.

## 일일 실행

```bash
python3 code/run_daily_jobs.py
```

`run_daily_jobs.py`는 같은 날짜에 이미 성공한 실행이 있으면 다시 실행하지 않습니다. 부팅 직후에는 10분 동안 실행을 건너뜁니다.

## 생성 파일

다음 파일들은 실행 중 생성되는 로컬 산출물이므로 Git에서 제외합니다.

- `sap_abap_jobs.xlsx`
- `code/job_postings_state.json`
- `code/daily_run_state.json`
- `code/job_update.log`
- `code/*.plist`
- `__pycache__/`

## 참고

이 프로젝트는 Python 표준 라이브러리만 사용합니다. 별도 패키지 설치는 필요하지 않습니다.
