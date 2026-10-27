# 종정보 SQLite 전환

종정보 Excel은 관리자 입력 자료로만 사용하고, BioFlow 실행 및 배포에는 완성된
`data/master/species_master.sqlite`를 사용한다.

## 동작 방식

- 앱 시작 시 전체 45,542행을 DataFrame이나 사전으로 적재하지 않는다.
- 국명과 학명 조회는 SQLite 인덱스를 사용한다.
- 종정보 관리 화면은 처음 200행만 미리보기로 읽는다.
- 사용자가 새 Excel을 등록하면 임시 SQLite를 만든 뒤 무결성, 스키마, 행 수,
  종고유ID 중복을 검증하고 전체 DB를 교체한다.
- GitHub 자동 업데이트도 SHA-256과 행 수를 확인한 뒤 같은 방식으로 전체 교체한다.
- 이전 DB는 `species_master.sqlite.bak`으로 한 세대 보관한다.
- DB 연결은 읽기 전용 URI와 `PRAGMA query_only=ON`을 사용한다.

## 생성 및 배포

```powershell
python tools/build_species_db.py
```

원본 Excel은 PyInstaller 배포 목록에서 제외하고 SQLite만 포함한다. 원본은 새 DB를
만들고 결과를 검증한 뒤 별도로 보관하거나 삭제할 수 있다. GitHub의
`data/version.json`에는 `species_rows`와 `species_sha256`을 함께 갱신해야 한다.

생성 통계는 `docs/species-db-build.json`에 기록된다.

## 2026-10-27 검증 결과

| 항목 | 결과 |
|---|---:|
| 원본 Excel | 17,612,790 bytes |
| 원본 행 / 컬럼 | 45,542 / 53 |
| 최종 SQLite | 43,016,192 bytes |
| Excel 단독 ZIP | 15,339,784 bytes |
| SQLite 단독 ZIP | 11,166,755 bytes |
| 200행 미리보기 | 약 12.14 ms |
| 국명 조회 100회 평균 | 약 0.865 ms |
| 기존 배포 ZIP | 250,163,513 bytes |
| SQLite 배포 ZIP | 247,312,918 bytes |
| 배포 ZIP 감소 | 2,850,595 bytes |

배포 폴더의 `.xlsx` 3개는 최종 제출 양식 템플릿이며 종정보 원본이 아니다.
`species_master.xlsx`는 배포물에 포함되지 않는다.
