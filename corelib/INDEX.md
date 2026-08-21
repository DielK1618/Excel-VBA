# INDEX.md — corelib 기능→파일 매핑

일반적인 기능 탐색용 인덱스다. 모듈/기능명으로 파일 경로를 바로 찾을 때 참조한다.

> **이식(마이그레이션) 작업 시**에는 이 파일 대신 [`MIGRATE.md`](MIGRATE.md)를 참조한다 — 원본→am_ 이식 상태(⬜/✅) 체크리스트는 그쪽이 기준이다.
> 코딩 규칙은 [`CLAUDE.md`](CLAUDE.md), 진행 로그는 [`LOG.md`](LOG.md), 세션 시작 브리핑은 [`HANDOFF.md`](HANDOFF.md) 참조.

---

## 폴더 구조 요약

| 폴더 | 내용 | xlam 대상 여부 |
|---|---|---|
| `xlam/` | corelib.xlam 실제 모듈 소스 | 대상 (xlam 본체) |
| `cwb/` | cwb_01.xlsm(클라이언트 워크북) 실제 모듈 소스 | 아님 — xlam을 호출하는 소비자 |
| `refer/export/` | vba.xlsm 원본 소스(이식 전 참고용) | 아님 — 참고 자료 |

---

## xlam 대상 모듈 (`xlam/`)

corelib.xlam 본체를 구성하는 모듈이다. `am_` 접두사를 쓴다.

| 모듈 | 파일 경로 | 역할 |
|---|---|---|
| `ThisWorkbook` | `xlam/현재_통합_문서.cls` | xlam 열림/닫힘 이벤트 |
| `am_Core` | `xlam/am_Core.bas` | 전역 상수, Property, 초기화/정리 |
| `am_Path` | `xlam/am_Path.bas` | 경로 토큰 변환, 경로 정규화 |
| `am_File` | `xlam/am_File.bas` | 파일/폴더 생성·삭제·복사·검색, 다이얼로그 |
| `am_DB` | `xlam/am_DB.bas` | DB 연결, 쿼리 실행, 스키마 조회, Access/MySQL/MsSQL 처리 |
| `am_Range` | `xlam/am_Range.bas` | FindRange, FindCellsByColor, GetUsedRange |
| `am_Sheet` | `xlam/am_Sheet.bas` | 시트 백업, 표시/숨김, 정렬, SheetLock/SheetUnLock |
| `am_Table` | `xlam/am_Table.bas` | 테이블(ListObject) CRUD·필터·정렬·검색 |
| `am_Format` | `xlam/am_Format.bas` | 조건부 서식, 유효성 검사 |
| `am_Excel` | `xlam/am_Excel.bas` | 인쇄/내보내기, 차트, 도형 등 Excel 객체 모델 자동화 |
| `am_Automation` | `xlam/am_Automation.bas` | 키보드/마우스 입력 시뮬레이션, 창 탐색·활성화·버튼 제어(Win32 API, Excel 비종속) |
| `am_Utils` | `xlam/am_Utils.bas` | 배열·검사·코드생성·날짜·외부앱·수식 등 범용 유틸리티 |
| `am_Error` | `xlam/am_Error.bas` | 공통 에러 핸들링, 로그 기록 |

---

## CWB 전용 모듈 (`cwb/`) — xlam 대상 아님

cwb_01.xlsm(클라이언트 워크북) 쪽 모듈이다. xlam을 소비하는 쪽이므로 xlam 본체에는 포함되지 않는다.

### ref_ 모듈 — am_ 1:1 래퍼 (기본 포함, xlam 업데이트 시 통째로 교체)

| 모듈 | 파일 경로 | 대응 xlam 모듈 |
|---|---|---|
| `ref_Core` | `cwb/ref_Core.bas` | `am_Core` |
| `ref_Path` | `cwb/ref_Path.bas` | `am_Path` |
| `ref_File` | `cwb/ref_File.bas` | `am_File` |
| `ref_DB` | `cwb/ref_DB.bas` | `am_DB` |
| `ref_Range` | `cwb/ref_Range.bas` | `am_Range` |
| `ref_Sheet` | `cwb/ref_Sheet.bas` | `am_Sheet` |
| `ref_Table` | `cwb/ref_Table.bas` | `am_Table` |
| `ref_Format` | `cwb/ref_Format.bas` | `am_Format` |
| `ref_Excel` | `cwb/ref_Excel.bas` | `am_Excel` |
| `ref_Automation` | `cwb/ref_Automation.bas` | `am_Automation` |
| `ref_Utils` | `cwb/ref_Utils.bas` | `am_Utils` |
| `ref_Error` | `cwb/ref_Error.bas` | `am_Error` |

### tpl_ 모듈 — CWB 비즈니스 로직 (필요시 작성)

| 모듈 | 파일 경로 | 역할 |
|---|---|---|
| `ThisWorkbook` | `cwb/현재_통합_문서.cls` | xlam 로드, 이벤트 처리 |
| `tpl_Path` | `cwb/tpl_Path.bas` | 경로 관련 CWB 전용 유틸리티 |
| `tpl_Test` | `cwb/tpl_Test.bas` | 전 am_ 모듈 자동화 테스트 프로시저 모음 |

---

## frm_ 폼 모듈 — CWB 전용, xlam 대상 아님

`frm_` 접두사 모듈은 항상 특정 워크북(CWB)에 종속되는 사용자 정의 폼이며 xlam 이식 대상에서 원칙적으로 제외된다(`MIGRATE.md` "CWB 전용(이식 제외)" 참조). 현재는 `refer/export/`에 vba.xlsm 원본 소스로만 존재하고, `cwb/`에는 아직 임포트되지 않은 상태다.

| 폼 | 파일 경로 |
|---|---|
| `frm_AddList` | `refer/export/frm_AddList.frm` |
| `frm_AddReasonForAbsence` | `refer/export/frm_AddReasonForAbsence.frm` |
| `frm_DBinfo` | `refer/export/frm_DBinfo.frm` |
| `frm_GetPath` | `refer/export/frm_GetPath.frm` |
| `frm_IndividualCourseStatus` | `refer/export/frm_IndividualCourseStatus.frm` |
| `frm_Navigation` | `refer/export/frm_Navigation.frm` |
| `frm_RowsAddOrDelete` | `refer/export/frm_RowsAddOrDelete.frm` |
| `frm_Stop` | `refer/export/frm_Stop.frm` |
| `frm_TableArrayFilter` | `refer/export/frm_TableArrayFilter.frm` |
| `frm_TextBox` | `refer/export/frm_TextBox.frm` |
| `frm_UnProtect` | `refer/export/frm_UnProtect.frm` |

---

## 참고: refer/export/ 원본 소스 (이식 전, xlam 대상 아님)

vba.xlsm에서 내보낸 원본 모듈이다. 이식 여부·수정 사항은 [`refer/SOURCES.md`](refer/SOURCES.md)에 정리되어 있으며, 개별 함수 단위 이식 대상 여부는 이 파일이 아니라 `SOURCES.md`/`MIGRATE.md`를 기준으로 판단한다.

- `refer/export/*.cls` — DB_*, 시즌/챕터/학생명단 등 CWB 업무 클래스 (xlam 대상 아님)
- `refer/export/tpl_*.bas` — am_ 모듈로 이식되는 원본 템플릿 모듈 (일부만 이식, 세부는 `SOURCES.md`)
- `refer/export/prj_*.bas` — Selenium/Upload/Form 등 프로젝트 전용 모듈 (xlam 대상 아님)
- `refer/export/frm_*.frm/.frx` — 위 "frm_ 폼 모듈" 표 참조
- `refer/export/vba.xlsm` — 원본 워크북 (직접 열 필요 없음, `export/*.bas` Read로 대체)
