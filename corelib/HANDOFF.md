# HANDOFF.md — 다음 세션 시작 브리핑

> 새 채팅창에서 이 파일을 먼저 읽고 시작

---

## 프로젝트 개요

Excel VBA 범용 라이브러리 `corelib.xlam` 제작 프로젝트.
CWB(Client WorkBook)에서 `Workbooks.Open` 방식으로 xlam을 로드하고
`Application.Run "corelib.xlam!모듈.함수"` 방식으로 호출.

**프로젝트 경로:** `d:\CLOUD\OneDrive\DIEL_VAULT\DEVELOPMENT\Git\Excel-VBA\corelib\`

---

## 필수 참조 파일

| 파일 | 용도 |
|---|---|
| `CLAUDE.md` | 코딩 규칙, 네이밍, 에러 패턴, 설계 원칙 전체 |
| `LOG.md` | 모듈 완성 현황, 이식 현황, 설계 결정 사항 |
| `MIGRATE.md` | 이식 작업 인덱스 (⬜/✅ 체크리스트) |
| `refer/SOURCES.md` | tpl_ 소스별 이식 여부·수정 사항 분석 |
| `refer/export/*.bas` | vba.xlsm 원본 소스 (직접 Read 가능, xlsm 열 필요 없음) |

---

## 현재 완성 상태 (2026-06-02 기준)

### 완성된 am_ 모듈

| 모듈 | 주요 기능 |
|---|---|
| `am_Core` | 상수, 프로퍼티, 초기화, DPUpdate/Event/Calculate/WB_Lock |
| `am_Path` | 경로 토큰 변환, 경로 정규화 |
| `am_File` | 파일/폴더 생성·삭제·복사·검색·다이얼로그 |
| `am_DB` | DB 연결·쿼리 실행·스키마·타입 처리 (골격만, 확장 필요) |
| `am_Sheet` | 백업, 표시/숨김, 정렬, SheetLock/SheetUnLock |
| `am_Table` | 테이블 CRUD·필터·정렬·검색 |
| `am_Range` | FindRange, FindCellsByColor, GetUsedRange |
| `am_Format` | 조건부 서식, 유효성 검사 |

### 미완성

| 모듈 | 상태 |
|---|---|
| `am_Error` | ⬜ 미작성 (다음 작업) |
| `am_Utils` | ⬜ 미작성 |
| `am_Excel` | ⬜ 스텁만 있음 |

---

## 다음 작업 순서 (Phase 2부터)

### Phase 2 — 신규 모듈
1. **`am_Error` 신규** ← **바로 다음 작업**
   - 소스: `refer/export/tpl_Error.bas`
   - 참고: `refer/SOURCES.md` → am_Error 섹션
   - `HandleError` : `cl.*` 호출 제거, am_Core 에러 패턴으로 교체
   - `WriteLog` : `ThisWorkbook.Path` → `am_Core.XlamPath`

2. **`am_DB` 확장**
   - 소스: `tpl_MsSQL.bas`, `tpl_Access.bas`, `tpl_MySQL_Sub.bas`
   - 참고: `refer/SOURCES.md` → am_DB 섹션 (이식 여부 분석 완료)

### Phase 3 — am_Utils 신규
- 소스: tpl_Array, tpl_Check(일부), tpl_Code(일부), tpl_ExtApp, tpl_ReplaceText(일부), tpl_Tools(일부), tpl_Validation

### Phase 4 — am_Excel 완성
- 소스: tpl_ExportFile, tpl_Chart, tpl_Shapes, tpl_KeyBoard, tpl_Mouse

---

## 작업 방식

1. `MIGRATE.md` 에서 ⬜ 항목 선택
2. `refer/SOURCES.md` 에서 해당 모듈 이식 여부·수정 사항 확인
3. `refer/export/tpl_XXX.bas` Read
4. 기존 am_ 모듈 Read
5. 코드 작성
6. `MIGRATE.md` 해당 항목 ✅ 업데이트
7. `LOG.md` 변경 이력 추가

---

## 핵심 설계 원칙 (자주 잊는 것)

- **모듈 간 직접 호출 금지** → 필요하면 `prv_` 로 내부 구현
- **Late Binding** → `CreateObject(...)` 사용, 참조 추가 안 함
- **Optional 인수** → 반드시 필수 인수 뒤에 위치
- **색상 센티널** → `-1` (0은 검정으로 유효값)
- **배열 인덱스** → 항상 `LBound` / `UBound`
- **에러 패턴** → `On Error GoTo ErrHandler` + `MsgBox ... am_Core.AM_NAME`
- **CleanUp 패턴** → 개체 해제 필요 시 `GoTo CleanUp` 사용
