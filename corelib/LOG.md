# LOG.md — corelib.xlam 진행 로그

> 코딩 규칙 및 설계 원칙은 `CLAUDE.md` 를 참고하세요.

---

## 모듈 완성 현황

### corelib.xlam (`xlam/` 폴더)

| 파일 | 모듈 | 상태 | 비고 |
|---|---|---|---|
| `현재_통합_문서.cls` | `ThisWorkbook` | ✅ 완성 | |
| `am_Core.bas` | `am_Core` | ✅ 완성 | CM_TO_POINTS 상수, DPUpdate/Event/Calculate/WB_Lock 추가 |
| `am_Path.bas` | `am_Path` | ✅ 완성 | |
| `am_File.bas` | `am_File` | ✅ 완성 | |
| `am_DB.bas` | `am_DB` | ✅ 완성 | tpl_MsSQL/MySQL_Sub 미이식 |
| `am_Sheet.bas` | `am_Sheet` | ✅ 완성 | 워크북/시트 조작, 백업, 시트 보호 추가 |
| `am_Table.bas` | `am_Table` | ✅ 완성 | 테이블 CRUD/필터/정렬/검색 |
| `am_Range.bas` | `am_Range` | ✅ 완성 | |
| `am_Format.bas` | `am_Format` | ✅ 완성 | 조건부 서식, 유효성 검사 |
| `am_Excel.bas` | `am_Excel` | ⬜ 스텁 | 차트/버튼/폼/도형 (미이식) |
| `am_Utils.bas` | `am_Utils` | ⬜ 미작성 | |
| `am_Error.bas` | `am_Error` | ⬜ 미작성 | |

### cwb_01.xlsm (`cwb/` 폴더)

| 파일 | 모듈 | 상태 | 비고 |
|---|---|---|---|
| `현재_통합_문서.cls` | `ThisWorkbook` | ✅ 완성 | |
| `tpl_Path.bas` | `tpl_Path` | ✅ 완성 | |
| `tpl_Test.bas` | `tpl_Test` | ✅ 완성 | am_Path 테스트 포함 |

---

## 소스 모듈 이식 현황

원본 소스(tpl_ 모듈)를 am_ 모듈로 이식한 현황입니다.

| 원본 모듈 | 이식 대상 | 상태 |
|---|---|---|
| `tpl_Path` | `am_Path` | ✅ 완료 |
| `tpl_File` | `am_File` | ✅ 완료 |
| `tpl_Sheet` | `am_Sheet` | ✅ 완료 |
| `tpl_Table` | `am_Table` | ✅ 완료 |
| `tpl_Formatting` | `am_Format` | ✅ 완료 |
| `tpl_Const` (CM_TO_POINTS) | `am_Core` | ✅ 완료 |
| `Common.cls` (App 상태, WB 보호) | `am_Core` | ✅ 완료 |
| `Common.cls` (SheetLock) | `am_Sheet` | ✅ 완료 |
| `tpl_Find` | `am_Range` | ✅ 완료 |
| `tpl_Range` | `am_Range` | ✅ 완료 |
| `tpl_Error` | `am_Error` | ⬜ 미이식 |
| `tpl_MsSQL` | `am_DB` | ⬜ 미이식 |
| `tpl_MySQL_Sub` (일부) | `am_DB` | ⬜ 미이식 |
| `tpl_Access` | `am_DB` | ⬜ 미이식 |
| `tpl_Array` | `am_Utils` | ⬜ 미이식 |
| `tpl_Check` (일부) | `am_Utils` | ⬜ 미이식 |
| `tpl_Code` (일부) | `am_Utils` | ⬜ 미이식 |
| `tpl_ExtApp` | `am_Utils` | ⬜ 미이식 |
| `tpl_Media` (일부) | `am_Utils` | ⬜ 미이식 |
| `tpl_ReplaceText` (일부) | `am_Utils` | ⬜ 미이식 |
| `tpl_Tools` (일부) | `am_Utils` | ⬜ 미이식 |
| `tpl_Validation` | `am_Utils` | ⬜ 미이식 |
| `tpl_Chart` | `am_Excel` | ⬜ 미이식 |
| `tpl_ExportFile` | `am_Excel` | ⬜ 미이식 |
| `tpl_KeyBoard` | `am_Excel` | ⬜ 미이식 |
| `tpl_Mouse` | `am_Excel` | ⬜ 미이식 |
| `tpl_Shapes` | `am_Excel` | ⬜ 미이식 |
| `tpl_Buttons`, `tpl_Buttons_other`, `tpl_Buttons_Top` | — | ❌ CWB 전용 |
| `tpl_Form`, `tpl_TestBed`, `tpl_Procedure` | — | ❌ CWB 전용 |
| `frm_*` (전체) | — | ❌ CWB 전용 (사용자 직접 제작) |

---

## 확정된 설계 결정 사항

| 항목 | 결정 내용 |
|---|---|
| xlam 파일명 | `corelib.xlam` |
| xlam 모듈 접두사 | `am_` (Add-in Macro) |
| CWB 템플릿 모듈 접두사 | `tpl_` (Template) |
| Private 함수 접두사 | `prv_` |
| CWB 약자 | `CWB` (Client WorkBook) |
| xlam 로드 방식 | `Workbooks.Open` (AddIns 등록 방식 ❌) |
| xlam 로드 확인 방식 | 오류 방식 (`On Error Resume Next`) |
| CWB 구분 방식 | 프로시저 인수에 명시적 전달 |
| RegisterWB 딕셔너리 | ❌ 폐기 (오류 시 초기화 문제) |
| 모듈 독립성 | 모듈 간 직접 호출 금지, 내부 전용 함수로 자체 구현 |
| 경로 관리 방식 | `Property Get` (런타임 동적 계산) |
| 고정값 관리 방식 | `Const` |
| 토큰 관리 위치 | `am_Path.prv_LoadTokens` 내부 직접 관리 |
| 색상 센티널 값 | `-1` (0은 검정으로 유효값) |
| GetValueValidationReport | ❌ 제거 |
| AutoIDs 배치 | `am_DB` 하단 보조 프로시저 |
| BackupSheet/BackupWorkbook 배치 | `am_Sheet` |
| ReverseText | ❌ 제거 (VBA.StrReverse 직접 호출) |
| GetPath blnToken 인수명 | `blnToken` (기존 blntWB 에서 변경) |
| am_File 파일 다이얼로그 배치 | `am_File` (am_Utils 아님) |
| am_Excel 분리 | am_Sheet / am_Table / am_Range / am_Format / am_Excel(스텁) |
| tpl_Find / tpl_Range 이식 대상 | `am_Range` (기존 am_Excel 아님) |
| tpl_Validation 이식 대상 | `am_Utils` (기존과 동일) |
| Common.cls 이식 방식 | 일반 모듈로 이식 (클래스 불필요, Application.Run 미지원 문제) |
| tpl_Procedure | ❌ CWB 전용 결정 (xlam에서 ThisWorkbook = xlam 자신, Trust Center 의존) |
| frm_* 모듈 | ❌ CWB 전용 결정 (사용자 정의 폼은 특정 파일 종속) |
| SheetLock 입력셀 규약 | 배경색 없는 셀(ColorIndex=xlNone) = 입력 가능 셀 |

---

## 주요 기술 이슈 및 해결 방법

| 이슈 | 해결 방법 |
|---|---|
| `Workbooks` 컬렉션에 xlam 미포함 | `On Error Resume Next` 로 `Workbooks("corelib.xlam")` 직접 접근 시도 |
| xlam 로드 직후 프로시저 호출 컴파일 오류 | `Application.Run` 사용 (문자열로 전달하여 컴파일 시점 스캔 우회) |
| `RegisterWB` 딕셔너리 오류 시 초기화 | 딕셔너리 방식 폐기, 프로시저 인수로 CWB 명시적 전달 |
| AddIns 등록 시 관련 없는 파일에도 영향 | `Workbooks.Open` 방식으로 변경 |
| Optional 인수가 필수 인수 앞에 위치 시 오류 | Optional 인수는 반드시 필수 인수 뒤에 배치 |
| 기존 .bas/.cls 파일 한글 깨짐 | EUC-KR → UTF-8 재작성 |

---

## 변경 이력

| 날짜 | 내용 |
|---|---|
| 2026-05-21 | 프로젝트 시작, 기본 구조 설계 |
| 2026-05-21 | am_Core, am_Path, am_File, am_DB, am_Excel 완성 |
| 2026-05-21 | tpl_Path, tpl_Test, ThisWorkbook (xlam/cwb) 완성 |
| 2026-05-21 | tpl_Formatting (ConditionalFormattingFormula 등) am_Excel 에 추가 |
| 2026-05-21 | CLAUDE.md, LOG.md 생성 |
| 2026-06-02 | am_Excel → am_Sheet / am_Table / am_Range / am_Format / am_Excel(스텁) 으로 분리 |
| 2026-06-02 | 전체 .bas/.cls 파일 인코딩 EUC-KR → UTF-8 재작성 |
| 2026-06-02 | am_Core 보완: CM_TO_POINTS 상수, DPUpdate/Event/Calculate/WB_Lock 추가 (Common.cls, tpl_Const 이식) |
| 2026-06-02 | am_Range 완성: FindCellsByColor (tpl_Find), GetUsedRange (tpl_Range) 추가 |
| 2026-06-02 | am_Sheet 보완: SheetLock/SheetUnLock 추가 (Common.cls 재설계, prv_GetUsedRange·prv_FindCellsByColor 내부 구현) |
