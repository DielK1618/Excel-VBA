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
| `am_DB.bas` | `am_DB` | ✅ 완성 | tpl_MsSQL/MySQL_Sub/Access 이식 완료 |
| `am_Sheet.bas` | `am_Sheet` | ✅ 완성 | 워크북/시트 조작, 백업, 시트 보호 추가 |
| `am_Table.bas` | `am_Table` | ✅ 완성 | 테이블 CRUD/필터/정렬/검색 |
| `am_Range.bas` | `am_Range` | ✅ 완성 | |
| `am_Format.bas` | `am_Format` | ✅ 완성 | 조건부 서식, 유효성 검사 |
| `am_Excel.bas` | `am_Excel` | ✅ 완성 | tpl_ExportFile/Chart/Shapes 이식 완료 |
| `am_Automation.bas` | `am_Automation` | ✅ 완성 | tpl_KeyBoard/Mouse(am_Excel에서 재이동) + 외부 제공 소스(창/버튼 제어, UIA) 이식 완료 |
| `am_Utils.bas` | `am_Utils` | ✅ 완성 | tpl_Array/Check/Code/ExtApp/Media/ReplaceText/Tools/Validation 이식 완료 |
| `am_Error.bas` | `am_Error` | ✅ 완성 | HandleError / WriteLog |

### cwb_01.xlsm (`cwb/` 폴더)

#### ref_ 모듈 (기본 포함 — am_ 래퍼)

| 파일 | 모듈 | 상태 |
|---|---|---|
| `ref_Core.bas` | `ref_Core` | ✅ 완성 |
| `ref_Path.bas` | `ref_Path` | ✅ 완성 |
| `ref_File.bas` | `ref_File` | ✅ 완성 |
| `ref_DB.bas` | `ref_DB` | ✅ 완성 |
| `ref_Excel.bas` | `ref_Excel` | ✅ 완성 |
| `ref_Automation.bas` | `ref_Automation` | ✅ 완성 |
| `ref_Range.bas` | `ref_Range` | ✅ 완성 |
| `ref_Sheet.bas` | `ref_Sheet` | ✅ 완성 |
| `ref_Table.bas` | `ref_Table` | ✅ 완성 |
| `ref_Format.bas` | `ref_Format` | ✅ 완성 |
| `ref_Utils.bas` | `ref_Utils` | ✅ 완성 |
| `ref_Error.bas` | `ref_Error` | ✅ 완성 |

#### tpl_ 모듈 (필요시 작성)

| 파일 | 모듈 | 상태 | 비고 |
|---|---|---|---|
| `현재_통합_문서.cls` | `ThisWorkbook` | ✅ 완성 | |
| `tpl_Path.bas` | `tpl_Path` | ✅ 완성 | |
| `tpl_Test.bas` | `tpl_Test` | ✅ 완성 | am_Core/Path/File/Utils 전 모듈 테스트 포함 |

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
| `tpl_Error` | `am_Error` | ✅ 완료 |
| `tpl_MsSQL` | `am_DB` | ✅ 완료 |
| `tpl_MySQL_Sub` (일부) | `am_DB` | ✅ 완료 |
| `tpl_Access` | `am_DB` | ✅ 완료 |
| `tpl_Array` | `am_Utils` | ✅ 완료 |
| `tpl_Check` (일부) | `am_Utils` | ✅ 완료 |
| `tpl_Code` (일부) | `am_Utils` | ✅ 완료 |
| `tpl_ExtApp` | `am_Utils` | ✅ 완료 |
| `tpl_Media` (일부) | `am_Utils` | ✅ 완료 |
| `tpl_ReplaceText` (일부) | `am_Utils` | ✅ 완료 |
| `tpl_Tools` (일부) | `am_Utils` | ✅ 완료 |
| `tpl_Validation` | `am_Utils` | ✅ 완료 |
| `tpl_Chart` | `am_Excel` | ✅ 완료 |
| `tpl_ExportFile` | `am_Excel` | ✅ 완료 |
| `tpl_Shapes` | `am_Excel` | ✅ 완료 |
| `tpl_KeyBoard` | `am_Automation` | ✅ 완료 (am_Excel → am_Automation 재이동) |
| `tpl_Mouse` | `am_Automation` | ✅ 완료 (am_Excel → am_Automation 재이동) |
| 외부 제공 소스(창/버튼 제어, UIA) | `am_Automation` | ✅ 완료 |
| `tpl_Buttons`, `tpl_Buttons_other`, `tpl_Buttons_Top` | — | ❌ CWB 전용 |
| `tpl_Form`, `tpl_TestBed`, `tpl_Procedure` | — | ❌ CWB 전용 |
| `frm_*` (전체) | — | ❌ CWB 전용 (사용자 직접 제작) |

---

## 확정된 설계 결정 사항

| 항목 | 결정 내용 |
|---|---|
| xlam 파일명 | `corelib.xlam` |
| xlam 모듈 접두사 | `am_` (Add-in Macro) |
| CWB ref_ 모듈 접두사 | `ref_` (Reference) — am_ 1:1 래퍼, 기본 포함 |
| CWB tpl_ 모듈 접두사 | `tpl_` (Template) — CWB 비즈니스 로직, 필요시 작성 |
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
| SheetLock 입력셀 규약 | 배경색 없는 셀(Interior.Pattern=xlNone) = 입력 가능 셀 |
| am_Excel vs am_Automation 경계 | am_Excel=Excel 객체 모델(시트/차트/도형) 전용, am_Automation=OS 레벨 입력·창 제어(키보드/마우스 시뮬레이션, FindWindow/UIA) 전용으로 분리 |
| UIA(UI Automation) 바인딩 방식 | **Early Binding**(`New CUIAutomation`, VBE 참조에 UIAutomationClient 추가) — Late Binding(`CreateObject`)은 이 환경에서 런타임 429로 실패 확인되어 Late Binding 원칙의 명시적 예외로 확정 (2026-07-24 실측, TROUBLESHOOTING.md TS-V16) |

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
| 2026-07-24 | **ClickButtonByUIA — Late Binding → Early Binding 확정**: 실제 개발 환경(Excel 2010 32bit)에서 `CreateObject("UIAutomationClient.CUIAutomation")` 테스트 결과 런타임 429 오류("ActiveX 구성 요소는 개체를 만들 수 없습니다") 재현 확인. `New CUIAutomation` + `IUIAutomationElement`/`IUIAutomationCondition`/`IUIAutomationInvokePattern` Early Binding으로 전환(VBE 참조에 UIAutomationClient 라이브러리 추가 필요) — 프로젝트 Late Binding 원칙의 명시적 예외로 `CLAUDE.md`(§Late Binding, §UIA 바인딩 방식) 반영. `TROUBLESHOOTING.md`에 창/버튼 자동화 트러블슈팅 6건(TS-V12~V17) + 설계 원칙 신규 기록 |
| 2026-07-24 | **am_Automation 신규 모듈 분리 및 창/버튼 제어 이식** — am_Excel의 키보드/마우스 섹션(tpl_KeyBoard·tpl_Mouse, Windows API 선언·Enum·ExecuteKeyAction/Sequence·ProcessRangeWithKeySequence·GetMousePosition·ClickAtPosition·WaitTime 전체)이 Excel 객체 모델과 무관한 OS 레벨 자동화라 판단하여 am_Automation.bas로 재이동(am_Excel은 인쇄/차트/도형만 남김). 동시에 외부 제공 소스(원본 파일명 미상, `_BT` 접미사 네이밍)에서 창 탐색/버튼 클릭/강제 활성화/UI Automation 기능을 이식: `ForceActivateWindow`(AttachThreadInput), `ClickButtonByText`(SendMessage BM_CLICK), `ClickButtonByRect`/`ClickButtonByRect_Retry`(좌표 클릭, ClickAtPosition 내부 재사용), `ClickButtonByUIA`(UI Automation Invoke). `_BT` 접미사 전체 제거, Public 전역변수(`g_hWndFound`/`g_strTargetText`) → 모듈 Private(`m_`) 전환, `EnumChildProc` → `prv_EnumChildProc`. UIA는 원본 Early Binding(`New CUIAutomation`, 참조 라이브러리 필요)을 Late Binding(`CreateObject("UIAutomationClient.CUIAutomation")`)으로 전환하여 프로젝트 원칙(참조 추가 없이 CreateObject 사용) 준수. 진단용 `Diag_*` 4개(WMC SCDK 하드코딩 등)는 CWB 전용/이식 제외. `ref_Excel.bas`에서 키보드/마우스 래퍼 제거, `ref_Automation.bas` 신규 작성(전 함수 래핑, ParamArray/ByRef 함수 3개는 기존 관례대로 래핑 불가 명시). `MIGRATE.md`/`refer/SOURCES.md`/`CLAUDE.md`(폴더구조·모듈목록) 동반 갱신 |
| 2026-06-16 | **LOCK_PW 상수 도입 — 시트/워크북 보호 기본 비밀번호**: am_Core에 `Public Const LOCK_PW As String = "1234"` 선언, SheetLock/SheetUnLock(am_Sheet) · WB_Lock/WB_UnLock(am_Core) strPW를 Optional로 변경(기본값 LOCK_PW). ref_Core에 동일 상수 선언, ref_Sheet/ref_Core 래퍼 동일 패턴 적용. 비밀번호 미입력 시 LOCK_PW 자동 사용, 필요 시 명시 전달 가능. |
| 2026-06-16 | **Optional 파라미터 패턴 전체 적용** — ws/wb/tbl 필수 인수를 Optional로 변경(기본: ActiveSheet/ActiveWorkbook/ActiveSheet.ListObjects(1)), 파라미터 재정렬(Optional은 뒤로). 변경 대상: am_Sheet(8개), am_Table(4개), am_Excel(ExportSheetToCSV), am_DB(DelExcelRecQuery), ref_Sheet(8개), ref_Table(4개), ref_Excel(ExportSheetToCSV), ref_DB(DelExcelRecQuery). ref_ 래퍼는 Optional 해소 후 명시값 전달 패턴 적용. cwb ref_ 전 모듈(11개) 프로시저 헤더 주석 추가 — 목적/인수/반환/예시 형식 통일. am_Range.FindCellsByColor 버그 수정: Find(What:="") 이전 검색어 재사용 문제 → Find(What:="*") + SpecialCells(xlCellTypeBlanks) 하이브리드로 교체. am_Sheet.SheetLock 전면 재설계: (1) Interior.ColorIndex → Interior.Pattern 으로 변경(테마/RGB 색상 감지 누락 수정), (2) prv_GetUsedRange(내용 기반) → ws.UsedRange(서식 포함, 빈 회색 셀 포함)로 교체, (3) 셀 루프 잠금 → SearchFormat+SpecialCells 하이브리드로 최적화(내용셀 네이티브 Find, 빈셀만 루프, rngUnlock 벌크 일괄 해제) |
| 2026-06-12 | corelib_manual.html 리뉴얼 완성 — Lamborghini 디자인(진블랙·골드), 반응형 드로어 사이드바, 전체 개요↔모듈 앵커 내비게이션, 누락 프로시저 42개 전체 문서화(am_DB 9개 포함, SelectExcelQuery·GetDbInfo 미구현 경고 표기) |
| 2026-06-11 | corelib_manual_cd.html 신규 — corelib_manual.html 을 Claude Design 기반으로 리뉴얼한 사용자 매뉴얼 완성 |
| 2026-06-11 | corelib_manual.html 신규 — 11개 모듈 전체 프로시저 설명·코드예시·사이드바·전체 개요 섹션(이름+한 줄 설명) 포함 HTML 매뉴얼 제작 |
| 2026-06-11 | 전 모듈 테스트 완료 — am_Core/Path/File/Utils/Error/Range/Format/Table/Sheet/Excel 전체 OK |
| 2026-06-11 | am_Core.IsReady 버그 수정: VBE 리셋 후 m_blnReady=False 유지 → lazy-init(미초기화 시 Initialize() 자동 호출)으로 해결 |
| 2026-06-11 | am_Table.AutoTableFilter_Arr 버그 수정: arrWildCards() As String → As Variant (Application.Run Variant 배열 수신 오류 방지) |
| 2026-06-11 | am_File.DelFolder 버그 수정: Kill+RmDir → FSO.DeleteFolder(True) 재귀 삭제 (중첩 폴더 미삭제 문제 해결) |
| 2026-06-11 | tpl_Test.Test_EdgeCases 수정: 빈 문자열·잘못된 경로 케이스 PrintResult("","") → PrintBool(=""확인) 로 변경 |
| 2026-06-11 | am_Error 테스트 추가: ENABLE_ERROR_LOG Const → m_blnLogEnabled + SetLogEnabled/GetLogEnabled 으로 변경, Test_Error 신규 (6건 전체 OK) |
| 2026-06-11 | am_Sheet.BackupWorkbook 버그 수정: prv_MkFolder 호출 누락 — 폴더 없을 때 SaveCopyAs 실패 방지 |
| 2026-06-10 | Application.Run 호환성 버그 수정 및 Test_Range 전체 통과 — am_Core `Property Get` 4개 → `Public Function` 변환(`XlamPath`/`XlamFullName`/`Version`/`IsReady`), am_Path `ReplacePath` MsgBox 제거·`prv_IsDriveAccessible` 신규, am_Sheet.`BackupSheet`·am_Excel.`ExportSheetToCSV` `DisplayAlerts` 래핑 추가, am_Range 스칼라 래퍼 6개 신규(`GetUsedRange_IsValid`/`RowCount`/`ColCount`·`FindRange_IsValid`/`CellValue`·`FindCellsByColor_Count`), tpl_Test 상수명 충돌 수정(`TEST_SHEET`→`TEST_SHEET_NM`)·`Setup_TestSheet` `NumberFormat "@"` 추가·`Test_Range` 전면 재작성 |
| 2026-06-09 | am_Sheet.GetSheetNames 버그 수정: `Application.WorksheetFunction.Transpose` 제거 → 1차원 배열 직접 반환. tpl_Test.Test_Sheet에서 arrNames(i) 단일 인덱스 접근 시 2D 배열 오류 방지 |
| 2026-06-08 | tpl_Test 전 모듈 테스트 전면 재작성: Setup/Teardown_TestSheet(자동 시트·테이블 생성·정리), am_Core/Path/File/Utils/Range/Format/Table/Sheet/Excel 전 모듈 커버, RunGetRng 헬퍼(Range 반환 안전처리), prv_CountHiddenSheets/CountVisibleSheets 추가 |
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
| 2026-06-04 | am_Error 신규 작성: HandleError / WriteLog (tpl_Error 이식, cl.* 제거, am_Core.AM_NAME·XlamPath 으로 교체) |
| 2026-06-04 | am_DB 확장: tpl_MsSQL/MySQL_Sub/Access 이식 완료 — DelExcelRecQuery(wb 인수화), GetFields(센티널 -1), GetFieldAndType/GetFieldNameConnection(ReplaceFields·GetDbInfo 의존성 제거), CreateAccessTable/DeleteAccessTable/CreateAccessTableADOX/DeleteAccessFields 추가 |
| 2026-06-05 | am_Utils 신규 작성: tpl_Array/Check/Code/ExtApp/Media/ReplaceText/Tools/Validation 이식 — IsRangeMerged 버그 2종 수정(IsSelectionMerged 오타·Selection Is Nothing), GenerateRandomCode 버그 수정(j>maxAttempts), tpl_Procedure·SyncCodeNamesToSheetNames는 VBProject/Trust Center 의존으로 기존 결정대로 제외 |
| 2026-06-05 | am_Excel 완성: tpl_ExportFile/Chart/Shapes/KeyBoard/Mouse 이식 — xlCSVUTF8→xlCSV(Excel 2010 호환), Call MkFolder→prv_MkFolder(모듈 독립성), SetChartDataRange에 ws 인수 추가, ClickAtPosition 타입 수정(BoolLeft As String→blnLeft As Boolean), prv_WaitMs로 내부 대기 통합, GetMousePosition을 Function→Sub으로 변경 |
