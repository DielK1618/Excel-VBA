# MIGRATE.md — xlam 이식 작업 인덱스

> 작업 시 vba.xlsm 직접 참조 대신 이 파일을 기준으로 진행
> 완료 시 `⬜ → ✅` 로 변경, LOG.md 변경 이력에도 기록

---

## 이식 대기 목록

### am_Core
| 원본 | 상태 | 비고 |
|---|---|---|
| `tpl_Const` | ✅ 완료 | CM_TO_POINTS 상수 이식 |
| `Common.cls` | ✅ 완료 | DPUpdate / Event / Calculate / WB_Lock 이식 |

### am_Error (신규 모듈)
| 원본 | 상태 | 비고 |
|---|---|---|
| `tpl_Error` | ✅ 완료 | 공통 에러 핸들링 |

### am_DB
| 원본 | 상태 | 비고 |
|---|---|---|
| `tpl_Access` | ✅ 완료 | Access DB 연결/쿼리 |
| `tpl_MsSQL` | ✅ 완료 | MsSQL 연결/쿼리 |
| `tpl_MySQL_Sub` | ✅ 완료 | MySQL 보조 프로시저 (CWB 종속 함수 제외) |

### am_Range
| 원본 | 상태 | 비고 |
|---|---|---|
| `tpl_Find` | ✅ 완료 | 범위 검색 |
| `tpl_Range` | ✅ 완료 | 범위 조작 |

### am_Excel
| 원본 | 상태 | 비고 |
|---|---|---|
| `tpl_Chart` | ✅ 완료 | SetChartDataRange (ws 인수 추가) |
| `tpl_ExportFile` | ✅ 완료 | SetPrintPage, ExportPDF, ExportSheetToCSV (xlCSV, prv_MkFolder 내부 구현) |
| `tpl_Shapes` | ✅ 완료 | RunShpMacro, GetShapeTextSafe (prv_GetShapeTextSafe_GItem) |

### am_Automation (신규 모듈, 2026-07-24 분리)
| 원본 | 상태 | 비고 |
|---|---|---|
| `tpl_KeyBoard` | ✅ 완료 | Enum + API + ExecuteKeyAction/Sequence/Range (prv_SendKey/Combo/WaitMs). 최초 am_Excel 이식(2026-06-05) 후 2026-07-24 am_Automation으로 재이동 — Excel 객체 모델과 무관한 OS 레벨 입력 시뮬레이션이라 분리 |
| `tpl_Mouse` | ✅ 완료 | GetMousePosition, ClickAtPosition, WaitTime (타입 수정, prv_WaitMs 공용). 위와 동일 사유로 am_Automation 재이동 |
| (외부 제공 소스, 원본 파일명 미상) | ✅ 완료 | Win32 SendMessage/좌표 클릭 버튼 제어(ClickButtonByText/Rect/Rect_Retry), AttachThreadInput 강제 창 활성화(ForceActivateWindow), UI Automation 버튼 Invoke(ClickButtonByUIA). `_BT` 접미사 제거, UIA는 Early Binding → Late Binding(`CreateObject("UIAutomationClient.CUIAutomation")`) 전환, 진단용 `Diag_*` 4개(WMC SCDK 하드코딩 등)는 CWB 전용/제외 처리 |

### am_Utils (신규 모듈)
| 원본 | 상태 | 비고 |
|---|---|---|
| `tpl_Array` | ✅ 완료 | ConvertToArrData, prv_FlattenArray, prv_CountElements |
| `tpl_Check` | ✅ 완료 | IsArrayEmpty, IsCells, IsTableRange, IsRangeMerged(버그수정), IsValidFileName, GetValidationType (AccessTableExists → am_DB 배치) |
| `tpl_Code` | ✅ 완료 | CreateUniqueID, GenerateRandomCode, prv_CheckUniqueID (BtCreateCodes/GetExistingCodes 제외) |
| `tpl_ExtApp` | ✅ 완료 | OpenAddressInGoogleMaps, GetVideoLength |
| `tpl_Media` | ✅ 완료 | GetVideoLength는 tpl_ExtApp 것으로 통합, GetVideoDuration 제외 |
| `tpl_ReplaceText` | ✅ 완료 | ConvertToExcelSerialDate, ExtractValues (ReplaceText 제외) |
| `tpl_Tools` | ✅ 완료 | CheckSelectionType, WaitMs (VBProject 의존 함수 전체 제외) |
| `tpl_Validation` | ✅ 완료 | EvaluateFormula, SetIfValTrue |

---

## CWB 전용 (이식 제외)

| 모듈 | 사유 |
|---|---|
| `tpl_Buttons` | 특정 워크북 버튼 UI |
| `tpl_Buttons_other` | 특정 워크북 버튼 UI |
| `tpl_Buttons_Top` | 특정 워크북 버튼 UI |
| `tpl_Form` | 특정 워크북 전용 |
| `tpl_Procedure` | VBA 메타프로그래밍 — xlam에서 ThisWorkbook이 xlam 자신을 가리킴, 보안 설정 의존 |
| `tpl_TestBed` | 테스트 전용 |
| `frm_*` (전체) | 사용자 정의 폼은 특정 파일 종속 |
| `Diag_UIA_TestCreate_EarlyBinding_BT`, `Diag_UIA_FindWindow_BT`, `Diag_UIA_ListButtons_BT`, `Diag_CountWmcWindows_BT` (외부 제공 소스) | 진단/디버그 전용, `WMC SCDK` 등 특정 앱명 하드코딩 |

---

## 이식 완료

| 원본 | am_ 모듈 | 완료일 |
|---|---|---|
| `tpl_File` | `am_File` | 2026-05-21 |
| `tpl_Path` | `am_Path` | 2026-05-21 |
| `tpl_Sheet` | `am_Sheet` | 2026-06-02 |
| `Common.cls` (sht_Lock) | `am_Sheet` | 2026-06-02 |
| `tpl_Table` | `am_Table` | 2026-06-02 |
| `tpl_Formatting` | `am_Format` | 2026-06-02 |
| `tpl_Error` | `am_Error` | 2026-06-04 |
| `tpl_MsSQL` | `am_DB` | 2026-06-04 |
| `tpl_MySQL_Sub` (일부) | `am_DB` | 2026-06-04 |
| `tpl_Access` | `am_DB` | 2026-06-04 |
| `tpl_Array` | `am_Utils` | 2026-06-05 |
| `tpl_Check` (일부) | `am_Utils` | 2026-06-05 |
| `tpl_Code` (일부) | `am_Utils` | 2026-06-05 |
| `tpl_ExtApp` | `am_Utils` | 2026-06-05 |
| `tpl_Media` (일부) | `am_Utils` | 2026-06-05 |
| `tpl_ReplaceText` (일부) | `am_Utils` | 2026-06-05 |
| `tpl_Tools` (일부) | `am_Utils` | 2026-06-05 |
| `tpl_Validation` | `am_Utils` | 2026-06-05 |
| `tpl_Chart` | `am_Excel` | 2026-06-05 |
| `tpl_ExportFile` | `am_Excel` | 2026-06-05 |
| `tpl_Shapes` | `am_Excel` | 2026-06-05 |
| `tpl_KeyBoard` | `am_Automation` | 2026-06-05 (am_Excel → am_Automation 재이동 2026-07-24) |
| `tpl_Mouse` | `am_Automation` | 2026-06-05 (am_Excel → am_Automation 재이동 2026-07-24) |
| (외부 제공 소스, 원본 파일명 미상) | `am_Automation` | 2026-07-24 |
