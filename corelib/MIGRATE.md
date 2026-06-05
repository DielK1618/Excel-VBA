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
| `tpl_KeyBoard` | ✅ 완료 | Enum + API + ExecuteKeyAction/Sequence/Range (prv_SendKey/Combo/WaitMs) |
| `tpl_Mouse` | ✅ 완료 | GetMousePosition, ClickAtPosition, WaitTime (타입 수정, prv_WaitMs 공용) |
| `tpl_Shapes` | ✅ 완료 | RunShpMacro, GetShapeTextSafe (prv_GetShapeTextSafe_GItem) |

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
| `tpl_KeyBoard` | `am_Excel` | 2026-06-05 |
| `tpl_Mouse` | `am_Excel` | 2026-06-05 |
| `tpl_Shapes` | `am_Excel` | 2026-06-05 |
