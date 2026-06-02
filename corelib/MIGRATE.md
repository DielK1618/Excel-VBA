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
| `tpl_Error` | ⬜ 대기 | 공통 에러 핸들링 |

### am_DB
| 원본 | 상태 | 비고 |
|---|---|---|
| `tpl_Access` | ⬜ 대기 | Access DB 연결/쿼리 |
| `tpl_MsSQL` | ⬜ 대기 | MsSQL 연결/쿼리 |
| `tpl_MySQL_Sub` | ⬜ 대기 | MySQL 보조 프로시저 |

### am_Range
| 원본 | 상태 | 비고 |
|---|---|---|
| `tpl_Find` | ✅ 완료 | 범위 검색 |
| `tpl_Range` | ✅ 완료 | 범위 조작 |

### am_Excel
| 원본 | 상태 | 비고 |
|---|---|---|
| `tpl_Chart` | ⬜ 대기 | 차트 생성/조작 |
| `tpl_ExportFile` | ⬜ 대기 | 파일 내보내기 |
| `tpl_KeyBoard` | ⬜ 대기 | 단축키 등록/해제 |
| `tpl_Mouse` | ⬜ 대기 | 마우스 이벤트 |
| `tpl_Shapes` | ⬜ 대기 | 도형 조작 |

### am_Utils (신규 모듈)
| 원본 | 상태 | 비고 |
|---|---|---|
| `tpl_Array` | ⬜ 대기 | 배열 유틸리티 |
| `tpl_Check` | ⬜ 대기 | 값 검사 유틸리티 |
| `tpl_Code` | ⬜ 대기 | 코드 변환 유틸리티 |
| `tpl_ExtApp` | ⬜ 대기 | 외부 앱 실행 |
| `tpl_Media` | ⬜ 대기 | 미디어(사운드 등) |
| `tpl_ReplaceText` | ⬜ 대기 | 문자열 치환 |
| `tpl_Tools` | ⬜ 대기 | 기타 도구 |
| `tpl_Validation` | ⬜ 대기 | 입력 유효성 검사 |

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
