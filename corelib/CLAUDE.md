# CLAUDE.md — corelib.xlam 프로젝트 지침

> 진행 현황, 작업 로그, 모듈별 완성 상태는 `LOG.md` 를 참고하세요.

---

## 1. 프로젝트 개요

Excel VBA 범용 라이브러리 xlam 파일을 제작하는 프로젝트입니다.
다양한 Excel 프로그램(CWB)에서 `corelib.xlam` 을 열어서 공통 함수를 사용하는 구조입니다.

### 핵심 구조

```
corelib.xlam       ← 공통 함수 라이브러리 (서비스 제공자)
cwb_01.xlsm        ← 클라이언트 워크북 (CWB, 사용자)
```

- **CWB (Client WorkBook)**: `corelib.xlam` 을 열어서 사용하는 Excel 파일
- xlam 은 `Workbooks.Open` 방식으로 로드 (AddIns 등록 방식 사용 안 함)
- `Workbooks` 컬렉션에 xlam 이 포함되지 않으므로 로드 확인은 **오류 방식**으로 체크

---

## 2. 폴더 구조

```
corelib/
├── CLAUDE.md                  ← 코딩 지침 (현재 파일)
├── LOG.md                     ← 진행 로그 (작업 현황 및 결정 사항)
│
├── xlam/                      ← corelib.xlam 모듈 소스
│   ├── 현재_통합_문서.cls     ← ThisWorkbook (corelib.xlam)
│   ├── am_Core.bas
│   ├── am_Path.bas
│   ├── am_File.bas
│   ├── am_DB.bas
│   ├── am_Excel.bas
│   ├── am_Utils.bas           ← 미작성
│   └── am_Error.bas           ← 미작성
│
└── cwb/                       ← cwb_01.xlsm 모듈 소스
    ├── 현재_통합_문서.cls     ← ThisWorkbook (cwb_01.xlsm)
    ├── tpl_Path.bas
    └── tpl_Test.bas
```

### 파일 규칙

- 모듈 파일명 = 모듈명 (예: `am_Core.bas`, `tpl_Path.bas`)
- `ThisWorkbook` 은 `.cls` 확장자 사용
- 일반 모듈은 `.bas` 확장자 사용

---

## 3. 모듈 구성

### corelib.xlam 모듈 목록

| 모듈 | 파일 | 역할 |
|---|---|---|
| `ThisWorkbook` | `현재_통합_문서.cls` | xlam 열림/닫힘 이벤트 |
| `am_Core` | `am_Core.bas` | 전역 상수, Property, 초기화/정리 |
| `am_Path` | `am_Path.bas` | 경로 토큰 변환, 경로 정규화 |
| `am_File` | `am_File.bas` | 파일/폴더 생성, 삭제, 복사, 검색, 다이얼로그 |
| `am_DB` | `am_DB.bas` | DB 연결, 쿼리 실행, 스키마 조회, 타입 처리 |
| `am_Excel` | `am_Excel.bas` | 시트, 테이블, 조건부 서식, 유효성 검사 등 엑셀 자동화 |
| `am_Utils` | `am_Utils.bas` | 문자열, 배열, 날짜 등 순수 범용 유틸리티 |
| `am_Error` | `am_Error.bas` | 공통 에러 핸들링 |

### cwb 템플릿 모듈 목록

| 모듈 | 파일 | 역할 |
|---|---|---|
| `ThisWorkbook` | `현재_통합_문서.cls` | xlam 로드, 이벤트 처리 |
| `tpl_Path` | `tpl_Path.bas` | `ResolvePath` 래퍼, 경로 관련 유틸리티 |
| `tpl_Test` | `tpl_Test.bas` | 테스트 프로시저 모음 |

---

## 4. 네이밍 규칙

### 모듈 접두사

| 접두사 | 소속 | 의미 |
|---|---|---|
| `am_` | corelib.xlam | Add-in Macro |
| `tpl_` | CWB 템플릿 | Template |

### 프로시저 접두사

| 선언 | 접두사 | 용도 |
|---|---|---|
| `Public` | 없음 | 외부 호출 가능한 함수 |
| `Private` | `prv_` | 모듈 내부 전용 함수 |

```vba
' 외부 호출 가능
Public Sub MkFolder(...)

' 모듈 내부 전용
Private Sub prv_MkFolder(...)
Private Function prv_GetExt(...) As String
```

### 변수 접두사 (Hungarian Notation)

#### 기본 데이터 타입

| 접두사 | 타입 |
|---|---|
| `bln` | Boolean |
| `byt` | Byte |
| `int` | Integer |
| `lng` | Long |
| `sng` | Single |
| `dbl` | Double |
| `cur` | Currency |
| `dec` | Decimal |
| `str` | String |
| `dt` | Date |
| `vnt` | Variant |

#### 구조 및 자료구조

| 접두사 | 타입 |
|---|---|
| `obj` | Object |
| `col` | Collection |
| `dic` | Dictionary |
| `arr` | Array |

#### 엑셀 개체

| 접두사 | 타입 |
|---|---|
| `wb` | Workbook |
| `ws` | Worksheet |
| `rng` | Range |
| `cel` | Cell |
| `cht` | Chart |
| `tbl` | ListObject |
| `shp` | Shape |

### UserForm 컨트롤 접두사

| 접두사 | 컨트롤 |
|---|---|
| `chk` | CheckBox |
| `cbo` | ComboBox |
| `cmd` | CommandButton |
| `fra` | Frame |
| `frm` | UserForm |
| `img` | Image |
| `lbl` | Label |
| `lst` | ListBox |
| `opt` | OptionButton |
| `txt` | TextBox |
| `tgl` | ToggleButton |

---

## 5. 코드 스타일

### 섹션 구분자

```vba
' ══════════════════════════════════════════════════════════
'  섹션명
' ══════════════════════════════════════════════════════════
```

### 단락 구분자

```vba
' ── 1. 단락명 ───────────────────────────────────────────────
```

### 모듈 헤더

```vba
Option Explicit

' ┌─────────────────────────────────────────────────────────┐
' │  am_ModuleName                                          │
' │  역할 : 모듈 역할 설명                                  │
' └─────────────────────────────────────────────────────────┘
```

### 프로시저 헤더 주석

```vba
' 목적   : 프로시저 설명
' 인수   : strPath - 설명
'          blnFlag - 설명
' 반환   : String - 설명
' 예시   : FunctionName("value") → "result"
Public Function FunctionName(...) As String
```

### 에러 핸들링 패턴

```vba
Public Sub SomeSub()

    On Error GoTo ErrHandler

    ' 코드

    Exit Sub

ErrHandler:
    MsgBox "오류 " & Err.Number & ": " & Err.Description, _
           vbCritical, am_Core.AM_NAME

End Sub
```

### CleanUp 패턴 (개체 해제 필요 시)

```vba
Public Function SomeFunc() As Boolean

    On Error GoTo ErrHandler

    ' 코드

    SomeFunc = True
    GoTo CleanUp

ErrHandler:
    SomeFunc = False

CleanUp:
    On Error Resume Next
    If Not objConn Is Nothing Then
        If objConn.State = 1 Then objConn.Close
    End If
    Set objConn = Nothing

End Function
```

---

## 6. 설계 원칙

### 모듈 독립성

- **모듈 간 직접 호출 금지**
- 각 모듈은 독립적으로 동작 가능해야 함
- 다른 모듈에 있는 함수와 동일한 기능이 필요하면 `prv_` 접두사로 내부에 직접 구현

```vba
' ❌ 금지 - 다른 모듈 직접 호출
Call am_File.MkFolder(strPath)

' ✅ 허용 - 내부 전용 함수로 자체 구현
Private Sub prv_MkFolder(ByVal strPath As String)
    ' 직접 구현
End Sub
```

### CWB 경로 전달

- CWB 를 구분할 때 **프로시저 인수에 명시적으로 전달**
- `RegisterWB` 딕셔너리 방식 사용 안 함 (오류 시 초기화 문제)

```vba
' ✅ CWB 경로 명시적 전달
Application.Run "corelib.xlam!am_Path.ReplacePath", _
                strPath, _
                ThisWorkbook.Path, _
                ThisWorkbook.FullName
```

### Late Binding

- 외부 라이브러리 참조 없이 `CreateObject` 방식 사용

```vba
Set objConn = CreateObject("ADODB.Connection")
Set objFSO  = CreateObject("Scripting.FileSystemObject")
Set dicData = CreateObject("Scripting.Dictionary")
```

### 배열 인덱스

- 항상 `LBound` / `UBound` 사용

```vba
For i = LBound(arrData) To UBound(arrData)
```

### 센티널 값

- 색상 인수에서 `0` 은 검정(유효값)이므로 `-1` 을 "미적용" 센티널로 사용

```vba
Optional ByVal lngFontColor As Long = -1
If lngFontColor <> -1 Then .Font.Color = lngFontColor
```

### Optional 인수 순서

- `Optional` 인수는 반드시 필수 인수 뒤에 위치

```vba
Public Sub DelTableFilteredRows(ByVal strFieldName     As String, _
                                ByVal strFilterPattern As String, _
                                Optional ByVal tbl     As ListObject)
```

---

## 7. 경로 토큰 시스템

### 고정 토큰

| 토큰 | 의미 |
|---|---|
| `{cPath}` | CWB 폴더 경로 |
| `{cFile}` | CWB 전체 파일 경로 |
| `{xPath}` | xlam 폴더 경로 |
| `{xFile}` | xlam 전체 파일 경로 |

### 커스텀 토큰

`am_Path` 의 `prv_LoadTokens` 에서 직접 관리

```vba
arrTokens = Array( _
    "xDB",  am_Core.XlamPath & "\DB", _
    "xBak", am_Core.XlamPath & "\Backup" _
)
```

---

## 8. xlam 로드 패턴

### CWB 의 `Workbook_Open` 표준 패턴

```vba
Private Sub Workbook_Open()

    On Error GoTo ErrHandler

    Dim strXlamPath As String
    Dim wbXlam      As Workbook

    strXlamPath = ThisWorkbook.Path & "\corelib.xlam"

    If Dir(strXlamPath) = "" Then
        MsgBox "corelib.xlam 파일을 찾을 수 없습니다." & vbCrLf & _
               strXlamPath, vbCritical, "corelib"
        Exit Sub
    End If

    ' Workbooks 컬렉션에 xlam 미포함 → 오류 방식으로 체크
    On Error Resume Next
    Set wbXlam = Workbooks("corelib.xlam")
    On Error GoTo ErrHandler

    If wbXlam Is Nothing Then
        Workbooks.Open strXlamPath
    End If

    Exit Sub

ErrHandler:
    MsgBox "corelib.xlam 로드 중 오류가 발생했습니다." & vbCrLf & _
           "오류 " & Err.Number & ": " & Err.Description, _
           vbCritical, "corelib"

End Sub
```

### xlam 프로시저 호출 패턴

```vba
' 반환값 없는 경우
Application.Run "corelib.xlam!am_Path.ReplacePath", arg1, arg2

' 반환값 있는 경우
Dim strResult As String
strResult = Application.Run("corelib.xlam!am_Path.ReplacePath", arg1, arg2)
```

---

## 9. am_Core 상수

```vba
Public Const AM_VERSION As String = "1.0.0"
Public Const AM_NAME    As String = "corelib"
Public Const AM_FILE    As String = "corelib.xlam"
```

---

## 10. am_DB 상수

```vba
Public Const DB_ODBC_DRIVER As String = "MySQL ODBC 8.2 Unicode Driver"
Public Const DB_EXCEL_VER   As String = "12.0"
Public Const DB_ACCESS_VER  As String = "12.0"

Public Const DB_TYPE_SERVER As String = "서버"
Public Const DB_TYPE_EXCEL  As String = "엑셀"
Public Const DB_TYPE_ACCESS As String = "엑세스"
```

---

## 11. 개발 환경

- Excel 2010 32-bit
- Microsoft ADO Ext. 2.8
- Microsoft ActiveX Data Objects 2.8 Library
- Late Binding 방식 (라이브러리 참조 없이 동작)
- 코드 주석 언어: **한국어**

---

## 12. 작업 진행 방식

- 소스 모듈을 하나씩 제공하면 분석 → 협의 → 확정 → 코드 작성 순서로 진행
- 뒤에 제공된 코드로 앞 모듈 수정이 필요하면 해당 모듈 전체를 다시 제공
- 중복 함수 발견 시 통합 또는 제거 후 알림
- 시트 참조 코드는 가능한 VBA 내부로 전환, 방향이 애매하면 먼저 협의
- 진행 현황 및 결정 사항은 `LOG.md` 에 기록
