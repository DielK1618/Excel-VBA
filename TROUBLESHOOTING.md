# TROUBLESHOOTING.md — scripts-vba

오류 발생 시 이 문서를 먼저 확인한다. 해결책이 없으면 LOG.md / 웹 검색 순으로 진행하고, 해결 후 이 문서에 추가한다.

---

## xlam 로드 / 호출

### TS-V01 · `Workbooks("corelib.xlam")` — 개체를 찾을 수 없음 오류

**증상**: xlam이 열려 있는데도 `Workbooks("corelib.xlam")`가 오류 발생

**원인**: `Workbooks` 컬렉션에 xlam 파일이 포함되지 않는 Excel 동작

**해결**: `On Error Resume Next`로 접근 시도 후 `Is Nothing` 확인

```vba
On Error Resume Next
Set wbXlam = Workbooks("corelib.xlam")
On Error GoTo ErrHandler
If wbXlam Is Nothing Then Workbooks.Open strXlamPath
```

---

### TS-V02 · xlam 로드 직후 함수 호출 — 컴파일 오류

**증상**: `corelib.xlam` 로드 직후 xlam 함수를 직접 호출하면 컴파일 시점 오류

**원인**: VBA 컴파일러가 로드 전 참조를 스캔

**해결**: `Application.Run`으로 문자열 전달 (컴파일 시점 스캔 우회)

```vba
Application.Run "corelib.xlam!am_Path.ReplacePath", arg1, arg2
```

---

### TS-V03 · `Property Get` 함수 — `Application.Run` 호출 불가

**증상**: `Application.Run "corelib.xlam!am_Core.XlamPath"` 오류

**원인**: `Application.Run`은 `Property Get`을 지원하지 않음

**해결**: `Property Get` → `Public Function`으로 변경

```vba
' ❌
Public Property Get XlamPath() As String

' ✅
Public Function XlamPath() As String
```

**적용**: am_Core의 `XlamPath` / `XlamFullName` / `Version` / `IsReady` (반영 완료)

---

### TS-V04 · `Optional` 인수가 필수 인수 앞에 위치 — 컴파일 오류

**증상**: 프로시저 선언 시 "선택적 인수가 필수 인수 앞에 올 수 없습니다" 컴파일 오류

**원인**: VBA 규칙 — `Optional` 인수는 반드시 필수 인수 뒤에 위치해야 함

**해결**: 인수 순서 재정렬 — 필수 인수 먼저, Optional 인수는 뒤로

```vba
' ❌
Public Sub SheetLock(Optional ws As Worksheet, ByVal strPW As String)

' ✅
Public Sub SheetLock(ByVal strPW As String, Optional ws As Worksheet)
```

---

## 파일 인코딩

### TS-V05 · `.bas` / `.cls` 파일 한글 깨짐

**증상**: VBE에서 가져온 모듈 파일의 한글이 깨짐

**원인**: VBE 기본 내보내기가 EUC-KR로 저장됨

**해결**: 파일을 UTF-8(BOM 없이)로 재작성 후 VBE에서 다시 가져오기. 신규 모듈은 처음부터 UTF-8로 작성.

---

## 배열 / 타입 오류

### TS-V06 · `Application.Run` — `Variant` 배열 수신 오류

**증상**: `Application.Run`으로 배열을 전달받을 때 타입 불일치 오류

**원인**: `Application.Run`은 `String()` 배열 등 특정 타입 배열을 제대로 전달하지 못함

**해결**: 배열 파라미터는 `As Variant`로 선언

```vba
' ❌
Public Sub AutoTableFilter_Arr(arrWildCards() As String)

' ✅
Public Sub AutoTableFilter_Arr(arrWildCards As Variant)
```

---

### TS-V07 · `GetSheetNames` — 2차원 배열로 반환되어 인덱스 오류

**증상**: `arrNames(i)` 접근 시 "인덱스가 유효 범위에 없습니다" 오류

**원인**: `WorksheetFunction.Transpose`가 1차원 배열을 2차원 배열로 변환

**해결**: `Transpose` 제거 → 1차원 배열 직접 구성 후 반환

---

## 파일 / 폴더

### TS-V08 · `DelFolder` — 중첩 폴더 미삭제

**증상**: 하위 폴더가 있는 폴더 삭제 시 오류 또는 일부만 삭제됨

**원인**: `Kill + RmDir`는 빈 폴더만 삭제 가능, 중첩 구조 처리 불가

**해결**: FSO `DeleteFolder(True)`로 재귀 삭제

```vba
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FolderExists(strPath) Then
    objFSO.DeleteFolder strPath, True
End If
```

---

## Range / 검색

### TS-V09 · `Find(What:="")` — 이전 검색어 재사용

**증상**: `Find(What:="")`로 빈 셀 검색 시 이전 Find의 검색어로 실행됨 (예상치 못한 셀 반환)

**원인**: Excel `Find`는 `What:=""`일 때 이전 Find 다이얼로그의 검색어 재사용

**해결**: `Find(What:="*")` (내용 있는 셀) + `SpecialCells(xlCellTypeBlanks)` (빈 셀) 하이브리드 사용

---

## 시트 보호

### TS-V10 · `SheetLock` — `Interior.ColorIndex`로 입력 셀 감지 누락

**증상**: RGB 색상 또는 테마 색상으로 칠한 셀이 입력 가능 셀로 인식되지 않고 잠김

**원인**: `Interior.ColorIndex`는 테마 색상과 일부 RGB 색상을 감지하지 못함

**해결**: `Interior.Pattern = xlNone` 기준 사용 (배경 없음 = 입력 가능 셀)

```vba
If cel.Interior.Pattern = xlNone Then
    cel.Locked = False
End If
```

---

## 초기화

### TS-V11 · `am_Core.IsReady` — VBE 리셋 후 `False` 유지로 기능 불작동

**증상**: VBE 리셋(Stop/재시작) 후 xlam 함수 호출 시 "초기화되지 않음" 오류

**원인**: `m_blnReady = False`가 리셋 후에도 유지되나 `Initialize()`가 자동 호출되지 않음

**해결**: lazy-init 패턴 — `IsReady` 호출 시 미초기화 상태면 `Initialize()` 자동 호출

```vba
Public Function IsReady() As Boolean
    If Not m_blnReady Then Initialize
    IsReady = m_blnReady
End Function
```

---

## 창/버튼 자동화 (Win32 API / UI Automation)

> 배경: EDU LMS 다운로드 매크로 — 웹 모달(Selenium) → 브라우저 확인 팝업(Chrome 자체 UI) →
> WMC SCDK 데스크톱 앱(WinForms) 저장 버튼 → OS 표준 "다른 이름으로 저장" 대화상자로 이어지는
> 팝업 연쇄 처리 자동화 과정에서 발생. 원본 모듈 `mod_WinAPI` → `corelib.xlam`에는 `am_Automation`으로 이식.

### TS-V12 · 좌표 기반 마우스 클릭 — 모니터 환경(해상도/DPI/멀티모니터)에 따라 어긋남

**증상**: `SetCursorPos`/`GetCursorPos` 기반 좌표 클릭이 싱글/멀티 모니터, 해상도, DPI 배율에 따라 어긋남

**원인**: DPI 미인식 프로세스(Excel)가 배율이 다른 모니터에서 좌표 스케일링과 어긋나고, 보조 모니터 배치에 따라 좌표가 음수가 될 수 있음

**해결**: 좌표 하드코딩을 지양하고, 대상 요소를 텍스트/이름으로 찾아 그 창의 `GetWindowRect`로 매번 동적 계산하는 방식(TS-V14)으로 전환. 근본적으로는 "요소 직접 탐색"이 가능하면 좌표 클릭보다 항상 우선

---

### TS-V13 · Chrome 외부 프로토콜 확인 팝업("○○을(를) 여시겠습니까?")에 `SendKeys` 미작동

**증상**: Chrome이 그리는 외부 프로토콜 실행 확인 팝업에 `SendKeys`를 보내도 반응 없음

**원인**: 이 팝업은 DOM도 OS 네이티브 창도 아닌 Chrome 자체 렌더링 UI라 Selenium이 제어 불가. 또한 `SendKeys`는 대상이 OS 포그라운드(활성) 창이어야 하는데, VBA 실행 중인 Excel이 계속 포커스를 쥐고 있어 실패

**해결**: `SendKeys "{Left}", True`(팝업 내 기본 선택 버튼에서 인접 버튼으로 포커스 이동) + `SendKeys "{ENTER}", True` 조합으로 성공

```vba
SendKeys "{Left}", True
SendKeys "{ENTER}", True
```

**주의**: KM Link(다중 PC 키보드/마우스 공유) 환경에서는 이 방식이 실패함 → TS-V16 참고

---

### TS-V14 · WinForms 커스텀 그리기 버튼 — `SendMessageA(BM_CLICK)` 무반응

**증상**: 별도 데스크톱 앱(WinForms) 창 내부의 커스텀 버튼을 `SendMessageA(..., BM_CLICK, ...)`로 클릭해도 반응 없음

**원인**: `EnumChildWindows`로 확인한 해당 컨트롤의 클래스명이 표준 `BUTTON`이 아닌 범용 `Window` 클래스 — 커스텀 그리기 버튼이라 `BM_CLICK` 메시지를 처리하지 않음

**해결**: `GetWindowRect`로 컨트롤의 실제 화면 좌표를 매번 동적으로 계산해 중심점을 클릭 (`ClickButtonByRect` 계열). 좌표를 하드코딩하지 않고 매번 재계산하므로 모니터 환경과 무관

**함정**: 여러 팝업/창이 연쇄되는 자동화에서는, "지금 이 단계가 웹페이지(DOM)인지 / 브라우저 자체 UI인지 / 별도 데스크톱 앱 창인지 / OS 표준 대화상자인지"를 먼저 구분하는 것이 핵심 진단 단계다. 이미 Selenium으로 처리 완료된 이전 단계(웹 모달)를 Win32 API로 다시 찾으려다 반복 실패한 사례 있음 — 실행 순서 재확인으로 해결. 표준 `BUTTON` 클래스 버튼은 `BM_CLICK`이 정상 작동함 (TS-V15 참고)

---

### TS-V15 · OS "다른 이름으로 저장" 대화상자 — `SendKeys "{Enter}"` 연속 입력 시 저장 안 됨

**증상**: 경로 입력 후 `{Enter}`를 두 번 보내면 저장되지 않고 창이 닫힘 (물리 키보드로는 정상 동작)

**원인**: 두 번째 Enter 시점에 포커스가 파일 이름 입력란이 아닌 다른 컨트롤(주소 표시줄 등)에 남아있어, Enter가 저장이 아닌 다른 동작(취소/닫기)으로 해석됨

**해결**: `ClickButtonByText("다른 이름으로 저장", "저장")`로 표준 Win32 버튼을 `SendMessageA(BM_CLICK)`으로 직접 클릭. 이 대화상자는 OS 표준 컨트롤이라 `BM_CLICK`이 정상 작동

**추가 확인**: 저장 완료 후 창 닫기(Alt+F4)는 실제로 남는 창이 2개(저장 확인창 + 원래 앱 본체)라 1회가 아닌 **2회 호출** 필요

---

### TS-V16 · KM Link(다중 PC 키보드/마우스 공유) 환경에서 `SendKeys` 자동화 실패 — UI Automation Late Binding도 424 오류

**증상**: KM Link로 입력 포커스를 다른 PC로 넘긴 상태에서, 자동화 PC 쪽 `SendKeys` 기반 팝업 클릭(TS-V13)이 작동하지 않음

**원인**: `SendKeys`는 OS 포그라운드 창 상태에 의존하는데, KM Link의 전역 키보드 후킹이 자동화 PC의 활성 입력 세션 인식이나 VBA 합성 키 입력과 간섭

**1차 시도(미채택)**: `AttachThreadInput` 기반 강제 포커스 확보(`ForceActivateWindow`) — 근본 해결 아님, 참고용으로만 보유

**해결**: UI Automation(OS 접근성 API, 스크린리더 등과 동일한 인터페이스)으로 전환하여 키보드/마우스/포커스와 완전히 무관하게 버튼을 직접 호출(Invoke)

- `CreateObject("UIAutomationClient.CUIAutomation")` (Late Binding) → **런타임 424 오류** (해당 환경에서 ProgID가 COM 레지스트리에 등록되어 있지 않음)
- `CreateObject("new:{CLSID}")` → **런타임 424 오류** (`"new:"` 모니커는 `GetObject` 전용 문법이라 `CreateObject`에는 잘못된 사용)
- VBE 참조에 **UIAutomationClient** 라이브러리를 추가하고 `New CUIAutomation`(Early Binding)으로 전환 → 성공
- `IUIAutomationElement.FindFirst`로 대상 창 이름·버튼 이름을 탐색, `IUIAutomationInvokePattern.Invoke`로 클릭 → KM Link 상태와 무관하게 정상 동작 확인

```vba
' 참조: VBE > 도구 > 참조 > UIAutomationClient 체크 필요
Dim oAutomation As New CUIAutomation
Dim oWindowElement As IUIAutomationElement
Set oWindowElement = oAutomation.GetRootElement.FindFirst( _
    TreeScope_Descendants, _
    oAutomation.CreatePropertyCondition(UIA_NamePropertyId, strWindowName))
```

**✅ corelib.xlam 이식 결론 (2026-07-24 실측 확인)**: `am_Automation.ClickButtonByUIA`를 처음엔 프로젝트의 "Late Binding, 참조 추가 없이 `CreateObject` 사용" 원칙에 따라 `CreateObject("UIAutomationClient.CUIAutomation")`로 이식했으나, 실제 개발 환경(Excel 2010 32bit)에서 실행 테스트 결과 **런타임 429 오류**("ActiveX 구성 요소는 개체를 만들 수 없습니다")로 재현됨 — 이 항목의 424와 같은 계열(ProgID/CLSID 해석 실패). 이에 따라 `ClickButtonByUIA`는 **Early Binding으로 확정**(`New CUIAutomation` + `IUIAutomationElement` 등, VBE 참조에 `UIAutomationClient` 라이브러리 추가 필요) — 프로젝트 Late Binding 원칙의 **명시적 예외**로 `CLAUDE.md`/`LOG.md`에 기록

---

### TS-V17 · 반복 실행 구간에서만 간헐적으로 버튼을 못 찾음

**증상**: 처리가 1회만 일어나는 흐름은 안정적으로 성공하는데, 같은 처리를 여러 번 반복하는 흐름(예: 옵션 개수만큼 반복)에서는 동일한 클릭 로직이 종종 실패함

**원인 추정**: 반복 횟수가 누적될수록 대상 앱 재실행 속도가 불안정(리소스 누적, 이전 창 종료 지연 등)해져 고정된 대기시간만으로는 렌더링 완료 시점을 항상 보장하지 못함

**해결**:
1. `ClickButtonByRect_Retry` 함수로 실패 시 대기 후 재시도(기본 3회)
2. 매크로 전체 대기시간을 균일하게 연장 + 팝업 직후 구간은 별도로 더 길게 고정
3. **핵심**: 재시도 간격(`ClickButtonByRect_Retry`의 `strRetryWait` 등 기본값)은 함수 내부(모듈)에 있으므로, **호출부(매크로)에서만 대기시간을 늘려서는 재시도 간격 자체는 늘어나지 않는다** — 재시도 간격을 조정하려면 모듈 쪽 기본값도 함께 수정해야 함

---

### 설계 원칙 (재사용 가능한 교훈)

1. 좌표 하드코딩(`ClickAtPosition` 류)은 최후의 수단으로만 사용하고, 가능하면 요소를 이름/텍스트로 찾아 직접 호출하는 방식을 우선한다
2. 표준 Win32 버튼은 `SendMessageA(BM_CLICK)`으로, 커스텀 그리기 버튼은 `GetWindowRect` 기반 동적 좌표 클릭으로 구분 대응한다
3. 팝업/창이 연쇄되는 자동화에서는 "지금 단계가 웹페이지(DOM) / 브라우저 자체 UI / 별도 데스크톱 앱 창 / OS 표준 대화상자 중 무엇인지"를 먼저 구분하는 것이 가장 중요한 진단 단계다 — 종류에 따라 제어 가능한 방법(Selenium / SendKeys / Win32 API / UI Automation)이 전혀 다르다
4. 입력 공유 환경(KM Link 등)에서는 `SendKeys`/마우스 이벤트 기반 자동화가 근본적으로 불안정하므로, UI Automation(접근성 API) 기반 `Invoke` 호출이 가장 안정적이다 — 단 `CreateObject` ProgID 방식이 환경에 따라 424로 실패할 수 있음(TS-V16)
5. 반복 실행 중 간헐적 실패는 대부분 "고정 대기시간 vs 실제 렌더링 시간의 불일치" 문제이므로, 대기시간 연장과 재시도 로직을 함께 적용한다. 이때 재시도 간격이 별도 함수/모듈 내부에 하드코딩되어 있는지 반드시 확인한다
