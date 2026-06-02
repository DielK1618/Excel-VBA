VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "현재_통합_문서"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ── 1. CWB 열림 이벤트 ──────────────────────────────────────
Private Sub Workbook_Open()

    On Error GoTo ErrHandler

    Dim strXlamPath As String
    Dim wbXlam      As Workbook

    strXlamPath = ThisWorkbook.Path & "\corelib.xlam"

    ' ── 1-1. 파일 존재 확인 ──────────────────────────────────
    If Dir(strXlamPath) = "" Then
        MsgBox "corelib.xlam 파일을 찾을 수 없습니다." & vbCrLf & _
               strXlamPath, vbCritical, "corelib"
        Exit Sub
    End If

    ' ── 1-2. 로드 여부 확인 (오류 방식) ─────────────────────
    ' 참고 : Workbooks 컬렉션에 xlam 이 포함되지 않으므로
    '        오류 방식으로 체크
    On Error Resume Next
    Set wbXlam = Workbooks("corelib.xlam")
    On Error GoTo ErrHandler

    ' ── 1-3. 미로드 시 열기 ──────────────────────────────────
    If wbXlam Is Nothing Then
        Workbooks.Open strXlamPath
    End If

    Exit Sub

ErrHandler:
    MsgBox "corelib.xlam 로드 중 오류가 발생했습니다." & vbCrLf & _
           "오류 " & Err.Number & ": " & Err.Description, _
           vbCritical, "corelib"

End Sub

' ── 2. CWB 닫힘 이벤트 ──────────────────────────────────────
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' 필요 시 정리 코드 삽입 위치
End Sub
