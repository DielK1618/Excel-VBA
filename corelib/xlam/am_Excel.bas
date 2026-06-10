Attribute VB_Name = "am_Excel"
Option Explicit

' +---------------------------------------------------------+
' |  am_Excel                                               |
' |  역할 : 인쇄/내보내기, 차트, 도형, 키보드/마우스 자동화  |
' +---------------------------------------------------------+

' -- Windows API --
Private Declare PtrSafe Sub keybd_event Lib "user32" _
        (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare PtrSafe Sub mouse_event Lib "user32" _
        (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const KEYEVENTF_KEYDOWN     As Long = 0
Private Const KEYEVENTF_KEYUP       As Long = 2
Private Const MOUSEEVENTF_LEFTDOWN  As Long = &H2
Private Const MOUSEEVENTF_LEFTUP    As Long = &H4
Private Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Private Const MOUSEEVENTF_RIGHTUP   As Long = &H10

' 가상 키코드
Public Enum VirtualKeys
    VK_TAB     = &H9
    VK_RETURN  = &HD
    VK_CONTROL = &H11
    VK_MENU    = &H12
    VK_ESCAPE  = &H1B
    VK_SPACE   = &H20
    VK_LEFT    = &H25
    VK_UP      = &H26
    VK_RIGHT   = &H27
    VK_DOWN    = &H28
End Enum

' 키 동작 정의
Public Enum KeyActions
    ACTION_COPY        = 1
    ACTION_PASTE       = 2
    ACTION_TAB         = 3
    ACTION_ENTER       = 4
    ACTION_ALT_TAB     = 5
    ACTION_ESCAPE      = 6
    ACTION_ARROW_DOWN  = 7
    ACTION_ARROW_UP    = 8
    ACTION_ARROW_LEFT  = 9
    ACTION_ARROW_RIGHT = 10
End Enum

' ==========================================================
'  인쇄 / 내보내기
' ==========================================================

' 목적   : 워크시트 인쇄 설정
' 인수   : sht            - 대상 시트 (생략 시 ActiveSheet)
'          rngPrintArea   - 인쇄 영역 (생략 시 첫 번째 표 또는 UsedRange)
'          rngTitleRows   - 상단 반복 인쇄 행 (생략 시 설정 안 함)
'          blnAddArea     - True=기존 인쇄 영역에 추가, False=덮어씌움 (기본 False)
'          paperSize      - 용지 크기 (기본 xlPaperA4)
'          orientation    - 용지 방향 (기본 xlPortrait)
'          sngTopMargin   - 위쪽 여백 cm (기본 1.5)
'          sngBottomMargin - 아래쪽 여백 cm (기본 1.0)
'          sngLeftMargin  - 왼쪽 여백 cm (기본 1.0)
'          sngRightMargin - 오른쪽 여백 cm (기본 1.0)
'          CenterH        - 가로 가운데 (기본 True)
'          CenterV        - 세로 가운데 (기본 False)
'          fitToPage      - True=페이지 맞춤, False=배율% (기본 True)
'          intWide        - [fitToPage=True] 가로 페이지 수 (기본 1, 0=자동)
'          intTall        - [fitToPage=True] 세로 페이지 수 (기본 0=자동)
'          intPer         - [fitToPage=False] 인쇄 배율 % (기본 100)
Public Sub SetPrintPage(Optional ByVal sht As Worksheet, _
                        Optional ByVal rngPrintArea As Range, _
                        Optional ByVal rngTitleRows As Range, _
                        Optional ByVal blnAddArea As Boolean = False, _
                        Optional ByVal paperSize As XlPaperSize = xlPaperA4, _
                        Optional ByVal orientation As XlPageOrientation = xlPortrait, _
                        Optional ByVal sngTopMargin As Double = 1.5, _
                        Optional ByVal sngBottomMargin As Double = 1, _
                        Optional ByVal sngLeftMargin As Double = 1, _
                        Optional ByVal sngRightMargin As Double = 1, _
                        Optional ByVal CenterH As Boolean = True, _
                        Optional ByVal CenterV As Boolean = False, _
                        Optional ByVal fitToPage As Boolean = True, _
                        Optional ByVal intWide As Integer = 1, _
                        Optional ByVal intTall As Integer = 0, _
                        Optional ByVal intPer As Integer = 100)

    On Error GoTo ErrHandler

    If sht Is Nothing Then Set sht = ActiveSheet

    If rngPrintArea Is Nothing Then
        On Error Resume Next
        Set rngPrintArea = sht.ListObjects(1).Range
        On Error GoTo ErrHandler
        If rngPrintArea Is Nothing Then Set rngPrintArea = sht.UsedRange
    End If

    With sht.PageSetup

        If blnAddArea And .PrintArea <> "" Then
            .PrintArea = .PrintArea & "," & rngPrintArea.Address
        Else
            .PrintArea = rngPrintArea.Address
        End If

        If Not rngTitleRows Is Nothing Then .PrintTitleRows = rngTitleRows.Address

        .paperSize          = paperSize
        .orientation        = orientation
        .CenterHorizontally = CenterH
        .CenterVertically   = CenterV
        .TopMargin          = Application.CentimetersToPoints(sngTopMargin)
        .LeftMargin         = Application.CentimetersToPoints(sngLeftMargin)
        .RightMargin        = Application.CentimetersToPoints(sngRightMargin)
        .BottomMargin       = Application.CentimetersToPoints(sngBottomMargin)

        If fitToPage Then
            .Zoom = False
            .FitToPagesWide = IIf(intWide >= 1, intWide, False)
            .FitToPagesTall = IIf(intTall >= 1, intTall, False)
        Else
            .Zoom = intPer
        End If

    End With

    Exit Sub

ErrHandler:
    MsgBox "오류 " & Err.Number & ": " & Err.Description, vbCritical, am_Core.AM_NAME
End Sub

' 목적   : 워크시트를 PDF 파일로 내보내기
' 인수   : strFilePath  - 저장할 PDF 전체 경로 (.pdf 없으면 자동 추가)
'          sht          - PDF로 저장할 워크시트 (생략 시 ActiveSheet)
'          blnOpenAfter - True=저장 후 자동 열기 (기본 False)
'          xlQual       - PDF 품질 (기본 xlQualityStandard)
'          blnDocProps  - 문서 속성 포함 여부 (기본 True)
Public Sub ExportPDF(ByVal strFilePath As String, _
                     Optional ByVal sht As Worksheet, _
                     Optional ByVal blnOpenAfter As Boolean = False, _
                     Optional ByVal xlQual As XlFixedFormatQuality = xlQualityStandard, _
                     Optional ByVal blnDocProps As Boolean = True)

    On Error GoTo ErrHandler

    Dim strFolder As String

    If sht Is Nothing Then Set sht = ActiveSheet

    If LCase(Right(strFilePath, 4)) <> ".pdf" Then strFilePath = strFilePath & ".pdf"

    strFolder = Left(strFilePath, InStrRev(strFilePath, "\"))
    If strFolder <> "" And Dir(strFolder, vbDirectory) = "" Then
        prv_MkFolder strFolder
        If Dir(strFolder, vbDirectory) = "" Then
            MsgBox "폴더를 생성할 수 없습니다." & vbCrLf & strFolder, vbCritical, am_Core.AM_NAME
            Exit Sub
        End If
    End If

    sht.ExportAsFixedFormat Type:=xlTypePDF, _
                            Filename:=strFilePath, _
                            Quality:=xlQual, _
                            IncludeDocProperties:=blnDocProps, _
                            IgnorePrintAreas:=False, _
                            OpenAfterPublish:=blnOpenAfter
    Exit Sub

ErrHandler:
    MsgBox "PDF 저장 중 오류가 발생했습니다." & vbCrLf & _
           "경로: " & strFilePath & vbCrLf & vbCrLf & _
           "오류: " & Err.Description, vbCritical, am_Core.AM_NAME
End Sub

' 목적   : 워크시트를 CSV 파일로 내보내기 (xlCSV, 시스템 로케일 인코딩)
' 인수   : sht         - CSV로 저장할 워크시트
'          strPath     - 저장 폴더 경로 (없으면 자동 생성)
'          strFileName - 파일명 (.csv 확장자 있어도 자동 처리)
Public Sub ExportSheetToCSV(ByVal sht As Worksheet, _
                            ByVal strPath As String, _
                            ByVal strFileName As String)

    On Error GoTo ErrHandler

    Dim blnVisible As Boolean
    Dim wbCopy     As Workbook

    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    If LCase(Right(strFileName, 4)) = ".csv" Then
        strFileName = Left(strFileName, Len(strFileName) - 4)
    End If

    prv_MkFolder strPath

    blnVisible = sht.Visible
    If Not blnVisible Then sht.Visible = True

    sht.Copy
    Set wbCopy = ActiveWorkbook
    Application.DisplayAlerts = False
    wbCopy.SaveAs Filename:=strPath & strFileName & ".csv", FileFormat:=xlCSV
    Application.DisplayAlerts = True
    wbCopy.Close SaveChanges:=False

    GoTo CleanUp

ErrHandler:
    MsgBox "CSV 저장 중 오류가 발생했습니다." & vbCrLf & _
           "경로: " & strPath & strFileName & ".csv" & vbCrLf & vbCrLf & _
           "오류: " & Err.Description, vbCritical, am_Core.AM_NAME
    On Error Resume Next
    If Not wbCopy Is Nothing Then wbCopy.Close SaveChanges:=False
    On Error GoTo ErrHandler

CleanUp:
    sht.Visible = blnVisible
End Sub

' 목적   : 폴더 생성 (중간 경로 없을 경우 재귀 생성)
Private Sub prv_MkFolder(ByVal strPath As String)
    Dim objFSO    As Object
    Dim strParent As String

    On Error GoTo ErrHandler

    Set objFSO = CreateObject("Scripting.FileSystemObject")

    If Not objFSO.FolderExists(strPath) Then
        strParent = objFSO.GetParentFolderName(strPath)
        If strParent <> "" And Not objFSO.FolderExists(strParent) Then
            prv_MkFolder strParent
        End If
        If Not objFSO.FolderExists(strPath) Then objFSO.CreateFolder strPath
    End If

    GoTo CleanUp

ErrHandler:
    ' 생성 실패 시 무시 — 호출자에서 Dir()로 결과 판단

CleanUp:
    On Error Resume Next
    Set objFSO = Nothing
End Sub

' ==========================================================
'  차트
' ==========================================================

' 목적   : 차트 이름으로 데이터 범위 설정
' 인수   : strChart - ChartObjects 이름
'          rng      - 설정할 데이터 범위
'          ws       - 대상 워크시트 (생략 시 ActiveSheet)
Public Sub SetChartDataRange(ByVal strChart As String, _
                             ByVal rng As Range, _
                             Optional ByVal ws As Worksheet)
    If ws Is Nothing Then Set ws = ActiveSheet

    On Error Resume Next
    With ws.ChartObjects(strChart).Chart
        .SetSourceData Source:=rng
    End With
    On Error GoTo 0
End Sub

' ==========================================================
'  도형
' ==========================================================

' 목적   : 도형 텍스트와 일치하는 이름의 도형 OnAction 매크로 실행
' 인수   : strShpName - 찾을 도형 텍스트 (대소문자 무시)
'          ws         - 대상 워크시트 (생략 시 ActiveSheet)
Public Sub RunShpMacro(ByVal strShpName As String, _
                       Optional ByVal ws As Worksheet)
    If ws Is Nothing Then Set ws = ActiveSheet

    Dim shp      As Shape
    Dim strText  As String
    Dim strMacro As String

    For Each shp In ws.Shapes
        strText = GetShapeTextSafe(shp)
        If Len(Trim(strText)) > 0 Then
            If UCase(Trim(strText)) = UCase(strShpName) Then
                On Error Resume Next
                strMacro = shp.OnAction
                On Error GoTo 0
                If Len(strMacro) > 0 Then Application.Run strMacro
            End If
        End If
    Next shp
End Sub

' 목적   : 도형(또는 그룹 도형)의 텍스트를 안전하게 반환
' 인수   : shp - 대상 Shape
' 반환   : String - 도형 텍스트 (실패 시 "" 반환)
Public Function GetShapeTextSafe(ByVal shp As Shape) As String
    Dim s      As Shape
    Dim strOut As String

    On Error GoTo ErrHandler

    If shp.Type = msoGroup Then
        For Each s In shp.GroupItems
            strOut = strOut & " " & prv_GetShapeTextSafe_GItem(s)
        Next s
    Else
        strOut = prv_GetShapeTextSafe_GItem(shp)
    End If

    GetShapeTextSafe = Trim(strOut)
    Exit Function

ErrHandler:
    GetShapeTextSafe = Trim(strOut)
End Function

Private Function prv_GetShapeTextSafe_GItem(ByVal s As Shape) As String
    Dim strRes As String

    On Error Resume Next

    If Not s.TextFrame2 Is Nothing Then
        If s.TextFrame2.HasText = msoTrue Then
            strRes = s.TextFrame2.TextRange.Text
            GoTo Done
        End If
    End If

    If Not s.TextFrame Is Nothing Then
        If s.TextFrame.HasText = msoTrue Then
            strRes = s.TextFrame.Characters.Text
            GoTo Done
        End If
    End If

Done:
    On Error GoTo 0
    prv_GetShapeTextSafe_GItem = Trim(strRes)
End Function

' ==========================================================
'  키보드
' ==========================================================

' 목적   : KeyActions Enum에 정의된 키 동작 실행
' 인수   : action    - 실행할 키 동작 (KeyActions Enum 또는 Long)
'          waitAfter - 실행 후 대기 ms (기본 100)
Public Sub ExecuteKeyAction(ByVal action As KeyActions, _
                            Optional ByVal waitAfter As Long = 100)
    Select Case action
        Case ACTION_COPY        : prv_SendKeyCombo VK_CONTROL, 67   ' Ctrl+C
        Case ACTION_PASTE       : prv_SendKeyCombo VK_CONTROL, 86   ' Ctrl+V
        Case ACTION_TAB         : prv_SendKey VK_TAB
        Case ACTION_ENTER       : prv_SendKey VK_RETURN
        Case ACTION_ALT_TAB     : prv_SendKeyCombo VK_MENU, VK_TAB
        Case ACTION_ESCAPE      : prv_SendKey VK_ESCAPE
        Case ACTION_ARROW_DOWN  : prv_SendKey VK_DOWN
        Case ACTION_ARROW_UP    : prv_SendKey VK_UP
        Case ACTION_ARROW_LEFT  : prv_SendKey VK_LEFT
        Case ACTION_ARROW_RIGHT : prv_SendKey VK_RIGHT
    End Select

    If waitAfter > 0 Then prv_WaitMs waitAfter
End Sub

' 목적   : 여러 키 동작을 순서대로 실행
' 인수   : actions - KeyActions 또는 Array(KeyActions, waitMs) 의 ParamArray
Public Sub ExecuteKeySequence(ParamArray actions() As Variant)
    Dim i As Long

    For i = LBound(actions) To UBound(actions)
        If IsArray(actions(i)) Then
            ExecuteKeyAction actions(i)(0), actions(i)(1)
        Else
            ExecuteKeyAction actions(i)
        End If
    Next i
End Sub

' 목적   : 범위의 각 셀에 키 시퀀스 적용
' 인수   : rng     - 대상 범위
'          actions - ExecuteKeySequence에 전달할 동작 ParamArray
Public Sub ProcessRangeWithKeySequence(ByVal rng As Range, _
                                       ParamArray actions() As Variant)
    Dim cel As Range

    For Each cel In rng
        If Not IsEmpty(cel) Then
            cel.Select
            ExecuteKeySequence actions
        End If
    Next cel
End Sub

Private Sub prv_SendKey(ByVal bKeyCode As Byte, _
                        Optional ByVal blnRelease As Boolean = True)
    keybd_event bKeyCode, 0, KEYEVENTF_KEYDOWN, 0
    If blnRelease Then keybd_event bKeyCode, 0, KEYEVENTF_KEYUP, 0
End Sub

Private Sub prv_SendKeyCombo(ParamArray keyCodes() As Variant)
    Dim i As Long

    For i = LBound(keyCodes) To UBound(keyCodes)
        keybd_event CByte(keyCodes(i)), 0, KEYEVENTF_KEYDOWN, 0
    Next i

    prv_WaitMs 50

    For i = UBound(keyCodes) To LBound(keyCodes) Step -1
        keybd_event CByte(keyCodes(i)), 0, KEYEVENTF_KEYUP, 0
    Next i
End Sub

' 목적   : ms 단위 대기 (키보드/마우스 내부 공용)
Private Sub prv_WaitMs(ByVal lngMs As Long)
    Application.Wait Now + TimeSerial(0, 0, lngMs / 1000)
End Sub

' ==========================================================
'  마우스
' ==========================================================

' 목적   : 현재 마우스 커서 위치 반환
' 인수   : lngX - (반환) X 좌표
'          lngY - (반환) Y 좌표
Public Sub GetMousePosition(ByRef lngX As Long, ByRef lngY As Long)
    Dim pt As POINTAPI
    GetCursorPos pt
    lngX = pt.x
    lngY = pt.y
End Sub

' 목적   : 마우스를 지정 좌표로 이동 후 클릭
' 인수   : lngX     - X 좌표
'          lngY     - Y 좌표
'          blnLeft  - True=왼쪽 클릭, False=오른쪽 클릭 (기본 True)
'          strDelay - 이동 후 클릭 전 대기 시간 "HH:MM:SS" (생략 시 생략)
'          strWait  - 클릭 후 대기 시간 "HH:MM:SS" (생략 시 생략)
Public Sub ClickAtPosition(ByVal lngX As Long, _
                           ByVal lngY As Long, _
                           Optional ByVal blnLeft As Boolean = True, _
                           Optional ByVal strDelay As String = "", _
                           Optional ByVal strWait As String = "")
    SetCursorPos lngX, lngY

    If strDelay <> "" Then WaitTime strDelay

    If blnLeft Then
        mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
        mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    Else
        mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
        mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
    End If

    If strWait <> "" Then WaitTime strWait
End Sub

' 목적   : 입력된 시간만큼 대기 (Application.Wait 기반)
' 인수   : strTime - 대기 시간 "HH:MM:SS" 형식 (예: "00:00:03" = 3초)
Public Sub WaitTime(ByVal strTime As String)
    Application.Wait Now + TimeValue(strTime)
End Sub
