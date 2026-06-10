Attribute VB_Name = "am_Sheet"
Option Explicit

' ┌─────────────────────────────────────────────────────────┐
' │  am_Sheet                                               │
' │  역할 : 워크북/시트 조작 (백업, 표시/숨김, 정렬, 보호) │
' └─────────────────────────────────────────────────────────┘

' ══════════════════════════════════════════════════════════
'  워크북 / 시트 백업
' ══════════════════════════════════════════════════════════

' 목적   : 시트를 별도 파일로 백업
' 인수   : strPath - 백업 저장 폴더 경로
'          wb      - 호출한 CWB 워크북 개체
'          strFile - 저장 파일명 (기본: 자동 생성)
'          ws      - 백업할 시트 (기본: ActiveSheet)
' 예시   : BackupSheet("C:\Backup", ThisWorkbook)
Public Sub BackupSheet(ByVal strPath As String, _
                       ByVal wb As Workbook, _
                       Optional ByVal strFile As String = "", _
                       Optional ByVal ws As Worksheet = Nothing)

    If ws Is Nothing Then Set ws = wb.ActiveSheet

    ws.Copy

    strPath = IIf(Right(strPath, 1) <> "\", strPath & "\", strPath)
    Call prv_MkFolder(strPath)

    If strFile = "" Then
        strFile = "Bak_(" & Format(Now, "yyyymmdd_hhmmss") & ") " & _
                  Left(wb.Name, InStrRev(wb.Name, ".") - 1) & "_" & ws.Name & ".xlsx"
    End If

    Application.DisplayAlerts = False
    Select Case prv_GetExt(strFile)
        Case ".xlsx": ActiveWorkbook.SaveAs strPath & strFile, 51
        Case ".xlsm": ActiveWorkbook.SaveAs strPath & strFile, 52
        Case ".csv":  ActiveWorkbook.SaveAs strPath & strFile, 6
        Case ".txt":  ActiveWorkbook.SaveAs strPath & strFile, 4158
    End Select
    Application.DisplayAlerts = True

    ActiveWorkbook.Close False

End Sub

' 목적   : 전체 워크북을 복사본으로 백업
' 인수   : strPath - 백업 저장 폴더 경로
'          wb      - 호출한 CWB 워크북 개체
'          strFile - 저장 파일명 (기본: 자동 생성)
' 예시   : BackupWorkbook("C:\Backup", ThisWorkbook)
Public Sub BackupWorkbook(ByVal strPath As String, _
                          ByVal wb As Workbook, _
                          Optional ByVal strFile As String = "")

    strPath = IIf(Right(strPath, 1) <> "\", strPath & "\", strPath)
    If strFile = "" Then
        strFile = "Bak_(" & Format(Now, "yyyymmdd_hhmmss") & ") " & wb.Name
    End If

    wb.SaveCopyAs strPath & strFile

End Sub

' ══════════════════════════════════════════════════════════
'  시트 표시 / 숨김
' ══════════════════════════════════════════════════════════

' 목적   : 특정 시트만 표시하고 나머지 숨김
' 인수   : ws - 표시할 시트
'          wb - 대상 워크북
' 예시   : HideAllSheetsExceptOne(Sheet1, ThisWorkbook)
Public Sub HideAllSheetsExceptOne(ByVal ws As Worksheet, _
                                  ByVal wb As Workbook)

    On Error Resume Next

    ws.Visible = True
    ws.Activate

    With ActiveWindow
        .ScrollRow = 1
        .ScrollColumn = 1
    End With

    Dim objSht As Object
    For Each objSht In wb.Sheets
        If objSht.Name <> ws.Name Then
            objSht.Visible = xlSheetHidden
        End If
    Next objSht

    On Error GoTo 0

End Sub

' 목적   : 워크북 내 모든 시트 표시
' 인수   : wb - 대상 워크북
' 예시   : VisibleAllSheets(ThisWorkbook)
Public Sub VisibleAllSheets(ByVal wb As Workbook)

    On Error Resume Next

    Dim objSht As Object
    For Each objSht In wb.Sheets
        objSht.Visible = xlSheetVisible
    Next objSht

    On Error GoTo 0

End Sub

' ══════════════════════════════════════════════════════════
'  시트 정보 조회
' ══════════════════════════════════════════════════════════

' 목적   : 워크북 내 시트명 배열 반환
' 인수   : wb - 대상 워크북
' 반환   : Variant - 시트명 1차원 배열 (0-based)
' 예시   : arr = GetSheetNames(ThisWorkbook)
Public Function GetSheetNames(ByVal wb As Workbook) As Variant

    Dim arrResult() As Variant
    Dim objSht      As Object
    Dim i           As Long

    For Each objSht In wb.Sheets
        ReDim Preserve arrResult(i)
        arrResult(i) = objSht.Name
        i = i + 1
    Next objSht

    GetSheetNames = arrResult

End Function

' ══════════════════════════════════════════════════════════
'  시트 순서 정렬
' ══════════════════════════════════════════════════════════

' 목적   : 배열 순서대로 시트 탭 순서 재정렬
' 인수   : wb            - 대상 워크북
'          arrSheetNames - 정렬 순서 시트명 배열
' 예시   : SortSheets(ThisWorkbook, Array("Home", "입력", "DB"))
Public Sub SortSheets(ByVal wb As Workbook, _
                      ByVal arrSheetNames As Variant)

    Call VisibleAllSheets(wb)

    Dim vntName As Variant
    Dim objSht  As Object

    For Each vntName In arrSheetNames
        On Error Resume Next
        Set objSht = wb.Sheets(CStr(vntName))
        If Not objSht Is Nothing Then
            objSht.Move After:=wb.Sheets(wb.Sheets.Count)
        End If
        Set objSht = Nothing
        On Error GoTo 0
    Next vntName

End Sub

' ══════════════════════════════════════════════════════════
'  시트 보호
' ══════════════════════════════════════════════════════════

' 목적   : 시트 보호 설정
'          실제 사용 범위 전체 잠금 후, 배경색 없는 셀만 잠금 해제
'          → 배경색 없는 셀 = 입력 가능 셀 규약
' 인수   : ws    - 보호할 시트
'          strPW - 보호 비밀번호
' 예시   : SheetLock(ActiveSheet, "1234")
Public Sub SheetLock(ByVal ws    As Worksheet, _
                     ByVal strPW As String)

    Dim blnScreen  As Boolean
    Dim rngUsed    As Range
    Dim rngUnlock  As Range

    blnScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.Calculation    = xlCalculationManual
    Application.EnableEvents   = False

    On Error GoTo ErrHandler

    Set rngUsed = prv_GetUsedRange(ws)

    ws.Unprotect Password:=strPW

    If Not rngUsed Is Nothing Then
        rngUsed.Locked        = True
        rngUsed.FormulaHidden = True

        ' 배경색 없는 셀(xlNone) → 입력 가능 셀로 잠금 해제
        Set rngUnlock = prv_FindCellsByColor(xlNone, rngUsed, True)
        If Not rngUnlock Is Nothing Then
            rngUnlock.Locked        = False
            rngUnlock.FormulaHidden = False
        End If
    End If

    With ws
        .Protect Password:=strPW, Contents:=True, Scenarios:=True, _
                 AllowSorting:=True, AllowFiltering:=True
        .EnableSelection = xlNoRestrictions
    End With

    GoTo CleanUp

ErrHandler:
    MsgBox "오류 " & Err.Number & ": " & Err.Description, _
           vbCritical, "am_Sheet.SheetLock"

CleanUp:
    Application.FindFormat.Clear
    Application.Calculation    = xlCalculationAutomatic
    Application.EnableEvents   = True
    Application.ScreenUpdating = blnScreen

End Sub

' 목적   : 시트 보호 해제
' 인수   : ws    - 해제할 시트
'          strPW - 보호 비밀번호
' 예시   : SheetUnLock(ActiveSheet, "1234")
Public Sub SheetUnLock(ByVal ws    As Worksheet, _
                       ByVal strPW As String)

    ws.Unprotect Password:=strPW

End Sub

' ══════════════════════════════════════════════════════════
'  내부 전용 함수 (Private)
' ══════════════════════════════════════════════════════════

' 목적   : 폴더 생성 (내부 전용)
Private Sub prv_MkFolder(ByVal strPath As String)

    Dim arrPath() As String
    Dim strCur    As String
    Dim i         As Integer

    strPath = IIf(Right(strPath, 1) = "\", Left(strPath, Len(strPath) - 1), strPath)
    arrPath = Split(strPath, "\")

    For i = LBound(arrPath) To UBound(arrPath)
        strCur = strCur & arrPath(i) & "\"
        If Dir(strCur, vbDirectory) = "" Then MkDir strCur
    Next i

End Sub

' 목적   : 파일 확장자 반환 (내부 전용)
Private Function prv_GetExt(ByVal strFileName As String) As String
    prv_GetExt = Mid(strFileName, InStrRev(strFileName, "."))
End Function

' 목적   : 실제 사용 범위 계산 (내부 전용 — SheetLock 전용)
' 참고   : am_Range.GetUsedRange 와 동일 로직, 모듈 독립성 원칙에 따라 내부 구현
Private Function prv_GetUsedRange(ByVal ws As Worksheet) As Range

    Dim rng         As Range
    Dim rngFormulas As Range
    Dim tbl         As ListObject
    Dim shp         As Shape
    Dim bFirst      As Boolean
    Dim i           As Long
    Dim lngMaxRow   As Long
    Dim lngMaxCol   As Long

    bFirst = True

    On Error Resume Next

    Set rng = ws.Cells.SpecialCells(xlCellTypeConstants)
    bFirst  = (rng Is Nothing)

    Set rngFormulas = ws.Cells.SpecialCells(xlCellTypeFormulas)
    If Not rngFormulas Is Nothing Then
        If bFirst Then
            Set rng = rngFormulas
            bFirst  = False
        Else
            Set rng = Union(rng, rngFormulas)
        End If
    End If

    For Each tbl In ws.ListObjects
        If bFirst Then
            Set rng = tbl.Range
            bFirst  = False
        Else
            Set rng = Union(rng, tbl.Range)
        End If
    Next tbl

    For Each shp In ws.Shapes
        Dim rngShape As Range
        Set rngShape = ws.Range(shp.TopLeftCell, shp.BottomRightCell)
        If bFirst Then
            Set rng = rngShape
            bFirst  = False
        Else
            Set rng = Union(rng, rngShape)
        End If
    Next shp

    On Error GoTo 0

    If rng Is Nothing Then Exit Function

    lngMaxRow = 0
    lngMaxCol = 0

    For i = 1 To rng.Areas.Count
        With rng.Areas(i)
            If .Row    + .Rows.Count    - 1 > lngMaxRow Then lngMaxRow = .Row    + .Rows.Count    - 1
            If .Column + .Columns.Count - 1 > lngMaxCol Then lngMaxCol = .Column + .Columns.Count - 1
        End With
    Next i

    Set prv_GetUsedRange = ws.Range(ws.Cells(1, 1), ws.Cells(lngMaxRow, lngMaxCol))

End Function

' 목적   : 배경색 기준 셀 검색 (내부 전용 — SheetLock 전용)
' 참고   : am_Range.FindCellsByColor 와 동일 로직, 모듈 독립성 원칙에 따라 내부 구현
Private Function prv_FindCellsByColor(ByVal lngColor      As Long, _
                                      ByVal rng            As Range, _
                                      ByVal blnColorIndex  As Boolean) As Range

    Dim f         As Range
    Dim strFirst  As String
    Dim rngResult As Range

    Application.FindFormat.Clear

    If blnColorIndex Then
        Application.FindFormat.Interior.ColorIndex = lngColor
    Else
        Application.FindFormat.Interior.Color = lngColor
    End If

    On Error Resume Next

    With rng
        Set f = .Find(What:="", LookIn:=xlFormulas, _
                      LookAt:=xlPart, SearchOrder:=xlByRows, _
                      SearchDirection:=xlNext, SearchFormat:=True)

        If Not f Is Nothing Then
            strFirst = f.Address
            Set rngResult = f

            Do
                Set f = .Find(What:="", After:=f, LookIn:=xlFormulas, _
                              LookAt:=xlPart, SearchOrder:=xlByRows, _
                              SearchDirection:=xlNext, SearchFormat:=True)
                If f Is Nothing     Then Exit Do
                If f.Address = strFirst Then Exit Do
                Set rngResult = Union(rngResult, f)
            Loop
        End If
    End With

    On Error GoTo 0

    If Not rngResult Is Nothing Then Set prv_FindCellsByColor = rngResult

End Function
