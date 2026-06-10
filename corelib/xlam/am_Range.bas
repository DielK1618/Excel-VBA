Attribute VB_Name = "am_Range"
Option Explicit

' ┌─────────────────────────────────────────────────────────┐
' │  am_Range                                               │
' │  역할 : 범위 검색, 범위 조작                            │
' └─────────────────────────────────────────────────────────┘

' ══════════════════════════════════════════════════════════
'  범위 검색
' ══════════════════════════════════════════════════════════

' 목적   : 범위 내 값 검색 후 셀 반환
' 인수   : rngFind   - 검색 범위
'          strValue  - 검색 값
'          blnAddRow - True: 미발견 시 마지막 행 다음 셀 반환
'          LookAt    - 검색 방식 (기본: 전체 일치)
' 반환   : Range - 찾은 셀 / 미발견 시 Nothing 또는 다음 행 셀
' 예시   : Set f = FindRange(tbl.ListColumns("ID").DataBodyRange, "001")
Public Function FindRange(ByVal rngFind   As Range, _
                          ByVal strValue  As String, _
                          Optional ByVal blnAddRow As Boolean  = True, _
                          Optional ByVal LookAt    As XlLookAt = xlWhole) As Range

    On Error Resume Next

    Dim f As Range
    Set f = rngFind.Find(strValue, , xlValues, LookAt)

    If Not f Is Nothing Then
        Set FindRange = f
    Else
        If blnAddRow Then
            Set f = rngFind.Worksheet.Cells(Rows.Count, rngFind.Column).End(xlUp)
            Set FindRange = IIf(f.Value = "", f, f.Offset(1))
        Else
            Set FindRange = Nothing
        End If
    End If

    On Error GoTo 0

End Function

' 목적   : 배경색 기준으로 범위 내 셀 검색 후 합집합 반환
' 인수   : lngColor      - 검색할 색상값
'          rng           - 검색 범위 (기본: ActiveSheet 전체)
'          blnColorIndex - True: ColorIndex 기준 / False: Color(RGB) 기준
' 반환   : Range - 해당 색상 셀 전체 (없으면 Nothing)
' 예시   : Set rng = FindCellsByColor(xlNone, ws.UsedRange, True)
Public Function FindCellsByColor(ByVal lngColor      As Long, _
                                 Optional ByVal rng           As Range, _
                                 Optional ByVal blnColorIndex As Boolean = True) As Range

    Dim f          As Range
    Dim strFirst   As String
    Dim rngResult  As Range
    Dim ws         As Worksheet

    If rng Is Nothing Then
        Set ws  = ActiveSheet
        Set rng = ws.Cells
    Else
        Set ws  = rng.Parent
    End If

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

                If f Is Nothing Then Exit Do
                If f.Address = strFirst Then Exit Do

                Set rngResult = Union(rngResult, f)
            Loop
        End If
    End With

    On Error GoTo 0

    Application.FindFormat.Clear

    If Not rngResult Is Nothing Then Set FindCellsByColor = rngResult

End Function

' ══════════════════════════════════════════════════════════
'  범위 조작
' ══════════════════════════════════════════════════════════

' 목적   : 시트의 실제 사용 범위 계산 (상수·수식·표·도형 포함)
' 인수   : ws         - 대상 시트 (기본: ActiveSheet)
'          blnFromA1  - True: A1 기준 반환 / False: 실제 시작 셀 기준
' 반환   : Range - 사용 범위 / 내용 없으면 Nothing
' 예시   : Set rng = GetUsedRange(ActiveSheet)
'          Set rng = GetUsedRange(ws, False)  ' 실제 데이터 영역만
Public Function GetUsedRange(Optional ByVal ws        As Worksheet, _
                              Optional ByVal blnFromA1 As Boolean = True) As Range

    If ws Is Nothing Then Set ws = ActiveSheet

    Dim rng         As Range
    Dim rngFormulas As Range
    Dim tbl         As ListObject
    Dim shp         As Shape
    Dim rngShape    As Range
    Dim bFirst      As Boolean
    Dim i           As Long
    Dim lngStartRow As Long
    Dim lngStartCol As Long
    Dim lngMaxRow   As Long
    Dim lngMaxCol   As Long

    bFirst = True

    On Error Resume Next

    ' 1. 상수 셀
    Set rng = ws.Cells.SpecialCells(xlCellTypeConstants)
    bFirst = (rng Is Nothing)

    ' 2. 수식 셀
    Set rngFormulas = ws.Cells.SpecialCells(xlCellTypeFormulas)
    If Not rngFormulas Is Nothing Then
        If bFirst Then
            Set rng  = rngFormulas
            bFirst   = False
        Else
            Set rng = Union(rng, rngFormulas)
        End If
    End If

    ' 3. 표(ListObject) 범위
    For Each tbl In ws.ListObjects
        If bFirst Then
            Set rng = tbl.Range
            bFirst  = False
        Else
            Set rng = Union(rng, tbl.Range)
        End If
    Next tbl

    ' 4. 도형 점유 범위
    For Each shp In ws.Shapes
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

    ' 5. 바운딩 박스 계산
    If blnFromA1 Then
        lngStartRow = 1
        lngStartCol = 1
    Else
        lngStartRow = rng.Row
        lngStartCol = rng.Column
    End If

    lngMaxRow = 0
    lngMaxCol = 0

    For i = 1 To rng.Areas.Count
        With rng.Areas(i)
            If Not blnFromA1 Then
                If .Row    < lngStartRow Then lngStartRow = .Row
                If .Column < lngStartCol Then lngStartCol = .Column
            End If
            If .Row    + .Rows.Count    - 1 > lngMaxRow Then lngMaxRow = .Row    + .Rows.Count    - 1
            If .Column + .Columns.Count - 1 > lngMaxCol Then lngMaxCol = .Column + .Columns.Count - 1
        End With
    Next i

    Set GetUsedRange = ws.Range(ws.Cells(lngStartRow, lngStartCol), _
                                ws.Cells(lngMaxRow,   lngMaxCol))

End Function

' ══════════════════════════════════════════════════════════
'  Application.Run 호환 스칼라 래퍼
'  (Range 반환 함수는 Application.Run 경유 수신 불가)
' ══════════════════════════════════════════════════════════

' 목적   : GetUsedRange 결과 존재 여부 반환
' 반환   : Boolean - True: 범위 있음 / False: Nothing
Public Function GetUsedRange_IsValid(Optional ByVal ws As Worksheet) As Boolean
    GetUsedRange_IsValid = Not (GetUsedRange(ws) Is Nothing)
End Function

' 목적   : GetUsedRange 결과의 행 수 반환 (0=Nothing)
Public Function GetUsedRange_RowCount(Optional ByVal ws As Worksheet) As Long
    Dim rng As Range
    Set rng = GetUsedRange(ws)
    If Not rng Is Nothing Then GetUsedRange_RowCount = rng.Rows.Count
End Function

' 목적   : GetUsedRange 결과의 열 수 반환 (0=Nothing)
Public Function GetUsedRange_ColCount(Optional ByVal ws As Worksheet) As Long
    Dim rng As Range
    Set rng = GetUsedRange(ws)
    If Not rng Is Nothing Then GetUsedRange_ColCount = rng.Columns.Count
End Function

' 목적   : FindRange 결과 존재 여부 반환
' 반환   : Boolean - True: 찾음 / False: Nothing
Public Function FindRange_IsValid(ByVal rngFind As Range, _
                                  ByVal strValue As String, _
                                  Optional ByVal blnAddRow As Boolean = False) As Boolean
    FindRange_IsValid = Not (FindRange(rngFind, strValue, blnAddRow) Is Nothing)
End Function

' 목적   : FindRange 결과 셀 값 반환 ("" = Nothing)
Public Function FindRange_CellValue(ByVal rngFind As Range, _
                                    ByVal strValue As String, _
                                    Optional ByVal blnAddRow As Boolean = False) As String
    Dim f As Range
    Set f = FindRange(rngFind, strValue, blnAddRow)
    If Not f Is Nothing Then FindRange_CellValue = CStr(f.Value)
End Function

' 목적   : FindCellsByColor 결과 셀 수 반환 (0=Nothing)
Public Function FindCellsByColor_Count(ByVal lngColor As Long, _
                                       ByVal rng As Range, _
                                       Optional ByVal blnUnion As Boolean = True) As Long
    Dim r As Range
    Set r = FindCellsByColor(lngColor, rng, blnUnion)
    If Not r Is Nothing Then FindCellsByColor_Count = r.Cells.Count
End Function
