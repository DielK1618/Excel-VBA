Attribute VB_Name = "tpl_Range"
Option Explicit
Function GetTwbRange(ByVal strRange As String) As Range

    Dim tWb As Workbook
    Dim aWb As Workbook
    
    Set tWb = ThisWorkbook
    Set aWb = ActiveWorkbook
    
    If tWb.Name = aWb.Name Then
        Set GetTwbRange = Range(strRange)
    Else
        tWb.Activate
        Set GetTwbRange = Range(strRange)
        aWb.Activate
    End If
    
End Function
Function GetUsedRange(Optional ByVal ws As Worksheet, _
                      Optional ByVal blnFromA1 As Boolean = True) As Range
    If ws Is Nothing Then Set ws = ActiveSheet
    
    Dim rng As Range
    Dim tbl As ListObject
    Dim shp As Shape
    Dim bFirstRange As Boolean
    
    bFirstRange = True
    
    On Error Resume Next
    
    ' 1. 상수 추가
    Set rng = ws.Cells.SpecialCells(xlCellTypeConstants)
    bFirstRange = (rng Is Nothing)
    
    ' 2. 수식 합치기 (첫 범위 처리 개선)
    Dim rngFormulas As Range
    Set rngFormulas = ws.Cells.SpecialCells(xlCellTypeFormulas)
    
    If Not rngFormulas Is Nothing Then
        If bFirstRange Then
            Set rng = rngFormulas
            bFirstRange = False
        Else
            Set rng = Union(rng, rngFormulas)
        End If
    End If
    
    ' 3. 표 합치기
    For Each tbl In ws.ListObjects
        If bFirstRange Then
            Set rng = tbl.Range
            bFirstRange = False
        Else
            Set rng = Union(rng, tbl.Range)
        End If
    Next tbl
    
    ' 4. 도형 합치기
    For Each shp In ws.Shapes
        Dim rngShape As Range
        Set rngShape = ws.Range(shp.TopLeftCell, shp.BottomRightCell)
        If bFirstRange Then
            Set rng = rngShape
            bFirstRange = False
        Else
            Set rng = Union(rng, rngShape)
        End If
    Next shp
    
    On Error GoTo 0
    
    ' 5. 범위 반환
    If Not rng Is Nothing Then
        Dim i As Long
        Dim lngStartRow As Long, lngStartCol As Long
        Dim lngMaxRow As Long, lngMaxCol As Long
        
        If blnFromA1 Then
            lngStartRow = 1
            lngStartCol = 1
        Else
            lngStartRow = rng.Row
            lngStartCol = rng.Column
        End If
        
        lngMaxRow = 0
        lngMaxCol = 0
        
        For i = 1 To rng.Areas.count
            With rng.Areas(i)
                If Not blnFromA1 Then
                    If .Row < lngStartRow Then lngStartRow = .Row
                    If .Column < lngStartCol Then lngStartCol = .Column
                End If
                If .Row + .Rows.count - 1 > lngMaxRow Then lngMaxRow = .Row + .Rows.count - 1
                If .Column + .Columns.count - 1 > lngMaxCol Then lngMaxCol = .Column + .Columns.count - 1
            End With
        Next i
        
        Set GetUsedRange = ws.Range(ws.Cells(lngStartRow, lngStartCol), ws.Cells(lngMaxRow, lngMaxCol))
    Else
        Set GetUsedRange = Nothing
    End If
End Function
