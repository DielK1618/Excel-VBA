Attribute VB_Name = "tpl_Find"
Option Explicit
Function FindCellsByColor(ByVal intColor As Long, _
                          Optional ByVal rng As Range, _
                          Optional ByVal blnColorIndex As Boolean = True) As Range
    Dim f As Range
    Dim firstAddr As String
    Dim resultRng As Range
    Dim ws As Worksheet
    
    ' 검색 범위 기본값 및 시트 보존
    If rng Is Nothing Then
        Set ws = ActiveSheet
        Set rng = ws.Cells
    Else
        Set ws = rng.Parent  ' 원본 시트 저장
    End If
    
    Application.FindFormat.Clear
    
    ' 색상 방식 선택
    If blnColorIndex Then
        Application.FindFormat.Interior.ColorIndex = intColor
    Else
        Application.FindFormat.Interior.Color = intColor
    End If
    
    On Error Resume Next
    
    With rng
        Set f = .Find(What:="", LookIn:=xlFormulas, _
                      LookAt:=xlPart, SearchOrder:=xlByRows, _
                      SearchDirection:=xlNext, SearchFormat:=True)
        
        If Not f Is Nothing Then
            firstAddr = f.Address
            Set resultRng = f  ' 첫 번째 셀 저장
            
            Do
                Set f = .Find(What:="", After:=f, LookIn:=xlFormulas, _
                              LookAt:=xlPart, SearchOrder:=xlByRows, _
                              SearchDirection:=xlNext, SearchFormat:=True)
                
                If f Is Nothing Then Exit Do
                If f.Address = firstAddr Then Exit Do
                
                ' 시트 명시적으로 지정하여 Union
                Set resultRng = Union(resultRng, f)
            Loop
        End If
    End With
    
    On Error GoTo 0
    
    Application.FindFormat.Clear  ' 포맷 정리
    
    ' 결과 반환 (시트 유지)
    If Not resultRng Is Nothing Then
        Set FindCellsByColor = resultRng
    End If
End Function
