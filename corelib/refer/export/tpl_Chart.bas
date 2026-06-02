Attribute VB_Name = "tpl_Chart"
Option Explicit
Sub SetChartDataRange(ByVal strChart As String, _
                                                    ByVal rng As Range)
    
    On Error Resume Next
    With ActiveSheet.ChartObjects(strChart).Chart
        .SetSourceData Source:=rng
    End With
    On Error GoTo 0
    
End Sub
