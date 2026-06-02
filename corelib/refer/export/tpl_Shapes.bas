Attribute VB_Name = "tpl_Shapes"
Option Explicit
'<<< 도형의 메크로 실행 ========================================
Sub ExampleRunShpMacro()
    Call RunShpMacro("SEARCH")
End Sub
Sub RunShpMacro(ByVal strShpName As String, Optional ByVal ws As Worksheet)
    If ws Is Nothing Then Set ws = ActiveSheet
    
    Dim shp As Shape
    Dim txt As String
    Dim strMacro As String
    
    For Each shp In ws.Shapes
        txt = GetShapeTextSafe(shp)
        If Len(Trim(txt)) > 0 Then
            If UCase(Trim(txt)) = UCase(strShpName) Then
                On Error Resume Next
                strMacro = shp.OnAction
                On Error GoTo 0
                If Len(strMacro) > 0 Then
                    Application.Run strMacro
                Else
                    ' 연결 매크로가 없으면 원하면 메시지 표시
                    ' MsgBox "도형에 연결된 매크로가 없습니다: " & shp.Name
                End If
            End If
        End If
    Next shp
End Sub

' 도형(또는 그룹 내부)의 텍스트를 안전하게 반환하는 함수
Function GetShapeTextSafe(shp As Shape) As String
    Dim s As Shape
    Dim txtOut As String
    txtOut = ""
    
    On Error GoTo CleanFail
    
    ' 그룹이면 그룹 내부 아이템 순회
    If shp.Type = msoGroup Then
        For Each s In shp.GroupItems
            txtOut = txtOut & " " & GetShapeTextSafe_GItem(s)
        Next s
    Else
        txtOut = GetShapeTextSafe_GItem(shp)
    End If

CleanFail:
    GetShapeTextSafe = Trim(txtOut)
    Exit Function
End Function

' 개별 Shape에 대해 안전하게 텍스트를 얻음 (보조 함수)
Private Function GetShapeTextSafe_GItem(s As Shape) As String
    Dim res As String
    res = ""
    
    On Error Resume Next
    ' 먼저 TextFrame2 사용 가능하고 텍스트가 있는지 확인
    If Not s.TextFrame2 Is Nothing Then
        If s.TextFrame2.HasText = msoTrue Then
            res = s.TextFrame2.TextRange.Text
            GoTo Done
        End If
    End If
    ' 레거시 TextFrame 체크
    If Not s.TextFrame Is Nothing Then
        If s.TextFrame.HasText = msoTrue Then
            res = s.TextFrame.Characters.Text
            GoTo Done
        End If
    End If

Done:
    On Error GoTo 0
    GetShapeTextSafe_GItem = Trim(res)
End Function

