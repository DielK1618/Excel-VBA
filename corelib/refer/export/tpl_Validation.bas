Attribute VB_Name = "tpl_Validation"
Option Explicit
Function EvaluateFormula(ByVal strFormula As String) As Boolean
    
    Dim result As Variant
    
    On Error GoTo ErrorHandler
    
    ' 수식이 =로 시작하지 않으면 추가
    If Left(strFormula, 1) <> "=" Then
        strFormula = "=" & strFormula
    End If
    
    ' 수식을 직접 평가
    result = Application.Evaluate(strFormula)
    
    ' 결과가 TRUE인지 확인
    If IsError(result) Then
        EvaluateFormula = False
    ElseIf VarType(result) = vbBoolean Then
        EvaluateFormula = result
    ElseIf IsNumeric(result) Then
        EvaluateFormula = (result <> 0)
    Else
        EvaluateFormula = False
    End If
    
    Exit Function
    
ErrorHandler:
    EvaluateFormula = False
End Function
Sub SetIfValTrue(ByVal rng As Range, _
                                ByVal vntValue As Variant, _
                                Optional blnMsg As Boolean = False)
    
    On Error GoTo ErrorHandler
    
    Dim cel As Range
    
    For Each cel In rng
        If cel.Validation.Type <> 7 Then GoTo ErrorHandler
        
        Dim strFormula As Variant
        strFormula = cel.Validation.Formula1
        
        If EvaluateFormula(strFormula) Then
            cel = vntValue
        ElseIf blnMsg Then
            MsgBox "유효성 검사 조건 불일치!", vbCritical
        End If
    Next cel
    
    Exit Sub
ErrorHandler:
    rng = vntValue
End Sub
