Attribute VB_Name = "am_Format"
Option Explicit

' ┌─────────────────────────────────────────────────────────┐
' │  am_Format                                              │
' │  역할 : 조건부 서식, 유효성 검사                        │
' └─────────────────────────────────────────────────────────┘

' ══════════════════════════════════════════════════════════
'  조건부 서식
' ══════════════════════════════════════════════════════════

' 목적   : 수식 기반 조건부 서식 추가
' 인수   : rng             - 적용할 셀 범위
'          strFormula      - 수식 조건 (예: "=$A1>10")
'          lngFontColor    - 폰트 색상 (-1: 미적용)
'          lngBackColor    - 배경 색상 (-1: 미적용)
'          xlPtn           - 배경 패턴 종류
'          lngPatternColor - 패턴 색상 (-1: 미적용)
'          lngBorderColor  - 테두리 색상 (-1: 미적용)
'          blnBorderTop    - 위쪽 테두리 적용 여부
'          blnBorderBottom - 아래쪽 테두리 적용 여부
'          blnBorderLeft   - 왼쪽 테두리 적용 여부
'          blnBorderRight  - 오른쪽 테두리 적용 여부
'          blnStopIfTrue   - 조건 충족 시 이후 규칙 중단 여부
'          intPriority     - 서식 규칙 우선순위
' 예시   : ConditionalFormattingFormula Range("A1:A10"), "=$A1>100", _
'              lngFontColor:=RGB(255,0,0), lngBackColor:=RGB(255,255,0)
Public Sub ConditionalFormattingFormula( _
        ByVal rng As Range, _
        ByVal strFormula As String, _
        Optional ByVal lngFontColor As Long = -1, _
        Optional ByVal lngBackColor As Long = -1, _
        Optional ByVal xlPtn As XlPattern = xlPatternNone, _
        Optional ByVal lngPatternColor As Long = -1, _
        Optional ByVal lngBorderColor As Long = -1, _
        Optional ByVal blnBorderTop As Boolean = False, _
        Optional ByVal blnBorderBottom As Boolean = False, _
        Optional ByVal blnBorderLeft As Boolean = False, _
        Optional ByVal blnBorderRight As Boolean = False, _
        Optional ByVal blnStopIfTrue As Boolean = False, _
        Optional ByVal intPriority As Integer = 1)

    On Error GoTo ErrHandler

    If Len(Trim(strFormula)) = 0 Then
        MsgBox "strFormula(수식 조건)은 필수 입력값입니다.", _
               vbExclamation, am_Core.AM_NAME
        Exit Sub
    End If

    Dim objFC As FormatCondition
    Set objFC = rng.FormatConditions.Add(Type:=xlExpression, Formula1:=strFormula)

    With objFC
        If lngFontColor <> -1 Then .Font.Color = lngFontColor
        If lngBackColor <> -1 Then .Interior.Color = lngBackColor

        If xlPtn <> xlPatternNone Then
            .Interior.Pattern = xlPtn
            If lngPatternColor <> -1 Then .Interior.PatternColor = lngPatternColor
        End If

        If lngBorderColor <> -1 Then
            With .Borders
                If blnBorderTop Then .Item(xlEdgeTop).Color = lngBorderColor
                If blnBorderBottom Then .Item(xlEdgeBottom).Color = lngBorderColor
                If blnBorderLeft Then .Item(xlEdgeLeft).Color = lngBorderColor
                If blnBorderRight Then .Item(xlEdgeRight).Color = lngBorderColor
            End With
        End If

        .StopIfTrue = blnStopIfTrue
        .Priority = intPriority
    End With

    Set objFC = Nothing
    Exit Sub

ErrHandler:
    MsgBox "ConditionalFormattingFormula 오류 " & Err.Number & ": " & Err.Description, _
           vbCritical, am_Core.AM_NAME
    Set objFC = Nothing

End Sub

' 목적   : 조건부 서식 전체 삭제
' 인수   : rng - 삭제할 셀 범위
' 예시   : ClearConditionalFormatting(Range("A1:A10"))
Public Sub ClearConditionalFormatting(ByVal rng As Range)
    On Error Resume Next
    rng.FormatConditions.Delete
    On Error GoTo 0
End Sub

' 목적   : 색상 배율 조건부 서식 추가 (2단 또는 3단)
' 인수   : rng         - 적용할 셀 범위
'          lngMinColor - 최솟값 색상
'          lngMaxColor - 최댓값 색상
'          lngMidColor - 중간값 색상 (-1: 2단 척도)
'          intPriority - 서식 규칙 우선순위
' 예시   : ConditionalFormattingColorScale Range("A1:A10"), RGB(255,0,0), RGB(0,255,0)
'          ConditionalFormattingColorScale Range("A1:A10"), RGB(255,0,0), RGB(0,255,0), RGB(255,255,0)
Public Sub ConditionalFormattingColorScale( _
        ByVal rng As Range, _
        ByVal lngMinColor As Long, _
        ByVal lngMaxColor As Long, _
        Optional ByVal lngMidColor As Long = -1, _
        Optional ByVal intPriority As Integer = 1)

    On Error GoTo ErrHandler

    Dim objCS As ColorScale

    If lngMidColor = -1 Then
        Set objCS = rng.FormatConditions.AddColorScale(ColorScaleType:=2)
        With objCS
            .ColorScaleCriteria(1).Type = xlConditionValueLowestValue
            .ColorScaleCriteria(1).FormatColor.Color = lngMinColor
            .ColorScaleCriteria(2).Type = xlConditionValueHighestValue
            .ColorScaleCriteria(2).FormatColor.Color = lngMaxColor
            .Priority = intPriority
        End With
    Else
        Set objCS = rng.FormatConditions.AddColorScale(ColorScaleType:=3)
        With objCS
            .ColorScaleCriteria(1).Type = xlConditionValueLowestValue
            .ColorScaleCriteria(1).FormatColor.Color = lngMinColor
            .ColorScaleCriteria(2).Type = xlConditionValuePercentile
            .ColorScaleCriteria(2).Value = 50
            .ColorScaleCriteria(2).FormatColor.Color = lngMidColor
            .ColorScaleCriteria(3).Type = xlConditionValueHighestValue
            .ColorScaleCriteria(3).FormatColor.Color = lngMaxColor
            .Priority = intPriority
        End With
    End If

    Set objCS = Nothing
    Exit Sub

ErrHandler:
    MsgBox "ConditionalFormattingColorScale 오류 " & Err.Number & ": " & Err.Description, _
           vbCritical, am_Core.AM_NAME
    Set objCS = Nothing

End Sub

' 목적   : 데이터 막대 조건부 서식 추가
' 인수   : rng          - 적용할 셀 범위
'          lngBarColor  - 막대 색상
'          blnShowValue - 셀 값 표시 여부 (기본: True)
'          blnGradient  - 그라디언트 채우기 여부 (기본: True)
'          intPriority  - 서식 규칙 우선순위
' 예시   : ConditionalFormattingDataBar Range("A1:A10"), RGB(0,112,192)
Public Sub ConditionalFormattingDataBar( _
        ByVal rng As Range, _
        ByVal lngBarColor As Long, _
        Optional ByVal blnShowValue As Boolean = True, _
        Optional ByVal blnGradient As Boolean = True, _
        Optional ByVal intPriority As Integer = 1)

    On Error GoTo ErrHandler

    Dim objDB As Databar
    Set objDB = rng.FormatConditions.AddDatabar

    With objDB
        .BarColor.Color = lngBarColor
        .BarFillType = IIf(blnGradient, xlDataBarFillGradient, xlDataBarFillSolid)
        .ShowValue = blnShowValue
        .MinPoint.Modify xlConditionValueAutomaticMin
        .MaxPoint.Modify xlConditionValueAutomaticMax
        .Priority = intPriority
    End With

    Set objDB = Nothing
    Exit Sub

ErrHandler:
    MsgBox "ConditionalFormattingDataBar 오류 " & Err.Number & ": " & Err.Description, _
           vbCritical, am_Core.AM_NAME
    Set objDB = Nothing

End Sub

' ══════════════════════════════════════════════════════════
'  유효성 검사
' ══════════════════════════════════════════════════════════

' 목적   : 드롭다운 목록 유효성 검사 간편 설정
' 인수   : rngTarget - 유효성 검사 적용 범위
'          arrValues - 목록 값 배열
' 예시   : ValidationList(Sheet1.Range("B2:B10"), Array("승인", "반려", "대기"))
Public Sub ValidationList(ByVal rngTarget As Range, _
                          ByVal arrValues As Variant)

    Dim strValues As String
    strValues = Join(arrValues, ",")

    With rngTarget.Validation
        .Delete
        If strValues <> "" Then
            .Add xlValidateList, xlValidAlertStop, xlBetween, strValues
        End If
    End With

End Sub

' 목적   : 범위 셀 유효성 검사 설정
' 인수   : rng               - 적용할 셀 범위
'          strFormula1       - 유효성 검사가 되는 조건
'          strFormula2       - 최댓값 (xlBetween 경우 사용)
'          xlVldType         - 유효성 검사 유형
'          xlAlertStyle      - 경고 알림 스타일
'          xlOpr             - 값 연산자
'          blnIgnoreBlank    - 빈 셀 무시 여부
'          blnInCellDropdown - 셀 내 드롭다운 표시 여부
'          strInputTitle     - 입력 메시지 제목
'          strInputMessage   - 입력 메시지 내용
'          strErrorTitle     - 오류 메시지 제목
'          strErrorMessage   - 오류 메시지 내용
'          xlIME             - IME 입력기 모드
'          blnShowInput      - 입력 메시지 표시 여부
'          blnShowError      - 오류 메시지 표시 여부
' 예시   : SetValidation Range("A1:A10"), "남,여", , xlValidateList
Public Sub SetValidation( _
        ByVal rng As Range, _
        Optional ByVal strFormula1 As String = "", _
        Optional ByVal strFormula2 As String = "", _
        Optional ByVal xlVldType As XlDVType = xlValidateInputOnly, _
        Optional ByVal xlAlertStyle As XlDVAlertStyle = xlValidAlertStop, _
        Optional ByVal xlOpr As XlFormatConditionOperator = xlBetween, _
        Optional ByVal blnIgnoreBlank As Boolean = True, _
        Optional ByVal blnInCellDropdown As Boolean = True, _
        Optional ByVal strInputTitle As String = "", _
        Optional ByVal strInputMessage As String = "", _
        Optional ByVal strErrorTitle As String = "", _
        Optional ByVal strErrorMessage As String = "", _
        Optional ByVal xlIME As XlIMEMode = xlIMEModeNoControl, _
        Optional ByVal blnShowInput As Boolean = True, _
        Optional ByVal blnShowError As Boolean = True)

    On Error GoTo ErrHandler

    If xlVldType <> xlValidateInputOnly And Len(Trim(strFormula1)) = 0 Then
        MsgBox "목록 외 유효성 검사는 strFormula1(조건값/범위)이 필요합니다.", _
               vbExclamation, am_Core.AM_NAME
        Exit Sub
    End If

    With rng.Validation
        .Delete

        If Len(Trim(strFormula2)) > 0 Then
            .Add Type:=xlVldType, AlertStyle:=xlAlertStyle, _
                 Operator:=xlOpr, Formula1:=strFormula1, Formula2:=strFormula2
        ElseIf Len(Trim(strFormula1)) > 0 Then
            .Add Type:=xlVldType, AlertStyle:=xlAlertStyle, _
                 Operator:=xlOpr, Formula1:=strFormula1
        Else
            .Add Type:=xlVldType
        End If

        .IgnoreBlank = blnIgnoreBlank
        .InCellDropdown = blnInCellDropdown
        .InputTitle = strInputTitle
        .InputMessage = strInputMessage
        .ErrorTitle = strErrorTitle
        .ErrorMessage = strErrorMessage
        .IMEMode = xlIME
        .ShowInput = blnShowInput
        .ShowError = blnShowError
    End With

    Exit Sub

ErrHandler:
    MsgBox "SetValidation 오류 " & Err.Number & ": " & Err.Description, _
           vbCritical, am_Core.AM_NAME

End Sub

' 목적   : 범위 유효성 검사 삭제
' 인수   : rng - 삭제할 셀 범위
' 예시   : ClearValidation(Range("A1:A10"))
Public Sub ClearValidation(ByVal rng As Range)
    On Error Resume Next
    rng.Validation.Delete
    On Error GoTo 0
End Sub
