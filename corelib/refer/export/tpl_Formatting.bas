Attribute VB_Name = "tpl_Formatting"
Option Explicit
'===============================================================================
' [ConditionalFormattingFormula] 조건부 서식 추가 (수식 기반)
'-------------------------------------------------------------------------------
' 설명  : 수식(Formula)을 기반으로 지정 범위에 조건부 서식을 적용합니다.
' 작성  : KyuCheol | 최종수정 : 2026-03-25
'-------------------------------------------------------------------------------
' 매개변수:
'   rng             - (필수) 조건부 서식을 적용할 셀 범위
'   strFormula      - (필수) 조건 수식 (예: "=$A1>10")
'   lngFontColor    - (선택) 폰트 색상. RGB() 또는 색상 상수 사용. 생략 시 미적용
'   lngBackColor    - (선택) 배경 색상. 생략 시 미적용
'   xlPtn           - (선택) 배경 무늬 패턴 (XlPattern 열거형). 생략 시 미적용
'   lngPatternColor - (선택) 무늬 색상. xlPtn 지정 시 함께 사용
'   lngBorderColor  - (선택) 테두리 색상. 생략 시 미적용
'   blnBorderTop    - (선택) 위쪽 테두리 적용 여부 (기본: False)
'   blnBorderBottom - (선택) 아래쪽 테두리 적용 여부 (기본: False)
'   blnBorderLeft   - (선택) 왼쪽 테두리 적용 여부 (기본: False)
'   blnBorderRight  - (선택) 오른쪽 테두리 적용 여부 (기본: False)
'   blnStopIfTrue   - (선택) 조건 충족 시 이후 서식 규칙 중단 여부 (기본: False)
'   intPriority     - (선택) 서식 규칙 우선순위. 숫자가 낮을수록 우선 (기본: 1)
'-------------------------------------------------------------------------------
' 사용 예시:
'   ' ① 기본 - 폰트/배경색만 지정
'   ConditionalFormattingFormula Range("A1:A10"), "=$A1>100", _
'       lngFontColor:=RGB(255,0,0), lngBackColor:=RGB(255,255,0)
'
'   ' ② 테두리 포함
'   ConditionalFormattingFormula Range("B1:B10"), "=$B1=""완료""", _
'       lngFontColor:=RGB(0,0,0), lngBorderColor:=RGB(0,0,255), _
'       blnBorderTop:=True, blnBorderBottom:=True
'
'   ' ③ 배경 무늬 포함
'   ConditionalFormattingFormula Range("C1:C5"), "=$C1<0", _
'       xlPtn:=xlPatternGray25, lngPatternColor:=RGB(255,0,0)
'===============================================================================
Sub ConditionalFormattingFormula( _
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

    ' 입력값 방어: 수식이 비어있으면 중단
    If Len(Trim(strFormula)) = 0 Then
        MsgBox "strFormula(조건 수식)는 필수 입력값입니다.", vbExclamation, "입력 오류"
        Exit Sub
    End If

    Dim objFC As FormatCondition
    Set objFC = rng.FormatConditions.Add(Type:=xlExpression, Formula1:=strFormula)

    With objFC
        ' 폰트 색상: -1은 미적용 센티널 값 → 0(검정)도 정상 적용됨
        If lngFontColor <> -1 Then .Font.Color = lngFontColor

        ' 배경 색상
        If lngBackColor <> -1 Then .Interior.Color = lngBackColor

        ' 배경 무늬 패턴
        If xlPtn <> xlPatternNone Then
            .Interior.Pattern = xlPtn
            If lngPatternColor <> -1 Then .Interior.PatternColor = lngPatternColor
        End If

        ' 테두리: 색상이 지정된 경우에만 각 방향 적용
        If lngBorderColor <> -1 Then
            With .Borders
                If blnBorderTop Then .item(xlEdgeTop).Color = lngBorderColor
                If blnBorderBottom Then .item(xlEdgeBottom).Color = lngBorderColor
                If blnBorderLeft Then .item(xlEdgeLeft).Color = lngBorderColor
                If blnBorderRight Then .item(xlEdgeRight).Color = lngBorderColor
            End With
        End If

        .StopIfTrue = blnStopIfTrue

        ' 우선순위: 다른 속성 설정 완료 후 마지막에 적용
        .Priority = intPriority
    End With

    Set objFC = Nothing
    Exit Sub

ErrHandler:
    MsgBox "ConditionalFormattingFormula 오류 " & Err.Number & ": " & Err.Description, _
           vbCritical, "오류"
    Set objFC = Nothing
End Sub


'===============================================================================
' [SetValidation] 데이터 유효성 검사 설정
'-------------------------------------------------------------------------------
' 설명  : 지정 범위에 데이터 유효성 검사 규칙을 추가합니다.
'         기존 유효성 검사는 자동 삭제 후 재설정됩니다.
' 작성  : KyuCheol | 최종수정 : 2026-03-25
'-------------------------------------------------------------------------------
' 매개변수:
'   rng              - (필수) 유효성 검사를 적용할 셀 범위
'   strFormula1      - (선택) 유효성 기준값 또는 목록 수식
'                      · xlValidateList    → 목록 수식  (예: "사과,배,감" 또는 "=$A$1:$A$5")
'                      · xlValidateWholeNumber → 최솟값 (예: "1")
'                      · xlValidateCustom  → 사용자 정의 수식 (예: "=ISNUMBER(A1)")
'   strFormula2      - (선택) xlBetween/xlNotBetween 사용 시 최댓값 (예: "100")
'   xlVldType        - (선택) 유효성 검사 유형 (기본: xlValidateInputOnly = 제한없음)
'   xlAlertStyle     - (선택) 오류 알림 스타일 (기본: xlValidAlertStop)
'   xlOpr            - (선택) 비교 연산자 (기본: xlBetween)
'   blnIgnoreBlank   - (선택) 빈 셀 무시 여부 (기본: True)
'   blnInCellDropdown- (선택) 셀 내 드롭다운 표시 여부 (기본: True)
'   strInputTitle    - (선택) 입력 메시지 제목
'   strInputMessage  - (선택) 입력 메시지 내용
'   strErrorTitle    - (선택) 오류 메시지 제목
'   strErrorMessage  - (선택) 오류 메시지 내용
'   xlIME            - (선택) IME 입력기 모드 (기본: xlIMEModeNoControl)
'   blnShowInput     - (선택) 입력 메시지 표시 여부 (기본: True)
'   blnShowError     - (선택) 오류 메시지 표시 여부 (기본: True)
'-------------------------------------------------------------------------------
' 사용 예시:
'   ' ① 목록형 유효성 (직접 입력)
'   SetValidation Range("A1:A10"), _
'       strFormula1:="""사과,배,감""", _
'       xlVldType:=xlValidateList
'
'   ' ② 목록형 유효성 (범위 참조)
'   SetValidation Range("B1:B10"), _
'       strFormula1:="=$D$1:$D$5", _
'       xlVldType:=xlValidateList, _
'       strInputTitle:="선택", strInputMessage:="목록에서 선택하세요"
'
'   ' ③ 정수 범위 제한 (1 ~ 100)
'   SetValidation Range("C1:C10"), _
'       strFormula1:="1", strFormula2:="100", _
'       xlVldType:=xlValidateWholeNumber, _
'       xlOpr:=xlBetween, _
'       strErrorTitle:="입력 오류", strErrorMessage:="1~100 사이의 정수만 입력 가능합니다."
'
'   ' ④ 사용자 정의 수식
'   SetValidation Range("D1:D10"), _
'       strFormula1:="=ISNUMBER(D1)", _
'       xlVldType:=xlValidateCustom, _
'       strErrorMessage:="숫자만 입력 가능합니다."
'===============================================================================
Sub SetValidation( _
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

    ' Formula1이 필요한 유형인데 비어있으면 경고
    Dim blnNeedsFormula As Boolean
    blnNeedsFormula = (xlVldType <> xlValidateInputOnly)

    If blnNeedsFormula And Len(Trim(strFormula1)) = 0 Then
        MsgBox "선택한 유효성 유형은 strFormula1(기준값/수식)이 필요합니다.", _
               vbExclamation, "입력 오류"
        Exit Sub
    End If

    With rng.Validation
        .Delete

        ' Formula1/Formula2 유무에 따라 분기
        If Len(Trim(strFormula2)) > 0 Then
            ' Between / NotBetween 등 두 값이 필요한 경우
            .Add Type:=xlVldType, AlertStyle:=xlAlertStyle, _
                 operator:=xlOpr, Formula1:=strFormula1, Formula2:=strFormula2
        ElseIf Len(Trim(strFormula1)) > 0 Then
            ' 단일 기준값 또는 목록/사용자정의 수식
            .Add Type:=xlVldType, AlertStyle:=xlAlertStyle, _
                 operator:=xlOpr, Formula1:=strFormula1
        Else
            ' xlValidateInputOnly: 제한 없음 (입력 메시지만 사용)
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
           vbCritical, "오류"
End Sub

