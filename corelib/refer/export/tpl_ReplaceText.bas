Attribute VB_Name = "tpl_ReplaceText"
Option Explicit
Function ReplaceText(ByVal strValue As String, _
                                        Optional ByVal rngBase As Range) As String '문자 변환
                                
    Dim strThisWorkbookPath, Path As String
    strThisWorkbookPath = tPath
    
    '기본
    strValue = Replace(strValue, "{오늘}", Date)
    strValue = Replace(strValue, "{지금}", Now)
    
    '공통
    Dim r, rng As Range
    
    Set rng = GetTwbRange("T_문자변환관리_공통[문자]")
    
    For Each r In rng
        If r <> "" Then
            strValue = Replace(strValue, r, Range(r.Offset(, 1)))
        End If
    Next
    
    '개별
    If Not rngBase Is Nothing Then
        
        Set rng = GetTwbRange("T_문자변환관리_개별[문자]")
        
        On Error Resume Next
        For Each r In rng
            If r <> "" Then
                With rngBase
                    strValue = Replace(strValue, "{" & r & "}", .Offset(, intOffset(rngBase, CStr(r))))
                End With
            End If
        Next
        On Error GoTo 0
    End If
    
    ReplaceText = strValue
        
End Function
Function ConvertToExcelSerialDate(ByVal strDate As String, _
                                                                Optional ByVal strTime As String = "00:00:00") As Double
                                                                
    Dim dtFullDateTime As Date ' 날짜와 시간을 결합하여 날짜 변수 생성
    dtFullDateTime = CDate(strDate & " " & strTime)
    ' 엑셀 시리얼 데이트 넘버로 변환
    ' 엑셀의 날짜 기준은 1900년 1월 1일이며, 정수 부분은 날짜, 소수 부분은 시간을 나타냄
    ConvertToExcelSerialDate = DateValue(strDate) + TimeValue(strTime)
    
End Function
Function ExtractValues(ByVal strValue As String, _
                                          ByVal strPattern As String) As Variant
    Dim regEx As Object
    Dim matches As Object
    Dim arrValues() As Variant
    Dim varValue As Variant
    Dim dateTimeStr As String
    Dim i, j As Long
    
    ' 정규 표현식 객체 생성
    Set regEx = CreateObject("VBScript.RegExp")
    
    regEx.Pattern = strPattern '"(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})" (8자리 날짜 + 6자리 시간)
    regEx.Global = False
    
    ' 정규 표현식 실행
    Set matches = regEx.Execute(strValue)

    ' 매치가 있으면 처리
    If matches.count > 0 Then
        '각 부분 추출
        For i = 0 To matches.count - 1
            For Each varValue In matches(i).SubMatches
                ReDim Preserve arrValues(j)
                arrValues(j) = varValue
                j = j + 1
            Next
        Next
    End If
    
    ExtractValues = arrValues
    
End Function
