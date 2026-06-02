Attribute VB_Name = "tpl_Test"
Option Explicit

' ┌─────────────────────────────────────────────────────────┐
' │  tpl_Test                                               │
' │  역할 : am_Path.ReplacePath 테스트 프로시저 모음        │
' └─────────────────────────────────────────────────────────┘

' ── 0. 테스트 공통 ───────────────────────────────────────────
Private Sub PrintResult(ByVal strCase   As String, _
                        ByVal strInput  As String, _
                        ByVal strOutput As String)
    Debug.Print "--------------------------------------------"
    Debug.Print "[" & strCase & "]"
    Debug.Print "  입력 : " & strInput
    Debug.Print "  결과 : " & strOutput
    Debug.Print "  유효 : " & IIf(strOutput <> "", "OK", "FAIL - 빈 문자열")
End Sub

Private Function Run(ByVal strPath As String) As String
    Run = Application.Run("corelib.xlam!am_Path.ReplacePath", _
                          strPath, _
                          ThisWorkbook.Path, _
                          ThisWorkbook.FullName)
End Function

' ══════════════════════════════════════════════════════════
'  전체 테스트 일괄 실행
' ══════════════════════════════════════════════════════════
Public Sub RunAllTests()
    Debug.Print "============================================"
    Debug.Print "  am_Path.ReplacePath 전체 테스트 시작"
    Debug.Print "============================================"
    Call Test_FixedTokens
    Call Test_CustomTokens
    Call Test_AbsolutePath
    Call Test_NetworkPath
    Call Test_DriveMapping
    Call Test_EdgeCases
    Debug.Print "============================================"
    Debug.Print "  전체 테스트 완료"
    Debug.Print "============================================"
End Sub

' ══════════════════════════════════════════════════════════
'  테스트 1 - 고정 토큰
' ══════════════════════════════════════════════════════════
Public Sub Test_FixedTokens()

    Debug.Print vbCrLf & "▶ 테스트 1 - 고정 토큰"
    Dim strIn  As String
    Dim strOut As String

    ' {cPath} - CWB 폴더 경로 대체
    strIn  = "{cPath}"
    strOut = Run(strIn)
    Call PrintResult("cPath 토큰", strIn, strOut)

    ' {cFile} - CWB 전체 파일 경로
    strIn  = "{cFile}"
    strOut = Run(strIn)
    Call PrintResult("cFile 토큰", strIn, strOut)

    ' {xPath} - xlam 폴더 경로 대체
    strIn  = "{xPath}"
    strOut = Run(strIn)
    Call PrintResult("xPath 토큰", strIn, strOut)

    ' {xFile} - xlam 전체 파일 경로
    strIn  = "{xFile}"
    strOut = Run(strIn)
    Call PrintResult("xFile 토큰", strIn, strOut)

End Sub

' ══════════════════════════════════════════════════════════
'  테스트 2 - 커스텀 토큰
' ══════════════════════════════════════════════════════════
Public Sub Test_CustomTokens()

    Debug.Print vbCrLf & "▶ 테스트 2 - 커스텀 토큰"
    Dim strIn  As String
    Dim strOut As String

    ' {xDB} - xlam 하위 DB 폴더
    strIn  = "{xDB}"
    strOut = Run(strIn)
    Call PrintResult("xDB 토큰", strIn, strOut)

    ' {xBak} - xlam 하위 백업 폴더
    strIn  = "{xBak}"
    strOut = Run(strIn)
    Call PrintResult("xBak 토큰", strIn, strOut)

    ' 미등록 토큰 (치환 안되고 그대로 반환되어야 함)
    strIn  = "{미등록토큰}\test.xlsx"
    strOut = Run(strIn)
    Call PrintResult("미등록 토큰", strIn, strOut)

End Sub

' ══════════════════════════════════════════════════════════
'  테스트 3 - 절대 경로
' ══════════════════════════════════════════════════════════
Public Sub Test_AbsolutePath()

    Debug.Print vbCrLf & "▶ 테스트 3 - 절대 경로"
    Dim strIn  As String
    Dim strOut As String

    ' 유효한 절대 경로
    strIn  = ThisWorkbook.Path
    strOut = Run(strIn)
    Call PrintResult("유효한 절대 경로", strIn, strOut)

    ' 슬래시 혼용 경로
    strIn  = Replace(ThisWorkbook.Path, "\", "/")
    strOut = Run(strIn)
    Call PrintResult("슬래시 혼용", strIn, strOut)

    ' 드라이브 없는 절대 경로
    strIn  = Mid(ThisWorkbook.Path, 3)
    strOut = Run(strIn)
    Call PrintResult("드라이브 없는 경로", strIn, strOut)

End Sub

' ══════════════════════════════════════════════════════════
'  테스트 4 - 네트워크 경로
' ══════════════════════════════════════════════════════════
Public Sub Test_NetworkPath()

    Debug.Print vbCrLf & "▶ 테스트 4 - 네트워크 경로"
    Dim strIn  As String
    Dim strOut As String

    ' UNC 경로 (실제 환경에 맞게 수정 필요)
    strIn  = "\\서버명\공유폴더명"
    strOut = Run(strIn)
    Call PrintResult("UNC 경로", strIn, strOut)

End Sub

' ══════════════════════════════════════════════════════════
'  테스트 5 - 드라이브 매핑
' ══════════════════════════════════════════════════════════
Public Sub Test_DriveMapping()

    Debug.Print vbCrLf & "▶ 테스트 5 - 드라이브 매핑"
    Dim strIn  As String
    Dim strOut As String

    ' 존재하지 않는 드라이브 (A~Z 추정 후 확인)
    strIn  = "Z:\임시경로\test.xlsx"
    strOut = Run(strIn)
    Call PrintResult("없는 드라이브", strIn, strOut)

    ' blnChangeDrive = True 테스트
    strIn  = "Z:\" & Mid(ThisWorkbook.Path, 4)
    strOut = Application.Run("corelib.xlam!am_Path.ReplacePath", _
                             strIn, _
                             ThisWorkbook.Path, _
                             ThisWorkbook.FullName, _
                             True)
    Call PrintResult("드라이브 강제 교체", strIn, strOut)

End Sub

' ══════════════════════════════════════════════════════════
'  테스트 6 - 엣지 케이스
' ══════════════════════════════════════════════════════════
Public Sub Test_EdgeCases()

    Debug.Print vbCrLf & "▶ 테스트 6 - 엣지 케이스"
    Dim strIn  As String
    Dim strOut As String

    ' 빈 문자열
    strIn  = ""
    strOut = Run(strIn)
    Call PrintResult("빈 문자열", strIn, strOut)

    ' 토큰만 있는 경로
    strIn  = "{cPath}"
    strOut = Run(strIn)
    Call PrintResult("토큰만 있는 경로", strIn, strOut)

    ' 경로가 아닌 문자열
    strIn  = "이건경로가아닙니다"
    strOut = Run(strIn)
    Call PrintResult("잘못된 경로", strIn, strOut)

    ' 토큰 연속 사용
    strIn  = "{cPath}{xPath}"
    strOut = Run(strIn)
    Call PrintResult("토큰 연속 사용", strIn, strOut)

End Sub
