Attribute VB_Name = "prj_Upload"
Option Explicit
Public blnStop As Boolean
Function CheckBadgeNo(ByVal ActColor As Long, _
                                            Optional ByVal sht As Worksheet) As Integer
    
    If sht Is Nothing Then Set sht = ActiveSheet
    
    Dim shp As Shape
    
    On Error Resume Next
    For Each shp In sht.Shapes
        If shp.Name Like "B[0-9]" Then
            If shp.Fill.ForeColor.RGB = ActColor Then
                CheckBadgeNo = CInt(Replace(shp.Name, "B", ""))
                Exit Function
            End If
        End If
    Next
    On Error GoTo 0
    
End Function
Sub MarkBadge(ByVal intNo As Integer, _
                            ByVal ActColor As Long, _
                            ByVal DeActColor As Long, _
                            Optional ByVal sht As Worksheet)
    
    If sht Is Nothing Then Set sht = ActiveSheet
    
    Dim shp As Shape
    
    On Error Resume Next
    For Each shp In sht.Shapes
        If shp.Name Like "B[0-9]" Then
            If intNo = 0 Or shp.Name <> "B" & intNo Then
                shp.Fill.ForeColor.RGB = DeActColor
            Else
                shp.Fill.ForeColor.RGB = ActColor
            End If
        End If
    Next
    On Error GoTo 0
    
End Sub
'<<< 버튼 모음 >>> ===========================================================================
Sub BtCheckFolders()

    If MsgBox("실행하시겠습니까?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    Dim strPath As String
    Dim strCatPath As String
    Dim arrCat As Variant, vntCat As Variant
    Dim r As Range, rng As Range
    
    strPath = ReplacePath("{tPath}\00_다운파일\")
    arrCat = Array("학생명단", "수강현황")
    
    For Each vntCat In arrCat
        If vntCat = "학생명단" Then
            Set rng = GetTwbRange("DB_시즌[코드]")
        ElseIf vntCat = "수강현황" Then
            Set rng = GetTwbRange("DB_챕터[시즌코드]")
        End If
        
        For Each r In rng
            If vntCat = "학생명단" Then
                strCatPath = strPath & vntCat & "\" & r
            ElseIf vntCat = "수강현황" Then
                strCatPath = strPath & vntCat & "\" & r & "\" & r.Offset(, 1)
            End If
            Call MkFolder(strCatPath)
        Next r
    Next vntCat
    
    Call MarkBadge(1, uiColor.TC02, uiColor.TC16)
    MsgBox "Completion", vbInformation

Exit Sub
ErrorHandler:
    Call HandleError("CheckFolders")

End Sub
Sub BtDeleteFiles()
 
    Dim sht As Worksheet
    Set sht = ActiveSheet

    Dim intNo As Integer
    intNo = 1
    
    If CheckBadgeNo(uiColor.TC02, sht) < intNo Then
        MsgBox "이전 단계를 먼저 실행하세요!", vbInformation
        Exit Sub
    End If
    
    If MsgBox("실행하시겠습니까?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    Dim vntFile As Variant, arrFiles As Variant
    arrFiles = GetFilesList(ReplacePath("{tPath}\00_다운파일\"))
    
    If IsArrayEmpty(arrFiles) = False Then
        For Each vntFile In arrFiles
            Call DelFile(vntFile)
        Next
    End If
    
    Call MarkBadge(intNo + 1, uiColor.TC02, uiColor.TC16, sht)
    MsgBox "Completion", vbInformation

Exit Sub
ErrorHandler:
    Call HandleError("BtDeleteFiles")
    
End Sub
Sub BtDownloadFiles()

    Dim sht As Worksheet
    Set sht = ActiveSheet

    Dim intNo As Integer
    intNo = 2
    
    If CheckBadgeNo(uiColor.TC02, sht) < intNo Then
        MsgBox "이전 단계를 먼저 실행하세요!", vbInformation
        Exit Sub
    End If
    
    If MsgBox("실행하시겠습니까?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    
    '<<< 프로시저===
'    On Error GoTo ErrorHandler
    
    OpenEduLmsSite strURL 'EDU LMS 사이트 오픈
    
    '================================================================================
    ' [변수 선언부] : 반복문/조건문 안에서 선언되던 것들을 "사용 범위" 기준으로 위로 정리
    '================================================================================
    
    '--- 시즌/카테고리 루프에 공통으로 쓰이는 변수 -------------------------------
    Dim arrCat As Variant                 '카테고리 목록: ("학생명단", "수강현황")
    Dim vntCat As Variant                 '카테고리 반복용
    Dim r As Range                        '시즌 테이블 한 행(코드 셀) 반복용
    Dim rng As Range                      '시즌 코드 범위(DB_시즌[코드])
    Dim strPath As String, strFolder As String '다운로드 저장 경로 (카테고리별로 붙여서 사용)
    Dim strSeason As String               '시즌 코드(문자)
    Dim intSeason As Integer              '시즌 번호(시즌 코드 옆 컬럼 값)
    
    '--- 수강현황(강좌 옵션 반복)에서만 쓰이는 변수 -------------------------------
    Dim Dropdown As Selenium.WebElement   '[수강현황] 강좌 선택 드롭다운
    Dim strOptions As String              '[수강현황] 드롭다운 텍스트에서 옵션 문자열 추출용
    Dim arrOptions() As String            '[수강현황] 옵션 배열
    Dim i As Long, j As Long     '[수강현황] 옵션 반복 인덱스(i), 저장 폴더 번호(j)

    Dim checkTable As Selenium.WebElement '[수강현황] listGrid 테이블 존재 확인
    Dim secondTr As Selenium.WebElement   '[수강현황] 테이블 tbody 두번째 tr 확인(데이터 유무)
    Dim trValue As String                 '[수강현황] secondTr.Text 값 저장(빈 값이면 다운로드 스킵)
    
    '※ 참고: 아래 Dim r As Range, rng As Range 처럼 "Dim ... As Range"는
    '  Excel 오브젝트라 선언은 위, Set은 사용 직전에 하는 편이 안전.
    
    '================================================================================

    Set rng = GetTwbRange("DB_시즌[코드]")
    arrCat = Array("학생명단", "수강현황")
    strPath = ReplacePath("{tPath}\00_다운파일\")
    
    With sl
    
        For Each r In rng
            If Date <= r.Offset(, intOffset(r, "시작일")) Then GoTo PassCat '시작전 시즌은 패스
            
            strSeason = r
            intSeason = r.Offset(, 1)
    
            For Each vntCat In arrCat
            
                '<< 관리자 페이지 이동 ===
                .Wait 1000
                .Get "https://edulms.watv.org/admin/mainV2.wmc"
                .Wait 2000    '=== 관리자 페이지 이동 >>
                
                '카테고리별 저장 경로 구성 (기존 로직 유지)
                strFolder = strPath & vntCat & "\" & r & "\"
                
                If vntCat = "학생명단" Then
    
                    '학생관리 명단 다운로드 ==========================================================
                    .FindElementByCss("body > div.navLeftAdminWrap > div.navGroup > button:nth-child(9)").Click
                    .FindElementByCss("body > div.navLeftAdminWrap > div.navGroup > div:nth-child(10) > button").Click
                    .Wait 3000
    
                    '아이프레임 시작
                    .SwitchToFrame .FindElementByCss(".imIfram")
                    
                    '전공 선택 -----------------------------------------------------------------------
                    With .FindElementByXPath("/html/body/div[7]/div[2]/div[1]/div[1]/div[1]/table/tbody/tr[2]/td/div/div/input[1]")
                        If .value <> "" Then
                            sl.Wait 2000
                            .Click  ' 포커스 확보
                            .SendKeys sl.Keys.Control & "a"  ' 전체 선택
                            .SendKeys sl.Keys.Delete          ' 삭제
                        End If

                        .SendKeys "봉사자교육 > 행정담당자교육 > 건물담당자교육 > 건물관리 시즌" & intSeason
                        .SendKeys sl.Keys.Enter
                    End With
                    
                    .Wait 1000
                    
                    '엑셀 다운로드 -------------------------------------------------------------------
                    .FindElementByXPath("/html/body/div[7]/div[2]/div[1]/div[3]/button[1]").Click
                    .FindElementByXPath("/html/body/div[9]/div[3]/button[2]").Click

                    '마우스와 키보드 제어
                    ClickAtPosition 1377, 269, , "00:00:02" 'WMC_SCDK열기
                    ClickAtPosition 775, 366, , "00:00:03" '내보내기
                    ClickAtPosition 1223, 376, , "00:00:01", "00:00:01" '엑셀 저장하기
                    SendKeys "%{D}", True '파일 주소 표시줄 선택
                    SendKeys strFolder, True  '파일 경로 지정
                    SendKeys "{Enter}", True '파일 경로 이동
                    ClickAtPosition 1718, 1009, , "00:00:02" '저장
                    SendKeys "{Enter}", True
                    .Wait 200
                    SendKeys "%{F4}", True
                    .FindElementByXPath("/html/body/div[9]/div[3]/button[1]").Click
    
                ElseIf vntCat = "수강현황" Then
    
                    .FindElementByCss("body > div.navLeftAdminWrap > div.navGroup > button.btnNav.btnNavDown.lineTop").Click
                    .Wait 1000
                    .FindElementByCss("body > div.navLeftAdminWrap > div.navGroup > div:nth-child(4) > button:nth-child(4)").Click
                    .Wait 3000
    
                    '아이프레임 시작
                    .SwitchToFrame .FindElementByCss("#iframe_수강현황")
    
                    '전공 선택 -----------------------------------------------------------------------
                    With .FindElementByXPath("/html/body/div[7]/div[2]/div[1]/table/tbody/tr[1]/td/div/div/input[1]")
                        .ExecuteScript "document.getElementsByClassName('dhxcombo_input')[0].value = '';"
                        .SendKeys "봉사자교육 > 행정담당자교육 > 건물담당자교육 > 건물관리 시즌" & intSeason
                        .SendKeys sl.Keys.Enter
                    End With
                    
                    '강좌 선택 -----------------------------------------------------------------------
                    Set Dropdown = .FindElementByXPath("/html/body/div[7]/div[2]/div[1]/table/tbody/tr[2]/td/select")
                    
                    .Wait 1000
                    '옵션 목록 파싱(기존 로직 유지)
                    strOptions = Replace(Dropdown.Text, "선택", "")
                    strOptions = Replace(strOptions, ")", ")|")
                    strOptions = Mid(strOptions, 1, Len(strOptions) - 1)
                    arrOptions = Split(strOptions, "|")
    
                    '옵션 반복 다운로드 ---------------------------------------------------------------
                    j = 0 '※ 시즌/카테고리 내에서 폴더 번호를 1부터 다시 쓰려면 여기서 초기화 (기존 로직 의도 반영)
    
                    For i = UBound(arrOptions) To 0 Step -1
    
                        j = j + 1
    
                        .Wait 2000
                        Dropdown.SendKeys arrOptions(i)
                        Dropdown.SendKeys sl.Keys.Enter
                        .Wait 3000
                        
                        ' 테이블 찾기 - 지정된 XPath 사용
                        Set checkTable = .FindElementByXPath("//*[@id=""listGrid""]/div/div[3]/div[2]/table")
    
                        If Not checkTable Is Nothing Then
                            ' 두번째 tr 찾기 - 정확한 XPath 사용
                            Set secondTr = .FindElementByXPath("//*[@id=""listGrid""]/div/div[3]/div[2]/table/tbody/tr[2]")
                            If Not secondTr Is Nothing Then
                                ' tr의 텍스트 값 확인
                                trValue = Trim(secondTr.Text)
    
                                ' 값이 비어있으면 다음으로 넘어가기
                                If trValue = "" Then
                                    Debug.Print arrOptions(i) & "; 값이; 비어; 있어; 다운로드하지; 않습니다!"
                                    GoTo PassOption
                                End If
                            Else
                                Debug.Print "두번째 TR을 찾을 수 없습니다: " & arrOptions(i)
                                GoTo PassOption
                            End If
                        Else
                            Debug.Print "테이블을 찾을 수 없습니다: " & arrOptions(i)
                            GoTo PassOption
                        End If
    
                        .FindElementByXPath("/html/body/div[7]/div[2]/div[1]/div/button[5]").Click
                        .FindElementByXPath("/html/body/div[9]/div[3]/button[2]").Click
    
                        '마우스와 키보드 제어
                        ClickAtPosition 1377, 269, , "00:00:02" 'WMC_SCDK열기
                        ClickAtPosition 775, 366, , "00:00:03" '내보내기
                        ClickAtPosition 1223, 376, , "00:00:01", "00:00:01" '엑셀 저장하기
                        SendKeys "%{D}", True '파일 주소 표시줄 선택
                        SendKeys strFolder & j & "\", True  '파일 경로 지정 (수강현황은 강좌별 폴더)
                        SendKeys "{Enter}", True '파일 경로 이동
                        ClickAtPosition 1718, 1009, , "00:00:01", "00:00:02"  '저장
                        SendKeys "{Enter}", True
                        .Wait 200
                        SendKeys "%{F4}", True
    
                        .FindElementByXPath("/html/body/div[9]/div[3]/button[1]").Click
PassOption:
                    Next i
                End If
            Next vntCat
PassCat:
        Next r
        .Quit                          '브라우저 종료
        AppActivate Application.Caption  '엑셀로 포커스 이동
    End With 'sl
    
    Call MarkBadge(intNo + 1, uiColor.TC02, uiColor.TC16, sht)
    MsgBox "Completion", vbInformation
    
    Exit Sub
ErrorHandler:
    Call HandleError("BtDownloadFiles")
End Sub
Sub BtCheckFiles()
    
    Dim sht As Worksheet
    Set sht = ActiveSheet

    Dim intNo As Integer
    intNo = 3
    
    If CheckBadgeNo(uiColor.TC02, sht) < intNo Then
        MsgBox "이전 단계를 먼저 실행하세요!", vbInformation
        Exit Sub
    End If
    
    If MsgBox("실행하시겠습니까?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    Dim arrFiles As Variant
    
    arrFiles = GetFilesList(ReplacePath("{tPath}\00_다운파일\"))

    If IsArrayEmpty(arrFiles) Then
        MsgBox "파일이 없습니다. 파일을 다운로드하여 폴더에 넣으세요!", vbInformation
        Exit Sub
    End If
    
    Dim colFiles As New Collection
    Dim arrParts As Variant
    Dim strFile As String, strCat As String, strSeason As String, strChapter As String, strKey As String, strValue As String
    Dim i As Long
    
    On Error Resume Next
    For i = 0 To UBound(arrFiles)
        strFile = arrFiles(i)
        arrParts = Split(strFile, "\")
        
        If InStr(strFile, "학생명단") > 0 Then
            strCat = arrParts(UBound(arrParts) - 2)
            strSeason = arrParts(UBound(arrParts) - 1)
            strChapter = ""
        ElseIf InStr(strFile, "수강현황") > 0 Then
            strCat = arrParts(UBound(arrParts) - 3)
            strSeason = arrParts(UBound(arrParts) - 2)
            strChapter = arrParts(UBound(arrParts) - 1)
        End If
        
        strKey = strCat & strSeason & strChapter
        strValue = strKey
        
        colFiles.Add strValue, strKey
    Next
    On Error GoTo 0
    
    i = 0
    
    Dim arrCat As Variant, vntCat As Variant
    Dim arrNothings() As Variant
    Dim strNothing As String
    Dim r As Range, rng As Range

    arrCat = Array("학생명단", "수강현황")

    For Each vntCat In arrCat
        If vntCat = "학생명단" Then
            Set rng = GetTwbRange("DB_시즌[코드]")
        ElseIf vntCat = "수강현황" Then
            Set rng = GetTwbRange("DB_챕터[시즌코드]")
        End If

        For Each r In rng
            If vntCat = "학생명단" Then
                strKey = vntCat & r
                strNothing = "[" & vntCat & "] 중 [" & r & "] 시즌"
            ElseIf vntCat = "수강현황" Then
                strKey = vntCat & r & r.Offset(, 1)
                strNothing = "[" & vntCat & "] 중 [" & r & "] 시즌 [" & r.Offset(, 2) & "] 챕터"
            End If
            
            If CheckUniqueID(colFiles, strKey) Then
                ReDim Preserve arrNothings(i)
                arrNothings(i) = strNothing
                i = i + 1
            End If
        Next r
    Next vntCat
    
    If IsArrayEmpty(arrNothings) Then
        MsgBox "Completion", vbInformation
    Else
        If MsgBox(Join(arrNothings, vbNewLine) & vbNewLine & vbNewLine & "상기 파일이 누락되어 있습니다." & vbNewLine & "그래도 이 단계를 완료하시겠습니까??", vbExclamation + vbYesNo) <> vbYes Then GoTo pass
    End If
    
    Call MarkBadge(intNo + 1, uiColor.TC02, uiColor.TC16, sht)
pass:
Exit Sub
ErrorHandler:
    Call HandleError("CheckFolders")
    
End Sub
Sub BtUploadData()

    Dim sht As Worksheet
    Set sht = ActiveSheet

    Dim intNo As Integer
    intNo = 4
    
    If CheckBadgeNo(uiColor.TC02, sht) < intNo Then
        MsgBox "이전 단계를 먼저 실행하세요!", vbInformation
        Exit Sub
    End If

    If MsgBox("실행하시겠습니까?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    '엑셀 데이터를 배열로 담는 변수
    Dim arrCat As Variant, vntCat As Variant
    Dim dvExcel As DBvar
    Dim dvAccess As DBvar
    Dim strCat As String
    Dim arrFiles As Variant
    Dim arr As Variant, arrTemp As Variant, arrData() As Variant, arrClasses() As Variant
    Dim arrParts() As String
    Dim i As Long

    '배열을 엑세스로 보내는 변수
    Dim arrDB As Variant
    Dim r As Range, rngBM As Range, rngEL As Range
    Dim cn As Object, rs As Object
    Dim t As Long, v As Long
    Dim totalRecords As Long

    arrCat = Array("학생명단", "수강현황")
    Set rngBM = GetTwbRange("T_건물관리자[코드]")
    Set rngEL = GetTwbRange("T_통계제외명단[코드]")

    For Each vntCat In arrCat

        ' 루프마다 변수 초기화
        Set cn = Nothing
        Set rs = Nothing
        totalRecords = 0
        Erase arrData
        Erase arrClasses

        strCat = vntCat

        '파일 경로를 배열로 변환 ==========================================================
        arrFiles = GetFilesList(ReplacePath("{tPath}\00_다운파일\" & strCat), "*export_*.xlsx")
        If IsArrayEmpty(arrFiles) Then Exit Sub

        '엑셀 데이터를 배열로 변환 ==========================================================

        With dvExcel

            .Type = "엑셀"
            .Table = "[Sheet1$]"
            .Query = "SELECT * FROM " & .Table
            .arrQuery = Array(.Query)

            For i = 0 To UBound(arrFiles)
                .File = arrFiles(i)
                arr = SelectQueryArr(.arrQuery, .Type, .File, .Server, .Port, .db, .ID, .PW, True)
                arrParts = Split(arrFiles(i), "\")

                ReDim Preserve arrData(i)
                If strCat = "학생명단" Then
                    ReDim Preserve arrClasses(0, i)
                    arrClasses(0, i) = arrParts(UBound(arrParts) - 1)
                ElseIf strCat = "수강현황" Then
                    ReDim Preserve arrClasses(1, i)
                    arrClasses(0, i) = arrParts(UBound(arrParts) - 2)
                    arrClasses(1, i) = arrParts(UBound(arrParts) - 1)
                End If

                arrTemp = arr(0)
                arrData(i) = arrTemp

                Erase arrParts
                Erase arr
                Erase arrTemp

            Next
        End With

        '배열의 값을 엑세스 테이블로 입력 =========================================================

        With dvAccess
            .Alias = "EDU_LMS"

            arrDB = GetDbInfo(.Alias, False)
            .Type = arrDB(0)
            .File = arrDB(3)

            If strCat = "학생명단" Then
                .Table = "lms_students"
            ElseIf strCat = "수강현황" Then
                .Table = "lms_course_status"
            End If

            ' Access DB 연결
            Set cn = CreateObject("ADODB.Connection")
            cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & .File & ";"

            ' 테이블의 모든 데이터 삭제
            cn.Execute "DELETE FROM " & .Table
            ' 트랜잭션 시작
            cn.BeginTrans
            ' Recordset 열기
            Set rs = CreateObject("ADODB.Recordset")
            rs.Open "SELECT * FROM " & .Table, cn, 2, 3  ' adOpenDynamic, adLockOptimistic

            totalRecords = 0

            For i = 0 To UBound(arrData)
                For v = 1 To UBound(arrData(i), 2)
                    Application.StatusBar = ""
                    rs.AddNew
                    For t = 0 To UBound(arrData(i))
                        rs.fields("SEASON_CD").value = arrClasses(0, i)
                        If strCat = "수강현황" Then rs.fields("CHAPTER_NO").value = arrClasses(1, i)
                        rs.fields(arrData(i)(t, 0)).value = arrData(i)(t, v)
                    Next
                    rs.Update
                    totalRecords = totalRecords + 1
                Next
            Next

            cn.CommitTrans
            '직분, 직책 수정
            cn.Execute "UPDATE " & .Table & " SET  POSITION_NM = '당회장' WHERE POSITION_NM = '교구장';"
            cn.Execute "UPDATE " & .Table & " SET  DUTY_NM = '형제' WHERE DUTY_NM = '';"
            
            '건물관리자 직책 수정
            For Each r In rngBM
                If Trim(r) <> "" Then
                    cn.Execute "UPDATE " & .Table & " SET  POSITION_NM = '건물관리자' WHERE WATV_ID = '" & r & "';"
                End If
            Next

            '통계 제외자 삭제
            For Each r In rngEL
                If Trim(r) <> "" Then
                    cn.Execute "DELETE FROM " & .Table & " WHERE WATV_ID = '" & r & "';"
                End If
            Next
            
            rs.Close
            cn.Close
            Set rs = Nothing
            Set cn = Nothing
            Application.StatusBar = False
        End With
    Next vntCat

    Call RecDBUpdateDateTime("LmsDataUpload")
    Call SyncDB("DB_업데이트로그")
    
    Call MarkBadge(intNo + 1, uiColor.TC02, uiColor.TC16, sht)
    MsgBox "Completion", vbInformation

    
    Exit Sub
ErrorHandler:
    Call HandleError("CheckFolders")
End Sub
Sub GetReports()

    Dim sht As Worksheet
    Set sht = ActiveSheet

    Dim intNo As Integer
    intNo = 5
    
    If CheckBadgeNo(uiColor.TC02, sht) < intNo Then
        MsgBox "이전 단계를 먼저 실행하세요!", vbInformation
        Exit Sub
    End If
    
    If MsgBox("실행하시겠습니까?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    cl.Event_Off
    cl.Calculate_Off
    cl.sht_UnLock

    Call BtUpdateDBonThisPage(False)
    
    Call MarkBadge(intNo + 1, uiColor.TC02, uiColor.TC16, sht)
    MsgBox "Completion", vbInformation
    
    cl.sht_Lock
    cl.Calculate_On
    cl.Event_On

Exit Sub
ErrorHandler:
    Call HandleError("CheckFolders")
End Sub
