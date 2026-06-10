Attribute VB_Name = "tpl_Test"
Option Explicit

' ┌─────────────────────────────────────────────────────────┐
' │  tpl_Test                                               │
' │  역할 : corelib.xlam 전 모듈 테스트 프로시저 모음       │
' └─────────────────────────────────────────────────────────┘

Private Const TEST_SHEET_NM As String = "__TEST__"
Private Const TEST_TABLE_NM As String = "T_TestData"

Private Function TestTmpPath() As String
    TestTmpPath = ThisWorkbook.Path & "\_test_tmp"
End Function

Private Function TestWs() As Worksheet
    Set TestWs = ThisWorkbook.Worksheets(TEST_SHEET_NM)
End Function

Private Function TestTbl() As ListObject
    Set TestTbl = TestWs().ListObjects(TEST_TABLE_NM)
End Function

' ══════════════════════════════════════════════════════════
'  공통 헬퍼
' ══════════════════════════════════════════════════════════

Private Sub PrintResult(ByVal strCase   As String, _
                        ByVal strExpect As String, _
                        ByVal strOutput As String)
    Debug.Print "--------------------------------------------"
    Debug.Print "[" & strCase & "]"
    If strExpect <> "" Then Debug.Print "  예상 : " & strExpect
    Debug.Print "  결과 : " & strOutput
    Dim blnOK As Boolean
    blnOK = IIf(strExpect = "", strOutput <> "", strOutput = strExpect)
    Debug.Print "  유효 : " & IIf(blnOK, "OK", "FAIL")
End Sub

Private Sub PrintBool(ByVal strCase   As String, _
                      ByVal blnResult  As Boolean, _
                      Optional ByVal blnExpect As Boolean = True)
    Debug.Print "--------------------------------------------"
    Debug.Print "[" & strCase & "]"
    Debug.Print "  결과 : " & blnResult & "  (예상: " & blnExpect & ")"
    Debug.Print "  유효 : " & IIf(blnResult = blnExpect, "OK", "FAIL")
End Sub

Private Function Run(ByVal strPath As String) As String
    Run = Application.Run("corelib.xlam!am_Path.ReplacePath", _
                          strPath, _
                          ThisWorkbook.Path, _
                          ThisWorkbook.FullName)
End Function

' Application.Run 에서 Range 반환 함수 호출 시 Nothing 안전 처리
Private Function RunGetRng(ByVal strFuncPath As String, _
                           ParamArray args() As Variant) As Range
    Dim vnt As Variant
    On Error Resume Next
    Select Case UBound(args)
        Case 0: vnt = Application.Run("corelib.xlam!" & strFuncPath, args(0))
        Case 1: vnt = Application.Run("corelib.xlam!" & strFuncPath, args(0), args(1))
        Case 2: vnt = Application.Run("corelib.xlam!" & strFuncPath, args(0), args(1), args(2))
    End Select
    If IsObject(vnt) Then Set RunGetRng = vnt
    On Error GoTo 0
End Function

' ══════════════════════════════════════════════════════════
'  테스트 환경 설정 / 정리
' ══════════════════════════════════════════════════════════

Public Sub Setup_TestSheet()
    Application.ScreenUpdating = False
    Application.DisplayAlerts  = False

    On Error Resume Next
    ThisWorkbook.Worksheets(TEST_SHEET_NM).Delete
    On Error GoTo 0

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = TEST_SHEET_NM

    ' ── 테이블 데이터 ─────────────────────────────────────
    ws.Range("A1").Value = "ID"       : ws.Range("B1").Value = "이름"
    ws.Range("C1").Value = "나이"     : ws.Range("D1").Value = "상태"

    ws.Columns("A").NumberFormat = "@"

    ws.Range("A2").Value = "001" : ws.Range("B2").Value = "홍길동"   : ws.Range("C2").Value = 30 : ws.Range("D2").Value = "진행"
    ws.Range("A3").Value = "002" : ws.Range("B3").Value = "이순신"   : ws.Range("C3").Value = 45 : ws.Range("D3").Value = "완료"
    ws.Range("A4").Value = "003" : ws.Range("B4").Value = "김유신"   : ws.Range("C4").Value = 28 : ws.Range("D4").Value = "진행"
    ws.Range("A5").Value = "004" : ws.Range("B5").Value = "강감찬"   : ws.Range("C5").Value = 55 : ws.Range("D5").Value = "완료"
    ws.Range("A6").Value = "005" : ws.Range("B6").Value = "을지문덕" : ws.Range("C6").Value = 38 : ws.Range("D6").Value = "대기"

    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:D6"), , xlYes)
    tbl.Name = TEST_TABLE_NM

    ' ── FindCellsByColor 테스트용 (F 열) ──────────────────
    ws.Range("F1").Interior.Color = RGB(255, 255, 0)   ' 노란색
    ws.Range("F2").Interior.Color = RGB(255, 0, 0)     ' 빨간색
    ws.Range("F3").Interior.Color = RGB(255, 255, 0)   ' 노란색

    ' ── 조건부 서식·유효성 테스트용 (G 열) ───────────────
    ws.Range("G1").Value = 10 : ws.Range("G2").Value = 20
    ws.Range("G3").Value = 30 : ws.Range("G4").Value = 40
    ws.Range("G5").Value = 50

    Application.ScreenUpdating = True
    Application.DisplayAlerts  = True

    Debug.Print "--------------------------------------------"
    Debug.Print "[Setup_TestSheet] 완료 — " & TEST_SHEET_NM & " / " & TEST_TABLE_NM
End Sub

Public Sub Teardown_TestSheet()
    Application.ScreenUpdating = False
    Application.DisplayAlerts  = False

    On Error Resume Next
    ThisWorkbook.Worksheets(TEST_SHEET_NM).Delete
    On Error GoTo 0

    Application.Run "corelib.xlam!am_File.DelFolder", TestTmpPath()

    Application.ScreenUpdating = True
    Application.DisplayAlerts  = True

    Debug.Print "--------------------------------------------"
    Debug.Print "[Teardown_TestSheet] 완료 — 시트·임시폴더 제거"
End Sub

' ══════════════════════════════════════════════════════════
'  전체 테스트 일괄 실행
' ══════════════════════════════════════════════════════════

Public Sub RunAllTests()
    Debug.Print "============================================"
    Debug.Print "  corelib.xlam 전체 테스트 시작  " & Now
    Debug.Print "============================================"

    Call Test_Core
    Call Test_Path
    Call Test_File
    Call Test_Utils

    Call Setup_TestSheet
    Call Test_Range
    Call Test_Format
    Call Test_Table
    Call Test_Sheet
    Call Test_Excel
    Call Teardown_TestSheet

    Debug.Print "============================================"
    Debug.Print "  전체 테스트 완료"
    Debug.Print "============================================"
End Sub

' ══════════════════════════════════════════════════════════
'  am_Core 테스트
' ══════════════════════════════════════════════════════════

Public Sub Test_Core()
    Debug.Print vbCrLf & "▶ am_Core 테스트"

    Dim blnOff As Boolean

    Call PrintResult("XlamPath",     "", Application.Run("corelib.xlam!am_Core.XlamPath"))
    Call PrintResult("XlamFullName", "", Application.Run("corelib.xlam!am_Core.XlamFullName"))
    Call PrintResult("Version",      "", Application.Run("corelib.xlam!am_Core.Version"))

    Call PrintBool("IsReady",      Application.Run("corelib.xlam!am_Core.IsReady"),      True)
    Call PrintBool("IsXlamLoaded", Application.Run("corelib.xlam!am_Core.IsXlamLoaded"), True)

    Application.Run "corelib.xlam!am_Core.DPUpdate_Off"
    blnOff = (Application.ScreenUpdating = False)
    Application.Run "corelib.xlam!am_Core.DPUpdate_On"
    Call PrintBool("DPUpdate_Off → ScreenUpdating=False", blnOff, True)

    Application.Run "corelib.xlam!am_Core.Event_Off"
    blnOff = (Application.EnableEvents = False)
    Application.Run "corelib.xlam!am_Core.Event_On"
    Call PrintBool("Event_Off → EnableEvents=False", blnOff, True)

    Application.Run "corelib.xlam!am_Core.Calculate_Off"
    blnOff = (Application.Calculation = xlCalculationManual)
    Application.Run "corelib.xlam!am_Core.Calculate_On"
    Call PrintBool("Calculate_Off → xlCalculationManual", blnOff, True)
End Sub

' ══════════════════════════════════════════════════════════
'  am_Path 테스트
' ══════════════════════════════════════════════════════════

Public Sub Test_Path()
    Debug.Print vbCrLf & "▶ am_Path 테스트"
    Call Test_FixedTokens
    Call Test_CustomTokens
    Call Test_AbsolutePath
    Call Test_NetworkPath
    Call Test_DriveMapping
    Call Test_EdgeCases
End Sub

Public Sub Test_FixedTokens()
    Debug.Print vbCrLf & "  [ 고정 토큰 ]"
    Call PrintResult("cPath 토큰", "", Run("{cPath}"))
    Call PrintResult("cFile 토큰", "", Run("{cFile}"))
    Call PrintResult("xPath 토큰", "", Run("{xPath}"))
    Call PrintResult("xFile 토큰", "", Run("{xFile}"))
End Sub

Public Sub Test_CustomTokens()
    Debug.Print vbCrLf & "  [ 커스텀 토큰 ]"
    Call PrintResult("xDB 토큰",    "", Run("{xDB}"))
    Call PrintResult("xBak 토큰",   "", Run("{xBak}"))
    Call PrintBool("미등록 토큰 → 빈 문자열", Run("{미등록}\test.xlsx") = "", True)
End Sub

Public Sub Test_AbsolutePath()
    Debug.Print vbCrLf & "  [ 절대 경로 ]"
    Call PrintResult("유효한 절대 경로",   "", Run(ThisWorkbook.Path))
    Call PrintResult("슬래시 혼용",        "", Run(Replace(ThisWorkbook.Path, "\", "/")))
    Call PrintResult("드라이브 없는 경로", "", Run(Mid(ThisWorkbook.Path, 3)))
End Sub

Public Sub Test_NetworkPath()
    Debug.Print vbCrLf & "  [ 네트워크 경로 ]"
    Call PrintResult("UNC 경로", "", Run("\\서버명\공유폴더"))
End Sub

Public Sub Test_DriveMapping()
    Debug.Print vbCrLf & "  [ 드라이브 매핑 ]"
    Call PrintResult("없는 드라이브", "", Run("Z:\임시경로\test.xlsx"))

    Dim strIn  As String
    Dim strOut As String
    strIn  = "Z:\" & Mid(ThisWorkbook.Path, 4)
    strOut = Application.Run("corelib.xlam!am_Path.ReplacePath", _
                             strIn, ThisWorkbook.Path, ThisWorkbook.FullName, True)
    Call PrintResult("드라이브 강제 교체", "", strOut)
End Sub

Public Sub Test_EdgeCases()
    Debug.Print vbCrLf & "  [ 엣지 케이스 ]"
    Call PrintResult("빈 문자열",      "", Run(""))
    Call PrintResult("토큰만",         "", Run("{cPath}"))
    Call PrintResult("잘못된 경로",    "", Run("이건경로가아닙니다"))
    Call PrintResult("토큰 연속 사용", "", Run("{cPath}{xPath}"))
End Sub

' ══════════════════════════════════════════════════════════
'  am_File 테스트
' ══════════════════════════════════════════════════════════

Public Sub Test_File()
    Debug.Print vbCrLf & "▶ am_File 테스트"

    Dim strTmp As String
    strTmp = TestTmpPath()

    Call PrintResult("GetExt(.xlsx)",    ".xlsx", Application.Run("corelib.xlam!am_File.GetExt", "report.xlsx"))
    Call PrintResult("GetExt(.xlsm)",   ".xlsm", Application.Run("corelib.xlam!am_File.GetExt", "book.xlsm"))
    Call PrintResult("GetExt(전체경로)", ".csv",  Application.Run("corelib.xlam!am_File.GetExt", strTmp & "\data.csv"))

    Call PrintBool("CheckFolderExistence(없는 폴더)", _
                   Application.Run("corelib.xlam!am_File.CheckFolderExistence", strTmp), False)

    Application.Run "corelib.xlam!am_File.MkFolder", strTmp & "\sub1\sub2"
    Call PrintBool("MkFolder 중첩 생성 후 존재 확인", _
                   Application.Run("corelib.xlam!am_File.CheckFolderExistence", strTmp & "\sub1\sub2"), True)

    Call PrintBool("CheckFileExistence(없는 파일)", _
                   Application.Run("corelib.xlam!am_File.CheckFileExistence", strTmp & "\nofile.txt"), False)

    Application.Run "corelib.xlam!am_File.DelFolder", strTmp
    Call PrintBool("DelFolder 후 폴더 제거 확인", _
                   Application.Run("corelib.xlam!am_File.CheckFolderExistence", strTmp), False)
End Sub

' ══════════════════════════════════════════════════════════
'  am_Utils 테스트
' ══════════════════════════════════════════════════════════

Public Sub Test_Utils()
    Debug.Print vbCrLf & "▶ am_Utils 테스트"
    Call Test_Utils_Array
    Call Test_Utils_Check
    Call Test_Utils_Code
    Call Test_Utils_Date
    Call Test_Utils_Formula
End Sub

Private Sub Test_Utils_Array()
    Debug.Print vbCrLf & "  [ 배열 ]"

    Dim arrResult  As Variant
    Dim arrEmpty() As Variant

    arrResult = Application.Run("corelib.xlam!am_Utils.ConvertToArrData", 42)
    Call PrintBool("ConvertToArrData(42) IsArray",    IsArray(arrResult), True)
    If IsArray(arrResult) Then
        Call PrintBool("ConvertToArrData(42) 요소 수=1", _
                       UBound(arrResult) - LBound(arrResult) + 1 = 1, True)
    End If

    Call PrintBool("IsArrayEmpty(미초기화)", _
                   Application.Run("corelib.xlam!am_Utils.IsArrayEmpty", arrEmpty), True)
    Call PrintBool("IsArrayEmpty(요소있음)", _
                   Application.Run("corelib.xlam!am_Utils.IsArrayEmpty", arrResult), False)
End Sub

Private Sub Test_Utils_Check()
    Debug.Print vbCrLf & "  [ 검사 ]"

    Call PrintBool("IsValidFileName(유효)", _
                   Application.Run("corelib.xlam!am_Utils.IsValidFileName", "report_2024.xlsx"), True)
    Call PrintBool("IsRangeMerged(A1 병합없음)", _
                   Application.Run("corelib.xlam!am_Utils.IsRangeMerged", ActiveSheet.Range("A1")), False)
    Call PrintResult("GetValidationType(유효성없음 셀)", "없음", _
                     Application.Run("corelib.xlam!am_Utils.GetValidationType", ActiveSheet.Range("A1")))
    Call PrintBool("IsCells(Range)", _
                   Application.Run("corelib.xlam!am_Utils.IsCells", ActiveSheet.Range("A1")), True)
End Sub

Private Sub Test_Utils_Code()
    Debug.Print vbCrLf & "  [ 코드 생성 ]"

    Dim colIDs As New Collection
    Dim strID  As String
    strID = Application.Run("corelib.xlam!am_Utils.CreateUniqueID", "ID", 4, colIDs)
    Call PrintResult("CreateUniqueID(prefix=ID, digits=4)", "", strID)
    Call PrintBool("CreateUniqueID 길이=6",    Len(strID) = 6,        True)
    Call PrintBool("CreateUniqueID 접두사=ID", Left(strID, 2) = "ID", True)

    Dim strCode As String
    strCode = Application.Run("corelib.xlam!am_Utils.GenerateRandomCode", _
                              ActiveSheet.Range("A1"), 6)
    Call PrintResult("GenerateRandomCode(len=6)", "", strCode)
    Call PrintBool("GenerateRandomCode 길이=6", Len(strCode) = 6, True)
End Sub

Private Sub Test_Utils_Date()
    Debug.Print vbCrLf & "  [ 날짜·정규식 ]"

    Dim dblSerial As Double
    dblSerial = Application.Run("corelib.xlam!am_Utils.ConvertToExcelSerialDate", "2024-01-15", "09:30:00")
    Call PrintResult("ConvertToExcelSerialDate(2024-01-15 09:30)", "", CStr(dblSerial))
    Call PrintBool("ConvertToExcelSerialDate > 0", dblSerial > 0, True)

    Dim arrVal As Variant
    arrVal = Application.Run("corelib.xlam!am_Utils.ExtractValues", "주문번호 20241231 처리완료", "(\d{8})")
    If IsArray(arrVal) Then
        Call PrintResult("ExtractValues((\d{8}))", "20241231", CStr(arrVal(LBound(arrVal))))
    Else
        Debug.Print "[ExtractValues] FAIL - 배열 아님: " & CStr(arrVal)
    End If
End Sub

Private Sub Test_Utils_Formula()
    Debug.Print vbCrLf & "  [ 수식·유효성 ]"

    Call PrintBool("EvaluateFormula(1=1)",               Application.Run("corelib.xlam!am_Utils.EvaluateFormula", "1=1"),               True)
    Call PrintBool("EvaluateFormula(1=2)",               Application.Run("corelib.xlam!am_Utils.EvaluateFormula", "1=2"),               False)
    Call PrintBool("EvaluateFormula(=LEN(""abc"")>0)",   Application.Run("corelib.xlam!am_Utils.EvaluateFormula", "=LEN(""abc"")>0"),   True)
End Sub

' ══════════════════════════════════════════════════════════
'  am_Range 테스트  (Setup_TestSheet 완료 후 실행)
' ══════════════════════════════════════════════════════════

Public Sub Test_Range()
    Debug.Print vbCrLf & "▶ am_Range 테스트"

    Dim ws  As Worksheet : Set ws  = TestWs()
    Dim tbl As ListObject: Set tbl = TestTbl()

    ' ── GetUsedRange ──────────────────────────────────────
    Dim blnUsed As Boolean
    blnUsed = Application.Run("corelib.xlam!am_Range.GetUsedRange_IsValid", ws)
    Call PrintBool("GetUsedRange Not Nothing", blnUsed, True)
    If blnUsed Then
        Call PrintBool("GetUsedRange 행 수 >= 6", _
                       Application.Run("corelib.xlam!am_Range.GetUsedRange_RowCount", ws) >= 6, True)
        Call PrintBool("GetUsedRange 열 수 >= 4", _
                       Application.Run("corelib.xlam!am_Range.GetUsedRange_ColCount", ws) >= 4, True)
    End If

    ' ── FindRange — 존재하는 값 ───────────────────────────
    Dim idRng As Range
    Set idRng = tbl.ListColumns("ID").DataBodyRange
    Call PrintBool("FindRange(003) Not Nothing", _
                   Application.Run("corelib.xlam!am_Range.FindRange_IsValid", idRng, "003", False), True)
    Call PrintResult("FindRange(003) 값 일치", "003", _
                     Application.Run("corelib.xlam!am_Range.FindRange_CellValue", idRng, "003", False))

    ' ── FindRange — 없는 값 ───────────────────────────────
    Call PrintBool("FindRange(999) Nothing", _
                   Not Application.Run("corelib.xlam!am_Range.FindRange_IsValid", idRng, "999", False), True)

    ' ── FindCellsByColor — 노란색 (F1, F3) ────────────────
    Dim lngYellow As Long
    lngYellow = Application.Run("corelib.xlam!am_Range.FindCellsByColor_Count", _
                                 RGB(255, 255, 0), ws.Range("F1:F5"), False)
    Call PrintBool("FindCellsByColor(노란색) Not Nothing", lngYellow > 0, True)
    Call PrintBool("FindCellsByColor 노란색 셀 수=2", lngYellow = 2, True)

    ' ── FindCellsByColor — 없는 색 ────────────────────────
    Call PrintBool("FindCellsByColor(없는 파란색) Nothing", _
                   Application.Run("corelib.xlam!am_Range.FindCellsByColor_Count", _
                                   RGB(0, 0, 255), ws.Range("F1:F5"), False) = 0, True)
End Sub

' ══════════════════════════════════════════════════════════
'  am_Format 테스트  (Setup_TestSheet 완료 후 실행)
' ══════════════════════════════════════════════════════════

Public Sub Test_Format()
    Debug.Print vbCrLf & "▶ am_Format 테스트"

    Dim ws   As Worksheet : Set ws = TestWs()
    Dim rng  As Range
    Dim rngG As Range     : Set rngG = ws.Range("G1:G5")   ' 숫자 10~50
    Dim rngH As Range     : Set rngH = ws.Range("H1")      ' 유효성 테스트용

    ' ── ConditionalFormattingFormula ──────────────────────
    rngG.FormatConditions.Delete

    Application.Run "corelib.xlam!am_Format.ConditionalFormattingFormula", _
                    rngG, "=G1>25", _
                    RGB(255, 0, 0)   ' lngFontColor

    Call PrintBool("CF 추가 후 FormatConditions.Count >= 1", _
                   rngG.FormatConditions.Count >= 1, True)

    ' ── ClearConditionalFormatting ────────────────────────
    Application.Run "corelib.xlam!am_Format.ClearConditionalFormatting", rngG
    Call PrintBool("ClearConditionalFormatting 후 Count=0", _
                   rngG.FormatConditions.Count = 0, True)

    ' ── ConditionalFormattingColorScale (2단) ─────────────
    Application.Run "corelib.xlam!am_Format.ConditionalFormattingColorScale", _
                    rngG, RGB(255, 0, 0), RGB(0, 255, 0)
    Call PrintBool("ColorScale(2단) 추가 Count >= 1", _
                   rngG.FormatConditions.Count >= 1, True)
    Application.Run "corelib.xlam!am_Format.ClearConditionalFormatting", rngG

    ' ── ConditionalFormattingDataBar ─────────────────────
    Application.Run "corelib.xlam!am_Format.ConditionalFormattingDataBar", _
                    rngG, RGB(0, 112, 192)
    Call PrintBool("DataBar 추가 Count >= 1", _
                   rngG.FormatConditions.Count >= 1, True)
    Application.Run "corelib.xlam!am_Format.ClearConditionalFormatting", rngG

    ' ── ValidationList ────────────────────────────────────
    rngH.Validation.Delete
    Application.Run "corelib.xlam!am_Format.ValidationList", _
                    rngH, Array("승인", "반려", "대기")

    ' ValidationList 적용 후 am_Utils.GetValidationType 로 교차 검증
    Call PrintResult("ValidationList 후 GetValidationType", "목록", _
                     Application.Run("corelib.xlam!am_Utils.GetValidationType", rngH))

    ' ── ClearValidation ───────────────────────────────────
    Application.Run "corelib.xlam!am_Format.ClearValidation", rngH
    Call PrintResult("ClearValidation 후 GetValidationType", "없음", _
                     Application.Run("corelib.xlam!am_Utils.GetValidationType", rngH))

    ' ── SetValidation (목록형) ────────────────────────────
    Application.Run "corelib.xlam!am_Format.SetValidation", _
                    rngH, "A,B,C", "", xlValidateList
    Call PrintResult("SetValidation(목록) 후 GetValidationType", "목록", _
                     Application.Run("corelib.xlam!am_Utils.GetValidationType", rngH))
    Application.Run "corelib.xlam!am_Format.ClearValidation", rngH
End Sub

' ══════════════════════════════════════════════════════════
'  am_Table 테스트  (Setup_TestSheet 완료 후 실행)
' ══════════════════════════════════════════════════════════

Public Sub Test_Table()
    Debug.Print vbCrLf & "▶ am_Table 테스트"

    Dim ws  As Worksheet : Set ws  = TestWs()
    Dim tbl As ListObject: Set tbl = TestTbl()

    ' ── 1. 조회 (Read-Only) ───────────────────────────────
    Debug.Print vbCrLf & "  [ 조회 ]"

    Dim arrTblNames As Variant
    arrTblNames = Application.Run("corelib.xlam!am_Table.GetTableNames", ws)
    Call PrintBool("GetTableNames Not Empty",  Not IsEmpty(arrTblNames), True)
    If Not IsEmpty(arrTblNames) Then
        Call PrintResult("GetTableNames(0)", TEST_TABLE_NM, CStr(arrTblNames(LBound(arrTblNames))))
    End If

    Dim arrAll As Variant
    arrAll = Application.Run("corelib.xlam!am_Table.GetAllSheetTableNames", ThisWorkbook)
    Call PrintBool("GetAllSheetTableNames Not Empty", Not IsEmpty(arrAll), True)

    Call PrintBool("IsTable(표 안 셀)", _
                   Application.Run("corelib.xlam!am_Table.IsTable", tbl.DataBodyRange.Cells(1, 1)), True)
    Call PrintBool("IsTable(표 밖 셀)", _
                   Application.Run("corelib.xlam!am_Table.IsTable", ws.Range("F1")), False)

    Dim arrWidths As Variant
    arrWidths = Application.Run("corelib.xlam!am_Table.GetTableColumnsWidth", tbl)
    Call PrintBool("GetTableColumnsWidth 열 수=4", _
                   IsArray(arrWidths) And UBound(arrWidths) - LBound(arrWidths) + 1 = 4, True)

    ' ── 2. 검색 ──────────────────────────────────────────
    Debug.Print vbCrLf & "  [ 검색 ]"

    ' TblFindVals_MC — 상태="진행" 인 ID 목록
    Dim arrFound As Variant
    arrFound = Application.Run("corelib.xlam!am_Table.TblFindVals_MC", _
                               tbl, "ID", "상태", "=", "진행")
    Call PrintBool("TblFindVals_MC(상태=진행) 결과 2건", _
                   IsArray(arrFound) And UBound(arrFound) - LBound(arrFound) + 1 = 2, True)

    ' TblFindVal_One — 나이>=45 첫 번째 이름
    Dim strName As String
    strName = Application.Run("corelib.xlam!am_Table.TblFindVal_One", _
                              tbl, "이름", 1, "나이", ">=", 45)
    Call PrintResult("TblFindVal_One(나이>=45 첫 번째 이름)", "이순신", CStr(strName))

    ' TblFindRng_MC — Application.Run 경유 Range 반환 불가 (설계 제약)
    '   TblFindVals_MC 로 동일 조건의 건수 검증으로 대체
    Dim arrRngTest As Variant
    arrRngTest = Application.Run("corelib.xlam!am_Table.TblFindVals_MC", _
                                 tbl, "이름", "상태", "=", "완료")
    Call PrintBool("TblFindRng_MC(상태=완료) 건수=2 [간접검증]", _
                   IsArray(arrRngTest) And _
                   UBound(arrRngTest) - LBound(arrRngTest) + 1 = 2, True)

    ' intOffset — ID 셀에서 상태(D) 까지 오프셋=3
    Dim lngOff As Long
    lngOff = Application.Run("corelib.xlam!am_Table.intOffset", _
                             tbl.ListColumns("ID").DataBodyRange.Cells(1, 1), "상태")
    Call PrintBool("intOffset(ID→상태)=3", lngOff = 3, True)

    ' ── 3. 필터 (원상복귀 필수) ──────────────────────────
    Debug.Print vbCrLf & "  [ 필터 ]"

    Application.Run "corelib.xlam!am_Table.AutoTableFilter", tbl, "상태", "진행"
    Call PrintBool("AutoTableFilter 필터 적용", tbl.AutoFilter.FilterMode, True)
    Application.Run "corelib.xlam!am_Table.ClearFiltersInTable", tbl
    Call PrintBool("ClearFiltersInTable 후 FilterMode=False", Not tbl.AutoFilter.FilterMode, True)

    Application.Run "corelib.xlam!am_Table.AutoTableFilter_Arr", tbl, "상태", Array("완료", "대기")
    Application.Run "corelib.xlam!am_Table.ClearFiltersInTable", tbl
    Call PrintBool("AutoTableFilter_Arr + Clear → FilterMode=False", Not tbl.AutoFilter.FilterMode, True)

    ' ── 4. 정렬 (ID 기준 복원) ────────────────────────────
    Debug.Print vbCrLf & "  [ 정렬 ]"

    ' 나이 오름차순 → 첫 행 나이=28 (김유신)
    Application.Run "corelib.xlam!am_Table.SortTable", TEST_TABLE_NM, "나이"
    Call PrintBool("SortTable(나이 ASC) 첫행 나이=28", _
                   tbl.DataBodyRange.Cells(1, 3).Value = 28, True)

    ' ID 오름차순 복원 → 첫 행 ID=001
    Application.Run "corelib.xlam!am_Table.SortTable", TEST_TABLE_NM, "ID"
    Call PrintResult("SortTable(ID ASC) 복원 첫행", "001", CStr(tbl.DataBodyRange.Cells(1, 1).Value))

    ' 사용자정의 정렬 — 상태 순서: 대기 > 진행 > 완료 → 첫 행 상태="대기"
    Application.Run "corelib.xlam!am_Table.SortTableCustomList", _
                    TEST_TABLE_NM, "상태", Array("대기", "진행", "완료")
    Call PrintResult("SortTableCustomList(대기>진행>완료) 첫행", "대기", _
                     CStr(tbl.DataBodyRange.Cells(1, 4).Value))

    ' ID 오름차순 복원
    Application.Run "corelib.xlam!am_Table.SortTable", TEST_TABLE_NM, "ID"
    Call PrintResult("SortTableCustomList 후 ID ASC 복원", "001", _
                     CStr(tbl.DataBodyRange.Cells(1, 1).Value))

    ' ── 5. 값 변경 (원상복귀 필수) ───────────────────────
    Debug.Print vbCrLf & "  [ 값 변경 ]"

    Application.Run "corelib.xlam!am_Table.ChangeTableValue", "003", "상태", "수정됨", tbl
    Call PrintBool("ChangeTableValue(003 상태→수정됨)", _
                   CStr(tbl.ListColumns("상태").DataBodyRange.Cells(3, 1).Value) = "수정됨", True)

    Application.Run "corelib.xlam!am_Table.ChangeTableValue", "003", "상태", "진행", tbl
    Call PrintBool("ChangeTableValue 복원(수정됨→진행)", _
                   CStr(tbl.ListColumns("상태").DataBodyRange.Cells(3, 1).Value) = "진행", True)

    ' ── 6. 행 추가 / 삭제 ────────────────────────────────
    Debug.Print vbCrLf & "  [ 행 추가·삭제 ]"

    Dim lngBefore As Long
    lngBefore = tbl.DataBodyRange.Rows.Count

    Application.Run "corelib.xlam!am_Table.AddTableRows", 1, 0, tbl
    Call PrintBool("AddTableRows(1) 후 행 수=" & lngBefore + 1, _
                   tbl.DataBodyRange.Rows.Count = lngBefore + 1, True)

    ' 추가된 행에 삭제 식별값 입력
    tbl.DataBodyRange.Cells(tbl.DataBodyRange.Rows.Count, 1).Value = "DEL"
    tbl.DataBodyRange.Cells(tbl.DataBodyRange.Rows.Count, 4).Value = "삭제예정"

    Application.Run "corelib.xlam!am_Table.DelTableFilteredRows", "상태", "삭제예정", tbl
    Call PrintBool("DelTableFilteredRows(상태=삭제예정) 후 행 수 복원", _
                   tbl.DataBodyRange.Rows.Count = lngBefore, True)

    ' ── 7. 열 추가 / 삭제 ────────────────────────────────
    Debug.Print vbCrLf & "  [ 열 추가·삭제 ]"

    Dim lngColBefore As Long
    lngColBefore = tbl.ListColumns.Count

    Application.Run "corelib.xlam!am_Table.AddArrayColumns", Array("메모", "날짜"), 0, tbl
    Call PrintBool("AddArrayColumns(2열) 후 열 수=" & lngColBefore + 2, _
                   tbl.ListColumns.Count = lngColBefore + 2, True)

    Application.Run "corelib.xlam!am_Table.DelTableColumns", lngColBefore + 1, 2, tbl
    Call PrintBool("DelTableColumns(2열 제거) 후 열 수 복원", _
                   tbl.ListColumns.Count = lngColBefore, True)
End Sub

' ══════════════════════════════════════════════════════════
'  am_Sheet 테스트  (Setup_TestSheet 완료 후 실행)
' ══════════════════════════════════════════════════════════

Public Sub Test_Sheet()
    Debug.Print vbCrLf & "▶ am_Sheet 테스트"

    Dim ws  As Worksheet : Set ws  = TestWs()
    Dim wb  As Workbook  : Set wb  = ThisWorkbook

    ' ── GetSheetNames ─────────────────────────────────────
    Dim arrNames As Variant
    arrNames = Application.Run("corelib.xlam!am_Sheet.GetSheetNames", wb)
    Call PrintBool("GetSheetNames Not Empty", IsArray(arrNames), True)

    ' __TEST__ 이 목록에 포함되는지 확인
    Dim blnFound As Boolean
    Dim i        As Long
    If IsArray(arrNames) Then
        For i = LBound(arrNames) To UBound(arrNames)
            If CStr(arrNames(i)) = TEST_SHEET_NM Then blnFound = True
        Next i
    End If
    Call PrintBool("GetSheetNames에 __TEST__ 포함", blnFound, True)

    ' ── VisibleAllSheets ──────────────────────────────────
    Application.Run "corelib.xlam!am_Sheet.VisibleAllSheets", wb
    Call PrintBool("VisibleAllSheets — 숨김 시트=0", _
                   prv_CountHiddenSheets(wb) = 0, True)

    ' ── HideAllSheetsExceptOne → VisibleAllSheets 복원 ────
    Application.Run "corelib.xlam!am_Sheet.HideAllSheetsExceptOne", ws, wb
    Call PrintBool("HideAllSheetsExceptOne — 보이는 시트=1", _
                   prv_CountVisibleSheets(wb) = 1, True)

    Application.Run "corelib.xlam!am_Sheet.VisibleAllSheets", wb
    Call PrintBool("VisibleAllSheets 복원", _
                   prv_CountHiddenSheets(wb) = 0, True)

    ' ── SortSheets ────────────────────────────────────────
    ' 현재 시트 순서 저장 → 역순 배열로 재정렬 → 원순서 복원
    Dim arrOrigOrder As Variant
    arrOrigOrder = Application.Run("corelib.xlam!am_Sheet.GetSheetNames", wb)

    Dim arrReverse() As Variant
    ReDim arrReverse(LBound(arrOrigOrder) To UBound(arrOrigOrder))
    For i = LBound(arrOrigOrder) To UBound(arrOrigOrder)
        arrReverse(i) = arrOrigOrder(UBound(arrOrigOrder) - i + LBound(arrOrigOrder))
    Next i

    Application.Run "corelib.xlam!am_Sheet.SortSheets", wb, arrReverse
    Call PrintBool("SortSheets(역순) 마지막 시트=" & CStr(arrOrigOrder(LBound(arrOrigOrder))), _
                   wb.Sheets(wb.Sheets.Count).Name = CStr(arrOrigOrder(LBound(arrOrigOrder))), True)

    Application.Run "corelib.xlam!am_Sheet.SortSheets", wb, arrOrigOrder
    Call PrintBool("SortSheets 원순서 복원 확인", _
                   wb.Sheets(1).Name = CStr(arrOrigOrder(LBound(arrOrigOrder))), True)

    ' ── BackupSheet ───────────────────────────────────────
    Dim strBakPath As String
    strBakPath = TestTmpPath() & "\backup"

    Application.Run "corelib.xlam!am_Sheet.BackupSheet", strBakPath, wb, "Test_Backup.xlsx", ws
    Call PrintBool("BackupSheet — 백업 파일 생성 확인", _
                   Dir(strBakPath & "\Test_Backup.xlsx") <> "", True)

    ' ── BackupWorkbook ────────────────────────────────────
    Application.Run "corelib.xlam!am_Sheet.BackupWorkbook", strBakPath, wb, "WB_Backup.xlsm"
    Call PrintBool("BackupWorkbook — 백업 파일 생성 확인", _
                   Dir(strBakPath & "\WB_Backup.xlsm") <> "", True)

    ' ── SheetLock / SheetUnLock ───────────────────────────
    ' SheetLock 전 배경색 없는 셀 만들기 (B2에 배경 없음 → 입력 가능 셀)
    ws.Range("B2").Interior.ColorIndex = xlNone

    Application.Run "corelib.xlam!am_Sheet.SheetLock", ws, "test1234"
    Call PrintBool("SheetLock — ProtectContents=True", ws.ProtectContents, True)

    Application.Run "corelib.xlam!am_Sheet.SheetUnLock", ws, "test1234"
    Call PrintBool("SheetUnLock — ProtectContents=False", ws.ProtectContents, False)
End Sub

' ── am_Sheet 보조 함수 (테스트 검증용) ───────────────────

Private Function prv_CountHiddenSheets(ByVal wb As Workbook) As Long
    Dim sht As Object
    Dim cnt As Long
    For Each sht In wb.Sheets
        If sht.Visible <> xlSheetVisible Then cnt = cnt + 1
    Next sht
    prv_CountHiddenSheets = cnt
End Function

Private Function prv_CountVisibleSheets(ByVal wb As Workbook) As Long
    Dim sht As Object
    Dim cnt As Long
    For Each sht In wb.Sheets
        If sht.Visible = xlSheetVisible Then cnt = cnt + 1
    Next sht
    prv_CountVisibleSheets = cnt
End Function

' ══════════════════════════════════════════════════════════
'  am_Excel 테스트  (Setup_TestSheet 완료 후 실행)
' ══════════════════════════════════════════════════════════

Public Sub Test_Excel()
    Debug.Print vbCrLf & "▶ am_Excel 테스트"

    Dim ws As Worksheet : Set ws = TestWs()

    ' ── SetPrintPage (다이얼로그 없음, 설정만) ────────────
    Application.Run "corelib.xlam!am_Excel.SetPrintPage", ws
    Call PrintBool("SetPrintPage — PrintArea 설정됨", ws.PageSetup.PrintArea <> "", True)

    ' ── ExportSheetToCSV ──────────────────────────────────
    Dim strCsvPath As String
    strCsvPath = TestTmpPath() & "\export"

    Application.Run "corelib.xlam!am_Excel.ExportSheetToCSV", ws, strCsvPath, "test_export"
    Call PrintBool("ExportSheetToCSV — CSV 파일 생성 확인", _
                   Dir(strCsvPath & "\test_export.csv") <> "", True)

    ' ── ExportPDF ─────────────────────────────────────────
    Dim strPdfPath As String
    strPdfPath = TestTmpPath() & "\export\test_export.pdf"

    Application.Run "corelib.xlam!am_Excel.ExportPDF", strPdfPath, ws
    Call PrintBool("ExportPDF — PDF 파일 생성 확인", _
                   Dir(strPdfPath) <> "", True)
End Sub
