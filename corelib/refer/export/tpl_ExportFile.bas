Attribute VB_Name = "tpl_ExportFile"
Option Explicit
'============================================================
' [프로시저] SetPrintPage
' [용도]    워크시트의 인쇄 페이지를 일괄 설정하는 범용 함수
'
' ─────────────────────────────────────────────────────────
' [인수 설명]
'
' ▶ 대상 시트 / 영역
'   sht          (Worksheet)  : 설정할 워크시트
'                               생략 시 → ActiveSheet 자동 적용
'
'   rngPrintArea (Range)      : 인쇄 영역으로 지정할 셀 범위
'                               생략 시 → 시트 내 첫 번째 표(ListObject) 범위 자동 적용
'                                         표가 없으면 UsedRange(사용된 전체 범위) 적용
'
'   rngTitleRows (Range)      : 모든 페이지 상단에 반복 인쇄할 행 범위
'                               생략 시 → 반복 행 없음
'                               예시)  rngTitleRows:=ws.Rows("1:2")
'
'   blnAddArea   (Boolean)    : 인쇄 영역 추가 여부
'                               False → 기존 인쇄 영역을 덮어쓰기 (기본값)
'                               True  → 기존 인쇄 영역에 새 범위를 추가
'
' ─────────────────────────────────────────────────────────
' ▶ 용지 / 방향
'   paperSize    (XlPaperSize)       : 인쇄 용지 크기
'                                      기본값 → xlPaperA4 (A4)
'                                      예시)  xlPaperA3, xlPaperLetter
'
'   orientation  (XlPageOrientation) : 인쇄 방향
'                                      기본값 → xlPortrait (세로)
'                                      예시)  xlLandscape (가로)
'
' ─────────────────────────────────────────────────────────
' ▶ 여백 (단위: cm)
'   sngTopMargin    (Double) : 위쪽 여백   / 기본값 → 1.5 cm
'   sngBottomMargin (Double) : 아래쪽 여백 / 기본값 → 1.0 cm
'   sngLeftMargin   (Double) : 왼쪽 여백   / 기본값 → 1.0 cm
'   sngRightMargin  (Double) : 오른쪽 여백 / 기본값 → 1.0 cm
'
' ─────────────────────────────────────────────────────────
' ▶ 페이지 정렬
'   CenterH (Boolean) : 가로 가운데 맞춤 / 기본값 → True
'   CenterV (Boolean) : 세로 가운데 맞춤 / 기본값 → False
'
' ─────────────────────────────────────────────────────────
' ▶ 배율 / 페이지 맞춤
'   fitToPage (Boolean) : True  → 페이지 수 기준으로 자동 축소 (FitToPages 모드)
'                         False → 지정 배율(%) 사용 (Zoom 모드)
'                         기본값 → True
'
'   intWide   (Integer) : [fitToPage = True 일 때] 가로 페이지 수
'                         기본값 → 1 (가로 1페이지에 맞춤)
'                         0 이하 → 가로 페이지 수 제한 없음
'
'   intTall   (Integer) : [fitToPage = True 일 때] 세로 페이지 수
'                         기본값 → 0 (세로 페이지 수 제한 없음)
'                         1 이상 → 해당 페이지 수에 맞춤
'
'   intPer    (Integer) : [fitToPage = False 일 때] 인쇄 배율 (%)
'                         기본값 → 100 (100%)
'                         예시)  85 → 85% 축소 출력
'
' ─────────────────────────────────────────────────────────
' [호출 예시]
'
'   ' 기본 호출 (A4, 세로, 가로 1페이지 맞춤)
'   Call SetPrintPage(sht:=Sheets("Sheet1"))
'
'   ' 인쇄 범위 + 반복 행 + 가로 방향 지정
'   Call SetPrintPage(sht:=Sheets("Sheet1"), _
'                     rngPrintArea:=Sheets("Sheet1").Range("A1:J100"), _
'                     rngTitleRows:=Sheets("Sheet1").Rows("1:2"), _
'                     orientation:=xlLandscape)
'
'   ' 기존 인쇄 영역에 새 범위 추가
'   Call SetPrintPage(sht:=Sheets("Sheet1"), _
'                     rngPrintArea:=Sheets("Sheet1").Range("A200:J300"), _
'                     blnAddArea:=True)
'
'   ' 배율 모드로 85% 출력
'   Call SetPrintPage(sht:=Sheets("Sheet1"), _
'                     fitToPage:=False, _
'                     intPer:=85)
'============================================================
Sub SetPrintPage(Optional ByVal sht As Worksheet, _
                 Optional ByVal rngPrintArea As Range, _
                 Optional ByVal rngTitleRows As Range, _
                 Optional ByVal blnAddArea As Boolean = False, _
                 Optional ByVal paperSize As XlPaperSize = xlPaperA4, _
                 Optional ByVal orientation As XlPageOrientation = xlPortrait, _
                 Optional ByVal sngTopMargin As Double = 1.5, _
                 Optional ByVal sngBottomMargin As Double = 1, _
                 Optional ByVal sngLeftMargin As Double = 1, _
                 Optional ByVal sngRightMargin As Double = 1, _
                 Optional ByVal CenterH As Boolean = True, _
                 Optional ByVal CenterV As Boolean = False, _
                 Optional ByVal fitToPage As Boolean = True, _
                 Optional ByVal intWide As Integer = 1, _
                 Optional ByVal intTall As Integer = 0, _
                 Optional ByVal intPer As Integer = 100)

    ' 시트 기본값 처리
    If sht Is Nothing Then Set sht = ActiveSheet

    ' 인쇄 영역 기본값 처리 (테이블 없는 시트 방어 코드)
    If rngPrintArea Is Nothing Then
        On Error Resume Next
        Set rngPrintArea = sht.ListObjects(1).Range
        On Error GoTo 0

        ' ListObjects(1) 없을 경우 UsedRange로 폴백
        If rngPrintArea Is Nothing Then
            Set rngPrintArea = sht.UsedRange
        End If
    End If

    With sht.PageSetup

        ' 인쇄 영역 설정 (추가 or 덮어쓰기 분기)
        If blnAddArea And .PrintArea <> "" Then
            .PrintArea = .PrintArea & "," & rngPrintArea.Address  ' 기존 영역에 추가
        Else
            .PrintArea = rngPrintArea.Address                     ' 덮어쓰기
        End If

        ' 반복 인쇄 행 설정
        If Not rngTitleRows Is Nothing Then .PrintTitleRows = rngTitleRows.Address

        ' 인쇄 용지 설정
        .paperSize = paperSize
        .orientation = orientation
        .CenterHorizontally = CenterH
        .CenterVertically = CenterV
        .TopMargin = Application.CentimetersToPoints(sngTopMargin)
        .LeftMargin = Application.CentimetersToPoints(sngLeftMargin)
        .RightMargin = Application.CentimetersToPoints(sngRightMargin)
        .BottomMargin = Application.CentimetersToPoints(sngBottomMargin)

        ' 페이지 배율 설정
        If fitToPage Then
            .Zoom = False

            If intWide >= 1 Then
                .FitToPagesWide = intWide
            Else
                .FitToPagesWide = False
            End If

            If intTall >= 1 Then
                .FitToPagesTall = intTall
            Else
                .FitToPagesTall = False
            End If

        Else
            ' 배율(%) 모드: Zoom에 직접 숫자 대입
            .Zoom = intPer

        End If
    End With

End Sub
'============================================================
' [프로시저] ExportPDF
' [용도]    워크시트를 PDF 파일로 내보내는 범용 함수
'
' ─────────────────────────────────────────────────────────
' [인수 설명]
'
' ▶ 대상 시트 / 파일 경로
'   strFilePath  (String)    : 저장할 PDF 파일 전체 경로
'                              확장자 생략 시 → .pdf 자동 추가
'                              존재하지 않는 폴더 → 자동 생성
'                              예시) "C:\Reports\결과물.pdf"
'
'   sht          (Worksheet) : PDF로 내보낼 워크시트
'                              생략 시 → ActiveSheet 자동 적용
'
' ─────────────────────────────────────────────────────────
' ▶ 출력 옵션
'   blnOpenAfter (Boolean)   : PDF 저장 후 자동 열기 여부
'                              기본값 → False (열지 않음)
'                              True   → 저장 완료 후 PDF 자동 실행
'
'   xlQual       (XlFixedFormatQuality) : PDF 출력 품질
'                              기본값 → xlQualityStandard (표준 품질)
'                              예시)  xlQualityMinimum (최소 품질 / 파일 크기 작음)
'
'   blnDocProps  (Boolean)   : 문서 속성(제목, 작성자 등) 포함 여부
'                              기본값 → True
'
' ─────────────────────────────────────────────────────────
' [호출 예시]
'
'   ' 기본 호출 (ActiveSheet → 표준 품질 PDF 저장)
'   Call ExportPDF(strFilePath:="C:\Reports\결과물.pdf")
'
'   ' 특정 시트 지정 + 저장 후 자동 열기
'   Call ExportPDF(strFilePath:="C:\Reports\결과물.pdf", _
'                  sht:=Sheets("Sheet1"), _
'                  blnOpenAfter:=True)
'
'   ' 최소 품질로 저장 (파일 크기 축소)
'   Call ExportPDF(strFilePath:="C:\Reports\결과물", _
'                  xlQual:=xlQualityMinimum)
'============================================================
Sub ExportPDF(ByVal strFilePath As String, _
              Optional ByVal sht As Worksheet, _
              Optional ByVal blnOpenAfter As Boolean = False, _
              Optional ByVal xlQual As XlFixedFormatQuality = xlQualityStandard, _
              Optional ByVal blnDocProps As Boolean = True)

    Dim strFolder As String

    ' 시트 기본값 처리
    If sht Is Nothing Then Set sht = ActiveSheet

    ' 확장자 자동 보정 (.pdf 없으면 추가)
    If LCase(Right(strFilePath, 4)) <> ".pdf" Then
        strFilePath = strFilePath & ".pdf"
    End If

    ' 폴더 존재 여부 확인 → 없으면 자동 생성
    strFolder = Left(strFilePath, InStrRev(strFilePath, "\"))
    If strFolder <> "" And Dir(strFolder, vbDirectory) = "" Then
        On Error Resume Next
        MkDir strFolder
        On Error GoTo 0

        ' 폴더 생성 실패 시 중단
        If Dir(strFolder, vbDirectory) = "" Then
            MsgBox "폴더를 생성할 수 없습니다." & vbCrLf & strFolder, vbCritical, "ExportPDF 오류"
            Exit Sub
        End If
    End If

    ' PDF 내보내기 (파일 잠금 등 오류 처리 포함)
    On Error GoTo ErrHandler
    sht.ExportAsFixedFormat Type:=xlTypePDF, _
                            Filename:=strFilePath, _
                            Quality:=xlQual, _
                            IncludeDocProperties:=blnDocProps, _
                            IgnorePrintAreas:=False, _
                            OpenAfterPublish:=blnOpenAfter
    On Error GoTo 0
    Exit Sub

ErrHandler:
    MsgBox "PDF 저장 중 오류가 발생했습니다." & vbCrLf & _
           "경로: " & strFilePath & vbCrLf & vbCrLf & _
           "오류: " & Err.Description, vbCritical, "ExportPDF 오류"

End Sub
'============================================================
' [프로시저] ExportSheetToCSV
' [용도]    워크시트를 CSV 파일로 내보내는 범용 함수
'
' ─────────────────────────────────────────────────────────
' [인수 설명]
'
' ▶ 대상 시트 / 저장 경로
'   sht          (Worksheet) : CSV로 내보낼 워크시트
'
'   strPath      (String)    : 저장할 폴더 경로
'                              끝에 "\" 없어도 자동 보정
'                              존재하지 않는 폴더 → MkFolder로 자동 생성
'                              예시) "C:\Reports\" 또는 "C:\Reports"
'
'   strFileName  (String)    : 저장할 파일명 (확장자 제외)
'                              .csv 확장자 중복 입력 시 자동 제거
'                              예시) "결과물" → "결과물.csv" 로 저장
'
' ─────────────────────────────────────────────────────────
' [호출 예시]
'
'   ' 기본 호출
'   Call ExportSheetToCSV(sht:=Sheets("Sheet1"), _
'                         strPath:="C:\Reports\", _
'                         strFileName:="결과물")
'
'   ' 확장자 포함해도 자동 보정
'   Call ExportSheetToCSV(sht:=Sheets("Sheet1"), _
'                         strPath:="C:\Reports", _
'                         strFileName:="결과물.csv")
'============================================================
Sub ExportSheetToCSV(ByVal sht As Worksheet, _
                     ByVal strPath As String, _
                     ByVal strFileName As String)

    Dim blnVisible  As Boolean
    Dim wbCopy      As Workbook

    ' Path 끝 "\" 자동 보정
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"

    ' 확장자 중복 자동 보정 (.csv.csv 방지)
    If LCase(Right(strFileName, 4)) = ".csv" Then
        strFileName = Left(strFileName, Len(strFileName) - 4)
    End If

    ' 폴더 없으면 자동 생성
    Call MkFolder(strPath)

    ' 원본 시트 Visible 상태 저장 (복원용)
    blnVisible = sht.Visible

    ' 숨김 시트인 경우 임시로 표시
    If Not blnVisible Then sht.Visible = True

    ' CSV 저장 (오류 처리 포함)
    On Error GoTo ErrHandler
    sht.Copy
    Set wbCopy = ActiveWorkbook
    wbCopy.SaveAs Filename:=strPath & strFileName & ".csv", FileFormat:=xlCSVUTF8
    wbCopy.Close SaveChanges:=False
    On Error GoTo 0

CleanUp:
    ' 원본 시트 Visible 상태 복원
    sht.Visible = blnVisible
    Exit Sub

ErrHandler:
    MsgBox "CSV 저장 중 오류가 발생했습니다." & vbCrLf & _
           "경로: " & strPath & strFileName & ".csv" & vbCrLf & vbCrLf & _
           "오류: " & Err.Description, vbCritical, "ExportSheetToCSV 오류"

    ' 오류 발생 시 복사된 워크북 강제 닫기
    If Not wbCopy Is Nothing Then
        wbCopy.Close SaveChanges:=False
    End If

    Resume CleanUp

End Sub
