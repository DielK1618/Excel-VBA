Attribute VB_Name = "ref_Excel"
Option Explicit

' ┌─────────────────────────────────────────────────────────┐
' │  ref_Excel                                              │
' │  역할 : am_Excel Application.Run 래퍼                  │
' └─────────────────────────────────────────────────────────┘

Private Const REF As String = "corelib.xlam!am_Excel."

' ── 인쇄 / 내보내기 ──────────────────────────────────────────

' 목적   : 시트 인쇄 영역, 여백, 용지 크기, 맞춤 설정 적용
' 인수   : sht           - 대상 시트 (기본: ActiveSheet)
'          rngPrintArea  - 인쇄 영역 Range (기본: UsedRange)
'          rngTitleRows  - 반복 인쇄 행 Range (없으면 Nothing)
'          blnAddArea    - True: 기존 인쇄 영역에 추가
'          paperSize     - 용지 크기 (기본: xlPaperA4)
'          orientation   - 방향 (기본: xlPortrait 세로)
'          sngTopMargin  - 상 여백 cm (기본: 1.5)
'          sngBottomMargin - 하 여백 cm (기본: 1)
'          sngLeftMargin - 좌 여백 cm (기본: 1)
'          sngRightMargin - 우 여백 cm (기본: 1)
'          CenterH       - 가로 가운데 맞춤 (기본: True)
'          CenterV       - 세로 가운데 맞춤 (기본: False)
'          fitToPage     - 페이지 맞춤 사용 (기본: True)
'          intWide       - 가로 페이지 수 (기본: 1)
'          intTall       - 세로 페이지 수 (기본: 0 = 자동)
'          intPer        - 확대/축소 비율 % (기본: 100)
' 예시   : SetPrintPage sht:=ActiveSheet, rngPrintArea:=Sheet1.Range("A1:H50")
Public Sub SetPrintPage(Optional ByVal sht As Worksheet, _
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
    Application.Run REF & "SetPrintPage", _
                    sht, rngPrintArea, rngTitleRows, blnAddArea, paperSize, orientation, _
                    sngTopMargin, sngBottomMargin, sngLeftMargin, sngRightMargin, _
                    CenterH, CenterV, fitToPage, intWide, intTall, intPer
End Sub

' 목적   : 시트를 PDF 파일로 내보내기
' 인수   : strFilePath  - 저장할 PDF 전체 파일 경로
'          sht          - 내보낼 시트 (기본: ActiveSheet)
'          blnOpenAfter - True: 내보낸 후 PDF 자동 열기
'          xlQual       - 품질 (기본: xlQualityStandard)
'          blnDocProps  - True: 문서 속성 포함
' 예시   : ExportPDF "C:\출력\report.pdf", ActiveSheet, blnOpenAfter:=True
Public Sub ExportPDF(ByVal strFilePath As String, _
                     Optional ByVal sht As Worksheet, _
                     Optional ByVal blnOpenAfter As Boolean = False, _
                     Optional ByVal xlQual As XlFixedFormatQuality = xlQualityStandard, _
                     Optional ByVal blnDocProps As Boolean = True)
    Application.Run REF & "ExportPDF", strFilePath, sht, blnOpenAfter, xlQual, blnDocProps
End Sub

' 목적   : 시트를 CSV 파일로 내보내기 (쉼표 구분, UTF-8)
' 인수   : strPath     - 저장 폴더 경로
'          strFileName - 저장 파일명 (확장자 포함)
'          sht         - 내보낼 시트 (기본: ActiveSheet)
' 예시   : ExportSheetToCSV "C:\출력", "data.csv"
'          ExportSheetToCSV "C:\출력", "data.csv", Sheet2
Public Sub ExportSheetToCSV(ByVal strPath As String, _
                            ByVal strFileName As String, _
                            Optional ByVal sht As Worksheet = Nothing)
    If sht Is Nothing Then Set sht = ActiveSheet
    Application.Run REF & "ExportSheetToCSV", strPath, strFileName, sht
End Sub

' ── 차트 ─────────────────────────────────────────────────────

' 목적   : 차트 개체의 데이터 범위 변경
' 인수   : strChart - 차트 개체 이름
'          rng      - 새 데이터 범위
'          ws       - 차트가 있는 시트 (기본: ActiveSheet)
' 예시   : SetChartDataRange("차트 1", Sheet1.Range("A1:C10"), Sheet1)
Public Sub SetChartDataRange(ByVal strChart As String, _
                             ByVal rng As Range, _
                             Optional ByVal ws As Worksheet)
    Application.Run REF & "SetChartDataRange", strChart, rng, ws
End Sub

' ── 도형 ─────────────────────────────────────────────────────

' 목적   : 도형에 연결된 매크로 실행
' 인수   : strShpName - 도형 이름
'          ws         - 도형이 있는 시트 (기본: ActiveSheet)
' 예시   : RunShpMacro("btn_실행", ActiveSheet)
Public Sub RunShpMacro(ByVal strShpName As String, _
                       Optional ByVal ws As Worksheet)
    Application.Run REF & "RunShpMacro", strShpName, ws
End Sub

' 목적   : 도형 텍스트 안전하게 읽기 (오류 시 "" 반환)
' 인수   : shp - 텍스트를 읽을 도형 개체
' 반환   : String - 도형 텍스트 ("" = 텍스트 없음 또는 오류)
' 예시   : GetShapeTextSafe(ActiveSheet.Shapes("lbl_Status")) → "완료"
Public Function GetShapeTextSafe(ByVal shp As Shape) As String
    GetShapeTextSafe = Application.Run(REF & "GetShapeTextSafe", shp)
End Function

' 키보드/마우스 래퍼는 ref_Automation 으로 이전됨 (2026-07-24)
