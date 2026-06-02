Attribute VB_Name = "tpl_Const"
Option Explicit
Public cl As New Common '클래스모듈
Public Const PW = "qlalfqjsgh!^^" '시트잠금 비밀번호
Public Const CM_TO_POINTS As Double = 28.3465 'cm를 Point 단위로 변환
Public dicPub As Object
Public dicPubRev As Object
Public celDC As Range '더블클릭셀
Public wsPub As Worksheet '활성시트 외에 다른 시트를 잠글 때 사용, sht_Lock와 연동하여 사용
Public blnPartEvents As Boolean '편집을 위해 일부 이벤트를 비활성화 시킬 때 사용
Public MsoDilogType As MsoFileDialogType
Public vntTemp As Variant '임시로 필요할 때 사용
Public Type DBvar
        Type As String
        Alias As String
        Token As String
        File As String
        Server As String
        Port As String
        db As String
        ID As String
        PW As String
        fields As String
        Table As String
        QType As String
        Query As String
        QPlus As String
        cstQuery As String
        arrFields() As Variant
        arrData() As Variant
        arrQuery() As Variant
        tbl As ListObject
        Target As Range
End Type

Public Enum uiColor
    TC01 = 5971721
    TC02 = 11357748
    TC03 = 13735535
    TC04 = 15455930
    TC05 = 16770256
    TC06 = 16117997
    TC07 = 14676466
    TC08 = 15727103
    TC09 = 15528695
    TC10 = 1974161
    TC11 = 6802427
    TC12 = 2500134
    TC13 = 9211020
    TC14 = 12435391
    TC15 = 16053492
    TC16 = 16448250
End Enum
