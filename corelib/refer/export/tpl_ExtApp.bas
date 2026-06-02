Attribute VB_Name = "tpl_ExtApp"
Option Explicit
Sub OpenAddressInGoogleMaps(ByVal strAddress As String)
    Dim url As String
    ' 주소를 URL 인코딩 (공백 → %20)
    strAddress = Replace(strAddress, " ", "+")
    ' 구글 지도 검색 URL 생성
    url = "https://www.google.com/maps/search/?api=1&query=" & strAddress
    ' 기본 브라우저에서 열기
    ThisWorkbook.FollowHyperlink url
End Sub
