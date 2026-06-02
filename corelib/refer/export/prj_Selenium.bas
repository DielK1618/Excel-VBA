Attribute VB_Name = "prj_Selenium"
Option Explicit
Public sl As New Selenium.WebDriver
Public Const strID As String = "modest1440"
Public Const strPW As String = "5tgbnhy6!"
Public Const strPIN As String = "009725"
Public Const strURL As String = "https://edulms.watv.org/login.wmc"
Sub OpenEduLmsSite(ByVal strURL As String)
    
    With sl
        ' WebDriver가 초기화되지 않았거나 창이 닫힌 경우 새로 시작
        On Error Resume Next
        Dim isActive As Boolean
        isActive = Len(.Title) > 0 ' 창이 열려 있는지 확인
        On Error GoTo 0
        
        If Not isActive Then
        
            .Start "chrome" ' Chrome 브라우저 시작
            .Window.Maximize ' 창 최대화
            .Get strURL ' URL로 이동

            '로그인
            .FindElementByCss("#id").SendKeys strID
            .FindElementByCss("#pass").SendKeys strPW
            .FindElementByCss("#login").Click
            .WaitForScript "document.readyState === 'complete'", 10000
    
            '개인 비밀번호 입력
            .FindElementByCss(".inputPWLine").SendKeys strPIN
            .FindElementByCss("#btnChkPw").Click
            On Error Resume Next
            .FindElementByCss(".inputPWLine").SendKeys strPIN
            .FindElementByCss("#btnChkPw").Click
            On Error GoTo 0
            .WaitForScript "document.readyState === 'complete'", 10000
        End If
        
    End With

End Sub


