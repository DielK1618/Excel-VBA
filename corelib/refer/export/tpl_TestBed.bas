Attribute VB_Name = "tpl_TestBed"
Option Explicit
Sub BtTestSQLQuery()
    
    On Error GoTo ErrorHandler
    
    cl.DPUpdate_Off
    cl.Event_Off
    
    Dim dv As DBvar
    Dim celDBinfo As Range
    Dim strProcType As String


    With dv
        .Type = [C2]
        .File = [D2]
        .Server = [E2]
        .Port = [F2]
        .db = [G2]
        .ID = [H2]
        .PW = [I2]
        .Query = [J2]
        
        Set .Target = [A5]
        .Target.CurrentRegion.ClearContents
        .arrQuery = Array(.Query)
        strProcType = [A2]

        Select Case strProcType
        Case "Select"
            Call SelectQuery(.Target, .arrQuery, .Type, .File, .Server, .Port, .db, .ID, .PW, , , True)
        Case "SelectArr"
            Dim arr As Variant
            arr = SelectQueryArr(.arrQuery, .Type, .File, .Server, .Port, .db, .ID, .PW, True, True)
        Case "Execute"
            Call ExecuteQueryArr(.arrQuery, .Type, .File, .Server, .Port, .db, .ID, .PW)
        Case Else
            MsgBox "프로시저를 먼저 선택하세요!", vbInformation
        End Select
    End With
    
    cl.DPUpdate_On
    cl.Event_On
    
    Exit Sub
ErrorHandler:
    
    Call HandleError("BtTestSQLQuery")

    cl.DPUpdate_On
    cl.Event_On
End Sub

' ========================================
' 암호 입력 폼을 동적 생성 후 통합 문서 보호 해제
' ========================================
Sub UnprotectWorkbook()

    Dim sPassword As String

    ' 암호 입력 (* 마스킹 처리된 커스텀 InputBox)
    sPassword = GetMaskedInput("해제할 암호를 입력하세요.", "통합 문서 보호 해제")

    ' 취소 또는 빈 값이면 종료
    If sPassword = "" Then
        MsgBox "암호를 입력하지 않아 취소되었습니다.", vbExclamation
        Exit Sub
    End If

    ' 보호 해제 시도
    On Error GoTo WrongPassword
    ThisWorkbook.Unprotect Password:=sPassword

    MsgBox "통합 문서 보호가 해제되었습니다.", vbInformation
    Exit Sub

WrongPassword:
    MsgBox "암호가 올바르지 않습니다.", vbCritical

End Sub


' ========================================
' * 마스킹 InputBox ? UserForm 동적 생성
' ========================================
Function GetMaskedInput(sPrompt As String, sTitle As String) As String

    Dim oForm       As Object
    Dim oLabel      As Object
    Dim oTextBox    As Object
    Dim oBtnOK      As Object
    Dim oBtnCancel  As Object

    ' UserForm 동적 생성
    Set oForm = ThisWorkbook.VBProject.VBComponents.Add(3)  ' 3 = vbext_ct_MSForm

    With oForm
        .Properties("Caption") = sTitle
        .Properties("Width") = 240
        .Properties("Height") = 120
    End With

    ' 라벨
    Set oLabel = oForm.Designer.Controls.Add("Forms.Label.1")
    With oLabel
        .Caption = sPrompt
        .Left = 10
        .Top = 10
        .Width = 210
        .Height = 20
    End With

    ' 텍스트박스 (PasswordChar = "*" 로 마스킹)
    Set oTextBox = oForm.Designer.Controls.Add("Forms.TextBox.1")
    With oTextBox
        .Name = "txtPassword"
        .PasswordChar = "*"
        .Left = 10
        .Top = 35
        .Width = 210
        .Height = 20
    End With

    ' 확인 버튼
    Set oBtnOK = oForm.Designer.Controls.Add("Forms.CommandButton.1")
    With oBtnOK
        .Name = "btnOK"
        .Caption = "확인"
        .Left = 60
        .Top = 65
        .Width = 70
        .Height = 25
    End With

    ' 취소 버튼
    Set oBtnCancel = oForm.Designer.Controls.Add("Forms.CommandButton.1")
    With oBtnCancel
        .Name = "btnCancel"
        .Caption = "취소"
        .Left = 140
        .Top = 65
        .Width = 70
        .Height = 25
    End With

    ' 버튼 이벤트 코드 삽입
    Dim sCode As String
    sCode = "Sub btnOK_Click()" & vbCrLf & _
            "    Me.Tag = Me.txtPassword.Value" & vbCrLf & _
            "    Me.Hide" & vbCrLf & _
            "End Sub" & vbCrLf & _
            "Sub btnCancel_Click()" & vbCrLf & _
            "    Me.Tag = """"" & vbCrLf & _
            "    Me.Hide" & vbCrLf & _
            "End Sub"

    oForm.CodeModule.AddFromString sCode

    ' 폼 표시 및 결과 수집
    VBA.UserForms.Add(oForm.Name).Show

    Dim sResult As String
    sResult = VBA.UserForms(VBA.UserForms.count - 1).Tag  ' Tag에 저장된 암호 회수

    ' 폼 제거 (흔적 남기지 않음)
    ThisWorkbook.VBProject.VBComponents.Remove oForm

    GetMaskedInput = sResult

End Function


