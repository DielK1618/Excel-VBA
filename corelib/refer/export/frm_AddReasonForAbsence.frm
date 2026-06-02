VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_AddReasonForAbsence 
   Caption         =   "교육면제등록"
   ClientHeight    =   6570
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   4812
   OleObjectBlob   =   "frm_AddReasonForAbsence.frx":0000
End
Attribute VB_Name = "frm_AddReasonForAbsence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub UserForm_Initialize()
    '<<<<< 폼 위치
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    '>>>>> 폼 위치
    
    Me.txt_StartYear.SetFocus
    
End Sub
Private Sub cmb_Code_Change()

    On Error Resume Next
    Dim arr As Variant
    Dim strCode As String, strID  As String
    
    strCode = Me.cmb_Code.value
    strID = celDC.Offset(, intOffset(celDC, "코드"))
    
    If strCode = "" Then GoTo EmptyValue
    
    Dim dv As DBvar
    With dv
        '1. 데이터베이스 ============
        .Alias = "EDU_LMS"
        .Table = "lms_edu_exempt"
        
        Dim arrDB As Variant
        arrDB = GetDbInfo(.Alias, False)
        
        .Type = arrDB(0)
        .Token = arrDB(2)
        .File = arrDB(3)
        .db = arrDB(4)
        .Server = arrDB(5)
        .Port = arrDB(6)
        .ID = arrDB(7)
        .PW = arrDB(8)
    
        .Query = "SELECT * FROM " & .Table & " WHERE EXEMPT_CD = '" & strCode & "' AND WATV_ID = '" & strID & "'"
        .arrQuery = Array(.Query)
        
        .arrData = SelectQueryArr(.arrQuery, .Type, .File)
        arr = .arrData(0)
    End With
        
EmptyValue:

    With Me
        If IsArrayEmpty(arr) Then
        
            .txt_StartYear.Enabled = True
            .txt_StartMonth.Enabled = True
            .txt_EndYear.Enabled = True
            .txt_EndMonth.Enabled = True
        
            .txt_StartYear.value = ""
            .txt_StartMonth.value = ""
            .txt_EndYear.value = ""
            .txt_EndMonth.value = ""
            .txt_Reason.value = ""
            .txt_Memo.value = ""
        Else
            .txt_StartYear.value = arr(3, 0)
            .txt_StartMonth.value = arr(4, 0)
            .txt_EndYear.value = arr(5, 0)
            .txt_EndMonth.value = arr(6, 0)
            .txt_Reason.value = arr(7, 0)
            .txt_Memo.value = arr(8, 0)
            
            .txt_StartYear.Enabled = False
            .txt_StartMonth.Enabled = False
            .txt_EndYear.Enabled = False
            .txt_EndMonth.Enabled = False
        End If
    End With
    
End Sub
' 시작연도 유효성 검사 (1900 이상의 숫자만 입력)
Private Sub txt_StartYear_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    '비어 있으면 종료
    If txt_StartYear.Text = "" Then Exit Sub
    
    '정수 확인
    If IsNumeric(txt_StartYear.Text) Then
        txt_StartYear.Text = Int(txt_StartYear.Text)
        
        If txt_StartYear.Text < 1900 Then
            MsgBox "1900이상의 정수를 입력해 주세요!", vbInformation
            txt_StartYear.Text = ""
            Cancel = True
        End If
    Else
        MsgBox "정수만 입력 가능합니다!", vbInformation
        txt_StartYear.Text = ""
        Cancel = True
        Exit Sub
    End If
End Sub
Private Sub txt_StartMonth_Exit(ByVal Cancel As MSForms.ReturnBoolean)
        
    If txt_StartMonth.Text = "" Then Exit Sub
    
    If txt_StartYear.Text = "" Then
        MsgBox "시작연을 먼저 입력해 주세요!", vbInformation
        txt_StartMonth.Text = ""
        Cancel = True
        Exit Sub
    End If
    
    '정수 확인
    If IsNumeric(txt_StartMonth.Text) Then
        txt_StartMonth.Text = Int(txt_StartMonth.Text)
        
        If txt_StartMonth.Text < 1 Or txt_StartMonth.Text > 12 Then
            MsgBox "1에서 12까지의 정수를 입력해 주세요!", vbInformation
            txt_StartMonth.Text = ""
            Cancel = True
        End If
    Else
        MsgBox "정수만 입력 가능합니다!", vbInformation
        txt_StartMonth.Text = ""
        Cancel = True
        Exit Sub
    End If
End Sub
Private Sub txt_EndYear_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If txt_EndYear.Text = "" Then Exit Sub
    
    If txt_StartMonth.Text = "" Then
        MsgBox "시작연월을 먼저 입력해 주세요!", vbInformation
        txt_EndYear.Text = ""
        Cancel = True
        Exit Sub
    End If
    
    If txt_EndYear.Text < txt_StartYear.Text Then
        MsgBox "종료연은 시작연보다 작을 수 없습니다!", vbInformation
        txt_EndYear.Text = ""
        Cancel = True
        Exit Sub
    End If
    
    '정수 확인
    If IsNumeric(txt_EndYear.Text) Then
        txt_EndYear.Text = Int(txt_EndYear.Text)
        
        If txt_EndYear.Text < 1900 Then
            MsgBox "1900이상의 정수를 입력해 주세요!", vbInformation
            txt_EndYear.Text = ""
            Cancel = True
        End If
    Else
        MsgBox "정수만 입력 가능합니다!", vbInformation
        txt_EndYear.Text = ""
        Cancel = True
        Exit Sub
    End If
End Sub
Private Sub txt_EndMonth_Exit(ByVal Cancel As MSForms.ReturnBoolean)
        
    If txt_EndMonth.Text = "" Then Exit Sub
    
    If txt_EndYear.Text = "" Then
        MsgBox "종료연을 먼저 입력해 주세요!", vbInformation
        txt_EndMonth.Text = ""
        Cancel = True
        Exit Sub
    End If
    
    If txt_EndYear.Text = "" Then
        MsgBox "종료연을 먼저 입력해 주세요!", vbInformation
        txt_EndMonth.Text = ""
        Cancel = True
        Exit Sub
    End If
    
    
    '정수 확인
    If IsNumeric(txt_EndMonth.Text) Then
        txt_EndMonth.Text = Int(txt_EndMonth.Text)
        
        If txt_EndMonth.Text < 1 Or txt_EndMonth.Text > 12 Then
            MsgBox "1에서 12까지의 정수를 입력해 주세요!", vbInformation
            txt_EndMonth.Text = ""
            Cancel = True
        End If
    Else
        MsgBox "정수만 입력 가능합니다!", vbInformation
        txt_EndMonth.Text = ""
        Cancel = True
        Exit Sub
    End If
    
    Dim DateStart, DateEnd As Date
    
    DateStart = DateSerial(txt_StartYear, txt_StartMonth, 1)
    DateEnd = DateSerial(txt_EndYear, txt_EndMonth, 1)
    
    If DateStart > DateEnd Then
        MsgBox "종료연월이 시작연월보다 작을 수 없습니다!", vbInformation
        txt_EndMonth.Text = ""
        Cancel = True
        Exit Sub
    End If
    
End Sub
Private Sub bt_New_Click()
    Me.cmb_Code.value = ""
End Sub
Private Sub bt_Save_Click()
    
    With Me
        Dim strCode As String, strID  As String
        
        strCode = Me.cmb_Code.value
        strID = Me.txt_ID
        
        If .txt_ID = "" Or .txt_StartYear = "" Or .txt_StartMonth = "" Or .txt_EndYear = "" Or .txt_EndMonth = "" Then
            MsgBox "면제 기간을 확인하세요!", vbInformation
            Exit Sub
        End If
        
        If strCode = "" Then strCode = Format(DateSerial(.txt_StartYear, .txt_StartMonth, 1), "yymm") & "-" & Format(DateSerial(.txt_EndYear, .txt_EndMonth, 1), "yymm")
    End With
    
    Dim dv As DBvar
    With dv
        '데이터베이스
        .Alias = "EDU_LMS"
        .Table = "lms_edu_exempt"
        
        Dim arrDB As Variant
        arrDB = GetDbInfo(.Alias, False)
        
        .Type = arrDB(0)
        .Token = arrDB(2)
        .File = arrDB(3)
        
        '등록코드가 기존 데이터 베이스에 있는지 확인
        .Query = "SELECT [EXEMPT_CD] FROM " & .Table & " WHERE [EXEMPT_CD] = '" & strCode & "' AND [WATV_ID] = '" & strID & "'"
        .arrQuery = Array(.Query)
        
        .arrData = SelectQueryArr(.arrQuery, .Type, .File)
        
        Dim arr As Variant
        arr = .arrData(0)
        
        Erase .arrData
        Erase .arrQuery
        
        '기존에 있으면 업데이트, 없으면 인설트
        If IsArrayEmpty(arr) Then
            .Query = "INSERT INTO " & .Table & " ([EXEMPT_CD], [WATV_ID], [ST_YYYY], [ST_MM], [ED_YYYY], [ED_MM], [REASON_NM], [MEMO]) VALUES (" & _
            "'" & strCode & _
            "','" & strID & _
            "'," & Me.txt_StartYear & _
            "," & Me.txt_StartMonth & _
            "," & Me.txt_EndYear & _
            "," & Me.txt_EndMonth & _
            ",'" & Me.txt_Reason & _
            "','" & Me.txt_Memo & "')"
        Else
            .Query = "UPDATE " & .Table & " SET" & _
            " [REASON_NM] = " & FormatValueForSQLByDBType(Me.txt_Reason, "VARCHAR") & _
            ", [MEMO] = " & FormatValueForSQLByDBType(Me.txt_Memo, "LONGTEXT") & _
            " WHERE [WATV_ID] = '" & Me.txt_ID & "' AND [EXEMPT_CD] = '" & strCode & "'"
        End If
        
        .arrQuery = Array(.Query)
        Call ExecuteQueryArr(.arrQuery, .Type, .File)
    End With
    
    Call SyncDB("DB_교육면제")
    
    MsgBox "Completion", vbInformation
    Unload Me
    
End Sub
Private Sub bt_Delete_Click()
    
    Dim strCode As String, strID As String
    
    strCode = Me.cmb_Code.value
    strID = Me.txt_ID.value
    
    If strCode = "" Or strID = "" Then
        MsgBox "[등록코드] 또는 [개인코드] 값을 확인하세요!", vbInformation
        Exit Sub
    End If
    
    If MsgBox(strCode & vbNewLine & "상기 코드의 레코드를 삭제하시겠습니까?" & vbNewLine & "삭제 후에는 복구가 불가능합니다!", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    Dim dv As DBvar
    With dv
        '1. 데이터베이스 ============
        .Alias = "EDU_LMS"
        .Table = "lms_edu_exempt"
        
        Dim arrDB As Variant
        arrDB = GetDbInfo(.Alias, False)
        
        .Type = arrDB(0)
        .Token = arrDB(2)
        .File = arrDB(3)
        
        .Query = "DELETE FROM " & .Table & " WHERE [WATV_ID] ='" & strID & "' AND [EXEMPT_CD] = '" & strCode & "';"
        .arrQuery = Array(.Query)
        
        Call ExecuteQueryArr(.arrQuery, .Type, .File)
    End With
    
    Call SyncDB("DB_교육면제")
    MsgBox "[" & strCode & "]" & vbNewLine & "상기 코드의 레코드가 삭제되었습니다!", vbInformation
    Unload Me
        
End Sub
Private Sub bt_Cancel_Click()
    Unload Me
End Sub
