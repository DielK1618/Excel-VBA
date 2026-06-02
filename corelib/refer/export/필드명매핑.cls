VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "필드명매핑"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub cboTable_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = 13 Then '엔터 키(KeyCode = 13) 입력 시 이벤트 처리
        Call RunShpMacro("BtSearch_AccessTableFields")
    End If
    
End Sub
Private Sub cboDB_DropButtonClick()
    With Me.cboTable
        .Clear
        .value = ""
    End With
    
    If Me.cboDB.ListIndex = -1 Then
        Call FormSearchFilter(Me.cboDB)
    End If
End Sub
Private Sub cboTable_DropButtonClick()

    If Me.cboDB.ListIndex > -1 And Me.cboTable.ListIndex = -1 Then
        
        Me.cboTable.Clear
        
        Dim dv As DBvar
        
        With dv
        Dim arrDB As Variant
        arrDB = GetDbInfo(Me.cboDB.value, False)
        
        .Type = arrDB(0)
        .Token = arrDB(2)
        .File = arrDB(3)
        .db = arrDB(4)
        
        Dim arr
        arr = GetDbTables("엑세스", .File)
        
        If IsArrayEmpty(arr) Then Exit Sub
        
        Dim vnt As Variant
        
        For Each vnt In arr
            If vnt Like "*" & Me.cboTable.value & "*" Then
                Me.cboTable.AddItem vnt
            End If
        Next
        End With
    End If

End Sub
Private Sub cboDB_Change()
    
    Dim celTarget As Range
    Set celTarget = [U3]
    
    With Me.cboDB
        If .ListIndex = -1 Then
            celTarget = IIf(.value = "", "", "%" & .value & "%")
        Else
            celTarget = .value
        End If
    End With
    
End Sub
Private Sub cboTable_Change()
    
    Dim celTarget As Range
    Set celTarget = [v3]
    
    With Me.cboTable
        If .ListIndex = -1 Then
            celTarget = IIf(.value = "", "", "%" & .value & "%")
        Else
            celTarget = .value
        End If
    End With
    
End Sub


