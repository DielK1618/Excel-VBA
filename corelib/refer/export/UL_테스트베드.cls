VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UL_테스트베드"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)

    Dim dv As DBvar
    Dim rng As Range, t As Range
    Set rng = Intersect(Target, Range("T_테스트베드[별칭]")(1, 1))
    
    If Not rng Is Nothing Then
        
        For Each t In Target
            If Not Intersect(t, rng) Is Nothing Then
            With dv
                Dim intReq As Integer
                
                '1. 데이터베이스 ============
                .Alias = GetTwbRange("T_테스트베드[별칭]")(1, 1)
                
                If .Alias = "" Then
                    Range("T_테스트베드[[타입]:[패스워드]]").ClearContents
                    Exit Sub
                End If
                
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
                
                cl.DPUpdate_Off
                
                t.Offset(, 1) = .Type
                t.Offset(, 2) = .File
                t.Offset(, 3) = .Server
                t.Offset(, 4) = .Port
                t.Offset(, 5) = .db
                t.Offset(, 6) = .PW
                
                cl.DPUpdate_On
            End With
            
            End If
        Next
    End If
End Sub
