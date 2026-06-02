VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "±³À°ÀÚ·á"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub SearchBox_Change()
    Dim celTarget As Range
    Set celTarget = [U3]
    With Me.SearchBox
    If .ListIndex = -1 Then
        celTarget = IIf(.value = "", "", "%" & .value & "%")
    Else
        celTarget = .value
    End If
    End With
End Sub
Private Sub SearchBox_DropButtonClick()
    If Me.SearchBox.ListIndex = -1 Then
        Call FormSearchFilter
    End If
End Sub
Private Sub SearchBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then '¿£ÅÍ Å°(KeyCode = 13) ÀÔ·Â ½Ã ÀÌº¥Æ® Ã³¸®
        Call RunShpMacro("SEARCH")
    End If
End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    Dim t As Range, rng As Range
    Set t = Target(1, 1)
    Set rng = Intersect(t, Union(Range("T_±³À°ÀÚ·á[Ã©ÅÍ¸í]"), Range("TS_±³À°ÀÚ·á_½ÃÁðÃ©ÅÍ[Ã©ÅÍ¸í]")))

    If Not rng Is Nothing Then

        Dim tbl As ListObject
        Set tbl = ThisWorkbook.Worksheets("DB_Ã©ÅÍ").ListObjects(1)

        Dim strSeason As String
        strSeason = t.Offset(, intOffset(t, "½ÃÁðÄÚµå"))

        If strSeason <> "" Then
            Dim arr As Variant
            arr = TblFindVals_MC(tbl, "¸íÄª", "½ÃÁðÄÚµå", "=", strSeason)
            
            cl.sht_UnLock
            Call ValidationList(t, arr)
            cl.sht_Lock
            
        End If
    End If
    
End Sub
