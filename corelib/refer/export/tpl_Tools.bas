Attribute VB_Name = "tpl_Tools"
Option Explicit
Sub SetHotKeys(Optional blnMsgBox As Boolean = False)
    
    Dim strFile As String
    strFile = Replace(ThisWorkbook.FullName, ThisWorkbook.Path, "") & "'!"
    strFile = "'" & Mid(strFile, 2)
    
    With Application
        .OnKey "^+b", strFile & "BackGroundColor"
        .OnKey "^+f", strFile & "FontColor"
        .OnKey "^+m", strFile & "BtNavigation"
        .OnKey "^+s", strFile & "BtBakupWorkBook"
        .OnKey "^+e", strFile & "TglEvents"
        .OnKey "^+p", strFile & "TglPartEvents"
        .OnKey "^+q", strFile & "RefreshAllPowerQueryTables"
        .OnKey "^+i", strFile & "ShowDBinfoForm"
        .OnKey "^+v", strFile & "VisibleAllSheets"
    End With
    
    If blnMsgBox Then MsgBox "Completion!", vbInformation
    
End Sub
Sub BackGroundColor()

    cl.sht_UnLock
    
    Dim inputValue As String
    Dim intColor As Long
    Dim f, rngFind As Range
    Dim objSelection As Object
    
    Set objSelection = Selection

    inputValue = Application.InputBox("ÄÃ·¯ ¹øÈ£¸¦ ÀÔ·ÂÇÏ¼¼¿ä.", "ÄÃ·¯Àû¿ë", Type:=1)
    If inputValue = False Then Exit Sub
    
    Set rngFind = GetTwbRange("T_Å×¸¶ÄÃ·¯[No]")
    Set f = rngFind.Find(inputValue, , xlValues, xlWhole)

    If Not f Is Nothing Then

        intColor = f.Offset(, 2).Interior.Color
        
        With objSelection
            On Error Resume Next
            .Interior.Color = intColor
            If Err.Number <> 0 Then
                MsgBox "¹üÀ§ ¶Ç´Â µµÇü ¿ÜÀÇ ´Ù¸¥ °³Ã¼´Â Àû¿ëµÇÁö ¾Ê½À´Ï´Ù!", vbCritical
                Exit Sub
            End If
            On Error GoTo 0
        End With
    Else
        
        MsgBox "ÇØ´ç ÄÃ·¯°¡ ¾ø½À´Ï´Ù!", vbCritical
            
    End If
    
    cl.sht_Lock
    
End Sub
Sub FontColor()
    
    cl.sht_UnLock
    
    Dim inputValue As String
    Dim intColor As Long
    Dim f, rngFind As Range
    Dim objSelection As Object
    
    Set objSelection = Selection
    
    inputValue = Application.InputBox("ÄÃ·¯ ¹øÈ£¸¦ ÀÔ·ÂÇÏ¼¼¿ä.", "ÄÃ·¯Àû¿ë", Type:=1)
        
    If inputValue = False Then Exit Sub
    
    Set rngFind = GetTwbRange("T_Å×¸¶ÄÃ·¯[No]")
    Set f = rngFind.Find(inputValue, , xlValues, xlWhole)
    
    If Not f Is Nothing Then
        intColor = f.Offset(, 2).Interior.Color
        
        With objSelection
            On Error Resume Next
            .Font.Color = intColor
            If Err.Number <> 0 Then
                MsgBox "¹üÀ§ ¶Ç´Â µµÇü ¿ÜÀÇ ´Ù¸¥ °³Ã¼´Â Àû¿ëµÇÁö ¾Ê½À´Ï´Ù!", vbCritical
                Exit Sub
            End If
            On Error GoTo 0
        End With
        
    Else
        
        MsgBox "ÇØ´ç ÄÃ·¯°¡ ¾ø½À´Ï´Ù!", vbCritical
        
    End If
    
    cl.sht_Lock
    
End Sub
Sub FindAndModifyShapes() 'Æ¯Á¤ µµÇüÀÇ ¸ÅÅ©·Î ¿¬°áÀ» ÀÏ°ý ¼öÁ¤
    
    Dim sht As Worksheet
    Dim shp As Shape
    Dim strShpName As String
    
    For Each sht In ThisWorkbook.Worksheets
        For Each shp In sht.Shapes
            On Error Resume Next
            strShpName = shp.TextFrame.Characters.Text
            If strShpName = "ADD ROWS" Then 'µµÇü¿¡ ÀÔ·ÂµÈ ÅØ½ºÆ®·Î Ã£À½
                Debug.Print sht.Name & " : " & strShpName
                shp.OnAction = "AddTableRows_BT"
            End If
            On Error GoTo 0
        Next
    Next

End Sub
Sub ChangeSizeForSpecificTextShapes() 'µµÇüÀÇ Å©±â¸¦ ÀÏ°ý ¼öÁ¤

    Dim ws As Worksheet
    Dim shp As Shape
    Dim txt As String
    Dim targetTexts As Variant
    Dim item As Variant
    Dim targetHeight As Double
    Dim targetWidth As Double
    
    ' ¹è¿­ ¼±¾ð
    targetTexts = Array("µµÇüÀÌ¸§")
    
    ' ³ôÀÌ¿Í °¡·ÎÆø ¼³Á¤ (´ÜÀ§: cm)
    targetHeight = 0.85 * 28.35  ' 0.85 cm -> points·Î º¯È¯ (1cm = 28.35 points)
    targetWidth = 3 * 28.35      ' 3 cm -> points·Î º¯È¯

    ' ¸ðµç ½ÃÆ®¸¦ ¼øÈ¸
    For Each ws In ThisWorkbook.Worksheets
        ' °¢ ½ÃÆ®ÀÇ ¸ðµç µµÇü¿¡ ´ëÇØ ¹Ýº¹
        For Each shp In ws.Shapes
            If Not shp.TextFrame2.HasText Then GoTo NextShape
            
            ' µµÇü¿¡ ÀÔ·ÂµÈ ÅØ½ºÆ® °¡Á®¿À±â
            On Error Resume Next
            txt = shp.TextFrame2.TextRange.Text
            On Error GoTo 0
            
            ' ÅØ½ºÆ®°¡ ¸ñÇ¥ ¸®½ºÆ®¿¡ ÀÖ´ÂÁö È®ÀÎ
            For Each item In targetTexts
                If StrComp(txt, item, vbTextCompare) = 0 Then
                
                    ' ÅØ½ºÆ®°¡ ¸ñ·Ï¿¡ ÀÖÀ¸¸é ³ôÀÌ¿Í °¡·ÎÆø ¼³Á¤
                    shp.Height = targetHeight
                    shp.Width = targetWidth
                    
                    Exit For
                End If
            Next item
            
NextShape:
        Next shp
    Next ws
End Sub
Sub CheckSelectionType()
    If Selection.Type = xlRange Then
        MsgBox "¼±ÅÃµÈ °ÍÀº ¼¿(Range)ÀÔ´Ï´Ù."
    ElseIf TypeName(Selection) = "DrawingObjects" Or TypeName(Selection) = "Shape" Then
        MsgBox "¼±ÅÃµÈ °ÍÀº µµÇü(Shape)ÀÔ´Ï´Ù."
    Else
        MsgBox "´Ù¸¥ À¯ÇüÀÇ °³Ã¼°¡ ¼±ÅÃµÇ¾ú½À´Ï´Ù."
    End If
End Sub
Sub ResetDB()
    Dim arrQuery()
    Dim strTable As String
    
    strTable = bdbEvTable
    arrQuery = Array("DELETE FROM " & strTable, "ALTER TABLE " & strTable & " ALTER COLUMN ID COUNTER(1,1)")
    Call ExecuteQueryArr(bdbType, arrQuery, bdbFile)
    
End Sub
Sub ResetAutoNumber()
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim maxID As Long
    
    Set db = CurrentDb
    Set tdf = db.TableDefs("Å×ÀÌºí¸í")
    
    ' ÇöÀç ÃÖ´ë ID °ª ±¸ÇÏ±â
    maxID = DMax("ID", "Å×ÀÌºí¸í")
    
    ' AutoNumber ½Ãµå°ª Àç¼³Á¤
    Set fld = tdf.fields("ID")
    fld.Properties("Seed") = maxID + 1
    
    ' Á¤¸®
    Set fld = Nothing
    Set tdf = Nothing
    Set db = Nothing
    
    MsgBox "AutoNumber°¡ " & (maxID + 1) & "ºÎÅÍ ½ÃÀÛÇÏµµ·Ï ¼³Á¤µÇ¾ú½À´Ï´Ù."
End Sub
Sub ChangeAllShapesFont() ' ¸ðµç µµÇüÀÇ ÆùÆ®¸¦ Pretendard ExtraBold·Î º¯°æÇÏ´Â ¸ÅÅ©·Î
    Dim ws As Worksheet
    Dim shp As Shape
    Dim shapeCount As Integer
    
    shapeCount = 0
    
    ' ÅëÇÕ½ÃÆ®ÀÇ ¸ðµç ½ÃÆ® ¼øÈ¸
    For Each ws In ThisWorkbook.Sheets
        ' °¢ ½ÃÆ®ÀÇ ¸ðµç µµÇü ¼øÈ¸
        On Error Resume Next
        For Each shp In ws.Shapes
            ' µµÇü¿¡ ÅØ½ºÆ®°¡ ÀÖ´Â °æ¿ì¸¸ Ã³¸®
            If shp.HasTextFrame Then
                With shp.TextFrame.Characters.Font
                    .Name = "Pretendard ExtraBold"
                End With
                shapeCount = shapeCount + 1
            End If
        Next shp
        On Error GoTo 0
    Next ws
    
    MsgBox "¿Ï·á! " & shapeCount & "°³ÀÇ µµÇü ÆùÆ®°¡ º¯°æµÇ¾ú½À´Ï´Ù.", vbInformation
    
End Sub
Sub SetAllUserFormFontsToPretendard() '¸ðµç »ç¿ëÀÚ Á¤ÀÇ ÆûÀÇ ÆùÆ®¸¦ ¹Ù²Ù´Â ÄÚµå
    Dim cmp As Object
    Dim uf As Object
    Dim ctl As Object
    
    For Each cmp In ThisWorkbook.VBProject.VBComponents
        If cmp.Type = vbext_ct_MSForm Then
            Set uf = cmp.Designer
            
            For Each ctl In uf.Controls
                On Error Resume Next
                ctl.Font.Name = "Pretendard Medium"
                On Error GoTo 0
            Next ctl
        End If
    Next cmp
End Sub
Private Sub WaitMs(ByVal ms As Long)
    If ms <= 0 Then Exit Sub
    Dim t As Single
    t = Timer
    Do While Timer - t < (ms / 1000#)
        DoEvents
    Loop
End Sub
Sub SyncCodeNamesToSheetNames()
    ' [ÁÖÀÇ] µµ±¸ > ÂüÁ¶¿¡¼­ "Microsoft Visual Basic for Applications Extensibility" Ã¼Å© ÇÊ¿ä

    Dim vbProj     As VBIDE.VBProject
    Dim vbComp     As VBIDE.VBComponent
    Dim ws         As Worksheet
    Dim strNewName As String

    Set vbProj = ThisWorkbook.VBProject

    For Each ws In ThisWorkbook.Worksheets
        strNewName = ws.Name

        ' À¯È¿ÇÏÁö ¾ÊÀº ¹®ÀÚ Á¦°Å (°ø¹é, Æ¯¼ö¹®ÀÚ ¡æ ¾ð´õ¹Ù·Î Ä¡È¯)
        strNewName = CleanCodeName(strNewName)

        ' CodeName º¯°æ
        Set vbComp = vbProj.VBComponents(ws.CodeName)
        vbComp.Name = strNewName

    Next ws

    MsgBox "¿Ï·á: CodeNameÀ» ½ÃÆ® ÅÇ ÀÌ¸§°ú µ¿ÀÏÇÏ°Ô º¯°æÇß½À´Ï´Ù.", vbInformation
End Sub
' -----------------------------------------------
' CodeName¿¡ »ç¿ëÇÒ ¼ö ¾ø´Â ¹®ÀÚ¸¦ ¾ð´õ¹Ù·Î Ä¡È¯
' Ã¹ ±ÛÀÚ°¡ ¼ýÀÚÀÌ¸é ¾Õ¿¡ "_" Ãß°¡
' -----------------------------------------------
Private Function CleanCodeName(strName As String) As String
    Dim i      As Integer
    Dim strOut As String
    Dim strChr As String

    strOut = ""
    For i = 1 To Len(strName)
        strChr = Mid(strName, i, 1)
        If strChr Like "[A-Za-z0-9°¡-ÆR_]" Then
            strOut = strOut & strChr
        Else
            strOut = strOut & "_"
        End If
    Next i

    ' Ã¹ ±ÛÀÚ°¡ ¼ýÀÚÀÎ °æ¿ì ¾Õ¿¡ _ Ãß°¡
    If Len(strOut) > 0 Then
        If Left(strOut, 1) Like "[0-9]" Then
            strOut = "_" & strOut
        End If
    End If

    CleanCodeName = strOut
End Function
