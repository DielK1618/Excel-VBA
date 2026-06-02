Attribute VB_Name = "tpl_Buttons_other"
Option Explicit
Sub TglEvents()
    If Application.EnableEvents Then
        Application.EnableEvents = False
        MsgBox "이벤트 비활성!", vbInformation
    Else
        Application.EnableEvents = True
        MsgBox "이벤트 활성!", vbInformation
    End If
End Sub
Sub TglPartEvents()
    
    Dim strCurrentMode As String, strChangeMode As String, strPrompt As String
    
    strCurrentMode = IIf(blnPartEvents, "사용", "편집")
    strChangeMode = IIf(blnPartEvents, "편집", "사용")
    strPrompt = "■ 현재 [" & strCurrentMode & "모드] 입니다." & vbNewLine & _
                          "     [" & strChangeMode & "모드]로 변경하시겠습니까?"
    
    Dim strPassword As String
    strPassword = GetUnProtectPassword(strPrompt)
    
    If strPassword = "" Then Exit Sub
    If strPassword <> PW Then
        MsgBox "비밀번호를 확인해 주세요!", vbCritical
        Exit Sub
    End If
    
    cl.DPUpdate_Off
    cl.Calculate_Off
    
    Dim ws As Worksheet
    
    If blnPartEvents Then
        Application.EnableEvents = True
        blnPartEvents = False
        
        For Each ws In ThisWorkbook.Worksheets
            Set wsPub = ws
            cl.sht_UnLock
        Next
        
        Set wsPub = Nothing
        cl.WB_UNLock
        
        MsgBox "[편집모드]" & vbNewLine & "일부 이벤트가 비활성", vbInformation
    Else
        blnPartEvents = True
        Application.EnableEvents = True
        
        For Each ws In ThisWorkbook.Worksheets
            Set wsPub = ws
            cl.sht_Lock
        Next
        
        Set wsPub = Nothing
        cl.WB_Lock
        
        MsgBox "[사용모드]" & vbNewLine & "모든 이벤트가 활성", vbInformation
    End If
    
    cl.Calculate_On
    cl.DPUpdate_On
End Sub
Sub ShowDBinfoForm()
    frm_DBinfo.Show 0
End Sub
Sub ClearContents(ByVal celTarget As Range)
    celTarget.CurrentRegion.ClearContents
End Sub
Sub GetAllDbTableNames(ByVal celTarget As Range) '외부테이블 업데이트

    '데이터베이스 정보 가져오기
    Dim arrDBInfo()
    Dim r, rng As Range
    Dim i As Long
    
    Set rng = GetTwbRange("T_데이터베이스관리[예약어]")

    For Each r In rng
        If r <> "TWB_DB" Then
            ReDim Preserve arrDBInfo(i)
            arrDBInfo(i) = GetDbInfo(CStr(r))
            i = i + 1
        End If
    Next
    
    Dim dv As DBvar
    
    '데이터베이스 테이블 목록 가져오기
    Dim arrTables As Variant
    Dim arrResult()
    Dim j As Long
    
    For i = 0 To UBound(arrDBInfo)
        
        On Error GoTo pass
        
        Set celTarget = celTarget.Worksheet.Cells(Rows.count, celTarget.Column).End(xlUp)
        Set celTarget = IIf(celTarget = "", celTarget, celTarget.Offset(1))

        dv.Type = arrDBInfo(i)(0)
        dv.Alias = arrDBInfo(i)(1)
        dv.Token = arrDBInfo(i)(2)
        dv.File = arrDBInfo(i)(3)
        dv.db = arrDBInfo(i)(4)
        dv.Server = arrDBInfo(i)(5)
        dv.Port = arrDBInfo(i)(6)
        dv.ID = arrDBInfo(i)(7)
        dv.PW = arrDBInfo(i)(8)

        Select Case True
        Case dv.Type = "엑셀" Or dv.Type = "엑세스"
            arrTables = GetDbTables(dv.Type, dv.File)
            
            If IsArrayEmpty(arrTables) Then GoTo pass
            
            For j = LBound(arrTables) To UBound(arrTables) ' 결과 출력
                ReDim Preserve arrResult(4, j)
                arrResult(0, j) = dv.Type
                arrResult(1, j) = dv.Alias
                arrResult(2, j) = dv.Token
                arrResult(3, j) = dv.db
                arrResult(4, j) = arrTables(j)
            Next j
            celTarget.Resize(UBound(arrResult, 2) + 1, UBound(arrResult) + 1) = Application.WorksheetFunction.Transpose(arrResult)
        Case dv.Type = "서버"
            dv.Query = "SELECT """ & dv.Type & """, """ & dv.Alias & """, """ & dv.Token & """, """ & dv.db & """,  table_name FROM information_schema.tables WHERE table_name LIKE 'mdl\_%' OR table_name LIKE 'v\_%' OR table_name LIKE 'cu\_%' OR table_name LIKE 'ext\_%';"
            Call SelectQuery(celTarget, Array(dv.Query), dv.Type, , dv.Server, dv.Port, dv.db, dv.ID, dv.PW)
        End Select
pass:
    Next
    
End Sub
Sub ResetTargetPages()
    
    Dim arrSheets As Variant
    arrSheets = TblFindVals_MC(GetTwbRange("T_페이지설정").ListObject, "명칭", "테이블리셋", "=", 1)
    
    If IsArrayEmpty(arrSheets) Then Exit Sub
    
    Dim i As Long
    
    For i = 0 To UBound(arrSheets)
        Call BtClearSearchSheet(ThisWorkbook.Worksheets(arrSheets(i)), False)
    Next
    
End Sub
Sub BtGetVideoLength()
    
    If MsgBox("실행하시겠습니까?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    
    cl.DPUpdate_Off
    cl.Event_Off
    cl.Calculate_Off
    cl.sht_UnLock
    
    Dim r, rng As Range
    Dim Path As String
    Dim intColPath As Integer
    
    Set rng = Range("T_교육자료[시간]")
    intColPath = intOffset(rng, "파일")
    
    For Each r In rng
        If r.Offset(, intColPath) <> "" Then
        
            Path = ReplacePath(r.Offset(, intColPath))
        
        If Path <> "" Then
            If CheckFileExistence(Path) Then
                
                On Error Resume Next
                r.value = GetVideoLength(Path)
                On Error GoTo pass
                
            End If
        End If
        End If
pass:
    Next
    
    MsgBox "Completion", vbInformation
    
    cl.sht_Lock
    cl.Calculate_On
    cl.Event_On
    cl.DPUpdate_On
    
End Sub
