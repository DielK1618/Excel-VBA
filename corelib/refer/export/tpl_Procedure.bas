Attribute VB_Name = "tpl_Procedure"
Option Explicit
'Microsoft Visual Basic for Applications Extensibility 5.3 참조 추가
Sub BtCreateColorProcedure() '테마컬러 테이블을 기준으로 Enum uiColor을 생성 또는 수정합니다.

    Dim r As Range
    Dim rng As Range
    Dim strModule As String
    Dim strProcedure As String
    Dim strHeader As String
    Dim strBody As String
    Dim strTerminator As String
    Dim strLineText As String
    Dim strProc As String
        
    Set rng = GetTwbRange("T_테마컬러[컬러]")
    
    strModule = "tpl_Const"
    strProcedure = "uiColor"
    strHeader = "Public Enum uiColor"
    strTerminator = "End Enum"
    
    For Each r In rng
        strLineText = vbTab & r.Offset(, -1) & " = " & r.Interior.Color
        strBody = strBody & IIf(strBody = "", strLineText, vbNewLine & strLineText)
    Next
    
    strProc = strHeader & vbNewLine & strBody & vbNewLine & strTerminator
    
    Call DelProcedure(strModule, strProcedure)
    Call AddProcedure(strProc, strModule)
    
    MsgBox "Completion!", vbInformation
End Sub
Function CheckProc(strModuleName As String, _
                                    Optional strProcName As String)
                                    
    Dim vbComp As VBComponent
    Dim vbMod As CodeModule
    Dim longNum As Long
    
    On Error GoTo NotFound
    
    '모듈 확인
    Set vbComp = ThisWorkbook.VBProject.VBComponents(strModuleName)
    Set vbMod = vbComp.CodeModule
    
    '프로시저 확인

    If strProcName <> vbNullString Then
        longNum = vbMod.ProcStartLine(strProcName, vbext_pk_Proc)
    End If
    
    CheckProc = True
    Exit Function

NotFound:
    '선언형인지 확인
    Dim strLineText As String
    Dim strEnd As String
    
    Dim i As Long
    
    For i = 1 To vbMod.CountOfLines
        strLineText = Trim(vbMod.Lines(i, 1))
        If strLineText Like "*Enum " & strProcName Or strLineText Like "*Type " & strProcName Then
            CheckProc = True
            Exit Function
        End If
    Next

    CheckProc = False
End Function
Sub AddProcedure(strCodeText As String, _
                                  strModuleName As String, _
                                  Optional lngLinesNum As Long = 0)
                                  
    Dim vbComp As VBComponent
    Dim vbMod As CodeModule
    Dim asdfa As Integer
    ' 모듈 없으면 새로 생성
    If CheckProc(strModuleName) Then
        Set vbComp = ThisWorkbook.VBProject.VBComponents(strModuleName)
    Else
        Set vbComp = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_StdModule)
        vbComp.Name = strModuleName
    End If
    
    ' 코드 삽입
    Set vbMod = vbComp.CodeModule
    If lngLinesNum = 0 Then lngLinesNum = vbMod.CountOfLines + 1
    vbMod.InsertLines lngLinesNum, strCodeText
End Sub
Sub DelProcedure(strModuleName As String, _
                                 strProcName As String)
    
    If CheckProc(strModuleName, strProcName) Then
        Dim vbMod As CodeModule
        Set vbMod = ThisWorkbook.VBProject.VBComponents(strModuleName).CodeModule
    Else
        Exit Sub
    End If
    
    Dim lngLineStart As Long, lngLinesNum As Long
    
    On Error Resume Next
    lngLineStart = vbMod.ProcStartLine(strProcName, vbext_pk_Proc)
    lngLinesNum = vbMod.ProcCountLines(strProcName, vbext_pk_Proc)
    If Err.Number <> 0 Then
    
        Dim strLineText As String
        Dim strEndText As String
        Dim i As Long

        For i = 1 To vbMod.CountOfLines
            strLineText = Trim(vbMod.Lines(i, 1))
            
            If strEndText = vbNullString Then
                Select Case True
                Case strLineText Like "*Enum " & strProcName
                    strEndText = "End Enum"
                    lngLineStart = i
                Case strLineText Like "*Type " & strProcName
                    strEndText = "End Type"
                    lngLineStart = i
                End Select
            Else
                If strLineText Like strEndText Then
                    lngLinesNum = i - lngLineStart + 1
                    Exit For
                End If
            End If
        Next
    End If
    On Error GoTo 0
    
    ' 삭제 실행
    If lngLineStart > 0 And lngLinesNum > 0 Then
        vbMod.DeleteLines lngLineStart, lngLinesNum
    End If
End Sub
Function ReplaceModuleLine(ByVal strModuleName As String, _
                          ByVal strSearchText As String, _
                          ByVal strNewLine As String) As Boolean
    ' 모듈에서 특정 텍스트가 있는 줄을 찾아 전체 줄을 교체
    ' strModuleName: 모듈 이름
    ' strSearchText: 찾을 텍스트
    ' strNewLine: 새로운 줄 내용
    ' 반환값: True=성공, False=실패
    
    On Error GoTo ErrorHandler
    
    Dim vbComp As Object
    Dim vbMod As Object
    Dim lineNum As Long
    Dim lineCount As Long
    Dim lineText As String
    Dim found As Boolean
    
    ' VBProject 접근 권한 확인 필요
    ' (도구 > 매크로 > 보안 > 매크로 설정 > "VBA 프로젝트 개체 모델에 대한 액세스 신뢰" 체크)
    
    Set vbComp = ThisWorkbook.VBProject.VBComponents(strModuleName)
    Set vbMod = vbComp.CodeModule
    
    lineCount = vbMod.CountOfLines
    found = False
    
    ' 모든 줄을 검색
    For lineNum = 1 To lineCount
        lineText = vbMod.Lines(lineNum, 1)
        
        ' 특정 텍스트가 포함된 줄 찾기
        If InStr(1, lineText, strSearchText, vbTextCompare) > 0 Then
            ' 해당 줄을 삭제하고 새 내용으로 교체
            vbMod.ReplaceLine lineNum, strNewLine
            found = True
            Debug.Print "Line " & lineNum & " replaced: " & strNewLine
            Exit For ' 첫 번째 매칭만 교체하려면 Exit For 사용
        End If
    Next lineNum
    
    If found Then
        ReplaceModuleLine = True
    Else
        ReplaceModuleLine = False
        MsgBox "'" & strSearchText & "'를 찾을 수 없습니다.", vbExclamation
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "오류 발생: " & Err.Description & vbNewLine & _
           "VBA 프로젝트 개체 모델 접근 권한을 확인하세요.", vbCritical
    ReplaceModuleLine = False
End Function

Function ReplaceAllLinesInModule(ByVal strModuleName As String, _
                                 ByVal strSearchText As String, _
                                 ByVal strNewLine As String) As Long
    ' 모듈에서 특정 텍스트가 있는 모든 줄을 교체
    ' 반환값: 교체된 줄의 개수
    
    On Error GoTo ErrorHandler
    
    Dim vbComp As Object
    Dim vbMod As Object
    Dim lineNum As Long
    Dim lineCount As Long
    Dim lineText As String
    Dim replaceCount As Long
    
    Set vbComp = ThisWorkbook.VBProject.VBComponents(strModuleName)
    Set vbMod = vbComp.CodeModule
    
    lineCount = vbMod.CountOfLines
    replaceCount = 0
    
    ' 모든 줄을 검색 (역순으로 검색하면 줄 번호 변경 걱정 없음)
    For lineNum = lineCount To 1 Step -1
        lineText = vbMod.Lines(lineNum, 1)
        
        If InStr(1, lineText, strSearchText, vbTextCompare) > 0 Then
            vbMod.ReplaceLine lineNum, strNewLine
            replaceCount = replaceCount + 1
            Debug.Print "Line " & lineNum & " replaced"
        End If
    Next lineNum
    
    ReplaceAllLinesInModule = replaceCount
    
    If replaceCount > 0 Then
        MsgBox replaceCount & "개의 줄이 교체되었습니다.", vbInformation
    Else
        MsgBox "'" & strSearchText & "'를 찾을 수 없습니다.", vbExclamation
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "오류 발생: " & Err.Description, vbCritical
    ReplaceAllLinesInModule = 0
End Function

Function ReplaceTextInModule(ByVal strModuleName As String, _
                            ByVal strSearchText As String, _
                            ByVal strReplaceText As String) As Long
    ' 모듈에서 특정 텍스트를 찾아 교체 (줄 전체가 아닌 텍스트만)
    ' 반환값: 교체된 횟수
    
    On Error GoTo ErrorHandler
    
    Dim vbComp As Object
    Dim vbMod As Object
    Dim lineNum As Long
    Dim lineCount As Long
    Dim lineText As String
    Dim newLineText As String
    Dim replaceCount As Long
    
    Set vbComp = ThisWorkbook.VBProject.VBComponents(strModuleName)
    Set vbMod = vbComp.CodeModule
    
    lineCount = vbMod.CountOfLines
    replaceCount = 0
    
    For lineNum = 1 To lineCount
        lineText = vbMod.Lines(lineNum, 1)
        
        If InStr(1, lineText, strSearchText, vbTextCompare) > 0 Then
            newLineText = Replace(lineText, strSearchText, strReplaceText)
            vbMod.ReplaceLine lineNum, newLineText
            replaceCount = replaceCount + 1
            Debug.Print "Line " & lineNum & ": " & newLineText
        End If
    Next lineNum
    
    ReplaceTextInModule = replaceCount
    
    If replaceCount > 0 Then
        MsgBox replaceCount & "개의 항목이 교체되었습니다.", vbInformation
    Else
        MsgBox "'" & strSearchText & "'를 찾을 수 없습니다.", vbExclamation
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "오류 발생: " & Err.Description, vbCritical
    ReplaceTextInModule = 0
End Function




