Attribute VB_Name = "tpl_File"
Option Explicit
'<<<<< 기본문
Function GetPath(Optional DilogType As MsoFileDialogType = msoFileDialogFilePicker, _
                                Optional ByVal InitialPath As String = "", _
                                Optional ByVal blnMultiSelect As Boolean = False, _
                                Optional ByVal strExtComment As String = "Select Item", _
                                Optional ByVal strFileName As String, _
                                Optional ByVal strExt As String = "*.*", _
                                Optional ByVal strCancelPath As String = "", _
                                Optional ByVal blntWB As Boolean = True) As String '<<< 파일다이얼로그 (다중 선택시 결과값에 ";"로 구분되어 있어서 Split 함수로 나누어 사용 가능
                                
    ' 파일 선택 창
    Dim arrSelections() As String
    Dim fdg As FileDialog
    Dim i As Integer
    
    Set fdg = Application.FileDialog(DilogType)
    If InitialPath = "" Then InitialPath = ReplacePath(ThisWorkbook.Path)

    With fdg
        On Error Resume Next
'        .Title = "FilePick"    ' 상단 타이틀
        .Filters.Clear
        .Filters.Add strExtComment, strExt    ' 파일 확장자 세팅
        .InitialView = msoFileDialogViewLargeIcons    ' 파일 아이콘 보기 모드
        .InitialFileName = InitialPath     ' 초기 폴더 위치
        .AllowMultiSelect = blnMultiSelect    ' 파일 다중 선택 세팅 False는 하나만 선택, True면 다중 선택 (SelectedItems.Count로 선택된 파일 수 알 수 있음 for~next와 함께 사용하여 정보 얻음)
        .InitialFileName = strFileName
        On Error GoTo 0
        
        If .Show = -1 Then ' 파일 선택되면 -1, 미선택시 0
            ReDim arrSelections(1 To .SelectedItems.count)
            For i = 1 To .SelectedItems.count
                arrSelections(i) = .SelectedItems(i)
            Next i
            GetPath = Join(arrSelections, ";") ' 다중 선택된 파일들을 세미콜론으로 구분하여 반환
        Else
            
            GetPath = ""
            
            If strCancelPath <> "" Then
                If MsgBox("선택된 값이 없습니다." & vbNewLine & "[" & strCancelPath & "] 이 경로를 유지하시겠습니까?", vbQuestion + vbYesNo) = vbYes Then
                    GetPath = strCancelPath  '선택이 취소되었을 경우 빈 문자열 반환
                    Exit Function
                End If
            End If
        End If
    End With
    
    GetPath = ReplacePath(GetPath)
    
    If blntWB Then GetPath = Replace(LCase(GetPath), LCase(ReplacePath(ThisWorkbook.Path)), "{tPath}")
    Set fdg = Nothing
    
End Function
Sub BackupSheet(ByVal Path As String, _
                                   Optional ByVal Filename As String, _
                                   Optional ByVal sht As Worksheet)
    
    
    Dim arrPath() As String
    
    If Not sht Is Nothing Then
        sht.Copy
    Else
        Set sht = ActiveSheet
    End If
    
    sht.Copy
    
    Path = IIf(Right(Path, 1) <> "\", Path & "\", Path)
    Call MkFolder(Path)
    
    If Filename = "" Then Filename = "Bak_(" & Format(Now, "yyyymmdd_hhmmss") & ") " & Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1) & "_" & sht.Name & ".xlsx"
    
    Select Case GetExt(Filename)
    Case ".xlsx"
        ActiveWorkbook.SaveAs Path & Filename, 51
    Case ".xlsm"
        ActiveWorkbook.SaveAs Path & Filename, 52
    Case ".csv"
        ActiveWorkbook.SaveAs Path & Filename, 6
    Case ".txt"
        ActiveWorkbook.SaveAs Path & Filename, 4158
    End Select
    
    ActiveWorkbook.Close False

End Sub
Sub BackupWorkbook(ByVal Path As String, _
                                            Optional ByVal Filename As String) '<<현재 워크북 백업>>
                                            
    Dim originalWB As Workbook
    Set originalWB = ThisWorkbook
    
    Path = IIf(Right(Path, 1) <> "\", Path & "\", Path)
    If Filename = "" Then Filename = "Bak_(" & Format(Now, "yyyymmdd_hhmmss") & ") " & ThisWorkbook.Name
    
    originalWB.SaveCopyAs Path & Filename
    
End Sub
Function GetExt(ByVal strFileName As String) As String
    GetExt = Mid(strFileName, InStrRev(strFileName, "."))
End Function
Function ReverseText(ByVal strText As String) As String
    ReverseText = VBA.StrReverse(strText)
End Function
Function CheckFileExistence(ByVal Path As String) As Boolean '<<파일존재확인>>
    On Error Resume Next
    If Dir(Path) <> "" Then
        CheckFileExistence = True
    Else
        CheckFileExistence = False
    End If
    If Err.Number <> 0 Then CheckFileExistence = False
    On Error GoTo 0
End Function
Function CheckFolderExistence(ByVal Path As String) As Boolean '<<폴더존재확인>>
    On Error Resume Next
    If Dir(Path, vbDirectory) <> "" Then
        CheckFolderExistence = True
    Else
        CheckFolderExistence = False
    End If
    If Err.Number <> 0 Then CheckFolderExistence = False
    On Error GoTo 0
End Function
Sub MkFolder(ByVal Path As String) '<<폴더생성>>

    Dim arrPath() As String
    Dim strPath As String
    Dim i As Integer

    Path = IIf(Right(Path, 1) = "\", Left(Path, Len(Path) - 1), Path)
    arrPath = Split(Path, "\")

    For i = 0 To UBound(arrPath)
        strPath = strPath & arrPath(i) & "\"

        If Dir(strPath, vbDirectory) = "" Then '폴더명 마지막에 \의 유무 상관 없음"
            MkDir (strPath)
        End If
    Next

End Sub
Sub DelFolder(ByVal Path As String) '<<폴더삭제>>

    If Dir(Path, vbDirectory) <> "" Then
        On Error Resume Next
        Kill Path & "\*.*" ' 폴더 내의 모든 파일 삭제
        RmDir Path ' 폴더 삭제
        On Error GoTo 0
    End If
    
End Sub
Sub DelFile(ByVal Path As String) '<<파일삭제>>

    If Dir(Path) <> "" Then
        On Error Resume Next
        Kill Path
        On Error GoTo 0
    End If
    
End Sub
Sub RenFile(ByVal OldFileName As String, _
                                ByVal NewFileName As String) '<<파일이름변경>>>
                                
    Dim Path As String
    
    Path = Left(NewFileName, InStrRev(NewFileName, "\") - 1)
    Call MkFolder(Path)
    
    Name OldFileName As NewFileName
    
End Sub
Sub CopyFile(ByVal filePath As String, _
                            ByVal CopyPath As String) '<<파일복사>>
    
    Dim Path As String
    
    Path = Left(CopyPath, InStrRev(CopyPath, "\") - 1)
    Call MkFolder(Path)
    
    FileCopy filePath, CopyPath
    
End Sub
Sub CopyFolder(ByVal folderPath As String, _
                                ByVal CopyPath As String) '<<폴더복사>>
    
    Call MkFolder(CopyPath)
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    fso.CopyFolder folderPath, CopyPath
    Set fso = Nothing
    
End Sub
'기본문 >>>>>
'<<<<< 파일 및 폴더 검색
Sub GetFoldersList(ByVal folderPath As String, _
                               ByRef arrFolders() As Variant, _
                               Optional ByVal blnSubFolers As Boolean = True, _
                               Optional ByVal strFilter As String = "*", _
                               Optional ByRef i As Long = 0)
    
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objSubfolder As Object
    Dim strFolder As String

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(folderPath)
    
    ' 현재 폴더의 하위 폴더들을 배열에 추가
    For Each objSubfolder In objFolder.SubFolders
        strFolder = objSubfolder.Path ' 전체 경로를 저장
        
        If objSubfolder.Name Like strFilter Then
            ReDim Preserve arrFolders(i)
            arrFolders(i) = strFolder
            i = i + 1
        End If
        
        ' 각 하위 폴더에 대해 재귀적으로 함수 호출
        If blnSubFolers Then GetFoldersList objSubfolder.Path, arrFolders, , strFilter, i
    Next objSubfolder
    
End Sub
Function GetFilesList(ByVal folderPath As String, _
                                        Optional ByVal strFilter As String = "*", _
                                        Optional ByVal blnSubFolders As Boolean = True) As Variant

    Dim arrFolders() As Variant
    Dim varFolder As Variant
    
    ReDim arrFolders(0)
    arrFolders(0) = folderPath
    
    GetFoldersList folderPath, arrFolders, blnSubFolders, , 1
        
    Dim fso As Object
    Dim Folder As Object
    Dim File As Object
    Dim arrFiles() As Variant
    Dim strFile As String
    Dim i As Long

    Set fso = CreateObject("Scripting.FileSystemObject")

    For Each varFolder In arrFolders
    
        Set Folder = fso.GetFolder(varFolder)
        
        For Each File In Folder.Files
            strFile = File.Path
            If strFile Like strFilter Then
                ReDim Preserve arrFiles(i)
                arrFiles(i) = strFile
                i = i + 1
            End If
        Next File
    Next varFolder
    
    GetFilesList = arrFiles
    
End Function






