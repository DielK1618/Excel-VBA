Attribute VB_Name = "tpl_Path"
Option Explicit
Public Property Get tFile() As String
    tFile = ReplacePath(ThisWorkbook.FullName)
End Property
Public Property Get tPath() As String
    tPath = ReplacePath(ThisWorkbook.Path) '& "\"
End Property
' ============================================
' 메인 함수 - 경로 변환
' ============================================
Function ReplacePath(ByVal Path As String, _
                    Optional blnChangeDrive As Boolean = False, _
                    Optional blnAddDrive As Boolean = True) As String
    

    If Path = "" Then Exit Function
    On Error GoTo ErrorHandler
    
    ' 기존 처리 과정
    Path = TokenConverter(Path)
    Path = ProcessWorkbookToken(Path)
    Path = ProcessNetworkDrives(Path)
    
    ' 단순화된 클라우드 경로 처리
    Dim processedPath As String
    processedPath = ProcessCloudPathsSimple(Path)
    
    ' 변환 실패 확인
    If processedPath = "" Then
        MsgBox "경로를 찾을 수 없습니다." & vbCrLf & vbCrLf & _
               "원본 경로: " & Path & vbCrLf & vbCrLf & _
               "모든 OneDrive 계정을 확인했으나 유효한 경로를 찾지 못했습니다.", _
               vbExclamation, "경로 오류"
        ReplacePath = ""
        Exit Function
    End If
    
    Path = processedPath
    Path = NormalizePath(Path)
    Path = ProcessDriveMapping(Path, blnChangeDrive, blnAddDrive)
    
    ' 유효성 검증 및 드라이브 대체 로직
    Path = ValidateAndFixDrivePath(Path)
    
    ReplacePath = Path
    Exit Function
    
ErrorHandler:
    MsgBox "경로 변환 중 오류가 발생했습니다." & vbCrLf & vbCrLf & _
           "오류 내용: " & Err.Description & vbCrLf & _
           "원본 경로: " & Path, vbCritical, "경로 변환 오류"
    ReplacePath = ""
    On Error GoTo 0
End Function
' ============================================
' 단순화된 클라우드 경로 처리
' ============================================
Private Function ProcessCloudPathsSimple(ByVal Path As String) As String
    On Error Resume Next
    
    Dim originalPath As String
    originalPath = Path
    
    ' 1단계: 원본 경로가 이미 유효한지 확인
    If IsValidPath(Path) Then
        ProcessCloudPathsSimple = Path
        Exit Function
    End If
    
    ' 2단계: URL이면 모든 OneDrive 경로로 변환 시도
    If InStr(Path, "http") > 0 Or InStr(Path, "sharepoint.com") > 0 Or _
       InStr(Path, "d.docs.live.net") > 0 Or InStr(Path, "1drv.ms") > 0 Then
        
        Dim convertedPath As String
        convertedPath = TryAllOneDriveConversions(Path)
        
        If convertedPath <> "" Then
            ProcessCloudPathsSimple = convertedPath
            Exit Function
        End If
    End If
    
    ' 3단계: 네트워크 경로면 OneDrive 로컬 경로로 변환 시도
    If InStr(Path, "\\") = 1 Then
        convertedPath = TryNetworkToOneDriveConversion(Path)
        
        If convertedPath <> "" Then
            ProcessCloudPathsSimple = convertedPath
            Exit Function
        End If
    End If
    
    ' 모든 시도 실패
    ProcessCloudPathsSimple = ""
    On Error GoTo 0
End Function
' ============================================
' 모든 OneDrive 경로로 URL 변환 시도
' ============================================
Private Function TryAllOneDriveConversions(ByVal webUrl As String) As String
    On Error Resume Next
    
    Dim oneDrivePaths As Collection
    Dim odPath As Variant
    Dim convertedPath As String
    Dim decodedUrl As String
    
    ' URL 디코딩
    decodedUrl = URLDecode(webUrl)
    
    ' 모든 OneDrive 경로 가져오기
    Set oneDrivePaths = GetAllOneDrivePaths()
    
    ' 각 OneDrive 경로로 변환 시도
    For Each odPath In oneDrivePaths
        ' SharePoint/OneDrive Business URL 변환 시도
        convertedPath = ConvertSharePointUrl(decodedUrl, CStr(odPath))
        If convertedPath <> "" And IsValidPath(convertedPath) Then
            TryAllOneDriveConversions = convertedPath
            Exit Function
        End If
        
        ' OneDrive Personal URL 변환 시도
        convertedPath = ConvertPersonalUrl(decodedUrl, CStr(odPath))
        If convertedPath <> "" And IsValidPath(convertedPath) Then
            TryAllOneDriveConversions = convertedPath
            Exit Function
        End If
    Next odPath
    
    ' 변환 실패
    TryAllOneDriveConversions = ""
    On Error GoTo 0
End Function
' ============================================
' 네트워크 경로를 OneDrive 로컬 경로로 변환 시도
' ============================================
Private Function TryNetworkToOneDriveConversion(ByVal networkPath As String) As String
    On Error Resume Next
    
    Dim oneDrivePaths As Collection
    Dim odPath As Variant
    Dim convertedPath As String
    
    ' 네트워크 경로 매핑 정의 (필요시 추가)
    Dim networkMappings As Variant
    networkMappings = Array("\\my04004.waffice.org\wf20024")
    
    Set oneDrivePaths = GetAllOneDrivePaths()
    
    ' 각 네트워크 매핑과 OneDrive 경로 조합으로 시도
    Dim mapping As Variant
    For Each mapping In networkMappings
        For Each odPath In oneDrivePaths
            convertedPath = Replace(networkPath, CStr(mapping), CStr(odPath), , , vbTextCompare)
            
            If convertedPath <> networkPath And IsValidPath(convertedPath) Then
                TryNetworkToOneDriveConversion = convertedPath
                Exit Function
            End If
        Next odPath
    Next mapping
    
    ' 변환 실패
    TryNetworkToOneDriveConversion = ""
    On Error GoTo 0
End Function
' ============================================
' SharePoint/OneDrive Business URL 변환
' ============================================
Private Function ConvertSharePointUrl(ByVal webUrl As String, ByVal oneDrivePath As String) As String
    On Error Resume Next
    
    Dim pathPart As String
    Dim pos As Long
    Dim result As String
    
    ' "/Documents/" 패턴 찾기
    pos = InStr(webUrl, "/Documents/")
    If pos > 0 Then
        pathPart = Mid(webUrl, pos + 11)  ' "/Documents/" 길이 = 11
        pathPart = Replace(pathPart, "/", "\")
        result = oneDrivePath & "\" & pathPart
        
        ConvertSharePointUrl = result
        Exit Function
    End If
    
    ' "/_layouts/" 패턴도 시도
    pos = InStr(webUrl, "/_layouts/")
    If pos > 0 Then
        ' 복잡한 SharePoint URL은 일단 스킵
        ConvertSharePointUrl = ""
        Exit Function
    End If
    
    ConvertSharePointUrl = ""
    On Error GoTo 0
End Function
' ============================================
' OneDrive Personal URL 변환
' ============================================
Private Function ConvertPersonalUrl(ByVal webUrl As String, ByVal oneDrivePath As String) As String
    On Error Resume Next
    
    Dim pathPart As String
    Dim pos As Long
    Dim nextPos As Long
    Dim result As String
    
    ' "d.docs.live.net/" 패턴 찾기
    pos = InStr(webUrl, "d.docs.live.net/")
    If pos = 0 Then
        ConvertPersonalUrl = ""
        Exit Function
    End If
    
    ' ID 부분 건너뛰기
    pathPart = Mid(webUrl, pos + 16)  ' "d.docs.live.net/" 길이 = 16
    
    ' 첫 번째 "/" 이후가 실제 경로
    nextPos = InStr(pathPart, "/")
    If nextPos > 0 Then
        pathPart = Mid(pathPart, nextPos + 1)
        pathPart = Replace(pathPart, "/", "\")
        
        result = oneDrivePath & "\" & pathPart
        ConvertPersonalUrl = result
        Exit Function
    End If
    
    ConvertPersonalUrl = ""
    On Error GoTo 0
End Function
' ============================================
' 모든 OneDrive 경로 수집
' ============================================
Private Function GetAllOneDrivePaths() As Collection
    On Error Resume Next
    
    Dim paths As New Collection
    Dim envPath As String
    
    ' 환경 변수에서 OneDrive 경로 수집
    envPath = Environ("OneDrive")
    If envPath <> "" Then paths.Add envPath
    
    envPath = Environ("OneDriveCommercial")
    If envPath <> "" And Not ContainsPath(paths, envPath) Then paths.Add envPath
    
    envPath = Environ("OneDriveConsumer")
    If envPath <> "" And Not ContainsPath(paths, envPath) Then paths.Add envPath
    
    ' 레지스트리에서 추가 OneDrive 계정 찾기
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    
    Dim i As Integer
    Dim regPath As String
    
    ' OneDrive Personal 계정들
    For i = 1 To 5
        On Error Resume Next
        regPath = ""
        regPath = wsh.RegRead("HKEY_CURRENT_USER\Software\Microsoft\OneDrive\Accounts\Personal_" & i & "\UserFolder")
        If regPath <> "" And Not ContainsPath(paths, regPath) Then
            paths.Add regPath
        End If
        On Error GoTo 0
    Next i
    
    ' OneDrive Business 계정들
    For i = 1 To 5
        On Error Resume Next
        regPath = ""
        regPath = wsh.RegRead("HKEY_CURRENT_USER\Software\Microsoft\OneDrive\Accounts\Business" & i & "\UserFolder")
        If regPath <> "" And Not ContainsPath(paths, regPath) Then
            paths.Add regPath
        End If
        On Error GoTo 0
    Next i
    
    Set wsh = Nothing
    Set GetAllOneDrivePaths = paths
    
    On Error GoTo 0
End Function
' ============================================
' Collection에 경로가 이미 있는지 확인
' ============================================
Private Function ContainsPath(ByVal paths As Collection, ByVal searchPath As String) As Boolean
    On Error Resume Next
    
    Dim p As Variant
    ContainsPath = False
    
    For Each p In paths
        If StrComp(CStr(p), searchPath, vbTextCompare) = 0 Then
            ContainsPath = True
            Exit Function
        End If
    Next p
    
    On Error GoTo 0
End Function
' ============================================
' URL 디코딩
' ============================================
Private Function URLDecode(ByVal strEncode As String) As String
    Dim strDecode As String
    
    strDecode = strEncode
    strDecode = Replace(strDecode, "+", " ")
    strDecode = Replace(strDecode, "%20", " ")
    strDecode = Replace(strDecode, "%2F", "/")
    strDecode = Replace(strDecode, "%5C", "\")
    strDecode = Replace(strDecode, "%3A", ":")
    
    URLDecode = strDecode
End Function
' ============================================
' 토큰 변환
' ============================================
Private Function TokenConverter(ByVal Path As String) As String
    On Error Resume Next
    Dim f, rng As Range
    Dim strExtract As String
    strExtract = Mid(Path, InStr(Path, "{") + 1, InStr(Path, "}") - InStr(Path, "{") - 1)
    
    Set rng = GetTwbRange("T_데이터베이스관리[예약어]")
    Set f = rng.Find(strExtract, , xlValues, xlWhole)
    
    TokenConverter = Path

    If Not f Is Nothing Then
        TokenConverter = Replace(Path, "{" & strExtract & "}", f.Offset(, intOffset(f, "파일")))
    End If
    On Error GoTo 0
End Function
' ============================================
' 워크북 토큰 처리
' ============================================
Private Function ProcessWorkbookToken(ByVal Path As String) As String
    On Error Resume Next
    ProcessWorkbookToken = Replace(Path, "{tPath}", ThisWorkbook.Path, , , vbTextCompare)
    ProcessWorkbookToken = Replace(ProcessWorkbookToken, "{tFile}", ThisWorkbook.FullName, , , vbTextCompare)
    On Error GoTo 0
End Function
' ============================================
' 네트워크 드라이브 매핑 처리
' ============================================
Private Function ProcessNetworkDrives(ByVal Path As String) As String
    On Error Resume Next
    
    Dim objNetwork As Object
    Dim objDrives As Object
    Dim i As Integer
    Dim twbPath As String
    
    Set objNetwork = CreateObject("WScript.Network")
    If objNetwork Is Nothing Then GoTo CleanUp
    
    Set objDrives = objNetwork.EnumNetworkDrives
    If objDrives Is Nothing Then GoTo CleanUp
    
    twbPath = ThisWorkbook.Path
    
    For i = 0 To objDrives.count - 1 Step 2
        If objDrives.item(i + 1) <> "" Then
            twbPath = Replace(twbPath, objDrives.item(i + 1), objDrives.item(i), , , vbTextCompare)
            Path = Replace(Path, objDrives.item(i + 1), objDrives.item(i), , , vbTextCompare)
        End If
    Next i
    
CleanUp:
    If Not objDrives Is Nothing Then Set objDrives = Nothing
    If Not objNetwork Is Nothing Then Set objNetwork = Nothing
    
    ProcessNetworkDrives = Path
    On Error GoTo 0
End Function
' ============================================
' 경로 정규화
' ============================================
Private Function NormalizePath(ByVal Path As String) As String
    NormalizePath = Replace(Path, "/", "\")
End Function
' ============================================
' 드라이브 매핑 처리
' ============================================
Private Function ProcessDriveMapping(ByVal Path As String, _
                                   ByVal blnChangeDrive As Boolean, _
                                   ByVal blnAddDrive As Boolean) As String
    On Error Resume Next
    
    Dim twbPath As String
    Dim twbPathDrive As String
    
    twbPath = ThisWorkbook.Path
    twbPathDrive = Left(twbPath, InStr(twbPath, ":"))
    
    If blnChangeDrive Then
        If InStr(Path, ":") > 0 Then
            Dim PathTemp As String
            PathTemp = Mid(Path, InStr(Path, ":") + 1)
            Path = twbPathDrive & PathTemp
        End If
    End If
    
    If blnAddDrive Then
        If InStr(Path, ":") = 0 Then
            Path = twbPathDrive & Path
        End If
    End If
    
    ProcessDriveMapping = Path
    On Error GoTo 0
End Function
' ============================================
' 유효성 검증 및 드라이브 대체 로직
' ============================================
Private Function ValidateAndFixDrivePath(ByVal Path As String) As String
    On Error Resume Next
    
    ' 1차: 원본 경로 유효성 검사
    If IsValidPath(Path) Then
        ValidateAndFixDrivePath = Path
        Exit Function
    End If
    
    ' 2차: 현재 워크북 드라이브로 대체 시도
    Dim fixedPath As String
    fixedPath = ReplaceWithWorkbookDrive(Path)
    
    If IsValidPath(fixedPath) Then
        ValidateAndFixDrivePath = fixedPath
        Exit Function
    End If
    
    ' 3차: A~Z 모든 드라이브 시도
    fixedPath = TryAllDrives(Path)
    
    If IsValidPath(fixedPath) Then
        ValidateAndFixDrivePath = fixedPath
    Else
        ' 모든 시도 실패 시 원본 경로 반환
        ValidateAndFixDrivePath = Path
    End If
    
    On Error GoTo 0
End Function
' ============================================
' 현재 워크북 드라이브로 경로 대체
' ============================================
Private Function ReplaceWithWorkbookDrive(ByVal Path As String) As String
    On Error Resume Next
    
    If InStr(Path, ":") = 0 Then
        ReplaceWithWorkbookDrive = Path
        Exit Function
    End If
    
    Dim twbDrive As String
    Dim pathWithoutDrive As String
    
    twbDrive = Left(ThisWorkbook.Path, 2)
    pathWithoutDrive = Mid(Path, 3)
    
    ReplaceWithWorkbookDrive = twbDrive & pathWithoutDrive
    
    On Error GoTo 0
End Function
' ============================================
' A~Z 모든 드라이브 시도
' ============================================
Private Function TryAllDrives(ByVal Path As String) As String
    On Error Resume Next
    
    If InStr(Path, ":") = 0 Then
        TryAllDrives = Path
        Exit Function
    End If
    
    Dim pathWithoutDrive As String
    Dim testPath As String
    Dim driveChar As String
    Dim i As Integer
    
    pathWithoutDrive = Mid(Path, 3)
    
    ' A부터 Z까지 모든 드라이브 시도
    For i = 65 To 90  ' ASCII: A=65, Z=90
        driveChar = Chr(i)
        testPath = driveChar & ":" & pathWithoutDrive
        
        If IsValidPath(testPath) Then
            TryAllDrives = testPath
            Exit Function
        End If
    Next i
    
    ' 모든 드라이브 실패 시 원본 반환
    TryAllDrives = Path
    
    On Error GoTo 0
End Function
' ============================================
' 경로 유효성 검사
' ============================================
Private Function IsValidPath(ByVal Path As String) As Boolean
    On Error Resume Next
    
    ' 빈 경로 체크
    If Len(Trim(Path)) = 0 Then
        IsValidPath = False
        Exit Function
    End If
    
    ' 경로 길이 제한 체크
    If Len(Path) > 260 Then
        IsValidPath = False
        Exit Function
    End If
    
    ' 파일 또는 폴더 존재 여부 확인
    Dim result As String
    result = Dir(Path, vbDirectory)
    
    If result <> "" Then
        IsValidPath = True
    Else
        result = Dir(Path)
        IsValidPath = (result <> "")
    End If
    
    On Error GoTo 0
End Function


