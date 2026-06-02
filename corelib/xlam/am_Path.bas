Attribute VB_Name = "am_Path"
Option Explicit

' ┌─────────────────────────────────────────────────────────┐
' │  am_Path                                                │
' │  역할 : 경로 토큰 치환 및 경로 정규화 처리             │
' └─────────────────────────────────────────────────────────┘

' ── 1. 모듈 변수 선언 ────────────────────────────────────────
Private m_dicTokens As Object

' ══════════════════════════════════════════════════════════
'  PUBLIC - 공개 함수
' ══════════════════════════════════════════════════════════

' 목적   : 토큰이 포함된 경로를 실제 로컬 경로로 반환
' 인수   : strPath        - 변환할 경로 (토큰 포함 가능)
'          strCwbPath     - 호출한 CWB 의 폴더 경로
'          strCwbFile     - 호출한 CWB 의 전체 파일 경로
'          blnChangeDrive - True: 드라이브를 CWB 드라이브로 강제 교체
'          blnAddDrive    - True: 드라이브 없는 경로에 CWB 드라이브 추가
' 반환   : String - 변환된 로컬 경로, 실패 시 ""
' 예시   : am_Path.ReplacePath("{cPath}\DB\main.accdb", _
'                               ThisWorkbook.Path, _
'                               ThisWorkbook.FullName)
Public Function ReplacePath(ByVal strPath As String, _
                            ByVal strCwbPath As String, _
                            ByVal strCwbFile As String, _
                            Optional ByVal blnChangeDrive As Boolean = False, _
                            Optional ByVal blnAddDrive As Boolean = True) As String

    On Error GoTo ErrHandler

    If strPath = "" Then Exit Function

    ' ── 1. 고정 토큰 치환 ────────────────────────────────────
    strPath = prv_ResolveFixedTokens(strPath, strCwbPath, strCwbFile)

    ' ── 2. 커스텀 토큰 치환 ──────────────────────────────────
    strPath = prv_ResolveCustomTokens(strPath)

    ' ── 3. 네트워크 드라이브 치환 ────────────────────────────
    strPath = prv_ProcessNetworkDrives(strPath, strCwbPath)

    ' ── 4. 클라우드 경로 변환 ────────────────────────────────
    strPath = prv_ProcessCloudPaths(strPath)

    ' ── 5. 경로 정규화 ───────────────────────────────────────
    strPath = prv_NormalizePath(strPath)

    ' ── 6. 드라이브 보정 처리 ────────────────────────────────
    strPath = prv_ProcessDriveMapping(strPath, strCwbPath, blnChangeDrive, blnAddDrive)

    ' ── 7. 유효성 검증 및 드라이브 추정 ─────────────────────
    strPath = prv_ValidateAndFixDrive(strPath)

    If strPath = "" Then
        MsgBox "경로를 찾을 수 없습니다." & vbCrLf & _
               "입력 경로: " & strPath, vbExclamation, am_Core.AM_NAME
    End If

    ReplacePath = strPath
    Exit Function

ErrHandler:
    MsgBox "[am_Path.ReplacePath] 오류 발생" & vbCrLf & _
           "오류 " & Err.Number & ": " & Err.Description & vbCrLf & _
           "경로: " & strPath, vbCritical, am_Core.AM_NAME
    ReplacePath = ""

End Function

' ══════════════════════════════════════════════════════════
'  PRIVATE - 토큰 처리
' ══════════════════════════════════════════════════════════

' 목적   : {cPath}, {cFile}, {xPath}, {xFile} 치환
Private Function prv_ResolveFixedTokens(ByVal strPath As String, _
                                        ByVal strCwbPath As String, _
                                        ByVal strCwbFile As String) As String

    strPath = Replace(strPath, "{cPath}", strCwbPath, , , vbTextCompare)
    strPath = Replace(strPath, "{cFile}", strCwbFile, , , vbTextCompare)
    strPath = Replace(strPath, "{xPath}", am_Core.XlamPath, , , vbTextCompare)
    strPath = Replace(strPath, "{xFile}", am_Core.XlamFullName, , , vbTextCompare)

    prv_ResolveFixedTokens = strPath

End Function

' 목적   : LoadTokens 에 등록된 커스텀 토큰 치환
Private Function prv_ResolveCustomTokens(ByVal strPath As String) As String

    If m_dicTokens Is Nothing Then Call prv_LoadTokens
    If m_dicTokens.Count = 0 Then Call prv_LoadTokens

    Dim vntKey As Variant
    For Each vntKey In m_dicTokens.Keys
        strPath = Replace(strPath, "{" & vntKey & "}", _
                          m_dicTokens(vntKey), , , vbTextCompare)
    Next vntKey

    prv_ResolveCustomTokens = strPath

End Function

' 목적   : 커스텀 토큰 배열을 딕셔너리에 등록
' 참고   : 토큰 추가/수정은 arrTokens 배열 직접 수정
'          배열 형식: ("토큰명", "경로", "토큰명", "경로", ...)
Private Sub prv_LoadTokens()

    Dim arrTokens As Variant
    arrTokens = Array( _
        "xDB",  am_Core.XlamPath & "\DB", _
        "xBak", am_Core.XlamPath & "\Backup" _
    )

    If (UBound(arrTokens) - LBound(arrTokens) + 1) Mod 2 <> 0 Then
        MsgBox "[am_Path.prv_LoadTokens] 토큰 배열이 올바르지 않습니다.", _
               vbCritical, am_Core.AM_NAME
        Exit Sub
    End If

    Set m_dicTokens = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = LBound(arrTokens) To UBound(arrTokens) Step 2
        m_dicTokens(arrTokens(i)) = arrTokens(i + 1)
    Next i

End Sub

' ══════════════════════════════════════════════════════════
'  PRIVATE - 경로 변환 처리
' ══════════════════════════════════════════════════════════

' 목적   : UNC 경로를 매핑된 드라이브 문자로 치환
Private Function prv_ProcessNetworkDrives(ByVal strPath As String, _
                                          ByVal strCwbPath As String) As String

    On Error Resume Next

    Dim objNetwork As Object
    Dim objDrives  As Object
    Dim i          As Integer

    Set objNetwork = CreateObject("WScript.Network")
    If objNetwork Is Nothing Then GoTo CleanUp

    Set objDrives = objNetwork.EnumNetworkDrives
    If objDrives Is Nothing Then GoTo CleanUp

    For i = 0 To objDrives.Count - 1 Step 2
        If objDrives.Item(i + 1) <> "" Then
            strCwbPath = Replace(strCwbPath, objDrives.Item(i + 1), _
                                 objDrives.Item(i), , , vbTextCompare)
            strPath = Replace(strPath, objDrives.Item(i + 1), _
                              objDrives.Item(i), , , vbTextCompare)
        End If
    Next i

CleanUp:
    Set objDrives = Nothing
    Set objNetwork = Nothing

    prv_ProcessNetworkDrives = strPath
    On Error GoTo 0

End Function

' 목적   : OneDrive/SharePoint URL 을 로컬 경로로 변환
Private Function prv_ProcessCloudPaths(ByVal strPath As String) As String

    On Error Resume Next

    If prv_IsValidPath(strPath) Then
        prv_ProcessCloudPaths = strPath
        Exit Function
    End If

    Dim blnIsUrl As Boolean
    blnIsUrl = (InStr(strPath, "http") > 0 Or _
                InStr(strPath, "sharepoint.com") > 0 Or _
                InStr(strPath, "d.docs.live.net") > 0)

    If blnIsUrl Then
        Dim strDecoded   As String
        Dim strConverted As String
        strDecoded = prv_URLDecode(strPath)
        strConverted = prv_TryOneDriveConversions(strDecoded)
        If strConverted <> "" Then
            prv_ProcessCloudPaths = strConverted
            Exit Function
        End If
    End If

    If InStr(strPath, "\\") = 1 Then
        Dim strUncConverted As String
        strUncConverted = prv_TryUncToOneDrive(strPath)
        If strUncConverted <> "" Then
            prv_ProcessCloudPaths = strUncConverted
            Exit Function
        End If
    End If

    prv_ProcessCloudPaths = strPath
    On Error GoTo 0

End Function

' 목적   : 슬래시(/) 를 백슬래시(\) 로 통일
Private Function prv_NormalizePath(ByVal strPath As String) As String
    prv_NormalizePath = Replace(strPath, "/", "\")
End Function

' 목적   : 드라이브 없거나 다를 때 CWB 드라이브로 보정
Private Function prv_ProcessDriveMapping(ByVal strPath As String, _
                                         ByVal strCwbPath As String, _
                                         ByVal blnChangeDrive As Boolean, _
                                         ByVal blnAddDrive As Boolean) As String

    On Error Resume Next

    Dim strCwbDrive As String
    strCwbDrive = Left(strCwbPath, InStr(strCwbPath, ":"))

    If blnChangeDrive Then
        If InStr(strPath, ":") > 0 Then
            strPath = strCwbDrive & Mid(strPath, InStr(strPath, ":") + 1)
        End If
    End If

    If blnAddDrive Then
        If InStr(strPath, ":") = 0 Then
            strPath = strCwbDrive & strPath
        End If
    End If

    prv_ProcessDriveMapping = strPath
    On Error GoTo 0

End Function

' 목적   : 경로 유효성 검사, 실패 시 A~Z 드라이브 순서로 추정
Private Function prv_ValidateAndFixDrive(ByVal strPath As String) As String

    If prv_IsValidPath(strPath) Then
        prv_ValidateAndFixDrive = strPath
        Exit Function
    End If

    Dim strFixed As String
    strFixed = prv_ReplaceWithCwbDrive(strPath)
    If prv_IsValidPath(strFixed) Then
        prv_ValidateAndFixDrive = strFixed
        Exit Function
    End If

    strFixed = prv_TryAllDrives(strPath)
    If prv_IsValidPath(strFixed) Then
        prv_ValidateAndFixDrive = strFixed
        Exit Function
    End If

    prv_ValidateAndFixDrive = ""

End Function

' ══════════════════════════════════════════════════════════
'  PRIVATE - 보조 함수
' ══════════════════════════════════════════════════════════

Private Function prv_TryOneDriveConversions(ByVal strUrl As String) As String

    On Error Resume Next

    Dim colPaths     As Collection
    Dim vntPath      As Variant
    Dim strConverted As String

    Set colPaths = prv_GetOneDrivePaths()

    For Each vntPath In colPaths
        strConverted = prv_ConvertSharePointUrl(strUrl, CStr(vntPath))
        If strConverted <> "" And prv_IsValidPath(strConverted) Then
            prv_TryOneDriveConversions = strConverted
            Exit Function
        End If
        strConverted = prv_ConvertPersonalUrl(strUrl, CStr(vntPath))
        If strConverted <> "" And prv_IsValidPath(strConverted) Then
            prv_TryOneDriveConversions = strConverted
            Exit Function
        End If
    Next vntPath

    prv_TryOneDriveConversions = ""
    On Error GoTo 0

End Function

Private Function prv_TryUncToOneDrive(ByVal strPath As String) As String

    On Error Resume Next

    Dim colPaths     As Collection
    Dim vntPath      As Variant
    Dim vntMapping   As Variant
    Dim strConverted As String

    Dim arrMappings As Variant
    arrMappings = Array("\\my04004.waffice.org\wf20024")

    Set colPaths = prv_GetOneDrivePaths()

    For Each vntMapping In arrMappings
        For Each vntPath In colPaths
            strConverted = Replace(strPath, CStr(vntMapping), _
                                   CStr(vntPath), , , vbTextCompare)
            If strConverted <> strPath And prv_IsValidPath(strConverted) Then
                prv_TryUncToOneDrive = strConverted
                Exit Function
            End If
        Next vntPath
    Next vntMapping

    prv_TryUncToOneDrive = ""
    On Error GoTo 0

End Function

Private Function prv_GetOneDrivePaths() As Collection

    On Error Resume Next

    Dim colPaths As New Collection
    Dim strPath  As String
    Dim objWsh   As Object
    Dim strReg   As String
    Dim i        As Integer

    strPath = Environ("OneDrive")
    If strPath <> "" Then colPaths.Add strPath

    strPath = Environ("OneDriveCommercial")
    If strPath <> "" And Not prv_IsInCollection(colPaths, strPath) Then
        colPaths.Add strPath
    End If

    strPath = Environ("OneDriveConsumer")
    If strPath <> "" And Not prv_IsInCollection(colPaths, strPath) Then
        colPaths.Add strPath
    End If

    Set objWsh = CreateObject("WScript.Shell")

    For i = 1 To 5
        On Error Resume Next
        strReg = objWsh.RegRead( _
            "HKEY_CURRENT_USER\Software\Microsoft\OneDrive\Accounts\Personal_" & i & "\UserFolder")
        If strReg <> "" And Not prv_IsInCollection(colPaths, strReg) Then
            colPaths.Add strReg
        End If
        strReg = objWsh.RegRead( _
            "HKEY_CURRENT_USER\Software\Microsoft\OneDrive\Accounts\Business" & i & "\UserFolder")
        If strReg <> "" And Not prv_IsInCollection(colPaths, strReg) Then
            colPaths.Add strReg
        End If
        On Error GoTo 0
    Next i

    Set objWsh = Nothing
    Set prv_GetOneDrivePaths = colPaths
    On Error GoTo 0

End Function

Private Function prv_ConvertSharePointUrl(ByVal strUrl As String, _
                                          ByVal strOdPath As String) As String

    On Error Resume Next

    Dim strPart As String
    Dim lngPos  As Long

    lngPos = InStr(strUrl, "/Documents/")
    If lngPos > 0 Then
        strPart = Mid(strUrl, lngPos + 11)
        strPart = Replace(strPart, "/", "\")
        prv_ConvertSharePointUrl = strOdPath & "\" & strPart
        Exit Function
    End If

    prv_ConvertSharePointUrl = ""
    On Error GoTo 0

End Function

Private Function prv_ConvertPersonalUrl(ByVal strUrl As String, _
                                        ByVal strOdPath As String) As String

    On Error Resume Next

    Dim strPart As String
    Dim lngPos  As Long
    Dim lngNext As Long

    lngPos = InStr(strUrl, "d.docs.live.net/")
    If lngPos = 0 Then
        prv_ConvertPersonalUrl = ""
        Exit Function
    End If

    strPart = Mid(strUrl, lngPos + 16)
    lngNext = InStr(strPart, "/")

    If lngNext > 0 Then
        strPart = Mid(strPart, lngNext + 1)
        strPart = Replace(strPart, "/", "\")
        prv_ConvertPersonalUrl = strOdPath & "\" & strPart
        Exit Function
    End If

    prv_ConvertPersonalUrl = ""
    On Error GoTo 0

End Function

Private Function prv_URLDecode(ByVal strEncoded As String) As String
    strEncoded = Replace(strEncoded, "+", " ")
    strEncoded = Replace(strEncoded, "%20", " ")
    strEncoded = Replace(strEncoded, "%2F", "/")
    strEncoded = Replace(strEncoded, "%5C", "\")
    strEncoded = Replace(strEncoded, "%3A", ":")
    prv_URLDecode = strEncoded
End Function

Private Function prv_ReplaceWithCwbDrive(ByVal strPath As String) As String

    On Error Resume Next

    If InStr(strPath, ":") = 0 Then
        prv_ReplaceWithCwbDrive = strPath
        Exit Function
    End If

    prv_ReplaceWithCwbDrive = Left(am_Core.XlamPath, 2) & Mid(strPath, 3)

    On Error GoTo 0

End Function

Private Function prv_TryAllDrives(ByVal strPath As String) As String

    On Error Resume Next

    If InStr(strPath, ":") = 0 Then
        prv_TryAllDrives = strPath
        Exit Function
    End If

    Dim strWithoutDrive As String
    Dim strTestPath     As String
    Dim i               As Integer

    strWithoutDrive = Mid(strPath, 3)

    For i = 65 To 90
        strTestPath = Chr(i) & ":" & strWithoutDrive
        If prv_IsValidPath(strTestPath) Then
            prv_TryAllDrives = strTestPath
            Exit Function
        End If
    Next i

    prv_TryAllDrives = ""
    On Error GoTo 0

End Function

Private Function prv_IsValidPath(ByVal strPath As String) As Boolean

    On Error Resume Next

    If Len(Trim(strPath)) = 0 Then Exit Function
    If Len(strPath) > 260 Then Exit Function

    Dim strResult As String
    strResult = Dir(strPath, vbDirectory)

    If strResult <> "" Then
        prv_IsValidPath = True
    Else
        prv_IsValidPath = (Dir(strPath) <> "")
    End If

    On Error GoTo 0

End Function

Private Function prv_IsInCollection(ByVal colTarget As Collection, _
                                    ByVal strSearch As String) As Boolean

    On Error Resume Next

    Dim vnt As Variant
    For Each vnt In colTarget
        If StrComp(CStr(vnt), strSearch, vbTextCompare) = 0 Then
            prv_IsInCollection = True
            Exit Function
        End If
    Next vnt

    On Error GoTo 0

End Function
