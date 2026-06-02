Attribute VB_Name = "tpl_Media"
Option Explicit
Sub GetVideoDuration()
    Dim wmp As Object
    Dim media As Object
    Dim filePath As String
    Dim videoLength As Double
    
    ' 비디오 파일 경로 설정
    filePath = "K:\002 건물유지관리\02 EDU LMS 운영\통합관리\00 영상자료\건물\BC1000_건물관리란.mp4" ' 파일 경로를 수정하세요.
    
    ' Windows Media Player 객체 생성
    Set wmp = CreateObject("WMPlayer.OCX")
    
    ' 미디어 파일 로드
    Set media = wmp.MediaCollection.Add(filePath)
    
    ' 미디어 길이 확인
    videoLength = media.Duration ' 초 단위로 반환됨
    
    ' 분과 초로 변환하여 결과 표시
    MsgBox "Video Duration: " & Int(videoLength \ 60) & " min " & Int(videoLength Mod 60) & " sec"
    
    ' 객체 정리
    Set media = Nothing
    Set wmp = Nothing
End Sub
Function GetVideoLength(ByVal Path As String) As String '영상 파일의 길이를 확인
    Dim objShell As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim selectedCell As Range
    
    ' Shell 객체 생성
    Set objShell = CreateObject("Shell.Application")
        
    ' 파일이 존재하는지 확인
    If Dir(Path) = "" Then
        MsgBox "파일을 찾을 수 없습니다: " & Path, vbExclamation
        Exit Function
    End If
    
    ' 파일 확장자 확인
    If LCase(Right(Path, 3)) <> "mp4" And _
       LCase(Right(Path, 3)) <> "avi" And _
       LCase(Right(Path, 3)) <> "mov" And _
       LCase(Right(Path, 3)) <> "wmv" Then
        MsgBox "지원되지 않는 파일 형식입니다.", vbExclamation
        Exit Function
    End If
    
    ' 파일 정보 가져오기
    Set objFolder = objShell.Namespace(Left(Path, InStrRev(Path, "\")))
    Set objFile = objFolder.ParseName(Mid(Path, InStrRev(Path, "\") + 1))
    
    ' 영상 길이 표시 (27번 속성이 길이를 나타냄)
    GetVideoLength = objFolder.GetDetailsOf(objFile, 27)
    
    ' 객체 해제
    Set objFile = Nothing
    Set objFolder = Nothing
    Set objShell = Nothing
End Function



