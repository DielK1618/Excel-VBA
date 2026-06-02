VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Common"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Public cl As New COMMON << 일반 모듈에 삽입
'Public wsPub As Worksheet << 활성시트 외에 다른 시트를 잠글 때 사용, sht_Lock와 연동하여 사용
Sub DPUpdate_Off()
    Application.ScreenUpdating = False    '진행과정 숨기기 / True = 보이기
    Application.DisplayAlerts = False    '경고창 안뜨게 하기 / True = 뜨게 하기
End Sub
Sub DPUpdate_On()
    Application.ScreenUpdating = True    '진행과정 숨기기 / True = 보이기
    Application.DisplayAlerts = True    '경고창 안뜨게 하기 / True = 뜨게 하기
End Sub
Sub Event_Off()
    Application.EnableEvents = False
End Sub
Sub Event_On()
    Application.EnableEvents = True
End Sub
Sub Calculate_Off()
    Application.Calculation = xlCalculationManual '수식계산 수동
End Sub
Sub Calculate_On()
    Application.Calculation = xlCalculationAutomatic
End Sub
Sub sht_Lock()
    If blnPartEvents = False Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = IIf(wsPub Is Nothing, ActiveSheet, wsPub)
    
    If InStr(ws.Name, "DB_") = 0 And InStr(ws.Name, "H_") = 0 And InStr(ws.Name, "UL_") = 0 Then
    
    Dim blnScreenUpdating As Boolean
    blnScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim usedRng As Range
    Set usedRng = GetUsedRange(ws)  ' [개선] 사용 범위 변수로 저장 (FindCellsByColor에서도 재사용)

    ws.Unprotect Password:=PW
    usedRng.Locked = True           ' [개선] 전체 셀 대신 사용 범위만 잠금
    usedRng.FormulaHidden = True    ' [개선] 전체 셀 대신 사용 범위만 숨김

    Dim unlockRng As Range
    Set unlockRng = FindCellsByColor(xlNone, usedRng, True)  ' [개선] tmpRng 제거, 직접 대입
    
    If Not unlockRng Is Nothing Then
        unlockRng.Locked = False
        unlockRng.FormulaHidden = False
    End If

    With ws
        .Protect Password:=PW, Contents:=True, Scenarios:=True, _
                 AllowSorting:=True, AllowFiltering:=True
        .EnableSelection = xlNoRestrictions
    End With

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = blnScreenUpdating
    Application.FindFormat.Clear
    
    End If
End Sub
Sub sht_UnLock()
    If wsPub Is Nothing Then Set wsPub = ActiveSheet
    wsPub.Unprotect Password:=PW    '시트잠금풀기
    Set wsPub = Nothing
End Sub
Sub WB_Lock()
    If blnPartEvents = False Then Exit Sub
    ActiveWorkbook.Protect Password:=PW, Structure:=True
End Sub
Sub WB_UNLock()
    ActiveWorkbook.Unprotect Password:=PW
End Sub
