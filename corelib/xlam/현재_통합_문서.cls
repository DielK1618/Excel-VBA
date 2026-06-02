VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "현재_통합_문서"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ── 1. xlam 열림 이벤트 ─────────────────────────────────────
Private Sub Workbook_Open()
    Call am_Core.Initialize
End Sub

' ── 2. xlam 닫힘 이벤트 ─────────────────────────────────────
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call am_Core.CleanUp
End Sub
