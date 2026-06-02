VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_IndividualCourseStatus 
   Caption         =   "개인별 수강 현황"
   ClientHeight    =   9570.001
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   5892
   OleObjectBlob   =   "frm_IndividualCourseStatus.frx":0000
End
Attribute VB_Name = "frm_IndividualCourseStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0 '포지션 수동 설정
        .Move Application.Left + (Application.Width - .Width) / 2, Application.Top + (Application.Height - .Height) / 2 ''폼의 위치를 중앙으로 이동
    End With
End Sub
