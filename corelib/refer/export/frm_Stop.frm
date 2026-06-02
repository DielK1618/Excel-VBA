VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Stop 
   ClientHeight    =   870
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   1944
   OleObjectBlob   =   "frm_Stop.frx":0000
End
Attribute VB_Name = "frm_Stop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub UserForm_Initialize()
    '<<<<< Æû À§Ä¡
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    '>>>>> Æû À§Ä¡
End Sub
Private Sub cmd_Stop_Click()
    blnStop = True
    Unload Me
End Sub

