VERSION 5.00
Begin VB.Form frmPreview 
   BorderStyle     =   0  'None
   ClientHeight    =   2310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2940
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2310
   ScaleWidth      =   2940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' frmPreview - Shaped form preview form.
'
' Version   Date        User            Notes
'   1.0     10/02/01    Mel Grubb II    Initial Version
'   1.1     10/29/01    Mel Grubb II    Eliminated PictureBox control
'===============================================================================
Option Explicit

'===============================================================================
' API Declarations
'===============================================================================
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

'===============================================================================
' Member Constants
'===============================================================================
Private Const mc_strModuleID = "frmPreview."


'===============================================================================
' Form_KeyPress
'===============================================================================
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandler
    
    If KeyAscii = vbKeyEscape Then Unload Me
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Form_KeyPress(" & KeyAscii & ")"

End Sub


'===============================================================================
' Form_MouseDown
'===============================================================================
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrorHandler
    
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
    End If
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Form_MouseDown(" & Button & ", " & Shift & ", " & X & ", " & Y & ")"

End Sub
