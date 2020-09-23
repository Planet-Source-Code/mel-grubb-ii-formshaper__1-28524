VERSION 5.00
Begin VB.Form dlgStatus 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1095
   ClientLeft      =   2760
   ClientTop       =   3465
   ClientWidth     =   2820
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   2820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picStatus 
      Height          =   255
      Left            =   60
      ScaleHeight     =   195
      ScaleWidth      =   2655
      TabIndex        =   0
      Top             =   300
      Width           =   2715
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1320
         TabIndex        =   1
         Top             =   0
         Width           =   315
      End
      Begin VB.Shape shpStatus 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Label lblMessage 
      Caption         =   "Press ESC to close preview form when finished."
      Height          =   375
      Left            =   60
      TabIndex        =   3
      Top             =   660
      Width           =   2655
   End
   Begin VB.Label lblCaption 
      Caption         =   "Progress..."
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   2655
   End
End
Attribute VB_Name = "dlgStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' dlgStatus - Status/Progress dialog
'
' Version   Date        User            Notes
'   1.0     10/02/01    Mel Grubb II    Initial Version
'==============================================================================
Option Explicit

'==============================================================================
' Member Constants
'==============================================================================
Private Const mc_strModuleID As String = "dlgStatus."

'==============================================================================
' Member Variables
'==============================================================================
Private m_intPercentage As Integer


'===============================================================================
' Percentage - Get/Let the status percentage
'
' Notes:
'===============================================================================
Public Property Get Percentage() As Integer
    On Error GoTo ErrorHandler
    
    Percentage = m_intPercentage
    Exit Sub
    
ErrorHandler:
    ProcessError Err, mc_strModuleID & "GetPercentage"

End Property

Public Property Let Percentage(ByVal iPercentage As Integer)
    On Error GoTo ErrorHandler
    
    m_intPercentage = iPercentage
    shpStatus.Width = (m_intPercentage / 100) * picStatus.ScaleWidth
    lblStatus.Caption = iPercentage & "%"
    lblStatus.Left = (picStatus.ScaleWidth - lblStatus.Width) / 2
    picStatus.Refresh
    Exit Property
    
ErrorHandler:
    ProcessError Err, mc_strModuleID & "LetPercentage"

End Property


'===============================================================================
' Reset - Reset the status text and percentage to nothing
'
' Arguments: None
'
' Notes:
'===============================================================================
Public Sub Reset()
    On Error GoTo ErrorHandler
    
    lblStatus = ""
    Me.Percentage = 0
    Exit Sub
    
ErrorHandler:
    ProcessError Err, mc_strModuleID & "Reset"

End Sub


'===============================================================================
' Text - Change the status caption
'
' Arguments:
'   Value - The new text value for the progress dialog.
'
' Notes:
'===============================================================================
Public Property Let Text(ByVal Value As String)
    On Error GoTo ErrorHandler

    lblCaption.Caption = Value
    Exit Property

ErrorHandler:
    ProcessError Err, mc_strModuleID & "LetText"

End Property

Public Property Get Text() As String
    On Error GoTo ErrorHandler
    
    Caption = lblCaption.Caption
    Exit Sub
    
ErrorHandler:
    ProcessError Err, mc_strModuleID & "GetText"

End Property
