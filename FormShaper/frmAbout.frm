VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3555
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4800
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   237
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1980
      Left            =   240
      Picture         =   "frmAbout.frx":000C
      ScaleHeight     =   132
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   1
      Top             =   240
      Width           =   420
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3345
      TabIndex        =   0
      Top             =   2625
      Width           =   1380
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   3360
      TabIndex        =   2
      Top             =   3075
      Width           =   1365
   End
   Begin VB.Label lblCopyright 
      Caption         =   "lblCopyright (Displays App.LegalCopyright)"
      Height          =   825
      Left            =   255
      TabIndex        =   6
      Top             =   2625
      Width           =   2970
   End
   Begin VB.Label lblDescription 
      Caption         =   "lblDescription (Displays App.Comments)"
      Height          =   1050
      Left            =   750
      TabIndex        =   5
      Top             =   1125
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   6
      X2              =   316
      Y1              =   163
      Y2              =   163
   End
   Begin VB.Label lblTitle 
      Caption         =   "lblTitle (Displays App.ProductName)"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   750
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   7
      X2              =   316
      Y1              =   164
      Y2              =   164
   End
   Begin VB.Label lblVersion 
      Caption         =   "lblVersion (Displays App.Major, .Minor, .Revision)"
      Height          =   225
      Left            =   750
      TabIndex        =   4
      Top             =   780
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' frmAbout - Project About form
'
' (c) 2001 Mel Grubb II
'
' Version   Date        User            Notes
'   1.0     02/06/01    Mel Grubb II    Initial version
'                                       Adapted from standard VB About box to
'                                       use cRegistry for registry access
'===============================================================================
Option Explicit

'===============================================================================
' Member Constants
'===============================================================================
Private Const mc_strModuleID = "frmAbout."
Private Const mc_strREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Private Const mc_strREGVALSYSINFOLOC = "MSINFO"
Private Const mc_strREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Private Const mc_strREGVALSYSINFO = "PATH"


'===============================================================================
' cmdSysInfo_Click - Display the system info application
'
' Arguments: None
'
' Returns: None
'===============================================================================
Private Sub cmdSysInfo_Click()
    On Error GoTo ErrorHandler
    
    StartSysInfo
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "cmdSysInfo_Click()"

End Sub


'===============================================================================
' cmdOK_Click - Close the about form
'
' Arguments: None
'
' Returns: None
'===============================================================================
Private Sub cmdOK_Click()
    On Error GoTo ErrorHandler
    
    Unload Me
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "cmdOK_Click()"

End Sub


'===============================================================================
' Form_Load - Set up the About form for display
'
' Arguments: None
'
' Returns: None
'===============================================================================
Private Sub Form_Load()
    On Error GoTo ErrorHandler

    ' Configure the about form
    Me.Caption = "About " & App.Title
    Me.Icon = frmMain.Icon
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.ProductName
    lblDescription.Caption = App.Comments
    lblCopyright.Caption = App.LegalCopyright
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Form_Load()"

End Sub


'===============================================================================
' StartSysInfo - Open the system information application
'
' Arguments: None
'
' Returns: None
'===============================================================================
Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
    Dim objRegistry As cRegistry
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    Set objRegistry = New cRegistry
    With objRegistry
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = mc_strREGKEYSYSINFO
        .ValueKey = mc_strREGVALSYSINFO
        SysInfoPath = .Value
        If Len(SysInfoPath) = 0 Then
            ' Try To Get System Info Program Path Only From Registry...
            .SectionKey = mc_strREGKEYSYSINFOLOC
            .ValueKey = mc_strREGVALSYSINFOLOC
            SysInfoPath = .Value
            
            ' Validate Existance Of Known 32 Bit File Version
            If (Len(Dir$(SysInfoPath & "\MSINFO32.EXE")) > 0) Then
                ' File found
                SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            Else
                ' Error - File Can Not Be Found...
                GoTo SysInfoErr
            End If
        End If
    End With
    Set objRegistry = Nothing
    
    ' Now launch the program
    Shell SysInfoPath, vbNormalFocus
    Exit Sub
    
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly

End Sub
