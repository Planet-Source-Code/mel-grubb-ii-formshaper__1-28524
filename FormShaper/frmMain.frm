VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000C&
   Caption         =   "frmMain"
   ClientHeight    =   3060
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3825
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmMain.frx":0000
   ScaleHeight     =   3060
   ScaleWidth      =   3825
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2580
      Left            =   60
      MouseIcon       =   "frmMain.frx":0152
      ScaleHeight     =   170
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   0
      Top             =   420
      Width           =   3700
   End
   Begin MSComctlLib.ImageList imgToolbar 
      Left            =   3180
      Top             =   -240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":02A4
            Key             =   "OpenPicture"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0400
            Key             =   "OpenRegion"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":055C
            Key             =   "SaveRegion"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":06B8
            Key             =   "Pipette"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0820
            Key             =   "Preview"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbStandard 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "imgToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OpenPicture"
            Object.ToolTipText     =   "Open picture"
            ImageKey        =   "OpenPicture"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OpenRegion"
            Object.ToolTipText     =   "Open region file"
            ImageKey        =   "OpenRegion"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SaveRegion"
            Object.ToolTipText     =   "Save region file"
            ImageKey        =   "SaveRegion"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Preview"
            Object.ToolTipText     =   "Preview"
            ImageKey        =   "Preview"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pipette"
            Object.ToolTipText     =   "Choose mask color"
            ImageKey        =   "Pipette"
         EndProperty
      EndProperty
      Begin VB.PictureBox picMaskColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF00FF&
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1905
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   2
         Top             =   45
         Width           =   240
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpenPicture 
         Caption         =   "&Open Picture"
      End
      Begin VB.Menu mnuFileOpenRegion 
         Caption         =   "Open &Region file"
      End
      Begin VB.Menu mnuFileSaveRegion 
         Caption         =   "&Save region data"
      End
      Begin VB.Menu mnuFileSaveRegionAs 
         Caption         =   "Save region data &As"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsSelectMaskColor 
         Caption         =   "&Select mask color"
      End
      Begin VB.Menu mnuToolsPreview 
         Caption         =   "&Preview"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' frmMain - Form Shaper main form
'
' Version   Date        User            Notes
'   1.0     10/02/01    Mel Grubb II    Initial Version
'   1.1     10/29/01    Mel Grubb II    Cleaned up and commented code
'===============================================================================
Option Explicit

'===============================================================================
' API Declarations
'===============================================================================
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

'===============================================================================
' Member constants
'===============================================================================
Private Const mc_strModuleID = "frmMain."
Private Const mc_strRegionFilter As String = "Region files (*.rgn)|*.rgn|All Files (*.*)|*.*"
Private Const mc_strImageFilter As String = "Picture Files|*.bmp;*.dib;*.gif;*.jpg;*.ico;*.cur|Bitmaps (*.bmp)|*.bmp|GIF Images (*.gif)|*.gif|JPEG Images (*.jpg)|*.jpg|Icons (*.ico; *.cur)|*.ico;*.cur|All Files(*.*)|*.*"
Private Const mc_strCaption As String = "Form Shaper"

'===============================================================================
' Member variables
'===============================================================================
Private m_lngMaskColor As Long              ' Currently selected mask color
Private m_objMRU As cMRU                    ' Most-recently-used list handler
Private WithEvents m_objRegion As cRegion   ' The reason we're here
Attribute m_objRegion.VB_VarHelpID = -1
Private m_strFileName As String             ' Current mask filename


'===============================================================================
' GenerateRegion - Generate region data based on the current picture
'
' Arguments: None
'
' Notes:
'===============================================================================
Private Sub GenerateRegion()
    On Error GoTo ErrorHandler

    If Picture1.Picture <> 0 Then
        With dlgStatus
            .Text = "Generating region data"
            .Percentage = 0
            .Show , Me
            DoEvents
        End With
        
        m_objRegion.RegionFromPicture Picture1, picMaskColor.BackColor
        Unload dlgStatus
    End If
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "GenerateRegion()"

End Sub


'===============================================================================
' Form_KeyPress
'===============================================================================
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandler

    If (KeyAscii = vbKeyEscape) And (Picture1.MousePointer = vbCustom) Then
        picMaskColor.BackColor = m_lngMaskColor
        Picture1.MousePointer = vbDefault
    End If
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Form_KeyPress(" & KeyAscii & ")"

End Sub


'===============================================================================
' Form_Load - The form is about to be shown for the first time
'===============================================================================
Private Sub Form_Load()
    On Error GoTo ErrorHandler

    ' Set up the most-recently-used file list
    Set m_objMRU = New cMRU
    With m_objMRU
        .Init mnuFileMRU
        .FillMenu
    End With

    ' Set up the region we will be using
    Set m_objRegion = New cRegion
    
    EnableControls
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Form_Load()"

End Sub


'===============================================================================
' Form_Resize - The form size is being changed, resize controls
'
' Arguments: None
'
' Notes:
'===============================================================================
Private Sub Form_Resize()
    On Error GoTo ErrorHandler
    Dim lngX As Long
    Dim lngY As Long
    
    With Picture1
        lngX = (Me.ScaleWidth - .Width) / 2
        .Left = IIf(lngX > 0, lngX, 0)
        lngY = (Me.ScaleHeight - tlbStandard.Height - .Height) / 2
        .Top = IIf(lngY > 0, lngY + tlbStandard.Height, tlbStandard.Height)
        .Refresh
    End With
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Form_Resize()"

End Sub


'===============================================================================
' Form_Terminate - Destroy outstanding object references
'
' Arguments: None
'
' Notes:
'===============================================================================
Private Sub Form_Terminate()
    On Error GoTo ErrorHandler
    
    Set m_objMRU = Nothing
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Form_Terminate()"

End Sub


'===============================================================================
' Form_Unload
'
' Arguments:
'   Cancel - Integer that determines whether the form is removed from the screen.
'       If cancel is 0, the form is removed. Setting cancel to any nonzero value
'       prevents the form from being removed.
'
' Notes:
'===============================================================================
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorHandler
    
    m_objMRU.Save
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Form_Unload(" & Cancel & ")"

End Sub


'===============================================================================
' mnuFileExit_Click - Exit the application
'
' Arguments: None
'
' Notes:
'===============================================================================
Private Sub mnuFileExit_Click()
    On Error GoTo ErrorHandler
    
    Unload Me
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "mnuFileExit_Click()"

End Sub


'===============================================================================
' mnuFileMRU_Click - The user has clicked on an MRU entry, load it
'
' Arguments:
'   Index (IN) - The index number of the MRU entry that was clicked
'
' Notes:
'===============================================================================
Private Sub mnuFileMRU_Click(Index As Integer)
    On Error GoTo ErrorHandler
    
    OpenPicture mnuFileMRU(Index).Caption
    EnableControls
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "mnuFileMRU_Click(" & Index & ")"

End Sub


'===============================================================================
' mnuFileOpenPicture - Allow the user to select a picture to load
'
' Arguments: None
'
' Notes:
'===============================================================================
Private Sub mnuFileOpenPicture_Click()
    On Error GoTo ErrorHandler
    Dim objCDLG As New cCommonDialog
    Dim strFileName As String

    ' Load an image into the picturebox
    Set objCDLG = New cCommonDialog
    objCDLG.VBGetOpenFileName strFileName, , , , True, True, mc_strImageFilter, , App.Path
    OpenPicture strFileName
    
    ' Reset the region object
    Set m_objRegion = Nothing
    Set m_objRegion = New cRegion
    
    EnableControls
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "mnuFileOpenPicture_Click()"

End Sub


'===============================================================================
' mnuFileOpenRegion_Click - Allow the user to load a region file from disk
'
' Arguments: None
'
' Notes:
'===============================================================================
Private Sub mnuFileOpenRegion_Click()
    On Error GoTo ErrorHandler
    Dim objCDLG As cCommonDialog
    Dim strFileName As String

    Set objCDLG = New cCommonDialog
    objCDLG.VBGetOpenFileName strFileName, , , , , True, mc_strRegionFilter, , App.Path, , "*.rgn", Me.hWnd
    If Len(strFileName) > 0 Then m_objRegion.RegionFromFile strFileName
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "mnuFileOpenRegion_Click()"

End Sub


'===============================================================================
' mnuFileSaveRegion_Click - Allow the user to save region data to a file
'
' Arguments: None
'
' Notes:
'===============================================================================
Private Sub mnuFileSaveRegion_Click()
    On Error GoTo ErrorHandler
    
    If Len(m_strFileName) = 0 Then
        mnuFileSaveRegionAs_Click
    ElseIf m_objRegion.hRgn = 0 Then
        ' Region has not been generated yet
        GenerateRegion
        m_objRegion.RegionToFile m_strFileName
    End If

    EnableControls
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "mnuFileSaveRegion_Click()"

End Sub


'===============================================================================
' mnuFileSaveRegionAs_Click - Allow the user to save region data to a file
'
' Arguments: None
'
' Notes: Prompts for a filename as needed
'===============================================================================
Private Sub mnuFileSaveRegionAs_Click()
    On Error GoTo ErrorHandler
    Dim objCDLG As cCommonDialog
    Dim strFileName As String
    
    Set objCDLG = New cCommonDialog
    objCDLG.VBGetSaveFileName strFileName, , , mc_strRegionFilter, , App.Path, , "*.rgn", Me.hWnd
    If Len(strFileName) > 0 Then
        m_strFileName = strFileName
        mnuFileSaveRegion_Click
    End If
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "mnuFileSaveRegionAs_Click()"

End Sub


'===============================================================================
' mnuHelpAbout_Click - Show the about screen
'
' Arguments: None
'
' Notes:
'===============================================================================
Private Sub mnuHelpAbout_Click()
    On Error GoTo ErrorHandler

    frmAbout.Show vbModal, Me
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "mnuHelpAbout()"

End Sub


'===============================================================================
' mnuToolsPreview_Click - Show the preview form
'
' Arguments: None
'
' Notes:
'===============================================================================
Private Sub mnuToolsPreview_Click()
    On Error GoTo ErrorHandler
    
    ' Generate region if needed
    If m_objRegion.hRgn = 0 Then GenerateRegion

    With frmPreview
        .Picture = Picture1.Picture
        .Width = Picture1.ScaleWidth * Screen.TwipsPerPixelX
        .Height = Picture1.ScaleHeight * Screen.TwipsPerPixelY
        
        ' The temporary clone will automatically delete the region
        ' data when it terminates.
        m_objRegion.Clone.Apply .hWnd

        Me.Hide
        .Show vbModal, Me
        Me.Show
    End With
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "mnuToolsPreview_Click()"

End Sub


'===============================================================================
' mnuToolsSelectMaskColor
'
' Arguments: None
'
' Notes:
'===============================================================================
Private Sub mnuToolsSelectMaskColor_Click()
    On Error GoTo ErrorHandler
    
    If Picture1 <> 0 Then
        m_lngMaskColor = picMaskColor.BackColor
        Picture1.MousePointer = 99
    End If
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "mnuToolsSelectMaskColor_Click()"

End Sub


'===============================================================================
' m_objRegion_Progress - Update the progress display
'
' Arguments:
'   Percentage - Indicates the amount of the work that has been completed.
'
' Notes:
'===============================================================================
Private Sub m_objRegion_Progress(Percentage As Integer)
    On Error GoTo ErrorHandler
    
    dlgStatus.Percentage = Percentage
    Exit Sub
    
ErrorHandler:
    ProcessError Err, mc_strModuleID & "m_objRegion_Progress(" & Percentage & ")"

End Sub


'===============================================================================
' picMaskColor_Click - Allow the user to select a color using the common dialog
'
' Arguments: None
'
' Notes:
'===============================================================================
Private Sub picMaskColor_Click()
    On Error GoTo ErrorHandler
    Dim objCDLG As cCommonDialog
    Dim lngMaskColor As Long
    
    m_lngMaskColor = picMaskColor.BackColor
    lngMaskColor = m_lngMaskColor
    Set objCDLG = New cCommonDialog
    objCDLG.VBChooseColor lngMaskColor, , True, , Me.hWnd
    If lngMaskColor > 0 Then
        picMaskColor.BackColor = lngMaskColor

        ' Reset the region object
        Set m_objRegion = Nothing
        Set m_objRegion = New cRegion
    End If
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "picMaskColor_Click()"

End Sub


'===============================================================================
' Picture1_MouseDown - The user has clicked inside the picture box
'
' Arguments:
'   Button - An integer that identifies the button that was pressed to cause the
'       event.
'   Shift - An integer that corresponds to the state of the SHIFT, CTRL, and ALT
'       keys when the button specified in the button argument is pressed.
'   X, Y - Specifies the current location of the mouse pointer.
'
' Notes: We apply a clone of the region to the preview form because the act of
' using a region seems to consume it.  I have not looked into the cause for this
' yet, but it definitely appears that you can only use a region once.
'===============================================================================
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrorHandler
    Dim intX As Integer
    Dim intY As Integer
    Dim lngTemp As Long

    With Picture1
        If .MousePointer = vbCustom Then
            ' We are in "pipette" mode, get the color
            picMaskColor.BackColor = GetPixel(.hDC, X, Y)
            .MousePointer = vbDefault
            
            ' Reset the region object
            Set m_objRegion = Nothing
            Set m_objRegion = New cRegion
        End If
    End With
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Picture1_MouseDown(" & Button & ", " & Shift & ", " & X & ", " & Y & ")"

End Sub


'===============================================================================
' Picture1_MouseMove - The mouse is moving over the picture.
'
' Arguments:
'   Button - An integer that corresponds to the state of the mouse buttons.
'   Shift - An integer that corresponds to the state of the SHIFT, CTRL, and ALT
'       keys.
'   X, Y - Specifies the current location of the mouse pointer.
'
' Notes: I am using this event to sync up the mask color on the toolbar.
'===============================================================================
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrorHandler
    
    If Picture1.MousePointer = vbCustom Then picMaskColor.BackColor = GetPixel(Picture1.hDC, X, Y)
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Picture1_MouseMove(" & Button & ", " & Shift & ", " & X & ", " & Y & ")"

End Sub


'===============================================================================
' tbStandard_ButtonClick - The user has clicked a toolbar button
'
' Arguments:
'   Button - A reference to the clicked Button object.
'
' Notes:
'===============================================================================
Private Sub tlbStandard_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo ErrorHandler
    
    Select Case Button.Key
        Case "OpenPicture"
            mnuFileOpenPicture_Click
        
        Case "OpenRegion"
            mnuFileOpenRegion_Click
        
        Case "SaveRegion"
            mnuFileSaveRegion_Click
        
        Case "Preview"
            mnuToolsPreview_Click

        Case "Pipette"
            mnuToolsSelectMaskColor_Click
    
    End Select
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "tlbStandard_ButtonClick(" & Button.Key & ")"

End Sub


'===============================================================================
' EnableControls - Enable/Disable controls as appropriate
'
' Arguments: None
'
' Notes:
'===============================================================================
Private Sub EnableControls()
    On Error GoTo ErrorHandler
    Dim blnPicture As Boolean

    blnPicture = Not (Picture1.Picture = 0)
    Picture1.Visible = blnPicture
    
    ' Enable/Disable menu items
    mnuFileSaveRegion.Enabled = blnPicture
    mnuFileSaveRegionAs.Enabled = blnPicture
    mnuToolsSelectMaskColor.Enabled = blnPicture
    mnuToolsPreview.Enabled = blnPicture

    ' Enable/Disable toolbar buttons
    With tlbStandard.Buttons
        .Item("SaveRegion").Enabled = blnPicture
        .Item("Preview").Enabled = blnPicture
        .Item("Pipette").Enabled = blnPicture
    End With
    
    ' Set caption as appropriate
    Me.Caption = App.Title & IIf(Len(m_strFileName) > 0, " [" & m_strFileName & "]", "")
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "EnableControls()"

End Sub


'===============================================================================
' OpenPicture - Open a file, and load it into Picture1
'
' Arguments:
'   FileName - The full path and filename to load
'
' Notes:
'===============================================================================
Public Sub OpenPicture(FileName As String)
    On Error GoTo ErrorHandler

    If Len(FileName) > 0 Then
        Picture1.Picture = LoadPicture(FileName)
        m_objMRU.Add FileName
        Form_Resize
    End If
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "OpenPicture('" & FileName & "')"

End Sub
