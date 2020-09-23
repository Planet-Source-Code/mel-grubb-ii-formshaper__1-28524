VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' cRegion - Windows region support class
'
' Provides support for creating, persisting, and applying region definitions to
' objects.
'
' Version   Date        User            Notes
'   1.0     10/04/01    Mel Grubb II    Initial Version
'===============================================================================
Option Explicit

'===============================================================================
' API Declarations
'===============================================================================
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As Long
Private Declare Function GetRegionData Lib "gdi32" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As Any) As Long
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Const RGN_DIFF = 4

'==============================================================================
' Member constants
'==============================================================================
Private Const c_intMaxPath = 260

'==============================================================================
' Member variables
'==============================================================================
Private m_lngRegion As Long

'===============================================================================
' Events
'===============================================================================
Public Event Progress(Percentage As Integer)


'===============================================================================
' Apply - Applies a region to a window
'
' Arguments:
'   hWnd - Handle to the window to receive the region
'
' Notes:
'===============================================================================
Public Sub Apply(ByVal hWnd As Long)
    SetWindowRgn hWnd, m_lngRegion, True
End Sub


'===============================================================================
' Clone - Returns a copy of this instance
'
' Notes: Applying a region to a window actually consumes or destroys it
'===============================================================================
Public Function Clone() As cRegion
    Dim objRegion As cRegion
    
    Set objRegion = New cRegion
    objRegion.RegionFromArray RegionToArray
    Set Clone = objRegion
End Function


'===============================================================================
' hRgn - Get/Set the handle to a region
'
' Notes: Use this with extreme caution.  When an instance of this class goes out
' of scope, it will delete the region data from memory.  If you set this to an
' existing region, it will be destroyed when this instance falls out or is
' set to Nothing.
'===============================================================================
Public Property Get hRgn() As Long
    hRgn = m_lngRegion
End Property

Public Property Let hRgn(Handle As Long)
    m_lngRegion = Handle
End Property


'===============================================================================
' RegionFromArray - Creates a region from a byte array and returns a handle to it
'
' Arguments:
'   RegionData - Byte array containing region data
'===============================================================================
Public Sub RegionFromArray(RegionData() As Byte)
    If m_lngRegion <> 0 Then DeleteObject m_lngRegion
    m_lngRegion = ExtCreateRegion(ByVal 0&, UBound(RegionData) + 1, RegionData(0))
End Sub


'===============================================================================
' RegionFromFile - Load region data from disk
'
' Arguments:
'   FileName - Path and file name of the file to load region data from
'
' Notes:
'===============================================================================
Public Sub RegionFromFile(FileName As String)
    Dim abytRegion() As Byte
    Dim intFile As Integer
    Dim lngLength As Long
    
    lngLength = FileLen(FileName)
    If lngLength > 0 Then
        If m_lngRegion <> 0 Then DeleteObject m_lngRegion
        
        ReDim abytRegion(lngLength)
        intFile = FreeFile
        Open FileName For Binary Access Read As #intFile
        Get #intFile, , abytRegion
        Close #intFile

        RegionFromArray abytRegion
    End If
End Sub


'===============================================================================
' RegionFromPicture - Creates a region based on a picture and a transparency
' color
'
' Arguments:
'   Source - The picture containing the picture
'   Color - The color to make transparent
'
' Notes:
'===============================================================================
Public Sub RegionFromPicture(Source As PictureBox, MaskColor As Long)
    Dim lngRegion As Long       ' Working region
    Dim lngRegionTemp As Long   ' Temporary region
    Dim lngRun As Long          ' Length of run
    Dim lngHeight As Long       ' Picture height
    Dim lngWidth As Long        ' Picture width
    Dim lngX As Long            ' Current X
    Dim lngY As Long            ' Current Y
    Dim intProgress As Integer  ' For reporting percent complete
    Dim intProgress2 As Integer
    
    If Source.Picture <> 0 Then
        If m_lngRegion <> 0 Then DeleteObject m_lngRegion
        With Source
            lngHeight = .ScaleHeight
            lngWidth = .ScaleWidth
            lngRegion = CreateRectRgn(0, 0, lngWidth, lngHeight)
    
            ' Check for horizontal runs of transparent pixels
            For lngY = 0 To lngHeight
                For lngX = 0 To lngWidth
                    If GetPixel(.hDC, lngX, lngY) = MaskColor Then
                        ' Increase the length of the run
                        lngRun = lngRun + 1
                    Else
                        ' Indicate there is a run to be processed
                        lngRun = -lngRun
                    End If
    
                    If (lngRun < 0) Or (lngX = lngWidth) Then
                        ' We've just ended a run, subtract it from the region
                        lngRegionTemp = CreateRectRgn(lngX - Abs(lngRun), lngY, lngX, lngY + 1)
                        CombineRgn lngRegion, lngRegion, lngRegionTemp, RGN_DIFF
                        DeleteObject lngRegionTemp
    
                        ' Start a new run
                        lngRun = 0
                    End If
                Next lngX
    
                intProgress = ((lngY * 100) \ lngHeight)
                If intProgress <> intProgress2 Then
                    RaiseEvent Progress(intProgress2)
                    intProgress2 = intProgress
                End If
            Next lngY
        End With
        m_lngRegion = lngRegion
    End If
End Sub


'===============================================================================
' RegionFromResource - Load a region definition from the resource and return a
' handle to it.
'
' Arguments:
'   ResID - Resource ID
'   ResType - Resource type ID
'
' Notes:
'===============================================================================
Public Sub RegionFromResource(ResID As Integer, ResType As String)
    Dim abytRegion() As Byte

    ' Pull the region data from the resource
    abytRegion = LoadResData(ResID, ResType)
    If UBound(abytRegion) > 0 Then
        If m_lngRegion <> 0 Then DeleteObject m_lngRegion
        m_lngRegion = ExtCreateRegion(ByVal 0&, UBound(abytRegion) + 1, abytRegion(0))
    End If
End Sub


'===============================================================================
' RegionToArray - Retrieve the byte array containing the region data
'
' Arguments:
'   hRgn - Handle to the region in memory
'
' Notes:
'===============================================================================
Public Function RegionToArray() As Byte()
    Dim abytRegion() As Byte
    Dim lngLength As Long

    ' Get the length of the region data
    lngLength = GetRegionData(m_lngRegion, 0, ByVal 0&)
    If lngLength > 0 Then
        ' Fill the byte array with the region data
        ReDim abytRegion(lngLength - 1)
        If lngLength <> GetRegionData(hRgn, lngLength, abytRegion(0)) Then Erase abytRegion
        RegionToArray = abytRegion
    End If
End Function


'===============================================================================
' RegionToFile - Save region data to a file on disk
'
' Arguments:
'   hRgn - Handle to the region data
'   FileName - File to save the region data in
'
' Notes:
'===============================================================================
Public Sub RegionToFile(ByRef FileName As String)
    Dim intFile As Integer
    Dim abytRegion() As Byte
    
    abytRegion = RegionToArray
    If UBound(abytRegion) > 0 Then
        On Error Resume Next
        Kill FileName
        On Error GoTo 0
        
        intFile = FreeFile
        Open FileName For Binary Access Write Lock Read As #intFile
        Put #intFile, , abytRegion
        Close #intFile
    End If
End Sub


'===============================================================================
' Class_Terminate - Destroy outstanding object references
'===============================================================================
Private Sub Class_Terminate()
    If m_lngRegion <> 0 Then DeleteObject m_lngRegion
End Sub
