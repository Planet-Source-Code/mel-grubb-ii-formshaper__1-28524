VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMRU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' cMRU - Most recently used file support class
'
' Supplies Most-Recently-Used list functionality for VB programs.
'
' Version   Date        User            Notes
'   1.0     02/05/01    Mel Grubb II    Initial Version
'   1.1     06/28/01    Mel Grubb II    Fixed collection keying to allow for spaces
'                                       in filenames.
'   1.2     08/29/01    Mel Grubb II    Eliminated dependency on cPreferences
'                                       which fixes the "Failed to delete..."
'                                       problem.
'===============================================================================
Option Explicit

'===============================================================================
' Constants
'===============================================================================
Private Const mc_strModuleID = "cMRU."
Private Const mc_strSectionKeyBase = "Software\"
Private Const mc_strMRUItems = "MRUItems"
Private Const mc_strMRUKey = "Recent File List"

'===============================================================================
' Private Member Variables
'===============================================================================
Private m_colMRUFiles As Collection
Private m_intMaxItems As Integer        ' The maximum bumber of items allowed in an MRU
Private m_intMRUStart As Integer        ' The array index at which we start our MRU
Private m_objMenu As Object             ' The menu array object
Private m_strRegSectionBase As String   ' Holds the registry key under which information is stored


'===============================================================================
' Add - Add an item to the MRU list
'===============================================================================
Public Sub Add(ByRef FileName As String)
    On Error GoTo ErrorHandler
    Dim strFileName As String
    Dim strKey As String
    
    strFileName = Trim$(FileName)
    If Len(strFileName) > 0 Then
        ' Replace spaces with underscores to generate key
        strKey = Replace(strFileName, " ", "_")
        
        ' Remove the item from the list if it is already present
        On Error Resume Next
        m_colMRUFiles.Remove strKey
        On Error GoTo 0

        If m_colMRUFiles.Count = 0 Then
            ' The list is empty, just add the item
            m_colMRUFiles.Add FileName, strKey
        Else
            ' There are items in the list, insert before the first item
            m_colMRUFiles.Add FileName, strKey, 1
        End If

        ' Now, make sure we don't have too many items
        While m_colMRUFiles.Count > m_intMaxItems
            m_colMRUFiles.Remove m_intMaxItems + 1
        Wend
        FillMenu
    End If
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Add('" & FileName & "')"

End Sub


'===============================================================================
' FillMenu - Build the menu items
'===============================================================================
Public Function FillMenu()
    On Error GoTo ErrorHandler
    Dim intIndex      As Integer
    Dim strFile  As Variant
    Dim intPos   As Integer
    
    ' Remove any existing entries
    On Error Resume Next ' We don't care if there are holes in the array
    For intIndex = m_objMenu.UBound To m_intMRUStart + 1 Step -1
        Unload m_objMenu(intIndex)
    Next
    
    On Error GoTo 0
    ' Display or hide bar as needed
    If m_objMenu(m_intMRUStart).Caption = "-" Then
        m_objMenu(m_intMRUStart).Visible = (m_colMRUFiles.Count <> 0)
    End If
    
    ' Add new entries
    For Each strFile In m_colMRUFiles
        intPos = m_objMenu.UBound + 1
        Load m_objMenu(intPos)
        m_objMenu(intPos).Caption = strFile
        m_objMenu(intPos).Visible = True
        m_objMenu(intPos).Enabled = True
    Next
    Exit Function

ErrorHandler:
    ProcessError Err, mc_strModuleID & "FillMenu()"

End Function


'===============================================================================
' Init - Set up the MRU list
'===============================================================================
Public Sub Init(MenuItem As Object)
    On Error GoTo ErrorHandler
    Dim intIndex As Integer
    Dim strTemp As String
    Dim objRegistry As cRegistry

    ' Set up the registry object
    Set objRegistry = New cRegistry
    With objRegistry
        ' Set defaults
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = m_strRegSectionBase & "Settings"
        .ValueKey = mc_strMRUItems
        .Default = 4
        m_intMaxItems = .Value

        ' Pull the file list from the registry
        .SectionKey = m_strRegSectionBase & mc_strMRUKey
        .ValueType = REG_SZ
        .Default = ""
        For intIndex = 1 To m_intMaxItems
            .ValueKey = intIndex
            strTemp = .Value
            If Len(strTemp) > 0 Then m_colMRUFiles.Add strTemp, Replace(strTemp, " ", "_")
        Next intIndex
    End With
    Set objRegistry = Nothing
    
    ' Memorize the starting point of the menu array as it stands now
    Set m_objMenu = MenuItem
    m_intMRUStart = MenuItem.UBound
    Exit Sub
    
ErrorHandler:
    ProcessError Err, mc_strModuleID & "Init(OBJECT)"

End Sub


'===============================================================================
' MaxItems - The maximum number of items the class will handle
'===============================================================================
Public Property Get MaxItems() As Integer
    On Error GoTo ErrorHandler
    
    MaxItems = m_intMaxItems
    Exit Property

ErrorHandler:
    ProcessError Err, mc_strModuleID & "MaxItems_Get()"

End Property

Public Property Let MaxItems(ByVal Items As Integer)
    On Error GoTo ErrorHandler
    
    m_intMaxItems = Items
    Exit Property

ErrorHandler:
    ProcessError Err, mc_strModuleID & "MasItems_Let(" & Items & ")"

End Property


'===============================================================================
' RegSectionBase - The base registry key under which settings will be saved
'===============================================================================
Public Property Get RegSectionBase() As String
    On Error GoTo ErrorHandler
    
    Trace mc_strModuleID & "RegSectionBase_Get()"
    RegSectionBase = m_strRegSectionBase
    Exit Property

ErrorHandler:
    ProcessError Err, mc_strModuleID & "RegSectionBase_Get()"

End Property

Public Property Let RegSectionBase(NewBase As String)
    On Error GoTo ErrorHandler
    
    Trace mc_strModuleID & "RegSectionBase_Let('" & NewBase & "')"
    m_strRegSectionBase = NewBase
    Exit Property

ErrorHandler:
    ProcessError Err, mc_strModuleID & "RegSectionBase_Let('" & NewBase & "')"

End Property
    

'===============================================================================
' Save - Save the MRU to the registry
'===============================================================================
Public Sub Save()
    On Error GoTo ErrorHandler
    Dim intIndex As Integer
    Dim strKey As String
    Dim objRegistry As cRegistry

    ' Persist the max number of items
    Set objRegistry = New cRegistry
    With objRegistry
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = m_strRegSectionBase & "Settings"
        .ValueKey = mc_strMRUItems
        .ValueType = REG_DWORD
        .Value = m_intMaxItems
        
        ' Clear existing contents
        .SectionKey = m_strRegSectionBase & mc_strMRUKey
        If .KeyExists Then .DeleteKey

        ' Write new file list
        .ValueType = REG_SZ
        For intIndex = 1 To m_colMRUFiles.Count
            .ValueKey = intIndex
            .Value = m_colMRUFiles(intIndex)
        Next intIndex
    End With
    Set objRegistry = Nothing
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Save"

End Sub


'===============================================================================
' Class_Initialize
'===============================================================================
Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    ' Set up member objects
    Set m_colMRUFiles = New Collection

    m_strRegSectionBase = mc_strSectionKeyBase & App.CompanyName & "\" & App.ProductName & "\"
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Class_Initialize()"

End Sub


'===============================================================================
' Class_Terminate
'===============================================================================
Private Sub Class_Terminate()
    On Error GoTo ErrorHandler
    
    ' Persist the current settings
    Save
    
    ' Destroy member objects
    Set m_colMRUFiles = Nothing
    Exit Sub

ErrorHandler:
    ProcessError Err, mc_strModuleID & "Class_Terminate()"

End Sub
