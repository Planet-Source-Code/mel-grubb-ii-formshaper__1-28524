VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'===============================================================================
' cRegistry - Registry support class
' A nice class wrapper around the registry functions
' Allows searching,deletion,modification and addition
' of Keys or Values.
'
' Version   Date        User            Notes
'   1.0     02/21/97    Steve McMahon   Initial version
'   1.1     04/29/98    Steve McMahon   Fixed GPF in EnumerateValues
'                                       Added support for all registry types, not just strings
'                                       Put all declares in local class
'                                       Added VB5 Enums
'                                       Added CreateKey and DeleteKey methods
'   1.2     01/02/99    Steve McMahon   The CreateExeAssociation method failed to set up the
'                                       association correctly if the optional document icon
'                                       was not provided.
'                                       Added new parameters to CreateExeAssociation to set up
'                                       other standard handlers: Print, Add, New
'                                       Provided the CreateAdditionalEXEAssociations method
'                                       to allow non-standard menu items to be added (for example,
'                                       right click on a .VBP file.  VB installs Run and Make
'                                       menu items).
'   1.3     02/08/00    Steve McMahon   Ensure CreateExeAssociation and related items sets up the
'                                       registry keys in the HKEY_LOCAL_MACHINE\SOFTWARE\Classes
'                                       branch as well as the HKEY_CLASSES_ROOT branch.
'   1.4     02/05/01    Mel Grubb II    Fixed CreateAdditionalEXEAssociations
'                                       It was seaching for existing keys in the wrong place
'   1.5     02/07/01    Mel Grubb II    Fixed CreateExeAssociation DefaultIcon to eliminate extra
'                                       quotes around executable name.
'   1.6     08/29/01    Mel Grubb II    Added ValueExists property
'===============================================================================
Option Explicit


'===============================================================================
' Member Constants
'===============================================================================
'Registry Specific Access Rights
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_ALL_ACCESS = &H3F

'Open/Create Options
Private Const REG_OPTION_NON_VOLATILE = 0&
Private Const REG_OPTION_VOLATILE = &H1

'Key creation/open disposition
Private Const REG_CREATED_NEW_KEY = &H1
Private Const REG_OPENED_EXISTING_KEY = &H2

'masks for the predefined standard access types
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL = &HFFFF

'Define severity codes
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_ACCESS_DENIED = 5
Private Const ERROR_INVALID_DATA = 13&
Private Const ERROR_MORE_DATA = 234 '  dderror
Private Const ERROR_NO_MORE_ITEMS = 259

'Structures Needed For Registry Prototypes
Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type


'===============================================================================
' Enumerations
'===============================================================================
Public Enum ERegistryClassConstants
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
End Enum

Public Enum ERegistryValueTypes
'Predefined Value Types
    REG_NONE = (0)                         'No value type
    REG_SZ = (1)                           'Unicode nul terminated string
    REG_EXPAND_SZ = (2)                    'Unicode nul terminated string w/enviornment var
    REG_BINARY = (3)                       'Free form binary
    REG_DWORD = (4)                        '32-bit number
    REG_DWORD_LITTLE_ENDIAN = (4)          '32-bit number (same as REG_DWORD)
    REG_DWORD_BIG_ENDIAN = (5)             '32-bit number
    REG_LINK = (6)                         'Symbolic Link (unicode)
    REG_MULTI_SZ = (7)                     'Multiple Unicode strings
    REG_RESOURCE_LIST = (8)                'Resource list in the resource map
    REG_FULL_RESOURCE_DESCRIPTOR = (9)     'Resource list in the hardware description
    REG_RESOURCE_REQUIREMENTS_LIST = (10)
End Enum


'===============================================================================
' API Declarations
'===============================================================================
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, ByVal lpType As Long, ByVal lpData As Long, ByVal lpcbData As Long) As Long
Private Declare Function RegEnumValueByte Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegEnumValueLong Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegEnumValueStr Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegQueryValueExByte Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, szData As Byte, ByRef lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, szData As Long, ByRef lpcbData As Long) As Long
Private Declare Function RegQueryValueExStr Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal szData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueExByte Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, szData As Byte, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, szData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExStr Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long


'===============================================================================
' Member variables
'===============================================================================
Private m_hClassKey As Long
Private m_sSectionKey As String
Private m_sValueKey As String
Private m_vValue As Variant
Private m_sSetValue As String
Private m_vDefault As Variant
Private m_eValueType As ERegistryValueTypes


'===============================================================================
' KeyExists - Tests for the existance of a registry key
'
' Arguments: None
'
' Returns: True if the key exists, False otherwise
'
' Notes:
'===============================================================================
    Public Property Get KeyExists() As Boolean
        On Error GoTo ErrorHandler
        Dim hKey As Long
    
1       If RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, 1, hKey) = ERROR_SUCCESS Then
2           KeyExists = True
3           RegCloseKey hKey
4       Else
5           KeyExists = False
        End If
        Exit Property

ErrorHandler:
        With Err
6           .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
        End With

    End Property


'===============================================================================
' KeyExists - Tests for the existance of a registry value
'
' Arguments: None
'
' Returns: True if the value exists, False otherwise
'
' Notes:
'===============================================================================
    Public Property Get ValueExists() As Boolean
        On Error GoTo ErrorHandler
        Dim lngHKey As Long
        Dim lngResult As Long
        Dim lngOrdType As Long
        Dim lngData As Long

        ' Determine what kind of value this will be
1       lngResult = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, KEY_QUERY_VALUE, lngHKey)
2       lngResult = RegQueryValueExLong(lngHKey, m_sValueKey, 0&, lngOrdType, 0&, lngData)
3       ValueExists = Not ((lngResult <> 0) And (lngResult <> ERROR_MORE_DATA))
        Exit Property

ErrorHandler:
        With Err
4           .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
        End With

    End Property


'===============================================================================
' CreateKey - Creates a new registry key
'
' Arguments: None
'
' Returns: True if key creation succeeds, False otherwise
'
' Notes:
'===============================================================================
    Public Function CreateKey() As Boolean
        On Error GoTo ErrorHandler
        Dim tSA As SECURITY_ATTRIBUTES
        Dim hKey As Long
        Dim lCreate As Long
        Dim e As Long

        'Open or Create the key
1       e = RegCreateKeyEx(m_hClassKey, m_sSectionKey, 0, "", REG_OPTION_NON_VOLATILE, _
            KEY_ALL_ACCESS, tSA, hKey, lCreate)
2       If e Then
3           Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to create registry Key: '" & m_sSectionKey
4       Else
5           CreateKey = (e = ERROR_SUCCESS)
            'Close the key
6           RegCloseKey hKey
        End If
        Exit Function

ErrorHandler:
        With Err
7           .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
        End With
    
    End Function


'===============================================================================
' DeleteKey - Delete a key from the registry
'
' Arguments: None
'
' Returns: True if key deletion succeeds, False otherwise
'
' Notes:
'===============================================================================
    Public Function DeleteKey() As Boolean
        On Error GoTo ErrorHandler
        Dim e As Long

1       e = RegDeleteKey(m_hClassKey, m_sSectionKey)
2       If e Then
3           Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to delete registry Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey
4       Else
5           DeleteKey = (e = ERROR_SUCCESS)
        End If
        Exit Function

ErrorHandler:
        With Err
6           .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
        End With

    End Function


'===============================================================================
' DeleteValue - Delete a value from the registry
'
' Arguments: None
'
' Returns: True if value deletion succeeds, False otherwise
'
' Notes:
'===============================================================================
    Public Function DeleteValue() As Boolean
        On Error GoTo ErrorHandler
        Dim e As Long
        Dim hKey As Long

1       e = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, KEY_ALL_ACCESS, hKey)
2       If e Then
3           Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to open key '" & m_hClassKey & "',Section: '" & m_sSectionKey & "' for delete access"
4       Else
5           e = RegDeleteValue(hKey, m_sValueKey)
6           If e Then
7               Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to delete registry Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey & "',Key: '" & m_sValueKey
8           Else
9               DeleteValue = (e = ERROR_SUCCESS)
            End If
        End If
        Exit Function

ErrorHandler:
        With Err
10          .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
        End With

    End Function


'===============================================================================
' Value_Get - Retrieve a value from the registry
'
' Arguments: None
'
' Returns: The value stored in the registry
'
' Notes:
'===============================================================================
    Public Property Get Value() As Variant
        On Error GoTo ErrorHandler
        Dim lngData As Long
        Dim lngHKey As Long
        Dim lngOrdType As Long
        Dim lngResult As Long
        Dim strData As String

        ' Determine what kind of value this will be
1       lngResult = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, KEY_QUERY_VALUE, lngHKey)
2       lngResult = RegQueryValueExLong(lngHKey, m_sValueKey, 0&, lngOrdType, 0&, lngData)
3       If lngResult And lngResult <> ERROR_MORE_DATA Then
4           Value = m_vDefault
            Exit Property
        End If
    
5       m_eValueType = lngOrdType
        Select Case lngOrdType
        Case REG_DWORD, REG_DWORD_LITTLE_ENDIAN
            Dim iData As Long
6           lngResult = RegQueryValueExLong(lngHKey, m_sValueKey, 0&, lngOrdType, iData, lngData)
7           Value = CLng(iData)
            
        Case REG_DWORD_BIG_ENDIAN  ' Unlikely, but you never know
            Dim dwData As Long
8           lngResult = RegQueryValueExLong(lngHKey, m_sValueKey, 0&, lngOrdType, dwData, lngData)
9           Value = SwapEndian(dwData)
            
        Case REG_SZ, REG_MULTI_SZ ' Same thing to Visual Basic
10          strData = String$(lngData - 1, 0)
11          lngResult = RegQueryValueExStr(lngHKey, m_sValueKey, 0&, lngOrdType, strData, lngData)
12          Value = strData
            
        Case REG_EXPAND_SZ
13          strData = String$(lngData - 1, 0)
14          lngResult = RegQueryValueExStr(lngHKey, m_sValueKey, 0&, lngOrdType, strData, lngData)
15          Value = ExpandEnvStr(strData)
            
            ' Catch REG_BINARY and anything else
        Case Else
            Dim abData() As Byte
16          ReDim abData(lngData)
17          lngResult = RegQueryValueExByte(lngHKey, m_sValueKey, 0&, lngOrdType, abData(0), lngData)
18          Value = abData
        End Select
        Exit Property

ErrorHandler:
        With Err
19          .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
        End With
    
    End Property


'===============================================================================
' Value_Let - Set a value in the registry
'
' Arguments:
'   NewValue - The value to store in the registry
'
' Returns: None
'
' Notes:
'===============================================================================
    Public Property Let Value(ByVal NewValue As Variant)
        On Error GoTo ErrorHandler
        Dim lngOrdType As Long
        Dim lngData As Long
        Dim lngHKey As Long
        Dim lngResult As Long
        Dim lngCreate As Long
        Dim typSA As SECURITY_ATTRIBUTES

        ' Open or Create the key
1       lngResult = RegCreateKeyEx(m_hClassKey, m_sSectionKey, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, typSA, lngHKey, lngCreate)
    
2       If lngResult Then
3           Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to set registry value Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey & "',Key: '" & m_sValueKey & "' to value: '" & m_vValue & "'"
4       Else
            Select Case m_eValueType
            Case REG_BINARY
5               If (VarType(NewValue) = vbArray + vbByte) Then
                    Dim ab() As Byte
6                   ab = NewValue
7                   lngOrdType = REG_BINARY
8                   lngData = UBound(ab) - LBound(ab) - 1
9                   lngResult = RegSetValueExByte(lngHKey, m_sValueKey, 0&, lngOrdType, ab(0), lngData)
10              Else
11                  Err.Raise 26001
                End If
            Case REG_DWORD, REG_DWORD_BIG_ENDIAN, REG_DWORD_LITTLE_ENDIAN
12              If (VarType(NewValue) = vbInteger) Or (VarType(NewValue) = vbLong) Then
                    Dim i As Long
13                  i = NewValue
14                  lngOrdType = REG_DWORD
15                  lngResult = RegSetValueExLong(lngHKey, m_sValueKey, 0&, lngOrdType, i, 4)
                End If
            Case REG_SZ, REG_EXPAND_SZ
                Dim s As String, iPos As Long
16              s = NewValue
17              lngOrdType = REG_SZ
                ' Assume anything with two non-adjacent percents is expanded string
18              iPos = InStr(s, "%")
19              If iPos Then
20                  If InStr(iPos + 2, s, "%") Then lngOrdType = REG_EXPAND_SZ
                End If
21              lngData = Len(s) + 1
22              lngResult = RegSetValueExStr(lngHKey, m_sValueKey, 0&, lngOrdType, s, lngData)
                
                ' User should convert to a compatible type before calling
            Case Else
23              lngResult = ERROR_INVALID_DATA
                
            End Select
        
24          If Not lngResult Then
25              m_vValue = NewValue
26          Else
27              Err.Raise vbObjectError + 1048 + 26001, App.EXEName & ".cRegistry", "Failed to set registry value Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey & "',Key: '" & m_sValueKey & "' to value: '" & m_vValue & "'"
            End If
        
            'Close the key
28          RegCloseKey lngHKey
    
        End If
        Exit Property

ErrorHandler:
        With Err
29          .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
        End With

    End Property


'===============================================================================
' EnumerateValues -
'===============================================================================
    Public Function EnumerateValues(ByRef sKeyNames() As String, ByRef iKeyCount As Long) As Boolean
        On Error GoTo ErrorHandler
        Dim lResult As Long
        Dim hKey As Long
        Dim sName As String
        Dim lNameSize As Long
        Dim sData As String
        Dim lIndex As Long
        Dim cJunk As Long
        Dim cNameMax As Long
        Dim ft As Currency
   
1       iKeyCount = 0
2       Erase sKeyNames()
    
3       lIndex = 0
4       lResult = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, KEY_QUERY_VALUE, hKey)
5       If (lResult = ERROR_SUCCESS) Then
6           lResult = RegQueryInfoKey(hKey, "", cJunk, 0, cJunk, cJunk, cJunk, cJunk, cNameMax, cJunk, cJunk, ft)
            Do While lResult = ERROR_SUCCESS
                'Set buffer space
7               lNameSize = cNameMax + 1
8               sName = String$(lNameSize, 0)
9               If (lNameSize = 0) Then lNameSize = 1

                'Get value name:
10              lResult = RegEnumValue(hKey, lIndex, sName, lNameSize, 0&, 0&, 0&, 0&)
11              If (lResult = ERROR_SUCCESS) Then
                    ' Although in theory you can also retrieve the actual
                    ' value and type here, I found it always (ultimately) resulted in
                    ' a GPF, on Win95 and NT.  Why?  Can anyone help?
12                  sName = Left$(sName, lNameSize)
                  
13                  iKeyCount = iKeyCount + 1
14                  ReDim Preserve sKeyNames(1 To iKeyCount) As String
15                  sKeyNames(iKeyCount) = sName
                End If
16              lIndex = lIndex + 1
17          Loop
        End If
18      If (hKey <> 0) Then
19          RegCloseKey hKey
        End If
20      EnumerateValues = True
        Exit Function

ErrorHandler:
21      If (hKey <> 0) Then
22          RegCloseKey hKey
        End If
    
        With Err
23          .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
        End With

    End Function


'===============================================================================
' EnumerateSections -
'===============================================================================
    Public Function EnumerateSections(ByRef sSect() As String, ByRef iSectCount As Long) As Boolean
        On Error GoTo ErrorHandler
        Dim lResult As Long
        Dim hKey As Long
        Dim dwReserved As Long
        Dim szBuffer As String
        Dim lBuffSize As Long
        Dim lIndex As Long
        Dim lType As Long
        Dim sCompKey As String
        Dim iPos As Long

1       iSectCount = 0
2       Erase sSect
    
3       lIndex = 0
    
4       lResult = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, KEY_ENUMERATE_SUB_KEYS, hKey)
        Do While lResult = ERROR_SUCCESS
            'Set buffer space
5           szBuffer = String$(255, 0)
6           lBuffSize = Len(szBuffer)
        
            'Get next value
7           lResult = RegEnumKey(hKey, lIndex, szBuffer, lBuffSize)
                               
8           If (lResult = ERROR_SUCCESS) Then
9               iSectCount = iSectCount + 1
10              ReDim Preserve sSect(1 To iSectCount) As String
11              iPos = InStr(szBuffer, Chr$(0))
12              If (iPos > 0) Then
13                  sSect(iSectCount) = Left(szBuffer, iPos - 1)
14              Else
15                  sSect(iSectCount) = Left(szBuffer, lBuffSize)
                End If
            End If
         
16          lIndex = lIndex + 1
17      Loop
18      If (hKey <> 0) Then
19          RegCloseKey hKey
        End If
20      EnumerateSections = True
        Exit Function

ErrorHandler:
21      If (hKey <> 0) Then
22          RegCloseKey hKey
        End If
   
23      Err.Raise vbObjectError + 1048 + 26002, App.EXEName & ".cRegistry", Err.Description

    End Function


'===============================================================================
' pSetClassValue
'===============================================================================
    Private Sub pSetClassValue(ByVal sValue As String)
        On Error GoTo ErrorHandler
        Dim sSection As String
    
1       ClassKey = HKEY_CLASSES_ROOT
2       Value = sValue
3       sSection = SectionKey
4       ClassKey = HKEY_LOCAL_MACHINE
5       SectionKey = "SOFTWARE\Classes\" & sSection
6       Value = sValue
7       SectionKey = sSection
        Exit Sub

ErrorHandler:
        With Err
8           .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
        End With

    End Sub


'===============================================================================
' CreateEXEAssociation -
'===============================================================================
    Public Sub CreateEXEAssociation( _
            ByVal sExePath As String, _
            ByVal sClassName As String, _
            ByVal sClassDescription As String, _
            ByVal sAssociation As String, _
            Optional ByVal sOpenMenuText As String = "&Open", _
            Optional ByVal bSupportPrint As Boolean = False, _
            Optional ByVal sPrintMenuText As String = "&Print", _
            Optional ByVal bSupportNew As Boolean = False, _
            Optional ByVal sNewMenuText As String = "&New", _
            Optional ByVal bSupportInstall As Boolean = False, _
            Optional ByVal sInstallMenuText As String = "", _
            Optional ByVal lDefaultIconIndex As Long = -1)
        On Error GoTo ErrorHandler

        ' Check if path is wrapped in quotes:
1       sExePath = Trim$(sExePath)
2       If (Left$(sExePath, 1) <> """") Then
3           sExePath = """" & sExePath
        End If
4       If (Right$(sExePath, 1) <> """") Then
5           sExePath = sExePath & """"
        End If
    
        ' Create the .File to Class association:
6       SectionKey = "." & sAssociation
7       ValueType = REG_SZ
8       ValueKey = ""
9       pSetClassValue sClassName
   
        ' Create the Class shell open command:
10      SectionKey = sClassName
11      pSetClassValue sClassDescription
    
12      SectionKey = sClassName & "\shell\open"
13      If (sOpenMenuText = "") Then sOpenMenuText = "&Open"
14      ValueKey = ""
15      pSetClassValue sOpenMenuText
16      SectionKey = sClassName & "\shell\open\command"
17      ValueKey = ""
18      pSetClassValue sExePath & " ""%1"""
    
19      If (bSupportPrint) Then
20          SectionKey = sClassName & "\shell\print"
21          If (sPrintMenuText = "") Then sPrintMenuText = "&Print"
22          ValueKey = ""
23          pSetClassValue sPrintMenuText
24          SectionKey = sClassName & "\shell\print\command"
25          ValueKey = ""
26          pSetClassValue sExePath & " /p ""%1"""
        End If
    
27      If (bSupportInstall) Then
28          If (sInstallMenuText = "") Then
29              sInstallMenuText = "&Install " & sAssociation
            End If
30          SectionKey = sClassName & "\shell\add"
31          ValueKey = ""
32          pSetClassValue sInstallMenuText
33          SectionKey = sClassName & "\shell\add\command"
34          ValueKey = ""
35          pSetClassValue sExePath & " /a ""%1"""
        End If
    
36      If (bSupportNew) Then
37          SectionKey = sClassName & "\shell\new"
38          ValueKey = ""
39          If (sNewMenuText = "") Then sNewMenuText = "&New"
40          pSetClassValue sNewMenuText
41          SectionKey = sClassName & "\shell\new\command"
42          ValueKey = ""
43          pSetClassValue sExePath & " /n ""%1"""
        End If
    
        ' Check if path is wrapped in quotes:
44      If (Left$(sExePath, 1) = """") Then
45          sExePath = Mid$(sExePath, 2)
        End If
46      If (Right$(sExePath, 1) = """") Then
47          sExePath = Left$(sExePath, Len(sExePath) - 1)
        End If
    
48      If lDefaultIconIndex > -1 Then
49          SectionKey = sClassName & "\DefaultIcon"
50          ValueKey = ""
51          pSetClassValue sExePath & "," & CStr(lDefaultIconIndex)
        End If
        Exit Sub

ErrorHandler:
        With Err
52          .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
        End With

    End Sub


'===============================================================================
' CreateAdditionalEXEAssociations -
'===============================================================================
    Public Sub CreateAdditionalEXEAssociations(ByVal sClassName As String, ParamArray vItems() As Variant)
        Dim hClassKey As Long
        Dim iItems As Long
        Dim iItem As Long
   
        On Error Resume Next
1       iItems = UBound(vItems) + 1
2       If (iItems Mod 3) <> 0 Or (Err.Number <> 0) Then
3           Err.Raise vbObjectError + 1048 + 26004, App.EXEName & ".cRegistry", "Invalid parameter list passed to CreateAdditionalEXEAssociations - expected Name/Text/Command"
4       Else
            ' Check if it exists:
5           hClassKey = m_hClassKey
6           m_hClassKey = HKEY_CLASSES_ROOT
7           SectionKey = sClassName
8           If Not (KeyExists) Then
9               m_hClassKey = hClassKey
10              Err.Raise vbObjectError + 1048 + 26005, App.EXEName & ".cRegistry", "Error - attempt to create additional associations before class defined."
11          Else
                For iItem = 0 To iItems - 1 Step 3
12                  ValueType = REG_SZ
13                  SectionKey = sClassName & "\shell\" & vItems(iItem)
14                  ValueKey = ""
15                  pSetClassValue vItems(iItem + 1)
16                  SectionKey = sClassName & "\shell\" & vItems(iItem) & "\command"
17                  ValueKey = ""
18                  pSetClassValue vItems(iItem + 2)
19              Next iItem
20              m_hClassKey = hClassKey
            End If
        End If
    End Sub


'===============================================================================
' ValueType_Get -
'===============================================================================
    Public Property Get ValueType() As ERegistryValueTypes
1       ValueType = m_eValueType
    End Property

'===============================================================================
' ValueType_Let -
'===============================================================================
    Public Property Let ValueType(ByVal eValueType As ERegistryValueTypes)
1       m_eValueType = eValueType
    End Property


'===============================================================================
' ClassKey_Get -
'===============================================================================
    Public Property Get ClassKey() As ERegistryClassConstants
1       ClassKey = m_hClassKey
    End Property

'===============================================================================
' ClassKey_Let -
'===============================================================================
    Public Property Let ClassKey(ByVal eKey As ERegistryClassConstants)
1       m_hClassKey = eKey
    End Property


'===============================================================================
' SectionKey_Get -
'===============================================================================
    Public Property Get SectionKey() As String
1       SectionKey = m_sSectionKey
    End Property

'===============================================================================
' SectionKey_Let -
'===============================================================================
    Public Property Let SectionKey(ByVal sSectionKey As String)
1       m_sSectionKey = sSectionKey
    End Property


'===============================================================================
' ValueKey_Get -
'===============================================================================
    Public Property Get ValueKey() As String
1       ValueKey = m_sValueKey
    End Property

'===============================================================================
' ValueKey_Let -
'===============================================================================
    Public Property Let ValueKey(ByVal sValueKey As String)
1       m_sValueKey = sValueKey
    End Property


'===============================================================================
' Default_Get -
'===============================================================================
    Public Property Get Default() As Variant
1       Default = m_vDefault
    End Property

'===============================================================================
' Default_Let -
'===============================================================================
    Public Property Let Default(ByVal vDefault As Variant)
1       m_vDefault = vDefault
    End Property


'===============================================================================
' SwapEndian - Swap the byte order from Big-Endian to Little-Endian or
'   vice-versa.
'
' Arguments:
'   lngData - The value to byte-swap
'
' Returns: Byte-swapped value
'
' Notes:
'===============================================================================
    Private Function SwapEndian(ByVal lngData As Long) As Long
1       CopyMemory ByVal VarPtr(SwapEndian) + 3, lngData, 1
2       CopyMemory ByVal VarPtr(SwapEndian) + 2, ByVal VarPtr(lngData) + 1, 1
3       CopyMemory ByVal VarPtr(SwapEndian) + 1, ByVal VarPtr(lngData) + 2, 1
4       CopyMemory SwapEndian, ByVal VarPtr(lngData) + 3, 1
    End Function


'===============================================================================
' ExpandEnvStr - Expands provided string containing environment variables
'
' Arguments:
'   strData - The string to expand
'
' Returns: Expanded string
'
' Notes:
'===============================================================================
    Private Function ExpandEnvStr(strData As String) As String
        Dim c As Long, s As String
        ' Get the length
1       s = "" ' Needed to get around Windows 95 limitation
2       c = ExpandEnvironmentStrings(strData, s, c)
        ' Expand the string
3       s = String$(c - 1, 0)
4       c = ExpandEnvironmentStrings(strData, s, c)
5       ExpandEnvStr = s
    End Function
