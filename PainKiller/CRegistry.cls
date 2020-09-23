VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'* Description : Class for working with the system registry.

Option Explicit

' Error handling definitions
Private Const E_ERR_BASE = 18000 + vbObjectError
Public Enum EErrRegistry
    eErrRegistry_InvalidKeyName = E_ERR_BASE + 1
    eErrRegistry_InvalidValueName
    eErrRegistry_ComponentFailure
End Enum
Private Const S_ERR_InvalidKeyName = "Invalid KeyName value"
Private Const S_ERR_InvalidValueName = "Invalid value name"
Private Const S_ERR_ComponentFailure = "CRegistry component failure"

' Public class enums
Public Enum ERegRoot
    eRegRoot_HKeyClassesRoot = &H80000000
    eRegRoot_HKeyCurrentUser = &H80000001
    eRegRoot_HKeyLocalMachine = &H80000002
    eRegRoot_HKeyUsers = &H80000003
    eRegRoot_HKeyCurrentConfig = &H80000005
    eRegRoot_HKeyDynData = &H80000006
End Enum

Public Enum ERegValue
    eRegValue_None = 0
    eRegValue_Sz = 1
    eRegValue_ExpandSz = 2
    eRegValue_Binary = 3
    eRegValue_DWord = 4
    eRegValue_DWordLittleEndian = 4
    eRegValue_DWordBigEndian = 5
    eRegValue_Link = 6
    eRegValue_MultiSz = 7
    eRegValue_ResourceList = 8
    eRegValue_FullResourceDescriptor = 9
    eRegValue_ResourceRequirementsList = 10
End Enum

' Private class constants
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_ALL_ACCESS = &H3F
Private Const REG_OPTION_NON_VOLATILE = 0&
Private Const REG_OPTION_VOLATILE = &H1
Private Const REG_CREATED_NEW_KEY = &H1
Private Const REG_OPENED_EXISTING_KEY = &H2
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL = &HFFFF
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_ACCESS_DENIED = 5
Private Const ERROR_INVALID_DATA = 13&
Private Const ERROR_MORE_DATA = 234
Private Const ERROR_NO_MORE_ITEMS = 259
Private Const MAX_PATH = 256

' Private class type definitions
Private Type SECURITY_ATTRIBUTES
    nLength                  As Long
    lpSecurityDescriptor     As Long
    bInheritHandle           As Boolean
End Type
Private Type FILETIME
    dwLowDateTime    As Long
    dwHighDateTime   As Long
End Type

' Private class API function declarations
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValueExStr Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, szData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExByte Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, szData As Byte, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueExStr Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal szData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, szData As Long, ByRef lpcbData As Long) As Long
Private Declare Function RegQueryValueExByte Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, szData As Byte, ByRef lpcbData As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, ByVal lpType As Long, ByVal lpData As Long, ByVal lpcbData As Long) As Long
Private Declare Function RegEnumValueLong Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegEnumValueStr Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegEnumValueByte Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long

' Private variables to hold property values
Private m_RootKey As ERegRoot


'*****************************************************************************************
'* Function    : Init
'* Notes       : Use this routine for basic object initialization.
'*****************************************************************************************
Public Function Init(RootKey As ERegRoot)
    On Error GoTo hComponentFailure
    
    m_RootKey = RootKey

    Exit Function

hComponentFailure:
    Err.Raise eErrRegistry_ComponentFailure, App.EXEName & ".CRegistry", S_ERR_ComponentFailure
End Function


'*****************************************************************************************
'* Property    : RootKey
'* Notes       : Returns or sets the key that will be used as root key.
'*****************************************************************************************
Public Property Get RootKey() As ERegRoot
    On Error GoTo hComponentFailure
    
    RootKey = m_RootKey

    Exit Property

hComponentFailure:
    Err.Raise eErrRegistry_ComponentFailure, App.EXEName & ".CRegistry", S_ERR_ComponentFailure
End Property

Public Property Let RootKey(ByVal eKey As ERegRoot)
    On Error GoTo hComponentFailure
    
    m_RootKey = eKey

    Exit Property

hComponentFailure:
    Err.Raise eErrRegistry_ComponentFailure, App.EXEName & ".CRegistry", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Function    : CreateKey
'* Notes       : Creates the specified registry key.
'*               Returns true if the key was created, false otherwise.
'*****************************************************************************************
Public Function CreateKey(KeyName As String) As Boolean
    On Error GoTo hComponentFailure
    
    Dim tSA      As SECURITY_ATTRIBUTES
    Dim hKey     As Long
    Dim lCreate  As Long
    Dim lRet     As Long
    
    If Len(KeyName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrRegistry_InvalidKeyName, App.EXEName & ".CRegistry", S_ERR_InvalidKeyName
    End If

    lRet = RegCreateKeyEx(m_RootKey, KeyName, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, tSA, hKey, lCreate)
    If lRet = ERROR_SUCCESS Then RegCloseKey hKey
    
    CreateKey = (lRet = ERROR_SUCCESS)

    Exit Function

hComponentFailure:
    Err.Raise eErrRegistry_ComponentFailure, App.EXEName & ".CRegistry", S_ERR_ComponentFailure
End Function


'*****************************************************************************************
'* Function    : DeleteKey
'* Notes       : Deletes  the specified registry key.
'*               Returns true if the key was deleted, false otherwise.
'*****************************************************************************************
Public Function DeleteKey(KeyName As String) As Boolean
    On Error GoTo hComponentFailure
    
    Dim lRet As Long
    
    If Len(KeyName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrRegistry_InvalidKeyName, App.EXEName & ".CRegistry", S_ERR_InvalidKeyName
    End If
    
    lRet = RegDeleteKey(m_RootKey, KeyName)
    DeleteKey = (lRet = ERROR_SUCCESS)

    Exit Function

hComponentFailure:
    Err.Raise eErrRegistry_ComponentFailure, App.EXEName & ".CRegistry", S_ERR_ComponentFailure
End Function


'*****************************************************************************************
'* Function    : DeleteValue
'* Notes       : Removes a named value from the specified registry key.
'*               Returns true if the key was deleted, false otherwise.
'*****************************************************************************************
Public Function DeleteValue(ByVal KeyName As String, ByVal ValueName As String) As Boolean
    On Error GoTo hComponentFailure
    
    Dim lRet As Long
    Dim hKey As Long
    
    If Len(KeyName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrRegistry_InvalidKeyName, App.EXEName & ".CRegistry", S_ERR_InvalidKeyName
    End If
    
    DeleteValue = False

    lRet = RegOpenKeyEx(m_RootKey, KeyName, 0, KEY_ALL_ACCESS, hKey)
    
    If lRet = ERROR_SUCCESS Then
        lRet = RegDeleteValue(hKey, ValueName)
        DeleteValue = (lRet = ERROR_SUCCESS)
    End If

    Exit Function

hComponentFailure:
    Err.Raise eErrRegistry_ComponentFailure, App.EXEName & ".CRegistry", S_ERR_ComponentFailure
End Function


'*****************************************************************************************
'* Function    : GetAllSubKeys
'* Notes       : Retrieves all the subkeys belonging to a registry key.
'*               Returns a long integer value containing the number of retrieved subkeys.
'*****************************************************************************************
Public Function GetAllSubKeys(ByVal KeyName As String, ByRef SubKeys() As String) As Long
    On Error GoTo hComponentFailure
    
    Dim Count        As Long
    Dim dwReserved   As Long
    Dim hKey         As Long
    Dim iPos         As Long
    Dim lenBuffer    As Long
    Dim lIndex       As Long
    Dim lRet         As Long
    Dim lType        As Long
    Dim sCompKey     As String
    Dim szBuffer     As String
        
    Erase SubKeys
    
    Count = 0
    lIndex = 0

    lRet = RegOpenKeyEx(m_RootKey, KeyName, 0, KEY_ENUMERATE_SUB_KEYS, hKey)
    
    Do While lRet = ERROR_SUCCESS
        
        szBuffer = String$(MAX_PATH, 0)
        lenBuffer = Len(szBuffer)
        
        lRet = RegEnumKey(hKey, lIndex, szBuffer, lenBuffer)
        
        If (lRet = ERROR_SUCCESS) Then
            
            Count = Count + 1
            
            ReDim Preserve SubKeys(1 To Count) As String
            iPos = InStr(szBuffer, Chr$(0))
            
            If (iPos > 0) Then
                SubKeys(Count) = Left$(szBuffer, iPos - 1)
            Else
                SubKeys(Count) = Left$(szBuffer, lenBuffer)
            End If
        
        End If
        
        lIndex = lIndex + 1
    
    Loop
    
    If (hKey <> 0) Then RegCloseKey hKey
    GetAllSubKeys = Count

    Exit Function

hComponentFailure:
    If (hKey <> 0) Then RegCloseKey hKey
    GetAllSubKeys = 0
    Err.Raise eErrRegistry_ComponentFailure, App.EXEName & ".CRegistry", S_ERR_ComponentFailure
End Function


'*****************************************************************************************
'* Function    : GetAllValueNames
'* Notes       : Retrieves all the  value names belonging to a registry key.
'*               Returns a long integer containing the number of retrieved names.
'*****************************************************************************************
Public Function GetAllValueNames(ByVal KeyName As String, ByRef ValueNames() As String) As Long
    On Error GoTo hComponentFailure
    
    Dim ft           As Currency
    Dim cJunk        As Long
    Dim cNameMax     As Long
    Dim Count        As Long
    Dim hKey         As Long
    Dim lIndex       As Long
    Dim lNameSize    As Long
    Dim lRet         As Long
    Dim sData        As String
    Dim sName        As String
    
    Count = 0
    Erase ValueNames()
    
    lIndex = 0
    lRet = RegOpenKeyEx(m_RootKey, KeyName, 0, KEY_QUERY_VALUE, hKey)
    
    If (lRet = ERROR_SUCCESS) Then
        
        lRet = RegQueryInfoKey(hKey, "", cJunk, 0, cJunk, cJunk, cJunk, cJunk, cNameMax, cJunk, cJunk, ft)
        
        Do While lRet = ERROR_SUCCESS
            
            lNameSize = cNameMax + 1
            sName = String$(lNameSize, 0)
            
            If (lNameSize = 0) Then lNameSize = 1
            lRet = RegEnumValue(hKey, lIndex, sName, lNameSize, 0&, 0&, 0&, 0&)
            
            If (lRet = ERROR_SUCCESS) Then
                sName = Left$(sName, lNameSize)
                Count = Count + 1
                ReDim Preserve ValueNames(1 To Count) As String
                ValueNames(Count) = sName
            End If
            
            lIndex = lIndex + 1
        
        Loop
    
    End If
    
    If (hKey <> 0) Then RegCloseKey hKey
    GetAllValueNames = Count

    Exit Function

hComponentFailure:
    If (hKey <> 0) Then RegCloseKey hKey
    GetAllValueNames = False
    Err.Raise eErrRegistry_ComponentFailure, App.EXEName & ".CRegistry", S_ERR_ComponentFailure
End Function


'*****************************************************************************************
'* Function    : GetValue
'* Notes       : Returns the value for the specified value name stored in a registry key.
'*               If there is no value stored in that key the Default value is returned.
'*****************************************************************************************
Public Function GetValue(ByVal KeyName As String, ByVal ValueName As String, ByVal Default As Variant) As Variant
    On Error GoTo hComponentFailure
    
    Dim abData()     As Byte
    Dim cData        As Long
    Dim dwData       As Long
    Dim hKey         As Long
    Dim lData        As Long
    Dim lRet         As Long
    Dim ordType      As Long
    Dim sData        As String
    Dim vValue       As Variant
    
    vValue = Default

    lRet = RegOpenKeyEx(m_RootKey, KeyName, 0, KEY_QUERY_VALUE, hKey)
    lRet = RegQueryValueExLong(hKey, ValueName, 0&, ordType, 0&, cData)
    
    If lRet And lRet <> ERROR_MORE_DATA Then
        GetValue = vValue
        Exit Function
    End If
    
    Select Case ordType
        
        Case eRegValue_DWord, eRegValue_DWordLittleEndian
            lRet = RegQueryValueExLong(hKey, ValueName, 0&, ordType, lData, cData)
            vValue = CLng(lData)
        
        Case eRegValue_DWordBigEndian
            lRet = RegQueryValueExLong(hKey, ValueName, 0&, ordType, dwData, cData)
            vValue = SwapEndian(dwData)
        
        Case eRegValue_Sz, eRegValue_MultiSz
            sData = String$(cData - 1, 0)
            lRet = RegQueryValueExStr(hKey, ValueName, 0&, ordType, sData, cData)
            vValue = sData
        
        Case eRegValue_ExpandSz
            sData = String$(cData - 1, 0)
            lRet = RegQueryValueExStr(hKey, ValueName, 0&, ordType, sData, cData)
            vValue = ExpandEnvStr(sData)
        
        Case Else
            ReDim abData(cData)
            lRet = RegQueryValueExByte(hKey, ValueName, 0&, ordType, abData(0), cData)
            vValue = abData
    
    End Select
    
    GetValue = vValue

    Exit Function

hComponentFailure:
    Err.Raise eErrRegistry_ComponentFailure, App.EXEName & ".CRegistry", S_ERR_ComponentFailure
End Function


'*****************************************************************************************
'* Function    : KeyExists
'* Notes       : Returns true if the specified registry key exists, false otherwise.
'*****************************************************************************************
Public Function KeyExists(KeyName As String) As Boolean
    On Error GoTo hComponentFailure
    
    ' Returns:  true if the key exists.
    Dim hKey As Long
    
    If Len(KeyName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrRegistry_InvalidKeyName, App.EXEName & ".CRegistry", S_ERR_InvalidKeyName
    End If
    
    If RegOpenKeyEx(m_RootKey, KeyName, 0, 1, hKey) = ERROR_SUCCESS Then
        KeyExists = True
        RegCloseKey hKey
    Else
        KeyExists = False
    End If

    Exit Function

hComponentFailure:
    Err.Raise eErrRegistry_ComponentFailure, App.EXEName & ".CRegistry", S_ERR_ComponentFailure
End Function


'*****************************************************************************************
'* Function    : SetValue
'* Notes       : Sets the value for the specified value name stored in a registry key.
'*               Returns true if the value is set, false otherwise.
'*****************************************************************************************
Public Function SetValue(ByVal KeyName As String, ByVal ValueName As String, ByVal Value As Variant, Optional ValueType As ERegValue = eRegValue_Sz) As Boolean
    On Error GoTo hComponentFailure
    
    Dim ab()         As Byte
    Dim C            As Long
    Dim i            As Long
    Dim iPos         As Long
    Dim hKey         As Long
    Dim lCreate      As Long
    Dim lRet         As Long
    Dim ordType      As Long
    Dim tSA          As SECURITY_ATTRIBUTES
    Dim s            As String
    
    If Len(KeyName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrRegistry_InvalidKeyName, App.EXEName & ".CRegistry", S_ERR_InvalidKeyName
    End If
    
    SetValue = False

    lRet = RegCreateKeyEx(m_RootKey, KeyName, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, tSA, hKey, lCreate)
    
    If lRet = ERROR_SUCCESS Then
        
        Select Case ValueType
            
            Case eRegValue_Binary
                If (VarType(Value) = vbArray + vbByte) Then
                    ab = Value
                    ordType = eRegValue_Binary
                    C = UBound(ab) - LBound(ab) - 1
                    
                    lRet = RegSetValueExByte(hKey, ValueName, 0&, ordType, ab(0), C)
                End If
            
            Case eRegValue_DWord, eRegValue_DWordBigEndian, eRegValue_DWordLittleEndian
                If (VarType(Value) = vbInteger) Or (VarType(Value) = vbLong) Then
                    i = Value
                    ordType = eRegValue_DWord
                    
                    lRet = RegSetValueExLong(hKey, ValueName, 0&, ordType, i, 4)
                End If
            
            Case eRegValue_Sz, eRegValue_ExpandSz
                s = Value
                ordType = eRegValue_Sz
                iPos = InStr(s, "%")
                
                If iPos Then
                    If InStr(iPos + 2, s, "%") Then ordType = eRegValue_ExpandSz
                End If
                
                If Len(s) > 0 Then
                    C = Len(s) + 1
                Else
                    s = vbNullChar
                    C = Len(s)
                End If
                
                lRet = RegSetValueExStr(hKey, ValueName, 0&, ordType, s, C)
            
            Case Else
                lRet = ERROR_INVALID_DATA
        
        End Select
        
        If lRet = ERROR_SUCCESS Then SetValue = True
        
        RegCloseKey hKey
    
    End If

    Exit Function

hComponentFailure:
    Err.Raise eErrRegistry_ComponentFailure, App.EXEName & ".CRegistry", S_ERR_ComponentFailure
End Function


'*****************************************************************************************
'* Function    : ValueType
'* Notes       : Returns a number containing the type of the value stored under the
'*               specified name in a registry key.
'*****************************************************************************************
Public Function ValueType(ByVal KeyName As String, ByVal ValueName As String) As ERegValue
    On Error GoTo hComponentFailure
    
    Dim cData    As Long
    Dim hKey     As Long
    Dim lRet     As Long
    Dim ordType  As Long
    Dim sData    As String
    Dim vValue   As Variant
    
    If Len(KeyName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrRegistry_InvalidKeyName, App.EXEName & ".CRegistry", S_ERR_InvalidKeyName
    End If
    
    lRet = RegOpenKeyEx(m_RootKey, KeyName, 0, KEY_QUERY_VALUE, hKey)

    lRet = RegQueryValueExLong(hKey, ValueName, 0&, ordType, 0&, cData)
    If lRet And lRet <> ERROR_MORE_DATA Then
        ValueType = eRegValue_None
    Else
        ValueType = ordType
    End If

    Exit Function

hComponentFailure:
    Err.Raise eErrRegistry_ComponentFailure, App.EXEName & ".CRegistry", S_ERR_ComponentFailure
End Function


'*****************************************************************************************
'* Function    : ExpandEnvStr
'* Notes       : Expands environment-variable strings and replaces them with their defined
'*               values.
'*****************************************************************************************
Private Function ExpandEnvStr(sData As String) As String
    On Error GoTo hComponentFailure
    
    Dim lRet     As Long
    Dim sTemp    As String
    
    sTemp = ""
    lRet = ExpandEnvironmentStrings(sData, sTemp, lRet)
    sTemp = String$(lRet - 1, 0)
    lRet = ExpandEnvironmentStrings(sData, sTemp, lRet)
    
    ExpandEnvStr = sTemp

    Exit Function

hComponentFailure:
    Err.Raise eErrRegistry_ComponentFailure, App.EXEName & ".CRegistry", S_ERR_ComponentFailure
End Function


'*****************************************************************************************
'* Function    : SwapEndian
'* Notes       : Swaps the highest byte with the lowest byte for the BIG_ENDIAN format.
'*****************************************************************************************
Private Function SwapEndian(ByVal dw As Long) As Long
    On Error GoTo hComponentFailure
    
    CopyMemory ByVal VarPtr(SwapEndian) + 3, dw, 1
    CopyMemory ByVal VarPtr(SwapEndian) + 2, ByVal VarPtr(dw) + 1, 1
    CopyMemory ByVal VarPtr(SwapEndian) + 1, ByVal VarPtr(dw) + 2, 1
    CopyMemory SwapEndian, ByVal VarPtr(dw) + 3, 1

    Exit Function

hComponentFailure:
    Err.Raise eErrRegistry_ComponentFailure, App.EXEName & ".CRegistry", S_ERR_ComponentFailure
End Function


'*****************************************************************************************
'* Sub         : Class_Initialize
'* Notes       : Class data space initialization.
'*****************************************************************************************
Private Sub Class_Initialize()
    On Error GoTo hComponentFailure
    
    m_RootKey = eRegRoot_HKeyClassesRoot

    Exit Sub

hComponentFailure:
    Err.Raise eErrRegistry_ComponentFailure, App.EXEName & ".CRegistry", S_ERR_ComponentFailure
End Sub
