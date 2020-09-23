Attribute VB_Name = "Reg_MOD"
'Registry Key/Value Enumeration Functions

'By Max Raskin 29 August 2000


Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long


Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long

Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey& Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String)


Private Const KEY_QUERY_VALUE = &H1
Private Const MAX_PATH = 260

Enum RegDataTypes
    REG_SZ = 1                         ' Unicode nul terminated string
    REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
    REG_DWORD = 4                      ' 32-bit number
End Enum

Enum RegistryKeys
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_DYN_DATA = &H80000006
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_USERS = &H80000003
End Enum

Enum ValKey
    Values = 0
    Keys = 1
End Enum

Private Type ByteArray
  FirstByte As Byte
  ByteBuffer(255) As Byte
End Type


Global Const ERROR_NONE = 0
Global Const ERROR_BADDB = 1
Global Const ERROR_BADKEY = 2
Global Const ERROR_CANTOPEN = 3
Global Const ERROR_CANTREAD = 4
Global Const ERROR_CANTWRITE = 5
Global Const ERROR_OUTOFMEMORY = 6
Global Const ERROR_INVALID_PARAMETER = 7
Global Const ERROR_ACCESS_DENIED = 8
Global Const ERROR_INVALID_PARAMETERS = 87
Global Const ERROR_NO_MORE_ITEMS = 259

Global Const KEY_ALL_ACCESS = &H3F

Global Const REG_OPTION_NON_VOLATILE = 0



Dim baData As ByteArray

Public Function OpenKey(RegistryKey As RegistryKeys, Optional SubKey As String) As Long
    If OpenKey <> 0 Then RegCloseKey (OpenKey)
    RegOpenKeyEx RegistryKey, SubKey, 0, KEY_QUERY_VALUE, OpenKey
End Function

Public Function GetCount(RegisteryKeyHandle As Long, ValuesOrKeys As ValKey) As Long
    If ValuesOrKeys = Keys Then RegQueryInfoKey RegisteryKeyHandle, "", 0, 0, GetCount, 0, 0, 0, 0, 0, 0, 0
    If ValuesOrKeys = Values Then RegQueryInfoKey RegisteryKeyHandle, "", 0, 0, 0, 0, 0, GetCount, 0, MAX_PATH + 1, 0, 0
End Function

Public Function EnumKey(RegisteryKeyHandle As Long, KeyIndex As Long) As String
    EnumKey = Space(MAX_PATH + 1)
    RegEnumKey RegisteryKeyHandle, KeyIndex, EnumKey, MAX_PATH + 1
    EnumKey = Trim(EnumKey)
End Function

Public Function EnumValue(RegisteryKeyHandle As Long, KeyIndex As Long) As String
    Dim lBufferLen As Long, i As Integer
    For i = 0 To 255
      baData.ByteBuffer(i) = 0
    Next
    lBufferLen = 255
    EnumValue = Space(MAX_PATH + 1)
    RegQueryInfoKey RegisteryKeyHandle, "", 0, 0, 0, 0, 0, 0, lValNameLen, lValLen, 0, 0
    RegEnumValue RegisteryKeyHandle, KeyIndex, EnumValue, MAX_PATH + 1, 0, 0, baData.FirstByte, lBufferLen
    EnumValue = Trim(EnumValue)
End Function


Public Function SetValue(RegisteryKeyHandle As RegistryKeys, SubRegistryKey As String, KeyName As String, NewValue As String, Optional datatype As RegDataTypes)
    Dim lRetVal As Long
    lRetVal = OpenKey(RegisteryKeyHandle, SubRegistryKey)
    If datatype = 0 Then datatype = REG_SZ
    RegSetValueEx lRetVal, KeyName, 0, datatype, NewValue, LenB(StrConv(SubKeyValue, vbFromUnicode))
End Function

Public Function GetKeyValue(hKey As Long, KeyName As String) As String
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    
    Dim hDepth As Long                                      '
    Dim sKeyVal As String
    Dim lKeyValType As Long                                 ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, KeyName, 0, _
                         lKeyValType, tmpVal, KeyValSize)
    GetKeyValue = Trim(tmpVal)
End Function

Public Function DeleteValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)
' Description:
'   This Function will delete a value
'
' Syntax:
'   DeleteValue Location, KeyName, ValueName
'
'   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
'   , HKEY_USERS
'
'   KeyName is the name of the key that the value you wish to delete is in
'   , it may include subkeys (example "Key1\SubKey1")
'
'   ValueName is the name of value you wish to delete

       Dim lRetVal As Long         'result of the SetValueEx function
       Dim hKey As Long         'handle of open key

       'open the specified key

       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
       lRetVal = RegDeleteValue(hKey, sValueName)
       RegCloseKey (hKey)
End Function
Public Function DeleteKey(lPredefinedKey As Long, sKeyName As String)
' Description:
'   This Function will Delete a key
'
' Syntax:
'   DeleteKey Location, KeyName
'
'   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
'   , HKEY_USERS
'
'   KeyName is name of the key you wish to delete, it may include subkeys (example "Key1\SubKey1")


    Dim lRetVal As Long         'result of the SetValueEx function
    Dim hKey As Long         'handle of open key
    
    'open the specified key
    
    'lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lRetVal = RegDeleteKey(lPredefinedKey, sKeyName)
    'RegCloseKey (hKey)
End Function
Public Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
    Dim lValue As Long
    Dim sValue As String

    Select Case lType
        Case REG_SZ
            sValue = vValue
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
        End Select

End Function

Public Function CreateNewKey(lPredefinedKey As Long, sNewKeyName As String)
' Description:
'   This Function will create a new key
'
' Syntax:
'   QueryValue Location, KeyName
'
'   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
'   , HKEY_USERS
'
'   KeyName is name of the key you wish to create, it may include subkeys (example "Key1\SubKey1")

    
    
    Dim hNewKey As Long         'handle to the new key
    Dim lRetVal As Long         'result of the RegCreateKeyEx function
    
    lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetVal)
    RegCloseKey (hNewKey)
End Function


Sub Main()
    'Examples of each function:
    'CreateNewKey HKEY_CURRENT_USER, "TestKey\SubKey1\SubKey2"
    'SetKeyValue HKEY_CURRENT_USER, "TestKey\SubKey1", "Test", "Testing, Testing", REG_SZ
    'MsgBox QueryValue(HKEY_CURRENT_USER, "TestKey\SubKey1", "Test")
    'DeleteKey HKEY_CURRENT_USER, "TestKey\SubKey1\SubKey2"
    'DeleteValue HKEY_CURRENT_USER, "TestKey\SubKey1", "Test"
End Sub


Public Function SetKeyValue(lPredefinedKey As Long, sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)
' Description:
'   This Function will set the data field of a value
'
' Syntax:
'   QueryValue Location, KeyName, ValueName, ValueSetting, ValueType
'
'   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
'   , HKEY_USERS
'
'   KeyName is the key that the value is under (example: "Key1\SubKey1")
'
'   ValueName is the name of the value you want create, or set the value of (example: "ValueTest")
'
'   ValueSetting is what you want the value to equal
'
'   ValueType must equal either REG_SZ (a string) Or REG_DWORD (an integer)

       Dim lRetVal As Long         'result of the SetValueEx function
       Dim hKey As Long         'handle of open key

       'open the specified key

       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
       lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
       RegCloseKey (hKey)

End Function





