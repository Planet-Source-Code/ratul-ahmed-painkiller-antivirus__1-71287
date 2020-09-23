VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEnvironment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'* Description : Class for retrieving environment parameters.

Option Explicit

' Error handling definitions
Private Const E_ERR_BASE = 17380 + vbObjectError
Public Enum EErrEnvironment
    eErrEnvironment_CannotGetEnvironmentVariable = E_ERR_BASE + 1
    eErrEnvironment_CannotSetEnvironmentVariable
    eErrEnvironment_CannotGetOsName
    eErrEnvironment_UnknownOperatingSystem
    eErrEnvironment_ComponentFailure
End Enum
Private Const S_ERR_CannotGetEnvironmentVariable = "Cannot get environment variable"
Private Const S_ERR_CannotSetEnvironmentVariable = "Cannot set environment variable"
Private Const S_ERR_CannotGetOsName = "Cannot get operating system name"
Private Const S_ERR_UnknwonOperatingSystem = "Unknown operating system"
Private Const S_ERR_ComponentFailure = "CEnvironment component failure"

' Public class enums
Public Enum EVbAppRunMode
    eVbAppRunMode_Compiled = 1
    eVbAppRunMode_FromIDE
    eVbAppRunMode_Unknown
End Enum

' Private class type definitions
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wReserved(1) As Integer
End Type

' Private class constants
Private Const MAX_LENGTH = 512
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

' Private class API function declarations
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal Path As String, ByVal cbBytes As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Function APISetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

' Private variables to hold property values
Private m_OSVersion As OSVERSIONINFOEX


'*****************************************************************************************
'* Property    : ComputerName
'* Notes       : Returns the computer name of the current system.
'*****************************************************************************************
Public Property Get ComputerName() As String
    On Error GoTo hComponentFailure
    
    Dim s       As String
    Dim apiRet  As Long
    Dim lSize   As Long
    
    s = Space$(MAX_LENGTH)
    lSize = Len(s)
    
    apiRet = GetComputerName(s, lSize)
    If apiRet Then
        If lSize > Len(s) Then
            s = Space$(lSize + 1)
            lSize = Len(s)
            apiRet = GetComputerName(s, lSize)
        End If
    End If
    
    ComputerName = IIf(lSize > 0, Left$(s, InStr(s, vbNullChar) - 1), "")

    Exit Property

hComponentFailure:
    Err.Raise eErrEnvironment_ComponentFailure, App.EXEName & ".CEnvironment", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Property    : OsName
'* Notes       : Returns a string value containing the operating system's name.
'*               Possible return values are WinNT4, Win95, Win98 etc.
'*****************************************************************************************
Public Property Get OsName() As String
    On Error GoTo hComponentFailure
    
    Dim sTemp As String
    
    sTemp = ""
    
    If GetOsVersion Then
        
        Select Case m_OSVersion.dwPlatformId
            
            Case VER_PLATFORM_WIN32_NT
                sTemp = "WinNT" & m_OSVersion.dwMajorVersion
            
            Case VER_PLATFORM_WIN32_WINDOWS
                If ((m_OSVersion.dwMajorVersion > 4) Or ((m_OSVersion.dwMajorVersion = 4) And (m_OSVersion.dwMinorVersion > 0))) Then
                    sTemp = "Win98"
                Else
                    sTemp = "Win95"
                End If
            
            Case VER_PLATFORM_WIN32_WINDOWS
                sTemp = "Win32s"
                
            Case Else
                On Error GoTo 0
                Err.Raise eErrEnvironment_UnknownOperatingSystem, App.EXEName & ".CEnvironment", S_ERR_UnknwonOperatingSystem
        
        End Select
    
    Else
        On Error GoTo 0
        Err.Raise eErrEnvironment_CannotGetOsName, App.EXEName & ".CEnvironment", S_ERR_CannotGetOsName
    End If
    
    OsName = sTemp

    Exit Property

hComponentFailure:
    Err.Raise eErrEnvironment_ComponentFailure, App.EXEName & ".CEnvironment", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Property    : Path
'* Notes       : Returns a string value containing the current search path.
'*****************************************************************************************
Public Property Get Path() As String
    On Error GoTo hComponentFailure
    
    Path = GetEnvironmentVariable("%Path%")

    Exit Property

hComponentFailure:
    Err.Raise eErrEnvironment_ComponentFailure, App.EXEName & ".CEnvironment", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Property    : SystemDirectory
'* Notes       : Returns a string value containing the path of the system directory.
'*****************************************************************************************
Public Property Get SystemDirectory() As String
    On Error GoTo hComponentFailure
    
    Dim s As String
    Dim C As Long
    
    s = String$(MAX_LENGTH, 0)
    C = GetSystemDirectory(s, MAX_LENGTH)
    
    If C > 0 Then
        If C > Len(s) Then
            s = Space$(C + 1)
            C = GetSystemDirectory(s, MAX_LENGTH)
        End If
    End If
    
    SystemDirectory = IIf(C > 0, Left$(s, C), "")

    Exit Property

hComponentFailure:
    Err.Raise eErrEnvironment_ComponentFailure, App.EXEName & ".CEnvironment", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Property    : TempDirectory
'* Notes       : Returns a string value containing the path of the directory designated
'*               for temporary files.
'*****************************************************************************************
Public Property Get TempDirectory() As String
    On Error GoTo hComponentFailure
    
    Dim s As String
    Dim C As Long
    
    s = Space$(MAX_LENGTH)
    C = GetTempPath(MAX_LENGTH, s)
    
    If C > 0 Then
        If C > Len(s) Then
            s = Space$(C + 1)
            C = GetTempPath(MAX_LENGTH, s)
        End If
    End If
    
    TempDirectory = IIf(C > 0, Left$(s, C), "")

    Exit Property

hComponentFailure:
    Err.Raise eErrEnvironment_ComponentFailure, App.EXEName & ".CEnvironment", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Property    : UserName
'* Notes       : Returns the user name of the current thread. This is the name of the user
'*               currently logged onto the system.
'*****************************************************************************************
Public Property Get UserName() As String
    On Error GoTo hComponentFailure
    
    Dim s       As String
    Dim apiRet  As Long
    Dim lSize   As Long
    
    s = Space$(MAX_LENGTH)
    lSize = Len(s)
    
    apiRet = GetUserName(s, lSize)
    If apiRet Then
        If lSize > Len(s) Then
            s = Space$(lSize + 1)
            lSize = Len(s)
            apiRet = GetUserName(s, lSize)
        End If
    End If
    
    UserName = IIf(lSize > 0, Left$(s, InStr(s, vbNullChar) - 1), "")

    Exit Property

hComponentFailure:
    Err.Raise eErrEnvironment_ComponentFailure, App.EXEName & ".CEnvironment", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Property    : VbAppRunMode
'* Notes       : Returns a constant specifying if the current Visual Basic program runs
'*               under the VB IDE  or not.
'*****************************************************************************************
Public Property Get VbAppRunMode(Optional VbExeName As String = "VB6.EXE") As EVbAppRunMode
    On Error GoTo hComponentFailure
    
    Dim lRet       As Long
    Dim sBuffer    As String
    
    sBuffer = Space$(2048)
    lRet = GetModuleFileName(0&, sBuffer, Len(sBuffer))
     
    If lRet = 0 Then
        VbAppRunMode = eVbAppRunMode_Unknown
    Else
        sBuffer = UCase$(Left$(sBuffer, lRet))
        
        If Right$(sBuffer, Len(VbExeName) + 1) = ("\" & VbExeName) Then
            VbAppRunMode = eVbAppRunMode_FromIDE
        Else
            VbAppRunMode = eVbAppRunMode_Compiled
        End If
    End If
    
    Exit Property

hComponentFailure:
    Err.Raise eErrEnvironment_ComponentFailure, App.EXEName & ".CEnvironment", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Property    : WindowsDirectory
'* Notes       : Returns a string containing the path of the Windows directory.
'*****************************************************************************************
Public Property Get WindowsDirectory() As String
    On Error GoTo hComponentFailure
    
    Dim s As String
    Dim C As Long
    
    s = String$(MAX_LENGTH, 0)
    C = GetWindowsDirectory(s, MAX_LENGTH)
    
    If C > 0 Then
        If C > Len(s) Then
            s = Space$(C + 1)
            C = GetWindowsDirectory(s, MAX_LENGTH)
        End If
    End If
    
    WindowsDirectory = IIf(C > 0, Left$(s, C), "")

    Exit Property

hComponentFailure:
    Err.Raise eErrEnvironment_ComponentFailure, App.EXEName & ".CEnvironment", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Function    : GetEnvironmentVariable
'* Notes       : Returns a string value filled with the contents of an environment
'*               variable.
'*****************************************************************************************
Public Function GetEnvironmentVariable(Name As String) As String
    On Error GoTo hComponentFailure
    
    Dim lRet As Long
    Dim sRet As String
    
    lRet = 0
    sRet = ""
    
    lRet = ExpandEnvironmentStrings(Name, sRet, lRet)
    
    If lRet = 0 Then
        On Error GoTo 0
        Err.Raise eErrEnvironment_CannotGetEnvironmentVariable, App.EXEName & ".CEnvironment", S_ERR_CannotGetEnvironmentVariable
    End If
    
    sRet = String$(lRet - 1, 0)
    
    lRet = ExpandEnvironmentStrings(Name, sRet, lRet)
    
    If lRet = 0 Then
        On Error GoTo 0
        Err.Raise eErrEnvironment_CannotGetEnvironmentVariable, App.EXEName & ".CEnvironment", S_ERR_CannotGetEnvironmentVariable
    End If
    
    If Right$(sRet, 1) = vbNullChar Then sRet = Left$(sRet, Len(sRet) - 1)
    
    GetEnvironmentVariable = sRet

    Exit Function

hComponentFailure:
    Err.Raise eErrEnvironment_ComponentFailure, App.EXEName & ".CEnvironment", S_ERR_ComponentFailure
End Function


'*****************************************************************************************
'* Function    : SetEnvironmentVariable
'* Notes       : Sets the value of the specified environment variable for the current
'*               process. The operating system creates the environment variable if it
'*               does not exist.
'*****************************************************************************************
Public Sub SetEnvironmentVariable(Name As String, Value As String)
    On Error GoTo hComponentFailure
    
    If APISetEnvironmentVariable(Name, Value) = 0 Then
        On Error GoTo 0
        Err.Raise eErrEnvironment_CannotSetEnvironmentVariable, App.EXEName & ".CEnvironment", S_ERR_CannotSetEnvironmentVariable
    End If
    
    Exit Sub

hComponentFailure:
    Err.Raise eErrEnvironment_ComponentFailure, App.EXEName & ".CEnvironment", S_ERR_ComponentFailure
End Sub


'*****************************************************************************************
'* Function    : GetOsVersion
'* Notes       : Obtains extended information about the version of the operating system
'*               that is currently running.
'*****************************************************************************************
Private Function GetOsVersion() As Boolean
    On Error GoTo hComponentFailure
    
    Dim lRet As Long
    Dim osV  As OSVERSIONINFO
    
    GetOsVersion = False
    m_OSVersion.dwOSVersionInfoSize = Len(m_OSVersion)
    
    If GetVersionEx(m_OSVersion) Then
        GetOsVersion = True
    Else
        m_OSVersion.dwOSVersionInfoSize = Len(osV)
        If GetVersionEx(m_OSVersion) Then GetOsVersion = True
    End If

    Exit Function

hComponentFailure:
    Err.Raise eErrEnvironment_ComponentFailure, App.EXEName & ".CEnvironment", S_ERR_ComponentFailure
End Function
