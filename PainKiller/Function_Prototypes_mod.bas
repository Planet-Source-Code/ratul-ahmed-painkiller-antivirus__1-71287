Attribute VB_Name = "Function_Prototypes_mod"
Public Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal _
lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal _
lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName _
As String) As Long
Private Const EWX_REBOOT As Long = 2
Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long

Public Sub pSuspend(filepath As String, objective As String)
Dim path As String
Dim comd As String
Dim chk As String
On Error Resume Next

chk = Right(App.path, 1)
    
    If chk = "\" Then
        path = App.path & "data\thirdparty\process.exe"
    Else
        path = App.path & "\data\thirdparty\process.exe"
    End If
    
    comd = path & " -" & objective & " " & Chr(34) & filepath & Chr(34)
    
    'MsgBox comd
    
    DOShell comd, vbHide
    
    
End Sub

Public Sub pKill(FileName As String)
Dim path As String
Dim comd As String
Dim chk As String
On Error Resume Next
chk = Right(App.path, 1)
    
    If chk = "\" Then
        path = App.path & "data\thirdparty\process.exe"
    Else
        path = App.path & "\data\thirdparty\process.exe"
    End If
    
    comd = path & " -k" & " " & FileName
    
    DOShell comd, vbHide
    
End Sub

Public Sub pDelete(FileName As String)
Dim path As String
Dim comd As String
Dim chk As String
On Error Resume Next
chk = Right(App.path, 1)
    
    If chk = "\" Then
        path = App.path & "data\thirdparty\del.exe /nologo /nr /nw "
    Else
        path = App.path & "\data\thirdparty\del.exe /nologo /nr /nw "
    End If
    
    comd = path & " " & Chr(34) & FileName & Chr(34)

    DOShell comd, vbHide
    
End Sub
'//////////////////////////////////////////////////////////////////////////////////////////////
Public Function ReadINI(strsection As String, strkey As String, strfullpath As String) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'read ini                                             '
'x = readini("Example INI", "Example", "C:\Example.ini")'
'MsgBox X                                             '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Dim strbuffer As String
   Let strbuffer$ = String$(750, Chr$(0&))
   Let ReadINI$ = Left$(strbuffer$, GetPrivateProfileString(strsection$, ByVal LCase$(strkey$), _
   "", strbuffer, Len(strbuffer), strfullpath$))
End Function

Public Sub WriteINI(strsection As String, strkey As String, strkeyvalue As String, strfullpath _
As String)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'write ini                                                       '
'Call writeini("Example INI", "Example", "Yes", "C:\Example.ini")'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call WritePrivateProfileString(strsection$, UCase$(strkey$), strkeyvalue$, strfullpath$)
End Sub
'/////////////////////////////////////////////////////////////////////////////////////////////
'.........................................................................[ Read The INF File............
'////////////////////////////////////////////////////////////////////////////////////////////
Public Function readinf(path As String) As String
Dim infpath As String
Dim appname As String
infpath = path
appname = ReadINI("autorun", "open", infpath) '.............Get Output
readinf = appname
'MsgBox appname
'----------------------------Conditions Here
End Function
'////////////////////////////////////////////////////////////////////////////////////////////
Public Function GetTarget(strPath As String) As String

    'Gets target path from a shortcut file
    On Error GoTo Error_Loading
    Dim wshShell As Object
    Dim wshLink As Object
    Set wshShell = CreateObject("WScript.Shell")
    Set wshLink = wshShell.CreateShortcut(strPath)
    GetTarget = wshLink.TargetPath
    Set wshLink = Nothing
    Set wshShell = Nothing
    Exit Function
Error_Loading:
    GetTarget = "Error occured."
End Function
