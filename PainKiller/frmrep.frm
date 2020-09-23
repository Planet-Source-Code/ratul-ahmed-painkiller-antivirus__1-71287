VERSION 5.00
Begin VB.Form frmrep 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmrep.frx":0000
   ScaleHeight     =   3240
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer boom 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   1920
   End
   Begin VB.Timer tmrmain 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   240
      Top             =   1080
   End
   Begin VB.Timer R1TMR 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   360
      Top             =   480
   End
   Begin PainKiller.StylerButton canbut 
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Cancel (10)"
      CaptionDisableColor=   12236471
      CaptionEffectColor=   16777215
      Theme           =   4
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PainKiller.XandersXPProgressBar pb 
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   2700
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   128
      Max             =   10
      Scrolling       =   5
      Value           =   100
   End
   Begin VB.OptionButton rboot 
      BackColor       =   &H00D6B487&
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   1660
      Width           =   375
   End
   Begin VB.OptionButton wrreg 
      BackColor       =   &H00D6B487&
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1280
      Width           =   375
   End
   Begin VB.OptionButton rscon 
      BackColor       =   &H00D6B487&
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   880
      Width           =   375
   End
   Begin VB.OptionButton rstart 
      BackColor       =   &H00D6B487&
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rebooting your computer is hardly recommended."
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   2280
      Width           =   4215
   End
End
Attribute VB_Name = "frmrep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Dim l As Integer


Private Sub repstrt()
'Dim x As Integer
rstart.Value = True
pb.Value = 0
'Sleep 100
Label1.Caption = "Please wait. While processing algorithms..."
tmrmain.Enabled = True
End Sub
Private Sub ressys()
Dim v As Integer
Dim sysdir As String
Dim mypath As String
Dim exepath As String
On Error Resume Next
sysdir = GetSystemDirectory

rscon.Value = True
pb.Value = 0

CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run"
CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce"
CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnceEx"

CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents"
CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\IMAIL"
CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\MAPI"
CreateNewKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\MSFS"
CreateNewKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices"
CreateNewKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServicesOnce"

SetKeyValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\IMAIL", "Installed", "1", REG_SZ
SetKeyValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\MAPI", "Installed", "1", REG_SZ
SetKeyValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\MAPI", "NoChange", "1", REG_SZ
SetKeyValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\MSFS", "Installed", "1", REG_SZ

pb.Value = 1

'============================HKEY_LOCAL_MACHINE===========================================
CreateNewKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run"
CreateNewKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunOnce"
CreateNewKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunOnceEx"


CreateNewKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents"
CreateNewKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\IMAIL"
CreateNewKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\MAPI"
CreateNewKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\MSFS"

SetKeyValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\IMAIL", "Installed", "1", REG_SZ
SetKeyValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\MAPI", "Installed", "1", REG_SZ
SetKeyValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\MAPI", "NoChange", "1", REG_SZ
SetKeyValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\MSFS", "Installed", "1", REG_SZ
CreateNewKey HKEY_USERS, ".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Run"

pb.Value = 2


'GENARATE AUTORUN Killing WITH CMD-----------------------------------------------------------------------
CreateNewKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Command Processor"
DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Command Processor", "AutoRun"
'--------------------------------------------------------------------------------------------------

pb.Value = 3
wrreg.Value = True
For v = 0 To 7
    Sleep 300
    pb.Value = pb.Value + 1
    'CREATING VALU TO ENABLE UHide
    SetKeyValue HKEY_LOCAL_MACHINE, _
    "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Hidden", _
    "00000001", REG_DWORD
    
    'CREATING VALU TO ENABLE Hide ext
    SetKeyValue HKEY_LOCAL_MACHINE, _
    "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "HideFileExt", _
    "00000000", REG_DWORD

    'CREATING VALU TO Enable Super hide
    SetKeyValue HKEY_LOCAL_MACHINE, _
    "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "SuperHidden", _
    "00000001", REG_DWORD
    'CREATING VALU TO Enable Super hide
    SetKeyValue HKEY_LOCAL_MACHINE, _
    "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden", _
    "00000001", REG_DWORD


    'Desable Folder Option
    CreateNewKey HKEY_CURRENT_USER, _
    "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    SetKeyValue HKEY_CURRENT_USER, _
    "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOtions", _
    "00000000", REG_DWORD

    'CREATING VALU TO DISABLE TASKMAN
    CreateNewKey HKEY_CURRENT_USER, _
    "Software\Microsoft\Windows\CurrentVersion\Policies\System"
    SetKeyValue HKEY_CURRENT_USER, _
    "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", _
    "00000000", REG_DWORD
    'CREATING VALU FOR DISABLING REGEDIT
    CreateNewKey HKEY_CURRENT_USER, _
    "Software\Microsoft\Windows\CurrentVersion\Policies\System"
    SetKeyValue HKEY_CURRENT_USER, _
    "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", _
    "00000000", REG_DWORD
    'CREATING VALU TO ENABLE UHide
    SetKeyValue HKEY_CURRENT_USER, _
    "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Hidden", _
    "00000001", REG_DWORD

    'CREATING VALU TO ENABLE Hide ext
    SetKeyValue HKEY_CURRENT_USER, _
    "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "HideFileExt", _
    "00000000", REG_DWORD

    'CREATING VALU TO Enable Super hide
    SetKeyValue HKEY_CURRENT_USER, _
    "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "SuperHidden", _
    "00000001", REG_DWORD
    'CREATING VALU TO Enable Super hide
    SetKeyValue HKEY_CURRENT_USER, _
    "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden", _
    "00000001", REG_DWORD
Next v
Label1.Caption = "Rebooting your computer is hardly recommended."
Me.MousePointer = vbDefault
'Me.MousePointer  = vbHourglass
boom.Enabled = True
rboot.Value = True
canbut.Enabled = True
'Me.MousePointer  = vbNormal
'Me.MousePointer  = vbHourglass
End Sub
Private Sub makeback()
Dim path As String
Dim comd As String
Dim chk As String
On Error Resume Next
chk = Right(App.path, 1)
    
    If chk = "\" Then
        path = App.path & "data\thirdparty\reg.exe"
    Else
        path = App.path & "\data\thirdparty\reg.exe"
    End If
    
    comd = Chr(34) & path & Chr(34) & " EXPORT" & " HKLM\Software\Microsoft\Windows\CurrentVersion\Run " & Chr(34) & App.path & "\backup\regback1.reg" & Chr(34)
    'MsgBox comd
    DOShell comd, vbHide
    pb.Value = 1
    Sleep 1000
    comd = Chr(34) & path & Chr(34) & " EXPORT" & " HKLM\Software\Microsoft\Windows\CurrentVersion\RunOnce " & Chr(34) & App.path & "\backup\regback2.reg" & Chr(34)
    'MsgBox comd
    DOShell comd, vbHide
    pb.Value = 2
    Sleep 1000
    comd = Chr(34) & path & Chr(34) & " EXPORT" & " HKLM\Software\Microsoft\Windows\CurrentVersion\RunOnceEx " & Chr(34) & App.path & "\backup\regback3.reg" & Chr(34)
    'MsgBox comd
    DOShell comd, vbHide
    pb.Value = 3
    Sleep 1000
    comd = Chr(34) & path & Chr(34) & " EXPORT" & " HKLM\Software\Microsoft\Windows\CurrentVersion\RunServices " & Chr(34) & App.path & "\backup\regback4.reg" & Chr(34)
    'MsgBox comd
    DOShell comd, vbHide
    pb.Value = 4
    Sleep 1000
    comd = Chr(34) & path & Chr(34) & " EXPORT" & " HKLM\Software\Microsoft\Windows\CurrentVersion\RunServicesOnce " & Chr(34) & App.path & "\backup\regback5.reg" & Chr(34)
    'MsgBox comd
    DOShell comd, vbHide
    pb.Value = 5
    Sleep 1000
      comd = Chr(34) & path & Chr(34) & " EXPORT" & " HKCU\Software\Microsoft\Windows\CurrentVersion\Run " & Chr(34) & App.path & "\backup\regback6.reg" & Chr(34)
    'MsgBox comd
    DOShell comd, vbHide
    pb.Value = 6
    Sleep 1000
    comd = Chr(34) & path & Chr(34) & " EXPORT" & " HKCU\Software\Microsoft\Windows\CurrentVersion\RunOnce " & Chr(34) & App.path & "\backup\regback7.reg" & Chr(34)
    'MsgBox comd
    DOShell comd, vbHide
    pb.Value = 7
    Sleep 1000
    comd = Chr(34) & path & Chr(34) & " EXPORT" & " HKCU\Software\Microsoft\Windows\CurrentVersion\RunOnceEx " & Chr(34) & App.path & "\backup\regback8.reg" & Chr(34)
    'MsgBox comd
    DOShell comd, vbHide
    pb.Value = 8
    Sleep 1000
    comd = Chr(34) & path & Chr(34) & " EXPORT" & " HKCU\Software\Microsoft\Windows\CurrentVersion\RunServices " & Chr(34) & App.path & "\backup\regback9.reg" & Chr(34)
    'MsgBox comd
    DOShell comd, vbHide
    pb.Value = 9
    Sleep 1000
    comd = Chr(34) & path & Chr(34) & " EXPORT" & " HKCU\Software\Microsoft\Windows\CurrentVersion\RunServicesOnce " & Chr(34) & App.path & "\backup\regback10.reg" & Chr(34)
    'MsgBox comd
    DOShell comd, vbHide
    Sleep 1000
    comd = Chr(34) & path & Chr(34) & " EXPORT" & " HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced " & Chr(34) & App.path & "\backup\regback11.reg" & Chr(34)
    'MsgBox comd
    DOShell comd, vbHide
    Sleep 1000
    comd = Chr(34) & path & Chr(34) & " EXPORT" & " HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced " & Chr(34) & App.path & "\backup\regback12.reg" & Chr(34)
    'MsgBox comd
    DOShell comd, vbHide
    Sleep 1000
    comd = Chr(34) & path & Chr(34) & " EXPORT" & " HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer " & Chr(34) & App.path & "\backup\regback13.reg" & Chr(34)
    'MsgBox comd
    DOShell comd, vbHide
    Sleep 1000
    comd = Chr(34) & path & Chr(34) & " EXPORT" & " HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer " & Chr(34) & App.path & "\backup\regback14.reg" & Chr(34)
    'MsgBox comd
    DOShell comd, vbHide
    Sleep 1000
    comd = Chr(34) & path & Chr(34) & " EXPORT" & " HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\System " & Chr(34) & App.path & "\backup\regback15.reg" & Chr(34)
    'MsgBox comd
    DOShell comd, vbHide
    Sleep 1000
    comd = Chr(34) & path & Chr(34) & " EXPORT" & " HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System " & Chr(34) & App.path & "\backup\regback16.reg" & Chr(34)
    'MsgBox comd
    DOShell comd, vbHide
    pb.Value = 10
    Sleep 1000
End Sub



Private Sub canbut_Click()
boom.Enabled = False
l = 0
MsgBox "It is recommended to reboot your computer, please reboot manually!", vbInformation, "PianKiller"
frmmain.Visible = True
Unload Me
End Sub

Private Sub Form_Load()
'Me.MousePointer  = vbNormal
'''''Me.MousePointer = vbHourglass
repstrt
l = 0
End Sub

Private Sub R1TMR_Timer()
R1TMR.Enabled = False
pb.Value = 0
rscon.Value = True

For X = 0 To 10
Sleep 300
    pb.Value = pb.Value + 1
'MsgBox x
'====================================== All virus Reg here
'1 AUTOEXEC.COM
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RunJava"
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RunJava2"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RunJava"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RunJava2"
'2 KRAG
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "krag"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "krag"
'3 LILF
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Winsock2 driver"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Winsock2 driver"
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce", "Winsock2 driver"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce", "Winsock2 driver"
'4 m1t8ta
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "amva"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "amva"
'5 RevMon
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SVCHOST"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SVCHOST"
'6 Setupexe
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "MyApp"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "MyApp"
'7 smss-funnymst
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Runonce"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Runonce"
'8 Setupmp4
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RunJava"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RunJava2"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RunJava"
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RunJava2"
'9 tip
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RunJava"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RunJava2"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RunJava"
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RunJava2"
'10 system-4msamir
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SYS1"
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SYS2"
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SYS3"
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SYS4"
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Msmsgs"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SYS1"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SYS2"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SYS3"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SYS4"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Msmsgs"
'11 SSVICHOSST
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "A:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "C:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "D:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "E:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "F:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "G:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "H:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "I:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "J:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "K:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "L:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "M:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "N:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "O:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "P:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "Q:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "R:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "S:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "T:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "U:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "V:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "W:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "X:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "Y:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "Z:\SSVICHOSST.exe"
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Yahoo Messengger"
DeleteValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell"
SetKeyValue HKEY_CURRENT_USER, _
"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" _
, "Shell", "Explorer.exe", REG_SZ

DeleteValue HKEY_USERS, _
"S-1-5-21-1343024091-1682526488-1801674531-1003\Software\Microsoft\Windows\CurrentVersion\Run", "Yahoo Messengger"

DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "A:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "C:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "D:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "E:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "F:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "G:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "H:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "I:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "J:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "K:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "L:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "M:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "N:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "O:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "P:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "Q:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "R:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "S:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "T:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "U:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "V:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "W:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "X:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "Y:\SSVICHOSST.exe"
DeleteValue HKEY_LOCAL_MACHINE, _
"Software\Microsoft\Windows\ShellNoRoam\MUICache", "Z:\SSVICHOSST.exe"

DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Yahoo Messengger"
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell"
SetKeyValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" _
, "Shell", "Explorer.exe", REG_SZ
'Flashy Bot
DeleteValue HKEY_LOCAL_MACHINE, _
"System\controlSet001\Services", "Flashy Bot"
DeleteValue HKEY_CURRENT_USER, _
"System\controlSet001\Services", "Flashy Bot"
'12 KALSHI spammer trojan registry entry
DeleteValue HKEY_LOCAL_MACHINE, _
"System\controlSet001\Services", "MassSender"
'13 msblaster registry entry
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "windows auto update"
'14 welchia registry entry
DeleteValue HKEY_LOCAL_MACHINE, _
"SYSTEM\CurrentControlSet\Services", "RpcPatch"
DeleteValue HKEY_LOCAL_MACHINE, _
"SYSTEM\CurrentControlSet\Services", "RpcTftpd"

'15 p spider backdoor
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "mssysint"
        
'16 yaha worm
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "MicrosoftServiceManager"
         
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "MicrosoftServiceManager"
        
'17 lala backdoor
        
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "PNtask Services"

'18 nibu backdoor
DeleteValue HKEY_LOCAL_MACHINE, _
"\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "load32"

'19 love virus registry entry
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "MSKernel32"
                
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Win32DLL"
                
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "WIN-BUGSFIX"
                
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "WWinFAT32=WinFAT32.EXE"
        
'20 cone keylogger registry entries
        
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects", "{1E1B2879-88FF-11D3-8D96-D7ACAC95951A}"
        
DeleteValue HKEY_LOCAL_MACHINE, _
"CLASSES\CLSID", "{1E1B2879-88FF-11D3-8D96-D7ACAC95951A}"

DeleteValue HKEY_LOCAL_MACHINE, _
"CLASSES\Interface", "{1E1B2879-88FF-11D3-8D96-D7ACAC95951A}"

DeleteValue HKEY_LOCAL_MACHINE, _
"CLASSES\TypeLib", "{1E1B2879-88FF-11D3-8D96-D7ACAC95951A}"

DeleteValue HKEY_LOCAL_MACHINE, _
"CLASSES\TypeLib", "{1E1B2879-88FF-11D3-8D96-D7ACAC95951A}"
        
'21 datom worm
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "MSVXD"

'22 sircam worm
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices", "Driver32."
    
'23 intruzzo trojan
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "HPSFD %System%\GLIDELOAD.exe /s"
    
'24 sworpta trojan
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Internet Explorer\Main\", "Start Page"
        
DeleteValue HKEY_LOCAL_MACHINE, _
"SOFTWARE\Microsoft\Internet Explorer\Main\", "Startpagina"
    
'''''below is how to delete a full key put all full
'''''key deletions under this for easy reference

'25 sub seven registry removal
DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\ENC"
    
'26 sircam worm
DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\SirCam"
    
'27 irc rpc bot
DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\TFTPD32"
    
'28 ms blast whole key kill?
DeleteKey HKEY_LOCAL_MACHINE, "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\windows auto update"
    
'29 sworpta trojan
DeleteKey HKEY_LOCAL_MACHINE, "HKEY_CURRENT_USER\Software\SWCaller\"


'============================HKEY_CURRENT_USER============================================
'DeleteKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce"
'DeleteKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnceEx"
'DeleteKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\IMAIL"
'DeleteKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\MAPI"
'DeleteKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\MSFS"
'DeleteKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents"
'DeleteKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run"
'============================HKEY_LOCAL_MACHINE===========================================
'DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunOnce"
'DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunOnceEx"
'DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\IMAIL"
'DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\MAPI"
'DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents\MSFS"
'DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\OptionalComponents"
'DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run"
'DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices"
'DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServicesOnce"
'=================================HKEY_USER=============================================
'DeleteKey HKEY_USERS, ".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Run"

Next X

ressys
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub tmrmain_Timer()
tmrmain.Enabled = False
makeback
R1TMR.Enabled = True
End Sub
Private Sub boom_Timer()
Dim path As String
Dim comd As String
Dim chk As String
On Error Resume Next
chk = Right(App.path, 1)
    
    If chk = "\" Then
        path = App.path & "data\thirdparty\sd.exe"
    Else
        path = App.path & "\data\thirdparty\sd.exe"
    End If
    
    comd = path & " 4 FORCE"
    
l = l + 1
pb.Value = l
canbut.Caption = "Cancel (" & (10 - l) & ")"
If l = 10 Then DOShell comd, vbHide
If l = 19 Then MsgBox "It seems like Painkiller can't reboot your pc automatically, please reboot manually!", vbExclamation, "PainKiller"
If l = 20 Then End
End Sub
