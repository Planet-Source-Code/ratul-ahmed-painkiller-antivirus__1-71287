VERSION 5.00
Begin VB.Form frmregback 
   BorderStyle     =   0  'None
   Caption         =   "backup"
   ClientHeight    =   1350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   Picture         =   "frmregback.frx":0000
   ScaleHeight     =   1350
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   6480
      Top             =   240
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1560
      Width           =   12135
   End
   Begin PainKiller.XandersXPProgressBar pb 
      Height          =   255
      Left            =   100
      TabIndex        =   0
      Top             =   720
      Width           =   5055
      _ExtentX        =   8916
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
      Color           =   4210752
      Max             =   10
      Scrolling       =   5
      Value           =   100
   End
End
Attribute VB_Name = "frmregback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
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
    pb.Value = 10
    Sleep 1000
    Me.MousePointer = vbDefault
    'Me.MousePointer  = vbHourglass
    MsgBox "Registry Backup Completed!!", vbInformation, "Painkiller"
    frmmain.Visible = True
    Unload Me
End Sub

Private Sub Form_Load()
pb.Value = 0
'''''Me.MousePointer = vbHourglass
End Sub

Private Sub Timer1_Timer()
makeback
Timer1.Enabled = False
End Sub
