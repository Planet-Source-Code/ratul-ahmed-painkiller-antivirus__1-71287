VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9e.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmscan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Scanner"
   ClientHeight    =   10305
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12285
   Icon            =   "frmscan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmscan.frx":74F2
   ScaleHeight     =   10305
   ScaleWidth      =   12285
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox autonmlst 
      Height          =   645
      Left            =   8880
      TabIndex        =   25
      Top             =   3960
      Width           =   3135
   End
   Begin VB.ListBox autolst 
      Height          =   645
      Left            =   8880
      TabIndex        =   24
      Top             =   3240
      Width           =   3135
   End
   Begin VB.ListBox lstdrv 
      Height          =   645
      Left            =   8880
      TabIndex        =   21
      Top             =   9240
      Width           =   3135
   End
   Begin PainKiller.StylerButton butucall 
      Height          =   375
      Left            =   6120
      TabIndex        =   20
      Top             =   6045
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Uncheck All"
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
   Begin PainKiller.StylerButton butchkall 
      Height          =   375
      Left            =   4920
      TabIndex        =   19
      Top             =   6045
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Check All"
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
   Begin VB.FileListBox lststr 
      Height          =   675
      Left            =   8880
      TabIndex        =   18
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Timer delay 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   0
      Top             =   7440
   End
   Begin PainKiller.StylerButton btnCancel 
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   6045
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Cancel"
      CaptionDisableColor=   12236471
      CaptionEffectColor=   16777215
      Theme           =   4
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
   Begin PainKiller.StylerButton btnDisinfect 
      Height          =   375
      Left            =   7320
      TabIndex        =   16
      Top             =   6050
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Disinfect"
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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   600
      TabIndex        =   15
      Top             =   7920
      Width           =   975
   End
   Begin VB.ListBox fregname 
      Height          =   1035
      Left            =   8880
      TabIndex        =   14
      Top             =   7560
      Width           =   3135
   End
   Begin VB.ListBox regname 
      Height          =   1035
      Left            =   8880
      TabIndex        =   13
      Top             =   5400
      Width           =   3135
   End
   Begin VB.ListBox apploc 
      Height          =   1035
      Left            =   8880
      TabIndex        =   12
      Top             =   6480
      Width           =   3135
   End
   Begin VB.ListBox appname 
      Height          =   645
      Left            =   8880
      TabIndex        =   11
      Top             =   8640
      Width           =   3135
   End
   Begin VB.CheckBox chkVscan 
      Caption         =   "Vscan"
      Height          =   255
      Left            =   8880
      TabIndex        =   7
      Top             =   1320
      Width           =   3015
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "Check All"
      Height          =   255
      Left            =   8880
      TabIndex        =   6
      Top             =   1080
      Width           =   3015
   End
   Begin VB.CheckBox chkautodel 
      Caption         =   "Auto Delete"
      Height          =   255
      Left            =   8880
      TabIndex        =   5
      Top             =   840
      Width           =   3015
   End
   Begin VB.CheckBox chkMedia 
      Caption         =   "Check Media"
      Height          =   255
      Left            =   8880
      TabIndex        =   4
      Top             =   600
      Width           =   3015
   End
   Begin VB.CheckBox chkdrv 
      Caption         =   "Check Drive"
      Height          =   255
      Left            =   8880
      TabIndex        =   3
      Top             =   360
      Width           =   3015
   End
   Begin VB.CheckBox chksystem 
      Caption         =   "Check System"
      Height          =   255
      Left            =   8880
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin MSComctlLib.ListView lstvr 
      Height          =   3735
      Index           =   0
      Left            =   3120
      TabIndex        =   1
      Top             =   1875
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   6588
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "nam"
         Object.Tag             =   "1"
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "loc"
         Object.Tag             =   "2"
         Text            =   "Location"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "reg"
         Object.Tag             =   "3"
         Text            =   "Regname"
         Object.Width           =   1764
      EndProperty
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash swfs 
      Height          =   3615
      Left            =   360
      TabIndex        =   0
      Top             =   1920
      Width           =   2535
      _cx             =   4471
      _cy             =   6376
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "NoScale"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin MSComctlLib.ListView lstvr 
      Height          =   855
      Index           =   1
      Left            =   8880
      TabIndex        =   10
      Top             =   1680
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1508
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "nam"
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "loc"
         Text            =   "Location"
         Object.Width           =   6350
      EndProperty
   End
   Begin VB.Label autopth 
      Caption         =   "autopth"
      Height          =   255
      Left            =   4560
      TabIndex        =   23
      Top             =   7800
      Width           =   4215
   End
   Begin VB.Label autonm 
      Caption         =   "autonm"
      Height          =   255
      Left            =   4560
      TabIndex        =   22
      Top             =   7440
      Width           =   4215
   End
   Begin VB.Label temptxt 
      Caption         =   "temptxt"
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   6960
      Width           =   8655
   End
   Begin VB.Label lblpath 
      Caption         =   "lblpath"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   6600
      Width           =   8655
   End
End
Attribute VB_Name = "frmscan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Const MF_BYPOSITION = &H400&
Private ReadyToClose As Boolean

Dim reg As CRegistry
Dim env As CEnvironment
Dim chkpath As String
Dim hKey As Long, LCount As Long, i As Long
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Private Sub RemoveMenus(frm As Form, _
    remove_restore As Boolean, _
    remove_move As Boolean, _
    remove_size As Boolean, _
    remove_minimize As Boolean, _
    remove_maximize As Boolean, _
    remove_seperator As Boolean, _
    remove_close As Boolean)
Dim hMenu As Long
    
    ' Get the form's system menu handle.
    hMenu = GetSystemMenu(hWnd, False)
    
    If remove_close Then DeleteMenu hMenu, 6, MF_BYPOSITION
    If remove_seperator Then DeleteMenu hMenu, 5, MF_BYPOSITION
    If remove_maximize Then DeleteMenu hMenu, 4, MF_BYPOSITION
    If remove_minimize Then DeleteMenu hMenu, 3, MF_BYPOSITION
    If remove_size Then DeleteMenu hMenu, 2, MF_BYPOSITION
    If remove_move Then DeleteMenu hMenu, 1, MF_BYPOSITION
    If remove_restore Then DeleteMenu hMenu, 0, MF_BYPOSITION
End Sub
Private Sub btnCancel_Click()
frmscan.Visible = False
frmmain.Visible = True
ReadyToClose = True
Unload frmscan
End Sub

Private Sub btnDisinfect_Click()
Dim v As Integer
If btnDisinfect.Caption = "Disinfect" Then process_Handler
If btnDisinfect.Caption = "OK" Then
    If chkVscan.Value = 1 Then
        For v = 0 To lstdrv.ListCount
            frmmscan.lstdir.AddItem frmscan.lstdrv.List(v)
            frmmscan.Visible = True
        Next v
    Else
        frmscan.Visible = False
        frmmain.Visible = True
    End If
''frmx.Visible = True
frmscan.Visible = False
frmscan.Visible = False
'frmmain.Visible = True
ReadyToClose = True
Unload frmscan
Unload frmscan
End If


End Sub

Private Sub butchkall_Click()
Dim v As Integer
On Error Resume Next
For v = 0 To lstvr(0).ListItems.count
    lstvr(0).ListItems.Item(v).Checked = True
Next v
End Sub

Private Sub butucall_Click()
Dim v As Integer
On Error Resume Next
For v = 0 To lstvr(0).ListItems.count
    lstvr(0).ListItems.Item(v).Checked = False
Next v
End Sub

Private Sub Command1_Click()
Dim v As Integer
Dim a As Integer
On Error Resume Next
For v = 0 To lstvr(0).ListItems.count

lstvr(0).ListItems.Item(v).Checked = True
 
Next v

End Sub

Private Sub delay_Timer()
c_all
delay.Enabled = False
End Sub

Private Sub Form_Load()
Unload frmdt
frmmain.Visible = False
'Me.Visible = False
RemoveMenus Me, False, False, _
        False, False, False, True, True
chksystem.Value = 1
'''''Me.MousePointer = vbHourglass

If chksystem.Value = 1 Then
    Me.Height = 6885
    Me.Width = 8670
    Me.Visible = True
    chkpath = Right(App.path, 1)
    
    If chkpath = "\" Then
    swfs.Movie = App.path & "data\anim\scan.swf"
    Else
    swfs.Movie = App.path & "\data\anim\scan.swf"
    End If
    delay.Enabled = True
Else
    'MsgBox "You havent Select me"
    'End
    'NOTHING
End If
End Sub

Private Sub c_all()
Dim exename As String
Dim exepath As String
Dim exename1 As String
Dim exepath1 As String
Dim file2chk As String
Dim yesauto As String
Dim j As Integer
Dim X As Integer
Dim C As Long
Dim q
Dim w
Dim e
Dim nm As String
Dim bin As String
Dim dinint As Integer
Dim add As String

On Error Resume Next


'================================================================================================
                                                                       '\ Genarate For HKLM RUN |
                                                                         '-----------------------
hKey = OpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run")
LCount = GetCount(hKey, Values)
For j = 0 To LCount - 1
    C = j
    exename = EnumValue(hKey, C) '------------------------String Of Startup program
    exepath = GetKeyValue(hKey, EnumValue(hKey, C)) '-----Path Of Startup Program
    
                       '_____________________
                      '/ Get Final Location /__/
    '--------------------------------------/
    lblpath = exepath                   '--
    chkpath = Left(lblpath, 1)          '--
    If chkpath = Chr(34) Then           '--
        X = Len(lblpath)                '--
        lblpath = Right(lblpath, X - 1) '--
        X = Len(lblpath)                '--
        lblpath = Left(lblpath, X - 1)  '--
        'MsgBox lblpath                 '--
    Else                                '--
        'MsgBox lblpath                 '--
    End If                              '--
    '--------------------------------------
                '____________________________
               '/ Add items to the List Box /__/
    '--------------------------------------/
    apploc.AddItem lblpath              '--
    regname.AddItem exename             '--
    With lstvr(1).ListItems.add         '--
        .Text = exename                 '--
        .SubItems(1) = lblpath          '--
    End With                            '--
    '--------------------------------------
Next j

'================================================================================================
                                                                   '\ Genarate For HKLM RUNONCE |
                                                                     '---------------------------

hKey = OpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunOnce")
LCount = GetCount(hKey, Values)
For j = 0 To LCount - 1
    C = j
    exename = EnumValue(hKey, C) '------------------------String Of Startup program
    exepath = GetKeyValue(hKey, EnumValue(hKey, C)) '-----Path Of Startup Program
    
                       '_____________________
                      '/ Get Final Location /__/
    '--------------------------------------/
    lblpath = exepath                   '--
    chkpath = Left(lblpath, 1)          '--
    If chkpath = Chr(34) Then           '--
        X = Len(lblpath)                '--
        lblpath = Right(lblpath, X - 1) '--
        X = Len(lblpath)                '--
        lblpath = Left(lblpath, X - 1)  '--
        'MsgBox lblpath                 '--
    Else                                '--
        'MsgBox lblpath                 '--
    End If                              '--
    '--------------------------------------
                '____________________________
               '/ Add items to the List Box /__/
    '--------------------------------------/
    apploc.AddItem lblpath              '--
    regname.AddItem exename             '--
    With lstvr(1).ListItems.add         '--
        .Text = exename                 '--
        .SubItems(1) = lblpath          '--
    End With                            '--
    '--------------------------------------
Next j

'================================================================================================
                                                                 '\ Genarate For HKLM RUNONCEEX |
                                                                   '-----------------------------
                                                                         
hKey = OpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunOnceEx")
LCount = GetCount(hKey, Values)
For j = 0 To LCount - 1
    C = j
    exename = EnumValue(hKey, C) '------------------------String Of Startup program
    exepath = GetKeyValue(hKey, EnumValue(hKey, C)) '-----Path Of Startup Program
    
                       '_____________________
                      '/ Get Final Location /__/
    '--------------------------------------/
    lblpath = exepath                   '--
    chkpath = Left(lblpath, 1)          '--
    If chkpath = Chr(34) Then           '--
        X = Len(lblpath)                '--
        lblpath = Right(lblpath, X - 1) '--
        X = Len(lblpath)                '--
        lblpath = Left(lblpath, X - 1)  '--
        'MsgBox lblpath                 '--
    Else                                '--
        'MsgBox lblpath                 '--
    End If                              '--
    '--------------------------------------
                '____________________________
               '/ Add items to the List Box /__/
    '--------------------------------------/
    apploc.AddItem lblpath              '--
    regname.AddItem exename             '--
    With lstvr(1).ListItems.add         '--
        .Text = exename                 '--
        .SubItems(1) = lblpath          '--
    End With                            '--
    '--------------------------------------
Next j


'================================================================================================
                                                               '\ Genarate For HKLM RunServices |
                                                                 '-------------------------------
                                                                         
hKey = OpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices")
LCount = GetCount(hKey, Values)
For j = 0 To LCount - 1
    C = j
    exename = EnumValue(hKey, C) '------------------------String Of Startup program
    exepath = GetKeyValue(hKey, EnumValue(hKey, C)) '-----Path Of Startup Program
    
                       '_____________________
                      '/ Get Final Location /__/
    '--------------------------------------/
    lblpath = exepath                   '--
    chkpath = Left(lblpath, 1)          '--
    If chkpath = Chr(34) Then           '--
        X = Len(lblpath)                '--
        lblpath = Right(lblpath, X - 1) '--
        X = Len(lblpath)                '--
        lblpath = Left(lblpath, X - 1)  '--
        'MsgBox lblpath                 '--
    Else                                '--
        'MsgBox lblpath                 '--
    End If                              '--
    '--------------------------------------
                '____________________________
               '/ Add items to the List Box /__/
    '--------------------------------------/
    apploc.AddItem lblpath              '--
    regname.AddItem exename             '--
    With lstvr(1).ListItems.add         '--
        .Text = exename                 '--
        .SubItems(1) = lblpath          '--
    End With                            '--
    '--------------------------------------
Next j


'================================================================================================
                                                           '\ Genarate For HKLM RunServicesOnce |
                                                             '-----------------------------------
                                                                         
hKey = OpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServicesOnce")
LCount = GetCount(hKey, Values)
For j = 0 To LCount - 1
    C = j
    exename = EnumValue(hKey, C) '------------------------String Of Startup program
    exepath = GetKeyValue(hKey, EnumValue(hKey, C)) '-----Path Of Startup Program
    
                       '_____________________
                      '/ Get Final Location /__/
    '--------------------------------------/
    lblpath = exepath                   '--
    chkpath = Left(lblpath, 1)          '--
    If chkpath = Chr(34) Then           '--
        X = Len(lblpath)                '--
        lblpath = Right(lblpath, X - 1) '--
        X = Len(lblpath)                '--
        lblpath = Left(lblpath, X - 1)  '--
        'MsgBox lblpath                 '--
    Else                                '--
        'MsgBox lblpath                 '--
    End If                              '--
    '--------------------------------------
                '____________________________
               '/ Add items to the List Box /__/
    '--------------------------------------/
    apploc.AddItem lblpath              '--
    regname.AddItem exename             '--
    With lstvr(1).ListItems.add         '--
        .Text = exename                 '--
        .SubItems(1) = lblpath          '--
    End With                            '--
    '--------------------------------------
Next j

'################################################################################################

'================================================================================================
                                                                       '\ Genarate For HKCU RUN |
                                                                         '-----------------------
hKey = OpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run")
LCount = GetCount(hKey, Values)
For j = 0 To LCount - 1
    C = j
    exename = EnumValue(hKey, C) '------------------------String Of Startup program
    exepath = GetKeyValue(hKey, EnumValue(hKey, C)) '-----Path Of Startup Program
    
                       '_____________________
                      '/ Get Final Location /__/
    '--------------------------------------/
    lblpath = exepath                   '--
    chkpath = Left(lblpath, 1)          '--
    If chkpath = Chr(34) Then           '--
        X = Len(lblpath)                '--
        lblpath = Right(lblpath, X - 1) '--
        X = Len(lblpath)                '--
        lblpath = Left(lblpath, X - 1)  '--
        'MsgBox lblpath                 '--
    Else                                '--
        'MsgBox lblpath                 '--
    End If                              '--
    '--------------------------------------
                '____________________________
               '/ Add items to the List Box /__/
    '--------------------------------------/
    apploc.AddItem lblpath              '--
    regname.AddItem exename             '--
    With lstvr(1).ListItems.add         '--
        .Text = exename                 '--
        .SubItems(1) = lblpath          '--
    End With                            '--
    '--------------------------------------
Next j

'================================================================================================
                                                                   '\ Genarate For HKCU RUNONCE |
                                                                     '---------------------------

hKey = OpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce")
LCount = GetCount(hKey, Values)
For j = 0 To LCount - 1
    C = j
    exename = EnumValue(hKey, C) '------------------------String Of Startup program
    exepath = GetKeyValue(hKey, EnumValue(hKey, C)) '-----Path Of Startup Program
    
                       '_____________________
                      '/ Get Final Location /__/
    '--------------------------------------/
    lblpath = exepath                   '--
    chkpath = Left(lblpath, 1)          '--
    If chkpath = Chr(34) Then           '--
        X = Len(lblpath)                '--
        lblpath = Right(lblpath, X - 1) '--
        X = Len(lblpath)                '--
        lblpath = Left(lblpath, X - 1)  '--
        'MsgBox lblpath                 '--
    Else                                '--
        'MsgBox lblpath                 '--
    End If                              '--
    '--------------------------------------
                '____________________________
               '/ Add items to the List Box /__/
    '--------------------------------------/
    apploc.AddItem lblpath              '--
    regname.AddItem exename             '--
    With lstvr(1).ListItems.add         '--
        .Text = exename                 '--
        .SubItems(1) = lblpath          '--
    End With                            '--
    '--------------------------------------
Next j

'================================================================================================
                                                                 '\ Genarate For HKCU RUNONCEEX |
                                                                   '-----------------------------
                                                                         
hKey = OpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnceEx")
LCount = GetCount(hKey, Values)
For j = 0 To LCount - 1
    C = j
    exename = EnumValue(hKey, C) '------------------------String Of Startup program
    exepath = GetKeyValue(hKey, EnumValue(hKey, C)) '-----Path Of Startup Program
    
                       '_____________________
                      '/ Get Final Location /__/
    '--------------------------------------/
    lblpath = exepath                   '--
    chkpath = Left(lblpath, 1)          '--
    If chkpath = Chr(34) Then           '--
        X = Len(lblpath)                '--
        lblpath = Right(lblpath, X - 1) '--
        X = Len(lblpath)                '--
        lblpath = Left(lblpath, X - 1)  '--
        'MsgBox lblpath                 '--
    Else                                '--
        'MsgBox lblpath                 '--
    End If                              '--
    '--------------------------------------
                '____________________________
               '/ Add items to the List Box /__/
    '--------------------------------------/
    apploc.AddItem lblpath              '--
    regname.AddItem exename             '--
    With lstvr(1).ListItems.add         '--
        .Text = exename                 '--
        .SubItems(1) = lblpath          '--
    End With                            '--
    '--------------------------------------
Next j


'================================================================================================
                                                               '\ Genarate For HKCU RunServices |
                                                                 '-------------------------------
                                                                         
hKey = OpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunServices")
LCount = GetCount(hKey, Values)
For j = 0 To LCount - 1
    C = j
    exename = EnumValue(hKey, C) '------------------------String Of Startup program
    exepath = GetKeyValue(hKey, EnumValue(hKey, C)) '-----Path Of Startup Program
    
                       '_____________________
                      '/ Get Final Location /__/
    '--------------------------------------/
    lblpath = exepath                   '--
    chkpath = Left(lblpath, 1)          '--
    If chkpath = Chr(34) Then           '--
        X = Len(lblpath)                '--
        lblpath = Right(lblpath, X - 1) '--
        X = Len(lblpath)                '--
        lblpath = Left(lblpath, X - 1)  '--
        'MsgBox lblpath                 '--
    Else                                '--
        'MsgBox lblpath                 '--
    End If                              '--
    '--------------------------------------
                '____________________________
               '/ Add items to the List Box /__/
    '--------------------------------------/
    apploc.AddItem lblpath              '--
    regname.AddItem exename             '--
    With lstvr(1).ListItems.add         '--
        .Text = exename                 '--
        .SubItems(1) = lblpath          '--
    End With                            '--
    '--------------------------------------
Next j


'================================================================================================
                                                           '\ Genarate For HKCU RunServicesOnce |
                                                             '-----------------------------------
                                                                         
hKey = OpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunServicesOnce")
LCount = GetCount(hKey, Values)
For j = 0 To LCount - 1
    C = j
    exename = EnumValue(hKey, C) '------------------------String Of Startup program
    exepath = GetKeyValue(hKey, EnumValue(hKey, C)) '-----Path Of Startup Program
    
                       '_____________________
                      '/ Get Final Location /__/
    '--------------------------------------/
    lblpath = exepath                   '--
    chkpath = Left(lblpath, 1)          '--
    If chkpath = Chr(34) Then           '--
        X = Len(lblpath)                '--
        lblpath = Right(lblpath, X - 1) '--
        X = Len(lblpath)                '--
        lblpath = Left(lblpath, X - 1)  '--
        'MsgBox lblpath                 '--
    Else                                '--
        'MsgBox lblpath                 '--
    End If                              '--
    '--------------------------------------
                '____________________________
               '/ Add items to the List Box /__/
    '--------------------------------------/
    apploc.AddItem lblpath              '--
    regname.AddItem exename             '--
    With lstvr(1).ListItems.add         '--
        .Text = exename                 '--
        .SubItems(1) = lblpath          '--
    End With                            '--
    '--------------------------------------
Next j

'================================================================================================
                                                                   '\ Genarate For User Startup |
                                                                     '---------------------------
                                                             
lststr.FileName = CheckFolderID(StartUp)
hKey = CheckFolderID(StartUp)
LCount = lststr.ListCount

For j = 0 To LCount - 1
    C = j
    exename = lststr.List(j)  '------------------------String Of Startup program
    exepath = CheckFolderID(StartUp) & "\" & lststr.List(j)  '-----Path Of Startup Program
    
                       '_____________________
                      '/ Get Final Location /__/
    '--------------------------------------/
    lblpath = exepath                   '--
    chkpath = Left(lblpath, 1)          '--
    If chkpath = Chr(34) Then           '--
        X = Len(lblpath)                '--
        lblpath = Right(lblpath, X - 1) '--
        X = Len(lblpath)                '--
        lblpath = Left(lblpath, X - 1)  '--
        'MsgBox lblpath                 '--
    Else                                '--
        'MsgBox lblpath                 '--
    End If                              '--
    '--------------------------------------
    
    temptxt = GetTarget(lblpath)
    q = Mid(temptxt, InStrRev(temptxt, "\") + 1)
    e = InStrRev(q, ".") - 1
    w = Mid(q, InStrRev(q, ".") - e)
    exename = w
                '____________________________
               '/ Add items to the List Box /__/
    '--------------------------------------/
    apploc.AddItem temptxt              '--
    regname.AddItem exename             '--
    With lstvr(1).ListItems.add         '--
        .Text = exename                 '--
        .SubItems(1) = lblpath          '--
    End With                            '--
    '--------------------------------------
Next j

'================================================================================================
                                                                   '\ Genarate For User Startup |
                                                                     '---------------------------
                                                             
lststr.FileName = CheckFolderID(Common_StartUp)
hKey = CheckFolderID(Common_StartUp)
LCount = lststr.ListCount

For j = 0 To LCount - 1
    C = j
    exename = lststr.List(j)  '------------------------String Of Startup program
    exepath = CheckFolderID(Common_StartUp) & "\" & lststr.List(j)  '-----Path Of Startup Program
    
                       '_____________________
                      '/ Get Final Location /__/
    '--------------------------------------/
    lblpath = exepath                   '--
    chkpath = Left(lblpath, 1)          '--
    If chkpath = Chr(34) Then           '--
        X = Len(lblpath)                '--
        lblpath = Right(lblpath, X - 1) '--
        X = Len(lblpath)                '--
        lblpath = Left(lblpath, X - 1)  '--
        'MsgBox lblpath                 '--
    Else                                '--
        'MsgBox lblpath                 '--
    End If                              '--
    '--------------------------------------
    
    temptxt = GetTarget(lblpath)
    q = Mid(temptxt, InStrRev(temptxt, "\") + 1)
    e = InStrRev(q, ".") - 1
    w = Mid(q, InStrRev(q, ".") - e)
    exename = w
    
                '____________________________
               '/ Add items to the List Box /__/
    '--------------------------------------/
    apploc.AddItem lblpath              '--
    regname.AddItem exename             '--
    With lstvr(1).ListItems.add         '--
        .Text = exename                 '--
        .SubItems(1) = lblpath          '--
    End With                            '--
    '--------------------------------------
Next j


'================================================================================================
                                                                   '\ Genarate For Drive Autorun |
                                                                     '---------------------------
                                                             

'hKey = CheckFolderID(Common_StartUp)
LCount = lstdrv.ListCount

For j = 0 To LCount - 1
    C = j
    exename1 = lstdrv.List(j)  '------------------------String Of Startup program
    exepath1 = lstdrv.List(j) & "autorun.inf"   '-----Path Of Startup Program
    
    
    file2chk = lstdrv.List(j) & "autorun.inf"  '--------------------File Path declaretion
    yesauto = FileExists(file2chk) '------------------------------Cheack the file Existence
    If yesauto = True Then
        Open file2chk For Binary As #1
        bin = Space$(LOF(1))
        Get #1, , bin
        Close #1
        nm = readinf(file2chk)
        add = lstdrv.List(j) & nm
        autopth = lstdrv.List(j) & "autorun.inf"
        autonm = nm
        autolst.AddItem autopth
        autonmlst.AddItem autonm
    End If
      
     
Next j

'================================================================================================
listVir
Exit Sub
End Sub

Private Sub listVir()
Dim X As Integer
Dim j As Integer
Dim a As String
Dim d As String
Dim n As String
Dim wordcut As Integer
Dim fnl As String
Dim chkdrv As String
For j = 0 To apploc.ListCount

                            '_____________________________________
                           '/ Genarate True Mistrusted Component /__/
    '-----------------------------------------------------------/
    temptxt = apploc.List(j)                                 '--
    X = Len(temptxt)                                         '--
    chkdrv = Left(temptxt, 3)                                '--
    chkpath = Left(temptxt, 16)                              '--
    If chkpath = chkdrv & "Program Files" Then               '--
    'NOTHING                                                 '--
    Else                                                     '--
        If apploc.List(j) = "" Then                          '--
        'NOTHING                                             '--
        Else                                                 '--
            fregname.AddItem regname.List(j)                 '--
            a = Mid(temptxt, InStrRev(temptxt, "\") + 1)     '--
            d = InStrRev(a, ".") - 1                         '--
            n = Mid(a, InStrRev(a, ".") - d)                 '--
            wordcut = d + 4                                  '--
            fnl = Left(n, wordcut)                           '--
            appname.AddItem fnl                              '--\\\
                             '____________________________      '--
                            '/ Add items to the List Box /__/   '--
                 '--------------------------------------/       '--
                 With lstvr(0).ListItems.add         '--        '--
                     .Text = fnl                     '--        '--
                     .SubItems(1) = temptxt          '--        '--
                     .SubItems(2) = regname.List(j)  '--        '--
                 End With                            '--        '--
                '--------------------------------------         '--
        End If                                               '--///
    End If                                                   '--
    '-----------------------------------------------------------
    
Next j

j = 0

For j = 0 To autolst.ListCount

    If autolst.List(j) <> "" Then
        apploc.AddItem autolst.List(j)
        chkdrv = Left(autolst.List(j), 3)
        apploc.AddItem chkdrv & autonmlst.List(j)
        appname.AddItem autonmlst.List(j)
        appname.AddItem autonmlst.List(j)
                    '____________________________
                   '/ Add items to the List Box /__/
        '--------------------------------------/
        With lstvr(0).ListItems.add             '--
            .Text = autonmlst.List(j)           '--
            .SubItems(1) = autolst.List(j)      '--
            .SubItems(2) = autonmlst.List(j)    '--
        End With                                '--
        '-------------------------------------------
                    '____________________________
                   '/ Add items to the List Box /__/
        '--------------------------------------/
        With lstvr(0).ListItems.add                     '--
            .Text = autonmlst.List(j)                   '--
            .SubItems(1) = chkdrv & autonmlst.List(j)   '--
            .SubItems(2) = autonmlst.List(j)            '--
        End With                                        '--
        '-------------------------------------------
        
    End If
Next j


swfs.Stop
swfs.GotoFrame 0
btnDisinfect.Enabled = True
If appname.ListCount = 0 Then
    MsgBox "PainKiller had Finished searching your system and Found" & vbNewLine & "no Mistrusted Components!!", vbInformation, "PainKiller"
    btnDisinfect.Caption = "OK"
End If
If appname.ListCount <> 0 Then
    MsgBox "PainKiller had Finished searching your system and Found" & vbNewLine & appname.ListCount & " Mistrusted Components!!", vbExclamation, "PainKiller"
    btnDisinfect.Caption = "Disinfect"
    butchkall.Enabled = True
    butucall.Enabled = True
End If
Exit Sub
End Sub

Private Sub process_Handler()
Dim j As Integer
Dim v As Integer
Dim m As Integer
On Error Resume Next
appname.Clear
apploc.Clear
regname.Clear
fregname.Clear
lstvr(1).ListItems.Clear
For v = 0 To lstvr(0).ListItems.count
    If lstvr(0).ListItems.Item(v).Checked = True Then
        appname.AddItem lstvr(0).ListItems.Item(v).Text
        apploc.AddItem lstvr(0).ListItems.Item(v).SubItems(1)
        fregname.AddItem lstvr(0).ListItems.Item(v).SubItems(2)
    Else
        'NOTHING
    End If
Next v

For j = 0 To appname.ListCount
    If appname.List(j) = "" Then
    'NOTHING
    Else
    pSuspend appname.List(j), "s"
    End If
    'Sleep 100
Next j

j = 0

For j = 0 To appname.ListCount
    If appname.List(j) = "" Then
    'NOTHING
    Else
    pKill appname.List(j)
    End If
    'Sleep 100
Next j
delay.Enabled = False
File_Handler
End Sub

Private Sub File_Handler()
Dim j As Integer
Dim v As Integer
Dim fdback As String
For j = 0 To fregname.ListCount

    DeleteValue HKEY_LOCAL_MACHINE, _
    "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", fregname.List(j)
    
    DeleteValue HKEY_CURRENT_USER, _
    "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", fregname.List(j)
    
    DeleteValue HKEY_LOCAL_MACHINE, _
    "SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce", fregname.List(j)
    
    DeleteValue HKEY_CURRENT_USER, _
    "SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce", fregname.List(j)
    
    DeleteValue HKEY_LOCAL_MACHINE, _
    "SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx", fregname.List(j)
    
    DeleteValue HKEY_CURRENT_USER, _
    "SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx", fregname.List(j)
    
    DeleteValue HKEY_LOCAL_MACHINE, _
    "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices", fregname.List(j)
    
    DeleteValue HKEY_CURRENT_USER, _
    "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices", fregname.List(j)
    
    DeleteValue HKEY_LOCAL_MACHINE, _
    "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServicesOnce", fregname.List(j)
    
    DeleteValue HKEY_CURRENT_USER, _
    "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServicesOnce", fregname.List(j)
'MsgBox fregname.List(j)
Next j

j = 0

For j = 0 To apploc.ListCount
Sleep 250
If apploc.List(j) <> "" Then
    fdback = FileExists(apploc.List(j))
        If fdback = 0 Then MsgBox "The File " & Chr(34) & apploc.List(j) & Chr(34) & " cannot be located on the Path!!", vbExclamation, "File Dosen't Exists"
    pDelete (apploc.List(j))
    fdback = FileExists(apploc.List(j))
        If fdback = 1 Then pDelete (apploc.List(j))
    fdback = FileExists(apploc.List(j))
        If fdback = "True" Then MsgBox "The File " & Chr(34) & apploc.List(j) & Chr(34) & " can not be Deleted!!", vbCritical, "Error"
End If
delay.Enabled = False
Next j
Me.MousePointer = vbDefault
If apploc.ListCount = 0 Then
    MsgBox "No Mistrusted Components had been selected to Disinfect!!", vbInformation, "painKiller"
    Me.MousePointer = vbDefault
    Exit Sub
Else
    MsgBox apploc.ListCount & " Mistrusted Components had been Disinfected!!", vbInformation, "painKiller"
    Me.MousePointer = vbDefault
    lstvr(0).ListItems.Clear
        If chkVscan.Value = 1 Then
            For j = 0 To lstdrv.ListCount
                frmmscan.lstdir.AddItem frmscan.lstdrv.List(j)
            Next j
            frmscan.Visible = False
            frmmscan.Visible = True
        Else
            frmscan.Visible = False
            frmmain.Visible = True
        End If
    ''frmx.Visible = True
    frmscan.Visible = False
    Exit Sub
End If
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = Not ReadyToClose
End Sub
