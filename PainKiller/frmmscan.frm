VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmmscan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Scanner"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9450
   Icon            =   "frmmscan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmmscan.frx":74F2
   ScaleHeight     =   6480
   ScaleWidth      =   9450
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrmc 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2640
      Top             =   240
   End
   Begin VB.Timer strup 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   2160
      Top             =   240
   End
   Begin PainKiller.XandersXPProgressBar pber 
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   6120
      Width           =   5415
      _ExtentX        =   9551
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
      Scrolling       =   5
      Value           =   85
   End
   Begin PainKiller.StylerButton btncan 
      Height          =   375
      Left            =   8280
      TabIndex        =   4
      Top             =   6090
      Width           =   1095
      _ExtentX        =   1931
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
   Begin PainKiller.StylerButton btnDone 
      Height          =   375
      Left            =   7080
      TabIndex        =   3
      Top             =   6090
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Done"
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
   Begin VB.TextBox cmdx 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   8160
      Width           =   9495
   End
   Begin VB.ListBox lstdir 
      Height          =   1425
      Left            =   0
      TabIndex        =   1
      Top             =   6600
      Width           =   9495
   End
   Begin RichTextLib.RichTextBox txtOutputs 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8705
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmmscan.frx":A2B9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblsts 
      BackStyle       =   0  'Transparent
      Caption         =   "Scanning..."
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   6120
      Width           =   1245
   End
End
Attribute VB_Name = "frmmscan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Private Const WM_CLOSE = &H10
Dim winHwnd As Long, RetVal As Long
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Dim IngSuccess As Long

Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Const MF_BYPOSITION = &H400&
Private ReadyToClose As Boolean
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

Private Sub btnDone_Click()
frmmscan.Visible = False
ReadyToClose = True
Unload frmscan
frmrep.Visible = True
Unload Me
End Sub

'/////////////////////////////////////////////////////////////////////////Form Unload Events/////
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim scanpath As String
Dim chk As String
On Error Resume Next

    If chk = "\" Then
        scanpath = App.path & "data\thirdparty\scan_eng\Scan.exe"
    Else
        scanpath = App.path & "\data\thirdparty\scan_eng\Scan.exe"
    End If
    
    winHwnd = FindWindow(vbNullString, scanpath)
    PostMessage winHwnd, WM_CLOSE, 0&, 0&
'End
'frmmscan.Visible = False
'frmmain.Visible = True
Cancel = Not ReadyToClose
Unload Me
End Sub
Private Sub Form_Terminate()
Dim scanpath As String
Dim chk As String
On Error Resume Next
    
    If chk = "\" Then
        scanpath = App.path & "data\thirdparty\scan_eng\Scan.exe"
    Else
        scanpath = App.path & "\data\thirdparty\scan_eng\Scan.exe"
    End If
    winHwnd = FindWindow(vbNullString, scanpath)
    PostMessage winHwnd, WM_CLOSE, 0&, 0&
'End

End Sub
Private Sub Form_Unload(Cancel As Integer)
Dim scanpath As String
Dim chk As String
On Error Resume Next

    If chk = "\" Then
        scanpath = App.path & "data\thirdparty\scan_eng\Scan.exe"
    Else
        scanpath = App.path & "\data\thirdparty\scan_eng\Scan.exe"
    End If
    winHwnd = FindWindow(vbNullString, scanpath)
    PostMessage winHwnd, WM_CLOSE, 0&, 0&
'End
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub btncan_Click()
ReadyToClose = True
frmmscan.Visible = False
Unload frmscan
frmmain.Visible = True
Unload Me
End Sub

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub strup_Timer()
startscan

End Sub

Private Sub tmrmc_Timer()
Dim File As String
Dim bin As String
Dim scanpath As String
Dim chk As String
Dim R As String
On Error Resume Next
chk = Right(App.path, 1)
    
   
If chk = "\" Then
    scanpath = App.path & "data\thirdparty\scan_eng\Scan.exe"
    R = App.path & "data\thirdparty\scan_eng\rep.txt"
Else
    scanpath = App.path & "\data\thirdparty\scan_eng\Scan.exe"
    R = App.path & "\data\thirdparty\scan_eng\rep.txt"
End If
    
        File = R
        Open File For Binary As #1
        bin = Space$(LOF(1))
        Get #1, , bin
        Close #1
        txtOutputs.Text = bin

        winHwnd = FindWindow(vbNullString, scanpath)


        If winHwnd <> 0 Then
            lblsts.Caption = "Scanning..."
        Else
            lblsts.Caption = "Scan Finished"
            pber.Value = 100
            pber.AutoScroll = False
            tmrmc.Enabled = False
            Me.Visible = False
            Me.MousePointer = vbDefault
            MsgBox "Your Selected Directories had been scanned Successfully, Use the Scrollbar" & vbNewLine & "to see the results!", vbInformation, "Painkiller"
            Me.Visible = True
            btnDone.Enabled = True
        End If
        
End Sub


Private Sub startscan()
Dim sPath As String
Dim cmd As String
Dim X As Integer
Dim scanpath As String
Dim chk As String
Dim cline As String
Dim cline1 As String
Dim cline2 As String
Dim R As String
On Error Resume Next
cline = ""


For X = 0 To lstdir.ListCount
    If lstdir.List(X) = "" Then
        'NULL
    Else
        sPath = sPath & " " & Chr(34) & lstdir.List(X) & Chr(34)
    End If
Next X

    If chk = "\" Then
        scanpath = App.path & "data\thirdparty\scan_eng\Scan.exe"
        R = App.path & "data\thirdparty\scan_eng\rep.txt"
    Else
        scanpath = App.path & "\data\thirdparty\scan_eng\Scan.exe"
        R = App.path & "\data\thirdparty\scan_eng\rep.txt"
    End If
    
    cline = " /DEL"
    cline1 = " /CLEAN"
    cline2 = " /MOVE" & " " & Chr(34) & "C:\QUARANTINE" & Chr(34)
        
            cmd = scanpath & " " & sPath & cline & cline1 & cline2 & " /MIME /SUB /UNZIP /RPTCOR /RPTERR /RPTALL /STREAMS /REPORT " & Chr(34) & R & Chr(34)
            winHwnd = FindWindow(vbNullString, scanpath)
            If winHwnd <> 0 Then
                Sleep 1000
                tmrmc.Enabled = True
                strup.Enabled = False
                pber.AutoScroll = True
            Else
                DOShell cmd, vbHide
                Sleep 1200
                tmrmc.Enabled = True
                strup.Enabled = False
                pber.AutoScroll = True
            End If
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////

Private Sub Form_Load()
Dim scanpath As String
Dim chk As String
Dim fxist As String
On Error Resume Next
Me.Height = "6930"
Me.Width = "9540"
chk = Right(App.path, 1)
    
If chk = "\" Then
    scanpath = App.path & "data\thirdparty\scan_eng\Scan.exe"
Else
    scanpath = App.path & "\data\thirdparty\scan_eng\Scan.exe"
End If
    
fxist = FileExists(scanpath)

If fxist <> "True" Then
    MsgBox ("Please download Mcafee Update file from http://www.mcafee.com/apps/downloads/security_updates/superdat.asp?region=us&segment=enterprise" & vbNewLine & "then extract it in \painkiller\data\thirdparty\scan_eng\' folder using '/e' commandline"), vbInformation, "PainKiller"
    btnDone.Enabled = True
    Exit Sub
End If
    

'''''Me.MousePointer = vbHourglass
strup.Enabled = True
btnDone.Enabled = False
pber.Value = 0
RemoveMenus Me, False, False, _
        False, False, False, True, True
End Sub
