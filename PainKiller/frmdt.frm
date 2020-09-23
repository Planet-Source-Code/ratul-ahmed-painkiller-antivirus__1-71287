VERSION 5.00
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form frmdt 
   BorderStyle     =   0  'None
   Caption         =   "Detecting.."
   ClientHeight    =   4185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   Picture         =   "frmdt.frx":0000
   ScaleHeight     =   4185
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer sttmr 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   2280
   End
   Begin VB.ListBox lstr 
      Height          =   1425
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   855
   End
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   960
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "frmdt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

Private Sub Form_Load()
Me.Height = "1500"
Me.Width = "3000"
'Sleep 100
'Call CheckRMedias
sttmr.Enabled = True
End Sub



Private Sub sttmr_Timer()
Call CheckRMedias
End Sub

Private Sub SysInfo1_DeviceArrival(ByVal DeviceType As Long, ByVal DeviceID As Long, ByVal DeviceName As String, ByVal DeviceData As Long)
Call CheckRMedias
End Sub

Private Sub SysInfo1_DeviceRemoveComplete(ByVal DeviceType As Long, ByVal DeviceID As Long, ByVal DeviceName As String, ByVal DeviceData As Long)
Call CheckRMedias
End Sub

Sub CheckRMedias()
Dim FSO As New Scripting.FileSystemObject, drv As Scripting.Drive
On Error Resume Next
'check if a drive is present..
For Each drv In FSO.Drives

    'check if drive exist..
    If drv.IsReady Then

        'if drive is a removable drive..
        If drv.DriveType = Removable Then

            lstr.AddItem drv.DriveLetter & ":\"

        End If

    End If

Next
Sleep 100
dodrv
End Sub

Private Sub dodrv()
Dim i As Integer
On Error Resume Next
If lstr.ListCount = 0 Then
    MsgBox "No Device found, please check connection!!", vbExclamation, "PainKiller"
    sttmr.Enabled = False
    frmdt.Visible = False
    frmmain.Visible = True
    Unload frmdt
Else

        For i = 0 To lstr.ListCount
            If lstr.List(i) <> "" Then frmscan.lstdrv.AddItem lstr.List(i)
        Next i
        'Sleep 100
        frmscan.chkAll.Value = 1
        frmscan.chkVscan.Value = 1
        frmscan.Visible = True
        'Exit Sub
        Unload frmdt

End If
End Sub
