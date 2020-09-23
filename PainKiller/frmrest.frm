VERSION 5.00
Begin VB.Form frmrest 
   BorderStyle     =   0  'None
   Caption         =   "restore"
   ClientHeight    =   7935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmrest.frx":0000
   ScaleHeight     =   7935
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1440
      Width           =   5295
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   2760
   End
   Begin PainKiller.XandersXPProgressBar pb 
      Height          =   255
      Left            =   90
      TabIndex        =   0
      Top             =   840
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
      Max             =   16
      Scrolling       =   5
   End
End
Attribute VB_Name = "frmrest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Private Sub Form_Load()
Me.Height = 1350
Me.Width = 5250
'Me.MousePointer  = vbNormal
'''''Me.MousePointer = vbHourglass
Timer1.Enabled = True
End Sub

Private Sub makere()
Dim path As String
Dim comd As String
Dim chk As String
Dim v As Integer
Dim path1 As String
On Error Resume Next

chk = Right(App.path, 1)
    
    If chk = "\" Then
        path = App.path & "backup\"
        path1 = App.path & "data\thirdparty\reg.exe"
    Else
        path = App.path & "\backup\"
        path1 = App.path & "\data\thirdparty\reg.exe"
    End If
    Sleep 100
For v = 0 To 16
    comd = Chr(34) & path1 & Chr(34) & " IMPORT " & Chr(34) & path & "regback" & v & ".reg" & Chr(34)
    'MsgBox comd
    'Text1 = comd
    DOShell comd, vbNormal
    pb.Value = pb.Value + 1
    Sleep 1000
Next v
Me.MousePointer = vbDefault
'Me.MousePointer  = vbHourglass
    MsgBox "All Registry settings had been restored successfully!", vbInformation, "PainKiller"
    Unload Me
End Sub


Private Sub Timer1_Timer()
Timer1.Enabled = False
makere
End Sub
