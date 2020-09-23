VERSION 5.00
Object = "{6A2F01E2-9EA2-48FD-B765-9F256FCEAFCF}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form frmsdrv 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Drive"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5130
   Icon            =   "frmsdrv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin vkUserContolsXP.vkListBox lstdrv 
      Height          =   2535
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   4471
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Sorted          =   0
      StyleCheckBox   =   -1  'True
      ListType        =   3
      IconSize        =   32
   End
   Begin PainKiller.StylerButton btnstart 
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Start"
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
   Begin PainKiller.StylerButton btncan 
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   2400
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
   Begin VB.OptionButton chkqick 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Perform Quick Scan"
      ForeColor       =   &H00784A32&
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   1800
      Width           =   2895
   End
   Begin VB.OptionButton chkfull 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Perform Full Scan"
      ForeColor       =   &H00784A32&
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Scanning Option :"
      ForeColor       =   &H00784A32&
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label lbltxt 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "PainKiller Drive Scanner lets you to scan throu your selected drives. Please select the drives you want to scan.."
      ForeColor       =   &H00784A32&
      Height          =   615
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmsdrv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btncan_Click()
frmsdrv.Visible = False
frmmain.Visible = True
Unload Me
End Sub

Private Sub btnStart_Click()
On Error Resume Next
Dim v As Integer
Dim i As Long
If frmsdrv.chkfull.Value = True Then
'txtshow.Picture = txtfull.Picture
    For v = 0 To lstdrv.ListCount
    i = v
        'MsgBox frmsdrv.lstdrv.List(i)
        If lstdrv.Checked(i) = True Then frmscan.lstdrv.AddItem frmsdrv.lstdrv.List(i)
        
    Next v
    frmscan.chkAll.Value = 1
    frmscan.chkVscan.Value = 1
        If frmscan.lstdrv.ListCount = 0 Then
            Unload frmscan
            MsgBox "No Drive(s) had been selected!", vbExclamation, "PainKiller"
            Exit Sub
        Else
            frmsdrv.Visible = False
            frmscan.Visible = True
        End If
End If

If chkqick.Value = True Then
'txtshow.Picture = txtfull.Picture
    For v = 0 To lstdrv.ListCount
    i = v
        If lstdrv.Checked(i) = True Then frmscan.lstdrv.AddItem frmsdrv.lstdrv.List(i)
    Next v
    frmscan.chkAll.Value = 1
    frmscan.chkVscan.Value = 0
    If frmscan.lstdrv.ListCount = 0 Then
            Unload frmscan
            MsgBox "No Drive(s) had been selected!", vbExclamation, "PainKiller"
            Exit Sub
        Else
            frmsdrv.Visible = False
            frmscan.Visible = True
        End If
End If

If chkfull.Value = False And chkqick.Value = False Then Exit Sub



Unload Me
End Sub
