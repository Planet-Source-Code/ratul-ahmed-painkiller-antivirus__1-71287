VERSION 5.00
Begin VB.Form frmfld 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Folder"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6015
   Icon            =   "frmfld.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin PainKiller.StylerButton btncan 
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   3240
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
   Begin PainKiller.StylerButton btnStart 
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
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
   Begin PainKiller.StylerButton btnAdd 
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Add"
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
   Begin VB.ListBox lstfldr 
      Height          =   2400
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   5775
   End
   Begin VB.Label lbltxt 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "PainKiller Folder Scanner lets you to scan throu your selected folder. Please select the folders you want to scan.."
      ForeColor       =   &H00784A32&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmfld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sfol As String

Private Sub btnAdd_Click()
sfol = BrowseForFolder("", "Select Filder")
If sfol <> "" Then lstfldr.AddItem sfol
End Sub

Private Sub btncan_Click()
frmfld.Visible = False
frmmain.Visible = True
Unload Me
End Sub

Private Sub btnStart_Click()
On Error Resume Next
Dim v As Integer

If lstfldr.ListCount = 0 Then
    MsgBox "No Folder or directory had been selected!", vbExclamation, "PainKiller"
    Exit Sub
Else
'txtshow.Picture = txtfull.Picture
    For v = 0 To lstfldr.ListCount
        frmscan.lstdrv.AddItem frmfld.lstfldr.List(v)
        'MsgBox frmfld.lstfldr.List(v)
        'MsgBox lvwDrives.ListItems.Item(v).Text
    Next v
    frmscan.chkAll.Value = 0
    frmscan.chkVscan.Value = 1
    frmfld.Visible = False
    frmscan.Visible = True
    Unload Me
End If
End Sub

