VERSION 5.00
Begin VB.Form Frmmain 
   AutoRedraw      =   -1  'True
   Caption         =   "Styler Button 2007"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   6810
   Icon            =   "Frmmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   464
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   454
   StartUpPosition =   3  'Windows Default
   Begin StylerButton2007.StylerButton SB3 
      Height          =   1650
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   2910
      Caption         =   "Styler Button!"
      ForeColor       =   16777215
      CaptionDisableColor=   12236471
      CaptionEffectColor=   14789504
      CaptionEffect   =   4
      Theme           =   5
      FocusDottedRect =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAlignment=   8
   End
   Begin StylerButton2007.StylerButton SB2 
      Height          =   1650
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   2910
      Caption         =   "IS   "
      ForeColor       =   16777215
      CaptionDisableColor=   12236471
      CaptionEffectColor=   14789504
      CaptionEffect   =   4
      Theme           =   4
      FocusDottedRect =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAlignment=   8
   End
   Begin StylerButton2007.StylerButton SB1 
      Height          =   1650
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   2910
      Caption         =   "Time  "
      ForeColor       =   16777215
      CaptionDisableColor=   13153946
      CaptionEffectColor=   14789504
      CaptionEffect   =   4
      Theme           =   3
      FocusDottedRect =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAlignment=   8
   End
   Begin StylerButton2007.StylerButton SB 
      Height          =   1650
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   2910
      Caption         =   "It    "
      ForeColor       =   16777215
      CaptionDisableColor=   12236471
      CaptionEffectColor=   14789504
      CaptionEffect   =   4
      Theme           =   2
      FocusDottedRect =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAlignment=   8
   End
End
Attribute VB_Name = "Frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim a As String
a = App.Major & "." & App.Minor & "." & App.Revision
Me.Caption = "Styler Button 2007 vr." & a & " BY UMAIR 11D"
SB.IconSource = App.Path & "\1.png"
SB1.IconSource = App.Path & "\2.png"
SB2.IconSource = App.Path & "\3.png"
SB3.IconSource = App.Path & "\4.png"

SB.Icon = True
SB1.Icon = True
SB2.Icon = True
SB3.Icon = True

End Sub

Private Sub SB_Click()
SB.About
End Sub
Private Sub SB1_Click()
SB.About
End Sub
Private Sub SB2_Click()
SB.About
End Sub
Private Sub SB3_Click()
SB.About
End Sub
