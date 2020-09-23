VERSION 5.00
Begin VB.Form frmabt 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5370
   Icon            =   "frmabt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmabt.frx":628A
   ScaleHeight     =   3210
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   360
      X2              =   5040
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Painkiller © by Ratul Ahmed. All Right reserved 2008."
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   3000
      Width           =   4095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   6
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   5
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Ahmed Rizawan Shams Ratul"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   4
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Ahmed Rizawan Shams Ratul"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   3
      Top             =   1920
      Width           =   3735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designed by"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Coded by"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   690
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Painkiller is a strong worm remover which can search through Your computer and find various worms to disinfect them."
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   5055
   End
End
Attribute VB_Name = "frmabt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmmain.Visible = True
End Sub
