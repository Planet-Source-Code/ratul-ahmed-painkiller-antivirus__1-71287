VERSION 5.00
Begin VB.Form frmx 
   BorderStyle     =   0  'None
   Caption         =   "Exiting.."
   ClientHeight    =   210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   ScaleHeight     =   210
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   600
      Top             =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait..."
      Height          =   375
      Left            =   -720
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Dim v As Integer
For v = 0 To 5
Unload frmscan
Unload frmmscan
Next v
Timer1.Enabled = False
frmmain.Visible = True
Unload Me
Unload frmx
End Sub
