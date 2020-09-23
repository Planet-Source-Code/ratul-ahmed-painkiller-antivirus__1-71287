VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "PainKiller"
   ClientHeight    =   10005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10005
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   3135
      Left            =   5400
      ScaleHeight     =   3075
      ScaleWidth      =   6435
      TabIndex        =   16
      Top             =   6720
      Width           =   6495
      Begin MSComctlLib.ListView lvwDrives 
         Height          =   3015
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   6390
         _ExtentX        =   11271
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Drive"
            Object.Width           =   1041
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Bus Type"
            Object.Width           =   1481
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Removable"
            Object.Width           =   1834
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Name"
            Object.Width           =   5186
         EndProperty
      End
   End
   Begin VB.PictureBox txtmedia 
      Height          =   375
      Left            =   1560
      Picture         =   "frmmain.frx":74F2
      ScaleHeight     =   315
      ScaleWidth      =   195
      TabIndex        =   9
      Top             =   7200
      Width           =   255
   End
   Begin VB.PictureBox txtfldr 
      Height          =   375
      Left            =   1320
      Picture         =   "frmmain.frx":A5A8
      ScaleHeight     =   315
      ScaleWidth      =   195
      TabIndex        =   8
      Top             =   7200
      Width           =   255
   End
   Begin VB.PictureBox txtdrv 
      Height          =   375
      Left            =   1080
      Picture         =   "frmmain.frx":F773
      ScaleHeight     =   315
      ScaleWidth      =   195
      TabIndex        =   7
      Top             =   7200
      Width           =   255
   End
   Begin VB.PictureBox txtfull 
      Height          =   375
      Left            =   840
      Picture         =   "frmmain.frx":12751
      ScaleHeight     =   315
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   7200
      Width           =   255
   End
   Begin VB.PictureBox txtnor 
      Height          =   375
      Left            =   600
      Picture         =   "frmmain.frx":16A31
      ScaleHeight     =   315
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   7200
      Width           =   255
   End
   Begin VB.PictureBox imgMask 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   0
      Picture         =   "frmmain.frx":1C120
      ScaleHeight     =   6135
      ScaleWidth      =   10935
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin PainKiller.GUI_Rollover btnMedia 
         Height          =   600
         Left            =   1680
         TabIndex        =   4
         Top             =   4320
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   1058
         Selectable      =   0   'False
         ImageNormal     =   "frmmain.frx":FFE64
         ImageHover      =   "frmmain.frx":100A76
         ImageDown       =   "frmmain.frx":10195C
         ImageDisabled   =   "frmmain.frx":10256E
         ImageMask       =   "frmmain.frx":103180
         ImageSelected   =   "frmmain.frx":103D92
         ImageSelectedHover=   "frmmain.frx":1049A4
      End
      Begin PainKiller.GUI_Rollover btnFldr 
         Height          =   600
         Left            =   1680
         TabIndex        =   3
         Top             =   3720
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   1058
         Selectable      =   0   'False
         ImageNormal     =   "frmmain.frx":1055B6
         ImageHover      =   "frmmain.frx":1061AD
         ImageDown       =   "frmmain.frx":10707D
         ImageDisabled   =   "frmmain.frx":107C74
         ImageMask       =   "frmmain.frx":10886B
         ImageSelected   =   "frmmain.frx":109462
         ImageSelectedHover=   "frmmain.frx":10A059
      End
      Begin PainKiller.GUI_Rollover btnDrv 
         Height          =   600
         Left            =   1680
         TabIndex        =   2
         Top             =   3120
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   1058
         Selectable      =   0   'False
         ImageNormal     =   "frmmain.frx":10AC50
         ImageHover      =   "frmmain.frx":10B7C2
         ImageDown       =   "frmmain.frx":10C5EC
         ImageDisabled   =   "frmmain.frx":10D15E
         ImageMask       =   "frmmain.frx":10DCD0
         ImageSelected   =   "frmmain.frx":10E842
         ImageSelectedHover=   "frmmain.frx":10F3B4
      End
      Begin PainKiller.GUI_Rollover btnFull 
         Height          =   600
         Left            =   1680
         TabIndex        =   1
         Top             =   2520
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   1058
         Selectable      =   0   'False
         ImageNormal     =   "frmmain.frx":10FF26
         ImageHover      =   "frmmain.frx":110A5E
         ImageDown       =   "frmmain.frx":111851
         ImageDisabled   =   "frmmain.frx":112389
         ImageMask       =   "frmmain.frx":112EC1
         ImageSelected   =   "frmmain.frx":1139F9
         ImageSelectedHover=   "frmmain.frx":114531
      End
      Begin PainKiller.GUI_Rollover btnMin 
         Height          =   270
         Left            =   8400
         TabIndex        =   14
         Top             =   760
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   476
         Enabled         =   0   'False
         Selectable      =   0   'False
         ImageNormal     =   "frmmain.frx":115069
         ImageHover      =   "frmmain.frx":117380
         ImageDown       =   "frmmain.frx":119751
         ImageDisabled   =   "frmmain.frx":11BB22
         ImageMask       =   "frmmain.frx":11DE39
         ImageSelected   =   "frmmain.frx":120150
         ImageSelectedHover=   "frmmain.frx":122467
      End
      Begin PainKiller.GUI_Rollover btnExit 
         Height          =   270
         Left            =   8790
         TabIndex        =   15
         Top             =   765
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   476
         Selectable      =   0   'False
         ImageNormal     =   "frmmain.frx":12477E
         ImageHover      =   "frmmain.frx":126DBD
         ImageDown       =   "frmmain.frx":12944F
         ImageDisabled   =   "frmmain.frx":12BAE1
         ImageMask       =   "frmmain.frx":12E120
         ImageSelected   =   "frmmain.frx":13075F
         ImageSelectedHover=   "frmmain.frx":132D9E
      End
      Begin VB.Label lblabt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "About"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   8760
         TabIndex        =   13
         Top             =   1395
         Width           =   420
      End
      Begin VB.Label lblhelp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Help"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   7920
         TabIndex        =   12
         Top             =   1395
         Width           =   330
      End
      Begin VB.Label lbltools 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tools"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   6960
         TabIndex        =   11
         Top             =   1395
         Width           =   390
      End
      Begin VB.Label lbloptions 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Options"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   5880
         TabIndex        =   10
         Top             =   1395
         Width           =   540
      End
      Begin VB.Image txtshow 
         Height          =   2415
         Left            =   5160
         Top             =   2520
         Width           =   3975
      End
      Begin VB.Image BodyIMG 
         Appearance      =   0  'Flat
         Height          =   6480
         Left            =   0
         Picture         =   "frmmain.frx":1353DD
         Top             =   0
         Width           =   10800
      End
   End
   Begin VB.Label ttext 
      Caption         =   "Label1"
      Height          =   255
      Left            =   840
      TabIndex        =   18
      Top             =   6480
      Width           =   4935
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'////////////////////////////////////////////////////////////////////////////////////////////////
'/////Coder: Ratul Ahmed/////////////////////////////////////////////////////////////////////////
'/////Thanx 2 so many peoples////////////////////////////////////////////////////////////////////
'/////I Love my sweet Bangladesh/////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////
'/////Read and Learn/////////////////////////////////////////////////////////////////////////////

Option Explicit
Implements iSubclass
Private Declare Function GetLogicalDriveStrings Lib "Kernel32" _
Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private m_clsSubcls As cSubclass
Dim FormLoad As Boolean
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Sub BodyIMG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long

txtshow.Picture = txtnor.Picture

lblabt.ForeColor = &H808000
lblhelp.ForeColor = &H808000
lbloptions.ForeColor = &H808000
lbltools.ForeColor = &H808000

If Button = 1 Then
    Call ReleaseCapture
    lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub



Private Sub btnDrv_OnMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtshow.Picture = txtdrv.Picture
End Sub

Private Sub btnFldr_OnMouseClick()
frmfld.Visible = True
frmmain.Visible = False
End Sub

Private Sub btnFldr_OnMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtshow.Picture = txtfldr.Picture
End Sub

Private Sub btnFull_OnMouseClick()
On Error Resume Next
Unload frmscan
Dim v As Integer
Load frmscan
'txtshow.Picture = txtfull.Picture
For v = 0 To lvwDrives.ListItems.count
    frmscan.lstdrv.AddItem lvwDrives.ListItems.Item(v).Text
    frmscan.chkVscan.Value = 1
    'MsgBox lvwDrives.ListItems.Item(v).Text
Next v
'Load frmscan
frmscan.Visible = True
frmmain.Visible = False
End Sub
Private Sub btnDrv_OnMouseClick()
On Error Resume Next
'Dim v As Integer
'txtshow.Picture = txtfull.Picture
'For v = 0 To lvwDrives.ListItems.count
'    'frmsdrv.lstdrv.ListItems.add.Text lvwDrives.ListItems.Item(v).Text
'    With frmsdrv.lstdrv.ListItems.add         '--
'        .Text = lvwDrives.ListItems.Item(v).Text                '--
'    End With
'    'MsgBox
'Next v
frmsdrv.Visible = True
frmmain.Visible = False
End Sub
Private Sub btnFull_OnMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

txtshow.Picture = txtfull.Picture

End Sub

Private Sub btnMedia_OnMouseClick()
On Error Resume Next
MsgBox "Please connect your device first and then press OK!", vbInformation, "PainKiller"
frmmain.Visible = False
frmdt.Visible = True
End Sub

Private Sub btnMedia_OnMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtshow.Picture = txtmedia.Picture
End Sub

Private Sub Form_Load()
Dim WindowRegion As Long
'///////////////////////////////////////[ Interface Codes Here ]////
FormLoad = True                                                 '///
imgMask.ScaleMode = vbPixels                                    '///
imgMask.AutoRedraw = True                                       '///
imgMask.AutoSize = True                                         '///
imgMask.BorderStyle = vbBSNone                                  '///
frmmain.BorderStyle = vbBSNone                                  '///
frmmain.Width = imgMask.Width                                   '///
frmmain.Height = imgMask.Height                                 '///
WindowRegion = MakeRegion(imgMask)                              '///
SetWindowRgn frmmain.hWnd, WindowRegion, True                   '///
'///////////////////////////////////////////////////////////////////
txtshow.Picture = txtnor.Picture


  Set m_clsSubcls = New cSubclass
    
    m_clsSubcls.Subclass Me.hWnd, Me
    m_clsSubcls.AddMsg Me.hWnd, WM_DEVICECHANGE
    
    RefreshDriveList

End Sub



Private Sub lblabt_Click()
frmmain.Visible = False
frmabt.Visible = True
End Sub

Private Sub lblabt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblabt.ForeColor = &H4080&
End Sub

Private Sub lblhelp_Click()
On Error Resume Next
Dim fchk As String
Dim yesauto As String
MsgBox "sorry couldn't have time to make one."
'fchk = App.path & "\data\help.html"
'ttext = fchk
'yesauto = FileExists(ttext)

'If yesauto = True Then
'    Shella fchk
'Else
'    MsgBox "The help file " & fchk & " could not be located on the path!!", vbExclamation, "PainKiller"
'End If
End Sub

Private Sub lblhelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblhelp.ForeColor = &H4080&
End Sub

Private Sub lbloptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbloptions.ForeColor = &H4080&
End Sub

Private Sub lbltools_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbltools.ForeColor = &H4080&
End Sub

Private Sub lbltools_Click()
PopupMenu frmpopup.Tools
End Sub

Private Sub lbloptions_Click()
PopupMenu frmpopup.Options
End Sub

Private Sub btnExit_OnMouseClick()
End
End Sub


Private Sub RefreshDriveList()
    Dim strDriveBuffer  As String
    Dim strDrives()     As String
    Dim i               As Long
    Dim udtInfo         As DEVICE_INFORMATION
    
    strDriveBuffer = Space(240)
    strDriveBuffer = Left$(strDriveBuffer, GetLogicalDriveStrings(Len(strDriveBuffer), strDriveBuffer))
    strDrives = Split(strDriveBuffer, Chr$(0))

    lvwDrives.ListItems.Clear

    For i = 0 To UBound(strDrives) - 1
        With lvwDrives.ListItems.add(Text:=strDrives(i))
            udtInfo = GetDevInfo(strDrives(i))
            
            If udtInfo.Valid Then
                Select Case udtInfo.BusType
                    Case BusTypeUsb:        .SubItems(1) = "USB"
                    Case BusType1394:       .SubItems(1) = "1394"
                    Case BusTypeAta:        .SubItems(1) = "ATA"
                    Case BusTypeAtapi:      .SubItems(1) = "ATAPI"
                    Case BusTypeFibre:      .SubItems(1) = "Fibre"
                    Case BusTypeRAID:       .SubItems(1) = "RAID"
                    Case BusTypeScsi:       .SubItems(1) = "SCSI"
                    Case BusTypeSsa:        .SubItems(1) = "SSA"
                    Case BusTypeUnknown:    .SubItems(1) = "Unknown"
                End Select
                
                .SubItems(2) = IIf(udtInfo.Removable, "yes", "no")
                .SubItems(3) = Trim$(udtInfo.VendorID & " " & udtInfo.ProductID & " " & udtInfo.ProductRevision)
                
                .Tag = strDrives(i)
            End If
        End With
    Next
End Sub

Private Sub iSubclass_WndProc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As eMsg, ByVal wParam As Long, ByVal lParam As Long, lParamUser As Long)
    If uMsg = WM_DEVICECHANGE Then RefreshDriveList
End Sub
