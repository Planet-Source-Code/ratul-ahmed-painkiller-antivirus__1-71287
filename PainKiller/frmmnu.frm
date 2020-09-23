VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmpopup 
   Caption         =   "frmmnu"
   ClientHeight    =   5430
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvwDrives 
      Height          =   3015
      Left            =   2520
      TabIndex        =   0
      Top             =   1920
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
   Begin VB.Menu Options 
      Caption         =   "Options"
      Begin VB.Menu mnu1 
         Caption         =   "# Full Scan"
         Index           =   0
      End
      Begin VB.Menu mnu2 
         Caption         =   "# Drive Scan"
         Index           =   1
      End
      Begin VB.Menu mnu3 
         Caption         =   "# Folder Scan"
         Index           =   2
      End
      Begin VB.Menu mnu4 
         Caption         =   "# Media Scan"
         Index           =   3
      End
      Begin VB.Menu mnu5 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnu6 
         Caption         =   "Exit"
         Index           =   5
      End
   End
   Begin VB.Menu Tools 
      Caption         =   "Tools"
      Begin VB.Menu mnua1 
         Caption         =   "Clean Autoruns"
         Index           =   6
      End
      Begin VB.Menu mnua2 
         Caption         =   "Repair System"
         Index           =   7
      End
      Begin VB.Menu mnua3 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnua4 
         Caption         =   "Backup Registry"
         Index           =   9
      End
      Begin VB.Menu mnua5 
         Caption         =   "Restore Registry"
         Index           =   10
      End
   End
End
Attribute VB_Name = "frmpopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Form_Load()
Set m_clsSubcls = New cSubclass
    
    m_clsSubcls.Subclass Me.hWnd, Me
    m_clsSubcls.AddMsg Me.hWnd, WM_DEVICECHANGE
    
    RefreshDriveList
End Sub

Private Sub mnu1_Click(Index As Integer)
On Error Resume Next
Dim v As Integer
'txtshow.Picture = txtfull.Picture
For v = 0 To lvwDrives.ListItems.count
    frmscan.lstdrv.AddItem lvwDrives.ListItems.Item(v).Text
    frmscan.chkVscan.Value = 1
    'MsgBox lvwDrives.ListItems.Item(v).Text
Next v
frmscan.Visible = True
frmmain.Visible = False
End Sub

Private Sub mnu2_Click(Index As Integer)
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

Private Sub mnu3_Click(Index As Integer)
frmfld.Visible = True
frmmain.Visible = False
End Sub

Private Sub mnu4_Click(Index As Integer)
On Error Resume Next
MsgBox "Please connect your device first and then press OK!", vbInformation, "PainKiller"
frmmain.Visible = False
frmdt.Visible = True
End Sub

Private Sub mnu6_Click(Index As Integer)
End
End Sub

Private Sub mnua1_Click(Index As Integer)
On Error Resume Next
Dim v As Integer
'txtshow.Picture = txtfull.Picture
For v = 0 To lvwDrives.ListItems.count
    frmscan.lstdrv.AddItem lvwDrives.ListItems.Item(v).Text
    frmscan.chkVscan.Value = 0
    'MsgBox lvwDrives.ListItems.Item(v).Text
Next v
frmscan.Visible = True
frmmain.Visible = False
End Sub

Private Sub mnua2_Click(Index As Integer)
frmmain.Visible = False
frmrep.Visible = True
End Sub

Private Sub mnua4_Click(Index As Integer)
frmregback.Visible = True
End Sub

Private Sub mnua5_Click(Index As Integer)
Dim a
a = MsgBox("Are you sure wanted to restore your registry settings?", vbYesNo, "PainKiller")
If a = vbYes Then
    frmrest.Visible = True
Else
    'NONE
End If
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

