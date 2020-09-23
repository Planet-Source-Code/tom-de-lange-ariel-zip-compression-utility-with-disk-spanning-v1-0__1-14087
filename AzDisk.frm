VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAzDisk 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Disk"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4260
      TabIndex        =   5
      ToolTipText     =   "Cancel selection and close window"
      Top             =   3660
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   5460
      TabIndex        =   3
      ToolTipText     =   "Accept selection and close window"
      Top             =   3660
      Width           =   1095
   End
   Begin VB.CommandButton cmdDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Details"
      Height          =   375
      Left            =   3060
      TabIndex        =   2
      ToolTipText     =   "Cancel selection and close window"
      Top             =   3660
      Width           =   1095
   End
   Begin VB.Timer tmr 
      Interval        =   2000
      Left            =   60
      Top             =   3660
   End
   Begin MSComctlLib.StatusBar StBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   4125
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1138
            MinWidth        =   1147
            Key             =   "Size"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2117
            MinWidth        =   2117
            Key             =   "Volume"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7726
            Key             =   "Info"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   609
            MinWidth        =   609
            Key             =   "Pic"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   2715
      Left            =   60
      TabIndex        =   4
      Top             =   840
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4789
      SortKey         =   -2
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   -9
         Key             =   "Name"
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         Key             =   "Modified"
         Text            =   "Modified"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         Key             =   "Size"
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image imgGrey 
      Height          =   240
      Left            =   900
      Picture         =   "AzDisk.frx":0000
      Top             =   3780
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgRed 
      Height          =   240
      Left            =   600
      Picture         =   "AzDisk.frx":038A
      Top             =   3780
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgGreen 
      Height          =   240
      Left            =   1200
      Picture         =   "AzDisk.frx":0714
      Top             =   3780
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "                                                                                                                 "
      Height          =   195
      Left            =   1155
      TabIndex        =   0
      Top             =   300
      Width           =   5115
   End
   Begin VB.Image imgDisk 
      Height          =   720
      Left            =   180
      Picture         =   "AzDisk.frx":0A9E
      Top             =   60
      Width           =   720
   End
End
Attribute VB_Name = "frmAzDisk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------
'Module     : frmAzDisk
'Description: Disk CHange Dialog
'Release    : 2001 VB6
'Copyright  : © T De Lange
'----------------------------------------------------------------
Option Base 0
Option Explicit
DefLng H-N
DefBool O

Const ModName = "Disk Change Dialog"
Public ShowDetails As Boolean
Private DriveReady As Boolean

'----------------------------------------------
'The following variables are
'set by the Zip_ChangeDisk() event
'If SourceFile is given, the client (this form)
'must locate the file (or cancel the operation)
'----------------------------------------------
Public Drive As Drive
Public SourceFile As String
Public Sub CheckDriveStatus(Optional ByVal Force As Boolean = False)
'------------------------------------
'Check Drive Status
'------------------------------------
If Drive.IsReady <> DriveReady Or Force Then
  DriveReady = Drive.IsReady
  If DriveReady Then
    StBar.Panels("Pic").Picture = imgGreen.Picture
    GetDriveStats
  Else
    StBar.Panels("Pic").Picture = imgRed.Picture
    ClearStats
  End If
  CheckFileStatus
  Me.Refresh
End If

End Sub
Public Sub CheckFileStatus()
'------------------------------------
'Check File Status
'------------------------------------
Dim fso As New FileSystemObject
Dim FileOk As Boolean

If SourceFile <> "" Then
  If DriveReady Then
    FileOk = fso.FileExists(SourceFile)
  Else
    FileOk = False
  End If
Else
  FileOk = DriveReady
  If FileOk Then
    StBar.Panels("Info") = "All existing files will be erased!"
  Else
    StBar.Panels("Info") = "Please insert disk..."
  End If
End If
cmdOk.Enabled = FileOk

End Sub

Sub ClearStats()
'--------------------------------------------
'Clear all stats
'--------------------------------------------
StBar.Panels("Size") = ""
StBar.Panels("Volume") = ""
StBar.Panels("Info") = "Please insert disk..."
lvw.ListItems.Clear

End Sub

Sub GetDriveStats()
'--------------------------------------------
'Get drive capacity, free space & root files
'--------------------------------------------
Dim DrvCap As Long, DrvFree As Long, DrvUsed As Long

DrvCap = Drive.TotalSize
DrvFree = Drive.AvailableSpace
DrvUsed = DrvCap - DrvFree

If DrvCap = 1457664 Then
  StBar.Panels("Size") = "1.44Mb"
ElseIf DrvCap = 730112 Then
  StBar.Panels("Size") = "720Kb"
Else
  StBar.Panels("Size") = Format(DrvCap / 1024, "##0") & "Kb"
End If

StBar.Panels("Volume") = Drive.VolumeName

StBar.Panels("Info") = Format(DrvFree / 1024, "#,##0") & "Kb free  " & _
      Format(DrvUsed / 1024, "#,##0") & "Kb used"
      
FillLvw

End Sub

Sub Resize()
If ShowDetails Then
  Me.Height = 4815
  lvw.Visible = True
  cmdDetails.Move 4260, 3660
  cmdOk.Move 5460, 3660
  cmdDetails.Caption = "Hide"
Else
  Me.Height = 2010
  lvw.Visible = False
  cmdDetails.Move 4260, 840
  cmdOk.Move 5460, 840
  cmdDetails.Caption = "Details"
End If
Me.Refresh

End Sub

Private Sub cmdCancel_Click()
DlgOk = False
Unload Me
End Sub

Private Sub cmdDetails_Click()
ShowDetails = Not (ShowDetails)
Resize
End Sub

Private Sub cmdOk_Click()
DlgOk = True
Unload Me
End Sub

Sub FillLvw()
'----------------------------------------------
'Scan rootfiles
'----------------------------------------------
Dim Item As ListItem, i, n, Ok
Dim File As Scripting.File

On Error GoTo FillLvwFileErr
Screen.MousePointer = vbArrowHourglass
lvw.Sorted = False
lvw.ListItems.Clear

For Each File In Drive.Rootfolder.Files
  Set Item = lvw.ListItems.Add()
  Item.Text = File.Name
  Item.SubItems(1) = RightFormat(File.DateLastModified, "yyyy/mm/dd hh:nn")
  Item.SubItems(2) = RightFormat(File.Size, "#,###,##0")
Next
lvw.SortKey = 0
lvw.Sorted = True
Screen.MousePointer = vbNormal
Exit Sub

FillLvwFileErr:
Screen.MousePointer = vbNormal
Exit Sub

End Sub

Private Sub Form_Load()
'---------------------------------------------------
'Load & Resize
'---------------------------------------------------
Dim i, Key As String, wdth

On Error Resume Next
For i = 1 To lvw.ColumnHeaders.Count
  Key = "ColumnHeader " & CStr(i)
  wdth = lvw.ColumnHeaders(i).Width
  lvw.ColumnHeaders(i).Width = GetProfile(ModName, Key, wdth)
Next
ShowDetails = GetProfile(ModName, "Show Details", False)
Resize

End Sub

Private Sub Form_Unload(Cancel As Integer)
'-----------------------------------------------------------
'Save settings
'-----------------------------------------------------------
Dim i, Key As String

For i = 1 To lvw.ColumnHeaders.Count
  Key = "ColumnHeader " & CStr(i)
  SaveProfile ModName, Key, lvw.ColumnHeaders(i).Width
Next
SaveProfile ModName, "Show Details", ShowDetails

End Sub


Private Sub tmr_Timer()
CheckDriveStatus
'Me.Refresh
DoEvents

End Sub


