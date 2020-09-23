VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAzMain 
   Caption         =   "Ariel Zip"
   ClientHeight    =   6165
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AzMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picStatus 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8340
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   6
      Top             =   5895
      Width           =   255
   End
   Begin MSComDlg.CommonDialog CdlOpen 
      Left            =   1920
      Top             =   5340
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Select files to add to zip archive"
      Filter          =   "All Files (*.*)|*.*|"
      MaxFileSize     =   20000
   End
   Begin VB.PictureBox picIconDefault 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1380
      Picture         =   "AzMain.frx":08CA
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   5280
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1080
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   5280
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picProgress 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   2985
      TabIndex        =   1
      Top             =   5880
      Width           =   2985
   End
   Begin MSComctlLib.StatusBar StBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   5850
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Key             =   "Progress"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            Key             =   "Percent"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8352
            MinWidth        =   176
            Key             =   "Info"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   635
            MinWidth        =   635
            Picture         =   "AzMain.frx":0C0C
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
   Begin MSComctlLib.ImageList imgToolbar 
      Left            =   6660
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AzMain.frx":11B0
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AzMain.frx":2004
            Key             =   "Folder Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AzMain.frx":2E58
            Key             =   "Item New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AzMain.frx":3CAC
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AzMain.frx":4248
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AzMain.frx":47E4
            Key             =   "Edit Prices"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AzMain.frx":5638
            Key             =   "Book Red"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AzMain.frx":5BD4
            Key             =   "Book Blue"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AzMain.frx":6170
            Key             =   "Book Cyan"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AzMain.frx":670C
            Key             =   "Book Brown"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AzMain.frx":6CA8
            Key             =   "Book Purple"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AzMain.frx":7244
            Key             =   "Import"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AzMain.frx":77E0
            Key             =   "Permanent"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AzMain.frx":8634
            Key             =   "Folder New"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AzMain.frx":9488
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AzMain.frx":9A24
            Key             =   "Extract"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AzMain.frx":9FC0
            Key             =   "Ariel1"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AzMain.frx":A89C
            Key             =   "Ariel"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AzMain.frx":B6F0
            Key             =   "FileAdd"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AzMain.frx":BC8C
            Key             =   "Folder Add"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AzMain.frx":CAE0
            Key             =   "Refresh"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolbar 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   1005
      ButtonWidth     =   1561
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "New"
            Object.ToolTipText     =   "Create new Ariel zip file"
            ImageKey        =   "Folder New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Key             =   "Open"
            Object.ToolTipText     =   "Open an existing Ariel zip file"
            ImageKey        =   "Folder Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add Folder"
            Key             =   "Add Folder"
            Object.ToolTipText     =   "Add a folder to an Ariel zip file"
            ImageKey        =   "Folder Add"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add Files"
            Key             =   "Add Files"
            Object.ToolTipText     =   "Add files to the archive"
            ImageKey        =   "FileAdd"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh existing files in archive"
            ImageKey        =   "Refresh"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete files from the archive list"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Zip"
            Key             =   "Zip"
            Object.ToolTipText     =   "Save the Ariel zip file"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Unzip"
            Key             =   "Unzip"
            Object.ToolTipText     =   "Unzip the Ariel zip file"
            ImageKey        =   "Extract"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   4605
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   8123
      SortKey         =   -1
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   -8
         Key             =   "Name"
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   "Size"
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   "Zipped"
         Text            =   "Zipped"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   "Ratio"
         Text            =   "Ratio"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Key             =   "Modified"
         Text            =   "Modified"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   -4
         Key             =   "Path"
         Text            =   "Path"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   480
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin VB.Image imgBlue 
      Height          =   240
      Left            =   8100
      Picture         =   "AzMain.frx":D07C
      Top             =   5280
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgGreen 
      Height          =   240
      Left            =   7740
      Picture         =   "AzMain.frx":D606
      Top             =   5280
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgRed 
      Height          =   240
      Left            =   8460
      Picture         =   "AzMain.frx":DB90
      Top             =   5280
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgGrey 
      Height          =   240
      Left            =   7440
      Picture         =   "AzMain.frx":E11A
      Top             =   5280
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New..."
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFileZip 
         Caption         =   "&Zip..."
      End
      Begin VB.Menu mnuFileUnzip 
         Caption         =   "&Unzip..."
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Actions"
      Begin VB.Menu mnuActionAddFolder 
         Caption         =   "&Add Folder..."
      End
      Begin VB.Menu mnuActionAddFiles 
         Caption         =   "Add &Files..."
      End
      Begin VB.Menu mnuActionRefresh 
         Caption         =   "&Refresh Files"
      End
      Begin VB.Menu mnuActionSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActionDelete 
         Caption         =   "&Delete"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewText 
         Caption         =   "Toolbar Te&xt"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpReg 
         Caption         =   "&Register Ariel Files"
      End
      Begin VB.Menu mnuHelpUnreg 
         Caption         =   "&Unregister Ariel Files"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmAzMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------
'Module     : frmArZip
'Description: Ariel Zip App Main Window
'Release    : 2001 VB6
'Copyright  : © T De Lange
'----------------------------------------------------------------
Option Base 1
Option Explicit
DefLng H-N
DefBool O

Const ModName = "Ariel Zip Main"

Dim ReadyZip As Boolean   'Ready to zip
Dim ReadyUnzip As Boolean 'Ready to unzip
Dim Busy As Boolean       'Busy unzipping/zipping
Public WithEvents Zip As ArZip
Attribute Zip.VB_VarHelpID = -1



Sub AddFiles()
'---------------------------------------------------------
'Add files to file list to include in archive
'---------------------------------------------------------
Dim Ok, n
Dim FileList() As String

Ok = SelectFiles(CdlOpen, FileList())
If Ok Then
  n = Zip.AddFiles(FileList())
  FillLvw
  UpdateInfoPanel False
End If

End Sub

Sub AddFolder()
'---------------------------------------------------------
'Add all files in a folder to the archive list
'---------------------------------------------------------

frmAzAddFldr.Show vbModal, Me
If DlgOk Then
  Zip.AddFolder rg.AddFolder, rg.IncludeSubfolders
  FillLvw
  UpdateInfoPanel False
End If

End Sub

Sub NewArchive()
'----------------------------------------------
'Create a new archive
'----------------------------------------------
frmAzNew.Show vbModal, Me
If DlgOk Then
  Zip.NewArchive rg.ZipFile, rg.ZipFolder, rg.IncludeFiles, rg.IncludeSubfolders
  FillLvw
  UpdateInfoPanel False
End If

End Sub

Sub OpenArchive()
'----------------------------------------------
'Open an existing archive
'----------------------------------------------
frmAzOpen.Show vbModal, Me
If DlgOk Then
  Zip.OpenArchive rg.ZipFile
  FillLvw
  UpdateInfoPanel False
End If

End Sub

Sub RefreshFiles()
'---------------------------------------------------------
'Refresh files in the file list. If a file has been
'modified since it was last zipped, re-zip it. Also
'update the original size, date/time & associated icon
'---------------------------------------------------------
Zip.RefreshFiles
FillLvw
UpdateInfoPanel False

End Sub

Sub Resize()
'--------------------------------------
'Resize form
'--------------------------------------
Dim Wd, Hg
Dim Ok

Wd = Me.Width
Hg = Me.Height
If Me.WindowState <> vbMinimized Then
  Ok = True
  If Wd < 6000 Then
    Me.Width = 6000
    Ok = False
  End If
  If Hg < 3500 Then
    Me.Height = 3500
    Ok = False
  End If
  If Ok Then
    Wd = Me.ScaleWidth
    Hg = Me.ScaleHeight
    'lvw
    If mnuViewToolbar.Checked Then
      lvw.Move 60, tbToolbar.Height, Wd - 120, Hg - tbToolbar.Height - StBar.Height - 30
    Else
      lvw.Move 60, 0, Wd - 120, Hg - StBar.Height - 30
    End If
    'picProgress
    picProgress.Move StBar.Panels("Progress").Left, Hg - StBar.Height + 45
    'picStatus
    picStatus.Move Wd - picStatus.Width - 300, Hg - StBar.Height + 45
  End If
End If

End Sub

Function SelectFiles(CD As CommonDialog, FileNames() As String) As Boolean
'--------------------------------------------------------------------
'Select files to add to archive using common dialog box
'Return Ok = True, Cancel = False
'--------------------------------------------------------------------
On Error GoTo SelectFilesErr
    
CD.Filter = "All Files (*.*)|*.*|"
CD.FilterIndex = 1
CD.flags = cdlOFNAllowMultiselect Or cdlOFNFileMustExist Or cdlOFNExplorer _
        Or cdlOFNNoDereferenceLinks
CD.DialogTitle = "Select one or more files to add to archive"
CD.MaxFileSize = 20000
CD.FileName = ""
CD.CancelError = True   'Exit if user presses Cancel
CD.ShowOpen
'Get filenames and return as array of strings
'Element 0 contains the path, the rest contain only the filename(s)
FileNames() = Split(CD.FileName, vbNullChar)
SelectFiles = True
Exit Function

SelectFilesErr:
SelectFiles = False

End Function


Sub FillLvw()
'----------------------------------------------
'Scan folders & files & fill listview
'----------------------------------------------
Dim Item As ListItem, i, n, Ok

On Error GoTo FillLvwFileErr
Screen.MousePointer = vbArrowHourglass
lvw.Sorted = False
lvw.ListItems.Clear
n = Zip.NoFiles
For i = 1 To n
  Set Item = lvw.ListItems.Add()
  Item.Key = Zip.Key(i)
  Item.SmallIcon = Zip.IconKey(i)
  Item.Text = Zip.Name(i)
  Item.Tag = CStr(i)
  With lvw
    Item.SubItems(1) = RightFormat(Zip.Size(i), "###,###,##0")
    Item.SubItems(2) = RightFormat(Zip.ZipSize(i), "###,###,##0")
    Item.SubItems(3) = RightFormat(Zip.Ratio(i), "##0.00%")
    Item.SubItems(4) = Format(Zip.Modified(i), "yyyy/mm/dd hh:nn")
    Item.SubItems(5) = Zip.RelativePath(i)
  End With
Next
lvw.Sorted = True
Screen.MousePointer = vbNormal
Exit Sub

FillLvwFileErr:
Ok = ReportErrorCont("FillLvwFiles()", ModName, Err, Error)
If Ok Then
  Resume Next
Else
  Exit Sub
End If

End Sub

Private Sub Progress(Value As Single, Info As String, Optional Show As Boolean = True)
'-------------------------------------------
'Update progress bar
'Value = 0 to 1
'-------------------------------------------
Dim x
Static OldValue As Single
Static OldShow As Boolean
Static OldInfo As String

If Value = 0 Then
  picProgress.Cls
  StBar.Panels("Percent").Text = ""
ElseIf Value <> OldValue Then
  'picProgress.ForeColor = RGB(0, 128, 128)
  x = picProgress.Width * Value
  picProgress.Line (0, 0)-(x, picProgress.Height), , BF
  picProgress.Refresh
  StBar.Panels("Percent").Text = Format(Value, "##0%")
  OldValue = Value
End If
If Info <> OldInfo Then
  StBar.Panels("Info").Text = Info
  OldInfo = Info
End If
If Show <> OldShow Then
  picProgress.Visible = Show
  If Not (Show) Then
    StBar.Panels("Percent").Text = ""
  End If
  OldShow = Show
End If

End Sub

Sub UnzipFile()
'---------------------------------------------------------
'Unzip archive to selected folder
'---------------------------------------------------------
frmAzUnzip.Show vbModal, Me
If DlgOk Then
  Zip.UnzipFiles rg.UnzipFolder, rg.UnzipAll, rg.Overwrite
  UpdateInfoPanel True
End If

End Sub

Sub UpdateInfoPanel(IncludeTime As Boolean)
'----------------------------------------------------------------
'Update the info panel in the status bar
'IncludeTime: Add the timeelapsed property
'----------------------------------------------------------------
Select Case Zip.Status
Case azsReady, azsCreated
  If IncludeTime Then
    StBar.Panels("Info") = Format(Zip.NoFiles, "#,##0") & " file(s)  " & Format(Zip.TotalSize / 1024, "#,##0") & " Kb  " & Format(Zip.TotalZipSize / 1024, "#,##0") & " Kb zipped  Ratio " & Format(Zip.TotalZipRatio, "##0.00%") & " in " & Format(Zip.ElapsedTime, "#,##0.0") & " sec"
  Else
    StBar.Panels("Info") = Format(Zip.NoFiles, "#,##0") & " file(s)  " & Format(Zip.TotalSize / 1024, "#,##0") & " Kb  " & Format(Zip.TotalZipSize / 1024, "#,##0") & " Kb zipped  Ratio " & Format(Zip.TotalZipRatio, "##0.00%")
  End If
  Me.Caption = "Ariel Zip - " & rg.ZipFile
Case Else
  Me.Caption = "Ariel Zip Program"
  StBar.Panels("Info") = ""
End Select

End Sub

Sub UpdateStatus(Optional Force As Boolean = False)
'------------------------------------
'Check & Update status
'------------------------------------
Static OldStatus As ArZipStatus

If Zip.Status <> OldStatus Or Force Then
  OldStatus = Zip.Status
  Select Case OldStatus
  Case azsEmpty
    picStatus.Picture = imgGrey.Picture
    StBar.Panels("Percent").Text = ""
    'Update toolbar status
    tbToolbar.Buttons("Add Files").Enabled = False
    tbToolbar.Buttons("Add Folder").Enabled = False
    tbToolbar.Buttons("Refresh").Enabled = False
    tbToolbar.Buttons("Delete").Enabled = False
    tbToolbar.Buttons("Zip").Enabled = False
    tbToolbar.Buttons("Unzip").Enabled = False
    'Update Menu status
    mnuActionAddFiles.Enabled = False
    mnuActionAddFolder.Enabled = False
    mnuActionRefresh.Enabled = False
    mnuActionDelete.Enabled = False
    mnuFileZip.Enabled = False
    mnuFileUnzip.Enabled = False
  Case azsReady     'Ready to zip archive
    picStatus.Picture = imgGreen.Picture
    StBar.Panels("Percent").Text = ""
    tbToolbar.Buttons("Add Files").Enabled = True
    tbToolbar.Buttons("Add Folder").Enabled = True
    tbToolbar.Buttons("Refresh").Enabled = True
    tbToolbar.Buttons("Delete").Enabled = True
    tbToolbar.Buttons("Zip").Enabled = True
    tbToolbar.Buttons("Unzip").Enabled = False
    'Update Menu status
    mnuActionAddFiles.Enabled = True
    mnuActionAddFolder.Enabled = True
    mnuActionRefresh.Enabled = True
    mnuActionDelete.Enabled = True
    mnuFileZip.Enabled = True
    mnuFileUnzip.Enabled = False
  Case azsCreated     'Ready to unzip archive
    picStatus.Picture = imgBlue.Picture
    StBar.Panels("Percent").Text = ""
    tbToolbar.Buttons("Add Files").Enabled = True
    tbToolbar.Buttons("Add Folder").Enabled = True
    tbToolbar.Buttons("Refresh").Enabled = True
    tbToolbar.Buttons("Delete").Enabled = True
    tbToolbar.Buttons("Zip").Enabled = True
    tbToolbar.Buttons("Unzip").Enabled = True
    'Update Menu status
    mnuActionAddFiles.Enabled = True
    mnuActionAddFolder.Enabled = True
    mnuActionRefresh.Enabled = True
    mnuActionDelete.Enabled = True
    mnuFileZip.Enabled = True
    mnuFileUnzip.Enabled = True
  Case azsBusy
    picStatus.Picture = imgRed.Picture
  End Select
  StBar.Refresh
  picProgress.Visible = (OldStatus = azsBusy)
End If

End Sub

Sub ZipFile()
'----------------------------------------------
'Zip all files in list to the selected archive
'----------------------------------------------
frmAzZip.Show vbModal, Me
If DlgOk Then
  Zip.ZipFiles rg.ZipFile, rg.CompressLevel, rg.Spanning, rg.SpanOption
  FillLvw
  UpdateInfoPanel True
End If

End Sub

Private Sub Form_Load()
'---------------------------------------------------
'Load & Resize
'---------------------------------------------------
Dim l, t, w, h
Dim i, Key As String, wdth

On Error Resume Next
If Me.WindowState = vbNormal Then
  l = GetProfile(ModName, "Left", Me.Left)
  t = GetProfile(ModName, "Top", Me.Top)
  w = GetProfile(ModName, "Width", Me.Width)
  h = GetProfile(ModName, "Height", Me.Height)
  Me.Move l, t, w, h
End If
For i = 1 To lvw.ColumnHeaders.Count
  Key = "ColumnHeader " & CStr(i)
  wdth = lvw.ColumnHeaders(i).Width
  lvw.ColumnHeaders(i).Width = GetProfile(ModName, Key, wdth)
Next
lvw.SortKey = GetProfile(ModName, "Sortkey", 0)
lvw.SortOrder = GetProfile(ModName, "Sortorder", lvwAscending)
'Set Sorted to True to sort the list.
lvw.Sorted = True


mnuViewToolbar.Checked = Not (GetProfile(ModName, "Toolbar", 0))
mnuViewToolbar_Click
mnuViewText.Checked = Not (GetProfile(ModName, "Text", 0))
mnuViewText_Click
Set Zip = New ArZip
Zip.Initialise lvw, imlIcons, picIcon, picIconDefault
UpdateStatus True

End Sub

Private Sub Form_Resize()
'--------------------------------------
'Resize form
'--------------------------------------
Resize

End Sub

Function FileExist(FileName As String) As Boolean
'----------------------------------------------
'Check if file exists
'----------------------------------------------
On Error GoTo FileExistErr
Call FileLen(FileName)
FileExist = True
Exit Function
  
FileExistErr:
FileExist = False
  
End Function


Private Sub Form_Unload(Cancel As Integer)
'-----------------------------------------------------------
'Save settings
'-----------------------------------------------------------
Dim i, Key As String

If Me.WindowState = vbNormal Then
  SaveProfile ModName, "Left", Me.Left
  SaveProfile ModName, "Top", Me.Top
  SaveProfile ModName, "Width", Me.Width
  SaveProfile ModName, "Height", Me.Height
End If
For i = 1 To lvw.ColumnHeaders.Count
  Key = "ColumnHeader " & CStr(i)
  SaveProfile ModName, Key, lvw.ColumnHeaders(i).Width
Next
SaveProfile ModName, "Sortkey", lvw.SortKey
SaveProfile ModName, "Sortorder", lvw.SortOrder
SaveProfile ModName, "Toolbar", IIf(mnuViewToolbar.Checked, -1, 0)
SaveProfile ModName, "Text", IIf(mnuViewText.Checked, -1, 0)
Set Zip = Nothing
PrgExit


End Sub

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'---------------------------------------------------------------
'Sort on Columnheader
'---------------------------------------------------------------
Screen.MousePointer = vbArrowHourglass
DoEvents
If ColumnHeader.Index - 1 = lvw.SortKey Then
  If lvw.SortOrder = lvwAscending Then
    lvw.SortOrder = lvwDescending
  Else
    lvw.SortOrder = lvwAscending
  End If
Else
  lvw.SortKey = ColumnHeader.Index - 1
End If
'Set Sorted to True to sort the list.
lvw.Sorted = True
Screen.MousePointer = vbNormal
DoEvents

End Sub

Private Sub lvw_KeyDown(KeyCode As Integer, Shift As Integer)
'------------------------------------------------------------
'Handle keys
'------------------------------------------------------------
Select Case KeyCode
Case vbKeyF2
  'Edit
  'lvw.StartLabelEdit
Case vbKeyF5
  'RefreshLvw
Case vbKeyReturn
  'EditProperties
Case vbKeyDelete
  DeleteItems
End Select

End Sub

Public Sub DeleteItems()
'------------------------------------------------
'Delete selected listview items
'------------------------------------------------
Dim i, Index, Ok
Dim Item As ListItem

On Error GoTo CmdDelErr
Screen.MousePointer = vbArrowHourglass
lvw.Sorted = False
For i = lvw.ListItems.Count To 1 Step -1
  Set Item = lvw.ListItems(i)
  If Item.Selected Then
    Index = Val(Item.Tag)
    lvw.ListItems.Remove i
    Zip.RemoveFile Index
  End If
Next

'Reload icon images
For Each Item In lvw.ListItems
  Index = Val(Item.Tag)
  Item.SmallIcon = Zip.IconKey(Index)
  'Item.Text = Zip.Name(Index)
  'Item.Tag = CStr(Index)
Next
lvw.Sorted = True
lvw.Refresh
UpdateInfoPanel False
Screen.MousePointer = vbNormal
Exit Sub

CmdDelErr:
Ok = ReportErrorCont("DeleteItems()", ModName, Err, Error)
If Ok Then
  Resume Next
Else
  Exit Sub
End If

End Sub

Private Sub mnuActionAddFiles_Click()
'--------------------------------------
'Add files to the archive
'--------------------------------------
AddFiles

End Sub

Private Sub mnuActionAddFolder_Click()
'--------------------------------------
'Add a folder (& subfolders) to the archive
'--------------------------------------
AddFolder

End Sub

Private Sub mnuActionDelete_Click()
'---------------------------------------------------------
'Delete selected files from list (but not from archive!)
'---------------------------------------------------------
DeleteItems

End Sub

Private Sub mnuActionRefresh_Click()
'--------------------------------------
'Refresh files in the archive
'--------------------------------------
RefreshFiles

End Sub

Private Sub mnuFileExit_Click()
'-----------------------------------------
'Exit program
'-----------------------------------------
Unload Me
End Sub

Private Sub mnuFileNew_Click()
'---------------------------------------------------------
'Create a new archive
'---------------------------------------------------------
NewArchive

End Sub

Private Sub mnuFileOpen_Click()
'----------------------------------------------
'Open an existing archive
'----------------------------------------------
OpenArchive

End Sub

Private Sub mnuFileUnzip_Click()
'--------------------------------------------
'Unzip files to selected folder
'--------------------------------------------
UnzipFile

End Sub

Private Sub mnuFileZip_Click()
'---------------------------------------------
'Compress files and save archive
'---------------------------------------------
ZipFile

End Sub

Private Sub mnuHelpAbout_Click()
'---------------------------------------------
'Show about box with pride! :)
'---------------------------------------------
frmAzAbout.Show vbModal, Me

End Sub

Private Sub mnuHelpReg_Click()
'-----------------------------------------------
'This registers the .azp file extension
'-----------------------------------------------
Zip.RegisterArielFileTypes

End Sub

Private Sub mnuHelpUnreg_Click()
'-----------------------------------------------
'This unregisters the .azp/azs file extensions
'-----------------------------------------------
Zip.UnregisterArielFileTypes

End Sub


Private Sub mnuViewText_Click()
'--------------------------------------------------
'Show/hide toolbar text
'--------------------------------------------------
Dim Button As Button

mnuViewText.Checked = Not (mnuViewText.Checked)
For Each Button In tbToolbar.Buttons
  If mnuViewText.Checked Then
    Button.Caption = Button.Key
  Else
    Button.Caption = ""
  End If
Next
DoEvents
Resize

End Sub

Private Sub mnuViewToolbar_Click()
'-----------------------------------------
'Show/hide toolbar
'-----------------------------------------
mnuViewToolbar.Checked = Not (mnuViewToolbar.Checked)
tbToolbar.Visible = mnuViewToolbar.Checked
Resize

End Sub

Private Sub tbToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
'---------------------------------------------------------
'Select items from toolbar
'---------------------------------------------------------
On Error Resume Next

Select Case Button.Key
Case "New"
  NewArchive
Case "Open"
  OpenArchive
Case "Add Folder"
  AddFolder
Case "Add Files"
  AddFiles
Case "Refresh"
  RefreshFiles
Case "Delete"
  DeleteItems
Case "Zip"
  ZipFile
Case "Unzip"
  UnzipFile
End Select

End Sub

Private Sub Zip_ChangeDisk(Drive As Scripting.Drive, Message As String, File As String)
'---------------------------------------------
'Show disk change form
'---------------------------------------------
Load frmAzDisk
Set frmAzDisk.Drive = Drive
frmAzDisk.lblInfo = Message
frmAzDisk.SourceFile = File
frmAzDisk.CheckDriveStatus True
frmAzDisk.Show vbModal, Me
'Set the return status
Zip.Cancel = Not (DlgOk)

End Sub

Private Sub Zip_Progress(Value As Single, Info As String)
'-----------------------------------------
'Update the progress bar
'-----------------------------------------
Progress Value, Info

End Sub

Private Sub Zip_StatusChange(NewStatus As ArZipStatus)
'-----------------------------------------
'Update status of zip archive
'-----------------------------------------
UpdateStatus

End Sub


