VERSION 5.00
Object = "{C1C2430B-978A-11D4-9744-004F490561B3}#11.0#0"; "ARIEL BROWSE CTRL.OCX"
Begin VB.Form frmAzZip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Save Ariel Zip File"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AzSave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4620
      TabIndex        =   3
      ToolTipText     =   "Cancel selection and close window"
      Top             =   2220
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Ok"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3420
      TabIndex        =   2
      ToolTipText     =   "Accept selection and close window"
      Top             =   2220
      Width           =   1095
   End
   Begin VB.Frame fr 
      Height          =   2115
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5715
      Begin VB.ComboBox cmbSpanOption 
         Height          =   315
         ItemData        =   "AzSave.frx":058A
         Left            =   3000
         List            =   "AzSave.frx":05AC
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1380
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CheckBox chkSpan 
         Alignment       =   1  'Right Justify
         Caption         =   "Disk spanning"
         Height          =   195
         Left            =   300
         TabIndex        =   6
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ComboBox cmbLevel 
         Height          =   315
         ItemData        =   "AzSave.frx":061F
         Left            =   3000
         List            =   "AzSave.frx":0641
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Select the compression level - low levels are faster but give lower compression"
         Top             =   840
         Width           =   2415
      End
      Begin ArielBrowseCtrl.ArielBrowseFile ArFile 
         Height          =   315
         Left            =   960
         TabIndex        =   8
         Top             =   300
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Select an Ariel zip file"
         Proper          =   -1  'True
         FileDialogType  =   1
         Filter          =   "Ariel Zip Files (*.azp) | *.azp|"
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Zip file"
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   9
         Top             =   360
         Width           =   465
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Compression Level"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   5
         Top             =   900
         Width           =   1335
      End
      Begin VB.Label lblOption 
         AutoSize        =   -1  'True
         Caption         =   "Span size"
         Height          =   195
         Left            =   2160
         TabIndex        =   1
         Top             =   1440
         Visible         =   0   'False
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmAzZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------
'Module     : frmAzNew
'Description: New Zip File Window
'Release    : 2001 VB6
'Copyright  : © T De Lange
'----------------------------------------------------------------
Option Base 0
Option Explicit
DefLng H-N
DefBool O

Const ModName = "New Zip File"
Dim Loading As Boolean
Dim fso As New FileSystemObject

Sub CheckDrive()
'------------------------------------
'Validate ctrls
'------------------------------------
Dim Drive As Drive, DriveSpec As String
If chkSpan = 1 Then
  cmbSpanOption.Visible = True
  lblOption.Visible = True
  DriveSpec = fso.GetDriveName(fso.GetAbsolutePathName(ArFile.File))
  Set Drive = fso.GetDrive(DriveSpec)
  FillCmbSpanOption cmbSpanOption, Drive.DriveType, rg.SpanOption
Else
  cmbSpanOption.Visible = False
  lblOption.Visible = False
End If

End Sub





Sub FillCmbSpanOption(c As ComboBox, DriveType As DriveTypeConst, Optional Default As Variant)
'-------------------------------------------------
'Fill CmbSpanOption with options depending on
'drive type
'-------------------------------------------------
c.Clear
c.AddItem "1.44Mb"
c.AddItem "1.40Mb"
c.AddItem "1.20Mb"
c.AddItem "1.00Mb"
c.AddItem "720 kb"
c.AddItem "700 kb"
If DriveType = Removable Then
  c.AddItem "100% Capacity"
  c.AddItem "99% Capacity"
  c.AddItem "98% Capacity"
  c.AddItem "95% Capacity"
  c.AddItem "90% Capacity"
End If
If Not (IsMissing(Default)) Then
  If Default < 0 Then
    c.ListIndex = 0
  ElseIf Default < c.ListCount Then
    c.ListIndex = Default
  Else
    c.ListIndex = 0
  End If
Else
  c.ListIndex = 0
End If

End Sub

Private Sub chkSpan_Click()
CheckDrive
End Sub

Private Sub cmdCancel_Click()
'-------------------------------------------
'Cancel changes & close
'-------------------------------------------
DlgOk = False
Unload Me

End Sub

Private Sub cmdOk_Click()
'-------------------------------------------
'Accept changes & close
'-------------------------------------------
rg.ZipFile = ArFile.File
rg.CompressLevel = cmbLevel.ListIndex
rg.Spanning = (chkSpan = 1)
rg.SpanOption = cmbSpanOption.ListIndex
DlgOk = True
Unload Me

End Sub


Private Sub Form_Load()
'---------------------------------------------------
'Load & Resize
'---------------------------------------------------
Dim l, t
Dim i

On Error Resume Next
Loading = True
If Me.WindowState = vbNormal Then
  l = GetProfile(ModName, "Left", Me.Left)
  t = GetProfile(ModName, "Top", Me.Top)
  'w = GetProfile(ModName, "Width", Me.Width)
  'h = GetProfile(ModName, "Height", Me.Height)
  Me.Move l, t  ', w, h
End If
ArFile.File = rg.ZipFile
cmbLevel.ListIndex = rg.CompressLevel
chkSpan = Abs(rg.Spanning)
CheckDrive
'cmbSpanOption.ListIndex = rg.SpanOption
Loading = False

End Sub

Private Sub ArFile_Change(Text As String)
ArFile.File = Text
cmdOk.Enabled = ArFile.FileName <> ""
If Right(ArFile.FileName, 4) <> ".azp" Then
  ArFile.File = ArFile.File & ".azp"
End If
CheckDrive
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
'-----------------------------------------------------------
'Save settings
'-----------------------------------------------------------
Dim i, Key As String

If Me.WindowState = vbNormal Then
  SaveProfile ModName, "Left", Me.Left
  SaveProfile ModName, "Top", Me.Top
  'SaveProfile ModName, "Width", Me.Width
  'SaveProfile ModName, "Height", Me.Height
End If

End Sub


