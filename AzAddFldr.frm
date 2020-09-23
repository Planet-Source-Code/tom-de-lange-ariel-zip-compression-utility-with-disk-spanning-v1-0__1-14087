VERSION 5.00
Object = "{C1C2430B-978A-11D4-9744-004F490561B3}#11.0#0"; "ARIEL BROWSE CTRL.OCX"
Begin VB.Form frmAzAddFldr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Folder"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AzAddFldr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      ToolTipText     =   "Cancel selection and close window"
      Top             =   1740
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Ok"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      ToolTipText     =   "Click to accept changes.  If no check boxes are selected, an emtpy list will be created."
      Top             =   1740
      Width           =   1095
   End
   Begin VB.Frame fr 
      Height          =   1635
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   6435
      Begin VB.CheckBox chkAddSub 
         Caption         =   "Include subfolders"
         Height          =   195
         Left            =   1800
         TabIndex        =   5
         ToolTipText     =   "If checked, subfolders will also be included in the archive."
         Top             =   1020
         Width           =   2535
      End
      Begin ArielBrowseCtrl.ArielBrowseFolder ArFolder 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         ToolTipText     =   "The rootfolder is required for the relative path reference."
         Top             =   420
         Width           =   4335
         _ExtentX        =   7646
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
         RootFolder      =   2
         Caption         =   "Select a rootfolder for the archive"
         Object.ToolTipText     =   "The rootfolder is required for the relative path reference."
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Folder to add"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   2
         Top             =   480
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmAzAddFldr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------
'Module     : frmAzAddFldr
'Description: Add Folder/subfolders
'Release    : 2001 VB6
'Copyright  : © T De Lange
'----------------------------------------------------------------
Option Base 0
Option Explicit
DefLng H-N
DefBool O

Const ModName = "Add Folder"
Dim Loading As Boolean

Sub CheckCtrls()
'------------------------------------
'Validate ctrls
'------------------------------------
If FolderExist(ArFolder.Text) Then
  cmdOk.Enabled = True
Else
  cmdOk.Enabled = False
  If Not (Loading) Then
    MsgBox "Folder doesn't exist.", vbOKOnly
  End If
End If

End Sub

Private Sub ArFolder_Change(SelectedPath As String)
'--------------------------------------------
'Check if path exists
'--------------------------------------------
CheckCtrls

End Sub

Private Sub ArFolder_Click(SelectedPath As String)
'--------------------------------------------
'Check if path exists
'--------------------------------------------
CheckCtrls

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
rg.AddFolder = ArFolder.Text
rg.IncludeSubfolders = (chkAddSub = 1)
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
If rg.AddFolder <> "" Then
  ArFolder.Text = rg.AddFolder
Else
  ArFolder.Text = CurDir
End If
chkAddSub = Abs(rg.IncludeSubfolders)
Loading = False

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


