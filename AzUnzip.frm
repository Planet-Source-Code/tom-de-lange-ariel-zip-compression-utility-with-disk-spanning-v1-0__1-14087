VERSION 5.00
Object = "{C1C2430B-978A-11D4-9744-004F490561B3}#11.0#0"; "ARIEL BROWSE CTRL.OCX"
Begin VB.Form frmAzUnzip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Unzip Ariel File"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AzUnzip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      ToolTipText     =   "Cancel selection and close window"
      Top             =   2340
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Ok"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      ToolTipText     =   "Accept selection and close window"
      Top             =   2340
      Width           =   1095
   End
   Begin VB.Frame fr 
      Height          =   2235
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   5715
      Begin VB.CheckBox chkOverwrite 
         Caption         =   "Overwrite existing files"
         Height          =   195
         Left            =   1140
         TabIndex        =   7
         ToolTipText     =   "If checked, extracts all files. If not, skips files that already exist in the destination folder"
         Top             =   1560
         Width           =   2655
      End
      Begin VB.OptionButton opSel 
         Caption         =   "Unzip selected files"
         Height          =   195
         Index           =   1
         Left            =   3240
         TabIndex        =   6
         Top             =   1020
         Width           =   1755
      End
      Begin VB.OptionButton opSel 
         Caption         =   "Unzip all files"
         Height          =   195
         Index           =   0
         Left            =   1140
         TabIndex        =   5
         Top             =   1020
         Width           =   1755
      End
      Begin ArielBrowseCtrl.ArielBrowseFolder ArFolder 
         Height          =   315
         Left            =   1140
         TabIndex        =   1
         Top             =   420
         Width           =   4275
         _ExtentX        =   7541
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
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Destination Folder"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   825
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmAzUnzip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------
'Module     : frmAzUnzip
'Description: Unzip File Window
'Release    : 2001 VB6
'Copyright  : © T De Lange
'----------------------------------------------------------------
Option Base 0
Option Explicit
DefLng H-N
DefBool O

Const ModName = "Unzip File"
Dim Loading As Boolean

Sub CheckCtrls()
'------------------------------------
'Validate ctrls
'------------------------------------
If ArFolder.Text <> "" Then
  cmdOk.Enabled = True
Else
  cmdOk.Enabled = False
  If Not (Loading) Then
    MsgBox "Please provide a valid path.", vbOKOnly Or vbInformation
  End If
End If

End Sub


Private Sub ArFolder_Change(SelectedPath As String)
'--------------------------------------------
'Handle folder changes
'--------------------------------------------
CheckCtrls

End Sub

Private Sub ArFolder_Click(SelectedPath As String)
'--------------------------------------------
'Handle folder changes
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
Dim n
rg.UnzipFolder = ArFolder.Text
If Right(rg.UnzipFolder, 1) = "\" Then
  n = Len(rg.UnzipFolder)
  rg.UnzipFolder = Mid(rg.UnzipFolder, n - 1)
End If
rg.UnzipAll = opSel(0)
rg.Overwrite = True

DlgOk = True
Unload Me

End Sub


Private Sub Form_Load()
'---------------------------------------------------
'Load & Resize
'---------------------------------------------------
Dim l, t

On Error Resume Next
Loading = True
If Me.WindowState = vbNormal Then
  l = GetProfile(ModName, "Left", Me.Left)
  t = GetProfile(ModName, "Top", Me.Top)
  Me.Move l, t
End If
If rg.UnzipFolder = "" Then
  ArFolder.Text = rg.ZipFolder
Else
  ArFolder.Text = rg.UnzipFolder
End If
If rg.UnzipAll Then
  opSel(0).Value = True
Else
  opSel(1).Value = True
End If
chkOverwrite = Abs(rg.Overwrite)

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
End If

End Sub


