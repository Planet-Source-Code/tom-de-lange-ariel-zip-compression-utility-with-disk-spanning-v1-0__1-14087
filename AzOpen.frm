VERSION 5.00
Object = "{C1C2430B-978A-11D4-9744-004F490561B3}#11.0#0"; "ARIEL BROWSE CTRL.OCX"
Begin VB.Form frmAzOpen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open Ariel Zip File"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   Icon            =   "AzOpen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4620
      TabIndex        =   4
      ToolTipText     =   "Cancel open and close the window"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Ok"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3420
      TabIndex        =   3
      ToolTipText     =   "Open the selected zip file"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Frame fr 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5715
      Begin ArielBrowseCtrl.ArielBrowseFile ArFile 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   360
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
         Filter          =   "Ariel Zip Files (*.azp) | *.azp|"
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Zip file"
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   2
         Top             =   420
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmAzOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------
'Module     : frmAzOpen
'Description: Open Ariel Zip File
'Release    : 2001 VB6
'Copyright  : © T De Lange
'----------------------------------------------------------------
Option Base 0
Option Explicit
DefLng H-N
DefBool O

Const ModName = "Open Archive"

Private Sub ArFile_Change(Text As String)
cmdOk.Enabled = ArFile.FileName <> ""
If Right(ArFile.FileName, 4) <> ".azp" Then
  ArFile.File = ArFile.File & ".azp"
End If
  
End Sub

Private Sub cmdCancel_Click()
DlgOk = False
Unload Me
End Sub

Private Sub cmdOk_Click()
rg.ZipFile = ArFile.File
DlgOk = True
Unload Me
End Sub

Private Sub Form_Load()
'---------------------------------------------------
'Load & Resize
'---------------------------------------------------
Dim l, t

On Error Resume Next
If Me.WindowState = vbNormal Then
  l = GetProfile(ModName, "Left", Me.Left)
  t = GetProfile(ModName, "Top", Me.Top)
  Me.Move l, t
End If
ArFile.File = rg.ZipFile

End Sub


Private Sub Form_Unload(Cancel As Integer)
'-----------------------------------------------------------
'Save settings
'-----------------------------------------------------------
If Me.WindowState = vbNormal Then
  SaveProfile ModName, "Left", Me.Left
  SaveProfile ModName, "Top", Me.Top
End If

End Sub


