VERSION 5.00
Begin VB.Form frmAzAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Ariel Zip"
   ClientHeight    =   4605
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   8355
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AzAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3178.453
   ScaleMode       =   0  'User
   ScaleWidth      =   7845.775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fr 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4515
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   8235
      Begin VB.CommandButton cmdOK 
         Cancel          =   -1  'True
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   6720
         TabIndex        =   0
         ToolTipText     =   "Close window"
         Top             =   3960
         Width           =   1275
      End
      Begin VB.PictureBox picIcon 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         Left            =   300
         Picture         =   "AzAbout.frx":0E42
         ScaleHeight     =   1369.55
         ScaleMode       =   0  'User
         ScaleWidth      =   1380.085
         TabIndex        =   2
         Top             =   480
         Width           =   1965
      End
      Begin VB.Label lblMail 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "E-mail: tomdl@attglobal.net"
         Height          =   195
         Left            =   6000
         TabIndex        =   11
         Tag             =   "Company"
         Top             =   3360
         Width           =   1995
      End
      Begin VB.Label lblTel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tel. +27-83-304-0934"
         Height          =   195
         Left            =   6390
         TabIndex        =   10
         Tag             =   "Company"
         Top             =   3135
         Width           =   1605
      End
      Begin VB.Label lblWarning 
         AutoSize        =   -1  'True
         Caption         =   $"AzAbout.frx":558C
         Height          =   585
         Left            =   300
         TabIndex        =   9
         Tag             =   "Warning"
         Top             =   3600
         Width           =   4230
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCopyright 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Copyright"
         Height          =   195
         Left            =   7290
         TabIndex        =   8
         Tag             =   "Copyright"
         Top             =   2700
         Width           =   705
      End
      Begin VB.Label lblCompany 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Company"
         Height          =   195
         Left            =   7320
         TabIndex        =   7
         Tag             =   "Company"
         Top             =   2925
         Width           =   675
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Product Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   2400
         TabIndex        =   6
         Tag             =   "Product"
         Top             =   540
         Width           =   2985
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         Caption         =   "Description (App.comments)"
         Height          =   195
         Left            =   2460
         TabIndex        =   5
         Top             =   1320
         Width           =   5040
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Platform"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6735
         TabIndex        =   4
         Tag             =   "Platform"
         Top             =   2280
         Width           =   1230
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2460
         TabIndex        =   3
         Tag             =   "Version"
         Top             =   1860
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmAzAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------
'Module     : AzAbout
'Description: About Ariel Zip
'Release    : 2000 VB6
'Copyright  : © T De Lange
'--------------------------------------------------------------------
Option Explicit
Option Base 0
DefLng H-N
DefBool O

Private Sub cmdOk_Click()
'---------------------------------
'Unload form
'---------------------------------
Unload Me

End Sub

Private Sub Form_Load()
'-----------------------------------------
'Load defaults
'-----------------------------------------
Me.Caption = "About " & App.Title
lblVersion.Caption = "Version " & App.Major & "." & Format(App.Minor, "00") & "." & Format(App.Revision, "00")
lblProductName.Caption = App.Title
lblPlatform.Caption = "for Windows 95/98/Me"
lblDescription.Caption = App.Comments
lblCopyright.Caption = "© " & App.LegalCopyright
lblCompany.Caption = App.CompanyName
  
End Sub

