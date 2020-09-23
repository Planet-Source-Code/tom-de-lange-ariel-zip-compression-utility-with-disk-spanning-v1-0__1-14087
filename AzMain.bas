Attribute VB_Name = "modAzMain"
'------------------------------------------------------------------
'Module Name: AzMain.bas
'Description: Main module file for Arzip App
'Version    : V1.00 Dec 2000
'Release    : 2000
'Copyright  : © T de Lange
'------------------------------------------------------------------
Option Explicit
Option Base 0
DefLng H-N
DefBool O

'Registry entries are saved in an ini file, not registry!
Public Type RegistryEntries
  ZipFile As String                 'Open dialog box
  ZipFolder As String               'Open dialog box/Add folder dialog box
  IncludeFiles As Boolean           'Open dialog box
  IncludeSubfolders As Boolean      'Open dialog box
  AddFolder As String               'Add folder dialog box
  UnzipFolder As String
  CompressLevel As Integer
  Spanning As Boolean
  SpanOption As ArSpanOption
  UnzipAll As Boolean
  Overwrite As Boolean
  RegFileType As Boolean
End Type

Public rg As RegistryEntries
Public DlgOk As Boolean

'------------------------------------------------
'DLL Declarations
'------------------------------------------------
'Ini files
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Function CheckPath(ByVal Path As String) As String
'--------------------------------------------------
'Checks if path ends with "\". If not, add it.
'--------------------------------------------------
If Right(Path, 1) <> "\" Then
  CheckPath = Path & "\"
Else
  CheckPath = Path
End If

End Function
Function FolderExist(Path As String) As Boolean
'-------------------------------------------------------
'Check if folder exists on hard disk
'-------------------------------------------------------
Dim fso As New FileSystemObject

On Error GoTo FolderExistErr
FolderExist = fso.FolderExists(Path)
Exit Function

FolderExistErr:
FolderExist = False

End Function
Public Function GetProfile(ByVal Section As String, ByVal KeyName As String, ByVal Default As Variant) As Variant
'---------------------------------------------------------------------------------------
'Get a profile item
'---------------------------------------------------------------------------------------
Dim AppName As String, FileName As String, RetString As String
Dim DefSt As String
Dim Size, RetSize

AppName = App.Title
FileName = App.Path & "\" & AppName & ".ini"
DefSt = CStr(Default)
RetString = Space(128)
Size = Len(RetString)
RetSize = GetPrivateProfileString(Section, KeyName, DefSt, RetString, Size, FileName)
GetProfile = Left(RetString, RetSize)

End Function

Function RightFormat(Var As Variant, Frmt As Variant) As String
'-------------------------------------------------------
'Formats a string Right Justified in a listview by
'padding spaces to the left
'Var    : Variable/Expression to format
'Frmt   : Format String
'-------------------------------------------------------
Dim n As Integer

n = Len(Frmt)
RightFormat = Right(Space(n) & Format(Var, Frmt), n)

End Function
Public Sub SaveProfile(ByVal Section As String, ByVal KeyName As String, ByVal Value As Variant)
'---------------------------------------------------------------------------------------
'Write a profile item to an ini file
'---------------------------------------------------------------------------------------
Dim AppName As String, FileName As String
Dim ValSt As String, Valid

AppName = App.Title
FileName = App.Path & "\" & AppName & ".ini"
ValSt = CStr(Value)
Valid = WritePrivateProfileString(Section, KeyName, ValSt, FileName)

End Sub



Sub ReadRegistry()
'-----------------------------------------------
'Read Registry Settings
'-----------------------------------------------
Dim Sec As String       'Section

'Format = GetProfile(Section, Key, Default)
Sec = "General"
rg.ZipFile = GetProfile(Sec, "Zip File", "")
rg.ZipFolder = GetProfile(Sec, "Zip Folder", "")
rg.IncludeFiles = GetProfile(Sec, "Include Files", True)
rg.IncludeSubfolders = GetProfile(Sec, "Include Subfolders", True)
rg.UnzipFolder = GetProfile(Sec, "Unzip Folder", "")
rg.CompressLevel = GetProfile(Sec, "Compress Level", 5)
rg.Spanning = GetProfile(Sec, "Spanning", 0)
rg.SpanOption = GetProfile(Sec, "Span Option", 0)
rg.UnzipAll = GetProfile(Sec, "Unzip All", True)
rg.Overwrite = GetProfile(Sec, "Overwrite", True)
rg.RegFileType = GetProfile(Sec, "Register File Type", False)

End Sub
Sub ReportError(Proc As String, Module As String, ErrNo As Integer, ErrSt As String)
'----------------------------------------------------------------------------------------
'Report an error showing Procedure and Module which caused the error
'Display the error number, errorstring and Ok button
'----------------------------------------------------------------------------------------
Dim ErrResponse As Integer
Dim MousePointer As Integer

MousePointer = Screen.MousePointer
Screen.MousePointer = vbNormal
ErrResponse = MsgBox("Error found in the " & Proc & " procedure of the " & Module & " module." & vbCrLf & _
"Error No" & Str(ErrNo) & " : " & ErrSt, vbExclamation, "Error")
Screen.MousePointer = MousePointer

End Sub

Function ReportErrorCont(Proc As String, Module As String, ErrNo As Integer, ErrSt As String) As Boolean
'----------------------------------------------------------------------------------------
'Report an error showing Procedure and Module which caused the error
'Display the error number and errorstring
'Ask user if program should continue
'----------------------------------------------------------------------------------------
Dim ErrResponse As Integer
Dim MousePointer As Integer

MousePointer = Screen.MousePointer
Screen.MousePointer = vbNormal
ErrResponse = MsgBox("Error found in the " & Proc & " procedure of the " & Module & " module." & vbCrLf & _
"Error No" & Str(ErrNo) & " : " & ErrSt & vbCrLf & "Continue?", _
vbCritical + vbOKCancel, "Error")
Screen.MousePointer = MousePointer
ReportErrorCont = (ErrResponse = vbOK)
End Function

Function ReportErrorAbort(Proc As String, Module As String, ErrNo As Integer, ErrSt As String) As Long
'----------------------------------------------------------------------------------------
'Report an error showing Procedure and Module which caused the error
'Display the error number and errorstring
'Ask user if program should abort/retry/ignore
'----------------------------------------------------------------------------------------
Dim ErrResponse As Integer
Dim MousePointer As Long

MousePointer = Screen.MousePointer
Screen.MousePointer = vbNormal
ErrResponse = MsgBox("Error found in the " & Proc & " procedure of the " & Module & " module." & vbCrLf & _
"Error No" & Str(ErrNo) & " : " & ErrSt, _
vbCritical + vbAbortRetryIgnore, "Error")
Screen.MousePointer = MousePointer
ReportErrorAbort = ErrResponse

End Function

Sub SaveRegistry()
'------------------------------------------------
'Object   : Write info to Project ini file
'------------------------------------------------
Dim Sec As String

Sec = "General"
Call SaveProfile(Sec, "Zip File", rg.ZipFile)
Call SaveProfile(Sec, "Zip Folder", rg.ZipFolder)
Call SaveProfile(Sec, "Include Files", rg.IncludeFiles)
Call SaveProfile(Sec, "Include Subfolders", rg.IncludeSubfolders)
Call SaveProfile(Sec, "Unzip Folder", rg.UnzipFolder)
Call SaveProfile(Sec, "Compress Level", rg.CompressLevel)
Call SaveProfile(Sec, "Spanning", rg.Spanning)
Call SaveProfile(Sec, "Span Option", rg.SpanOption)
Call SaveProfile(Sec, "Unzip All", rg.UnzipAll)
Call SaveProfile(Sec, "Overwrite", rg.Overwrite)
Call SaveProfile(Sec, "Register File Type", rg.RegFileType)

End Sub

Sub Main()
'--------------------------------------------
'Main starting procedure
'--------------------------------------------
ReadRegistry
frmAzMain.Show
If Not (rg.RegFileType) Then
  'Register file types
  frmAzMain.Zip.RegisterArielFileTypes
  rg.RegFileType = True
End If
'Check command string
If Command <> "" Then
  rg.ZipFile = Command
  With frmAzMain
    .Zip.OpenArchive rg.ZipFile
    .FillLvw
    .UpdateInfoPanel False
  End With
End If

End Sub

Sub PrgExit()
'------------------------------------------------
'Object   : Exit Program
'------------------------------------------------
Dim f As Form

SaveRegistry
For Each f In Forms
  If f.Name <> "frmAzMain" Then
    Unload f
  End If
Next
Unload frmAzMain
'End -------------> Don't use AT ALL!!! Ends program without unloading frmMain, causing prg not to end execution
                    'and therefore to stay memory resident!!!
End Sub

