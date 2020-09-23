VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------
'Module     : ArZip Class
'Description: Ariel Compression Class
'Release    : 2001 VB6
'Copyright  : © T De Lange, 2000
'------------------------------------------------------------------
'Credits:
'a) Peter Meier, Planet Source Code for the icon extraction technique
'   in his 'DelRecent' posting
'b) Recursive folder list: Brad V (Planet Source Code)
'   Please see http://www.Planet-Source-Code.com/xq/ASP/txtCodeId.11429/lngWId.1/qx/vb/scripts/ShowCode.htm
'   for details.
'------------------------------------------------------------------
'Notes: Add a reference to Microsoft Scripting Runtime dll
'       in the project references box
'------------------------------------------------------------------
Option Base 1
Option Explicit
DefLng A-N, P-Z
DefBool O

'----------------------------------------------------------
'Private Constants
'----------------------------------------------------------
Const ModName = "ArielZip Class"
'Icon Sizes in pixels
Private Const LARGE_ICON As Integer = 32    'Icon Size, pixels
Private Const SMALL_ICON As Integer = 16    'Icon size
Private Const MAX_PATH = 260                'Max length of path+file name
Private Const ILD_TRANSPARENT = &H1         'Display icon transparent
Private Const BMP_SIZE As Long = 822        'Size in bytes of icon bitmaps
Private Const CHUNK_SIZE As Long = 1048576  'Chunk size 1Mb to write archive file

'ShellInfo Flags
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000 'System icon index
Private Const SHGFI_LARGEICON = &H0       'Large icon
Private Const SHGFI_SMALLICON = &H1       'Small icon
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400

Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
        Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
        Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

'----------------------------------------------------------
'Public Enums
'----------------------------------------------------------
Public Enum ArZipStatus
  azsEmpty          'Empty
  azsReady          'File & Rootfolder specified
  azsCreated        'Zipfile created (ready for extraction)
  azsBusy
End Enum

Public Enum ArSpanOption
  azo1440     '1.44Mb
  azo1400     '1.40Mb
  azo1200     '1.20Mb
  azo1000     '1.00Mb
  azo720      '720kb
  azo700      '700kb
  azo100p     '100% Capacity (Removable drives only)
  azo99p      '99% Capacity     "
  azo98p      '98% Capacity     "
  azo95p      '95% Capacity     "
  azo90p      '90% Capacity     "
End Enum
  
Private Enum ZLibErrors
  'Positive numbers are normal occasions, neg's are errors
  zlOk = 0
  zlStreamEnd = 1
  zlNeedDict = 2
  zlErrNo = -1
  zlStreamError = -2      'Invalid compression level parameter
  zlDataError = -3        'Input data corrupted
  zlMemError = -4         'Not Enough Memory
  zlBufError = -5         'Not enough space in output buffer
  zlVersionError = -6
End Enum
  
'----------------------------------------------------------
'Private Type Defs
'----------------------------------------------------------
Private Type SHFILEINFO              'As required by ShInfo
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type

Private Type ArielFile
  Key As String         'Path\File.ext
  Name As String        'Filename including extension, without path
  Ext As String         'File extension
  Path As String        'Path, eg C:\My Documents (to locate file on HDD)
  RelPath As String     'Relative path in archive e.g. \Memos
  Zipped As Boolean     'Has file been zipped already?
  Size As Long          'Original file size
  ZipSize As Long       'Zip size
  Ratio As Single       'Zip ratio 0 to 1
  Modified As Date
  IconKey As String
  Offset As Long        'Offset of file in CURRENT volume
  ChkSum As Byte        'XOR Checksum from original file back to unzipped file
  ChkSumZip As Byte     'XOR Checksum from zipped file write to read
  Selected As Boolean
End Type

Private Type ArHeader
  'Added at the start of each span volume
  'Id As String * 3     'Don't use fixed/variable strings in copymem()
  Major As Integer
  Minor As Integer
  VolNo As Byte         'Span Volume no, starting at 1
  NextVol As Byte       'Indicates another volume follows (1), or not (0)
  ZipSize As Long       'Total zip file size, prior to spanning
  NoFiles As Long       'No of files in archive
  NoIcons As Long       'No of unique icons (bitmaps) in archive
  Offset  As Long       'Offset to start of file list
End Type

Private Type ArFileRec
  Offset As Long
  OrigSize As Long
  ZipSize As Long
  Modified As Date
  ChkSum As Byte
  ChkSumZip As Byte
End Type
  
'----------------------------------------------------------
'Functions & Procedures
'----------------------------------------------------------
'Functions to extract icons & place them in a picture box
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
    (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long

Private Declare Function ImageList_Draw Lib "comctl32.dll" _
    (ByVal himl&, ByVal i&, ByVal hDCDest&, _
    ByVal x&, ByVal y&, ByVal flags&) As Long

'General Kernel.dll functions
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'Sleep & Timer functions - see Wait()
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

'ZLib 1.1.3 functions
'Private Declare Function CompressDef Lib "zlib.dll" Alias "compress" (Dest As Any, DestLen As Long, Source As Any, ByVal SourceLen As Long) As Long
Private Declare Function Compress Lib "zlib.dll" Alias "compress2" (Dest As Any, DestLen As Long, Source As Any, ByVal SourceLen As Long, ByVal Level As Long) As Long
Private Declare Function Uncompress Lib "zlib.dll" Alias "uncompress" (Dest As Any, DestLen As Long, Source As Any, ByVal SourceLen As Long) As Long

'----------------------------------------------------------
'Private property variables
'----------------------------------------------------------
Private mRootFolder As String
Private mStatus As ArZipStatus
Private mExist As Boolean         'Does a current zip file exist?
Private mZipFileNew As String     'Zip path & file name
Private mZipFile As String     'Current zip file
Private mTempFolder As String     'Scratch area (used to save icon bitmaps)
Private mTempFile As String       'For use only during disk spanning, placed in tempfolder
Private mBackupFile As String     'Used for storage of existing data, during zipping
Private mUnzipFolder As String
Private mElapsedTime As Single
Private mCompressLevel As Integer
Private mSpanning As Boolean      'Must new zip file be spanned?
Private mSpanned As Boolean       'Has current zip file been spanned?
Private mSpanOption As Long       'Span option for new zip file
Private mSpanSize As Long         'Size of selected span option (new file)
Private mTotalZipSize As Long
Private mCancel As Boolean        'Used to cancel zip/unzip operation

'----------------------------------------------------------
'Events
'----------------------------------------------------------
Public Event StatusChange(NewStatus As ArZipStatus)
Public Event Progress(Value As Single, Info As String)
Public Event ChangeDisk(Drive As Drive, Message As String, File As String)

'----------------------------------------------------------
'Private variables
'----------------------------------------------------------
Private ShInfo As SHFILEINFO
Private mLvwFiles As ListView
Private mIml As ImageList                 'Imagelist containing small icons
Private mPic As PictureBox                'Temporay container for small icon
Private mPicDef As PictureBox             'Default picture for <None> icons
Private mFso As New FileSystemObject      'File system object
Private aFiles() As ArielFile             'Array of files
Private nFiles As Long                    'No of files in array
Private mProgress As Single               'Progress 0 to 1
Private nIcons As Long                    'No of unique icons
Private nPos As Long                      'Temporary posistion in archive file
Private Function AddByte(ByRef Bytes() As Byte, Value As Byte) As Long
'---------------------------------------------------------------------
'Add a byte value to the byte array
'---------------------------------------------------------------------
Dim ns

On Error Resume Next
ns = UBound(Bytes)
If ns = 0 Then
  ReDim Bytes(1) As Byte      'Necessary, to get rid of 0 lower bound
Else
  ReDim Preserve Bytes(ns + 1) As Byte
End If
Call CopyMem(Bytes(ns + 1), Value, 1)
AddByte = UBound(Bytes)

End Function

Private Function AddFile(ByRef File As File) As Boolean
'---------------------------------------------------------------
'Adds a file to the array
'Returns true if successful
'---------------------------------------------------------------
'Notes
'a) The file to be included is validated for existence at source
'   and if it is not the current archive file (or spanned volume)
'b) If it already is in the list (and even in the archive) the
'   file is 'freshened' with the current source info. I.e. by
'   default the source file is to be considered a later/newer
'   version. The zipped flag is set to false to prevent the
'   current zipped version being included in the archive.
'---------------------------------------------------------------
Dim Ok, n

Ok = ValidateFile(File)
If Ok Then
  n = FileInList(File)
  If n = 0 Then
    'Add file to the list
    nFiles = nFiles + 1
    ReDim Preserve aFiles(nFiles) As ArielFile
    n = nFiles
    aFiles(n).ZipSize = File.Size       'Keep the same as original size for now
    aFiles(n).Ratio = 1
    aFiles(n).Zipped = False
  Else
    'Freshen file
    aFiles(n).Zipped = False
  End If
  aFiles(n).Name = File.Name
  aFiles(n).Modified = File.DateLastModified
  aFiles(n).Size = File.Size
  aFiles(n).Path = File.ParentFolder
  aFiles(n).RelPath = RelPath(File.ParentFolder, mRootFolder)
  aFiles(n).Ext = FileExtension(File.Name)
  aFiles(n).Key = aFiles(nFiles).RelPath & "\" & aFiles(nFiles).Name
  aFiles(n).IconKey = GetIconKey(File)
  AddFile = True
Else
  AddFile = False
End If

End Function
Private Function AddBlankFileRecords(Bytes() As Byte) As Long
'---------------------------------------------------
'Add an empty file header to the byte array for
'each file
'a) This to be updated later
'b) The string portions are added in a separater sub
'   after the file list has been written
'Return the ubound() of Bytes()
'---------------------------------------------------
Dim Afr As ArFileRec
Dim ns, n

On Error Resume Next
ns = UBound(Bytes)
n = Len(Afr) * nFiles
If ns = 0 Then
  ReDim Bytes(n) As Byte          'Necessary to get rid of a 0 lower bound
Else
  ReDim Preserve Bytes(ns + n) As Byte
End If
AddBlankFileRecords = UBound(Bytes)

End Function

Private Function AddFileName(Bytes() As Byte, i As Long) As Long
'-----------------------------------------------------------------
'Add the string portion of a file to the byte array
'a) The numeric portions are dealt with separately
'   as they need to be updated at the end of the zip file
'b) All strings are variable length, so add a vbCr at the end
'Return the ubound() of Bytes()
'-----------------------------------------------------------------
On Error Resume Next

AddString Bytes, aFiles(i).Name, True         'File name incl ext
AddString Bytes, aFiles(i).RelPath, True      'Relative path (to rootfolder)
AddString Bytes, aFiles(i).Path, True         'Absolute path
AddString Bytes, aFiles(i).IconKey, True      'Icon key in image list
AddFileName = UBound(Bytes)

End Function

Public Function AddFiles(FileList() As String) As Long
'------------------------------------------------
'Add files to the archive list
'FileList :  Contains the file name(s)
'Return   :  NoFiles added
'------------------------------------------------
'Note: if only 1 file was added, the file
'      path & name is given in element 0
'      if multiselection was done, element 0
'      contains the path and the others the
'      file names, excl path
'------------------------------------------------
Dim Path As String      'Path of files
Dim File As File
Dim i, n, m, Ok

On Error Resume Next
Screen.MousePointer = vbArrowHourglass
SetStatus azsBusy
'mlvwfiles.SmallIcons=nothing
n = UBound(FileList)
If n = 0 Then
  'Single File
  Set File = mFso.GetFile(FileList(0))
  Ok = AddFile(File)
  m = IIf(Ok, 1, 0)
Else
  'Multiple Files
  'Get the path (element 0)
  Path = CheckPath(FileList(0))
  For i = 1 To n
    Set File = mFso.GetFile(Path & FileList(i))
    Ok = AddFile(File)
    If Ok Then m = m + 1
    'DoEvents
  Next
  AddFiles = m          'Set the return parameter
End If
mLvwFiles.SmallIcons = mIml       'Set the smallicons imagelist
Screen.MousePointer = vbNormal
SetStatus azsReady

End Function

Public Function AddFolder(Rootfolder As String, IncludeSubfolders As Boolean) As Long
'---------------------------------------------------------------------------------------------------------------
'Add a folder (and subfolders) to the archive list
'Rootfolder         : Path of folder to add
'IncludeSubfolders  : Include subfolders' contents in the archive
'Return             : Total NoFiles now in list
'---------------------------------------------------------------------------------------------------------------

On Error Resume Next
Screen.MousePointer = vbArrowHourglass
SetStatus azsBusy
ScanTree Rootfolder, IncludeSubfolders    'Recursive tree scan
mLvwFiles.SmallIcons = mIml               'Set the smallicons imagelist
Screen.MousePointer = vbNormal
SetStatus azsReady
AddFolder = nFiles                        'Set the return parameter

End Function

Private Function AddIconRecord(Bytes() As Byte, Bmp() As Byte, Key As String) As Long
'---------------------------------------------------
'Add an icon record (key & bmp) to the Byte array
'Return the ubound() of Bytes()
'---------------------------------------------------
On Error Resume Next

AddString Bytes, Key, True      'Icon key (see imglist)
AddBytes Bytes, Bmp
AddIconRecord = UBound(Bytes)

End Function

Private Function AddInt(ByRef Bytes() As Byte, Value As Integer) As Long
'---------------------------------------------------------------------
'Add integer value to the byte array
'---------------------------------------------------------------------
Dim ns

On Error Resume Next
ns = UBound(Bytes)
If ns = 0 Then
  ReDim Bytes(2) As Byte      'Necessary, to get rid of 0 lower bound
Else
  ReDim Preserve Bytes(ns + 2) As Byte
End If
Call CopyMem(Bytes(ns + 1), Value, 2)
AddInt = UBound(Bytes)

End Function

Private Function AddLong(ByRef Bytes() As Byte, Value As Long) As Long
'---------------------------------------------------------------------
'Add long value to the byte array
'---------------------------------------------------------------------
Dim ns

On Error Resume Next
ns = UBound(Bytes)
If ns = 0 Then
  ReDim Bytes(4) As Byte      'Necessary, to get rid of 0 lower bound
Else
  ReDim Preserve Bytes(ns + 4) As Byte
End If
Call CopyMem(Bytes(ns + 1), Value, 4)
AddLong = UBound(Bytes)

End Function

Private Function AddString(ByRef Bytes() As Byte, ByVal St As String, Optional VarLen As Boolean = False) As Long
'---------------------------------------------------------------------
'Add string St to the byte array
'If VarLen is true, add vbCr to end of string
'Return the ubound (No) of bytes
'---------------------------------------------------------------------
Dim StrBytes() As Byte      '0-based
Dim n, ns

On Error Resume Next
ns = UBound(Bytes)
If VarLen Then St = St & vbCr
StrBytes = StrConv(St, vbFromUnicode)   '0-based
n = Len(St)
If ns = 0 Then
  ReDim Bytes(n) As Byte      'Necessary, to get rid of 0 lower bound
Else
  ReDim Preserve Bytes(ns + n) As Byte
End If
CopyMem Bytes(ns + 1), StrBytes(0), n
AddString = UBound(Bytes)

End Function

Private Sub BackupZipFile()
'-------------------------------------------------------------
'Make a backup of a normal (unspanned) zipfile to the tempfolder
'This is similar to the unspanzipfile() routine
'-------------------------------------------------------------
Dim Drive As Drive          'FSO Drive object
Dim DriveSpec As String     'Required to resolve the drive object
Dim IsFDD As Boolean        'Is destination drive removable?
Dim nfd, nfs                'File Numbers
Dim SourceFile As String    'Name of sourcefile, incl path
Dim DestFile As String      'Complete name of destination file, incl path & ext
Dim nLeft As Long           'Remaining length of file to process
Dim nDest As Long           'Chunck to write
Dim Bytes() As Byte         'Bytes array
Dim SwapDisk As Boolean     'Flag indicating that a disk needs to be inserted/changed
Dim Sum, Total              'Progress indicator
Dim Ans

On Error GoTo BackupZipFileErr
'--------------------------------------------------
'Determine type of drive to read from
'This is necessary since if it is a removable drive
'(i.e. FDD), a dialog must be shown to enter the disk
'--------------------------------------------------
DriveSpec = mFso.GetDriveName(mFso.GetAbsolutePathName(mZipFile))
Set Drive = mFso.GetDrive(DriveSpec)
IsFDD = (Drive.DriveType = Removable)

'---------------------------------------------------
'Get the source file name & initialise other var's
'---------------------------------------------------
SourceFile = mZipFile     'Current zip file
Sum = 0                   'Progress indicators
Total = 0
SetProgress 0, 1, ""

'---------------------------------------------------
'Open the destination file in the temp folder
'---------------------------------------------------
mBackupFile = CheckPath(mTempFolder) & "_Backup.azp"
DestFile = mBackupFile

On Error Resume Next
Kill DestFile                   'Erase file to prevent carry over
On Error GoTo BackupZipFileErr
nfd = FreeFile
Open DestFile For Binary As #nfd

'---------------------------------------------------
'Notes on FDD usage
'---------------------------------------------------
'If a FDD is used, an event is raised indicating to
'the client that the disk needs to be changed. The
'event also indicates which file is required.
'Control should not be returned to this routine
'until a) the correct file has been located or
'      b) the user has cancelled the operation
'Notes
'a) The client must verify the file's existence
'b) The first FDD disk is probably already inserted
'   when the user read the file info. (But it could
'   have been removed in the mean time)
'---------------------------------------------------
'Determine if a disk needs to be changed/inserted
'---------------------------------------------------
StartLoop:
If IsFDD Then
  If Not (Drive.IsReady) Then
    SwapDisk = True
  Else
    SwapDisk = Not (mFso.FileExists(SourceFile))
  End If
  If SwapDisk Then
    'Ask to insert disk
    Screen.MousePointer = vbNormal
    RaiseEvent ChangeDisk(Drive, "Please insert disk with the file " & SourceFile & " into drive " & Drive.DriveLetter, SourceFile)
    Screen.MousePointer = vbArrowHourglass
    DoEvents
    'Allow for client canceling swap disk dialogue
    If mCancel Then
      mCancel = False         'Reset cancel property
      SetStatus azsReady      'Show green light
      Screen.MousePointer = vbNormal
      Exit Sub
    End If  'Cancel
  End If    'SwapDisk
End If      'IsFDD
  
'-----------------------------------------------------
'Open the source file (spanned volume)
'-----------------------------------------------------
nfs = FreeFile
Open SourceFile For Binary As #nfs

'Repeat this loop until all bytes have been read
nLeft = LOF(nfs) - Loc(nfs)    'Remaining bytes in source
Do
  '-----------------------------------------------------
  'Read a chunk of data
  '-----------------------------------------------------
  SetProgress Sum, Total, "Backing up zip file to hard disk..."
  If nLeft > CHUNK_SIZE Then
    nDest = CHUNK_SIZE
  Else
    nDest = nLeft
  End If
  ReDim Bytes(nDest) As Byte
  Get #nfs, , Bytes
  
  '-----------------------------------------------------
  'Write chunk of data to destination
  '-----------------------------------------------------
  Put #nfd, Loc(nfd) + 1, Bytes
  nLeft = LOF(nfs) - Loc(nfs)    'Remaining bytes in source
  Sum = Sum + nDest
Loop While nLeft > 0
  
'-------------------------------------------------------
'Close Files
'-------------------------------------------------------
Close #nfs
Close #nfd
SetProgress Total, Total, ""
Exit Sub

BackupZipFileErr:
Ans = ReportErrorAbort("BackupZipFile()", ModName, Err, Error)
Select Case Ans
Case vbCancel, vbAbort
  Reset       'Closes all files
  Exit Sub
Case vbRetry
  If IsFDD Then
    Close #nfs
    Resume StartLoop
  Else
    Resume
  End If
Case vbOK, vbIgnore
  Resume Next
End Select

End Sub

Property Let Cancel(vCancel As Boolean)
mCancel = vCancel

End Property

Property Get Cancel() As Boolean
Cancel = mCancel

End Property

Property Get Exist() As Boolean
'----------------------------------------------
'Check if a current zip file exists
'----------------------------------------------
Exist = mExist

End Property

Private Function ExtractFile(ByVal i As Long, Bytes() As Byte) As Boolean
'-----------------------------------------------------------------
'Extract a file (no i) from the backed up zip file
'Returns the data in a byte array starting with 1
'If no data is available, then ubound is set to 0
'Returns True if everything went Ok
'-----------------------------------------------------------------
Dim nfs                     'File Numbers
Dim SourceFile As String    'Complete name of source file, incl path & ext
Dim nSource                 'No of bytes to read from source
Dim ChkSum As Byte          'Original Checksum as calculated, for each file
Dim ChkSumZip As Byte       'Zipped checksum
Dim Result As ZLibErrors    'Zlib error code
Dim Ans, n

On Error GoTo ExtractFileErr
'---------------------------------------------------
'Open the source file
'---------------------------------------------------
SourceFile = mBackupFile
nfs = FreeFile
Open SourceFile For Binary As #nfs

'-----------------------------------------------------
'Unzip file
'-----------------------------------------------------
'Get the no of bytes to read for file(i)
nSource = aFiles(i).ZipSize
If nSource > 0 Then
  '-------------------------------------------------
  'Read data for file(i) to unzip
  '-------------------------------------------------
  ReDim Bytes(nSource) As Byte
  n = aFiles(i).Offset          'Get offset position
  Get #nfs, n + 1, Bytes        'Read data to buffer
      
  '-------------------------------------------------
  'Validate zip checksum
  '-------------------------------------------------
  ChkSumZip = CheckSum(Bytes)
  If ChkSumZip <> aFiles(i).ChkSumZip Then
    Err.Raise 703, "ArZip.UnzipFiles()", "Invalid zip checksum in file " & aFiles(i).Name
  End If
  '-------------------------------------------------
  'Decompress bytes using zlib.dll 1.1.3
  'The original size needs to be given in order to
  'make enough space (an extra 1% + 12 bytes is added
  'as a temporary measure, whereafter the buffer is
  'resized to the original size. The size parameter
  'is passed by value, so no need to protect it.
  '-------------------------------------------------
  Result = DecompressBytes(Bytes, aFiles(i).Size)
  If Result <> zlOk Then
    Select Case Result
    Case zlStreamEnd    '1
      Err.Raise 801, "ArZip.UnzipFiles()", "Error decompressing file " & aFiles(i).Name & " : Stream end"
    Case zlNeedDict     '2
      Err.Raise 802, "ArZip.UnzipFiles()", "Error decompressing file " & aFiles(i).Name & " : Dictionary required"
    Case zlErrNo        '-1
      Err.Raise 803, "ArZip.UnzipFiles()", "Error decompressing file " & aFiles(i).Name & " : Unknown error"
    Case zlStreamError  '-2
      Err.Raise 804, "ArZip.UnzipFiles()", "Error decompressing file " & aFiles(i).Name & " : Stream error"
    Case zlDataError    '-3
      Err.Raise 805, "ArZip.UnzipFiles()", "Error decompressing file " & aFiles(i).Name & " : Input data corrupted"
    Case zlMemError     '-4
      Err.Raise 806, "ArZip.UnzipFiles()", "Error decompressing file " & aFiles(i).Name & " : Insufficient memory"
    Case zlBufError     '-5
      Err.Raise 804, "ArZip.UnzipFiles()", "Error decompressing file " & aFiles(i).Name & " : Insufficient space in output buffer"
    Case zlVersionError '-6
      Err.Raise 805, "ArZip.UnzipFiles()", "Error decompressing file " & aFiles(i).Name & " : zlib version error"
    End Select
  End If
      
  '-------------------------------------------------
  'Validate checksum
  '-------------------------------------------------
  ChkSum = CheckSum(Bytes)
  If ChkSum <> aFiles(i).ChkSum Then
    Err.Raise 704, "ArZip.UnzipFiles()", "Invalid checksum in file " & aFiles(i).Name
  End If
Else
  ReDim Bytes(0 To 0) As Byte
End If
'-----------------------------------------------------
'Close Source (Zip) file
'-----------------------------------------------------
ExtractFile = True
Close #nfs
Exit Function

ExtractFileErr:
Ans = ReportErrorAbort("UnzipFiles()", ModName, Err, Error)
Select Case Ans
Case vbCancel, vbAbort
  Close #nfs
  ExtractFile = False
  Exit Function
Case vbRetry
  Resume
Case vbOK, vbIgnore
  Resume Next
End Select

End Function

Private Function FileInList(File As File) As Long
'---------------------------------------------------------------
'Checks if a file is already in the list
'Returns the position in the array, or 0 if it doesn't exist
'---------------------------------------------------------------
Dim Key As String
Dim Path As String        'Relative Path
Dim i

FileInList = 0
Path = RelPath(File.ParentFolder, mRootFolder)
Key = Path & "\" & File.Name
For i = 1 To nFiles
  If aFiles(i).Key = Key Then
    FileInList = i
    Exit For
  End If
Next

End Function

Public Sub RefreshFiles()
'-------------------------------------------
'Refresh files in archive
'Notes
'a) The file to be refreshed is validated for existence at source
'b) If the file exists, the last modified date/time is used to
'   determine if the file should be freshened (re-zipped)
'---------------------------------------------------------------
Dim i, Ok
Dim File As File

On Error Resume Next
Screen.MousePointer = vbArrowHourglass
SetStatus azsBusy
For i = 1 To nFiles
  Set File = mFso.GetFile(aFiles(i).Path & "\" & aFiles(i).Name)
  If ValidateFile(File) Then
    If File.DateLastModified > aFiles(i).Modified Then
      aFiles(i).Zipped = False
      aFiles(i).Modified = File.DateLastModified
      aFiles(i).Size = File.Size
      aFiles(i).IconKey = GetIconKey(File)
    End If
  End If
Next
mLvwFiles.SmallIcons = mIml       'Set the smallicons imagelist
Screen.MousePointer = vbNormal
SetStatus azsReady

End Sub

Private Sub RefreshSelection(ByVal AllFiles As Boolean)
'-----------------------------------------
'Refresh selected file list based on
'selected items in listview
'-----------------------------------------
Dim i, Item As ListItem

If AllFiles Then
  For i = 1 To nFiles
    aFiles(i).Selected = True
  Next
Else
  For Each Item In mLvwFiles.ListItems
    i = Val(Item.Tag)
    aFiles(i).Selected = Item.Selected
  Next
End If

End Sub
Public Sub RemoveFile(Index As Long)
'----------------------------------------------------------------
'Removes a file from the list
'----------------------------------------------------------------
'Notes
'a) The file is removed from the list at this stage
'b) When zipping (save button), the file is simply skipped when
'   copying the files in the current archive, effectively removing
'   it from the archive.
'----------------------------------------------------------------
Dim Key As String           'Icon Key
Dim UniqueKey As Boolean    'Is Icon Key unique?
Dim imgObj As ListImage           'Single image object in mIml.listimages collection
Dim Item As ListItem
Dim i, Ans

On Error GoTo RemoveFileErr
'Determine if IconKey is unique
UniqueKey = True
Key = aFiles(Index).IconKey
For i = 1 To nFiles
  If i <> Index And aFiles(i).IconKey = Key Then
    UniqueKey = False
    Exit For
  End If
Next
'If Key is unique, remove it from the icon image list
If UniqueKey Then
  mLvwFiles.SmallIcons = Nothing    'Remove the image list
  mIml.ListImages.Remove Key
  If mIml.ListImages.Count > 0 Then
    Set mLvwFiles.SmallIcons = mIml   'Set the smallicons imagelist
    nIcons = mIml.ListImages.Count
  Else
    nIcons = 0
  End If
End If
  
'Remove file from array
For i = Index To nFiles - 1
  aFiles(i) = aFiles(i + 1)
Next
nFiles = nFiles - 1
If nFiles > 0 Then
  ReDim Preserve aFiles(nFiles) As ArielFile
Else
  ReDim aFiles(0 To 0) As ArielFile
  mExist = False            'No further need to create a temp read file
End If

'Update tag references in lvw
mLvwFiles.Sorted = False
For i = Index To nFiles
  Set Item = mLvwFiles.ListItems(aFiles(i).Key)
  Item.Tag = CStr(i)
Next
mLvwFiles.Sorted = True

Exit Sub

RemoveFileErr:
Ans = ReportErrorAbort("ArZip.RemoveFile()", ModName, Err, Error)
Select Case Ans
Case vbCancel, vbAbort
  Exit Sub
Case vbRetry
  Resume
Case vbOK, vbIgnore
  Resume Next
End Select

End Sub

Private Function St2Byte(s As String) As Byte()
'-------------------------------------------------------
'Convert a string to byte array
'Not used but keep it for reference purposes
'-------------------------------------------------------
St2Byte = StrConv(s, vbFromUnicode)

End Function

Private Function Byte2St(Bytes() As Byte) As String
'-------------------------------------------------------
'Convert a byte array to string
'Not used but keep it for reference purposes
'-------------------------------------------------------
Dim St As String
St = Bytes
Byte2St = StrConv(St, vbUnicode)

End Function

Private Function CheckArchive() As Boolean
'--------------------------------------------------------
'Check the archive to determine if it is a normal
'or spanned file.
'Then check if Vol 1 is available
'--------------------------------------------------------
Dim Drive As Drive              'FSO Drive object (see scripting.dll)
Dim DriveSpec As String         'Required to get drive object
Dim SourceFile As String        'Name of ariel zip file to check
Dim IsFDD As Boolean            'Reading from a removable drive (Floppy)?
Dim nfs                         'File handle
Dim Hdr As ArHeader             'Archive Header Record
Dim Id As String * 3            'Archive file Id
Dim Ans                         'Error handler response

On Error GoTo CheckArchiveErr
SetStatus azsBusy

'Determine type of drive to read from
'This is necessary since if it is a removable drive (i.e. FDD)
'a dialog must be shown to switch disks
DriveSpec = mFso.GetDriveName(mFso.GetAbsolutePathName(mZipFile))
Set Drive = mFso.GetDrive(DriveSpec)
IsFDD = (Drive.DriveType = Removable)

'-------------------------------------------
'If FDD and drive is not ready, ask to
'insert disk
'-------------------------------------------
If IsFDD Then
  If Not (Drive.IsReady) Then
    'Ask to insert disk
    RaiseEvent ChangeDisk(Drive, "Please insert disk with file " & mZipFile & " into drive " & Drive.DriveLetter, mZipFile)
    'Allow for client to cancel current operation
    If mCancel Then
      mCancel = False         'Reset cancel property
      SetStatus azsReady      'Show green light
      CheckArchive = False    'Exit function
      Exit Function
    End If
  End If
End If
'-------------------------------------------
'Check if file exists
'-------------------------------------------
Do While Not (mFso.FileExists(mZipFile))
  '-------------------------------------------
  'File not found
  'If FDD ask to change disks
  '-------------------------------------------
  If IsFDD Then
    'Ask to insert disk
    RaiseEvent ChangeDisk(Drive, "File not found. Please insert disk with file " & mZipFile & " into drive " & Drive.DriveLetter, mZipFile)
    'Allow for client to cancel current operation
    If mCancel Then
      mCancel = False         'Reset cancel property
      SetStatus azsReady      'Show green light
      CheckArchive = False    'Exit function
      Exit Function
    End If
  Else
    MsgBox "The file " & mZipFile & " could not be found." & vbCrLf & _
           "Please select another file.", vbExclamation, "Error opening Ariel zip file"
    SetStatus azsReady
    CheckArchive = False
    Exit Function
  End If
Loop
'-------------------------------------------
'Open Ariel Archive File and read Id string
'-------------------------------------------
nfs = FreeFile
Open mZipFile For Binary As #nfs
Get #nfs, , Id            'Filetype ID: 'azp' No spanning, 'azs' Spanning
Select Case Id
Case "azp"                'Normal Archive
  mSpanning = False
Case "azs"                'Spanned Archive
  mSpanning = True
Case Else                 'Invalid file
  MsgBox "The file " & mZipFile & " is not a valid Ariel Zip file." & vbCrLf & _
         "The action will be cancelled.", vbExclamation, "Error opening Ariel zip file"
  Close #nfs
  SetStatus azsReady
  CheckArchive = False
  Exit Function
End Select

'-------------------------------------------
'Read Archive Header & check for Vol 1
'Vol 1 contains the file info & icons
'which is read by the ReadArchiveInfo() sub
'-------------------------------------------
Get #nfs, , Hdr
If Hdr.VolNo <> 1 Then
  MsgBox "The file " & mZipFile & " is volume " & Format(Hdr.VolNo, "000") & " of another Ariel Zip file." & vbCrLf & _
         "Please reselect another Ariel zip file.", vbExclamation, "Invalid Ariel zip file"
  Close #nfs
  SetStatus azsReady
  CheckArchive = False
  Exit Function
End If
'-----------------------------------------
'Ok, the correct Ariel zip file (Vol 1)
'has been found. Return Ok
'-----------------------------------------
CheckArchive = True
Close #nfs
Exit Function

CheckArchiveErr:
Ans = ReportErrorAbort("CheckArchive()", ModName, Err, Error)
Select Case Ans
Case vbCancel, vbAbort
  Reset       'Closes all files
  Exit Function
Case vbRetry
  Resume
Case vbOK, vbIgnore
  Resume Next
End Select

End Function

Private Function AddBytes(Bytes() As Byte, Temp() As Byte) As Long
'--------------------------------------------------------------------
'Add Temp() Bytes (1-based) to Bytes() (1-based)
'Return the ubound() posisiton
'--------------------------------------------------------------------
Dim ns, n

On Error Resume Next
ns = UBound(Bytes)
n = UBound(Temp)
If ns = 0 Then
  ReDim Bytes(n) As Byte      'Necessary, to get rid of 0 lower bound
Else
  ReDim Preserve Bytes(ns + n) As Byte
End If
Call CopyMem(Bytes(ns + 1), Temp(1), n)
AddBytes = UBound(Bytes)

End Function

Private Function CheckPath(ByVal Path As String) As String
'--------------------------------------------------
'Checks if path ends with "\". If not, add it.
'--------------------------------------------------
If Right(Path, 1) <> "\" Then
  CheckPath = Path & "\"
Else
  CheckPath = Path
End If

End Function

Public Sub Clear()
'--------------------------------------------------
'Clear all objects & erase temp files
'--------------------------------------------------
On Error Resume Next

nFiles = 0
ReDim aFiles(nFiles) As ArielFile
mLvwFiles.ListItems.Clear
mLvwFiles.SmallIcons = Nothing
mIml.ListImages.Clear
nIcons = 0
Kill mTempFolder & "\*.bmp"     'Delete all icons
Kill mTempFolder & "\*.azp"     'Delete all azp files
SetStatus azsEmpty

End Sub

Private Function CheckSum(Bytes() As Byte) As Byte
'-------------------------------------------------------
'Calculate a simple XOR checksum of the given byte array
'-------------------------------------------------------
Dim ChkSum As Byte
Dim i, n

n = UBound(Bytes)
For i = 1 To n
  ChkSum = ChkSum Xor Bytes(i)
Next
CheckSum = ChkSum

End Function

Private Function AddHeader(ByRef Bytes() As Byte, Hdr As ArHeader) As Long
'----------------------------------------------------------------------
'Add the file header to the byte array
'Each span volume has its own similar header
'Note that strings cannot be used in a udt structure so therefore, all
'strings have to be added 'externally'.
'At this stage, the NextVol, ZipSize & Header offset is not known, but
'space needs to be reserved for them.
'Return the ubound() of Bytes()
'----------------------------------------------------------------------
Dim ns

On Error Resume Next
'When zipping to a temp file, a normal archive is used
'A different routine is used when spanning, so all files
'created with this routine gets an azp Id
ns = AddString(Bytes, "azp")          'Bytes 1 - 3 File Id for Ariel archive (normal)
'Cannot use copymem to place contenst of Hdr in bytes directly
'so do it the long way
ns = AddInt(Bytes, Hdr.Major)
ns = AddInt(Bytes, Hdr.Minor)
ns = AddByte(Bytes, Hdr.VolNo)
ns = AddByte(Bytes, Hdr.NextVol)
ns = AddLong(Bytes, Hdr.ZipSize)
ns = AddLong(Bytes, Hdr.NoFiles)
ns = AddLong(Bytes, Hdr.NoIcons)
ns = AddLong(Bytes, Hdr.Offset)

AddHeader = UBound(Bytes)

End Function

Private Function CompressBytes(Data() As Byte, ByVal CompressionLevel As Long) As Long
'-----------------------------------------------------------------------------------------
'Compress the Bytes in data() using zlib 1.1.3
'Compression level 0-None, 1-Best Speed ... 9-Best Compression, or -1 Default (about level 6)
'-----------------------------------------------------------------------------------------
Dim SourceSize As Long, DestSize As Long
Dim Buffer() As Byte
Dim Result As Long, Ans

On Error GoTo CompressBytesErr

'Allocate memory for byte array
SourceSize = UBound(Data)
DestSize = SourceSize * 1.01 + 12      'As per definition, add 1% + 12 bytes
ReDim Buffer(DestSize) As Byte


'Compress byte array. Adjusts DestSize to contain the actual no of compressed bytes
Result = Compress(Buffer(1), DestSize, Data(1), SourceSize, CompressionLevel)

'Truncate data to compressed size
ReDim Preserve Data(DestSize) As Byte
CopyMem Data(1), Buffer(1), DestSize

'Cleanup
Erase Buffer

'Return error code (if any)
CompressBytes = Result
Exit Function

CompressBytesErr:
Ans = ReportErrorAbort("CompressBytes()", ModName, Err, Error)
Select Case Ans
Case vbCancel, vbAbort
  CompressBytes = Err
  Exit Function
Case vbRetry
  Resume
Case vbOK, vbIgnore
  Resume Next
End Select

End Function

Property Let CompressLevel(vCompressLevel As Integer)
If vCompressLevel < -1 Then
  mCompressLevel = -1
ElseIf vCompressLevel > 9 Then
  mCompressLevel = 9
Else
  mCompressLevel = vCompressLevel
End If

End Property

Property Get CompressLevel() As Integer
CompressLevel = mCompressLevel
End Property

Private Sub CreateFolders(ByVal Path As String)
'----------------------------------------------
'Create all folders belong in path, if they
'don't exist
'----------------------------------------------
Dim nSlash            'Position of slash
Dim Fldr As String    'Folder to check
Dim Folder As Folder

On Error Resume Next
nSlash = 1
Do While nSlash > 0
  nSlash = InStr(nSlash + 1, Path, "\")
  If nSlash > 0 Then
    Fldr = Left(Path, nSlash - 1)
    'Set Folder = mFso.GetFolder(Fldr)
    'Fldr = mFso.BuildPath(Folder.Path, Folder.Name)
    If Not (mFso.FolderExists(Fldr)) Then
      mFso.CreateFolder Fldr
    End If
  End If
Loop
If Not (mFso.FolderExists(Path)) Then
  mFso.CreateFolder Path
End If
    
End Sub

Private Function DecompressBytes(Data() As Byte, ByVal OrigSize As Long) As Long
'-----------------------------------------------------------------------------------------
'Compress the Bytes in data() using zlib 1.1.3
'-----------------------------------------------------------------------------------------
Dim SourceSize As Long, DestSize As Long
Dim Buffer() As Byte    'Allocate memory for buffers
Dim Result As Long, Ans

SourceSize = UBound(Data)             'Compressed Size
DestSize = OrigSize * 1.01 + 12       'Add 1% + 12 bytes to original size for appropriate buffer, as per specs
ReDim Buffer(DestSize) As Byte

'Decompress data
'Usage: Result = Uncompress(Dest, DestLen, Source, SourceLen)
'Upon exit DestSize contains the actual unzipped length
Result = Uncompress(Buffer(1), DestSize, Data(1), SourceSize)

'Truncate buffer to compressed size
ReDim Preserve Data(DestSize) As Byte
CopyMem Data(1), Buffer(1), DestSize

'Return error code (if any)
DecompressBytes = Result
Exit Function

CompressBytesErr:
Ans = ReportErrorAbort("DecompressBytes()", ModName, Err, Error)
Select Case Ans
Case vbCancel, vbAbort
  DecompressBytes = Err
  Exit Function
Case vbRetry
  Resume
Case vbOK, vbIgnore
  Resume Next
End Select

End Function

Property Get ElapsedTime() As Single
ElapsedTime = mElapsedTime

End Property

Private Sub EraseDrive(Drive As Drive)
'---------------------------------------------
'Erases all files & folders on disk
'---------------------------------------------
On Error Resume Next
mFso.DeleteFolder Drive.Rootfolder & "*.*", True
mFso.DeleteFile Drive.Rootfolder & "*.*", True
DoEvents

End Sub

Public Property Get Ext(Index As Long) As String
'------------------------------------------------
'Returns the file extension description
'e.g. 'txt' for .txt
'------------------------------------------------
Ext = aFiles(Index).Ext
End Property

Private Function FileExtension(ByVal File As String) As String
Dim Ext As String, n
n = InStrRev(File, ".")
If n > 0 Then
  FileExtension = LCase(Mid(File, n + 1))
Else
  FileExtension = "None"
End If

End Function

Private Function GetIconKey(File As File) As String
'---------------------------------------------------------------------
'Extract an individual small icon associated with the given file
'Scan the image list for the extension key
'If it exists in the image list already, return the existing key
'If not, add the icon and key to the image list
'Returns the IconKey in the image list
'---------------------------------------------------------------------
Dim hSmallIcon As Long        'Handle to small Icon
Dim imgObj As ListImage       'Single image object in imagelist.listimages collection
Dim Key As String
Dim Ret As Long               'Return value
Dim FileName As String
Dim Ext As String

On Error Resume Next
'Get a handle to the small icon
FileName = File.Path
Ext = FileExtension(File.Name)
If InStr("exe|pif|lnk|ico|cur", Ext) > 0 Then
  Key = LCase(File.Name)
Else
  Key = Ext
End If
hSmallIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
'If the handle exists, load it into the picture box(es)
If hSmallIcon <> 0 Then
  'Small Icon
  With mPic
    Set .Picture = LoadPicture("")
    Ret = ImageList_Draw(hSmallIcon, ShInfo.iIcon, mPic.hDC, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With
  'Try to access the image in the image list, if it doesn't exist then
  'an error is returned
  Set imgObj = mIml.ListImages(Key)
  If Err <> 0 Then
    'Add the image if it doesn't exist.
    Err = 0
    Set imgObj = mIml.ListImages.Add(, Key, mPic.Image)
    'Save to temp file
    SavePicture mPic.Image, mTempFolder & "\" & Key & ".bmp"
    nIcons = nIcons + 1
  End If
Else
  'Try to access the image in the image list, if it doesn't exist then
  'an error is returned
  Set imgObj = mIml.ListImages(Key)
  If Err <> 0 Then
    'Add the image if it doesn't exist.
    Err = 0
    Set imgObj = mIml.ListImages.Add(, Key, mPicDef.Image)
    'Save to temp file
    SavePicture mPicDef.Image, mTempFolder & "\" & Key & ".bmp"
    nIcons = nIcons + 1
  End If
End If
GetIconKey = Key

End Function

Private Function GetVolumeSize(Optional ByVal SpaceAvail As Long = 1457644) As Long
'---------------------------------------------------------------------------
'Determine the chunksize to write to during diskspanning
'This is determined by a) the SpanOption selected and
'b) for removable drives, the available space
'By default, if writing span disks to a HDD, the max capacity of a
'volume is 1457644 bytes
'The 1.44Mb for FDD's is actually a misrepresentation as the available
'space for an empty formatted FDD is about 1.39Mb or 1,457,644 bytes
'Therefore, scale down the other options appropriately
'---------------------------------------------------------------------------
Dim Size

Select Case mSpanOption
Case azo1440      '1.44Mb
  Size = 1457644  'Actually 1.39 Mb (as reported by empty FDD's)
Case azo1400      '1.40Mb
  Size = 1457644 * 1.4 / 1.44
Case azo1200      '1.20Mb
  Size = 1457644 * 1.2 / 1.44
Case azo1000      '1.00Mb
  Size = 1048576  '1024x1024 bytes
Case azo720       '720kb
  Size = 730112   'Actually 713kb
Case azo700       '700kb
  Size = 716800
Case azo100p      '100% Capacity (Removable drives only)
  Size = SpaceAvail
Case azo99p       '99% Capacity     "
  Size = SpaceAvail * 0.99
Case azo98p       '98% Capacity     "
  Size = SpaceAvail * 0.98
Case azo95p       '95% Capacity     "
  Size = SpaceAvail * 0.95
Case azo90p       '90% Capacity     "
  Size = SpaceAvail * 0.9
Case Else         'Default -1
  mSpanOption = azo1440
  Size = 1457644  'Actually 1.39 Mb (as reported by empty FDD's)
End Select
'Set span size
mSpanSize = Size
If Size > SpaceAvail Then
  Size = SpaceAvail
End If
GetVolumeSize = Size

End Function

Private Function GetTempFolder() As String
'----------------------------------------------------------
'Get the system temporary folder
'----------------------------------------------------------
Dim lngReturnVal As Long
Dim strTempPath As String * 255

On Error Resume Next
lngReturnVal = GetTempPath(255, strTempPath)

GetTempFolder = StrConv(Left(strTempPath, InStr(strTempPath, Chr(0)) - 1), vbProperCase)

End Function

Public Property Get IconKey(Index As Long) As String
'------------------------------------------------
'Returns the icon key in the image list associated
'with the file
'------------------------------------------------
IconKey = aFiles(Index).IconKey
End Property

Public Sub Initialise(lvwFiles As ListView, imlSmall As ImageList, picSmall As PictureBox, picDefault As PictureBox)
'-----------------------------------------------
'Initialise the controls
'lvwFile  : Contains the files
'imlSmall : Contains the small 16x16 icons
'picSmall : Temporary holder for small icons
'-----------------------------------------------
On Local Error Resume Next

Set mLvwFiles = lvwFiles
Set mIml = imlSmall
Set mPic = picSmall

'Initialise picture box
mPic.Width = (SMALL_ICON) * Screen.TwipsPerPixelX
mPic.Height = (SMALL_ICON) * Screen.TwipsPerPixelY
mPic.AutoRedraw = True

'Initialise Default Picture
Set mPicDef = picDefault

'Break the link to iml lists
mLvwFiles.ListItems.Clear
mLvwFiles.Icons = Nothing
mLvwFiles.SmallIcons = Nothing

'Clear the image lists
mIml.ListImages.Clear

End Sub

Public Property Get Key(Index As Long) As String
'------------------------------------------------
'Returns the file key (path\filename)
'------------------------------------------------
Key = aFiles(Index).Key
End Property

Private Function Maximum(t1 As Variant, ParamArray t() As Variant) As Variant
'----------------------------------------------------
'Determine the maximum of all values
'----------------------------------------------------
Dim x As Variant, i As Long

x = t1
For i = 0 To UBound(t)
  If t(i) > x Then
    x = t(i)
  End If
Next
Maximum = x

End Function

Private Function Minimum(t1 As Variant, ParamArray t() As Variant) As Variant
'----------------------------------------------------
'Determine the minimum of all values
'----------------------------------------------------
Dim x As Variant, i As Long

x = t1
For i = 0 To UBound(t)
  If t(i) < x Then
    x = t(i)
  End If
Next
Minimum = x

End Function

Public Property Get Modified(Index As Long) As Date
'------------------------------------------------
'Returns the last modification date of the file
'------------------------------------------------
Modified = aFiles(Index).Modified
End Property

Public Property Get Name(Index As Long) As String
'------------------------------------------------
'Returns the file name (incl extension), excl path
'------------------------------------------------
Name = aFiles(Index).Name
End Property

Property Get NoFiles() As Long
'----------------------------------------
'Returns the number of files in the list
'----------------------------------------
NoFiles = nFiles

End Property

Public Function OpenArchive(sZipFile As String) As Long
'------------------------------------------------
'Opens the archive and reads the file/icon lists
'sZipFile   : Zipfile to open (incl path & ext)
'Return     : NoFiles
'------------------------------------------------
Dim Ok

On Error Resume Next
Screen.MousePointer = vbArrowHourglass
mZipFile = sZipFile
Clear
SetStatus azsBusy
Ok = CheckArchive
If Ok Then
  Ok = ReadArchiveInfo
End If
If Ok Then
  mLvwFiles.SmallIcons = mIml     'Set the smallicons imagelist
  OpenArchive = nFiles            'Set the return parameter
Else
  OpenArchive = 0
End If
Screen.MousePointer = vbNormal
SetStatus azsCreated
End Function

Public Property Get Path(Index As Long) As String
'------------------------------------------------
'Returns the folderpath of the specified file
'------------------------------------------------
Path = aFiles(Index).Path
End Property
Public Property Get RelativePath(Index As Long) As String
'-----------------------------------------------------------------
'Returns the relative folderpath of the specified file
'-----------------------------------------------------------------
RelativePath = aFiles(Index).RelPath

End Property

Private Function Proper(txt As Variant) As String
'-------------------------------------------------------
'Returns a proper string with first letters capitalized
'Capitalises 1st letter after space or start
'Not used here - keep for record purposes
'see also strconv function
'-------------------------------------------------------
Dim n, i, s As String, bSp As Boolean, c As String

s = ""
txt = Trim(txt)
n = Len(txt)
bSp = True
For i = 1 To n
  c = Mid(txt, i, 1)
  If InStr(" `~!@#$%^&*()_+|-=\[]{};':<>,/?", c) > 0 Then
    bSp = True
  Else
    If bSp Then
      c = UCase(c)
      bSp = False
    Else
      c = LCase(c)
    End If
  End If
  s = s + c
Next
Proper = s

End Function

Public Property Get Ratio(Index As Long) As Single
'------------------------------------------------
'Returns the zipped/unzipped ratio in a fraction
'------------------------------------------------
Ratio = aFiles(Index).Ratio

End Property

Private Function ReadArchiveInfo() As Boolean
'---------------------------------------------------------------------------
'Reads all file & icon info from the azp or azs archive
'Notes:
'a) Variable length strings are used to reduce space
'   As the get statement cannot be used for these, the input
'   statement must be used. Also, a CR character need to
'   be added to identify the end of the string.
'b) Refer to the ZipFiles() sub for the archive file structure
'c) It is assumed that all of the archive info are contained in the
'   first volume. This is valid as the total info to read is normally
'   very small, from 2Kb to 100Kb
'---------------------------------------------------------------------------
Dim nfs, nfd                      'File Numbers
Dim SourceFile As String          'Name of source file, incl path
Dim DestFile As String            'Name of dest file, incl path (for icon bitmaps)
Dim Bytes() As Byte               'Main source buffer
Dim Buffer() As Byte              'Temporary buffer
Dim Hdr As ArHeader               'Archive header info
Dim Afr As ArFileRec              'Ariel file record
Dim imgObj As ListImage           'Single image object in mIml.listimages collection
Dim Key As String                 'Key for mIml.listimage
Dim Id As String * 3              'Ariel Zip file Id
Dim Id7 As String * 7             'Position tracking in zip file
Dim i, n
Dim Ans

On Error GoTo ReadArchiveInfoErr
SetStatus azsBusy
Screen.MousePointer = vbArrowHourglass
SetProgress 0, 1, "Reading archive info..."
'---------------------------------------------------
'Open the source file
'---------------------------------------------------
nfs = FreeFile
SourceFile = mZipFile
mTotalZipSize = FileLen(SourceFile)   'Must correspond with Hdr.Zipsize
Open SourceFile For Binary As #nfs

'---------------------------------------------------
'Process Header
'---------------------------------------------------
Get #nfs, , Id
If Id <> "azp" And Id <> "azs" Then
  Err.Raise 601, "ArZip.ReadArchiveInfo()", "Invalid Ariel Zip File Id"
End If
Get #nfs, , Hdr               'Contains Version, vol info, zip file size, no files, no icons etc.
nFiles = Hdr.NoFiles          'No of files
nIcons = Hdr.NoIcons          'No of icons (16x16 bitmaps)
mSpanned = (Id = "azs")       'Has this file been spanned?
Input #nfs, mRootFolder       'Used as default folder for extraction
'Note: Use Input statement for variable strings (delimited with vbCr)
ReDim aFiles(nFiles) As ArielFile

'-----------------------------------------------
'File Info
'-----------------------------------------------
Get #nfs, , Id7               '[Files]
If Id7 <> "[Files]" Then
  Err.Raise 602, "ArZip.ReadArchiveInfo()", "Can't locate file list in Ariel zip archive"
End If
For i = 1 To nFiles
  'Read file info (numeric & date)
  Get #nfs, , Afr
  aFiles(i).Offset = Afr.Offset
  aFiles(i).Zipped = True
  aFiles(i).Size = Afr.OrigSize
  aFiles(i).ZipSize = Afr.ZipSize
  aFiles(i).Modified = Afr.Modified
  aFiles(i).ChkSum = Afr.ChkSum
  aFiles(i).ChkSumZip = Afr.ChkSumZip
  If aFiles(i).Size > 0 Then
    aFiles(i).Ratio = aFiles(i).ZipSize / aFiles(i).Size
  Else
    aFiles(i).Ratio = 1
  End If
Next

'-----------------------------------------------
'File Names
'-----------------------------------------------
Get #nfs, , Id7               '[Names]
If Id7 <> "[Names]" Then
  Err.Raise 603, "ArZip.ReadArchiveInfo()", "Can't locate file names in Ariel zip archive"
End If
For i = 1 To nFiles
  'Read file info (string info).
  'Use the input function for variable strings (delimited with vbCr)
  Input #nfs, aFiles(i).Name            'File name
  Input #nfs, aFiles(i).RelPath         'Relative path (to rootfolder)
  Input #nfs, aFiles(i).Path            'Absolute path
  Input #nfs, aFiles(i).IconKey         'Icon key in image list
  'Compile file key (uniquely identifies each file for use in client listview object)
  aFiles(i).Key = aFiles(i).RelPath & "\" & aFiles(i).Name
Next

'-----------------------------------------------
'Icons
'-----------------------------------------------
Get #nfs, , Id7               '[Icons]
If Id7 <> "[Icons]" Then
  Err.Raise 604, "ArZip.ReadArchiveInfo()", "Can't locate the icons in Ariel zip archive"
End If
ReDim Buffer(BMP_SIZE) As Byte
'Delete all bmp files in the temp folder
On Error Resume Next
Kill CheckPath(mTempFolder) & "*.bmp"
On Error GoTo ReadArchiveInfoErr
For i = 1 To nIcons
  Input #nfs, Key                       'Icon Key
  Get #nfs, , Buffer                    'Icon 16x16 Bitmap (822 bytes)
  'Save the bitmap to a temp file
  nfd = FreeFile
  DestFile = CheckPath(mTempFolder) & Key & ".bmp"
  Open DestFile For Binary As #nfd
  Put #nfd, , Buffer
  Close #nfd
  'Add the bitmap picture to the image list, using the Key as identifier
  'This list is identified in the Initialise() sub
  Set imgObj = mIml.ListImages.Add(, Key, LoadPicture(DestFile))
Next

'-----------------------------------------------------
'Close Source (Zip) file
'-----------------------------------------------------
Close #nfs
mExist = True             'Set current zip file flag
mZipFileNew = mZipFile    'Set the new zip file to the current one
ReadArchiveInfo = True
Screen.MousePointer = vbNormal
Exit Function

ReadArchiveInfoErr:
Ans = ReportErrorAbort("ReadArchiveInfo()", ModName, Err, Error)
Select Case Ans
Case vbCancel, vbAbort
  Reset             'Closes all files
  ReadArchiveInfo = False
  mExist = False             'Set current zip file flag
  Exit Function
Case vbRetry
  Resume
Case vbOK, vbIgnore
  Resume Next
End Select

End Function

Public Sub UnregisterArielFileTypes()
'--------------------------------------------------------
'Unregisters the '.azp/.azs' extensions
'This requires the regobj.dll in the system folder
'Also, add a reference to 'Registration Manipulation Classes'
'(the regobj.dll) in this project. See the References item
'on the Project menu in the VB IDE.
'--------------------------------------------------------
Dim objRootKey As RegKey     'See objreg.dll documentation
  
On Error Resume Next
'Get the key from its string reference
Set objRootKey = RegKeyFromHKey(HKEY_CLASSES_ROOT)
'Remove the azp and azs keys (and all subkeys)
objRootKey.SubKeys.Remove ".azp"
objRootKey.SubKeys.Remove ".azs"

Set objRootKey = Nothing

End Sub

Private Sub UnspanZipfile()
'-------------------------------------------------------------
'Build the Ariel Zip Archive from its spanned components
'The rebuilt file is saved in the tempfolder
'-------------------------------------------------------------
Dim Drive As Drive          'FSO Drive object
Dim DriveSpec As String     'Required to resolve the drive object
Dim IsFDD As Boolean        'Is destination drive removable?
Dim nfd, nfs                'File Numbers
Dim SourceFile As String    'Name of sourcefile, incl path
Dim SourceBase As String    'Dest (zip) file name component, excl path, vol & ext
Dim SourceFolder As String  'Destination folder, excluding last "\"
Dim nSource As Long         'Length of source spanvolume, bytes
Dim DestFile As String      'Complete name of destination file, incl path & ext
Dim nDest As Long           'Length of destination file to write
Dim Bytes() As Byte         'Bytes array
Dim Id As String * 3        'Archive Id ("azs" or "azp")
Dim Hdr As ArHeader         'Archive header
Dim VolNo As Byte           'Span (Volume) No to read, appended to zipfile name
Dim NextVol As Byte         'Control flag for volumes to read
Dim SwapDisk As Boolean     'Flag indicating that a disk needs to be inserted/changed
Dim Sum, Total              'Progress indicator
Dim Ans

On Error GoTo UnspanZipfileErr
'--------------------------------------------------
'Determine type of drive to read from
'This is necessary since if it is a removable drive
'(i.e. FDD), a dialog must be shown to switch disks
'--------------------------------------------------
DriveSpec = mFso.GetDriveName(mFso.GetAbsolutePathName(mZipFile))
Set Drive = mFso.GetDrive(DriveSpec)
IsFDD = (Drive.DriveType = Removable)

'---------------------------------------------------
'Determine source folder & core source file name
'---------------------------------------------------
If IsFDD Then
  SourceFolder = DriveSpec
Else
  SourceFolder = mFso.GetParentFolderName(mZipFile)
End If
'Core name of zipfile e.g. "test" for test.azp
SourceBase = mFso.GetBaseName(mZipFile)

'---------------------------------------------------
'Initialise other variables
'---------------------------------------------------
VolNo = 1         'Volume No to read, starting with 1
Sum = 0           'Progrss indicators
Total = 0
SetProgress 0, 1, ""

'---------------------------------------------------
'Open the destination file in the temp folder
'---------------------------------------------------
DestFile = CheckPath(mTempFolder) & SourceBase & ".azp"
On Error Resume Next
Kill DestFile                   'Erase file to prevent carry over
On Error GoTo UnspanZipfileErr
nfd = FreeFile
Open DestFile For Binary As #nfd

StartLoop:
'Repeat this loop until all bytes have been read
Do
  '---------------------------------------------------
  'Notes on FDD usage
  '---------------------------------------------------
  'If a FDD is used, an event is raised indicating to
  'the client that the disk needs to be changed. The
  'event also indicates which file is required.
  'Control should not be returned to this routine
  'until a) the correct file has been located or
  '      b) the user has cancelled the operation
  'Notes
  'a) The client must verify the file's existence
  'b) The first FDD disk is probably already inserted
  '   when the user read the file info. (But it could
  '   have been removed in the mean time)
  '---------------------------------------------------
  'Determine which source file to open
  '---------------------------------------------------
  If VolNo <= 1 Then
    SourceFile = CheckPath(SourceFolder) & SourceBase & ".azp"
  Else
    SourceFile = CheckPath(SourceFolder) & SourceBase & "-" & Format(VolNo, "000") & ".azs"
  End If
  '---------------------------------------------------
  'Determine if a disk needs to be changed/inserted
  '---------------------------------------------------
  If IsFDD Then
    If Not (Drive.IsReady) Then
      SwapDisk = True
    Else
      SwapDisk = Not (mFso.FileExists(SourceFile))
    End If
    If SwapDisk Then
      'Ask to insert disk
      Screen.MousePointer = vbNormal
      RaiseEvent ChangeDisk(Drive, "Please insert disk #" & Format(VolNo, "##0") & " into drive " & Drive.DriveLetter, SourceFile)
      Screen.MousePointer = vbArrowHourglass
      DoEvents
      'Allow for client canceling swap disk dialogue
      If mCancel Then
        mCancel = False         'Reset cancel property
        SetStatus azsReady      'Show green light
        Screen.MousePointer = vbNormal
        Exit Sub
      End If  'Cancel
    End If    'SwapDisk
  End If      'IsFDD
  
  '-----------------------------------------------------
  'Open the source file (spanned volume)
  '-----------------------------------------------------
  nfs = FreeFile
  Open SourceFile For Binary As #nfs

  '-----------------------------------------------------
  'Read header & place in dest file (only Vol 1)
  '-----------------------------------------------------
  Get #nfs, 1, Id
  If Id <> "azs" Then
    Err.Raise 901, "ArZip.UnspanZipfile()", "Invalid Ariel Zip File Id"
  End If
  Get #nfs, 4, Hdr
  If VolNo <> Hdr.VolNo Then
    Err.Raise 902, "ArZip.UnspanZipfile()", "Volume No doesn't match - looking for vol " & Format(VolNo, "000") & "."
  End If
  NextVol = Hdr.NextVol       'Save for loop control
  Hdr.VolNo = 1               'Change volumes for unspanned zip file
  Hdr.NextVol = 0             'Change next vol flag "        "    "
  If VolNo = 1 Then
    Put #nfd, 1, "azp"
    Put #nfd, 4, Hdr
  End If
  
  '-----------------------------------------------------
  'Read data in one go
  '-----------------------------------------------------
  SetProgress Sum, Total, "Unspanning volume " & Format(VolNo, "000") & " to hard disk..."
  nDest = LOF(nfs) - Loc(nfs)    'Remaining bytes in source
  ReDim Bytes(nDest) As Byte
  Get #nfs, , Bytes
  Close #nfs
  
  '-----------------------------------------------------
  'Write data in one go
  '-----------------------------------------------------
  Put #nfd, Loc(nfd) + 1, Bytes
  Sum = Sum + nDest
  VolNo = VolNo + 1
  If NextVol = 0 Then
    Exit Do
  End If
Loop
  
'-------------------------------------------------------
'Close Destination (temp zip file)
'-------------------------------------------------------
Close #nfd
'To prevent re-reading of FDD's etc, change the reference to the
'current zipfile to the unspanned volume in the tempfolder
mZipFile = DestFile
mSpanned = False
SetProgress Total, Total, ""
Exit Sub

UnspanZipfileErr:
Ans = ReportErrorAbort("UnspanZipfile()", ModName, Err, Error)
Select Case Ans
Case vbCancel, vbAbort
  Reset       'Closes all files
  Exit Sub
Case vbRetry
  If IsFDD Then
    Resume StartLoop
  Else
    Resume
  End If
Case vbOK, vbIgnore
  Resume Next
End Select

End Sub
Private Function RelPath(ByVal FullPath As String, ByVal Root As String) As String
'---------------------------------------------------------------------
'Strips the relative path from the full path, given the root
'Assumes fullpath always CONTAINS the root
'Returns "" if fullpath=root or "SubFolder\SubSubFolder"
'---------------------------------------------------------------------
Dim n
Dim Full As String, Rt As String
Rt = LCase(Root)
Full = LCase(FullPath)
n = InStr(Full, Rt)
If n > 0 Then
  RelPath = Mid(FullPath, Len(Root) + 1)
Else
  RelPath = ""
End If

End Function

Property Get Rootfolder() As String
'---------------------------------------------
'Returns the root folder from which the
'list was built
'---------------------------------------------
Rootfolder = mRootFolder

End Property

Property Let Rootfolder(vRootFolder As String)
'---------------------------------------------
'Sets the rootfolder from which the list is
'to be built
'---------------------------------------------
mRootFolder = vRootFolder

End Property

Public Function NewArchive(ZipFile As String, Rootfolder As String, IncludeFiles As Boolean, IncludeSubfolders As Boolean) As Long
'---------------------------------------------------------------------------------------------------------------
'Open a new archive
'ZipFile            : Name & path of zip archive
'Rootfolder         : Path of folder to use as rootfolder reference (for relative path)
'IncludeFiles       : Include all files in the rootfolder in the archive
'IncludeSubfolders  : Include subfolders' contents in the archive
'Return             : NoFiles added
'---------------------------------------------------------------------------------------------------------------
'a) Clears current list
'b) Populates zipfile & rootfolder properties
'c) Adds all files in the root & subfolders (if true) to the filelist
'd) Updates the mExist flag
'---------------------------------------------------------------------------------------------------------------

On Error Resume Next
Screen.MousePointer = vbArrowHourglass
Clear
SetStatus azsBusy
mZipFileNew = ZipFile
mRootFolder = Rootfolder
mExist = False
If IncludeFiles Then
  ScanTree mRootFolder, IncludeSubfolders         'Recursive tree scan
End If
mLvwFiles.SmallIcons = mIml       'Set the smallicons imagelist
NewArchive = nFiles               'Set the return parameter
Screen.MousePointer = vbNormal
SetStatus azsReady

End Function

Private Sub SetProgress(ByVal Sum As Long, ByVal Total As Long, ByVal Info As String)
'----------------------------------------------------
'Calculates the progress as a fraction 0 to 1
'The user should convert it to a % value
'----------------------------------------------------
If Total <= 0 Then
  mProgress = 1
Else
  mProgress = Sum / Total
  If mProgress < 0 Then mProgress = 0
  If mProgress > 1 Then mProgress = 1
End If
RaiseEvent Progress(mProgress, Info)
DoEvents

End Sub

Private Sub ScanTree(ByVal Path As String, ByVal IncludeSubfolders As Boolean)
'-----------------------------------------------------------------
'Recursive subroutine to add subfolders & files to the list of
'folders
'Path     : The path to scan for subfolders & files
'Called by: ScanFiles()
'-----------------------------------------------------------------
'Method
'a) The folder indicated in the Path parameter becomes the
'   current root folder.
'b) Add all files in the root folder to the aFile() array
'c) Do a recursive ScanTree call to scan the subfolders
'-----------------------------------------------------------------
'Credits  : Brad V, Planet Source Code
'-----------------------------------------------------------------
Dim Root As Folder                'Fso Folder object
Dim SubFolder As Folder           'Fso Folder object
Dim File As File                  'Fso File object
Dim sPath As String               'Current folder ending with '\'
Dim Ok

'Set the root foler
sPath = CheckPath(Path)           'Add '\'
Set Root = mFso.GetFolder(sPath)
'Scan the files in the root folder
For Each File In Root.Files
  Ok = AddFile(File)
Next
If IncludeSubfolders Then
  'Do a recursive scan on each subfolder
  For Each SubFolder In Root.SubFolders
    Call ScanTree(sPath & SubFolder.Name, True)
  Next
End If
'Clean up
Set Root = Nothing
Set SubFolder = Nothing
Set File = Nothing

End Sub

Private Sub SetStatus(NewStatus As ArZipStatus)
'----------------------------------------------
'Set/Change the zip status
'----------------------------------------------
mStatus = NewStatus
RaiseEvent StatusChange(mStatus)
DoEvents

End Sub

Private Sub SpanZipFile()
'----------------------------------------------------------------------
'Split the tempzip file into the desired portions
'and copy with user interaction
'The first portion carries the zipfile name and is vol 1 (e.g. myzip.azp)
'The second portion is vol 2 and is identified as myzip-002.azs
'The last portion is vol n and is identified as   myzip-00n.azs
'----------------------------------------------------------------------
Dim Drive As Drive        'FSO Drive object
Dim DriveSpec As String   'Required to resolve the drive object
Dim IsFDD As Boolean      'Writing to a removable drive (Floppy)?
Dim nfd, nfs              'File Numbers
Dim nb                    'No of bytes in Bytes() Bytes
Dim SourceFile As String  'Name of sourcefile, incl path
Dim nSource As Long       'Length of source (zip) file, bytes
Dim nLeft As Long         'Remaining portion of zip file to write
Dim DestFile As String    'Complete name of destination file, incl path & ext
Dim DestFiles As String   'Series of destination files
Dim DestBase As String    'Dest (zip) file name component, excl path, vol & ext
Dim DestFolder As String  'Destination folder, excluding last "\"
Dim nDest As Long         'Length of destination file to write
Dim Bytes() As Byte
Dim Hdr As ArHeader       'Archive header
Dim nHdr As Long          'Size of Volume Header (len(Hdr)+3 bytes for azs Id)
Dim VolNo As Byte         'Span (Volume) No to write, appended to zipfile name
Dim Sum, Total            'Progress indicator
Dim DestReady As Boolean  'Flag indicating that everything is ready to write to dest
Dim DestOpen As Boolean   'Flag indicating that dest file is already open
Dim Ans, ErrTxt As String

On Error GoTo SpanZipFileErr
'--------------------------------------------------
'Determine type of drive to write to
'This is necessary since if it is a removable drive
'(i.e. FDD), a dialog must be shown to switch disks
'--------------------------------------------------
DriveSpec = mFso.GetDriveName(mFso.GetAbsolutePathName(mZipFileNew))
Set Drive = mFso.GetDrive(DriveSpec)
IsFDD = (Drive.DriveType = Removable)

'---------------------------------------------------
'Initialise other variables
'---------------------------------------------------
nHdr = Len(Hdr) + 3   'Size of header, bytes (incl space for 'azs' id)
Sum = 0               'Total bytes processed
Total = 0             'Total bytes to process
VolNo = 1             'Number of current volume to write

'---------------------------------------------------
'Determine destination folder
'---------------------------------------------------
If IsFDD Then
  DestFolder = Drive.Path
Else
  DestFolder = mFso.GetParentFolderName(mZipFileNew)
End If
DestBase = mFso.GetBaseName(mZipFileNew)       'Core name of zipfile e.g. "test" for test.azp
DestFiles = CheckPath(DestFolder) & DestBase & "-*.azs"

'---------------------------------------------------
'Open source file
'---------------------------------------------------
nfs = FreeFile
SourceFile = mTempFile
Open SourceFile For Binary As #nfs
nSource = LOF(nfs)          'Total length to write
nLeft = nSource             'Portion remaining
Get #nfs, 4, Hdr            'Read header
Total = nSource             'Update progress target

'-----------------------------------------------------
'Main loop for spanning volume(s)
'-----------------------------------------------------
Do While nLeft > 0
  '---------------------------------------------------
  'Determine which destination file to open
  '---------------------------------------------------
  If VolNo <= 1 Then
    DestFile = CheckPath(DestFolder) & DestBase & ".azp"
  Else
    DestFile = CheckPath(DestFolder) & DestBase & "-" & Format(VolNo, "000") & ".azs"
  End If
  
  '---------------------------------------------------
  'Notes on writing archives to Floppy Disk
  '---------------------------------------------------
  'If a FDD is used, an event is raised indicating to
  'the client that the disk needs to be changed.
  'Control should not be returned to this routine
  'until a) the disk is ready, or
  '      b) the user has cancelled the operation
  'Notes
  'a) All files & folders will be erased
  'b) The first FDD disk is probably already inserted
  '   when the user read the file info, but if the
  '   user had typed the dest name, rather than
  '   selecting it, the drive is probably empty
  'c) For the first volume, always ask for a disk
  '   (to prevent accidental overwriting of data)
  '---------------------------------------------------
  If IsFDD Then
    'Ask to insert disk
    Screen.MousePointer = vbNormal
    RaiseEvent ChangeDisk(Drive, "Please insert disk #" & Format(VolNo, "##0") & " into drive " & Drive.DriveLetter, "")
    Screen.MousePointer = vbArrowHourglass
    DoEvents
    'Allow for client canceling swap disk dialogue
    If mCancel Then
      mCancel = False         'Reset cancel property
      Close #nfs              'Close source file
      SetStatus azsReady      'Show green light
      Exit Sub
    End If
    EraseDrive Drive          'Delete all files & folders on the disk

    '------------------------------------------------------------------------
    'Determine the volume size to write to. This is the minimum of:
    'a) The specified spansize
    'b) Disk capacity, and
    'c) Remaining portion of source file to write
    'If the specified size (or available capacity) is to be utilised in
    'full, subtract the header size to determine the amount of data
    'that can be handled
    '------------------------------------------------------------------------
    nDest = Minimum(GetVolumeSize(Drive.AvailableSpace) - nHdr, nLeft)
  Else
    'Kill any file with the same name as destination file
    'to prevent a left-over included in the archive
    If VolNo <= 1 Then
      On Error Resume Next
      Kill DestFile
      Kill DestFiles
      On Error GoTo SpanZipFileErr
    End If
    nDest = Minimum(GetVolumeSize - nHdr, nLeft)
  End If
  
  '------------------------------------------------------
  'Read the required amt of data from the source
  '------------------------------------------------------
  ReDim Bytes(nDest) As Byte
  Get #nfs, , Bytes
  nLeft = nLeft - nDest
  'Make space for header
  ReDim Preserve Bytes(nDest + nHdr) As Byte
  CopyMem Bytes(nHdr + 1), Bytes(1), nDest
  
  '------------------------------------------------------
  'Add the header to the Byte() array
  'Note: It is preferable to write to the disk in one
  '      go, so that if an error occurrs, the operation
  '      can be repeated with another disk (if FDD)
  '------------------------------------------------------
  Bytes(1) = Asc("a")
  Bytes(2) = Asc("z")
  Bytes(3) = Asc("s")
  Hdr.VolNo = VolNo                     'Current volume no
  Hdr.NextVol = IIf(nLeft > 0, 1, 0)    'More to follow? Yes=1, No=0
  CopyMem Bytes(4), Hdr.Major, 2        'Note: Can't copymem Hdr record in one go!
  CopyMem Bytes(6), Hdr.Minor, 2
  CopyMem Bytes(8), Hdr.VolNo, 1
  CopyMem Bytes(9), Hdr.NextVol, 1
  CopyMem Bytes(10), Hdr.ZipSize, 4
  CopyMem Bytes(14), Hdr.NoFiles, 4
  CopyMem Bytes(18), Hdr.NoIcons, 4
  CopyMem Bytes(22), Hdr.Offset, 4
 
  '------------------------------------------------------
  'Open the destination file, write the contents of the
  'buffer and close the file
  '------------------------------------------------------
  SetProgress Sum, Total, "Spanning volume " & Format(VolNo, "000") & "..."
  DestReady = True

SpanZipFile1:
  nfd = FreeFile
  Open DestFile For Binary As #nfd
  DestOpen = True
SpanZipFile2:
  Put #nfd, , Bytes                     'Write contents
  Close #nfd                            'Close dest file
  DestOpen = False
  DestReady = False
  Sum = Sum + nDest                     'Update progress indicator
  VolNo = VolNo + 1                     'Increment Volume no
Loop
  
'--------------------------------------------------------
'Close source
'--------------------------------------------------------
Close #nfs
'On Error Resume Next
'Kill SourceFile                         'Erase temp zip file
'On Error GoTo SpanZipFileErr
SetProgress Total, Total, ""
Screen.MousePointer = vbNormal
Exit Sub

SpanZipFileErr:
If IsFDD And DestReady Then
  If DestOpen Then Close nfd
  Screen.MousePointer = vbNormal
  ErrTxt = "An error has occurred with the current disk (" & Error & "). Please insert another disk."
  RaiseEvent ChangeDisk(Drive, ErrTxt, "")
  Screen.MousePointer = vbArrowHourglass
  DoEvents
  EraseDrive Drive    'Delete all files & folders on the disk
  DoEvents
  If DestOpen Then
    Resume SpanZipFile1
  Else
    Resume SpanZipFile2
  End If
Else
  Ans = ReportErrorAbort("ZipFiles()", ModName, Err, Error)
  Select Case Ans
  Case vbCancel, vbAbort
    Reset                 'Close all files
    Exit Sub
  Case vbRetry
    Resume
  Case vbOK, vbIgnore
    Resume Next
  End Select
End If

End Sub

Property Get TempFolder() As String
'---------------------------------------------
'Returns the temp folder from which the
'list was built
'---------------------------------------------
TempFolder = mTempFolder

End Property

Property Let TempFolder(vTempFolder As String)
'---------------------------------------------
'Sets the tempfolder from which the list is
'to be built
'---------------------------------------------
mTempFolder = vTempFolder

End Property

Public Property Get TotalSize() As Long
'--------------------------------------------
'Returns the total unzipped size of all files
'--------------------------------------------
Dim i, Total

For i = 1 To nFiles
  Total = Total + aFiles(i).Size
Next
TotalSize = Total

End Property

Public Property Get TotalZipRatio() As Single
'--------------------------------------------
'Returns the total (average) compression ratio
'The lower the ratio, the higher the compression
'achieved (i.e. less disk space occupied)
'--------------------------------------------
Dim Unzipped As Long
Dim Zipped As Long

Zipped = TotalZipSize
Unzipped = TotalSize
If Unzipped > 0 Then
  TotalZipRatio = Zipped / Unzipped
Else
  TotalZipRatio = 1
End If

End Property

Public Property Get TotalZipSize() As Long
'---------------------------------------------
'Returns the total compressed size (incl
'overhead) = FileLen
'---------------------------------------------
Dim i, Total

For i = 1 To nFiles
  Total = Total + aFiles(i).ZipSize
Next
TotalZipSize = Total

End Property

Property Get UnzipFolder() As String
'---------------------------------------------
'Returns the unzip folder to which the archive
'must be restored
'---------------------------------------------
UnzipFolder = mUnzipFolder

End Property

Property Let UnzipFolder(vUnzipFolder As String)
'---------------------------------------------
'Sets the folder to which the archive must
'be restored
'---------------------------------------------
mUnzipFolder = vUnzipFolder

End Property

Public Property Get Size(Index As Long) As Long
'------------------------------------------------
'Returns the filesize (unpacked) in bytes
'for file(i)
'------------------------------------------------
Size = aFiles(Index).Size

End Property

Public Property Get Spanning() As Boolean
'------------------------------------------------
'Returns the disk spanning flag
'------------------------------------------------
Spanning = mSpanning

End Property

Public Property Let Spanning(vSpanning As Boolean)
'------------------------------------------------
'Sets the disk spanning flag
'------------------------------------------------
mSpanning = vSpanning

End Property

Public Property Get SpanSize() As Long
'------------------------------------------------
'Returns the disk spanning size (approximate)
'------------------------------------------------
SpanSize = mSpanSize

End Property


Public Property Get SpanOption() As ArSpanOption
'------------------------------------------------
'Returns the disk spanning size option
'------------------------------------------------
SpanOption = mSpanOption

End Property

Public Property Let SpanOption(vSpanOption As ArSpanOption)
'------------------------------------------------
'Sets the disk spanning size option
'------------------------------------------------
mSpanOption = vSpanOption

End Property

Public Property Get Status() As ArZipStatus
'------------------------------------------------
'Returns the current status (busy, ready etc)
'------------------------------------------------
Status = mStatus

End Property

Public Property Let Status(NewStatus As ArZipStatus)
'------------------------------------------------
'Sets the current status (busy, ready etc)
'(Not used in this app - can remove)
'------------------------------------------------
mStatus = NewStatus

End Property

Sub UnzipFiles(ByVal sUnzipFolder As String, ByVal bAllFiles As Boolean, ByVal bOverwrite As Boolean)
'-----------------------------------------------------------------
'Unzip the files to the specified folder
'sUnzipFolder : Destination folder (will be created if it doesn't exist)
'bAllFiles    : True-Unzip all files, False-selection only
'bOverwrite   : True-Overwrite existing files, false - skip existing files
'a) If spanned disks, unspan the volumes to a temp zip file
'b) Skip file info & start reading & unzipping files
'-----------------------------------------------------------------
Dim Drive As Drive
Dim DriveSpec As String
Dim IsFDD As Boolean        'Writing to a removable drive (Floppy)?
Dim nfd, nfs                'File Numbers
Dim SourceFile As String    'Complete name of source file, incl path & ext
Dim nSource                 'No of bytes to read from source
Dim DestFile As String      'Complete name of destination file, incl path & ext
Dim DestFolder As String    'Destination folder (excl last "\")
Dim SwapDisk As Boolean     'Flag indicating that a disk needs to be inserted/changed
Dim Bytes() As Byte         'Main buffer
Dim VolNo As Long           'Volume no to read, starting at 1
Dim Id As String * 3        'Ariel Zip File Id
Dim Hdr As ArHeader         'Zip file header
Dim ChkSum As Byte          'Original Checksum as calculated, for each file
Dim ChkSumZip As Byte       'Zipped checksum
Dim Sum As Long             'Total bytes handled from source, for progress
Dim Total As Long           'Total bytes to handle from source, for progress
Dim TmrStart As Single      'Start time stamp, to calc elapsed time
Dim Result As ZLibErrors    'Zlib error code
Dim i, n, Ok
Dim Ans


On Error GoTo UnzipFilesErr
'--------------------------------------------------
'Show busy light & set file selection
'--------------------------------------------------
Screen.MousePointer = vbArrowHourglass
SetStatus azsBusy
mUnzipFolder = sUnzipFolder
RefreshSelection bAllFiles

'--------------------------------------------------
'Determine type of drive to read from
'This is necessary since if it is a removable drive
'(i.e. FDD) a dialog must be shown to insert the disk
'--------------------------------------------------
DriveSpec = mFso.GetDriveName(mFso.GetAbsolutePathName(mZipFile))
Set Drive = mFso.GetDrive(DriveSpec)
IsFDD = (Drive.DriveType = Removable)

'---------------------------------------------------
'Initialise other variables
'---------------------------------------------------
VolNo = 1             'Volume No
Sum = 0               'Total bytes processed
Total = TotalZipSize  'Total bytes to process

'---------------------------------------------------
'Determine the source file to open
'---------------------------------------------------
SourceFile = mZipFile

'---------------------------------------------------
'Notes on FDD usage
'---------------------------------------------------
'If a FDD is used, an event is raised indicating to
'the client that the disk needs to be changed. The
'event also indicates which file is required.
'Control should not be returned to this routine
'until a) the correct file has been located or
'      b) the user has cancelled the operation
'Notes
'a) The client must verify the file's existence
'b) The first FDD disk is probably already inserted
'   when the user read the file info. (But it could
'   have been removed in the mean time)
'---------------------------------------------------
'Determine if a disk needs to be changed/inserted
'---------------------------------------------------
If IsFDD Then
  If Not (Drive.IsReady) Then
    SwapDisk = True
  Else
    SwapDisk = Not (mFso.FileExists(SourceFile))
  End If
  If SwapDisk Then
    'Ask to insert disk
    Screen.MousePointer = vbNormal
    RaiseEvent ChangeDisk(Drive, "Please insert disk #" & Format(VolNo, "##0") & " into drive " & Drive.DriveLetter, SourceFile)
    Screen.MousePointer = vbArrowHourglass
    DoEvents
    'Allow for client canceling swap disk dialogue
    If mCancel Then
      mCancel = False         'Reset cancel property
      SetStatus azsReady      'Show green light
      Screen.MousePointer = vbNormal
      Exit Sub
    End If  'Cancel
  End If    'SwapDisk
End If      'IsFDD
  

Do                    'Do this loop either once or twice
  '-----------------------------------------------------
  'Open the source file (could be a spanned volume)
  '-----------------------------------------------------
  nfs = FreeFile
  Open SourceFile For Binary As #nfs

  '-----------------------------------------------------
  'Read Id & test for spanned volume
  '-----------------------------------------------------
  Get #nfs, 1, Id
  Select Case Id
  Case "azs"
    '---------------------------------------------------
    'Spanned archive, unspan to temp file & reopen
    '---------------------------------------------------
    Close #nfs
    UnspanZipfile
    SourceFile = mZipFile
    IsFDD = False
  Case "azp"
    '---------------------------------------------------
    'Normal archive, exit loop & continue
    '---------------------------------------------------
    Exit Do
  Case Else
    '---------------------------------------------------
    'Invalid archive, raise error
    '---------------------------------------------------
    Err.Raise 701, "ArZip.UnzipFiles()", "Invalid Ariel Zip File Id"
    Exit Do
  End Select
Loop

'--------------------------------------------------
'Start timer
'--------------------------------------------------
TmrStart = Timer
  
'---------------------------------------------------
'Read Header & check volume no
'---------------------------------------------------
Get #nfs, 4, Hdr
If VolNo <> Hdr.VolNo Then
  Err.Raise 702, "ArZip.UnzipFiles()", "Volume No doesn't match - looking for vol " & Format(VolNo, "000") & "."
End If

'---------------------------------------------------
'Create destination rootfolder (if necessary)
'---------------------------------------------------
CreateFolders mUnzipFolder

'-----------------------------------------------------
'Unzip files
'-----------------------------------------------------
For i = 1 To nFiles
  '-------------------------------------------------
  'If overwrite=false, check if file exists.
  'If it does, skip it
  '-------------------------------------------------
  DestFile = mUnzipFolder & aFiles(i).Key
  If bOverwrite Then
    Ok = aFiles(i).Selected
  Else
    If mFso.FileExists(DestFile) Then
      Ok = False
    Else
      Ok = aFiles(i).Selected
    End If
  End If

  If Ok Then
    SetProgress Sum, Total, "Unzipping file " & aFiles(i).Name & "..."
    '---------------------------------------------------
    'Create destination folder if not root
    '---------------------------------------------------
    If aFiles(i).RelPath <> "" Then
      DestFolder = CheckPath(mUnzipFolder) & aFiles(i).RelPath
      CreateFolders DestFolder
    End If
    'Get the no of bytes to read for file(i)
    nSource = aFiles(i).ZipSize
    If nSource > 0 Then
      '-------------------------------------------------
      'Read data for file(i) to unzip
      '-------------------------------------------------
      ReDim Bytes(nSource) As Byte
      n = aFiles(i).Offset          'Get offset position
      Get #nfs, n + 1, Bytes        'Read data to buffer
      Sum = Sum + nSource           'Update progress indicator
      
      '-------------------------------------------------
      'Validate zip checksum
      '-------------------------------------------------
      ChkSumZip = CheckSum(Bytes)
      If ChkSumZip <> aFiles(i).ChkSumZip Then
        Err.Raise 703, "ArZip.UnzipFiles()", "Invalid zip checksum in file " & DestFile
      End If
      '-------------------------------------------------
      'Decompress bytes using zlib.dll 1.1.3
      'The original size needs to be given in order to
      'make enough space (an extra 1% + 12 bytes is added
      'as a temporary measure, whereafter the buffer is
      'resized to the original size. The size parameter
      'is passed by value, so no need to protect it.
      '-------------------------------------------------
      Result = DecompressBytes(Bytes, aFiles(i).Size)
      If Result <> zlOk Then
        Select Case Result
        Case zlStreamEnd    '1
          Err.Raise 801, "ArZip.UnzipFiles()", "Error decompressing file " & DestFile & " : Stream end"
        Case zlNeedDict     '2
          Err.Raise 802, "ArZip.UnzipFiles()", "Error decompressing file " & DestFile & " : Dictionary required"
        Case zlErrNo        '-1
          Err.Raise 803, "ArZip.UnzipFiles()", "Error decompressing file " & DestFile & " : Unknown error"
        Case zlStreamError  '-2
          Err.Raise 804, "ArZip.UnzipFiles()", "Error decompressing file " & DestFile & " : Stream error"
        Case zlDataError    '-3
          Err.Raise 805, "ArZip.UnzipFiles()", "Error decompressing file " & DestFile & " : Input data corrupted"
        Case zlMemError     '-4
          Err.Raise 806, "ArZip.UnzipFiles()", "Error decompressing file " & DestFile & " : Insufficient memory"
        Case zlBufError     '-5
          Err.Raise 804, "ArZip.UnzipFiles()", "Error decompressing file " & DestFile & " : Insufficient space in output buffer"
        Case zlVersionError '-6
          Err.Raise 805, "ArZip.UnzipFiles()", "Error decompressing file " & DestFile & " : zlib version error"
        End Select
      End If
      
      '-------------------------------------------------
      'Validate checksum
      '-------------------------------------------------
      ChkSum = CheckSum(Bytes)
      If ChkSum <> aFiles(i).ChkSum Then
        Err.Raise 704, "ArZip.UnzipFiles()", "Invalid checksum in file " & DestFile
      End If
      
      '-------------------------------------------------
      'Write file to disk
      '-------------------------------------------------
      nfd = FreeFile
      'DestFile = mUnzipFolder & aFiles(i).Key
      On Error Resume Next
      Kill DestFile
      On Error GoTo UnzipFilesErr
      Open DestFile For Binary As #nfd
      Put #nfd, , Bytes
      Close #nfd
    Else
      'Write 0-byte file to disk
      nfd = FreeFile
      DestFile = CheckPath(mUnzipFolder) & aFiles(i).Key
      On Error Resume Next
      Kill DestFile
      On Error GoTo UnzipFilesErr
      Open DestFile For Binary As #nfd
      Close #nfd
    End If
  Else
    'Skip file (not selected)
    Sum = Sum + aFiles(i).ZipSize
  End If
Next            'Next File

'-----------------------------------------------------
'Close Source (Zip) file
'-----------------------------------------------------
SetProgress Total, Total, ""
Close #nfs
SetStatus azsCreated
mElapsedTime = Timer - TmrStart
Screen.MousePointer = vbNormal
Exit Sub

UnzipFilesErr:
Ans = ReportErrorAbort("UnzipFiles()", ModName, Err, Error)
Select Case Ans
Case vbCancel, vbAbort
  Reset       'Closes all files
  Exit Sub
Case vbRetry
  Resume
Case vbOK, vbIgnore
  Resume Next
End Select

End Sub
Private Function ValidateFile(File As File) As Boolean
'---------------------------------------------------------------
'Validates a file for inclusion in the archive
'Can fail for two reasons:
'a) File is part/same of the current archive (.azp or .azs)
'b) File doesn't exist on hdd (i.e. unzipped version)
'Returns true if successful
'---------------------------------------------------------------
Dim Key As String
Dim Path As String        'Relative Path
Dim FileBase As String
Dim ZipBase As String
Dim i, Ok As Boolean

'Test for existence
Ok = mFso.FileExists(File.Path)

'Test for "azp" file
If Ok And FileExtension(File.Path) = "azp" Then
  ZipBase = LCase(mFso.GetBaseName(mZipFileNew))    'Core name of zipfile e.g. "test" for test.azp
  FileBase = LCase(mFso.GetBaseName(File.Name))
  If FileBase = ZipBase Then Ok = False
End If
If Ok And FileExtension(File.Path) = "azs" Then
  ZipBase = LCase(mFso.GetBaseName(mZipFileNew)) & "-"    'Core name of zipfile e.g. "test-" for test.azs
  FileBase = LCase(mFso.GetBaseName(File.Name))
  If InStr(FileBase, ZipBase) > 0 Then Ok = False
End If
ValidateFile = Ok

End Function
Private Function Wait(ByVal MsToWait As Long)
'--------------------------------------------------
'Waits for specified milliseconds while yielding to
'other events. Added the sleep function to
'cut down on processor usage
'Not used - keep for reference
'--------------------------------------------------
Dim EndTime As Long

EndTime = GetTickCount + MsToWait
Do Until GetTickCount > EndTime
  DoEvents
  Call Sleep(250)
Loop

End Function

Property Let ZipFile(vZipFile As String)
'----------------------------------------------
'Set the name & path of the current zip file
'----------------------------------------------
mZipFile = vZipFile

End Property

Property Get ZipFile() As String
'---------------------------------------------
'Returns the path & name of the current zip file
'---------------------------------------------
ZipFile = mZipFile

End Property

Public Sub ZipFiles(ByVal sNewZipFile As String, ByVal iCompressLevel As Integer, ByVal bSpanning As Boolean, ByVal lSpanOption As Long)
'------------------------------------------------------------------------------------
'Compress the files to the NewZipFile using Byte Arrays
'NewZipFile   : Name & Path of new zip file. Must have .azp. extension
'CompressLevel: Level of compression, 0-9
'Spanning     : If true, spans zip file in volumes
'SpanOption   : Capacity options for spanned volumes (see frmAzZip combobox)
'------------------------------------------------------------------------------------
'Notes:
'a) If a current archive is being updated (mExist) then make a backup
'   to the file 'Temp Old.azp' in the tempfolder (if not already done)
'   When a previously zipped file needs to be included in the archive,
'   it can thus be copied from the backup file
'b) If disk spanning has been set, write to a new temp file 'Temp New.azp' in
'   the temp folder. Once the zip file has been fully written, the spanning
'   can be done.
'c) If writing to a removable FDD drive, automatically turn disk spanning on,
'   and write to the temp file first
'd) The zip file is written in chunks (currently 1Mb). See the CHUNK_SIZE
'   constant
'------------------------------------------------------------------------------------
'Overall structure of zip file
'a) File Id             "azp"               3 bytes  1-3
'b) Archive Header      22 bytes
'                       Major RevNo   Int   2 bytes  4-5
'                       Minor RevNo   Int   2 bytes  6-7
'                       VolNo         Byte  1 bytes  8
'                       NextVol       Byte  1 bytes  9
'                       ZipSize       Long  4 bytes  10-13
'                       NoFiles       Long  4 bytes  14-17
'                       NoIcons       Long  4 bytes  18-21
'                       Offset        Long  4 bytes  22-25
'c) Rootfolder          var length string
'd) "[Files]"           Identifies start of file list
'e) File Header         Header for each file (1 to nFiles) (22 bytes)
'                       Offset        Long  4 bytes  (offset in zip file, from start)
'                       OrigSize      Long  4 bytes  (uncompressed size, bytes)
'                       ZipSize       Long  4 bytes  (compressed size, bytes)
'                       Modified      Date  8 bytes  (last file modification date)
'                       Checksum      Byte  1 bytes  (original to unzipped)
'                       ChecksumZip   Byte  1 bytes  (zip write to zip read)
'f) "[Names]"           Identifies start of file name list
'g) File Names          String list for each file (1 to nFiles) (variable length)
'                       File Name     Var String + vbCr
'                       Rel Path      Var String + vbCr
'                       Path          Var String + vbCr
'                       Icon Key      Var String + vbCr
'h) "[Icons]"           Identifies start of icons
'i) Icons               Icon info for each  unique icon (1 to nIcons)
'                       Icon Key      Var String  (used in image list)
'                       Bitmap        16x16x256 (822 bytes)
'j) Files               Compressed data for each file (1-nFiles)
'                       Note: 0-length files are not zipped (but are incl in list)
'------------------------------------------------------------------------------------
Dim Drive As Drive        'FSO Drive object
Dim DriveSpec As String   'Required to resolve the drive object
Dim IsFDD As Boolean      'Is destination drive removable?
Dim nfd, nfs              'File Numbers
Dim nb                    'Byte position in buffer
Dim SourceFile As String  'Name of sourcefile, incl path
Dim DestFile As String    'Name of dest file, incl path (zip archive)
Dim DestBase As String    'Dest file name base component, excl path, vol no & ext
Dim DestFolder As String  'Dest folder, excluding last "\"
Dim Bytes() As Byte       'Archive Data written to disk
Dim Buffer() As Byte      'Temporary buffer
Dim Hdr As ArHeader       'Used for the numeric part of the header
Dim Afr As ArFileRec      'File header (one for each source file)
Dim imgObj As ListImage   'Single image object in mIml.listimages collection
Dim Result As ZLibErrors  'Zlib error code
Dim Sum As Long           'Total bytes handled (from source)
Dim Total As Long         'Total bytes to handle (from source)
Dim TmrStart As Single
Dim i, n, Ok
Dim Ans

On Error GoTo ZipFilesErr
'--------------------------------------------------
'Start timer & show busy light
'--------------------------------------------------
TmrStart = Timer
SetStatus azsBusy
Screen.MousePointer = vbArrowHourglass

'--------------------------------------------------
'Set entry properties
'--------------------------------------------------
mZipFileNew = sNewZipFile
mCompressLevel = iCompressLevel
mSpanning = bSpanning
mSpanOption = lSpanOption

'--------------------------------------------------
'Make a backup / unspan volumes if required
'--------------------------------------------------
If mExist Then
  If mSpanned Then
    UnspanZipfile
  End If
  BackupZipFile
End If

'--------------------------------------------------
'Determine type of drive to write to
'This is necessary since if it is a removable drive
'(i.e. FDD), a dialog must be shown to switch disks
'--------------------------------------------------
DriveSpec = mFso.GetDriveName(mFso.GetAbsolutePathName(mZipFileNew))
Set Drive = mFso.GetDrive(DriveSpec)
IsFDD = (Drive.DriveType = Removable)

'---------------------------------------------------
'Initialise other variables
'---------------------------------------------------
Sum = 0                   'Total bytes processed
Total = 0                 'Total bytes to process
'Populate header
Hdr.Major = App.Major     'Int  2x bytes
Hdr.Minor = App.Minor     'Int  2x bytes
Hdr.VolNo = 1             'Byte 1x bytes
Hdr.NextVol = 0           'Byte 1x bytes    Used for spanning, not here
Hdr.ZipSize = 0           'Long 4x bytes    Not known now, done later
Hdr.NoFiles = nFiles      'Long 4x bytes
Hdr.NoIcons = nIcons      'Long 4x bytes
Hdr.Offset = 0            'Long 4x bytes    Points to start of file list (done later)
                          '                 (only applicable to volume 0)

'---------------------------------------------------
'Determine destination folder & file name
'---------------------------------------------------
DestBase = mFso.GetBaseName(mZipFileNew)    'Core name of zipfile e.g. "test" for test.azp
If IsFDD Then mSpanning = True
If mSpanning Then
  DestFolder = mTempFolder
  DestFile = CheckPath(DestFolder) & DestBase & ".azp"
  'Save as tempfile for use later
  mTempFile = DestFile
Else
  DestFolder = mFso.GetParentFolderName(mZipFileNew)
  DestFile = CheckPath(DestFolder) & DestBase & ".azp"
  mTempFile = ""
End If

'---------------------------------------------------
'Open the destination file
'---------------------------------------------------
'Kill any file with the same name as destination file
'to prevent a left-over included in the archive
On Error Resume Next
Kill DestFile
On Error GoTo ZipFilesErr
nfd = FreeFile
Open DestFile For Binary As #nfd

'---------------------------------------------------
'Header
'---------------------------------------------------
SetProgress 0, 1, "Adding header..."
nb = AddHeader(Bytes, Hdr)                 'Add header
nb = AddString(Bytes, mRootFolder, True)   'Rootfolder (variable length)

'---------------------------------------------------
'File List (Numeric portion)
'---------------------------------------------------
nb = AddString(Bytes, "[Files]")
'Save the current offset for use later when updating the file list
Hdr.Offset = nb
'Make space for a file header for each file
'Contains Offset,OrigSize,ZipSize,Date Modified & ChkSum
'Actual data is only added at the end of the zip operation
nb = AddBlankFileRecords(Bytes)

'---------------------------------------------------
'File Names (String portion)
'---------------------------------------------------
Call AddString(Bytes, "[Names]")
For i = 1 To nFiles
  'Make space for Name, Rel Path, Path & Icon Key
  nb = AddFileName(Bytes, i)
Next

'---------------------------------------------------
'Icons
'---------------------------------------------------
'Open each icon (bmp) file and add data to Bytes
'together with icon key (used in imagelist)
Call AddString(Bytes, "[Icons]")
For Each imgObj In mIml.ListImages
  nfs = FreeFile
  SourceFile = CheckPath(mTempFolder) & imgObj.Key & ".bmp"
  Open SourceFile For Binary As #nfs
  n = LOF(nfs)
  ReDim Buffer(n) As Byte
  'Read the bitmap
  Get #nfs, , Buffer
  Close #nfs
  'Add the bitmap and icon key to Bytes()
  nb = AddIconRecord(Bytes, Buffer, imgObj.Key)
Next

'--------------------------------------------------
'Files
'--------------------------------------------------
'Get total length of source files to write
'for progress indicator
Total = Me.TotalSize
Sum = 0
For i = 1 To nFiles
  '----------------------------------------------------------
  'Determine if file needs to be zipped from its source or
  'if it needs to be copied from the old (current) zip file
  '----------------------------------------------------------
  If aFiles(i).Zipped Then
    'This file does not require freshening and exists in the
    'current zip file. Therefore, extract the bytes from the
    'current zip file. If an error occurs, set no bytes to 0
    Ok = ExtractFile(i, Buffer)              'Decompress also
    If Ok Then
      n = UBound(Buffer)                     'No of bytes in Buffer
    Else
      n = 0
    End If
  Else
    'This file either does not exist in the zip file, or
    'requires 'freshening'. So read it from source
    nfs = FreeFile
    SourceFile = CheckPath(aFiles(i).Path) & aFiles(i).Name
    Open SourceFile For Binary As #nfs
    n = LOF(nfs)                              'No of bytes in the file
    If n > 0 Then                             'Source file contains data
      'Read the file into Buffer()
      ReDim Buffer(n) As Byte                 '1-based
      Get #nfs, , Buffer()
      Close #nfs
    End If
  End If
   
  If n > 0 Then
    '----------------------------------------------
    'Do compression using zlib.dll
    '----------------------------------------------
    'Calculate crude XOR check sum of original file
    aFiles(i).ChkSum = CheckSum(Buffer)
    'Compress the byte array using zlib.dll call
    Result = CompressBytes(Buffer, mCompressLevel)
    If Result <> 0 Then
      Err.Raise Result, "ArZip.ZipFiles()", "Error compressing byte array"
    End If
    'Calculate crude XOR check sum of zipped file
    aFiles(i).ChkSumZip = CheckSum(Buffer)
    'Collect file info
    aFiles(i).Offset = Loc(nfd) + nb                    'Offset in dest file, from start
    aFiles(i).Size = n                                  'Original size (bytes)
    aFiles(i).ZipSize = UBound(Buffer)                   'Zip size
    aFiles(i).Ratio = aFiles(i).ZipSize / aFiles(i).Size 'Zip ratio
    aFiles(i).Zipped = True
    'Add compressed data to buffer
    nb = AddBytes(Bytes, Buffer)
  Else
    'Collect file info
    aFiles(i).Offset = Loc(nfd) + nb                    'Offset in dest file, from start
    aFiles(i).Size = 0                                  'Original size (empty)
    aFiles(i).ZipSize = 0                                'Zip size (empty)
    aFiles(i).Ratio = 1                                 'Zip ratio 100%
    aFiles(i).ChkSum = 0
    aFiles(i).ChkSumZip = 0
    aFiles(i).Zipped = True
    'No need to add empty bytes to Bytes() array
  End If
  '-------------------------------------------------------
  'Write Bytes() to destination if chunk size exceeded
  '-------------------------------------------------------
  If nb >= CHUNK_SIZE Then
    Put #nfd, , Bytes()             'Write Bytes() to disk
    ReDim Bytes(0 To 0) As Byte     'To indicate 0 bytes in buffer
    nb = 0                          'Reset buffer offset
  End If
  Sum = Sum + n
  SetProgress Sum, Total, "Adding file " & aFiles(i).Name
Next

'-----------------------------------------------------
'Write remaining Bytes() to Destination (Zip) file
'-----------------------------------------------------
Put #nfd, , Bytes()                 'Loc(nfd) + 1

'-----------------------------------------------------
'Update Header & File Headers
'-----------------------------------------------------
'Header
Hdr.ZipSize = Loc(nfd)
Put #nfd, 4, Hdr                    'Right after "azs" or "azp" Id

'Debugging:
'ReDim Buffer(Len(Hdr) + 3) As Byte
'Get #nfd, 1, Buffer
'For i = 1 To 25: Debug.Print i, Buffer(i), Chr(Buffer(i)): Next

'File Headers
n = Hdr.Offset
For i = 1 To nFiles
  Afr.Offset = aFiles(i).Offset
  Afr.OrigSize = aFiles(i).Size
  Afr.ZipSize = aFiles(i).ZipSize
  Afr.Modified = aFiles(i).Modified
  Afr.ChkSum = aFiles(i).ChkSum
  Afr.ChkSumZip = aFiles(i).ChkSumZip
  Put #nfd, n + 1, Afr
  n = n + Len(Afr)
Next

'-----------------------------------------------------
'Close destination file
'-----------------------------------------------------
Close #nfd
mTotalZipSize = FileLen(DestFile)   'Must correspond with Hdr.Zipsize
mExist = True                       'Update the zip file exists flag
mZipFile = DestFile
SetProgress Total, Total, ""
mElapsedTime = Timer - TmrStart

'----------------------------------------------------
'Diskspanning
'----------------------------------------------------
If mSpanning Then
  SpanZipFile
End If
Screen.MousePointer = vbNormal
SetStatus azsCreated
Exit Sub

ZipFilesErr:
Ans = ReportErrorAbort("ZipFiles()", ModName, Err, Error)
Select Case Ans
Case vbCancel, vbAbort
  Reset       'Closes all files
  Exit Sub
Case vbRetry
  Resume
Case vbOK, vbIgnore
  Resume Next
End Select

End Sub

Public Sub RegisterArielFileTypes()
'--------------------------------------------------------
'Registers the '.azp' extension to allow azp files to
'be opened with this application
'Also registers the 'azs' file type as an Ariel Spanned Zip Volume
'This requires regobj.dll in the system folder
'Also, add a reference to 'Registration Manipulation Classes'
'(the regobj.dll) in this project. See the References item
'on the Project menu in the VB IDE.
'--------------------------------------------------------
Dim Path As String          'Path of exe file, i.e. 'C:\program files\ariel zip\ariel zip.exe'
Dim FileName As String      'exe file name to run i.e. 'ariel zip.exe'
Dim Ext As String           'The file extension, '.azp'
Dim Desc As String          'File extionsion description
Dim CmdSt As String         'Command string:  Exepath & " %1"
Dim IcnSt As String         'Default Icon  :  Exepath & ",0"
Dim objRegKey As RegKey     'See objreg.dll documentation
Dim objRootKey As RegKey
Dim objSubKey As RegKey
  
  
On Error Resume Next        'To prevent adding duplicate key errors

'AZP file
Ext = ".azp"
FileName = App.EXEName & ".exe"
Path = App.Path & "\" & FileName
Desc = "Ariel Zip File"
CmdSt = Path & " %1"
IcnSt = Path & ",0"

'Collect the parent of the new key - in this case HKEY_CLASSES_ROOT
Set objRootKey = RegKeyFromHKey(HKEY_CLASSES_ROOT)
'Add a new file extension SubKey
Call objRootKey.SubKeys.Add(Ext)
'Get a reference to the new subkey
'Use: Set objRegKey = RegKeyFromString("\HKEY_CLASSES_ROOT\" & Ext)
'or
Set objRegKey = objRootKey.SubKeys(Ext)
'Set the description to the default value
objRegKey.Value = Desc


'Add the Shell Command
Call objRegKey.SubKeys.Add("shell")           'Create the shell subkey
Set objSubKey = objRegKey.SubKeys("shell")    'Get the shell subkey
Call objSubKey.SubKeys.Add("open")            'Create the open subkey
Set objSubKey = objSubKey.SubKeys("open")     'Get the open subkey
Call objSubKey.SubKeys.Add("command")         'Create the command subkey
Set objSubKey = objSubKey.SubKeys("command")  'Get the command subkey
objSubKey.Value = CmdSt                       'Set the command key value

'Default Icon
Call objRegKey.SubKeys.Add("DefaultIcon")         'Create the DefaultIcon key
Set objSubKey = objRegKey.SubKeys("DefaultIcon")  'Get the subkey
objSubKey.Value = IcnSt                           'Set the default value

'-------------------------------------
'azs file
'-------------------------------------
Ext = ".azs"
Desc = "Ariel Spanned Zip Volume"

'Add a new file extension SubKey & get reference
Call objRootKey.SubKeys.Add(Ext)
Set objRegKey = objRootKey.SubKeys(Ext)
'Set the description to the default value
objRegKey.Value = Desc
'Default Icon
Call objRegKey.SubKeys.Add("DefaultIcon")         'Create the DefaultIcon key
Set objSubKey = objRegKey.SubKeys("DefaultIcon")  'Get the subkey
objSubKey.Value = IcnSt                           'Set the default value

Set objRootKey = Nothing
Set objRegKey = Nothing
Set objSubKey = Nothing

End Sub

Public Property Get ZipSize(Index As Long) As Long
'------------------------------------------------
'Returns the zipped filesize in bytes
'------------------------------------------------
ZipSize = aFiles(Index).ZipSize

End Property

Private Sub Class_Initialize()
'---------------------------------------------------
'Initialise the class
'---------------------------------------------------
mTempFolder = GetTempFolder

End Sub

Private Sub Class_Terminate()
'----------------------------------------------
'Clears the memory
'----------------------------------------------
On Error Resume Next

Kill mTempFolder & "\*.bmp"     'Delete all icons
Kill mTempFolder & "\*.azp"     'Delete all azp files
Set mIml = Nothing
Set mLvwFiles = Nothing
Set mPic = Nothing

End Sub



