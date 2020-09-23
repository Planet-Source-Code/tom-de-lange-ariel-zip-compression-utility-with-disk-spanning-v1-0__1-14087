Ariel Zip Compression Utility with Disk Spanning V1.0
-----------------------------------------------------
A powerful and very quick compression program using the 32bit zlib.dll (V1.1.3) compression library originally written by Jean-loup Gailly and Mark Adler (variation of LZ77 Lempel-Ziv 1977 algoritm). The application is implemented through an extensive ArielZip class and a Winzip like client interface.

Features
--------
* Powerful and very fast - compresses 6000 kb data files to 1500 kb (25%) in 2.1 seconds! * Multiple disk spanning with automatic sensing of disk capacity * User friendly floppy disk change dialogue showing contents of disk to be overwritten * Single and multiple file extraction * Extracted icons included in archive as bitmaps * 9 compression levels * Add folders and subfolders through recursive scanning of FSO objects * Add files with multiple file selection * Delete files from list and archive * Refresh files * File association of .azp extension with default icon using regobj.dll (included in zip file).

ArielZip Class
--------------
Public methods: AddFiles(), AddFolder(), Clear(), Initialise(), NewArchive(), OpenArchive(), RefreshFiles(), RegisterArielFileTypes(), RemoveFile(), UnregisterArielFileTypes(), UnzipFiles() and Unzipfiles()

Events: StatusChange(), Progress(), ChangeDisk()

Properties: Cancel, CompressLevel, ElapsedTime, Exist, Ext, IconKey, Key, Modified, Name, NoFiles, Path, Ratio, RelativePath, Rootfolder, Size, Spanning, SpanOption, SpanSize, Status, TempFolder, TotalSize, TotalZipRatio, TotalZipSize, Unzipfolder, Unzipfile

Other Programming Features
--------------------------
- About box referencing application object (revision etc)
- Custom Folder browse control implementing BrowseForFolder (ArielBrowseControl)
- Extensive use of FileSystemObject (requires scripting runtime dll)
- Automatic sensing of floppy disk insertion/removal
- Demonstration of toolbar control
- Small icon extraction using SHGetFileInfo calls in shell32.dll lib
- Extensive usage of CopyMem (Kernel32)
- Access the system temp folder through GetTempPath (kernel32)
- Registering of azp file type and associated default icon through regobj.dll
- Ini file manipulation using GetProfile and SaveProfile functions

Installation of Source Code
---------------------------
a) Unzip file to a designated folder
b) Double-click on the Easy-Register.reg file to add right mouse register/unregister functionality to explorer (You can use other methods of registering dll's and ocx's as well.
c) Copy the files regobj.dll, zlib.dll and Ariel Browse Ctrl.ocx to your windows system directory
d) NB - Register these files using the easy-register.reg utility (or any other). 
e) The source was compiled with VB6 SP4. If you are using VB5 you may have to replace
the mscomctl.ocx reference in the vbp file with the version you are using (contains listview and image list objects) To do this, simply oben the vbp file with notepad and replace the following line with one from your project files
   Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0; MSCOMCTL.OCX
f) If you still get errors when loading the Ariel Zip project files, continue to load all the files. Then remove all the forms, modules and classes. Do not save any files while doing this. Once you have only the project file loaded, make sure that the following components are checked (See Project/Components (Ctrl-T) and Projects/References menus):
Components:  	Ariel System Browse Controls (Ariel Browse Ctrl.ocx)
		Microsoft Common Dialog Control 6.0 (SP3) (comdlg32.ocx)
		Microsoft Windows COmmon Contols 6.0 (SP4) (mscomctl.ocx)
References:	Microsoft Scripting Runtime (scrrun.dll)
		Registration Manipulation Classes (regobj.dll)
Once you have located and checked these components and dlls, add the forms to the project. Then save the project.

Using the Ariel Zip Program
---------------------------
Usage is similar to Winzip with one exception: the zip file is not created when adding the files and/or folders, but only when the 'Save' button is clicked. This gives more flexibility in that the same file can be zipped to differenct locations.
The order of usage is as follows: 
a) Click New to create a new archive - provide the archive name and root folder. Select if you want to include all files in the rootfolder and subfolders.
b) Click Add Folder to add another folder to the archive (with/without subfolders)
c) Click Add Files to add one or more files to the archive
d) If you want to remove files from the file list, select them (using the multiple selection of the listview) and click Delete
e) Finally click 'Zip' (Save icon) to create the zip file. Here you set the compression level and disk spanning options.
To Unzip a file:
a) Click Open to load the contents of an archive (reads only the header and file info)
b) Click Unzip to unzip the files (using either all or selected files options)
The colored ball at the bottom right of the status bar means the following:
a) Grey - zip file empty (no archive specified)
b) Green - ready to zip (archive has been specified and files have been added to list)
c) Blue - ready to unzip (archive has been created)
d) Red - busy with operation - do not interrupt.

Compression Files
-----------------
All files to be included in the archive are added to a file list, using an array of user define type. This udt includes the file name, path, relative path, original size, zipped size, key, iconkey, date modified etc. When adding the files, they are checked if they exist and for invalid files (like a previous version of the archive the user is attempting to create).

If an existing archive has been opened, a backup of it is made in the windows temp folder. This is required since some files that need to be included in the archive may not have been unzipped. So when the zip archive is created (zip/save button), the files that have now been added are read from its original source and all the others are read from the backup zip file.

When spanning is specified, a temp zip file is created in the windows temp folder. Once created, the temp file is broken into the different pieces, called volumes. A nice feature is that when spanning is done to floppy disks, a dialogue to change disk is shown. This dialogue automatically senses when a disk has been inserted and displays the current contents. The user can therefore decide not to use this disk and insert another. 
Another nice feature is that disks with different capacities can be mixed - i.e. 1.44Mb and 720kb disks can be interchanged, even if a 1.44 Mb span size has been specified. In all cases removable disks (floppies) will be erased prior to spanning.
When unzipping a spanned volume, the same temp file is created. This file is used as the source for additions, deletions etc, so no need to insert the floppy disks more than once.

Archive File Structure
----------------------
The structure of the archive is summarised as follows (more detail can be found in the source code - see the ZipFiles() subroutine in the ArZip class module):
a) File Id          "azp" (Normal) or "azs" (Spanned volume) (3 bytes)
b) Archive Header   Revision No, Vol No, ZipFileSize, NoFiles, NoIcons (22 bytes) 
c) Rootfolder       var length string
d) File Header	    22 bytes per file (offset, origsize,zipsize,checksum etc)
e) File Names       File Name and path info per file (variable len strings)
f) Icons	    Icon info for each unique icon in the archive incl 16x16x256 bitmap
g) Files            Compressed data for each file

Limitations
-----------
a) Since the compression is done in memory for a single file as a whole in one go, the amount of RAM in your computer will determine the maximum file size that can be included in the archive. Copying and spanning of archive files are done in chunks of 1Mb (and can be changed by simply changing the CHUNK_SIZE constant)
b) Apparently there is a physical limit to the string length that can be used (about 2Gb), but I don't think this has an affect in the program.
c) Since a backup archive is made, be sure to have at least 2x the amount of disk space available than required for an archive (3x for spanned volumes).

Credits
-------
I am in debt to the work of others, whose source code can mainly be found on Planet Source Code. PSC is a truly great concept. Specific credit goes to these PSC contributors:
a) Dan Davis (Compression Utility) - copymem usage, objreg.dll usage (file association)
b) Dan Davis (Self Extracting Compressed Files) - lots of small things
c) Mark Withers (Cybercript 5.0) - usage of zlib.dll, copymem usage, byte manipulation, registry manipulation
d) Fredrik Qvarfort (Huffman Encoding/Decoding Class) - byte manipulation
e) Vasilis Sagonas (Self extractor, no compression) - general file handling
f) DarrynB (Daz B) (Smartbasket) - good drag/drop interface concept to packing (no compression)
g) Peter Meier (DelRecent) - icon extraction
h) Josh B FLinders (Common Dialog Example) - multiple file selection
i) Doug Gaede (String, array and file compression with zlib.dll) - using zlib

Other sites:
a) Mark Le Voi at Info-zip	ftp://ftp.freesoftware.com/pub/infozip/index.html
b) zlib home page		ftp://ftp.freesoftware.com/pub/infozip/zlib/zlib.html
c) zlib.dll			http://www.winimage.com/zLibDll/
				To download, click on zlib113adll.zip link
d) Richard Southey (ZipIt) - read Winzip file info (not used, maybe later)
   http://www.richsoftcomputing.btinternet.co.uk/index.html

Final Word
----------
Why bother with your own compression utility? Well, I want to include a miniture version in a share/equity analysis program for purpose of backing up data files. Once started, I decided to add a lot more features so as to make it a generally more useful project. My real motive for doing this was simply to go for the gold... So, if you like this submission, please vote!
If you need more info on the Ariel Browse Control, search for 'Ariel' in Planet source code
at http://www.planet-source-code.com/vb/    I've also written a neat color picker.

For more info on zlib, contact me by e-mail at tomdl@attglobal.net 
Enjoy!
===============================
Tom de Lange
Centurion, South Africa
January 2001
===============================

