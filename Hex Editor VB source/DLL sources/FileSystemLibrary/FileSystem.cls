VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileSystem"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' =======================================================
'
' Hex Editor VB
' Coded by violent_ken (Alain Descotes)
'
' =======================================================
'
' A complete hexadecimal editor for Windows ©
' (Editeur hexadécimal complet pour Windows ©)
'
' Copyright © 2006-2007 by Alain Descotes.
'
' This file is part of Hex Editor VB.
'
' Hex Editor VB is free software; you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation; either version 2 of the License, or
' (at your option) any later version.
'
' Hex Editor VB is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with Hex Editor VB; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
' =======================================================


Option Explicit
'BUG LIGNE 1157 RESOLU

'=======================================================
'CONSTANTES
'=======================================================
Private Const GENERIC_READ                  As Long = &H80000000
Private Const FILE_SHARE_READ               As Long = &H1
Private Const FILE_SHARE_WRITE              As Long = &H2
Private Const OPEN_EXISTING                 As Long = 3
Private Const GENERIC_WRITE                 As Long = &H40000000
Private Const BIF_RETURNONLYFSDIRS          As Long = 1
Private Const BIF_DONTGOBELOWDOMAIN         As Long = 2
Private Const CREATE_NEW                    As Long = 1
Private Const FILE_END                      As Long = 2&
Private Const FILE_BEGIN                    As Long = 0
Private Const CREATE_ALWAYS                 As Long = 2
Private Const FILE_FLAG_BACKUP_SEMANTICS    As Long = &H2000000
Private Const FILE_ATTRIBUTE_DIRECTORY      As Long = &H10
Private Const FILE_SHARE_DELETE             As Long = &H4
Private Const FILE_FLAG_NO_BUFFERING        As Long = &H20000000
Private Const SEE_MASK_INVOKEIDLIST         As Long = &HC
Private Const SEE_MASK_NOCLOSEPROCESS       As Long = &H40
Private Const SEE_MASK_FLAG_NO_UI           As Long = &H400
Private Const FO_DELETE                     As Long = &H3
Private Const FOF_ALLOWUNDO                 As Long = &H40
Private Const FO_COPY                       As Long = &H2&
Private Const FO_MOVE                       As Long = &H1&
Private Const FOF_SIMPLEPROGRESS            As Long = &H100&
Private Const DJM_SMALLICON                 As Long = &H6605
Private Const PICTYPE_ICON                  As Long = 3
Private Const INVALID_HANDLE_VALUE          As Long = -1
Private Const IOCTL_DISK_GET_DRIVE_GEOMETRY As Long = &H70000
Private Const IOCTL_DISK_GET_PARTITION_INFO As Long = &H74004
Private Const IDENTIFY_BUFFER_SIZE          As Long = 512
Private Const OUTPUT_DATA_SIZE              As Long = IDENTIFY_BUFFER_SIZE + 16
Private Const DFP_RECEIVE_DRIVE_DATA        As Long = &H7C088
Private Const FILE_TYPE_CHAR                As Long = &H2
Private Const FILE_TYPE_DISK                As Long = &H1
Private Const FILE_TYPE_PIPE                As Long = &H3
Private Const FILE_TYPE_UNKNOWN             As Long = &H0
Private Const LBUFSIZE                      As Long = 5000000


'=======================================================
'TYPES & ENUMS
'=======================================================
'=======================================================
'Attributs possible d'un fichier
'=======================================================
Public Enum FileAttributes
    Alias = 1024
    Archive = 32
    Compressed = 2048
    Directory = 16
    Hidden = 2
    Normal = 0
    ReadOnly = 1
    System = 4
    Volume = 8
End Enum
'=======================================================
'Type possible pour un drive
'=======================================================
Public Enum DriveType
    [CD Rom] = 5
    [Drive Fixed] = 3
    [RAM Disk] = 6
    [Remote] = 2
    [Removable Drive] = 2
    [UnknowType] = 0
End Enum
'=======================================================
'Rértoire spéciaux
'=======================================================
Public Enum SpecialFolderConst
    CSIDL_DESKTOP = &H0                 '{desktop}
    CSIDL_INTERNET = &H1                'Internet Explorer (icon on desktop)
    CSIDL_PROGRAMS = &H2                'Start Menu\Programs
    CSIDL_CONTROLS = &H3                'My Computer\Control Panel
    CSIDL_PRINTERS = &H4                'My Computer\Printers
    CSIDL_PERSONAL = &H5                'My Documents
    CSIDL_FAVORITES = &H6               '{user}\Favourites
    CSIDL_STARTUP = &H7                 'Start Menu\Programs\Startup
    CSIDL_RECENT = &H8                  '{user}\Recent
    CSIDL_SENDTO = &H9                  '{user}\SendTo
    CSIDL_BITBUCKET = &HA               '{desktop}\Recycle Bin
    CSIDL_STARTMENU = &HB               '{user}\Start Menu
    CSIDL_DESKTOPDIRECTORY = &H10       '{user}\Desktop
    CSIDL_DRIVES = &H11                 'My Computer
    CSIDL_NETWORK = &H12                'Network Neighbourhood
    CSIDL_NETHOOD = &H13                '{user}\nethood
    CSIDL_FONTS = &H14                  'windows\fonts
    CSIDL_TEMPLATES = &H15
    CSIDL_COMMON_STARTMENU = &H16       'All Users\Start Menu
    CSIDL_COMMON_PROGRAMS = &H17        'All Users\Programs
    CSIDL_COMMON_STARTUP = &H18         'All Users\Startup
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19 'All Users\Desktop
    CSIDL_APPDATA = &H1A                '{user}\Application Data
    CSIDL_PRINTHOOD = &H1B              '{user}\PrintHood
    CSIDL_LOCAL_APPDATA = &H1C          '{user}\Local Settings\Application Data (non roaming)
    CSIDL_ALTSTARTUP = &H1D             'non localized startup
    CSIDL_COMMON_ALTSTARTUP = &H1E      'non localized common startup
    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_COOKIES = &H21
    CSIDL_HISTORY = &H22
    CSIDL_COMMON_APPDATA = &H23          'All Users\Application Data
    CSIDL_WINDOWS = &H24                 'GetWindowsDirectory()
    CSIDL_SYSTEM = &H25                  'GetSystemDirectory()
    CSIDL_PROGRAM_FILES = &H26           'C:\Program Files
    CSIDL_MYPICTURES = &H27              'C:\Program Files\My Pictures
    CSIDL_PROFILE = &H28                 'USERPROFILE
    CSIDL_SYSTEMX86 = &H29               'x86 system directory on RISC
    CSIDL_PROGRAM_FILESX86 = &H2A        'x86 C:\Program Files on RISC
    CSIDL_PROGRAM_FILES_COMMON = &H2B    'C:\Program Files\Common
    CSIDL_PROGRAM_FILES_COMMONX86 = &H2C 'x86 Program Files\Common on RISC
    CSIDL_COMMON_TEMPLATES = &H2D        'All Users\Templates
    CSIDL_COMMON_DOCUMENTS = &H2E        'All Users\Documents
    CSIDL_COMMON_ADMINTOOLS = &H2F       'All Users\Start Menu\Programs\Administrative Tools
    CSIDL_ADMINTOOLS = &H30              '{user}\Start Menu\Programs\Administrative Tools
End Enum
'=======================================================
'Type de média (disque dur)
'=======================================================
Public Enum MediaType
    Unknown = 0
    F5_1Pt2_512 = 1
    F3_1Pt44_512 = 2
    F3_2Pt88_512 = 3
    F3_20Pt8_512 = 4
    F3_720_512 = 5
    F5_360_512 = 6
    F5_320_512 = 7
    F5_320_1024 = 8
    F5_180_512 = 9
    F5_160_512 = 10
    Removable = 11
    FixedMedia = 12
    F3_120M_512 = 13
    F3_640_512 = 14
    F5_640_512 = 15
    F5_720_512 = 16
    F3_1Pt2_512 = 17
    F3_1Pt23_1024 = 18
    F5_1Pt23_1024 = 19
    F3_128Mb_512 = 20
    F3_230Mb_512 = 21
    F8_256_128 = 22
    F3_200Mb_512 = 23
    F3_240M_512 = 24
    F3_32M_512 = 25
End Enum
'=======================================================
'Dates d'un fichier
'=======================================================
Public Type FileDates
    DateCreated As String
    DateLastAccessed As String
    DateLastModified As String
End Type
'=======================================================
'Dates d'un fichier
'=======================================================
Public Enum FileDatesEnum
    DateCreated
    DateLastAccessed
    DateLastModified
End Enum
'=======================================================
'Tailles d'un fichier
'=======================================================
Public Type FileSizes
    FileSize As Currency
    CompressedFileSize As Currency
End Type
'=======================================================
'Taille d'un icone
'=======================================================
Public Enum IconSize
    Size16
    Size32
End Enum
'=======================================================
'Géométrie d'un disque dur
'=======================================================
Public Type DiskGeometry
   Cylinders As Currency
   MediaType  As Long
   TracksPerCylinder As Long
   SectorsPerTrack As Long
   BytesPerSector As Long
   strMediaType As String
End Type
'=======================================================
'Tailles diverses pour un disque logique
'=======================================================
Public Type DriveSizes
    SectorPerCluster As Long
    FreeClusters  As Long
    TotalClusters  As Long
    BytesPerSector  As Long
    TotalSpace  As Currency
    UsedClusters As Long
    BytesPerCluster As Long
    FreeSpace As Currency
    UsedSpace As Currency
End Type
'=======================================================
'Type de découpe
'=======================================================
Public Type CutMethod
    tMethode As CutMethodEnum
    lParam As Long
End Type
Public Enum CutMethodEnum
    [Fixed Size]
    [Fixed Number Of Files]
End Enum
'=======================================================
'Infos sur la partition
'=======================================================
Public Type PartitionInfo
    PartitionLength As Currency
    HiddenSectors As Long
    StartingOffset As Currency
End Type
'=======================================================
'Type et enums sur la recherche par critère
'=======================================================
Public Enum SearchCriteria
    [SearchSize] = 2
    [SearchName] = 4
    [SearchDateCreated] = 8
    [SearchDateLastAccessed] = 16
    [SearchDateLastModified] = 32
End Enum
Public Enum Operator
    [ > ]
    [ >= ]
    [ < ]
    [ <= ]
    [ = ]
End Enum
Public Type SearchDefinition
    FolderName As String
    SearchName As String
    RespectCase As Boolean
    Criteria As SearchCriteria
    SizeOperator As Operator
    SizeValue As Currency
    DateCreatedOperator As Operator
    DateCreatedValue As String
    DateLastAccessedOperator As Operator
    DateLastAccessedValue As String
    DateLastModifiedOperator As Operator
    DateLastModifiedValue As String
    SubFolder As Boolean
End Type
'=======================================================
'Infos sur le volume
'=======================================================
Public Type VolumeInfo
    VolumeName As String
    VolumeSerialNumber As String
    FileSystemName As String
End Type
'=======================================================
'Type de partition
'=======================================================
Public Enum PARTITION_TYPE
    PARTITION_ENTRY_UNUSED = &H0
    PARTITION_FAT_12 = &H1
    PARTITION_XENIX_1 = &H2
    PARTITION_XENIX_2 = &H3
    PARTITION_FAT_16 = &H4
    PARTITION_EXTENDED = &H5
    PARTITION_HUGE = &H6
    PARTITION_IFS = &H7
    PARTITION_OS2BOOTMGR = &HA
    PARTITION_FAT32 = &HB
    PARTITION_FAT32_XINT13 = &HC
    PARTITION_XINT13 = &HE
    PARTITION_XINT13_EXTENDED = &HF
    PARTITION_PREP = &H41
    PARTITION_LDM = &H42
    PARTITION_UNIX = &H63
End Enum
'=======================================================
'Attribut des fichiers
'=======================================================
Public Enum FILE_ATTRIBUTE
    FILE_ATTRIBUTE_ARCHIVE = &H20
    FILE_ATTRIBUTE_HIDDEN = &H2
    FILE_ATTRIBUTE_NORMAL = &H80
    FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = &H2000
    FILE_ATTRIBUTE_OFFLINE = &H1000
    FILE_ATTRIBUTE_READONLY = &H1
    FILE_ATTRIBUTE_SYSTEM = &H4
    FILE_ATTRIBUTE_TEMPORARY = &H100
    FILE_ATTRIBUTE_INVISIBLE_SYSTEM = 39
End Enum
'=======================================================
'Résultat d'une comparaison de fichiers
'=======================================================
Public Enum FileComparisonResult
    [Files are identical]
    [Files are different]
End Enum
'=======================================================
'Flags pour les boites de dialogue
'=======================================================
Public Enum OPENFILENAME_FLAGS
    OFN_READONLY = &H1
    OFN_OVERWRITEPROMPT = &H2
    OFN_HIDEREADONLY = &H4
    OFN_NOCHANGEDIR = &H8
    OFN_SHOWHELP = &H10
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_NOVALIDATE = &H100
    OFN_ALLOWMULTISELECT = &H200
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_PATHMUSTEXIST = &H800
    OFN_CREATEPROMPT = &H2000
    OFN_SHAREAWARE = &H4000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_SHAREFALLTHROUGH = 2
    OFN_SHARENOWARN = 1
    OFN_SHAREWARN = 0
    OFN_FILEMUSTEXIST = &H1000
    OFN_EXPLORER = &H80000
End Enum

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Private Type BY_HANDLE_FILE_INFORMATION
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    dwVolumeSerialNumber As Long
    nFileSizeHigh As Long
    nFileSizeLow As Long
    nNumberOfLinks As Long
    nFileIndexHigh As Long
    nFileIndexLow As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
End Type
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type
Private Type Fichier    'utilisé pour envoyer vers la corbeille
     hWnd As Long
     wFunc As Long
     pFrom As String
     pTo As String
     fFlags As Integer
     fAnyOperationsAborted As Boolean
     hNameMappings As Long
     lpszProgressTitle As String
End Type
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
'Structure contenant des infos sur une image
'Normalement cette structure contient une union mais VB ne les gère pas
Private Type PICTDESC
    cbSizeofStruct As Long  ' Taille de cette structure
    picType As Long  ' Type d'image : ICON, BITMAP, METAFILE, ENHMETAFILE
    ' Début union
    hImage As Long  ' Handle de l'image
    xExt As Long  ' Taille x de l'image pour une METAFILE
    yExt As Long  ' Taille y de l'image pour une METAFILE
End Type
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As Long
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
End Type
Public Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersion As Long
    dwFileVersionMS As Long
    dwFileVersionLS As Long
    dwProductVersionMS As Long
    dwProductVersionLS As Long
    dwFileFlagsMask As Long
    dwFileFlags As Long
    dwFileOS As Long
    dwFileType As Long
    dwFileSubtype As Long
    dwFileDateMS As Long
    dwFileDateLS As Long
End Type
Public Type FileVersionInfos
    FileVersion As String
    FileDescription As String
    CompanyName As String
    InternalName As String
    Copyright As String
    OriginalFileName As String
    ProductName As String
    ProductVersion As String
End Type
Private Type DISK_GEOMETRY
   Cylinders         As Currency
   MediaType         As Long
   TracksPerCylinder As Long
   SectorsPerTrack   As Long
   BytesPerSector    As Long
End Type
Private Type OVERLAPPED
    ternal As Long
    ternalHigh As Long
    Offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type
Private Type PARTITION_INFORMATION
  StartingOffset As Currency
  PartitionLength As Currency
  HiddenSectors As Long
  PartitionNumber As Currency
  PartitionType As PARTITION_TYPE
  BootIndicator As Byte
  RecognizedPartition As Byte
  RewritePartition As Byte
End Type
Private Type IDEREGS
    bFeaturesReg As Byte ' // Used for specifying SMART "commands".
    bSectorCountReg As Byte ' // IDE sector count register
    bSectorNumberReg As Byte ' // IDE sector number register
    bCylLowReg As Byte ' // IDE low order cylinder value
    bCylHighReg As Byte ' // IDE high order cylinder value
    bDriveHeadReg As Byte ' // IDE drive/head register
    bCommandReg As Byte ' // Actual IDE command.
    bReserved As Byte ' // reserved for future use. Must be zero.
End Type
Private Type SENDCMDINPARAMS
    cBufferSize As Long ' Buffer size in bytes
    irDriveRegs As IDEREGS ' Structure with drive register values.
    bDriveNumber As Byte ' Physical drive number to send command to(0,1,2,3).
    bReserved(2) As Byte ' Bytes reserved
    dwReserved(3) As Long ' DWORDS reserved
    bBuffer() As Byte ' Input buffer.
End Type
Private Type IDSECTOR
    wGenConfig As Integer
    wNumCyls As Integer
    wReserved As Integer
    wNumHeads As Integer
    wBytesPerTrack As Integer
    wBytesPerSector As Integer
    wSectorsPerTrack As Integer
    wVendorUnique(2) As Integer
    sSerialNumber(19) As Byte
    wBufferType As Integer
    wBufferSize As Integer
    wECCSize As Integer
    sFirmwareRev(7) As Byte
    sModelNumber(39) As Byte
    wMoreVendorUnique As Integer
    wDoubleWordIO As Integer
    wCapabilities As Integer
    wReserved1 As Integer
    wPIOTiming As Integer
    wDMATiming As Integer
    wBS As Integer
    wNumCurrentCyls As Integer
    wNumCurrentHeads As Integer
    wNumCurrentSectorsPerTrack As Integer
    ulCurrentSectorCapacity As Long
    wMultSectorStuff As Integer
    ulTotalAddressableSectors As Long
    wSingleWordDMA As Integer
    wMultiWordDMA As Integer
    bReserved(127) As Byte
End Type


'=======================================================
'APIS
'=======================================================
Private Declare Function CompareFileTime Lib "kernel32" (lpFileTime1 As Currency, lpFileTime2 As Currency) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function SetVolumeLabelAPI Lib "kernel32" Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long
Private Declare Function EncryptFileAPI Lib "ADVAPI32" Alias "EncryptFileA" (ByVal lpFileName As String) As Long
Private Declare Function DecryptFileAPI Lib "ADVAPI32" Alias "DecryptFileA" (ByVal lpFileName As String, ByVal dwReserved As Long) As Long
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare Function GetFileTypeAPI Lib "kernel32" Alias "GetFileType" (ByVal hFile As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32.dll" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function DeviceIoControl Lib "kernel32.dll" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, lpOverlapped As OVERLAPPED) As Long
Private Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function GetDriveTypeAPI Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function DeviceIoControl2 Lib "kernel32" Alias "DeviceIoControl" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function RemoveDirectory Lib "kernel32" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function LockFile Lib "kernel32.dll" (ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToLockLow As Long, ByVal nNumberOfBytesToLockHigh As Long) As Long
Private Declare Function UnlockFile Lib "kernel32.dll" (ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToUnlockLow As Long, ByVal nNumberOfBytesToUnlockHigh As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function SetFilePointerEx Lib "kernel32" (ByVal hFile As Long, ByVal liDistanceToMove As Currency, ByRef lpNewFilePointer As Currency, ByVal dwMoveMethod As Long) As Long
Private Declare Sub OleCreatePictureIndirect Lib "oleaut32.dll" (ByRef lpPictDesc As PICTDESC, ByRef riid As GUID, ByVal fOwn As Long, ByRef lplpvObj As IPictureDisp)
Private Declare Function ExtractAssociatedIcon Lib "Shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Private Declare Function ImageList_GetIcon Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal Flags&) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Boolean
Private Declare Function SHGetFileInfo Lib "Shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function SHFileOperation Lib "Shell32.dll" Alias "SHFileOperationA" (lpFileOp As Fichier) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "Shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetFileSizeEx Lib "kernel32" (ByVal hFile As Long, lpFileSize As Currency) As Boolean
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function DeleteFileAPI Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileSpec As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare Function PathIsRoot Lib "shlwapi.dll" Alias "PathIsRootA" (ByVal pszPath As String) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function GetFileInformationByHandle Lib "kernel32" (ByVal hFile As Long, lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
Private Declare Function FindExecutable Lib "Shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Declare Function GetCompressedFileSizeAPI Lib "kernel32" Alias "GetCompressedFileSizeA" (ByVal lpFileName As String, lpFileSizeHigh As Long) As Long
Private Declare Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function PathFindExtension Lib "shlwapi" Alias "PathFindExtensionA" (ByVal pPath As String) As Long
Private Declare Function lstrcpyA Lib "kernel32" (ByVal RetVal As String, ByVal ptr As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal ptr As Any) As Long
Private Declare Sub PathStripPath Lib "shlwapi.dll" Alias "PathStripPathA" (ByVal pszPath As String)
Private Declare Function PathRemoveFileSpec Lib "shlwapi" Alias "PathRemoveFileSpecA" (ByVal pPath As String) As Long
Private Declare Function PathRemoveBackslash Lib "shlwapi.dll" Alias "PathRemoveBackslashA" (ByVal pszPath As String) As Long
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function ShellExecuteEX Lib "Shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long
Private Declare Function SHRunDialog Lib "shell32" Alias "#61" (ByVal hOwner As Long, ByVal Unknown1 As Long, ByVal Unknown2 As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal uFlags As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function MoveFileAPI Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long


Private IEvent As New ISearchEvent  'pour libérer l'event de l'interface
Private varDates(2) As Currency     'pour les dates pour la recherche


'=======================================================
'//Propriétés de la classe
'=======================================================
Public Property Get Drives() As Drives
    Set Drives = Me.EnumDrives(True)
End Property
Public Property Get PhysicalDisks() As PhysicalDisks
    Set PhysicalDisks = Me.EnumPhysicalDisks(True)
End Property

'=======================================================
'Renvoie le nombre de disques physiques
'=======================================================
Public Property Get NumberOfPhysicalDrives() As Byte
Dim bytVar As Byte
Dim hDrive As Long

    bytVar = 0: hDrive = 1

    Do While hDrive <> INVALID_HANDLE_VALUE
        'tente d'ouvrir les Drives 0,1,2,3,.....
        's'arrête à la première erreur (hdrive=-1)
        hDrive = CreateFile("\\.\PHYSICALDRIVE" & CStr(bytVar), GENERIC_READ, _
            FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
        
        Call CloseHandle(hDrive)
        bytVar = bytVar + 1
    Loop
    
    NumberOfPhysicalDrives = bytVar - 1
End Property



'=======================================================
'//Méthodes
'=======================================================

'=======================================================
'Copie un fichier
'=======================================================
Public Function CopyFile(ByVal Source As String, ByVal Destination As String) As Long
Dim tFile As Fichier

    'définit l'objet tFile
    With tFile
        .wFunc = FO_COPY
        .pFrom = Source
        .pTo = Destination
        .fFlags = FOF_ALLOWUNDO And FOF_SIMPLEPROGRESS
    End With
    
    'procède à la copie
    CopyFile = SHFileOperation(tFile)
    
End Function

'=======================================================
'Copie un dossier
'=======================================================
Public Function CopyFolder(ByVal Source As String, ByVal Destination As String) As Long
Dim tFolder As Fichier
    
    'vire le dernier \
    If Right$(Source, 1) = "\" Then Source = Left$(Source, Len(Source) - 1)
    
    'définit l'objet tFolder
    With tFolder
        .wFunc = FO_COPY
        .pFrom = Source
        .pTo = Destination
        .fFlags = FOF_ALLOWUNDO And FOF_SIMPLEPROGRESS
    End With
    
    'procède à la copie
    CopyFolder = SHFileOperation(tFolder)
End Function

'=======================================================
'Créé un dossier
'=======================================================
Public Function CreateFolder(ByVal Path As String) As Long
Dim secSec As SECURITY_ATTRIBUTES

    CreateFolder = CreateDirectory(Path, secSec)
End Function

'=======================================================
'Fonction permettant d'ouvrir un fichier
'=======================================================
Public Function ShellOpenFile(ByVal File As String, ByVal hWnd As Long, _
    Optional Parameters As String = vbNullString, Optional ByVal Directory _
    As String = vbNullString) As Long
    
    ShellOpenFile = ShellExecute(hWnd, "open", File, Parameters, _
        Directory, 1)
End Function

'=======================================================
'Créé un fichier vide
'=======================================================
Public Function CreateEmptyFile(ByVal Path As String, Optional ByVal OverWrite As _
    Boolean = False) As Long
    
Dim PARAM As Long
    
    If OverWrite Then
        'fichier déjà existant ==> overwrite
        PARAM = CREATE_ALWAYS
    Else
        PARAM = CREATE_NEW
    End If
    
    'créé le fichier
    CreateEmptyFile = CreateFile(Path, GENERIC_WRITE, FILE_SHARE_WRITE Or _
        FILE_SHARE_READ, ByVal 0&, PARAM, ByVal 0&, ByVal 0&)
    
    CloseHandle CreateEmptyFile 'ferme le handle...
End Function

'=======================================================
'Supprime un fichier
'=======================================================
Public Function DeleteFile(ByVal Source As String) As Long

    SetFileAttributes Source, FILE_ATTRIBUTE_NORMAL
    'permet de virer un attribut qui empecherait la suppression
    
    DeleteFile = DeleteFileAPI(Source)
End Function

'=======================================================
'Supprime un dossier (vierge)
'=======================================================
Public Function DeleteEmptyFolder(ByVal Source As String) As Long
    DeleteEmptyFolder = RemoveDirectory(Source)
End Function

'=======================================================
'Renvoie true si le fichier existe
'=======================================================
Public Function FileExists(ByVal Source As String) As Boolean
    FileExists = CBool(PathFileExists(Source) And Not (Me.FolderExists(Source)))
End Function

'=======================================================
'Renvoie true si le dossier existe
'=======================================================
Public Function FolderExists(ByVal Source As String) As Boolean
    FolderExists = CBool(PathIsDirectory(Source))
End Function

'=======================================================
'Renvoie true si le drive existe
'=======================================================
Public Function DriveExists(ByVal Source As String) As Boolean
    DriveExists = FolderExists(Source)
End Function

'=======================================================
'Récupère les infos sur un fichier
'=======================================================
Public Function GetFile(ByVal Path As String) As File

    Set GetFile = New File
    
    Call GetFile.SetPath(Path, True)
    
End Function

'=======================================================
'Récupère les infos sur un drive
'=======================================================
Public Function GetDrive(ByVal DriveLetter As String) As Drive

    Set GetDrive = New Drive
    
    Call GetDrive.SetVolumeLetter(DriveLetter, True)
End Function

'=======================================================
'Récupère le nom du drive depuis le path
'=======================================================
Public Function GetDriveName(ByVal Path As String) As String
    GetDriveName = Left$(Path, 1) & ":\"
End Function

'=======================================================
'Récupère le nom du dossier depuis le path
'=======================================================
Public Function GetFolderName(ByVal Path As String) As String

    Call PathRemoveFileSpec(Path)
    
    GetFolderName = StringWithoutNullChar(Path)
    
    If Right$(GetFolderName, 1) <> "\" Then GetFolderName = GetFolderName & "\"
End Function

'=======================================================
'Récupère le nom du fichier depuis le path
'=======================================================
Public Function GetFileName(ByVal Path As String) As String

    Call PathStripPath(Path)
    
    GetFileName = StringWithoutNullChar(Path)
    
End Function

'=======================================================
'Récupère l'extension du fichier depuis un path
'=======================================================
Public Function GetFileExtension(ByVal Path As String) As String

    GetFileExtension = String$(lstrlenA(ByVal PathFindExtension(Path)), 0)
    
    Call lstrcpyA(ByVal GetFileExtension, ByVal PathFindExtension(Path))
    
End Function

'=======================================================
'Récupère les infos sur le dossier
'=======================================================
Public Function GetFolder(ByVal Path As String) As Folder

    Set GetFolder = New Folder
    
    Call GetFolder.SetPath(Path, True)
    
End Function

'=======================================================
'Récupère le dossier parent (string)
'=======================================================
Public Function GetParentFolderName(ByVal Path As String) As String
Dim l As Long

    Call PathRemoveBackslash(Path)
    
    l = InStrRev(StringWithoutNullChar(Path), "\")
    
    If l = 0 Then GetParentFolderName = vbNullString Else _
        GetParentFolderName = Left$(Path, l)
End Function

'=======================================================
'Récupère un path spécial
'=======================================================
Public Function GetSpecialFolder(ByVal SpecialFolder As SpecialFolderConst) As String
Dim r As Long
Dim sPath As String
Dim IDL As ITEMIDLIST

    'obtient le path
    r = SHGetSpecialFolderLocation(100, SpecialFolder, IDL)
    
    If r = 0 Then
    
        'création d'un buffer
        sPath = Space$(512)
        
        'obtient le path depuis l'itemidlist
        r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
        
        'formate le texte
        GetSpecialFolder = Left$(sPath, InStr(sPath, vbNullChar) - 1)
        
        Exit Function
        
    End If
    
    GetSpecialFolder = vbNullString
End Function

'=======================================================
'Déplace des fichiers
'=======================================================
Public Function MoveFile(ByVal Source As String, ByVal Destination As String) As Long
Dim tFile As Fichier

    'définit l'objet tFile
    With tFile
        .wFunc = FO_MOVE
        .pFrom = Source
        .pTo = Destination
        .fFlags = FOF_ALLOWUNDO And FOF_SIMPLEPROGRESS
    End With
    
    'procède au déplacement
    MoveFile = SHFileOperation(tFile)
End Function

'=======================================================
'Déplace des dossiers
'=======================================================
Public Function MoveFolder(ByVal Source As String, ByVal Destination As String) As Long
Dim tFolder As Fichier

    'vire le dernier \
    If Right$(Source, 1) = "\" Then Source = Left$(Source, Len(Source) - 1)
    
    'définit l'objet tFolder
    With tFolder
        .wFunc = FO_MOVE
        .pFrom = Source
        .pTo = Destination
        .fFlags = FOF_ALLOWUNDO And FOF_SIMPLEPROGRESS
    End With
    
    'procède au déplacement
    MoveFolder = SHFileOperation(tFolder)
End Function

'=======================================================
'Obtient des infos sur un disque physique
'=======================================================
Public Function GetPhysicalDisk(ByVal DiskNumber As Byte) As PhysicalDisk

    Set GetPhysicalDisk = New PhysicalDisk
    
    Call GetPhysicalDisk.SetDiskNumber(DiskNumber, True)
End Function

'=======================================================
'Affiche la boite de dialogue 'Choix du dossier'
'=======================================================
Public Function BrowseForFolder(ByVal Title As String, hWnd As Long) As String
Dim lIDList As Long
Dim sBuffer As String
Dim tBrowseInfo As BrowseInfo

    'affectation des propriétés définies par les arguments de la fonction
    With tBrowseInfo
        .hwndOwner = hWnd
        .lpszTitle = Title
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    
    'affichage de la boite de dialogue
    lIDList = SHBrowseForFolder(tBrowseInfo)
    
    'récupération du résultat
    If lIDList Then
        sBuffer = String$(260, vbNullChar)
        SHGetPathFromIDList lIDList, sBuffer
        BrowseForFolder = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        CoTaskMemFree lIDList
    End If
    
End Function

'=======================================================
'Vérifie que la partition est accessible
'=======================================================
Public Function IsDriveAvailable(ByVal DriveLetter As String) As Boolean
Dim hDrive As Long
Dim strDrive As String
'temp variables
Dim a As Long, b As Long, c As Long, e As Long
    
    IsDriveAvailable = False

    'obtient un path valide de drive
    strDrive = "\\.\" & UCase$(DriveLetter) & ":"
    
    '//obtient le handle du drive
    hDrive = CreateFile(strDrive, GENERIC_READ, FILE_SHARE_READ Or _
        FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    
    'récupère des infos sur la taille
    Call GetDiskFreeSpace(Right$(strDrive, 2) & "\", a, e, b, c)
    
    If c = 0 Or e = 0 Or a = 0 Or hDrive = INVALID_HANDLE_VALUE Then
        Call CloseHandle(hDrive)
        Exit Function
    End If

    IsDriveAvailable = True

    Call CloseHandle(hDrive)
End Function

'=======================================================
'Vérifie que la partition est accessible
'=======================================================
Public Function IsPhysicalDiskAvailable(ByVal DiskNumber As Byte) As Boolean
Dim hDrive As Long
Dim bytesReturned As Long
Dim strDrive As String
Dim tGeom As DISK_GEOMETRY
Dim ovrTemp As OVERLAPPED

    'obtient un path valide de drive
    strDrive = "\\.\PHYSICALDRIVE" & CStr(DiskNumber)

    '//obtient le handle du drive
    hDrive = CreateFile(strDrive, GENERIC_READ, FILE_SHARE_READ Or _
        FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)

    'obtient les infos souhaitées par l'API DeviceIoControl ; ControlCode=Geometry
    Call DeviceIoControl(hDrive, IOCTL_DISK_GET_DRIVE_GEOMETRY, ByVal 0&, _
        0&, tGeom, Len(tGeom), bytesReturned, ovrTemp)
    
    If bytesReturned = 0 Or hDrive = INVALID_HANDLE_VALUE Or _
        tGeom.BytesPerSector = 0 Or tGeom.Cylinders = 0 Or _
        tGeom.SectorsPerTrack = 0 Or tGeom.TracksPerCylinder = 0 _
        Then
        Call CloseHandle(hDrive)
        Exit Function
    End If

    IsPhysicalDiskAvailable = True

    CloseHandle hDrive
End Function

'=======================================================
'Change les attributs d'un fichier
'=======================================================
Public Function ChangeFileAttributes(ByVal Source As String, ByVal NewAttributes _
    As FileAttributes)
    
    ChangeFileAttributes = SetFileAttributes(Source, NewAttributes)
End Function

'=======================================================
'Renvoie true si le fichier est dispo
'=======================================================
Public Function IsFileAvailable(ByVal Path As String) As Boolean
Dim lFile As Long
    
    'obtient un handle vers le fichier à ouvrir
    lFile = CreateFile(Path, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, _
        ByVal 0&, OPEN_EXISTING, 0, 0)
    
    If lFile <> INVALID_HANDLE_VALUE Then
    
        'fichier existant
        IsFileAvailable = True
        
        'referme le handle
        CloseHandle lFile
        
    End If

End Function

'=======================================================
'Renvoie true si le dossier est dispo
'=======================================================
Public Function IsFolderAvailable(ByVal Folder As String) As Boolean
Dim lFolder As Long
    
    'obtient un handle vers le fichier à ouvrir
    lFolder = CreateFile(Folder, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, _
        ByVal 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0)
    
    If lFolder <> INVALID_HANDLE_VALUE Then
    
        'fichier existant
        IsFolderAvailable = True
        
        'referme le handle
        CloseHandle lFolder
        
    End If

End Function

'=======================================================
'Récupère la géométrie d'un disque
'=======================================================
Public Function GetPhysicalDiskGeometry(ByVal DiskNumber As Byte) As DiskGeometry
Dim hDrive As Long
Dim strDrive As String
Dim tGeom As DISK_GEOMETRY
Dim bytesReturned As Long

    strDrive = "\\.\PHYSICALDRIVE" & CStr(DiskNumber)
    
    '//obtient le handle du disque
    hDrive = CreateFile(strDrive, GENERIC_READ Or _
        GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0, 0)
    
    If hDrive = INVALID_HANDLE_VALUE Then Exit Function
    
    
    '//obtient les infos sur la géométrie du disque
        Dim ovrTemp As OVERLAPPED
        
        'obtient les infos souhaitées par l'API DeviceIoControl ; ControlCode=Geometry
        Call DeviceIoControl(hDrive, IOCTL_DISK_GET_DRIVE_GEOMETRY, ByVal 0&, _
            0&, tGeom, Len(tGeom), bytesReturned, ovrTemp)
        
        With tGeom
            GetPhysicalDiskGeometry.BytesPerSector = .BytesPerSector
            GetPhysicalDiskGeometry.Cylinders = 10000 * .Cylinders
            GetPhysicalDiskGeometry.MediaType = .MediaType
            GetPhysicalDiskGeometry.TracksPerCylinder = .TracksPerCylinder
            GetPhysicalDiskGeometry.SectorsPerTrack = .SectorsPerTrack
            GetPhysicalDiskGeometry.strMediaType = GetMediaType(.MediaType)
        End With
    
    'referme le handle
    Call CloseHandle(hDrive)
End Function

'=======================================================
'Récupère la géométrie d'un drive
'=======================================================
Public Function GetDriveGeometry(ByVal DriveLetter As String) As DiskGeometry
Dim strDrive As String
Dim hDrive As Long
Dim bytesReturned As Long
Dim ovrTemp As OVERLAPPED

    strDrive = "\\.\" & UCase$(DriveLetter) & ":"
    
    '//obtient le handle du drive
    hDrive = CreateFile(strDrive, GENERIC_READ, FILE_SHARE_READ Or _
        FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    
    If hDrive = INVALID_HANDLE_VALUE Then Exit Function
    
    '//obtient les infos sur la géométrie du disque
        Dim tGeom As DISK_GEOMETRY
        
        'obtient les infos souhaitées par l'API DeviceIoControl ; ControlCode=Geometry
        Call DeviceIoControl(hDrive, IOCTL_DISK_GET_DRIVE_GEOMETRY, ByVal 0&, _
            0&, tGeom, Len(tGeom), bytesReturned, ovrTemp)
        
        With tGeom
            GetDriveGeometry.BytesPerSector = .BytesPerSector
            GetDriveGeometry.Cylinders = 10000 * .Cylinders
            GetDriveGeometry.MediaType = .MediaType
            GetDriveGeometry.strMediaType = GetMediaType(.MediaType)
            GetDriveGeometry.TracksPerCylinder = .TracksPerCylinder
            GetDriveGeometry.SectorsPerTrack = .SectorsPerTrack
        End With
    
    'referme le handle
    Call CloseHandle(hDrive)
End Function

'=======================================================
'Renvoie diverses infos sur les tailles d'un disque
'=======================================================
Public Function GetDriveSizes(ByVal DriveLetter As String) As DriveSizes
Dim strDrive As String
Dim tGeom As DiskGeometry
Dim hDrive As Long
Dim a As Long, b As Long, c As Long, e As Long

    strDrive = "\\.\" & UCase$(DriveLetter) & ":"
    
    '//obtient le handle du drive
    hDrive = CreateFile(strDrive, GENERIC_READ, FILE_SHARE_READ Or _
        FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    
    If hDrive = INVALID_HANDLE_VALUE Then Exit Function
    
    'récupère quelques info géométriques
    tGeom = Me.GetDriveGeometry(DriveLetter)
    
    'calcule les tailles
    Call GetDiskFreeSpace(Right$(strDrive, 2) & "\", a, e, b, c)
    
    With GetDriveSizes
        .SectorPerCluster = a
        .FreeClusters = b
        .TotalClusters = c
        .BytesPerSector = e
        .TotalSpace = CCur(CCur(.BytesPerSector) * CCur(.SectorPerCluster) * _
            CCur(.TotalClusters))
        .UsedClusters = .TotalClusters - .FreeClusters
        .BytesPerCluster = CLng(.BytesPerSector * .SectorPerCluster)
        .FreeSpace = CCur(CCur(.BytesPerSector) * CCur(.SectorPerCluster) * _
            CCur(.FreeClusters))
        .UsedSpace = .TotalSpace - .FreeSpace
    End With
    
End Function

'=======================================================
'Récupère le type de disque
'=======================================================
Public Function GetDriveType(ByVal DiskLetter As String) As Long
    GetDriveType = GetDriveTypeAPI(DiskLetter & ":\")
End Function

'=======================================================
'Récupère les infos partition d'un disque
'=======================================================
Public Function GetDrivePartitionInfo(ByVal DriveLetter As String) As PartitionInfo
Dim tPart As PARTITION_INFORMATION
Dim ovrTemp As OVERLAPPED
Dim bytesReturned As Long
Dim strDrive As String
Dim hDrive As Long
    
    strDrive = "\\.\" & UCase$(DriveLetter) & ":"
    
    '//obtient le handle du drive
    hDrive = CreateFile(strDrive, GENERIC_READ, FILE_SHARE_READ Or _
        FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    
    If hDrive = INVALID_HANDLE_VALUE Then Exit Function
    
    'ControlCode=IOCTL_DISK_GET_PARTITION_INFO
    Call DeviceIoControl(hDrive, IOCTL_DISK_GET_PARTITION_INFO, ByVal 0&, 0&, _
        tPart, Len(tPart), bytesReturned, ovrTemp)
    
    With tPart
        GetDrivePartitionInfo.PartitionLength = 10000 * .PartitionLength
        GetDrivePartitionInfo.HiddenSectors = .HiddenSectors
        GetDrivePartitionInfo.StartingOffset = .StartingOffset * 10000
    End With
    
    'ferme le handle
    Call CloseHandle(hDrive)
    
End Function

'=======================================================
'Récupère les infos sur le volume
'=======================================================
Public Function GetDriveVolumeInfo(ByVal DriveLetter As String) As VolumeInfo
Dim strDrive As String
Dim hDrive As Long
Dim s As String
Dim s2 As String
Dim d As Long

    strDrive = "\\.\" & UCase$(DriveLetter) & ":"
    
    '//obtient le handle du drive
    hDrive = CreateFile(strDrive, GENERIC_READ, FILE_SHARE_READ Or _
        FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    
    If hDrive = INVALID_HANDLE_VALUE Then Exit Function
    
    'buffer
    s = Space$(255)
    s2 = s
        
    'get infos
    Call GetVolumeInformation(Right$(strDrive, 2) & "\", s, 255, d, 0, 0, s2, 255)
    
    With GetDriveVolumeInfo
        If InStr(1, s, vbNullChar) > 1 Then .VolumeName = Left$(s, InStr(1, s, vbNullChar) - 1)
        If InStr(1, s2, vbNullChar) > 1 Then .FileSystemName = Left$(s2, InStr(1, s2, vbNullChar) - 1)
        .VolumeSerialNumber = d
    End With
    
    'ferme le handle
    Call CloseHandle(hDrive)
End Function

'=======================================================
'Renvoie true si le path est un drive
'=======================================================
Public Function IsRoot(ByVal Path As String) As Boolean
    IsRoot = CBool(PathIsRoot(Path))
End Function

'=======================================================
'Renomme le fichier/dossier
'=======================================================
Public Function Rename(ByVal Path As String, ByVal NewName As String) As Long
    Rename = MoveFileAPI(Path, NewName)
End Function

'=======================================================
'Sauvegarde la string dans un fichier à la suite
'=======================================================
Public Function SaveDataInFile(ByVal Path As String, ByVal strStringToWrite As _
    String, Optional ByVal DeleteActualFile As Boolean = False)
    
Dim lFile As Long
Dim Ret As Long
    
    'si le fichier est inexistant ou bien si on doit l'overwrite
    If Me.FileExists(Path) = False Or DeleteActualFile = True Then
        Call Me.CreateEmptyFile(Path, True)
    End If

    'écrit dans le fichier
    'obtient un handle vers le fichier à écrire
    'ouverture en ECRITURE, avec overwrite si déjà existant (car déjà demandé confirmation avant)
    lFile = CreateFile(Path, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, _
        0&, OPEN_EXISTING, 0, 0)

    If lFile = INVALID_HANDLE_VALUE Then Exit Function 'fichier indisponible

    'bouge le pointeur sur le fichier à la fin du fichier
    Ret = SetFilePointerEx(lFile, 0&, 0&, FILE_END)

    'écriture dans le fichier
    SaveDataInFile = WriteFile(lFile, ByVal strStringToWrite, Len(strStringToWrite), Ret, _
        ByVal 0&)

    'ferme le handle du fichier écrit
    CloseHandle lFile
    
End Function

'=======================================================
'Affiche les propriétés windows d'un fichier
'=======================================================
Public Function ShowFileProperty(ByVal Path As String, ByVal hWnd As Long) As Long
Dim SEI As SHELLEXECUTEINFO
    
    'définition des paramètres de la boite de dialogue
    With SEI
        'défnition
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or _
            SEE_MASK_FLAG_NO_UI
        .hWnd = hWnd
        .lpVerb = "properties"
        .lpFile = Path
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = 0
        .lpIDList = 0
    End With
    
    ShowFileProperty = ShellExecuteEX(SEI)   'affichage
    
End Function

'=======================================================
'Affiche les propriétés windows d'un dossier
'=======================================================
Public Function ShowFolderProperty(ByVal Path As String, ByVal hWnd As Long) As Long
    ShowFolderProperty = Me.ShowFileProperty(Path, hWnd)
End Function

'=======================================================
'Affiche les propriétés windows d'un drive
'=======================================================
Public Function ShowDriveProperty(ByVal DriveLetter As String, ByVal hWnd As Long) As Long
    ShowDriveProperty = Me.ShowFileProperty(DriveLetter & ":\", hWnd)
End Function

'=======================================================
'Affichage de la boite de dialogue Ouvrir
'Renvoie le FileName
'Multiples fichiers stockés en dernier paramètre Byref, de 1 à Ubound
'=======================================================
Public Function ShowOpen(ByVal Title As String, ByVal hWnd As Long, _
    Optional ByVal Filter As String, Optional ByVal InitDir As String, _
    Optional ByVal InitFile As String, Optional ByRef HasPushedCancel As _
    Boolean, Optional ByVal DefIndex As Long, Optional ByVal Flags As _
    OPENFILENAME_FLAGS, Optional ByRef Files As Variant) As String
    
Dim OFName As OPENFILENAME
Dim s() As String
Dim s2 As String
Dim x As Long

    With OFName
    
        'taille du type
        .lStructSize = Len(OFName)
        
        'affecte les valeurs des variables de configuration de la boite de dialogue au type
        .hwndOwner = hWnd
        .hInstance = App.hInstance
        Filter = Replace$(Filter, "|", vbNullChar) 'convertit la string
        .nMaxCustFilter = 0 ' Index de filtre personnalisé (non géré)
        .nFilterIndex = DefIndex  ' Index du filtre à utiliser par défaut
        .lpstrFilter = Filter
        
        .lpstrFile = Left$(InitFile & String$(2048, vbNullChar), 2048)
        .nMaxFile = Len(.lpstrFile) - 1 ' Longueur du nom de fichier
        .lpstrFileTitle = .lpstrFile ' Nom et extension du fichier (sans chemin)
        .nMaxFileTitle = .nMaxFile ' Taille de la chaîne précédente
        .lpstrInitialDir = InitDir ' Répertoire initial
        .lpstrTitle = Title ' Titre de la fenêtre
        
        .Flags = Flags
    End With

    'ouvre la boite de dialogue
    HasPushedCancel = Not (CBool(GetOpenFileName(OFName))) 'récupère si l'utilisateur à cliqué sur "annuler"
    If HasPushedCancel Then Exit Function
    
    'récupère le FileName
    ShowOpen = Left$(OFName.lpstrFile, InStr(OFName.lpstrFile, vbNullChar) - 1)
    
    If IsMissing(Files) Then Exit Function
           
           
    '//récupère tous les items sélectionnés
    s2 = OFName.lpstrFile
    
    'récupère les multiples fichiers ==> vire les derniers vbnullchar
    Do
        s2 = Left$(s2, Len(s2) - 1)
    Loop While Right$(s2, 1) = vbNullChar
    
    'créé un tableau avec les fichiers
    s() = Split(s2, vbNullChar, , vbBinaryCompare)
    
    ReDim Files(0) As String
    For x = 1 To UBound(s())
        s(x) = s(0) & "\" & s(x) 'ajoute le path aux fichiers
        ReDim Preserve Files(x) As String   'agrandit le tableau
        Files(x) = s(x) 'stocke
    Next x

End Function

'=======================================================
'Affichage de la boite de dialogue Enregistrer...
'Renvoie le FileName AVEC l'extension choisie
'(ne double pas l'extension si comprise dans le nom)
'=======================================================
Public Function ShowSave(ByVal Title As String, ByVal hWnd As Long, _
    Optional ByVal Filter As String, Optional ByVal InitDir As String, _
    Optional ByRef HasPushedCancel As Boolean) As String
    
Dim OFName As OPENFILENAME
Dim s As String
Dim s2() As String
Dim s3 As String
    
    With OFName
        'taille du type
        .lStructSize = Len(OFName)
        'définit les paramètres au type
        .hwndOwner = hWnd
        .hInstance = App.hInstance
        .lpstrFilter = Replace$(Filter, "|", vbNullChar) 'convertit la string
        .nMaxFile = 255
        .nMaxFileTitle = 255
        .lpstrFileTitle = Space$(254)
        .lpstrInitialDir = InitDir
        .lpstrTitle = Title
        .lpstrFile = Space$(254)    'créé un buffer pour le fichier. Ne pas modifier 254.
    End With

    'affichage de la boite de dialogue
    HasPushedCancel = Not (CBool(GetSaveFileName(OFName)))
    
    'récupère le filename formaté
    s3 = Left$(OFName.lpstrFile, InStr(OFName.lpstrFile, vbNullChar) - 1)
    
    'récupère le filename
    'on récupère l'extension choisie
    s2() = Split(Filter, "|", , vbBinaryCompare)
    s = Right$(s2(OFName.nFilterIndex * 2 - 1), Len(s2(OFName.nFilterIndex * 2 - 1)) - InStrRev(s2(OFName.nFilterIndex * 2 - 1), "*", , vbBinaryCompare))
    
    If Not LCase$(Right$(s3, Len(s))) = LCase$(s) Then
        'alors on ajoute le filtre
        ShowSave = s3 & s
    Else
        'pas besoin de filtre
        ShowSave = s3
    End If
End Function

'=======================================================
'Affiche la boite de dialogue Run
'=======================================================
Public Function ShowRunBox(ByVal hWnd As Long, ByVal Title As String, _
    ByVal Message As String) As Long
    
    ShowRunBox = SHRunDialog(hWnd, 0, 0, StrConv(Title, vbUnicode), _
        StrConv(Message, vbUnicode), 0)
        
End Function

'=======================================================
'Génère un nom de fichier temporaire
'=======================================================
Public Function GetTempName(ByVal FileName As String) As String
Dim sBuf As String
Dim s As String

    '//obtient le répertoire temporaire
    sBuf = String$(256, vbNullChar)
    
    'obtient le dossier temp
    Call GetTempPath(256, sBuf)
    
    'formate le path
    sBuf = Left$(sBuf, InStr(sBuf, vbNullChar) - 1)
    
    '//obtient un path unique
    'buffer
    s = String$(256, vbNullChar)
    
    'obtient le dossier temp
    Call GetTempFileName(sBuf, FileName, 0, s)
    
    'formate le path
    s = Left$(s, InStr(s, vbNullChar) - 1)

    GetTempName = s
    
End Function

'=======================================================
'Déplace les objets vers la corbeille
'=======================================================
Public Function MoveToTrash(ByVal Source As String) As Long
Dim tObj As Fichier

    'définit l'objet tObj
    With tObj
        .wFunc = FO_DELETE
        .pFrom = Source
        .fFlags = FOF_ALLOWUNDO
    End With
    
    'procède à la suppression
    MoveToTrash = SHFileOperation(tObj)
End Function

'=======================================================
'Ecrit dans un fichier
'=======================================================
Public Function WriteFileString(ByVal Source As String, ByVal strString As _
    String, Optional ByVal StartingOffset As Currency = 0, Optional ByVal _
    WriteToEnd As Boolean = False) As String

Dim Ret As Long
Dim lFile As Long
       
    'obtient un handle vers le fichier à écrire
    'ouverture en ECRITURE, avec overwrite si déjà existant (car déjà demandé confirmation avant)
    lFile = CreateFile(Source, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, _
        0&, OPEN_EXISTING, 0, 0)

    If lFile = INVALID_HANDLE_VALUE Then Exit Function 'fichier indisponible
    
    If WriteToEnd = False Then
        'bouge le pointeur sur le fichier au bon emplacement
        Ret = SetFilePointerEx(lFile, StartingOffset / 10000, 0&, FILE_BEGIN)
    Else
        'bouge le pointeur sur le fichier à la fin du fichier
        Ret = SetFilePointerEx(lFile, 0&, 0&, FILE_END)
    End If

    'écriture dans le fichier
    WriteFileString = WriteFile(lFile, ByVal strString, Len(strString), Ret, ByVal 0&)

    'ferme le handle du fichier écrit
    CloseHandle lFile
End Function

'=======================================================
'Ecrit dans un drive
'=======================================================
Public Function WritePhysicalDiskString(ByVal DiskNumber As Byte, ByVal StartingSector As _
    Currency, ByVal nBytes As Long, ByVal BytesPerSector As Long, ByVal _
    strStringToWrite As String) As Long

'/!\ StartingSector et nbytes doivent être des multiples de la taille d'un secteur (généralement 512 octets)
Dim Pointeur As Currency
Dim Ret As Long
Dim sDisk As String
Dim hDevice As Long
Dim lLowPart As Long
Dim lHighPart As Long
  
    'obtient un path valide pour l'API CreateFIle si nécessaire
    sDisk = "\\.\PHYSICALDRIVE" & CStr(DiskNumber)

    'ouvre le drive
    hDevice = CreateFile(sDisk, GENERIC_WRITE, FILE_SHARE_READ Or _
        FILE_SHARE_WRITE, 0&, OPEN_EXISTING, FILE_FLAG_NO_BUFFERING, 0&)
   
    'quitte si le handle n'est pas valide
    If hDevice = INVALID_HANDLE_VALUE Then Exit Function
   
    'détermine le byte de départ du secteur
    Pointeur = CCur(StartingSector) * CCur(BytesPerSector)
    
    'transforme un currency en 2 long pour une structure LARGE_INTEGER
    Call GetLargeInteger(Pointeur, lLowPart, lHighPart)

    'déplace, dans le fichier (ici un disque) pointé par hDevice, le "curseur" au premier
    'byte que l'on veut lire (donné par deux long)
    Ret = SetFilePointer(hDevice, lLowPart, lHighPart, FILE_BEGIN)  'FILE_BEGIN ==> part du début du fichier pour décompter la DistanceToMove
    If Ret = INVALID_HANDLE_VALUE Then Exit Function
    
    'verrouilage de la zone du disque à écrire
    Call LockFile(hDevice, lLowPart, lHighPart, nBytes, 0)
    
    'écriture disque
    WritePhysicalDiskString = WriteFile(hDevice, ByVal strStringToWrite, nBytes, Ret, _
        ByVal 0&)
    
    'on vide les buffers internes et on dévérouille la zone
    Call FlushFileBuffers(hDevice)
    Call UnlockFile(hDevice, lLowPart, lHighPart, nBytes, 0)

    'ferme le handle
    CloseHandle hDevice
End Function

'=======================================================
'Ecrit dans un disque physique
'=======================================================
Public Function WriteDriveString(ByVal DriveLetter As String, ByVal _
    StartingSector As Currency, ByVal nBytes As Long, ByVal BytesPerSector As Long, ByVal _
    strStringToWrite As String) As Long

'/!\ StartingSector et nbytes doivent être des multiples de la taille d'un secteur (généralement 512 octets)
Dim Pointeur As Currency
Dim Ret As Long
Dim sDrive As String
Dim hDevice As Long
Dim lLowPart As Long
Dim lHighPart As Long
  
    'obtient un path valide pour l'API CreateFIle si nécessaire
    sDrive = "\\.\" & UCase$(DriveLetter) & ":"

    'ouvre le drive
    hDevice = CreateFile(sDrive, GENERIC_WRITE, FILE_SHARE_READ Or _
        FILE_SHARE_WRITE, 0&, OPEN_EXISTING, FILE_FLAG_NO_BUFFERING, 0&)
   
    'quitte si le handle n'est pas valide
    If hDevice = INVALID_HANDLE_VALUE Then Exit Function
   
    'détermine le byte de départ du secteur
    Pointeur = CCur(StartingSector) * CCur(BytesPerSector)
    
    'transforme un currency en 2 long pour une structure LARGE_INTEGER
    Call GetLargeInteger(Pointeur, lLowPart, lHighPart)

    'déplace, dans le fichier (ici un disque) pointé par hDevice, le "curseur" au premier
    'byte que l'on veut lire (donné par deux long)
    Ret = SetFilePointer(hDevice, lLowPart, lHighPart, FILE_BEGIN)  'FILE_BEGIN ==> part du début du fichier pour décompter la DistanceToMove
    If Ret = INVALID_HANDLE_VALUE Then Exit Function
    
    'verrouilage de la zone du disque à écrire
    Call LockFile(hDevice, lLowPart, lHighPart, nBytes, 0)
    
    'écriture disque
    WriteDriveString = WriteFile(hDevice, ByVal strStringToWrite, nBytes, _
        Ret, ByVal 0&)
    
    'on vide les buffers internes et on dévérouille la zone
    Call FlushFileBuffers(hDevice)
    Call UnlockFile(hDevice, lLowPart, lHighPart, nBytes, 0)

    'ferme le handle
    CloseHandle hDevice

End Function

'=======================================================
'Lit dans un fichier
'=======================================================
Public Function ReadFileString(ByVal Source As String, ByVal Size As Currency, _
    ByVal StartingOffset As Currency) As String
    
Dim tmpText As String
Dim Ret As Long
Dim lFile As Long
    
    'obtient un handle vers le fichier à ouvrir
    lFile = CreateFile(Source, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, _
        ByVal 0&, OPEN_EXISTING, 0, 0)
    
    If lFile = INVALID_HANDLE_VALUE Then
        'fichier inexistant, ou en tout cas inaccessible
        Exit Function
    End If
        
    'créé un buffer qui contiendra les valeurs
    tmpText = String$(Size, 0)

    'bouge le pointeur sur lr fichier au bon emplacement
    Ret = SetFilePointerEx(lFile, StartingOffset / 10000, 0&, FILE_BEGIN)  'divise par 10000 pour
    'pouvoir renvoyer une currency DECIMALE (cad du genre 1.4567 pour l'offset 14567)
    
    'prend un morceau du fichier
    Ret = ReadFile(lFile, ByVal tmpText, Len(tmpText), Ret, ByVal 0&)
    
    'affecte à la fonction
    ReadFileString = tmpText

    'referme le handle
    CloseHandle lFile
    
End Function

'=======================================================
'Lit dans un drive
'=======================================================
Public Function ReadDriveString(ByVal DriveLetter As String, ByVal StartingSector _
    As Currency, ByVal nBytes As Long, ByVal BytesPerSector As Long) As String
    
'/!\ StartingSector et nbytes doivent être des multiples de la taille
'd'un secteur (généralement 512 octets)
Dim BytesRead As Long
Dim Pointeur As Currency
Dim Ret As Long
Dim hDevice As Long
Dim lLowPart As Long
Dim sDrive As String
Dim lHighPart As Long
Dim sBufferOut As String
   
    'obtient un path valide pour l'API CreateFIle si nécessaire
    sDrive = "\\.\" & UCase$(DriveLetter) & ":"

    'ouvre le drive
    hDevice = CreateFile(sDrive, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
   
    'quitte si le handle n'est pas valide
    If hDevice = INVALID_HANDLE_VALUE Then Exit Function
   
    'détermine le byte de départ du secteur
    Pointeur = CCur(StartingSector) * CCur(BytesPerSector)
    
    'transforme un currency en 2 long pour une structure LARGE_INTEGER
    Call GetLargeInteger(Pointeur, lLowPart, lHighPart)

    'déplace, dans le fichier (ici un disque) pointé par hDevice, le "curseur" au premier
    'byte que l'on veut lire (donné par deux long)
    Ret = SetFilePointer(hDevice, lLowPart, lHighPart, FILE_BEGIN)   'FILE_BEGIN ==> part du début du fichier pour décompter la DistanceToMove
    If Ret = INVALID_HANDLE_VALUE Then
        'ferme le handle
        CloseHandle hDevice
        Exit Function
    End If
    
    'création d'un buffer
    sBufferOut = Space$(nBytes)

    'obtention de la string
    Ret = ReadFile(hDevice, ByVal sBufferOut, nBytes, BytesRead, 0&)

    'ferme le handle
    CloseHandle hDevice
    
    ReadDriveString = sBufferOut
End Function

'=======================================================
'Lit dans un disque physique
'=======================================================
Public Function ReadPhysicalDiskString(ByVal DiskNumber As Byte, ByVal _
    StartingSector As Currency, ByVal nBytes As Long, ByVal BytesPerSector As _
    Long) As String
    
'/!\ StartingSector et nbytes doivent être des multiples de la taille
'd'un secteur (généralement 512 octets)
Dim BytesRead As Long
Dim Pointeur As Currency
Dim Ret As Long
Dim hDevice As Long
Dim lLowPart As Long
Dim sDisk As String
Dim lHighPart As Long
Dim sBufferOut As String
   
    'obtient un path valide pour l'API CreateFIle si nécessaire
    sDisk = "\\.\PHYSICALDRIVE" & CStr(DiskNumber)

    'ouvre le drive
    hDevice = CreateFile(sDisk, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
   
    'quitte si le handle n'est pas valide
    If hDevice = INVALID_HANDLE_VALUE Then Exit Function
   
    'détermine le byte de départ du secteur
    Pointeur = CCur(StartingSector) * CCur(BytesPerSector)
    
    'transforme un currency en 2 long pour une structure LARGE_INTEGER
    Call GetLargeInteger(Pointeur, lLowPart, lHighPart)

    'déplace, dans le fichier (ici un disque) pointé par hDevice, le "curseur" au premier
    'byte que l'on veut lire (donné par deux long)
    Ret = SetFilePointer(hDevice, lLowPart, lHighPart, FILE_BEGIN)  'FILE_BEGIN ==> part du début du fichier pour décompter la DistanceToMove
    If Ret = INVALID_HANDLE_VALUE Then
        
        'ferme le handle
        CloseHandle hDevice
        Exit Function
    End If
    
    'création d'un buffer
    sBufferOut = Space$(nBytes)

    'obtention de la string
    Ret = ReadFile(hDevice, ByVal sBufferOut, nBytes, BytesRead, 0&)

    'ferme le handle
    CloseHandle hDevice
    
    ReadPhysicalDiskString = sBufferOut
End Function

'=======================================================
'Renvoie des infos sur le fichier EXE/DLL/OCX...
'=======================================================
Public Function GetFileVersionInfos(ByVal Source As String) As FileVersionInfos
    Dim Vffi As VS_FIXEDFILEINFO  ' version info structure
    Dim Buffer() As Byte          ' buffer for info ressource
    Dim pData As Long             ' pointer to info data
    Dim lDataLen As Long          ' length of info pointed at by pData
    Dim cpl(0 To 3) As Byte       ' buffer for code page & language
    Dim cplstr As String          ' 8-digit hex string of cpl
    Dim strVersionInfo As String
    ' Prépare la structure
    GetFileVersionInfos.CompanyName = ""
    GetFileVersionInfos.Copyright = ""
    GetFileVersionInfos.FileDescription = ""
    GetFileVersionInfos.FileVersion = ""
    GetFileVersionInfos.InternalName = ""
    GetFileVersionInfos.OriginalFileName = ""
    GetFileVersionInfos.ProductName = ""
    GetFileVersionInfos.ProductVersion = ""
    ' Contrôle si le fichier contient des informations récupérables
    lDataLen = GetFileVersionInfoSize(Source, pData)
    If lDataLen = 0 Then
        Exit Function
    End If
    ReDim Buffer(0 To lDataLen - 1) As Byte
    ' Get the version information resource.
    If GetFileVersionInfo(Source, 0, lDataLen, Buffer(0)) <> 0 Then
        ' Get a pointer to a structure that holds a bunch of data.
        If VerQueryValue(Buffer(0), "\", pData, lDataLen) <> 0 Then
            ' Copy that structure into the one we can access.
            MoveMemory Vffi, ByVal pData, lDataLen
            ' ---------------------------------------
            ' Récupération de la version du fichier
            GetFileVersionInfos.FileVersion = Trim$(Str$(HIWORD(Vffi.dwFileVersionMS))) & "." & _
                                            Trim$(Str$(LOWORD(Vffi.dwFileVersionMS))) & "." & _
                                            Trim$(Str$(HIWORD(Vffi.dwFileVersionLS))) & "." & _
                                            Trim$(Str$(LOWORD(Vffi.dwFileVersionLS)))
            ' Récupération du code page et du langage
            If VerQueryValue(Buffer(0), "\VarFileInfo\Translation", pData, lDataLen) <> 0 Then
                ' Copy that information into the byte array.
                MoveMemory cpl(0), ByVal pData, 4
                ' Convert those four bytes into a 8-digit hexadecimal string.
                cplstr = FixedHex(cpl(1), 2) & FixedHex(cpl(0), 2) & FixedHex(cpl(3), 2) & FixedHex(cpl(2), 2)
                ' cplstr contient maintenant le code page et le langage
                ' ---------------------------------------
                ' Récupération des autres infos
                strVersionInfo = "CompanyName"
                If VerQueryValue(Buffer(0), "\StringFileInfo\" & cplstr & "\" & strVersionInfo, pData, lDataLen) <> 0 Then
                        If lDataLen > 0 Then lDataLen = lDataLen - 1
                        GetFileVersionInfos.CompanyName = Space$(lDataLen)
                        lstrcpy GetFileVersionInfos.CompanyName, pData
                End If
                strVersionInfo = "FileDescription"
                If VerQueryValue(Buffer(0), "\StringFileInfo\" & cplstr & "\" & strVersionInfo, pData, lDataLen) <> 0 Then
                        If lDataLen > 0 Then lDataLen = lDataLen - 1
                        GetFileVersionInfos.FileDescription = Space$(lDataLen)
                        lstrcpy GetFileVersionInfos.FileDescription, pData
                End If
                strVersionInfo = "FileVersion"
                If VerQueryValue(Buffer(0), "\StringFileInfo\" & cplstr & "\" & strVersionInfo, pData, lDataLen) <> 0 Then
                        If lDataLen > 0 Then lDataLen = lDataLen - 1
                        GetFileVersionInfos.FileVersion = Space$(lDataLen)
                        lstrcpy GetFileVersionInfos.FileVersion, pData
                End If
                strVersionInfo = "InternalName"
                If VerQueryValue(Buffer(0), "\StringFileInfo\" & cplstr & "\" & strVersionInfo, pData, lDataLen) <> 0 Then
                        If lDataLen > 0 Then lDataLen = lDataLen - 1
                        GetFileVersionInfos.InternalName = Space$(lDataLen)
                        lstrcpy GetFileVersionInfos.InternalName, pData
                End If
                strVersionInfo = "LegalCopyright"
                If VerQueryValue(Buffer(0), "\StringFileInfo\" & cplstr & "\" & strVersionInfo, pData, lDataLen) <> 0 Then
                        If lDataLen > 0 Then lDataLen = lDataLen - 1
                        GetFileVersionInfos.Copyright = Space$(lDataLen)
                        lstrcpy GetFileVersionInfos.Copyright, pData
                End If
                strVersionInfo = "OriginalFileName"
                If VerQueryValue(Buffer(0), "\StringFileInfo\" & cplstr & "\" & strVersionInfo, pData, lDataLen) <> 0 Then
                        If lDataLen > 0 Then lDataLen = lDataLen - 1
                        GetFileVersionInfos.OriginalFileName = Space$(lDataLen)
                        lstrcpy GetFileVersionInfos.OriginalFileName, pData
                End If
                strVersionInfo = "ProductName"
                If VerQueryValue(Buffer(0), "\StringFileInfo\" & cplstr & "\" & strVersionInfo, pData, lDataLen) <> 0 Then
                        If lDataLen > 0 Then lDataLen = lDataLen - 1
                        GetFileVersionInfos.ProductName = Space$(lDataLen)
                        lstrcpy GetFileVersionInfos.ProductName, pData
                End If
                strVersionInfo = "ProductVersion"
                If VerQueryValue(Buffer(0), "\StringFileInfo\" & cplstr & "\" & strVersionInfo, pData, lDataLen) <> 0 Then
                        If lDataLen > 0 Then lDataLen = lDataLen - 1
                        GetFileVersionInfos.ProductVersion = Space$(lDataLen)
                        lstrcpy GetFileVersionInfos.ProductVersion, pData
                End If
            End If
        End If
    End If
    
    With GetFileVersionInfos
        .CompanyName = StringWithoutNullChar(.CompanyName)
        .Copyright = StringWithoutNullChar(.Copyright)
        .FileDescription = StringWithoutNullChar(.FileDescription)
        .FileVersion = StringWithoutNullChar(.FileVersion)
        .InternalName = StringWithoutNullChar(.InternalName)
        .OriginalFileName = StringWithoutNullChar(.OriginalFileName)
        .ProductName = StringWithoutNullChar(.ProductName)
        .ProductVersion = StringWithoutNullChar(.ProductVersion)
    End With
    
End Function

'=======================================================
'Renvoie les dates d'un fichier
'=======================================================
Public Function GetFileDates(ByVal Source As String) As FileDates
Dim lFile As Long
Dim FileInfo As BY_HANDLE_FILE_INFORMATION

    'handle vers le fichier désiré
    lFile = CreateFile(Source, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, _
        ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    
    If lFile = INVALID_HANDLE_VALUE Then Exit Function
    
    'obtient les infos sur le fichier
    Call GetFileInformationByHandle(lFile, FileInfo)
    
    'referme immédiatement le handle
    Call CloseHandle(lFile)

    With GetFileDates
        .DateCreated = FileTimeToString(FileInfo.ftCreationTime)
        .DateLastAccessed = FileTimeToString(FileInfo.ftLastAccessTime)
        .DateLastModified = FileTimeToString(FileInfo.ftLastWriteTime)
    End With

End Function

'=======================================================
'Renvoie une date de fichier sous forme de currency
'=======================================================
Public Function GetFileDateAsCurrency(ByVal Source As String, TypeOfDate As _
    FileDatesEnum) As Currency
    
Dim lFile As Long
Dim FileInfo As BY_HANDLE_FILE_INFORMATION
Dim MyTime As FILETIME

    'handle vers le fichier désiré
    lFile = CreateFile(Source, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, _
        ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    
    If lFile = INVALID_HANDLE_VALUE Then Exit Function
    
    'obtient les infos sur le fichier
    Call GetFileInformationByHandle(lFile, FileInfo)
    
    'referme immédiatement le handle
    Call CloseHandle(lFile)
    
    'ne pas oublier de convertir en local
    Select Case TypeOfDate
    
        Case DateCreated
            Call FileTimeToLocalFileTime(FileInfo.ftCreationTime, _
                FileInfo.ftCreationTime)
            MyTime = FileInfo.ftCreationTime
            
        Case DateLastAccessed
            Call FileTimeToLocalFileTime(FileInfo.ftLastAccessTime, _
                FileInfo.ftLastAccessTime)
            MyTime = FileInfo.ftLastAccessTime
            
        Case DateLastModified
            Call FileTimeToLocalFileTime(FileInfo.ftLastWriteTime, _
                FileInfo.ftLastWriteTime)
            MyTime = FileInfo.ftLastWriteTime
            
    End Select
    
    'convertit MyTime en Currency
    Call CopyMemory(GetFileDateAsCurrency, MyTime, LenB(MyTime))

End Function

'=======================================================
'Renvoie une date de dossier sous forme de currency
'=======================================================
Public Function GetFolderDateAsCurrency(ByVal Source As String, TypeOfDate As _
    FileDatesEnum) As Currency
    
Dim lFolder As Long
Dim FileInfo As BY_HANDLE_FILE_INFORMATION
Dim MyTime As FILETIME

    'handle vers le dossier désiré
    lFolder = CreateFile(Source, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or _
        FILE_SHARE_DELETE, 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0&)
    
    If lFolder = INVALID_HANDLE_VALUE Then Exit Function
    
    'obtient les infos sur le dossier
    Call GetFileInformationByHandle(lFolder, FileInfo)
    
    'referme immédiatement le handle
    Call CloseHandle(lFolder)
    
    'ne pas oublier de convertir en local
    Select Case TypeOfDate
    
        Case DateCreated
            Call FileTimeToLocalFileTime(FileInfo.ftCreationTime, _
                FileInfo.ftCreationTime)
            MyTime = FileInfo.ftCreationTime
            
        Case DateLastAccessed
            Call FileTimeToLocalFileTime(FileInfo.ftLastAccessTime, _
                FileInfo.ftLastAccessTime)
            MyTime = FileInfo.ftLastAccessTime
            
        Case DateLastModified
            Call FileTimeToLocalFileTime(FileInfo.ftLastWriteTime, _
                FileInfo.ftLastWriteTime)
            MyTime = FileInfo.ftLastWriteTime
            
    End Select
    
    'convertit MyTime en Currency
    Call CopyMemory(GetFolderDateAsCurrency, MyTime, LenB(MyTime))

End Function

'=======================================================
'Renvoie les dates d'un fichier
'=======================================================
Public Function GetFolderDates(ByVal Source As String) As FileDates
Dim lFolder As Long
Dim FileInfo As BY_HANDLE_FILE_INFORMATION

    'handle vers le dossier désiré
    lFolder = CreateFile(Source, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or _
        FILE_SHARE_DELETE, 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0&)
    
    If lFolder = INVALID_HANDLE_VALUE Then Exit Function
    
    'obtient les infos sur le fichier
    Call GetFileInformationByHandle(lFolder, FileInfo)
    
    'referme immédiatement le handle
    Call CloseHandle(lFolder)

    With GetFolderDates
        .DateCreated = FileTimeToString(FileInfo.ftCreationTime)
        .DateLastAccessed = FileTimeToString(FileInfo.ftLastAccessTime)
        .DateLastModified = FileTimeToString(FileInfo.ftLastWriteTime)
    End With

End Function

'=======================================================
'Renvoie la taille normale
'=======================================================
Public Function GetFileSize(ByVal Source As String) As Currency
Dim lngFile As Long
Dim curSize As Currency

    'obtient le handle du fichier
    lngFile = CreateFile(Source, GENERIC_READ, FILE_SHARE_WRITE Or FILE_SHARE_READ, _
        ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    
    If lngFile = INVALID_HANDLE_VALUE Then Exit Function
    
    'obtient la taille normale par API
    Call GetFileSizeEx(lngFile, curSize)
    
    'ferme le handle ouvert
    Call CloseHandle(lngFile)
        
    GetFileSize = curSize * 10000  'multiplie par 10^4 pour obtenir un nombre entier

End Function

'=======================================================
'Récupère la taille compressée
'=======================================================
Public Function GetCompressedFileSize(Source As String) As Currency
    GetCompressedFileSize = GetCompressedFileSizeAPI(Source, ByVal 0&)
End Function

'=======================================================
'Récupère l'attribut d'un fichier
'=======================================================
Public Function GetFileAttributes(ByVal Source As String) As FileAttributes
Dim lFile As Long
Dim FileInfo As BY_HANDLE_FILE_INFORMATION

    'handle vers le fichier désiré
    lFile = CreateFile(Source, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, _
        ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
        
    If lFile = INVALID_HANDLE_VALUE Then Exit Function
    
    'obtient les infos sur le fichier
    Call GetFileInformationByHandle(lFile, FileInfo)
    
    'referme le handle
    CloseHandle lFile
    
    GetFileAttributes = FileInfo.dwFileAttributes
    
End Function

'=======================================================
'Récupère l'attribut d'un dossier
'=======================================================
Public Function GetFolderAttributes(ByVal Source As String) As FileAttributes
Dim lFolder As Long
Dim FileInfo As BY_HANDLE_FILE_INFORMATION

    'handle vers le dossier désiré
    lFolder = CreateFile(Source, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or _
        FILE_SHARE_DELETE, 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0&)
        
    If lFolder = INVALID_HANDLE_VALUE Then Exit Function
    
    'obtient les infos sur le fichier
    Call GetFileInformationByHandle(lFolder, FileInfo)
    
    'referme le handle
    CloseHandle lFolder
    
    GetFolderAttributes = FileInfo.dwFileAttributes
    
End Function

'=======================================================
'Récupère le nom court
'=======================================================
Public Function GetShortName(ByVal Source As String, Optional ByVal Size As _
    Long = 164) As String
    
Dim sBuffer As String

    sBuffer = String$(Size, 0)   'buffer
    Call GetShortPathName(Source, sBuffer, Size)
    GetShortName = StringWithoutNullChar(sBuffer)
End Function

'=======================================================
'Récupère le path court
'=======================================================
Public Function GetShortPath(ByVal Source As String, Optional ByVal Size As _
    Long = 164) As String

Dim sBuffer As String

    sBuffer = String$(Size, 0)   'buffer
    GetShortPath = GetShortPathName(Me.GetFolderName(Source), sBuffer, Size)
    GetShortPath = StringWithoutNullChar(sBuffer)
End Function

'=======================================================
'Récupère le programme associé
'=======================================================
Public Function GetAssociatedProgram(ByVal Source As String) As String
Dim sBuf As String

    sBuf = String$(255, 0)  'création d'un buffer
    
    'trouve l'éxécutable associé
    Call FindExecutable(Source, vbNullString, sBuf)
    
    'formate la string
    GetAssociatedProgram = StringWithoutNullChar(sBuf)

End Function

'=======================================================
'Récupère le type de fichier
'=======================================================
Public Function GetFileType(ByVal Source As String) As String
Dim hFile As Long

    'handle du fichier
    hFile = CreateFile(Source, GENERIC_READ, FILE_SHARE_READ, 0&, _
        OPEN_EXISTING, 0&, 0&)
    
    'récupère le type
    Select Case GetFileTypeAPI(hFile)
        Case FILE_TYPE_UNKNOWN
            GetFileType = "Unknow"
        Case FILE_TYPE_DISK
            GetFileType = "Disk File"
        Case FILE_TYPE_CHAR
            GetFileType = "Character file"
        Case FILE_TYPE_PIPE
            GetFileType = "Pipe"
    End Select
    
    'referme le handle
    Call CloseHandle(hFile)
End Function

'=======================================================
'Récupère l'icone du fichier
'=======================================================
Public Function GetIcon(ByVal Source As String, ByVal Size As IconSize) As IPictureDisp
Dim hIco As Long
Dim hIcon As Long
Dim SHFinfo As SHFILEINFO
    
    If Size = Size16 Then
    
        'La variable qui va contenir l'identifiant de l'icône
        hIco = SHGetFileInfo(Source, 0&, SHFinfo, Len(SHFinfo), DJM_SMALLICON)
        
        hIcon = ImageList_GetIcon(hIco, SHFinfo.iIcon, 0)
        
        'convertit en IPictureDisp
        Set GetIcon = GetIconFromHandle(hIcon)
        
        'Efface l'icône qui a été extrait
        Call DestroyIcon(hIcon)
        
    ElseIf Size = Size32 Then
    
        'La variable qui va contenir l'identifiant de l'icône
        hIcon = ExtractAssociatedIcon(App.hInstance, Source, 2)
        
        Set GetIcon = GetIconFromHandle(hIcon)
        
        'Efface l'icône qui a été extrait
        Call DestroyIcon(hIcon)
    
    End If
    
End Function

'=======================================================
'Enumère les fichiers du dossier
'=======================================================
Public Function EnumFiles(ByVal Folder As String, Optional ByVal SubFolder As _
    Boolean = False, Optional ByVal RefreshInfos As Boolean = False) As Files
    
Dim x As Long
Dim s() As String

    'instancie la classe
    Set EnumFiles = New Files
    
    If Right$(Folder, 1) <> "\" Then Folder = Folder & "\"
    
    'lance l'énumération
    ReDim s(0)
    Call EnumFilesFromFolder(Folder, s(), True, SubFolder)
    
    'créé un objet dans la collection pour chaque item
    For x = 1 To UBound(s())
        Call EnumFiles.AddItem(s(x), RefreshInfos)
    Next x
    
End Function

'=======================================================
'Enumère les disques logiques
'=======================================================
Public Function EnumDrives(Optional ByVal RefreshInfos As Boolean = False) As Drives
Dim strBuffer As String
Dim x As Byte, lngLen As Long
Dim s() As String

    'créé un buffer qui recevra les noms de partition
    strBuffer = Space$(255)
    
    'obtient les noms
    lngLen = GetLogicalDriveStrings(255, strBuffer)
    
    'formate la string
    strBuffer = Left$(strBuffer, lngLen)
    
    'redimensionne le tableau
    lngLen = (lngLen \ 4) - 1
    
    'redimensionne le tableau
    ReDim s(lngLen)
    
    For x = 0 To lngLen
        s(x) = Mid$(strBuffer, 4 * x + 1, 3)
    Next x
    
    'ajoute à la collection
    Set EnumDrives = New Drives
    For x = 0 To lngLen
        Call EnumDrives.AddItem(Left$(s(x), 1), RefreshInfos)
    Next x

End Function

'=======================================================
'Enumère les disques physiques
'=======================================================
Public Function EnumPhysicalDisks(Optional ByVal RefreshInfos As Boolean = False) _
    As PhysicalDisks
    
Dim bytVar As Byte
Dim hDrive As Long

    bytVar = 0: hDrive = 1: ReDim bt(0)
    
    Set EnumPhysicalDisks = New PhysicalDisks

    Do While hDrive <> INVALID_HANDLE_VALUE
        'tente d'ouvrir les Drives 0,1,2,3,.....
        's'arrête à la première erreur (hdrive=-1)
        hDrive = CreateFile("\\.\PHYSICALDRIVE" & CStr(bytVar), GENERIC_READ, _
            FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
        
        If hDrive = INVALID_HANDLE_VALUE Then Exit Do
        
        Call EnumPhysicalDisks.AddItem(CInt(bytVar), RefreshInfos)
            
        Call CloseHandle(hDrive)
        bytVar = bytVar + 1
    Loop

End Function

'=======================================================
'Enumère les disques logiques sous forme de string
'De 0 à Ubound
'=======================================================
Public Function EnumDrivesStr() As String()
Dim strBuffer As String
Dim x As Byte, lngLen As Long
Dim s() As String

    'créé un buffer qui recevra les noms de partition
    strBuffer = Space$(255)
    
    'obtient les noms
    lngLen = GetLogicalDriveStrings(255, strBuffer)
    
    'formate la string
    strBuffer = Left$(strBuffer, lngLen)
    
    'redimensionne le tableau
    lngLen = (lngLen \ 4) - 1
    
    'redimensionne le tableau
    ReDim s(lngLen)
    
    For x = 0 To lngLen
        s(x) = Mid$(strBuffer, 4 * x + 1, 3)
    Next x
    
    EnumDrivesStr = s
End Function

'=======================================================
'Enumère les disques logiques physiques sous forme de string
'De 0 à Ubound
'=======================================================
Public Function EnumPhysicalDisksStr() As Byte()
Dim bytVar As Byte
Dim hDrive As Long
Dim bt() As Byte

    bytVar = 0: hDrive = 1: ReDim bt(0)
    
    Do While hDrive <> INVALID_HANDLE_VALUE
        'tente d'ouvrir les Drives 0,1,2,3,.....
        's'arrête à la première erreur (hdrive=-1)
        hDrive = CreateFile("\\.\PHYSICALDRIVE" & CStr(bytVar), GENERIC_READ, _
            FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
        
        If hDrive = INVALID_HANDLE_VALUE Then Exit Do
        
        ReDim Preserve bt(bytVar)
        bt(UBound(bt())) = bytVar
            
        Call CloseHandle(hDrive)
        bytVar = bytVar + 1
    Loop
    
    EnumPhysicalDisksStr = bt
End Function

'=======================================================
'Enumère les sous dossiers du dossier
'=======================================================
Public Function EnumFolders(ByVal Folder As String, Optional ByVal SubFolder As _
    Boolean = False, Optional ByVal RefreshInfos As Boolean = False) As Folders
    
Dim x As Long
Dim s() As String

    'instancie la classe
    Set EnumFolders = New Folders

    If Right$(Folder, 1) <> "\" Then Folder = Folder & "\"
    
    'lance l'énumération
    ReDim s(0)
    Call EnumFoldersPrivate(Folder, s(), True, SubFolder)
    
    'créé un objet dans la collection pour chaque item
    For x = 1 To UBound(s())
        Call EnumFolders.AddItem(s(x), RefreshInfos)
    Next x
    
End Function

'=======================================================
'Enumère les fichiers du dossier en renvoyant une liste string
'=======================================================
Public Function EnumFilesStr(ByVal Folder As String, Optional ByVal SubFolder As _
    Boolean = False) As String()
    
Dim s() As String
    
    If Right$(Folder, 1) <> "\" Then Folder = Folder & "\"
        
    'lance l'énumération
    ReDim s(0)
    Call EnumFilesFromFolder(Folder, s(), True, SubFolder)
    
    EnumFilesStr = s
    
End Function

'=======================================================
'Renvoie le nom de disque d'un disque dur physique
'=======================================================
Public Function GetPhysicalDiskName(ByVal DiskNumber As Long) As String
Dim SCIP As SENDCMDINPARAMS
Dim IDSEC As IDSECTOR
Dim bArrOut(OUTPUT_DATA_SIZE - 1) As Byte
Dim lpcbBytesReturned As Long
Dim hDrive As Long
    
    'récupère un handle valide sur ce drive physique
    hDrive = CreateFile("\\.\PHYSICALDRIVE" & CStr(DiskNumber), GENERIC_READ Or _
        GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0, 0)

    'affecte toutes les properties nécessaires à cette structure
    With SCIP
        .bDriveNumber = CByte(DiskNumber)
        .cBufferSize = IDENTIFY_BUFFER_SIZE
        With .irDriveRegs
            .bCylLowReg = 0
            .bFeaturesReg = 0
            .bCommandReg = &HEC
            .bSectorCountReg = 1
            .bSectorNumberReg = 1
            .bCylHighReg = 0
            .bDriveHeadReg = &HA0
        End With
    End With
    
    'on récupère les infos souhaitées dans la structure IDSEC gràce au DeviceIoControl
    If DeviceIoControl2(hDrive, DFP_RECEIVE_DRIVE_DATA, SCIP, Len(SCIP) - 4, bArrOut(0), _
        OUTPUT_DATA_SIZE, lpcbBytesReturned, ByVal 0&) Then

        CopyMemory IDSEC, bArrOut(16), Len(IDSEC)
        GetPhysicalDiskName = PermuteString(StrConv(IDSEC.sModelNumber, vbUnicode))
    End If
    
    'referme le handle
    Call CloseHandle(hDrive)
End Function

'=======================================================
'Enumère les sous dossiers du dossier en renvoyant une liste string
'=======================================================
Public Function EnumFoldersStr(ByVal Folder As String, Optional ByVal SubFolder As _
    Boolean = False) As String()
    
Dim s() As String

    If Right$(Folder, 1) <> "\" Then Folder = Folder & "\"
    
    'lance l'énumération
    ReDim s(0)
    Call EnumFoldersPrivate(Folder, s(), True, SubFolder)
    Dim o As Long
    o = UBound(s())
    EnumFoldersStr = s
    
End Function





'=======================================================
'Fonctions et procédures privées
'=======================================================

'=======================================================
'Enlève le NullChar
'=======================================================
Private Function StringWithoutNullChar(ByVal strString As String) As String
Dim lIn As Long
    
    lIn = InStr(strString, vbNullChar)
    
    If lIn Then StringWithoutNullChar = Left$(strString, lIn - 1) Else _
        StringWithoutNullChar = strString

End Function

'=======================================================
'Covertit une date FILETIME en string
'=======================================================
Private Function FileTimeToString(fDate As FILETIME) As String
Dim sDate As SYSTEMTIME
Dim sDay As String
Dim sMonth As String
Dim sYear As String
Dim sHour As String
Dim sMinute As String
Dim sSecond As String
Dim s As String

    'conversion en LocalFileTime (temps universel ==> temps local)
    FileTimeToLocalFileTime fDate, fDate
    
    'conversion en SystemTime
    FileTimeToSystemTime fDate, sDate
    
    'conversion en string vers un format du genre 24/04/2000 09:50:59
    sDay = Trim$(IIf(sDate.wDay < 10, "0" & Trim$(Str$(sDate.wDay)), Trim$(Str$(sDate.wDay))))
    sMonth = Trim$(IIf(sDate.wMonth < 10, "0" & Trim$(Str$(sDate.wMonth)), Trim$(Str$(sDate.wMonth))))
    sHour = Trim$(IIf(sDate.wHour < 10, "0" & Trim$(Str$(sDate.wHour)), Trim$(Str$(sDate.wHour))))
    sMinute = Trim$(IIf(sDate.wMinute < 10, "0" & Trim$(Str$(sDate.wMinute)), Trim$(Str$(sDate.wMinute))))
    sSecond = Trim$(IIf(sDate.wSecond < 10, "0" & Trim$(Str$(sDate.wSecond)), Trim$(Str$(sDate.wSecond))))
    sYear = sDate.wYear
    
    s = sDay & "/" & sMonth & "/" & sYear & " " & sHour & ":" & sMinute & ":" & sSecond
    
    FileTimeToString = s

End Function

'=======================================================
'Récupère une icone depuis un handle
'=======================================================
Private Function GetIconFromHandle(hIcon As Long) As IPictureDisp
'Le REFIID de IPictureDisp (=GUID)
Dim IID_IPictureDisp As GUID
'Infos sur l'icone
Dim lpIcon As PICTDESC

    'on met place l'IID de IPictureDisp
    IID_IPictureDisp.Data1 = &H7BF80981
    IID_IPictureDisp.Data2 = &HBF32
    IID_IPictureDisp.Data3 = &H101A
    IID_IPictureDisp.Data4(0) = &H8B
    IID_IPictureDisp.Data4(1) = &HBB
    IID_IPictureDisp.Data4(2) = &H0
    IID_IPictureDisp.Data4(3) = &HAA
    IID_IPictureDisp.Data4(4) = &H0
    IID_IPictureDisp.Data4(5) = &H30
    IID_IPictureDisp.Data4(6) = &HC
    IID_IPictureDisp.Data4(7) = &HAB
    
    ' On initialise les champs requis :
    lpIcon.cbSizeofStruct = Len(lpIcon)
    lpIcon.hImage = hIcon
    
    ' le type : ICON
    lpIcon.picType = PICTYPE_ICON
    
    'on demande la création d'une interface de type IPictureDisp pour le handle d'icône
    'et en indiquant de ne pas effacer le HICON quand l'interface est libéré
    OleCreatePictureIndirect lpIcon, IID_IPictureDisp, 0, GetIconFromHandle
    
End Function

'Création d'une chaine Hexadecimale pour représenter un nombre
Private Function FixedHex(ByVal hexval As Long, ByVal nDigits As Long) As String
    FixedHex = Right$("00000000" & Hex(hexval), nDigits)
End Function
Private Function HIWORD(ByVal dwValue As Long) As Long
Dim hexstr As String
    hexstr = Right$("00000000" & Hex$(dwValue), 8)
    HIWORD = CLng("&H" & Left$(hexstr, 4))
End Function
Private Function LOWORD(ByVal dwValue As Long) As Long
Dim hexstr As String
    hexstr = Right$("00000000" & Hex$(dwValue), 8)
    LOWORD = CLng("&H" & Right$(hexstr, 4))
End Function

'=======================================================
'divise une currency en 2 long ==> créé une LARGE_INTEGER
'=======================================================
Private Sub GetLargeInteger(ByVal curVar As Currency, ByRef lngLowPart As Long, _
    ByRef lngHighPart As Long)
    
Dim tblbyte(0 To 7) As Byte
Dim pt10000 As Currency

    pt10000 = curVar / 10000
    
    Call CopyMemory(tblbyte(0), pt10000, 8)
    Call CopyMemory(lngHighPart, tblbyte(4), 4)
    Call CopyMemory(lngLowPart, tblbyte(0), 4)
        
End Sub

'=======================================================
'Enumère les fichiers d'un dossier
'De 1 à ubound
'=======================================================
Private Sub EnumFilesFromFolder(ByVal Directory As String, Files() As String, _
    Optional Begin As Boolean = False, Optional SubFolder As Boolean = True)
    
'infos d'un fichier et handle d'énumération
Dim FileInfo As WIN32_FIND_DATA, hFind As Long
'limite de Files
Static ub As Long

    'si on commence la limite est 0, pas d'enregistrement dans le tableau
    If Begin = True Then ub = 0
    
    'ouvre le dossier pour récupérer la liste de ses fichiers
    hFind = FindFirstFile(Directory & "*", FileInfo)
    If hFind <> -1 Then
        'si le fichier est un dossier
        If (FileInfo.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            'si ce n'est ni . ni ..
            If InStr(FileInfo.cFileName, ".") <> 1 And SubFolder Then
                'on récupère la liste des fichiers de ce sous dossier
                EnumFilesFromFolder Directory & Mid$(FileInfo.cFileName, 1, InStr(FileInfo.cFileName, vbNullChar) - 1) & "\", Files, False
            End If
        'sinon on l'ajoute à la fin de la liste
        Else
            ub = ub + 1
            ReDim Preserve Files(ub)
            'chemin complet
            Files(ub) = Directory & Mid$(FileInfo.cFileName, 1, InStr(FileInfo.cFileName, vbNullChar) - 1)
        End If
        'tant qu'il y a des fichiers dans le dossier en cours
        Do While FindNextFile(hFind, FileInfo)
            'on peut arrêter à tout moment
            DoEvents
            
            'si le fichier est un dossier
            If (FileInfo.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
                'si ce n'est ni . ni ..
                If InStr(FileInfo.cFileName, ".") <> 1 And SubFolder Then
                    'on récupère la liste des fichiers de ce sous dossier
                    EnumFilesFromFolder Directory & Mid$(FileInfo.cFileName, 1, InStr(FileInfo.cFileName, vbNullChar) - 1) & "\", Files, False
                End If
            'sinon on l'ajoute à la fin de la liste
            Else
                ub = ub + 1
                ReDim Preserve Files(ub)
                'chemin complet
                Files(ub) = Directory & Mid$(FileInfo.cFileName, 1, InStr(FileInfo.cFileName, vbNullChar) - 1)
            End If
        Loop
    End If
    'ferme l'énumèration
    FindClose hFind
End Sub

'=======================================================
'Enumère les sous dossiers d'un dossier
'De 1 à ubound
'=======================================================
Private Sub EnumFoldersPrivate(ByVal Directory As String, s() As String, _
    Optional Begin As Boolean = False, Optional SubFolder As Boolean = True)
    
'infos d'un fichier et handle d'énumération
Dim FileInfo As WIN32_FIND_DATA, hFind As Long, sDir As String
'limite de Files
Static ub As Long

    'si on commence la limite est 0, pas d'enregistrement dans le tableau
    If Begin = True Then ub = 0
    
    'ouvre le dossier pour récupérer la liste de ses fichiers
    hFind = FindFirstFile(Directory & "*", FileInfo)
    If hFind <> -1 Then
        'si le fichier est un dossier
        If (FileInfo.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            'si ce n'est ni . ni ..
            If InStr(FileInfo.cFileName, ".") <> 1 And SubFolder Then
                'on récupère la liste des fichiers de ce sous dossier
                EnumFoldersPrivate Directory & Mid$(FileInfo.cFileName, 1, InStr(FileInfo.cFileName, vbNullChar) - 1) & "\", s, False
            End If
        End If
        'tant qu'il y a des fichiers dans le dossier en cours
        Do While FindNextFile(hFind, FileInfo)
            'on peut arrêter à tout moment
            DoEvents
            
            'si le fichier est un dossier
            If (FileInfo.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
                'si ce n'est ni . ni ..
                If InStr(FileInfo.cFileName, ".") <> 1 And SubFolder Then
                    'on récupère la liste des fichiers de ce sous dossier
                    EnumFoldersPrivate Directory & Mid$(FileInfo.cFileName, 1, InStr(FileInfo.cFileName, vbNullChar) - 1) & "\", s, False
                End If
                sDir = Directory & Mid$(FileInfo.cFileName, 1, InStr(FileInfo.cFileName, vbNullChar) - 1)
                If Right$(sDir, 3) <> "\.." Then
                    ub = ub + 1
                    ReDim Preserve s(ub)
                    'chemin complet
                    s(ub) = sDir
                End If
            End If
        Loop
    End If
    'ferme l'énumèration
    FindClose hFind
End Sub

'=======================================================
'Permet de permuter une string
'Exemple : BADCFE devient ABCDEF
'=======================================================
Private Function PermuteString(ByRef s As String) As String
Dim x As Long

    PermuteString = s
    
    For x = 1 To Len(s) Step 2
        Mid$(PermuteString, x, 2) = Mid$(s, x + 1, 1) & Mid$(s, x, 1)
    Next x
    
End Function

'=======================================================
'Renvoie un type de drive en String correspondant au type en long
'=======================================================
Private Function GetMediaType(ByVal bytType As Byte) As String

   Select Case bytType
      Case 0:               GetMediaType = "Format unknown"
      Case F5_1Pt2_512:     GetMediaType = "5.25, 1.2MB, 512 bytes/sector"
      Case F3_1Pt44_512:    GetMediaType = "3.5, 1.44MB, 512 bytes/sector"
      Case F3_2Pt88_512:    GetMediaType = "3.5, 2.88MB, 512 bytes/sector"
      Case F3_20Pt8_512:    GetMediaType = "3.5, 20.8MB, 512 bytes/sector"
      Case F3_720_512:      GetMediaType = "3.5, 720KB, 512 bytes/sector"
      Case F5_360_512:      GetMediaType = "5.25, 360KB, 512 bytes/sector"
      Case F5_320_512:      GetMediaType = "5.25, 320KB, 512 bytes/sector"
      Case F5_320_1024:     GetMediaType = "5.25, 320KB, 1024 bytes/sector"
      Case F5_180_512:      GetMediaType = "5.25, 180KB, 512 bytes/sector"
      Case F5_160_512:      GetMediaType = "5.25, 160KB, 512 bytes/sector"
      Case Removable:       GetMediaType = "Removable media other than floppy"
      Case FixedMedia:      GetMediaType = "Fixed hard disk"
      Case F3_120M_512:     GetMediaType = "3.5, 120M Floppy"
      Case F3_640_512:      GetMediaType = "3.5, 640KB, 512 bytes/sector"
      Case F5_640_512:      GetMediaType = "5.25, 640KB, 512 bytes/sector"
      Case F5_720_512:      GetMediaType = "5.25, 720KB, 512 bytes/sector"
      Case F3_1Pt2_512:     GetMediaType = "3.5, 1.2Mb, 512 bytes/sector"
      Case F3_1Pt23_1024:   GetMediaType = "3.5, 1.23Mb, 1024 bytes/sector"
      Case F5_1Pt23_1024:   GetMediaType = "5.25, 1.23MB, 1024 bytes/sector"
      Case F3_128Mb_512:    GetMediaType = "3.5 MO 128Mb, 512 bytes/sector"
      Case F3_230Mb_512:    GetMediaType = "3.5 MO 230Mb, 512 bytes/sector"
      Case F8_256_128:      GetMediaType = "8, 256KB, 128 bytes/sector"
      Case F3_200Mb_512:    GetMediaType = "3.5, 200M Floppy (HiFD)"
      Case F3_240M_512:     GetMediaType = "3.5, 240Mb Floppy (HiFD)"
      Case F3_32M_512:      GetMediaType = "3.5, 32Mb Floppy"
      Case Else:            GetMediaType = "Format unknown"
   End Select

End Function

'=======================================================
'Renvoie un type de drive en String correspondant au type en long
'=======================================================
Public Function GetStringDriveType(ByVal bytType As Byte) As String

    Select Case bytType
        Case 1: GetStringDriveType = "No Root Directory"
        Case 2: GetStringDriveType = "Removeable Disk"
        Case 3: GetStringDriveType = "Drive fixed"
        Case 4: GetStringDriveType = "Remote"
        Case 5: GetStringDriveType = "CD-ROM"
        Case 6: GetStringDriveType = "RamDisk"
        Case Else: GetStringDriveType = "Unknown"
    End Select
    
End Function

'=======================================================
'Télécharge un fichier depuis une URL
'=======================================================
Public Function DownloadFile(ByVal URL As String, ByVal FileToSave As String) _
    As Long
    
    DownloadFile = URLDownloadToFile(0, URL, FileToSave, 0, 0)
    
End Function

'=======================================================
'Encrypte un fichier
'=======================================================
Public Function EncryptFile(ByVal Source As String) As Long
    EncryptFile = EncryptFileAPI(Source)
End Function

'=======================================================
'Encrypte un dossier
'=======================================================
Public Function EncryptFolder(ByVal Source As String) As Long
    EncryptFolder = EncryptFileAPI(Source)
End Function

'=======================================================
'Décrypte un fichier
'=======================================================
Public Function DecryptFile(ByVal Source As String) As Long
    DecryptFile = DecryptFileAPI(Source, 0)
End Function

'=======================================================
'Décrypte un dossier
'=======================================================
Public Function DecryptFolder(ByVal Source As String) As Long
    DecryptFolder = DecryptFileAPI(Source, 0)
End Function

'=======================================================
'Récupère le contenu d'un fichier dans une string
'=======================================================
Public Function LoadFileInString(ByVal Source As String, Optional ByRef IsOk As Long) As String
Dim hFile As Long
Dim lSize As Long
Dim Ret As Long
Dim s As String
    
    'obtient le handle du fichier
    hFile = CreateFile(Source, GENERIC_READ, FILE_SHARE_READ Or _
        FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    
    'obtient sa taille pour créer un buffer
    lSize = Me.GetFileSize(Source)
 
    'initialise le pointeur sur le fichier (au début)
    Call SetFilePointer(hFile, 0, 0, FILE_BEGIN)
    
    'créé le buffer
    s = String$(lSize, 0)
    
    'charge le fichier
    Ret = ReadFile(hFile, ByVal s, lSize, Ret, ByVal 0&)
     
    LoadFileInString = s
    
    If Not (IsMissing(IsOk)) Then
        'alors on renvoie si la fonction a marché ou pas
        IsOk = IIf(Ret > 0 And hFile > 0, 1, 0)
    End If
    
    'ferme le handle
    CloseHandle hFile
End Function

'=======================================================
'Lance la sanitization d'un fichier... le pauvre...
'=======================================================
Public Function SanitizeFile(Source As String) As Long
Dim s As String
Dim hFile As Long
Dim curSize As Currency
Dim nbBuf As Long
Dim lLastSize As Long
Dim y As Long
Dim Tbl(2097151) As Byte
Dim cASM As CAsmProc
Dim sH55() As Byte
Dim sHAA() As Byte
Dim pAA As Long
Dim x As Long
Dim p55 As Long

    '//génère les tableaux
    
    'les redimensionne
    ReDim sH55(2097151)
    ReDim sHAA(2097151)
    
    'les remplit
    For x = 0 To 2097151
        sH55(x) = 85
        sHAA(x) = 170
    Next x
    
    'récupère les pointeurs
    pAA = VarPtr(sHAA(0))
    p55 = VarPtr(sH55(0))
    
    'instancie la classe
    Set cASM = New CAsmProc
    
    'on fait çà par buffer de 2Mo
    curSize = Me.GetFileSize(Source)
    
    If curSize < 2097152 Then
    
        'un seul buffer, on génère une string aléatoire unique
        Call cASM.bnAlloc2MoAlea(Tbl(0))
        
        'on récupère le handle du fichier
        hFile = CreateFile(Source, GENERIC_WRITE, FILE_SHARE_READ Or _
            FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
        
        If hFile = INVALID_HANDLE_VALUE Then
            SanitizeFile = -1
            Exit Function
        End If

        'on écrit dans le fichier
        '// &H55
        Call WriteBytesToFileHandle(hFile, p55, 0, curSize)

        '// &HAA
        Call WriteBytesToFileHandle(hFile, pAA, 0, curSize)

        '//random string
        Call WriteBytesToFileHandle(hFile, VarPtr(Tbl(0)), 0, curSize)

        'rend la main
        DoEvents
                        
        'referme le handle
        Call CloseHandle(hFile)
        
    Else
    
        'plusieurs buffers
        
        'calcule le nombre de buffers et la taille du dernier buffer
        nbBuf = Int(curSize / 2097152)
        lLastSize = curSize - 2097152 * nbBuf
        
        'on récupère le handle du fichier
        hFile = CreateFile(Source, GENERIC_WRITE, FILE_SHARE_READ Or _
            FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
            
        If hFile = INVALID_HANDLE_VALUE Then
            SanitizeFile = -1
            Exit Function
        End If
        
        For y = 1 To nbBuf

            'on récupère un pointeur sur une string de 2Mo
            Call cASM.bnAlloc2MoAlea(Tbl(0))
    
            'on écrit dans le fichier
            '// &H55
             Call WriteBytesToFileHandle(hFile, p55, (y - 1) * 2097152, _
                2097152)
    
            '// &HAA
            Call WriteBytesToFileHandle(hFile, pAA, (y - 1) * 2097152, _
                2097152)
    
            '//random string
            Call WriteBytesToFileHandle(hFile, VarPtr(Tbl(0)), (y - 1) * 2097152, 2097152)
    
            'rend la main de tps en tps
            DoEvents
        Next y
        
        's'occupe du dernier buffer (plus petit)
        Call cASM.bnAlloc2MoAlea(Tbl(0))

        '// &H55
        Call WriteBytesToFileHandle(hFile, p55, (nbBuf) * 2097152, _
            lLastSize)

        '// &HAA
        Call WriteBytesToFileHandle(hFile, pAA, (nbBuf) * 2097152, _
            lLastSize)

        '//random string
        Call WriteBytesToFileHandle(hFile, VarPtr(Tbl(0)), (nbBuf) * 2097152, lLastSize)
        
        'referme le handle
        Call CloseHandle(hFile)
        
        DoEvents

    End If

    Set cASM = Nothing
    
    SanitizeFile = 1
    
End Function

'=======================================================
'Lance la sanitization d'un drive... le pauvre...
'=======================================================
Public Function SanitizeDrive(DriveLetter As String) As Long
Dim cDriv As Drive
Dim x As Long
Dim secPerString As Long
Dim s As String
Dim pt As Long
Dim hDevice As Long
Dim bPerSec As Long
Dim curOp As Currency
Dim bMax As Long
Dim Tbl(2097151) As Byte
Dim cASM As CAsmProc
Dim sH55() As Byte
Dim sHAA() As Byte
Dim pAA As Long
Dim p55 As Long

    'vérifie que le disque est accessible
    If Me.IsDriveAvailable(DriveLetter) = False Then
        SanitizeDrive = -1
        Exit Function
    End If
    
    '//génère les tableaux
    
    'les redimensionne
    ReDim sH55(2097151)
    ReDim sHAA(2097151)
    
    'les remplit
    For x = 0 To 2097151
        sH55(x) = 85
        sHAA(x) = 170
    Next x
    
    'récupère les pointeurs
    pAA = VarPtr(sHAA(0))
    p55 = VarPtr(sH55(0))
    
    'instancie la classe
    Set cASM = New CAsmProc

    'récupère les infos sue le disque
    Set cDriv = Me.GetDrive(DriveLetter)

    'nombre de secteurs pour une string de 2Mo
    secPerString = 2097152 / cDriv.BytesPerSector

    'handle du disque
    hDevice = CreateFile("\\.\" & UCase$(DriveLetter) & ":", GENERIC_WRITE, _
        FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)

    bPerSec = cDriv.BytesPerSector
    bMax = Int(cDriv.TotalPhysicalSectors / secPerString) + 1 '+1

    'pour chaque secteur
    For x = 1 To bMax

        'on récupère un pointeur sur une string de 2Mo
        Call cASM.bnAlloc2MoAlea(Tbl(0))

        'calcul unique
        curOp = CCur(x * secPerString)

        'on écrit dans le disque
        '// &H55
        Call DirectWritePtHandle(hDevice, curOp, 2097152, bPerSec, p55)

        '// &HAA
        Call DirectWritePtHandle(hDevice, curOp, 2097152, bPerSec, pAA)

        '//random string
        Call DirectWritePtHandle(hDevice, curOp, 2097152, bPerSec, VarPtr(Tbl(0)))

        'rend la main de tps en tps
        If (x Mod 5) = 0 Then DoEvents

    Next x

    '//s'occupe de l'entête du disque
        Call cASM.bnAlloc2MoAlea(Tbl(0))

        'on écrit dans le disque
        '// &H55
        Call DirectWritePtHandle(hDevice, 0, 2097152, bPerSec, p55)

        '// &HAA
        Call DirectWritePtHandle(hDevice, 0, 2097152, bPerSec, pAA)

        '//random string
        Call DirectWritePtHandle(hDevice, 0, 2097152, bPerSec, VarPtr(Tbl(0)))

    'referme le handle
    Call CloseHandle(hDevice)

    'libère classe
    Set cDriv = Nothing

    SanitizeDrive = 1
End Function


'=======================================================
'Lance la sanitization d'un disque physical... OMFG !
'=======================================================
Public Function SanitizePhysicalDisk(DiskNumber As Byte) As Long
Dim cDriv As PhysicalDisk
Dim x As Long
Dim secPerString As Long
Dim s As String
Dim pt As Long
Dim hDevice As Long
Dim bPerSec As Long
Dim curOp As Currency
Dim bMax As Long
Dim Tbl(2097151) As Byte
Dim cASM As CAsmProc
Dim sH55() As Byte
Dim sHAA() As Byte
Dim pAA As Long
Dim p55 As Long

    'vérifie que le disque est accessible
    If Me.IsPhysicalDiskAvailable(DiskNumber) = False Then
        SanitizePhysicalDisk = -1
        Exit Function
    End If
    
    '//génère les tableaux
    
    'les redimensionne
    ReDim sH55(2097151)
    ReDim sHAA(2097151)
    
    'les remplit
    For x = 0 To 2097151
        sH55(x) = 85
        sHAA(x) = 170
    Next x
    
    'récupère les pointeurs
    pAA = VarPtr(sHAA(0))
    p55 = VarPtr(sH55(0))
    
    'instancie la classe
    Set cASM = New CAsmProc

    'récupère les infos sue le disque
    Set cDriv = Me.GetPhysicalDisk(DiskNumber)

    'nombre de secteurs pour une string de 2Mo
    secPerString = 2097152 / cDriv.BytesPerSector

    'handle du disque
    hDevice = CreateFile("\\.\PHYSICALDRIVE" & CStr(DiskNumber), GENERIC_WRITE, _
        FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)

    bPerSec = cDriv.BytesPerSector
    bMax = Int(cDriv.TotalPhysicalSectors / secPerString) + 1  '+1

    'pour chaque secteur
    For x = 1 To bMax

        'on récupère un pointeur sur une string de 2Mo
        Call cASM.bnAlloc2MoAlea(Tbl(0))

        'calcul unique
        curOp = CCur(x * secPerString)

        'on écrit dans le disque
        '// &H55
        Call DirectWritePtHandle(hDevice, curOp, 2097152, bPerSec, p55)

        '// &HAA
        Call DirectWritePtHandle(hDevice, curOp, 2097152, bPerSec, pAA)

        '//random string
        Call DirectWritePtHandle(hDevice, curOp, 2097152, bPerSec, VarPtr(Tbl(0)))

        'rend la main de tps en tps
        If (x Mod 5) = 0 Then DoEvents

    Next x

    '//s'occupe de l'entête du disque
        Call cASM.bnAlloc2MoAlea(Tbl(0))

        'on écrit dans le disque
        '// &H55
        Call DirectWritePtHandle(hDevice, 0, 2097152, bPerSec, p55)

        '// &HAA
        Call DirectWritePtHandle(hDevice, 0, 2097152, bPerSec, pAA)

        '//random string
        Call DirectWritePtHandle(hDevice, 0, 2097152, bPerSec, VarPtr(Tbl(0)))

    'referme le handle
    Call CloseHandle(hDevice)

    'libère classe
    Set cDriv = Nothing

    SanitizePhysicalDisk = 1
End Function

'=======================================================
'Recherche multicritères de fichiers
'=======================================================
Public Function SearchForFiles(Search As SearchDefinition, Optional _
    RefreshInfos As Boolean = False, Optional FormEvent As Object) As Files
    
Dim s() As String
Dim cFiles As Files
Dim x As Long
    
    'FormEvent implémente l'interface
    If Not IsMissing(FormEvent) Then Set IEvent = FormEvent
    Set cFiles = New Files
    
    ReDim s(0)
    'on recherche tous les fichiers dans le dossier désigné
    s() = Me.EnumFilesStr(Search.FolderName, Search.SubFolder)
    
    'on récupère une bonne fois pour toutes les dates des critères sous forme
    'de currency
    varDates(0) = DateString2Currency(Search.DateCreatedValue)
    varDates(1) = DateString2Currency(Search.DateLastAccessedValue)
    varDates(2) = DateString2Currency(Search.DateLastModifiedValue)
    
    'maintenant pour chaque fichier, on vérifie qu'il correspond aux critères
    For x = 1 To UBound(s())
        If IsFileOk(s(x), Search) Then
            
            'ajoute l'item
            Call cFiles.AddItem(s(x), RefreshInfos)
            
            'lance un event dans la classe implémentant l'interface
            If Not IsMissing(FormEvent) Then Call IEvent.ItemFound(s(x))
        End If
            
        If (x Mod 1000) = 0 Then DoEvents
    Next x
    
    Set SearchForFiles = cFiles
    
End Function

'=======================================================
'Recherche multicritères de dossiers
'=======================================================
Public Function SearchForFolders(Search As SearchDefinition, Optional _
    RefreshInfos As Boolean = False, Optional FormEvent As Object) As Folders
    
Dim s() As String
Dim cFolders As Folders
Dim x As Long
    
    'FormEvent implémente l'interface
    If Not IsMissing(FormEvent) Then Set IEvent = FormEvent
    Set cFolders = New Folders
    
    ReDim s(0)
    'on recherche tous les fichiers dans le dossier désigné
    s() = Me.EnumFoldersStr(Search.FolderName, Search.SubFolder)
    
    'on récupère une bonne fois pour toutes les dates des critères sous forme
    'de currency
    varDates(0) = DateString2Currency(Search.DateCreatedValue)
    varDates(1) = DateString2Currency(Search.DateLastAccessedValue)
    varDates(2) = DateString2Currency(Search.DateLastModifiedValue)
    
    'maintenant pour chaque fichier, on vérifie qu'il correspond aux critères
    For x = 1 To UBound(s())
        If IsFolderOk(s(x), Search) Then
            
            'ajoute l'item
            Call cFolders.AddItem(s(x), RefreshInfos)
            
            'lance un event dans la classe implémentant l'interface
            If Not IsMissing(FormEvent) Then Call IEvent.ItemFound(s(x))
        End If
            
        If (x Mod 1000) = 0 Then DoEvents
    Next x
    
    Set SearchForFolders = cFolders
    
End Function

'=======================================================
'Recherche multicritères de fichiers (sortie string)
'=======================================================
Public Function SearchForFilesStr(Search As SearchDefinition, Optional _
    FormEvent As Object) As String()

Dim s() As String
Dim n As Long
Dim s2() As String
Dim x As Long
    
    'FormEvent implémente l'interface
    If Not IsMissing(FormEvent) Then Set IEvent = FormEvent
    
    ReDim s(0)
    ReDim s2(0)
    n = 0
    'on recherche tous les fichiers dans le dossier désigné
    s() = Me.EnumFilesStr(Search.FolderName, Search.SubFolder)
    
    'on récupère une bonne fois pour toutes les dates des critères sous forme
    'de currency
    varDates(0) = DateString2Currency(Search.DateCreatedValue)
    varDates(1) = DateString2Currency(Search.DateLastAccessedValue)
    varDates(2) = DateString2Currency(Search.DateLastModifiedValue)
    
    'maintenant pour chaque fichier, on vérifie qu'il correspond aux critères
    For x = 1 To UBound(s())
        If IsFileOk(s(x), Search) Then
            
            'ajoute l'item
            n = n + 1
            ReDim Preserve s2(n)
            s2(n) = s(x)
            
            'lance un event dans la classe implémentant l'interface
            If Not IsMissing(FormEvent) Then Call IEvent.ItemFound(s(x))
        End If
            
        If (x Mod 10) = 0 Then DoEvents
    Next x
    
    SearchForFilesStr = s2
    
    'recherche terminée
    Call IEvent.SearchIsFinished
    
End Function

'=======================================================
'Recherche multicritères de dossiers (sortie string)
'=======================================================
Public Function SearchForFoldersStr(Search As SearchDefinition, Optional _
    FormEvent As Object) As String()

Dim s() As String
Dim n As Long
Dim s2() As String
Dim x As Long
    
    'FormEvent implémente l'interface
    If Not IsMissing(FormEvent) Then Set IEvent = FormEvent
    
    ReDim s(0)
    ReDim s2(0)
    n = 0
    'on recherche tous les fichiers dans le dossier désigné
    s() = Me.EnumFoldersStr(Search.FolderName, Search.SubFolder)
    
    'on récupère une bonne fois pour toutes les dates des critères sous forme
    'de currency
    varDates(0) = DateString2Currency(Search.DateCreatedValue)
    varDates(1) = DateString2Currency(Search.DateLastAccessedValue)
    varDates(2) = DateString2Currency(Search.DateLastModifiedValue)
    
    'maintenant pour chaque fichier, on vérifie qu'il correspond aux critères
    For x = 1 To UBound(s())
        If IsFolderOk(s(x), Search) Then
            
            'ajoute l'item
            n = n + 1
            ReDim Preserve s2(n)
            s2(n) = s(x)
            
            'lance un event dans la classe implémentant l'interface
            If Not IsMissing(FormEvent) Then Call IEvent.ItemFound(s(x))
        End If
            
        If (x Mod 10) = 0 Then DoEvents
    Next x
    
    SearchForFoldersStr = s2
    
    'recherche terminée
    Call IEvent.SearchIsFinished
    
End Function

'=======================================================
'Change le label d'un disque
'=======================================================
Public Function SetVolumeLabel(DiskLetter As String, NewLabel As String) As Long
    SetVolumeLabel = SetVolumeLabelAPI(DiskLetter & ":\", NewLabel)
End Function

'=======================================================
'Affecte des dates à un fichier
'=======================================================
Public Function SetFileDates(Source As String, Optional ByVal DateCreated As _
    String = vbNullString, Optional ByVal DateLastAccess As String = vbNullString, _
    Optional ByVal DateLastModification As String = vbNullString) As Long
    
Dim hFile As Long
Dim bModif As Boolean
Dim ftCreated As FILETIME
Dim ftAccess As FILETIME
Dim ftModified As FILETIME
Dim stCreated As SYSTEMTIME
Dim stAccess As SYSTEMTIME
Dim stModified As SYSTEMTIME

    bModif = False
    
    If DateCreated <> vbNullString Then
        'on change la date de création
        
        bModif = True
        
        With stCreated
            .wDay = Day(DateCreated)
            .wHour = Hour(DateCreated)
            .wMinute = Minute(DateCreated)
            .wMonth = Month(DateCreated)
            .wSecond = Second(DateCreated)
            .wYear = Year(DateCreated)
        End With
        
        'convertit en FILETIME
        Call SystemTimeToFileTime(stCreated, ftCreated)
        'convertit en temps local
        Call LocalFileTimeToFileTime(ftCreated, ftCreated)
    End If

    If DateLastAccess <> vbNullString Then
        'on change la date de dernier access
        
        bModif = True
        
        With stAccess
            .wDay = Day(DateLastAccess)
            .wHour = Hour(DateLastAccess)
            .wMinute = Minute(DateLastAccess)
            .wMonth = Month(DateLastAccess)
            .wSecond = Second(DateLastAccess)
            .wYear = Year(DateLastAccess)
        End With
        
        'convertit en FILETIME
        Call SystemTimeToFileTime(stAccess, ftAccess)
        'convertit en temps local
        Call LocalFileTimeToFileTime(ftAccess, ftAccess)
    End If
    
    If DateLastModification <> vbNullString Then
        'on change la date de création
        
        bModif = True
        
        With stModified
            .wDay = Day(DateLastModification)
            .wHour = Hour(DateLastModification)
            .wMinute = Minute(DateLastModification)
            .wMonth = Month(DateLastModification)
            .wSecond = Second(DateLastModification)
            .wYear = Year(DateLastModification)
        End With
        
        'convertit en FILETIME
        Call SystemTimeToFileTime(stModified, ftModified)
        'convertit en temps local
        Call LocalFileTimeToFileTime(ftModified, ftModified)
    End If
    
    'si pas de modifs à faire
    If bModif = False Then Exit Function
        
    'récupère le handle
    hFile = CreateFile(Source, GENERIC_WRITE, FILE_SHARE_READ Or _
        FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)

    If hFile <> INVALID_HANDLE_VALUE Then
    
        'change les dates
        SetFileDates = SetFileTime(hFile, ftCreated, ftAccess, ftModified)
        
        'referme le handle
        Call CloseHandle(hFile)
    End If

End Function
    
'=======================================================
'Affecte des dates à un dossier
'=======================================================
Public Function SetFolderDates(Source As String, Optional ByVal DateCreated As _
    String = vbNullString, Optional ByVal DateLastAccess As String = vbNullString, _
    Optional ByVal DateLastModification As String = vbNullString) As Long
    
Dim hFolder As Long
Dim bModif As Boolean
Dim ftCreated As FILETIME
Dim ftAccess As FILETIME
Dim ftModified As FILETIME
Dim stCreated As SYSTEMTIME
Dim stAccess As SYSTEMTIME
Dim stModified As SYSTEMTIME

    bModif = False
    
    If DateCreated <> vbNullString Then
        'on change la date de création
        
        bModif = True
        
        With stCreated
            .wDay = Day(DateCreated)
            .wHour = Hour(DateCreated)
            .wMinute = Minute(DateCreated)
            .wMonth = Month(DateCreated)
            .wSecond = Second(DateCreated)
            .wYear = Year(DateCreated)
        End With
        
        'convertit en FILETIME
        Call SystemTimeToFileTime(stCreated, ftCreated)
        'convertit en temps local
        Call LocalFileTimeToFileTime(ftCreated, ftCreated)
    End If

    If DateLastAccess <> vbNullString Then
        'on change la date de dernier access
        
        bModif = True
        
        With stAccess
            .wDay = Day(DateLastAccess)
            .wHour = Hour(DateLastAccess)
            .wMinute = Minute(DateLastAccess)
            .wMonth = Month(DateLastAccess)
            .wSecond = Second(DateLastAccess)
            .wYear = Year(DateLastAccess)
        End With
        
        'convertit en FILETIME
        Call SystemTimeToFileTime(stAccess, ftAccess)
        'convertit en temps local
        Call LocalFileTimeToFileTime(ftAccess, ftAccess)
    End If
    
    If DateLastModification <> vbNullString Then
        'on change la date de création
        
        bModif = True
        
        With stModified
            .wDay = Day(DateLastModification)
            .wHour = Hour(DateLastModification)
            .wMinute = Minute(DateLastModification)
            .wMonth = Month(DateLastModification)
            .wSecond = Second(DateLastModification)
            .wYear = Year(DateLastModification)
        End With
        
        'convertit en FILETIME
        Call SystemTimeToFileTime(stModified, ftModified)
        'convertit en temps local
        Call LocalFileTimeToFileTime(ftModified, ftModified)
    End If
    
    'si pas de modifs à faire
    If bModif = False Then Exit Function
        
    'récupère le handle
    hFolder = CreateFile(Source, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or _
        FILE_SHARE_DELETE, 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0&)

    If hFolder <> INVALID_HANDLE_VALUE Then
    
        'change les dates
        SetFolderDates = SetFileTime(hFolder, ftCreated, ftAccess, ftModified)
        
        'referme le handle
        Call CloseHandle(hFolder)
    End If

End Function

'=======================================================
'Compare deux fichiers
'=======================================================
Public Function CompareFiles(File1 As String, File2 As String) As FileComparisonResult
Dim curSize As Currency
Dim hFile1 As Long
Dim hFile2 As Long
Dim nbBuffer As Long
Dim lastBufferSize As Currency
Dim x As Long
Dim curLastPt As Currency

    CompareFiles = [Files are different]

    'on vérifie d'abord les tailles
    curSize = Me.GetFileSize(File1)
    If curSize <> Me.GetFileSize(File2) Then Exit Function
    
    'récupère un handle du fichier 1 et 2
    hFile1 = CreateFile(File1, GENERIC_READ, FILE_SHARE_READ Or _
        FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    hFile2 = CreateFile(File2, GENERIC_READ, FILE_SHARE_READ Or _
        FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    
    If hFile1 = INVALID_HANDLE_VALUE Or hFile2 = INVALID_HANDLE_VALUE Then _
        Exit Function
    
    'détermine le nombre de buffers de 5Mo et la taille du dernier buffer
    nbBuffer = Int(curSize / 5242880)
    lastBufferSize = CCur(-5242.88 * nbBuffer) * 1000
    lastBufferSize = CCur(lastBufferSize + curSize)
    
    For x = 1 To nbBuffer
        
        curLastPt = 5242.88 * (x - 1)
        curLastPt = curLastPt * 1000
        
        If GetBytesFromFileHandle(hFile1, 5242880, curLastPt) <> _
            GetBytesFromFileHandle(hFile2, 5242880, curLastPt) Then
            
            'alors fichier différents
            
            'referme les handles
            Call CloseHandle(hFile1)
            Call CloseHandle(hFile2)
            
            Exit Function
        End If
        
        If (x Mod 20) = 0 Then DoEvents

    Next x
    
    'le dernier buffer
    curLastPt = 5242880
    curLastPt = CCur(curLastPt * (nbBuffer - 1))
    
    If GetBytesFromFileHandle(hFile1, lastBufferSize, curLastPt) = _
        GetBytesFromFileHandle(hFile2, lastBufferSize, curLastPt) Then
        
        'fichiers identiques
        CompareFiles = [Files are identical]
    End If
    
    'referme les handles
    Call CloseHandle(hFile1)
    Call CloseHandle(hFile2)
    
End Function

'=======================================================
'Découpe un fichier
'=======================================================
Public Function CutFile(Source As String, DestinationFolder As String, _
    Method As CutMethod) As Long
    
    CutFile = Me.CutFileFriend(Source, DestinationFolder, Method)
End Function

'=======================================================
'Fusionne des fichiers
'=======================================================
Public Function FusionFiles(GrupFile As String, FolderDestination As String) _
    As Long
    
    FusionFiles = Me.PasteFile(GrupFile, FolderDestination)
End Function

'=======================================================
'Créé un ISO depuis un disque
'=======================================================
Public Function CreateIsoFromDrive(DriveLetter As String, IsoDestination As _
    String) As Long

Dim lngSec As Long
Dim s As String
Dim x As Long
Dim hFile As Long
Dim hDrive As Long
Dim lBPC As Long
Dim lTLS As Long
Dim hDisk As Long
    
    If Me.FileExists(IsoDestination) Then Exit Function
    
    'créé un fichier vide
    Call Me.CreateEmptyFile(IsoDestination, True)
    
    'récupère les handles
    hFile = CreateFile(IsoDestination, GENERIC_WRITE, FILE_SHARE_READ Or _
        FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    hDrive = CreateFile("\\.\" & UCase$(DriveLetter) & ":", GENERIC_READ, _
        FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    
    If hFile = INVALID_HANDLE_VALUE Or hDrive = INVALID_HANDLE_VALUE Then _
        Exit Function
    
    lBPC = Me.GetDriveGeometry_HANDLE(hDrive).BytesPerSector
    lTLS = Me.GetDrive(DriveLetter).TotalPhysicalSectors - 1
    
    'pour chaque secteur
    For lngSec = 0 To lTLS
        
        'sauve dans le fichier la string lue dans le disque
        s = DirectReadSHandle(hDrive, _
            lngSec, lBPC, lBPC)
        Call WriteBytesToFileEndHandle(hFile, s)
        
        If (lngSec Mod 500) = 0 Then DoEvents 'rend la main de tps en tps
             
    Next lngSec

    'referme les handles
    Call CloseHandle(hFile)
    Call CloseHandle(hDrive)
    
    CreateIsoFromDrive = hFile
    
End Function










'=======================================================
'//FRIENDS FUNCTIONS
'=======================================================


'=======================================================
'Récupère la géométrie d'un drive par le handle en param
'=======================================================
Friend Function GetDriveGeometry_HANDLE(ByVal hDrive As Long) As DiskGeometry
Dim bytesReturned As Long
Dim ovrTemp As OVERLAPPED
Dim tGeom As DISK_GEOMETRY

    
    '//obtient les infos sur la géométrie du disque
        
        'obtient les infos souhaitées par l'API DeviceIoControl ; ControlCode=Geometry
        Call DeviceIoControl(hDrive, IOCTL_DISK_GET_DRIVE_GEOMETRY, ByVal 0&, _
            0&, tGeom, Len(tGeom), bytesReturned, ovrTemp)
        
        With tGeom
            GetDriveGeometry_HANDLE.BytesPerSector = .BytesPerSector
            GetDriveGeometry_HANDLE.Cylinders = 10000 * .Cylinders
            GetDriveGeometry_HANDLE.MediaType = .MediaType
            GetDriveGeometry_HANDLE.strMediaType = GetMediaType(.MediaType)
            GetDriveGeometry_HANDLE.TracksPerCylinder = .TracksPerCylinder
            GetDriveGeometry_HANDLE.SectorsPerTrack = .SectorsPerTrack
        End With

End Function

'=======================================================
'Renvoie diverses infos sur les tailles d'un disque par le handle en param
'=======================================================
Friend Function GetDriveSizes_HANDLE(ByVal hDrive As String, _
    sDiskLetter As String) As DriveSizes
    
Dim strDrive As String
Dim tGeom As DiskGeometry
Dim a As Long, b As Long, c As Long, e As Long
    
    'récupère quelques info géométriques
    tGeom = Me.GetDriveGeometry_HANDLE(hDrive)
    
    'calcule les tailles
    Call GetDiskFreeSpace(sDiskLetter & ":\", a, e, b, c)
    
    With GetDriveSizes_HANDLE
        .SectorPerCluster = a
        .FreeClusters = b
        .TotalClusters = c
        .BytesPerSector = e
        .TotalSpace = CCur(CCur(.BytesPerSector) * CCur(.SectorPerCluster) * _
            CCur(.TotalClusters))
        .UsedClusters = .TotalClusters - .FreeClusters
        .BytesPerCluster = CLng(.BytesPerSector * .SectorPerCluster)
        .FreeSpace = CCur(CCur(.BytesPerSector) * CCur(.SectorPerCluster) * _
            CCur(.FreeClusters))
        .UsedSpace = .TotalSpace - .FreeSpace
    End With
    
End Function

'=======================================================
'Récupère les infos partition d'un disque par le handle en param
'=======================================================
Friend Function GetDrivePartitionInfo_HANDLE(ByVal hDrive As String) As PartitionInfo
Dim tPart As PARTITION_INFORMATION
Dim ovrTemp As OVERLAPPED
Dim bytesReturned As Long
    
    'ControlCode=IOCTL_DISK_GET_PARTITION_INFO
    Call DeviceIoControl(hDrive, IOCTL_DISK_GET_PARTITION_INFO, ByVal 0&, 0&, _
        tPart, Len(tPart), bytesReturned, ovrTemp)
    
    With tPart
        GetDrivePartitionInfo_HANDLE.PartitionLength = 10000 * .PartitionLength
        GetDrivePartitionInfo_HANDLE.HiddenSectors = .HiddenSectors
        GetDrivePartitionInfo_HANDLE.StartingOffset = .StartingOffset * 10000
    End With
    
End Function

'=======================================================
'Récupère la géométrie d'un disque par le handle en param
'=======================================================
Friend Function GetPhysicalDiskGeometry_HANDLE(ByVal hDrive As Long) As DiskGeometry
Dim tGeom As DISK_GEOMETRY
Dim bytesReturned As Long
Dim ovrTemp As OVERLAPPED

    
    '//obtient les infos sur la géométrie du disque
        
        'obtient les infos souhaitées par l'API DeviceIoControl ; ControlCode=Geometry
        Call DeviceIoControl(hDrive, IOCTL_DISK_GET_DRIVE_GEOMETRY, ByVal 0&, _
            0&, tGeom, Len(tGeom), bytesReturned, ovrTemp)
        
        With tGeom
            GetPhysicalDiskGeometry_HANDLE.BytesPerSector = .BytesPerSector
            GetPhysicalDiskGeometry_HANDLE.Cylinders = 10000 * .Cylinders
            GetPhysicalDiskGeometry_HANDLE.MediaType = .MediaType
            GetPhysicalDiskGeometry_HANDLE.TracksPerCylinder = .TracksPerCylinder
            GetPhysicalDiskGeometry_HANDLE.SectorsPerTrack = .SectorsPerTrack
            GetPhysicalDiskGeometry_HANDLE.strMediaType = GetMediaType(.MediaType)
        End With

End Function

'=======================================================
'Renvoie le nom de disque d'un disque dur physique par le handle en param
'=======================================================
Friend Function GetPhysicalDiskName_HANDLE(ByVal hDrive As Long, ByVal DiskNumber As Byte) As String
Dim SCIP As SENDCMDINPARAMS
Dim IDSEC As IDSECTOR
Dim bArrOut(OUTPUT_DATA_SIZE - 1) As Byte
Dim lpcbBytesReturned As Long

    'affecte toutes les properties nécessaires à cette structure
    With SCIP
        .bDriveNumber = CByte(DiskNumber)
        .cBufferSize = IDENTIFY_BUFFER_SIZE
        With .irDriveRegs
            .bCylLowReg = 0
            .bFeaturesReg = 0
            .bCommandReg = &HEC
            .bSectorCountReg = 1
            .bSectorNumberReg = 1
            .bCylHighReg = 0
            .bDriveHeadReg = &HA0
        End With
    End With
    
    'on récupère les infos souhaitées dans la structure IDSEC gràce au DeviceIoControl
    If DeviceIoControl2(hDrive, DFP_RECEIVE_DRIVE_DATA, SCIP, Len(SCIP) - 4, bArrOut(0), _
        OUTPUT_DATA_SIZE, lpcbBytesReturned, ByVal 0&) Then

        CopyMemory IDSEC, bArrOut(16), Len(IDSEC)
        GetPhysicalDiskName_HANDLE = PermuteString(StrConv(IDSEC.sModelNumber, _
            vbUnicode))
    End If
End Function

'=======================================================
'Renvoie les dates d'un fichier par le handle en param
'=======================================================
Friend Function GetFileDates_HANDLE(ByVal hFile As Long) As FileDates
Dim FileInfo As BY_HANDLE_FILE_INFORMATION
    
    'obtient les infos sur le fichier
    Call GetFileInformationByHandle(hFile, FileInfo)

    With GetFileDates_HANDLE
        .DateCreated = FileTimeToString(FileInfo.ftCreationTime)
        .DateLastAccessed = FileTimeToString(FileInfo.ftLastAccessTime)
        .DateLastModified = FileTimeToString(FileInfo.ftLastWriteTime)
    End With

End Function

'=======================================================
'Renvoie les dates d'un fichier par le handle en param
'=======================================================
Friend Function GetFolderDates_HANDLE(ByVal hFolder As Long) As FileDates
Dim FileInfo As BY_HANDLE_FILE_INFORMATION
    
    'obtient les infos sur le fichier
    Call GetFileInformationByHandle(hFolder, FileInfo)

    With GetFolderDates_HANDLE
        .DateCreated = FileTimeToString(FileInfo.ftCreationTime)
        .DateLastAccessed = FileTimeToString(FileInfo.ftLastAccessTime)
        .DateLastModified = FileTimeToString(FileInfo.ftLastWriteTime)
    End With

End Function

'=======================================================
'Renvoie les tailles par le handle en param
'=======================================================
Friend Function GetFileSizes_HANDLE(ByVal hFile As Long, ByVal Source As String) _
    As FileSizes
    
Dim curSize As Currency
    
    'obtient la taille normale par API
    Call GetFileSizeEx(hFile, curSize)
        
    With GetFileSizes_HANDLE
        .FileSize = curSize * 10000  'multiplie par 10^4 pour obtenir un nombre entier
        .CompressedFileSize = GetCompressedFileSizeAPI(Source, ByVal 0&)
    End With

End Function

'=======================================================
'Récupère l'attribut d'un fichier par le handle en param
'=======================================================
Friend Function GetFileAttributes_HANDLE(ByVal hFile As Long) As FileAttributes
Dim FileInfo As BY_HANDLE_FILE_INFORMATION
    
    'obtient les infos sur le fichier
    Call GetFileInformationByHandle(hFile, FileInfo)
    
    GetFileAttributes_HANDLE = FileInfo.dwFileAttributes
    
End Function

'=======================================================
'Récupère l'attribut d'un dossier par le handle en param
'=======================================================
Friend Function GetFolderAttributes_HANDLE(ByVal hFolder As Long) As FileAttributes
Dim FileInfo As BY_HANDLE_FILE_INFORMATION
    
    'obtient les infos sur le fichier
    Call GetFileInformationByHandle(hFolder, FileInfo)
    
    GetFolderAttributes_HANDLE = FileInfo.dwFileAttributes
    
End Function

'=======================================================
'Récupère des bytes dans un fichier
'=======================================================
Friend Function GetBytesFromFileHandle(ByVal hFile As Long, ByVal curSize As _
    Currency, ByVal curOffset As Currency) As String
    
Dim tmpText As String
Dim Ret As Long
        
    'créé un buffer qui contiendra les valeurs
    tmpText = String$(curSize, 0)

    'bouge le pointeur sur lr fichier au bon emplacement
    Ret = SetFilePointerEx(hFile, curOffset / 10000, 0&, FILE_BEGIN)  'divise par 10000 pour
    'pouvoir renvoyer une currency DECIMALE (cad du genre 1.4567 pour l'offset 14567)
    
    'prend un morceau du fichier
    Ret = ReadFile(hFile, ByVal tmpText, Len(tmpText), Ret, ByVal 0&)
    
    'affecte à la fonction
    GetBytesFromFileHandle = tmpText

End Function

'=======================================================
'Permet de lire des bytes directement dans le disque
'Sortie en String
'Demande un handle
'=======================================================
Friend Function DirectReadSHandle(ByVal hDevice As Long, ByVal iStartSec As _
    Currency, ByVal nBytes As Long, ByVal lBytesPerSector As Long) As String

'/!\ iStartsec et nbytes doivent être des multiples de la taille d'un secteur (généralement 512 octets)
Dim BytesRead As Long
Dim Pointeur As Currency
Dim Ret As Long
Dim lLowPart As Long, lHighPart As Long
   
    'détermine le byte de départ du secteur
    Pointeur = CCur(iStartSec) * CCur(lBytesPerSector)
    
    'transforme un currency en 2 long pour une structure LARGE_INTEGER
    Call GetLargeInteger(Pointeur, lLowPart, lHighPart)

    'déplace, dans le fichier (ici un disque) pointé par hDevice, le "curseur" au premier
    'byte que l'on veut lire (donné par deux long)
    Ret = SetFilePointer(hDevice, lLowPart, lHighPart, FILE_BEGIN)  'FILE_BEGIN ==> part du début du fichier pour décompter la DistanceToMove
    
    'création d'un buffer
    DirectReadSHandle = Space$(nBytes)

    'obtention de la string
    Ret = ReadFile(hDevice, ByVal DirectReadSHandle, nBytes, BytesRead, 0&)

End Function

'=======================================================
'Ecrit des bytes dans un fichier (à la fin du fichier) avec en entrée un handle
'=======================================================
Friend Function WriteBytesToFileEndHandle(ByVal lngFile As Long, ByVal _
    sString As String) As String
    
Dim tmpText As String
Dim Ret As Long
        
    'bouge le pointeur sur le fichier à la fin du fichier
    Ret = SetFilePointerEx(lngFile, 0&, 0&, FILE_END) '

    'écriture dans le fichier
    Call WriteFile(lngFile, ByVal sString, Len(sString), Ret, ByVal 0&)

End Function

'=======================================================
'Fonction de découpe de fichier
'=======================================================
Friend Function CutFileFriend(ByVal sFile As String, ByVal sFolderOut As _
    String, tMethode As CutMethod) As Long
    
Dim lFileCount As Long
Dim lLastFileSize As Currency
Dim curSize As Currency
Dim x As Long
Dim i As Long
Dim j As Long
Dim a As Long
Dim lBuf As Long
Dim sFileStr As String
Dim sBuf As String
Dim lBuf2 As Long
Dim sFic As String
Dim lTime As Long
Dim lNormalSize As Currency
Dim k As Currency
Dim k2 As Currency
    
    '//VERIFICATIONS
    'vérifie que le fichier existe bien
    If Me.FileExists(sFile) = False Then
        'fichier manquant
        CutFileFriend = -1
        Exit Function
    End If
    
    'vérifie que le dossier résultat existe bien
    If Me.FolderExists(sFolderOut) = False Then
        'dossier résultat inexistant
        CutFileFriend = -1
        Exit Function
    End If
    
    'récupère le nom du fichier
    sFileStr = Me.GetFileName(sFile)
    
    'récupère la taille du fichier
    curSize = Me.GetFileSize(sFile)
    If curSize = 0 Or Me.IsFileAvailable(sFile) = False Then
        'fichier vide ou inaccessible
        CutFileFriend = -1
        Exit Function
    End If
    
    '//LANCE LE DECOUPAGE
    If tMethode.tMethode = [Fixed Size] Then
        'alors on découpe en fixant la taille
        
        'calcule le nombre de fichiers nécessaire et la taille du dernier
        lFileCount = Int(curSize / tMethode.lParam) + IIf(Mod2(curSize, tMethode.lParam) = 0, 0, 1)
        lLastFileSize = curSize     'taille du dernier fichier
        k = (lFileCount - 1)
        k = k * tMethode.lParam
        lLastFileSize = lLastFileSize - k 'évite les dépassement de capacité
        
        
        'lance la découpe
        'utilisation de l'API ReadFile pour plus d'efficacité
        'prend des buffers de 5Mo maximum
        If tMethode.lParam <= LBUFSIZE Then
            'alors tout rentre dans un seul buffer
                        
            k2 = lFileCount - 1
            k2 = k2 * tMethode.lParam
            
            For i = 1 To lFileCount - 1 'pas le dernier qui est fait à part
                
                'fichier résultat i
                sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(i))
                
                'créé le fichier résultat
                Call Me.CreateEmptyFile(sFic, True)
                
                'récupère le buffer
                k = i - 1
                k = k * tMethode.lParam
                sBuf = GetBytesFromFile(sFile, CCur(tMethode.lParam), _
                    CCur(k))
                
                'on écrit dans le fichier résultat
                Call WriteBytesToFile(sFic, sBuf, 0)
                
                DoEvents
            Next i
            
            'maintenant le dernier fichier
            sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(lFileCount))
            
            'créé le fichier résultat
            Call Me.CreateEmptyFile(sFic, True)
            
            'récupère le buffer
            sBuf = GetBytesFromFile(sFile, lLastFileSize, CCur(k2))
            
            'on écrit dans le fichier résultat
            Call WriteBytesToFile(sFic, sBuf, 0)
            
        Else
            'alors plusieurs buffer
            
            'calcule le nombre de buffers nécessaires pour chaque fichier
            lBuf2 = Int(tMethode.lParam / LBUFSIZE) + IIf(Mod2(tMethode.lParam, LBUFSIZE) = 0, 0, 1)
            
            k2 = lFileCount - 1
            k2 = k2 * tMethode.lParam
            
            For i = 1 To lFileCount - 1 'pas le dernier qui est fait à part
                
                'fichier résultat i
                sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(i))
                
                'créé le fichier résultat
                Call Me.CreateEmptyFile(sFic, True)
                
                k = i - 1
                k = k * tMethode.lParam
                
                For j = 1 To lBuf2 - 1
 
                    'récupère le buffer
                    sBuf = GetBytesFromFile(sFile, LBUFSIZE, CCur((j - 1) * LBUFSIZE) + k)
                    
                    'on écrit dans le fichier résultat
                    Call WriteBytesToFileEnd(sFic, sBuf)   ', 5242880 * (j - 1)
                    
                    DoEvents

                Next j
                
                'le dernier buffer
                sBuf = GetBytesFromFile(sFile, tMethode.lParam - (lBuf2 - 1) * LBUFSIZE, CCur((lBuf2 - 1) * LBUFSIZE) + k)

                'on écrit dans le fichier résultat
                Call WriteBytesToFileEnd(sFic, sBuf)  ', 5242880 * (lBuf2 - 1)
                
                DoEvents
            
            Next i

            'recalcule le nombre de buffers dans le dernier fichier
            lBuf2 = Int(lLastFileSize / LBUFSIZE) + IIf(Mod2(lLastFileSize, LBUFSIZE) = 0, 0, 1)

            'maintenant le dernier fichier
            sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(lFileCount))
            
            'créé le fichier résultat
            Call Me.CreateEmptyFile(sFic, True)
            
            For j = 1 To lBuf2 - 1
                    
                'récupère le buffer
                sBuf = GetBytesFromFile(sFile, LBUFSIZE, CCur(k2 + (j - 1) * LBUFSIZE))
                
                'on écrit dans le fichier résultat
                Call WriteBytesToFileEnd(sFic, sBuf)  ', 5242880 * (j - 1)
            
            Next j
                
            'récupère le dernier buffer
            sBuf = GetBytesFromFile(sFile, lLastFileSize - (lBuf2 - 1) * LBUFSIZE, CCur(k2 + (lBuf2 - 1) * LBUFSIZE))
            
            'on écrit dans le fichier résultat
            Call WriteBytesToFileEnd(sFic, sBuf)  ', 0
  
        End If
        
        
        'on créé le fichier groupeur
        Call Me.CreateEmptyFile(sFolderOut & "\" & sFileStr & ".grp", True)
        Call Me.SaveDataInFile(sFolderOut & "\" & sFileStr & ".grp", _
            sFileStr & "|" & Str$(lFileCount))
       
    Else
        'alors nombre de fichiers fixé

        'nombre de fichiers
        lFileCount = tMethode.lParam
        
        'calcule la taille de chaque fichier
        lNormalSize = Int(curSize / lFileCount)
        lLastFileSize = lNormalSize + curSize
        k = lNormalSize
        k = k * lFileCount 'taille du dernier fichier (plus quelques octets, au maximum 1 par fichier)
        lLastFileSize = lLastFileSize - k       'évite le dépassement de capacité
        
        
        'lance la découpe
        'utilisation de l'API ReadFile pour plus d'efficacité
        'prend des buffers de 5Mo maximum
        If lNormalSize <= LBUFSIZE Then
            'alors tout rentre dans un seul buffer
            
            k2 = lFileCount - 1
            k2 = k2 * lNormalSize
            
            For i = 1 To lFileCount - 1 'pas le dernier qui est fait à part
                
                'fichier résultat i
                sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(i))
                
                'créé le fichier résultat
                Call Me.CreateEmptyFile(sFic, True)
                
                'récupère le buffer
                k = i - 1
                k = k * lNormalSize
                sBuf = GetBytesFromFile(sFile, lNormalSize, CCur(k))
                
                'on écrit dans le fichier résultat
                Call WriteBytesToFile(sFic, sBuf, 0)
                
                DoEvents
            Next i
            
            'maintenant le dernier fichier
            sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(lFileCount))
            
            'créé le fichier résultat
            Call Me.CreateEmptyFile(sFic, True)
            
            'récupère le buffer
            sBuf = GetBytesFromFile(sFile, lLastFileSize, CCur(k2))
            
            'on écrit dans le fichier résultat
            Call WriteBytesToFile(sFic, sBuf, 0)
            
        Else
            'alors plusieurs buffer
            
            'calcule le nombre de buffers nécessaires pour chaque fichier
            lBuf2 = Int(lNormalSize / LBUFSIZE) + IIf(Mod2(lNormalSize, LBUFSIZE) = 0, 0, 1)
            
            k2 = lFileCount - 1
            k2 = k2 * lNormalSize
            
            For i = 1 To lFileCount - 1 'pas le dernier qui est fait à part
                
                'fichier résultat i
                sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(i))
                
                'créé le fichier résultat
                Call Me.CreateEmptyFile(sFic, True)
                
                DoEvents
                
                k = i - 1
                k = k * lNormalSize
                
                For j = 1 To lBuf2 - 1
                
                    'récupère le buffer
                    sBuf = GetBytesFromFile(sFile, LBUFSIZE, CCur((j - 1) * LBUFSIZE) + k)
                    
                    'on écrit dans le fichier résultat
                    Call WriteBytesToFileEnd(sFic, sBuf)  ', 5242880 * (j - 1)
                    
                    DoEvents
                Next j

                'le dernier buffer
                sBuf = GetBytesFromFile(sFile, lNormalSize - (lBuf2 - 1) * LBUFSIZE, CCur((lBuf2 - 1) * LBUFSIZE) + k)

                'on écrit dans le fichier résultat
                Call WriteBytesToFileEnd(sFic, sBuf)  ', 5242880 * (lBuf2 - 1)
                
                DoEvents
            Next i

            'recalcule le nombre de buffers dans le dernier fichier
            lBuf2 = Int(lLastFileSize / LBUFSIZE) + IIf(Mod2(lLastFileSize, LBUFSIZE) = 0, 0, 1)

            'maintenant le dernier fichier
            sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(lFileCount))
            
            'créé le fichier résultat
            Call Me.CreateEmptyFile(sFic, True)
            
            For j = 1 To lBuf2 - 1
            
                'récupère le buffer
                sBuf = GetBytesFromFile(sFile, LBUFSIZE, CCur(k2 + (j - 1) * LBUFSIZE))
                
                'on écrit dans le fichier résultat
                Call WriteBytesToFileEnd(sFic, sBuf)  ', 5242880 * (j - 1)
            
            Next j
            
            'récupère le dernier buffer
            sBuf = GetBytesFromFile(sFile, lLastFileSize - (lBuf2 - 1) * LBUFSIZE, CCur(k2 + (lBuf2 - 1) * LBUFSIZE))
            
            'on écrit dans le fichier résultat
            Call WriteBytesToFileEnd(sFic, sBuf)  ', 0
            
        End If
        
        
        'on créé le fichier groupeur
        Call Me.CreateEmptyFile(sFolderOut & "\" & sFileStr & ".grp", True)
        Call Me.SaveDataInFile(sFolderOut & "\" & sFileStr & ".grp", _
            sFileStr & "|" & Str$(lFileCount))
 
    End If
    
    'terminé
    CutFileFriend = 1

End Function


'=======================================================
'fonction de fusion de fichier
'=======================================================
Friend Function PasteFile(ByVal sFileGroup As String, ByVal sFolderOut As _
    String) As Long
    
Dim lFileCount As Long
Dim x As Long
Dim i As Long
Dim j As Long
Dim lBuf As Long
Dim sFileStr As String
Dim sBuf As String
Dim lBuf2 As Long
Dim sFic As String
Dim bOk As Boolean
Dim curSize As Currency
Dim a As Long
Dim lTime As Long
    
    '//VERIFICATIONS
        'vérifie que le fichier existe bien
        If Me.FileExists(sFileGroup) = False Then
            'fichier manquant
            PasteFile = -1
            Exit Function
        End If
        
        'vérifie que le dossier résultat existe bien
        If Me.FolderExists(sFolderOut) = False Then
            'dossier résultat inexistant
            PasteFile = -1
            Exit Function
        End If
        sBuf = Me.LoadFileInString(sFileGroup)
        'récupère le nom du fichier
        sFileStr = Mid$(sBuf, 1, InStr(1, sBuf, "|") - 1)
    
        '//LANCE LA FUSION
        'récupère le nombre de fichiers concernés
        lFileCount = Val(Right$(sBuf, Len(sBuf) - InStr(1, sBuf, "|")))
        
        'vérifie l'existence de chaque fichier
        bOk = True
        For i = 1 To lFileCount
            If Me.FileExists(Me.GetFolderName(sFileGroup) & "\" & sFileStr & "." & Trim$(Str$(i))) = False Then
                bOk = False
            End If
        Next i
        If Not (bOk) Then
            'alors un fichier est absent
            PasteFile = -1
            Exit Function
        End If
    
    
    'créé le fichier résultat
    Call Me.CreateEmptyFile(sFolderOut & "\" & sFileStr, True)
    
    'alors tout est OK, on peut commencer à coller les données par buffer de 5Mo
    If Me.GetFileSize(sFolderOut & "\" & sFileStr & ".1") <= LBUFSIZE Then
        'alors tout rentre dans un buffer de 5Mo
    
        For i = 1 To lFileCount
            'écrit les bytes lus
            Call WriteBytesToFileEnd(sFolderOut & "\" & sFileStr, Me.LoadFileInString(Me.GetFolderName(sFileGroup) & "\" & sFileStr & "." & Trim$(Str$(i))))
            DoEvents
        Next i
        
    Else
    
        'alors il faut plusieurs buffers de 5Mo par fichier
        
        'détermine le nombre de buffers nécessaire
        lBuf2 = Int(Me.GetFileSize(sFolderOut & "\" & sFileStr & ".1") / LBUFSIZE) + IIf(Mod2(Me.GetFileSize(sFolderOut & "\" & sFileStr & ".1"), LBUFSIZE) = 0, 0, 1)
                
        For i = 1 To lFileCount - 1
        
            'le fichier que l'on lit
            sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(i))
            
            For j = 1 To lBuf2 - 1
                'sbuf contient 5Mo lus
                sBuf = GetBytesFromFile(sFic, LBUFSIZE, LBUFSIZE * (j - 1))
                
                'écrit les bytes dans le fichier résultat
                Call WriteBytesToFileEnd(sFolderOut & "\" & sFileStr, sBuf)
                
                DoEvents
            Next j
            
            'le dernier buffer
            a = Me.GetFileSize(sFic) - (lBuf2 - 1) * LBUFSIZE      'taille du dernier buffer
            sBuf = GetBytesFromFile(sFic, a, LBUFSIZE * (lBuf2 - 1))
            
            'écrit les bytes dans le fichier résultat
            Call WriteBytesToFileEnd(sFolderOut & "\" & sFileStr, sBuf)
            
            DoEvents
        Next i
        
        'fait le dernier fichier
        sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(lFileCount))
        lBuf2 = Int(Me.GetFileSize(sFic) / LBUFSIZE) + IIf(Mod2(Me.GetFileSize(sFic), LBUFSIZE) = 0, 0, 1)      'nouveau buffer
            
        For j = 1 To lBuf2 - 1
            'sbuf contient 5Mo lus
            sBuf = GetBytesFromFile(sFic, LBUFSIZE, LBUFSIZE * (j - 1))
            
            'écrit les bytes dans le fichier résultat
            Call WriteBytesToFileEnd(sFolderOut & "\" & sFileStr, sBuf)
        Next j
        
        'le dernier buffer
        a = Me.GetFileSize(sFic) - (lBuf2 - 1) * LBUFSIZE
        sBuf = GetBytesFromFile(sFic, a, LBUFSIZE * (lBuf2 - 1))
        
        'écrit les bytes dans le fichier résultat
        Call WriteBytesToFileEnd(sFolderOut & "\" & sFileStr, sBuf)
        
        DoEvents
        
    End If

    PasteFile = 1

End Function

'=======================================================
'Effectue un modulo sans dépassement de capacité
'assez peu optimisé, mais utile pour les grandes valeurs de cur
'=======================================================
Friend Function Mod2(ByVal cur As Currency, lng As Long) As Currency
    Mod2 = cur - Int(cur / lng) * lng
End Function

'=======================================================
'Récupère des bytes dans un fichier
'=======================================================
Friend Function GetBytesFromFile(ByVal sFile As String, ByVal curSize As _
    Currency, ByVal curOffset As Currency) As String
    
Dim tmpText As String
Dim Ret As Long
Dim lFile As Long
    
    'obtient un handle vers le fichier à ouvrir
    lFile = CreateFile(sFile, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    
    If lFile = INVALID_HANDLE_VALUE Then
        'fichier inexistant, ou en tout cas inaccessible
        Exit Function
    End If
        
    'créé un buffer qui contiendra les valeurs
    tmpText = String$(curSize, 0)

    'bouge le pointeur sur lr fichier au bon emplacement
    Ret = SetFilePointerEx(lFile, curOffset / 10000, 0&, FILE_BEGIN)  'divise par 10000 pour
    'pouvoir renvoyer une currency DECIMALE (cad du genre 1.4567 pour l'offset 14567)
    
    'prend un morceau du fichier
    Ret = ReadFile(lFile, ByVal tmpText, Len(tmpText), Ret, ByVal 0&)
    
    'affecte à la fonction
    GetBytesFromFile = tmpText

    'referme le handle
    CloseHandle lFile
    
End Function

'=======================================================
'Ecrit des bytes dans un fichier
'=======================================================
Friend Function WriteBytesToFile(ByVal sFile As String, ByVal sString As String, ByVal curOffset As Currency) As String
Dim tmpText As String
Dim Ret As Long
Dim lFile As Long
       
    'obtient un handle vers le fichier à écrire
    'ouverture en ECRITURE, avec overwrite si déjà existant (car déjà demandé confirmation avant)
    lFile = CreateFile(sFile, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)

    If lFile = INVALID_HANDLE_VALUE Then Exit Function 'fichier indisponible
    
    'bouge le pointeur sur le fichier au bon emplacement
    Ret = SetFilePointerEx(lFile, curOffset / 10000, 0&, FILE_BEGIN)
    'a divisé par 10^4 pour obtenir un nombre décimal de Currency

    'écriture dans le fichier
    Call WriteFile(lFile, ByVal sString, Len(sString), Ret, ByVal 0&)

    'ferme le handle du fichier écrit
    CloseHandle lFile
    
End Function

'=======================================================
'Ecrit des bytes dans un fichier (à la fin du fichier)
'=======================================================
Friend Function WriteBytesToFileEnd(ByVal sFile As String, ByVal sString As _
    String) As String
    
Dim tmpText As String
Dim Ret As Long
Dim lFile As Long
       
    'obtient un handle vers le fichier à écrire
    'ouverture en ECRITURE, avec overwrite si déjà existant (car déjà demandé confirmation avant)
    lFile = CreateFile(sFile, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)

    If lFile = INVALID_HANDLE_VALUE Then Exit Function 'fichier non dispo
    
    'bouge le pointeur sur le fichier à la fin du fichier
    Ret = SetFilePointerEx(lFile, 0&, 0&, FILE_END) '

    'écriture dans le fichier
    Call WriteFile(lFile, ByVal sString, Len(sString), Ret, ByVal 0&)

    'ferme le handle du fichier écrit
    CloseHandle lFile

End Function

'=======================================================
'écrire des bytes dans un fichier
'=======================================================
Friend Sub WriteBytesToFileHandle(ByVal hFile As Long, ByVal pt As Long, _
    ByVal curOffset As Currency, ByVal lLen As Long)
    
Dim tmpText As String
Dim Ret As Long
    
    'bouge le pointeur sur le fichier au bon emplacement
    Ret = SetFilePointerEx(hFile, curOffset / 10000, 0&, FILE_BEGIN)

    'écriture dans le fichier
    Call WriteFile(hFile, ByVal pt, lLen, Ret, ByVal 0&)
    
End Sub

'=======================================================
'Permet de d'écrire de manière directe dans le disque
'avec en entrée un pointeur et un handle de disque
'=======================================================
Friend Sub DirectWritePtHandle(ByVal hDevice As Long, ByVal iStartSec As _
    Currency, ByVal nBytes As Long, ByVal lBytesPerSector As Long, _
    ByRef pt As Long)

'/!\ iStartsec et nbytes doivent être des multiples de la taille d'un secteur (généralement 512 octets)

Dim BytesRead As Long
Dim Pointeur As Currency
Dim Ret As Long
Dim lLowPart As Long
Dim lHighPart As Long
   
    'détermine le byte de départ du secteur
    Pointeur = CCur(iStartSec) * CCur(lBytesPerSector)
    
    'transforme un currency en 2 long pour une structure LARGE_INTEGER
    Call GetLargeInteger(Pointeur, lLowPart, lHighPart)

    'déplace, dans le fichier (ici un disque) pointé par hDevice, le "curseur" au premier
    'byte que l'on veut lire (donné par deux long)
    Ret = SetFilePointer(hDevice, lLowPart, lHighPart, FILE_BEGIN)  'FILE_BEGIN ==> part du début du fichier pour décompter la DistanceToMove
    
    'verrouilage de la zone du disque à écrire
    Call LockFile(hDevice, lLowPart, lHighPart, nBytes, 0)
    
    'écriture disque
    Ret = WriteFile(hDevice, ByVal pt, nBytes, Ret, ByVal 0&)
    
    'on vide les buffers internes et on dévérouille la zone
    Call FlushFileBuffers(hDevice)
    Call UnlockFile(hDevice, lLowPart, lHighPart, nBytes, 0)
    
End Sub












'=======================================================
'PRIVATE FUNCTIONS
'=======================================================

'=======================================================
'Renvoie True si le fichier correspond aux critères de recherche
'=======================================================
Private Function IsFileOk(sFile As String, tS As SearchDefinition) As Boolean
Dim curSize As Currency
Dim curV As Currency
Dim dblSize As Long
Dim tSOp As Operator
Dim Ret As Long
Dim sStr As String

    'recherche par nom
    If (tS.Criteria And SearchName) = SearchName Then
        sStr = Me.GetFileName(sFile)    'récupère le nom
        If tS.RespectCase Then
            If InStr(1, sStr, tS.SearchName) = 0 Then
                'raté
                IsFileOk = False
                Exit Function
            End If
        Else
            If InStr(1, LCase$(sStr), LCase$(tS.SearchName)) = 0 Then
                'raté
                IsFileOk = False
                Exit Function
            End If
        End If
    End If
    
    'recherche par taille
    If (tS.Criteria And SearchSize) = SearchSize Then
        
        dblSize = tS.SizeValue
        tSOp = tS.SizeOperator
        
        'recupère la taille du fichier
        curSize = Me.GetFileSize(sFile)
        
        If tSOp = [ < ] Then
            '<
            If curSize >= dblSize Then
                IsFileOk = False
                Exit Function
            End If
        ElseIf tSOp = [ > ] Then
            '>
            If curSize <= dblSize Then
                IsFileOk = False
                Exit Function
            End If
        ElseIf tSOp = [ = ] Then
            '=
            If curSize <> dblSize Then
                IsFileOk = False
                Exit Function
            End If
        ElseIf tSOp = [ <= ] Then
            '<=
            If curSize > dblSize Then
                IsFileOk = False
                Exit Function
            End If
        ElseIf tSOp = [ >= ] Then
            '>=
            If curSize < dblSize Then
                IsFileOk = False
                Exit Function
            End If
        End If
    End If
    
    
    'recherche par date de création
    If (tS.Criteria And SearchDateCreated) = SearchDateCreated Then
        
        'recupère la date en currency
        curV = Me.GetFileDateAsCurrency(sFile, DateCreated)
        
        tSOp = tS.DateCreatedOperator
        
        If curV Then
            'compare avec la date à rechercher
            '-1==> curDate < curDateReal
            '0 ==> curDate = curDateReal
            '1 ==> curDate > curDateReal
            Ret = CompareFileTime(varDates(0), curV)
            
            If Ret = 0 Then
                If (tSOp = [ = ]) Or (tSOp = [ <= ]) Or (tSOp = [ >= ]) Then _
                    GoTo DateCreatedIsOk
            ElseIf Ret = 1 Then
                If (tSOp = [ < ]) Or (tSOp = [ <= ]) Then GoTo DateCreatedIsOk
            Else
                If (tSOp = [ > ]) Or (tSOp = [ >= ]) Then GoTo DateCreatedIsOk
            End If
        End If
        
        'si on est là, c'est que la comparaison n'est pas bonne
        IsFileOk = False
        Exit Function
    End If
    
DateCreatedIsOk:

    If (tS.Criteria And SearchDateLastAccessed) = SearchDateLastAccessed Then
        
        'recupère la date en currency
        curV = Me.GetFileDateAsCurrency(sFile, DateLastAccessed)
        
        tSOp = tS.DateLastAccessedOperator
        
        If curV Then
            'compare avec la date à rechercher
            '-1==> curDate < curDateReal
            '0 ==> curDate = curDateReal
            '1 ==> curDate > curDateReal
            Ret = CompareFileTime(varDates(1), curV)
            
            If Ret = 0 Then
                If (tSOp = [ = ]) Or (tSOp = [ <= ]) Or (tSOp = [ >= ]) Then _
                    GoTo DateLastAccessedIsOk
            ElseIf Ret = 1 Then
                If (tSOp = [ < ]) Or (tSOp = [ <= ]) Then _
                    GoTo DateLastAccessedIsOk
            Else
                If (tSOp = [ > ]) Or (tSOp = [ >= ]) Then _
                    GoTo DateLastAccessedIsOk
            End If
        End If
        
        'si on est là, c'est que la comparaison n'est pas bonne
        IsFileOk = False
        Exit Function
    End If
    
DateLastAccessedIsOk:
    
    If (tS.Criteria And SearchDateLastModified) = SearchDateLastModified Then
        
        'recupère la date en currency
        curV = Me.GetFileDateAsCurrency(sFile, DateLastModified)
        
        tSOp = tS.DateLastModifiedOperator
        
        If curV Then
            'compare avec la date à rechercher
            '-1==> curDate < curDateReal
            '0 ==> curDate = curDateReal
            '1 ==> curDate > curDateReal
            Ret = CompareFileTime(varDates(2), curV)
            
            If Ret = 0 Then
                If (tSOp = [ = ]) Or (tSOp = [ <= ]) Or (tSOp = [ >= ]) Then _
                    GoTo DateLastModifiedIsOk
            ElseIf Ret = 1 Then
                If (tSOp = [ < ]) Or (tSOp = [ <= ]) Then _
                    GoTo DateLastModifiedIsOk
            Else
                If (tSOp = [ > ]) Or (tSOp = [ >= ]) Then _
                    GoTo DateLastModifiedIsOk
            End If
        End If
        
        'si on est là, c'est que la comparaison n'est pas bonne
        IsFileOk = False
        Exit Function
    End If

DateLastModifiedIsOk:
    
    IsFileOk = True     'parcours du combattant terminé !
    
End Function

'=======================================================
'Renvoie True si le dossier correspond aux critères de recherche
'=======================================================
Private Function IsFolderOk(sFolder As String, tS As SearchDefinition) As Boolean
Dim curSize As Currency
Dim curV As Currency
Dim dblSize As Long
Dim tSOp As Operator
Dim Ret As Long
Dim sStr As String

    'recherche par nom
    If (tS.Criteria And SearchName) = SearchName Then
        sStr = Me.GetFileName(sFolder)      'récupère le nom
        If tS.RespectCase Then
            If InStr(1, sStr, tS.SearchName) = 0 Then
                'raté
                IsFolderOk = False
                Exit Function
            End If
        Else
            If InStr(1, LCase$(sStr), LCase$(tS.SearchName)) = 0 Then
                'raté
                IsFolderOk = False
                Exit Function
            End If
        End If
    End If
    
    
    'recherche par date de création
    If (tS.Criteria And SearchDateCreated) = SearchDateCreated Then
        
        'recupère la date en currency
        curV = Me.GetFolderDateAsCurrency(sFolder, DateCreated)
        
        tSOp = tS.DateCreatedOperator
        
        If curV Then
            'compare avec la date à rechercher
            '-1==> curDate < curDateReal
            '0 ==> curDate = curDateReal
            '1 ==> curDate > curDateReal
            Ret = CompareFileTime(varDates(0), curV)
            
            If Ret = 0 Then
                If (tSOp = [ = ]) Or (tSOp = [ <= ]) Or (tSOp = [ >= ]) Then _
                    GoTo DateCreatedIsOk
            ElseIf Ret = 1 Then
                If (tSOp = [ < ]) Or (tSOp = [ <= ]) Then GoTo DateCreatedIsOk
            Else
                If (tSOp = [ > ]) Or (tSOp = [ >= ]) Then GoTo DateCreatedIsOk
            End If
        End If
        
        'si on est là, c'est que la comparaison n'est pas bonne
        IsFolderOk = False
        Exit Function
    End If
    
DateCreatedIsOk:

    If (tS.Criteria And SearchDateLastAccessed) = SearchDateLastAccessed Then
        
        'recupère la date en currency
        curV = Me.GetFolderDateAsCurrency(sFolder, DateLastAccessed)
        
        tSOp = tS.DateCreatedOperator
        
        If curV Then
            'compare avec la date à rechercher
            '-1==> curDate < curDateReal
            '0 ==> curDate = curDateReal
            '1 ==> curDate > curDateReal
            Ret = CompareFileTime(varDates(0), curV)
            
            If Ret = 0 Then
                If (tSOp = [ = ]) Or (tSOp = [ <= ]) Or (tSOp = [ >= ]) Then _
                    GoTo DateLastAccessedIsOk
            ElseIf Ret = 1 Then
                If (tSOp = [ < ]) Or (tSOp = [ <= ]) Then _
                    GoTo DateLastAccessedIsOk
            Else
                If (tSOp = [ > ]) Or (tSOp = [ >= ]) Then _
                    GoTo DateLastAccessedIsOk
            End If
        End If
        
        'si on est là, c'est que la comparaison n'est pas bonne
        IsFolderOk = False
        Exit Function
    End If
    
DateLastAccessedIsOk:
    
    If (tS.Criteria And SearchDateLastModified) = SearchDateLastModified Then
        
        'recupère la date en currency
        curV = Me.GetFolderDateAsCurrency(sFolder, DateLastModified)
        
        tSOp = tS.DateCreatedOperator
        
        If curV Then
            'compare avec la date à rechercher
            '-1==> curDate < curDateReal
            '0 ==> curDate = curDateReal
            '1 ==> curDate > curDateReal
            Ret = CompareFileTime(varDates(0), curV)
            
            If Ret = 0 Then
                If (tSOp = [ = ]) Or (tSOp = [ <= ]) Or (tSOp = [ >= ]) Then _
                    GoTo DateLastModifiedIsOk
            ElseIf Ret = 1 Then
                If (tSOp = [ < ]) Or (tSOp = [ <= ]) Then _
                    GoTo DateLastModifiedIsOk
            Else
                If (tSOp = [ > ]) Or (tSOp = [ >= ]) Then _
                    GoTo DateLastModifiedIsOk
            End If
        End If
        
        'si on est là, c'est que la comparaison n'est pas bonne
        IsFolderOk = False
        Exit Function
    End If

DateLastModifiedIsOk:
    
    IsFolderOk = True     'parcours du combattant terminé !
    
End Function

'=======================================================
'Transforme une string (date) en currency
'=======================================================
Private Function DateString2Currency(ByVal sDate As String) As Currency
Dim FT As FILETIME
Dim d As Date
Dim t As Date
Dim ST As SYSTEMTIME
    
    If sDate = vbNullString Then Exit Function
    
    d = DateValue(sDate)
    t = TimeValue(sDate)

    'transfome déjà en systemtime
    With ST
        .wDay = Day(d)
        .wMonth = Month(d)
        .wYear = Year(d)
        .wMinute = Minute(t)
        .wHour = Hour(t)
        .wSecond = Second(t)
    End With
    
    'passe en filetime
    Call SystemTimeToFileTime(ST, FT)
    
    'passe en heure locale
    'Call FileTimeToLocalFileTime(FT, FT)
    
    'passe en currency
    Call CopyMemory(DateString2Currency, FT, LenB(FT))
        
End Function
