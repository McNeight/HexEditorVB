VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileInfos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

'=======================================================
'//MODULE DE CLASSE PERMETTANT DE RECUPERER LES INFOS SUR LES FICHIERS
'// by violent_ken
'=======================================================



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
Private Const FILE_BEGIN                    As Long = 0
Private Const HKEY_CLASSES_ROOT             As Long = &H80000000
Private Const CREATE_ALWAYS                 As Long = 2
Private Const FILE_FLAG_BACKUP_SEMANTICS = &H200000
Private Const FILE_ATTRIBUTE_DIRECTORY      As Long = &H10
Private Const cFILE_ATTRIBUTE_TEMPORARY     As Long = &H100
Private Const cFILE_ATTRIBUTE_SYSTEM        As Long = &H4
Private Const cFILE_ATTRIBUTE_READONLY      As Long = &H1
Private Const cFILE_ATTRIBUTE_NORMAL        As Long = &H80
Private Const cFILE_ATTRIBUTE_HIDDEN        As Long = &H2
Private Const cFILE_ATTRIBUTE_COMPRESSED    As Long = &H800
Private Const cFILE_ATTRIBUTE_ARCHIVE       As Long = &H20
Private Const SEE_MASK_INVOKEIDLIST         As Long = &HC
Private Const SEE_MASK_NOCLOSEPROCESS       As Long = &H40
Private Const SEE_MASK_FLAG_NO_UI           As Long = &H400
Private Const FO_DELETE                     As Long = &H3
Private Const FOF_ALLOWUNDO                 As Long = &H40
Private Const FO_COPY                       As Long = &H2&
Private Const FO_MOVE                       As Long = &H1&
Private Const FOF_SIMPLEPROGRESS            As Long = &H100&
Private Const OFN_EXPLORER                  As Long = &H80000


'=======================================================
'TYPES & ENUMS
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
End Enum
Public Enum SPECIAL_FOLDER_TYPE
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
Public Enum TypeInfo
  CompanyName
  FileDescription
  FileVersion
  InternalName
  LegalCopyright
  OriginalFileName
  ProductName
  ProductVersion
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
    pszDisplayName As Long
    lpszTitle As Long
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
    FLAGS As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
End Type
Private Type LARGE_INTEGER
    LowDWORD As Long
    HighDWORD As Long
End Type


'=======================================================
'APIS
'=======================================================
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As Fichier) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetFileSizeEx Lib "kernel32" (ByVal hFile As Long, lpFileSize As Currency) As Boolean
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileSpec As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare Function PathIsRoot Lib "shlwapi.dll" Alias "PathIsRootA" (ByVal pszPath As String) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function GetFileInformationByHandle Lib "kernel32" (ByVal hFile As Long, lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long
Private Declare Function PathStripToRoot Lib "shlwapi.dll" Alias "PathStripToRootA" (ByVal pszPath As String) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Declare Function GetCompressedFileSize Lib "kernel32" Alias "GetCompressedFileSizeA" (ByVal lpFileName As String, lpFileSizeHigh As Long) As Long
Private Declare Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function PathFindExtension Lib "shlwapi" Alias "PathFindExtensionA" (ByVal pPath As String) As Long
Private Declare Function lstrcpyA Lib "kernel32" (ByVal RetVal As String, ByVal Ptr As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal Ptr As Any) As Long
Private Declare Sub PathStripPath Lib "shlwapi.dll" Alias "PathStripPathA" (ByVal pszPath As String)
Private Declare Function PathRemoveFileSpec Lib "shlwapi" Alias "PathRemoveFileSpecA" (ByVal pPath As String) As Long
Private Declare Function PathRemoveBackslash Lib "shlwapi.dll" Alias "PathRemoveBackslashA" (ByVal pszPath As String) As Long
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long
Private Declare Function SHRunDialog Lib "shell32" Alias "#61" (ByVal hOwner As Long, ByVal Unknown1 As Long, ByVal Unknown2 As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal uFlags As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



'=======================================================
'//PUBLIC PROCEDURES & FUNCTIONS
'=======================================================


'=======================================================
'liste les sous-dossiers contenus dans un dossier
'=======================================================
Public Function GetSubFolders(ByVal sFolder As String, ByRef Folders() As String, Optional ByVal SubFolder As Boolean = True) As Long

    On Error Resume Next    'évite les erreurs si le tableau est vide
    
    If Right$(sFolder, 1) <> "\" Then sFolder = sFolder & "\"     'ajoute le "\" si nécessaire
    
    If Me.FolderExists(sFolder) = False Then
        'alors pas la peine de chercher
        GetSubFolders = 0
        Exit Function
    Else
        EnumFolders sFolder, Folders(), True, SubFolder   'énumère les fichiers
        GetSubFolders = UBound(Folders())
    End If
End Function

'=======================================================
'liste les fichiers contenus dans un dossier
'de 1 à ubound
'=======================================================
Public Function EnumFilesFromFolder(ByVal sFolder As String, Files() As String, Optional SubFolder As Boolean = True) As Long

    On Error Resume Next    'évite les erreurs si le tableau est vide
    
    'ajoute un "\" si nécessaire
    If Right$(sFolder, 1) <> "\" Then sFolder = sFolder & "\"
    
    If Me.FolderExists(sFolder) = False Then
        'alors pas la peine de chercher
        EnumFilesFromFolder = 0
        Exit Function
    Else
        EnumFiles sFolder, Files(), True, SubFolder   'énumère les fichiers
        EnumFilesFromFolder = UBound(Files())
    End If
End Function

'=======================================================
'change l'attribut fichier/dossier
'=======================================================
Public Function ChangeAttributes(ByVal sPath As String, lParam As FILE_ATTRIBUTE) As Long
    ChangeAttributes = SetFileAttributes(sPath, lParam)
End Function

'=======================================================
'copie la string dans un fichier
'=======================================================
Public Function SaveStringInfile(ByVal sFile As String, ByVal sContenuFichier As String, Optional ByVal bDeleteActualFile As Boolean = False) As Long
Dim lFile As Long
    
    'si le fichier est inexistant ou bien si on doit l'overwrite
    If Me.FileExists(sFile) = False Or bDeleteActualFile = True Then
        Me.CreateEmptyFile sFile, True
    End If

    'récupère un numéro libre
    lFile = FreeFile
    
    'ouvre le fichier
    Open sFile For Append As lFile
    
    'sauvegarde la string
    Print #lFile, sContenuFichier
    
    'referme le fichier
    Close lFile
    
End Function

'=======================================================
'copie la string dans un fichier (type de donnée binaire)
'=======================================================
Public Function SaveDATAinFile(ByVal sFile As String, ByVal sContenuFichier As String, Optional ByVal bDeleteActualFile As Boolean = False) As Long
Dim lFile As Long
    
    'si le fichier est inexistant ou bien si on doit l'overwrite
    If Me.FileExists(sFile) = False Or bDeleteActualFile = True Then
        Me.CreateEmptyFile sFile, True
    End If

    'récupère un numéro libre
    lFile = FreeFile
    
    'ouvre le fichier
    Open sFile For Binary Access Write As #lFile

    'sauvegarde la string
    Put #lFile, , sContenuFichier
    
    'referme le fichier
    Close lFile
    
End Function

'=======================================================
'obtient les infos sur un dossier
'=======================================================
Public Function GetFolder(ByVal sFolder As String) As clsFolder
Dim tmp As clsFolder
Dim lFolder As Long
Dim FolderInfo As BY_HANDLE_FILE_INFORMATION
Dim sBuf As String

    Set tmp = New clsFolder

    On Error GoTo ErrGestion
    
    'handle vers le fichier désiré
    lFolder = CreateFile(sFolder, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0&)
    
    'obtient les infos sur le fichier
    GetFileInformationByHandle lFolder, FolderInfo
    
    With tmp
        .Folder = sFolder
        .CreationDate = FileTimeToString(FolderInfo.ftCreationTime, True)
        .LastAccessDate = FileTimeToString(FolderInfo.ftLastAccessTime, True)
        .LastModificationDate = FileTimeToString(FolderInfo.ftLastWriteTime, True)
        .FolderAttributes = FolderInfo.dwFileAttributes
        .ParentDir = Me.GetParentDirectory(sFolder)
        .FolderDrive = Me.GetDriveFormPath(sFolder)
        Me.GetShortName sFolder, sBuf
        .ShortPath = sBuf
        .IsHidden = (.FolderAttributes And vbHidden) = vbHidden
        .IsSystem = (.FolderAttributes And vbSystem) = vbSystem
        .IsReadOnly = (.FolderAttributes And vbReadOnly) = vbReadOnly
        .IsNormal = (.FolderAttributes And vbNormal) = vbNormal And Not (.IsHidden) And Not (.IsReadOnly) And Not (.IsSystem)
    End With
    Set GetFolder = tmp
    
ErrGestion:
    CloseHandle lFolder   'ferme le handle
End Function

'=======================================================
'obtient les infos sur un fichier
'=======================================================
Public Function GetFile(ByVal sFile As String) As clsFile
Dim tmp As clsFile
Dim lFile As Long
Dim FileInfo As BY_HANDLE_FILE_INFORMATION
Dim sBuf As String

    Set tmp = New clsFile
    Set GetFile = New clsFile

   ' On Error GoTo ErrGestion
    
    'handle vers le fichier désiré
    lFile = CreateFile(sFile, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, ByVal &O0, ByVal 0&)
    
    'obtient les infos sur le fichier
    GetFileInformationByHandle lFile, FileInfo
    CloseHandle lFile

    With tmp
        .File = sFile
        .CreationDate = FileTimeToString(FileInfo.ftCreationTime, True)
        .LastAccessDate = FileTimeToString(FileInfo.ftLastAccessTime, True)
        .LastModificationDate = FileTimeToString(FileInfo.ftLastWriteTime, True)
        .FileAttributes = FileInfo.dwFileAttributes
        .FileSize = Me.GetFileSize(sFile)
        .FileExtension = Me.GetFileExtension(sFile)
        .FileDirectory = Me.GetParentDirectory(sFile)
        .FileDrive = Me.GetDriveFormPath(sFile)
        Me.GetShortName sFile, sBuf
        .ShortName = sBuf
        Me.GetShortPath sFile, sBuf
        .ShortPath = sBuf
        .IsHidden = (.FileAttributes And vbHidden) = vbHidden
        .IsSystem = (.FileAttributes And vbSystem) = vbSystem
        .IsReadOnly = (.FileAttributes And vbReadOnly) = vbReadOnly
        .IsNormal = (.FileAttributes And vbNormal) = vbNormal And Not (.IsHidden) And Not (.IsReadOnly) And Not (.IsSystem)
        .FileType = Me.GetFileType(sFile)
        .AssociatedExecutableProgram = Me.GetAssociatedExecutableProgram(sFile)
        .EXECompanyName = GetEXEdescription(sFile, CompanyName)
        .EXEFileDescription = GetEXEdescription(sFile, FileDescription)
        .EXEFileVersion = GetEXEdescription(sFile, FileVersion)
        .EXEInternalName = GetEXEdescription(sFile, InternalName)
        .EXELegalCopyright = GetEXEdescription(sFile, LegalCopyright)
        .EXEOriginalFileName = GetEXEdescription(sFile, OriginalFileName)
        .EXEProductName = GetEXEdescription(sFile, ProductName)
        .EXEProductVersion = GetEXEdescription(sFile, ProductVersion)
        .FileCompressedSize = GetCompressedFileSize(sFile, ByVal 0&)
    End With
    Set GetFile = tmp
    
ErrGestion:
    CloseHandle lFile   'ferme le handle
End Function

'=======================================================
'obtient les infos sur un fichier de manière rapide (conçu pour la recherche de fichiers)
'0==> creation
'1==> dernier accès
'2==> dernière modification
'=======================================================
Public Function GetFileDate(ByVal sFile As String, ByVal lType As Long, Optional ByVal LocalTime As Boolean = True) As Currency
Dim lFile As Long
Dim FileInfo As BY_HANDLE_FILE_INFORMATION
Dim FT As FILETIME

    'handle vers le fichier désiré
    lFile = CreateFile(sFile, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, ByVal &O0, ByVal 0&)
    
    'obtient les infos sur le fichier
    GetFileInformationByHandle lFile, FileInfo
    CloseHandle lFile

    'change les FILE_TIME en Currency
    With FileInfo
        If lType = 0 Then
            FT = .ftCreationTime
        ElseIf lType = 1 Then
            FT = .ftLastAccessTime
        Else
            FT = .ftLastWriteTime
        End If
    End With
    
    If LocalTime Then
        'convertit au temps local
        FileTimeToLocalFileTime FT, FT
    End If

    'renvoie le résultat en currency
    GetFileDate = LI2Currency(FT)
    
End Function

'=======================================================
'loade un fichier en string
'=======================================================
Public Function LoadFileInString(ByVal sFile As String, Optional ByRef sLoadIsOk As Long) As String
Dim hFile As Long
Dim lSize As Long
Dim Ret As Long
Dim s As String
    
    'On Error GoTo ErrGestion
    
    'obtient le handle du fichier
    hFile = CreateFile(sFile, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    
    'obtient sa taille pour créer un buffer
    lSize = Me.GetFileSize(sFile)
 
    'initialise le pointeur sur le fichier (au début)
    SetFilePointer hFile, 0, 0, FILE_BEGIN
    
    'créé le buffer
    s = String$(lSize, 0)
    
    'charge le fichier
    Ret = ReadFile(hFile, ByVal s, lSize, Ret, ByVal 0&)
     
    LoadFileInString = s
    
    If Not (IsMissing(sLoadIsOk)) Then
        'alors on renvoie si la fonction a marché ou pas
        sLoadIsOk = IIf(Ret > 0 And hFile > 0, 1, 0)
    End If
    
ErrGestion:
    'ferme le handle
    CloseHandle hFile
End Function

'=======================================================
'renomme un fichier/dossier
'=======================================================
Public Sub Rename(ByVal sOldName As String, ByVal sNewName As String)
    On Error Resume Next
    Name sOldName As sNewName
End Sub

'=======================================================
'obtient le dossier parent du dossier spécifié
'=======================================================
Public Function GetParentDirectory(ByVal sFolder As String) As String
Dim l As Long

    PathRemoveBackslash sFolder
    
    l = InStrRev(FormatedString(sFolder), "\")
    If l = 0 Then GetParentDirectory = vbNullString Else GetParentDirectory = Left$(sFolder, l)
End Function

'=======================================================
'créé un fichier vierge et renvoie son handle
'=======================================================
Public Function CreateEmptyFile(ByVal sFile As String, Optional ByVal OverWrite As Boolean = False) As Long
Dim PARAM As Long
Dim i As Long
Dim lBuf2 As Long
Dim lFile As Long
    
    If OverWrite Then
        'fichier déjà existant
        PARAM = CREATE_ALWAYS
    Else
        PARAM = CREATE_NEW
    End If
    
    'créé le fichier
    CreateEmptyFile = CreateFile(sFile, GENERIC_WRITE, FILE_SHARE_WRITE Or FILE_SHARE_READ, ByVal 0&, PARAM, ByVal 0&, ByVal 0&)
    CloseHandle CreateEmptyFile 'ferme le handle...

End Function

'=======================================================
'détermine si le path est un root
'=======================================================
Public Function IsRoot(ByVal sPath As String) As Boolean
    IsRoot = CBool(PathIsRoot(sPath))
End Function

'=======================================================
'obtient le programme associé à un fichier
'=======================================================
Public Function GetAssociatedExecutableProgram(ByVal sFile As String) As String
Dim sBuf As String

    sBuf = String$(255, 0)  'création d'un buffer
    
    'trouve l'éxécutable associé
    FindExecutable sFile, vbNullString, sBuf
    
    'formate la string
    GetAssociatedExecutableProgram = FormatedString(sBuf)

End Function

'=======================================================
'obtient le nom court
'=======================================================
Public Function GetShortName(ByVal sPath As String, ByRef sOutputBuffer As String, Optional ByVal lSize As Long = 164) As Long
Dim s As String
    s = String$(lSize, 0)   'buffer
    GetShortName = GetShortPathName(sPath, s, lSize)
    sOutputBuffer = FormatedString(s)
End Function

'=======================================================
'obtient le path court
'=======================================================
Public Function GetShortPath(ByVal sPath As String, ByRef sOutputBuffer As String, Optional ByVal lSize As Long = 164) As Long
Dim s As String
    s = String$(lSize, 0)   'buffer
    GetShortPath = GetShortPathName(Me.GetFolderFromPath(sPath), s, lSize)
    sOutputBuffer = FormatedString(s)
End Function

'=======================================================
'obtient les dossier depuis un path
'=======================================================
Public Function GetFolderFromPath(ByVal sPath As String) As String
    PathRemoveFileSpec sPath
    GetFolderFromPath = FormatedString(sPath)
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
End Function

'=======================================================
'obtient le type du fichier
'=======================================================
Public Function GetFileType(ByVal sFile As String) As String
Dim sExt As String
Dim hKey As Long
Dim hKey2 As Long

    On Error Resume Next
    
    sExt = Right$(Me.GetFileExtension(sFile), Len(Me.GetFileExtension(sFile)) - 1) 'obtient l'extension
    
    'ouvre la clé du registre et récupère la string concernée
    If RegOpenKey(HKEY_CLASSES_ROOT, "." & sExt, hKey) = 0 Then
        GetFileType = RegQueryStringValue(hKey, vbNullString)
        If RegOpenKey(HKEY_CLASSES_ROOT, GetFileType, hKey2) = 0 Then
            GetFileType = RegQueryStringValue(hKey2, vbNullString)
        End If
    End If
    
    'referme les clés
    RegCloseKey hKey2
    RegCloseKey hKey
End Function

'=======================================================
'affiche la boite de dialogue "Choix du dossier Dossier..."
'=======================================================
Public Function BrowseForFolder(ByVal sTitle As String, ByVal hWnd As Long) As String
Dim lIDList As Long
Dim sBuffer As String
Dim tBrowseInfo As BrowseInfo

    'affectation des propriétés définies par les arguments de la fonction
    With tBrowseInfo
        .hwndOwner = hWnd
        .lpszTitle = lstrcat(sTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    
    'affichage de la boite de dialogue
    lIDList = SHBrowseForFolder(tBrowseInfo)
    
    'récupération du résultat
    If (lIDList) Then
        sBuffer = String$(260, vbNullChar)
        SHGetPathFromIDList lIDList, sBuffer
        BrowseForFolder = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        CoTaskMemFree lIDList
    End If
End Function

'=======================================================
'obtient la taille d'un fichier, même si celle ci est supérieure à 4Go (long)
'=======================================================
Public Function GetFileSize(ByVal sFile As String) As Currency
Dim lngFile As Long
Dim curSize As Currency

    'obtient le handle du fichier
    lngFile = CreateFile(sFile, GENERIC_READ, FILE_SHARE_WRITE Or FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    
    'obtient la taille par API
    GetFileSizeEx lngFile, curSize
    
    'ferme le handle ouvert
    CloseHandle lngFile
    
    GetFileSize = curSize * 10000 'multiplie par 10^4 pour obtenir un nombre entier
End Function

'=======================================================
'teste la disponibilité d'un fichier (et aussi sa présence)
'=======================================================
Public Function IsFileAvailable(ByVal sFile As String) As Boolean
Dim lFile As Long
    
    IsFileAvailable = True
    
    'obtient un handle vers le fichier à ouvrir
    lFile = CreateFile(sFile, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    
    If lFile = -1 Then
        'fichier inexistant, ou en tout cas inaccessible
        IsFileAvailable = False
    End If

    'referme le handle
    CloseHandle lFile
    
End Function

'=======================================================
'teste l'existence d'un dossier
'=======================================================
Public Function FolderExists(ByVal sFolder As String) As Boolean
    FolderExists = CBool(PathIsDirectory(sFolder))
End Function

'=======================================================
'teste l'existence un fichier
'=======================================================
Public Function FileExists(ByVal sFile As String) As Boolean
    FileExists = CBool(PathFileExists(sFile) And Not (Me.FolderExists(sFile)))
End Function

'=======================================================
'renvoie la terminaison d'un fichier/path
'=======================================================
Public Function GetFileExtension(ByVal sFile As String) As String
    GetFileExtension = String$(lstrlenA(ByVal PathFindExtension(sFile)), 0)
    Call lstrcpyA(ByVal GetFileExtension, ByVal PathFindExtension(sFile))
End Function

'=======================================================
'obtient un fichier depuis un path
'=======================================================
Public Function GetFileFromPath(ByVal sPath As String) As String
    PathStripPath sPath
    GetFileFromPath = FormatedString(sPath)
End Function

'=======================================================
'obtient un drive depuis un path
'=======================================================
Public Function GetDriveFormPath(ByVal sPath As String) As String
    PathStripToRoot sPath
    GetDriveFormPath = FormatedString(sPath)
End Function

'=======================================================
'tue un fichier
'=======================================================
Public Function KillFile(ByVal sFile As String) As Long
    SetFileAttributes sFile, FILE_ATTRIBUTE_NORMAL  'permet de virer un attribut qui empecherait la suppression
    KillFile = DeleteFile(sFile)
End Function

'=======================================================
'obtient une info sur un executable
'=======================================================
Public Function GetEXEFileInfo(ByVal sFile As String, Information As TypeInfo) As String
    GetEXEFileInfo = GetEXEdescription(sFile, Information)
End Function

'=======================================================
'obtient un path spécial
'=======================================================
Public Function GetSpecialFolder(Folder As SPECIAL_FOLDER_TYPE, Optional ByVal hWnd As Long) As String
Dim r As Long
Dim sPath As String
Dim IDL As ITEMIDLIST

    If IsMissing(hWnd) Then hWnd = 100

    'obtient le path
    r = SHGetSpecialFolderLocation(hWnd, Folder, IDL)
    
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
'créé un dossier
'=======================================================
Public Function CreateFolder(ByVal sFolder As String) As Long
Dim secSec As SECURITY_ATTRIBUTES
    CreateFolder = CreateDirectory(sFolder, secSec)
End Function

'=======================================================
'énumère les sous dossiers d'un dossier
'de 1 à ubound
'=======================================================
Public Sub EnumFolders(ByVal Directory As String, Folders() As String, Optional Begin As Boolean = False, Optional SubFolder As Boolean = True)
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
                EnumFolders Directory & Mid$(FileInfo.cFileName, 1, InStr(FileInfo.cFileName, vbNullChar) - 1) & "\", Folders, False
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
                    EnumFolders Directory & Mid$(FileInfo.cFileName, 1, InStr(FileInfo.cFileName, vbNullChar) - 1) & "\", Folders, False
                End If
                sDir = Directory & Mid$(FileInfo.cFileName, 1, InStr(FileInfo.cFileName, vbNullChar) - 1)
                If Right$(sDir, 3) <> "\.." Then
                    ub = ub + 1
                    ReDim Preserve Folders(ub)
                    'chemin complet
                    Folders(ub) = sDir
                End If
            End If
        Loop
    End If
    'ferme l'énumèration
    FindClose hFind
End Sub

'=======================================================
'fonction permettant d'afficher une boite de dialogue "propriété"
'=======================================================
Public Function ShowFileProperty(ByVal sFile As String, ByVal lHwnd As Long) As Long
Dim SEI As SHELLEXECUTEINFO

    On Error GoTo ErrGestion
    
    'définition des paramètres de la boite de dialogue
    With SEI
        'défnition
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hWnd = lHwnd
        .lpVerb = "properties"
        .lpFile = sFile
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = 0
        .lpIDList = 0
    End With
    ShowFileProperty = ShellExecuteEX(SEI)   'affichage
    
ErrGestion:
End Function

'=======================================================
'fonction permettant d'ouvrir un fichier
'=======================================================
Public Function ShellOpenFile(ByVal sFile As String, ByVal hWnd As Long, Optional sParameters As String = vbNullString, Optional ByVal sDirectory As String = vbNullString) As Long
    ShellOpenFile = ShellExecute(hWnd, "open", sFile, sParameters, sDirectory, 1)
End Function

'=======================================================
'déplace un fichier ou un dossier
'=======================================================
Public Function MoveFileOrFolder(ByVal sFileOrFolder As String, ByVal sDestination As String) As Long
Dim sFileToDelete As Fichier

    'définit l'objet Fichier
    With sFileToDelete
        .wFunc = FO_MOVE
        .pFrom = sFileOrFolder
        .pTo = sDestination
        .fFlags = FOF_ALLOWUNDO And FOF_SIMPLEPROGRESS
    End With
    
    'procède à la suppression
    MoveFileOrFolder = SHFileOperation(sFileToDelete)
End Function

'=======================================================
'suppression d'un dossier
'=======================================================
Public Function KillFolder(ByVal sFileOrFolder As String) As Long
Dim sFileToDelete As Fichier

    'définit l'objet Fichier
    With sFileToDelete
        .wFunc = FO_DELETE
        .pFrom = sFileOrFolder
        .fFlags = FOF_ALLOWUNDO And FOF_SIMPLEPROGRESS
    End With
    
    'procède à la suppression
    KillFolder = SHFileOperation(sFileToDelete)
End Function

'=======================================================
'copie un fichier/dossier
'=======================================================
Public Function CopyFileOrFolder(ByVal sFileOrFolder As String, ByVal sDestination As String) As Long
Dim sFileToDelete As Fichier

    'définit l'objet Fichier
    With sFileToDelete
        .wFunc = FO_COPY
        .pFrom = sFileOrFolder
        .pTo = sDestination
        .fFlags = FOF_ALLOWUNDO And FOF_SIMPLEPROGRESS
    End With
    
    'procède à la suppression
    CopyFileOrFolder = SHFileOperation(sFileToDelete)
End Function

'=======================================================
'déplace vers la corbeille un fichier/dossier
'=======================================================
Public Function MoveToTrash(ByVal sFileOrFolder As String) As Long
Dim sFileToDelete As Fichier

    'définit l'objet Fichier
    With sFileToDelete
        .wFunc = FO_DELETE
        .pFrom = sFileOrFolder
        .fFlags = FOF_ALLOWUNDO
    End With
    
    'procède à la suppression
    MoveToTrash = SHFileOperation(sFileToDelete)
End Function

'=======================================================
'affichage de la boite de dialogue Executer...
'=======================================================
Public Function ShowRunBox(ByVal hWnd As Long) As Long
    ShowRunBox = SHRunDialog(hWnd, 0, 0, StrConv("Exécuter", vbUnicode), _
        StrConv("Sélectionnez un élément à lancer (dossier, document, exécutable...) et Process Guardian l'ouvrira pour vous.", _
        vbUnicode), 0)
End Function

'=======================================================
'affichage de la boite de dialogue Ouvrir...
'renvoie le FileName
'=======================================================
Public Function ShowOpen(ByVal sTitle As String, ByVal hWnd As Long, _
    Optional ByVal sFilter As String, Optional ByVal sInitDir As String, _
    Optional ByRef bHasPushedCancel As Boolean) As String
    
Dim OFName As OPENFILENAME

    With OFName
        'taille du type
        .lStructSize = Len(OFName)
        
        'affecte les valeurs des variables de configuration de la boite de dialogue au type
        .hwndOwner = hWnd
        .hInstance = App.hInstance
        sFilter = Replace$(sFilter, "|", vbNullChar) 'convertit la string
        .lpstrFilter = sFilter
        .nMaxFile = 255
        .nMaxFileTitle = 255
        .lpstrFileTitle = Space$(254)
        .lpstrInitialDir = sInitDir
        .lpstrTitle = sTitle
        
        .lpstrFile = Space$(254)    'créé un buffer /!\ NE PAS CHANGER la taille du buffer

    End With

    'ouvre la boite de dialogue
    bHasPushedCancel = Not (CBool(GetOpenFileName(OFName))) 'récupère si l'utilisateur à cliqué sur "annuler"
    If bHasPushedCancel Then Exit Function
    
    'récupère le FileName
    ShowOpen = Left$(OFName.lpstrFile, InStr(OFName.lpstrFile, vbNullChar) - 1)
End Function

'=======================================================
'affichage de la boite de dialogue Enregistrer...
'renvoie le FileName AVEC l'extension choisie (ne double pas l'extension si comprise dans le nom)
'=======================================================
Public Function ShowSave(ByVal sTitle As String, ByVal hWnd As Long, _
    Optional ByVal sFilter As String, Optional ByVal sInitDir As String, _
    Optional ByVal sDefaultFileName As String, _
    Optional ByRef bHasPushedCancel As Boolean) As String
    
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
        .lpstrFilter = Replace$(sFilter, "|", vbNullChar) 'convertit la string
        .nMaxFile = 255
        .nMaxFileTitle = 255
        .lpstrFileTitle = Space$(254)
        .lpstrInitialDir = sInitDir
        .lpstrTitle = sTitle
        .lpstrFile = Space$(254)    'créé un buffer pour le fichier. Ne pas modifier 254.
    End With

    'affichage de la boite de dialogue
    bHasPushedCancel = Not (CBool(GetSaveFileName(OFName)))
    
    'récupère le filename formaté
    s3 = Left$(OFName.lpstrFile, InStr(OFName.lpstrFile, vbNullChar) - 1)
    
    'récupère le filename
    'on récupère l'extension choisie
    s2() = Split(sFilter, "|", , vbBinaryCompare)
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
'//PRIVATE PROCEDURES & FUNCTIONS
'=======================================================
'=======================================================
'transforme une date en FILETIME vers une date en string
'=======================================================
Private Function FileTimeToString(fDate As FILETIME, Optional ByVal bConvertToLocal As Boolean = True) As String
Dim sDate As SYSTEMTIME
Dim sDay As String
Dim sMonth As String
Dim sYear As String
Dim sHour As String
Dim sMinute As String
Dim sSecond As String
Dim s As String

    If bConvertToLocal Then
        'conversion en LocalFileTime (temps universel ==> temps local)
        FileTimeToLocalFileTime fDate, fDate
    End If
    
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
'lit une valeur dans le registre
'=======================================================
Private Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
Dim lValueType As Long
Dim strBuf As String
Dim lDataBufSize As Long
Dim lResult As Long
Dim strData As String
Dim lData As Long

    'obtiention des infos sur la clé
    If 0 = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal 0&, lDataBufSize) Then
        'buffer
        strBuf = Space$(lDataBufSize)
        'contenu de la clé
        lResult = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal strBuf, lDataBufSize)
        If lResult = 0 Then
            'formatage de la string
            RegQueryStringValue = Left$(strBuf, InStr(strBuf, vbNullChar) - 1)
        End If
    End If
End Function

'=======================================================
'obtient une info sur un executable - from MSDN
'=======================================================
Private Function GetEXEdescription(ByVal Cible As String, ByVal TypeInfo As TypeInfo) As String
Dim Buffer As String, Lang_Charset_String As String
Dim Rc As Long, HexNumber As Long, p As Long
Dim strVersionInfo As String, strtemp As String
Dim bufferlen As Long, Dummy As Long
Dim sBuffer() As Byte
Dim ByteBuffer(255) As Byte
Dim Cible2 As String

    Select Case TypeInfo
        Case CompanyName
            strVersionInfo = "CompanyName"
        Case FileDescription
            strVersionInfo = "FileDescription"
        Case FileVersion
            strVersionInfo = "FileVersion"
        Case InternalName
            strVersionInfo = "InternalName"
        Case LegalCopyright
            strVersionInfo = "LegalCopyright"
        Case OriginalFileName
            strVersionInfo = "OriginalFileName"
        Case ProductName
            strVersionInfo = "ProductName"
        Case ProductVersion
            strVersionInfo = "ProductVersion"
    End Select
    
    Cible2 = String$(255, 0)  'buffer
    'trouve l'éxécutable
    FindExecutable Cible, vbNullString, Cible2
    
    'Vérifie si les informations sur la version existent.
    bufferlen = GetFileVersionInfoSize(FormatedString(Cible2), Dummy)
    If bufferlen < 1 Then Exit Function
    
    ReDim sBuffer(bufferlen)
    Rc = GetFileVersionInfo(Cible, 0&, bufferlen, sBuffer(0))
    If Rc = 0 Then Exit Function
    
    '"\VarFileInfo\Translation" permet de récupérer la langue utilisée et
    'le type de caractère:
    'Par exemple, on peut récupérer la valeur 040C1200 où 040C identifie la
    'langue française et 1200 identifie le jeu de caractères Unicode
    '(Les valeurs des identifiants de langue et de jeu de caractères sont
    'données dans l'aide WIN SDK 32 HELP pour la structure VERSIONINFO).
    Rc = _
    VerQueryValue(sBuffer(0), "\VarFileInfo\Translation", p, bufferlen)
    
    If Rc = 0 Then Exit Function
    
    MoveMemory ByteBuffer(0), p, bufferlen
    
    HexNumber = ByteBuffer(2) + ByteBuffer(3) * &H100 + ByteBuffer(0) * _
    &H10000 + ByteBuffer(1) * &H1000000
    
    Lang_Charset_String = Hex(HexNumber)
    
    Do While Len(Lang_Charset_String) < 8
        Lang_Charset_String = "0" & Lang_Charset_String
    Loop
    
    Buffer = String$(255, 0)
    strtemp = "\StringFileInfo\" & Lang_Charset_String & "\" & strVersionInfo
    Rc = VerQueryValue(sBuffer(0), strtemp, p, bufferlen)
    
    If Rc = 0 Then Exit Function
    
    lstrcpy Buffer, p
    Buffer = Mid$(Buffer, 1, InStr(Buffer, vbNullChar) - 1)
    
    GetEXEdescription = Buffer
End Function

'=======================================================
'formatage de string
'=======================================================
Private Function FormatedString(ByVal sString As String) As String

    If InStr(sString, vbNullChar) Then _
    FormatedString = Left$(sString, InStr(sString, vbNullChar) - 1) Else _
    FormatedString = sString
    
    FormatedString = Trim$(FormatedString)
End Function

'=======================================================
'énumère les fichiers d'un dossier
'de 1 à ubound
'=======================================================
Private Sub EnumFiles(ByVal Directory As String, Files() As String, Optional Begin As Boolean = False, Optional SubFolder As Boolean = True)
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
                EnumFiles Directory & Mid$(FileInfo.cFileName, 1, InStr(FileInfo.cFileName, vbNullChar) - 1) & "\", Files, False
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
                    EnumFiles Directory & Mid$(FileInfo.cFileName, 1, InStr(FileInfo.cFileName, vbNullChar) - 1) & "\", Files, False
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
'transforme un largeinterger en currency
'=======================================================
Private Function LI2Currency(liInput As FILETIME) As Currency
    CopyMemory LI2Currency, liInput, LenB(liInput)
End Function
