Attribute VB_Name = "mdlDeclarations"
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
'//MODULE DE DECLARATION
'//DES APIS/ENUM/TYPES/CONSTANTES
'=======================================================



'=======================================================
'//CONSTANTES
'=======================================================

'constantes d'accès à un processus (pour OpenProcess)
Public Const SYNCHRONIZE                        As Long = &H100000
Public Const PROCESS_SUSPEND_RESUME             As Long = &H800
Public Const PROCESS_QUERY_INFORMATION          As Long = 1024&
Public Const PROCESS_VM_READ                    As Long = 16&
Public Const PROCESS_VM_WRITE                   As Long = &H20
Public Const PROCESS_VM_OPERATION               As Long = &H8
Public Const STANDARD_RIGHTS_REQUIRED           As Long = &HF0000 'aussi pour d'autres accès que les processus
Public Const PROCESS_ALL_ACCESS                 As Long = (STANDARD_RIGHTS_REQUIRED Or _
        SYNCHRONIZE Or &HFFF)
Public Const SPECIFIC_RIGHTS_ALL                As Long = &HFFFF
Public Const PROCESS_READ_WRITE_QUERY           As Long = PROCESS_VM_READ Or PROCESS_VM_WRITE Or _
        PROCESS_VM_OPERATION Or PROCESS_QUERY_INFORMATION

'constantes utilisées pour la gestion des menus
Public Const MIIM_ID                            As Long = &H2
Public Const MIIM_TYPE                          As Long = &H10
Public Const MIIM_STATE                         As Long = &H1
Public Const MIIM_SUBMENU                       As Long = &H4
Public Const TPM_LEFTALIGN                      As Long = &H0&
Public Const TPM_RETURNCMD                      As Long = &H100&
Public Const TPM_RIGHTBUTTON                    As Long = &H2&
Public Const MFT_RADIOCHECK                     As Long = &H200&
Public Const MFT_CHECKED                        As Long = &H8&
Public Const MFT_STRING                         As Long = &H0
Public Const MFS_ENABLED                        As Long = &H0
Public Const MF_BYCOMMAND                       As Long = &H0
Public Const MIIM_DATA                          As Long = &H20
Public Const MF_BYPOSITION                      As Long = &H400
Public Const MF_OWNERDRAW                       As Long = &H100
Public Const MFT_OWNERDRAW                      As Long = MF_OWNERDRAW
Public Const MFS_DEFAULT                        As Long = &H1000
Public Const MF_STRING                          As Long = &H0&

Public Const MEM_PUBLIC                         As Long = &H20000
Public Const MEM_COMMIT                         As Long = &H1000
Public Const MEM_RELEASE                        As Long = &H8000
Public Const MEM_DECOMMIT                       As Long = &H4000
Public Const MEM_RESERVE                        As Long = &H2000
Public Const MEM_RESET                          As Long = &H80000
Public Const MEM_TOP_DOWN                       As Long = &H100000

Public Const PAGE_READWRITE                     As Long = &H4
Public Const PAGE_READONLY                      As Long = &H2
Public Const PAGE_EXECUTE                       As Long = &H10
Public Const PAGE_EXECUTE_READ                  As Long = &H20
Public Const PAGE_EXECUTE_READWRITE             As Long = &H40
Public Const PAGE_GUARD                         As Long = &H100
Public Const PAGE_NOACCESS                      As Long = &H1
Public Const PAGE_NOCACHE                       As Long = &H200

'constantes utilisées pour afficher une boite de dialogue Windows
Public Const BIF_RETURNONLYFSDIRS               As Long = 1&
Public Const BIF_DONTGOBELOWDOMAIN              As Long = 2&

'constantes utilisées avec SHGetFileInfos (récupération de l'icone d'un fichier)
Public Const SHGFI_USEFILEATTRIBUTES            As Long = &H10
Public Const SHGFI_DISPLAYNAME                  As Long = &H200
Public Const SHGFI_TYPENAME                     As Long = &H400
Public Const SHGFI_LARGEICON                    As Long = &H0
Public Const SHGFI_ICON                         As Long = &H100
Public Const SHGFI_EXETYPE                      As Long = &H2000
Public Const SHGFI_SYSICONINDEX                 As Long = &H4000
Public Const SHGFI_SMALLICON                    As Long = &H1
Public Const SHGFI_SHELLICONSIZE                As Long = &H4
Public Const BASIC_SHGFI_FLAGS                  As Long = SHGFI_TYPENAME Or _
        SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
Public Const ILD_TRANSPARENT                    As Long = &H1
Public Const IID_IICON                          As String = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"

'messages pour les fenêtres
Public Const WM_MENUSELECT                      As Long = &H11F
Public Const WM_CLOSE                           As Long = &H10
Public Const WM_SHOWWINDOW                      As Long = &H18
Public Const VISIBLEFLAGS                       As Long = &H2 Or &H1 Or &H40 Or &H10
Public Const WM_GETMINMAXINFO                   As Long = &H24
Public Const GWL_WNDPROC                        As Long = -4&

'constantes pour l'impression
Public Const CCHDEVICENAME                      As Long = 32&
Public Const CCHFORMNAME                        As Long = 32&
Public Const PD_RETURNDC                        As Long = &H100
Public Const PD_RETURNIC                        As Long = &H200
Public Const NULL_PTR                           As Long = 0&

'paramètres pour l'API CreateFile
Public Const FILE_READ_ACCESS                   As Long = &H1
Public Const FILE_BEGIN                         As Long = 0&
Public Const FILE_SHARE_READ                    As Long = &H1
Public Const FILE_SHARE_WRITE                   As Long = &H2
Public Const CREATE_NEW                         As Long = 1&
Public Const OPEN_EXISTING                      As Long = 3&
Public Const GENERIC_WRITE                      As Long = &H40000000
Public Const GENERIC_READ                       As Long = &H80000000
Public Const CREATE_ALWAYS                      As Long = 2&
Public Const FILE_END                           As Long = 2&
Public Const TRUNCATE_EXISTING                  As Long = 5&
Public Const FILE_FLAG_WRITE_THROUGH            As Long = &H80000000
Public Const FILE_FLAG_NO_BUFFERING             As Long = &H20000000

'constantes d'attributs de fichiers
Public Const FILE_ATTRIBUTE_TEMPORARY           As Long = &H100
'Public Const FILE_ATTRIBUTE_NORMAL              As Long = &H80     'pas déclaré car déjà dans un enum public de classe
Public Const FILE_ATTRIBUTE_DIRECTORY           As Long = &H10

'constantes nécessaires à la fonction DisplayFileProperty (type SHELLEXECUTEINFO)
Public Const SEE_MASK_INVOKEIDLIST              As Long = &HC
Public Const SEE_MASK_NOCLOSEPROCESS            As Long = &H40
Public Const SEE_MASK_FLAG_NO_UI                As Long = &H400

'pour la récupération de la carte des clusters des fichiers
Public Const ERROR_MORE_DATA                    As Long = 234&
Public Const FSCTL_GET_RETRIEVAL_POINTERS       As Long = 589939

'constantes contenant mes couleurs publiques
Public Const GREEN_COLOR                        As Long = &HC000&
Public Const RED_COLOR                          As Long = &HC0&

'constante générale
Public Const INVALID_HANDLE_VALUE               As Long = -1&





'=======================================================
'//APIs
'=======================================================

'systèmes de temps
Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Public Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Public Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Public Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Public Declare Function CompareFileTime Lib "kernel32" (lpFileTime1 As Currency, lpFileTime2 As Currency) As Long

'APIS pour les menus
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Public Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal hWnd As Long, ByVal lptpm As Any) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemInfoStr Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal ByPosition As Long, ByRef lpMenuItemInfo As MENUITEMINFO_STRINGDATA) As Boolean

'APIs sur les processus
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Public Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

'APIs sur les fichiers/disques
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Public Declare Function ReadFileEx Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpOverlapped As OVERLAPPED, ByVal lpCompletionRoutine As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Public Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Public Declare Function SetFilePointerEx Lib "kernel32" (ByVal hFile As Long, ByVal liDistanceToMove As Currency, ByRef lpNewFilePointer As Currency, ByVal dwMoveMethod As Long) As Long
Public Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Public Declare Function LockFile Lib "kernel32.dll" (ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToLockLow As Long, ByVal nNumberOfBytesToLockHigh As Long) As Long
Public Declare Function UnlockFile Lib "kernel32.dll" (ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToUnlockLow As Long, ByVal nNumberOfBytesToUnlockHigh As Long) As Long

'APIs sur les fichiers
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetFileSizeEx Lib "kernel32" (ByVal hFile As Long, lpFileSize As Currency) As Boolean
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Public Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Public Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Public Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Public Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileSpec As String, ByVal dwFileAttributes As Long) As Long
Public Declare Function GetFileType Lib "kernel32" (ByVal hFile As Long) As Long

'APIs sur les disques
Public Declare Function DeviceIoControl Lib "kernel32.dll" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function GetLogicalDriveStrings Lib "kernel32.dll" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'APIs sur la mémoire
Public Declare Function VirtualAlloc Lib "kernel32.dll" (ByRef lpAddress As Any, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Public Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Public Declare Function VirtualLock Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long) As Long
Public Declare Function VirtualUnlock Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function PathCompactPath Lib "shlwapi.dll" Alias "PathCompactPathA" (ByVal hdc As Long, ByVal pszPath As String, ByVal dx As Long) As Long

'APIs sur la manipulation de strings
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpStringDest As String, ByVal lpStringSrc As Long) As Long
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Public Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As GUID) As Long

'APIs graphiques (création/destroy d'icones)
Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Public Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PICTDESC, riid As GUID, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Public Declare Function ImageList_Draw Lib "Comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal FLAGS&) As Long

'APIs pour l'affichage/gestion des fenêtres
Public Declare Sub InvalidateRect Lib "user32" (ByVal hWnd As Long, ByVal t As Long, ByVal bErase As Long)
Public Declare Sub ValidateRect Lib "user32" (ByVal hWnd As Long, ByVal t As Long)
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'APIs pour "l'environnement Windows"
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHRunDialog Lib "shell32" Alias "#61" (ByVal hOwner As Long, ByVal Unknown1 As Long, ByVal Unknown2 As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal uFlags As Long) As Long
Public Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As Any, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long
Public Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTER_INFO) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long



'=======================================================
'//ENUMS
'=======================================================

'type de table à afficher dans frmTable
Public Enum TableType
    AllTables = 1
    HEX_ASCII = 2
End Enum

'type de passe pour le collage de bytes sur un disque
Public Enum PASSE_ENUM
    FixedByte = 0
    RandomByte = 1
    ListByte = 2
End Enum

'type d'affichage de la form (AlwaysOntop ou pas)
Public Enum ModePlan
    MettreAuPremierPlan = True
    MettreNormal = False
End Enum

'liste des versions de Windows
Public Enum WINDOWS_VERSION
    [Windows Vista]
    [Windows Server 2003]
    [Windows XP]
    [Windows 2000]
    [Windows Me]
    [Windows 98]
    [Windows 95]
    [UnKnown_OS]
End Enum

'définit la méthode de découpe des fichiers (splitter)
Public Enum CUT_METHOD_ENUM
    [Taille fixe]
    [Nombre fichiers fixe]
End Enum

'type de recherche (recherche de fichiers)
Public Enum TYPE_OF_FILE_SEARCH
    [Recherche de fichiers]
    [Recherche de dossiers]
    [Recherche de contenu de fichier]
End Enum



'=======================================================
'//TYPES
'=======================================================

'pour l'obtention des zones mémoire d'un processus
Public Type MEMORY_BASIC_INFORMATION ' 28 bytes
    BaseAddress As Long
    AllocationBase As Long
    AllocationProtect As Long
    RegionSize As Long
    State As Long
    Protect As Long
    lType As Long
End Type

'informations système
Public Type SYSTEM_INFO ' 36 Bytes
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    wProcessorLevel As Integer
    wProcessorRevision As Integer
End Type

'overlapped passé pour SetFilePointer...etc
Public Type OVERLAPPED
    ternal As Long
    ternalHigh As Long
    Offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type

'contient les résultats des recherche de string
Public Type SearchResult
    curOffset As Currency
    strString As String
End Type

'équivalent à un currency
Public Type LARGE_INTEGER
    LowDWORD As Long
    HighDWORD As Long
End Type

'types pour la récupération de la carte des clusters d'un fichier
Public Type STARTING_VCN_INPUT_BUFFER
    StartingVcn As LARGE_INTEGER
End Type
Public Type Extent
    NextVcn As LARGE_INTEGER
    LCN As LARGE_INTEGER
End Type
Public Type RETRIEVAL_POINTERS_BUFFER
    ExtentCount As Long
    Padding As Long
    StartingVcn As LARGE_INTEGER
    Extents(1023) As Extent
End Type
Public Type FileClusters
    File As String
    Moveable As Long
    ExtentsCount As Long
    Extents() As Extent
End Type
Public Type FileClusters2
    File As String
    ExtentsCount As Long
End Type

'temps système
Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

'temps fichier (intervalles de 100 nanosecondes depuis le 1/1/1601)
Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

'pour la recherche de fichiers
Public Type WIN32_FIND_DATA
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

'pour la récupération des répertoires spéciaux
Public Type SHITEMID
    cb As Long
    abID As Byte
End Type
Public Type ITEMIDLIST
    mkid As SHITEMID
End Type

'type concernant la définition de la boite de dialogue Windows à afficher
Public Type SHELLEXECUTEINFO
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

'pour la création d'une IPicture
Public Type GUID
    Data1           As Long
    Data2           As Integer
    Data3           As Integer
    Data4(7)        As Byte
End Type
Public Type PICTDESC
    dwSize          As Long
    dwType          As Long
    hImage          As Long
    xExt            As Long
    yExt            As Long
End Type

'pour la récupération d'icone de fichier (avec SHGetFileInfo)
Public Type SHFILEINFO
    hIcon           As Long
    iIcon           As Long
    dwAttributes    As Long
    szDisplayName   As String * 260
    szTypeName      As String * 80
End Type

'types utilisés pour l'impression
Public Type PRINTER_INFO
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hdc As Long
    FLAGS As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type
Public Type DEVMODE
    dmDeviceName As String * 32
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * 32
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
Public Type DEVNAMES
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
End Type

'contient le type de passe à appliquer à une écriture sur disque
Public Type PASSE_TYPE
    tType As PASSE_ENUM
    sData1 As String
    sData2 As String
End Type

'utilisés pour la création de menus dynamiques
Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type
Public Type POINTAPI
    x As Long
    y As Long
End Type

'utilisé pour subclasser le resize des form
Public Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

'pour la détermination de la version de Windows
Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

'définit la méthode de découpe (splitter de fichiers)
Public Type CUT_METHOD
    tMethode As CUT_METHOD_ENUM
    lParam As Long
End Type

'contient les résultats de la recherche de fichiers
Public Type FILE_SEARCH_RESULT
    sF() As String
End Type

'type de définition des menus
Public Type MENUITEMINFO_STRINGDATA
   cbSize As Long
   fMask As Long
   fType As Long
   fState As Long
   wID As Long
   hSubMenu As Long
   hbmpChecked As Long
   hbmpUnchecked As Long
   dwItemData As Long
   dwTypeData As String
   cch As Long
End Type
