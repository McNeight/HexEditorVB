VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcess"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -----------------------------------------------
'
' Hex Editor VB
' Coded by violent_ken (Alain Descotes)
'
' -----------------------------------------------
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
' -----------------------------------------------


Option Explicit

'-------------------------------------------------------
'//CLASSE DE RECUPERATION D'INFOS SUR LES PROCESSUS
'//Par violent_ken, le 11/11/2006
'-------------------------------------------------------


'-------------------------------------------------------
'CONSTANTES
'-------------------------------------------------------
'constantes utilisées pour l'accès à un snapshot avec l'API CreateToolHelpSnapshot
Private Const TH32CS_SNAPHEAPLIST           As Long = &H1  'heap des processus
Private Const TH32CS_SNAPPROCESS            As Long = &H2   'processus
Private Const TH32CS_SNAPTHREAD             As Long = &H4    'threads
Private Const TH32CS_SNAPMODULE             As Long = &H8    'modules du processus
Private Const TH32CS_SNAPALL                As Long = (TH32CS_SNAPHEAPLIST Or _
                                            TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or _
                                            TH32CS_SNAPMODULE)
Private Const TH32CS_INHERIT                As Long = &H80000000
Private Const SYNCHRONIZE                   As Long = &H100000
'constantes d'accès à un processus (pour OpenProcess)
Private Const PROCESS_SUSPEND_RESUME        As Long = &H800
Private Const PROCESS_QUERY_INFORMATION     As Long = 1024
Private Const PROCESS_VM_READ               As Long = 16
Private Const PROCESS_VM_WRITE              As Long = &H20
Private Const PROCESS_VM_OPERATION          As Long = &H8
Private Const STANDARD_RIGHTS_REQUIRED      As Long = &HF0000 'aussi pour d'autres accès que les processus
Private Const PROCESS_ALL_ACCESS            As Long = (STANDARD_RIGHTS_REQUIRED Or _
                                            SYNCHRONIZE Or &HFFF)
Private Const SPECIFIC_RIGHTS_ALL           As Long = &HFFFF
Private Const PROCESS_READ_WRITE_QUERY      As Long = PROCESS_VM_READ + PROCESS_VM_WRITE + PROCESS_VM_OPERATION + PROCESS_QUERY_INFORMATION
'constantes utilisées pour le changement de privileges
Private Const TOKEN_ASSIGN_PRIMARY          As Long = &H1
Private Const TOKEN_DUPLICATE               As Long = &H2
Private Const TOKEN_IMPERSONATE             As Long = &H4
Private Const TOKEN_QUERY                   As Long = &H8
Private Const TOKEN_QUERY_SOURCE            As Long = &H10
Private Const TOKEN_ADJUST_PRIVILEGES       As Long = &H20
Private Const TOKEN_ADJUST_GROUPS           As Long = &H40
Private Const TOKEN_ADJUST_DEFAULT          As Long = &H80
Private Const TOKEN_ALL_ACCESS              As Long = TOKEN_ASSIGN_PRIMARY + TOKEN_DUPLICATE + _
                                            TOKEN_IMPERSONATE + TOKEN_QUERY + _
                                            TOKEN_QUERY_SOURCE + TOKEN_ADJUST_PRIVILEGES + _
                                            TOKEN_ADJUST_GROUPS + TOKEN_ADJUST_DEFAULT
Private Const SE_DEBUG_NAME                 As String = "SeDebugPrivilege"
Private Const SE_PRIVILEGE_ENABLED          As Long = &H2
'utilisé pour killer un processus
Private Const PROCESS_TERMINAT              As Long = &H1
'constante pour le changement de priority (OpenProcess)
Private Const PROCESS_SET_INFORMATION       As Long = &H200
Private Const MAX_COMPUTERNAME_LENGTH       As Long = 31


'-------------------------------------------------------
'APIs
'-------------------------------------------------------
'changer les droits (privileges) sur les processus
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
'quitter un processus
Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
'fermer un handle
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'obtenir un handle d'un processus à ouvrir
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
'pour obtenir le nom d'un processus (string)
Private Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
'renvoit le premier processus du snap
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
'renvoit les processus suivants
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
'création d'un snapshot process (pour le listing des processus)
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
'créé un snapshot (ici pour les threads)
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32.dll" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
'obtien un handle pour chaque module d'un processus
Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
'obtient le process courant
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
'obtient le répertoire de Windows
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'changement de priorité de processus/thread
Private Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
'obtient des infos sur les heures/temps des processus
Private Declare Function GetProcessTimes Lib "kernel32" (ByVal hProcess As Long, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As Currency, lpUserTime As Currency) As Long
'utilisés pour le listing des modules
'renvoit le premier module du snap
Private Declare Function Module32First Lib "kernel32.dll" (ByVal hSnapshot As Long, ByRef lppe As MODULEENTRY32) As Long
'renvoit les modules suivants
Private Declare Function Module32Next Lib "kernel32.dll" (ByVal hSnapshot As Long, ByRef lpme As MODULEENTRY32) As Long
'suspend un processus
Private Declare Function NtSuspendProcess Lib "Ntdll.dll" (ByVal hProc As Long) As Long
'relache un processus
Private Declare Function NtResumeProcess Lib "Ntdll.dll" (ByVal hProc As Long) As Long
'obtient des infos mémoire sur un processus en particulier
Private Declare Function GetProcessMemoryInfo Lib "PSAPI.DLL" (ByVal hProcess As Long, ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
'obtient le ComputerName
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long


'-------------------------------------------------------
'TYPES ET ENUMS
'-------------------------------------------------------
'type concernant les informations sur les processus
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
    childWnd As Integer
End Type
Private Type MODULEENTRY32
    dwSize As Long                  'taille de cette structure (à initialiser avant l'appel à Module32First ou Module32Next)
    th32ModuleID As Long            'ID du module
    th32ProcessID As Long           'ID du processus qui utilise le module
    GlblcntUsage As Long            'compteur d'usage global pour toutes les instances présentes dans la mémoire du système
    ProccntUsage As Long            'commteur d'usage du module pour le processus
    modBaseAddr As Long             'adresse de début du module en mémoire
    modBaseSize As Long             'taille du module en mémoire
    hModule As Long                 'HMODULE du module
    szModule As String * 256        'nom du module
    szExeFile As String * 260       'path du module
End Type
'type concernant les informations mémoire des processus
Private Type PROCESS_MEMORY_COUNTERS
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
End Type
'types utilisés pour les privileges
Private Type LUID
    LowPart As Long
    HighPart As Long
End Type
Private Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type
Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type
'types contenant des dates
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


'-------------------------------------------------------
'FUNCTIONS AND PROCEDURES
'-------------------------------------------------------


'-------------------------------------------------------
'suspend un processus
'-------------------------------------------------------
Public Function SuspendProcess(ByVal PID As Long) As Long
Dim lProcess As Long

    'obtient le handle
    lProcess = OpenProcess(PROCESS_SUSPEND_RESUME, 0&, PID)
    
    SuspendProcess = NtSuspendProcess(lProcess)
    
    CloseHandle lProcess
End Function

'-------------------------------------------------------
'relache un processus suspendu
'-------------------------------------------------------
Public Function ResumeProcess(ByVal PID As Long) As Long
Dim lProcess As Long

    'obtient le handle
    lProcess = OpenProcess(PROCESS_SUSPEND_RESUME, 0&, PID)
    
    ResumeProcess = NtResumeProcess(lProcess)
    
    CloseHandle lProcess
End Function

'-------------------------------------------------------
'fonction renvoyant le nom complet du fichier en fonction du PID du processus
'-------------------------------------------------------
Public Function GetPathFromPID(ByVal PID As Long) As String
Dim lHprcss As Long
Dim Ret As Long
Dim sResult As String
Dim hModule As Long

    On Error GoTo ErrGest

    If PID = 0 Then
        GetPathFromPID = "[System Process]"
    ElseIf PID = 4 Then
        GetPathFromPID = "System"
    Else
        'création d'un buffer
        sResult = Space$(512)
        
        lHprcss = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, PID)    'recupère un handle du processus
        
        If lHprcss = 0 Then
        'erreur d'accès
        'MsgBox GetLastError = 5
            Exit Function   'echec
        End If
        
        'handle du module du fichier *.exe
        EnumProcessModules lHprcss, hModule, 4&, Ret
        'buffer
        sResult = Space(260) 'max path
        'obtient le path
        GetModuleFileNameExA lHprcss, hModule, sResult, 260
        'ferme le handle
        CloseHandle lHprcss
        'formate la string
        GetPathFromPID = FormatedPath(FormatedString(sResult))  'Left$(sResult, InStr(sResult, vbNullChar) - 1)
    End If
    
ErrGest:
    CloseHandle lHprcss
End Function

'-------------------------------------------------------
'renvoie le nombre de processus
'-------------------------------------------------------
Public Property Get ProcessNumber() As Long
Dim hSnapshot As Long
Dim uProcess As PROCESSENTRY32
Dim R As Long
Dim x As Long

    'création du snapshot (liste)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    If hSnapshot = 0 Then Exit Function
    
    'récupère la taille
    uProcess.dwSize = Len(uProcess)
    
    'premier process à être vu
    R = ProcessFirst(hSnapshot, uProcess)
    
    x = 0
    Do While R
        R = ProcessNext(hSnapshot, uProcess)
        x = x + 1
    Loop
    
    'fermeture du handle
    Call CloseHandle(hSnapshot)
    
    ProcessNumber = x
End Property
    
'-------------------------------------------------------
'renvoie le nom du processus en fonction de son PID
'-------------------------------------------------------
Public Function GetProcessFromPID(ByVal PID As Long) As String
Dim hSnapshot As Long
Dim uProcess As PROCESSENTRY32
Dim R As Long

    On Error Resume Next
    
    'création du snapshot (liste)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    If hSnapshot = 0 Then Exit Function
    
    'récupère la taille
    uProcess.dwSize = Len(uProcess)
    
    'premier process à être vu
    R = ProcessFirst(hSnapshot, uProcess)
    
    Do While R
        If uProcess.th32ProcessID = PID Then
            GetProcessFromPID = FormatedString(uProcess.szExeFile)
            Exit Do
        End If
        'prochain process
        R = ProcessNext(hSnapshot, uProcess)
    Loop
    
    'fermeture du handle
    Call CloseHandle(hSnapshot)
    
End Function

'-------------------------------------------------------
'termine un processus en fonction de son PID
'-------------------------------------------------------
Public Function TerminateProc(ByVal PID As Long, Optional ByVal Confirm As Boolean = False) As Long
Dim lhwndProcess As Long
Dim lRep As Long
Dim lExitCode As Long

    If Confirm Then
        'demande de confirmation
        lRep = MsgBox("Êtes vous sur de vouloir terminer ces processus ?" & vbNewLine & "Même les processus système seront fermés, ce qui peut entrainer des problèmes.", vbInformation + vbYesNo, "Attention")
        If Not (lRep = vbYes) Then Exit Function
    End If
    
    'ouvre le process en question
    lhwndProcess = OpenProcess(PROCESS_TERMINAT, 0, PID)
    
    'kill le processus
    TerminateProc = TerminateProcess(lhwndProcess, lExitCode)
    
    'ferme le handle
    CloseHandle lhwndProcess

End Function

'-------------------------------------------------------
'obtient les infos sur le temps processeur
'-------------------------------------------------------
Public Sub GetProcTimes(ByVal PID As Long, ByRef KernelTime As Double, ByRef UserTime As Double, ByRef ProcessorTime As Double)
'permet de récupérer les temps kernel, user et processor d'un processus en fonction de son PID
'Le temps processeur (visible dans TaskMgr) est le temps noyau + temps utilisateur
Dim T0 As FILETIME  'cette variable n'est pas utilisées, mais est nécessaire comme argument de l'API
Dim curTime As Currency 'Le temps qui nous intéresse
Dim curTime2 As Currency  'l'autre
Dim lProcess As Double  'handle du processus
Dim lTime As Double 'temps que l'on récupère
Dim lTime2 As Double 'temps que l'on récupère

    'obtient le handle vers le process
    lProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, PID)
    
    
    GetProcessTimes lProcess, T0, T0, curTime, curTime2
    
    'Renvoit le temps en secondes
    lTime = curTime / 1000
    lTime2 = curTime2 / 1000
    
    'Ferme le handle ouvert par OpenProcess
    CloseHandle lProcess
    
    
    KernelTime = lTime
    UserTime = lTime2
    'Le temps processeur est la somme des temps kernel et user
    ProcessorTime = UserTime + KernelTime
    
End Sub


'-------------------------------------------------------
'donne le privilege Debug à l'application (permet d'accéder à tout les processus avec OpenProcess)
'-------------------------------------------------------
Private Sub GetAllPrivileges()
Dim lhTokenHandle As Long
Dim tLuid As LUID
Dim tTokenPriv As TOKEN_PRIVILEGES
Dim tTokenPrivNew As TOKEN_PRIVILEGES
Dim lBufferNeeded As Long
Dim hProc As Long
Dim lBufferLen As Long
Dim sComputerName As String

    'créé un buffer pour obtenir le nom du PC
    lBufferLen = MAX_COMPUTERNAME_LENGTH + 1
    sComputerName = String$(lBufferLen, "0")
    'obtient le nom
    GetComputerName sComputerName, lBufferLen
    'formate la string
    sComputerName = Left$(sComputerName, lBufferLen)
    
    'obtient le handle de ce process
    hProc = GetCurrentProcess
    
    'change les droits de cette application
    OpenProcessToken hProc, TOKEN_ALL_ACCESS, lhTokenHandle
    LookupPrivilegeValue sComputerName, SE_DEBUG_NAME, tLuid
    
    'détermine le nombre de privileges à changer
    tTokenPriv.PrivilegeCount = 1
    tTokenPriv.TheLuid = tLuid
    tTokenPriv.Attributes = SE_PRIVILEGE_ENABLED
    
    'autorise le privilege SE_DEBUG_NAME
    AdjustTokenPrivileges lhTokenHandle, False, tTokenPriv, Len(tTokenPrivNew), tTokenPrivNew, lBufferNeeded
    
    'ferme les handles
    CloseHandle lhTokenHandle
    CloseHandle hProc
End Sub

'-------------------------------------------------------
'obtient un processus ==> renseignera sur ses propriétés
'-------------------------------------------------------
Public Function GetProcess(ByVal PID As Long, Optional ByVal EnumerateParent As Boolean = False, Optional ByVal CountModules As Boolean = False, Optional ByVal bGetMemoryInfo As Boolean = False) As ProcessItem
Dim hSnapshot As Long
Dim uProcess As PROCESSENTRY32
Dim R As Long
Dim mdl() As MODULEENTRY32
Dim procTemp As ProcessItem
Dim procMem As PROCESS_MEMORY_COUNTERS
Dim NulProc As ProcessItem

    Set NulProc = New ProcessItem
    Set procTemp = New ProcessItem
    
    'création du snapshot (liste)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    If hSnapshot = 0 Then Exit Function
    
    'récupère la taille
    uProcess.dwSize = Len(uProcess)
    
    'premier process à être vu
    R = ProcessFirst(hSnapshot, uProcess)
    
    Do While R
        If uProcess.th32ProcessID = PID Then
            'c'est le PID souhaité
            
            'on rensiegne dessus
            With procTemp
                .childWnd = uProcess.childWnd
                .cntThreads = uProcess.cntThreads
                .cntUsage = uProcess.cntUsage
                .dwFlags = uProcess.dwFlags
                .dwSize = uProcess.dwSize
                .pcPriClassBase = uProcess.pcPriClassBase
                .th32ModuleID = uProcess.th32ModuleID
                .th32DefaultHeapID = uProcess.th32DefaultHeapID
                .szExeFile = FormatedString(uProcess.szExeFile)
                .th32ParentProcessID = uProcess.th32ParentProcessID
                .th32ProcessID = PID
                .szImagePath = GetPathFromPID(uProcess.th32ProcessID)
                If .th32ProcessID <> 0 And EnumerateParent Then
                    'obtient le processus parent de ce processus (récursivement)
                    Set .procParentProcess = GetProcess(.th32ParentProcessID)
                Else
                    'processus inexistant
                    Set .procParentProcess = NulProc
                End If
                
                If CountModules Then
                    'alors compte le nombre de modules
                    .procModulesCount = NumberOfModules(.th32ProcessID)
                    
                    Dim m() As ProcessModuleItem
                    Me.EnumerateModules PID, m()
                    
                    .procModules = m()
                    
                End If

                If bGetMemoryInfo Then
                    'récupère les infos mémoire
                    
                    Set .procMemory = New ProcessMemoryItem
                    
                    procMem = GetMemoryInfo(.th32ProcessID)
                    .procMemory.cb = procMem.cb
                    .procMemory.PageFaultCount = procMem.PageFaultCount
                    .procMemory.PagefileUsage = procMem.PagefileUsage
                    .procMemory.PeakPagefileUsage = procMem.PeakPagefileUsage
                    .procMemory.PeakWorkingSetSize = procMem.PeakWorkingSetSize
                    .procMemory.QuotaNonPagedPoolUsage = procMem.QuotaNonPagedPoolUsage
                    .procMemory.QuotaPagedPoolUsage = procMem.QuotaPagedPoolUsage
                    .procMemory.QuotaPeakNonPagedPoolUsage = procMem.QuotaPeakNonPagedPoolUsage
                    .procMemory.QuotaPeakPagedPoolUsage = procMem.QuotaPeakPagedPoolUsage
                    .procMemory.WorkingSetSize = procMem.WorkingSetSize
                End If
                    .IsProc = True
            End With
            Exit Do
        End If
        
        'prochain process
        R = ProcessNext(hSnapshot, uProcess)
    Loop
    
    Set GetProcess = procTemp
    
    'fermeture du handle
    Call CloseHandle(hSnapshot)

End Function

'-------------------------------------------------------
'créé une liste des modules d'un processus
'-------------------------------------------------------
Private Sub CreateModuleList(ByVal PID As Long, ByRef mdList() As MODULEENTRY32)
Dim lSnap As Long
Dim x As Long
Dim mdMOD As MODULEENTRY32
Dim mdTemp() As MODULEENTRY32

    x = 0

    'création du snapshot des modules
    lSnap = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, PID)

    mdMOD.dwSize = Len(mdMOD)
    
    's'occupe du premier module rencontré dans le snap
    If Module32First(lSnap, mdMOD) Then
        'alors on a trouvé un module, on va pouvoir continuer plus loin
        
        'on dimensionne notre tableau
        ReDim mdTemp(0)
        'on stocke les infos du premier module dans la liste, à l'emplacement 0
        mdTemp(0) = mdMOD
        mdMOD.dwSize = Len(mdMOD)
        
        'continue tant qu'il y a des nouveaux modules dans le snapshot
        Do While Module32Next(lSnap, mdMOD)

            'on redimensionne le tableau de 1 plus grand, pour pouvoir acceullir la liste temporaire du module en cours
            ReDim Preserve mdTemp(x)
            
            'formatage des strings directement dans cette Sub (car ces infos ne servent qu'à l'affichage direct)
            mdMOD.szExeFile = FormatedString(mdMOD.szExeFile)
            mdMOD.szModule = FormatedString(mdMOD.szModule)
            'les autres infos sont des Long, donc pas de formatage
            
            'stocke les infos du module en cours à l'emplacement x
            mdTemp(x) = mdMOD
            'prépare la taille pour le prochain module
            mdMOD.dwSize = Len(mdMOD)
            x = x + 1
        Loop
    Else
        ReDim mdTemp(1)
    End If
    
    ReDim Preserve mdTemp(UBound(mdTemp()) - 1)
    'stockage de la liste des modules
    mdList = mdTemp 'on aurait directement pu travailler sur mdList sans passer par mdTemp, mais il est plus propre de ne changer mdList (d'ailleurs toute autre valeur renvoyée par une sub/fonction) uniquement A LA FIN de cette sub/fonction
    
    'on le handle du snap
    CloseHandle lSnap
End Sub

'-------------------------------------------------------
'renvoie la liste des modules d'un processus
'-------------------------------------------------------
Public Function EnumerateModules(ByVal PID As Long, ByRef m() As ProcessModuleItem) As Long
Dim mdl() As MODULEENTRY32
Dim x As Long

    On Error Resume Next

    CreateModuleList PID, mdl() 'créé la liste
    
    ReDim m(UBound(mdl()))
    
    For x = 0 To UBound(mdl())
        Set m(x) = New ProcessModuleItem
        With m(x)
            .dwSize = mdl(x).dwSize
            .GlblcntUsage = mdl(x).GlblcntUsage
            .hModule = mdl(x).hModule
            .modBaseAddr = mdl(x).modBaseAddr
            .modBaseSize = mdl(x).modBaseSize
            .ProccntUsage = mdl(x).ProccntUsage
            .szExeFile = FormatedString(mdl(x).szExeFile) 'Left$(mdl(x).szExeFile, InStr(mdl(x).szExeFile, Chr(32)) - 1)
            .szModule = FormatedString(mdl(x).szModule)
            .th32ModuleID = mdl(x).th32ModuleID
            .th32ProcessID = mdl(x).th32ProcessID
        End With
    Next x
    
    EnumerateModules = UBound(m())
    
End Function

'-------------------------------------------------------
'énumère la liste des processus et la stocke dans p
'-------------------------------------------------------
Public Function EnumerateProcesses(ByRef p() As ProcessItem, Optional ByVal EnumerateParent As Boolean = False, Optional ByVal CountModules As Boolean = False, Optional ByVal bGetMemoryInfo As Boolean = False) As Long
Dim hSnapshot As Long
Dim uProcess As PROCESSENTRY32
Dim R As Long
Dim x As Long
Dim NOPROCESS As ProcessItem
Dim procMem As PROCESS_MEMORY_COUNTERS
    
    
    Set NOPROCESS = New ProcessItem 'c'est le process nul
    
    x = 0
        
    'redimensionne le tableau de la liste des processus
    ReDim p(0)
    
    'création du snapshot (liste)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    If hSnapshot = 0 Then
        ReDim p(0)
        Exit Function
    End If
    
    'récupère la taille
    uProcess.dwSize = Len(uProcess)
    
    'premier process à être vu
    R = ProcessFirst(hSnapshot, uProcess)
    
    Do While R
        Set p(x) = New ProcessItem
        With p(x)
            .th32ProcessID = uProcess.th32ProcessID
            .szExeFile = FormatedString(uProcess.szExeFile)
            .th32ParentProcessID = uProcess.th32ParentProcessID
            .childWnd = uProcess.childWnd
            .cntThreads = uProcess.cntThreads
            .cntUsage = uProcess.cntUsage
            .dwFlags = uProcess.dwFlags
            .dwSize = uProcess.dwSize
            .pcPriClassBase = uProcess.pcPriClassBase
            .th32ModuleID = uProcess.th32ModuleID
            .th32DefaultHeapID = uProcess.th32DefaultHeapID
            .szImagePath = GetPathFromPID(uProcess.th32ProcessID)
            If .th32ProcessID <> 0 And EnumerateParent Then
                'obtient le processus parent de ce processus
                Set .procParentProcess = GetProcess(.th32ParentProcessID, EnumerateParent, CountModules)
            Else
                Set .procParentProcess = NOPROCESS
            End If
            
            If CountModules Then
                'alors compte le nombre de modules
                p(x).procModulesCount = NumberOfModules(.th32ProcessID)
            End If
        
            If bGetMemoryInfo Then
                'récupère les infos mémoire
                                    
                Set .procMemory = New ProcessMemoryItem
                    
                procMem = GetMemoryInfo(.th32ProcessID)
                .procMemory.cb = procMem.cb
                .procMemory.PageFaultCount = procMem.PageFaultCount
                .procMemory.PagefileUsage = procMem.PagefileUsage
                .procMemory.PeakPagefileUsage = procMem.PeakPagefileUsage
                .procMemory.PeakWorkingSetSize = procMem.PeakWorkingSetSize
                .procMemory.QuotaNonPagedPoolUsage = procMem.QuotaNonPagedPoolUsage
                .procMemory.QuotaPagedPoolUsage = procMem.QuotaPagedPoolUsage
                .procMemory.QuotaPeakNonPagedPoolUsage = procMem.QuotaPeakNonPagedPoolUsage
                .procMemory.QuotaPeakPagedPoolUsage = procMem.QuotaPeakPagedPoolUsage
                .procMemory.WorkingSetSize = procMem.WorkingSetSize
            End If
                .IsProc = True
        End With
        
        'prochain process
        R = ProcessNext(hSnapshot, uProcess)
        
        ReDim Preserve p(UBound(p()) + 1)
        x = x + 1
    Loop
    
    'fermeture du handle
    Call CloseHandle(hSnapshot)
    
    EnumerateProcesses = x

End Function

'-------------------------------------------------------
'compte le nombre de modules d'un processus
'-------------------------------------------------------
Private Function NumberOfModules(ByVal PID As Long) As Long
Dim lSnap As Long
Dim x As Long
Dim mdMOD As MODULEENTRY32
Dim mdTemp() As MODULEENTRY32

    'création du snapshot des modules
    lSnap = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, PID)
    
    x = 0
    mdMOD.dwSize = Len(mdMOD)
    
    's'occupe du premier module rencontré dans le snap
    If Module32First(lSnap, mdMOD) Then
        'alors on a trouvé un module, on va pouvoir continuer plus loin

        mdMOD.dwSize = Len(mdMOD)
        
        'continue tant qu'il y a des nouveaux modules dans le snapshot
        Do While Module32Next(lSnap, mdMOD)
        
            'prépare la taille pour le prochain module
            mdMOD.dwSize = Len(mdMOD)
            x = x + 1
        Loop
    End If
    
    NumberOfModules = x - 1

    'on le handle du snap
    CloseHandle lSnap
End Function

'-------------------------------------------------------
'change la priorité d'un processus
'-------------------------------------------------------
Public Sub ChangePriority(ByVal PID As Long, ByVal NewPriority As PRIORITY_TYPE)
'changement de priorité d'un processus en fonction de son PID
Dim hProcess As Long

    'obtient le thread et processus du processus désiré
    hProcess = OpenProcess(PROCESS_SET_INFORMATION, 0, PID)
    
    'changement de priorité
    SetPriorityClass hProcess, NewPriority
    
    'ferme le handle
    CloseHandle hProcess
End Sub

'-------------------------------------------------------
'renvoie vrai ou faux si le pid existe ou non
'-------------------------------------------------------
Public Function DoesPIDExist(PID As Long) As Boolean
Dim hSnapshot As Long
Dim uProcess As PROCESSENTRY32
Dim R As Long

    DoesPIDExist = False
    
    'création du snapshot (liste)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    If hSnapshot = 0 Then Exit Function
    
    'récupère la taille
    uProcess.dwSize = Len(uProcess)
    
    'premier process à être vu
    R = ProcessFirst(hSnapshot, uProcess)
    
    Do While R
        
        If uProcess.th32ProcessID = PID Then
            DoesPIDExist = True
            Exit Do
        End If
        
        'prochain process
        R = ProcessNext(hSnapshot, uProcess)
    Loop
    
    'fermeture du handle
    Call CloseHandle(hSnapshot)

End Function

'-------------------------------------------------------
'obtient le répertoire de windows
'-------------------------------------------------------
Private Function WindowsDirectory() As String
Dim sPath As String

    'buffer
    sPath = Space(255)
    
    'get the directory
    GetWindowsDirectory sPath, Len(sPath)
    
    'formate la string
    sPath = Left$(sPath, InStr(sPath, vbNullChar) - 1)
    
    WindowsDirectory = sPath
End Function

'-------------------------------------------------------
'convertit le chemin sPath en chemin existant (correct)
'-------------------------------------------------------
Private Function FormatedPath(ByVal sPath As String) As String
Dim x As Long
Dim s As String

    If Len(sPath) < 1 Then Exit Function
    
    'modifie le path si commence par SystemRoot
    'len("SystemRoot")=10
    If Left$(sPath, 10) = "SystemRoot" Then
        'obtient le répertoire de windows
        sPath = WindowsDirectory & "\" & Right$(sPath, Len(sPath) - 10)
    End If
    'len("\SystemRoot")=11
    If Left$(sPath, 11) = "\SystemRoot" Then
        'obtient le répertoire de windows
        sPath = WindowsDirectory & "\" & Right$(sPath, Len(sPath) - 11)
    End If
    
    s = sPath
    While ((Asc(UCase(Left$(s, 1))) < 65 Or Asc(UCase(Left$(s, 1))) > 90) And Len(s) > 3)
        'alors ce n'est pas une lettre valide ==> on enlève cette lettre
        s = Right$(s, Len(s) - 1)
        DoEvents
    Wend
    
    'enlève deux antislash successifs et les remplace par un seul
    x = InStr(1, s, "\\")
    If x > 0 Then
        s = Left$(s, x - 1) & "\" & Right$(s, Len(s) - Len(Left$(s, x - 1)) - 2)
    End If
    
    'sPath = \??\c:\...
   ' If Left$(sPath, 4) = "\??\" Then
   '     'obtient le répertoire de windows
   '     sPath = Right$(sPath, Len(sPath) - 4)
   ' End If
    
    
    FormatedPath = s
End Function

'-------------------------------------------------------
'initialisation de la classe
'-------------------------------------------------------
Private Sub Class_Initialize()
    GetAllPrivileges    'obtient les privileges nécessaires
End Sub

'-------------------------------------------------------
'obtient les infos sur la mémoire d'un processus
'-------------------------------------------------------
Private Function GetMemoryInfo(ByVal PID As Long) As PROCESS_MEMORY_COUNTERS
Dim procMem As PROCESS_MEMORY_COUNTERS
Dim hProcess As Long

    'définit la taille du type
    procMem.cb = LenB(procMem)

    'obtient le handle du processus
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, PID)
    
    'obtient les infos
    GetProcessMemoryInfo hProcess, procMem, procMem.cb
    
    'stocke les infos
    With GetMemoryInfo
        .cb = procMem.cb
        .PageFaultCount = procMem.PageFaultCount
        .PagefileUsage = procMem.PagefileUsage
        .PeakPagefileUsage = procMem.PeakPagefileUsage
        .PeakWorkingSetSize = procMem.PeakWorkingSetSize
        .QuotaNonPagedPoolUsage = procMem.QuotaNonPagedPoolUsage
        .QuotaPagedPoolUsage = procMem.QuotaPagedPoolUsage
        .QuotaPeakNonPagedPoolUsage = procMem.QuotaPeakNonPagedPoolUsage
        .QuotaPeakPagedPoolUsage = procMem.QuotaPeakPagedPoolUsage
        .WorkingSetSize = procMem.WorkingSetSize
    End With
    
    'ferme le handle
    CloseHandle hProcess
End Function

'-------------------------------------------------------
'formatage de string
'-------------------------------------------------------
Private Function FormatedString(ByVal sString As String) As String
Dim s As String

    s = sString
    
    'enlève le vbnullchar de fin si nécessaire
    If InStr(s, vbNullChar) Then s = Left$(s, InStr(s, vbNullChar) - 1)
    
    'enlève les espaces inutiles
    s = Trim$(s)
    
    FormatedString = s
End Function


