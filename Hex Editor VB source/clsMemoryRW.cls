VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMemoryRW"
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
'//CLASSE DE GESTION DES PROCESSUS ET DE LEUR MEMOIRE
'//by Violent_ken
'=======================================================

'=======================================================
'CONSTANTES
'=======================================================
'constantes d'accès à un processus (pour OpenProcess)
Private Const SYNCHRONIZE                   As Long = &H100000
Private Const STANDARD_RIGHTS_REQUIRED      As Long = &HF0000 'aussi pour d'autres accès que les processus
Private Const PROCESS_ALL_ACCESS            As Long = (STANDARD_RIGHTS_REQUIRED Or _
                                            SYNCHRONIZE Or &HFFF)
Private Const PROCESS_VM_READ               As Long = 16
Private Const PROCESS_VM_WRITE              As Long = &H20
Private Const PROCESS_VM_OPERATION          As Long = &H8
Private Const PROCESS_QUERY_INFORMATION     As Long = 1024
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
Private Const TOKEN_ALL_ACCESS              As Long = TOKEN_ASSIGN_PRIMARY + _
                                            TOKEN_DUPLICATE + TOKEN_IMPERSONATE + _
                                            TOKEN_QUERY + TOKEN_QUERY_SOURCE + _
                                            TOKEN_ADJUST_PRIVILEGES + TOKEN_ADJUST_GROUPS + _
                                            TOKEN_ADJUST_DEFAULT
Private Const SE_DEBUG_NAME                 As String = "SeDebugPrivilege"
Private Const SE_PRIVILEGE_ENABLED          As Long = &H2

'constantes utilisées pour déterminer le type de zone mémoire d'un processus
Private Const MEM_PRIVATE                   As Long = &H20000
Private Const MEM_COMMIT                    As Long = &H1000

Private Const INVALID_HANDLE_VALUE          As Long = -1


'=======================================================
'APIs
'=======================================================
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function VirtualQueryEx& Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long)



'=======================================================
'TYPES ET ENUMS
'=======================================================
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
Private Type MEMORY_BASIC_INFORMATION ' 28 bytes
    BaseAddress As Long
    AllocationBase As Long
    AllocationProtect As Long
    RegionSize As Long
    State As Long
    Protect As Long
    lType As Long
End Type
Private Type SYSTEM_INFO ' 36 Bytes
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
Private Type T_RESULT   'privé pour éviter d'avoir à déclarer un type de donnée public dans une classe pour des modules, ou dans un module pour cette classe
    curOffset As Currency
    strString As String
End Type


'=======================================================
'FUNCTIONS AND PROCEDURES
'=======================================================

'=======================================================
'donne le privilege Debug à l'application (permet d'accéder à tout les processus avec OpenProcess)
'=======================================================
Private Sub GetAllPrivileges()
Dim lhTokenHandle As Long
Dim tLuid As LUID
Dim tTokenPriv As TOKEN_PRIVILEGES
Dim tTokenPrivNew As TOKEN_PRIVILEGES
Dim lBufferNeeded As Long
Dim hProc As Long

    'obtient le handle de ce process
    hProc = GetCurrentProcess
    
    'change les droits de cette application
    Call OpenProcessToken(hProc, TOKEN_ALL_ACCESS, lhTokenHandle)
    Call LookupPrivilegeValue(vbNullString, SE_DEBUG_NAME, tLuid)
    
    'détermine le nombre de privileges à changer
    With tTokenPriv
        .PrivilegeCount = 1
        .TheLuid = tLuid
        .Attributes = SE_PRIVILEGE_ENABLED
    End With
    
    'autorise le privilege SE_DEBUG_NAME
    Call AdjustTokenPrivileges(lhTokenHandle, False, tTokenPriv, _
        Len(tTokenPrivNew), tTokenPrivNew, lBufferNeeded)
    
    'ferme les handles
    Call CloseHandle(lhTokenHandle)
End Sub

'=======================================================
'renvoie un handle de processus permettant la lecture écriture en mémoire
'=======================================================
Private Function GetValidHandle(ByVal PID As Long) As Long
    GetValidHandle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
End Function

'=======================================================
'lit lSize bytes dans la zone virtuelle de la mémoire d'un processus
'demande un PID
'=======================================================
Public Function ReadBytes(ByVal PID As Long, ByVal lngOffset As Long, _
    ByVal lngSize As Long) As String
    
Dim sBuf As String
Dim lByte As Long
Dim lHandle As Long
    
    '/!\ Un long suffit pour pouvoir parcourir l'ensemble des offsets possibles
    'car la taille max de la zone virtuelle est 2Go
    
    'créé un buffer
    sBuf = String$(lngSize, 0)
    
    'obtient le handle du processus
    lHandle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
    
    'lit les bytes et stocke dans le buffer
    Call ReadProcessMemory(lHandle, lngOffset, sBuf, lngSize, lByte)
    
    'referme le handle
    Call CloseHandle(lHandle)
    
    ReadBytes = sBuf
End Function

'=======================================================
'lit lSize bytes dans la zone virtuelle de la mémoire d'un processus
'demande un handle
'=======================================================
Public Function ReadBytesH(ByVal lHandle As Long, ByVal lngOffset As Long, _
    ByVal lngSize As Long) As String
    
Dim sBuf As String
Dim lByte As Long
Dim lRet As Long
    
    '/!\ Un long suffit pour pouvoir parcourir l'ensemble des offsets possibles
    'car la taille max de la zone virtuelle est 2Go
    
    'créé un buffer
    sBuf = String$(lngSize, 0)
    
    'lit les bytes et stocke dans le buffer
    lRet = ReadProcessMemory(lHandle, lngOffset, sBuf, lngSize, lByte)
    
    ReadBytesH = sBuf
End Function

'=======================================================
'écrit une string dans la mémoire virtuelle d'un processus
'=======================================================
Public Function WriteBytes(ByVal PID As Long, ByVal lngOffset As Long, _
    ByVal strStringToWrite As String) As Long
    
Dim lHandle As Long

    'obtient le handle du processus
    lHandle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)

    WriteBytes = WriteProcessMemory(lHandle, lngOffset, ByVal strStringToWrite, _
        Len(strStringToWrite), 0&)
    
    'referme le handle
    Call CloseHandle(lHandle)
End Function

'=======================================================
'écrit une string dans la mémoire virtuelle d'un processus (à partir d'un handle)
'=======================================================
Public Function WriteBytesH(ByVal lngHandle As Long, ByVal lngOffset As Long, _
    ByVal strStringToWrite As String) As Long
    
    WriteBytesH = WriteProcessMemory(lngHandle, lngOffset, ByVal _
        strStringToWrite, Len(strStringToWrite), 0&)
End Function

'=======================================================
'obtient les différentes zones mémoire d'un processus utilisées dans la zone virtuelle
'stocke de 1 à Ubound
'=======================================================
Public Sub RetrieveMemRegions(ByVal PID As Long, ByRef lBaseAdress() As Long, _
    ByRef lRegionSize() As Long)
    
Dim lHandle As Long
Dim lPosMem As Long
Dim lRet As Long
Dim lLenMBI As Long
Dim mbi As MEMORY_BASIC_INFORMATION
Dim si As SYSTEM_INFO

    'initialise les tableaux
    ReDim lBaseAdress(0)
    ReDim lRegionSize(0)
                    
    'obtient le handle du processus
    lHandle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)

    lLenMBI = Len(mbi)  'taille de la structure
    Call GetSystemInfo(si)     'obtient les infos sur les adresses de début et de fin de la plage mémoire maximum
    
    lPosMem = si.lpMinimumApplicationAddress  'adresse la plus petite ==> part de là
    
    Do While lPosMem < si.lpMaximumApplicationAddress 'tant que l'adresse est inférieure à l'adresse maximale
        mbi.RegionSize = 0
        
        'obtient les infos sur les régions mémoire du processus définit par son handle hProcess
        lRet = VirtualQueryEx(lHandle, ByVal lPosMem, mbi, lLenMBI)
        
        If lRet = lLenMBI Then
        
            If (mbi.lType = MEM_PRIVATE) And (mbi.State = MEM_COMMIT) Then
                'alors utilisé par le processus
                
                If mbi.RegionSize > 0 Then
                    'région non nulle, alors on la stocke dans les tableaux résultats
                    'les redimensionne
                    ReDim Preserve lBaseAdress(UBound(lBaseAdress) + 1)
                    ReDim Preserve lRegionSize(UBound(lRegionSize) + 1)
                    'stocke à la fin
                    lRegionSize(UBound(lRegionSize)) = mbi.RegionSize
                    lBaseAdress(UBound(lBaseAdress)) = mbi.BaseAddress
                End If
                
            End If
            'continue la recherche des régions (ajoute la taille de la région à l'adresse de départ ==> donne la prochaine adresse de départ)
            
            On Error GoTo ErrCapacityGestion  'dépassement de capacité pour la dernière adresse+regiosize
            
            lPosMem = mbi.BaseAddress + mbi.RegionSize    'fait l'ajout
        
        Else
            'recherche terminée
            Exit Do
        End If
    Loop
    
ErrCapacityGestion:
   Call CloseHandle(lHandle)  'ferme le handle du processus
End Sub

'=======================================================
'fonction de recherche de string dans un processus
'de 1 à Ubound
'=======================================================
Public Sub SearchForStringMemory(ByVal PID As Long, ByVal sMatch As String, _
    ByVal bCasse As Boolean, ByRef tRes() As Long, Optional ByVal PGB As pgrBar)
'Utilisation de l'API CreateFile et ReadFileEx pour une lecture rapide
Dim x As Long
Dim bytAsc As Byte
Dim strBufT As String
Dim i As Long
Dim lHandle As Long
Dim LB() As Long
Dim LS() As Long

    On Error GoTo ErrGestion

    'initialise le tableau
    ReDim tRes(0)

    'on obtient les différentes régions de la mémoire du processus
    Call RetrieveMemRegions(PID, LB(), LS())
    
    If Not (PGB Is Nothing) Then
        'on initialise la progressabr
        PGB.Min = 0
        PGB.Value = 0
        PGB.Max = UBound(LS())
    End If
    
    'on obtient le handle depuis le PID
    lHandle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)

    If bCasse = False Then sMatch = LCase$(sMatch)  'ne cherche que les minuscules

    For x = 1 To UBound(LS())   'pour chaque zone mémoire
    
        'obtient la string de la plage visualisée
        strBufT = ReadBytesH(lHandle, LB(x), LS(x))
        
        If bCasse = False Then strBufT = LCase$(strBufT)    'cherche que des minuscules (pas de casse respectée)
        
        'tant que la string contient le match
        While InStr(1, strBufT, sMatch, vbBinaryCompare) <> 0
            'trouvé une string ==> l'ajoute
            ReDim Preserve tRes(UBound(tRes) + 1)
                            
            tRes(UBound(tRes)) = LB(x) + InStr(1, strBufT, sMatch, vbBinaryCompare) + LS(x) - Len(strBufT) - 1
            
            'raccourci le buffer
            strBufT = Right$(strBufT, Len(strBufT) - InStr(1, strBufT, sMatch, vbBinaryCompare) - Len(sMatch) + 1)
        Wend
        
        If Not (PGB Is Nothing) Then PGB.Value = x    'refresh progressbar
        DoEvents    'rend la main
        
    Next x
    
    If Not (PGB Is Nothing) Then PGB.Value = PGB.Max
        
    Let strBufT = vbNullString

ErrGestion:
    'referme le handle
    Call CloseHandle(lHandle)
End Sub

'=======================================================
'détermine si un byte est considéré comme convenable en fonction
'des paramètres Afficher : min, MAJ, nbres, signes
'function utilisée directement avec les procédures de SearchStringIn...
'=======================================================
Private Function IsCharConsideredInAString(ByVal bytChar As Byte, _
    ByVal bSigns As Boolean, ByVal bMaj As Boolean, ByVal bMin As Boolean, _
    ByVal bNumbers As Boolean, ByVal bAccent As Boolean) As Boolean
    
    If bMaj Then
        IsCharConsideredInAString = (bytChar >= 65 And bytChar <= 90)
        If IsCharConsideredInAString Then Exit Function
    End If
    If bMin Then
        IsCharConsideredInAString = (bytChar >= 97 And bytChar <= 122)
        If IsCharConsideredInAString Then Exit Function
    End If
    If bNumbers Then
        IsCharConsideredInAString = (bytChar >= 48 And bytChar <= 57)
        If IsCharConsideredInAString Then Exit Function
    End If
    If bSigns Then
        IsCharConsideredInAString = (bytChar >= 33 And bytChar <= 47) Or _
        (bytChar >= 58 And bytChar <= 64) Or (bytChar >= 91 And bytChar <= 96) Or _
        (bytChar >= 123 And bytChar <= 126)
        If IsCharConsideredInAString Then Exit Function
    End If
    If bytChar = 32 Or bytChar = 39 Then    'espace ou apostrophe
        IsCharConsideredInAString = True
        If IsCharConsideredInAString Then Exit Function
    End If
    If bAccent Then
        IsCharConsideredInAString = (bytChar >= 192)
        If IsCharConsideredInAString Then Exit Function
    End If
    
    IsCharConsideredInAString = False
End Function

'=======================================================
'fonction de recherche de string complètes dans un processus
'ne cherche pas une string en particulier, mais toutes les strings qui ressemblent à des mots
'stocke dans un tableau de 1 à Ubound
'=======================================================
Public Sub SearchEntireStringMemory(ByVal PID As Long, ByVal lMinimalLength _
    As Long, ByVal bSigns As Boolean, ByVal bMaj As Boolean, ByVal bMin As _
    Boolean, ByVal bNumbers As Boolean, ByVal bAccent As Boolean, _
    ByRef lngRes() As Long, ByRef strRes() As String, Optional PGB As pgrBar)
    
Dim strCtemp As String
Dim x As Long
Dim lngLen As Long
Dim bytAsc As Byte
Dim strBuffer As String
Dim curByte As Currency
Dim i As Long
Dim tRes() As T_RESULT
Dim lHandle As Long
Dim LB() As Long
Dim LS() As Long

    On Error GoTo ErrGestion
    
    'on obtient les différentes régions de la mémoire du processus
    Call RetrieveMemRegions(PID, LB(), LS())
    
    'calcule la taille max (lngLen)
    lngLen = 0
    For i = 1 To UBound(LS())
        lngLen = lngLen + LS(i)
    Next i
    
    If Not (PGB Is Nothing) Then
        'on initialise la progressabr
        PGB.Min = 0
        PGB.Value = 0
        PGB.Max = UBound(LS())
    End If
    
    'on obtient le handle depuis le PID
    lHandle = GetValidHandle(PID)

    'initialise le tableau
    ReDim tRes(0)

    'vérifie que le handle est valide
    If lHandle = INVALID_HANDLE_VALUE Then
        ReDim lngRes(0)
        ReDim strRes(0)
        Exit Sub
    End If
    
    strCtemp = vbNullString: x = 1: curByte = 0
    
    'pour chaque zone mémoire
    For x = 1 To UBound(LS())
        
        'on récupère la string complète de cette zone mémoire
        strBuffer = Me.ReadBytesH(lHandle, LB(x), LS(x))

        strCtemp = vbNullString
        
        'effectue la recherche dans la string
        For i = 0 To LS(x) - 1
            If (i Mod 300000) = 0 Then DoEvents   'rend un peu la main
            bytAsc = Asc(Mid$(strBuffer, i + 1, 1)) 'prend un byte
            
            If IsCharConsideredInAString(bytAsc, bSigns, bMaj, bMin, bNumbers, bAccent) Then
                'caractère x est valide
                strCtemp = strCtemp & Chr_(bytAsc)
            Else
                strCtemp = Trim$(strCtemp)
                If Len(strCtemp) > lMinimalLength Then
                    'trouvé la chaine correspondante
                    ReDim Preserve tRes(UBound(tRes) + 1)
                    tRes(UBound(tRes)).curOffset = i + LB(x) - Len(strCtemp) + 1
                    tRes(UBound(tRes)).strString = strCtemp
                End If
                strCtemp = vbNullString
            End If
        Next i
            
        If Len(strCtemp) > lMinimalLength Then
            'trouvé la dernière chaine possible (dernier byte compris dans cette chaine)
            ReDim Preserve tRes(UBound(tRes) + 1)
            tRes(UBound(tRes)).curOffset = LS(x) + LB(x) - Len(strCtemp) + 1
            tRes(UBound(tRes)).strString = strCtemp
        End If
        
        If Not (PGB Is Nothing) Then PGB.Value = x    'refresh progressbar
        DoEvents    'rend la main
     Next x
    
    
    If Not (PGB Is Nothing) Then PGB.Value = PGB.Max
    
    Let strBuffer = vbNullString
    Call CloseHandle(lHandle)  'ferme le handle du process
    
    'maintenant, stocke dans les arrays de sortie
    ReDim lngRes(UBound(tRes()))
    ReDim strRes(UBound(tRes()))
    For i = 1 To UBound(tRes())
        If (i Mod 2000) = 0 Then DoEvents 'rend la main
        lngRes(i) = tRes(i).curOffset
        strRes(i) = tRes(i).strString
    Next i
    
    Exit Sub
    
ErrGestion:

    Call CloseHandle(lHandle)  'ferme le handle du process
    clsERREUR.AddError "mdlDisk.SearchStringInFile", True
End Sub

'=======================================================
'procédure d'initialisation de la classe
'=======================================================
Private Sub Class_Initialize()
    Call GetAllPrivileges    'obtient le privileges debug
End Sub
