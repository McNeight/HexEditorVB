Attribute VB_Name = "mdlDisk"
' =======================================================
'
' Hex Editor VB
' Coded by violent_ken (Alain Descotes)
'
' =======================================================
'
' A complete hexadecimal editor for Windows �
' (Editeur hexad�cimal complet pour Windows �)
'
' Copyright � 2006-2007 by Alain Descotes.
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
'//MODULE CONTENANT LES FONCTIONS POUR L'OUVERTURE D'UN DISQUE
'=======================================================


'=======================================================
'renvoie un drive compatible avec l'api CreateFile
'=======================================================
Public Function BuildDrive(ByVal sDrive As String) As String
    BuildDrive = "\\.\" & UCase$(Left$(sDrive, 2))
End Function

'=======================================================
'lecture de lLen bytes � l'offset lOffset dans le drive sDrive
'=======================================================
'Public Sub ReadB(ByVal sDrive As String, ByVal lOffset As Currency, ByVal lLen As Long, ByRef lResult() As Byte)
'Dim lDrive As Long
'Dim crPointeur As Currency
'Dim tOver As OVERLAPPED
'Dim Ret As Long, ret2 As Long

'    On Error GoTo DiskErr

    'obtient un path valide pour l'API CreateFIle si n�cessaire
'    If Len(sDrive) <> 6 Then sDrive = BuildDrive(sDrive)
    
    'initialise le tableau r�sultat
'    ReDim lResult(0)
        
    'obtient un handle vers le Drive
'    lDrive = CreateFile(sDrive, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    
    'si le handle est correct
'    If lDrive <> INVALID_HANDLE_VALUE Then
        
        'move pointer
'        ret2 = SetFilePointerEx(lDrive, lOffset, 0&, FILE_BEGIN) ', , lOffset   'positionne au Offset dans le disque
        
        'redimensionne le tableau � la taille convenable du r�sultat
'        ReDim lResult(lLen)
        
        'obtient les bytes d�sir�s
'        Ret = ReadFile(lDrive, lResult(1), lLen, 0&, ByVal 0&)
'        Debug.Print "setfilepointerex=" & ret2 & "  readfile=" & Ret
'    End If
    
'DiskErr:
    
    'ferme le handle ouvert
'    CloseHandle lDrive
'End Sub

'=======================================================
'permet de lire des bytes directement dans le disque
'=======================================================
Public Sub DirectRead(ByVal sDrive As String, ByVal iStartSec As Currency, ByVal nBytes As Long, ByVal lBytesPerSector As Long, ByRef ReadOctet() As Byte)
' Attention le nombre d'octets lus ou �crits ainsi que l'offset du premier octet lu ou �crit
' doivent imp�rativement �tre un multiple de la taille d'un secteur de disque
' Istartsec et nbytes doivent �tre des multiples de 512 ( taille standard des secteurs des disques)
Dim BytesRead As Long
Dim Pointeur As Currency
Dim Ret As Long
Dim hDevice As Long
Dim lLowPart As Long, lHighPart As Long

    On Error GoTo dskerror
    
    'obtient un path valide pour l'API CreateFIle si n�cessaire
    If Len(sDrive) <> 6 Then sDrive = BuildDrive(sDrive)

    'ouvre le drive
    hDevice = CreateFile(sDrive, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
   
    'quitte si le handle n'est pas valide
    If hDevice = INVALID_HANDLE_VALUE Then Exit Sub
   
    'd�termine le byte de d�part du secteur
    Pointeur = CCur(iStartSec) * CCur(lBytesPerSector)
    
    'transforme un currency en 2 long pour une structure LARGE_INTEGER
    GetLargeInteger Pointeur, lLowPart, lHighPart

    'd�place, dans le fichier (ici un disque) point� par hDevice, le "curseur" au premier
    'byte que l'on veut lire (donn� par deux long)
    Ret = SetFilePointer(hDevice, lLowPart, lHighPart, FILE_BEGIN)  'FILE_BEGIN ==> part du d�but du fichier pour d�compter la DistanceToMove
    If Ret = -1 Then GoTo dskerror
           
    'redimensionne le tableaux r�sultant
    ReDim ReadOctet(0 To nBytes - 1) 'contient les nBytes lus, de 0 � Ubound-1
    
    'appelle l'API de lecture
    Ret = ReadFile(hDevice, ReadOctet(0), nBytes, BytesRead, 0&)
    
dskerror:

    'ferme le handle
    CloseHandle hDevice
End Sub

'=======================================================
'permet de lire des bytes directement dans le disque PHYSIQUE
'=======================================================
Public Sub DirectReadPhys(ByVal bytDrive As Byte, ByVal iStartSec As Currency, ByVal nBytes As Long, ByVal lBytesPerSector As Long, ByRef ReadOctet() As Byte)
' Attention le nombre d'octets lus ou �crits ainsi que l'offset du premier octet lu ou �crit
' doivent imp�rativement �tre un multiple de la taille d'un secteur de disque
' Istartsec et nbytes doivent �tre des multiples de 512 ( taille standard des secteurs des disques)
Dim BytesRead As Long
Dim Pointeur As Currency
Dim Ret As Long
Dim hDevice As Long
Dim lLowPart As Long, lHighPart As Long

    On Error GoTo dskerror

    'ouvre le drive
    hDevice = CreateFile("\\.\PHYSICALDRIVE" & CStr(bytDrive), GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
   
    'quitte si le handle n'est pas valide
    If hDevice = INVALID_HANDLE_VALUE Then Exit Sub
   
    'd�termine le byte de d�part du secteur
    Pointeur = CCur(iStartSec) * CCur(lBytesPerSector)
    
    'transforme un currency en 2 long pour une structure LARGE_INTEGER
    GetLargeInteger Pointeur, lLowPart, lHighPart

    'd�place, dans le fichier (ici un disque) point� par hDevice, le "curseur" au premier
    'byte que l'on veut lire (donn� par deux long)
    Ret = SetFilePointer(hDevice, lLowPart, lHighPart, FILE_BEGIN)  'FILE_BEGIN ==> part du d�but du fichier pour d�compter la DistanceToMove
    If Ret = -1 Then GoTo dskerror
           
    'redimensionne le tableaux r�sultant
    ReDim ReadOctet(0 To nBytes - 1) 'contient les nBytes lus, de 0 � Ubound-1
    
    'appelle l'API de lecture
    Ret = ReadFile(hDevice, ReadOctet(0), nBytes, BytesRead, 0&)
    
dskerror:

    'ferme le handle
    CloseHandle hDevice
End Sub

'=======================================================
'permet de d'�crire de mani�re directe dans le disque
'=======================================================
Public Sub DirectWriteS(ByVal sDrive As String, ByVal iStartSec As Currency, ByVal nBytes As Long, ByVal lBytesPerSector As Long, ByRef sStringToWrite As String)
'/!\ iStartsec et nbytes doivent �tre des multiples de la taille d'un secteur (g�n�ralement 512 octets)
Dim BytesRead As Long
Dim Pointeur As Currency
Dim Ret As Long
Dim hDevice As Long
Dim lLowPart As Long, lHighPart As Long

    'On Error GoTo dskerror
    
    'obtient un path valide pour l'API CreateFIle si n�cessaire
    If Len(sDrive) <> 6 Then sDrive = BuildDrive(sDrive)

    'ouvre le drive
    hDevice = CreateFile(sDrive, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, FILE_FLAG_NO_BUFFERING, 0&)
   
    'quitte si le handle n'est pas valide
    If hDevice = INVALID_HANDLE_VALUE Then Exit Sub
   
    'd�termine le byte de d�part du secteur
    Pointeur = CCur(iStartSec) * CCur(lBytesPerSector)
    
    'transforme un currency en 2 long pour une structure LARGE_INTEGER
    GetLargeInteger Pointeur, lLowPart, lHighPart

    'd�place, dans le fichier (ici un disque) point� par hDevice, le "curseur" au premier
    'byte que l'on veut lire (donn� par deux long)
    Ret = SetFilePointer(hDevice, lLowPart, lHighPart, FILE_BEGIN)  'FILE_BEGIN ==> part du d�but du fichier pour d�compter la DistanceToMove
    If Ret = -1 Then GoTo dskerror
    
    'verrouilage de la zone du disque � �crire
    Call LockFile(hDevice, lLowPart, lHighPart, nBytes, 0)
    
    '�criture disque
    Ret = WriteFile(hDevice, ByVal sStringToWrite, nBytes, Ret, ByVal 0&)
    
    'on vide les buffers internes et on d�v�rouille la zone
    Call FlushFileBuffers(hDevice)
    Call UnlockFile(hDevice, lLowPart, lHighPart, nBytes, 0)
    
dskerror:

    'ferme le handle
    CloseHandle hDevice
End Sub

'=======================================================
'permet de d'�crire de mani�re directe dans le disque
'avec en entr�e un pointeur et un handle de disque
'=======================================================
Public Sub DirectWritePtHandle(ByVal hDevice As Long, ByVal iStartSec As Currency, _
    ByVal nBytes As Long, ByVal lBytesPerSector As Long, _
    ByRef pt As Long)

'/!\ iStartsec et nbytes doivent �tre des multiples de la taille d'un secteur (g�n�ralement 512 octets)

Dim BytesRead As Long
Dim Pointeur As Currency
Dim Ret As Long
Dim lLowPart As Long
Dim lHighPart As Long
   
    'd�termine le byte de d�part du secteur
    Pointeur = CCur(iStartSec) * CCur(lBytesPerSector)
    
    'transforme un currency en 2 long pour une structure LARGE_INTEGER
    GetLargeInteger Pointeur, lLowPart, lHighPart

    'd�place, dans le fichier (ici un disque) point� par hDevice, le "curseur" au premier
    'byte que l'on veut lire (donn� par deux long)
    Ret = SetFilePointer(hDevice, lLowPart, lHighPart, FILE_BEGIN)  'FILE_BEGIN ==> part du d�but du fichier pour d�compter la DistanceToMove
    
    'verrouilage de la zone du disque � �crire
    Call LockFile(hDevice, lLowPart, lHighPart, nBytes, 0)
    
    '�criture disque
    Ret = WriteFile(hDevice, ByVal pt, nBytes, Ret, ByVal 0&)
    
    'on vide les buffers internes et on d�v�rouille la zone
    Call FlushFileBuffers(hDevice)
    Call UnlockFile(hDevice, lLowPart, lHighPart, nBytes, 0)
    
End Sub

'=======================================================
'permet de lire des bytes directement dans le disque
'sortie en String
'=======================================================
Public Sub DirectReadS(ByVal sDrive As String, ByVal iStartSec As Currency, ByVal nBytes As Long, ByVal lBytesPerSector As Long, ByRef sBufferOut As String)
'/!\ iStartsec et nbytes doivent �tre des multiples de la taille d'un secteur (g�n�ralement 512 octets)
Dim BytesRead As Long
Dim Pointeur As Currency
Dim Ret As Long
Dim hDevice As Long
Dim lLowPart As Long, lHighPart As Long

    On Error GoTo dskerror
    
    'obtient un path valide pour l'API CreateFIle si n�cessaire
    If Len(sDrive) <> 6 Then sDrive = BuildDrive(sDrive)

    'ouvre le drive
    hDevice = CreateFile(sDrive, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
   
    'quitte si le handle n'est pas valide
    If hDevice = INVALID_HANDLE_VALUE Then Exit Sub
   
    'd�termine le byte de d�part du secteur
    Pointeur = CCur(iStartSec) * CCur(lBytesPerSector)
    
    'transforme un currency en 2 long pour une structure LARGE_INTEGER
    GetLargeInteger Pointeur, lLowPart, lHighPart

    'd�place, dans le fichier (ici un disque) point� par hDevice, le "curseur" au premier
    'byte que l'on veut lire (donn� par deux long)
    Ret = SetFilePointer(hDevice, lLowPart, lHighPart, FILE_BEGIN)  'FILE_BEGIN ==> part du d�but du fichier pour d�compter la DistanceToMove
    If Ret = -1 Then GoTo dskerror
    
    'cr�ation d'un buffer
    sBufferOut = Space$(nBytes)

    'obtention de la string
    Ret = ReadFile(hDevice, ByVal sBufferOut, nBytes, BytesRead, 0&)

dskerror:

    'ferme le handle
    CloseHandle hDevice
End Sub

'=======================================================
'r�cup�re un handle de disque valide pour la lecture
'=======================================================
Public Function GetDiskHandleRead(ByVal sDrive As String) As Long

    'obtient un path valide pour l'API CreateFIle si n�cessaire
    If Len(sDrive) <> 6 Then sDrive = BuildDrive(sDrive)

    'ouvre le drive
    GetDiskHandleRead = CreateFile(sDrive, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
End Function

'=======================================================
'r�cup�re un handle de disque valide pour l'�criture
'=======================================================
Public Function GetDiskHandleWrite(ByVal sDrive As String) As Long

    'obtient un path valide pour l'API CreateFIle si n�cessaire
    If Len(sDrive) <> 6 Then sDrive = BuildDrive(sDrive)

    'ouvre le drive
    GetDiskHandleWrite = CreateFile(sDrive, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
End Function

'=======================================================
'permet de lire des bytes directement dans le disque
'sortie en String
'demande un handle
'=======================================================
Public Sub DirectReadSHandle(ByVal hDevice As Long, ByVal iStartSec As Currency, ByVal nBytes As Long, ByVal lBytesPerSector As Long, ByRef sBufferOut As String)
'/!\ iStartsec et nbytes doivent �tre des multiples de la taille d'un secteur (g�n�ralement 512 octets)
Dim BytesRead As Long
Dim Pointeur As Currency
Dim Ret As Long
Dim lLowPart As Long, lHighPart As Long
   
    'd�termine le byte de d�part du secteur
    Pointeur = CCur(iStartSec) * CCur(lBytesPerSector)
    
    'transforme un currency en 2 long pour une structure LARGE_INTEGER
    GetLargeInteger Pointeur, lLowPart, lHighPart

    'd�place, dans le fichier (ici un disque) point� par hDevice, le "curseur" au premier
    'byte que l'on veut lire (donn� par deux long)
    Ret = SetFilePointer(hDevice, lLowPart, lHighPart, FILE_BEGIN)  'FILE_BEGIN ==> part du d�but du fichier pour d�compter la DistanceToMove
    
    'cr�ation d'un buffer
    sBufferOut = Space$(nBytes)

    'obtention de la string
    Ret = ReadFile(hDevice, ByVal sBufferOut, nBytes, BytesRead, 0&)

End Sub

'=======================================================
'permet de lire des bytes directement dans le disque PHYSIQUE
'sortie en String
'=======================================================
Public Sub DirectReadSPhys(ByVal bytDrive As Byte, ByVal iStartSec As Currency, ByVal nBytes As Long, ByVal lBytesPerSector As Long, ByRef sBufferOut As String)
'/!\ iStartsec et nbytes doivent �tre des multiples de la taille d'un secteur (g�n�ralement 512 octets)
Dim BytesRead As Long
Dim Pointeur As Currency
Dim Ret As Long
Dim hDevice As Long
Dim lLowPart As Long, lHighPart As Long

    On Error GoTo dskerror

    'ouvre le drive
    hDevice = CreateFile("\\.\PHYSICALDRIVE" & CStr(bytDrive), GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
   
    'quitte si le handle n'est pas valide
    If hDevice = INVALID_HANDLE_VALUE Then Exit Sub
   
    'd�termine le byte de d�part du secteur
    Pointeur = CCur(iStartSec) * CCur(lBytesPerSector)
    
    'transforme un currency en 2 long pour une structure LARGE_INTEGER
    GetLargeInteger Pointeur, lLowPart, lHighPart

    'd�place, dans le fichier (ici un disque) point� par hDevice, le "curseur" au premier
    'byte que l'on veut lire (donn� par deux long)
    Ret = SetFilePointer(hDevice, lLowPart, lHighPart, FILE_BEGIN)  'FILE_BEGIN ==> part du d�but du fichier pour d�compter la DistanceToMove
    If Ret = -1 Then GoTo dskerror
    
    'cr�ation d'un buffer
    sBufferOut = Space$(nBytes)

    'obtention de la string
    Ret = ReadFile(hDevice, ByVal sBufferOut, nBytes, BytesRead, 0&)

dskerror:

    'ferme le handle
    CloseHandle hDevice
End Sub

'=======================================================
'lecture de lLen bytes � l'offset lOffset dans le drive sDrive
'=======================================================
Public Sub ReadDiskBytes(ByVal sDrive As String, ByVal lOffset As Currency, ByVal lLen As Long, ByRef lResult() As Byte, ByVal lBytesPerSector As Long)
Dim lDrive As Long
Dim crPointeur As Currency
Dim tOver As OVERLAPPED
Dim crHi32 As Currency

    On Error GoTo DiskErr

    'obtient un path valide pour l'API CreateFIle si n�cessaire
    If Len(sDrive) <> 6 Then sDrive = BuildDrive(sDrive)
    
    'initialise le tableau r�sultat
    ReDim lResult(0)
        
    'obtient un handle vers le Drive
    lDrive = CreateFile(sDrive, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    
    'si le handle est correct
    If lDrive <> INVALID_HANDLE_VALUE Then
        
        'calcule le "move" � appliquer � lDrive
        'crPointeur = CCur(lOffset) * CCur(lBytesPerSector)
        
        'move pointer
        'Call SetFilePointerEx(lDrive, crPointeur, 0, FILE_BEGIN)
        
        'redimensionne le tableau � la taille convenable du r�sultat
        ReDim lResult(lLen - 1)
        
        crHi32 = 0  'pas de HighOffset par d�faut
        
        'on ajoute 1 au HighOffset si crPointer>2^32
        'car on doit stocker cette valeur Currency en une LARGE_INTEGER
        'pour la structure OverLapped
        GetLargeInteger crPointeur, tOver.Offset, tOver.OffsetHigh
        
        'affecte les valeurs de l'offset (constitu� de la partie High et de la partie Low)
        '� la structure OverLapped
        tOver.Offset = CLng(crPointeur): tOver.OffsetHigh = CLng(crHi32)
        
        'obtient les bytes d�sir�s
        ReadFileEx lDrive, ByVal VarPtr(lResult(0)), lLen, tOver, AddressOf CallBackFunction
    End If
    
DiskErr:
    
    'ferme le handle ouvert
    CloseHandle lDrive
End Sub

'=======================================================
'callback fonction appel�e par l'API ReadFileEx juste au dessus
'fonction non utilis�e, mais sa pr�sence est n�anmoins n�cessaire
'=======================================================
Public Function CallBackFunction()
    Rem N'est pas utile en soit
End Function

'=======================================================
'identique � ReadDiskBytes, mais diff�rent
'=======================================================
'Public Sub DirectReadDriveNT(ByVal sDrive As String, ByVal iStartSec As Currency, ByVal iOffset As Currency, ByVal cBytes As Long, ByVal BytesPerSector As Long, ByRef abResult() As Byte)
'Dim hDevice As Long
'Dim abBuff() As Byte
'Dim nSectors As Currency
'Dim nRead As Long

'    On Error GoTo ErrGestion

    'obtient un path valide pour l'API CreateFIle si n�cessaire
'    If Len(sDrive) <> 6 Then sDrive = BuildDrive(sDrive)

    'calcule le num�ro du secteur lu
'    nSectors = Int((iOffset + cBytes - 1) / BytesPerSector) + 1
    
    'ouvre le drive
'    hDevice = CreateFile(sDrive, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    
    'quitte si le handle n'est pas valide
'    If hDevice = INVALID_HANDLE_VALUE Then Exit Sub
    
    'move pointer
'    Call SetFilePointer(hDevice, iStartSec * BytesPerSector, 0, FILE_BEGIN)
    
    'redimensionne les tableaux r�sultants
'    ReDim abResult(cBytes - 1)
'    ReDim abBuff(nSectors * BytesPerSector - 1)
    
    'appel l'API de lecture
'    Call ReadFile(hDevice, abBuff(0), UBound(abBuff) + 1, nRead, 0&)
    
    'ferme le handle
'    CloseHandle hDevice
    
    'stocke le r�sultat dans le tableau
'    CopyMemory abResult(0), abBuff(iOffset), cBytes
    
'    Exit Sub
'ErrGestion:

    'ferme le handle
'    CloseHandle hDevice
    
'    clsERREUR.AddError "mdlDisk.DirectReadDriveNT", True
'End Sub

'=======================================================
'fonction de recherche de string compl�tes dans un fichier
'stocke dans un tableau de 1 � Ubound
'=======================================================
Public Sub SearchStringInFile(ByVal sFile As String, ByVal lMinimalLength As Long, ByVal bSigns As Boolean, ByVal bMaj As Boolean, ByVal bMin As Boolean, ByVal bNumbers As Boolean, ByVal bAccent As Boolean, ByRef tRes() As SearchResult, Optional PGB As pgrBar)
'Utilisation de l'API CreateFile et ReadFileEx pour une lecture rapide
Dim s As String
Dim strCtemp As String
Dim x As Long
Dim lngLen As Long
Dim bytAsc As Byte
Dim lngFile As Long
Dim strBuffer As String
Dim curByte As Currency
Dim tOver As OVERLAPPED
Dim i As Long

    On Error GoTo ErrGestion
    
    'taille du fichier
    lngLen = cFile.GetFileSize(sFile)
    
    If Not (PGB Is Nothing) Then
        'on initialise la progressabr
        PGB.Min = 0
        PGB.Value = 0
        PGB.Max = lngLen
    End If

    'initialise le tableau
    ReDim tRes(0)

    'obtient le handle du fichier
    lngFile = CreateFile(sFile, GENERIC_READ, FILE_SHARE_READ, 0&, OPEN_EXISTING, 0&, 0&)
    
    'v�rifie que le handle est valide
    If lngFile = INVALID_HANDLE_VALUE Then Exit Sub
    
    strCtemp = vbNullString: x = 1: curByte = 0
    
    'va faire tout le fichier pour tenter de d�nicher des strings
    'cr�� un buffer de 50Ko
    '/!\ le fichier va �tre interpr�t� comme 'fichiers' de 50Ko ==> ne trouve pas les chaines
    'entrecoup�es entre 'fichiers' /!\
    
    strBuffer = String$(51200, 0) 'buffer de 50K
      
    Do Until curByte > lngLen  'tant que le fichier n'est pas fini
    
        x = x + 1
    
        'pr�pare le type OVERLAPPED - obtient 2 long � la place du Currency
        GetLargeInteger curByte, tOver.Offset, tOver.OffsetHigh
        
        'obtient la string sur le buffer
        ReadFileEx lngFile, ByVal strBuffer, 51200, tOver, AddressOf CallBackFunction
    
        strCtemp = vbNullString
        
        'effectue la recherche dans la string
        For i = 0 To 51199
        
            bytAsc = Asc(Mid$(strBuffer, i + 1, 1)) 'prend un byte
            
            If IsCharConsideredInAString(bytAsc, bSigns, bMaj, bMin, bNumbers, bAccent) Then
                'caract�re x est valide
                strCtemp = strCtemp & Chr(bytAsc)
            Else
                strCtemp = Trim$(strCtemp)
                If Len(strCtemp) > lMinimalLength Then
                    'trouv� la chaine correspondante
                    ReDim Preserve tRes(UBound(tRes) + 1)
                    tRes(UBound(tRes)).curOffset = curByte + i - Len(strCtemp) + 1
                    tRes(UBound(tRes)).strString = strCtemp
                End If
                strCtemp = vbNullString
            End If
        Next i
        
        If Len(strCtemp) > lMinimalLength Then
            'trouv� la derni�re chaine possible (dernier byte compris dans cette chaine)
            ReDim Preserve tRes(UBound(tRes) + 1)
            tRes(UBound(tRes)).curOffset = curByte + 51199 - Len(strCtemp) + 1
            tRes(UBound(tRes)).strString = strCtemp
        End If
        
        If (x Mod 10) = 0 Then
            If Not (PGB Is Nothing) Then PGB.Value = curByte    'refresh progressbar
            DoEvents    'rend la main
        End If
        
        curByte = curByte + 51200   'incr�mente la position
    Loop
    
    If Not (PGB Is Nothing) Then PGB.Value = lngLen
    
    Let strBuffer = vbNullString
    CloseHandle lngFile 'ferme le handle du fichier
    
    Exit Sub
ErrGestion:

    CloseHandle lngFile 'ferme le handle du fichier
    
    clsERREUR.AddError "mdlDisk.SearchStringInFile", True
End Sub

'=======================================================
'fonction de recherche de string dans un fichier
'de 1 � Ubound
'=======================================================
Public Sub SearchForStringFile(ByVal sFile As String, ByVal sMatch As String, ByVal bCasse As Boolean, ByRef tRes() As Long, Optional PGB As pgrBar)
'Utilisation de l'API CreateFile et ReadFileEx pour une lecture rapide
Dim s As String
Dim x As Long
Dim lngLen As Long
Dim bytAsc As Byte
Dim lngFile As Long
Dim strBuffer As String
Dim strBuffer2 As String
Dim strBufT As String
Dim curByte As Currency
Dim tOver As OVERLAPPED
Dim i As Long

    On Error GoTo ErrGestion

    'taille du fichier
    lngLen = cFile.GetFileSize(sFile)
    
    If Not (PGB Is Nothing) Then
        'on initialise la progressabr
        PGB.Min = 0
        PGB.Value = 0
        PGB.Max = lngLen
    End If

    'initialise le tableau
    ReDim tRes(0)

    'obtient le handle du fichier
    lngFile = CreateFile(sFile, GENERIC_READ, FILE_SHARE_READ, 0&, OPEN_EXISTING, 0&, 0&)
    
    'v�rifie que le handle est valide
    If lngFile = INVALID_HANDLE_VALUE Then Exit Sub
    
    x = 1: curByte = 0
    
    'va faire tout le fichier pour tenter de d�nicher des strings
    'cr�� un buffer de 50Ko
    '/!\ le fichier va �tre interpr�t� comme 'fichiers' de 50Ko ==> ne trouve pas les chaines
    'entrecoup�es entre 'fichiers' /!\
    
    strBuffer = String$(51200, 0) 'buffer de 50K
    
    If bCasse = False Then sMatch = LCase$(sMatch)  'cherche que les minuscules
      
    Do Until curByte > lngLen  'tant que le fichier n'est pas fini
    
        x = x + 1
        
        If bCasse = False Then strBuffer = LCase$(strBuffer)
        
        strBuffer2 = Replace$(strBuffer, sMatch, vbNullString) 'affecte l'ancien buffer � la partie qui sera concaten�e
        'devant le nouveau buffer ==> permet de rechercher dans tout le fichier
        'en prenant en compte les strings coup�es entre 2 buffers
        'Enl�ve le r�sultat pr�c�dent (avec Replace) pour �viter lse doublons
    
        'pr�pare le type OVERLAPPED - obtient 2 long � la place du Currency
        GetLargeInteger curByte, tOver.Offset, tOver.OffsetHigh
        
        'obtient la string sur le buffer
        ReadFileEx lngFile, ByVal strBuffer, 51200, tOver, AddressOf CallBackFunction
    
        
        strBufT = strBuffer2 & strBuffer 'concat�nation de l'ancien et du nouveau buffer
     
        If bCasse = False Then strBufT = LCase$(strBufT)   'minuscules only
         
        'tant que la string contient le match
        While InStr(1, strBufT, sMatch, vbBinaryCompare) <> 0
            'trouv� une string ==> l'ajoute
            ReDim Preserve tRes(UBound(tRes) + 1)
            tRes(UBound(tRes)) = curByte + InStr(1, strBufT, sMatch, vbBinaryCompare) + Len(strBuffer) - Len(strBufT) - 1
            
            'raccourci le buffer
            strBufT = Right$(strBufT, Len(strBufT) - InStr(1, strBufT, sMatch, vbBinaryCompare) - Len(sMatch) + 1)
        Wend
        
        If (x Mod 10) = 0 Then
            If Not (PGB Is Nothing) Then PGB.Value = curByte    'refresh progressbar
            DoEvents    'rend la main
        End If
        
        curByte = curByte + Len(strBuffer2) + Len(strBuffer) 'incr�mente la position
    Loop
    
    If Not (PGB Is Nothing) Then PGB.Value = lngLen
    
    Let strBufT = vbNullString
    Let strBuffer2 = vbNullString
    Let strBuffer = vbNullString
    
    CloseHandle lngFile 'ferme le handle du fichier

    Exit Sub
ErrGestion:

    CloseHandle lngFile 'ferme le handle du fichier
    
    clsERREUR.AddError "mdlDisk.SearchForStringFile", True
End Sub

'=======================================================
'fonction de recherche de string dans un disque
'de 1 � Ubound
'=======================================================
Public Sub SearchForStringDisk(ByVal sDrive As String, ByVal sMatch As String, ByVal bCasse As Boolean, ByRef tRes() As Long, Optional PGB As pgrBar, Optional ByVal IsPhys As Boolean = False)
'Utilisation de l'API CreateFile et ReadFileEx pour une lecture rapide
Dim x As Long
Dim r() As Byte
Dim bytAsc As Byte
Dim strDrive As String
Dim strBufT As String
Dim curByte As Currency
Dim tOver As OVERLAPPED
Dim i As Currency
Dim btPerSec As Long
Dim nbSec As Currency
Dim cDrive As clsDrive
Dim clsDrive As clsDiskInfos

    On Error GoTo ErrGestion

    'initialise les tableaux
    ReDim tRes(0): ReDim r(0)
    
    'r�-obtient les infos sur les secteurs (nombre et taille)
    Set clsDrive = New clsDiskInfos
        
    If IsPhys Then
        Set cDrive = clsDrive.GetPhysicalDrive(Val(sDrive))
    Else
        'formate le nom du disque
        strDrive = BuildDrive(Right$(sDrive, 3))

        Set cDrive = clsDrive.GetLogicalDrive(strDrive)
    End If
    
    'affecte les infos sur les secteurs aux variables
    nbSec = cDrive.TotalLogicalSectors
    btPerSec = cDrive.BytesPerSector
    
    If Not (PGB Is Nothing) Then
        'on initialise la progressabr
        PGB.Min = 0
        PGB.Value = 0
        PGB.Max = nbSec
    End If

    If bCasse = False Then sMatch = LCase$(sMatch)  'cherche que les minuscules

    For i = 0 To nbSec Step 20000  'pour chaque secteur logique
        
        If IsPhys Then
            'obtient les bytes du secteur visualis� en partie
            DirectReadSPhys Val(sDrive), CCur(i), 20000 * btPerSec, btPerSec, strBufT
        Else
            'obtient les bytes du secteur visualis� en partie
            DirectReadS strDrive, CCur(i), 20000 * btPerSec, btPerSec, strBufT
        End If
        
        If bCasse = False Then strBufT = LCase$(strBufT)    'cherche que des minuscules (pas de casse respect�e)
         
        'tant que la string contient le match
        While InStr(1, strBufT, sMatch, vbBinaryCompare) <> 0
            'trouv� une string ==> l'ajoute
            ReDim Preserve tRes(UBound(tRes) + 1)
            tRes(UBound(tRes)) = i * btPerSec + InStr(1, strBufT, sMatch, vbBinaryCompare) + 10240000 - Len(strBufT) - 1
            
            'raccourci le buffer
            strBufT = Right$(strBufT, Len(strBufT) - InStr(1, strBufT, sMatch, vbBinaryCompare) - Len(sMatch) + 1)
        Wend
        
        If Not (PGB Is Nothing) Then PGB.Value = i    'refresh progressbar
        DoEvents    'rend la main
        
    Next i
    
    If Not (PGB Is Nothing) Then PGB.Value = nbSec
    
    Let strBufT = vbNullString

    Exit Sub
ErrGestion:
    clsERREUR.AddError "mdlDisk.SearchForStringFile", True
End Sub

'=======================================================
'd�termine si un byte est consid�r� comme convenable en fonction
'des param�tres Afficher : min, MAJ, nbres, signes
'function utilis�e directement avec les proc�dures de SearchStringIn...
'=======================================================
Public Function IsCharConsideredInAString(ByVal bytChar As Byte, ByVal bSigns As Boolean, ByVal bMaj As Boolean, ByVal bMin As Boolean, ByVal bNumbers As Boolean, ByVal bAccent As Boolean) As Boolean
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
'�criture de bytes dans un disque physique
'=======================================================
Public Sub DirectWrite(ByVal sDrive As String, ByVal iStartSec As Currency, ByVal nBytes As Long, ByVal lBytesPerSector As Long, ByRef sStringToWrite As String)
'/!\ iStartsec et nbytes doivent �tre des multiples de la taille d'un secteur (g�n�ralement 512 octets)
Dim BytesRead As Long
Dim Pointeur As Currency
Dim Ret As Long
Dim OVER As OVERLAPPED
Dim hDevice As Long
Dim lLowPart As Long, lHighPart As Long

    'On Error GoTo dskerror
    
    'obtient un path valide pour l'API CreateFIle si n�cessaire
    If Len(sDrive) <> 6 Then sDrive = BuildDrive(sDrive)

    'ouvre le drive
    hDevice = CreateFile(sDrive, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, FILE_FLAG_OVERLAPPED, 0&)
   
    'quitte si le handle n'est pas valide
    If hDevice = INVALID_HANDLE_VALUE Then Exit Sub
   
    'd�termine le byte de d�part du secteur
    Pointeur = CCur(iStartSec) * CCur(lBytesPerSector)

    'transforme un currency en 2 long pour une structure LARGE_INTEGER
    GetLargeInteger Pointeur, lLowPart, lHighPart
    
    'd�finit le OVERLAPPED
    With OVER
        .Offset = lLowPart
        .OffsetHigh = lHighPart
    End With
    
    'd�place, dans le fichier (ici un disque) point� par hDevice, le "curseur" au premier
    'byte que l'on veut lire (donn� par deux long)
    
    'Ret = SetFilePointer(hDevice, lLowPart, lHighPart, FILE_BEGIN)  'FILE_BEGIN ==> part du d�but du fichier pour d�compter la DistanceToMove
    'If Ret = -1 Then GoTo dskerror
    
    'verrouilage de la zone du disque � �crire
    Call LockFile(hDevice, lLowPart, lHighPart, nBytes, 0)
    
    '�criture disque
    Ret = WriteFileEx(hDevice, ByVal sStringToWrite, nBytes, OVER, AddressOf CallBackFunction)
    
    If Ret = 0 Then Stop
    
    'on vide les buffers internes et on d�v�rouille la zone
    Call FlushFileBuffers(hDevice)
    Call UnlockFile(hDevice, lLowPart, lHighPart, nBytes, 0)
    
dskerror:

    'ferme le handle
    CloseHandle hDevice
End Sub

'=======================================================
'efface compl�tement un fichier du disque dur
'=======================================================
Public Function ShreddFile(ByVal sFile As String, ByVal nPass As Integer, _
    PGB As ProgressBar_OCX.pgrBar) As Boolean
Dim hFile As Long
Dim sFile2 As String
Dim Ret As Long
Dim ret2 As Long
Dim x As Long
Dim tTime As FILETIME
Dim tsTime As SYSTEMTIME
    
    On Error GoTo ErrGestion

    ShreddFile = False

    '�tapes de shredd
    '1) on efface tous les bytes du disk qui �taient utilis�s pour un fichier
    '2) on renomme le fichier avec un nom al�atoire
    '3) on change la date
    '4) on efface le fichier renomm�
  
    'affecte l'attribut normal (car le kill de VB ne fonctionne que pour l'attribut normal)
    ret2 = SetFileAttributes(sFile, FILE_ATTRIBUTE_NORMAL)

    'obtient le handle du fichier
    hFile = CreateFile(sFile, GENERIC_WRITE, FILE_SHARE_WRITE, 0, TRUNCATE_EXISTING, FILE_FLAG_NO_BUFFERING Or FILE_FLAG_WRITE_THROUGH, 0)
    
    'initialise le PGB
    With PGB
        .Max = nPass * 3
        .Min = 0
        .Value = 0
    End With
        
    '//effectue les diff�rentes passes de la sanitization des fichiers
    For x = 1 To nPass
        
        '&H55
        Call WriteBytesToFile(sFile, String$(cFile.GetFileSize(sFile), 85), 0)
        
        PGB.Value = PGB.Value + 1
        
        '&HAA
        Call WriteBytesToFile(sFile, String$(cFile.GetFileSize(sFile), 170), 0)
        
        PGB.Value = PGB.Value + 1
        
        'Random
        
        
        PGB.Value = PGB.Value + 1
        
        'flush buffers
        Ret = FlushFileBuffers(hFile)
    Next x
    
    'ferme le handle ouvert
    CloseHandle hFile
    
    'renomme le fichier de mani�re bidon (car le nom reste quand m�me dans le fichier MFT)
    Randomize   'pour obtenir un pseudo-hasard
    sFile2 = Left$(sFile, 3) & Replace(CStr(Rnd), ",", vbNullString) & ".temp" 'd�place � la racine, mais peu importe car suppression. N�cessite une extension
    cFile.Rename sFile, sFile2
    
    'r�-obtient le handle du fichier
    hFile = CreateFile(sFile2, GENERIC_WRITE, FILE_SHARE_WRITE, 0, TRUNCATE_EXISTING, FILE_FLAG_NO_BUFFERING Or FILE_FLAG_WRITE_THROUGH, 0)
 
    'change la date
    With tsTime
        .wYear = 1999
        .wMonth = 1
        .wDay = 1
        .wDayOfWeek = Weekday("1/1/1999")
    End With
    SystemTimeToFileTime tsTime, tTime
    SetFileTime hFile, tTime, tTime, tTime
    
    'referme d�finitivement le handle
    CloseHandle hFile
    
    'on efface le fichier (deux suppressions si renommage rat�)
    cFile.KillFile sFile2
    cFile.KillFile sFile
    
    'v�rifie que toutes les �tapes sont OK
    If cFile.FileExists(sFile) Or cFile.FileExists(sFile2) Then Exit Function   'rat�
    If Ret = 0 Or ret2 = 0 Then Exit Function  'rat�
    If hFile = -1 Then Exit Function 'rat�
    
    ShreddFile = True

    Exit Function
ErrGestion:
    clsERREUR.AddError "mdlDisk.ShreddFile", True
End Function

'=======================================================
'obtient les infos sur l'emplacement (clusters) du fichier sur le disque
'=======================================================
Public Function GetFileBitmap(File As String) As FileClusters
Dim hFile As Long 'handle de fichier dont on veut la carte des clusters
Dim FileBitmap As RETRIEVAL_POINTERS_BUFFER 'carte des clusters du fichier
Dim nExtents As Long 'nombre d'extents (fragments) du fichier
Dim StartingAddress As LARGE_INTEGER 'VCN de d�but de la carte du fichier
Dim bt As Long 'nombre d'octets renvoy�s
Dim status As Long '�tat de l'op�ration
Dim x As Long 'compteur

    'ouvre le fichier avec les droits de la d�placer juste pour voir si on pourrait le d�placer
    hFile = CreateFile(File, FILE_READ_ACCESS Or &H10000, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    
    'copie le nom du fichier
    GetFileBitmap.File = File
    
    'si on ne peut pas l'ouvrir pour d�placement
    If hFile = -1 Then
        'pas d�placable
        GetFileBitmap.Moveable = False
        'on essaie de l'ouvrir en lecture
        hFile = CreateFile(File, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
        'si pas possible : fichier syst�me vital
        If hFile = -1 Then Exit Function
    Else
        'd�placable
        GetFileBitmap.Moveable = True
    End If
    
    'on demande la carte compl�te du fichier donc depuis le d�but
    StartingAddress.HighDWORD = 0
    StartingAddress.LowDWORD = 0
    
    'on demande la carte du fichier tant qu'il y en a encore � r�cup�rer
    Do
        'demande un morceau de 512 fragments
        DeviceIoControl hFile, FSCTL_GET_RETRIEVAL_POINTERS, StartingAddress, 8&, FileBitmap, Len(FileBitmap), bt, 0&
        status = Err.LastDllError
        
        'si la partie contient des fragments
        If FileBitmap.ExtentCount Then
            'ajoute le nombre de fragments de la partie de carte au nombre de fragments du fichier
            GetFileBitmap.ExtentsCount = GetFileBitmap.ExtentsCount + FileBitmap.ExtentCount
            'fait de la place pour ajouter les fragments
            ReDim Preserve GetFileBitmap.Extents(GetFileBitmap.ExtentsCount - 1)
        End If
        
        'si le nombre de fragments est > 512
        If FileBitmap.ExtentCount > 512 Then
            'on copie les 512 premier fragments, car notre structure allou�e ne peut pas en contenir plus
            CopyMemory GetFileBitmap.Extents(nExtents), FileBitmap.Extents(0), 512& * 16&
            'on avance de 512 fragments
            nExtents = nExtents + 512
        'sinon s'il y a moins de 512 fragments dans la partie de carte
        ElseIf FileBitmap.ExtentCount Then
            'on les copies
            CopyMemory GetFileBitmap.Extents(nExtents), FileBitmap.Extents(0), FileBitmap.ExtentCount * 16&
            'on avance du nombre de fragments renvoy�s
            nExtents = nExtents + FileBitmap.ExtentCount
        End If
        
        'on avance dans le fichier jusqu'� l'offset (depuis le d�but du fichier) du prochain fragment apr�s ceux que l'on a d�j� obtenus
        StartingAddress.LowDWORD = FileBitmap.Extents(511).NextVcn.LowDWORD
        StartingAddress.HighDWORD = FileBitmap.Extents(511).NextVcn.HighDWORD
        
    'tant que l'on n'est pas � la fin des fragments du fichier
    Loop While status = ERROR_MORE_DATA
    
    CloseHandle hFile
End Function

'Public Function GetFileBitmap(File As String) As FileClusters
'Dim hFile As Long 'handle de fichier dont on veut la carte des clusters
'Dim FileBitmap As RETRIEVAL_POINTERS_BUFFER 'carte des clusters du fichier
'Dim nExtents As Long
'Dim StartingAddress As LARGE_INTEGER 'VCN de d�but de la carte du fichier
'Dim status As Long '�tat de l'op�ration
'Dim x As Long
'Dim tmp As FileClusters

'    On Error GoTo ErrGestion
    
'    hFile = CreateFile(File, FILE_READ_ACCESS, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
'    tmp.File = File
'    If hFile = -1 Then
'        tmp.Moveable = False
'        hFile = CreateFile(File, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
'        If hFile = -1 Then Exit Function
'    Else
'        tmp.Moveable = True
'    End If
    
    'on demande la carte compl�te du fichier
'    StartingAddress.HighDWORD = 0
'    StartingAddress.LowDWORD = 0
    
    'on demande la carte
'    Do
'        DeviceIoControl hFile, FSCTL_GET_RETRIEVAL_POINTERS, StartingAddress, 8&, FileBitmap, Len(FileBitmap), 0&, 0&
'        status = Err.LastDllError
        'If (FileBitmap.StartingVcn.LowDWORD = 0) And (FileBitmap.StartingVcn.HighDWORD = 0) Then
'            If FileBitmap.ExtentCount Then
'                tmp.ExtentsCount = tmp.ExtentsCount + FileBitmap.ExtentCount
'                ReDim Preserve tmp.Extents(tmp.ExtentsCount - 1)
'            End If
        'End If
'        If FileBitmap.ExtentCount > 1024 Then
'            CopyMemory tmp.Extents(nExtents), FileBitmap.Extents(0), 1024& * 16&
'            nExtents = nExtents + 1024
'        ElseIf FileBitmap.ExtentCount Then
'            CopyMemory tmp.Extents(nExtents), FileBitmap.Extents(0), FileBitmap.ExtentCount * 16&
'            nExtents = nExtents + FileBitmap.ExtentCount
'        End If
'        StartingAddress.LowDWORD = StartingAddress.LowDWORD + FileBitmap.Extents(1023).NextVcn.LowDWORD
'        StartingAddress.HighDWORD = StartingAddress.HighDWORD + FileBitmap.Extents(1023).NextVcn.HighDWORD
'    Loop While status = ERROR_MORE_DATA
'    CloseHandle hFile
    
'    GetFileBitmap = tmp
    
'    Exit Function
'ErrGestion:
'    clsERREUR.AddError "mdlDisk.GetFileBitMap", True
'End Function

'=======================================================
'version simplifi�e de GetFileBitmap ==> n'obtient que le nombre
'de fragments d'un fichier
'=======================================================
'Public Function GetFileFragmentCount(File As String) As FileClusters2
'Dim hFile As Long 'handle de fichier dont on veut la carte des clusters
'Dim FileBitmap As RETRIEVAL_POINTERS_BUFFER 'carte des clusters du fichier
'Dim StartingAddress As LARGE_INTEGER 'VCN de d�but de la carte du fichier
'Dim status As Long '�tat de l'op�ration
'Dim tmp As FileClusters2

'    On Error GoTo ErrGestion
    
'    hFile = CreateFile(File, FILE_READ_ACCESS Or GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
'    tmp.File = File
    
    'on demande la carte compl�te du fichier
'    StartingAddress.HighDWORD = 0
'    StartingAddress.LowDWORD = 0
    
    'on demande la carte
'    Do
'        DeviceIoControl hFile, FSCTL_GET_RETRIEVAL_POINTERS, StartingAddress, 8&, FileBitmap, Len(FileBitmap), 0&, 0&
'        status = Err.LastDllError
'            If FileBitmap.ExtentCount Then
'                tmp.ExtentsCount = tmp.ExtentsCount + FileBitmap.ExtentCount
'            End If
                
'    Loop While status = ERROR_MORE_DATA
'    CloseHandle hFile
    
'    GetFileFragmentCount = tmp
    
'    Exit Function
'ErrGestion:
'    clsERREUR.AddError "mdlDisk.GetFileFragmentCount", True
'End Function

'=======================================================
'obtient le nombre de fragments pour chaque fichier d'un drive
'=======================================================
'Public Function GetVolumeFilesBitmap(Volume As String, Optional Progress As pgrbar, Optional SubFolder As Boolean = True) As FileClusters()
'Dim tmp() As FileClusters2
'Dim Files() As String
'Dim x As Long, ub As Long
    
'    On Error GoTo ErrGestion
'    DoEvents
'    GetVolumeFiles Volume, Files, True, SubFolder
'    ub = UBound(Files)
'    ReDim tmp(ub)
'
'    If IsMissing(Progress) = False Then
'        Progress.Min = 0
'        Progress.Max = ub + 1
'        Progress.Value = 0
'    End If
'    For x = 0 To ub
'        tmp(x) = GetFileFragmentCount(Files(x))
'        If (x Mod 250) = 0 Then
'            Progress.Value = IIf(Progress.Value + 250 < Progress.Max, Progress.Value + 250, Progress.Max)
'            DoEvents
'        End If
'    Next
'    GetVolumeFilesBitmap = tmp

'    Exit Function
'ErrGestion:
'    clsERREUR.AddError "mdlDisk.GetVolumeFilesBitMap", True
'End Function

'=======================================================
'liste tous les fichiers d'un drive
'=======================================================
'Public Sub GetVolumeFiles(ByVal Directory As String, Files() As String, Optional Begin As Boolean = False, Optional SubFolder As Boolean = True)
'Dim FileInfo As WIN32_FIND_DATA, hFind As Long
'Static ub As Long
'    If Begin = True Then ub = 0
'    DoEvents
'    hFind = FindFirstFile(Directory & "*", FileInfo)
'    If hFind <> -1 Then
'        If (FileInfo.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
'            If InStr(FileInfo.cFileName, ".") <> 1 And SubFolder Then GetVolumeFiles Directory & Mid$(FileInfo.cFileName, 1, InStr(FileInfo.cFileName, vbNullChar) - 1) & "\", Files, False
'        Else
'            ub = ub + 1
'            ReDim Preserve Files(ub)
'            Files(ub) = Directory & Mid$(FileInfo.cFileName, 1, InStr(FileInfo.cFileName, vbNullChar) - 1)
'        End If
'        Do While FindNextFile(hFind, FileInfo)
'            If (FileInfo.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
'                If InStr(FileInfo.cFileName, ".") <> 1 And SubFolder Then GetVolumeFiles Directory & Mid$(FileInfo.cFileName, 1, InStr(FileInfo.cFileName, vbNullChar) - 1) & "\", Files, False
'            Else
'                ub = ub + 1
'                ReDim Preserve Files(ub)
'                Files(ub) = Directory & Mid$(FileInfo.cFileName, 1, InStr(FileInfo.cFileName, vbNullChar) - 1)
'            End If
'        Loop
'    End If
'    FindClose hFind
'End Sub
