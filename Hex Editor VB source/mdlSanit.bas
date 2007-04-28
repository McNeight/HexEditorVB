Attribute VB_Name = "mdlSanit"
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
'MODULE POUR LA SANITIZATION
'=======================================================

'=======================================================
'lance la sanitization des fichiers
'=======================================================
Public Sub SanitFilesNow(LV As ListView, PGB As pgrBar)
Dim x As Long
Dim s As String
Dim pt As Long
Dim hFile As Long
Dim cFile As filesystemlibrary.FileSystem
Dim curSize As Currency
Dim nbBuf As Long
Dim lLastSize As Long
Dim y As Long

    Call SetCurrentDirectoryA(App.Path)
    
    Set cFile = New filesystemlibrary.FileSystem

    With PGB
        .Max = LV.ListItems.Count
        .Min = 0
        .Value = 0
    End With

    'pour chaque fichier
    For x = 1 To LV.ListItems.Count
        
TestPres:
        If cFile.FileExists(LV.ListItems.Item(x).Text) = False Then x = x + 1: _
            GoTo TestPres
        
        'on fait çà par buffer de 2Mo
        curSize = cFile.GetFileSize(LV.ListItems.Item(x).Text)
        
        If curSize < 2097152 Then
        
            'un seul buffer, on génère une string aléatoire unique
            pt = GetPtRandomString
            
            'on récupère le handle du fichier
            hFile = GetFileHandleWrite(LV.ListItems.Item(x))
    
            'on écrit dans le fichier
            '// &H55
            Call WriteBytesToFileHandle(hFile, p55, 0, curSize)
    
            '// &HAA
            Call WriteBytesToFileHandle(hFile, pAA, 0, curSize)
    
            '//random string
            Call WriteBytesToFileHandle(hFile, pt, 0, curSize)
            Call FreePtRandomString(pt)
    
            'rend la main
            PGB.Value = x
            DoEvents
                            
            'referme le handle
            CloseHandle hFile
            
        Else
        
            'plusieurs buffers
            
            'calcule le nombre de buffers et la taille du dernier buffer
            nbBuf = Int(curSize / 2097152)
            lLastSize = curSize - 2097152 * nbBuf
            
            For y = 1 To nbBuf

                'on récupère le handle du fichier
                hFile = GetFileHandleWrite(LV.ListItems.Item(x))
                
                'on récupère un pointeur sur une string de 2Mo
                pt = GetPtRandomString
        
                'on écrit dans le fichier
                '// &H55
                 Call WriteBytesToFileHandle(hFile, p55, (y - 1) * 2097152, _
                    2097152)
        
                '// &HAA
                Call WriteBytesToFileHandle(hFile, pAA, (y - 1) * 2097152, _
                    2097152)
        
                '//random string
                Call WriteBytesToFileHandle(hFile, pt, (y - 1) * 2097152, 2097152)
        
                'on libère les 2Mo
                Call FreePtRandomString(pt)
        
                'rend la main de tps en tps
                DoEvents
            Next y
            
            's'occupe du dernier buffer (plus petit)
            pt = GetPtRandomString

            '// &H55
            Call WriteBytesToFileHandle(hFile, p55, nbBuf * 2097152, _
                lLastSize)
    
            '// &HAA
            Call WriteBytesToFileHandle(hFile, pAA, nbBuf * 2097152, _
                lLastSize)
    
            '//random string
            Call WriteBytesToFileHandle(hFile, pt, nbBuf * 2097152, lLastSize)
            Call FreePtRandomString(pt)
            
            'referme le handle
            CloseHandle hFile
            
            DoEvents
            PGB.Value = x

        End If
        
    Next x

    PGB.Value = PGB.Max

    'libère classe + affiche message
    MsgBox frmSanitization.Lang.GetString("_SanitOk"), vbInformation, frmSanitization.Lang.GetString("_SanitOk")

End Sub

'=======================================================
'lance la sanitization du disque
'=======================================================
Public Sub SanitDiskNow(ByVal sDisk As String, PGB As pgrBar)
Dim cDriv As clsDrive
Dim clsDriv As clsDiskInfos
Dim x As Long
Dim secPerString As Long
Dim s As String
Dim pt As Long
Dim hDevice As Long
Dim bPerSec As Long
Dim curOp As Currency
Dim bMax As Long

    Call SetCurrentDirectoryA(App.Path)

    'vérifie que le disque est accessible
    Set clsDriv = New clsDiskInfos
    With frmSanitization.Lang
        If clsDriv.IsLogicalDriveAccessible(sDisk) = False Then
            MsgBox .GetString("_DiskNotR"), vbCritical, .GetString("_War")
            Exit Sub
        End If
    End With
    
    'récupère les infos sue le disque
    Set cDriv = clsDriv.GetLogicalDrive(sDisk)
    
    'nombre de secteurs pour une string de 2Mo
    secPerString = 2097152 / cDriv.BytesPerSector
    
    'handle du disque
    hDevice = GetDiskHandleWrite(sDisk)
    
    With PGB
        .Max = cDriv.TotalPhysicalSectors / secPerString
        .Min = 0
        .Value = 0
    End With
    
    bPerSec = cDriv.BytesPerSector
    bMax = Int(cDriv.TotalPhysicalSectors / secPerString) + 1 '+1
    
    'pour chaque secteur
    For x = 1 To bMax
    
        'on récupère un pointeur sur une string de 2Mo
        pt = GetPtRandomString
        
        'calcul unique
        curOp = CCur(x * secPerString)
        
        'on écrit dans le disque
        '// &H55
        Call DirectWritePtHandle(hDevice, curOp, 2097152, bPerSec, p55)
        
        '// &HAA
        Call DirectWritePtHandle(hDevice, curOp, 2097152, bPerSec, pAA)
            
        '//random string
        Call DirectWritePtHandle(hDevice, curOp, 2097152, bPerSec, pt)
        
        'on libère les 2Mo
        Call FreePtRandomString(pt)
        
        'rend la main de tps en tps
        If (x Mod 5) = 0 Then
            PGB.Value = x
            DoEvents
        End If
        
    Next x
    
    '//s'occupe de l'entête du disque
        pt = GetPtRandomString
        
        'on écrit dans le disque
        '// &H55
        Call DirectWritePtHandle(hDevice, 0, 2097152, bPerSec, p55)
        
        '// &HAA
        Call DirectWritePtHandle(hDevice, 0, 2097152, bPerSec, pAA)
            
        '//random string
        Call DirectWritePtHandle(hDevice, 0, 2097152, bPerSec, pt)
        
        'on libère les 2Mo
        Call FreePtRandomString(pt)
    
    
    PGB.Value = PGB.Max
    
    'referme le handle
    CloseHandle hDevice
    
    'libère classe + affiche message
    Set clsDriv = Nothing
    MsgBox frmSanitization.Lang.GetString("_SanitOk"), vbInformation, frmSanitization.Lang.GetString("_SanitOk")

End Sub

'=======================================================
'récupère un pointeur sur une string aléatoire de 2Mo
'(2*1024^2 octets) générée par la dll bnAlloc
'/!\ ne pas oublier de libérer la mémoire une fois la string
'utilisée et plus utile
'=======================================================
Public Function GetPtRandomString() As Long
    GetPtRandomString = bnAlloc2MoAlea
End Function
Public Sub FreePtRandomString(pt As Long)
    Call bnFreeAlloc(pt)
End Sub
