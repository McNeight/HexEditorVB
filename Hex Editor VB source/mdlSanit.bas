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
Dim hFile As Long
Dim curSize As Currency
Dim nbBuf As Long
Dim lLastSize As Long
Dim y As Long
Dim cASM As CAsmProc
Dim Tbl(2097151) As Byte

    'instancie les classes
    Set cASM = New CAsmProc

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
            Call cASM.bnAlloc2MoAlea(Tbl(0))
            
            'on récupère le handle du fichier
            hFile = GetFileHandleWrite(LV.ListItems.Item(x))
    
            'on écrit dans le fichier
            '// &H55
            Call WriteBytesToFileHandle(hFile, p55, 0, curSize)
    
            '// &HAA
            Call WriteBytesToFileHandle(hFile, pAA, 0, curSize)
    
            '//random string
            Call WriteBytesToFileHandle(hFile, VarPtr(Tbl(0)), 0, curSize)
    
            'rend la main
            PGB.Value = x
            DoEvents
                            
            'referme le handle
            Call CloseHandle(hFile)
            
        Else
        
            'plusieurs buffers
            
            'calcule le nombre de buffers et la taille du dernier buffer
            nbBuf = Int(curSize / 2097152)
            lLastSize = curSize - 2097152 * nbBuf
            
            For y = 1 To nbBuf

                'on récupère le handle du fichier
                hFile = GetFileHandleWrite(LV.ListItems.Item(x))
                
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
            Call WriteBytesToFileHandle(hFile, p55, nbBuf * 2097152, _
                lLastSize)
    
            '// &HAA
            Call WriteBytesToFileHandle(hFile, pAA, nbBuf * 2097152, _
                lLastSize)
    
            '//random string
            Call WriteBytesToFileHandle(hFile, VarPtr(Tbl(0)), nbBuf * 2097152, lLastSize)
            
            'referme le handle
            Call CloseHandle(hFile)
            
            DoEvents
            PGB.Value = x

        End If
        
    Next x

    PGB.Value = PGB.Max

    'libère classe + affiche message
    Set cASM = Nothing
    MsgBox frmSanitization.Lang.GetString("_SanitOk"), vbInformation, frmSanitization.Lang.GetString("_SanitOk")

End Sub

'=======================================================
'lance la sanitization du disque physique
'=======================================================
Public Sub SanitPhysDiskNow(ByVal DiskNumber As Byte, PGB As pgrBar)
Dim cDisk As FileSystemLibrary.PhysicalDisk
Dim x As Long
Dim secPerString As Long
Dim s As String
Dim hDevice As Long
Dim bPerSec As Long
Dim curOp As Currency
Dim bMax As Long
Dim cASM As CAsmProc
Dim Tbl(2097151) As Byte

    Set cASM = New CAsmProc

    'vérifie que le disque est accessible
    With frmSanitization.Lang
        If cFile.IsPhysicalDiskAvailable(DiskNumber) = False Then
            MsgBox .GetString("_DiskNotR"), vbCritical, .GetString("_War")
            Exit Sub
        End If
    End With
    
    'récupère les infos sur le disque
    Set cDisk = cFile.GetPhysicalDisk(DiskNumber)
    
    'nombre de secteurs pour une string de 2Mo
    secPerString = 2097152 / cDisk.BytesPerSector
    
    'handle du disque
    hDevice = GetPhysicalDiskHandleWrite(DiskNumber)
    
    With PGB
        .Max = cDisk.TotalPhysicalSectors / secPerString
        .Min = 0
        .Value = 0
    End With
    
    bPerSec = cDisk.BytesPerSector
    bMax = Int(cDisk.TotalPhysicalSectors / secPerString) + 1 '+1
    
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
        If (x Mod 5) = 0 Then
            PGB.Value = x
            DoEvents
        End If
        
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
    
    
    PGB.Value = PGB.Max
    
    'referme le handle
    Call CloseHandle(hDevice)
    
    'libère classes + affiche message
    Set cDisk = Nothing
    Set cASM = Nothing
    MsgBox frmSanitization.Lang.GetString("_SanitOk"), vbInformation, frmSanitization.Lang.GetString("_SanitOk")

End Sub

'=======================================================
'lance la sanitization du disque logique
'=======================================================
Public Sub SanitDiskNow(ByVal sDisk As String, PGB As pgrBar)
Dim cDriv As FileSystemLibrary.Drive
Dim x As Long
Dim secPerString As Long
Dim s As String
Dim hDevice As Long
Dim bPerSec As Long
Dim curOp As Currency
Dim bMax As Long
Dim Tbl(2097151) As Byte
Dim cASM As CAsmProc
    
    Set cASM = New CAsmProc

    'vérifie que le disque est accessible
    With frmSanitization.Lang
        If cFile.IsDriveAvailable(Left$(sDisk, 1)) = False Then
            MsgBox .GetString("_DiskNotR"), vbCritical, .GetString("_War")
            Exit Sub
        End If
    End With
    
    'récupère les infos sue le disque
    Set cDriv = cFile.GetDrive(Left$(sDisk, 1))
    
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
        Call cASM.bnAlloc2MoAlea(Tbl(0))
        
        'calcul unique
        curOp = CCur(x * secPerString)
        
        'on écrit dans le disque
        '// &H55
        Call DirectWritePtHandle(hDevice, curOp, 2097152, bPerSec, p55)
        
        '// &HAA
        Call DirectWritePtHandle(hDevice, curOp, 2097152, bPerSec, pAA)
            
        '//random string
        Call DirectWritePtHandle(hDevice, curOp, 2097152, bPerSec, VarPtr(cASM(0)))
        
        'rend la main de tps en tps
        If (x Mod 5) = 0 Then
            PGB.Value = x
            DoEvents
        End If
        
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
    
    PGB.Value = PGB.Max
    
    'referme le handle
    Call CloseHandle(hDevice)
    
    'libère classes + affiche message
    Set cDriv = Nothing
    Set cASM = Nothing
    MsgBox frmSanitization.Lang.GetString("_SanitOk"), vbInformation, frmSanitization.Lang.GetString("_SanitOk")

End Sub
