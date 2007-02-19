Attribute VB_Name = "mdlFusionCut"
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
'//MODULE DE GESTION DE LA DECOUPE/FUSION
'=======================================================

Public lBufSize As Long     'taille du buffer

'=======================================================
'fonction de découpe de fichier
'=======================================================
Public Function CutFile(ByVal sFile As String, ByVal sFolderOut As String, tMethode As CUT_METHOD) As Long
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

    'On Error GoTo ErrGestion
    
    lTime = GetTickCount
    
    '//VERIFICATIONS
    'vérifie que le fichier existe bien
    If cFile.FileExists(sFile) = False Then
        'fichier manquant
        MsgBox "Le fichier ne peut être découpé car il n'existe pas.", vbCritical, "Erreur critique"
        Exit Function
    End If
    
    'vérifie que le dossier résultat existe bien
    If cFile.FolderExists(sFolderOut) = False Then
        'dossier résultat inexistant
        MsgBox "L'emplacement résultant n'existe pas, vous devez spécifier le fichier groupeur dans un dossier existant.", vbCritical, "Erreur critique"
        Exit Function
    End If
    
    'récupère le nom du fichier
    sFileStr = cFile.GetFileFromPath(sFile)
    
    'vérifie que le fichier groupeur n'existe pas déjà
    If cFile.FileExists(sFolderOut & "\" & sFileStr & ".grp") Then
        'fichier déjà existant
        If MsgBox("Le fichier existe déjà, le remplacer ?", vbInformation + vbYesNo, "Attention") <> vbYes Then Exit Function
    End If
    
    
    'récupère la taille du fichier
    curSize = cFile.GetFileSize(sFile)
    If curSize = 0 Or cFile.IsFileAvailable(sFile) = False Then
        'fichier vide ou inaccessible
        MsgBox "Le fichier est vide ou inaccessible, l'opération n'a pas pu être terminée.", vbCritical, "Erreur critique"
        Exit Function
    End If
    
    'règle la progressbar
    With frmCut.pgb
        .Min = 0
        .Value = 0
    End With
    
    '//LANCE LE DECOUPAGE
    If tMethode.tMethode = [Taille fixe] Then
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
        If tMethode.lParam <= lBufSize Then
            'alors tout rentre dans un seul buffer
            
            frmCut.pgb.Max = lFileCount
            
            k2 = lFileCount - 1
            k2 = k2 * tMethode.lParam
            
            For i = 1 To lFileCount - 1 'pas le dernier qui est fait à part
                
                'fichier résultat i
                sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(i))
                
                'créé le fichier résultat
                cFile.CreateEmptyFile sFic, True
                
                'récupère le buffer
                k = i - 1
                k = k * tMethode.lParam
                sBuf = GetBytesFromFile(sFile, CCur(tMethode.lParam), CCur(k))
                
                'on écrit dans le fichier résultat
                WriteBytesToFile sFic, sBuf, 0
                
                frmCut.pgb.Value = i
                DoEvents
            Next i
            
            'maintenant le dernier fichier
            sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(lFileCount))
            
            'créé le fichier résultat
            cFile.CreateEmptyFile sFic, True
            
            'récupère le buffer
            sBuf = GetBytesFromFile(sFile, lLastFileSize, CCur(k2))
            
            'on écrit dans le fichier résultat
            WriteBytesToFile sFic, sBuf, 0
            
            frmCut.pgb.Value = frmCut.pgb.Max

        Else
            'alors plusieurs buffer
            
            'calcule le nombre de buffers nécessaires pour chaque fichier
            lBuf2 = Int(tMethode.lParam / lBufSize) + IIf(Mod2(tMethode.lParam, lBufSize) = 0, 0, 1)
            
            frmCut.pgb.Max = lFileCount * lBuf2
            k2 = lFileCount - 1
            k2 = k2 * tMethode.lParam
            
            For i = 1 To lFileCount - 1 'pas le dernier qui est fait à part
                
                'fichier résultat i
                sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(i))
                
                'créé le fichier résultat
                cFile.CreateEmptyFile sFic, True
                
                k = i - 1
                k = k * tMethode.lParam
                
                For j = 1 To lBuf2 - 1
 
                    'récupère le buffer
                    sBuf = GetBytesFromFile(sFile, lBufSize, CCur((j - 1) * lBufSize) + k)
                    
                    'on écrit dans le fichier résultat
                    WriteBytesToFileEnd sFic, sBuf ', 5242880 * (j - 1)
                    
                    frmCut.pgb.Value = frmCut.pgb.Value + 1: DoEvents

                Next j
                
                'le dernier buffer
                sBuf = GetBytesFromFile(sFile, tMethode.lParam - (lBuf2 - 1) * lBufSize, CCur((lBuf2 - 1) * lBufSize) + k)

                'on écrit dans le fichier résultat
                WriteBytesToFileEnd sFic, sBuf ', 5242880 * (lBuf2 - 1)
                
                frmCut.pgb.Value = frmCut.pgb.Value + 1: DoEvents
            
            Next i

            'recalcule le nombre de buffers dans le dernier fichier
            lBuf2 = Int(lLastFileSize / lBufSize) + IIf(Mod2(lLastFileSize, lBufSize) = 0, 0, 1)

            'maintenant le dernier fichier
            sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(lFileCount))
            
            'créé le fichier résultat
            cFile.CreateEmptyFile sFic, True
            
            For j = 1 To lBuf2 - 1
                    
                'récupère le buffer
                sBuf = GetBytesFromFile(sFile, lBufSize, CCur(k2 + (j - 1) * lBufSize))
                
                'on écrit dans le fichier résultat
                WriteBytesToFileEnd sFic, sBuf ', 5242880 * (j - 1)
            
            Next j
                
            'récupère le dernier buffer
            sBuf = GetBytesFromFile(sFile, lLastFileSize - (lBuf2 - 1) * lBufSize, CCur(k2 + (lBuf2 - 1) * lBufSize))
            
            'on écrit dans le fichier résultat
            WriteBytesToFileEnd sFic, sBuf ', 0
            
            frmCut.pgb.Value = frmCut.pgb.Max
  
        End If
        
        
        'on créé le fichier groupeur
        cFile.CreateEmptyFile sFolderOut & "\" & sFileStr & ".grp", True
        cFile.SaveStringInfile sFolderOut & "\" & sFileStr & ".grp", sFileStr & "|" & Str$(lFileCount)
       
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
        If lNormalSize <= lBufSize Then
            'alors tout rentre dans un seul buffer
            
            frmCut.pgb.Max = lFileCount
            k2 = lFileCount - 1
            k2 = k2 * lNormalSize
            
            For i = 1 To lFileCount - 1 'pas le dernier qui est fait à part
                
                'fichier résultat i
                sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(i))
                
                'créé le fichier résultat
                cFile.CreateEmptyFile sFic, True
                
                'récupère le buffer
                k = i - 1
                k = k * lNormalSize
                sBuf = GetBytesFromFile(sFile, lNormalSize, CCur(k))
                
                'on écrit dans le fichier résultat
                WriteBytesToFile sFic, sBuf, 0
                
                frmCut.pgb.Value = i
                DoEvents
            Next i
            
            'maintenant le dernier fichier
            sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(lFileCount))
            
            'créé le fichier résultat
            cFile.CreateEmptyFile sFic, True
            
            'récupère le buffer
            sBuf = GetBytesFromFile(sFile, lLastFileSize, CCur(k2))
            
            'on écrit dans le fichier résultat
            WriteBytesToFile sFic, sBuf, 0
            
            frmCut.pgb.Value = frmCut.pgb.Max

        Else
            'alors plusieurs buffer
            
            'calcule le nombre de buffers nécessaires pour chaque fichier
            lBuf2 = Int(lNormalSize / lBufSize) + IIf(Mod2(lNormalSize, lBufSize) = 0, 0, 1)
            
            frmCut.pgb.Max = lFileCount * lBuf2
            k2 = lFileCount - 1
            k2 = k2 * lNormalSize
            
            For i = 1 To lFileCount - 1 'pas le dernier qui est fait à part
                
                'fichier résultat i
                sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(i))
                
                'créé le fichier résultat
                cFile.CreateEmptyFile sFic, True
                
                DoEvents
                
                k = i - 1
                k = k * lNormalSize
                
                For j = 1 To lBuf2 - 1
                
                    'récupère le buffer
                    sBuf = GetBytesFromFile(sFile, lBufSize, CCur((j - 1) * lBufSize) + k)
                    
                    'on écrit dans le fichier résultat
                    WriteBytesToFileEnd sFic, sBuf ', 5242880 * (j - 1)
                    
                    frmCut.pgb.Value = frmCut.pgb.Value + 1: DoEvents
                Next j

                'le dernier buffer
                sBuf = GetBytesFromFile(sFile, lNormalSize - (lBuf2 - 1) * lBufSize, CCur((lBuf2 - 1) * lBufSize) + k)

                'on écrit dans le fichier résultat
                WriteBytesToFileEnd sFic, sBuf ', 5242880 * (lBuf2 - 1)
                frmCut.pgb.Value = frmCut.pgb.Value + 1: DoEvents
            Next i

            'recalcule le nombre de buffers dans le dernier fichier
            lBuf2 = Int(lLastFileSize / lBufSize) + IIf(Mod2(lLastFileSize, lBufSize) = 0, 0, 1)

            'maintenant le dernier fichier
            sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(lFileCount))
            
            'créé le fichier résultat
            cFile.CreateEmptyFile sFic, True
            
            For j = 1 To lBuf2 - 1
            
                'récupère le buffer
                sBuf = GetBytesFromFile(sFile, lBufSize, CCur(k2 + (j - 1) * lBufSize))
                
                'on écrit dans le fichier résultat
                WriteBytesToFileEnd sFic, sBuf ', 5242880 * (j - 1)
            
            Next j
            
            'récupère le dernier buffer
            sBuf = GetBytesFromFile(sFile, lLastFileSize - (lBuf2 - 1) * lBufSize, CCur(k2 + (lBuf2 - 1) * lBufSize))
            
            'on écrit dans le fichier résultat
            WriteBytesToFileEnd sFic, sBuf ', 0
            
            frmCut.pgb.Value = frmCut.pgb.Max
        End If
        
        
        'on créé le fichier groupeur
        cFile.CreateEmptyFile sFolderOut & "\" & sFileStr & ".grp", True
        cFile.SaveStringInfile sFolderOut & "\" & sFileStr & ".grp", sFileStr & "|" & Str$(lFileCount)
 
    End If
    
    
    'terminé
    MsgBox "Découpage terminé avec succès.", vbInformation + vbOKOnly, "Découpage réussi"
    CutFile = GetTickCount - lTime
    Exit Function

ErrGestion:
End Function


'=======================================================
'fonction de fusion de fichier
'=======================================================
Public Function PasteFile(ByVal sFileGroup As String, ByVal sFolderOut As String) As Long
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

    'On Error GoTo ErrGestion
    
    lTime = GetTickCount
    
    '//VERIFICATIONS
    'vérifie que le fichier existe bien
    If cFile.FileExists(sFileGroup) = False Then
        'fichier manquant
        MsgBox "Le fichier ne peut être créé car le fichier de fusion n'existe pas.", vbCritical, "Erreur critique"
        Exit Function
    End If
    
    'vérifie que le dossier résultat existe bien
    If cFile.FolderExists(sFolderOut) = False Then
        'dossier résultat inexistant
        MsgBox "L'emplacement résultant n'existe pas, vous devez spécifier le fichier créé dans un dossier existant.", vbCritical, "Erreur critique"
        Exit Function
    End If
    
    sBuf = cFile.LoadFileInString(sFileGroup)
    'récupère le nom du fichier
    sFileStr = Mid$(sBuf, 1, InStr(1, sBuf, "|") - 1)
    
    'vérifie que le fichier groupeur n'existe pas déjà
    If cFile.FileExists(sFolderOut & "\" & sFileStr) Then
        'fichier déjà existant
        If MsgBox("Le fichier existe déjà, le remplacer ?", vbInformation + vbYesNo, "Attention") <> vbYes Then Exit Function
    End If
    
    If cFile.IsFileAvailable(sFileGroup) = False Then
        'fichier groupe indisponible ou inexistant
        MsgBox "Le fichier groupeur est indisponible ou inexistant.", vbCritical, "Erreur critique"
    End If

    With frmCut.pgb
        .Min = 0
        .Value = 0
    End With

    '//LANCE LA FUSION
    'récupère le nombre de fichiers concernés
    lFileCount = Val(Right$(sBuf, Len(sBuf) - InStr(1, sBuf, "|")))
    
    'vérifie l'existence de chaque fichier
    bOk = True
    For i = 1 To lFileCount
        If cFile.FileExists(cFile.GetFolderFromPath(sFileGroup) & "\" & sFileStr & "." & Trim$(Str$(i))) = False Then
            bOk = False
        End If
    Next i
    If Not (bOk) Then
        'alors un fichier est absent
        MsgBox "Il manque un fichier.", vbCritical, "Opération de fusion impossible"
        Exit Function
    End If
    
    'créé le fichier résultat
    cFile.CreateEmptyFile sFolderOut & "\" & sFileStr, True
    
    'alors tout est OK, on peut commencer à coller les données par buffer de 5Mo
    If cFile.GetFileSize(sFolderOut & "\" & sFileStr & ".1") <= lBufSize Then
        'alors tout rentre dans un buffer de 5Mo
    
        frmCut.pgb.Max = lFileCount
        frmCut.pgb.Value = 0
        For i = 1 To lFileCount
            'écrit les bytes lus
            WriteBytesToFileEnd sFolderOut & "\" & sFileStr, cFile.LoadFileInString(cFile.GetFolderFromPath(sFileGroup) & "\" & sFileStr & "." & Trim$(Str$(i)))
            DoEvents: frmCut.pgb.Value = frmCut.pgb.Value + 1
        Next i
        frmCut.pgb.Value = frmCut.pgb.Max
        
    Else
    
        'alors il faut plusieurs buffers de 5Mo par fichier
        
        'détermine le nombre de buffers nécessaire
        lBuf2 = Int(cFile.GetFileSize(sFolderOut & "\" & sFileStr & ".1") / lBufSize) + IIf(Mod2(cFile.GetFileSize(sFolderOut & "\" & sFileStr & ".1"), lBufSize) = 0, 0, 1)
        
        frmCut.pgb.Max = lFileCount * lBuf2
        
        For i = 1 To lFileCount - 1
        
            'le fichier que l'on lit
            sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(i))
            
            For j = 1 To lBuf2 - 1
                'sbuf contient 5Mo lus
                sBuf = GetBytesFromFile(sFic, lBufSize, lBufSize * (j - 1))
                
                'écrit les bytes dans le fichier résultat
                WriteBytesToFileEnd sFolderOut & "\" & sFileStr, sBuf
                
                frmCut.pgb.Value = frmCut.pgb.Value + 1: DoEvents
            Next j
            
            'le dernier buffer
            a = cFile.GetFileSize(sFic) - (lBuf2 - 1) * lBufSize     'taille du dernier buffer
            sBuf = GetBytesFromFile(sFic, a, lBufSize * (lBuf2 - 1))
            
            'écrit les bytes dans le fichier résultat
            WriteBytesToFileEnd sFolderOut & "\" & sFileStr, sBuf
            
            frmCut.pgb.Value = frmCut.pgb.Value + 1
            DoEvents
        Next i
        
        'fait le dernier fichier
        sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(lFileCount))
        lBuf2 = Int(cFile.GetFileSize(sFic) / lBufSize) + IIf(Mod2(cFile.GetFileSize(sFic), lBufSize) = 0, 0, 1)    'nouveau buffer
            
        For j = 1 To lBuf2 - 1
            'sbuf contient 5Mo lus
            sBuf = GetBytesFromFile(sFic, lBufSize, lBufSize * (j - 1))
            
            'écrit les bytes dans le fichier résultat
            WriteBytesToFileEnd sFolderOut & "\" & sFileStr, sBuf
        Next j
        
        'le dernier buffer
        a = cFile.GetFileSize(sFic) - (lBuf2 - 1) * lBufSize
        sBuf = GetBytesFromFile(sFic, a, lBufSize * (lBuf2 - 1))
        
        'écrit les bytes dans le fichier résultat
        WriteBytesToFileEnd sFolderOut & "\" & sFileStr, sBuf
        
        frmCut.pgb.Value = frmCut.pgb.Max
        DoEvents
        
    End If
    
    'terminé
    MsgBox "Fusion terminée avec succès.", vbInformation + vbOKOnly, "Fusion réussie"
    
    PasteFile = GetTickCount - lTime
    Exit Function

ErrGestion:
End Function

'=======================================================
'effectue un modulo sans dépassement de capacité
'très peu optimisé, mais utile pour les grandes valeurs de cur
'=======================================================
Public Function Mod2(ByVal cur As Currency, lng As Long) As Currency
    Mod2 = cur - Int(cur / lng) * lng
End Function

