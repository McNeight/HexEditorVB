Attribute VB_Name = "mdlFusionCut"
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
'//MODULE DE GESTION DE LA DECOUPE/FUSION
'-------------------------------------------------------

'-------------------------------------------------------
'TYPES & ENUMS
'-------------------------------------------------------
Public Enum CUT_METHOD_ENUM
    [Taille fixe]
    [Nombre fichiers fixe]
End Enum
Public Type CUT_METHOD
    tMethode As CUT_METHOD_ENUM
    lParam As Long
End Type

'-------------------------------------------------------
'fonction de découpe de fichier
'-------------------------------------------------------
Public Sub CutFile(ByVal sFile As String, ByVal sFolderOut As String, tMethode As CUT_METHOD)
Dim lFileCount As Long
Dim lLastFileSize As Long
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

    On Error GoTo ErrGestion
    
    
    '//VERIFICATIONS
    'vérifie que le fichier existe bien
    If cFile.FileExists(sFile) = False Then
        'fichier manquant
        MsgBox "Le fichier ne peut être découpé car il n'existe pas.", vbCritical, "Erreur critique"
        Exit Sub
    End If
    
    'vérifie que le dossier résultat existe bien
    If cFile.FolderExists(sFolderOut) = False Then
        'dossier résultat inexistant
        MsgBox "L'emplacement résultant n'existe pas, vous devez spécifier le fichier groupeur dans un dossier existant.", vbCritical, "Erreur critique"
        Exit Sub
    End If
    
    'récupère le nom du fichier
    sFileStr = cFile.GetFileFromPath(sFile)
    
    'vérifie que le fichier groupeur n'existe pas déjà
    If cFile.FileExists(sFolderOut & "\" & sFileStr & ".grp") Then
        'fichier déjà existant
        If MsgBox("Le fichier existe déjà, le remplacer ?", vbInformation + vbYesNo, "Attention") <> vbYes Then Exit Sub
    End If
    
    
    'récupère la taille du fichier
    curSize = cFile.GetFileSize(sFile)
    If curSize = 0 Then
        'fichier vide ou inaccessible
        MsgBox "Le fichier est vide ou inaccessible, l'opération n'a pas pu être terminée.", vbCritical, "Erreur critique"
        Exit Sub
    End If

    
    '//LANCE LE DECOUPAGE
    If tMethode.tMethode = [Taille fixe] Then
        'alors on découpe en fixant la taille
        
        'calcul le nombre de fichiers nécessaire et la taille du dernier
        lFileCount = Int(curSize / tMethode.lParam) + IIf((curSize Mod tMethode.lParam) = 0, 0, 1)
        lLastFileSize = curSize - (lFileCount - 1) * tMethode.lParam 'taille du dernier fichier
        
        'lance la découpe
        'utilisation de l'API ReadFile pour plus d'efficacité
        'prend des buffers de 5Mo maximum
        If tMethode.lParam <= 5242880 Then
            'alors tout rentre dans un seul buffer
            
            For i = 1 To lFileCount - 1 'pas le dernier qui est fait à part
                
                'fichier résultat i
                sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(i))
                
                'créé le fichier résultat
                cFile.CreateEmptyFile sFic, True
                
                'récupère le buffer
                sBuf = GetBytesFromFile(sFile, CCur(tMethode.lParam), CCur((i - 1) * tMethode.lParam))
                
                'on écrit dans le fichier résultat
                WriteBytesToFile sFic, sBuf, 0
                
                DoEvents
            Next i
            
            'maintenant le dernier fichier
            sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(lFileCount))
            
            'créé le fichier résultat
            cFile.CreateEmptyFile sFic, True
            
            'récupère le buffer
            sBuf = GetBytesFromFile(sFile, lLastFileSize, CCur((lFileCount - 1) * tMethode.lParam))
            
            'on écrit dans le fichier résultat
            WriteBytesToFile sFic, sBuf, 0

        Else
            'alors plusieurs buffer
            
            'calcule le nombre de buffers nécessaires pour chaque fichier
            lBuf2 = Int(tMethode.lParam / 5242880) + IIf((tMethode.lParam Mod 5242880) = 0, 0, 1)
            
            For i = 1 To lFileCount - 1 'pas le dernier qui est fait à part
                
                'fichier résultat i
                sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(i))
                
                'créé le fichier résultat
                cFile.CreateEmptyFile sFic, True
                
                DoEvents
                
                For j = 1 To lBuf2 - 1
                
                    'récupère le buffer
                    sBuf = GetBytesFromFile(sFile, 5242880, CCur((j - 1) * 5242880))
                    
                    'on écrit dans le fichier résultat
                    WriteBytesToFileEnd sFic, sBuf ', 5242880 * (j - 1)

                Next j

                'le dernier buffer
                '==> CETTE ligne foire (espion InStr(1, sbuf, "àpäR") ==> indique que pique dans la fin du fichier
                'pour mettre dans le dernier buffer du premier fichier)
                sBuf = GetBytesFromFile(sFile, tMethode.lParam - (lBuf2 - 1) * 5242880, CCur((lBuf2 - 1) * 5242880))

                'on écrit dans le fichier résultat
                WriteBytesToFileEnd sFic, sBuf ', 5242880 * (lBuf2 - 1)
                
                
            '//BUG ==> a la fin du fichier 1 (sur 2), on a la fin du fichier total
            '==> prend la fin du fichier total dans le dernier buffer du fichier 1
            'A priori, on écrit à la fin du fichier à chaque fois, donc pas de bug
            'de Write au mauvais endroit, mais plutôt de Read au mauvais endroit.
            'A priori, le dernier fichier est bon
            
            
            Next i

            'recalcule le nombre de buffers dans le dernier fichier
            lBuf2 = Int(lLastFileSize / 5242880) + IIf((lLastFileSize Mod 5242880) = 0, 0, 1)

            'maintenant le dernier fichier
            sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(lFileCount))
            
            'créé le fichier résultat
            cFile.CreateEmptyFile sFic, True
            
            For j = 1 To lBuf2 - 1
            
                'récupère le buffer
                sBuf = GetBytesFromFile(sFile, 5242880, CCur((lFileCount - 1) * tMethode.lParam + j * 5242880))
                
                'on écrit dans le fichier résultat
                WriteBytesToFileEnd sFic, sBuf ', 5242880 * (j - 1)
            
            Next j
            
            'récupère le dernier buffer
            sBuf = GetBytesFromFile(sFile, lLastFileSize, CCur((lFileCount - 1) * tMethode.lParam))
            
            'on écrit dans le fichier résultat
            WriteBytesToFileEnd sFic, sBuf ', 0
  
        End If
        
        
        'on créé le fichier groupeur
        cFile.CreateEmptyFile sFolderOut & "\" & sFileStr & ".grp", True
        cFile.SaveStringInfile sFolderOut & "\" & sFileStr & ".grp", sFileStr & "|" & Str$(lFileCount)
       
    Else
        'alors nombre de fichiers fixé
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
    End If
    
    
    'terminé
    MsgBox "Découpage terminé avec succès.", vbInformation + vbOKOnly, "Découpage réussi"
    Exit Sub

ErrGestion:
    clsERREUR.AddError "mdlFusionCut.CutFile", True
End Sub


'-------------------------------------------------------
'fonction de fusion de fichier
'-------------------------------------------------------
Public Sub PasteFile(ByVal sFileGroup As String, ByVal sFolderOut As String)
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

    On Error GoTo ErrGestion
    
    
    '//VERIFICATIONS
    'vérifie que le fichier existe bien
    If cFile.FileExists(sFileGroup) = False Then
        'fichier manquant
        MsgBox "Le fichier ne peut être créé car le fichier de fusion n'existe pas.", vbCritical, "Erreur critique"
        Exit Sub
    End If
    
    'vérifie que le dossier résultat existe bien
    If cFile.FolderExists(sFolderOut) = False Then
        'dossier résultat inexistant
        MsgBox "L'emplacement résultant n'existe pas, vous devez spécifier le fichier créé dans un dossier existant.", vbCritical, "Erreur critique"
        Exit Sub
    End If
    
    sBuf = cFile.LoadFileInString(sFileGroup)
    'récupère le nom du fichier
    sFileStr = Mid$(sBuf, 1, InStr(1, sBuf, "|") - 1)
    
    'vérifie que le fichier groupeur n'existe pas déjà
    If cFile.FileExists(sFolderOut & "\" & sFileStr) Then
        'fichier déjà existant
        If MsgBox("Le fichier existe déjà, le remplacer ?", vbInformation + vbYesNo, "Attention") <> vbYes Then Exit Sub
    End If


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
        Exit Sub
    End If
    
    'créé le fichier résultat
    cFile.CreateEmptyFile sFolderOut & "\" & sFileStr, True
    
    'alors tout est OK, on peut commencer à coller les données par buffer de 5Mo
    If cFile.GetFileSize(sFolderOut & "\" & sFileStr & ".1") <= 5242880 Then
        'alors tout rentre dans un buffer de 5Mo
    
        For i = 1 To lFileCount
            'écrit les bytes lus
            WriteBytesToFileEnd sFolderOut & "\" & sFileStr, cFile.LoadFileInString(sFolderOut & "\" & sFileStr & "." & Trim$(Str$(i)))
            DoEvents
        Next i
        
    Else
    
        'alors il faut plusieurs buffers de 5Mo par fichier
        
        'détermine le nombre de buffers nécessaire
        lBuf2 = Int(cFile.GetFileSize(sFolderOut & "\" & sFileStr & ".1") / 5242880) + IIf((cFile.GetFileSize(sFolderOut & "\" & sFileStr & ".1") Mod 5242880) = 0, 0, 1)
        
        For i = 1 To lFileCount - 1
        
            'le fichier que l'on lit
            sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(i))
            
            For j = 1 To lBuf2 - 1
                'sbuf contient 5Mo lus
                sBuf = GetBytesFromFile(sFic, 5242880, 5242880 * (j - 1))
                
                'écrit les bytes dans le fichier résultat
                WriteBytesToFileEnd sFolderOut & "\" & sFileStr, sBuf
            Next j
            
            'le dernier buffer
            a = cFile.GetFileSize(sFic) - (lBuf2 - 1) * 5242880     'taille du dernier buffer
            sBuf = GetBytesFromFile(sFic, a, 5242880 * (lBuf2 - 1))
            
            'écrit les bytes dans le fichier résultat
            WriteBytesToFileEnd sFolderOut & "\" & sFileStr, sBuf
            
            DoEvents
        Next i
        
        'fait le dernier fichier
        sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(lFileCount))
        lBuf2 = Int(cFile.GetFileSize(sFic) / 5242880) + IIf((cFile.GetFileSize(sFic) Mod 5242880) = 0, 0, 1)   'nouveau buffer
            
        For j = 1 To lBuf2 - 1
            'sbuf contient 5Mo lus
            sBuf = GetBytesFromFile(sFic, 5242880, 5242880 * (j - 1))
            
            'écrit les bytes dans le fichier résultat
            WriteBytesToFileEnd sFolderOut & "\" & sFileStr, sBuf
        Next j
        
        'le dernier buffer
        a = cFile.GetFileSize(sFic) - (lBuf2 - 1) * 5242880
        sBuf = GetBytesFromFile(sFic, a, 5242880 * (lBuf2 - 1))
        
        'écrit les bytes dans le fichier résultat
        WriteBytesToFileEnd sFolderOut & "\" & sFileStr, sBuf
        
        DoEvents
        
    End If
    
    'terminé
    MsgBox "Fusion terminée avec succès.", vbInformation + vbOKOnly, "Fusion réussie"
   
    Exit Sub

ErrGestion:
    clsERREUR.AddError "mdlFusionCut.PasteFile", True
End Sub

