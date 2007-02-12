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
Dim lBuf As Long
Dim sFileStr As String
Dim sBuf As String
Dim sFic As String

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
        If MsgBox("Le fichier existe déjà, le rempalcer ?", vbInformation + vbYesNo, "Attention") <> vbYes Then Exit Sub
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
        'prend des buffers de 1Mo maximum
        If tMethode.lParam <= 1048576 Then
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
            
            
            
            
        End If
        
        
        'on créé le fichier groupeur
       
    Else
        'alors nombre de fichiers fixé
        
        
    End If
    
    
    'terminé
    MsgBox "Découpage terminé avec succès.", vbInformation + vbOKOnly, "Découpage réussi"
        
End Sub
