Attribute VB_Name = "mdlFusionCut"
' -----------------------------------------------
'
' Hex Editor VB
' Coded by violent_ken (Alain Descotes)
'
' -----------------------------------------------
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
'fonction de d�coupe de fichier
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
    'v�rifie que le fichier existe bien
    If cFile.FileExists(sFile) = False Then
        'fichier manquant
        MsgBox "Le fichier ne peut �tre d�coup� car il n'existe pas.", vbCritical, "Erreur critique"
        Exit Sub
    End If
    
    'v�rifie que le dossier r�sultat existe bien
    If cFile.FolderExists(sFolderOut) = False Then
        'dossier r�sultat inexistant
        MsgBox "L'emplacement r�sultant n'existe pas, vous devez sp�cifier le fichier groupeur dans un dossier existant.", vbCritical, "Erreur critique"
        Exit Sub
    End If
    
    'r�cup�re le nom du fichier
    sFileStr = cFile.GetFileFromPath(sFile)
    
    'v�rifie que le fichier groupeur n'existe pas d�j�
    If cFile.FileExists(sFolderOut & "\" & sFileStr & ".grp") Then
        'fichier d�j� existant
        If MsgBox("Le fichier existe d�j�, le rempalcer ?", vbInformation + vbYesNo, "Attention") <> vbYes Then Exit Sub
    End If
    
    
    'r�cup�re la taille du fichier
    curSize = cFile.GetFileSize(sFile)
    If curSize = 0 Then
        'fichier vide ou inaccessible
        MsgBox "Le fichier est vide ou inaccessible, l'op�ration n'a pas pu �tre termin�e.", vbCritical, "Erreur critique"
        Exit Sub
    End If

    
    '//LANCE LE DECOUPAGE
    If tMethode.tMethode = [Taille fixe] Then
        'alors on d�coupe en fixant la taille
        
        'calcul le nombre de fichiers n�cessaire et la taille du dernier
        lFileCount = Int(curSize / tMethode.lParam) + IIf((curSize Mod tMethode.lParam) = 0, 0, 1)
        lLastFileSize = curSize - (lFileCount - 1) * tMethode.lParam 'taille du dernier fichier
        
        'lance la d�coupe
        'utilisation de l'API ReadFile pour plus d'efficacit�
        'prend des buffers de 1Mo maximum
        If tMethode.lParam <= 1048576 Then
            'alors tout rentre dans un seul buffer
            
            For i = 1 To lFileCount - 1 'pas le dernier qui est fait � part
                
                'fichier r�sultat i
                sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(i))
                
                'cr�� le fichier r�sultat
                cFile.CreateEmptyFile sFic, True
                
                'r�cup�re le buffer
                sBuf = GetBytesFromFile(sFile, CCur(tMethode.lParam), CCur((i - 1) * tMethode.lParam))
                
                'on �crit dans le fichier r�sultat
                WriteBytesToFile sFic, sBuf, 0
                
            Next i
            
            'maintenant le dernier fichier
            sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(lFileCount))
            
            'cr�� le fichier r�sultat
            cFile.CreateEmptyFile sFic, True
            
            'r�cup�re le buffer
            sBuf = GetBytesFromFile(sFile, lLastFileSize, CCur((lFileCount - 1) * tMethode.lParam))
            
            'on �crit dans le fichier r�sultat
            WriteBytesToFile sFic, sBuf, 0

        Else
            'alors plusieurs buffer
            
            
            
            
        End If
        
        
        'on cr�� le fichier groupeur
       
    Else
        'alors nombre de fichiers fix�
        
        
    End If
    
    
    'termin�
    MsgBox "D�coupage termin� avec succ�s.", vbInformation + vbOKOnly, "D�coupage r�ussi"
        
End Sub
