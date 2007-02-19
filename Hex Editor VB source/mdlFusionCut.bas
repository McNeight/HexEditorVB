Attribute VB_Name = "mdlFusionCut"
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
'//MODULE DE GESTION DE LA DECOUPE/FUSION
'=======================================================

Public lBufSize As Long     'taille du buffer

'=======================================================
'fonction de d�coupe de fichier
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
    'v�rifie que le fichier existe bien
    If cFile.FileExists(sFile) = False Then
        'fichier manquant
        MsgBox "Le fichier ne peut �tre d�coup� car il n'existe pas.", vbCritical, "Erreur critique"
        Exit Function
    End If
    
    'v�rifie que le dossier r�sultat existe bien
    If cFile.FolderExists(sFolderOut) = False Then
        'dossier r�sultat inexistant
        MsgBox "L'emplacement r�sultant n'existe pas, vous devez sp�cifier le fichier groupeur dans un dossier existant.", vbCritical, "Erreur critique"
        Exit Function
    End If
    
    'r�cup�re le nom du fichier
    sFileStr = cFile.GetFileFromPath(sFile)
    
    'v�rifie que le fichier groupeur n'existe pas d�j�
    If cFile.FileExists(sFolderOut & "\" & sFileStr & ".grp") Then
        'fichier d�j� existant
        If MsgBox("Le fichier existe d�j�, le remplacer ?", vbInformation + vbYesNo, "Attention") <> vbYes Then Exit Function
    End If
    
    
    'r�cup�re la taille du fichier
    curSize = cFile.GetFileSize(sFile)
    If curSize = 0 Or cFile.IsFileAvailable(sFile) = False Then
        'fichier vide ou inaccessible
        MsgBox "Le fichier est vide ou inaccessible, l'op�ration n'a pas pu �tre termin�e.", vbCritical, "Erreur critique"
        Exit Function
    End If
    
    'r�gle la progressbar
    With frmCut.pgb
        .Min = 0
        .Value = 0
    End With
    
    '//LANCE LE DECOUPAGE
    If tMethode.tMethode = [Taille fixe] Then
        'alors on d�coupe en fixant la taille
        
        'calcule le nombre de fichiers n�cessaire et la taille du dernier
        lFileCount = Int(curSize / tMethode.lParam) + IIf(Mod2(curSize, tMethode.lParam) = 0, 0, 1)
        lLastFileSize = curSize     'taille du dernier fichier
        k = (lFileCount - 1)
        k = k * tMethode.lParam
        lLastFileSize = lLastFileSize - k '�vite les d�passement de capacit�
        
        
        'lance la d�coupe
        'utilisation de l'API ReadFile pour plus d'efficacit�
        'prend des buffers de 5Mo maximum
        If tMethode.lParam <= lBufSize Then
            'alors tout rentre dans un seul buffer
            
            frmCut.pgb.Max = lFileCount
            
            k2 = lFileCount - 1
            k2 = k2 * tMethode.lParam
            
            For i = 1 To lFileCount - 1 'pas le dernier qui est fait � part
                
                'fichier r�sultat i
                sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(i))
                
                'cr�� le fichier r�sultat
                cFile.CreateEmptyFile sFic, True
                
                'r�cup�re le buffer
                k = i - 1
                k = k * tMethode.lParam
                sBuf = GetBytesFromFile(sFile, CCur(tMethode.lParam), CCur(k))
                
                'on �crit dans le fichier r�sultat
                WriteBytesToFile sFic, sBuf, 0
                
                frmCut.pgb.Value = i
                DoEvents
            Next i
            
            'maintenant le dernier fichier
            sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(lFileCount))
            
            'cr�� le fichier r�sultat
            cFile.CreateEmptyFile sFic, True
            
            'r�cup�re le buffer
            sBuf = GetBytesFromFile(sFile, lLastFileSize, CCur(k2))
            
            'on �crit dans le fichier r�sultat
            WriteBytesToFile sFic, sBuf, 0
            
            frmCut.pgb.Value = frmCut.pgb.Max

        Else
            'alors plusieurs buffer
            
            'calcule le nombre de buffers n�cessaires pour chaque fichier
            lBuf2 = Int(tMethode.lParam / lBufSize) + IIf(Mod2(tMethode.lParam, lBufSize) = 0, 0, 1)
            
            frmCut.pgb.Max = lFileCount * lBuf2
            k2 = lFileCount - 1
            k2 = k2 * tMethode.lParam
            
            For i = 1 To lFileCount - 1 'pas le dernier qui est fait � part
                
                'fichier r�sultat i
                sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(i))
                
                'cr�� le fichier r�sultat
                cFile.CreateEmptyFile sFic, True
                
                k = i - 1
                k = k * tMethode.lParam
                
                For j = 1 To lBuf2 - 1
 
                    'r�cup�re le buffer
                    sBuf = GetBytesFromFile(sFile, lBufSize, CCur((j - 1) * lBufSize) + k)
                    
                    'on �crit dans le fichier r�sultat
                    WriteBytesToFileEnd sFic, sBuf ', 5242880 * (j - 1)
                    
                    frmCut.pgb.Value = frmCut.pgb.Value + 1: DoEvents

                Next j
                
                'le dernier buffer
                sBuf = GetBytesFromFile(sFile, tMethode.lParam - (lBuf2 - 1) * lBufSize, CCur((lBuf2 - 1) * lBufSize) + k)

                'on �crit dans le fichier r�sultat
                WriteBytesToFileEnd sFic, sBuf ', 5242880 * (lBuf2 - 1)
                
                frmCut.pgb.Value = frmCut.pgb.Value + 1: DoEvents
            
            Next i

            'recalcule le nombre de buffers dans le dernier fichier
            lBuf2 = Int(lLastFileSize / lBufSize) + IIf(Mod2(lLastFileSize, lBufSize) = 0, 0, 1)

            'maintenant le dernier fichier
            sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(lFileCount))
            
            'cr�� le fichier r�sultat
            cFile.CreateEmptyFile sFic, True
            
            For j = 1 To lBuf2 - 1
                    
                'r�cup�re le buffer
                sBuf = GetBytesFromFile(sFile, lBufSize, CCur(k2 + (j - 1) * lBufSize))
                
                'on �crit dans le fichier r�sultat
                WriteBytesToFileEnd sFic, sBuf ', 5242880 * (j - 1)
            
            Next j
                
            'r�cup�re le dernier buffer
            sBuf = GetBytesFromFile(sFile, lLastFileSize - (lBuf2 - 1) * lBufSize, CCur(k2 + (lBuf2 - 1) * lBufSize))
            
            'on �crit dans le fichier r�sultat
            WriteBytesToFileEnd sFic, sBuf ', 0
            
            frmCut.pgb.Value = frmCut.pgb.Max
  
        End If
        
        
        'on cr�� le fichier groupeur
        cFile.CreateEmptyFile sFolderOut & "\" & sFileStr & ".grp", True
        cFile.SaveStringInfile sFolderOut & "\" & sFileStr & ".grp", sFileStr & "|" & Str$(lFileCount)
       
    Else
        'alors nombre de fichiers fix�

        'nombre de fichiers
        lFileCount = tMethode.lParam
        
        'calcule la taille de chaque fichier
        lNormalSize = Int(curSize / lFileCount)
        lLastFileSize = lNormalSize + curSize
        k = lNormalSize
        k = k * lFileCount 'taille du dernier fichier (plus quelques octets, au maximum 1 par fichier)
        lLastFileSize = lLastFileSize - k       '�vite le d�passement de capacit�
        
        
        'lance la d�coupe
        'utilisation de l'API ReadFile pour plus d'efficacit�
        'prend des buffers de 5Mo maximum
        If lNormalSize <= lBufSize Then
            'alors tout rentre dans un seul buffer
            
            frmCut.pgb.Max = lFileCount
            k2 = lFileCount - 1
            k2 = k2 * lNormalSize
            
            For i = 1 To lFileCount - 1 'pas le dernier qui est fait � part
                
                'fichier r�sultat i
                sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(i))
                
                'cr�� le fichier r�sultat
                cFile.CreateEmptyFile sFic, True
                
                'r�cup�re le buffer
                k = i - 1
                k = k * lNormalSize
                sBuf = GetBytesFromFile(sFile, lNormalSize, CCur(k))
                
                'on �crit dans le fichier r�sultat
                WriteBytesToFile sFic, sBuf, 0
                
                frmCut.pgb.Value = i
                DoEvents
            Next i
            
            'maintenant le dernier fichier
            sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(lFileCount))
            
            'cr�� le fichier r�sultat
            cFile.CreateEmptyFile sFic, True
            
            'r�cup�re le buffer
            sBuf = GetBytesFromFile(sFile, lLastFileSize, CCur(k2))
            
            'on �crit dans le fichier r�sultat
            WriteBytesToFile sFic, sBuf, 0
            
            frmCut.pgb.Value = frmCut.pgb.Max

        Else
            'alors plusieurs buffer
            
            'calcule le nombre de buffers n�cessaires pour chaque fichier
            lBuf2 = Int(lNormalSize / lBufSize) + IIf(Mod2(lNormalSize, lBufSize) = 0, 0, 1)
            
            frmCut.pgb.Max = lFileCount * lBuf2
            k2 = lFileCount - 1
            k2 = k2 * lNormalSize
            
            For i = 1 To lFileCount - 1 'pas le dernier qui est fait � part
                
                'fichier r�sultat i
                sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(i))
                
                'cr�� le fichier r�sultat
                cFile.CreateEmptyFile sFic, True
                
                DoEvents
                
                k = i - 1
                k = k * lNormalSize
                
                For j = 1 To lBuf2 - 1
                
                    'r�cup�re le buffer
                    sBuf = GetBytesFromFile(sFile, lBufSize, CCur((j - 1) * lBufSize) + k)
                    
                    'on �crit dans le fichier r�sultat
                    WriteBytesToFileEnd sFic, sBuf ', 5242880 * (j - 1)
                    
                    frmCut.pgb.Value = frmCut.pgb.Value + 1: DoEvents
                Next j

                'le dernier buffer
                sBuf = GetBytesFromFile(sFile, lNormalSize - (lBuf2 - 1) * lBufSize, CCur((lBuf2 - 1) * lBufSize) + k)

                'on �crit dans le fichier r�sultat
                WriteBytesToFileEnd sFic, sBuf ', 5242880 * (lBuf2 - 1)
                frmCut.pgb.Value = frmCut.pgb.Value + 1: DoEvents
            Next i

            'recalcule le nombre de buffers dans le dernier fichier
            lBuf2 = Int(lLastFileSize / lBufSize) + IIf(Mod2(lLastFileSize, lBufSize) = 0, 0, 1)

            'maintenant le dernier fichier
            sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(lFileCount))
            
            'cr�� le fichier r�sultat
            cFile.CreateEmptyFile sFic, True
            
            For j = 1 To lBuf2 - 1
            
                'r�cup�re le buffer
                sBuf = GetBytesFromFile(sFile, lBufSize, CCur(k2 + (j - 1) * lBufSize))
                
                'on �crit dans le fichier r�sultat
                WriteBytesToFileEnd sFic, sBuf ', 5242880 * (j - 1)
            
            Next j
            
            'r�cup�re le dernier buffer
            sBuf = GetBytesFromFile(sFile, lLastFileSize - (lBuf2 - 1) * lBufSize, CCur(k2 + (lBuf2 - 1) * lBufSize))
            
            'on �crit dans le fichier r�sultat
            WriteBytesToFileEnd sFic, sBuf ', 0
            
            frmCut.pgb.Value = frmCut.pgb.Max
        End If
        
        
        'on cr�� le fichier groupeur
        cFile.CreateEmptyFile sFolderOut & "\" & sFileStr & ".grp", True
        cFile.SaveStringInfile sFolderOut & "\" & sFileStr & ".grp", sFileStr & "|" & Str$(lFileCount)
 
    End If
    
    
    'termin�
    MsgBox "D�coupage termin� avec succ�s.", vbInformation + vbOKOnly, "D�coupage r�ussi"
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
    'v�rifie que le fichier existe bien
    If cFile.FileExists(sFileGroup) = False Then
        'fichier manquant
        MsgBox "Le fichier ne peut �tre cr�� car le fichier de fusion n'existe pas.", vbCritical, "Erreur critique"
        Exit Function
    End If
    
    'v�rifie que le dossier r�sultat existe bien
    If cFile.FolderExists(sFolderOut) = False Then
        'dossier r�sultat inexistant
        MsgBox "L'emplacement r�sultant n'existe pas, vous devez sp�cifier le fichier cr�� dans un dossier existant.", vbCritical, "Erreur critique"
        Exit Function
    End If
    
    sBuf = cFile.LoadFileInString(sFileGroup)
    'r�cup�re le nom du fichier
    sFileStr = Mid$(sBuf, 1, InStr(1, sBuf, "|") - 1)
    
    'v�rifie que le fichier groupeur n'existe pas d�j�
    If cFile.FileExists(sFolderOut & "\" & sFileStr) Then
        'fichier d�j� existant
        If MsgBox("Le fichier existe d�j�, le remplacer ?", vbInformation + vbYesNo, "Attention") <> vbYes Then Exit Function
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
    'r�cup�re le nombre de fichiers concern�s
    lFileCount = Val(Right$(sBuf, Len(sBuf) - InStr(1, sBuf, "|")))
    
    'v�rifie l'existence de chaque fichier
    bOk = True
    For i = 1 To lFileCount
        If cFile.FileExists(cFile.GetFolderFromPath(sFileGroup) & "\" & sFileStr & "." & Trim$(Str$(i))) = False Then
            bOk = False
        End If
    Next i
    If Not (bOk) Then
        'alors un fichier est absent
        MsgBox "Il manque un fichier.", vbCritical, "Op�ration de fusion impossible"
        Exit Function
    End If
    
    'cr�� le fichier r�sultat
    cFile.CreateEmptyFile sFolderOut & "\" & sFileStr, True
    
    'alors tout est OK, on peut commencer � coller les donn�es par buffer de 5Mo
    If cFile.GetFileSize(sFolderOut & "\" & sFileStr & ".1") <= lBufSize Then
        'alors tout rentre dans un buffer de 5Mo
    
        frmCut.pgb.Max = lFileCount
        frmCut.pgb.Value = 0
        For i = 1 To lFileCount
            '�crit les bytes lus
            WriteBytesToFileEnd sFolderOut & "\" & sFileStr, cFile.LoadFileInString(cFile.GetFolderFromPath(sFileGroup) & "\" & sFileStr & "." & Trim$(Str$(i)))
            DoEvents: frmCut.pgb.Value = frmCut.pgb.Value + 1
        Next i
        frmCut.pgb.Value = frmCut.pgb.Max
        
    Else
    
        'alors il faut plusieurs buffers de 5Mo par fichier
        
        'd�termine le nombre de buffers n�cessaire
        lBuf2 = Int(cFile.GetFileSize(sFolderOut & "\" & sFileStr & ".1") / lBufSize) + IIf(Mod2(cFile.GetFileSize(sFolderOut & "\" & sFileStr & ".1"), lBufSize) = 0, 0, 1)
        
        frmCut.pgb.Max = lFileCount * lBuf2
        
        For i = 1 To lFileCount - 1
        
            'le fichier que l'on lit
            sFic = sFolderOut & "\" & sFileStr & "." & Trim$(Str$(i))
            
            For j = 1 To lBuf2 - 1
                'sbuf contient 5Mo lus
                sBuf = GetBytesFromFile(sFic, lBufSize, lBufSize * (j - 1))
                
                '�crit les bytes dans le fichier r�sultat
                WriteBytesToFileEnd sFolderOut & "\" & sFileStr, sBuf
                
                frmCut.pgb.Value = frmCut.pgb.Value + 1: DoEvents
            Next j
            
            'le dernier buffer
            a = cFile.GetFileSize(sFic) - (lBuf2 - 1) * lBufSize     'taille du dernier buffer
            sBuf = GetBytesFromFile(sFic, a, lBufSize * (lBuf2 - 1))
            
            '�crit les bytes dans le fichier r�sultat
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
            
            '�crit les bytes dans le fichier r�sultat
            WriteBytesToFileEnd sFolderOut & "\" & sFileStr, sBuf
        Next j
        
        'le dernier buffer
        a = cFile.GetFileSize(sFic) - (lBuf2 - 1) * lBufSize
        sBuf = GetBytesFromFile(sFic, a, lBufSize * (lBuf2 - 1))
        
        '�crit les bytes dans le fichier r�sultat
        WriteBytesToFileEnd sFolderOut & "\" & sFileStr, sBuf
        
        frmCut.pgb.Value = frmCut.pgb.Max
        DoEvents
        
    End If
    
    'termin�
    MsgBox "Fusion termin�e avec succ�s.", vbInformation + vbOKOnly, "Fusion r�ussie"
    
    PasteFile = GetTickCount - lTime
    Exit Function

ErrGestion:
End Function

'=======================================================
'effectue un modulo sans d�passement de capacit�
'tr�s peu optimis�, mais utile pour les grandes valeurs de cur
'=======================================================
Public Function Mod2(ByVal cur As Currency, lng As Long) As Currency
    Mod2 = cur - Int(cur / lng) * lng
End Function

