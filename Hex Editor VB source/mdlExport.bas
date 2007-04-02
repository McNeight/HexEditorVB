Attribute VB_Name = "mdlExport"
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
'MODULE CONTENANT LES PROCEDURES D'EXPORTS VERS LES DIFFERENTS FORMATS
'=======================================================


'=======================================================
'sauvegarde en HTML
'paramètres : sOutputFile (fichier de sortie)
'boolean pour les options
'sStringHex : contient la suite des valeurs hexa, ou le fichier d'entrée si fichier entier
'curFirstOffset : premier offset de la sélection (-1 si fichier entier)
'=======================================================
Public Sub SaveAsHTML(ByVal sOutputFile As String, ByVal bOffset As Boolean, _
    bString As Boolean, ByVal sStringHex As String, ByVal curFirstOffset As Currency, _
    Optional ByVal curSecondOffset As Currency, Optional ByVal lSize As Long = 3, _
    Optional ByVal bClip As Boolean = False)

Dim s As String
Dim curS As Currency
Dim x As Long
Dim y As Long
Dim s3 As String
Dim s2 As String
Dim z As Long
Dim s4 As String
Dim sRes As String
    
    'exemple de string au format HTML (contient une ligne avec offset, hexa et ASCII)
    '<font face="Courier New"><font size="1">1248A1ED1 </font><font color="#0000ff" size="1"
    '>4D5A 0000 0000 0000 0000 0000 0000 0000 </font><font color="#000000" size="1"> M
    'Z..................<BR>
    
    If frmContent.ActiveForm Is Nothing Then Exit Sub
    DoEvents
    
    'créé un nom de fichier temp si option "copie dans le clipboard"
    If bClip Then
        Call ObtainTempPathFile("temp_clip", sOutputFile, vbNullString)
    End If
    
    If curFirstOffset = -1 Then
        'alors c'est le fichier/disque/process entier
    
        'la méthode de sauvegarde dépend du type d'activeform
        Select Case TypeOfForm(frmContent.ActiveForm)
        
            Case "Fichier"
                'sauvegarde du fichier
                'lecture de 16kB en 16kB
                               
                'récupère la taille du fichier
                curS = cFile.GetFileSize(sStringHex)
                Call cFile.CreateEmptyFile(sOutputFile, True)
                
                
                For x = 1 To Int(curS / 16000)
                    'récupère les bytes
                    s = GetBytesFromFile(sStringHex, 16000, 16000 * (x - 1))
                    sRes = vbNullString
                    
                    'maintenant on créé le buffer avec les balises HTML
                    For y = 1 To Len(s) Step 16
                        'récupère 16 de long
                        s2 = Mid$(s, y, 16)
        
                        s3 = Space$(48)
                        'on récupère tous les valeurs hexa
                        For z = 1 To Len(s2)
                            Mid$(s3, 3 * z - 2, 3) = Str2Hex_(Mid$(s2, z, 1)) & " "
                        Next z
                        
                        s2 = Formated16String(s2)
                        s2 = Replace$(s2, "<", " &lt;")
                        s2 = Replace$(s2, ">", " &gt;")
                        s4 = ExtendedHex((16000 * (x - 1) + y - 1))
                        If Len(s4) < 10 Then s4 = String$(10 - Len(s4), "0") & s4
                        sRes = sRes & "<font face=|Courier New|><font size=|" & Str$(lSize) & "|>" & s4 & _
                            " " & "</font><font color=|#0000ff| size=|" & Str$(lSize) & "|>" & s3 & _
                            " </font><font color=|#000000| size=|" & Str$(lSize) & "|>" & s2 & _
                            "<BR>" & vbNewLine  'AVEC OPTIMISATION (BAD RESULT)

                    Next y
                    Call WriteBytesToFileEnd(sOutputFile, Replace$(sRes, "|", Chr$(34), , , _
                        vbBinaryCompare)): DoEvents
                Next x
                
                's'occupe de la dernière partie du fichier
                s = GetBytesFromFile(sStringHex, curS - 16000 * (x - 1), 16000 * (x - 1))
                sRes = vbNullString
                
                'maintenant on créé le buffer avec les balises HTML
                For y = 1 To Len(s) Step 16
                    'récupère 16 de long
                    s2 = Mid$(s, y, 16)
    
                    s3 = Space$(48)
                    'on récupère tous les valeurs hexa
                    For z = 1 To Len(s2)
                        Mid$(s3, 3 * z - 2, 3) = Str2Hex_(Mid$(s2, z, 1)) & " "
                    Next z
                    
                    s2 = Formated16String(s2)
                    s2 = Replace$(s2, "<", " &lt;")
                    s2 = Replace$(s2, ">", " &gt;")
                    s4 = ExtendedHex((16000 * (x - 1) + y - 1))
                    If Len(s4) < 10 Then s4 = String$(10 - Len(s4), "0") & s4
                    sRes = sRes & "<font face=|Courier New|><font size=|" & Str$(lSize) & "|>" & s4 & _
                        " " & "</font><font color=|#0000ff| size=|" & Str$(lSize) & "|>" & s3 & _
                        " </font><font color=|#000000| size=|" & Str$(lSize) & "|>" & s2 & _
                        "<BR>" & vbNewLine  'AVEC OPTIMISATION (BAD RESULT)
                    'sRes = sRes & "<font face=" & Chr$(34) & "Courier New" & Chr$(34) & _
                        "><font size=" & Chr$(34) & Str$(lSize) & Chr$(34) & ">" & s4 & _
                        " " & "</font><font color=" & Chr$(34) & "#0000ff" & Chr$(34) & _
                        " size=" & Chr$(34) & Str$(lSize) & Chr$(34) & ">" & s3 & _
                        " </font><font color=" & Chr$(34) & "#000000" & Chr$(34) & _
                        " size=" & Chr$(34) & Str$(lSize) & Chr$(34) & ">" & s2 & _
                        "<BR>" & vbNewLine  'SANS OPTIMISATION (NORMAL RESULT)
                Next y
                       
                Call WriteBytesToFileEnd(sOutputFile, sRes)
                
            Case "Disque"
            
            Case "Processus"
            
            Case "Disque physique"
            
            
            Case Else
                MsgBox "Form not defined", vbCritical, "Internal error"
                Exit Sub
        End Select
        
        
    Else
        'alors juste la sélection
    
        
        'la méthode de sauvegarde dépend du type d'activeform
        Select Case TypeOfForm(frmContent.ActiveForm)
        
            Case "Fichier"

            Case "Disque"
            
            Case "Processus"
            
            Case "Disque physique"
            
            
            Case Else
                MsgBox "Form not defined", vbCritical, "Internal error"
                Exit Sub
        End Select
        
    End If
    
    'on copie maintenant dans le clipboard si l'option est activée
    If bClip Then
        Call Clipboard.Clear
        Clipboard.SetText cFile.LoadFileInString(sOutputFile)
    End If
    
End Sub

'=======================================================
'sauvegarde en TEXTE SIMPLE
'paramètres : sOutputFile (fichier de sortie)
'boolean pour les options
'sStringHex : contient la suite des valeurs hexa, ou le fichier d'entrée si fichier entier
'curFirstOffset : premier offset de la sélection (-1 si fichier entier)
'=======================================================
Public Sub SaveAsTEXT(ByVal sOutputFile As String, ByVal bOffset As Boolean, _
    bString As Boolean, ByVal sStringHex As String, ByVal curFirstOffset As Currency, _
    Optional ByVal curSecondOffset As Currency, Optional ByVal bClip As Boolean = False)

Dim s As String
Dim curS As Currency
Dim x As Long
Dim y As Long
Dim s3 As String
Dim s2 As String
Dim z As Long
Dim s4 As String
Dim sRes As String
    
    'exemple de string au format TEXTE SIMPLE
    '012A45780124781
    
    If frmContent.ActiveForm Is Nothing Then Exit Sub
    DoEvents
    
    'créé un nom de fichier temp si option "copie dans le clipboard"
    If bClip Then
        Call ObtainTempPathFile("temp_clip", sOutputFile, vbNullString)
    End If
    
    If curFirstOffset = -1 Then
        'alors c'est le fichier/disque/process entier
    
        'la méthode de sauvegarde dépend du type d'activeform
        Select Case TypeOfForm(frmContent.ActiveForm)
        
            Case "Fichier"
                'sauvegarde du fichier
                'lecture de 16kB en 16kB
                               
                'récupère la taille du fichier
                curS = cFile.GetFileSize(sStringHex)
                Call cFile.CreateEmptyFile(sOutputFile, True)
                
                
                For x = 1 To Int(curS / 16000)
                    'récupère les bytes
                    s = GetBytesFromFile(sStringHex, 16000, 16000 * (x - 1))
                    sRes = vbNullString
                    
                    'maintenant on créé le buffer
                    For y = 1 To Len(s) Step 16
                        'récupère 16 de long
                        s2 = Mid$(s, y, 16)
        
                        s3 = Space$(48)
                        'on récupère tous les valeurs hexa
                        For z = 1 To Len(s2)
                            Mid$(s3, 3 * z - 2, 3) = Str2Hex_(Mid$(s2, z, 1)) & " "
                        Next z
                        
                        s2 = Formated16String(s2)
                        s4 = ExtendedHex((16000 * (x - 1) + y - 1))
                        If Len(s4) < 10 Then s4 = String$(10 - Len(s4), "0") & s4
                        If bOffset Then sRes = sRes & s4 & "   "
                        sRes = sRes & s3
                        If bString Then sRes = sRes & "   " & s2
                        sRes = sRes & vbNewLine
                        
                    Next y
                    Call WriteBytesToFileEnd(sOutputFile, sRes): DoEvents
                Next x
                
                's'occupe de la dernière partie du fichier
                s = GetBytesFromFile(sStringHex, curS - 16000 * (x - 1), 16000 * (x - 1))
                sRes = vbNullString
                
                'maintenant on créé le buffer
                For y = 1 To Len(s) Step 16
                    'récupère 16 de long
                    s2 = Mid$(s, y, 16)
    
                    s3 = Space$(48)
                    'on récupère tous les valeurs hexa
                    For z = 1 To Len(s2)
                        Mid$(s3, 3 * z - 2, 3) = Str2Hex_(Mid$(s2, z, 1)) & " "
                    Next z
                    
                    s2 = Formated16String(s2)
                    s4 = ExtendedHex((16000 * (x - 1) + y - 1))
                    If Len(s4) < 10 Then s4 = String$(10 - Len(s4), "0") & s4
                    If bOffset Then sRes = sRes & s4
                    sRes = sRes & "   " & s3
                    If bString Then sRes = sRes & "   " & s2
                    sRes = sRes & vbNewLine

                Next y
                
                Call WriteBytesToFileEnd(sOutputFile, sRes): DoEvents
                
            Case "Disque"
            
            Case "Processus"
            
            Case "Disque physique"
            
            
            Case Else
                MsgBox "Form not defined", vbCritical, "Internal error"
                Exit Sub
        End Select
        
        
    Else
        'alors juste la sélection
    
        
        'la méthode de sauvegarde dépend du type d'activeform
        Select Case TypeOfForm(frmContent.ActiveForm)
        
            Case "Fichier"
                
            Case "Disque"
            
            Case "Processus"
            
            Case "Disque physique"
            
            
            Case Else
                MsgBox "Form not defined", vbCritical, "Internal error"
                Exit Sub
        End Select
        
    End If
    
    'on copie maintenant dans le clipboard si l'option est activée
    If bClip Then
        Call Clipboard.Clear
        Clipboard.SetText cFile.LoadFileInString(sOutputFile)
    End If
    
End Sub

'=======================================================
'sauvegarde en CODE C
'paramètres : sOutputFile (fichier de sortie)
'sStringHex : contient la suite des valeurs hexa, ou le fichier d'entrée si fichier entier
'curFirstOffset : premier offset de la sélection (-1 si fichier entier)
'=======================================================
Public Sub SaveAsC(ByVal sOutputFile As String, ByVal sStringHex As String, _
    ByVal curFirstOffset As Currency, Optional ByVal curSecondOffset As Currency, _
    Optional ByVal bClip As Boolean = False)

Dim s As String
Dim curS As Currency
Dim x As Long
Dim y As Long
Dim s3 As String
Dim s2 As String
Dim z As Long
Dim s4 As String
Dim sRes As String
    
    'exemple de string au format CODE C
    '/* Source File: AUTHORS.TXT
    'Length: 118*/
    'unsigned char rawData[118] =
    '{
    '    0x43, 0x6F, 0x70, 0x79, 0x72, 0x69, 0x67, 0x68, 0x74, 0x20, 0x62, 0x79, 0x0D, 0x0A, 0x41, 0x6C,
    '    0x61, 0x69, 0x6E, 0x20, 0x44, 0x65, 0x73, 0x63, 0x6F, 0x74, 0x65, 0x73, 0x20, 0x28, 0x76, 0x69,
    '    0x6F, 0x6C, 0x65, 0x6E, 0x74, 0x5F, 0x6B, 0x65, 0x6E, 0x29, 0x0D, 0x0A, 0x3C, 0x61, 0x6C, 0x61,
    '    0x69, 0x6E, 0x64, 0x65, 0x73, 0x63, 0x6F, 0x74, 0x65, 0x73, 0x40, 0x68, 0x6F, 0x74, 0x6D, 0x61,
    '    0x69, 0x6C, 0x2E, 0x66, 0x72, 0x2C, 0x20, 0x61, 0x6C, 0x61, 0x69, 0x6E, 0x64, 0x65, 0x73, 0x63,
    '    0x6F, 0x74, 0x65, 0x73, 0x40, 0x67, 0x6D, 0x61, 0x69, 0x6C, 0x2E, 0x63, 0x6F, 0x6D, 0x2C, 0x20,
    '    0x68, 0x65, 0x78, 0x65, 0x64, 0x69, 0x74, 0x6F, 0x72, 0x76, 0x62, 0x40, 0x67, 0x6D, 0x61, 0x69,
    '    0x6C, 0x2E, 0x63, 0x6F, 0x6D, 0x3E,
    '} ;
    
    
    If frmContent.ActiveForm Is Nothing Then Exit Sub
    DoEvents
    
    'créé un nom de fichier temp si option "copie dans le clipboard"
    If bClip Then
        Call ObtainTempPathFile("temp_clip", sOutputFile, vbNullString)
    End If
    
    If curFirstOffset = -1 Then
        'alors c'est le fichier/disque/process entier
    
        'la méthode de sauvegarde dépend du type d'activeform
        Select Case TypeOfForm(frmContent.ActiveForm)
        
            Case "Fichier"
                'sauvegarde du fichier
                'lecture de 16kB en 16kB
                               
                'récupère la taille du fichier
                curS = cFile.GetFileSize(sStringHex)
                Call cFile.CreateEmptyFile(sOutputFile, True)
                
                'pose le header
                s = "/* Source File: " & cFile.GetFileFromPath(sStringHex)
                s = s & vbNewLine & "Length: " & Trim$(Str$(curS)) & "*/" & vbNewLine
                s = s & "unsigned char rawData[" & Trim$(Str$(curS)) & "] =" & vbNewLine & "{" & vbNewLine
                Call WriteBytesToFileEnd(sOutputFile, s)
                
                For x = 1 To Int(curS / 16000)
                    'récupère les bytes
                    s = GetBytesFromFile(sStringHex, 16000, 16000 * (x - 1))
                    sRes = vbNullString
                    
                    'maintenant on créé le buffer
                    For y = 1 To Len(s) Step 16
        
                    'on récupère toutes les valeurs hexa
                    sRes = sRes & "   "
                    s2 = Mid$(s, y, 16)
                    For z = 1 To Len(s2)
                        sRes = sRes & "0x" & Hex_(Asc(Mid$(s, y + z - 1, 1))) & ", "
                    Next z
                    sRes = sRes & vbNewLine
                        
                    Next y
                    Call WriteBytesToFileEnd(sOutputFile, sRes): DoEvents
                Next x
                
                's'occupe de la dernière partie du fichier
                s = GetBytesFromFile(sStringHex, curS - 16000 * (x - 1), 16000 * (x - 1))
                sRes = vbNullString
                
                'maintenant on créé le buffer
                For y = 1 To Len(s) Step 16
    
                    'on récupère toutes les valeurs hexa
                    sRes = sRes & "   "
                    s2 = Mid$(s, y, 16)
                    For z = 1 To Len(s2)
                        sRes = sRes & "0x" & Hex_(Asc(Mid$(s, y + z - 1, 1))) & ", "
                    Next z
                    sRes = sRes & vbNewLine
                    
                Next y
                Call WriteBytesToFileEnd(sOutputFile, sRes): DoEvents
                Call WriteBytesToFileEnd(sOutputFile, vbNewLine & "};")
            Case "Disque"
            
            Case "Processus"
            
            Case "Disque physique"
            
            
            Case Else
                MsgBox "Form not defined", vbCritical, "Internal error"
                Exit Sub
        End Select
        
        
    Else
        'alors juste la sélection
    
        
        'la méthode de sauvegarde dépend du type d'activeform
        Select Case TypeOfForm(frmContent.ActiveForm)
        
            Case "Fichier"
                
            Case "Disque"
            
            Case "Processus"
            
            Case "Disque physique"
            
            
            Case Else
                MsgBox "Form not defined", vbCritical, "Internal error"
                Exit Sub
        End Select
        
    End If
    
    'on copie maintenant dans le clipboard si l'option est activée
    If bClip Then
        Call Clipboard.Clear
        Clipboard.SetText cFile.LoadFileInString(sOutputFile)
    End If
    
End Sub

'=======================================================
'sauvegarde en CODE JAVA
'paramètres : sOutputFile (fichier de sortie)
'sStringHex : contient la suite des valeurs hexa, ou le fichier d'entrée si fichier entier
'curFirstOffset : premier offset de la sélection (-1 si fichier entier)
'=======================================================
Public Sub SaveAsJAVA(ByVal sOutputFile As String, ByVal sStringHex As String, _
    ByVal curFirstOffset As Currency, Optional ByVal curSecondOffset As Currency, _
    Optional ByVal bClip As Boolean = False)

Dim s As String
Dim curS As Currency
Dim x As Long
Dim y As Long
Dim s3 As String
Dim s2 As String
Dim z As Long
Dim s4 As String
Dim sRes As String
    
    'exemple de string au format CODE JAVA
    '/* Source File: AUTHORS.TXT
    'Length: 118*/
    'byte char rawData[] =
    '{
    '    0x43, 0x6F, 0x70, 0x79, 0x72, 0x69, 0x67, 0x68, 0x74, 0x20, 0x62, 0x79, 0x0D, 0x0A, 0x41, 0x6C,
    '    0x61, 0x69, 0x6E, 0x20, 0x44, 0x65, 0x73, 0x63, 0x6F, 0x74, 0x65, 0x73, 0x20, 0x28, 0x76, 0x69,
    '    0x6F, 0x6C, 0x65, 0x6E, 0x74, 0x5F, 0x6B, 0x65, 0x6E, 0x29, 0x0D, 0x0A, 0x3C, 0x61, 0x6C, 0x61,
    '    0x69, 0x6E, 0x64, 0x65, 0x73, 0x63, 0x6F, 0x74, 0x65, 0x73, 0x40, 0x68, 0x6F, 0x74, 0x6D, 0x61,
    '    0x69, 0x6C, 0x2E, 0x66, 0x72, 0x2C, 0x20, 0x61, 0x6C, 0x61, 0x69, 0x6E, 0x64, 0x65, 0x73, 0x63,
    '    0x6F, 0x74, 0x65, 0x73, 0x40, 0x67, 0x6D, 0x61, 0x69, 0x6C, 0x2E, 0x63, 0x6F, 0x6D, 0x2C, 0x20,
    '    0x68, 0x65, 0x78, 0x65, 0x64, 0x69, 0x74, 0x6F, 0x72, 0x76, 0x62, 0x40, 0x67, 0x6D, 0x61, 0x69,
    '    0x6C, 0x2E, 0x63, 0x6F, 0x6D, 0x3E,
    '} ;
    
    
    If frmContent.ActiveForm Is Nothing Then Exit Sub
    DoEvents
    
    'créé un nom de fichier temp si option "copie dans le clipboard"
    If bClip Then
        Call ObtainTempPathFile("temp_clip", sOutputFile, vbNullString)
    End If
    
    If curFirstOffset = -1 Then
        'alors c'est le fichier/disque/process entier
    
        'la méthode de sauvegarde dépend du type d'activeform
        Select Case TypeOfForm(frmContent.ActiveForm)
        
            Case "Fichier"
                'sauvegarde du fichier
                'lecture de 16kB en 16kB
                               
                'récupère la taille du fichier
                curS = cFile.GetFileSize(sStringHex)
                Call cFile.CreateEmptyFile(sOutputFile, True)
                
                'pose le header
                s = "/* Source File: " & cFile.GetFileFromPath(sStringHex)
                s = s & vbNewLine & "Length: " & Trim$(Str$(curS)) & "*/" & vbNewLine
                s = s & "byte rawData[] =" & vbNewLine & "{" & vbNewLine
                Call WriteBytesToFileEnd(sOutputFile, s)
                
                For x = 1 To Int(curS / 16000)
                    'récupère les bytes
                    s = GetBytesFromFile(sStringHex, 16000, 16000 * (x - 1))
                    sRes = vbNullString
                    
                    'maintenant on créé le buffer
                    For y = 1 To Len(s) Step 16
        
                    'on récupère toutes les valeurs hexa
                    sRes = sRes & "   "
                    s2 = Mid$(s, y, 16)
                    For z = 1 To Len(s2)
                        sRes = sRes & "0x" & Hex_(Asc(Mid$(s, y + z - 1, 1))) & ", "
                    Next z
                    sRes = sRes & vbNewLine
                        
                    Next y
                    Call WriteBytesToFileEnd(sOutputFile, sRes): DoEvents
                Next x
                
                's'occupe de la dernière partie du fichier
                s = GetBytesFromFile(sStringHex, curS - 16000 * (x - 1), 16000 * (x - 1))
                sRes = vbNullString
                
                'maintenant on créé le buffer
                For y = 1 To Len(s) Step 16
    
                    'on récupère toutes les valeurs hexa
                    sRes = sRes & "   "
                    s2 = Mid$(s, y, 16)
                    For z = 1 To Len(s2)
                        sRes = sRes & "0x" & Hex_(Asc(Mid$(s, y + z - 1, 1))) & ", "
                    Next z
                    sRes = sRes & vbNewLine
                    
                Next y
                Call WriteBytesToFileEnd(sOutputFile, sRes): DoEvents
                Call WriteBytesToFileEnd(sOutputFile, vbNewLine & "};")
            Case "Disque"
            
            Case "Processus"
            
            Case "Disque physique"
            
            
            Case Else
                MsgBox "Form not defined", vbCritical, "Internal error"
                Exit Sub
        End Select
        
        
    Else
        'alors juste la sélection
    
        
        'la méthode de sauvegarde dépend du type d'activeform
        Select Case TypeOfForm(frmContent.ActiveForm)
        
            Case "Fichier"
                
            Case "Disque"
            
            Case "Processus"
            
            Case "Disque physique"
            
            
            Case Else
                MsgBox "Form not defined", vbCritical, "Internal error"
                Exit Sub
        End Select
        
    End If
    
    'on copie maintenant dans le clipboard si l'option est activée
    If bClip Then
        Call Clipboard.Clear
        Clipboard.SetText cFile.LoadFileInString(sOutputFile)
    End If
    
End Sub

'=======================================================
'sauvegarde en CODE VB
'paramètres : sOutputFile (fichier de sortie)
'sStringHex : contient la suite des valeurs hexa, ou le fichier d'entrée si fichier entier
'curFirstOffset : premier offset de la sélection (-1 si fichier entier)
'=======================================================
Public Sub SaveAsVB(ByVal sOutputFile As String, ByVal sStringHex As String, _
    ByVal curFirstOffset As Currency, Optional ByVal curSecondOffset As Currency, _
    Optional ByVal sSep As String = vbNullString, Optional ByVal bClip As Boolean = False)

Dim s As String
Dim curS As Currency
Dim x As Long
Dim z2 As Long
Dim y As Long
Dim s3 As String
Dim s2 As String
Dim z As Long
Dim s4 As String
Dim o As Long
Dim ov As Long
Dim sRes As String
    
    'exemple de string au format CODE VB
    ''==========================================
    ''Source file: fichier.txt
    ''Length: 13
    '==========================================
    'Private Const HEX_VALUES = "48455820454449544F52205642"
    
    
    If frmContent.ActiveForm Is Nothing Then Exit Sub
    DoEvents
    
    'créé un nom de fichier temp si option "copie dans le clipboard"
    If bClip Then
        Call ObtainTempPathFile("temp_clip", sOutputFile, vbNullString)
    End If
    
    If curFirstOffset = -1 Then
        'alors c'est le fichier/disque/process entier
    
        'la méthode de sauvegarde dépend du type d'activeform
        Select Case TypeOfForm(frmContent.ActiveForm)
        
            Case "Fichier"
                'sauvegarde du fichier
                'lecture de 16kB en 16kB
                               
                'récupère la taille du fichier
                curS = cFile.GetFileSize(sStringHex)
                Call cFile.CreateEmptyFile(sOutputFile, True)
                
                'pose le header
                s = "'==========================================" & vbNewLine & "'Source File: " & cFile.GetFileFromPath(sStringHex)
                s = s & vbNewLine & "'Length: " & Trim$(Str$(curS)) & vbNewLine & "'==========================================" & vbNewLine
                s = s & "Private Const HEX_VALUES = " & Chr$(34)
                Call WriteBytesToFileEnd(sOutputFile, s)
                
                o = 0   'nombre de retours à la ligne
                ov = 0
                
                For x = 1 To Int(curS / 16000)
                    'récupère les bytes
                    s = GetBytesFromFile(sStringHex, 16000, 16000 * (x - 1))
                    sRes = vbNullString
                    z2 = 0
                    
                    'maintenant on créé le buffer
                    For y = 1 To Len(s) Step 16
        
                        'on récupère toutes les valeurs hexa
                        s2 = Mid$(s, y, 16)
                        For z = 1 To Len(s2)
                            sRes = sRes & Hex_(Asc(Mid$(s, y + z - 1, 1))) & sSep
                        Next z
                        
                        If Len(sRes) - z2 > 800 Then
                            'alors il faut faire un saut de ligne
                            z2 = Len(sRes)
                            sRes = sRes & Chr$(34) & IIf(o < 10, " & _" & vbNewLine & "    " & Chr$(34), vbNewLine)
                            o = o + 1
                        End If
                        
                        If o > 10 Then
                            'alors trop de retours à la ligne
                            ov = ov + 1
                            sRes = sRes & IIf(o <> 11, Chr$(34), vbNullString) & vbNewLine & "Private Const HEX_VALUES_" & Trim$(Str$(ov)) & " = " & Chr$(34)
                            o = 0
                        End If
                        
                    Next y
                    Call WriteBytesToFileEnd(sOutputFile, sRes): DoEvents
                Next x
                
                's'occupe de la dernière partie du fichier
                s = GetBytesFromFile(sStringHex, curS - 16000 * (x - 1), 16000 * (x - 1))
                sRes = vbNullString
                z2 = 0

                'maintenant on créé le buffer
                For y = 1 To Len(s) Step 16
    
                    'on récupère toutes les valeurs hexa
                    s2 = Mid$(s, y, 16)
                    For z = 1 To Len(s2)
                        sRes = sRes & Hex_(Asc(Mid$(s, y + z - 1, 1))) & sSep
                    Next z
                    
                    If Len(sRes) - z2 > 800 Then
                        'alors il faut faire un saut de ligne
                        z2 = Len(sRes)
                        sRes = sRes & Chr$(34) & IIf(o < 10, " & _" & vbNewLine & "    " & Chr$(34), vbNewLine)
                        o = o + 1
                    End If
                    
                    If o > 10 Then
                        'alors trop de retours à la ligne
                        ov = ov + 1
                        sRes = sRes & IIf(o <> 11, Chr$(34), vbNullString) & vbNewLine & "Private Const HEX_VALUES_" & Trim$(Str$(ov)) & " = " & Chr$(34)
                        o = 0
                    End If
                    
                Next y
                Call WriteBytesToFileEnd(sOutputFile, sRes): DoEvents
                Call WriteBytesToFileEnd(sOutputFile, Chr$(34))
            Case "Disque"
            
            Case "Processus"
            
            Case "Disque physique"
            
            
            Case Else
                MsgBox "Form not defined", vbCritical, "Internal error"
                Exit Sub
        End Select
        
        
    Else
        'alors juste la sélection
    
        
        'la méthode de sauvegarde dépend du type d'activeform
        Select Case TypeOfForm(frmContent.ActiveForm)
        
            Case "Fichier"
                
            Case "Disque"
            
            Case "Processus"
            
            Case "Disque physique"
            
            
            Case Else
                MsgBox "Form not defined", vbCritical, "Internal error"
                Exit Sub
        End Select
        
    End If
    
    'on copie maintenant dans le clipboard si l'option est activée
    If bClip Then
        Call Clipboard.Clear
        Clipboard.SetText cFile.LoadFileInString(sOutputFile)
    End If
    
End Sub


