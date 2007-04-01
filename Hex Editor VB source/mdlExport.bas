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
    Optional ByVal curSecondOffset As Currency, Optional ByVal lSize As Long = 3)

Dim s As String
Dim curS As Currency
Dim sObj As String
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
    
    If curFirstOffset = -1 Then
        'alors c'est le fichier/disque/process entier
    
        'la méthode de sauvegarde dépend du type d'activeform
        Select Case TypeOfForm(frmContent.ActiveForm)
        
            Case "Fichier"
                'sauvegarde du fichier
                'lecture de 16kB en 16kB
                
                sObj = frmContent.ActiveForm.Caption 'fichier
                
                'récupère la taille du fichier
                curS = cFile.GetFileSize(sObj)
                Call cFile.CreateEmptyFile(sOutputFile, True)
                
                
                For x = 1 To Int(curS / 16000)
                    'récupère les bytes
                    s = GetBytesFromFile(sObj, 16000, 16000 * (x - 1))
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
                        If Len(s4) < 16 Then s4 = String$(16 - Len(s4), "0") & s4
                        sRes = sRes & "<font face=|Courier New|><font size=|" & Str$(lSize) & "|>" & s4 & _
                            " " & "</font><font color=|#0000ff| size=|" & Str$(lSize) & "|>" & s3 & _
                            " </font><font color=|#000000| size=|" & Str$(lSize) & "|>" & s2 & _
                            "<BR>" & vbNewLine

                    Next y
                    Call WriteBytesToFileEnd(sOutputFile, Replace$(sRes, "|", Chr$(34), , , _
                        vbBinaryCompare)): DoEvents
                Next x
                
                's'occupe de la dernière partie du fichier
                s = GetBytesFromFile(sObj, curS - 16000 * (x - 1), 16000 * (x - 1))
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
                    If Len(s4) < 16 Then s4 = String$(16 - Len(s4), "0") & s4
                    sRes = sRes & "<font face=" & Chr$(34) & "Courier New" & Chr$(34) & _
                        "><font size=" & Chr$(34) & Str$(lSize) & Chr$(34) & ">" & s4 & _
                        " " & "</font><font color=" & Chr$(34) & "#0000ff" & Chr$(34) & _
                        " size=" & Chr$(34) & Str$(lSize) & Chr$(34) & ">" & s3 & _
                        " </font><font color=" & Chr$(34) & "#000000" & Chr$(34) & _
                        " size=" & Chr$(34) & Str$(lSize) & Chr$(34) & ">" & s2 & _
                        "<BR>" & vbNewLine
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
                'sauvegarde du fichier
                'lecture de 16 en 16 bytes
                
                'récupère la taille du fichier
                s = GetBytesFromFile(frmContent.ActiveForm.Caption, 16, 16 * x)
                
                
            Case "Disque"
            
            Case "Processus"
            
            Case "Disque physique"
            
            
            Case Else
                MsgBox "Form not defined", vbCritical, "Internal error"
                Exit Sub
        End Select
        
    End If
    
End Sub
