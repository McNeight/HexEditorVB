Attribute VB_Name = "mdlStrings"
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
'//MODULE DE GESTION DES STRINGS
'//FONCTIONS GENERIQUES D'OPERATIONS SUR DES STRINGS
'=======================================================


'=======================================================
'renvoie le nom du path sans le fichier
'=======================================================
Public Function GetFolderFormPath(ByVal sPath As String) As String
    GetFolderFormPath = Left(sPath, InStrRev(sPath, "\", Len(sPath)))
End Function

'=======================================================
'formate les 16 caractères d'une chaine de 16
'=======================================================
Public Function Formated16String(ByVal sString As String) As String
Dim x As Long
Dim s As String

    On Error Resume Next
    
    s = vbNullString
    
    For x = 1 To 16
        s = s & Byte2FormatedString(Asc(Mid$(sString, x, 1)))
    Next x
    
    Formated16String = s
End Function

'=======================================================
'formate les n caractères d'une chaine de n de long
'pourquoi utilisé Formated16String et Formated1String alors
'qu'il existe cette fonction ? Pour des raisons de performance.
'=======================================================
Public Function FormatednString(ByVal sString As String) As String
Dim x As Long
Dim curLen As Currency
Dim s As String

    s = vbNullString
    
    'longueur de la chaine à formater
    curLen = Len(sString)
    
    For x = 1 To curLen
        If (x Mod 2000) = 0 Then DoEvents
        s = s & Byte2FormatedString(Asc(Mid$(sString, x, 1)))
    Next x
    
    FormatednString = s
End Function

'=======================================================
'formate un caractère string vers quelque chose de lisible
'=======================================================
Public Function Formated1String(ByVal sString As String) As String

    Formated1String = Byte2FormatedString(Asc(sString))

End Function

'=======================================================
'formate la taille d'un fichier
'=======================================================
Public Function FormatedSize(ByVal LS As Currency, Optional ByVal lRoundNumber = 5) As String
Dim dS As Double
Dim n As Byte

    On Error Resume Next
    
    dS = LS: n = 0
    While (dS / 1024) > 1
        n = n + 1
        dS = dS / 1024
        DoEvents
    Wend
    
    dS = Round(dS, lRoundNumber)
    
    With frmContent.Lang
        If n = 0 Then FormatedSize = Str$(dS) & " " & .GetString("_Bytes")
        If n = 1 Then FormatedSize = Str$(dS) & " " & .GetString("_Ko")
        If n = 2 Then FormatedSize = Str$(dS) & " " & .GetString("_Mo")
        If n = 3 Then FormatedSize = Str$(dS) & " " & .GetString("_Go")
    End With
    
    FormatedSize = Trim$(FormatedSize)
    
End Function

'=======================================================
'colle la string s1 dans s2, à l'emplacement ldep
'=======================================================
Public Sub PasteS1inS2(ByVal s1 As String, ByRef s2 As String, ByVal lDep As Long)
Dim sAvant As String
Dim sApres As String

    'découpe s2 en sAvant, vide , sApres et fait la concaténation
    
    If lDep = 1 Then
        'pas de sAvant
        sAvant = vbNullString
        sApres = Mid$(s2, Len(s1) + 1, Len(s2) - Len(s1))
    ElseIf lDep = Len(s2) Then
        'pas de sApres
        sApres = vbNullString
        sAvant = Mid$(s2, 1, Len(s2) - Len(s1))
    Else
        'alors il y a sAvant ET sApres
        sAvant = Mid$(s2, 1, lDep - 1)
        sApres = Mid$(s2, lDep + Len(s1), 1 + Len(s2) - lDep - Len(s1))
    End If
    
    s2 = sAvant & s1 & sApres
End Sub

'=======================================================
'formate un byte en string en enlevant les caractères ASCII
'non représentés(ables) dans un visualisateur hexa
'=======================================================
Public Function Byte2FormatedString(ByVal bCar As Long) As String

    'renvoie un "." pour les caractères non affichables

    If bCar < 32 Or bCar > 255 Or bCar = 144 Or bCar = 143 Then
        'caractère non affichable
        Byte2FormatedString = "."
    Else
        'caractère OK
        Byte2FormatedString = Chr$(bCar)
    End If
   
End Function

'=======================================================
'renvoie un tableau de 1 à ubound de string
'qui contient les strings comprises entre un caractère défini
'=======================================================
Public Sub SplitString(ByVal strSeparator As String, ByVal strString As String, ByRef strArray() As String)
Dim s As String
Dim x As Long
Dim i As Long

    i = 0

    'redimensionne le tableau
    ReDim strArray(0)
    
    For x = 1 To Len(strString)
        If Mid$(strString, x, 1) = strSeparator Then
            If i = 0 Then
                'alors c'est celui de gauche
                i = x
            Else
                'alors c'est celui de droite ==> stocke dans le tableau le Mid$ de la string
                ReDim Preserve strArray(UBound(strArray()) + 1)
                strArray(UBound(strArray())) = Mid$(strString, i + 1, x - i - 1)
                i = 0 'on recommencera en prenant la position de separateur de gauche
            End If
        End If
    Next x

End Sub

'=======================================================
'renvoie une adresse (string) avec les 0 devant si nécessaire, pour avoir un longueur
'de string fixe (8)
'=======================================================
Public Function FormatedAdress(ByVal lNumber As Long, Optional ByVal lLongueur As Long = 8) As String
Dim s As String

    s = CStr(lNumber)
    
    While Len(s) < lLongueur
        s = "0" + s
    Wend

    FormatedAdress = s
End Function

'=======================================================
'transforme une date en FILETIME vers une date en string
'=======================================================
Public Function FileTimeToString(fDate As FILETIME, Optional ByVal bConvertToLocal As Boolean = True) As String
Dim sDate As SYSTEMTIME
Dim sDay As String
Dim sMonth As String
Dim sYear As String
Dim sHour As String
Dim sMinute As String
Dim sSecond As String
Dim s As String

    If bConvertToLocal Then
        'conversion en LocalFileTime (temps universel ==> temps local)
        FileTimeToLocalFileTime fDate, fDate
    End If
    
    'conversion en SystemTime
    FileTimeToSystemTime fDate, sDate
    
    'conversion en string vers un format du genre 24/04/2000 09:50:59
    sDay = Trim$(IIf(sDate.wDay < 10, "0" & Trim$(Str$(sDate.wDay)), Trim$(Str$(sDate.wDay))))
    sMonth = Trim$(IIf(sDate.wMonth < 10, "0" & Trim$(Str$(sDate.wMonth)), Trim$(Str$(sDate.wMonth))))
    sHour = Trim$(IIf(sDate.wHour < 10, "0" & Trim$(Str$(sDate.wHour)), Trim$(Str$(sDate.wHour))))
    sMinute = Trim$(IIf(sDate.wMinute < 10, "0" & Trim$(Str$(sDate.wMinute)), Trim$(Str$(sDate.wMinute))))
    sSecond = Trim$(IIf(sDate.wSecond < 10, "0" & Trim$(Str$(sDate.wSecond)), Trim$(Str$(sDate.wSecond))))
    sYear = sDate.wYear
    
    s = sDay & "/" & sMonth & "/" & sYear & " " & sHour & ":" & sMinute & ":" & sSecond
    FileTimeToString = s

End Function

'=======================================================
'convertit le chemin sPath en chemin existant (correct)
'=======================================================
Public Function FormatedPath(ByVal sPath As String) As String
Dim x As Long
Dim s As String

    If Len(sPath) < 1 Then Exit Function
    
    'modifie le path si commence par SystemRoot
    'len("SystemRoot")=10
    If Left$(sPath, 10) = "SystemRoot" Then
        'obtient le répertoire de windows
        sPath = cFile.GetSpecialFolder(CSIDL_WINDOWS) & "\" & Right$(sPath, Len(sPath) - 10)
    End If
    'len("\SystemRoor")=11
    If Left$(sPath, 11) = "\SystemRoot" Then
        'obtient le répertoire de windows
        sPath = cFile.GetSpecialFolder(CSIDL_WINDOWS) & "\" & Right$(sPath, Len(sPath) - 11)
    End If
    
    s = sPath
    While ((Asc(UCase(Left$(s, 1))) < 65 Or Asc(UCase(Left$(s, 1))) > 90) And Len(s) > 3)
        'alors ce n'est pas une lettre valide ==> on enlève cette lettre
        s = Right$(s, Len(s) - 1)
        DoEvents
    Wend
    
    'enlève deux antislash successifs et les remplace par un seul
    x = InStr(1, s, "\\")
    If x > 0 Then
        s = Left$(s, x - 1) & "\" & Right$(s, Len(s) - Len(Left$(s, x - 1)) - 2)
    End If
    
    
    FormatedPath = s
End Function

'=======================================================
'obtient le NOM de la priorité en STRING à partir du long correspondant
'=======================================================
Public Function PriorityFromLong(ByVal lp As Long) As String
Dim s As String

    With frmContent.Lang
        s = .GetString("_IdleMdl")
        If lp >= 6 Then s = .GetString("_BelowMdl")
        If lp >= 8 Then s = .GetString("_NormalMdl")
        If lp >= 10 Then s = .GetString("_AboveMdl")
        If lp >= 13 Then s = .GetString("_HighMdl")
        If lp >= 24 Then s = .GetString("_RealMdl")
    End With
    PriorityFromLong = s
    
End Function

'=======================================================
'créé une string contenant les données en langage HTML à sauvegarder
'prend en paramètre les infos à afficher (deux listviews contenant les infos disques)
'=======================================================
Public Function CreateMeHtmlString(lvPhys As ListView, lvLog As ListView) As String
Dim s As String
Dim x As Long

    s = "<html>" & vbNewLine & "<body>" & vbNewLine & "<font face=" & Chr$(34) & "courier new" & Chr$(34) & ">" & vbNewLine & "<H2>Disques physiques</H2>"
    
    'disques physiques
    With lvPhys
        For x = 1 To .ListItems.Count
        
            s = s & "<font color=red>" & vbNewLine & "<div align=center>" & "<HR size=3 align=center width=100%>"
            s = s & "<B>" & frmContent.Lang.GetString("_DiskMdl") & " " & Str$(.ListItems.Item(x).Text) & "</B> <BR>" & vbNewLine & "<HR size=3 align=center width=100%>" & vbNewLine & "<P>" & vbNewLine & "</font>" & vbNewLine & "</div>"
            
            s = s & vbNewLine & "<B>" & frmContent.Lang.GetString("_SizeMdl") & "</B> = [" & .ListItems.Item(x).SubItems(1) & "]<BR>"
            s = s & vbNewLine & "<B>" & frmContent.Lang.GetString("_CylMdl") & "</B> = [" & .ListItems.Item(x).SubItems(2) & "]<BR>"
            s = s & vbNewLine & "<B>" & frmContent.Lang.GetString("_TrackPerCylMdl") & "</B> = [" & .ListItems.Item(x).SubItems(3) & "]<BR>"
            s = s & vbNewLine & "<B>" & frmContent.Lang.GetString("_SecPerTMdl") & "</B> = [" & .ListItems.Item(x).SubItems(4) & "]<BR>"
            s = s & vbNewLine & "<B>" & frmContent.Lang.GetString("_BytePerSecMdl") & "</B> = [" & .ListItems.Item(x).SubItems(5) & "]<BR>"
            s = s & vbNewLine & "<B>" & frmContent.Lang.GetString("_TypeMdl") & "</B> = [" & .ListItems.Item(x).SubItems(6) & "]<BR>"

            s = s & vbNewLine & "<BR> <BR>"
 
        Next x
        s = s & vbNewLine & "<BR> <BR> <BR> <BR> <BR> <BR>"
    End With
    
    
    'disques logiques
    s = s & vbNewLine & "<H2>" & frmContent.Lang.GetString("_LogicalDiskStr") & "</H2>"
    
    With lvLog
        For x = 1 To .ListItems.Count
        
            s = s & "<font color=red>" & vbNewLine & "<div align=center>" & vbNewLine & "<HR size=3 align=center width=100%>" & vbNewLine
            s = s & "<B>Disque " & .ListItems.Item(x).Text & "</B> <BR>" & vbNewLine & "<HR size=3 align=center width=100%>" & vbNewLine & "<P>" & vbNewLine & "</font>" & vbNewLine & "</div>"
            
            s = s & vbNewLine & "<B>" & frmContent.Lang.GetString("_SizeMdl") & "</B> = [" & .ListItems.Item(x).SubItems(1) & "]<BR>"
            s = s & vbNewLine & "<B>" & frmContent.Lang.GetString("_PhysSizeMdl") & "</B> = [" & .ListItems.Item(x).SubItems(2) & "]<BR>"
            s = s & vbNewLine & "<B>" & frmContent.Lang.GetString("_UsedMdl") & "</B> = [" & .ListItems.Item(x).SubItems(3) & "]<BR>"
            s = s & vbNewLine & "<B>" & frmContent.Lang.GetString("_FreeMdl") & "</B> = [" & .ListItems.Item(x).SubItems(4) & "]<BR>"
            s = s & vbNewLine & "<B>" & frmContent.Lang.GetString("_PercentMdl") & "</B> = [" & .ListItems.Item(x).SubItems(5) & "]<BR>"
            s = s & vbNewLine & "<B>" & frmContent.Lang.GetString("_ClustSizeMdl") & "</B> = [" & .ListItems.Item(x).SubItems(6) & "]<BR>"
            s = s & vbNewLine & "<B>" & frmContent.Lang.GetString("_ClustUsedMdl") & "</B> = [" & .ListItems.Item(x).SubItems(7) & "]<BR>"
            s = s & vbNewLine & "<B>" & frmContent.Lang.GetString("_ClustFreeMdl") & "</B> = [" & .ListItems.Item(x).SubItems(8) & "]<BR>"
                        s = s & vbNewLine & "<B>" & frmContent.Lang.GetString("_ClustMdl") & "</B> = [" & .ListItems.Item(x).SubItems(9) & "]<BR>"
            s = s & vbNewLine & "<B>" & frmContent.Lang.GetString("_HidSecMdl") & "</B> = [" & .ListItems.Item(x).SubItems(10) & "]<BR>"
            s = s & vbNewLine & "<B>" & frmContent.Lang.GetString("_LogMdl") & "</B> = [" & .ListItems.Item(x).SubItems(11) & "]<BR>"
            s = s & vbNewLine & "<B>" & frmContent.Lang.GetString("_PhysMdl") & "</B> = [" & .ListItems.Item(x).SubItems(12) & "]<BR>"
            s = s & vbNewLine & "<B>" & frmContent.Lang.GetString("_TypeMdl") & "</B> = [" & .ListItems.Item(x).SubItems(13) & "]<BR>"
            s = s & vbNewLine & "<B>" & frmContent.Lang.GetString("_SerialMdl") & "</B> = [" & .ListItems.Item(x).SubItems(14) & "]<BR>"
            s = s & vbNewLine & "<B>" & frmContent.Lang.GetString("_BytePerSecMdl") & "</B> = [" & .ListItems.Item(x).SubItems(15) & "]<BR>"
            s = s & vbNewLine & "<B>" & frmContent.Lang.GetString("_SecPerClustMdl") & "</B> = [" & .ListItems.Item(x).SubItems(16) & "]<BR>"
            s = s & vbNewLine & "<B>" & frmContent.Lang.GetString("_TrackPerCylMdl") & "</B> = [" & .ListItems.Item(x).SubItems(17) & "]<BR>"
            s = s & vbNewLine & "<B>" & frmContent.Lang.GetString("_SecPerTMdl") & "</B> = [" & .ListItems.Item(x).SubItems(18) & "]<BR>"
            s = s & vbNewLine & "<B>" & frmContent.Lang.GetString("_OffDepMdl") & "</B> = [" & .ListItems.Item(x).SubItems(19) & "]<BR>"
            s = s & vbNewLine & "<B>" & frmContent.Lang.GetString("_FileFormMdl") & "</B> = [" & .ListItems.Item(x).SubItems(20) & "]<BR>"
            s = s & vbNewLine & "<B>" & frmContent.Lang.GetString("_DriveTypeMdl") & "</B> = [" & .ListItems.Item(x).SubItems(21) & "]<BR>"
            
            s = s & vbNewLine & "<BR> <BR>"
            
        Next x
    End With
    
    s = s & vbNewLine & "</body>" & vbNewLine & "</html>"

    CreateMeHtmlString = s
End Function
