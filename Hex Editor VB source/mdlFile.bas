Attribute VB_Name = "mdlFile"
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
'//MODULE DE GESTION DES FICHIERS
'=======================================================

'=======================================================
'obtient le path du fichier temp � cr�er
'=======================================================
Public Sub ObtainTempPathFile(ByVal sFile As String, ByRef sTempFile As String, _
    sExt As String)
    
Dim sBuf As String
Dim s As String

    '//obtient le r�pertoire temporaire
    'buffer
    sBuf = String$(256, vbNullChar)
    'obtient le dossier temp
    Call GetTempPath(256, sBuf)
    'formate le path
    sBuf = Left$(sBuf, InStr(sBuf, vbNullChar) - 1)
    
    '//obtient un path unique
    'buffer
    s = String$(256, vbNullChar)
    'obtient le dossier temp
    Call GetTempFileName(sBuf, sFile, 0, s)
    'formate le path
    s = Left$(s, InStr(s, vbNullChar) - 1)
    
    'ajoute l'extension
    s = s & "." & sExt
    
    'ajoute le fichier � la liste des fichiers temp
    ReDim Preserve TempFiles(UBound(TempFiles) + 1)
    TempFiles(UBound(TempFiles)) = sBuf
    
    sTempFile = s
    
End Sub

'=======================================================
'enregistrer sous ==> lance la cr�ation d'un fichier
'=======================================================
Public Function CreateAFile(Frm As Form, ByVal sPath As String) As Boolean
Dim sFile As String
Dim lFile As Long
    
    'sauvegarde le fichier
    Call cFile.DeleteFile(sPath)
    
    'cr�� ke fichier
    Call Frm.GetNewFile(sPath)

End Function

'=======================================================
'execute un fichier temporaire cr�� � partir
'des valeurs modifi�es du tableau hexad�cimal
'=======================================================
Public Function ExecuteTempFile(ByVal hWnd As Long, Frm As Form, sExt As String) As Long
Dim sTempFile As String

    'obtient un path temporaire
    Call ObtainTempPathFile("to_execute", sTempFile, sExt)
    
    'cr�� le fichier
    Call CreateAFile(Frm, sTempFile)
     
    'l'ex�cute
    ExecuteTempFile = cFile.ShellOpenFile(sTempFile, hWnd)
End Function

'=======================================================
'cr�� un raccourci dans 'envoyer vers...'
'=======================================================
Public Sub Shortcut(ByVal bCreate As Boolean)
Dim WSHShell As Object, Sh As Variant
Dim sPath As String

    If bCreate Then
        'le cr��
        sPath = cFile.GetSpecialFolder(CSIDL_SENDTO) 'contient le nom du path du shortcut
        
        Set WSHShell = CreateObject("Wscript.Shell")
        
        ' Cr�ation d'un objet raccourci sur le Bureau
        Set Sh = WSHShell.CreateShortcut(sPath & "\HexEditor.lnk")
        With Sh
            .TargetPath = App.Path & "\HexEditor.exe"
            .WorkingDirectory = WSHShell.ExpandEnvironmentStrings("%windir%")
            .WindowStyle = 4
            .IconLocation = WSHShell.ExpandEnvironmentStrings(App.Path & "\HexEditor.exe,0")
            .Save
        End With
    Else
        'le supprime
        sPath = cFile.GetSpecialFolder(CSIDL_SENDTO) 'contient le nom du path du shortcut
        
        Call cFile.DeleteFile(sPath & "\HexEditor.lnk")
    End If

End Sub

'=======================================================
'impression du fichier de l'activeform
'=======================================================
Public Sub PrintFile(ByVal curStartOffset As Currency, ByVal curEndOffset As Currency, _
ByVal bPrintHexa As Boolean, ByVal bPrintASCII As Boolean, ByVal bPrintOffset As Boolean, _
ByVal bPrintFileInfo As Boolean, ByVal lngTextSize As Long, ByVal tPrinter As Printer, Optional ByVal strTitle As String)

Dim x As Long
Dim y As Long

    Set Printer = tPrinter
    
    With Printer
    
        'd�finit la police
        .FontName = "courier"
        .FontSize = lngTextSize
        
        If curStartOffset < 0 Then curStartOffset = 0
        If curEndOffset > frmContent.ActiveForm.HW.MaxOffset Then curEndOffset = frmContent.ActiveForm.HW.MaxOffset
        
        Printer.Print vbNewLine & vbNewLine
        
        'proc�de � l'impression
        For x = By16(curStartOffset) To By16(curEndOffset) Step 16
        
            'offset
            .CurrentX = 300
            .ForeColor = frmContent.ActiveForm.HW.OffsetForeColor
            y = .CurrentY
            Printer.Print FormatedAdress(x)
            
            'valeurs hexa
            .CurrentX = 3000: .CurrentY = y
            .ForeColor = frmContent.ActiveForm.HW.HexForeColor
            Printer.Print "0H 45 12 E7 AA 12 35 00 00 FB 4F 7E 81 0D 38 11"
        Next x
        
        'fin de l'impression
        .EndDoc
    End With
    
End Sub

'=======================================================
'r�cup�rer des bytes dans un fichier
'=======================================================
Public Function GetBytesFromFile(ByVal sFile As String, ByVal curSize As _
    Currency, ByVal curOffset As Currency) As String
    
Dim tmpText As String
Dim Ret As Long
Dim lFile As Long
    
    'obtient un handle vers le fichier � ouvrir
    lFile = CreateFile(sFile, GENERIC_READ, FILE_SHARE_READ Or _
        FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    
    'fichier inexistant, ou en tout cas inaccessible
    If lFile = INVALID_HANDLE_VALUE Then Exit Function
        
    'cr�� un buffer qui contiendra les valeurs
    tmpText = String$(curSize, 0)

    'bouge le pointeur sur lr fichier au bon emplacement
    Ret = SetFilePointerEx(lFile, curOffset / 10000, 0&, FILE_BEGIN)  'divise par 10000 pour
    'pouvoir renvoyer une currency DECIMALE (cad du genre 1.4567 pour l'offset 14567)
    
    'prend un morceau du fichier
    Ret = ReadFile(lFile, ByVal tmpText, Len(tmpText), Ret, ByVal 0&)
    
    'affecte � la fonction
    GetBytesFromFile = tmpText
    
    'referme le handle
    Call CloseHandle(lFile)
    
End Function

'=======================================================
'�crire des bytes dans un fichier
'=======================================================
Public Function WriteBytesToFile(ByVal sFile As String, ByVal sString As _
    String, ByVal curOffset As Currency) As String
    
Dim tmpText As String
Dim Ret As Long
Dim lFile As Long
       
    'obtient un handle vers le fichier � �crire
    'ouverture en ECRITURE, avec overwrite si d�j� existant (car d�j� demand� confirmation avant)
    lFile = CreateFile(sFile, GENERIC_WRITE, FILE_SHARE_READ Or _
        FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)

    If lFile = INVALID_HANDLE_VALUE Then Exit Function 'fichier indisponible
    
    'bouge le pointeur sur le fichier au bon emplacement
    Ret = SetFilePointerEx(lFile, curOffset / 10000, 0&, FILE_BEGIN)
    'a divis� par 10^4 pour obtenir un nombre d�cimal de Currency

    '�criture dans le fichier
    Call WriteFile(lFile, ByVal sString, Len(sString), Ret, ByVal 0&)
    
    'ferme le handle du fichier �crit
    Call CloseHandle(lFile)
    
End Function

'=======================================================
'�crire des bytes dans un fichier
'=======================================================
Public Sub WriteBytesToFileHandle(ByVal hFile As Long, ByVal pt As Long, _
    ByVal curOffset As Currency, ByVal lLen As Long)
    
Dim tmpText As String
Dim Ret As Long
    
    'bouge le pointeur sur le fichier au bon emplacement
    Ret = SetFilePointerEx(hFile, curOffset / 10000, 0&, FILE_BEGIN)

    '�criture dans le fichier
    Call WriteFile(hFile, ByVal pt, lLen, Ret, ByVal 0&)
    
End Sub

'=======================================================
'�crire des bytes dans un fichier (� la fin du fichier)
'=======================================================
Public Function WriteBytesToFileEnd(ByVal sFile As String, ByVal sString As _
    String) As String
    
Dim tmpText As String
Dim Ret As Long
Dim lFile As Long
       
    'obtient un handle vers le fichier � �crire
    'ouverture en ECRITURE, avec overwrite si d�j� existant (car d�j� demand� confirmation avant)
    lFile = CreateFile(sFile, GENERIC_WRITE, FILE_SHARE_READ Or _
        FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)

    If lFile = INVALID_HANDLE_VALUE Then Exit Function 'fichier non dispo
    
    'bouge le pointeur sur le fichier � la fin du fichier
    Ret = SetFilePointerEx(lFile, 0&, 0&, FILE_END) '

    '�criture dans le fichier
    Call WriteFile(lFile, ByVal sString, Len(sString), Ret, ByVal 0&)
    
    'ferme le handle du fichier �crit
    Call CloseHandle(lFile)

End Function

'=======================================================
'r�cup�re un handle d'�criture vers un fichier
'=======================================================
Public Function GetFileHandleWrite(ByVal sFile As String) As Long
    GetFileHandleWrite = CreateFile(sFile, GENERIC_WRITE, FILE_SHARE_READ Or _
        FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
End Function

'=======================================================
'�crire des bytes dans un fichier (� la fin du fichier) avec en entr�e un handle
'=======================================================
Public Function WriteBytesToFileEndHandle(ByVal lngFile As Long, ByVal _
    sString As String) As String
    
Dim tmpText As String
Dim Ret As Long
        
    'bouge le pointeur sur le fichier � la fin du fichier
    Ret = SetFilePointerEx(lngFile, 0&, 0&, FILE_END) '

    '�criture dans le fichier
    Call WriteFile(lngFile, ByVal sString, Len(sString), Ret, ByVal 0&)

End Function
