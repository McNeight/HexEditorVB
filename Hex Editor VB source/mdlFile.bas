Attribute VB_Name = "mdlFile"
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
'//MODULE DE GESTION DES FICHIERS
'-------------------------------------------------------

'-------------------------------------------------------
'obtient le path du fichier temp à créer
'-------------------------------------------------------
Public Sub ObtainTempPathFile(ByVal sFile As String, ByRef sTempFile As String, sExt As String)
Dim sBuf As String
Dim s As String

    '//obtient le répertoire temporaire
    'buffer
    sBuf = String$(256, vbNullChar)
    'obtient le dossier temp
    GetTempPath 256, sBuf
    'formate le path
    sBuf = Left$(sBuf, InStr(sBuf, vbNullChar) - 1)
    
    '//obtient un path unique
    'buffer
    s = String$(256, vbNullChar)
    'obtient le dossier temp
    GetTempFileName sBuf, sFile, 0, s
    'formate le path
    s = Left$(s, InStr(s, vbNullChar) - 1)
    
    'ajoute l'extension
    s = s & "." & sExt
    
    'ajoute le fichier à la liste des fichiers temp
    ReDim Preserve TempFiles(UBound(TempFiles) + 1)
    TempFiles(UBound(TempFiles)) = sBuf
    
    sTempFile = s
    
End Sub

'-------------------------------------------------------
'enregistrer sous ==> lance la création d'un fichier
'-------------------------------------------------------
Public Function CreateAFile(Frm As Form, ByVal sPath As String) As Boolean
Dim sFile As String
Dim lFile As Long

    On Error GoTo GestionErr
    
    'sauvegarde le fichier
    cFile.KillFile sPath
    
    'créé ke fichier
    Call Frm.GetNewFile(sPath)

GestionErr:
End Function

'-------------------------------------------------------
'execute un fichier temporaire créé à partir
'des valeurs modifiées du tableau hexadécimal
'-------------------------------------------------------
Public Function ExecuteTempFile(ByVal hwnd As Long, Frm As Form, sExt As String) As Long
Dim sTempFile As String

    'obtient un path temporaire
    ObtainTempPathFile "to_execute", sTempFile, sExt
    
    'créé le fichier
    CreateAFile Frm, sTempFile
     
    'l'exécute
    ExecuteTempFile = ShellExecute(hwnd, "open", sTempFile, vbNullString, vbNullString, 1)
End Function

'-------------------------------------------------------
'créé un raccourci dans 'envoyer vers...'
'-------------------------------------------------------
Public Sub Shortcut(ByVal bCreate As Boolean)
Dim WSHShell As Object, Sh As Variant
Dim sPath As String

    If bCreate Then
        'le créé
        sPath = cFile.GetSpecialFolder(CSIDL_SENDTO) 'contient le nom du path du shortcut
        
        Set WSHShell = CreateObject("Wscript.Shell")
        
        ' Création d'un objet raccourci sur le Bureau
        Set Sh = WSHShell.CreateShortcut(sPath & "\HexEditor.lnk")
        Sh.TargetPath = App.Path & "\HexEditor.exe"
        Sh.WorkingDirectory = WSHShell.ExpandEnvironmentStrings("%windir%")
        Sh.WindowStyle = 4
        Sh.IconLocation = WSHShell.ExpandEnvironmentStrings(App.Path & "\HexEditor.exe,0")
        Sh.Save
    Else
        'le supprime
        sPath = cFile.GetSpecialFolder(CSIDL_SENDTO) 'contient le nom du path du shortcut
        
        cFile.KillFile sPath & "\HexEditor.lnk"
    End If

End Sub

'-------------------------------------------------------
'impression du fichier de l'activeform
'-------------------------------------------------------
Public Sub PrintFile(ByVal curStartOffset As Currency, ByVal curEndOffset As Currency, _
ByVal bPrintHexa As Boolean, ByVal bPrintASCII As Boolean, ByVal bPrintOffset As Boolean, _
ByVal bPrintFileInfo As Boolean, ByVal lngTextSize As Long, ByVal tPrinter As Printer, Optional ByVal strTitle As String)

Dim x As Long
Dim y As Long

    Set Printer = tPrinter
    
    With Printer
    
        'définit la police
        .FontName = "courier"
        .FontSize = lngTextSize
        
        If curStartOffset < 0 Then curStartOffset = 0
        If curEndOffset > frmContent.ActiveForm.HW.MaxOffset Then curEndOffset = frmContent.ActiveForm.HW.MaxOffset
        
        Printer.Print vbNewLine & vbNewLine
        
        'procède à l'impression
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

'-------------------------------------------------------
'récupérer des bytes dans un fichier
'-------------------------------------------------------
Public Function GetBytesFromFile(ByVal sFile As String, ByVal curSize As Currency, ByVal curOffset As Currency) As String
Dim tmpText As String
Dim Ret As Long
Dim lFile As Long

    On Error GoTo ErrGestion
    
    'obtient un handle vers le fichier à ouvrir
    lFile = CreateFile(sFile, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    
    If lFile = 0 Then
        'fichier inexistant, ou en tout cas inaccessible
        Exit Function
    End If
        
    'créé un buffer qui contiendra les valeurs
    tmpText = String$(curSize, 0)

    'bouge le pointeur sur lr fichier au bon emplacement
    Ret = SetFilePointerEx(lFile, curOffset / 10000, 0&, FILE_BEGIN)  'divise par 10000 pour
    'pouvoir renvoyer une currency DECIMALE (cad du genre 1.4567 pour l'offset 14567)
    
    'prend un morceau du fichier
    Ret = ReadFile(lFile, ByVal tmpText, Len(tmpText), Ret, ByVal 0&)
    
    'affecte à la fonction
    GetBytesFromFile = tmpText
    
    Exit Function
ErrGestion:
    clsERREUR.AddError "mdlFile.GetBytesFromFile", True
End Function

'-------------------------------------------------------
'écrire des bytes dans un fichier
'-------------------------------------------------------
Public Function WriteBytesToFile(ByVal sFile As String, ByVal sString As String, ByVal curOffset As Currency) As String
Dim tmpText As String
Dim Ret As Long
Dim lFile As Long

    On Error GoTo ErrGestion
       
    'obtient un handle vers le fichier à écrire
    'ouverture en ECRITURE, avec overwrite si déjà existant (car déjà demandé confirmation avant)
    lFile = CreateFile(sFile, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)

    'bouge le pointeur sur le fichier au bon emplacement
    Ret = SetFilePointerEx(lFile, curOffset / 10000, 0&, FILE_BEGIN)
    'a divisé par 10^4 pour obtenir un nombre décimal de Currency

    'écriture dans le fichier
    WriteFile lFile, ByVal sString, Len(sString), Ret, ByVal 0&
    
    'ferme le handle du fichier écrit
    CloseHandle lFile

    Exit Function
ErrGestion:
    clsERREUR.AddError "mdlFile.WriteBytesToFile", True
End Function

'-------------------------------------------------------
'écrire des bytes dans un fichier (à la fin du fichier)
'-------------------------------------------------------
Public Function WriteBytesToFileEnd(ByVal sFile As String, ByVal sString As String) As String
Dim tmpText As String
Dim Ret As Long
Dim lFile As Long

    On Error GoTo ErrGestion
       
    'obtient un handle vers le fichier à écrire
    'ouverture en ECRITURE, avec overwrite si déjà existant (car déjà demandé confirmation avant)
    lFile = CreateFile(sFile, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)

    'bouge le pointeur sur le fichier à la fin du fichier
    Ret = SetFilePointerEx(lFile, 0&, 0&, FILE_END) '

    'écriture dans le fichier
    WriteFile lFile, ByVal sString, Len(sString), Ret, ByVal 0&
    
    'ferme le handle du fichier écrit
    CloseHandle lFile

    Exit Function
ErrGestion:
    clsERREUR.AddError "mdlFile.WriteBytesToFileEnd", True
End Function
