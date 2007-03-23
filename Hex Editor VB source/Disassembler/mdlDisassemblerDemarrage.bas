Attribute VB_Name = "mdlDisassemblerDemarrage"
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
'MODULE DE DEMARRAGE
'=======================================================


'=======================================================
'VARIABLES PUBLIQUES
'=======================================================
Public cFile As clsFileInfos
Private AfManifest As AfClsManifest
Public tmpDir As String


'=======================================================
'//DEMARRAGE DU PROGRAMME
'=======================================================
Sub Main()
Dim s As String
Dim x As WINDOWS_VERSION
Dim y As Long

    
    '//vérifie la version de Windows
        x = GetWindowsVersion(s, y)
        If x <> [Windows Vista] And x <> [Windows XP] Then
            'OS non compatible
            MsgBox "Votre système d'exploitation est [" & s & "] build [" & Trim$(Str$(y)) & "]" & vbNewLine & "Ce logiciel n'est compatible qu'avec Windows XP et Windows Vista." & vbNewLine & "Hex Editor VB va donc se fermer", vbCritical, "Système d'exploitation non compatible"
            End
        End If
    
    '//applique le style XP (création d'un *.manifest si nécessaire)
        Set AfManifest = New AfClsManifest
        AfManifest.Run
        Set AfManifest = Nothing
    
    '//affiche des messages de warning si on n'a pas une version finale
        #If PRE_ALPHA_VERSION Then
            'version prealpha
            MsgBox "This file is a pre-alpha version, it means that functionnalities are missing and it may contains bugs." & vbNewLine & "This file is avalailable for testing purpose.", vbCritical, "Warning"
        #ElseIf BETA_VERSION Then
            'version beta
            MsgBox "This file is a beta version, it means that all principal functions are availables but there is still bugs." & vbNewLine & "This file is avalailable for testing purpose.", vbCritical, "Warning"
        #End If
    
    '//instancie les classes
        Set cFile = New clsFileInfos
    

    '//affiche la form principale
        frmDisAsm.Show
    
    '//gère le Command si nécessaire
        If Len(Command) > 0 Then
            Call frmDisAsm.DisAsmFile(Mid$(Command, 2, Len(Command) - 2))
        End If
    
End Sub


'=======================================================
'quitte le programme
'=======================================================
Public Sub EndProgram()
Dim sF() As String
Dim x As Long
Dim cp As String

    On Error Resume Next
    
    'décharge la form principale
    Unload frmDisAsm
    
    
    'vire tous les fichiers temp et le dossier temp
    ReDim sF(0)
    cp = cFile.GetParentDirectory(tmpDir)
    
    '//VERIFIE QUE L'ON KILL BIEN DES FICHIERS D'UN SOUS DOSSIER DU DOSSIER TEMP
    If Left$(cp, Len(cp) - 1) <> ObtainTempPath Then GoTo DONOTKILL
    
    'liste
    Call cFile.EnumFilesFromFolder(tmpDir, sF(), True)
    For x = 1 To UBound(sF())
        Call cFile.KillFile(sF(x))  'delete
    Next x
    Call RmDir(tmpDir)
    
    
DONOTKILL:

    'libère les classes
    Set cFile = Nothing

    'quitte
    End
End Sub
