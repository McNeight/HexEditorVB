VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AfClsManifest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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


'   =======================================================----------------------------------
'           [AFCK]              Classe Manifest v1.0.3             25 nov 2005
'   =======================================================----------------------------------
'
'                   v1.0.1          23 nov  2005
'                   v1.0.2          24 nov  2005
'                                               ~ compatibilit� EXE Win2000 registre
'                   v1.0.3          25 nov  2005
'                                               ~ test version WinXP num�rique
'
'   =======================================================----------------------------------
'
'
'       Cette classe sert � cr�er un MonProgramme.exe.manifest afin d'appliquer
'       facilement le style XP aux composants de votre appli, et modifie le
'       mode de compatibilit� pour Windows2000
'
'       NB :    Bug connu de Kro$oft
'               les OptionsButtons doivent �tre plac�es dans des PictureBox
'               pour ne pas �tre sur fond noir.
'
'       ATTENTION :     Si vous utilisez "If App.PrevInstance Then End"
'                       le re-lancement automatique ne fonctionnera pas
'                       (premi�re ex�cution. donc manuel! et apr�s la 2�me
'                       non-plus mais l�, c'est votre volont�... ;) )
'
'   =======================================================----------------------------------
'   derni�re version :                     http://www.vbfrance.com/code.aspx?ID=34784
'   =======================================================----------------------------------


Option Explicit

' CONSTANTES
'
Private Const HKEY_CURRENT_USER     As Long = &H80000001
Private Const KEY_ALL_ACCESS        As Long = &H3F
Private Const ERROR_SUCCESS         As Long = 0&
Private Const REG_SZ                As Long = 1&
'
'
' API
'
'
'
'-> s'assure que la DLL "Common Control" est bien charg�e
Private Declare Sub InitCommonControls Lib "Comctl32" ()
'
'-> d�termine si un fichier ou un dossier existe
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" _
                        (ByVal pszPath As String) As Long
'
'-> supprime un fichier
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
'
'-> retourne des infos sur la plateforme syst�me
Private Declare Function GetVersion Lib "kernel32" () As Long
'
'-> ouvre la cl� sp�cifi�e
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
                        (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
                        ByVal samDesired As Long, phkResult As Long) As Long
'
'-> relache le handle de la cl� sp�cifi�e
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'
'->  cr�� le cl� sp�cifi�e. l'ouvre si elle existe d�j�
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
                        (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'
'-> cr�� ou modifie le type et la donn�e d'une valeur de cl� ouverte sp�cifi�e
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" _
                        (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
                        ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long
'
'-> retourne le type et la donn�e d'une valeur de cl� ouverte sp�cifi�e
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" _
                        (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
                        ByRef lpType As Long, ByVal szData As String, ByRef lpcbData As Long) As Long
'
'
'
Private Const cDestKey   As String = "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers"
Private Const cDestValue As String = "WIN2000"

'
'
'   *- RUN -*
Public Sub Run(Optional ByVal bReLoad As Boolean = True)
' bReLoad ->    si le fichier n'existe pas (premier lancement), on quitte cette instance
'               pour que l'exe prenne en compte le nouveau manifest. relance de l'exe?

'   on quitte si on est pas en EXE (pas utile puisque ne fonctionne qu'en compil�)
    If Not CBool(App.LogMode) Then Exit Sub

'   on quitte si pas XP
    'If Not (IsWindowsXP) Then Exit Sub

    Dim sPathFileExe As String, sPathFileManifest As String, FF As Long
'   on ajoute le Slash au App.Path (si besoin)
    sPathFileExe = App.Path
    If RightB$(sPathFileExe, 2) <> "\" Then sPathFileExe = sPathFileExe & "\"

'   on ajoute le nom du fichier � cr�er
    sPathFileExe = sPathFileExe & App.EXEName & ".exe"
    sPathFileManifest = sPathFileExe & ".manifest"

'   on appelle l'API
    Call InitCommonControls

'   on rend l'exe compatible � Windows2000. Certains PC sous XP (dont le mien) emp�chent le lancement de l'exe li� � un .manifest si pas compatible. et cette action est faite ici au cas o� le manifest serait fourni.
'   (inutile de le founir, mais on connait les loustiks ^^). Merci � RenField (Forum VbF�606392) pour avoir trouv� rapidement la ligne du registre. on utilise les API ; CreateObject("Wscript.Shell") serait sans doute plus pratique, mais les BONS anti-virus devraient emp�cher l'�criture...
    On Error Resume Next
'       on regarde si la cl� existe
        Dim lRet As Long, lKeyHandle As Long
        
        lRet = RegOpenKeyEx(HKEY_CURRENT_USER, cDestKey, 0&, KEY_ALL_ACCESS, lKeyHandle)
        If lRet = ERROR_SUCCESS Then
'           elle existe, on ferme la lecture
            Call RegCloseKey(lKeyHandle)
        Else
'           elle n'existe pas, on la cr��
            Call RegCreateKey(HKEY_CURRENT_USER, cDestKey, 0&)
        End If

'       on r�-ouvre la cl� + v�rif erreur registre (logiquement inutile.....)
        lRet = RegOpenKeyEx(HKEY_CURRENT_USER, cDestKey, 0&, KEY_ALL_ACCESS, lKeyHandle)
        If lRet <> ERROR_SUCCESS Then Exit Sub

'       on r�cup�re la valeur (peut ne pas exister)
        Dim sBuffer As String, sValue As String
        sBuffer = Space$(2048)
        Call RegQueryValueEx(lKeyHandle, sPathFileExe, 0&, REG_SZ, sBuffer, Len(sBuffer))
    
        sValue = LeftB$(sBuffer, 14)
        sBuffer = vbNullChar

'       si la cl� n'existe pas ou n'est pas bonne, on la cr��
        If sValue <> cDestValue Then
            sValue = cDestValue & vbNullChar
            Call RegSetValueEx(lKeyHandle, sPathFileExe, 0&, REG_SZ, sValue, Len(sValue))
        End If

'       on ferme
        Call RegCloseKey(lKeyHandle)
        
    On Error GoTo 0 ' *FGE*
    Err.Clear
    
'   on quitte si Manifest existe d�j�
    If CBool(PathFileExists(sPathFileManifest)) Then Exit Sub
    
'   on cr�� le fichier
    FF = FreeFile
    On Error GoTo ErrHandler
        Open sPathFileManifest For Output As #FF
            Print #FF, "<?xml version=" & Chr$(34) & "1.0" & Chr$(34) & " encoding=" & Chr$(34) & "UTF-8" & Chr$(34) & " standalone=" & Chr$(34) & "yes" & Chr$(34) & "?>"
            Print #FF, ""
            Print #FF, "<assembly xmlns=" & Chr$(34) & "urn:schemas-microsoft-com:asm.v1" & Chr$(34) & " manifestVersion=" & Chr$(34) & "1.0" & Chr$(34) & "> <assemblyIdentity"
            Print #FF, "    name=" & Chr$(34) & "Process Explorer" & Chr$(34)
            Print #FF, "    processorArchitecture=" & Chr$(34) & "x86" & Chr$(34)
            Print #FF, "    version=" & Chr$(34) & "1.0.0.0" & Chr$(34)
            Print #FF, "    type=" & Chr$(34) & "win32" & Chr$(34) & "/>"
            Print #FF, "<description>Process Explorer</description> <dependency>"
            Print #FF, "    <dependentAssembly>"
            Print #FF, "        <assemblyIdentity"
            Print #FF, "            type=" & Chr$(34) & "win32" & Chr$(34)
            Print #FF, "            name=" & Chr$(34) & "Microsoft.Windows.Common-Controls" & Chr$(34)
            Print #FF, "            version=" & Chr$(34) & "6.0.0.0" & Chr$(34)
            Print #FF, "            processorArchitecture=" & Chr$(34) & "x86" & Chr$(34)
            Print #FF, "            publicKeyToken=" & Chr$(34) & "6595b64144ccf1df" & Chr$(34)
            Print #FF, "            language=" & Chr$(34) & "*" & Chr$(34)
            Print #FF, "        />"
            Print #FF, "    </dependentAssembly>"
            Print #FF, "</dependency>"
            Print #FF, "</assembly>"
        Close #FF
    On Error GoTo 0 ' *FGE*
    DoEvents

'   on relance l'exe? ne fonctionnera pas si PrevInstance. (cf ATTENTION d�but module)
    If bReLoad Then Call Shell(sPathFileExe, vbNormalFocus)

'   on quitte cette instance
    End
    Exit Sub

ErrHandler:
    Err.Clear
    Call DeleteFile(sPathFileManifest)
End Sub
