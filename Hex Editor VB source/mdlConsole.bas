Attribute VB_Name = "mdlConsole"
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
'//MODULE DE GESTION DE LA CONSOLE
'=======================================================

Public strConsoleText() As String    'contient les diff�rentes commances entr�es 1 � Ubound
Public lngConsolePos As Long  'num�ro de la commande de l'historique

'=======================================================
'refresh la liste des anciennes commandes en ajoutant la derni�re string
'=======================================================
Public Sub AddTextToConsole(ByVal sText As String)
    If sText = "Cls" Then Exit Sub
    With frmContent.txt
        .Text = .Text & IIf(Len(.Text) > 0, vbNewLine, vbNullString) & sText
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelColor = cPref.console_ForeColor
        .SelStart = Len(.Text)
    End With
End Sub

'=======================================================
'permet de r�cup�rer les strings depuis le fichier *.ini
'=======================================================
Private Function GetHelp(ByVal sSection As String) As String
Dim s As String
Dim l As Long
Dim l2 As Long

    On Error Resume Next

    'r�cup�re le contenu du fichier
    #If MODE_DEBUG Then
        s = cFile.LoadFileInString("C:\HEX EDITOR VB\Executable folder\ConsoleHelp.ini")
    #Else
        s = cFile.LoadFileInString(App.Path & "\ConsoleHelp.ini")
    #End If
    
    'r�cup�re la position de la section
    l = InStr(1, s, sSection, vbBinaryCompare)
    
    'r�cup�re la position du premier '|' apr�s la section
    l2 = InStr(l + 1, s, "|", vbBinaryCompare)
    
    GetHelp = Mid$(s, l + Len(sSection) + 1, l2 - l - Len(sSection) - 1)
    
End Function

'=======================================================
'lance la commande valid�e
'=======================================================
Public Sub LaunchCommand()
Dim s As String
Dim s2 As String
Dim Frm As Form
Dim sF() As String
Dim x As Long
Dim sPref As String

    'commence par ajouter � la liste la commande entr�e
    ReDim Preserve strConsoleText(UBound(strConsoleText()) + 1)
    strConsoleText(UBound(strConsoleText())) = frmContent.txtE.Text
    Call AddTextToConsole(frmContent.txtE.Text)
    
    'ex�cute la commande
    With frmContent.Lang
        s2 = .GetString("_InvalidCommand")
        
        'r�cup�re le chemin des pr�f�rences
        #If MODE_DEBUG Then
            sPref = "C:\Hex Editor VB\Executable Folder\Preferences\"
        #Else
            sPref = App.Path & "\Preferences\"
        #End If
        
        '//EXECUTION DES DIFFERENTES COMMANDES
        s = LCase$(frmContent.txtE.Text)
        If s = "help" Then
            'on affiche l'aide
            s2 = "Pour plus d'informations sur une commandes, tappez Help [commande]"
            s2 = s2 & vbNewLine & "ABOUT  A propos"
            s2 = s2 & vbNewLine & "BOOKMARK  Gestion des signet" 'TODO
            s2 = s2 & vbNewLine & "BUGREPORT  Affiche le rapport d'erreurs"
            s2 = s2 & vbNewLine & "CALC  Lance la calculatrice"
            s2 = s2 & vbNewLine & "CLOSE  Ferme les fen�tres ouvertes" 'TODO
            s2 = s2 & vbNewLine & "CLS  Efface la console"
            s2 = s2 & vbNewLine & "CONVERT  Afficher la fen�tre de conversion" 'TODO
            s2 = s2 & vbNewLine & "COPY  Copier la s�lection" 'TODO
            s2 = s2 & vbNewLine & "CUT  Couper la s�lection" 'TODO
            s2 = s2 & vbNewLine & "SEARCHFILE  Affiche la recherche de fichier"
            s2 = s2 & vbNewLine & "KILL  Supprimer un fichier"
            s2 = s2 & vbNewLine & "LICENSE  Affiche la license"
            s2 = s2 & vbNewLine & "MOVE  Effectue un d�placement de la vue" 'TODO
            s2 = s2 & vbNewLine & "NEWFILE  Cr�er un fichier" 'TODO
            s2 = s2 & vbNewLine & "OPEN  Ouvrir un fichier, processus ou disque" 'TODO
            s2 = s2 & vbNewLine & "OPTIONS  Affiche les options"
            s2 = s2 & vbNewLine & "PASTE  Coller la s�lection" 'TODO
            s2 = s2 & vbNewLine & "PRINT  Lancer une impression"
            s2 = s2 & vbNewLine & "PROCESS  Gestion des processus" 'TODO
            s2 = s2 & vbNewLine & "PROPERTY  Afficher les propri�t�s"
            s2 = s2 & vbNewLine & "QUIT  Quitter le programme"
            s2 = s2 & vbNewLine & "REDO  Refaire"
            s2 = s2 & vbNewLine & "REFRESH  Rafraichit les valeurs hexa"
            s2 = s2 & vbNewLine & "REPLACE  Remplacer" 'TODO
            s2 = s2 & vbNewLine & "RESETCONFIG  R�initialise la configuration"
            s2 = s2 & vbNewLine & "RUN  Affiche la boite de dialogue 'Ex�cuter'"
            s2 = s2 & vbNewLine & "SCRIPT  D�marre l'�diteur de script"
            s2 = s2 & vbNewLine & "SEARCH  Effectuer une recherche"
            s2 = s2 & vbNewLine & "SELECT  Effectuer une s�lection" 'TODO
            s2 = s2 & vbNewLine & "SHELL  Lance une commande"
            s2 = s2 & vbNewLine & "SHOWCOMMANDS  Affiche la liste des commandes de la console"
            s2 = s2 & vbNewLine & "START  D�marre une t�che"
            s2 = s2 & vbNewLine & "STAT  Affiche les statistiques du fichier"
            s2 = s2 & vbNewLine & "UNDO  D�faire"
            s2 = s2 & vbNewLine & "VERSION  Affiche la version du logiciel"
        ElseIf Left$(s, 5) = "help " And Len(s) > 5 Then
            'affiche l'aide de la commande
            s2 = GetHelp(s)
            If s2 = vbNullString Then s2 = "Aide non support�e"
        ElseIf s = "about" Then frmAbout.Show vbModal: s2 = "Fen�tre d'A propos affich�e"
        ElseIf s = "start" Then frmHome.Show: PremierPlan frmHome, MettreAuPremierPlan: s2 = "Fen�tre de d�marrage de t�che affich�e"
        ElseIf s = "quit" Then frmContent.mnuExit_Click: s2 = "Quitte le programme..."
        ElseIf s = "property" Then frmPropertyShow.Show: s2 = "Fen�tre de propri�t�s affich�e"
        ElseIf s = "print" Then frmPrint.Show vbModal: s2 = "Fen�tre d'impression affich�e"
        ElseIf s = "bugreport" Then frmLogErr.Show vbModal: s2 = "Fen�tre de log affich�e"
        ElseIf s = "calc" Then frmContent.mnuCalc_Click: s2 = "Calculatrice lanc�e"
        ElseIf s = "searchfile" Then frmFileSearch.Show: s2 = "Fen�tre de recherche de fichiers affich�e"
        ElseIf s = "stat" Then frmContent.mnuStats_Click: s2 = "Fen�tre de statistiques affich�e"
        ElseIf s = "search" Then frmSearch.Show: s2 = "Fen�tre de recherche affich�e"
        ElseIf s = "script" Then frmScript.Show: s2 = "Editeur de script lanc�"
        ElseIf s = "refresh" Then frmContent.mnuRefreh_Click: s2 = "Rafraichissement de la vue effectu�e"
        ElseIf s = "options" Then frmOptions.Show vbModal: s2 = "Fen�tre d'options affich�e"
        ElseIf s = "redo" Then frmContent.mnuRedo_Click: s2 = "Commande 'Redo' lanc�e"
        ElseIf s = "undo" Then frmContent.mnuUndo_Click: s2 = "Commande 'Undo' lanc�e"
        ElseIf Left$(s, 7) = "convert" Then
            If InStr(7, s, "-a", vbBinaryCompare) Then
                'alors c'est le convertisseur avanc�
                frmAdvancedConversion.Show
                s2 = "Fen�tre de conversion avanc�e affich�e"
            Else
                'converion simple
                frmConvert.Show
                s2 = "Fen�tre de conversion simple affich�e"
            End If
        ElseIf s = vbNullString Then
            s2 = vbNewLine
        ElseIf s = "license" Then Call cFile.ShellOpenFile(App.Path & "\license.txt", frmContent.hWnd): s2 = "License affich�e"
        ElseIf s = "cls" Then frmContent.txt.Text = vbNullString: s2 = "Cls"
        ElseIf s = "shell" Then s2 = "Param�tre manquant"
        ElseIf Left$(s, 6) = "shell " Then
            Shell Right$(s, Len(s) - 6), vbNormalFocus
            s2 = "Commance lanc�e"
        ElseIf Left$(s, 5) = "kill " Then
            If cFile.FileExists(Right$(s, Len(s) - 5)) = False Then
                s2 = "Fichier inexistant"
            Else
                cFile.DeleteFile Right$(s, Len(s) - 5)
                s2 = IIf(cFile.FileExists(Right$(s, Len(s) - 5)), "Fichier encore existant", "Fichier effac�")
            End If
        ElseIf Left$(s, 5) = "close" Then
            If InStr(1, s, "-a") Then
                'alors on ferme toutes les fen�tres
                If frmContent.ActiveForm Is Nothing Then
                    s2 = "Aucune fen�tre � fermer"
                Else
                    For Each Frm In Forms
                        If (TypeOf Frm Is Pfm) Or (TypeOf Frm Is diskPfm) Or (TypeOf Frm Is MemPfm) Or (TypeOf Frm Is physPfm) Then
                            SendMessage Frm.hWnd, WM_CLOSE, 0, 0
                        End If
                    Next Frm
                        
                    '/!\ NE PAS ENLEVER
                    '/!\ BUG NON RESOLU
                    '/!\ Apr�s d�chargement des form (juste en haut), des form nomm�es "Form1" (caption par
                    'd�faut) subsistent
                    For Each Frm In Forms
                        If Frm.Caption = "Form1" Then SendMessage Frm.hWnd, WM_CLOSE, 0, 0
                    Next Frm
                    s2 = "Toutes les fen�tres ont �t� ferm�es"
                End If
            Else
                'juste l'active
                If frmContent.ActiveForm Is Nothing Then
                    s2 = "Aucune fen�tre � fermer"
                Else
                
                    SendMessage frmContent.ActiveForm.hWnd, WM_CLOSE, 0, 0
        
                    '/!\ NE PAS ENLEVER
                    '/!\ BUG NON RESOLU
                    '/!\ Apr�s d�chargement des form (juste en haut), des form nomm�es "Form1" (caption par
                    'd�faut) subsistent
                    For Each Frm In Forms
                        If Frm.Caption = "Form1" Then SendMessage Frm.hWnd, WM_CLOSE, 0, 0
                    Next Frm
                    s2 = "Fen�tre ferm�e"
                End If
            End If
        ElseIf s = "cut" Then Call frmContent.mnuCut_Click: s2 = "S�lection coup�e"
        ElseIf s = "sourceforge" Then Call frmContent.mnuSourceForge_Click: s2 = "Ouverture de la page SourceForge du projet"
        ElseIf s = "vbfrance" Then Call frmContent.mnuVbfrance_Click: s2 = "Ouverture de la page VBfrance de l'auteur"
        ElseIf s = "options.ini" Or s = "config.ini" Or s = "config" Then
            Call cFile.ShellOpenFile(sPref & "config.ini", frmContent.hWnd)
            s2 = "Fichier d'options ouvert"
        ElseIf s = "console.ini" Or s = "consolehelp.ini" Or s = "showcommands" Then
            Call cFile.ShellOpenFile(App.Path & "\ConsoleHelp.ini", frmContent.hWnd)
            s2 = "Fichier d'aide de la console ouvert"
        ElseIf s = "version" Then MsgBox "Version " & Trim$(Str$(App.Major)) & "." & Trim$(Str$(App.Minor)) & "." & Trim$(Str$(App.Revision)), vbInformation + vbOKOnly, "Hex Editor VB": s2 = vbNullString
        ElseIf s = "test" Then MsgBox "Test", vbCritical, "Test": s2 = vbNullString
        ElseIf Left$(s, 7) = "msgbox " Then
            MsgBox Right$(s, Len(s) - 7), vbInformation + vbOKOnly, "Hex Editor VB Console Message"
            s2 = vbNullString
        ElseIf s = "maximize" Then frmContent.WindowState = vbMaximized: s2 = "Form minimize"
        ElseIf s = "minimize" Then frmContent.WindowState = vbMinimized: s2 = "Form maximize"
        ElseIf s = "resize" Then Call frmContent.MDIForm_Resize: s2 = "Form resize"
        ElseIf s = "run" Then ShowRunBox (frmContent.hWnd): s2 = "RunBox affich�e"
        ElseIf Left$(s, 12) = "resetconfig " Then
            s2 = "Param�tres invalides"
            If InStr(12, s, "-o") Then
                'reset les options
                            
                'r�cr�� le fichier
                cFile.SaveDataInFile sPref & "Config.ini", DEFAULT_INI, True
                
                'reloade les options
                Set cPref = clsPref.GetIniFile(sPref & "Config.ini")
                Call MAJoptions
                
                s2 = "R�initialisation effectu�e"
            End If
            If InStr(12, s, "-w") Then
                'reset les *.ini des forms
                
                '�num�re tous les fichiers *.ini
                sF() = cFile.EnumFilesStr(sPref, False)
                
                'vire tous les fichiers
                For x = 1 To UBound(sF())
                    If LCase$(sF(x)) <> "config.ini" Then
                        cFile.DeleteFile sF(x)
                    End If
                Next x
                
                s2 = "R�initialisation effectu�e"
            End If
        ElseIf s = "resetconfig" Then s2 = "Param�tre manquant"
        ElseIf s = "mailauthor" Then Shell "MailTo:hexeditorvb@gmail.com": s2 = "Ouverture du logiciel de messagerie"
        ElseIf s = "aboutthisfile" Then Call cFile.ShellOpenFile(App.Path & "\PLEASE READ ME (eng + fr).TXT", frmContent.hWnd): s2 = vbNullString
        
        
        End If
    End With


    'affiche � la console le r�sultat
    Call AddTextToConsole(s2)
End Sub

'=======================================================
'r�cup�re la commande correspondant � la position dans l'historique
'=======================================================
Public Function GetCommand() As String
    On Error Resume Next
    GetCommand = strConsoleText(lngConsolePos)
End Function

'=======================================================
'met � jour les options dans toutes les form
'=======================================================
Private Sub MAJoptions()
Dim x As Form

    For Each x In Forms
        If (TypeOf x Is Pfm) Or (TypeOf x Is diskPfm) Or (TypeOf x Is MemPfm) Or (TypeOf x Is physPfm) Then

                With x.HW
                    'on applique ces couleurs au HW de CETTE form
                    .BackColor = cPref.app_BackGroundColor
                    .OffsetForeColor = cPref.app_OffsetForeColor
                    .HexForeColor = cPref.app_HexaForeColor
                    .StringForeColor = cPref.app_StringsForeColor
                    .OffsetTitleForeColor = cPref.app_OffsetTitleForeColor
                    .BaseTitleForeColor = cPref.app_BaseForeColor
                    .TitleBackGround = cPref.app_TitleBackGroundColor
                    .LineColor = cPref.app_LinesColor
                    .SelectionColor = cPref.app_SelectionColor
                    .ModifiedItemColor = cPref.app_ModifiedItems
                    .ModifiedSelectedItemColor = cPref.app_ModifiedSelectedItems
                    .SignetColor = cPref.app_BookMarkColor
                    .Grid = cPref.app_Grid
                    .UseHexOffset = CBool(cPref.app_OffsetsHex)
                    .Refresh
                End With
                
                'change les Visible des frames de toutes les forms active
                x.FrameData.Visible = CBool(cPref.general_DisplayData)
                x.FrameInfos.Visible = CBool(cPref.general_DisplayInfos)
                If (TypeOf x Is diskPfm) Or (TypeOf x Is physPfm) Then x.FrameInfo2.Visible = CBool(cPref.general_DisplayInfos)
            'End If
        End If
    Next x
              
    On Error Resume Next
    
    'on change la taille du Explorer
    frmContent.pctExplorer.Height = cPref.explo_Height
    frmContent.LV.Height = cPref.explo_Height - 145
    
    'apparence de la console
    frmContent.pctConsole.BackColor = cPref.console_BackColor
    frmContent.txt.BackColor = cPref.console_BackColor
    frmContent.txtE.BackColor = cPref.console_BackColor
    frmContent.pctConsole.Height = cPref.console_Heigth
    With frmContent.txt
        .BackColor = cPref.console_BackColor
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelColor = cPref.console_ForeColor
        .SelStart = Len(.Text)
    End With
    With frmContent.txtE
        .BackColor = cPref.console_BackColor
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelColor = cPref.console_ForeColor
        .SelStart = Len(.Text)
    End With

    'cr�� ou supprime les menus contextuels de Windows en fonction des nouvelles prefs.
    If CBool(cPref.integ_FileContextual) = False Then
        'enl�ve
        RemoveContextMenu 1
    Else
        'ajoute
        AddContextMenu 1
    End If
    If CBool(cPref.integ_FolderContextual) = False Then
        'enl�ve
        RemoveContextMenu 0
    Else
        'ajoute
        AddContextMenu 0
    End If
    
    'cr�� ou pas le raccourci
    Shortcut CBool(cPref.integ_SendTo)


    'change les settings du Explorer
    With frmContent.LV
        .ShowEntirePath = CBool(cPref.explo_ShowPath)
        .ShowHiddenDirectories = CBool(cPref.explo_ShowHiddenFolders)
        .ShowHiddenFiles = CBool(cPref.explo_ShowHiddenFiles)
        .ShowSystemDirectories = CBool(cPref.explo_ShowSystemFodlers)
        .ShowSystemFiles = CBool(cPref.explo_ShowSystemFiles)
        .ShowReadOnlyDirectories = CBool(cPref.explo_ShowROFolders)
        .ShowReadOnlyFiles = CBool(cPref.explo_ShowROFiles)
        .AllowMultiSelect = CBool(cPref.explo_AllowMultipleSelection)
        .AllowFileDeleting = CBool(cPref.explo_AllowFileSuppression)
        .Pattern = cPref.explo_Pattern
        .HideColumnHeaders = CBool(cPref.explo_HideColumnTitle)
        Select Case cPref.explo_IconType
            Case 0
                .DisplayIcons = BasicIcons
            Case 1
                .DisplayIcons = FileIcons
            Case 2
                .DisplayIcons = NoIcons
        End Select
    End With
    
    Call frmContent.MDIForm_Resize
End Sub
