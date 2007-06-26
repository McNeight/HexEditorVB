Attribute VB_Name = "Demarrage"
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
'//SUB DE DEMARRAGE DU PROGRAMME
'lecture des préférences
'application du style XP
'création des raccourcis dans explorer
'gestion du Command
'=======================================================

Public Const DEFAULT_INI As String = "[Appearance]" & vbNewLine & "BackGroundColor=16777215" & vbNewLine & "OffsetForeColor=16737380" & vbNewLine & "HexaForeColor=7303023" & vbNewLine & "StringsForeColor=7303023" & vbNewLine & "BaseForeColor=16737380" & vbNewLine & "TitleBackGroundColor=16777215" & vbNewLine & "LinesColor=-2147483636" & vbNewLine & "SelectionColor=14737632" & vbNewLine & "ModifiedItems=255" & vbNewLine & "SelectedItems=0" & vbNewLine & "BookMarkColor=8421631" & vbNewLine & "ModifiedSelectedItems=255" & vbNewLine & "Grid=0" & vbNewLine & "OffsetTitleForeColor=16737380" & vbNewLine & "OffsetsHex=1" & vbNewLine & "[Integration]" & vbNewLine & "FileContextual=1" & vbNewLine & "FolderContextual=1" & vbNewLine & "SendTo=1" & vbNewLine & "[General]" & vbNewLine & "DisplayExplore=1" & vbNewLine & "ShowAvert=1" & vbNewLine & "MaximizeWhenOpen=1" & vbNewLine & "DisplayIcon=1" & vbNewLine & "DisplayInfos=1" & vbNewLine & "DisplayData=1" & vbNewLine & "QuickBackup=1" & vbNewLine & "ResoX=640" & _
    vbNewLine & "ResoY=480" & vbNewLine & "AllowMultipleInstances=0" & vbNewLine & "DoNotChangeDates=1" & vbNewLine & "OpenSubFiles=0" & vbNewLine & "CloseHomeWhenChosen=0" & vbNewLine & "Splash=1" & vbNewLine & "FormBackColor=16377305" & vbNewLine & "MenuBackColor=16117739" & vbNewLine & "ToolbarPCT=1" & vbNewLine & "[Environnement]" & vbNewLine & "OS=1" & vbNewLine & "Lang=" & vbNewLine & "[Historique]" & vbNewLine & "NumberOfSave=0" & vbNewLine & "[FileExplorer]" & vbNewLine & "ShowPath=0" & vbNewLine & "ShowHiddenFiles=1" & vbNewLine & "ShowHiddenFolders=1" & vbNewLine & "ShowSystemFiles=1" & vbNewLine & "ShowSystemFodlers=1" & vbNewLine & "ShowROFiles=1" & vbNewLine & "ShowROFolders=1" & vbNewLine & "AllowMultipleSelection=1" & vbNewLine & "AllowFileSuppression=1" & vbNewLine & "AllowFolderSuppression=0" & vbNewLine & "IconType=1" & vbNewLine & _
    "DefaultPath=Dossier du programme" & vbNewLine & "Pattern=*.*" & vbNewLine & "Height=2200" & vbNewLine & "HideColumnTitle=0" & vbNewLine & "[Executable]" & vbNewLine & "HasCrashed=0" & vbNewLine & "[Console]" & vbNewLine & "BackColor=0" & vbNewLine & "ForeColor=12632256" & vbNewLine & "Heigth=1250" & vbNewLine & "Load=1"
Public Const HEX_EDITOR_VB_VERSION As String = "Hex Editor VB v1.6 pre-Alpha 1"

Public AfManifest As AfClsManifest   'classe appliquant le style XP
Public TempFiles() As String    'contient tout les fichiers temporaires
Public JailedProcess() As ProcessItem   'contient la liste de tous les processus bloqués
Public bAcceptBackup As Boolean 'variable qui détermine si la création d'un backup a été acceptée
Public clsERREUR As clsGetionErreur
Public cFile As FileSystemLibrary.FileSystem
Public cMem As clsMemoryRW
Public clsConv As clsConvert
Public cProc As clsProcess
Public clsPref As clsIniFile
Public cPref As clsIniPref
Public lNbChildFrm As Long
Public bEndSplash As Boolean
Public lngTimeLoad As Long
Public sLang() As String
Public Chr_(255) As String  'contient la liste des char, pour gagner en vitesse

'pour la sanitization
Public sH55() As Byte
Public sHAA() As Byte
Public pAA As Long
Public p55 As Long


'=======================================================
'//DEMARRAGE DU PROGRAMME
'=======================================================
Sub Main()
Dim Frm As Form
Dim sFile() As String
Dim m() As String
Dim x As Long
Dim y As Long
Dim S As String

    'change le path actuel pour permettre la reconnaissance de bnAlloc.dll
    #If FINAL_VERSION Then
        Call SetCurrentDirectoryA(App.Path)
    #Else
        Call SetCurrentDirectoryA(EXE_PATH)
    #End If

    'récupère le temps mis pour charger le logiciel
    #If Not (FINAL_VERSION) Then
        lngTimeLoad = GetTickCount
    #End If
    
    'On Error GoTo ErrGestion
    
    '//on quitte si déjà une instance
    If App.PrevInstance Then End
    
    
    '//vérifie la version de Windows
        x = GetWindowsVersion(S, y)
        If x <> [Windows Vista] And x <> [Windows XP] And x <> [Windows 2000] Then
            'OS non compatible
            MsgBox "Votre système d'exploitation est [" & S & "] build [" & Trim$(Str$(y)) & "]" & vbNewLine & "Ce logiciel n'est compatible qu'avec Windows XP et Windows Vista." & vbNewLine & "Hex Editor VB va donc se fermer", vbCritical, "Système d'exploitation non compatible"
            End
        End If
        
    
    '//applique le style XP (création d'un *.manifest si nécessaire)
        Set AfManifest = New AfClsManifest
        Call AfManifest.Run
        Set AfManifest = Nothing
        
    
    '//affiche des messages de warning si on n'a pas une version finale
        #If PRE_ALPHA_VERSION Then
            'version prealpha
            MsgBox "This file is a pre-alpha version, it means that functionnalities are missing and it may contains bugs." & vbNewLine & "This file is avalailable for testing purpose.", vbCritical, "Warning"
        #ElseIf BETA_VERSION Then
            'version beta
            MsgBox "This file is a beta version, it means that all principal functions are availables but there is still bugs." & vbNewLine & "This file is avalailable for testing purpose.", vbCritical, "Warning"
        #End If
        
    
    '//initialisation de la gestion des erreurs
        Set clsERREUR = New clsGetionErreur 'instancie la classe de gestion des erreurs
        'affecte les properties à la classe
        clsERREUR.LogFile = App.Path & "\ErrLog.log"
        clsERREUR.MakeSoundIDE = True
        
    
    '//instancie les classes
        Set cFile = New FileSystemLibrary.FileSystem
        Set cMem = New clsMemoryRW
        Set clsPref = New clsIniFile
        Set cPref = New clsIniPref
        Set cProc = New clsProcess
        Set clsConv = New clsConvert
    
    '//initialise les tableaux
        ReDim JailedProcess(0)  'contient les process bloqués
        ReDim TempFiles(0)  'contient les fichiers temporaires à supprimer au déchargement du logiciel
        
        'on remplit le tableau Chr_()
        For x = 0 To 255
            Chr_(x) = Chr$(x)
        Next x
    
    '// récupère la langue
        'liste les fichiers de langue
        ReDim sLang(0): ReDim sFile(0)
        
        If App.LogMode = 0 Then
            'IDE
            S = LANG_PATH
        Else
            S = App.Path & "\Lang"
        End If
        sFile() = cFile.EnumFilesStr(S, False)
        
        'vire les fichiers qui ne sont pas *.ini et French.ini
        For x = 1 To UBound(sFile())
            If LCase$(Right$(sFile(x), 4)) = ".ini" Then
                'c'est un fichier de langue
                ReDim Preserve sLang(UBound(sLang()) + 1)
                sLang(UBound(sLang())) = sFile(x)
            End If
        Next x
        
    '//récupère le fichier d'aide
        If cFile.FileExists(App.Path & "\Help.chm") Then _
            App.HelpFile = App.Path & "\Help.chm"
        
    
    '//récupère les préférences
        #If MODE_DEBUG Then
            'alors on est dans la phase Debug, donc on a le dossier du source
            clsPref.sDefaultPath = cFile.GetParentFolderName(App.Path) & "\Executable folder\Preferences\config.ini"
        #Else
            'alors c'est plus la phase debug, donc plus d'IDE possible
            clsPref.sDefaultPath = App.Path & "\Preferences\config.ini" 'détermine le fichier de config par défaut
        #End If
        
        If cFile.FileExists(clsPref.sDefaultPath) = False Then
            'le fichier de configuration est inexistant
            'il est necesasire de le crér (par défaut)
            Call cFile.CreateEmptyFile(clsPref.sDefaultPath, True)
            
            'remplit le fichier
            Call cFile.SaveDataInFile(clsPref.sDefaultPath, DEFAULT_INI, False)
        End If
         
        Set cPref = clsPref.GetIniFile
        cPref.IniFilePath = clsPref.sDefaultPath
        
        bEndSplash = False
        'affiche le splash si souhaité
        If cPref.general_Splash Then
            frmSplash.Show
            DoEvents    '/!\ DO NOT REMOVE (permet d'afficher le splash screen correctement)
        End If
        
        frmSplash.lblState.Caption = "Configuration des options..."
        'détermine si le programme a crashé ou pas
        If cPref.exe_HasCrashed = 1 Then
            'alors on sort d'un crash ==> informe
            MsgBox "Le programme n'a pas été fermé correctement, il récupère probablement d'une erreur critique." & vbNewLine & "Merci de me contacter par mail en précisant le contexte et les causes, si possible, du crash." & vbNewLine & "Vous pouvez me contacter en cliquant sur 'Hex Editor VB sur Internet' dans le menu d'aide." & vbNewLine & "Vous pouvez également envoyer le rapport d'erreur (menu Aide ==> rapport d'erreur)." & vbNewLine & "Merci de votre contribution.", vbCritical + vbOKOnly, "Erreur critique lors de la précédente fermeture"
        End If
        'affecte la valeur True au crash
        cPref.exe_HasCrashed = 1
        'sauvegarde les pref (met à jour la valeur)
        Call clsPref.SaveIniFile(cPref) '//CHANGER CA ET NE SAUVER QUE LA VARIABLE CRASH
         
         
        frmSplash.lblState.Caption = "Génération de l'intégration dans Explorer..."
        'créé le raccourci 'envoyer vers...'
        'Shortcut True
        'ajoute au menu contextuel de windows les entrées de HexEditor
        'AddContextMenu 1    'fichiers
        ' AddContextMenu 0    'dossiers
         
        'ajout du type de fichier *.hescr à HexEditor VB.exe
        Call Reg_HESCR_file
        
    
        frmSplash.lblState.Caption = "Lancement du logiciel..."

    
    
    '//créé le tableau contenant la liste des commandes pour l'éditeur de script
        Call GetSplit
    
    
    '//Ouvre chaque fichier désigné par le path (gestion du Command)
        If Len(Command) > 0 Then
            'alors on ouvre un fichier/dossier (celui lancé avec Command)
           
            If InStrRev(Command, "shredd", , vbBinaryCompare) Then
                'alors on ouvrira la form de suppression si il y a l'argument shredd à la fin
                If Right$(Command, 8) = Chr_(34) & "shredd" & Chr_(34) Then
                    'alors c'est bon ==> suppression form
                    
                    ReDim sFile(0)   'contiendra les paths
           
                    'sépare Command en plusieurs path
                    Call SplitString(Chr_(34), Command, sFile())
                    
                    'affiche la form
                    frmShredd.Show
                    
                    For x = 1 To UBound(sFile())
                        'teste l'existence de chaque path
                    
                        If cFile.FileExists(sFile(x)) Then
                            'ouvre un fichier
                            frmShredd.LV.ListItems.Add Text:=sFile(x)
                        ElseIf cFile.FolderExists(sFile(x)) Then
                            'ouvre un dossier - liste les fichiers
                            m() = cFile.EnumFilesStr(sFile(x))
                            If UBound(m()) <> 0 Then
                                'les ouvre un par un
                                For y = 1 To UBound(m)
                                    If cFile.FileExists(m(y)) Then
                                        frmShredd.LV.ListItems.Add sFile(m(y))
                                        DoEvents
                                    End If
                                Next y
                            End If
                        End If
                    Next x
                End If
            ElseIf InStrRev(Command, "date", , vbBinaryCompare) Then
                If Right$(Command, 6) = Chr_(34) & "date" & Chr_(34) Then
                    'alors c'est bon ==> date form
                    
                    
                        MsgBox "date"
                        
                End If
            ElseIf InStrRev(Command, "viewfile", , vbBinaryCompare) Then
                If Right$(Command, 10) = Chr_(34) & "viewfile" & Chr_(34) Then
                    'alors c'est bon ==> visualise le fichier en mode File
                    
                    ReDim sFile(0)   'contiendra les paths
           
                    'sépare Command en plusieurs path
                    Call SplitString(Chr_(34), Command, sFile())
                    
                    For x = 1 To UBound(sFile())
                        'teste l'existence de chaque path
                    
                        If cFile.FileExists(sFile(x)) Then
                            'ouvre un fichier
                            Set Frm = New Pfm
                            Call Frm.GetFile(sFile(x))
                            Frm.Show
                        ElseIf cFile.FolderExists(sFile(x)) Then
                            'ouvre un dossier - liste les fichiers
                            m() = cFile.EnumFilesStr(sFile(x))
                            If UBound(m()) <> 0 Then
                                'les ouvre un par un
                                For y = 1 To UBound(m)
                                    If cFile.FileExists(m(y)) Then
                                        Set Frm = New Pfm
                                        Call Frm.GetFile(m(x))
                                        Frm.Show
                                        lNbChildFrm = lNbChildFrm + 1
                                        frmContent.Sb.Panels(2).Text = "Ouvertures=[" & CStr(lNbChildFrm) & "]"
                                        Set Frm = Nothing
                                        DoEvents
                                    End If
                                Next y
                             End If
                        End If
                    Next x
            
                End If
            Else
                'alors on ouvre normalement
                
                 ReDim sFile(0)   'contiendra les paths
        
                 'sépare Command en plusieurs path
                 Call SplitString(Chr_(34), Command, sFile())
                 
                 For x = 1 To UBound(sFile())
                     'teste l'existence de chaque path
                     If cFile.FileExists(sFile(x)) Then
                         'ouvre un fichier
                         Set Frm = New Pfm
                         Call Frm.GetFile(sFile(x))
                         Frm.Show
                     ElseIf cFile.FolderExists(sFile(x)) Then
                        'ouvre un dossier - liste les fichiers
                        m() = cFile.EnumFilesStr(sFile(x))
                        If UBound(m()) <> 0 Then                             'les ouvre un par un
                             For y = 1 To UBound(m)
                                 If cFile.FileExists(m(y)) Then
                                     Set Frm = New Pfm
                                     Call Frm.GetFile(m(x))
                                     Frm.Show
                                     lNbChildFrm = lNbChildFrm + 1
                                     frmContent.Sb.Panels(2).Text = "Ouvertures=[" & CStr(lNbChildFrm) & "]"
                                     Set Frm = Nothing
                                     DoEvents
                                 End If
                             Next y
                        End If
                     End If
                 Next x
            End If
                
            
        Else
            'pas de Command
            frmContent.Show
            
            
            'on récupère l'état dans lequel le logiciel était en partant
            If cPref.general_QuickBackup Then
            
                'alors on vérifie l'existence du fichier
                'If cFile.FileExists(App.Path & "\Preferences\QuickBackup.ini") = False Then Exit Sub
                
                'charge les données
                Call LoadQuickBackupINIFile

            End If
            
        End If

    Exit Sub
ErrGestion:
    clsERREUR.AddError "Demarrage.Main", True
End Sub

'=======================================================
'termine le programme
'=======================================================
Public Sub EndProgram()
Dim x As Long

    '//prévient des processus bloqués
        If UBound(JailedProcess()) > 0 Then
            'alors des processus bloqués
            If MsgBox(frmContent.Lang.GetString("_ProcessHaveBeenBlocked"), vbInformation + vbYesNo, frmContent.Lang.GetString("_War")) <> vbYes Then
                
                'alors on libère tout
                For x = 1 To UBound(JailedProcess())
                    cProc.ResumeProcess (JailedProcess(x).th32ProcessID)
                Next x
            End If
        End If
        

    '//supprime les fichiers temporaires de TempFiles
        For x = 1 To UBound(TempFiles())
            Call cFile.DeleteFile(TempFiles(x))
        Next x
    
    '//libère les classes
        Set clsERREUR = Nothing
        Set cFile = Nothing
        Set cMem = Nothing
        Set cProc = Nothing
        Set clsConv = Nothing
    
    '//affecte la valeur False au crash (car si on est là, c'est que c'est bien fermé)
        cPref.exe_HasCrashed = 0
        
        'sauvegarde les pref (met à jour la valeur)
        Call clsPref.SaveIniFile(cPref)
        
        'libère les dernières classes
        Set cPref = Nothing
        Set clsPref = Nothing
        
    DoEvents
    
    End 'quitte
End Sub

'=======================================================
'charge les données qui permettent de rendre le logiciel dans l'état dans lequel on a quitté
'=======================================================
Private Sub LoadQuickBackupINIFile()
Dim x As Long
Dim s2() As String
Dim s3 As String
Dim S As String
Dim s4() As String
Dim bIsOk As Long
Dim Frm As Form
Dim lFrom As Currency
Dim lTo As Currency

    On Error Resume Next
    
    'extrait la première ligne qui détermine le type de form à ouvrir
    Dim s8 As String
    S = cFile.LoadFileInString(App.Path & "\Preferences\QuickBackup.ini", bIsOk)
    
    If bIsOk = False Then Exit Sub  'fichier inacessible en lecture (ou inexistant)
    
    'extrait la première ligne
    s2() = Split(S, vbNewLine, , vbBinaryCompare) ' Left$(s, InStr(1, s, vbNewLine) - 1)
    s3 = Right$(s2(0), Len(s2(0)) - InStr(1, S, "|"))   'contient le PID, le disque ou le fichier
    
    With frmContent.Lang
        Select Case Left$(s2(0), 2)
            Case "Pr"
                'processus
                If cProc.DoesPIDExist(Val(s3)) = False Then
                    'process inexistant
                    MsgBox .GetString("_ProcessWithPID") & " " & s3 & " " & .GetString("_ProcessDoesNotE"), vbCritical, .GetString("_Error")
                    Exit Sub
                End If
                Set Frm = New MemPfm
                Call Frm.GetFile(Val(s3))   'le PID en paramètre
            Case "Di"
                'disque
                If cFile.FolderExists(s3) = False Or Len(s3) <> 3 Then
                    'disque inexistant
                    MsgBox .GetString("_TheDisk") & " " & s3 & " " & .GetString("_DoesNotEx"), vbCritical, .GetString("_Error")
                    Exit Sub
                End If
                If cFile.IsDriveAvailable(Left$(s3, 1)) = False Then
                    'disque inaccessible
                    MsgBox .GetString("_TheDisk") & " " & s3 & " " & .GetString("_IsNotAccessibl"), vbCritical, .GetString("_Error")
                    Exit Sub
                End If
                Set Frm = New diskPfm
                Call Frm.GetDrive(s3)
            Case "Fi"
                If cFile.FileExists(s3) = False Then
                    'fichier invalide
                    MsgBox .GetString("_TheFile") & " " & s3 & " " & .GetString("_IsNotValid"), vbCritical, .GetString("_Error")
                    Exit Sub
                End If
                'fichier
                Set Frm = New Pfm
                Call Frm.GetFile(s3)
            Case "Ph"
                'disque physique
                If cFile.IsPhysicalDiskAvailable(Val(s3)) = False Then
                    'disque inaccessible
                    MsgBox .GetString("_TheDisk") & " " & s3 & " " & .GetString("_NotAccessOrInex"), vbCritical, .GetString("_Error")
                    Exit Sub
                End If
                Set Frm = New physPfm
                Call Frm.GetDrive(Val(s3))   'le numéro du disque en paramètre
            Case Else
                'fichier non valide (trafiqué)
                Exit Sub
        End Select
        
        'affiche la form
        Frm.Show
        lNbChildFrm = lNbChildFrm + 1
        frmContent.Sb.Panels(2).Text = .GetString("_Openings") & CStr(lNbChildFrm) & "]"
    End With
    
    DoEvents    '/!\ IMPORTANT DO NOT REMOVE
    
    'extrait la seconde ligne (qui contient la sélection et le VS.Value)
    s4() = Split(s2(1), "|", , vbBinaryCompare)
    
    If UBound(s4()) <> 9 Then Exit Sub  'fichier corrompu
    
    With frmContent.ActiveForm.HW
    
        .FirstOffset = Val(s4(7))

        'change le VS.Value et refresh le HW
        lFrom = Val(s4(0)) + Val(s4(1))
        lTo = Val(s4(2)) + Val(s4(3)) + (s4(3) <> "1") '-1 pour corriger la valeur erronnée si Col>1
        
        'donne le focus au HW et positionne à la bonne place
        .Item.Offset = Val(s4(5))
        .Item.Col = Val(s4(6))
        .Item.Line = Val(s4(8))
        .Item.tType = Val(s4(9))
        Call frmContent.ActiveForm.HW_MouseDown(1, 0, 1, 1, .Item)
        
        'sélectionne la zone désirée
        .SelectZone 16 - (By16(lFrom) - lFrom), By16(lFrom) - 16, 17 - (By16(lTo) - lTo), _
            By16(lTo) - 16
        frmContent.ActiveForm.VS.Value = Val(s4(4))
        Call frmContent.ActiveForm.VS_Change(frmContent.ActiveForm.VS.Value)
        Call frmContent.ActiveForm.cmdMAJ_Click 'MAJ du fichier et de la sélection, Offset courant...

        .Refresh
    End With
    
    'extrait les signets et les ajoute
    
            
    
End Sub

'=======================================================
'sauve les données qui permettent de rendre le logiciel dans l'état dans lequel on a quitté
'=======================================================
Public Sub SaveQuickBackupINIFile()
Dim S As String
Dim x As Long

    If cPref.general_QuickBackup Then
        'on lance la sauvegarde de plusieurs choses : type de form, fichier/disque/processus
        'zone de sélection et signets éventuels
        
        'créé la string à enregistrer
        If Not (frmContent.ActiveForm Is Nothing) Then
        
            With frmContent.ActiveForm
                'sauvegarde le type de form et le path (ou PID) correspondant
                Select Case TypeOfForm(frmContent.ActiveForm)
                    Case "Processus"
                        S = "Process|" & Trim$(Str$(.Tag))
                    Case "Disque"
                        S = "Disk|" & Right$(.Caption, 3)
                    Case "Fichier"
                        S = "File|" & .Caption
                    Case "Disque physique"
                        S = "Phys|" & Trim$(Str$(.Tag))
                End Select
            
                'maintenant on sauve la zone sélectionnée et la valeur du VS
                S = S & vbNewLine & Trim$(Str$(.HW.FirstSelectionItem.Offset)) & "|" & _
                    Trim$(Str$(.HW.FirstSelectionItem.Col)) & "|" & _
                    Trim$(Str$(.HW.SecondSelectionItem.Offset)) & "|" & _
                    Trim$(Str$(.HW.SecondSelectionItem.Col)) & "|" & Trim$(Str$(.VS.Value)) & "|" & _
                    Trim$(Str$(.HW.Item.Offset)) & "|" & Trim$(Str$(.HW.Item.Col)) & "|" & _
                    Trim$(Str$(.HW.FirstOffset)) & "|" & Trim$(Str$(.HW.Item.Line)) & "|" & _
                    Trim$(Str$(.HW.Item.tType))
                
                'maintenant on sauvegarde tous les signets
                For x = 1 To .lstSignets.ListItems.Count
                    S = S & vbNewLine & .lstSignets.ListItems.Item(x) & "|" & .lstSignets.ListItems.Item(x).SubItems(1)
                Next x
            End With
            
            'lance la sauvegarde
            Call cFile.SaveDataInFile(App.Path & "\Preferences\QuickBackup.ini", S, True)
        Else
            'on delete le fichier
            Call cFile.DeleteFile(App.Path & "\Preferences\QuickBackup.ini")
        End If
    End If
End Sub
