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

Private Const DEFAULT_INI = "[Appearance]" & vbNewLine & "BackGroundColor=16777215" & vbNewLine & "OffsetForeColor=16737380" & vbNewLine & "HexaForeColor=7303023" & vbNewLine & "StringsForeColor=7303023" & vbNewLine & "BaseForeColor=16737380" & vbNewLine & "TitleBackGroundColor=16777215" & vbNewLine & "LinesColor=-2147483636" & vbNewLine & "SelectionColor=14737632" & vbNewLine & "ModifiedItems=255" & vbNewLine & "SelectedItems=0" & vbNewLine & "BookMarkColor=8421631" & vbNewLine & "ModifiedSelectedItems=255" & vbNewLine & "Grid=0" & vbNewLine & "OffsetTitleForeColor=16737380" & vbNewLine & "OffsetsHex=1" & vbNewLine & "[Integration]" & vbNewLine & "FileContextual=1" & vbNewLine & "FolderContextual=1" & vbNewLine & "SendTo=1" & vbNewLine & "[General]" & vbNewLine & "MaximizeWhenOpen=1" & vbNewLine & "DisplayIcon=1" & vbNewLine & "DisplayInfos=1" & vbNewLine & "DisplayData=1" & vbNewLine & "QuickBackup=1" & vbNewLine & "ResoX=640" & _
    vbNewLine & "ResoY=480" & vbNewLine & "AllowMultipleInstances=0" & vbNewLine & "DoNotChangeDates=1" & vbNewLine & "OpenSubFiles=0" & vbNewLine & "CloseHomeWhenChosen=0" & vbNewLine & "Splash=1" & vbNewLine & "[Environnement]" & vbNewLine & "OS=1" & vbNewLine & "Lang=" & vbNewLine & "[Historique]" & vbNewLine & "NumberOfSave=0" & vbNewLine & "[FileExplorer]" & vbNewLine & "ShowPath=0" & vbNewLine & "ShowHiddenFiles=1" & vbNewLine & "ShowHiddenFolders=1" & vbNewLine & "ShowSystemFiles=1" & vbNewLine & "ShowSystemFodlers=1" & vbNewLine & "ShowROFiles=1" & vbNewLine & "ShowROFolders=1" & vbNewLine & "AllowMultipleSelection=1" & vbNewLine & "AllowFileSuppression=1" & vbNewLine & "AllowFolderSuppression=0" & vbNewLine & "IconType=1" & vbNewLine & "DefaultPath=Dossier du programme" & vbNewLine & "Pattern=*.*" & vbNewLine & "Height=2200" & vbNewLine & "HideColumnTitle=0" & vbNewLine & "[Executable]" & vbNewLine & "HasCrashed=0"

Public AfManifest As AfClsManifest   'classe appliquant le style XP
Public TempFiles() As String    'contient tout les fichiers temporaires
Public JailedProcess() As ProcessItem   'contient la liste de tous les processus bloqués
Public bAcceptBackup As Boolean 'variable qui détermine si la création d'un backup a été acceptée
Public clsERREUR As clsGetionErreur
Public cFile As clsFileInfos
Public cMem As clsMemoryRW
Public cProc As clsProcess
Public cDisk As clsDiskInfos
Public clsPref As clsIniFile
Public cPref As clsIniPref
Public lNbChildFrm As Long
Public bEndSplash As Boolean


'=======================================================
'//DEMARRAGE DU PROGRAMME
'=======================================================
Sub Main()
Dim Frm As Form
Dim sFile() As String
Dim m() As String
Dim X As Long
Dim y As Long
Dim s As String

    On Error GoTo ErrGestion
    

    
    '//vérifie la version de Windows
        X = GetWindowsVersion(s, y)
        If X <> [Windows Vista] And X <> [Windows XP] Then
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
    
    '//initialisation de la gestion des erreurs
        Set clsERREUR = New clsGetionErreur 'instancie la classe de gestion des erreurs
        'affecte les properties à la classe
        clsERREUR.LogFile = App.Path & "\ErrLog.log"
        clsERREUR.MakeSoundIDE = True
    
    '//instancie les classes
        Set cFile = New clsFileInfos
        Set cMem = New clsMemoryRW
        Set cDisk = New clsDiskInfos
        Set clsPref = New clsIniFile
        Set cPref = New clsIniPref
        Set cProc = New clsProcess
    
    '//initialise les tableaux
        ReDim JailedProcess(0)  'contient les process bloqués
        ReDim TempFiles(0)  'contient les fichiers temporaires à supprimer au déchargement du logiciel

    
    '//récupère les préférences
         #If MODE_DEBUG Then
             'alors on est dans la phase Debug, donc on a le dossier du source
             clsPref.sDefaultPath = cFile.GetParentDirectory(App.Path) & "\Executable folder\Preferences\config.ini"
         #Else
             'alors c'est plus la phase debug, donc plus d'IDE possible
             clsPref.sDefaultPath = App.Path & "\Preferences\config.ini" 'détermine le fichier de config par défaut
         #End If
         
         If cFile.FileExists(clsPref.sDefaultPath) = False Then
             'le fichier de configuration est inexistant
             'il est necesasire de le crér (par défaut)
             cFile.CreateEmptyFile clsPref.sDefaultPath, True
             
             'remplit le fichier
             cFile.SaveStringInfile clsPref.sDefaultPath, DEFAULT_INI, False
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
                If Right$(Command, 8) = Chr$(34) & "shredd" & Chr$(34) Then
                    'alors c'est bon ==> suppression form
                    
                    ReDim sFile(0)   'contiendra les paths
           
                    'sépare Command en plusieurs path
                    SplitString Chr$(34), Command, sFile()
                    
                    'affiche la form
                    frmShredd.Show
                    
                    For X = 1 To UBound(sFile())
                        'teste l'existence de chaque path
                    
                        If cFile.FileExists(sFile(X)) Then
                            'ouvre un fichier
                            frmShredd.LV.ListItems.Add Text:=sFile(X)
                        ElseIf cFile.FolderExists(sFile(X)) Then
                            'ouvre un dossier - liste les fichiers
                            If cFile.EnumFilesFromFolder(sFile(X), m) <> 0 Then
                                'les ouvre un par un
                                For y = 1 To UBound(m)
                                    If cFile.FileExists(m(y)) Then
                                        frmShredd.LV.ListItems.Add sFile(m(y))
                                        DoEvents
                                    End If
                                Next y
                            End If
                        End If
                    Next X
                End If
            ElseIf InStrRev(Command, "date", , vbBinaryCompare) Then
                If Right$(Command, 6) = Chr$(34) & "date" & Chr$(34) Then
                    'alors c'est bon ==> date form
                    
                    
                        MsgBox "date"
                        
                End If
            ElseIf InStrRev(Command, "viewfile", , vbBinaryCompare) Then
                If Right$(Command, 10) = Chr$(34) & "viewfile" & Chr$(34) Then
                    'alors c'est bon ==> visualise le fichier en mode File
                    
                    ReDim sFile(0)   'contiendra les paths
           
                    'sépare Command en plusieurs path
                    SplitString Chr$(34), Command, sFile()
                    
                    For X = 1 To UBound(sFile())
                        'teste l'existence de chaque path
                    
                        If cFile.FileExists(sFile(X)) Then
                            'ouvre un fichier
                            Set Frm = New Pfm
                            Call Frm.GetFile(sFile(X))
                            Frm.Show
                        ElseIf cFile.FolderExists(sFile(X)) Then
                            'ouvre un dossier - liste les fichiers
                            If cFile.EnumFilesFromFolder(sFile(X), m) <> 0 Then
                                'les ouvre un par un
                                For y = 1 To UBound(m)
                                    If cFile.FileExists(m(y)) Then
                                        Set Frm = New Pfm
                                        Call Frm.GetFile(m(X))
                                        Frm.Show
                                        lNbChildFrm = lNbChildFrm + 1
                                        frmContent.Sb.Panels(2).Text = "Ouvertures=[" & CStr(lNbChildFrm) & "]"
                                        Set Frm = Nothing
                                        DoEvents
                                    End If
                                Next y
                             End If
                        End If
                    Next X
            
                End If
            Else
                'alors on ouvre normalement
                
                 ReDim sFile(0)   'contiendra les paths
        
                 'sépare Command en plusieurs path
                 SplitString Chr$(34), Command, sFile()
                 
                 For X = 1 To UBound(sFile())
                     'teste l'existence de chaque path
                     If cFile.FileExists(sFile(X)) Then
                         'ouvre un fichier
                         Set Frm = New Pfm
                         Call Frm.GetFile(sFile(X))
                         Frm.Show
                     ElseIf cFile.FolderExists(sFile(X)) Then
                         'ouvre un dossier - liste les fichiers
                         If cFile.EnumFilesFromFolder(sFile(X), m) <> 0 Then
                             'les ouvre un par un
                             For y = 1 To UBound(m)
                                 If cFile.FileExists(m(y)) Then
                                     Set Frm = New Pfm
                                     Call Frm.GetFile(m(X))
                                     Frm.Show
                                     lNbChildFrm = lNbChildFrm + 1
                                     frmContent.Sb.Panels(2).Text = "Ouvertures=[" & CStr(lNbChildFrm) & "]"
                                     Set Frm = Nothing
                                     DoEvents
                                 End If
                             Next y
                        End If
                     End If
                 Next X
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
Dim X As Long

    '//prévient des processus bloqués
        If UBound(JailedProcess()) > 0 Then
            'alors des processus bloqués
            If MsgBox("Ces processus on été bloqués, voulez vous quitter Hex Editor VB sans les débloquer ?", vbInformation + vbYesNo, "Attention") <> vbYes Then
                
                'alors on libère tout
                For X = 1 To UBound(JailedProcess())
                    cProc.ResumeProcess (JailedProcess(X).th32ProcessID)
                Next X
            End If
        End If
        

    '//supprime les fichiers temporaires de TempFiles
        For X = 1 To UBound(TempFiles())
            cFile.KillFile TempFiles(X)
        Next X
    
    '//libère les classes
        Set clsERREUR = Nothing
        Set cFile = Nothing
        Set cMem = Nothing
        Set cDisk = Nothing
        Set cProc = Nothing
    
    '//affecte la valeur False au crash (car si on est là, c'est que c'est bien fermé)
        cPref.exe_HasCrashed = 0
        
        'sauvegarde les pref (met à jour la valeur)
        clsPref.SaveIniFile cPref
        
        'libère les dernières classes
        Set cPref = Nothing
        Set clsPref = Nothing
        
    
    End 'quitte
End Sub

'=======================================================
'charge les données qui permettent de rendre le logiciel dans l'état dans lequel on a quitté
'=======================================================
Private Sub LoadQuickBackupINIFile()
Dim X As Long
Dim s2() As String
Dim s3 As String
Dim s As String
Dim s4() As String
Dim bIsOk As Long
Dim Frm As Form
Dim lFrom As Currency
Dim lTo As Currency

    On Error Resume Next
    
    'extrait la première ligne qui détermine le type de form à ouvrir
    Dim s8 As String
    s = cFile.LoadFileInString(App.Path & "\Preferences\QuickBackup.ini", bIsOk)
    
    If bIsOk = False Then Exit Sub  'fichier inacessible en lecture (ou inexistant)
    
    'extrait la première ligne
    s2() = Split(s, vbNewLine, , vbBinaryCompare) ' Left$(s, InStr(1, s, vbNewLine) - 1)
    s3 = Right$(s2(0), Len(s2(0)) - InStr(1, s, "|"))   'contient le PID, le disque ou le fichier
    
    Select Case Left$(s2(0), 1)
        Case "P"
            'processus
            Set Frm = New MemPfm
            Call Frm.GetFile(Val(s3))   'le PID en paramètre
        Case "D"
            'disque
            Set Frm = New diskPfm
            Call Frm.GetDrive(s3)
        Case "F"
            'fichier
            Set Frm = New Pfm
            Call Frm.GetFile(s3)
        Case Else
            'fichier non valide (trafiqué)
            Exit Sub
    End Select
    
    'affiche la form
    Frm.Show
    lNbChildFrm = lNbChildFrm + 1
    frmContent.Sb.Panels(2).Text = "Ouvertures=[" & CStr(lNbChildFrm) & "]"
    
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
Dim s As String
Dim X As Long

    If cPref.general_QuickBackup Then
        'on lance la sauvegarde de plusieurs choses : type de form, fichier/disque/processus
        'zone de sélection et signets éventuels
        
        'créé la string à enregistrer
        If Not (frmContent.ActiveForm Is Nothing) Then
        
            With frmContent.ActiveForm
                'sauvegarde le type de form et le path (ou PID) correspondant
                Select Case TypeOfForm(frmContent.ActiveForm)
                    Case "Processus"
                        s = "Process|" & Trim$(Str$(.Tag))
                    Case "Disque"
                        s = "Disk|" & Right$(.Caption, 3)
                    Case "Fichier"
                        s = "File|" & .Caption
                End Select
            
                'maintenant on sauve la zone sélectionnée et la valeur du VS
                s = s & vbNewLine & Trim$(Str$(.HW.FirstSelectionItem.Offset)) & "|" & _
                    Trim$(Str$(.HW.FirstSelectionItem.Col)) & "|" & _
                    Trim$(Str$(.HW.SecondSelectionItem.Offset)) & "|" & _
                    Trim$(Str$(.HW.SecondSelectionItem.Col)) & "|" & Trim$(Str$(.VS.Value)) & "|" & _
                    Trim$(Str$(.HW.Item.Offset)) & "|" & Trim$(Str$(.HW.Item.Col)) & "|" & _
                    Trim$(Str$(.HW.FirstOffset)) & "|" & Trim$(Str$(.HW.Item.Line)) & "|" & _
                    Trim$(Str$(.HW.Item.tType))
                
                'maintenant on sauvegarde tous les signets
                For X = 1 To .lstSignets.ListItems.Count
                    s = s & vbNewLine & .lstSignets.ListItems.Item(X) & "|" & .lstSignets.ListItems.Item(X).SubItems(1)
                Next X
            End With
            
            'lance la sauvegarde
            cFile.SaveStringInfile App.Path & "\Preferences\QuickBackup.ini", s, True
        Else
            'on delete le fichier
            cFile.KillFile App.Path & "\Preferences\QuickBackup.ini"
        End If
    End If
End Sub
