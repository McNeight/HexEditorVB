Attribute VB_Name = "Demarrage"
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
'//SUB DE DEMARRAGE DU PROGRAMME
'lecture des préférences
'application du style XP
'création des raccourcis dans explorer
'gestion du Command
'-------------------------------------------------------


Public AfManifest As New AfClsManifest  'classe appliquant le style XP
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

'-------------------------------------------------------
'//DEMARRAGE DU PROGRAMME
'-------------------------------------------------------
Sub Main()
Dim Frm As Form
Dim sFile() As String
Dim m() As String
Dim x As Long
Dim y As Long
Dim s As String

    On Error GoTo ErrGestion
    

    
    'vérifie la version de Windows
    x = GetWindowsVersion(s, y)
    If x <> [Windows Vista] And x <> [Windows XP] Then
        'OS non compatible
        MsgBox "Votre système d'exploitation est [" & s & "] build [" & Trim$(Str$(y)) & "]" & vbNewLine & "Ce logiciel n'est compatible qu'avec Windows XP et Windows Vista." & vbNewLine & "Hex Editor VB va donc se fermer", vbCritical, "Système d'exploitation non compatible"
        End
    End If
    
    'affiche des messages de warning si on n'a pas une version finale
    #If PRE_ALPHA_VERSION Then
        'version prealpha
        MsgBox "This file is a pre-alpha version, it means that functionnalities are missing and it may contains bugs." & vbNewLine & "This file is avalailable for testing purpose.", vbCritical, "Warning"
    #ElseIf BETA_VERSION Then
        'version beta
        MsgBox "This file is a beta version, it means that all principal functions are availables but there is still bugs." & vbNewLine & "This file is avalailable for testing purpose.", vbCritical, "Warning"
    #End If
    
    Set clsERREUR = New clsGetionErreur 'instancie la classe de gestion des erreurs
    'affecte les properties à la classe
    clsERREUR.LogFile = App.Path & "\ErrLog.log"
    clsERREUR.MakeSoundIDE = True
    
    'applique le style XP (création d'un *.manifest si nécessaire)
    AfManifest.Run
    
    'instancie les classes
    Set cFile = New clsFileInfos
    Set cMem = New clsMemoryRW
    Set cDisk = New clsDiskInfos
    Set clsPref = New clsIniFile
    Set cPref = New clsIniPref
    Set cProc = New clsProcess
    
    'récupère les préférences
    
    '//VERIFIE L'EXISTENCE DU FICHIER DE CONFIGURATION
    '//LE CREE SI INEXISTANT
    
    #If MODE_DEBUG Then
        'alors on est dans la phase Debug, donc on a le dossier du source
        clsPref.sDefaultPath = cFile.GetParentDirectory(App.Path) & "\Executable folder\Preferences\config.ini"
    #Else
        'alors c'est plus la phase debug, donc plus d'IDE possible
        clsPref.sDefaultPath = App.Path & "\Preferences\config.ini" 'détermine le fichier de config par défaut
    #End If
    
    
    Set cPref = clsPref.GetIniFile
    cPref.IniFilePath = clsPref.sDefaultPath
    
    bEndSplash = False
    'affiche le splash si souhaité
    cPref.general_Splash = 1
    If cPref.general_Splash Then
        frmSplash.Show
        DoEvents    '/!\ DO NOT REMOVE (permet d'afficher le splash screen correctement)
    End If
    
    'détermine si le programme a crashé ou pas
    If cPref.exe_HasCrashed = 1 Then
        'alors on sort d'un crash ==> informe
        MsgBox "Le programme n'a pas été fermé correctement, il récupère probablement d'une erreur critique." & vbNewLine & "Merci de me contacter par mail en précisant le contexte et les causes, si possible, du crash." & vbNewLine & "Vous pouvez me contacter en cliquant sur 'Hex Editor VB sur Internet' dans le menu d'aide." & vbNewLine & "Vous pouvez également envoyer le rapport d'erreur (menu Aide ==> rapport d'erreur)." & vbNewLine & "Merci de votre contribution.", vbCritical + vbOKOnly, "Erreur critique lors de la précédente fermeture"
    End If
    'affecte la valeur True au crash
    cPref.exe_HasCrashed = 1
    'sauvegarde les pref (met à jour la valeur)
    Call clsPref.SaveIniFile(cPref) '//CHANGER CA ET NE SAUVER QUE LA VARIABLE CRASH
    
    'créé le raccourci 'envoyer vers...'
    Shortcut True
    
    'ajoute au menu contextuel de windows les entrées de HexEditor
    'AddContextMenu 1    'fichiers
   ' AddContextMenu 0    'dossiers
    
    'ajout du type de fichier *.hescr à HexEditor VB.exe
    Call Reg_HESCR_file
    
    
    ReDim TempFiles(1) As String
    
    'lance la form
    
    
    'créé le tableau contenant la liste des commandes pour l'éditeur de script
    GetSplit
    
    
    '//Ouvre chaque fichier désigné par le path
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
                
                For x = 1 To UBound(sFile())
                    'teste l'existence de chaque path
                
                    If cFile.FileExists(sFile(x)) Then
                        'ouvre un fichier
                        frmShredd.LV.ListItems.Add Text:=sFile(x)
                    ElseIf cFile.FolderExists(sFile(x)) Then
                        'ouvre un dossier - liste les fichiers
                        If cFile.GetFolderFiles(sFile(x), m) <> 0 Then
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
                
                For x = 1 To UBound(sFile())
                    'teste l'existence de chaque path
                
                    If cFile.FileExists(sFile(x)) Then
                        'ouvre un fichier
                        Set Frm = New Pfm
                        Call Frm.GetFile(sFile(x))
                        Frm.Show
                    ElseIf cFile.FolderExists(sFile(x)) Then
                        'ouvre un dossier - liste les fichiers
                        If cFile.GetFolderFiles(sFile(x), m) <> 0 Then
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
             SplitString Chr$(34), Command, sFile()
             
             For x = 1 To UBound(sFile())
                 'teste l'existence de chaque path
                 If cFile.FileExists(sFile(x)) Then
                     'ouvre un fichier
                     Set Frm = New Pfm
                     Call Frm.GetFile(sFile(x))
                     Frm.Show
                 ElseIf cFile.FolderExists(sFile(x)) Then
                     'ouvre un dossier - liste les fichiers
                     If cFile.GetFolderFiles(sFile(x), m) <> 0 Then
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
    
        frmContent.Show
    End If

    Exit Sub
ErrGestion:
    clsERREUR.AddError "Demarrage.Main", True
End Sub

'-------------------------------------------------------
'termine le programme
'-------------------------------------------------------
Public Sub EndProgram()

    'prévient des processus bloqués

    
    'supprime les fichiers temporaires de TempFiles
    
    
    'libère les classes
    Set clsERREUR = Nothing
    Set cFile = Nothing
    Set cMem = Nothing
    Set cDisk = Nothing
    Set cProc = Nothing
    
    'affecte la valeur False au crash (car si on est là, c'est que c'est bien fermé)
    cPref.exe_HasCrashed = 0
    
    'sauvegarde les pref (met à jour la valeur)
    clsPref.SaveIniFile cPref
    
    Set cPref = Nothing
    Set clsPref = Nothing
    
    End 'quitte
End Sub
