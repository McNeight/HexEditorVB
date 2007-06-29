VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{16DCE99A-3937-4772-A07F-3BA5B09FCE6E}#1.1#0"; "vkUserControlsXP.ocx"
Begin VB.Form frmFileSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recherche de fichiers"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9780
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   37
   Icon            =   "frmFileSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Begin vkUserContolsXP.vkBar PGB 
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   5520
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   450
      Value           =   1
      BackPicture     =   "frmFileSearch.frx":058A
      FrontPicture    =   "frmFileSearch.frx":05A6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Type de recherche"
      Height          =   1455
      Index           =   0
      Left            =   83
      TabIndex        =   23
      Top             =   83
      Width           =   2895
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   0
         Left            =   120
         ScaleHeight     =   1095
         ScaleWidth      =   2655
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   2655
         Begin VB.OptionButton Option1 
            Caption         =   "Rechercher des fichiers"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   27
            Tag             =   "pref"
            ToolTipText     =   "Ne recherche que des fichiers"
            Top             =   120
            Value           =   -1  'True
            Width           =   2535
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Rechercher des dossiers"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   26
            Tag             =   "pref"
            ToolTipText     =   "Ne recherche que des dossiers"
            Top             =   480
            Width           =   2535
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Rechercher dans des fichiers"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   25
            Tag             =   "pref"
            ToolTipText     =   "Ne recherche que dans les fichiers (lent)"
            Top             =   840
            Width           =   2535
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Critères de recherche"
      Height          =   1815
      Index           =   1
      Left            =   3083
      TabIndex        =   10
      Top             =   83
      Width           =   6615
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1455
         Index           =   1
         Left            =   120
         ScaleHeight     =   1455
         ScaleWidth      =   6375
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   6375
         Begin VB.CheckBox chkName 
            Caption         =   "Par nom"
            Height          =   195
            Left            =   0
            TabIndex        =   21
            ToolTipText     =   "Ajoute le critère 'nom' à la recherche"
            Top             =   120
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.TextBox txtName 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1200
            TabIndex        =   20
            ToolTipText     =   "Nom à rechercher"
            Top             =   120
            Width           =   2415
         End
         Begin VB.CheckBox chkCasse 
            Caption         =   "Casse"
            Height          =   195
            Left            =   3840
            TabIndex        =   19
            ToolTipText     =   "Respecte ou non la casse"
            Top             =   120
            Width           =   975
         End
         Begin VB.CheckBox chkSize 
            Caption         =   "Par taille"
            Height          =   255
            Left            =   0
            TabIndex        =   18
            ToolTipText     =   "Ajoute le critère 'taille' à la recherche"
            Top             =   600
            Width           =   1095
         End
         Begin VB.ComboBox cdUnit 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmFileSearch.frx":05C2
            Left            =   3840
            List            =   "frmFileSearch.frx":05D2
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Tag             =   "pref lang_ok"
            ToolTipText     =   "Unité"
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtSize 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   2280
            TabIndex        =   16
            Tag             =   "pref"
            Text            =   "100"
            ToolTipText     =   "Taille"
            Top             =   600
            Width           =   1335
         End
         Begin VB.ComboBox cbOpSize 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmFileSearch.frx":05EA
            Left            =   1200
            List            =   "frmFileSearch.frx":05FD
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Tag             =   "pref lang_ok"
            ToolTipText     =   "Opérateur de recherche"
            Top             =   600
            Width           =   855
         End
         Begin VB.CheckBox chkDate 
            Caption         =   "Par date"
            Height          =   195
            Left            =   0
            TabIndex        =   14
            ToolTipText     =   "Ajoute le critère 'date' à la recherche"
            Top             =   1080
            Width           =   975
         End
         Begin VB.ComboBox cbOpDate 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmFileSearch.frx":0612
            Left            =   1200
            List            =   "frmFileSearch.frx":0625
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Tag             =   "pref lang_ok"
            ToolTipText     =   "Opérateur de recherche"
            Top             =   1080
            Width           =   855
         End
         Begin VB.ComboBox cbDateType 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmFileSearch.frx":063A
            Left            =   4440
            List            =   "frmFileSearch.frx":0647
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Tag             =   "pref lang_ok"
            ToolTipText     =   "Type de date"
            Top             =   1080
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker DT 
            Height          =   300
            Left            =   2280
            TabIndex        =   22
            Top             =   1080
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "dd/MM/yyyy hh:mm:ss"
            Format          =   62193667
            CurrentDate     =   39133.9583333333
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Emplacements"
      Height          =   3735
      Left            =   83
      TabIndex        =   6
      Top             =   1643
      Width           =   2895
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   120
         ScaleHeight     =   3375
         ScaleWidth      =   2655
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   2655
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Ajouter un dossier..."
            Height          =   375
            Left            =   0
            TabIndex        =   9
            ToolTipText     =   "Ajouter un dossier à la liste des emplacements où il faut rechercher"
            Top             =   0
            Width           =   2535
         End
         Begin ComctlLib.ListView LV 
            Height          =   2895
            Left            =   0
            TabIndex        =   8
            Tag             =   "lang_ok"
            Top             =   480
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   5106
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   0
            NumItems        =   2
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Dossier"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Sous dossiers"
               Object.Width           =   1764
            EndProperty
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Résultats"
      Height          =   2895
      Left            =   3083
      TabIndex        =   3
      Top             =   2483
      Width           =   6615
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   120
         ScaleHeight     =   2535
         ScaleWidth      =   6375
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   6375
         Begin ComctlLib.ListView LVres 
            Height          =   2535
            Left            =   0
            TabIndex        =   5
            Tag             =   "lang_ok"
            Top             =   0
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   4471
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   0
            NumItems        =   1
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Objet"
               Object.Width           =   12347
            EndProperty
         End
      End
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Lancer la recherche"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3083
      TabIndex        =   2
      ToolTipText     =   "Lance la recherche"
      Top             =   2003
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quitter"
      Height          =   375
      Left            =   8363
      TabIndex        =   1
      ToolTipText     =   "Ferme la fenêtre de recherche"
      Top             =   2003
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Sauvegarder les résultats..."
      Height          =   375
      Left            =   5363
      TabIndex        =   0
      ToolTipText     =   "Sauvegarde les résultats de la recherche"
      Top             =   2003
      Width           =   2775
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "mnuPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuSub 
         Caption         =   "&Recherche dans les sous dossiers"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu rmnuResPop 
      Caption         =   "mnuResPop"
      Visible         =   0   'False
      Begin VB.Menu mnuOpenFolder 
         Caption         =   "&Ouvrir le dossier contenant"
      End
      Begin VB.Menu mnuOpenFile 
         Caption         =   "&Ouvrir le fichier"
      End
      Begin VB.Menu mnuFileProp 
         Caption         =   "&Propriétés du fichier..."
      End
   End
End
Attribute VB_Name = "frmFileSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
'FORM DE RECHERCHE DE FICHIERS DANS LES DISQUES DURS
'=======================================================

Private Lang As New clsLang
Private It As ListItem
Private dblSize As Double    'taille à rechercher
Private curDate As Currency 'date à rechercher
Private bStop As Boolean

Private Sub chkDate_Click()
    If chkDate.Value Then
        cbOpDate.Enabled = True
        cbDateType.Enabled = True
        DT.Enabled = True
    Else
        cbOpDate.Enabled = False
        cbDateType.Enabled = False
        DT.Enabled = False
    End If
    Call CheckSearch 'vérifie qu'une recherche est possible
End Sub
Private Sub chkName_Click()
    If chkName.Value Then
        txtName.Enabled = True
        chkCasse.Enabled = True
    Else
        txtName.Enabled = False
        chkCasse.Enabled = False
    End If
    Call CheckSearch 'vérifie qu'une recherche est possible
End Sub
Private Sub chkSize_Click()
    If chkSize.Value Then
        cbOpSize.Enabled = True
        txtSize.Enabled = True
        cdUnit.Enabled = True
    Else
        cbOpSize.Enabled = False
        txtSize.Enabled = False
        cdUnit.Enabled = False
    End If
    Call CheckSearch 'vérifie qu'une recherche est possible
End Sub

Private Sub cmdAdd_Click()
Dim S As String
    S = cFile.BrowseForFolder(Lang.GetString("_AddFolder"), Me.hWnd)    'browse un dossier
    
    If cFile.FolderExists(S) Then
        'alors ajoute le dossier à la liste des emplacements
        LV.ListItems.Add Text:=S
        LV.ListItems.Item(LV.ListItems.Count).SubItems(1) = Lang.GetString("_NoSub")
    End If
    Call CheckSearch 'vérifie qu'une recherche est possible
End Sub

Private Sub cmdGo_Click()
'lance la recherche

    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_FileSearch"))
    
    If Option1(0).Value Then
        'alors recherche par nom de fichier
        Call LaunchSearch([Recherche de fichiers])
    ElseIf Option1(1).Value Then
        'alors recherche par nom de dossier
        Call LaunchSearch([Recherche de dossiers])
    Else
        'alors recherche par contenu de fichier
        Call LaunchSearch([Recherche de contenu de fichier])
    End If
    
    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_SearchOK"))
End Sub

Private Sub cmdQuit_Click()
    If cmdQuit.Caption = Lang.GetString("_cmdQuitCaption") Then
        Unload Me
    Else
        'annule la recherche
        bStop = True
    End If
End Sub

'=======================================================
'checke si la recherche est possible ou non
'=======================================================
Private Sub CheckSearch()

    cmdGO.Enabled = False
    If LV.ListItems.Count = 0 Then Exit Sub 'aucune zone à chercher
    
    If chkName.Value Then
        cmdGO.Enabled = True
    End If
    If chkSize.Value Then
        If Len(txtSize.Text) > 0 And cdUnit.ListIndex >= 0 And cbOpSize.ListIndex >= 0 Then
            cmdGO.Enabled = True
        Else
            cmdGO.Enabled = False
            Exit Sub
        End If
    End If
    If chkDate.Value Then
        If cbOpDate.ListIndex >= 0 And cbDateType.ListIndex >= 0 Then
            cmdGO.Enabled = True
        Else
            cmdGO.Enabled = False
        End If
    End If
    
End Sub

Private Sub cmdSave_Click()
'sauvegarde les résultats
Dim sFile As String
Dim X As Long
Dim lFile As Long

    If LVres.ListItems.Count = 0 Then
        MsgBox Lang.GetString("_NoRes"), vbInformation, Lang.GetString("_CanNotSave")
        Exit Sub
    End If
    
    On Error GoTo CancelPushed
    
    With frmContent.CMD
        .CancelError = True
        .DialogTitle = Lang.GetString("_SelFile")
        .Filter = Lang.GetString("_TxtFile") & "|*.txt|" & _
            Lang.GetString("_All") & "|*.*"
        .FileName = vbNullString
        .ShowSave
        sFile = .FileName
    End With
    
    If cFile.FileExists(sFile) Then
        'fichier déjà existant
        If MsgBox(Lang.GetString("_FileAlreadyExists"), vbInformation + vbYesNo, Lang.GetString("_War")) <> vbYes Then Exit Sub
    End If
        
    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_SavingRes"))
    
    'récupère un numéro vide
    lFile = FreeFile
    
    'ouvre le fichier
    Open sFile For Append As lFile
    
    For X = 1 To LVres.ListItems.Count
        'sauvegarde la string
        Print #lFile, LVres.ListItems.Item(X).Text
    Next X
    
    'referme le fichier
    Close lFile

    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_ResSaved"))
    
CancelPushed:
End Sub

Private Sub Form_Load()

    With Lang
        #If MODE_DEBUG Then
            If App.LogMode = 0 And CREATE_FRENCH_FILE Then
                'on créé le fichier de langue français
                .Language = "French"
                .LangFolder = LANG_PATH
                Call .WriteIniFileFormIDEform
            End If
        #End If
        
        If App.LogMode = 0 Then
            'alors on est dans l'IDE
            Lang.LangFolder = LANG_PATH
        Else
            Lang.LangFolder = App.Path & "\Lang"
        End If
        
        'applique la langue désirée aux controles
        Call .ActiveLang(Me): .Language = cPref.env_Lang
        Call .LoadControlsCaption
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set It = Nothing
    Set Lang = Nothing
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
'supprime les items sélectionnés
Dim r As Long

    If KeyCode = vbKeyDelete Then
        'touche suppr
        For r = LV.ListItems.Count To 1 Step -1
            If LV.ListItems.Item(r).Selected Then LV.ListItems.Remove r
        Next r
    End If

    Call CheckSearch 'vérifie qu'une recherche est possible
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'change l'attribut "sous dossier"

    Set It = LV.HitTest(X, Y)
    If It Is Nothing Then Exit Sub

    If Button = 2 Then
        Me.mnuSub.Checked = IIf(It.SubItems(1) = "Non", False, True)
        Me.PopupMenu Me.mnuPopup
    End If
End Sub

Private Sub LVres_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu Me.rmnuResPop
    End If
End Sub

Private Sub mnuFileProp_Click()
'affiche les propriétés du fichier sélectionné
    If LVres.SelectedItem Is Nothing Then Exit Sub
    Call cFile.ShowFileProperty(LVres.SelectedItem.Text, Me.hWnd)
End Sub

Private Sub mnuOpenFile_Click()
'ouvre le fichier sélectionné
    If LVres.SelectedItem Is Nothing Then Exit Sub
    Call cFile.ShellOpenFile(LVres.SelectedItem.Text, Me.hWnd)
End Sub

Private Sub mnuOpenFolder_Click()
'ouvre le dossier du fichier sélectionné
    If LVres.SelectedItem Is Nothing Then Exit Sub
    Shell "explorer.exe " & cFile.GetFolderName(LVres.SelectedItem.Text), _
        vbNormalFocus
End Sub

Private Sub mnuSub_Click()
'change le check
    If It Is Nothing Then Exit Sub
    
    Me.mnuSub.Checked = Not (Me.mnuSub.Checked)
    It.SubItems(1) = IIf(Me.mnuSub.Checked, Lang.GetString("_YesSub"), Lang.GetString("_NoSub"))
End Sub

Private Sub Option1_Click(Index As Integer)
    chkSize.Enabled = Option1(0).Value
    chkSize.Value = 0
    chkDate.Enabled = Option1(0).Value
    chkDate.Value = 0
    If Option1(0).Value = False Then
        cbOpSize.Enabled = False
        cbOpDate.Enabled = False
        txtSize.Enabled = False
        cdUnit.Enabled = False
        DT.Enabled = False
        cbDateType.Enabled = False
    End If
End Sub

Private Sub txtName_Change()
    Call CheckSearch 'vérifie qu'une recherche est possible
End Sub
Private Sub txtSize_Change()
    Call CheckSearch 'vérifie qu'une recherche est possible
End Sub
Private Sub cbOpDate_Click()
    Call CheckSearch 'vérifie qu'une recherche est possible
End Sub
Private Sub cbOpSize_Click()
    Call CheckSearch 'vérifie qu'une recherche est possible
End Sub
Private Sub cdUnit_Click()
    Call CheckSearch 'vérifie qu'une recherche est possible
End Sub
Private Sub cbDateType_Click()
    Call CheckSearch 'vérifie qu'une recherche est possible
End Sub

'=======================================================
'effectue la recherche
'=======================================================
Private Sub LaunchSearch(ByVal tMet As TYPE_OF_FILE_SEARCH)
Dim S() As FILE_SEARCH_RESULT
Dim i As Long
Dim X As Long
Dim lC As Long

    'efface le lv de resultats
    LVres.ListItems.Clear
    Frame3.Caption = Lang.GetString("_Res")
    
    Frame1(0).Enabled = False
    Frame1(1).Enabled = False
    Frame2.Enabled = False
    cmdSave.Enabled = False
    cmdGO.Enabled = False
    cmdQuit.Caption = Lang.GetString("_Cancel")
    bStop = False
    DoEvents    '/!\DO NOT REMOVE

    If tMet = [Recherche de fichiers] Then
        'alors recherche par nom de fichier
        
        '//recupère les infos sur le fichier à rechercher
        'on calcule la taille du fichier à rechercher
        If chkSize.Value Then
            dblSize = Abs(Val(txtSize.Text))
            With Lang
                If cdUnit.Text = .GetString("_Ko") Then dblSize = dblSize * 1024
                If cdUnit.Text = .GetString("_Mo") Then dblSize = (dblSize * 1024) * 1024
                If cdUnit.Text = .GetString("_Go") Then dblSize = ((dblSize * 1024) * 1024) * 1024
            End With
        End If
        
        'on calcule sa date
        If chkDate.Value Then curDate = DateString2Currency(DT.Value)
                
        Me.Caption = Lang.GetString("_IndexingFiles")
        
        '//indexe les fichiers
        'contiendra de 1 à ubound une liste de fichiers
        ReDim S(LV.ListItems.Count)
        
        With Me.PGB
            .Min = 0
            .Max = LV.ListItems.Count
            .Value = 0
        End With
        'indexation des fichiers
        For X = LV.ListItems.Count To 1 Step -1
            S(X).sF() = cFile.EnumFilesStr(LV.ListItems.Item(X).Text, _
                IIf(LV.ListItems.Item(X).SubItems(1) = _
                Lang.GetString("_YesSub"), True, False))
                
            Me.PGB.Value = LV.ListItems.Count - X + 1
            If bStop Then GoTo GStop
            
            DoEvents
        Next X
        
        
        '//recherche dans les fichiers indexés
        Me.Caption = Lang.GetString("_SearchingFiles")
            
        lC = 0
        'compte le nombre de fichiers
        For X = 1 To UBound(S())
            lC = lC + UBound(S(X).sF())
        Next X
        With Me.PGB
            .Max = lC
            .Min = 0
            .Value = 0
        End With

        'teste chaque élément
        lC = 0
        LVres.Visible = False
        For X = 1 To UBound(S())
            For i = 1 To UBound(S(X).sF())
                If IsOk(S(X).sF(i)) Then
                    'on ajoute
                    LVres.ListItems.Add Text:=S(X).sF(i)
                End If
                lC = lC + 1
                If (lC Mod 200) = 0 Then
                    DoEvents   'rend la main
                    PGB.Value = lC
                End If
                If bStop Then GoTo GStop
            Next i
        Next X
        PGB.Value = PGB.Max
        Frame3.Caption = Trim$(Str$(LVres.ListItems.Count)) & " " & _
            Lang.GetString("_ResS")
                

    ElseIf tMet = [Recherche de dossiers] Then
        'alors recherche par nom de dossier
        Me.Caption = Lang.GetString("_IndexingFolders")
        
        '//indexe les dossiers
        'contiendra de 1 à ubound une liste de fichiers
        ReDim S(LV.ListItems.Count)
        
        With Me.PGB
            .Min = 0
            .Max = LV.ListItems.Count
            .Value = 0
        End With
        'indexation des dossiers
        For X = LV.ListItems.Count To 1 Step -1
            S(X).sF() = cFile.EnumFoldersStr(LV.ListItems.Item(X).Text, _
                IIf(LV.ListItems.Item(X).SubItems(1) = _
                Lang.GetString("_YesSub"), True, False))
                
                Me.PGB.Value = LV.ListItems.Count - X + 1
                If bStop Then GoTo GStop
            DoEvents
        Next X
        
        
        '//recherche dans les dossiers indexés
        Me.Caption = Lang.GetString("_SearchingFiles")
        
        lC = 0
        'compte le nombre de dossiers
        For X = 1 To UBound(S())
            lC = lC + UBound(S(X).sF())
        Next X
        With Me.PGB
            .Max = lC
            .Min = 0
            .Value = 0
        End With

        'teste chaque élément
        lC = 0
        LVres.Visible = False
        For X = 1 To UBound(S())
            For i = 1 To UBound(S(X).sF())
                If IsOk(S(X).sF(i)) Then
                    'on ajoute
                    LVres.ListItems.Add Text:=S(X).sF(i)
                End If
                lC = lC + 1
                If (lC Mod 200) = 0 Then
                    DoEvents   'rend la main
                    PGB.Value = lC
                End If
                If bStop Then GoTo GStop
            Next i
        Next X
        PGB.Value = PGB.Max
        Frame3.Caption = Trim$(Str$(LVres.ListItems.Count)) & " " & Lang.GetString("_ResS")
        
    Else
        'alors recherche dans le contenu des fichiers
        
        
        
        
    End If
GStop:
    Frame1(0).Enabled = True
    Frame1(1).Enabled = True
    Frame2.Enabled = True
    cmdSave.Enabled = True
    cmdGO.Enabled = True
    cmdQuit.Caption = Lang.GetString("_cmdQuitCaption")
    LVres.Visible = True
End Sub

'=======================================================
'détermine si le fichier est OK pour la recherche
'=======================================================
Private Function IsOk(ByVal sFile As String) As Boolean
Dim l As Long
Dim l2 As Long
Dim Ret As Long
Dim curSize As Currency
Dim curDateReal As Currency

    IsOk = True
    
    'vérifie tout d'abord que le nom du fichier est OK
    If chkName.Value Then
        
        'si pas de texte à chercher, renvoie tous les fichiers
        If txtName.Text = vbNullString Then GoTo NoNameToS
        
        If chkCasse.Value Then l = vbBinaryCompare Else l = vbTextCompare
        
        If InStr(1, cFile.GetFileName(sFile), txtName.Text, l) = 0 Then
            IsOk = False
            Exit Function
        End If
    End If
    
NoNameToS:
    'on continue la recherche
    If chkSize.Value Then
        'alors on doit récupérer la taille du fichier
        curSize = cFile.GetFileSize(sFile)
        
        If cbOpSize.ListIndex = 2 Then
            '<
            If curSize >= dblSize Then
                IsOk = False
                Exit Function
            End If
        ElseIf cbOpSize.ListIndex = 1 Then
            '>
            If curSize <= dblSize Then
                IsOk = False
                Exit Function
            End If
        ElseIf cbOpSize.ListIndex = 0 Then
            '=
            If curSize <> dblSize Then
                IsOk = False
                Exit Function
            End If
        ElseIf cbOpSize.ListIndex = 4 Then
            '<=
            If curSize > dblSize Then
                IsOk = False
                Exit Function
            End If
        ElseIf cbOpSize.ListIndex = 3 Then
            '>=
            If curSize < dblSize Then
                IsOk = False
                Exit Function
            End If
        End If
    End If
    
    'continue la recherche
    If chkDate.Value Then
    
        l2 = cbOpDate.ListIndex
        
        'récupère la date en Currency
        Select Case cbDateType.ListIndex
            Case 0
                curDateReal = cFile.GetFileDateAsCurrency(sFile, DateCreated)
            Case 1
                curDateReal = cFile.GetFileDateAsCurrency(sFile, DateLastAccessed)
            Case 2
                curDateReal = cFile.GetFileDateAsCurrency(sFile, DateLastModified)
        End Select
        
        If curDateReal = 0 Then
            'date inaccessible
            IsOk = False
            Exit Function
        End If
        
        'compare avec la date à rechercher
        '-1==> curDate < curDateReal
        '0 ==> curDate = curDateReal
        '1 ==> curDate > curDateReal
        Ret = CompareFileTime(curDate, curDateReal)
        
        If Ret = 0 Then
            If (l2 = 0) Or (l2 = 3) Or (l2 = 4) Then Exit Function
        ElseIf Ret = 1 Then
            If (l2 = 2) Or (l2 = 4) Then Exit Function
        Else
            If (l2 = 1) Or (l2 = 3) Then Exit Function
        End If
        
        'si on est là, c'est que la comparaison n'est pas bonne
        IsOk = False
    End If
    
End Function
