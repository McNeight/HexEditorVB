VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{6ADE9E73-F694-428F-BF86-06ADD29476A5}#1.0#0"; "ProgressBar_OCX.ocx"
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
   Icon            =   "frmFileSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Begin ProgressBar_OCX.pgrBar pgb 
      Height          =   255
      Left            =   120
      TabIndex        =   27
      ToolTipText     =   "Avancement de la recherche"
      Top             =   5520
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   450
      BackColorTop    =   13027014
      BackColorBottom =   15724527
      Value           =   1
      BackPicture     =   "frmFileSearch.frx":08CA
      FrontPicture    =   "frmFileSearch.frx":08E6
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Sauvegarder les r�sultats..."
      Height          =   375
      Left            =   5400
      TabIndex        =   26
      ToolTipText     =   "Sauvegarde les r�sultats de la recherche"
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quitter"
      Height          =   375
      Left            =   8400
      TabIndex        =   25
      ToolTipText     =   "Ferme la fen�tre de recherche"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Lancer la recherche"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   24
      ToolTipText     =   "Lance la recherche"
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      Caption         =   "R�sultats"
      Height          =   2895
      Left            =   3120
      TabIndex        =   21
      Top             =   2520
      Width           =   6615
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   120
         ScaleHeight     =   2535
         ScaleWidth      =   6375
         TabIndex        =   22
         Top             =   240
         Width           =   6375
         Begin ComctlLib.ListView LVres 
            Height          =   2535
            Left            =   0
            TabIndex        =   23
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
   Begin VB.Frame Frame2 
      Caption         =   "Emplacements"
      Height          =   3735
      Left            =   120
      TabIndex        =   17
      Top             =   1680
      Width           =   2895
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   120
         ScaleHeight     =   3375
         ScaleWidth      =   2655
         TabIndex        =   18
         Top             =   240
         Width           =   2655
         Begin ComctlLib.ListView LV 
            Height          =   2895
            Left            =   0
            TabIndex        =   20
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
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Ajouter un dossier..."
            Height          =   255
            Left            =   0
            TabIndex        =   19
            ToolTipText     =   "Ajouter un dossier � la liste des emplacements o� il faut rechercher"
            Top             =   120
            Width           =   2175
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Crit�res de recherche"
      Height          =   1815
      Index           =   1
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   6615
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1455
         Index           =   1
         Left            =   120
         ScaleHeight     =   1455
         ScaleWidth      =   6375
         TabIndex        =   3
         Top             =   240
         Width           =   6375
         Begin VB.CommandButton cmdDate 
            Caption         =   "..."
            Height          =   255
            Left            =   4200
            TabIndex        =   29
            ToolTipText     =   "S�lectionner une date..."
            Top             =   1080
            Width           =   375
         End
         Begin VB.ComboBox cbDateType 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmFileSearch.frx":0902
            Left            =   4680
            List            =   "frmFileSearch.frx":090F
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Tag             =   "pref"
            ToolTipText     =   "Type de date"
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox txtDate 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   2280
            TabIndex        =   16
            ToolTipText     =   "Date (et heure)"
            Top             =   1080
            Width           =   1815
         End
         Begin VB.ComboBox cbOpDate 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmFileSearch.frx":093D
            Left            =   1200
            List            =   "frmFileSearch.frx":0950
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Tag             =   "pref"
            ToolTipText     =   "Op�rateur de recherche"
            Top             =   1080
            Width           =   855
         End
         Begin VB.CheckBox chkDate 
            Caption         =   "Par date"
            Height          =   195
            Left            =   0
            TabIndex        =   14
            ToolTipText     =   "Ajoute le crit�re 'date' � la recherche"
            Top             =   1080
            Width           =   975
         End
         Begin VB.ComboBox cbOpSize 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmFileSearch.frx":0965
            Left            =   1200
            List            =   "frmFileSearch.frx":0978
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Tag             =   "pref"
            ToolTipText     =   "Op�rateur de recherche"
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtSize 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   2280
            TabIndex        =   12
            Tag             =   "pref"
            Text            =   "100"
            ToolTipText     =   "Taille"
            Top             =   600
            Width           =   1335
         End
         Begin VB.ComboBox cdUnit 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmFileSearch.frx":098D
            Left            =   3840
            List            =   "frmFileSearch.frx":099D
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Tag             =   "pref"
            ToolTipText     =   "Unit�"
            Top             =   600
            Width           =   1095
         End
         Begin VB.CheckBox chkSize 
            Caption         =   "Par taille"
            Height          =   255
            Left            =   0
            TabIndex        =   10
            ToolTipText     =   "Ajoute le crit�re 'taille' � la recherche"
            Top             =   600
            Width           =   1095
         End
         Begin VB.CheckBox chkCasse 
            Caption         =   "Casse"
            Height          =   195
            Left            =   3840
            TabIndex        =   9
            ToolTipText     =   "Respecte ou non la casse"
            Top             =   120
            Width           =   975
         End
         Begin VB.TextBox txtName 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1200
            TabIndex        =   8
            ToolTipText     =   "Nom � rechercher"
            Top             =   120
            Width           =   2415
         End
         Begin VB.CheckBox chkName 
            Caption         =   "Par nom"
            Height          =   195
            Left            =   0
            TabIndex        =   7
            ToolTipText     =   "Ajoute le crit�re 'nom' � la recherche"
            Top             =   120
            Value           =   1  'Checked
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Type de recherche"
      Height          =   1455
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   0
         Left            =   120
         ScaleHeight     =   1095
         ScaleWidth      =   2655
         TabIndex        =   1
         Top             =   240
         Width           =   2655
         Begin VB.OptionButton Option1 
            Caption         =   "Rechercher dans des fichiers"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   6
            Tag             =   "pref"
            ToolTipText     =   "Ne recherche que dans les fichiers (lent)"
            Top             =   840
            Width           =   2535
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Rechercher des dossiers"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   5
            Tag             =   "pref"
            ToolTipText     =   "Ne recherche que des dossiers"
            Top             =   480
            Width           =   2535
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Rechercher des fichiers"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   4
            Tag             =   "pref"
            ToolTipText     =   "Ne recherche que des fichiers"
            Top             =   120
            Value           =   -1  'True
            Width           =   2535
         End
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "mnuPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuSub 
         Caption         =   "&Recherche dans les sous dossiers"
         Checked         =   -1  'True
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
'FORM DE RECHERCHE DE FICHIERS DANS LES DISQUES DURS
'=======================================================

Private It As ListItem
Private dblSize As Double    'taille � rechercher
Private curDate As Currency 'date � rechercher
Private bStop As Boolean

Private Sub chkDate_Click()
    If chkDate.Value Then
        txtDate.Enabled = True
        cbOpDate.Enabled = True
        cbDateType.Enabled = True
        cmdDate.Enabled = True
    Else
        txtDate.Enabled = False
        cbOpDate.Enabled = False
        cbDateType.Enabled = False
        cmdDate.Enabled = False
    End If
    Call CheckSearch 'v�rifie qu'une recherche est possible
End Sub
Private Sub chkName_Click()
    If chkName.Value Then
        txtName.Enabled = True
        chkCasse.Enabled = True
    Else
        txtName.Enabled = False
        chkCasse.Enabled = False
    End If
    Call CheckSearch 'v�rifie qu'une recherche est possible
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
    Call CheckSearch 'v�rifie qu'une recherche est possible
End Sub

Private Sub cmdAdd_Click()
Dim s As String
    s = cFile.BrowseForFolder("Ajouter un dossier", Me.hwnd)    'browse un dossier
    
    If cFile.FolderExists(s) Then
        'alors ajoute le dossier � la liste des emplacements
        LV.ListItems.Add Text:=s
        LV.ListItems.Item(LV.ListItems.Count).SubItems(1) = "Non"
    End If
    Call CheckSearch 'v�rifie qu'une recherche est possible
End Sub

Private Sub cmdGo_Click()
'lance la recherche
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
End Sub

Private Sub cmdQuit_Click()
    If cmdQuit.Caption = "Quitter" Then
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

    cmdGo.Enabled = False
    If LV.ListItems.Count = 0 Then Exit Sub 'aucune zone � chercher
    
    If chkName.Value Then
        cmdGo.Enabled = True
    End If
    If chkSize.Value Then
        If Len(txtSize.Text) > 0 And cdUnit.ListIndex >= 0 And cbOpSize.ListIndex >= 0 Then
            cmdGo.Enabled = True
        Else
            cmdGo.Enabled = False
            Exit Sub
        End If
    End If
    If chkDate.Value Then
        If Len(txtDate.Text) > 0 And cbOpDate.ListIndex >= 0 And cbDateType.ListIndex >= 0 Then
            cmdGo.Enabled = True
        Else
            cmdGo.Enabled = False
        End If
    End If
    
End Sub

Private Sub cmdSave_Click()
'sauvegarde les r�sultats


End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set It = Nothing
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
'supprime les items s�lectionn�s
Dim r As Long

    If KeyCode = vbKeyDelete Then
        'touche suppr
        For r = LV.ListItems.Count To 1 Step -1
            If LV.ListItems.Item(r).Selected Then LV.ListItems.Remove r
        Next r
    End If

    Call CheckSearch 'v�rifie qu'une recherche est possible
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'change l'attribut "sous dossier"

    Set It = LV.HitTest(x, y)
    If It Is Nothing Then Exit Sub

    If Button = 2 Then
        Me.mnuSub.Checked = IIf(It.SubItems(1) = "Non", False, True)
        Me.PopupMenu Me.mnuPopUp
    End If
End Sub

Private Sub mnuSub_Click()
'change le check
    If It Is Nothing Then Exit Sub
    
    Me.mnuSub.Checked = Not (Me.mnuSub.Checked)
    It.SubItems(1) = IIf(Me.mnuSub.Checked, "Oui", "Non")
End Sub

Private Sub Option1_Click(Index As Integer)
    chkSize.Enabled = Option1(0).Value
    chkSize.Value = 0
    chkDate.Enabled = Option1(0).Value
    chkDate.Value = 0
    If Option1(0).Value = False Then
        cbOpSize.Enabled = False
        cbOpDate.Enabled = False
        txtDate.Enabled = False
        txtSize.Enabled = False
        cmdDate.Enabled = False
        cdUnit.Enabled = False
        cbDateType.Enabled = False
    End If
End Sub

Private Sub txtDate_Change()
    Call CheckSearch 'v�rifie qu'une recherche est possible
End Sub
Private Sub txtName_Change()
    Call CheckSearch 'v�rifie qu'une recherche est possible
End Sub
Private Sub txtSize_Change()
    Call CheckSearch 'v�rifie qu'une recherche est possible
End Sub
Private Sub cbOpDate_Click()
    Call CheckSearch 'v�rifie qu'une recherche est possible
End Sub
Private Sub cbOpSize_Click()
    Call CheckSearch 'v�rifie qu'une recherche est possible
End Sub
Private Sub cdUnit_Click()
    Call CheckSearch 'v�rifie qu'une recherche est possible
End Sub
Private Sub cbDateType_Click()
    Call CheckSearch 'v�rifie qu'une recherche est possible
End Sub

'=======================================================
'effectue la recherche
'=======================================================
Private Sub LaunchSearch(ByVal tMet As TYPE_OF_FILE_SEARCH)
Dim s() As FILE_SEARCH_RESULT
Dim i As Long
Dim x As Long
Dim lC As Long

    'efface le lv de resultats
    LVres.ListItems.Clear
    Frame3.Caption = "R�sultats"
    
    Frame1(0).Enabled = False
    Frame1(1).Enabled = False
    Frame2.Enabled = False
    cmdSave.Enabled = False
    cmdGo.Enabled = False
    cmdQuit.Caption = "Annuler"
    bStop = False
    DoEvents    '/!\DO NOT REMOVE

    If tMet = [Recherche de fichiers] Then
        'alors recherche par nom de fichier
        
        '//recup�re les infos sur le fichier � rechercher
        'on calcule la taille du fichier � rechercher
        If chkSize.Value Then
            dblSize = Abs(Val(txtSize.Text))
            If cdUnit.Text = "Ko" Then dblSize = dblSize * 1024
            If cdUnit.Text = "Mo" Then dblSize = (dblSize * 1024) * 1024
            If cdUnit.Text = "Go" Then dblSize = ((dblSize * 1024) * 1024) * 1024
        End If
        
        'on calcule sa date
        curDate = DateString2Currency(txtDate.Text)
                
        Me.Caption = "Indexation des fichiers..."
        
        '//indexe les fichiers
        'contiendra de 1 � ubound une liste de fichiers
        ReDim s(LV.ListItems.Count)
        
        With Me.pgb
            .Min = 0
            .Max = LV.ListItems.Count
            .Value = 0
        End With
        'indexation des fichiers
        For x = LV.ListItems.Count To 1 Step -1
            Call cFile.GetFolderFiles(LV.ListItems.Item(x).Text, s(x).sF(), _
                IIf(LV.ListItems.Item(x).SubItems(1) = "Oui", True, False))
                Me.pgb.Value = LV.ListItems.Count - x + 1
                If bStop Then GoTo GStop
            DoEvents
        Next x
        
        
        '//recherche dans les fichiers index�s
        Me.Caption = "Recherche de fichiers"
            
        lC = 0
        'compte le nombre de fichiers
        For x = 1 To UBound(s())
            lC = lC + UBound(s(x).sF())
        Next x
        With Me.pgb
            .Max = lC
            .Min = 0
            .Value = 0
        End With

        'teste chaque �l�ment
        lC = 0
        LVres.Visible = False
        For x = 1 To UBound(s())
            For i = 1 To UBound(s(x).sF())
                If IsOk(s(x).sF(i)) Then
                    'on ajoute
                    LVres.ListItems.Add Text:=s(x).sF(i)
                End If
                lC = lC + 1
                If (lC Mod 200) = 0 Then
                    DoEvents   'rend la main
                    pgb.Value = lC
                End If
                If bStop Then GoTo GStop
            Next i
        Next x
        pgb.Value = pgb.Max
        Frame3.Caption = Trim$(Str$(LVres.ListItems.Count)) & " r�sultat(s)"
                

    ElseIf tMet = [Recherche de dossiers] Then
        'alors recherche par nom de dossier
        Me.Caption = "Indexation des dossiers..."
        
        '//indexe les dossiers
        'contiendra de 1 � ubound une liste de fichiers
        ReDim s(LV.ListItems.Count)
        
        With Me.pgb
            .Min = 0
            .Max = LV.ListItems.Count
            .Value = 0
        End With
        'indexation des dossiers
        For x = LV.ListItems.Count To 1 Step -1
            Call cFile.EnumFolders(LV.ListItems.Item(x).Text, s(x).sF(), True, _
                IIf(LV.ListItems.Item(x).SubItems(1) = "Oui", True, False))
                Me.pgb.Value = LV.ListItems.Count - x + 1
                If bStop Then GoTo GStop
            DoEvents
        Next x
        
        
        '//recherche dans les dossiers index�s
        Me.Caption = "Recherche de fichiers"
        
        lC = 0
        'compte le nombre de dossiers
        For x = 1 To UBound(s())
            lC = lC + UBound(s(x).sF())
        Next x
        With Me.pgb
            .Max = lC
            .Min = 0
            .Value = 0
        End With

        'teste chaque �l�ment
        lC = 0
        LVres.Visible = False
        For x = 1 To UBound(s())
            For i = 1 To UBound(s(x).sF())
                If IsOk(s(x).sF(i)) Then
                    'on ajoute
                    LVres.ListItems.Add Text:=s(x).sF(i)
                End If
                lC = lC + 1
                If (lC Mod 200) = 0 Then
                    DoEvents   'rend la main
                    pgb.Value = lC
                End If
                If bStop Then GoTo GStop
            Next i
        Next x
        pgb.Value = pgb.Max
        Frame3.Caption = Trim$(Str$(LVres.ListItems.Count)) & " r�sultat(s)"
        
    Else
        'alors recherche dans le contenu des fichiers
        
        
        
        
    End If
GStop:
    Frame1(0).Enabled = True
    Frame1(1).Enabled = True
    Frame2.Enabled = True
    cmdSave.Enabled = True
    cmdGo.Enabled = True
    cmdQuit.Caption = "Quitter"
    LVres.Visible = True
End Sub

'=======================================================
'd�termine si le fichier est OK pour la recherche
'=======================================================
Private Function IsOk(ByVal sFile As String) As Boolean
Dim l As Long
Dim l2 As Long
Dim Ret As Long
Dim curSize As Currency
Dim curDateReal As Currency

    IsOk = True
    
    'v�rifie tout d'abord que le nom du fichier est OK
    If chkName.Value Then
        
        'si pas de texte � chercher, renvoie tous les fichiers
        If txtName.Text = vbNullString Then GoTo NoNameToS
        
        If chkCasse.Value Then l = vbBinaryCompare Else l = vbTextCompare
        
        If InStr(1, cFile.GetFileFromPath(sFile), txtName.Text, l) = 0 Then
            IsOk = False
            Exit Function
        End If
    End If
    
NoNameToS:
    'on continue la recherche
    If chkSize.Value Then
        'alors on doit r�cup�rer la taille du fichier
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
        
        'r�cup�re la date en Currency
        curDateReal = cFile.GetFileDate(sFile, cbDateType.ListIndex, True)
        
        If curDateReal = 0 Then
            'date inaccessible
            IsOk = False
            Exit Function
        End If
        
        'compare avec la date � rechercher
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
        
        'si on est l�, c'est que la comparaison n'est pas bonne
        IsOk = False
    End If
    
End Function
