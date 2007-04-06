VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Lang Editor Tool for Hex Editor VB"
   ClientHeight    =   8205
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10290
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   10290
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ListView LV 
      Height          =   1935
      Left            =   1800
      TabIndex        =   0
      Top             =   2280
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   3413
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "New string"
         Object.Width           =   5644
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Model string"
         Object.Width           =   5644
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Type"
         Object.Width           =   5644
      EndProperty
   End
   Begin MSComDlg.CommonDialog CMD 
      Left            =   240
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuCreateLangFile 
         Caption         =   "&Create a new lang file"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpenFile 
         Caption         =   "&Open a lang file..."
         Enabled         =   0   'False
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save lang file..."
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuChooseModel 
         Caption         =   "&Choose a model..."
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edition"
      Begin VB.Menu mnuCopyModel 
         Caption         =   "&Copy model to clipboard"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuInsert 
         Caption         =   "&Insert from clipboard"
      End
   End
   Begin VB.Menu mnuChooseLang 
      Caption         =   "&Lang"
      Begin VB.Menu mnuEnglish 
         Caption         =   "&English"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFrench 
         Caption         =   "&Français"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmMain"
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
'FORM PRINCIPALE D'EDITION DE LANGUE
'=======================================================

Private sOpenModel As String
Private sSaveFile As String
Private sOpenFile As String
Private lRLen As Long
Private sMod() As String
Private bModelChosen As Boolean

Private Sub Form_Load()
Dim cManif As AfClsManifest

    'applique le style XP
    Set cManif = New AfClsManifest
    Call cManif.Run
    Set cManif = Nothing
    
    'en anglais par défaut
    sOpenFile = "Choose a file"
    sOpenModel = "Choose a model"
    bModelChosen = False
    sSaveFile = "Save lang file"
End Sub

Private Sub Form_Resize()
    With LV
        .Left = 0
        .Top = 0
        .Width = Me.Width - 140
        .Height = Me.Height - 830
    End With
End Sub

Private Sub LV_Click()
    Call LV.StartLabelEdit
End Sub

Private Sub mnuChooseModel_Click()
'choix d'un modèle
Dim s As String
Dim s2() As String
Dim x As Long
Dim cFile As clsFileInfos
Dim l As Long
Dim r As Long

    On Error GoTo CancelPushed
    
    'on charge le fichier que l'on vient de sélectionner
    Set cFile = New clsFileInfos
    
    With Me.CMD
        .CancelError = True
        .DialogTitle = sOpenModel
        .Filter = "Lang files|*.ini"
        .InitDir = App.Path & "\Lang"
        .ShowOpen
        s = .FileName
    End With

    'récupère le fichier *.ini dans une string
    s = cFile.LoadFileInString(s)
    
    'split chaque ligne
    s2() = Split(s, vbNewLine, , vbBinaryCompare)
    
    'ajoute chaque ligne dans le LV
    LV.ListItems.Clear: r = 0
    For x = 0 To UBound(s2())
        
        'on récupère la position du "="
        l = InStr(1, s2(x), "=", vbBinaryCompare)
        
        If l < Len(s2(x)) And Left$(s2(x), 1) <> "[" Then
            'élément OK
            r = r + 1
            LV.ListItems.Add Text:=vbNullString
            LV.ListItems.Item(r).SubItems(1) = Right$(s2(x), Len(s2(x)) - l)
            LV.ListItems.Item(r).SubItems(2) = Left$(s2(x), l - 1)
            LV.ListItems.Item(r).Tag = CStr(x)  'IMPORTANT
        End If
            
    Next x
    
    lRLen = UBound(s2())
    bModelChosen = True
    sMod() = s2()
    Me.mnuSave.Enabled = True
    Me.mnuOpenFile.Enabled = True
    Me.mnuCopyModel.Enabled = True
    
CancelPushed:

    'libère la classe
    Set cFile = Nothing
End Sub

Private Sub mnuCopyModel_Click()
'copie le modèle vers le presse papier
Dim x As Long
Dim s As String
Dim s2 As String

    For x = 1 To LV.ListItems.Count
    
        'formate la string pour Google traduction
        s = s & Replace$(LV.ListItems.Item(x).SubItems(1), "&", " & ", , _
            , vbBinaryCompare) & " | "

    Next x
    
    Clipboard.Clear
    Call Clipboard.SetText(s)
           
End Sub

Private Sub mnuCreateLangFile_Click()
'création d'un nouveau fichier de langue
Dim x As Long

    'on vire tous les textes de la colonne de gauche
    For x = 1 To LV.ListItems.Count
        LV.ListItems.Item(x).Text = vbNullString
    Next x
End Sub

Private Sub mnuEnglish_Click()
'met les menus en anglais
    mnuEnglish.Checked = True
    mnuFrench.Checked = False
    
    mnuChooseModel.Caption = "&Choose a model..."
    mnuCreateLangFile.Caption = "&Create a new lang file"
    mnuHelp.Caption = "&Help"
    mnuOpenFile.Caption = "&Open lang file..."
    mnuQuit.Caption = "&Quit"
    mnuFile.Caption = "&File"
    mnuChooseLang.Caption = "&Lang"
    mnuSave.Caption = "&Save lang file..."
    Me.mnuCopyModel.Caption = "&Copier le modèle dans le presse-papier"
    Me.mnuInsert.Caption = "&Insérer depuis le presse-papier"
    LV.ColumnHeaders.Item(2).Text = "Model string"
    LV.ColumnHeaders.Item(1).Text = "New string"
    sOpenFile = "Choose a file"
    sOpenModel = "Choose a model"
    sSaveFile = "Save lang file"
End Sub

Private Sub mnuFrench_Click()
'met les menus en français
    mnuFrench.Checked = True
    mnuEnglish.Checked = False
    
    mnuChooseModel.Caption = "&Choix d'un modèle..."
    mnuCreateLangFile.Caption = "&Créer un nouveau fichier de langue"
    mnuHelp.Caption = "&Aide"
    mnuOpenFile.Caption = "&Ouvrir un fichier..."
    mnuQuit.Caption = "&Quitter"
    mnuFile.Caption = "&Fichier"
    mnuChooseLang.Caption = "&Langue"
    mnuSave.Caption = "&Enregistrer le fichier de langue..."
    LV.ColumnHeaders.Item(2).Text = "Texte du modèle"
    LV.ColumnHeaders.Item(1).Text = "Nouveau texte"
    Me.mnuCopyModel.Caption = "&Copy model to clipboard"
    Me.mnuInsert.Caption = "&Insert from clipboard"
    sOpenFile = "Choix d'un fichier"
    sOpenModel = "Choix d'un modèle"
    sSaveFile = "Sauvegarder le fichier"
End Sub

Private Sub mnuHelp_Click()
'aide
Dim s As String
    
    If Me.mnuFrench.Checked Then
        'en français
        s = "Module d'édition de fichiers de langue pour Hex Editor VB" & vbNewLine
        s = s & vbNewLine & "Vous devez :" & vbNewLine & "1) Choisir un modèle, c'est à dire un fichier de langue déjà existant"
        s = s & vbNewLine & "2) Créer votre fichier de langue en traduisant dans la colonne de gauche les textes de droite"
        s = s & vbNewLine & "3) Sauvegarder votre nouveau fichier de langue"
        s = s & vbNewLine & "4) Vous pouvez 'Copier le modèle dans le presse papier, coller dans google traduction, copier le"
        s = s & vbNewLine & "résultat traduit, et 'Insérer depuis le presse papier'"
        s = s & vbNewLine & "Il est (très fortement) conseillé de faire une traduction depuis le modèle anglais."
        s = s & vbNewLine & vbNewLine & "By violent_ken"
    Else
        s = "Lang Editor Tool for Hex Editor VB" & vbNewLine
        s = s & vbNewLine & "You should :" & vbNewLine & "1) Choose a model, i.e. an existant lang file"
        s = s & vbNewLine & "2) Create your own lang file by translating in the left column the text of the right column"
        s = s & vbNewLine & "3) Save your new lang file"
        s = s & vbNewLine & "4) You can 'Copy model to clipboard', paste it in Google Traduction, copy it"
        s = s & vbNewLine & "to clipboard and 'Insert from clipboard' the translation"
        s = s & vbNewLine & "It is (very strongly) advised to make a translation from the English model."
        s = s & vbNewLine & vbNewLine & "By violent_ken"
    End If
    
    MsgBox s, vbInformation, "Lang Editor for Hex Editor VB"
        
End Sub

Private Sub mnuInsert_Click()
'insère depuis le presse papier
Dim x As Long
Dim s As String
Dim s2() As String

    'récupère depuis le clipboard
    s = Clipboard.GetText
    
    'formate (car google fait nimp des fois)
    s = Replace$(s, "|", " | ", , , vbBinaryCompare)
    s = Replace$(s, " & ", "&", , , vbBinaryCompare)
    s = Replace$(s, "  ", " ", , , vbBinaryCompare)
    
    'récupère chaque ligne
    s2() = Split(s, " | ", , vbBinaryCompare)
    
    If UBound(s2()) <> LV.ListItems.Count Then Exit Sub
    
    For x = 1 To LV.ListItems.Count
        LV.ListItems.Item(x).Text = s2(x - 1)
    Next x

End Sub

Private Sub mnuSave_Click()
'ouvre un fichier de langue
Dim s As String
Dim s2() As String
Dim sNew() As String
Dim x As Long
Dim cFile As clsFileInfos
Dim l As Long
Dim r As Long
Dim sFile As String

    'On Error GoTo CancelPushed
    
    
    'vérifie qu'un modèle existe bien
    If bModelChosen = False Then
        If Me.mnuFrench.Checked Then
            MsgBox "Vous devez d'abord choisir un modèle avant de pouvoir sauvegarder", vbCritical, "Erreur"
            GoTo CancelPushed
        Else
            MsgBox "You have to choose a model before saving your file", vbCritical, "Error"
            GoTo CancelPushed
        End If
    End If
    
    'on charge le fichier que l'on vient de sélectionner
    Set cFile = New clsFileInfos
    
    With Me.CMD
        .CancelError = True
        .DialogTitle = sSaveFile
        .Filter = "Lang files|*.ini"
        .InitDir = App.Path & "\Lang"
        .ShowSave
        sFile = .FileName
    End With
    
    'fichier déjà existant ?
    If cFile.FileExists(sFile) Then
        If Me.mnuFrench.Checked Then
            x = MsgBox("Le fichier existe déjà, le remplacer ?", _
                vbInformation + vbYesNo, "Attention")
            If x <> vbYes Then GoTo CancelPushed
        Else
            x = MsgBox("File already exists, overwrite it ?", _
                vbInformation + vbYesNo, "Warning")
            If x <> vbYes Then GoTo CancelPushed
        End If
    End If
       
    
    'ajoute chaque ligne dans le LV
    r = 0: ReDim sNew(lRLen): r = 0
    For x = 0 To lRLen
        
        'on récupère la position du "="
        l = InStr(1, sMod(x), "=", vbBinaryCompare)
        
        'on récupère dans sMod() la string de gauche de l'égalité
        s = Left$(sMod(x), l)
        
        If l < Len(sMod(x)) And Left$(sMod(x), 1) <> "[" Then
            'élément OK
            r = r + 1
            s = s & LV.ListItems.Item(r).Text
        End If
        
        If l Then
            sNew(x) = s
        Else
            sNew(x) = sMod(x)
        End If
    Next x
    
    'on sauvegarde maintenant toutes les strings de sNew() dans le fichier
    s = vbNullString
    For x = 0 To lRLen
        s = s & sNew(x) & vbNewLine
    Next x
    
    'lance l'enregistrement
    Call cFile.SaveDATAinFile(sFile, s, True)
    
    
CancelPushed:

    'libère la classe
    Set cFile = Nothing
End Sub

Private Sub mnuQuit_Click()
    Unload Me
End Sub

Private Sub mnuOpenFile_Click()
'ouvre le fichier
Dim s As String
Dim s2() As String
Dim x As Long
Dim cFile As clsFileInfos
Dim l As Long
Dim r As Long

    On Error GoTo CancelPushed
    
    'on charge le fichier que l'on vient de sélectionner
    Set cFile = New clsFileInfos
    
    With Me.CMD
        .CancelError = True
        .DialogTitle = sOpenFile
        .Filter = "Lang files|*.ini"
        .InitDir = App.Path & "\Lang"
        .ShowOpen
        s = .FileName
    End With
    
    'récupère le fichier *.ini dans une string
    s = cFile.LoadFileInString(s)
    
    'split chaque ligne
    s2() = Split(s, vbNewLine, , vbBinaryCompare)

    'ajoute chaque ligne dans le LV
    For x = 1 To LV.ListItems.Count
        r = Val(LV.ListItems.Item(x).Tag)
        l = InStr(1, s2(r), "=", vbBinaryCompare)
        LV.ListItems.Item(x).Text = Right$(s2(r), Len(s2(r)) - l)
    Next x

CancelPushed:

    'libère la classe
    Set cFile = Nothing
End Sub
