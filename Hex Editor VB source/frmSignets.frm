VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{16DCE99A-3937-4772-A07F-3BA5B09FCE6E}#1.1#0"; "vkUserControlsXP.ocx"
Begin VB.Form frmSignets 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gestionnaire de signets"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5475
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   22
   Icon            =   "frmSignets.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin vkUserContolsXP.vkCommand cmdSaveChanges 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Applique les changements � la fen�tre active"
      Top             =   5040
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      Caption         =   "Appliquer les changements"
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
   Begin vkUserContolsXP.vkFrame vkFrame2 
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2355
      Caption         =   "Signets"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin vkUserContolsXP.vkCommand cmdComment 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Ajouter un commentaite sur le signet s�lectionn�"
         Top             =   840
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   661
         Caption         =   "Ajouter un commentaite sur le signet s�lectionn�"
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
      Begin vkUserContolsXP.vkCommand cmdNew 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Cr�� un nouveau signet"
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "Nouveau"
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
      Begin vkUserContolsXP.vkCommand cmdDeleteSel 
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         ToolTipText     =   "Supprime les fichiers s�lectionn�s"
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Supprimer s�lection"
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
      Begin vkUserContolsXP.vkCommand cmdDeleteAll 
         Height          =   375
         Left            =   3720
         TabIndex        =   11
         ToolTipText     =   "Supprime tous les signets"
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "Supprimer tout"
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
   End
   Begin vkUserContolsXP.vkFrame vkFrame1 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1508
      Caption         =   "Liste des signets"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin vkUserContolsXP.vkCommand cmdSave 
         Height          =   375
         Left            =   3720
         TabIndex        =   7
         ToolTipText     =   "Sauvegarde la liste de signets actuelle"
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "Sauvegarder..."
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
      Begin vkUserContolsXP.vkCommand cmdAdd 
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         ToolTipText     =   "Ajoute une liste de signets"
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "Ajouter..."
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
      Begin vkUserContolsXP.vkCommand cmdOpen 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Ouvre une liste de signets"
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "Ouvrir..."
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
   End
   Begin ComctlLib.ListView lstSignets 
      Height          =   2535
      Left            =   48
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "lang_ok"
      Top             =   0
      Width           =   5340
      _ExtentX        =   9419
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Offset"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Commentaire"
         Object.Width           =   6703
      EndProperty
   End
   Begin vkUserContolsXP.vkCommand dmQuit 
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      ToolTipText     =   "Fermer cette fen�tre"
      Top             =   5040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Fermer"
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
End
Attribute VB_Name = "frmSignets"
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
'FORM DE GESTION DES SIGNETS
'=======================================================

Private Lang As New clsLang
Private mouseUped As Boolean

Private Sub cmdAdd_Click()
'ajoute une liste de signets
    Call AddSignetIn(False)
End Sub

Private Sub cmdComment_Click()
'ajoute un commentaire sur les signets s�lectionn�s
Dim x As Long
Dim S As String

    For x = Me.lstSignets.ListItems.Count To 1 Step -1
        If Me.lstSignets.ListItems.Item(x).Selected Then
            S = InputBox(Lang.GetString("_NewComment") & " " & _
                Me.lstSignets.ListItems.Item(x).Text, Lang.GetString("_AddComment"))
            If StrPtr(S) <> 0 Then _
                Me.lstSignets.ListItems.Item(x).SubItems(1) = S
        End If
    Next x
End Sub

Private Sub cmdDeleteAll_Click()
'supprime tous les signets
    Call lstSignets.ListItems.Clear
End Sub

Private Sub cmdDeleteSel_Click()
'supprime la s�lection
Dim x As Long

    For x = Me.lstSignets.ListItems.Count To 1 Step -1
        If Me.lstSignets.ListItems.Item(x).Selected Then _
            Me.lstSignets.ListItems.Remove x
    Next x
End Sub

Private Sub cmdNew_Click()
'nouveau signet
Dim S As String
Dim s2 As String

    S = InputBox(Lang.GetString("_OffsetNewSignet"), Lang.GetString("_AddSignet"))
    If StrPtr(S) = 0 Then Exit Sub
    s2 = InputBox(Lang.GetString("_NewCommentSignet"), Lang.GetString("_AddSignet"))
    If StrPtr(s2) = 0 Then Exit Sub
    
    'on ajoute le signet
    With Me.lstSignets
        .ListItems.Add Text:=S
        .ListItems.Item(.ListItems.Count).SubItems(1) = s2
    End With
    
End Sub

Private Sub cmdOpen_Click()
'ouvre une liste de signets
    Call AddSignetIn(True)
End Sub

Private Sub cmdSave_Click()
'enregistre la liste des signets de la form active
Dim S As String
Dim lFile As Long
Dim x As Long

    On Error GoTo ErrGestion
    
    If frmContent.ActiveForm Is Nothing Then Exit Sub
    If Me.lstSignets.ListItems.Count = 0 Then Exit Sub 'pas de signets
    
    'enregistrement ==> choix du fichier
    With frmContent.CMD
        .CancelError = True
        .FileName = frmContent.ActiveForm.Caption & ".sig"
        .DialogTitle = Lang.GetString("_ListSave")
        .Filter = Lang.GetString("_ListSignet") & "|*.sig|"
        .InitDir = App.Path
        .FileName = vbNullString
        .ShowSave
        S = .FileName
    End With

    If cFile.FileExists(S) Then
        'message de confirmation
        x = MsgBox(Lang.GetString("_FileAlreadyExists"), vbInformation + vbYesNo, Lang.GetString("_Warning"))
        If Not (x = vbYes) Then Exit Sub
    End If
    
    'ouvre le fchier
    lFile = FreeFile
    Open S For Output As lFile
    
    'enregistre les entr�es
    For x = 1 To lstSignets.ListItems.Count
        Write #lFile, lstSignets.ListItems.Item(x) & "|" & lstSignets.ListItems.Item(x).SubItems(1)
    Next x
    
    Close lFile

    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_SignetSaved"))
    
ErrGestion:
End Sub

Private Sub cmdSaveChanges_Click()
Dim x As Long

    'applique les changements � la forme active
 
    With frmContent.ActiveForm
        'on ajoute tous les signets qui sont affich�s dans l'activeform
        .lstSignets.ListItems.Clear
        For x = 1 To Me.lstSignets.ListItems.Count
            .lstSignets.ListItems.Add Text:=Me.lstSignets.ListItems.Item(x).Text
            .lstSignets.ListItems.Item(x).SubItems(1) = Me.lstSignets.ListItems.Item(x).SubItems(1)
        Next x
        
        'on vire les anciens signets du HW actif et on rajoute les nouveau
        Call .HW.RemoveAllSignets
        For x = 1 To Me.lstSignets.ListItems.Count
            .HW.AddSignet CCur(Val(Me.lstSignets.ListItems.Item(x).Text))
        Next x
        
        Call .HW.Refresh
        
    End With
    
End Sub

Private Sub dmQuit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim x As Long

    With Lang
        #If MODE_DEBUG Then
            If App.LogMode = 0 And CREATE_FRENCH_FILE Then
                'on cr�� le fichier de langue fran�ais
                .Language = "French"
                .LangFolder = LANG_PATH
                .WriteIniFileFormIDEform
            End If
        #End If
        
        If App.LogMode = 0 Then
            'alors on est dans l'IDE
            .LangFolder = LANG_PATH
        Else
            .LangFolder = App.Path & "\Lang"
        End If
        
        'applique la langue d�sir�e aux controles
        Call .ActiveLang(Me): .Language = cPref.env_Lang
        .LoadControlsCaption
    End With
    
    If frmContent.ActiveForm Is Nothing Then Exit Sub
    
    'on ajoute tous les signets qui sont affich�s dans l'activeform
    With lstSignets.ListItems
        For x = 1 To frmContent.ActiveForm.lstSignets.ListItems.Count
            .Add Text:=frmContent.ActiveForm.lstSignets.ListItems.Item(x).Text
            .Item(x).SubItems(1) = frmContent.ActiveForm.lstSignets.ListItems.Item(x).SubItems(1)
        Next x
    End With
    
End Sub

'=======================================================
'ajoute (ou ouvre si overwrite) une liste de signets
'=======================================================
Private Sub AddSignetIn(ByVal bOverWrite As Boolean)
Dim S As String
Dim lFile As Long
Dim x As Long
Dim sTemp As String
Dim l As Long

    On Error GoTo ErrGestion
    
    'ouverture ==> choix du fichier
    With frmContent.CMD
        .CancelError = True
        .DialogTitle = Lang.GetString("_OpenSignet")
        .Filter = Lang.GetString("_ListSignet") & "|*.sig|"
        .InitDir = App.Path
        .ShowOpen
        S = .FileName
    End With
    
    If bOverWrite Then lstSignets.ListItems.Clear
    
    'ouvre le fchier
    lFile = FreeFile
    Open S For Input As lFile
    While Not EOF(lFile)
        Input #lFile, sTemp
        l = InStr(1, sTemp, "|", vbBinaryCompare)
        If l <> 0 Then
            'ajoute aussi un commentaire
            lstSignets.ListItems.Add Text:=Left$(sTemp, l - 1)
            lstSignets.ListItems.Item(lstSignets.ListItems.Count).SubItems(1) = Right$(sTemp, Len(sTemp) - l)
        End If
    Wend
        
    Close lFile
    
    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_SignetAdded"))
ErrGestion:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Lang = Nothing
End Sub

Private Sub lstSignets_ItemClick(ByVal Item As ComctlLib.ListItem)
'va au signet

    If mouseUped Then
        With frmContent.ActiveForm
            .HW.FirstOffset = Val(Item.Text)
            .HW.Refresh
            .VS.Value = .HW.FirstOffset / 16
        End With
        mouseUped = False   '�vite de devoir bouger le HW si l'on s�lectionne pleins d'items
        'par exemple avec Shift
    End If
End Sub

Private Sub lstSignets_KeyDown(KeyCode As Integer, Shift As Integer)
'vire les signets si touche suppr
Dim r As Long

    mouseUped = True
    
    If KeyCode = vbKeyDelete Then
        'touche suppr
        If lstSignets.SelectedItem.Selected Then
            'alors on supprime quelque chose
        
            For r = lstSignets.ListItems.Count To 1 Step -1
                If lstSignets.ListItems.Item(r).Selected Then _
                    lstSignets.ListItems.Remove r
            Next r
        End If
    End If
        
End Sub

Private Sub lstSignets_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim tLst As ListItem
Dim S As String
Dim r As Long

    If Button = 2 Then
        'alors clic droit ==> on affiche la boite de dialogue "commentaire" sur le comment
        'qui a �t� s�lectionn�
        Set tLst = lstSignets.HitTest(x, y)
        If tLst Is Nothing Then Exit Sub
        S = InputBox(Lang.GetString("_AddCommentFor") & " " & tLst.Text, _
            Lang.GetString("_AddComment"))
        If StrPtr(S) <> 0 Then
            'ajoute le commentaire
            tLst.SubItems(1) = S
        End If
    End If
    
    If Button = 4 Then
        'mouse du milieu ==> on supprime le signet
        Set tLst = lstSignets.HitTest(x, y)
        If tLst Is Nothing Then Exit Sub
        
        'on enl�ve du listview
        lstSignets.ListItems.Remove tLst.Index
    End If
        
End Sub

Private Sub lstSignets_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'permet de ne pas changer le HW dans le cas de multiples s�lections
    mouseUped = True
End Sub
