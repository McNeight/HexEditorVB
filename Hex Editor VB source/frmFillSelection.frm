VERSION 5.00
Begin VB.Form frmFillSelection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insertion / remplissage de bytes"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4485
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   18
   Icon            =   "frmFillSelection.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   4485
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Méthode"
      Height          =   1455
      Index           =   0
      Left            =   135
      TabIndex        =   8
      Top             =   83
      Width           =   3735
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   0
         Left            =   120
         ScaleHeight     =   1095
         ScaleWidth      =   3495
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   3495
         Begin VB.OptionButton Option1 
            Caption         =   "Valeurs hexa"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   16
            Tag             =   "pref0"
            ToolTipText     =   "Remplit avec un byte fixe"
            Top             =   0
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Random entre "
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   15
            Tag             =   "pref1"
            ToolTipText     =   "Remplit avec byte sélectionné au hasard"
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Byte dans la liste"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   14
            Tag             =   "pref2"
            ToolTipText     =   "Remplit avec un byte choisi dans une liste"
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtByte 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1680
            TabIndex        =   13
            Tag             =   "pref"
            Text            =   "55"
            ToolTipText     =   "Valeur hexa de remplissement"
            Top             =   0
            Width           =   975
         End
         Begin VB.TextBox txtBorneInf 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            TabIndex        =   12
            Tag             =   "pref"
            Text            =   "0"
            ToolTipText     =   "Borne inférieure du random (0-254)"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtBorneSup 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   2760
            TabIndex        =   11
            Tag             =   "pref"
            Text            =   "FF"
            ToolTipText     =   "Borne supérieure du random (1-255)"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtList 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            TabIndex        =   10
            Tag             =   "pref"
            Text            =   "00 55 AA FF"
            ToolTipText     =   "Liste des bytes (séparer par un espace les paquets de 2)"
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "et"
            Height          =   255
            Left            =   2400
            TabIndex        =   17
            Top             =   360
            Width           =   255
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Passes"
      Height          =   1815
      Index           =   1
      Left            =   135
      TabIndex        =   2
      Top             =   1643
      Width           =   3015
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1455
         Index           =   1
         Left            =   120
         ScaleHeight     =   1455
         ScaleWidth      =   2775
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   2775
         Begin VB.ListBox lstPasses 
            Height          =   1425
            Left            =   0
            TabIndex        =   7
            ToolTipText     =   "Liste des passes à appliquer"
            Top             =   0
            Width           =   1575
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Ajouter"
            Height          =   375
            Left            =   1680
            TabIndex        =   6
            ToolTipText     =   "Ajouter la passe en cours"
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Enlever"
            Height          =   375
            Left            =   1680
            TabIndex        =   5
            ToolTipText     =   "Supprimer la passe sélectionnée"
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton cmdSanitization 
            Caption         =   "Sanitization"
            Height          =   375
            Left            =   1680
            TabIndex        =   4
            ToolTipText     =   "Ajouter les 3 passes de sanitization"
            Top             =   960
            Width           =   975
         End
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Appliquer"
      Height          =   375
      Left            =   3255
      TabIndex        =   1
      ToolTipText     =   "Appliquer les passes"
      Top             =   2243
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   3255
      TabIndex        =   0
      ToolTipText     =   "Fermer cette fenêtre"
      Top             =   2843
      Width           =   1095
   End
End
Attribute VB_Name = "frmFillSelection"
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
'FORM POUR REMPLIR LE BLOC SELECTIONNE
'=======================================================

'=======================================================
'VARIABLES PRIVEES
'=======================================================
Private Lang As New clsLang
Private clsPref As clsIniForm
Private tPasses() As PASSE_TYPE


Private Sub cmdAdd_Click()
'ajoute une passe
        
    lstPasses.AddItem Lang.GetString("_Pass") & " " & CStr(UBound(tPasses) + 1)

    'initialise la dernière passe
    With tPasses(UBound(tPasses))
        .sData1 = vbNullString
        .sData2 = vbNullString
        .tType = FixedByte
    End With
    
    ReDim Preserve tPasses(UBound(tPasses) + 1) 'ajoute une case à la liste

End Sub

Private Sub cmdApply_Click()
'applique les différentes passes
Dim cur1 As Currency
Dim cur2 As Currency
Dim sFile As String
    
    If frmContent.ActiveForm Is Nothing Then Exit Sub
    
    With frmContent.ActiveForm.HW
        'récupère les bornes de la sélection
        cur1 = .FirstSelectionItem.Offset + .FirstSelectionItem.Col - 1
        cur2 = .SecondSelectionItem.Offset + .SecondSelectionItem.Col - 1
    End With
    
    '>10Mo ==> demande confirmation
    If (cur2 - cur1) > 10000000 Then
        If MsgBox(Lang.GetString("_SelIsBig"), vbInformation + vbYesNo, _
            Lang.GetString("_War")) <> vbYes Then Exit Sub
    End If
    
    '//UNIQUEMENT POUR UN FICHIER !!!
    'affiche la demande de création de backup
    frmCreateBackup.Show vbModal
    
    If bAcceptBackup Then
        'alors l'user a accepté la création du backup
        
        'récupère un nom de fichier temp
        Call ObtainTempPathFile("temp_file", sFile, "tmp")
        
        'créé une copie (lance une sauvegarde)
        Call frmContent.ActiveForm.GetNewFile(sFile)
    Else
        Exit Sub
    End If
    
    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_PassCour"))
    
    
    Select Case TypeOfForm(frmContent.ActiveForm)
        Case "Fichier"
            
            'on applique au fichier les différentes passes au fichier
            'on fait çà en temporaire
            Call ApplyPass_File(cur1, cur2, frmContent.ActiveForm.HW, tPasses(), _
                sFile)
            
        Case "Processus"
        
        
        Case "Disque"
        
        
        Case "Disque physique"
        
        
        Case Else
            Exit Sub
    End Select
        

    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_PassOk"))
    
    Unload Me
End Sub

Private Sub cmdDelete_Click()
'enlève un élément de la liste
Dim l As Long
Dim X As Long
Dim tPTmp() As PASSE_TYPE

    On Error GoTo ErrGestion

    l = lstPasses.ListIndex 'éléments à enlever de tPasses
    tPTmp = tPasses 'backup
    
    If l < 0 Then Exit Sub
    
    'redimensionne le tableau
    ReDim tPasses(UBound(tPTmp) - 1)
    
    If UBound(tPTmp) = 0 Then Exit Sub   'rien à enlever
    
    For X = 0 To l - 1
        tPasses(X) = tPTmp(X)
    Next X
    For X = l + 1 To UBound(tPTmp) - 1
        tPasses(X - 1) = tPTmp(X)
    Next X
    
    Call lstPasses.Clear   'enlève les éléments de la liste
    
    'rajoute n-1 passes
    For X = 1 To UBound(tPTmp) - 1
        lstPasses.AddItem Lang.GetString("_Pass") & " " & CStr(X)
    Next X
    
    lstPasses.ListIndex = l - 1
    Call lstPasses_Click
    
    Exit Sub
ErrGestion:
    clsERREUR.AddError "disfrmFillSelection.cmdDeleteClick", True
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdSanitization_Click()
'sanitization
'procède à l'enregistrement de 3 passes successives sur la zone sélectionnée
'1) passe qui remplit avec la valeur hexa 0x55 (01010101 en binaire)
'2) passe qui remplit avec la valeur hexa 0xAA (10101010 en binaire)
'3) Random passe

    'suppression des passes actuelles
    ReDim tPasses(3)
    lstPasses.Clear
    
    'ajout des 3 passes
    With lstPasses
        .AddItem Lang.GetString("_Pass") & " 1"
        .AddItem Lang.GetString("_Pass") & " 2"
        .AddItem Lang.GetString("_Pass") & " 3"
    End With
    
    tPasses(0).sData1 = "55"
    tPasses(0).sData2 = ""
    tPasses(0).tType = FixedByte
    tPasses(1).sData1 = "AA"
    tPasses(1).sData2 = ""
    tPasses(1).tType = FixedByte
    tPasses(2).sData1 = "0"
    tPasses(2).sData2 = "FF"
    tPasses(2).tType = RandomByte
End Sub

Private Sub Form_Load()

    'loading des preferences
    Set clsPref = New clsIniForm
    clsPref.GetFormSettings App.Path & "\Preferences\FillSelection.ini", Me
    
    With Lang
        #If MODE_DEBUG Then
            If App.LogMode = 0 And CREATE_FRENCH_FILE Then
                'on créé le fichier de langue français
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
        
        'applique la langue désirée aux controles
        Call .ActiveLang(Me): .Language = cPref.env_Lang
        Call .LoadControlsCaption
    End With
    
    ReDim tPasses(0)   'initialize array

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'sauvegarde des preferences
    clsPref.SaveFormSettings App.Path & "\Preferences\FillSelection.ini", Me
    Set clsPref = Nothing
    Set Lang = Nothing
End Sub

Private Sub lstPasses_Click()
'affiche la passe sélectionnée

    If lstPasses.ListIndex < 0 Then Exit Sub
    
    Option1(tPasses(lstPasses.ListIndex).tType).Value = True    'sélection du type de passe
    
    'remplit les champs
    With tPasses(lstPasses.ListIndex)
        Select Case .tType
            Case FixedByte
                txtByte.Text = .sData1
            Case RandomByte
                txtBorneInf.Text = .sData1
                txtBorneSup.Text = .sData2
            Case Else
                txtList.Text = .sData1
        End Select
    End With
    
End Sub

Private Sub Option1_Click(Index As Integer)
'enabled ou pas certains éléments

    txtByte.Enabled = False
    txtBorneInf.Enabled = False
    txtBorneSup.Enabled = False
    txtList.Enabled = False
    
    If Index = 0 Then txtByte.Enabled = True
    If Index = 1 Then txtBorneInf.Enabled = True: txtBorneSup.Enabled = True
    If Index = 2 Then txtList.Enabled = True
    
    'change le type de passe de la passe sélectionnée
    If lstPasses.ListIndex <> -1 Then tPasses(lstPasses.ListIndex).tType = Index
    
End Sub

Private Sub txtBorneInf_Change()
'change le sData1 de la passe
    If lstPasses.ListIndex <> -1 Then tPasses(lstPasses.ListIndex).sData1 = txtBorneInf.Text
End Sub

Private Sub txtBorneSup_Change()
'change le sData2 de la passe
    If lstPasses.ListIndex <> -1 Then tPasses(lstPasses.ListIndex).sData2 = txtBorneSup.Text
End Sub

Private Sub txtByte_Change()
'change le sData1 de la passe
    If lstPasses.ListIndex <> -1 Then tPasses(lstPasses.ListIndex).sData1 = txtByte.Text
End Sub

Private Sub txtList_Change()
'change le sData1 de la passe
    If lstPasses.ListIndex <> -1 Then tPasses(lstPasses.ListIndex).sData1 = txtList.Text
End Sub
