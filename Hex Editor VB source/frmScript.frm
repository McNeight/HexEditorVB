VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmScript 
   Caption         =   "Editeur de script"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   9165
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   24
   Icon            =   "frmScript.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5610
   ScaleWidth      =   9165
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   360
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":058A
            Key             =   "Script|ExécuterF5"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":08DC
            Key             =   "Fichier|Ouvrir..."
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":0C2E
            Key             =   "Fichier|Nouveau"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":0F80
            Key             =   "Script|Vérifier la cohérenceF9"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":12D2
            Key             =   "Fichier|Imprimer..."
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":1624
            Key             =   "Fichier|Enregistrer..."
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":1976
            Key             =   "Aide|Aide...F1"
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lst 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      ItemData        =   "frmScript.frx":1CC8
      Left            =   5160
      List            =   "frmScript.frx":1DA7
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   720
      Width           =   3550
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4895
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmScript.frx":2307
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CMD 
      Left            =   360
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":2387
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":3D19
            Key             =   ""
            Object.Tag             =   "Copy"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":56AB
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":703D
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":89CF
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":A361
            Key             =   "Signet"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":BCF3
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":C28D
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":C827
            Key             =   "Left"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":E1B9
            Key             =   "Right"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":FB4B
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":114DD
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":12E6F
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":14801
            Key             =   "Trash"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":16193
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":17B25
            Key             =   "FolderOpen"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":194B7
            Key             =   "FileOpen"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":1AE49
            Key             =   "Computer"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":1C7DB
            Key             =   "Settings"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":1E16D
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":205BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":20AD0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "Créer un nouveau fichier"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OpenFile"
            Object.ToolTipText     =   "Ouvrir un fichier"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Sauvegarder l'objet ouvert"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Imprimer"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Search"
            Object.ToolTipText     =   "Effectuer une recherche dans l'objet"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Couper la sélection"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copier la sélection"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Coller le contenu du presse-papier"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Undo"
            Object.ToolTipText     =   "Défaire"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Redo"
            Object.ToolTipText     =   "Refaire"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.Menu rmnuFile 
      Caption         =   "&Fichier"
      Begin VB.Menu mnuNew 
         Caption         =   "&Nouveau"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Ouvrir..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "&Enregistrer..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Imprimer..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileTiret1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quitter"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu rmnuScript 
      Caption         =   "&Script"
      Begin VB.Menu mnuCheck 
         Caption         =   "&Vérifier la cohérence"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuLauch 
         Caption         =   "&Exécuter"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu rmnuHelp 
      Caption         =   "&Aide"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Aide..."
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmScript"
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
'FORM D'EDITION/CREATION DE SCRIPTS
'=======================================================

Private Lang As New clsLang
Private bIsModified As Boolean  'contient si le fichier est modifié ou non


Private Sub Form_Load()

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
        .LoadControlsCaption
    End With
    
    Call AddIconsToMenus(Me.hWnd, Me.ImageList2)    'ajoute les icones au menu
    bIsModified = False 'pas de modification actuellement

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With RTB
        .Left = 0
        .Top = 340
        .Height = Me.Height - 1140
        .Width = Me.Width - 3650
    End With
    With lst
        .Top = 340
        .Height = RTB.Height
        .Left = RTB.Width
    End With
End Sub

Private Sub lst_DblClick()
'ajoute au richtextbox la commande double cliquée
Dim s As String
Dim lg As Long
Dim ld As Long
Dim len1 As Long

    s = lst.List(lst.ListIndex)  'contient la commande à ajouter

    len1 = Len(RTB.Text)
    RTB.Text = RTB.Text & s
    
    'change la zone sélectionnée
    'détermine où se situe la première zone entre <> dans la string s
    lg = InStr(1, s, "<", vbBinaryCompare)
    ld = InStr(1, s, ">", vbBinaryCompare)
    
    RTB.SelStart = len1 + lg
    RTB.SelLength = IIf((ld - lg - 1) > 0, ld - lg - 1, 0)  'définit la taille de la sélection
    If (lg - ld) = 0 Then
        'alors pas de <>
        RTB.SelStart = len1 + Len(s)
        RTB.SelLength = 0
    End If
        
    RTB.SetFocus    'donne le focus au controle RTB
   
End Sub

Private Sub mnuCheck_Click()
'vérifie la cohérence du script
Dim pb As Long

    pb = IsScriptCorrect(RTB.Text)
    If pb = 0 Then
        'correct
        MsgBox Lang.GetString("_ScriptOk"), vbInformation + vbOKOnly, _
            Lang.GetString("_ScriptEd")
    Else
        'problème quelque part
        MsgBox Lang.GetString("_ScriptNot") & vbNewLine & _
            Lang.GetString("_PbLine") & " " & CStr(pb), vbCritical + vbOKOnly, _
            Lang.GetString("_ScriptEd")
    End If
End Sub

Private Sub mnuLauch_Click()
'exécute le script

    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_ScriptLau"))
End Sub

Private Sub mnuNew_Click()
'nouveau fichier
    If bIsModified Then
        'alors le fichier est modifié, on demande confirmation
        If MsgBox(Lang.GetString("_ScriptWillBeLost"), vbInformation + vbYesNo, Lang.GetString("_War")) <> vbYes Then Exit Sub
    End If
    
    RTB.Text = vbNullString 'efface le contenu
End Sub
Private Sub mnuOpen_Click()
'ouverture de fichier
Dim s As String
Dim x As Long

    On Error GoTo CancelPushed
    
    If bIsModified Then
        'alors le fichier est modifié, on demande confirmation
        If MsgBox(Lang.GetString("_ScriptWillBeLost"), vbInformation + _
            vbYesNo, Lang.GetString("_War")) <> vbYes Then Exit Sub
    End If
    
    With Me.CMD
        .CancelError = True
        .DialogTitle = Lang.GetString("_OpenFile")
        .Filter = "Hex Editor Script|*.hescr|Tous|*.*|"
        .ShowOpen
        s = .Filename
    End With
    
    'sauvegarde du fichier
    Call RTB.LoadFile(s)
    
    bIsModified = False 'le fichier a été sauvegardé
    
CancelPushed:
End Sub
Private Sub mnuPrint_Click()
'impression

End Sub
Private Sub mnuQuit_Click()

    If bIsModified Then
        'alors le fichier est modifié, on demande confirmation
        If MsgBox(Lang.GetString("_ScriptWillBeLost"), vbInformation + _
            vbYesNo, Lang.GetString("_War")) <> vbYes Then Exit Sub
    End If
    
    Unload Me
End Sub
Private Sub mnuSaveAs_Click()
'sauvegarde
Dim s As String
Dim x As Long

    On Error GoTo CancelPushed
    
    With Me.CMD
        .CancelError = True
        .DialogTitle = Lang.GetString("_SaveAs")
        .Filter = "Hex Editor Script|*.hescr|Tous|*.*|"
        .Filename = vbNullString
        .ShowSave
        s = .Filename
    End With
    
    If cFile.FileExists(s) Then
        'message de confirmation
        x = MsgBox(Lang.GetString("_FileAlreadyExists"), vbInformation + _
            vbYesNo, Lang.GetString("_War"))
        If Not (x = vbYes) Then Exit Sub
    End If
    
    'sauvegarde du fichier
    Call RTB.SaveFile(s)
    
    bIsModified = False 'le fichier a été sauvegardé
    
CancelPushed:
End Sub

Private Sub RTB_Change()
    bIsModified = True  'le fichier a été modifié
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'appui sur les icones

    Select Case Button.Key
    
        Case "OpenFile"
            Call mnuOpen_Click
        Case "New"
            Call mnuNew_Click
        Case "Copy"
            
        Case "Save"
            Call mnuSaveAs_Click
        Case "Print"
            Call mnuPrint_Click
        Case "Search"
            
        Case "Undo"
            RTB.SetFocus
            Call SendKeys("^Z")
        Case "Redo"
            RTB.SetFocus
            Call SendKeys("^Y")
    End Select

End Sub
