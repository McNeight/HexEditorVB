VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BC0A7EAB-09F8-454A-AB7D-447C47D14F18}#1.0#0"; "ProgressBar_OCX.ocx"
Object = "{C77F04DF-B546-4EBA-AFE7-F46C1BA9BCF4}#1.0#0"; "LanguageTranslator.ocx"
Begin VB.Form frmShredd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Effacement d�finitif de fichiers"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4860
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmShredd.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ProgressBar_OCX.pgrBar PGB 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4680
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      BackColorTop    =   13027014
      BackColorBottom =   15724527
      Value           =   1
      BackPicture     =   "frmShredd.frx":058A
      FrontPicture    =   "frmShredd.frx":05A6
   End
   Begin VB.TextBox txtPass 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Text            =   "3"
      ToolTipText     =   "D�signe le nombre de sanitizations qui seront effectu�es"
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      ToolTipText     =   "Fermer cette fen�tre"
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdProceed 
      Caption         =   "Supprimer d�finitivement"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2543
      TabIndex        =   1
      ToolTipText     =   "D�truit les fichiers (/!\ suppression IRRECUPERABLE)"
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CommandButton cmdAddFile 
      Caption         =   "Ajouter des fichiers..."
      Height          =   375
      Left            =   143
      TabIndex        =   0
      ToolTipText     =   "Permet l'ajout de fichiers � d�truire"
      Top             =   3600
      Width           =   2175
   End
   Begin ComctlLib.ListView LV 
      Height          =   3375
      Left            =   0
      TabIndex        =   3
      Tag             =   "lang_ok"
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDropMode     =   1
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      OLEDropMode     =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Fichier"
         Object.Width           =   7056
      EndProperty
   End
   Begin LanguageTranslator.ctrlLanguage Lang 
      Left            =   0
      Top             =   0
      _ExtentX        =   1402
      _ExtentY        =   1402
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre de sanitizations :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   2055
   End
End
Attribute VB_Name = "frmShredd"
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
'FORM POUR LA SUPPRESSION DEFINITIVE DE FICHIER
'=======================================================

Private Sub cmdAddFile_Click()
'ajoute un fichier � la liste � supprimer
Dim s() As String
Dim s2 As String
Dim x As Long

    ReDim s(0)
    s2 = cFile.ShowOpen(Lang.GetString("_FilesToKillSel"), Me.hWnd, _
        Lang.GetString("_All") & "|*.*", , , , , OFN_EXPLORER + _
        OFN_ALLOWMULTISELECT, s())
    
    For x = 1 To UBound(s())
        If cFile.FileExists(s(x)) Then
            LV.ListItems.Add Text:=s(x) 'ajoute l'�l�ment
        End If
    Next x
    
    'dans le cas d'un fichier simple
    If cFile.FileExists(s2) Then LV.ListItems.Add Text:=s2
    
    CheckBtn    'enable ou non le cmdProceed

ErrCancel:
End Sub

Private Sub cmdProceed_Click()
'proc�de � la suppression d�finitive
Dim x As Long

    'affiche un advertissement
    x = MsgBox(Lang.GetString("_FilesWillBeLost") & vbNewLine & Lang.GetString("_WannaKill"), vbYesNo + vbInformation, Lang.GetString("_War"))
    
    If Not (x = vbYes) Then Exit Sub
    
    If Abs(Int(Val(txtPass.Text))) < 1 Or Abs(Int(Val(txtPass.Text))) > 2048 Then
        'nombre de sanitizations incorrecte
        MsgBox Lang.GetString("_PassNot")
        Exit Sub
    End If
    
    For x = LV.ListItems.Count To 1 Step -1
        DoEvents    'rend quand m�me la main, si bcp de fichiers, c'est utile
        If ShreddFile(LV.ListItems.Item(x), Int(Val(txtPass.Text)), Me.PGB) Then   'proc�de � la suppression
            LV.ListItems.Remove (x) 'enl�ve l'item si la suppression � �chou�
        End If
    Next
    
    'affichage des r�sultats
    If LV.ListItems.Count > 0 Then
        'alors il reste au moins un fichier
        MsgBox Lang.GetString("_OneCannot"), vbInformation, Lang.GetString("_War")
    Else
        'OK
        MsgBox Lang.GetString("_DelOk"), vbOKOnly, Lang.GetString("_DelIsOk")
    End If

    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_DelFin"))
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    #If MODE_DEBUG Then
        If App.LogMode = 0 And CREATE_FRENCH_FILE Then
            'on cr�� le fichier de langue fran�ais
            Lang.Language = "French"
            Lang.LangFolder = LANG_PATH
            Lang.WriteIniFileFormIDEform
        End If
    #End If
    
    If App.LogMode = 0 Then
        'alors on est dans l'IDE
        Lang.LangFolder = LANG_PATH
    Else
        Lang.LangFolder = App.Path & "\Lang"
    End If
    
    'applique la langue d�sir�e aux controles
    Lang.Language = cPref.env_Lang
    Lang.LoadControlsCaption
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)

    If LV.SelectedItem Is Nothing Then Exit Sub
        
    If KeyCode = vbKeyDelete Then
        'alors enleve le fichiers s�lectionn�s
        LV.ListItems.Remove LV.SelectedItem.Index
    End If
    
    CheckBtn    'enable ou non le cmdProceed

End Sub

'=======================================================
'v�rifie que le bouton de suppression est enabled ou pas
'=======================================================
Private Sub CheckBtn()
    Me.cmdProceed.Enabled = (LV.ListItems.Count > 0)
End Sub

Private Sub LV_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long

    'gestion de la d�pose des fichiers sur le listview

    On Error GoTo BadFormat 'pas de drag and drop de fichier
    
    'ajoute les fichers du drag and drop � la liste
    For i = 1 To Data.Files.Count
        If cFile.FileExists(Data.Files.Item(i)) Then LV.ListItems.Add Text:=Data.Files.Item(i)
    Next i
    
BadFormat:
End Sub
