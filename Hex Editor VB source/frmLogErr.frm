VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmLogErr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rapport d'erreurs"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8475
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   44
   Icon            =   "frmLogErr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   7290
      TabIndex        =   3
      ToolTipText     =   "Fermer la fenêtre"
      Top             =   218
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   1125
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   98
      Width           =   6975
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset..."
      Height          =   375
      Left            =   7290
      TabIndex        =   0
      ToolTipText     =   "Supprimer toutes les entrées du rapport d'erreur"
      Top             =   818
      Width           =   1095
   End
   Begin ComctlLib.ListView LV 
      Height          =   2775
      Left            =   90
      TabIndex        =   2
      Tag             =   "lang_ok"
      Top             =   1298
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Date"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Heure"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Zone"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Source"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Numéro d'erreur"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Description"
         Object.Width           =   7056
      EndProperty
   End
End
Attribute VB_Name = "frmLogErr"
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
'FORM D'AFFICHAGE DU RAPPORT D'ERREUR
'=======================================================
Private Lang As New clsLang

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdReset_Click()
    'supprime le log
    If clsERREUR.DeleteLogFile <> 0 Then LV.ListItems.Clear
    Call Form_Load
    
    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_LogDel"))
End Sub

Private Sub Form_Load()
Dim var As Variant
Dim x As Long

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
    
    'obtient les infos sur les erreurs
    var = clsERREUR.GetErrors
    
    'affiche tout çà dans le LV
    LV.ListItems.Clear
    
    With LV.ListItems
        For x = 1 To clsERREUR.NumberOfErrorInLogFile
            .Add Text:=var(x).ErrDate
            .Item(x).SubItems(1) = var(x).ErrTime
            .Item(x).SubItems(2) = var(x).ErrZone
            .Item(x).SubItems(3) = var(x).ErrSource
            .Item(x).SubItems(4) = var(x).ErrNumber
            .Item(x).SubItems(5) = var(x).ErrDescription
        Next x
    End With
    
    With Text1
        If clsERREUR.NumberOfErrorInLogFile <> 0 Then
            'il y a des erreurs
            .ForeColor = RED_COLOR
            .Text = Lang.GetString("_ErrorRap") & vbNewLine & _
                Lang.GetString("_PleaseSend") & vbNewLine & _
                clsERREUR.LogFile & vbNewLine & Lang.GetString("_At") & _
                " hexeditorvb@gmail.com" & vbNewLine & Lang.GetString("_YouCont")
        Else
            'pas d'erreurs
            .ForeColor = GREEN_COLOR
            .Text = Lang.GetString("_NoError")
        End If
    End With
End Sub
