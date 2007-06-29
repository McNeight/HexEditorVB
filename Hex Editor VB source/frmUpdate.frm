VERSION 5.00
Object = "{16DCE99A-3937-4772-A07F-3BA5B09FCE6E}#1.1#0"; "vkUserControlsXP.ocx"
Begin VB.Form frmUpdate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mise à jour de Hex Editor VB"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5430
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin vkUserContolsXP.vkCommand cmdQuit 
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   3240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
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
   Begin vkUserContolsXP.vkCommand cmdCheckMAJ 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      Caption         =   "Vérifier l'existence d'une mise à jour"
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
   Begin vkUserContolsXP.vkTextBox txtInfo 
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5318
      BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2
      LegendText      =   "Informations de mise à jour"
      LegendForeColor =   12937777
      LegendType      =   1
   End
End
Attribute VB_Name = "frmUpdate"
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
'FORM DE MISE A JOUR
'=======================================================

Private cMAJ As clsInternet
Private Lang As New clsLang

Private Sub cmdCheckMAJ_Click()
Dim Ret As Long
Dim S As String
Dim sNotes As String
Dim sNew As String
Dim sRes As String
Dim PageDL As String
Dim OkForDL As Boolean

    cmdCheckMAJ.Enabled = False

    txtInfo.Text = txtInfo.Text & vbNewLine & vbNewLine & Lang.GetString("_RetInfo")
    DoEvents

    'lance la récupération des infos sur une MAJ
    Ret = cMAJ.CheckUpdate(sNotes, sNew, PageDL)
    
    DoEvents
    
    'on récupère les infos sur le téléchargement
    Select Case Ret
        Case -1
            S = Lang.GetString("_Cannot")
        Case 0
            S = Lang.GetString("_UpToDate")
        Case Else
            S = Lang.GetString("_NewAvailable") & vbNewLine & _
                vbNewLine & Lang.GetString("_RetNotes") & " " & sNew & "..."
            
            txtInfo.Text = txtInfo.Text & vbNewLine & vbNewLine & S
            DoEvents
            
            'on récupère les infos de mise à jour
            Ret = cMAJ.GetNotes(sNotes, sRes)
            
            If Ret = -1 Then
                'alors c'est raté
                S = Lang.GetString("_NotesCannot")
            Else
                S = vbNewLine & sRes
            End If
            
            OkForDL = True
            
    End Select
    txtInfo.Text = txtInfo.Text & vbNewLine & vbNewLine & S
    
    cmdCheckMAJ.Enabled = True
    
    If OkForDL Then
        Ret = MsgBox(Lang.GetString("_Wanna"), vbInformation + vbYesNo, Lang.GetString("_NewOne"))
        If Ret <> vbNo Then
            'alors on lance la page de téléchargement
            Call cMAJ.ShellOpenFile(PageDL, Me.hWnd, , App.Path)
        End If
    End If
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

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
    
    'instancie la classe
    Set cMAJ = New clsInternet
    
    'mise à jour du texte
    txtInfo.Text = Lang.GetString("_WelCome") & vbNewLine & vbNewLine & Lang.GetString("_HowTo")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'libère la classe
    Set cMAJ = Nothing
    Set Lang = Nothing
End Sub
