VERSION 5.00
Begin VB.Form frmNew 
   BackColor       =   &H00F9E5D9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nouveau fichier"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   2715
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
   Icon            =   "frmNew.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   2715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSize 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   150
      TabIndex        =   3
      Tag             =   "pref"
      Text            =   "100"
      ToolTipText     =   "Taille"
      Top             =   660
      Width           =   975
   End
   Begin VB.ComboBox cdUnit 
      Height          =   315
      ItemData        =   "frmNew.frx":058A
      Left            =   1230
      List            =   "frmNew.frx":059A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "pref lang_ok"
      ToolTipText     =   "Unité"
      Top             =   660
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Créer"
      Height          =   375
      Left            =   300
      TabIndex        =   1
      ToolTipText     =   "Créer le fichier (emplacement dans les fichiers temporaires)"
      Top             =   1140
      Width           =   975
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   1500
      TabIndex        =   0
      ToolTipText     =   "Fermer cette fenêtre"
      Top             =   1140
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Taille du fichier"
      Height          =   255
      Left            =   420
      TabIndex        =   4
      Top             =   180
      Width           =   1935
   End
End
Attribute VB_Name = "frmNew"
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
'FORM QUI INVITE A CREER UN NOUVEAU FICHIER DONT ON DEFINIT LA TAILLE
'=======================================================

Private Lang As New clsLang
Private clsPref As clsIniForm

Private Sub cmdNO_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
'créé le fichier
Dim Frm As Form
Dim sFile As String
Dim lFile As Long
Dim lLen As Double
Dim s As String
    
    On Error GoTo ErrGestion
    
    'affiche une nouvelle fenêtre
    Set Frm = New Pfm
    
    'calcule la taille du fichier
    If Len(txtSize.Text) = 0 Or Len(cdUnit.Text) = 0 Or Val(txtSize.Text) <= 0 Then
        'rien sélectionné
        MsgBox Lang.GetString("_HaveToSelValid"), vbInformation, Lang.GetString("_War")
        Exit Sub
    End If
    
    lLen = Abs(Val(txtSize.Text))
    With Lang
        If cdUnit.Text = .GetString("_Ko") Then lLen = lLen * 1024
        If cdUnit.Text = .GetString("_Mo") Then lLen = (lLen * 1024) * 1024
        If cdUnit.Text = .GetString("_Go") Then lLen = ((lLen * 1024) * 1024) * 1024
    End With
    
    lLen = Int(lLen)
    
    Unload Me
    
    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_CreateNewFile"))
        
    'obtient un path temporaire
    Call ObtainTempPathFile("new" & CStr(lNbChildFrm), sFile, vbNullString)
    
    'créé le fichier
    
    'obtient un numéro de fichier disponible
    lFile = FreeFile
    
    Open sFile For Binary Access Write As lFile
        Put lFile, , String$(lLen, vbNullChar)
    Close lFile
    
    Call Frm.GetFile(sFile)
    Frm.Show
    lNbChildFrm = lNbChildFrm + 1
    frmContent.Sb.Panels(2).Text = Lang.GetString("_Openings") & CStr(lNbChildFrm) & "]"
    
    Unload Me
    
ErrGestion:
End Sub

Private Sub Form_Load()

    Set clsPref = New clsIniForm
    
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
    
    'loading des preferences
    Call clsPref.GetFormSettings(App.Path & "\Preferences\NewFile.ini", Me)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'sauvegarde des preferences
    Call clsPref.SaveFormSettings(App.Path & "\Preferences\NewFile.ini", Me)
    Set clsPref = Nothing
End Sub

Private Sub txtSize_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdOk_Click
End Sub
