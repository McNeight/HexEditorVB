VERSION 5.00
Begin VB.Form frmNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "           Nouveau fichier"
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
   Icon            =   "frmNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   2715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNo 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   1470
      TabIndex        =   3
      ToolTipText     =   "Fermer cette fenêtre"
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Créer"
      Height          =   375
      Left            =   270
      TabIndex        =   2
      ToolTipText     =   "Créer le fichier (emplacement dans les fichiers temporaires)"
      Top             =   1200
      Width           =   975
   End
   Begin VB.ComboBox cdUnit 
      Height          =   315
      ItemData        =   "frmNew.frx":000C
      Left            =   1200
      List            =   "frmNew.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "pref"
      ToolTipText     =   "Unité"
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtSize 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Tag             =   "pref"
      Text            =   "100"
      ToolTipText     =   "Taille"
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Taille du fichier"
      Height          =   255
      Left            =   390
      TabIndex        =   4
      Top             =   240
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

Private clsPref As clsIniForm

Private Sub cmdNo_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
'créé le fichier
Dim Frm As Form
Dim sFile As String
Dim lFile As Long
Dim lLen As Double
Dim s As String
    
    
    'affiche une nouvelle fenêtre
    Set Frm = New Pfm
    
    'calcule la taille du fichier
    lLen = Abs(Val(txtSize.Text))
    If cdUnit.Text = "Ko" Then lLen = lLen * 1024
    If cdUnit.Text = "Mo" Then lLen = (lLen * 1024) * 1024
    If cdUnit.Text = "Go" Then lLen = ((lLen * 1024) * 1024) * 1024
    
    lLen = Int(lLen)
    
    Unload Me
        
    'obtient un path temporaire
    ObtainTempPathFile "new" & CStr(lNbChildFrm), sFile, vbNullString
    
    'créé le fichier
    
    'obtient un numéro de fichier disponible
    lFile = FreeFile
    
    Open sFile For Binary Access Write As lFile
        Put lFile, , String$(lLen, Chr$(0))
    Close lFile
    
    Call Frm.GetFile(sFile)
    Frm.Show
    lNbChildFrm = lNbChildFrm + 1
    frmContent.Sb.Panels(2).Text = "Ouvertures=[" & CStr(lNbChildFrm) & "]"
    
    Unload Me

End Sub

Private Sub Form_Load()
    'loading des preferences
    Set clsPref = New clsIniForm
    clsPref.GetFormSettings App.Path & "\Preferences\NewFile.ini", Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'sauvegarde des preferences
    clsPref.SaveFormSettings App.Path & "\Preferences\NewFile.ini", Me
    Set clsPref = Nothing
End Sub

Private Sub txtSize_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdOk_Click
End Sub
