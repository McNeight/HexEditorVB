VERSION 5.00
Begin VB.Form frmSave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sauvegarder le fichier"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3510
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDoNotShowAlert 
      Caption         =   "Ne plus afficher à l'avenir"
      Height          =   195
      Left            =   248
      TabIndex        =   3
      ToolTipText     =   $"frmSave.frx":058A
      Top             =   1800
      Width           =   3015
   End
   Begin VB.CommandButton cmdNO 
      Caption         =   "Non"
      Height          =   495
      Left            =   2048
      TabIndex        =   2
      ToolTipText     =   "Ne procède pas à la sauvegarde"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdYES 
      Caption         =   "Oui"
      Height          =   495
      Left            =   248
      TabIndex        =   1
      ToolTipText     =   "Procède à la sauvegarde et écrase le fichier"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Voulez vous réellement sauvegarder ? (le fichier actuel sera écrasé)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   128
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmSave"
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
'FORM QUI PROPOSE D'ECRASER UN FICHIER (SAUVEGARDE)
'=======================================================

Private Sub cmdNO_Click()
    SavePrefShowAlert   'sauvegarde (ou pas) les pref
    Unload Me
End Sub

Private Sub cmdYES_Click()
'procède à la sauvegarde
    
    
    
    
    Call cmdNO_Click    'quitte la form en sauvant les pref
End Sub

'=======================================================
'sauvegarde les préférences avec le checkbox si nécessaire
'=======================================================
Private Sub SavePrefShowAlert()

    If (cPref.general_ShowAlert + chkDoNotShowAlert.Value) <> 1 Then
        'alors les deux sont identiques (0+0=0 ou 1+1=2 ==> pas de 0+1 ou 1+0)
        cPref.general_ShowAlert = Abs(chkDoNotShowAlert.Value - 1)  '0==>1, 1==>0
        
        'lance la sauvegarde
        Call clsPref.SaveIniFile(cPref)
    End If
End Sub
