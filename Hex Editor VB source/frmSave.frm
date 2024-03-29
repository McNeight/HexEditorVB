VERSION 5.00
Object = "{16DCE99A-3937-4772-A07F-3BA5B09FCE6E}#1.1#0"; "vkUserControlsXP.ocx"
Begin VB.Form frmSave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sauvegarder le fichier"
   ClientHeight    =   1785
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
   HelpContextID   =   44
   Icon            =   "frmSave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin vkUserContolsXP.vkCommand cmdYES 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "Proc�de � la sauvegarde et �crase le fichier"
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Oui"
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
   Begin vkUserContolsXP.vkCommand cmdNO 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      ToolTipText     =   "Ne proc�de pas � la sauvegarde"
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Non"
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
   Begin vkUserContolsXP.vkCheck chkDoNotShowAlert 
      Height          =   255
      Left            =   615
      TabIndex        =   1
      ToolTipText     =   $"frmSave.frx":058A
      Top             =   1440
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   450
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "Ne plus afficher � l'avenir"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Voulez vous r�ellement sauvegarder ? (le fichier actuel sera �cras�)"
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
      Top             =   188
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
'FORM QUI PROPOSE D'ECRASER UN FICHIER (SAUVEGARDE)
'=======================================================
Private Lang As New clsLang

Private Sub cmdNO_Click()
    Call SavePrefShowAlert   'sauvegarde (ou pas) les pref
    Unload Me
End Sub

Private Sub cmdYES_Click()
'proc�de � la sauvegarde
    
    

    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_SaveOk"))
    
    Call cmdNO_Click    'quitte la form en sauvant les pref
End Sub

'=======================================================
'sauvegarde les pr�f�rences avec le checkbox si n�cessaire
'=======================================================
Private Sub SavePrefShowAlert()

    If (cPref.general_ShowAlert + chkDoNotShowAlert.Value) <> 1 Then
        'alors les deux sont identiques (0+0=0 ou 1+1=2 ==> pas de 0+1 ou 1+0)
        cPref.general_ShowAlert = Abs(chkDoNotShowAlert.Value - 1)  '0==>1, 1==>0
        
        'lance la sauvegarde
        Call clsPref.SaveIniFile(cPref)
    End If
End Sub

Private Sub Form_Load()

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
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Lang = Nothing
End Sub
