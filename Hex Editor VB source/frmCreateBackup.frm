VERSION 5.00
Object = "{BC0A7EAB-09F8-454A-AB7D-447C47D14F18}#1.0#0"; "ProgressBar_OCX.ocx"
Object = "{C77F04DF-B546-4EBA-AFE7-F46C1BA9BCF4}#1.0#0"; "LanguageTranslator.ocx"
Begin VB.Form frmCreateBackup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Attention"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCreateBackup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin LanguageTranslator.ctrlLanguage Lang 
      Left            =   0
      Top             =   0
      _ExtentX        =   1402
      _ExtentY        =   1402
   End
   Begin VB.Frame Frame1 
      Caption         =   "Avancement du backup"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   4455
      Begin ProgressBar_OCX.pgrBar pgrBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         BackColorTop    =   13027014
         BackColorBottom =   15724527
         Value           =   1
         BackPicture     =   "frmCreateBackup.frx":000C
         FrontPicture    =   "frmCreateBackup.frx":0028
      End
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "Non"
      Height          =   375
      Left            =   2873
      TabIndex        =   1
      ToolTipText     =   "Annule la proc�dure"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Oui"
      Height          =   375
      Left            =   713
      TabIndex        =   0
      ToolTipText     =   "Proc�de au changement"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   $"frmCreateBackup.frx":0044
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmCreateBackup"
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
'FORM PERMETTANT DE CREER UN BAKCUP DE FICHIER
'=======================================================


'=======================================================
'VARIABLES PRIVESS
'=======================================================
'Private strFile As String
'Private Frm As Pfm
'Private tAction As BACKUP_ACTION

'=======================================================
'ENUMS
'=======================================================
'Public Enum BACKUP_ACTION
'    DeleteZone
'    AddZone
'End Enum


    
Private Sub cmdNO_Click()
    bAcceptBackup = False 'accepte PAS le backup
    Unload Me
End Sub

Private Sub cmdOk_Click()
    'Me.Height = 3135
    bAcceptBackup = True    'on accepte le backup
    Unload Me   'la form sera recharg�e juste apr�s pour l'affichage de la progression
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

'=======================================================
'obtient le nom du fichier dont on doit faire le backup
'la form qui contient le HW et donc les modifs
'le type de modif � faire
'les param�tres des modifs
'=======================================================
'Public Sub GetAction(ByVal sFile As String, ByVal frmFrom As Pfm, _
'ByVal TypeOfAction As BACKUP_ACTION, Optional ByVal Param1 As Variant, _
'Optional ByVal Param2 As Variant, Optional ByVal Param3 As Variant)

    'sauvegarde les param�tres dans les variables priv�es
'    strFile = sFile
'    Set Frm = frmFrom
'    tAction = TypeOfAction
    
'End Sub

