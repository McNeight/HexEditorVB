VERSION 5.00
Object = "{9B9A881F-DBDC-4334-BC23-5679E5AB0DC6}#1.1#0"; "FileView_OCX.ocx"
Begin VB.Form frmDrive 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "                     Sélection d'un disque"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4125
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDrive.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FileView_OCX.FileView FV 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   6376
      AllowDirectoryEntering=   0   'False
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
   Begin VB.CommandButton cmdNo 
      Caption         =   "Fermer"
      Height          =   495
      Left            =   2355
      TabIndex        =   2
      ToolTipText     =   "Fermer cette fenêtre"
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ouvrir..."
      Height          =   495
      Left            =   675
      TabIndex        =   1
      ToolTipText     =   "Ouvrir ce lecteur"
      Top             =   3720
      Width           =   1095
   End
End
Attribute VB_Name = "frmDrive"
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
'FORM PERMETTANT LE CHOIX DU DRIVE A OUVRIR
'=======================================================

Private Sub Form_Load()
'prépare le FileView
    With FV
        .Path = Left$(App.Path, 3)  'affiche un drive existant (nécessaire à AddDrives)
        .ShowFiles = False
        .ShowDirectories = False
        .ShowDrives = True 'affiche QUE les drives
        .AllowDirectoryEntering = False 'ne RENTRE PAS dans les dossiers (drives)
        .View = lvwList 'pas de columns
        .AddDrives    'affiche les drives
    End With
End Sub


'=======================================================
'ouvre diskFrm
'=======================================================
Private Sub cmdOk_Click()
Dim Frm As Form
Dim sDrive As String
Dim cDr As clsDiskInfos

    'vérifié que le drive est accessible
    Set cDr = New clsDiskInfos
    If cDr.IsLogicalDriveAccessible(FV.ListItems.Item(FV.ListIndex).Text) = False Then
        Set cDr = Nothing   'inaccessible, alors on sort de cette procédure
        Exit Sub
    End If
    Set cDr = Nothing
    
    'affiche une nouvelle fenêtre
    Set Frm = New diskPfm
    
    Call Frm.GetDrive(FV.ListItems.Item(FV.ListIndex).Text) 'renseigne sur le path sélectionné
    
    Unload Me   'quitte cette form
    
    Frm.Show    'affiche la nouvelle
    lNbChildFrm = lNbChildFrm + 1
    frmContent.Sb.Panels(2).Text = "Ouvertures=[" & CStr(lNbChildFrm) & "]"
 
End Sub

Private Sub cmdNo_Click()
    Unload Me
End Sub

Private Sub FV_ItemDblSelection(Item As ComctlLib.ListItem)
    cmdOk_Click
End Sub
