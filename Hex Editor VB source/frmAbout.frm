VERSION 5.00
Object = "{BEF0F0EF-04C8-45BD-A6A9-68C01A66CB51}#1.1#0"; "vkUserControlsXP.ocx"
Begin VB.Form frmAbout 
   BackColor       =   &H00760401&
   BorderStyle     =   0  'None
   Caption         =   "About"
   ClientHeight    =   6390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7065
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmAbout.frx":000C
   ScaleHeight     =   6390
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin vkUserContolsXP.vkCommand cmdUnload 
      Height          =   375
      Left            =   5640
      TabIndex        =   9
      ToolTipText     =   "Fermer cette fenêtre"
      Top             =   5880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
   Begin VB.TextBox txt 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1660
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4040
      Width           =   6705
   End
   Begin vkUserContolsXP.vkCommand cmdLicense 
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Tag             =   "Afficher les informations de licence"
      Top             =   5880
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Caption         =   "Informations de licence"
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
   Begin VB.Label lblVersionWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "Pre Alpha version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   105
      TabIndex        =   7
      Top             =   3960
      Width           =   6855
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) 2006-2007 Alain Descotes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3937
      TabIndex        =   6
      Top             =   2872
      Width           =   3015
   End
   Begin VB.Label lblWARNING 
      BackStyle       =   0  'Transparent
      Caption         =   "Avertissement : ce logiciel est protégé par la license GNU General Public License"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   3585
      Width           =   6855
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5977
      TabIndex        =   4
      Top             =   2512
      Width           =   795
   End
   Begin VB.Label lblPlatform 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designed for Windows"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3307
      TabIndex        =   3
      Top             =   2152
      Width           =   3525
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hex Editor VB"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   2257
      TabIndex        =   2
      Top             =   1072
      Width           =   4500
   End
   Begin VB.Label lblLicenseTo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "License accordée à [NAME]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   90
      TabIndex        =   1
      Top             =   165
      Width           =   6855
   End
   Begin VB.Image Image1 
      Height          =   2265
      Left            =   127
      Picture         =   "frmAbout.frx":5AAEA
      Stretch         =   -1  'True
      Top             =   982
      Width           =   1815
   End
End
Attribute VB_Name = "frmAbout"
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

Private Lang As New clsLang

Private Sub cmdLicense_Click()
'affiche le ReadMe

    If cFile.FileExists(App.Path & "\License.txt") = False Then Exit Sub
    Call cFile.ShellOpenFile(App.Path & "\License.txt", Me.hWnd)
End Sub

Private Sub cmdUnload_Click()
    Call Unload(Me)
End Sub

Private Sub Form_Load()
Dim s As String

    With Lang
        #If MODE_DEBUG Then
            If App.LogMode = 0 And CREATE_FRENCH_FILE Then
                'on créé le fichier de langue français
                .Language = "French"
                .LangFolder = LANG_PATH
                Call .WriteIniFileFormIDEform
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
        Call .LoadControlsCaption
        
        'mise à jour de la version et de l'USER
        lblLicenseTo.Caption = .GetString("_LicenseTo") & " " & GetUserName
        lblVersion.Caption = .GetString("_Version") & " " & _
            Trim$(Str$(App.Major)) & "." & Trim$(Str$(App.Minor)) & "." & _
            Trim$(Str$(App.Revision))
    End With
    
    'écriture du texte
    s = "Hex Editor VB" & vbNewLine & _
        "Copyright (c) 2006-2007 Alain Descotes (violent_ken)"
    s = s & vbNewLine & vbNewLine & _
        cFile.LoadFileInString(App.Path & "\License.txt")
    txt.Text = s
    
End Sub
