VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{6ADE9E73-F694-428F-BF86-06ADD29476A5}#1.0#0"; "ProgressBar_OCX.ocx"
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recherche d'expressions"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8745
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   8745
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Type de recherche"
      Height          =   1095
      Index           =   5
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   3015
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         ScaleHeight     =   735
         ScaleWidth      =   1575
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
         Begin VB.OptionButton Option4 
            Caption         =   "Valeur ASCII"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   4
            Tag             =   "pref1"
            ToolTipText     =   "Rechercher une valeur ASCII"
            Top             =   480
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Valeur hexa"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   3
            Tag             =   "pref0"
            ToolTipText     =   "Recherche une valeur hexa"
            Top             =   120
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options de recherche"
      Height          =   2055
      Index           =   2
      Left            =   120
      TabIndex        =   25
      Top             =   4680
      Width           =   3015
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1695
         Index           =   4
         Left            =   120
         ScaleHeight     =   1695
         ScaleWidth      =   2775
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   240
         Width           =   2775
         Begin VB.CheckBox Check3 
            Caption         =   "Mot entier"
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            TabIndex        =   14
            Tag             =   "pref"
            ToolTipText     =   "Rechercher un mot entier"
            Top             =   840
            Width           =   2295
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Ajouter des signets"
            Height          =   255
            Left            =   0
            TabIndex        =   13
            Tag             =   "pref"
            ToolTipText     =   "Ajoute un signet pour chaque résultat trouvé"
            Top             =   480
            Width           =   2295
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Partir vers le haut"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   16
            Tag             =   "pref1"
            ToolTipText     =   "Commencer la recherche par en haut"
            Top             =   1460
            Width           =   1935
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Partir vers le bas"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   15
            Tag             =   "pref0"
            ToolTipText     =   "Commencer la recherche par en bas"
            Top             =   1160
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Respecter la casse"
            Height          =   255
            Left            =   0
            TabIndex        =   12
            Tag             =   "pref"
            ToolTipText     =   "Le respect de la casse est aussi valable pour des valeurs hexa"
            Top             =   120
            Width           =   2295
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Résultats"
      Height          =   4935
      Index           =   4
      Left            =   3240
      TabIndex        =   20
      Top             =   1800
      Width           =   5415
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   4560
         Left            =   30
         ScaleHeight     =   4560
         ScaleWidth      =   5325
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   260
         Width           =   5330
         Begin ComctlLib.ListView LV 
            Height          =   4575
            Left            =   45
            TabIndex        =   31
            Top             =   0
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   8070
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   0
            NumItems        =   1
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Offset"
               Object.Width           =   9128
            EndProperty
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rechercher"
      Height          =   1575
      Index           =   3
      Left            =   3240
      TabIndex        =   19
      Top             =   120
      Width           =   5415
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1215
         Index           =   2
         Left            =   120
         ScaleHeight     =   1215
         ScaleWidth      =   5175
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   240
         Width           =   5175
         Begin VB.CommandButton cmdQuit 
            Caption         =   "Fermer"
            Height          =   375
            Left            =   3720
            TabIndex        =   2
            ToolTipText     =   "Fermer cette fenêtre"
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Rechercher"
            Height          =   375
            Left            =   3720
            TabIndex        =   1
            ToolTipText     =   "lancer la recherche"
            Top             =   0
            Width           =   1335
         End
         Begin VB.TextBox txtSearch 
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   120
            TabIndex        =   0
            ToolTipText     =   $"frmSearch.frx":058A
            Top             =   240
            Width           =   3255
         End
         Begin ProgressBar_OCX.pgrBar PGB 
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   960
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   450
            BackColorTop    =   13027014
            BackColorBottom =   15724527
            Min             =   0
            Value           =   0
            BackPicture     =   "frmSearch.frx":062D
            FrontPicture    =   "frmSearch.frx":0649
         End
         Begin VB.Label Label2 
            Caption         =   "Expression à rechercher :"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   0
            Width           =   1935
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Zone de recherche"
      Height          =   1935
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   2640
      Width           =   3015
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1575
         Index           =   1
         Left            =   120
         ScaleHeight     =   1575
         ScaleWidth      =   2775
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   2775
         Begin VB.TextBox txtTo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1560
            TabIndex        =   9
            Tag             =   "pref"
            ToolTipText     =   "Offset supérieur"
            Top             =   400
            Width           =   1095
         End
         Begin VB.TextBox txtFrom 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Tag             =   "pref"
            ToolTipText     =   "Offset inférieur"
            Top             =   400
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Tout"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   11
            Tag             =   "pref"
            ToolTipText     =   "Rechercher de partout"
            Top             =   1200
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Sélection"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   10
            Tag             =   "pref1"
            ToolTipText     =   "Ne recherche que dans la sélection"
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Offset"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Tag             =   "pref0"
            ToolTipText     =   "Sélectionner une place d'offsets"
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "à"
            Height          =   255
            Left            =   1320
            TabIndex        =   24
            Top             =   400
            Width           =   135
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Type de recherche"
      Height          =   1215
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   1320
      Width           =   3015
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   855
         Index           =   0
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   2775
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   2775
         Begin VB.OptionButton Option1 
            Caption         =   "Expression regulière"
            Enabled         =   0   'False
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   5
            Tag             =   "pref3"
            ToolTipText     =   "Effectuer une recherche à l'aide d'une expression régulière"
            Top             =   120
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Expression simple"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   6
            Tag             =   "pref2"
            ToolTipText     =   "Effectuer une recherche simple"
            Top             =   480
            Value           =   -1  'True
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "frmSearch"
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
'FORM DE RECHERCHE
'=======================================================

Private clsPref As clsIniForm

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
'lance la recherche
Dim tRes() As Long
Dim X As Long
Dim s As String

    If txtSearch.Text = vbNullString Then Exit Sub
    
    LV.ListItems.Clear

    Select Case TypeOfForm(frmContent.ActiveForm)
        Case "Fichier"
            'recherche dans un fichier
            If Option4(1).Value And Option1(2).Value Then
                'alors c'est une string en recherche simple
                SearchForStringFile frmContent.ActiveForm.Caption, txtSearch.Text, Check1.Value, _
                tRes(), Me.pgb
            ElseIf Option1(2).Value Then
                'alors c'est une valeur hexa en recherche simple ==> on convertit d'abord en string
                s = HexValues2String(txtSearch.Text)
                
                'lance la recherche de la string
                SearchForStringFile frmContent.ActiveForm.Caption, s, Check1.Value, _
                tRes(), Me.pgb
            Else
                Exit Sub
            End If
            
        Case "Disque"
            'recherche dans un disque
            If Option4(1).Value And Option1(2).Value Then
                'alors c'est une string en recherche simple
                SearchForStringDisk frmContent.ActiveForm.Caption, txtSearch.Text, Check1.Value, _
                tRes(), Me.pgb
            ElseIf Option1(2).Value Then
                'alors c'est une valeur hexa en recherche simple ==> on convertit d'abord en string
                s = HexValues2String(txtSearch.Text)
                
                'lance la recherche de la string
                SearchForStringDisk frmContent.ActiveForm.Caption, s, Check1.Value, _
                tRes(), Me.pgb
            Else
                Exit Sub
            End If
            
            
        Case "Processus"
            'recherche dans la mémoire
            If Option4(1).Value And Option1(2).Value Then
                'alors c'est une string en recherche simple
                ' ==> le PID avait été stocké dans le Tag

                'lance la recherche
                cMem.SearchForStringMemory CLng(frmContent.ActiveForm.Tag), txtSearch.Text, Check1.Value, _
                tRes(), Me.pgb
            ElseIf Option1(2).Value Then
                'alors c'est une valeur hexa en recherche simple ==> on convertit d'abord en string
                s = HexValues2String(txtSearch.Text)
                
                'lance la recherche de la string
                cMem.SearchForStringMemory CLng(frmContent.ActiveForm.Tag), s, Check1.Value, _
                tRes(), Me.pgb
            Else
                Exit Sub
            End If
            
    End Select
    

    '//affiche les résultats
    For X = 1 To UBound(tRes())
        LV.ListItems.Add Text:="Trouvé à l'offset " & CStr(By16D(tRes(X)))
        If Check2.Value Then
            frmContent.ActiveForm.HW.AddSignet By16D(tRes(X))
            frmContent.ActiveForm.lstSignets.ListItems.Add Text:=Trim$(Str$(By16D(tRes(X))))
            frmContent.ActiveForm.lstSignets.ListItems.Item(frmContent.ActiveForm.lstSignets.ListItems.Count).SubItems(1) = "Found [" & Trim$(txtSearch.Text) & "]"
        End If
    Next X
    
    Frame1(4).Caption = "Résultats : " & CStr(UBound(tRes()))
        
End Sub

Private Sub Form_Load()
    'loading des preferences
    Set clsPref = New clsIniForm
    clsPref.GetFormSettings App.Path & "\Preferences\Search.ini", Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'sauvegarde des preferences
    clsPref.SaveFormSettings App.Path & "\Preferences\Search.ini", Me
    Set clsPref = Nothing
End Sub

Private Sub LV_ItemClick(ByVal Item As ComctlLib.ListItem)
'va dans le HW correspondant

    On Error Resume Next    'si jamais plus de ActiveForm ou je ne sais quoi d'autre...
    
    frmContent.ActiveForm.VS.Value = Val(Right$(Item.Text, Len(Item.Text) - 18)) / 16
    Call frmContent.ActiveForm.VS_Change(frmContent.ActiveForm.VS.Value)
    
End Sub
