VERSION 5.00
Object = "{BEF0F0EF-04C8-45BD-A6A9-68C01A66CB51}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recherche d'expressions"
   ClientHeight    =   7065
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
   HelpContextID   =   20
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   8745
   StartUpPosition =   2  'CenterScreen
   Begin vkUserContolsXP.vkFrame vkFrame6 
      Height          =   1695
      Left            =   3240
      TabIndex        =   21
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2990
      Caption         =   "Rechercher"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin vkUserContolsXP.vkBar PGB 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   450
         Value           =   1
         BackPicture     =   "frmSearch.frx":058A
         FrontPicture    =   "frmSearch.frx":05A6
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
      Begin VB.TextBox txtSearch 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   $"frmSearch.frx":05C2
         Top             =   600
         Width           =   3375
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Rechercher"
         Height          =   375
         Left            =   3840
         TabIndex        =   23
         ToolTipText     =   "lancer la recherche"
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Fermer"
         Height          =   375
         Left            =   3840
         TabIndex        =   22
         ToolTipText     =   "Fermer cette fenêtre"
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Expression à rechercher :"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   1935
      End
   End
   Begin vkUserContolsXP.vkFrame grdFrame1 
      Height          =   5055
      Left            =   3240
      TabIndex        =   20
      Top             =   1920
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   8916
      Caption         =   "Résultats - offsets"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin vkUserContolsXP.vkListBox Listbox 
         Height          =   4575
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   8070
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MultiSelect     =   0   'False
         Sorted          =   0
      End
   End
   Begin vkUserContolsXP.vkFrame vkFrame4 
      Height          =   2295
      Left            =   120
      TabIndex        =   14
      Top             =   4680
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4048
      Caption         =   "Options de recherche"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin vkUserContolsXP.vkOptionButton Option3 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   19
         ToolTipText     =   "Commencer la recherche par en haut"
         Top             =   1920
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Partir vers le haut"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Group           =   4
      End
      Begin vkUserContolsXP.vkOptionButton Option3 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   18
         ToolTipText     =   "Commencer la recherche par en bas"
         Top             =   1680
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Partir vers le bas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Group           =   4
      End
      Begin vkUserContolsXP.vkCheck Check1 
         Height          =   255
         Left            =   240
         TabIndex        =   17
         ToolTipText     =   "Le respect de la casse est aussi valable pour des valeurs hexa"
         Top             =   1200
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Respecter la casse"
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
      Begin vkUserContolsXP.vkCheck Check3 
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Mot entier"
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
      Begin vkUserContolsXP.vkCheck Check2 
         Height          =   255
         Left            =   240
         TabIndex        =   15
         ToolTipText     =   "Ajoute un signet pour chaque résultat trouvé"
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Ajouter des signets"
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
   End
   Begin vkUserContolsXP.vkFrame vkFrame3 
      Height          =   2055
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   3625
      Caption         =   "Zone de recherche"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin vkUserContolsXP.vkOptionButton Option2 
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Rechercher de partout"
         Top             =   1560
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Tout"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Group           =   3
      End
      Begin vkUserContolsXP.vkOptionButton Option2 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Ne recherche que dans la sélection"
         Top             =   1200
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Sélection"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Group           =   3
      End
      Begin vkUserContolsXP.vkOptionButton Option2 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Sélectionner une place d'offsets"
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Offset"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Group           =   3
      End
      Begin VB.TextBox txtFrom 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Tag             =   "pref"
         ToolTipText     =   "Offset inférieur"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtTo 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Tag             =   "pref"
         ToolTipText     =   "Offset supérieur"
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "à"
         Height          =   255
         Left            =   1320
         TabIndex        =   10
         Top             =   840
         Width           =   135
      End
   End
   Begin vkUserContolsXP.vkFrame vkFrame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1931
      Caption         =   "Type de recherche"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin vkUserContolsXP.vkOptionButton Option1 
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   5
         ToolTipText     =   "Effectuer une recherche simple"
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Expression simple"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   1
         Group           =   2
      End
      Begin vkUserContolsXP.vkOptionButton Option1 
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "Effectuer une recherche à l'aide d'une expression régulière"
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Expression regulière"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Group           =   2
      End
   End
   Begin vkUserContolsXP.vkFrame vkFrame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1931
      Caption         =   "Type de recherche"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin vkUserContolsXP.vkOptionButton Option4 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         ToolTipText     =   "Rechercher une valeur ASCII"
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Valeur ASCII"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   1
         Group           =   1
      End
      Begin vkUserContolsXP.vkOptionButton Option4 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         ToolTipText     =   "Recherche une valeur hexa"
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Valeur hexa"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Group           =   1
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

Private Lang As New clsLang
Private clsPref As clsIniForm

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
'lance la recherche
Dim tRes() As Currency
Dim x As Long
Dim s As String

    If txtSearch.Text = vbNullString Then Exit Sub
    
    Call Listbox.Clear
    txtSearch.Enabled = False

    Select Case TypeOfForm(frmContent.ActiveForm)
        Case "Fichier"
            'recherche dans un fichier
            If Option4(1).Value And Option1(2).Value Then
                'alors c'est une string en recherche simple
                
                'ajoute du texte à la console
                Call AddTextToConsole(Lang.GetString("_SearchCour"))
    
                Call SearchForStringFile(frmContent.ActiveForm.Caption, txtSearch.Text, _
                    Check1.Value, tRes(), Me.PGB)
                    
            ElseIf Option1(2).Value Then
                'alors c'est une valeur hexa en recherche simple ==> on convertit d'abord en string
                s = HexValues2String(txtSearch.Text)
                
                'ajoute du texte à la console
                Call AddTextToConsole(Lang.GetString("_SearchCour"))
                    
                'lance la recherche de la string
                Call SearchForStringFile(frmContent.ActiveForm.Caption, s, Check1.Value, _
                    tRes(), Me.PGB)
            Else
                Exit Sub
            End If
            
        Case "Disque"
            'recherche dans un disque
            If Option4(1).Value And Option1(2).Value Then
                'alors c'est une string en recherche simple
                
                'ajoute du texte à la console
                Call AddTextToConsole(Lang.GetString("_SearchCour"))
                
                Call SearchForStringDisk(frmContent.ActiveForm.Caption, txtSearch.Text, _
                    Check1.Value, tRes(), Me.PGB)
                
            ElseIf Option1(2).Value Then
                'alors c'est une valeur hexa en recherche simple ==> on convertit d'abord en string
                s = HexValues2String(txtSearch.Text)
                
                'ajoute du texte à la console
                Call AddTextToConsole(Lang.GetString("_SearchCour"))
                
                'lance la recherche de la string
                Call SearchForStringDisk(frmContent.ActiveForm.Caption, s, Check1.Value, _
                    tRes(), Me.PGB)
            Else
                Exit Sub
            End If
            
            
            
        Case "Disque physique"
            'recherche dans un disque PHYSIQUE
            If Option4(1).Value And Option1(2).Value Then
                'alors c'est une string en recherche simple
                
                'ajoute du texte à la console
                Call AddTextToConsole(Lang.GetString("_SearchCour"))
                
                Call SearchForStringDisk(frmContent.ActiveForm.Caption, _
                    txtSearch.Text, Check1.Value, tRes(), Me.PGB, True)
                
            ElseIf Option1(2).Value Then
                'alors c'est une valeur hexa en recherche simple ==> on convertit d'abord en string
                s = HexValues2String(txtSearch.Text)
                
                'ajoute du texte à la console
                Call AddTextToConsole(Lang.GetString("_SearchCour"))
                
                'lance la recherche de la string
                Call SearchForStringDisk(frmContent.ActiveForm.Caption, s, _
                    Check1.Value, tRes(), Me.PGB, True)
            Else
                Exit Sub
            End If
            
            
        Case "Processus"
            'recherche dans la mémoire
            If Option4(1).Value And Option1(2).Value Then
                'alors c'est une string en recherche simple
                ' ==> le PID avait été stocké dans le Tag

                'ajoute du texte à la console
                Call AddTextToConsole(Lang.GetString("_SearchCour"))
                
                'lance la recherche
                Call cMem.SearchForStringMemory(CLng(frmContent.ActiveForm.Tag), _
                    txtSearch.Text, Check1.Value, tRes(), Me.PGB)
            ElseIf Option1(2).Value Then
                'alors c'est une valeur hexa en recherche simple ==> on convertit d'abord en string
                s = HexValues2String(txtSearch.Text)

                'ajoute du texte à la console
                Call AddTextToConsole(Lang.GetString("_SearchCour"))
                
                'lance la recherche de la string
                Call cMem.SearchForStringMemory(CLng(frmContent.ActiveForm.Tag), _
                    s, Check1.Value, tRes(), Me.PGB)
            Else
                Exit Sub
            End If
            
    End Select
    
    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_DisplayRes"))

    '//affiche les résultats
    Listbox.UnRefreshControl = True
    For x = 1 To UBound(tRes())
        Call Listbox.AddItem(Caption:=Lang.GetString("_FoundAt") & " " & CStr(By16D(tRes(x))))
        If Check2.Value Then
            With frmContent.ActiveForm
                .HW.AddSignet By16D(tRes(x))
                .lstSignets.ListItems.Add Text:=Trim$(Str$(By16D(tRes(x))))
                .lstSignets.ListItems.Item(.lstSignets.ListItems.Count).SubItems(1) = "Found [" & Trim$(txtSearch.Text) & "]"
            End With
        End If
    Next x
    Listbox.UnRefreshControl = False
    Call Listbox.Refresh
    
    grdFrame1.Caption = Lang.GetString("_ResAre") & " " & CStr(UBound(tRes()))
    
    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_SearchComplete"))
    
    txtSearch.Enabled = True
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
    Call clsPref.GetFormSettings(App.Path & "\Preferences\Search.ini", Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'sauvegarde des preferences
    Call clsPref.SaveFormSettings(App.Path & "\Preferences\Search.ini", Me)
    Set clsPref = Nothing
End Sub

Private Sub Listbox_ItemClick(Item As vkUserContolsXP.vkListItem)
'va dans le HW correspondant

    On Error Resume Next    'si jamais plus de ActiveForm ou je ne sais quoi d'autre...
    
    frmContent.ActiveForm.VS.Value = Val(Right$(Item.Text, Len(Item.Text) - 18)) / 16
    Call frmContent.ActiveForm.VS_Change(frmContent.ActiveForm.VS.Value)

End Sub
