VERSION 5.00
Begin VB.Form frmSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sélectionner une zone"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   2745
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
   Icon            =   "frmSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFrom 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1545
      TabIndex        =   3
      ToolTipText     =   "Offset inférieur"
      Top             =   105
      Width           =   1095
   End
   Begin VB.TextBox txtTo 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1545
      TabIndex        =   2
      ToolTipText     =   "Offset supérieur"
      Top             =   465
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Sélectionner"
      Height          =   375
      Left            =   150
      TabIndex        =   1
      ToolTipText     =   "Procéder à la restriction"
      Top             =   945
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   1470
      TabIndex        =   0
      ToolTipText     =   "Fermer cette fenêtre"
      Top             =   945
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "A partir du byte"
      Height          =   255
      Index           =   0
      Left            =   105
      TabIndex        =   5
      Top             =   105
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "jusqu'au byte"
      Height          =   255
      Index           =   1
      Left            =   105
      TabIndex        =   4
      Top             =   465
      Width           =   1215
   End
End
Attribute VB_Name = "frmSelect"
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
'FORM POUR SELECTIONNER UNE ZONE PARTICULIERE
'=======================================================

Private Lang As New clsLang
Private byteFunc As Byte

Private Sub cmdOk_Click()
'valide
Dim lFrom As Currency
Dim lTo As Currency
Dim X As Currency

    'On Error GoTo ErrGestion
    
    'récupère les valeurs numériques
    lFrom = FormatedVal_(txtFrom.Text) + 1
    lTo = FormatedVal_(txtTo.Text)
    
    'fait en sorte que lFrom soit le plus petit
    If lFrom > lTo Then
        X = lFrom
        lFrom = lTo
        lTo = X
    End If
        
    If byteFunc = 0 Then    'il s'agit d'une sélection paramétrée

        'vérifie que la plage est OK
        If lFrom < frmContent.ActiveForm.HW.FirstOffset Or lTo > _
            frmContent.ActiveForm.HW.MaxOffset Then
            Unload Me
            Exit Sub
        End If
        If lFrom > frmContent.ActiveForm.HW.MaxOffset Or lTo < _
            frmContent.ActiveForm.HW.FirstOffset Then
            Unload Me
            Exit Sub
        End If
        
        'fait la sélection désirée
        With frmContent.ActiveForm
            .HW.SelectZone 16 - (By16(lFrom) - lFrom), _
                By16(lFrom) - 16, 17 - (By16(lTo) - lTo), By16(lTo) - 16
            
            'refresh le label qui contient la taille de la sélection
            .Sb.Panels(4).Text = Lang.GetString("_Sel") & _
                CStr(.HW.NumberOfSelectedItems) & " bytes]"
            .Label2(9) = .Sb.Panels(4).Text
        End With
        
    ElseIf byteFunc = 1 Then 'il s'agit d'une restriction d'affichage
    
        'formate les valeurs en terme d'offset
        lFrom = By16(lFrom)
        
        'vérifie que la plage est OK
        '/!\ vérifie la plage à l'aide des TAGS CURRENCY du HW (valeurs maximales autorisées)
        If lFrom < frmContent.ActiveForm.HW.curTag1 Or lTo > _
            frmContent.ActiveForm.HW.curTag2 Then
            Unload Me
            Exit Sub
        End If
        If lFrom > frmContent.ActiveForm.HW.curTag2 Or lTo < _
            frmContent.ActiveForm.HW.curTag1 Then
            Unload Me
            Exit Sub
        End If
        
        With frmContent.ActiveForm
            'ajoute une entrée à l'historique
            .AddHistoFrm actRestArea, , , .HW.FirstOffset, .HW.MaxOffset
            
            'change les valeurs dans la ActiveForm
            .HW.FirstOffset = lFrom
            .HW.MaxOffset = lTo
            .VS.Min = lFrom / 16
            .VS.Max = By16(lTo / 16)
            Call .VS_Change(.VS.Min)
        End With

    End If
    
    Call frmContent.ActiveForm.HW.Refresh
    
    Unload Me
    
    Exit Sub
ErrGestion:
    clsERREUR.AddError "frmSelect.cmdOkClick", True
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
    
    If frmContent.ActiveForm Is Nothing Then Unload Me
    
    'affiche l'élément actuellement sélectionné dans l'activeform
    txtFrom.Text = CStr(frmContent.ActiveForm.HW.Item.Offset + frmContent.ActiveForm.HW.Item.Col) - 1
End Sub

'=======================================================
'sub permettant de récupérer un nombre qui va spécifier
'a quoi la sélection servira
'0 = sélection paramétrée
'1 = affichage restreint
'=======================================================
Public Sub GetEditFunction(ByVal btFunction As Byte)
    
    'active la gestion des langues
    Call Lang.ActiveLang(Me)
    
    byteFunc = btFunction
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Lang = Nothing
End Sub
