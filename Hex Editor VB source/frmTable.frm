VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmTable 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Table"
   ClientHeight    =   4125
   ClientLeft      =   -72960
   ClientTop       =   360
   ClientWidth     =   4050
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   4050
   Visible         =   0   'False
   Begin ComctlLib.ListView LV 
      Height          =   4095
      Left            =   38
      TabIndex        =   0
      Top             =   15
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   7223
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "D�cimal"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Binaire"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Octal"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Hexad�cimal"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "ANSI ASCII"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Menu mnuVisible 
      Caption         =   "mnuVisible"
      Visible         =   0   'False
      Begin VB.Menu mnuZOrder 
         Caption         =   "&Affichage au premier plan"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' -----------------------------------------------
'
' Hex Editor VB
' Coded by violent_ken (Alain Descotes)
'
' -----------------------------------------------
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
' -----------------------------------------------


Option Explicit

'-------------------------------------------------------
'FORM AFFICHANT UNE TABLE DE CONVERSION
'-------------------------------------------------------

'-------------------------------------------------------
'cr�� une table de conversion
'param�tre : tType As TableType
'peut afficher une table Hex<-->ASCII
'ou bien afficher un listview avec toutes les
'valeurs en base octale, d�cimale, hexa, ascii et binaire
'-------------------------------------------------------
Public Sub CreateTable(ByVal tType As TableType)
Dim x As Long
Dim y As Long
    
    If Not (tType = AllTables) Then
        'pr�pare l'affichage des r�gles en ordonn�e / ascisse
        Me.Width = 4140
        Me.Height = 4500
        LV.Width = 4100
        LV.Height = 4500
        LV.Left = 20
        LV.Visible = False
        LV.Top = 25
        
        Me.CurrentX = 0
        Me.CurrentY = 0
        Me.ForeColor = 16737380
        Me.BackColor = vbWhite
        For x = 0 To 15
            Me.CurrentY = 240 + 240 * x
            Me.CurrentX = 0
            Me.Print Hex$(x) & "0"
            Me.CurrentX = 360 + 230 * x
            Me.CurrentY = 0
            Me.Print Hex$(x)
        Next x
        
        'affichage des valeurs
        Me.ForeColor = vbBlack
        For x = 0 To 15
            For y = 0 To 15
                Me.CurrentX = 360 + 230 * x
                Me.CurrentY = 240 + 240 * y
                Me.Print Chr$(Val(16 * y + x))
            Next y
        Next x
        
    Else
        'alors on remplit et affiche le listview
        
        Me.Cls
        Me.Width = 7050
        Me.Height = 7000
        LV.Width = 6900
        LV.Height = 6600
        LV.Left = 38
        LV.Top = 15
        
        With LV
            For x = 1 To 256
                .ListItems.Add Text:=IIf(Len(CStr(x - 1)) = 1, "00" & CStr(x - 1), _
                IIf(Len(CStr(x - 1)) = 2, "0" & CStr(x - 1), CStr(x - 1)))  'd�cimal
                .ListItems.Item(x).SubItems(1) = Dec2Bin(x - 1, 8) 'binaire
                .ListItems.Item(x).SubItems(2) = IIf(Len(Oct$(x - 1)) = 1, _
                "00" & Oct$(x - 1), IIf(Len(Oct$(x - 1)) = 2, "0" & Oct$(x - 1), _
                Oct$(x - 1))) 'octal
                .ListItems.Item(x).SubItems(3) = IIf(Len(Hex$(x - 1)) = 1, "0" & Hex$(x - 1), _
                Hex$(x - 1)) 'hexa
                .ListItems.Item(x).SubItems(4) = Chr$(x - 1) 'ANSI ASCII
            Next x
        End With
        LV.Visible = True
    End If
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2

End Sub

Private Sub Form_Activate()
'mise au premier plan
    PremierPlan Me, MettreAuPremierPlan
    Me.Visible = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'affichage du popu menu
    Me.SetFocus
    If Button = 2 Then Me.PopupMenu Me.mnuVisible
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'affichage du popu menu
    Me.SetFocus
    If Button = 2 Then Me.PopupMenu Me.mnuVisible
End Sub

Private Sub mnuZOrder_Click()
    Me.mnuZOrder.Checked = Not (Me.mnuZOrder.Checked)
    If Me.mnuZOrder.Checked Then
        'affichage au premier plan
        PremierPlan Me, MettreAuPremierPlan
    Else
        'pas au premier plan
        PremierPlan Me, MettreNormal
    End If
End Sub
