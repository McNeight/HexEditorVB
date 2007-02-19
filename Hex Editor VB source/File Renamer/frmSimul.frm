VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmSimul 
   Caption         =   "Simulation"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8460
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSimul.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   5550
   ScaleWidth      =   8460
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ListView LV 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8493
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Ancien nom"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Nouveau nom"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Path"
         Object.Width           =   12347
      EndProperty
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   3443
      TabIndex        =   1
      Top             =   5040
      Width           =   1575
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   2640
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSimul.frx":08CA
            Key             =   "Success"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSimul.frx":0C1C
            Key             =   "Failed"
         EndProperty
      EndProperty
   End
   Begin VB.Image oui 
      Height          =   240
      Left            =   960
      Picture         =   "frmSimul.frx":0F6E
      Top             =   5160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image non 
      Height          =   240
      Left            =   1440
      Picture         =   "frmSimul.frx":12B0
      Top             =   5160
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmSimul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =======================================================
'
' File Renamer VB (part of Hex Editor VB)
' Coded by violent_ken (Alain Descotes)
'
' =======================================================
'
' An Windows utility which allows to rename lots of file (part of Hex Editor VB)
'
' Copyright (c) 2006-2007 by Alain Descotes.
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

Private Sub cmdOk_Click()
'ok
    Unload Me
End Sub

Private Sub Form_Load()
'lance la simulation
Dim sOld() As String
Dim sNew() As String
Dim x As Long

    ReDim sOld(frmMain.FileR.ListCount)
    'remplit sOld (anciens noms)
    For x = 1 To frmMain.FileR.ListCount
        sOld(x) = frmMain.FileR.ListItems(x).Text
    Next x
    
    'procède au calcul des nouveaux noms
    RenameMyFiles frmMain.lstTransfo, sOld(), sNew()
    
    'affiche les résultats
    LV.Visible = False
    LV.ListItems.Clear
    For x = 1 To UBound(sOld())
        LV.ListItems.Add Text:=sOld(x), SmallIcon:=IIf(sOld(x) = sNew(x), "Failed", "Success")
        LV.ListItems.Item(x).SubItems(1) = sNew(x)
        LV.ListItems.Item(x).SubItems(2) = cFile.GetFolderFromPath(frmMain.FileR.ListItems.Item(x).Tag)  'le path
    Next x
    LV.Visible = True
End Sub
