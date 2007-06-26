VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BEF0F0EF-04C8-45BD-A6A9-68C01A66CB51}#1.1#0"; "vkUserControlsXP.ocx"
Begin VB.Form frmComponents 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Composants de Hex Editor VB"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6750
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmComponents.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pctIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6480
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   240
   End
   Begin vkUserContolsXP.vkCommand cmdOK 
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   2520
      Width           =   2055
      _ExtentX        =   3625
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
   Begin vkUserContolsXP.vkTextBox txtVersion 
      Height          =   2175
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   3836
      BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      LegendText      =   "Informations sur le fichier"
      LegendForeColor =   12937777
      LegendType      =   1
   End
   Begin vkUserContolsXP.vkListBox List 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   4895
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
      Path            =   ""
   End
   Begin ComctlLib.ImageList IMG 
      Left            =   2160
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
End
Attribute VB_Name = "frmComponents"
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
'FORM DE GESTION DES VERSIONS DES COMPOSANTS
'=======================================================

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim sPath As String
    
    'on ajoute à la liste tous les composants utilisés par Hex Editor VB
    
    'détermine le Path
    #If FINAL_VERSION Then
        sPath = App.Path & "\"
    #Else
        sPath = EXE_PATH
    #End If
    
    'on créé un FileListBox avec un Pattern spécial
    List.Path = Left$(sPath, Len(sPath) - 1)
    List.Pattern = "*.ocx;*.exe;*.dll"
    List.ListType = FileListBox
    Call List.SortItems(Alphabetical)

End Sub

'=======================================================
'détermine si une clé existe deja ou pas dans le IMG
'=======================================================
Private Function DoesKeyExist(ByVal sKey As String) As Boolean
'renvoie si la clé existe ou non deja dans IMG
Dim l As Long

    DoesKeyExist = False
    
    On Error GoTo ErrGest
    
    l = IMG.ListImages(sKey).Index

    DoesKeyExist = True
    
    Exit Function
    
ErrGest:
'la clé n'existait pas
End Function

'=======================================================
'ajoute une icone au IMG, en fonction du fichier (obtient l'icone de l'executable)
'=======================================================
Private Function AddIconToIMG(ByVal sFile As String, ByVal sKey As String) As Boolean
Dim lstImg As ListImage
Dim hIcon As Long
Dim ShInfo As SHFILEINFO
Dim pct As IPictureDisp

    On Error GoTo ErrGestion
    
    AddIconToIMG = False
    
    If cFile.FileExists(sFile) = False Then Exit Function 'fichier introuvable
    If DoesKeyExist(sKey) Then Exit Function 'clé existe déjà
    
    'obtient le handle de l'icone
    hIcon = SHGetFileInfo(sFile, 0&, ShInfo, Len(ShInfo), BASIC_SHGFI_FLAGS Or _
        SHGFI_SMALLICON)
        
    'prépare la picturebox
    pctIcon.Picture = Nothing
    
    'trace l'image
    Call ImageList_Draw(hIcon, ShInfo.iIcon, pctIcon.hdc, 0, 0, ILD_TRANSPARENT)
    
    'ajout de l'icone à l'imagelist
    IMG.ListImages.Add Key:=sKey, Picture:=pctIcon.Image
    
    AddIconToIMG = True

    Exit Function
ErrGestion:
    clsERREUR.AddError "frmComponents.AddIconToImg", True
End Function

Private Sub List_ItemClick(Item As vkUserContolsXP.vkListItem)
Dim s As String
Dim cFic As FileSystemLibrary.File
Dim sPath As String

    If Item Is Nothing Then Exit Sub

    'détermine le Path
    #If FINAL_VERSION Then
        sPath = App.Path & "\"
    #Else
        sPath = EXE_PATH
    #End If

    'on affiche les infos sur le fichier
    Set cFic = cFile.GetFile(sPath & Item.Text) 'récupère les infos
    
    With cFic
        s = frmContent.Lang.GetString("_SizeIs") & CStr(.FileSize) & " Octets  -  " & CStr(Round(.FileSize / 1024, 3)) & " Ko" & "]"
        s = s & vbNewLine & frmContent.Lang.GetString("_AttrIs") & CStr(.Attributes) & "]"
        s = s & vbNewLine & frmContent.Lang.GetString("_CreaIs") & .DateCreated & "]"
        s = s & vbNewLine & frmContent.Lang.GetString("_AccessIs") & .DateLastAccessed & "]"
        s = s & vbNewLine & frmContent.Lang.GetString("_ModifIs") & .DateLastModified & "]"
        s = s & vbNewLine & frmContent.Lang.GetString("_Version") & .FileVersionInfos.FileVersion & "]"
        s = s & vbNewLine & frmContent.Lang.GetString("_DescrIs") & .FileVersionInfos.FileDescription & "]"
        s = s & vbNewLine & "Copyright=[" & .FileVersionInfos.Copyright & "]"
    End With
    
    txtVersion.Text = s
    
    Set cFic = Nothing
    
End Sub
