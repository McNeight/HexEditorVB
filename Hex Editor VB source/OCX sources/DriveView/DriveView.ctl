VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.UserControl DriveView 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox pctIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3960
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   240
   End
   Begin ComctlLib.TreeView TV 
      Height          =   2175
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   3836
      _Version        =   327682
      Indentation     =   1411
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "IMG"
      Appearance      =   0
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
   Begin ComctlLib.ImageList IMG 
      Left            =   3120
      Top             =   2400
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
            Picture         =   "DriveView.ctx":0000
            Key             =   "PhysicalDrive"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DriveView.ctx":0352
            Key             =   "inaccessible_drive"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "DriveView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
'//CONTROLE D'AFFICHAGE DES DRIVES ET DES DISQUES PHYSIQUES
'=======================================================


'=======================================================
'PRIVATE VARIABLES
'=======================================================
Private bDisplayPhysicalDrives As Boolean
Private bDisplayLogicalDrives As Boolean
Private cDrives As clsDiskInfos
Private bReadyToRefresh As Boolean
Private sLog As String
Private sPhys As String


'=======================================================
'CONSTANTES
'=======================================================
Private Const SHGFI_DISPLAYNAME             As Long = &H200
Private Const SHGFI_EXETYPE                 As Long = &H2000
Private Const SHGFI_SYSICONINDEX            As Long = &H4000
Private Const SHGFI_LARGEICON               As Long = &H0
Private Const SHGFI_SMALLICON               As Long = &H1
Private Const SHGFI_SHELLICONSIZE           As Long = &H4
Private Const SHGFI_TYPENAME                As Long = &H400
Private Const BASIC_SHGFI_FLAGS             As Long = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or _
                                            SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or _
                                            SHGFI_EXETYPE
Private Const ILD_TRANSPARENT               As Long = &H1


'=======================================================
'ENUMS & TYPES
'=======================================================
Private Type SHFILEINFO 'utilisé pour récupérer les icones des fichiers
    hIcon        As Long
    iIcon        As Long
    dwAttributes    As Long
    szDisplayName As String * 260
    szTypeName  As String * 80
End Type


'=======================================================
'EVENTS
'=======================================================
Public Event Click()
Public Event Collapse(ByVal Node As ComctlLib.Node)
Public Event DblClick()
Public Event Expand(ByVal Node As ComctlLib.Node)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event NodeClick(ByVal Node As ComctlLib.Node)


'=======================================================
'APIS
'=======================================================
'obtient des infos (utilisé pour l'icone de l'executable) d'un fichier
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
'créé une image à partir d'un handle d'icone
Private Declare Function ImageList_Draw Lib "Comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal Flags&) As Long
'autorise l'affichage d'un listview
Private Declare Sub InvalidateRect Lib "user32" (ByVal hwnd As Long, ByVal t As Long, ByVal bErase As Long)
'bloque l'affichage d'un listview
Private Declare Sub ValidateRect Lib "user32" (ByVal hwnd As Long, ByVal t As Long)




'=======================================================
'USERCONTROL SUBS
'=======================================================
Private Sub UserControl_InitProperties()
    'valeurs par défaut
    Me.Sorted = True
    Me.Style = tvwTreelinesPlusMinusPictureText
    Me.LineStyle = tvwRootLines
    Me.BorderStyle = 0
    Me.Appearance = ccFlat
    Me.Indentation = 300
    Me.DisplayLogicalDrives = True
    Me.DisplayPhysicalDrives = True
    Me.HideSelection = False
    Me.LogicalDrivesString = "Disques logiques"
    Me.PhysicalDrivesString = "Disques physiques"
    
    Call ShowD  'affiche les drives
End Sub
Private Sub UserControl_Initialize()
    bReadyToRefresh = False
    Set cDrives = New clsDiskInfos
End Sub
Private Sub UserControl_Show()
    bReadyToRefresh = True  'alors on peut désormais rafraichir
    Call ShowD
End Sub

Private Sub UserControl_Terminate()
    Set cDrives = Nothing
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Sorted", Me.Sorted, True)
        Call .WriteProperty("Style", Me.Style, tvwTreelinesPlusMinusPictureText)
        Call .WriteProperty("LineStyle", Me.LineStyle, tvwRootLines)
        Call .WriteProperty("BorderStyle", Me.BorderStyle, 0)
        Call .WriteProperty("Appearance", Me.Appearance, ccFlat)
        Call .WriteProperty("Indentation", Me.Indentation, 300)
        Call .WriteProperty("DisplayLogicalDrives", Me.DisplayLogicalDrives, True)
        Call .WriteProperty("DisplayPhysicalDrives", Me.DisplayPhysicalDrives, True)
        Call .WriteProperty("HideSelection", Me.HideSelection, False)
        Call .WriteProperty("PhysicalDrivesString", Me.PhysicalDrivesString, "Disques physiques")
        Call .WriteProperty("LogicalDrivesString", Me.LogicalDrivesString, "Disques logiques")
    End With
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Me.Sorted = .ReadProperty("Sorted", True)
        Me.Style = .ReadProperty("Style", tvwTreelinesPlusMinusPictureText)
        Me.LineStyle = .ReadProperty("LineStyle", tvwRootLines)
        Me.BorderStyle = .ReadProperty("BorderStyle", 0)
        Me.Appearance = .ReadProperty("Appearance", ccFlat)
        Me.Indentation = .ReadProperty("Indentation", 300)
        Me.DisplayLogicalDrives = .ReadProperty("DisplayLogicalDrives", True)
        Me.DisplayPhysicalDrives = .ReadProperty("DisplayPhysicalDrives", True)
        Me.HideSelection = .ReadProperty("HideSelection", False)
        Me.LogicalDrivesString = .ReadProperty("LogicalDrivesString", "Disques logiques")
        Me.PhysicalDrivesString = .ReadProperty("PhysicalDrivesString", "Disques physiques")
    End With
End Sub
Private Sub UserControl_Resize()
    With TV     'resize du TV
        .Height = UserControl.Height
        .Width = UserControl.Width
        .Top = 0
        .Left = 0
    End With
End Sub



'=======================================================
'PROPERTIES
'=======================================================
Public Property Get Sorted() As Boolean: Sorted = TV.Sorted: End Property
Public Property Let Sorted(Sorted As Boolean): TV.Sorted = Sorted: End Property
Public Property Get SelectedItem() As Node: Set SelectedItem = TV.SelectedItem: End Property
Public Property Let SelectedItem(SelectedItem As Node): Set TV.SelectedItem = SelectedItem: End Property
Public Property Get Nodes() As Nodes: Set Nodes = TV.Nodes: End Property
Public Property Let Nodes(Nodes As Nodes): Set TV.Nodes = Nodes: End Property
Public Property Get Object() As Object: Set Object = TV.Object: End Property
Public Property Get Style() As TreeStyleConstants: Style = TV.Style: End Property
Public Property Let Style(Style As TreeStyleConstants): TV.Style = Style: End Property
Public Property Get LineStyle() As TreeLineStyleConstants: LineStyle = TV.LineStyle: End Property
Public Property Let LineStyle(LineStyle As TreeLineStyleConstants): TV.LineStyle = LineStyle: End Property
Public Property Get Index() As Integer: Index = TV.Index: End Property
Public Property Get hwnd() As Long: hwnd = TV.hwnd: End Property
Public Property Get BorderStyle() As ComctlLib.BorderStyleConstants: BorderStyle = TV.BorderStyle: End Property
Public Property Let BorderStyle(BorderStyle As ComctlLib.BorderStyleConstants): TV.BorderStyle = BorderStyle: End Property
Public Property Get Appearance() As AppearanceConstants: Appearance = TV.Appearance: End Property
Public Property Let Appearance(Appearance As AppearanceConstants): TV.Appearance = Appearance: End Property
Public Property Get Indentation() As Long: Indentation = TV.Indentation: End Property
Public Property Let Indentation(Indentation As Long): TV.Indentation = Indentation: End Property
Public Property Get DisplayLogicalDrives() As Boolean: DisplayLogicalDrives = bDisplayLogicalDrives: End Property
Public Property Let DisplayLogicalDrives(DisplayLogicalDrives As Boolean): bDisplayLogicalDrives = DisplayLogicalDrives: Call ShowD: End Property
Public Property Get DisplayPhysicalDrives() As Boolean: DisplayPhysicalDrives = bDisplayPhysicalDrives: End Property
Public Property Let DisplayPhysicalDrives(DisplayPhysicalDrives As Boolean): bDisplayPhysicalDrives = DisplayPhysicalDrives: Call ShowD: End Property
Public Property Get HideSelection() As Boolean: HideSelection = TV.HideSelection: End Property
Public Property Let HideSelection(HideSelection As Boolean): TV.HideSelection = HideSelection: End Property
Public Property Get LogicalDrivesString() As String: LogicalDrivesString = sLog: End Property
Public Property Let LogicalDrivesString(LogicalDrivesString As String): sLog = LogicalDrivesString: Call ShowD: End Property
Public Property Get PhysicalDrivesString() As String: PhysicalDrivesString = sPhys: End Property
Public Property Let PhysicalDrivesString(PhysicalDrivesString As String): sPhys = PhysicalDrivesString: Call ShowD: End Property




'=======================================================
'SIMPLE EVENTS
'=======================================================
Private Sub TV_Click()
    RaiseEvent Click
End Sub
Private Sub TV_Collapse(ByVal Node As ComctlLib.Node)
    RaiseEvent Collapse(Node)
End Sub
Private Sub TV_DblClick()
    RaiseEvent DblClick
End Sub
Private Sub TV_Expand(ByVal Node As ComctlLib.Node)
    RaiseEvent Expand(Node)
End Sub
Private Sub TV_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub TV_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub TV_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub TV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub
Private Sub TV_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub
Private Sub TV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub
Private Sub TV_NodeClick(ByVal Node As ComctlLib.Node)
    RaiseEvent NodeClick(Node)
End Sub




'=======================================================
'PUBLIC FUNCTIONS & PROCEDURES
'=======================================================

'=======================================================
'récupère le drive sélection de type Drive
'=======================================================
Public Function GetSelectedDrive() As clsDrive
Dim s As String

    s = TV.SelectedItem.Text
    If Mid$(s, 3, 1) = "\" Then
        'alors c'est un drive logique
        Set GetSelectedDrive = cDrives.GetLogicalDrive(s)
    Else
        'alors c'est un drive physique
        Set GetSelectedDrive = cDrives.GetPhysicalDrive(CByte(Val(Mid$(s, 3, 1))))
    End If
End Function

'=======================================================
'refresh la vue
'=======================================================
Public Sub Refresh()
    Call ShowD
End Sub

'=======================================================
'renvoie si le drive sélectionné est accessible ou pas
'=======================================================
Public Function IsSelectedDriveAccessible() As Boolean
Dim s As String

    s = TV.SelectedItem.Text
    If Mid$(s, 3, 1) = "\" Then
        'alors c'est un drive logique
        IsSelectedDriveAccessible = cDrives.IsLogicalDriveAccessible(s)
    Else
        'alors c'est un drive physique
        IsSelectedDriveAccessible = cDrives.IsPhysicalDriveAccessible(CByte(Val(Mid$(s, 3, 1))))
    End If
    
End Function

'=======================================================
'clear le TV
'=======================================================
Public Function Clear()
    TV.Nodes.Clear
End Function

'=======================================================
'renvoie le nombre d'items visibles
'=======================================================
Public Function GetVisibleCount() As Long
    GetVisibleCount = TV.GetVisibleCount
End Function

'=======================================================
'fonction HitTest
'=======================================================
Public Function HitTest(x As Single, y As Single) As Node
    Set HitTest = TV.HitTest(x, y)
End Function

'=======================================================
'pour récupérer les informations sur les drives
'=======================================================
Public Function Drives() As clsDiskInfos
    Set Drives = cDrives
End Function





'=======================================================
'PRIVATE FUNCTIONS & PROCEDURES
'=======================================================

'=======================================================
'affiche les drives
'=======================================================
Private Sub ShowD()
Dim y As Long
Dim x As Long
Dim s() As String
Dim Key As String
Dim cDrive As clsDrive


    If bReadyToRefresh = False Then Exit Sub
    
    TV.Nodes.Clear  'clear les nodes
    
    'si on affiche les disques logiques
    If bDisplayLogicalDrives Then
    
        'ajoute le premier node
        Call TV.Nodes.Add(, , "Logical", sLog, "PhysicalDrive")
    
        y = cDrives.GetLogicalDrivesList(s())
        
        For x = 0 To y - 1
            
            'on ajoute l'image récupérée dans le IMG
            Call AddIconToIMG(Left$(s(x), 3), "log" & Trim$(Str$(x)))
            
            'on ajoute le noeud
            Call TV.Nodes.Add("Logical", tvwChild, "log" & Trim$(Str$(x)), Left$(s(x), 3), "log" & Trim$(Str$(x)))
            
        Next x
    End If
    
    
    'si on affiche les disques physiques
    If bDisplayPhysicalDrives Then
    
        'ajoute le premier node
        Call TV.Nodes.Add(, , "Physical", sPhys, "PhysicalDrive")

        'on liste tous les drives logiques et physiques
        For y = 0 To cDrives.NumberOfPhysicalDrives - 1
            
            'obtient les infos sur le drive
            Set cDrive = cDrives.GetPhysicalDrive(y)
            
            
            If cDrives.IsPhysicalDriveAccessible(y) Then
                  
                'on ajoute le noeud
                Call TV.Nodes.Add("Physical", tvwChild, "phys" & Trim$(Str$(y)), "HD" & Trim$(Str$(y)) & "  " & cDrive.VolumeName, "PhysicalDrive")
                
            End If
        Next y
        
        'libère
        Set cDrive = Nothing
    End If
    

End Sub

'=======================================================
'ajoute une icone à l'ImageList (à partir du fichier sFile)
'=======================================================
Private Sub AddIconToIMG(ByVal sDrive As String, ByVal sKey As String)
Dim lstImg As ListImage
Dim hIcon As Long
Dim ShInfo As SHFILEINFO
Dim pct As IPictureDisp
    
    If DoesKeyExist(sKey) Then Exit Sub    'clé existe déjà
    
    'obtient le handle de l'icone
    hIcon = SHGetFileInfo(sDrive, 0&, ShInfo, Len(ShInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
        
    'prépare la picturebox
    pctIcon.Picture = Nothing
    
    'trace l'image
    ImageList_Draw hIcon, ShInfo.iIcon, pctIcon.hDC, 0, 0, ILD_TRANSPARENT
    
    'ajout de l'icone à l'imagelist
    IMG.ListImages.Add Key:=sKey, Picture:=pctIcon.Image

End Sub

'=======================================================
'renvoie si la clé existe ou non deja dans IMG
'=======================================================
Private Function DoesKeyExist(ByVal sKey As String) As Boolean
Dim l As Long

    DoesKeyExist = False
    
    On Error GoTo ErrGest
    
    l = IMG.ListImages(sKey).Index

    DoesKeyExist = True
    
    Exit Function
    
ErrGest:
'la clé n'existait pas
End Function
