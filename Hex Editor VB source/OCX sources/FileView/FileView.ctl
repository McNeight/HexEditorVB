VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.UserControl FileView 
   ClientHeight    =   6825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   6825
   ScaleWidth      =   9480
   Begin VB.PictureBox pctIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   480
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   240
   End
   Begin ComctlLib.ListView LV 
      Height          =   3135
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      SmallIcons      =   "IMG"
      ForeColor       =   -2147483640
      BackColor       =   16777215
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Nom"
         Object.Width           =   9701
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Taille"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Date de création"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Date de dernier accès"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Date de dernière modification"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Attribut"
         Object.Width           =   1235
      EndProperty
   End
   Begin ComctlLib.ImageList IMG 
      Left            =   360
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileView.ctx":0000
            Key             =   "File"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileView.ctx":0C52
            Key             =   "Directory"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileView.ctx":18A4
            Key             =   "Drive"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileView.ctx":33F6
            Key             =   "FileWithoutExtension"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FileView"
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
'//FILEVIEW USERCONTROL : permet d'avoir une liste des fichiers/dossiers du disque dur
'PAR violent_ken
'v2.0 codé le 14/01/2007
'=======================================================

'=======================================================
'//IMPORTANT NOTE : YOU HAVE TO ADD THE "MICROSOFT SCRIPTING RUNTIME" REFERENCE
'//IN YOUR PROJECT TO USE THIS USERCONTROL
'=======================================================




'=======================================================
'DECLARATIONS autre que les variables
'=======================================================
Private fs As FileSystemObject



'=======================================================
'APIs
'=======================================================
'déplacement des fichiers vers la corbeille
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As Fichier) As Long
'obtient des infos (utilisé pour l'icone de l'executable) d'un fichier
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
'créé une image à partir d'un handle d'icone
Private Declare Function ImageList_Draw Lib "Comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal Flags&) As Long
'autorise l'affichage d'un listview
Private Declare Sub InvalidateRect Lib "user32" (ByVal hwnd As Long, ByVal t As Long, ByVal bErase As Long)
'bloque l'affichage d'un listview
Private Declare Sub ValidateRect Lib "user32" (ByVal hwnd As Long, ByVal t As Long)
'Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
'pour les bench
'Private Declare Function GetTickCount Lib "kernel32" () As Long

'pour la fonction GetFileSize
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileSizeEx Lib "kernel32" (ByVal hFile As Long, lpFileSize As Currency) As Boolean
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


'=======================================================
'CONSTANTES
'=======================================================
Private Const FO_DELETE                     As Long = &H3
Private Const FOF_ALLOWUNDO                 As Long = &H40
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
'EVENTS
'=======================================================
Public Event ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
Public Event AfterLabelEdit(Cancel As Integer, OldString As String, NewString As String)
Public Event BeforeLabelEdit(Cancel As Integer)
Public Event Click()
Public Event DblClick()
Public Event ItemClick(ByVal Item As ComctlLib.ListItem)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event PathChange(sOldPath As String, sNewPath As String)
Public Event PatternChange(sOldPattern As String, sNewPattern As String)
Public Event ItemDblSelection(Item As ComctlLib.ListItem)


'=======================================================
'ENUMS AND TYPES
'=======================================================
Public Enum IconDisplay
    NoIcons = 0
    BasicIcons = 1
    FileIcons = 2
End Enum
Public Enum Item_Type
    File = 1
    Directory = 2
    Drive = 3
End Enum
Private Type Fichier    'utilisé pour envoyer vers la corbeille
     hwnd As Long
     wFunc As Long
     pFrom As String
     pTo As String
     fFlags As Integer
     fAnyOperationsAborted As Boolean
     hNameMappings As Long
     lpszProgressTitle As String
End Type
Private Type SHFILEINFO 'utilisé pour récupérer les icones des fichiers
    hIcon        As Long
    iIcon        As Long
    dwAttributes    As Long
    szDisplayName As String * 260
    szTypeName  As String * 80
End Type




'=======================================================
'VARIABLES
'=======================================================
Private bLabelWrap As Boolean
Private tView As ListViewConstants
Private bAppearence3D As Boolean
Private lListCount As Long
Private lListIndex As Long
Private bShowHiddenFiles As Boolean
Private bShowSystemFiles As Boolean
Private bShowReadOnlyFiles As Boolean
Private bShowDirectories As Boolean
Private bShowFiles As Boolean
Private bShowHiddenDirectories As Boolean
Private bShowSystemDirectories As Boolean
Private bShowReadOnlyDirectories As Boolean
Private bShowNormalFiles As Boolean
Private bShowNormalDirectories As Boolean
Private bAllowDirectoryEntering As Boolean
Private bAllowFileDeleting As Boolean
Private bAllowFileRenaming As Boolean
Private bAllowDirectoryRenaming As Boolean
Private bAllowDirectoryDeleting As Boolean
Private bDisplayIcons As IconDisplay
Private bAllowReorganisationByColumn As Boolean
Private lBackColor As OLE_COLOR
Private lForeColor As OLE_COLOR
Private bEnabled As Boolean
Private bAllowMultiSelect As Boolean
Private sPattern As String
Private bHideColumnHeaders As Boolean
Private sPath As String
Private iSizeDecimals As Integer
Private bStillOkForRefresh As Boolean
Private sExt() As String
Private lFiles As Long
Private lDirectories As Long
Private lDrives As Long
Private bShowEntirePath As Boolean
Private bShowDrives As Boolean
Private lItemWidth As Long


'=======================================================
'SIMPLE EVENTS
'=======================================================
Private Sub LV_Click(): RaiseEvent Click: End Sub
Private Sub LV_ItemClick(ByVal Item As ComctlLib.ListItem): RaiseEvent ItemClick(Item): End Sub
Private Sub LV_KeyPress(KeyAscii As Integer): RaiseEvent KeyPress(KeyAscii): End Sub
Private Sub LV_KeyUp(KeyCode As Integer, Shift As Integer): RaiseEvent KeyUp(KeyCode, Shift): End Sub
Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single): RaiseEvent MouseDown(Button, Shift, x, y): End Sub
Private Sub LV_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single): RaiseEvent MouseMove(Button, Shift, x, y): End Sub
Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single): RaiseEvent MouseUp(Button, Shift, x, y): End Sub



'=======================================================
'PROPERTIES (gros bloc....)
'=======================================================
Public Property Get LabelWrap() As Boolean: LabelWrap = bLabelWrap: End Property
Public Property Let LabelWrap(LabelWrap As Boolean): bLabelWrap = LabelWrap: LV.LabelWrap = LabelWrap: End Property
Public Property Get View() As ListViewConstants: View = tView: End Property
Public Property Let View(View As ListViewConstants): tView = View: LV.View = View: End Property
Public Property Get Appearence3D() As Boolean: Appearence3D = bAppearence3D: End Property
Public Property Let Appearence3D(Appearence3D As Boolean): bAppearence3D = Appearence3D: LV.Appearance = Abs(CLng(Appearence3D)): End Property
Public Property Get ListCount() As Long: lListCount = LV.ListItems.Count: ListCount = lListCount: End Property
Public Property Get ListIndex() As Long
If (LV.SelectedItem Is Nothing) Then
    lListIndex = -1
Else
    lListIndex = LV.SelectedItem.Index
End If
ListIndex = lListIndex
End Property
Public Property Get ShowHiddenFiles() As Boolean: ShowHiddenFiles = bShowHiddenFiles: End Property
Public Property Let ShowHiddenFiles(ShowHiddenFiles As Boolean): bShowHiddenFiles = ShowHiddenFiles: Refresh: End Property
Public Property Get ShowSystemFiles() As Boolean: ShowSystemFiles = bShowSystemFiles: End Property
Public Property Let ShowSystemFiles(ShowSystemFiles As Boolean): bShowSystemFiles = ShowSystemFiles: Refresh: End Property
Public Property Get ShowReadOnlyFiles() As Boolean: ShowReadOnlyFiles = bShowReadOnlyFiles: End Property
Public Property Let ShowReadOnlyFiles(ShowReadOnlyFiles As Boolean): bShowReadOnlyFiles = ShowReadOnlyFiles: Refresh: End Property
Public Property Get ShowDrives() As Boolean: ShowDrives = bShowDrives: End Property
Public Property Let ShowDrives(ShowDrives As Boolean): bShowDrives = ShowDrives: Refresh: End Property
Public Property Get ShowDirectories() As Boolean: ShowDirectories = bShowDirectories: End Property
Public Property Let ShowDirectories(ShowDirectories As Boolean): bShowDirectories = ShowDirectories: Refresh: End Property
Public Property Get ShowFiles() As Boolean: ShowFiles = bShowFiles: End Property
Public Property Let ShowFiles(ShowFiles As Boolean): bShowFiles = ShowFiles: Refresh: End Property
Public Property Get ShowHiddenDirectories() As Boolean: ShowHiddenDirectories = bShowHiddenDirectories: End Property
Public Property Let ShowHiddenDirectories(ShowHiddenDirectories As Boolean): bShowHiddenDirectories = ShowHiddenDirectories: Refresh: End Property
Public Property Get ShowSystemDirectories() As Boolean: ShowSystemDirectories = bShowSystemDirectories: End Property
Public Property Let ShowSystemDirectories(ShowSystemDirectories As Boolean): bShowSystemDirectories = ShowSystemDirectories: Refresh: End Property
Public Property Get ShowReadOnlyDirectories() As Boolean: ShowReadOnlyDirectories = bShowReadOnlyDirectories: End Property
Public Property Let ShowReadOnlyDirectories(ShowReadOnlyDirectories As Boolean): bShowReadOnlyDirectories = ShowReadOnlyDirectories: Refresh: End Property
Public Property Get AllowDirectoryEntering() As Boolean: AllowDirectoryEntering = bAllowDirectoryEntering: End Property
Public Property Let AllowDirectoryEntering(AllowDirectoryEntering As Boolean): bAllowDirectoryEntering = AllowDirectoryEntering: End Property
Public Property Get AllowFileDeleting() As Boolean: AllowFileDeleting = bAllowFileDeleting: End Property
Public Property Let AllowFileDeleting(AllowFileDeleting As Boolean): bAllowFileDeleting = AllowFileDeleting: End Property
Public Property Get AllowFileRenaming() As Boolean: AllowFileRenaming = bAllowFileRenaming: End Property
Public Property Let AllowFileRenaming(AllowFileRenaming As Boolean): bAllowFileRenaming = AllowFileRenaming: LV.LabelEdit = IIf(AllowFileRenaming Or bAllowDirectoryRenaming, lvwAutomatic, lvwManual): End Property
Public Property Get AllowDirectoryRenaming() As Boolean: AllowDirectoryRenaming = bAllowDirectoryRenaming: End Property
Public Property Let AllowDirectoryRenaming(AllowDirectoryRenaming As Boolean): bAllowDirectoryRenaming = AllowDirectoryRenaming: LV.LabelEdit = IIf(AllowDirectoryRenaming Or bAllowFileRenaming, lvwAutomatic, lvwManual): End Property
Public Property Get AllowDirectoryDeleting() As Boolean: AllowDirectoryDeleting = bAllowDirectoryDeleting: End Property
Public Property Let AllowDirectoryDeleting(AllowDirectoryDeleting As Boolean): bAllowDirectoryDeleting = AllowDirectoryDeleting: End Property
Public Property Get DisplayIcons() As IconDisplay: DisplayIcons = bDisplayIcons: End Property
Public Property Let DisplayIcons(DisplayIcons As IconDisplay): bDisplayIcons = DisplayIcons: Refresh: End Property
Public Property Get AllowReorganisationByColumn() As Boolean: AllowReorganisationByColumn = bAllowReorganisationByColumn: End Property
Public Property Let AllowReorganisationByColumn(AllowReorganisationByColumn As Boolean): bAllowReorganisationByColumn = AllowReorganisationByColumn: End Property
Public Property Get BackColor() As OLE_COLOR: BackColor = lBackColor: End Property
Public Property Let BackColor(BackColor As OLE_COLOR): lBackColor = BackColor: LV.BackColor = BackColor: End Property
Public Property Get ForeColor() As OLE_COLOR: ForeColor = lForeColor: End Property
Public Property Let ForeColor(ForeColor As OLE_COLOR): lForeColor = ForeColor: LV.ForeColor = ForeColor: End Property
Public Property Get Enabled() As Boolean: Enabled = bEnabled: End Property
Public Property Let Enabled(Enabled As Boolean): bEnabled = Enabled: LV.Enabled = Enabled: End Property
Public Property Get AllowMultiSelect() As Boolean: AllowMultiSelect = bAllowMultiSelect: End Property
Public Property Let AllowMultiSelect(AllowMultiSelect As Boolean): bAllowMultiSelect = AllowMultiSelect: LV.MultiSelect = AllowMultiSelect: End Property
Public Property Get Pattern() As String: Pattern = sPattern: End Property
Public Property Let Pattern(Pattern As String)
If sPattern <> Pattern Then RaiseEvent PatternChange(sPattern, Pattern)
sPattern = Pattern: Refresh: End Property
Public Property Get HideColumnHeaders() As Boolean: HideColumnHeaders = bHideColumnHeaders: End Property
Public Property Let HideColumnHeaders(HideColumnHeaders As Boolean): bHideColumnHeaders = HideColumnHeaders: LV.HideColumnHeaders = HideColumnHeaders: End Property
Public Property Get hwnd() As Long: hwnd = UserControl.hwnd: End Property
Public Property Get Path() As String: Path = sPath: End Property
Public Property Let Path(Path As String)
RaiseEvent PathChange(Me.Path, Path)
sPath = Path
Refresh
End Property
Public Property Get SizeDecimals() As Integer: SizeDecimals = iSizeDecimals: End Property
Public Property Let SizeDecimals(SizeDecimals As Integer): iSizeDecimals = SizeDecimals: Refresh: End Property
Public Property Get ShowNormalDirectories() As Boolean: ShowNormalDirectories = bShowNormalDirectories: End Property
Public Property Let ShowNormalDirectories(ShowNormalDirectories As Boolean): bShowNormalDirectories = ShowNormalDirectories: Refresh: End Property
Public Property Get ShowNormalFiles() As Boolean: ShowNormalFiles = bShowNormalFiles: End Property
Public Property Let ShowNormalFiles(ShowNormalFiles As Boolean): bShowNormalFiles = ShowNormalFiles: Refresh: End Property
Public Property Get Files() As Integer: Files = lFiles: End Property
Public Property Get Directories() As Integer: Directories = lDirectories: End Property
Public Property Get Drives() As Integer: Drives = lDrives: End Property
Public Property Get ListItems() As ListItems: Set ListItems = LV.ListItems: End Property
Public Property Get ShowEntirePath() As Boolean: ShowEntirePath = bShowEntirePath: End Property
Public Property Let ShowEntirePath(ShowEntirePath As Boolean): bShowEntirePath = ShowEntirePath: Refresh: End Property
Public Property Get ItemWidth() As Long: ItemWidth = lItemWidth: End Property
Public Property Let ItemWidth(ItemWidth As Long)
lItemWidth = ItemWidth
LV.ColumnHeaders.Item(1).Width = ItemWidth
End Property
Public Property Get Font() As StdFont: Set Font = UserControl.Font: End Property
Public Property Set Font(Font As StdFont)
    Set UserControl.Font = Font
    LV.Font = Font
End Property


'=======================================================
'USERCONTROL SUBS
'=======================================================
Private Sub UserControl_InitProperties()
'properties by default
    Me.AllowDirectoryDeleting = True
    Me.AllowDirectoryEntering = True
    Me.AllowFileDeleting = True
    Me.AllowMultiSelect = True
    Me.AllowFileRenaming = True
    Me.AllowDirectoryRenaming = True
    Me.AllowReorganisationByColumn = True
    Me.Appearence3D = False
    Me.BackColor = vbWhite
    Me.DisplayIcons = FileIcons
    Me.Enabled = True
    Me.ForeColor = vbBlack
    Me.HideColumnHeaders = False
    Me.LabelWrap = True
    Me.Path = App.Path
    Me.Pattern = "*.*"
    Me.ShowDirectories = True
    Me.ShowFiles = True
    Me.ShowHiddenDirectories = True
    Me.ShowHiddenFiles = True
    Me.ShowReadOnlyFiles = True
    Me.ShowNormalDirectories = True
    Me.ShowNormalFiles = True
    Me.ShowReadOnlyDirectories = True
    Me.ShowSystemDirectories = True
    Me.ShowSystemFiles = True
    Me.SizeDecimals = 3
    Me.View = lvwReport
    Me.ShowEntirePath = True
    bStillOkForRefresh = False
    Me.ShowDrives = True
    Me.ItemWidth = 5500
    Me.Font = Ambient.Font
    Refresh
End Sub
Private Sub UserControl_Resize()
'resize le Listview à la taille du usercontrol
    LV.Height = UserControl.Height
    LV.Width = UserControl.Width
    LV.Left = 0
    LV.Top = 0
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Me.ShowEntirePath = .ReadProperty("ShowEntirePath", True)
        Me.AllowDirectoryDeleting = .ReadProperty("AllowDirectoryDeleting", True)
        Me.AllowDirectoryEntering = .ReadProperty("AllowDirectoryEntering", True)
        Me.AllowFileDeleting = .ReadProperty("AllowFileDeleting", True)
        Me.AllowMultiSelect = .ReadProperty("AllowMultiSelect", True)
        Me.AllowFileRenaming = .ReadProperty("AllowFileRenaming", True)
        Me.AllowDirectoryRenaming = .ReadProperty("AllowDirectoryRenaming", True)
        Me.AllowReorganisationByColumn = .ReadProperty("AllowReorganisationByColumn", True)
        Me.Appearence3D = .ReadProperty("Appearence3D", False)
        Me.BackColor = .ReadProperty("BackColor", vbWhite)
        Me.DisplayIcons = .ReadProperty("DisplayIcons", FileIcons)
        Me.Enabled = .ReadProperty("Enabled", True)
        Me.ForeColor = .ReadProperty("ForeColor", vbBlack)
        Me.HideColumnHeaders = .ReadProperty("HideColumnHeaders", False)
        Me.LabelWrap = .ReadProperty("LabelWrap", True)
        Me.Path = .ReadProperty("Path", App.Path)
        Me.Pattern = .ReadProperty("Pattern", "*.*")
        Me.ShowDrives = .ReadProperty("ShowDrives", True)
        Me.ShowDirectories = .ReadProperty("ShowDirectories", True)
        Me.ShowFiles = .ReadProperty("ShowFiles", True)
        Me.ShowHiddenDirectories = .ReadProperty("ShowHiddenDirectories", True)
        Me.ShowHiddenFiles = .ReadProperty("ShowHiddenFiles", True)
        Me.ShowReadOnlyFiles = .ReadProperty("ShowReadOnlyFiles", True)
        Me.ShowReadOnlyDirectories = .ReadProperty("ShowReadOnlyDirectories", True)
        Me.ShowSystemDirectories = .ReadProperty("ShowSystemDirectories", True)
        Me.ShowSystemFiles = .ReadProperty("ShowSystemFiles", True)
        Me.ShowNormalDirectories = .ReadProperty("ShowNormalDirectories", True)
        Me.ShowNormalFiles = .ReadProperty("ShowNormalFiles", True)
        Me.View = .ReadProperty("View", lvwReport)
        Me.SizeDecimals = .ReadProperty("SizeDecimals", 3)
        Me.ItemWidth = .ReadProperty("ItemWidth", 5500)
        Set Me.Font = .ReadProperty("Font", Ambient.Font)
    End With
    
    'alors c'est bon, on rafraichit
    'ceci évite de rafraichir pour CHAQUE property à l'entrée dans le controle

    If bStillOkForRefresh Then
        bStillOkForRefresh = False  'on ne rafraichira plus à l'entrée au focus
        Refresh
    End If
End Sub
Private Sub UserControl_Terminate()
    Set fs = Nothing
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("ShowEntirePath", Me.ShowEntirePath, True)
        Call .WriteProperty("SizeDecimals", Me.SizeDecimals, 3)
        Call .WriteProperty("AllowDirectoryDeleting", Me.AllowDirectoryDeleting, True)
        Call .WriteProperty("AllowDirectoryEntering", Me.AllowDirectoryEntering, True)
        Call .WriteProperty("AllowFileDeleting", Me.AllowFileDeleting, True)
        Call .WriteProperty("AllowMultiSelect", Me.AllowMultiSelect, True)
        Call .WriteProperty("AllowFileRenaming", Me.AllowFileRenaming, True)
        Call .WriteProperty("AllowDirectoryRenaming", Me.AllowDirectoryRenaming, True)
        Call .WriteProperty("AllowReorganisationByColumn", Me.AllowReorganisationByColumn, True)
        Call .WriteProperty("Appearence3D", Me.Appearence3D, False)
        Call .WriteProperty("BackColor", Me.BackColor, vbWhite)
        Call .WriteProperty("DisplayIcons", Me.DisplayIcons, FileIcons)
        Call .WriteProperty("Enabled", Me.Enabled, True)
        Call .WriteProperty("ForeColor", Me.ForeColor, vbBlack)
        Call .WriteProperty("HideColumnHeaders", Me.HideColumnHeaders, False)
        Call .WriteProperty("LabelWrap", Me.LabelWrap, True)
        Call .WriteProperty("Path", Me.Path, App.Path)
        Call .WriteProperty("Pattern", Me.Pattern, "*.*")
        Call .WriteProperty("ShowDirectories", Me.ShowDirectories, True)
        Call .WriteProperty("ShowFiles", Me.ShowFiles, True)
        Call .WriteProperty("ShowHiddenDirectories", Me.ShowHiddenDirectories, True)
        Call .WriteProperty("ShowHiddenFiles", Me.ShowHiddenFiles, True)
        Call .WriteProperty("ShowNormalDirectories", Me.ShowNormalDirectories, True)
        Call .WriteProperty("ShowNormalFiles", Me.ShowNormalFiles, True)
        Call .WriteProperty("ShowReadOnlyDirectories", Me.ShowReadOnlyDirectories, True)
        Call .WriteProperty("ShowReadOnlyFiles", Me.ShowReadOnlyFiles, True)
        Call .WriteProperty("ShowSystemDirectories", Me.ShowSystemDirectories, True)
        Call .WriteProperty("ShowSystemFiles", Me.ShowSystemFiles, True)
        Call .WriteProperty("View", Me.View, lvwReport)
        Call .WriteProperty("ShowDrives", Me.ShowDrives, True)
        Call .WriteProperty("ItemWidth", Me.ItemWidth, 5500)
        Call .WriteProperty("Font", Me.Font, Ambient.Font)
    End With
End Sub
Private Sub UserControl_Initialize()
    bStillOkForRefresh = True   'alors on est prêt à attendre l'entrée en focus pour pouvoir refresh
    Set fs = New FileSystemObject   'définit l'objet fs
End Sub



'=======================================================
'EVENTS PARTICULIERS
'=======================================================
Private Sub LV_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
'classement (ou pas) des columns
Dim bOkForAjout As Boolean

    LV.Sorted = bAllowReorganisationByColumn

    If bAllowReorganisationByColumn Then
        'alors réorganisation des items
        'réorganisation par ordre alphanumérique (basique)
        
        'pas le temps de faire un tri réel des taille des fichiers et des dates
        'peut être que ce sera changé pour HexEditor
        'mais ce sera inclus de manière certaine à Process Guardian
        
        
        'on enlève l'item qui représente le "parent folder"
        bOkForAjout = False
        If LV.ListItems(1).Text = "..\" Then
            'il est tout au début
            bOkForAjout = True  'on pourra le réajouter après
            LV.ListItems.Remove 1
        End If
        
        If LV.SortKey = ColumnHeader.Index - 1 Then
            If Not LV.SortOrder = lvwAscending Then
                LV.SortOrder = lvwAscending
            Else
                LV.SortOrder = lvwDescending
            End If
        Else
            LV.SortKey = ColumnHeader.Index - 1
            LV.SortOrder = lvwAscending
        End If
        
        LV.Sorted = False
        
        'on ré-ajoute l'élément "../" tout au début
        If bOkForAjout Then
            'on ajoute
            If bDisplayIcons = NoIcons Then
                LV.ListItems.Add Index:=1, Text:="..\"
            Else
                LV.ListItems.Add Index:=1, Text:="..\", SmallIcon:="Directory"
            End If
        End If
    End If


RaiseEvent ColumnClick(ColumnHeader)
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
'rentre dans l'item si
'-c'est un dossier
'-KeyCode=13 (entrée)
Dim sOld As String
Dim s As String

    If KeyCode = 13 Then
        'si entrée dans les dossiers activée
        If Me.AllowDirectoryEntering = True Then
            'si éléments est un dossier
            If fs.FolderExists(LV.SelectedItem.Tag) And LV.SelectedItem.Text <> "..\" Then
                'alors on l'ouvre
                sOld = Me.Path
                s = LV.SelectedItem.Tag
                If Len(s) = 2 Then s = s & "\" 'rajoute un '\' dans le cas d'un drive
                Me.Path = s
                RaiseEvent PathChange(sOld, Me.Path)
            ElseIf LV.SelectedItem.Text = "..\" Then
                'alors dossier parent
                sOld = Me.Path
                
                If fs.DriveExists(sOld) Then
                    'alors il faut maintenant afficher tout les drives
                    Me.Clear
                    AddDrives
                Else
                    'remonte d'un niveau
                    GoParentFolder
                End If
                
                RaiseEvent PathChange(sOld, Me.Path)
            End If
        End If
    End If
    
    
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub LV_AfterLabelEdit(Cancel As Integer, NewString As String)
'renomme le fichier
Dim sOld As String

    sOld = LV.SelectedItem.Tag
    
    If (fs.FolderExists(sOld) And Me.AllowDirectoryRenaming) Or _
    (fs.FileExists(sOld) And Me.AllowFileRenaming) Then
        'alors on renomme l'item
        MustRename LV.SelectedItem.Tag, ParentFolder(LV.SelectedItem.Tag) & NewString
    End If
    
RaiseEvent AfterLabelEdit(Cancel, sOld, NewString)
End Sub

Private Sub LV_BeforeLabelEdit(Cancel As Integer)
    If LV.SelectedItem.Text = "..\" Or fs.DriveExists(LV.SelectedItem.Tag) Then Cancel = 1  'ne renomme pas les drive et les dossier "remonter d'un niveau"
    If (fs.FolderExists(LV.SelectedItem.Tag) And Me.AllowDirectoryRenaming = False) Or _
    (fs.FileExists(LV.SelectedItem.Tag) And Me.AllowFileRenaming = False) Then Cancel = 1 'n'autorise pas le renommage
    RaiseEvent BeforeLabelEdit(Cancel)
End Sub

Private Sub LV_DblClick()
'rentre dans un dossier si nécessaire
Dim sOld As String
Dim bOkForEvent As Boolean
Dim s As String

    bOkForEvent = True
    
    'gère l'event
    RaiseEvent DblClick
    
    'si entrée dans les dossiers activée
    If Me.AllowDirectoryEntering = True Then
        'si éléments est un dossier
        If fs.FolderExists(LV.SelectedItem.Tag) And LV.SelectedItem.Text <> "..\" Then
            'alors on l'ouvre
            sOld = Me.Path
            s = IIf(Right$(LV.SelectedItem.Tag, 1) = "\", LV.SelectedItem.Tag, LV.SelectedItem.Tag & "\")
            If Len(s) = 2 Then s = s & "\" 'rajoute un '\' dans le cas d'un drive
            Me.Path = s
            RaiseEvent PathChange(sOld, Me.Path)
        ElseIf LV.SelectedItem.Text = "..\" Then
            'alors dossier parent
            sOld = Me.Path
            
            If fs.DriveExists(sOld) Then
                'alors il faut maintenant afficher tout les drives
                Me.Clear
                AddDrives
                bOkForEvent = False
            Else
                'remonte d'un niveau
                GoParentFolder
            End If
            
            RaiseEvent PathChange(sOld, Me.Path)
        End If
    End If
    
    If bOkForEvent Then RaiseEvent ItemDblSelection(LV.SelectedItem)
       
End Sub



'=======================================================
'FONCTIONS ET PROCEDURES PUBLIQUES
'=======================================================

'=======================================================
'obtient le nom des fichiers sélectionnés
'/!\ type variant attendu, mais renvoie un type ListItem (ComctLib)
'=======================================================
Public Sub GetSelectedItems(ByRef sFileArray As Variant)   '/!\ ByRef sFileArray() As ComctlLib.ListItem  attendu
Dim x As Long
'Dim sFileSelection() As ComctlLib.ListItem
Dim sFileSelection() As ComctlLib.ListItem

    ReDim sFileSelection(0)
    
    'récupère toutes les sélections
    For x = LV.ListItems.Count To 1 Step -1
        If LV.ListItems.Item(x).Selected Then
            'on ajoute un élément sélectioné
            ReDim Preserve sFileSelection(UBound(sFileSelection) + 1)
            Set sFileSelection(UBound(sFileSelection)) = LV.ListItems.Item(x)
        End If
    Next x
    
    sFileArray = sFileSelection
    
End Sub

'=======================================================
'utilisé pour éviter les bugs d'affichage au loading d'une form
'=======================================================
Public Sub RefreshListViewOnly()
    LV.Refresh
End Sub

'=======================================================
'rafraichit le control (son contenu)
'=======================================================
Public Sub Refresh()

    If bStillOkForRefresh = True Then Exit Sub  'contrôle pas encore chargé

    LV.Visible = False
    
    'clear ListView
    Clear
    'ajoute les fichiers du répertoire courant
    AddFiles
    
    LV.Visible = True
    
End Sub

'=======================================================
'efface le cotenu
'=======================================================
Public Sub Clear()
    LV.ListItems.Clear
End Sub

'=======================================================
'ajoute manuellement une entrée au listview
'=======================================================
Public Sub AddItemManually(ByVal sItem As String, ByVal tTypeOfItem As Item_Type, Optional ByVal sSubItem1 As String = vbNullString, Optional ByVal sSubItem2 As String = vbNullString, Optional ByVal sSubItem3 As String = vbNullString, Optional ByVal sSubItem4 As String = vbNullString, Optional ByVal sSubItem5 As String = vbNullString, Optional ByVal sSubItem6 As String = vbNullString, Optional ByVal bFillSubItemsAuto As Boolean = True)
Dim g As Long
Dim l As Long
Dim s As String, key As String


    l = MustGetAttr(sItem)  'attribut

    If tTypeOfItem = Directory Then
        'dossier
        If bDisplayIcons <> NoIcons Then
            'affiche une icone
            LV.ListItems.Add Text:=IIf(bShowEntirePath, sItem, GetFileFromPath(sItem)), SmallIcon:="Directory"
            LV.ListItems.Item(LV.ListItems.Count).Tag = sItem
        Else
            LV.ListItems.Add Text:=IIf(bShowEntirePath, sItem, GetFileFromPath(sItem))
            LV.ListItems.Item(LV.ListItems.Count).Tag = sItem
        End If
    ElseIf tTypeOfItem = File Then
        'fichier
        If IsPatternOk(sItem) Then  'vérifie de suite le pattern
            If bDisplayIcons = NoIcons Then
                'pas d'icones
                AddFileToLV sItem, vbNullString, l
            ElseIf bDisplayIcons = BasicIcons Then
                'icones basiques
                AddFileToLV sItem, "File", l
            ElseIf bDisplayIcons = FileIcons Then
                'icone par fichier
                        
                'obtient l'extension du fichier
                g = InStrRev(sItem, ".", , vbBinaryCompare)
                        
                If g <> 0 Then
                    'alors il y a une extension
                    s = LCase(Right$(sItem, Len(sItem) - g)) 's=extension
                            
                    If s = "ico" Or s = "exe" Then
                        'il faut récupérer l'icone du fichier
                        key = sItem & "_"
                        AddIconToIMG sItem, key
                        AddFileToLV sItem, key, l
                    Else
                        'il faut récupérer l'icone associée au type de fichier
                        key = "_" & s & "_"    'clé unique d'extension
                        AddIconToIMG sItem, key
                        AddFileToLV sItem, key, l
                    End If
                Else
                    'pas d'extension
                    key = "FileWithoutExtension"
                    AddFileToLV sItem, key, l
                End If
            End If
        End If
    ElseIf tTypeOfItem = Drive Then
        If bDisplayIcons <> NoIcons Then
            'affiche une icone
            LV.ListItems.Add Text:=sItem, SmallIcon:="Drive"
        Else
            LV.ListItems.Add Text:=sItem
        End If
    End If
    
    If bFillSubItemsAuto Then
        'alors remplit automatiquement les champs
        If fs.FileExists(sItem) Then
            With LV.ListItems(LV.ListItems.Count)
                .SubItems(1) = FormatedSize(GetFileSize(sItem), iSizeDecimals)
                .SubItems(2) = fs.GetFile(sItem).Type
                .SubItems(3) = fs.GetFile(sItem).DateCreated
                .SubItems(4) = fs.GetFile(sItem).DateLastAccessed
                .SubItems(5) = fs.GetFile(sItem).DateLastModified
                .SubItems(6) = fs.GetFile(sItem).Attributes
            End With
        End If
    Else
        'ajoute les champs spécifiés
        With LV.ListItems.Item(LV.ListItems.Count)
            .SubItems(1) = sSubItem1
            .SubItems(2) = sSubItem2
            .SubItems(3) = sSubItem3
            .SubItems(4) = sSubItem4
            .SubItems(5) = sSubItem5
            .SubItems(6) = sSubItem6
        End With
    End If
End Sub

'=======================================================
'enlève un item définit par son index
'=======================================================
Public Sub RemoveItemManually(ByVal Index As Long)
    
    On Error Resume Next
    
    LV.ListItems.Remove Index
    
End Sub

'=======================================================
'efface le(s) fichier(s)/dossier(s) sélectionné(s) du disque dur
'peut aussi déplacer vers la corbeille
'peut afficher un message d'alerte
'=======================================================
Public Sub DeleteSelectedItemsFromDisk(ByVal bAskConfirmation As Boolean, Optional ByVal sConfirmationTitle As String = "Attention", Optional ByVal sConfirmationMessage As String = "Vous allez supprimer des éléments de votre disque dur. Continuer ?", Optional ByVal bMoveToTrash As Boolean = True, Optional ByVal Force As Boolean = False)
Dim lRet As Long
Dim x As Long
Dim sItem As String

    If bAskConfirmation Then
        'alors demande une confirmation avant de supprimer
        lRet = MsgBox(sConfirmationMessage, vbYesNo, sConfirmationTitle)
        If Not (lRet = vbYes) Then Exit Sub
    End If
    
    'procède à la suppression de tous les objets
    For x = LV.ListItems.Count To 1 Step -1
        sItem = LV.ListItems.Item(x).Tag
        
        'vérfie que l'on à l'autorisation de deleter
        If (fs.FolderExists(sItem) And Me.AllowDirectoryDeleting) Or _
        (fs.FileExists(sItem) And Me.AllowFileDeleting) Then
        
            If LV.ListItems.Item(x).Selected Then
                'le supprime
                If bMoveToTrash Then
                    'vers la corbeille
                    MoveToTrash sItem
                    If (fs.FolderExists(sItem) = False And fs.FileExists(sItem) = False) Then LV.ListItems.Remove x
                Else
                    'supprime directement
                    If fs.FolderExists(sItem) Then
                        'supprime un dossier
                        fs.DeleteFolder sItem, Force
                        If fs.FolderExists(sItem) = False Then
                            LV.ListItems.Remove x
                        End If
                    ElseIf fs.FileExists(sItem) Then
                        'supprime un fichier
                        fs.DeleteFile sItem, Force
                        If fs.FileExists(sItem) = False Then
                            LV.ListItems.Remove x
                        End If
                    End If
                End If
            End If
        End If
    Next x

End Sub

'=======================================================
'va au dossier parent
'=======================================================
Public Sub GoParentFolder()
    Me.Path = ParentFolder(Me.Path)
End Sub

'=======================================================
'obtient le dossier parent
'=======================================================
Public Function GetParentFolder() As String
    GetParentFolder = ParentFolder(Me.Path)
End Function

'=======================================================
'créé un dossier dans le path actuel
'=======================================================
Public Sub CreateDirectoryInCurrentPath(ByVal sDirectory_with_or_without_current_path As String)
Dim s As String

    s = sDirectory_with_or_without_current_path
    
    'path + dossier
    s = IIf(ParentFolder(s) = Me.Path, s, Me.Path & "\" & s)
    
    If fs.FolderExists(s) Then Exit Sub 'existe deja
    
    'créé le dossier
    fs.CreateFolder (s)

    'l'ajoute si il existe (si a réussi à le créé juste avant)
    If fs.FolderExists(s) Then Me.AddItemManually s, Directory
    
End Sub

'=======================================================
'créé un fichier dans le Path actuel
'=======================================================
Public Sub CreateFileInCurrentPath(ByVal sFile_with_or_without_path As String, Optional ByVal bOverWrite As Boolean = False)
Dim s As String

    'path + fichier
    s = Me.Path & "\" & GetFileFromPath(sFile_with_or_without_path)
    
    If fs.FileExists(s) And Not (bOverWrite) Then Exit Sub
    
    'créé le fichier
    fs.CreateTextFile s, bOverWrite
    
    'l'ajoute si il existe (si a réussi à le créé juste avant)
    If fs.FileExists(s) Then Me.AddItemManually s, File

End Sub



'=======================================================
'FONCTIONS DIRECTES SUR LE LISTVIEW
'=======================================================
Public Function GetFirstVisible() As ComctlLib.ListItem
    GetFirstVisible = LV.GetFirstVisible
End Function
Public Function HitTest(x As Single, y As Single) As ComctlLib.ListItem
    Set HitTest = LV.HitTest(x, y)
End Function
Public Function FindItem(sz As String, Optional Where As ListFindItemWhereConstants, Optional Index As Variant, Optional fPartial As ListFindItemHowConstants) As ComctlLib.ListItem
    FindItem = LV.FindItem(sz, Where, Index, fPartial)
End Function



'=======================================================
'FONCTIONS ET PROCEDURES PRIVEES AU USERCONTROL
'=======================================================

'=======================================================
'ajoute les fichiers du dossier sPath et les sous dossiers de ce dossier sPath dans le LV
'en fonction des paramètres du usercontrol
'=======================================================
Private Sub AddFiles()
Dim sDirectory() As String
Dim sFile() As String
Dim sF As String, key As String, s As String
Dim x As Long, l As Long, g As Long, i As Long
Dim objFolder As Object
Dim ssFolder As Object

    On Error GoTo ErrGestion
    
    'i = GetTickCount
    
    lFiles = 0: lDirectories = -1: lDrives = 0 'aucun objet dans le listview
    
    If fs.FolderExists(Me.Path) = False Then Exit Sub
    
    'désactive le rafraichissement du listview
    ValidateRect LV.hwnd, 0&
    'LockWindowUpdate LV.hWnd
    
    Set objFolder = fs.GetFolder(Me.Path)

    ReDim sFile(0)
    ReDim sDirectory(0)
    
    If Me.AllowDirectoryEntering Then
        'ajoute le répertoire "précédent"
        ReDim sDirectory(1)
        sDirectory(1) = "..\"
    Else
        lDirectories = 0
    End If

    If Me.ShowDirectories Then
    
        'énumère TOUS les dossier
        'utilisation de FSO pour avoir les dossiers systèmes...etc.
        'donc PAS DE DIR
        
        For Each ssFolder In objFolder.SubFolders
            ReDim Preserve sDirectory(UBound(sDirectory) + 1)
            sDirectory(UBound(sDirectory)) = ssFolder
        Next ssFolder
        
        'trie les dossiers en fonction des propriétés de l'user control
        For x = 1 To UBound(sDirectory)
        
            If (x Mod 500) = 0 Then DoEvents            'rend la main de temps en temps
                    
            l = MustGetAttr(sDirectory(x))
            
            lDirectories = lDirectories + 1
            
            If bDisplayIcons <> NoIcons Then
                'affiche une icone
                If IsAttrOK(l, bShowNormalDirectories, bShowHiddenDirectories, _
                bShowReadOnlyDirectories, bShowSystemDirectories) And bShowDirectories Then
                    LV.ListItems.Add Text:=IIf(bShowEntirePath, sDirectory(x), GetFileFromPath(sDirectory(x))), SmallIcon:="Directory"
                    LV.ListItems.Item(LV.ListItems.Count).Tag = sDirectory(x)
                End If
            Else
                'pas d'icone
                If IsAttrOK(l, bShowNormalDirectories, bShowHiddenDirectories, _
                bShowReadOnlyDirectories, bShowSystemDirectories) And bShowDirectories Then
                    LV.ListItems.Add Text:=IIf(bShowEntirePath, sDirectory(x), GetFileFromPath(sDirectory(x)))
                    LV.ListItems.Item(LV.ListItems.Count).Tag = sDirectory(x)
                End If
            End If
        Next x
        
    End If
        
       
    If Me.ShowFiles Then
    
        'énumère TOUS les fichiers
    
        For Each ssFolder In objFolder.Files
            ReDim Preserve sFile(UBound(sFile) + 1)
            sFile(UBound(sFile)) = ssFolder
        Next ssFolder
        
        'trie les fichiers en fonction des propriétés de l'user control
        For x = 1 To UBound(sFile)
            l = MustGetAttr(sFile(x))
            
            If (x Mod 500) = 0 Then DoEvents            'rend la main de temps en temps
            
            If IsPatternOk(sFile(x)) Then
                'vérifie tout de suite le pattern
            
                If bDisplayIcons = NoIcons Then
                    'pas d'icones
                    AddFileToLV sFile(x), vbNullString, l
                ElseIf bDisplayIcons = BasicIcons Then
                    'icones basiques
                    AddFileToLV sFile(x), "File", l
                ElseIf bDisplayIcons = FileIcons Then
                    'icone par fichier
                    
                    'obtient l'extension du fichier
                    g = InStrRev(sFile(x), ".", , vbBinaryCompare)
                    
                    If g <> 0 Then
                        'alors il y a une extension
                        s = LCase(Right$(sFile(x), Len(sFile(x)) - g)) 's=extension
                        
                        If s = "ico" Or s = "exe" Then
                            'il faut récupérer l'icone du fichier
                            key = sFile(x) & "_"
                            AddIconToIMG sFile(x), key
                            AddFileToLV sFile(x), key, l
                        Else
                            'il faut récupérer l'icone associée au type de fichier
                            key = "_" & s & "_"    'clé unique d'extension
                            AddIconToIMG sFile(x), key
                            AddFileToLV sFile(x), key, l
                        End If
                    Else
                        'pas d'extension
                        key = "FileWithoutExtension"
                        AddFileToLV sFile(x), key, l
                    End If
                End If
            End If
        Next x
    End If
    
    'réactive l'affichage du listview
    InvalidateRect hwnd, 0&, 0&
    'LockWindowUpdate 0&
    
    'MsgBox GetTickCount - i
    
    Exit Sub

ErrGestion:
    Me.GoParentFolder 'permission refusée ou autres erreurs d'accès
    
    'réactive l'affichage du listview
    InvalidateRect hwnd, 0&, 0&
    'LockWindowUpdate 0&
End Sub

'=======================================================
'ajoute les disques
'=======================================================
Public Sub AddDrives()
Dim sDirectory() As String
Dim sF As String
Dim x As Long, l As Long
Dim ssDrive As Object
Dim s As Object
Dim s2 As String
    
    If fs.DriveExists(Me.Path) = False Then Exit Sub

    On Error Resume Next    'pour les disques non prêts

    ReDim sDirectory(0)

    If Me.ShowDrives Then
    
        'énumère TOUS les Drives
        'utilisation de FSO pour avoir les dossiers systèmes...etc.
        'donc PAS DE DIR
        
        For Each ssDrive In fs.Drives
            ReDim Preserve sDirectory(UBound(sDirectory) + 1)
            s2 = fs.GetDrive(ssDrive).DriveLetter & ":\"
            s2 = s2 & " [" & fs.GetDrive(ssDrive).VolumeName & "][" & fs.GetDrive(ssDrive).FileSystem & "]"
            sDirectory(UBound(sDirectory)) = s2
        Next ssDrive
        
        'trie les dossiers en fonction des propriétés de l'user control
        For x = 1 To UBound(sDirectory)
            lDrives = lDrives + 1
            If bDisplayIcons <> NoIcons Then
                'affiche une icone
                LV.ListItems.Add Text:=sDirectory(x), SmallIcon:="Drive"
                LV.ListItems.Item(LV.ListItems.Count).Tag = ParentFolder(sDirectory(x))
            Else
                LV.ListItems.Add Text:=sDirectory(x)
                LV.ListItems.Item(LV.ListItems.Count).Tag = ParentFolder(sDirectory(x))
            End If
        Next x
        
    End If
    
End Sub

'=======================================================
'renvoie le dossier parent
'=======================================================
Private Function ParentFolder(ByVal sFolder As String) As String
Dim l As Long
Dim s As String

    If Right$(sFolder, 1) = "\" Then sFolder = Left$(sFolder, Len(sFolder) - 1) 'enlève le dernier '\' si présent
    
    l = InStrRev(sFolder, "\")
    
    If l = 0 Then Exit Function
    
    ParentFolder = Left$(sFolder, l)
End Function

'=======================================================
'déplace vers la corbeille un fichier
'=======================================================
Private Function MoveToTrash(sFile As String) As Long
Dim sFileToDelete As Fichier

    'définit l'objet Fichier
    With sFileToDelete
        .wFunc = FO_DELETE
        .pFrom = sFile
        .fFlags = FOF_ALLOWUNDO
    End With
    
    'procède à la suppression
    MoveToTrash = SHFileOperation(sFileToDelete)
End Function

'=======================================================
'formate la taille d'un fichier
'=======================================================
Private Function FormatedSize(ByVal lS As Currency, Optional ByVal lRoundNumber = 5) As String
Dim dS As Double
Dim n As Byte

    On Error Resume Next
    
    dS = lS: n = 0
    While (dS / 1024) > 1
        n = n + 1
        dS = dS / 1024
    Wend
    
    dS = Round(dS, lRoundNumber)
    
    If n = 0 Then FormatedSize = Str$(dS) & " Octets"
    If n = 1 Then FormatedSize = Str$(dS) & " Ko"
    If n = 2 Then FormatedSize = Str$(dS) & " Mo"
    If n = 3 Then FormatedSize = Str$(dS) & " Go"
    
    FormatedSize = Trim$(FormatedSize)
    
End Function

'=======================================================
'ajoute un fichier au listview
'=======================================================
Private Sub AddFileToLV(ByVal sFile As String, ByVal sImageKey As String, ByVal lAttribute As Long)
Dim g As Long
Dim l As Long

    l = lAttribute
    
    If sImageKey <> vbNullString Then
        'on ajoute une icone
        If IsAttrOK(l, bShowNormalFiles, bShowHiddenFiles, bShowReadOnlyFiles, _
        bShowSystemFiles) And bShowFiles Then
            LV.ListItems.Add Text:=IIf(bShowEntirePath, sFile, GetFileFromPath(sFile)), SmallIcon:=sImageKey
            LV.ListItems.Item(LV.ListItems.Count).Tag = sFile
            g = LV.ListItems.Count
            If fs.FileExists(sFile) Then
                lFiles = lFiles + 1
                With LV.ListItems.Item(g)
                    .SubItems(1) = FormatedSize(GetFileSize(sFile), iSizeDecimals)
                    .SubItems(2) = fs.GetFile(sFile).Type
                    .SubItems(3) = fs.GetFile(sFile).DateCreated
                    .SubItems(4) = fs.GetFile(sFile).DateLastAccessed
                    .SubItems(5) = fs.GetFile(sFile).DateLastModified
                    .SubItems(6) = fs.GetFile(sFile).Attributes
                End With
            End If
        End If
    Else
        'on ajoute pas d'icone
        
        If IsAttrOK(l, bShowNormalFiles, bShowHiddenFiles, bShowReadOnlyFiles, _
        bShowSystemFiles) And bShowFiles Then
            LV.ListItems.Add Text:=IIf(bShowEntirePath, sFile, GetFileFromPath(sFile))
            LV.ListItems.Item(LV.ListItems.Count).Tag = sFile
            g = LV.ListItems.Count
            If fs.FileExists(sFile) Then
                With LV.ListItems.Item(g)
                    .SubItems(1) = FormatedSize(GetFileSize(sFile), iSizeDecimals)
                    .SubItems(2) = fs.GetFile(sFile).Type
                    .SubItems(3) = fs.GetFile(sFile).DateCreated
                    .SubItems(4) = fs.GetFile(sFile).DateLastAccessed
                    .SubItems(5) = fs.GetFile(sFile).DateLastModified
                    .SubItems(6) = fs.GetFile(sFile).Attributes
                End With
            End If
        End If
    End If

End Sub

'=======================================================
'ajoute une icone à l'ImageList (à partir du fichier sFile)
'=======================================================
Private Sub AddIconToIMG(ByVal sFile As String, ByVal sKey As String)
Dim lstImg As ListImage
Dim hIcon As Long
Dim ShInfo As SHFILEINFO
Dim pct As IPictureDisp
    
    If fs.FileExists(sFile) = False Then Exit Sub   'fichier introuvable
    If DoesKeyExist(sKey) Then Exit Sub    'clé existe déjà
    
    'obtient le handle de l'icone
    hIcon = SHGetFileInfo(sFile, 0&, ShInfo, Len(ShInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
        
    'prépare la picturebox
    pctIcon.Picture = Nothing
    
    'trace l'image
    ImageList_Draw hIcon, ShInfo.iIcon, pctIcon.hDC, 0, 0, ILD_TRANSPARENT
    
    'ajout de l'icone à l'imagelist
    IMG.ListImages.Add key:=sKey, Picture:=pctIcon.Image

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

'=======================================================
'obtient un fichier depuis un path
'=======================================================
Private Function GetFileFromPath(ByVal sPath As String)
Dim l As Long

    If sPath = "..\" Then
        GetFileFromPath = sPath
        Exit Function
    End If

    l = InStrRev(sPath, "\", , vbBinaryCompare)
    
    If l = 0 Then
        'pas de path
        GetFileFromPath = sPath
        Exit Function
    End If
    
    GetFileFromPath = Right$(sPath, Len(sPath) - l)
    
End Function

'=======================================================
'vérifie que l'attribut est compatible avec les paramètres de l'usercontrol
'==> doit on l'afficher ou pas ?
'=======================================================
Private Function IsAttrOK(ByVal lAttribute As Long, ByVal bNormal As Boolean, ByVal bHidden As Boolean, ByVal bReadOnly As Boolean, ByVal bSystem As Boolean) As Boolean

    
    IsAttrOK = True
    
    'Normal  0   Fichier normal
    'ReadOnly    1   Fichier en lecture seule
    'Hidden  2   Fichier caché
    'System  4   Fichier système
    'Volume  8   Le nom " label " du lecteur
    'Directory   16  Répertoire
    'Archive 32  Fichier archive
    'Alias   64  Raccourci
    'Compressed  128 Fichier compressé

    If (lAttribute And vbSystem) = vbSystem And bSystem = False Then IsAttrOK = False
    If (lAttribute And vbNormal) = vbNormal And bNormal = False Then IsAttrOK = False
    If (lAttribute And vbReadOnly) = vbReadOnly And bReadOnly = False Then IsAttrOK = False
    If (lAttribute And vbHidden) = vbHidden And bHidden = False Then IsAttrOK = False

End Function

'=======================================================
'obtient l'attribut sans erreur
'=======================================================
Private Function MustGetAttr(ByVal sFile As String) As Long
        On Error Resume Next
        If fs.FileExists(sFile) Then MustGetAttr = fs.GetFile(sFile).Attributes
        If fs.FolderExists(sFile) Then MustGetAttr = fs.GetFolder(sFile).Attributes
End Function

'=======================================================
'renomme sans erreur
'=======================================================
Private Sub MustRename(ByVal sOldName As String, sNewName As String)
    On Error Resume Next
    
    If ParentFolder(sOldName) <> ParentFolder(sNewName) Then Exit Sub 'pas dans le même dossier
    
    Name sOldName As sNewName
End Sub

'=======================================================
'renvoie la terminaison d'un fichier/path
'=======================================================
Private Function GetFileExtension(ByVal sFileOrPath As String) As String
Dim x As Long

    sFileOrPath = GetFileFromPath(sFileOrPath)
    x = InStrRev(sFileOrPath, ".")
    If x > 0 Then GetFileExtension = Right(sFileOrPath, Len(sFileOrPath) - x)
End Function

'=======================================================
'détermine si le fichier convient au pattern
'=======================================================
Private Function IsPatternOk(ByVal sFile As String) As Boolean
Dim sExt As String

    sExt = GetFileExtension(sFile)  'extension du fichier
    
    If sPattern = "*.*" Or InStr(LCase(sPattern), "*." & LCase(sExt) & "|") Or _
    (Right(LCase(sPattern), Len("*." & LCase(sExt))) = "*." & LCase(sExt)) _
    Then IsPatternOk = True Else IsPatternOk = False
    
End Function

'=======================================================
'obtient la taille d'un fichier, même si celle ci est supérieure à 4Go (long)
'=======================================================
Private Function GetFileSize(ByVal strFile As String) As Currency
Dim lngFile As Long
Dim curSize As Currency

    'obtient le handle du fichier
    lngFile = CreateFile(strFile, &H80000000, &H1, ByVal 0&, 3, ByVal 0&, ByVal 0&)
    
    'obtient la taille par API
    GetFileSizeEx lngFile, curSize
    
    'ferme le handle ouvert
    CloseHandle lngFile
    
    GetFileSize = curSize * 10000 'multiplie par 10^4 pour obtenir un nombre entier
       
End Function

