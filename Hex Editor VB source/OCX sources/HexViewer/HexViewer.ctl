VERSION 5.00
Begin VB.UserControl HexViewer 
   AutoRedraw      =   -1  'True
   ClientHeight    =   5115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6900
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   5115
   ScaleWidth      =   6900
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2040
      Top             =   1680
   End
End
Attribute VB_Name = "HexViewer"
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
'//HEWVIEWER USERCONTROL
'AFFICHE UNE TABLE HEXADECIMALE
'=======================================================

'=======================================================
'//USERCONTROL BY Violent_ken
'//
'//NOTE : THE SUBCLASSING PART OF THIS CODE BELONGS TO PAUL CATON
'//FOR MORE INFORMATIONS ABOUT THE COPYRIGHT, PLEASE REFER TO
'//http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=64867&lngWId=1
'=======================================================

'=======================================================
'CONSTANTES
'=======================================================
Private Const WM_MOUSEWHEEL             As Long = &H20A
Private Const ALL_MESSAGES              As Long = -1         'All messages callback
Private Const MSG_ENTRIES               As Long = 32         'Number of msg table entries
Private Const WNDPROC_OFF               As Long = &H38       'Thunk offset to the WndProc execution address
Private Const GWL_WNDPROC               As Long = -4         'SetWindowsLong WndProc index
Private Const IDX_SHUTDOWN              As Long = 1          'Thunk data index of the shutdown flag
Private Const IDX_HWND                  As Long = 2          'Thunk data index of the subclassed hWnd
Private Const IDX_WNDPROC               As Long = 9          'Thunk data index of the original WndProc
Private Const IDX_BTABLE                As Long = 11         'Thunk data index of the Before table
Private Const IDX_ATABLE                As Long = 12         'Thunk data index of the After table
Private Const IDX_PARM_USER             As Long = 13         'Thunk data index of the User-defined callback parameter data index
Private Const WM_MOUSEMOVE              As Long = &H200
Private Const WM_MOUSELEAVE             As Long = &H2A3
Private Const WM_MOVING                 As Long = &H216
Private Const WM_SIZING                 As Long = &H214
Private Const WM_EXITSIZEMOVE           As Long = &H232

'=======================================================
'APIs
'=======================================================
Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long



'=======================================================
'TYPE & ENUMS
'=======================================================
Private Enum TRACKMOUSEEVENT_FLAGS
  TME_HOVER = &H1&
  TME_LEAVE = &H2&
  TME_QUERY = &H40000000
  TME_CANCEL = &H80000000
End Enum
Private Type TRACKMOUSEEVENT_STRUCT
  cbSize                      As Long
  dwFlags                     As TRACKMOUSEEVENT_FLAGS
  hwndTrack                   As Long
  dwHoverTime                 As Long
End Type
Public Enum GridType
    None = 0
    Horizontal = 1
    HorizontalHexOnly = 2
    VerticalHex = 3
    HorizontalHexOnly_VerticalHex = 4
    Horizontal_VerticalHex = 5
End Enum
Private Enum eMsgWhen                                                       'When to callback
  MSG_BEFORE = 1                                                            'Callback before the original WndProc
  MSG_AFTER = 2                                                             'Callback after the original WndProc
  MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                                'Callback before and after the original WndProc
End Enum
Private Type HexCase
    lOffset As Currency
    lCol As Long
End Type



'=======================================================
'SUBCLASS VARIABLE DECLARATION
'=======================================================
Private z_ScMem             As Long                                         'Thunk base address
Private z_Sc(64)            As Long                                         'Thunk machine-code initialised here
Private z_Funk              As Collection                                   'hWnd/thunk-address collection
Private bTrack                As Boolean
Private bTrackUser32          As Boolean
Private bInCtrl               As Boolean
Private bMoving               As Boolean


'=======================================================
'VARIABLES
'=======================================================
Private bStillOkForRefresh As Boolean
Private lNumberPerPage As Long   'nombre de lignes à afficher
Private lFirstOffset As Currency 'valeur du premier offset
Private Mhex() As String    'contient les valeurs hexa
Private Mstr() As String  'contient les strings
Private M_M() As Boolean   'case modifiée ?
Private xOld As Long    'x old selection
Private xOld2 As Long
Private yOld2 As Long
Private yOld As Long    'y old selection
Private bGrid As GridType    'grid ?
Private M_S() As Currency   'contient les offsets des signets
Private lNumberOfSelectedItems As Currency
'couleurs du controle
Private lBackColor As OLE_COLOR
Private lOffsetForeColor As OLE_COLOR
Private lHexForeColor As OLE_COLOR
Private lStringForeColor As OLE_COLOR
Private lTitleBackGround As OLE_COLOR
Private lOffsetTitleForeColor As OLE_COLOR
Private lBaseTitleForeColor As OLE_COLOR
Private lLineColor As OLE_COLOR
Private lSignetColor As OLE_COLOR
Private lSelectionColor As OLE_COLOR
Private lModifiedItemColor As OLE_COLOR
Private lModifiedSelectedItemColor As OLE_COLOR
Private cit As ItemElement  'current Item ==> renvoyé par la property Get
Private lMaxOffset As Currency
Private hexOldCase As HexCase
Private hexNewCase As HexCase
Private yZone As Single   'contient l'ordonnée du curseur de la souris sur le usercontrol
Private xZone As Single   'idem
Private lSpeed As Long  'vitesse de défilement
Private bUseHexOffset As Boolean    'utilise ou non l'affichage des offsets en hexa
'tags (permettent de stocker des infos dans le HW)
Private cur_Tag1 As Currency
Private cur_Tag2 As Currency
Private str_Tag1 As String
Private str_Tag2 As String
Private bDisableHexDisplay As Boolean
Private curFileSize As Currency     'pour ne pas pourvoir sélectionner les derniers bytes
'du tableau si > taille du fichier

   
   
'=======================================================
'EVENTS
'=======================================================
Public Event ItemClick(Item As ItemElement, Button As Integer)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single, Item As ItemElement)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single, Item As ItemElement)
Public Event Click()
Public Event DblClick()
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event Resize()
Public Event MouseWheel(ByVal lSens As Long)
Public Event UserMakeFirstOffsetChangeByMovingMouse()



'=======================================================
'PROPRIETES
'=======================================================
Public Property Get BackColor() As OLE_COLOR: BackColor = lBackColor: End Property
Public Property Let BackColor(BackColor As OLE_COLOR): lBackColor = BackColor: UserControl.BackColor = BackColor: Refresh: End Property
Public Property Get SignetColor() As OLE_COLOR: SignetColor = lSignetColor: End Property
Public Property Let SignetColor(SignetColor As OLE_COLOR): lSignetColor = SignetColor: Refresh: End Property
Public Property Get Grid() As GridType: Grid = bGrid: End Property
Public Property Let Grid(Grid As GridType): bGrid = Grid: Refresh: End Property
Public Property Get Speed() As Long: Speed = lSpeed: End Property
Public Property Let Speed(Speed As Long): lSpeed = Speed: Timer1.Interval = lSpeed: End Property
Public Property Get OffsetForeColor() As OLE_COLOR: OffsetForeColor = lOffsetForeColor: End Property
Public Property Let OffsetForeColor(OffsetForeColor As OLE_COLOR): lOffsetForeColor = OffsetForeColor: Refresh: End Property
Public Property Get HexForeColor() As OLE_COLOR: HexForeColor = lHexForeColor: End Property
Public Property Let HexForeColor(HexForeColor As OLE_COLOR): lHexForeColor = HexForeColor: Refresh: End Property
Public Property Get NumberOfSelectedItems() As Currency
'détermine le nombre d'items sélectionnés
Dim z As Currency
Dim NewCase As HexCase
Dim OldCase As HexCase

    'détermine les Case TEMPORAIRES de départ et d'arrivée de sélection
    'Old DOIT ETRE inférieur à New au niveau de l'offset
    If hexOldCase.lOffset > hexNewCase.lOffset Then
        OldCase.lCol = hexNewCase.lCol
        OldCase.lOffset = hexNewCase.lOffset
        NewCase.lCol = hexOldCase.lCol
        NewCase.lOffset = hexOldCase.lOffset
    Else
        NewCase.lCol = hexNewCase.lCol
        NewCase.lOffset = hexNewCase.lOffset
        OldCase.lCol = hexOldCase.lCol
        OldCase.lOffset = hexOldCase.lOffset
    End If
    
    z = CCur((NewCase.lOffset - OldCase.lOffset - 16) + (16 - OldCase.lCol) + NewCase.lCol)
    If z < 0 Then z = -z   'évite le bug lors de la sélection sur une même ligne de droite à gauche

    NumberOfSelectedItems = z + 1: lNumberOfSelectedItems = z + 1: End Property
Public Property Get StringForeColor() As OLE_COLOR: StringForeColor = lStringForeColor: End Property
Public Property Let StringForeColor(StringForeColor As OLE_COLOR): lStringForeColor = StringForeColor: Refresh: End Property
Public Property Get TitleBackGround() As OLE_COLOR: TitleBackGround = lTitleBackGround: End Property
Public Property Let TitleBackGround(TitleBackGround As OLE_COLOR): lTitleBackGround = TitleBackGround: Refresh: End Property
Public Property Get OffsetTitleForeColor() As OLE_COLOR: OffsetTitleForeColor = lOffsetTitleForeColor: End Property
Public Property Let OffsetTitleForeColor(OffsetTitleForeColor As OLE_COLOR): lOffsetTitleForeColor = OffsetTitleForeColor: Refresh: End Property
Public Property Get BaseTitleForeColor() As OLE_COLOR: BaseTitleForeColor = lBaseTitleForeColor: End Property
Public Property Let BaseTitleForeColor(BaseTitleForeColor As OLE_COLOR): lBaseTitleForeColor = BaseTitleForeColor: Refresh: End Property
Public Property Get LineColor() As OLE_COLOR: LineColor = lLineColor: End Property
Public Property Let LineColor(LineColor As OLE_COLOR): lLineColor = LineColor: Refresh: End Property
Public Property Get SelectionColor() As OLE_COLOR: SelectionColor = lSelectionColor: End Property
Public Property Let SelectionColor(SelectionColor As OLE_COLOR): lSelectionColor = SelectionColor: Refresh: End Property
Public Property Get ModifiedItemColor() As OLE_COLOR: ModifiedItemColor = lModifiedItemColor: End Property
Public Property Let ModifiedItemColor(ModifiedItemColor As OLE_COLOR): lModifiedItemColor = ModifiedItemColor: Refresh: End Property
Public Property Get ModifiedSelectedItemColor() As OLE_COLOR: ModifiedSelectedItemColor = lModifiedSelectedItemColor: End Property
Public Property Let ModifiedSelectedItemColor(ModifiedSelectedItemColor As OLE_COLOR): lModifiedSelectedItemColor = ModifiedSelectedItemColor: Refresh: End Property
Public Property Get FirstOffset() As Currency: FirstOffset = lFirstOffset: End Property
Public Property Let FirstOffset(FirstOffset As Currency)
'il faut reféfinir cit
cit.Offset = FirstOffset
lFirstOffset = FirstOffset
End Property
Public Property Get NumberPerPage() As Long: NumberPerPage = lNumberPerPage: End Property
Public Property Let NumberPerPage(NumberPerPage As Long): lNumberPerPage = NumberPerPage: ChangeValues: End Property
Public Property Get Item() As ItemElement: Set Item = cit: End Property
Public Property Get MaxOffset() As Currency: MaxOffset = lMaxOffset: End Property
Public Property Let MaxOffset(MaxOffset As Currency): lMaxOffset = MaxOffset: End Property
Public Property Get Value(Line As Long, Col As Long) As String: Value = Mhex(Col, Line): End Property
Public Property Let Item(Item As ItemElement): Set cit = Item: End Property
Public Property Get FirstSelectionItem() As ItemElement
Dim NewCase As HexCase
Dim OldCase As HexCase

    'détermine les Case TEMPORAIRES de départ et d'arrivée de sélection
    'Old DOIT ETRE inférieur à New au niveau de l'offset
    If hexOldCase.lOffset > hexNewCase.lOffset Then
        OldCase.lCol = hexNewCase.lCol
        OldCase.lOffset = hexNewCase.lOffset
        NewCase.lCol = hexOldCase.lCol
        NewCase.lOffset = hexOldCase.lOffset
    Else
        NewCase.lCol = hexNewCase.lCol
        NewCase.lOffset = hexNewCase.lOffset
        OldCase.lCol = hexOldCase.lCol
        OldCase.lOffset = hexOldCase.lOffset
    End If
    
Set FirstSelectionItem = New ItemElement
FirstSelectionItem.Offset = OldCase.lOffset
FirstSelectionItem.Col = OldCase.lCol
End Property
Public Property Get SecondSelectionItem() As ItemElement
Dim NewCase As HexCase
Dim OldCase As HexCase

    'détermine les Case TEMPORAIRES de départ et d'arrivée de sélection
    'Old DOIT ETRE inférieur à New au niveau de l'offset
    If hexOldCase.lOffset > hexNewCase.lOffset Then
        OldCase.lCol = hexNewCase.lCol
        OldCase.lOffset = hexNewCase.lOffset
        NewCase.lCol = hexOldCase.lCol
        NewCase.lOffset = hexOldCase.lOffset
    Else
        NewCase.lCol = hexNewCase.lCol
        NewCase.lOffset = hexNewCase.lOffset
        OldCase.lCol = hexOldCase.lCol
        OldCase.lOffset = hexOldCase.lOffset
    End If
    
Set SecondSelectionItem = New ItemElement
SecondSelectionItem.Offset = NewCase.lOffset
SecondSelectionItem.Col = NewCase.lCol
End Property
Public Property Get UseHexOffset() As Boolean: UseHexOffset = bUseHexOffset: End Property
Public Property Let UseHexOffset(UseHexOffset As Boolean): bUseHexOffset = UseHexOffset: End Property
Public Property Get curTag1() As Currency: curTag1 = cur_Tag1: End Property
Public Property Let curTag1(curTag1 As Currency): cur_Tag1 = curTag1: End Property
Public Property Get curTag2() As Currency: curTag2 = cur_Tag2: End Property
Public Property Let curTag2(curTag2 As Currency): cur_Tag2 = curTag2: End Property
Public Property Get strTag1() As String: strTag1 = str_Tag1: End Property
Public Property Let strTag1(strTag1 As String): str_Tag1 = strTag1: End Property
Public Property Get strTag2() As String: strTag2 = str_Tag2: End Property
Public Property Let strTag2(strTag2 As String): str_Tag2 = strTag2: End Property
Public Property Get DisableHexDisplay() As Boolean: DisableHexDisplay = bDisableHexDisplay: End Property
Public Property Let DisableHexDisplay(DisableHexDisplay As Boolean): bDisableHexDisplay = DisableHexDisplay: End Property
Public Property Get NumberOfSignets() As Long: NumberOfSignets = UBound(M_S()) - 1: End Property
Public Property Let FileSize(FileSize As Currency): curFileSize = FileSize: End Property


'=======================================================
'EVENEMENTS SIMPLES
'=======================================================
Private Sub UserControl_Click()
    RaiseEvent Click
End Sub
Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub



'=======================================================
'changement de FirstOffset lorsque la souris est tout en haut
'ou tout en bas du userontrol
'=======================================================
Private Sub Timer1_Timer()
    If yZone > UserControl.Height - 150 Then
        'alors on descend
        Me.FirstOffset = IIf(Me.FirstOffset + 16 < Me.MaxOffset, Me.FirstOffset + 16, Me.FirstOffset)
        Me.Refresh
        RaiseEvent UserMakeFirstOffsetChangeByMovingMouse
    ElseIf yZone < 300 Then
        'alors on monte
        Me.FirstOffset = IIf(Me.FirstOffset - 16 > 0, Me.FirstOffset - 16, 0)
        Me.Refresh
        RaiseEvent UserMakeFirstOffsetChangeByMovingMouse
    End If

End Sub



'=======================================================
'EVENEMENTS NON SIMPLES
'=======================================================
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim it As New ItemElement
Dim z As Long, z2 As Long
Dim yCase As Currency
Dim xCase As Long

    
          
    If Button <> 1 Then GoTo RaiseEv
    
    
    'annule toute la sélection
   ' For z = 1 To 16
   '     For z2 = 1 To Me.NumberPerPage
   '         Msel(z, z2) = False
   '     Next z2
   ' Next z
 
        
    'détermine colonne et ligne en fonction de x et y
    yCase = Round((Y - 180) / 260, 0) 'y en coordonnée de matrice
    
    If (x > 7500 And x < 9450) Then
        'alors c'est une zone de strings
        xCase = Int((x - 7515) / 120) + 1
    Else
        'alors c'est une zone de valeurs hexa
        xCase = Round((x - 1250) / 360, 0)
    End If
    
    'détermine si oui ou non la nouvelle position de la souris est
    'EN DEHORS DU FICHIER, c'est à dire si la zone sélectionnée est sur des bytes
    'existants ou non
    If (Me.FirstOffset + 16 * (yCase - 1) + xCase) > curFileSize Then Exit Sub 'dépasse du fichier
          
    
    'stocke dans la variable contenant la nouvelle case sélectionnée
    hexNewCase.lCol = xCase
    hexNewCase.lOffset = Me.FirstOffset + 16 * (yCase - 1)
    
    '//ZONE OFFSET
    If x <= 1300 And Y > 310 And yCase <= NumberPerPage Then
    'alors c'est un offset

        If Shift = 1 Then
            'maintenu la touche Shit, alors sélection multiple
            Call UserControl_MouseMove(Button, Shift, x, Y)
        End If
        
        If Shift <> 1 Then
            'xOld = xCase: yOld = yCase 'nouvelle case, car pas de Shift appuyé
            hexOldCase.lCol = hexNewCase.lCol
            hexOldCase.lOffset = hexNewCase.lOffset
        End If
        
        it.Col = 1
        it.Line = yCase
        it.tType = tOffset
        it.Value = lFirstOffset + (it.Line - 1) * 16
        
        'colore l'offset, les valeurs hex, et la string
        ColorItem tOffset, it.Line, 1, it.Value, lSelectionColor, False
        
        'Items hexa=selectionnés
       ' For z = 1 To 16
       '     Msel(z, it.Line) = True
       ' Next z
    End If
    
    
    '//ZONE HEXA
    If xCase >= 1 And Y > 310 And yCase <= NumberPerPage And xCase <= 16 Then
        'alors c'est une valeur hexa
        
        If Shift = 1 Then
            'maintenu la touche Shit, alors sélection multiple
            Call UserControl_MouseMove(Button, Shift, x, Y)
        End If
        
        it.Col = xCase
        it.Line = yCase
        it.tType = tHex
        it.Value = Mhex(xCase, yCase)
        
        If Shift <> 1 Then
            'xOld = xCase: yOld = yCase 'nouvelle case, car pas de Shift appuyé
            hexOldCase.lCol = hexNewCase.lCol
            hexOldCase.lOffset = hexNewCase.lOffset
        End If
        
        'colore l'offset, les valeurs hex, et la string
        ColorItem tHex, it.Line, xCase, it.Value, lSelectionColor, False
        
        'Items hexa=selectionnés
        'Msel(it.Col, it.Line) = True
    End If
    
    '//ZONE STRING
    If x > 7500 And Y > 310 And yCase <= NumberPerPage And x < 9450 Then
        'alors c'est une valeur string
        
        If Shift = 1 Then
            'maintenu la touche Shit, alors sélection multiple
            Call UserControl_MouseMove(Button, Shift, x, Y)
        End If
        
        'redéfinit xCase
        xCase = Int((x - 7515) / 120) + 1
        
        hexNewCase.lCol = xCase
        
        it.Col = xCase
        it.Line = yCase
        it.tType = tString
        it.Value = Mid$(Mstr(yCase), xCase, 1)
        
        If Shift <> 1 Then
            'xOld = xCase: yOld = yCase 'nouvelle case, car pas de Shift appuyé
            hexOldCase.lCol = hexNewCase.lCol
            hexOldCase.lOffset = hexNewCase.lOffset
        End If
        
        'colore l'offset, les valeurs hex, et la string
        ColorItem tHex, it.Line, xCase, it.Value, lSelectionColor, False

    End If

    RaiseEvent ItemClick(it, Button)
    RaiseEvent MouseDown(Button, Shift, x, Y, it)
    
    Set cit = it
    cit.Offset = (cit.Line - 1) * 16 + Me.FirstOffset
    
    Refresh

    Exit Sub
    
RaiseEv:
    
    'on définit ici it (car pas passé par les boucles ou çà a été défini)
    yCase = Round((Y - 180) / 260, 0) 'y en coordonnée de matrice
    xCase = Round((x - 1250) / 360, 0)  'idem pour x
        
    If x <= 1300 And Y > 310 And yCase <= NumberPerPage Then
        'alors c'est un offset
        it.Col = 1
        it.Line = yCase
        it.tType = tOffset
        it.Value = lFirstOffset + (it.Line - 1) * 16
    End If
    If xCase >= 1 And Y > 310 And yCase <= NumberPerPage And xCase <= 16 Then
        'alors c'est une valeur hexa
        it.Col = xCase
        it.Line = yCase
        it.tType = tHex
        it.Value = Mhex(xCase, yCase)
    End If
    If x > 7500 And Y > 310 And yCase <= NumberPerPage And x < 9450 Then
        'alors c'est une valeur string
        'redéfinit xCase
        xCase = Int((x - 7515) / 120) + 1
        it.Col = IIf(xCase < 16, xCase, 16)
        it.Line = yCase
        it.tType = tString
        it.Value = Mid$(Mstr(yCase), xCase, 1)
    End If
    
    Set cit = it
    cit.Offset = (cit.Line - 1) * 16 + Me.FirstOffset
    
    RaiseEvent MouseDown(Button, Shift, x, Y, it)
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'click sur le picturebox
Dim it As New ItemElement
Dim z As Long
Dim w As Long
Dim yCase As Currency
Dim xCase As Long
Dim lStep As Long


    yCase = Round((Y - 180) / 260, 0) 'y en coordonnée de matrice
    
    If (x > 7500 And x < 9450) Then
        'alors c'est une zone de strings
        xCase = Int((x - 7515) / 120) + 1
    Else
        'alors c'est une zone de valeurs hexa
        xCase = Round((x - 1250) / 360, 0)
    End If
    
    'détermine si oui ou non la nouvelle position de la souris est
    'EN DEHORS DU FICHIER, c'est à dire si la zone sélectionnée est sur des bytes
    'existants ou non
    If (Me.FirstOffset + 16 * (yCase - 1) + xCase) > curFileSize Then Exit Sub 'dépasse du fichier
    
    
    '//détermine si il faut ou non activer le changement d'offset
    'automatique lorsqu'on est vers les bords du controle
    yZone = Y: xZone = x
    If (yZone > UserControl.Height - 150) Or (yZone < 300) Then
        'alors on est dans la zone inférieure du controle
        'lance le timer pour le défilement
        Timer1.Enabled = (Button = 1)
    ElseIf Timer1.Enabled Then
        Timer1.Enabled = False
    End If
    
    
    If Button <> 1 Then GoTo RaiseEv
    
    If Shift <> 1 Then
        'xOld = xCase: yOld = yCase 'nouvelle case, car pas de Shift appuyé
        hexNewCase.lCol = xCase
        hexNewCase.lOffset = Me.FirstOffset + 16 * (yCase - 1)
    End If

    If x > 7500 Then
        'alors on est sur la zone string, il faut redéfinir xCase
        If Int((x - 7515) / 120) + 1 = xOld2 And yCase = yOld2 Then
            'alors inutile de rafraichir ==> on ne s'est pas déplacé de case
            RaiseEvent MouseMove(Button, Shift, x, Y, it)
            Exit Sub
        End If
    ElseIf xCase = xOld2 And yCase = yOld2 Then
        'alors inutile de rafraichir ==> on ne s'est pas déplacé de case
        RaiseEvent MouseMove(Button, Shift, x, Y, it)
        Exit Sub
    End If
    
    xOld2 = xCase: yOld2 = yCase
    
    
    'UserControl.Cls
    'UserControl.Picture = UserControl.MaskPicture
    Refresh2

    
    '//ZONE OFFSET
    If x <= 1300 And Y > 310 And yCase <= NumberPerPage Then
        'alors c'est un offset
        
        'évite les dépassements (ex xold=-1) en cas de problème
        '//UTILE - NE PAS ENLEVER
        If xOld < 1 Then xOld = 1
        If xOld > 16 Then xOld = 16
        If yOld < 1 Then yOld = 1
        If yOld > NumberPerPage Then yOld = NumberPerPage
        
        If hexOldCase.lOffset < Me.FirstOffset Then
            'alors la sélection provient d'avant ce qui est visible
            For z = 1 To yCase
                ColorItem tOffset, z, 1, FirstOffset + 16 * (z - 1), lSelectionColor, False
            Next z
        ElseIf Me.FirstOffset <= hexOldCase.lOffset And hexOldCase.lOffset <= (Me.FirstOffset + 16 * Me.NumberPerPage) Then
            'alors la sélection provient de la même page
            yOld = (hexOldCase.lOffset - Me.FirstOffset) / 16 + 1
            
            If yOld < 1 Then yOld = 1
            If yOld > NumberPerPage Then yOld = NumberPerPage
        
            lStep = IIf(yOld > yCase, -1, 1)
            For z = yOld To yCase Step lStep
                ColorItem tOffset, z, 1, FirstOffset + 16 * (z - 1), lSelectionColor, False
            Next z
        ElseIf hexOldCase.lOffset > (Me.FirstOffset + 16 * Me.NumberPerPage) Then
            'alors la sélection provient d'une page plus basse
            For z = Me.NumberPerPage To yCase Step -1
                ColorItem tOffset, z, 1, FirstOffset + 16 * (z - 1), lSelectionColor, False
            Next z
        End If
            
        
        it.Col = 1
        it.Line = yCase
        it.tType = tOffset
        it.Value = lFirstOffset + (it.Line - 1) * 16

    End If
    
    
    '//ZONE HEXA
    If xCase >= 1 And Y > 310 And yCase <= NumberPerPage And xCase <= 16 And yCase >= 1 Then
        'alors c'est une valeur hexa

        it.Col = xCase
        it.Line = yCase
        it.tType = tOffset
        it.Value = Mhex(xCase, yCase)
        
        'xOld et yOld sont les anciennes valeurs
        'remplit une sélection en fonction de xOld, xCase, yOld et yCase

        If hexOldCase.lOffset < Me.FirstOffset Then
            'alors la sélection provient d'avant ce qui est visible
            yOld = 1
            xOld = 1
        ElseIf Me.FirstOffset <= hexOldCase.lOffset And hexOldCase.lOffset <= (Me.FirstOffset + 16 * Me.NumberPerPage) Then
            'alors la sélection provient de la même page
            yOld = (hexOldCase.lOffset - Me.FirstOffset) / 16 + 1
            xOld = hexOldCase.lCol
        ElseIf hexOldCase.lOffset > (Me.FirstOffset + 16 * Me.NumberPerPage) Then
            'alors la sélection provient d'une page plus basse
            yOld = Me.NumberPerPage
            xOld = 16
        End If
        
        'évite les dépassements (ex xold=-1) en cas de problème
        '//UTILE - NE PAS ENLEVER
        If xOld < 1 Then xOld = 1
        If xOld > 16 Then xOld = 16
        If yOld < 1 Then yOld = 1
        If yOld > NumberPerPage Then yOld = NumberPerPage
        
        If yCase = yOld Then
            'remplit sur l'horizontale
            lStep = IIf(xCase > xOld, -1, 1)
            For z = xCase To xOld Step lStep
                'colore l'offset, les valeurs hex, et la string
                ColorItem tHex, yOld, z, Mhex(z, yOld), lSelectionColor, False
                'Items hexa=selectionnés
               ' Msel(z, yOld) = True
            Next z
        End If
        
        If yCase > yOld Then
            'finit première la ligne
            For z = xOld To 16
                'colore l'offset, les valeurs hex, et la string
                ColorItem tHex, yOld, z, Mhex(z, yOld), lSelectionColor, False
                'Items hexa=selectionnés
               ' Msel(z, yOld) = True
            Next z
            'finit la dernière ligne
            For z = xCase To 1 Step -1
                'colore l'offset, les valeurs hex, et la string
                ColorItem tHex, yCase, z, Mhex(z, yCase), lSelectionColor, False
                'Items hexa=selectionnés
               ' Msel(z, yCase) = True
            Next z
            'fait les lignes entre si nécessaire
            If (yCase - yOld) > 1 Then
                'il y a des lignes entre : les fait
                'remplit par LIGNE (pas par élément) ==> gagne du temps
                For z = yCase - 1 To yOld + 1 Step -1
                    ColorLine z, lSelectionColor, False
                   ' For w = 1 To 16
                   '     Msel(w, z) = True
                   ' Next w
                Next z
            End If
        End If
        
        If yCase < yOld Then
            For z = xCase To 16
                'colore l'offset, les valeurs hex, et la string
                ColorItem tHex, yCase, z, Mhex(z, yCase), lSelectionColor, False
                'Items hexa=selectionnés
               ' Msel(z, yCase) = True
            Next z
            For z = xOld To 1 Step -1
                'colore l'offset, les valeurs hex, et la string
                ColorItem tHex, yOld, z, Mhex(z, yOld), lSelectionColor, False
                'Items hexa=selectionnés
               ' Msel(z, yOld) = True
            Next z
            'fait les lignes entre si nécessaire
            If (yOld - yCase) > 1 Then
                'il y a des lignes entre : les fait
                'remplit par LIGNE (pas par élément) ==> gagne du temps
                For z = yOld - 1 To yCase + 1 Step -1
                    ColorLine z, lSelectionColor, False
                    'For w = 1 To 16
                   '     Msel(w, z) = True
                    'Next w
                Next z
            End If
        End If
    End If
    
    '//ZONE STRING
    If x > 7500 And Y > 310 And yCase <= NumberPerPage And x < 9450 Then
        'alors c'est une valeur string
        
        'redéfinit xCase
        xCase = Int((x - 7515) / 120) + 1
        
        If hexOldCase.lOffset < Me.FirstOffset Then
            'alors la sélection provient d'avant ce qui est visible
            yOld = 1
            xOld = 1
        ElseIf Me.FirstOffset <= hexOldCase.lOffset And hexOldCase.lOffset <= (Me.FirstOffset + 16 * Me.NumberPerPage) Then
            'alors la sélection provient de la même page
            yOld = (hexOldCase.lOffset - Me.FirstOffset) / 16 + 1
            xOld = hexOldCase.lCol
        ElseIf hexOldCase.lOffset > (Me.FirstOffset + 16 * Me.NumberPerPage) Then
            'alors la sélection provient d'une page plus basse
            yOld = Me.NumberPerPage
            xOld = 16
        End If
        
        'évite les dépassements (ex xold=-1) en cas de problème
        '//UTILE - NE PAS ENLEVER
        If xOld < 1 Then xOld = 1
        If xOld > 16 Then xOld = 16
        If yOld < 1 Then yOld = 1
        If yOld > NumberPerPage Then yOld = NumberPerPage
        
        it.Col = xCase
        it.Line = yCase
        it.tType = tString
        it.Value = Mid$(Mstr(yCase), xCase, 1)
        
        If xCase > 16 Then xCase = 16
        
        If yCase = yOld Then
            'remplit sur l'horizontale
            lStep = IIf(xCase > xOld, -1, 1)
            For z = xCase To xOld Step lStep
                'colore l'offset, les valeurs hex, et la string
                ColorItem tHex, yOld, z, Mhex(z, yOld), lSelectionColor, False
                'Items hexa=selectionnés
               ' Msel(z, yOld) = True
            Next z
        End If
        
        If yCase > yOld Then
            'finit première la ligne
            For z = xOld To 16
                'colore l'offset, les valeurs hex, et la string
                ColorItem tHex, yOld, z, Mhex(z, yOld), lSelectionColor, False
                'Items hexa=selectionnés
             '   Msel(z, yOld) = True
            Next z
            'finit la dernière ligne
            For z = xCase To 1 Step -1
                'colore l'offset, les valeurs hex, et la string
                ColorItem tHex, yCase, z, Mhex(z, yCase), lSelectionColor, False
                'Items hexa=selectionnés
               ' Msel(z, yCase) = True
            Next z
            'fait les lignes entre si nécessaire
            If (yCase - yOld) > 1 Then
                For z = yCase - 1 To yOld + 1 Step -1
                    ColorLine z, lSelectionColor, False
                  '  For w = 1 To 16
                  '      Msel(w, z) = True
                   ' Next w
                Next z
            End If
        End If
        
        If yCase < yOld Then
            For z = xCase To 16
                'colore l'offset, les valeurs hex, et la string
                ColorItem tHex, yCase, z, Mhex(z, yCase), lSelectionColor, False
                'Items hexa=selectionnés
                'Msel(z, yCase) = True
            Next z
            For z = xOld To 1 Step -1
                'colore l'offset, les valeurs hex, et la string
                ColorItem tHex, yOld, z, Mhex(z, yOld), lSelectionColor, False
                'Items hexa=selectionnés
                'Msel(z, yOld) = True
            Next z
            'fait les lignes entre si nécessaire
            If (yOld - yCase) > 1 Then
                For z = yOld - 1 To yCase + 1 Step -1
                    ColorLine z, lSelectionColor, False
                 '   For w = 1 To 16
                  '      Msel(w, z) = True
                  '  Next w
                Next z
            End If
        End If
    End If
    
    
    RaiseEvent ItemClick(it, Button)
    RaiseEvent MouseMove(Button, Shift, x, Y, it)
        
    TraceGrid
    TraceSignets
    
    Exit Sub
    
RaiseEv:
    
    'on définit ici it (car pas passé par les boucles on çà a été défini)
    yCase = Round((Y - 180) / 260, 0) 'y en coordonnée de matrice
    xCase = Round((x - 1250) / 360, 0)  'idem pour x
    If x <= 1300 And Y > 310 And yCase <= NumberPerPage Then
        'alors c'est un offset
        it.Col = 1
        it.Line = yCase
        it.tType = tOffset
        it.Value = lFirstOffset + (it.Line - 1) * 16
    End If
    If xCase >= 1 And Y > 310 And yCase <= NumberPerPage And xCase <= 16 Then
        'alors c'est une valeur hexa
        it.Col = xCase
        it.Line = yCase
        it.tType = tHex
        it.Value = Mhex(xCase, yCase)
    End If
    If x > 7500 And Y > 310 And yCase <= NumberPerPage And x < 9450 Then
        'alors c'est une valeur string
        'redéfinit xCase
        xCase = Int((x - 7515) / 120) + 1
        it.Col = IIf(xCase < 16, xCase, 16)
        it.Line = yCase
        it.tType = tString
        it.Value = Mid$(Mstr(yCase), xCase, 1)
    End If
    
    RaiseEvent MouseMove(Button, Shift, x, Y, it)
    
    
End Sub



'=======================================================
'USERCONTROL SUBS
'=======================================================
Private Sub UserControl_InitProperties()
'initialise les variables par défault (création du controle)
    Me.BackColor = vbWhite
    Me.OffsetForeColor = 16737380
    Me.HexForeColor = &H6F6F6F
    Me.StringForeColor = &H6F6F6F
    Me.TitleBackGround = &H8000000F
    Me.OffsetTitleForeColor = 16737380
    Me.BaseTitleForeColor = 16737380
    Me.LineColor = &H8000000C
    Me.FirstOffset = 0
    Me.NumberPerPage = 20
    Me.SelectionColor = &HE0E0E0
    Me.Grid = None
    xOld = 0: yOld = 0
    Me.SignetColor = &H8080FF
    Me.ModifiedItemColor = vbRed
    Me.ModifiedSelectedItemColor = vbRed
    Me.MaxOffset = 4096
    Me.Speed = 1
    Me.UseHexOffset = False
    Me.curTag1 = 0
    Me.curTag2 = 0
    Me.strTag1 = vbNullString
    Me.strTag2 = vbNullString
    Me.DisableHexDisplay = False
End Sub

Private Sub UserControl_Show()
    'alors c'est bon, on rafraichit
    'ceci évite de rafraichir 50 fois pour rien au loading

    If bStillOkForRefresh Then
        bStillOkForRefresh = False  'on ne rafraichira plus à l'entrée au focus
        'Refresh
    End If
End Sub

Private Sub UserControl_Terminate()
'enleve le hook
    sc_Terminate
End Sub
Private Sub UserControl_Initialize()
'initialisation du controle

    Set cit = New ItemElement
    
    ReDim M_S(1)
    
    ChangeValues
    
    bStillOkForRefresh = True   'alors on est prêt à attendre l'entrée en focus pour pouvoir refresh
    
End Sub
Private Sub UserControl_Resize()
'resize le composant Picturebox

    On Error Resume Next
    
    'pct.Height = UserControl.Height
    'pct.Width = UserControl.Width
    'pct.Left = 0
    'pct.Top = 0

    'détermine le nombre de lignes à afficher
    'lNumberPerPage = IIf(Me.NumberPerPage < (Int(pct.Height / 250) - 1), Me.NumberPerPage, Int(pct.Height / 250) - 1)
    
    ChangeValues    'redimensionne/redessine
    RaiseEvent Resize
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", Me.BackColor, vbWhite)
    Call PropBag.WriteProperty("SelectionColor", Me.SelectionColor, &HE0E0E0)
    Call PropBag.WriteProperty("OffsetForeColor", Me.OffsetForeColor, 16737380)
    Call PropBag.WriteProperty("HexForeColor", Me.HexForeColor, &H6F6F6F)
    Call PropBag.WriteProperty("StringForeColor", Me.StringForeColor, &H6F6F6F)
    Call PropBag.WriteProperty("TitleBackGround", Me.TitleBackGround, &H8000000F)
    Call PropBag.WriteProperty("OffsetTitleForeColor", Me.OffsetTitleForeColor, 16737380)
    Call PropBag.WriteProperty("BaseTitleForeColor", Me.BaseTitleForeColor, 16737380)
    Call PropBag.WriteProperty("LineColor", Me.LineColor, &H8000000C)
    Call PropBag.WriteProperty("SignetColor", Me.SignetColor, &H8080FF)
    Call PropBag.WriteProperty("FirstOffset", Me.FirstOffset, 0)
    Call PropBag.WriteProperty("NumberPerPage", Me.NumberPerPage, 20)
    Call PropBag.WriteProperty("Grid", Me.Grid, None)
    Call PropBag.WriteProperty("ModifiedItemColor", Me.ModifiedItemColor, vbRed)
    Call PropBag.WriteProperty("ModifiedSelectedItemColor", Me.ModifiedSelectedItemColor, vbRed)
    Call PropBag.WriteProperty("MaxOffset", Me.MaxOffset, 4096)
    Call PropBag.WriteProperty("Speed", Me.Speed, 1)
    Call PropBag.WriteProperty("UseHexOffset", Me.UseHexOffset, False)
    Call PropBag.WriteProperty("curTag1", Me.curTag1, 0)
    Call PropBag.WriteProperty("curTag2", Me.curTag2, 0)
    Call PropBag.WriteProperty("strTag1", Me.strTag1, vbNullString)
    Call PropBag.WriteProperty("strTag2", Me.strTag2, vbNullString)
    Call PropBag.WriteProperty("DisableHexDisplay", Me.DisableHexDisplay, False)
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Me.Speed = PropBag.ReadProperty("Speed", 1)
    Me.MaxOffset = PropBag.ReadProperty("MaxOffset", 4096)
    Me.BackColor = PropBag.ReadProperty("BackColor", vbWhite)
    Me.OffsetForeColor = PropBag.ReadProperty("OffsetForeColor", 16737380)
    Me.HexForeColor = PropBag.ReadProperty("HexForeColor", &H6F6F6F)
    Me.StringForeColor = PropBag.ReadProperty("StringForeColor", &H6F6F6F)
    Me.TitleBackGround = PropBag.ReadProperty("TitleBackGround", &H8000000F)
    Me.OffsetTitleForeColor = PropBag.ReadProperty("OffsetTitleForeColor", 16737380)
    Me.BaseTitleForeColor = PropBag.ReadProperty("BaseTitleForeColor", 16737380)
    Me.LineColor = PropBag.ReadProperty("LineColor", &H8000000C)
    Me.SignetColor = PropBag.ReadProperty("SignetColor", &H8080FF)
    Me.FirstOffset = PropBag.ReadProperty("FirstOffset", 0)
    Me.NumberPerPage = PropBag.ReadProperty("NumberPerPage", 20)
    Me.SelectionColor = PropBag.ReadProperty("SelectionColor", &HE0E0E0)
    Me.Grid = PropBag.ReadProperty("Grid", None)
    Me.ModifiedItemColor = PropBag.ReadProperty("ModifiedItemColor", vbRed)
    Me.ModifiedSelectedItemColor = PropBag.ReadProperty("ModifiedSelectedItemColor", vbRed)
    Me.UseHexOffset = PropBag.ReadProperty("UseHexOffset", False)
    Me.curTag1 = PropBag.ReadProperty("curTag1", 0)
    Me.curTag2 = PropBag.ReadProperty("curTag2", 0)
    Me.strTag1 = PropBag.ReadProperty("strTag1", 0)
    Me.strTag2 = PropBag.ReadProperty("strTag2", 0)
    Me.DisableHexDisplay = PropBag.ReadProperty("DisableHexDisplay", False)
    
    
    'c'est la bonne place pour commencer à subclasser
      If Ambient.UserMode Then  'If we're not in design mode
          bTrack = True
          bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")

          If Not bTrackUser32 Then
              If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then bTrack = False
          End If
      
      If bTrack Then
        'OS supports mouse leave, so let's subclass for it
      With UserControl
        'Subclass the UserControl
        sc_Subclass .hWnd
        sc_AddMsg .hWnd, WM_MOUSEMOVE
        sc_AddMsg .hWnd, WM_MOUSELEAVE
        sc_AddMsg .hWnd, &H20A
      End With
    End If
    
    'Subclass the parent form
    With UserControl.Parent
      sc_Subclass .hWnd
      sc_AddMsg .hWnd, WM_MOVING
      sc_AddMsg .hWnd, WM_SIZING
      sc_AddMsg .hWnd, WM_EXITSIZEMOVE
    End With
    End If
  
End Sub


'=======================================================
'PUBLIC PROCEDURES AND FUNCTIONS
'=======================================================

'=======================================================
'renvoie si l'élément est sélectionné
'=======================================================
Public Function IsSelected(ByVal lLine As Long, ByVal Col As Long) As Boolean
Dim TempOff As Long
Dim NewCase As HexCase
Dim OldCase As HexCase

    'détermine les Case TEMPORAIRES de départ et d'arrivée de sélection
    'Old DOIT ETRE inférieur à New au niveau de l'offset
    If hexOldCase.lOffset > hexNewCase.lOffset Then
        OldCase.lCol = hexNewCase.lCol
        OldCase.lOffset = hexNewCase.lOffset
        NewCase.lCol = hexOldCase.lCol
        NewCase.lOffset = hexOldCase.lOffset
    Else
        NewCase.lCol = hexNewCase.lCol
        NewCase.lOffset = hexNewCase.lOffset
        OldCase.lCol = hexOldCase.lCol
        OldCase.lOffset = hexOldCase.lOffset
    End If
    
    'détermine l'offset de la Line
    TempOff = Me.FirstOffset + 16 * (lLine - 1)
    
    If TempOff < NewCase.lOffset And TempOff > OldCase.lOffset Then
        IsSelected = True
        Exit Function
    ElseIf TempOff < OldCase.lOffset Or TempOff > NewCase.lOffset Then
        IsSelected = False
        Exit Function
    ElseIf TempOff = OldCase.lOffset Then
        IsSelected = Col >= OldCase.lCol
        Exit Function
    Else    'tempoff=newcase.loffset
        IsSelected = Col <= NewCase.lCol
    End If
    
End Function

'=======================================================
'ajoute un offset en signet
'=======================================================
Public Sub AddSignet(ByVal lOffset As Currency)

    ReDim Preserve M_S(UBound(M_S) + 1)
    M_S(UBound(M_S)) = lOffset
    
End Sub

'=======================================================
'enlève un offset des signets
'=======================================================
Public Sub RemoveSignet(ByVal lOffset As Currency)
Dim x As Long
Dim b As Boolean

    b = False

    'on augmente la taille de la liste de 1, pour permettre d'accéder à toutes les valeurs
    'et de dépasser sur la dernière (car M_S(dernière vraie valeur)=M_S(valeur inutile en plus))
    ReDim Preserve M_S(UBound(M_S) + 1)
    
    For x = 1 To UBound(M_S) - 1
        If M_S(x) = lOffset Then
            'alors c'est ce signet là qu'il faut virer
            'on va décaler toute la liste des signet (en effçant celui là)
            'puis on va raccourcir la liste des signets de 2
            '2=signet enlevé + valeur ajoutée au début
            b = True
        End If
        If b Then M_S(x) = M_S(x + 1)   'le dernier x (=ubound(m_s)+1) n'importe pas ==> il est supprimé ensuite
    Next x
    
    ReDim Preserve M_S(UBound(M_S) - 2)
    
    'refresh l'affichage des signets
    Refresh

End Sub

'=======================================================
'renvoie True si l'offset lOffset a un signet
'=======================================================
Public Function IsSignet(ByVal lOffset As Currency) As Boolean
Dim x As Long

    IsSignet = False
    
    For x = 2 To UBound(M_S)
        If M_S(x) = lOffset Then
            IsSignet = True
            Exit For
        End If
    Next x
End Function

'=======================================================
'enlève tout les signets
'=======================================================
Public Sub RemoveAllSignets()
    ReDim M_S(1)
    Me.Refresh
End Sub

'=======================================================
'renvoie l'offset suivant
'=======================================================
Public Function GetNextSignet(ByVal lOffset As Currency) As Currency
Dim lAfter As Currency
Dim lMin As Currency
Dim x As Long

    On Error Resume Next    'ici demeure un bug de dépassement de capacité après compilation
    'paix à son âme -_-
    
    lMin = lOffset
    lAfter = Int(Me.MaxOffset / 16) * 16 'offset max possible

    For x = 2 To UBound(M_S)
        If M_S(x) < lMin Then lMin = M_S(x) 'redéfinit le minimum
        If M_S(x) > lOffset And M_S(x) < lAfter Then lAfter = M_S(x) 'redéfinit l'offset suivant
    Next x
    
    If lAfter = Int(Me.MaxOffset / 16) * 16 Then
        'alors on est au dernier offset ==> on remet le premier
        GetNextSignet = lMin
    Else
        GetNextSignet = lAfter
    End If
    
End Function

'=======================================================
'renvoie le signet précédent
'=======================================================
Public Function GetPrevSignet(ByVal lOffset As Currency) As Currency
Dim lBefore As Currency
Dim lMax As Currency
Dim x As Long
    
    On Error Resume Next    'ici demeure un bug de dépassement de capacité après compilation
    'paix à son âme -_-

    lMax = lOffset
    lBefore = 0 'offset min possible
    For x = 2 To UBound(M_S)
        If M_S(x) > lMax Then lMax = M_S(x) 'redéfinit le maximum
        If M_S(x) < lOffset And M_S(x) > lBefore Then lBefore = M_S(x) 'redéfinit l'offset suivant
    Next x
    If lBefore = 0 Then
        'alors on est au premier offset ==> on remet le dernier
        GetPrevSignet = lMax
    Else
        GetPrevSignet = lBefore
    End If
End Function

'=======================================================
'rafraichit le controle entièrement (retrace les valeurs hexa et les lignes du tableau)
'=======================================================
Public Sub Refresh()
Dim x As Long
Dim Y As Long

    If bStillOkForRefresh = True Then Exit Sub  'contrôle pas encore chargé
    
    UserControl.Picture = LoadPicture() 'efface le contenu du controle (pas correct avec Cls)
    
    'alors on efface les anciennes sélections
    CreateBackGround
        
    'on trace les nouvelles sélections
    CreateSelections Me.Item.tType
    
    'remplit le texte
    FillText

    'affiche les signets
    TraceSignets
    
    'sauvegarde la maskpicture
    UserControl.MaskPicture = UserControl.Image
End Sub

'=======================================================
'colorise un élément
'1) applique un rectangle de couleur
'2) réaffiche le texte
'=======================================================
Public Sub ColorItem(ByVal tType As ItemType, ByVal lLine As Long, ByVal lCol As Long, ByVal vValue As Variant, ByVal lColor As Long, ByVal bEraseOtherSelection As Boolean, Optional ByVal bFillOffsetText As Boolean = True)

Dim x As Long
Dim Y As Long

    If bEraseOtherSelection Then
        'alors on efface les anciennes sélections
        Refresh
    End If

    Select Case tType
        Case tOffset
        
            'colorise un offset
            UserControl.Line (20, 260 * lLine + 100)-(1250, 260 * (lLine + 1) + 50), lColor, BF
            'reecrit le texte
            PasteOffset lLine, FormatedAdress(vValue, bUseHexOffset)
            
            'colorise toutes les valeurs hexa
            For x = 1 To 16
                'rectangle
                UserControl.Line (360 * x + 1100, 260 * lLine + 100)-(360 * (x + 1) + 1100, 260 * (lLine + 1) + 50), lColor, BF
                'reecrit le texte
                PasteHex lLine, x, Mhex(x, lLine)
            Next x
            
            'colorise la string
            UserControl.Line (7400, 260 * lLine + 100)-(9500, 260 * (lLine + 1) + 50), lColor, BF
            'reecrit la string
            PasteString lLine, Mstr(lLine)
            
            
        Case tHex
        
            If lCol > 16 Then lCol = 16

            'colorise une case hexa
            UserControl.Line (360 * lCol + 1050, 260 * lLine + 100)-(360 * (lCol + 1) + 1050, 260 * (lLine + 1) + 50), lColor, BF
            'reecrit le texte
            If bFillOffsetText Then PasteHex lLine, lCol, Mhex(lCol, lLine)
            'colorise une partie de la string
            UserControl.Line (7400 + 120 * (lCol), 260 * lLine + 100)-(7400 + 120 * (lCol + 1), 260 * (lLine + 1) + 50), lColor, BF
            'réécrit la string
            If bFillOffsetText Then PasteString lLine, Mstr(lLine)

    End Select
    
End Sub

'=======================================================
'colorise un élément
'1) applique un rectangle de couleur
'2) réaffiche le texte
'=======================================================
Private Sub ColorLine(ByVal lLine As Long, ByVal lColor As Long, ByVal bEraseOtherSelection As Boolean)
Dim x As Long
Dim Y As Long

    If bEraseOtherSelection Then
        'alors on efface les anciennes sélections
        Refresh
    End If

    'colorise une ligne hexa
    UserControl.Line (1410, 260 * lLine + 100)-(7170, 260 * (lLine + 1) + 50), lColor, BF
    
    'colorise une ligne de string
    UserControl.Line (7520, 260 * lLine + 100)-(9440, 260 * (lLine + 1) + 50), lColor, BF
    
    'réécrit la string
    PasteString lLine, Mstr(lLine)
    
    'reecrit le texte
    For x = 1 To 16
        PasteHex lLine, x, Mhex(x, lLine)
    Next x
        
End Sub

'=======================================================
'changement des offsets ==> redessine, redimensionne et vide les tableaux
'=======================================================
Public Sub ChangeValues()

    'détermine le nombre de lignes à afficher
    'lNumberPerPage = IIf(Me.NumberPerPage < (Int(pct.Height / 250) - 1), Me.NumberPerPage, Int(pct.Height / 250) - 1)
    
    'création des échelles et des titres
    CreateBackGround
    
    'initilisation - vidage des tableaux
    ReDim Mlng(lNumberPerPage + 1) As Long
    ReDim Mhex(16, lNumberPerPage + 1) As String
    'ReDim Msel(16, lNumberPerPage + 1) As Boolean
    ReDim Mstr(lNumberPerPage + 1) As String
    ReDim M_M(16, lNumberPerPage + 1) As Boolean
    
End Sub

'=======================================================
'obtient la string FORMATEE de la ligne lLine
'=======================================================
Public Function GetString(ByVal lLine As Long) As String
    GetString = Mstr(lLine)
End Function

'=======================================================
'obtient la string REELLE de la ligne lLine
'=======================================================
Public Function GetRealString(ByVal lLine As Long) As String
Dim x As Byte
Dim s As String
    s = vbNullString
    
    'prend les valeurs depuis la liste des valeurs HEXA
    For x = 1 To 16
        s = s & Chr$(Hex2Dec(Mhex(x, lLine)))
    Next x
    GetRealString = s
End Function

'=======================================================
'ajoute une valeur hexa au tableau
'=======================================================
Public Sub AddHexValue(ByVal lLine As Long, ByVal lCol As Long, ByVal sHexValue As String, Optional ByVal bIsModified = False)
    If lLine > lNumberPerPage Then Exit Sub
    Mhex(lCol, lLine) = sHexValue

    M_M(lCol, lLine) = bIsModified  'renseigne sur une modification potentielle de la case
End Sub

'=======================================================
'ajoute une valeur hexa au tableau
'=======================================================
Public Sub AddStringValue(ByVal lLine As Long, ByVal sString As String)
    If lLine > lNumberPerPage Then Exit Sub
    Mstr(lLine) = sString
End Sub

'=======================================================
'ajoute une valeur hexa au tableau
'=======================================================
Public Sub AddOneStringValue(ByVal lLine As Long, ByVal lCol As Long, ByVal sString As String)
Dim sAvant As String
Dim sApres As String

    If lLine > lNumberPerPage Then Exit Sub
    
    If lCol = 1 Then
        'pas de sAvant
        sAvant = vbNullString
        sApres = Mid$(Mstr(lLine), 2, 15)
    ElseIf lCol = 16 Then
        'pas de sApres
        sApres = vbNullString
        sAvant = Mid$(Mstr(lLine), 1, 15)
    Else
        'alors il y a sAvant ET sApres
        sApres = Mid$(Mstr(lLine), lCol + 1, 16 - lCol + 1)
        sAvant = Mid$(Mstr(lLine), 1, lCol - 1)
    End If
    
    Mstr(lLine) = sAvant & sString & sApres
End Sub

'=======================================================
'remplit le texte dans toutes les cases
'=======================================================
Public Sub FillText()
Dim x As Long
Dim Y As Long

    CreateBackGround
    
    For x = 1 To lNumberPerPage
        PasteString x, Mstr(x)  'affiche les strings
        For Y = 1 To 16
            PasteHex x, Y, Mhex(Y, x)
        Next Y
        PasteOffset x, FormatedAdress(lFirstOffset + (x - 1) * 16, bUseHexOffset)
    Next x

End Sub



'=======================================================
'SUBS PRIVEES
'=======================================================

'=======================================================
'affiche la sélection
'=======================================================
Private Sub CreateSelections(ByVal tType As ItemType)
Dim z As Long
Dim w As Long
Dim yCase As Currency
Dim xCase As Long
Dim NewCase As HexCase
Dim OldCase As HexCase
Dim lStep As Long


    'détermine les Case TEMPORAIRES de départ et d'arrivée de sélection
    'Old DOIT ETRE inférieur à New au niveau de l'offset
    If hexOldCase.lOffset > hexNewCase.lOffset Then
        OldCase.lCol = hexNewCase.lCol
        OldCase.lOffset = hexNewCase.lOffset
        NewCase.lCol = hexOldCase.lCol
        NewCase.lOffset = hexOldCase.lOffset
    Else
        NewCase.lCol = hexNewCase.lCol
        NewCase.lOffset = hexNewCase.lOffset
        OldCase.lCol = hexOldCase.lCol
        OldCase.lOffset = hexOldCase.lOffset
    End If
    
    xCase = NewCase.lCol
    yCase = (NewCase.lOffset - Me.FirstOffset) / 16 + 1
        
    If yCase < 1 Then Exit Sub
    If yCase > Me.NumberPerPage Then yCase = Me.NumberPerPage
    If xCase < 1 Then xCase = 1
    If OldCase.lCol < 1 Then OldCase.lCol = 1
   
    '//ZONE OFFSET
    If tType = tOffset Then
        'alors c'est un offset
        
        'on sélectionne tout entre xOld et x, yOld et y
        
        If OldCase.lOffset < Me.FirstOffset Then
            'alors la sélection provient d'avant ce qui est visible
            For z = 1 To yCase
                ColorItem tOffset, z, 1, FirstOffset + 16 * (z - 1), lSelectionColor, False, False
            Next z
        ElseIf Me.FirstOffset <= OldCase.lOffset And OldCase.lOffset <= (Me.FirstOffset + 16 * Me.NumberPerPage) Then
            'alors la sélection provient de la même page
            yOld = (hexOldCase.lOffset - Me.FirstOffset) / 16 + 1
            lStep = IIf(yOld > yCase, -1, 1)
            For z = yOld To yCase Step lStep
                ColorItem tOffset, z, 1, FirstOffset + 16 * (z - 1), lSelectionColor, False, False
            Next z
        'ElseIf OldCase.lOffset > (Me.FirstOffset + 16 * Me.NumberPerPage) Then
            'alors la sélection provient d'une page plus basse
            'For z = Me.NumberPerPage To yCase Step -1
            '    ColorItem tOffset, z, 1, FirstOffset + 16 * (z - 1), lSelectionColor, False, False
           ' Next z
        End If
    End If
    
    
    '//ZONE HEXA
    If tType = tHex Then
        'alors c'est une valeur hexa
        
        'xOld et yOld sont les anciennes valeurs
        'remplit une sélection en fonction de xOld, xCase, yOld et yCase

        If OldCase.lOffset < Me.FirstOffset Then
            'alors la sélection provient d'avant ce qui est visible
            yOld = 1
            xOld = 1
        ElseIf Me.FirstOffset <= OldCase.lOffset And OldCase.lOffset <= (Me.FirstOffset + 16 * Me.NumberPerPage) Then
            'alors la sélection provient de la même page
            yOld = (OldCase.lOffset - Me.FirstOffset) / 16 + 1
            xOld = OldCase.lCol
        ElseIf OldCase.lOffset > (Me.FirstOffset + 16 * Me.NumberPerPage) Then
            'alors la sélection provient d'une page plus basse
            yOld = Me.NumberPerPage
            xOld = 16
        End If
        
        If yCase = yOld Then
            'remplit sur l'horizontale
            lStep = IIf(xCase > xOld, -1, 1)
            For z = xCase To xOld Step lStep
                'colore l'offset, les valeurs hex, et la string
                ColorItem tHex, yOld, z, "", lSelectionColor, False
            Next z
        End If
        
        If yCase > yOld Then
            'finit première la ligne
            For z = xOld To 16
                'colore l'offset, les valeurs hex, et la string
                ColorItem tHex, yOld, z, "", lSelectionColor, False
            Next z
            'finit la dernière ligne
            For z = xCase To 1 Step -1
                'colore l'offset, les valeurs hex, et la string
                ColorItem tHex, yCase, z, "", lSelectionColor, False
            Next z
            'fait les lignes entre si nécessaire
            If (yCase - yOld) > 1 Then
                'il y a des lignes entre : les fait
                'remplit par LIGNE (pas par élément) ==> gagne du temps
                For z = yCase - 1 To yOld + 1 Step -1
                    ColorLine z, lSelectionColor, False
                Next z
            End If
        End If
        
        If yCase < yOld Then
            For z = xCase To 16
                'colore l'offset, les valeurs hex, et la string
                ColorItem tHex, yCase, z, "", lSelectionColor, False
            Next z
            For z = xOld To 1 Step -1
                'colore l'offset, les valeurs hex, et la string
                ColorItem tHex, yOld, z, "", lSelectionColor, False
            Next z
            'fait les lignes entre si nécessaire
            If (yOld - yCase) > 1 Then
                'il y a des lignes entre : les fait
                'remplit par LIGNE (pas par élément) ==> gagne du temps
                For z = yOld - 1 To yCase + 1 Step -1
                    ColorLine z, lSelectionColor, False
                Next z
            End If
        End If
    End If
    
    '//ZONE STRING
    If tType = tString Then
        'alors c'est une valeur string
        
        If OldCase.lOffset < Me.FirstOffset Then
            'alors la sélection provient d'avant ce qui est visible
            yOld = 1
            xOld = 1
        ElseIf Me.FirstOffset <= OldCase.lOffset And OldCase.lOffset <= (Me.FirstOffset + 16 * Me.NumberPerPage) Then
            'alors la sélection provient de la même page
            yOld = (OldCase.lOffset - Me.FirstOffset) / 16 + 1
            xOld = OldCase.lCol
        ElseIf OldCase.lOffset > (Me.FirstOffset + 16 * Me.NumberPerPage) Then
            'alors la sélection provient d'une page plus basse
            yOld = Me.NumberPerPage
            xOld = 16
        End If
        
        If xCase > 16 Then xCase = 16
        
        If yCase = yOld Then
            'remplit sur l'horizontale
            lStep = IIf(xCase > xOld, -1, 1)
            For z = xCase To xOld Step lStep
                'colore l'offset, les valeurs hex, et la string
                ColorItem tHex, yOld, z, "", lSelectionColor, False
            Next z
        End If
        
        If yCase > yOld Then
            'finit première la ligne
            For z = xOld To 16
                'colore l'offset, les valeurs hex, et la string
                ColorItem tHex, yOld, z, "", lSelectionColor, False
            Next z
            'finit la dernière ligne
            For z = xCase To 1 Step -1
                'colore l'offset, les valeurs hex, et la string
                ColorItem tHex, yCase, z, "", lSelectionColor, False
            Next z
            'fait les lignes entre si nécessaire
            If (yCase - yOld) > 1 Then
                For z = yCase - 1 To yOld + 1 Step -1
                    ColorLine z, lSelectionColor, False
                Next z
            End If
        End If
        
        If yCase < yOld Then
            For z = xCase To 16
                'colore l'offset, les valeurs hex, et la string
                ColorItem tHex, yCase, z, "", lSelectionColor, False
            Next z
            For z = xOld To 1 Step -1
                'colore l'offset, les valeurs hex, et la string
                ColorItem tHex, yOld, z, "", lSelectionColor, False
            Next z
            'fait les lignes entre si nécessaire
            If (yOld - yCase) > 1 Then
                For z = yOld - 1 To yCase + 1 Step -1
                    ColorLine z, lSelectionColor, False
                Next z
            End If
        End If
    End If
       
        
End Sub

'=======================================================
'renvoie une adresse (string) avec les 0 devant si nécessaire, pour avoir une longueur
'de string fixe (10)
'=======================================================
Private Function FormatedAdress(ByVal lNumber As Currency, Optional ByVal bAsHex As Boolean = False) As String
Dim s As String

    If bAsHex Then
        'formate en hexa
        
        s = ExtendedHex(lNumber)    'valeur hexa étendue
        
    Else
        'formate en long
        s = CStr(CCur(lNumber))
        
        While Len(s) < 10
            s = "0" + s
        Wend
    End If

    FormatedAdress = s
End Function

'=======================================================
'renvoie a^b
'=======================================================
Private Function AexpB(ByVal a As Long, ByVal b As Long) As Long
Dim x As Long
Dim l As Long

    If b = 0 Then
        AexpB = 1
        Exit Function
    End If
    
    l = 1
    For x = 1 To b
        l = l * a
    Next x
    AexpB = l

End Function

'=======================================================
'convertit une valeur hexa (string) en valeur décimale
'=======================================================
Private Function Hex2Dec(ByVal s As String) As Long
Dim x As Long
Dim l As Long

    For x = Len(s) To 1 Step -1
        l = l + HexVal(Mid$(s, Len(s) - x + 1, 1)) * AexpB(16, x - 1)
    Next x

    Hex2Dec = l
End Function

'=======================================================
'renvoie la valeur decimale d'une string de longueur 1 en hexa
'=======================================================
Private Function HexVal(ByVal s As String) As Long
    If s = "0" Then
        HexVal = 0
    ElseIf s = "1" Then
        HexVal = 1
    ElseIf s = "2" Then
        HexVal = 2
    ElseIf s = "3" Then
        HexVal = 3
    ElseIf s = "4" Then
        HexVal = 4
    ElseIf s = "5" Then
        HexVal = 5
    ElseIf s = "6" Then
        HexVal = 6
    ElseIf s = "7" Then
        HexVal = 7
    ElseIf s = "8" Then
        HexVal = 8
    ElseIf s = "9" Then
        HexVal = 9
    ElseIf LCase(s) = "a" Then
        HexVal = 10
    ElseIf LCase(s) = "b" Then
        HexVal = 11
    ElseIf LCase(s) = "c" Then
        HexVal = 12
    ElseIf LCase(s) = "d" Then
        HexVal = 13
    ElseIf LCase(s) = "e" Then
        HexVal = 14
    ElseIf LCase(s) = "f" Then
        HexVal = 15
    End If
End Function

'=======================================================
'affecte les valeurs CurrentX et CurrentY au usercontrol
'=======================================================
Private Sub Pos(ByVal CurrentX As Long, CurrentY As Long)
    UserControl.CurrentX = CurrentX
    UserControl.CurrentY = CurrentY
End Sub

'=======================================================
'écrit une String à la ligne lLine
'=======================================================
Private Sub PasteOffset(ByVal lLine As Long, ByVal sString As String)
    UserControl.ForeColor = lOffsetForeColor
    Pos 50, 260 * (lLine) + 100
    UserControl.Print sString
End Sub
'=======================================================
'écrit une String à la ligne lLine
'=======================================================
Private Sub PasteString(ByVal lLine As Long, ByVal sString As String)
    
    If (Me.FirstOffset + 16 * (lLine - 1) + 1) > curFileSize Then Exit Sub 'dépasse du fichier
        
    UserControl.ForeColor = lStringForeColor
    Pos 7500, 260 * (lLine) + 100
    
    'tronque la string si on est au bout du fichier
    If (Me.FirstOffset + 16 * lLine) > curFileSize Then
        sString = Left$(sString, curFileSize - Me.FirstOffset - 16 * (lLine - 1))
    End If
    
    UserControl.Print sString
End Sub
'=======================================================
'écrit une valeurs hexa a un endroit de la matrice
'=======================================================
Private Sub PasteHex(ByVal lLine As Long, ByVal lCol As Long, ByVal sString As String)

    If (Me.FirstOffset + 16 * (lLine - 1) + lCol) > curFileSize Then Exit Sub 'dépasse du fichier
    
    If M_M(lCol, lLine) = False Then
        'case normale
        UserControl.ForeColor = lHexForeColor
    Else
        'case modifiée
        UserControl.ForeColor = lModifiedItemColor
    End If
    Pos 360 * lCol + 1100, 260 * (lLine) + 100
    'If lCol > 8 Then pct.CurrentX = pct.CurrentX + 200
    UserControl.Print sString
End Sub

'=======================================================
'créé les titres et les lignes du tableau
'=======================================================
Private Sub CreateBackGround()
Dim x As Long

    'UserControl.Cls
    
    'grisage
    UserControl.Line (0, 0)-(9600, 300), lTitleBackGround, BF
    UserControl.Line (0, 300)-(9600, 300), lLineColor
    
    '"Offset"
    Pos 0, 50
    UserControl.ForeColor = lOffsetTitleForeColor
    UserControl.Print "   Offset"
    
    'ligne verticale 1
    UserControl.Line (1300, 0)-(1300, UserControl.Height), lLineColor

    '"0 2 3 4 5 ..15"
    For x = 0 To 15
        Pos 360 * x + 1500, 50
        If x > 9 Then UserControl.CurrentX = UserControl.CurrentX - 50
        UserControl.ForeColor = lBaseTitleForeColor
        UserControl.Print CStr(x)
    Next x
    
    'ligne verticale 2
    UserControl.Line (7300, 0)-(7300, UserControl.Height), lLineColor
    'ligne verticale 3
    UserControl.Line (9600, 0)-(9600, UserControl.Height), lLineColor
    
    TraceGrid
    
End Sub

'=======================================================
'Trace le Grid (lignes du tableau ENTRE les valeurs)
'=======================================================
Private Sub TraceGrid()
Dim x As Long
Dim l As Long
Dim lDeltaOffset As Long
Dim lDeltaString As Long

'trace (ou pas) le grid
    
    If bGrid = HorizontalHexOnly Or bGrid = HorizontalHexOnly_VerticalHex Then
        lDeltaOffset = 0
        lDeltaString = 0
    Else
        lDeltaOffset = -1300
        lDeltaString = 2300
    End If
    
    If bGrid = Horizontal Or bGrid = Horizontal_VerticalHex Or bGrid = HorizontalHexOnly Or bGrid = HorizontalHexOnly_VerticalHex Then
        'alors on trace un Grid horizontal
        For x = 1 To lNumberPerPage
            UserControl.Line (1300 + lDeltaOffset, 300 + 260 * x)-(7300 + lDeltaString, 300 + 260 * x), lLineColor
        Next x
    End If
    If bGrid = VerticalHex Or bGrid = HorizontalHexOnly_VerticalHex Or bGrid = Horizontal_VerticalHex Then
        'alors on trace un Grid vertical des valeurs hexa
        For x = 1 To 15
            UserControl.Line (1400 + 360 * x, 300)-(1400 + 360 * x, UserControl.Height), lLineColor
        Next x
    End If
    
End Sub

'=======================================================
'traçage des signets
'=======================================================
Public Sub TraceSignets()   'public pour permettre un refresh rapide après un ajout (car pour obliger de faire HW.Refresh)
Dim x As Long
Dim l As Currency

    For x = 2 To UBound(M_S)
        l = M_S(x) - Me.FirstOffset
        If l >= 0 And l <= (lNumberPerPage * 16) Then
            'alors çà se situe sur la partie affichée ==> on affiche
            UserControl.Line (10, 360 + l * 16.24)-(1260, 555 + l * 16.24), lSignetColor, B
        End If
    Next x
End Sub

'=======================================================
'efface uniquement la zone de sélection (réapplique la MaskPicture)
'=======================================================
Private Sub Refresh2()
    
    UserControl.Cls
    UserControl.Picture = UserControl.MaskPicture

End Sub

'=======================================================
'ajoute une sélection au tableau des sélections
'=======================================================
Public Sub AddSelection(ByVal lLine As Long, ByVal lCol As Long)
    'Msel(lCol, lLine) = True
End Sub

'=======================================================
'sélectionne manuellement une zone
'=======================================================
Public Sub SelectZone(Col1 As Long, Offset1 As Currency, Col2 As Long, Offset2 As Currency)
    hexOldCase.lCol = Col1
    hexOldCase.lOffset = Offset1
    hexNewCase.lCol = Col2
    hexNewCase.lOffset = Offset2
End Sub

'=======================================================
'permet de calculer la valeur hexa d'un nombre très grand
'(currency, jusqu'à 15*16^12+15*16^11+15*16^10+...)
'=======================================================
Private Function ExtendedHex(ByVal cVal As Currency) As String
Dim x As Long
Dim s As String
Dim table16(9) As Currency
Dim res(9) As Byte

    cVal = cVal + 1 'ajoute 1 pour que le résultat soit juste

    'contient la table des 16^n
    table16(0) = 1
    table16(1) = 16
    table16(2) = 256
    table16(3) = 4096
    table16(4) = 65536
    table16(5) = 1048576
    table16(6) = 16777216
    table16(7) = 268435456
    table16(8) = 4294967296#
    table16(9) = 68719476736#

    'enlève, en partant des plus grosses valeurs, un maximum de fois un 16^x
    For x = 9 To 0 Step -1
        While cVal > table16(x)
            cVal = cVal - table16(x)
            res(x) = res(x) + 1 'ajoute 1 à l'occurence de table16(x)
        Wend
    Next x
    
    'créé la string
    For x = 9 To 0 Step -1
        s = s & Hex(res(x))
    Next
    
    ExtendedHex = s
End Function













'=======================================================
'SUBS FOR SELF-SUBCLASSING
'cette partie n'est pas de moi
'thx to Paul Caton
'=======================================================

'Determine if the passed function is supported
Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
  Dim hmod        As Long
  Dim bLibLoaded  As Boolean

  hmod = GetModuleHandleA(sModule)

  If hmod = 0 Then
    hmod = LoadLibraryA(sModule)
    If hmod Then
      bLibLoaded = True
    End If
  End If

  If hmod Then
    If GetProcAddress(hmod, sFunction) Then
      IsFunctionExported = True
    End If
  End If

  If bLibLoaded Then
    FreeLibrary hmod
  End If
End Function

'-SelfSub code=======================================================-------------------------------------
Private Function sc_Subclass(ByVal lng_hWnd As Long, _
                    Optional ByVal lParamUser As Long = 0, _
                    Optional ByVal nOrdinal As Long = 1, _
                    Optional ByVal oCallback As Object = Nothing, _
                    Optional ByVal bIdeSafety As Boolean = True) As Boolean 'Subclass the specified window handle
'*************************************************************************************************
'* lng_hWnd   - Handle of the window to subclass
'* lParamUser - Optional, user-defined callback parameter
'* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
'* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
'* bIdeSafety - Optional, enable/disable IDE safety measures. NB: you should really only disable IDE safety in a UserControl for design-time subclassing
'*************************************************************************************************
Const CODE_LEN      As Long = 260                                           'Thunk length in bytes
Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1))            'Bytes to allocate per thunk, data + code + msg tables
Const PAGE_RWX      As Long = &H40&                                         'Allocate executable memory
Const MEM_COMMIT    As Long = &H1000&                                       'Commit allocated memory
Const MEM_RELEASE   As Long = &H8000&                                       'Release allocated memory flag
Const IDX_EBMODE    As Long = 3                                             'Thunk data index of the EbMode function address
Const IDX_CWP       As Long = 4                                             'Thunk data index of the CallWindowProc function address
Const IDX_SWL       As Long = 5                                             'Thunk data index of the SetWindowsLong function address
Const IDX_FREE      As Long = 6                                             'Thunk data index of the VirtualFree function address
Const IDX_BADPTR    As Long = 7                                             'Thunk data index of the IsBadCodePtr function address
Const IDX_OWNER     As Long = 8                                             'Thunk data index of the Owner object's vTable address
Const IDX_CALLBACK  As Long = 10                                            'Thunk data index of the callback method address
Const IDX_EBX       As Long = 16                                            'Thunk code patch index of the thunk data
Const SUB_NAME      As String = "sc_Subclass"                               'This routine's name
  Dim nAddr         As Long
  Dim nID           As Long
  Dim nMyID         As Long
  
  If IsWindow(lng_hWnd) = 0 Then                                            'Ensure the window handle is valid
    zError SUB_NAME, "Invalid window handle"
    Exit Function
  End If

  nMyID = GetCurrentProcessId                                               'Get this process's ID
  GetWindowThreadProcessId lng_hWnd, nID                                    'Get the process ID associated with the window handle
  If nID <> nMyID Then                                                      'Ensure that the window handle doesn't belong to another process
    zError SUB_NAME, "Window handle belongs to another process"
    Exit Function
  End If
  
  If oCallback Is Nothing Then                                              'If the user hasn't specified the callback owner
    Set oCallback = Me                                                      'Then it is me
  End If
  
  nAddr = zAddressOf(oCallback, nOrdinal)                                   'Get the address of the specified ordinal method
  If nAddr = 0 Then                                                         'Ensure that we've found the ordinal method
    zError SUB_NAME, "Callback method not found"
    Exit Function
  End If
    
  If z_Funk Is Nothing Then                                                 'If this is the first time through, do the one-time initialization
    Set z_Funk = New Collection                                             'Create the hWnd/thunk-address collection
    z_Sc(14) = &HD231C031: z_Sc(15) = &HBBE58960: z_Sc(17) = &H4339F631: z_Sc(18) = &H4A21750C: z_Sc(19) = &HE82C7B8B: z_Sc(20) = &H74&: z_Sc(21) = &H75147539: z_Sc(22) = &H21E80F: z_Sc(23) = &HD2310000: z_Sc(24) = &HE8307B8B: z_Sc(25) = &H60&: z_Sc(26) = &H10C261: z_Sc(27) = &H830C53FF: z_Sc(28) = &HD77401F8: z_Sc(29) = &H2874C085: z_Sc(30) = &H2E8&: z_Sc(31) = &HFFE9EB00: z_Sc(32) = &H75FF3075: z_Sc(33) = &H2875FF2C: z_Sc(34) = &HFF2475FF: z_Sc(35) = &H3FF2473: z_Sc(36) = &H891053FF: z_Sc(37) = &HBFF1C45: z_Sc(38) = &H73396775: z_Sc(39) = &H58627404
    z_Sc(40) = &H6A2473FF: z_Sc(41) = &H873FFFC: z_Sc(42) = &H891453FF: z_Sc(43) = &H7589285D: z_Sc(44) = &H3045C72C: z_Sc(45) = &H8000&: z_Sc(46) = &H8920458B: z_Sc(47) = &H4589145D: z_Sc(48) = &HC4836124: z_Sc(49) = &H1862FF04: z_Sc(50) = &H35E30F8B: z_Sc(51) = &HA78C985: z_Sc(52) = &H8B04C783: z_Sc(53) = &HAFF22845: z_Sc(54) = &H73FF2775: z_Sc(55) = &H1C53FF28: z_Sc(56) = &H438D1F75: z_Sc(57) = &H144D8D34: z_Sc(58) = &H1C458D50: z_Sc(59) = &HFF3075FF: z_Sc(60) = &H75FF2C75: z_Sc(61) = &H873FF28: z_Sc(62) = &HFF525150: z_Sc(63) = &H53FF2073: z_Sc(64) = &HC328&

    z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA")                    'Store CallWindowProc function address in the thunk data
    z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA")                     'Store the SetWindowLong function address in the thunk data
    z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree")                     'Store the VirtualFree function address in the thunk data
    z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr")                  'Store the IsBadCodePtr function address in the thunk data
  End If
  
  z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)                  'Allocate executable memory

  If z_ScMem <> 0 Then                                                      'Ensure the allocation succeeded
    On Error GoTo CatchDoubleSub                                            'Catch double subclassing
      z_Funk.Add z_ScMem, "h" & lng_hWnd                                    'Add the hWnd/thunk-address to the collection
    On Error GoTo 0
  
    If bIdeSafety Then                                                      'If the user wants IDE protection
      z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode")                          'Store the EbMode function address in the thunk data
    End If
    
    z_Sc(IDX_EBX) = z_ScMem                                                 'Patch the thunk data address
    z_Sc(IDX_HWND) = lng_hWnd                                               'Store the window handle in the thunk data
    z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN                                   'Store the address of the before table in the thunk data
    z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4)         'Store the address of the after table in the thunk data
    z_Sc(IDX_OWNER) = ObjPtr(oCallback)                                     'Store the callback owner's object address in the thunk data
    z_Sc(IDX_CALLBACK) = nAddr                                              'Store the callback address in the thunk data
    z_Sc(IDX_PARM_USER) = lParamUser                                        'Store the lParamUser callback parameter in the thunk data
    
    nAddr = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF)    'Set the new WndProc, return the address of the original WndProc
    If nAddr = 0 Then                                                       'Ensure the new WndProc was set correctly
      zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
      GoTo ReleaseMemory
    End If
        
    z_Sc(IDX_WNDPROC) = nAddr                                               'Store the original WndProc address in the thunk data
    RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                        'Copy the thunk code/data to the allocated memory
    sc_Subclass = True                                                      'Indicate success
  Else
    zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
  End If
  
  Exit Function                                                             'Exit sc_Subclass

CatchDoubleSub:
  zError SUB_NAME, "Window handle is already subclassed"
  
ReleaseMemory:
  VirtualFree z_ScMem, 0, MEM_RELEASE                                       'sc_Subclass has failed after memory allocation, so release the memory
End Function

'Terminate all subclassing
Private Sub sc_Terminate()
  Dim i As Long

  If Not (z_Funk Is Nothing) Then                                           'Ensure that subclassing has been started
    With z_Funk
      For i = .Count To 1 Step -1                                           'Loop through the collection of window handles in reverse order
        z_ScMem = .Item(i)                                                  'Get the thunk address
        If IsBadCodePtr(z_ScMem) = 0 Then                                   'Ensure that the thunk hasn't already released its memory
          sc_UnSubclass zData(IDX_HWND)                                     'UnSubclass
        End If
      Next i                                                                'Next member of the collection
    End With
    Set z_Funk = Nothing                                                    'Destroy the hWnd/thunk-address collection
  End If
End Sub

'UnSubclass the specified window handle
Private Sub sc_UnSubclass(ByVal lng_hWnd As Long)
  If z_Funk Is Nothing Then                                                 'Ensure that subclassing has been started
    zError "sc_UnSubclass", "Window handle isn't subclassed"
  Else
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                           'Ensure that the thunk hasn't already released its memory
      zData(IDX_SHUTDOWN) = -1                                              'Set the shutdown indicator
      zDelMsg ALL_MESSAGES, IDX_BTABLE                                      'Delete all before messages
      zDelMsg ALL_MESSAGES, IDX_ATABLE                                      'Delete all after messages
    End If
    z_Funk.Remove "h" & lng_hWnd                                            'Remove the specified window handle from the collection
  End If
End Sub

'Add the message value to the window handle's specified callback table
Private Sub sc_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    If When And MSG_BEFORE Then                                             'If the message is to be added to the before original WndProc table...
      zAddMsg uMsg, IDX_BTABLE                                              'Add the message to the before table
    End If
    If When And MSG_AFTER Then                                              'If message is to be added to the after original WndProc table...
      zAddMsg uMsg, IDX_ATABLE                                              'Add the message to the after table
    End If
  End If
End Sub

'Delete the message value from the window handle's specified callback table
Private Sub sc_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    If When And MSG_BEFORE Then                                             'If the message is to be deleted from the before original WndProc table...
      zDelMsg uMsg, IDX_BTABLE                                              'Delete the message from the before table
    End If
    If When And MSG_AFTER Then                                              'If the message is to be deleted from the after original WndProc table...
      zDelMsg uMsg, IDX_ATABLE                                              'Delete the message from the after table
    End If
  End If
End Sub

'Call the original WndProc
Private Function sc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    sc_CallOrigWndProc = _
        CallWindowProcA(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
  End If
End Function

'Get the subclasser lParamUser callback parameter
Private Property Get sc_lParamUser(ByVal lng_hWnd As Long) As Long
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    sc_lParamUser = zData(IDX_PARM_USER)                                    'Get the lParamUser callback parameter
  End If
End Property

'Let the subclasser lParamUser callback parameter
Private Property Let sc_lParamUser(ByVal lng_hWnd As Long, ByVal NewValue As Long)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    zData(IDX_PARM_USER) = NewValue                                         'Set the lParamUser callback parameter
  End If
End Property

'-The following routines are exclusively for the sc_ subclass routines----------------------------

'Add the message to the specified table of the window handle
Private Sub zAddMsg(ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count
  Dim nBase  As Long                                                        'Remember z_ScMem
  Dim i      As Long                                                        'Loop index

  nBase = z_ScMem                                                            'Remember z_ScMem so that we can restore its value on exit
  z_ScMem = zData(nTable)                                                    'Map zData() to the specified table

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being added to the table...
    nCount = ALL_MESSAGES                                                   'Set the table entry count to ALL_MESSAGES
  Else
    nCount = zData(0)                                                       'Get the current table entry count
    If nCount >= MSG_ENTRIES Then                                           'Check for message table overflow
      zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values"
      GoTo Bail
    End If

    For i = 1 To nCount                                                     'Loop through the table entries
      If zData(i) = 0 Then                                                  'If the element is free...
        zData(i) = uMsg                                                     'Use this element
        GoTo Bail                                                           'Bail
      ElseIf zData(i) = uMsg Then                                           'If the message is already in the table...
        GoTo Bail                                                           'Bail
      End If
    Next i                                                                  'Next message table entry

    nCount = i                                                              'On drop through: i = nCount + 1, the new table entry count
    zData(nCount) = uMsg                                                    'Store the message in the appended table entry
  End If

  zData(0) = nCount                                                         'Store the new table entry count
Bail:
  z_ScMem = nBase                                                           'Restore the value of z_ScMem
End Sub

'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count
  Dim nBase  As Long                                                        'Remember z_ScMem
  Dim i      As Long                                                        'Loop index

  nBase = z_ScMem                                                           'Remember z_ScMem so that we can restore its value on exit
  z_ScMem = zData(nTable)                                                   'Map zData() to the specified table

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being deleted from the table...
    zData(0) = 0                                                            'Zero the table entry count
  Else
    nCount = zData(0)                                                       'Get the table entry count
    
    For i = 1 To nCount                                                     'Loop through the table entries
      If zData(i) = uMsg Then                                               'If the message is found...
        zData(i) = 0                                                        'Null the msg value -- also frees the element for re-use
        GoTo Bail                                                           'Bail
      End If
    Next i                                                                  'Next message table entry
    
    zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
  End If
  
Bail:
  z_ScMem = nBase                                                           'Restore the value of z_ScMem
End Sub

'Error handler
Private Sub zError(ByVal sRoutine As String, ByVal sMsg As String)
  App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
  MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long
  zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                   'Get the specified procedure address
  Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
End Function

'Map zData() to the thunk address for the specified window handle
Private Function zMap_hWnd(ByVal lng_hWnd As Long) As Long
  If z_Funk Is Nothing Then                                                 'Ensure that subclassing has been started
    zError "zMap_hWnd", "Subclassing hasn't been started"
  Else
    On Error GoTo Catch                                                     'Catch unsubclassed window handles
    z_ScMem = z_Funk("h" & lng_hWnd)                                        'Get the thunk address
    zMap_hWnd = z_ScMem
  End If
  
  Exit Function                                                             'Exit returning the thunk address

Catch:
  zError "zMap_hWnd", "Window handle isn't subclassed"
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable
  Dim i     As Long                                                         'Loop index
  Dim j     As Long                                                         'Loop limit
  
  RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
  If Not zProbe(nAddr + &H1C, i, bSub) Then                                 'Probe for a Class method
    If Not zProbe(nAddr + &H6F8, i, bSub) Then                              'Probe for a Form method
      If Not zProbe(nAddr + &H7A4, i, bSub) Then                            'Probe for a UserControl method
        Exit Function                                                       'Bail...
      End If
    End If
  End If
  
  i = i + 4                                                                 'Bump to the next entry
  j = i + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
  Do While i < j
    RtlMoveMemory VarPtr(nAddr), i, 4                                       'Get the address stored in this vTable entry
    
    If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If

    RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If
    
    i = i + 4                                                             'Next vTable entry
  Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                            'Start address
  nLimit = nAddr + 32                                                       'Probe eight entries
  Do While nAddr < nLimit                                                   'While we've not reached our probe depth
    RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
    
    If nEntry <> 0 Then                                                     'If not an implemented interface
      RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
        nMethod = nAddr                                                     'Store the vTable entry
        bSub = bVal                                                         'Store the found method signature
        zProbe = True                                                       'Indicate success
        Exit Function                                                       'Return
      End If
    End If
    
    nAddr = nAddr + 4                                                       'Next vTable entry
  Loop
End Function

Private Property Get zData(ByVal nIndex As Long) As Long
  RtlMoveMemory VarPtr(zData), z_ScMem + (nIndex * 4), 4
End Property

Private Property Let zData(ByVal nIndex As Long, ByVal nValue As Long)
  RtlMoveMemory z_ScMem + (nIndex * 4), VarPtr(nValue), 4
End Property

'-Subclass callback, usually ordinal #1, the last method in this source file----------------------
Private Sub zWndProc1(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Long)
'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc and (if set to do so) the after
'*              original WndProc callback.
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter
'*************************************************************************************************

    If uMsg = WM_MOUSEWHEEL Then
        'capté le mousewheel
        RaiseEvent MouseWheel(wParam)   'on récuperera ensuite le signe de l'argument renvoyé
        'pour déterminer le sens du mouvement effectué par la molette
    End If
    
End Sub


