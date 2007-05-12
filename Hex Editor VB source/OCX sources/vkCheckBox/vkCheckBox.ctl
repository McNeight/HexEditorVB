VERSION 5.00
Begin VB.UserControl vkCheckBox 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "vkCheckBox.ctx":0000
   Begin VB.Image Image1 
      Height          =   195
      Left            =   480
      Picture         =   "vkCheckBox.ctx":0312
      Top             =   600
      Visible         =   0   'False
      Width           =   1170
   End
End
Attribute VB_Name = "vkCheckBox"
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
'CONSTANTES
'=======================================================
Private Const WM_KEYDOWN                    As Long = &H100
Private Const WM_KEYFIRST                   As Long = &H100
Private Const WM_KEYLAST                    As Long = &H108
Private Const WM_KEYUP                      As Long = &H101
Private Const WM_LBUTTONDBLCLK              As Long = &H203
Private Const WM_LBUTTONDOWN                As Long = &H201
Private Const WM_LBUTTONUP                  As Long = &H202
Private Const WM_MBUTTONDBLCLK              As Long = &H209
Private Const WM_MBUTTONDOWN                As Long = &H207
Private Const WM_MBUTTONUP                  As Long = &H208
Private Const WM_MOUSEFIRST                 As Long = &H200
Private Const WM_MOUSELAST                  As Long = &H209
Private Const WM_MOUSEHOVER                 As Long = &H2A1
Private Const WM_MOUSELEAVE                 As Long = &H2A3
Private Const WM_MOUSEMOVE                  As Long = &H200
Private Const WM_RBUTTONDBLCLK              As Long = &H206
Private Const WM_RBUTTONDOWN                As Long = &H204
Private Const WM_RBUTTONUP                  As Long = &H205
Private Const WM_MOUSEWHEEL                 As Long = &H20A
Private Const WM_PAINT                      As Long = &HF
Private Const GWL_WNDPROC                   As Long = -4&
Private Const TME_LEAVE                     As Long = &H2&
Private Const TME_HOVER                     As Long = &H1&
Private Const DT_CENTER                     As Long = &H1&
Private Const DT_LEFT                       As Long = &H0&
Private Const DT_RIGHT                      As Long = &H2&
Private Const DI_MASK                       As Long = &H1
Private Const DI_IMAGE                      As Long = &H2
Private Const DI_NORMAL                     As Long = DI_MASK Or DI_IMAGE
Private Const SRCCOPY                       As Long = 13369376


'=======================================================
'APIS
'=======================================================
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENTTYPE) As Long
Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetTabbedTextExtent Lib "user32" Alias "GetTabbedTextExtentA" (ByVal hdc As Long, ByVal lpString As String, ByVal nCount As Long, ByVal nTabPositions As Long, lpnTabStopPositions As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long


'=======================================================
'TYPES
'=======================================================
Private Type TRACKMOUSEEVENTTYPE
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type



'=======================================================
'ENUMS
'=======================================================
Public Enum BackStyleConstants
    Transparent = 0
    Opaque = 1
End Enum
Public Enum TextPositionConstants
    [Text_Left] = 0
    [Text_Center] = 1
    [Text_Right] = 2
End Enum
Public Enum WHEEL_SENS
    WHEEL_UP
    WHEEL_DOWN
End Enum
Public Enum GradientConstants
    None = 0
    Vertical = 1
    Horizontal = 2
End Enum
Public Enum PictureAligment
    [Left Justify]
    [Right Justify]
End Enum
 

'=======================================================
'VARIABLES PRIVEES
'=======================================================
Private mAsm(63) As Byte    'contient le code ASM
Private OldProc As Long     'adresse de l'ancienne window proc
Private objHwnd As Long     'handle de l'objet concerné
Private ET As TRACKMOUSEEVENTTYPE   'type pour le mouse_hover et le mouse_leave
Private IsMouseIn As Boolean    'si la souris est dans le controle

Private lBackStyle As BackStyleConstants
Private lForeColor As OLE_COLOR
Private bCol As OLE_COLOR
Private bEnable As Boolean
Private tVal As CheckBoxConstants
Private sCaption As String
Private bNotOk As Boolean
Private bHasFocus As Boolean
Private bNotOk2 As Boolean


'=======================================================
'EVENTS
'=======================================================
Public Event Change(Value As CheckBoxConstants)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseHover()
Public Event MouseMove()
Public Event MouseLeave()
Public Event MouseWheel(Sens As WHEEL_SENS)
Public Event MouseDown(Button As MouseButtonConstants)
Public Event MouseUp(Button As MouseButtonConstants)
Public Event MouseDblClick(Button As MouseButtonConstants)




'=======================================================
'USERCONTROL SUBS
'=======================================================
'=======================================================
' /!\ NE PAS DEPLACER CETTE FONCTION /!\ '
'=======================================================
' Cette fonction doit rester la premiere '
' fonction "public" du module de classe  '
'=======================================================
'          Fonction de CallBack          '
'             Par EBArtSoft              '
'                                        '
'=======================================================
Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  
    Select Case uMsg
        
        Case WM_LBUTTONDBLCLK
            RaiseEvent MouseDblClick(vbLeftButton)
        Case WM_LBUTTONDOWN
            Call ChangeValue
            RaiseEvent MouseDown(vbLeftButton)
        Case WM_LBUTTONUP
            RaiseEvent MouseUp(vbLeftButton)
        Case WM_MBUTTONDBLCLK
            RaiseEvent MouseDblClick(vbMiddleButton)
        Case WM_MBUTTONDOWN
            RaiseEvent MouseDown(vbMiddleButton)
        Case WM_MBUTTONUP
            RaiseEvent MouseUp(vbMiddleButton)
        Case WM_MOUSEHOVER
            If IsMouseIn = False Then
                RaiseEvent MouseHover
                IsMouseIn = True
                'on refresh l'icone
                If bEnable Then Call SplitIMGandShow: UserControl.Refresh
            End If
        Case WM_MOUSELEAVE
            RaiseEvent MouseLeave
            IsMouseIn = False
            'on refresh l'icone
            If bEnable Then Call SplitIMGandShow: UserControl.Refresh
        Case WM_MOUSEMOVE
            Call TrackMouseEvent(ET)
            RaiseEvent MouseMove
        Case WM_RBUTTONDBLCLK
            RaiseEvent MouseDblClick(vbRightButton)
        Case WM_RBUTTONDOWN
            RaiseEvent MouseDown(vbRightButton)
        Case WM_RBUTTONUP
            RaiseEvent MouseUp(vbRightButton)
        Case WM_MOUSEWHEEL
            If wParam < 0 Then
                RaiseEvent MouseWheel(WHEEL_DOWN)
            Else
                RaiseEvent MouseWheel(WHEEL_UP)
            End If
        Case WM_PAINT
            bNotOk = True  'évite le clignotement lors du survol de la souris
    End Select
    
    'appel de la routine standard pour les autres messages
    WindowProc = CallWindowProc(OldProc, hWnd, uMsg, wParam, lParam)
    
End Function

'=======================================================
'Change la value du control
'=======================================================
Private Sub ChangeValue()
    
    If tVal = vbChecked Then
        tVal = vbUnchecked
    ElseIf tVal = vbUnchecked Then
        tVal = vbChecked
    End If
    
    RaiseEvent Change(tVal)
    
    bNotOk = False: Call UserControl_Paint
    
End Sub

Private Sub UserControl_GotFocus()
'alors on va tracer un BÔ rectangle de sélection
Dim R As RECT
Dim y As Long
Dim x As Long

    If bEnable = False Then
        'on ne garde pas le focus
        Call SendKeys("{Tab}")
        Exit Sub
    End If
    
    'on a alors le focus
    bHasFocus = True
    
    '//on dessine le rectangle de focus
    'une zone rectangulaire
    y = (UserControl.ScaleHeight / 15 - GetCharHeight) / 2
    Call SetRect(R, 17, y - 1, TextWidth(sCaption) / 15 + 23, y + _
        GetCharHeight + 2)
    'dessine
    Call DrawFocusRect(UserControl.hdc, R)
End Sub

Private Sub UserControl_Initialize()
Dim Ofs As Long
Dim Ptr As Long
        
    'Recupere l'adresse de "Me.WindowProc"
    Call CopyMemory(Ptr, ByVal (ObjPtr(Me)), 4)
    Call CopyMemory(Ptr, ByVal (Ptr + 489 * 4), 4)
    
    'Crée la veritable fonction WindowProc (à optimiser)
    Ofs = VarPtr(mAsm(0))
    MovL Ofs, &H424448B            '8B 44 24 04          mov         eax,dword ptr [esp+4]
    MovL Ofs, &H8245C8B            '8B 5C 24 08          mov         ebx,dword ptr [esp+8]
    MovL Ofs, &HC244C8B            '8B 4C 24 0C          mov         ecx,dword ptr [esp+0Ch]
    MovL Ofs, &H1024548B           '8B 54 24 10          mov         edx,dword ptr [esp+10h]
    MovB Ofs, &H68                 '68 44 33 22 11       push        Offset RetVal
    MovL Ofs, VarPtr(mAsm(59))
    MovB Ofs, &H52                 '52                   push        edx
    MovB Ofs, &H51                 '51                   push        ecx
    MovB Ofs, &H53                 '53                   push        ebx
    MovB Ofs, &H50                 '50                   push        eax
    MovB Ofs, &H68                 '68 44 33 22 11       push        ObjPtr(Me)
    MovL Ofs, ObjPtr(Me)
    MovB Ofs, &HE8                 'E8 1E 04 00 00       call        Me.WindowProc
    MovL Ofs, Ptr - Ofs - 4
    MovB Ofs, &HA1                 'A1 20 20 40 00       mov         eax,RetVal
    MovL Ofs, VarPtr(mAsm(59))
    MovL Ofs, &H10C2               'C2 10 00             ret         10h
End Sub

Private Sub UserControl_InitProperties()
    'valeurs par défaut
    bNotOk2 = True
    With Me
        .BackColor = vbWhite '
        .BackStyle = Opaque
        .Caption = "Caption" '
        .Font = Ambient.Font '
        .ForeColor = vbBlack '
        .Enabled = True '
        .Value = False
    End With
    bNotOk2 = False
    Call UserControl_Paint  'refresh
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp, vbKeyLeft:
            Call SendKeys("+{Tab}")
        Case vbKeyDown, vbKeyRight:
            Call SendKeys("{Tab}")
        Case vbKeySpace
            Call ChangeValue
    End Select
    
    Call Refresh
End Sub

Private Sub UserControl_LostFocus()
    bHasFocus = False: bNotOk = False
    Call UserControl_Paint
End Sub

Private Sub UserControl_Terminate()
    'vire le subclassing
    If OldProc Then Call SetWindowLong(UserControl.hWnd, GWL_WNDPROC, OldProc)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("BackColor", Me.BackColor, &HC0C0C0)
        Call .WriteProperty("BackStyle", Me.BackStyle, Opaque)
        Call .WriteProperty("Caption", Me.Caption, "Caption")
        Call .WriteProperty("Font", Me.Font, Ambient.Font)
        Call .WriteProperty("ForeColor", Me.ForeColor, vbBlack)
        Call .WriteProperty("Enabled", Me.Enabled, True)
        Call .WriteProperty("Value", Me.Value, False)
    End With
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    bNotOk2 = True
    With PropBag
        Me.BackColor = .ReadProperty("BackColor", &HC0C0C0)
        Me.BackStyle = .ReadProperty("BackStyle", Opaque)
        Me.Caption = .ReadProperty("Caption", "Caption")
        Set Me.Font = .ReadProperty("Font", Ambient.Font)
        Me.ForeColor = .ReadProperty("ForeColor", vbBlack)
        Me.Enabled = .ReadProperty("Enabled", True)
        Me.Value = .ReadProperty("Value", False)
    End With
    bNotOk2 = False
    Call UserControl_Paint  'refresh
    
    'le bon endroit pour lancer le subclassing
    Call LaunchKeyMouseEvents
End Sub
Private Sub UserControl_Resize()
    bNotOk = False
    Call UserControl_Paint  'refresh
End Sub

'=======================================================
'lance le subclassing
'=======================================================
Public Sub LaunchKeyMouseEvents()
                
    If Ambient.UserMode Then

        OldProc = SetWindowLong(UserControl.hWnd, GWL_WNDPROC, _
            VarPtr(mAsm(0)))    'pas de AddressOf aujourd'hui ;)
            
        'prépare le terrain pour le mouse_over et mouse_leave
        With ET
            .cbSize = Len(ET)
            .hwndTrack = UserControl.hWnd
            .dwFlags = TME_LEAVE Or TME_HOVER
            .dwHoverTime = 1
        End With
        
        'démarre le tracking de l'entrée
        Call TrackMouseEvent(ET)
        
        'pas dedans par défaut
        IsMouseIn = False
        
    End If
    
End Sub



'=======================================================
'PROPERTIES
'=======================================================
Public Property Get hdc() As Long: hdc = UserControl.hdc: End Property
Public Property Get hWnd() As Long: hWnd = UserControl.hWnd: End Property
Public Property Get BackStyle() As BackStyleConstants: BackStyle = lBackStyle: End Property
Public Property Let BackStyle(BackStyle As BackStyleConstants): lBackStyle = BackStyle: UserControl.BackStyle = BackStyle: bNotOk = False: UserControl_Paint: End Property
Public Property Get Caption() As String: Caption = sCaption: End Property
Public Property Let Caption(Caption As String): sCaption = Caption: bNotOk = False: UserControl_Paint: bNotOk = True: End Property
Public Property Get ForeColor() As OLE_COLOR: ForeColor = lForeColor: End Property
Public Property Let ForeColor(ForeColor As OLE_COLOR): lForeColor = ForeColor: UserControl.ForeColor = ForeColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get BackColor() As OLE_COLOR: BackColor = bCol: End Property
Public Property Let BackColor(BackColor As OLE_COLOR)
UserControl.BackColor = BackColor
bCol = BackColor: bNotOk = False: UserControl_Paint:
End Property
Public Property Get Font() As StdFont: Set Font = UserControl.Font: End Property
Public Property Set Font(Font As StdFont): Set UserControl.Font = Font: bNotOk = False: UserControl_Paint: End Property
Public Property Get Enabled() As Boolean: Enabled = bEnable: End Property
Public Property Let Enabled(Enabled As Boolean)
bEnable = Enabled: bNotOk = False: UserControl_Paint
End Property
Public Property Get Value() As CheckBoxConstants: Value = tVal: End Property
Public Property Let Value(Value As CheckBoxConstants): tVal = Value: bNotOk = False: UserControl_Paint: End Property


Private Sub UserControl_Paint()

    If bNotOk Or bNotOk2 Then Exit Sub     'pas prêt à peindre
    
    Call Refresh    'on refresh
End Sub




'=======================================================
'PRIVATE SUBS
'=======================================================
'=======================================================
'copie un "byte"
'=======================================================
Private Sub MovB(Ofs As Long, ByVal Value As Long)
    Call CopyMemory(ByVal Ofs, Value, 1): Ofs = Ofs + 1
End Sub

'=======================================================
'copie un "long"
'=======================================================
Private Sub MovL(Ofs As Long, ByVal Value As Long)
    Call CopyMemory(ByVal Ofs, Value, 4): Ofs = Ofs + 4
End Sub

'=======================================================
'récupère la hauteur d'un caractère
'=======================================================
Private Function GetCharHeight() As Long
Dim Res As Long
    Res = GetTabbedTextExtent(UserControl.hdc, "A", 1, 0, 0)
    GetCharHeight = (Res And &HFFFF0000) \ &H10000
End Function


'=======================================================
'MAJ du controle
'=======================================================
Public Sub Refresh()
Dim R As RECT
Dim yVal As Long

    '//on efface
    Call UserControl.Cls
    
    '//on locke
    Call LockWindowUpdate(Me.hWnd)
    
    UserControl.MaskPicture = Nothing
    UserControl.ForeColor = lForeColor
    
    '//copnvertir les couleurs
    Call OleTranslateColor(bCol, 0, bCol)
    Call OleTranslateColor(lForeColor, 0, lForeColor)
    
    
    '//on va afficher l'image correspondant à l'état
    Call SplitIMGandShow
    
    '//on va tracer le rectangle de focus si on a le focus
    If bHasFocus Then Call UserControl_GotFocus
    
    '//affiche le texte
    yVal = (UserControl.ScaleHeight - GetCharHeight * 15) / 2
    'définit une zone pour le texte
    Call SetRect(R, 20, yVal / 15, TextWidth(sCaption) / 15 + 20, UserControl.ScaleHeight / 15)
    'dessine le texte
    Call DrawText(UserControl.hdc, sCaption, Len(sCaption), R, DT_CENTER)


    '//style
    If lBackStyle = [Transparent] Then
        'transparent
        With UserControl
            .BackStyle = 0
            .MaskColor = bCol
            Set MaskPicture = .Image
        End With
    Else
        UserControl.BackStyle = 1
    End If
    
    '//on délocke
    Call LockWindowUpdate(0)
    
    '//on refresh le control
    Call UserControl.Refresh
    
    bNotOk = True
End Sub

'=======================================================
'affiche une des 6 images en la découpant depuis l'image complète
'=======================================================
Private Sub SplitIMGandShow()
Dim SrcDC As Long
Dim SrcObj As Long
Dim y As Single
Dim lIMG As Long

    '0 rien
    '1 survol
    '2 enabled=false
    '3 value enable
    '4 value survol enable
    '5 enable=false OR gray
    
    If bEnable = False Then
        'grisé
        If tVal = vbUnchecked Then
            'pas checked
            lIMG = 2
        Else
            'checked et gris
            lIMG = 5
        End If
    Else
        'enabled=true
        If IsMouseIn Then
            'alors mouse survol
            If tVal = vbChecked Then
                'checked
                lIMG = 4
            ElseIf tVal = vbUnchecked Then
                'non checked
                lIMG = 1
            Else
                'gray
                lIMG = 5
            End If
        Else
            'pas de survol
            If tVal = vbChecked Then
                'checked
                lIMG = 3
            ElseIf tVal = vbUnchecked Then
                'non checked
                lIMG = 0
            Else
                'gray
                lIMG = 5
            End If
        End If
    End If
    
    'on découpe l'image correspondant à lIMG depuis Image1 et on blit
    'sur l'usercontrol
    SrcDC = CreateCompatibleDC(hdc)
    SrcObj = SelectObject(SrcDC, Image1.Picture)
    
    y = (UserControl.ScaleHeight / 15 - 13) / 2
    Call BitBlt(UserControl.hdc, 0, y, 13, 13, SrcDC, lIMG * 13, 0, SRCCOPY)

    Call DeleteDC(SrcDC)
    Call DeleteObject(SrcObj)
End Sub

