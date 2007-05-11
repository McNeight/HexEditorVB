VERSION 5.00
Begin VB.UserControl grdFrame 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "grdFrame.ctx":0000
   Begin VB.PictureBox pctG 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   2160
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image PCTcolor 
      Height          =   240
      Left            =   720
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image PCTgray 
      Height          =   240
      Left            =   360
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "grdFrame"
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


'=======================================================
'APIS
'=======================================================
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENTTYPE) As Long
Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, Lppoint As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetTabbedTextExtent Lib "user32" Alias "GetTabbedTextExtentA" (ByVal hdc As Long, ByVal lpString As String, ByVal nCount As Long, ByVal nTabPositions As Long, lpnTabStopPositions As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function FrameRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long


'=======================================================
'TYPES
'=======================================================
Private Type TRACKMOUSEEVENTTYPE
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type
Private Type RGB_COLOR
    R As Long
    G As Long
    B As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
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
Private lTextPos As TextPositionConstants
Private lTitleHeight As Long
Private lForeColor As OLE_COLOR
Private tCol1 As OLE_COLOR
Private tCol2 As OLE_COLOR
Private bCol1 As OLE_COLOR
Private bCol2 As OLE_COLOR
Private bShowTitle As Boolean
Private sCaption As String
Private bShowBackGround As Boolean
Private lTitleGradient As GradientConstants
Private lBackGradient As GradientConstants
Private bNotOk As Boolean
Private bNotOk2 As Boolean
Private bEnable As Boolean
Private lBorderColor As OLE_COLOR
Private bDisplayBorder As Boolean
Private bBreakCorner As Boolean
Private lCornerSize As Long
Private lBWidth As Long
Private pctAlign As PictureAligment
Private bPic As Boolean
Private lOffsetX As Long
Private lOffsetY As Long
Private bGray As Boolean

'=======================================================
'EVENTS
'=======================================================
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
            End If
        Case WM_MOUSELEAVE
            RaiseEvent MouseLeave
            IsMouseIn = False
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
        .BackColor1 = &HC0C0C0    '
        .BackColor2 = vbWhite '
        .BackGradient = Horizontal '
        .BackStyle = Opaque
        .Caption = "Caption" '
        .Font = Ambient.Font '
        .ForeColor = vbWhite '
        .ShowBackGround = True '
        .ShowTitle = True '
        .TextPosition = Text_Center '
        .TitleColor1 = vbBlue
        .TitleColor2 = vbWhite '
        .TitleGradient = Vertical '
        .TitleHeight = 300 '
        .Enabled = True '
        .BorderColor = &HFF8080    '
        .DisplayBorder = True '
        .BreakCorner = True '
        .BorderWidth = 1 '
        .RoundAngle = 7 '
        Set .Picture = Nothing
        .PictureAligment = [Left Justify]
        .DisplayPicture = True
        .PictureOffsetX = 0
        .PictureOffsetY = 0
        .GrayPictureWhenDisabled = True
    End With
    bNotOk2 = False
    Call UserControl_Paint  'refresh
End Sub

Private Sub UserControl_Terminate()
    'vire le subclassing
    If OldProc Then Call SetWindowLong(UserControl.hWnd, GWL_WNDPROC, OldProc)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("BackColor1", Me.BackColor1, &HC0C0C0)
        Call .WriteProperty("BackColor2", Me.BackColor2, vbWhite)
        Call .WriteProperty("BackGradient", Me.BackGradient, Horizontal)
        Call .WriteProperty("BackStyle", Me.BackStyle, Opaque)
        Call .WriteProperty("Caption", Me.Caption, "Caption")
        Call .WriteProperty("Font", Me.Font, Ambient.Font)
        Call .WriteProperty("ForeColor", Me.ForeColor, vbWhite)
        Call .WriteProperty("ShowBackGround", Me.ShowBackGround, True)
        Call .WriteProperty("ShowTitle", Me.ShowTitle, True)
        Call .WriteProperty("TextPosition", Me.TextPosition, Text_Center)
        Call .WriteProperty("TitleColor1", Me.TitleColor1, vbBlue)
        Call .WriteProperty("TitleColor2", Me.TitleColor2, vbWhite)
        Call .WriteProperty("TitleGradient", Me.TitleGradient, Vertical)
        Call .WriteProperty("TitleHeight", Me.TitleHeight, 300)
        Call .WriteProperty("Enabled", Me.Enabled, True)
        Call .WriteProperty("BorderColor", Me.BorderColor, &HFF8080)
        Call .WriteProperty("DisplayBorder", Me.DisplayBorder, True)
        Call .WriteProperty("BreakCorner", Me.BreakCorner, True)
        Call .WriteProperty("Picture", Me.Picture, Nothing)
        Call .WriteProperty("RoundAngle", Me.RoundAngle, 7)
        Call .WriteProperty("BorderWidth", Me.BorderWidth, 1)
        Call .WriteProperty("PictureAligment", Me.PictureAligment, [Left Justify])
        Call .WriteProperty("DisplayPicture", Me.DisplayPicture, True)
        Call .WriteProperty("PictureOffsetX", Me.PictureOffsetX, 0)
        Call .WriteProperty("PictureOffsetY", Me.PictureOffsetY, 0)
        Call .WriteProperty("GrayPictureWhenDisabled", Me.GrayPictureWhenDisabled, True)
    End With
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    bNotOk2 = True
    With PropBag
        Me.BackColor1 = .ReadProperty("BackColor1", &HC0C0C0)
        Me.BackColor2 = .ReadProperty("BackColor2", vbWhite)
        Me.BackGradient = .ReadProperty("BackGradient", Horizontal)
        Me.BackStyle = .ReadProperty("BackStyle", Opaque)
        Me.Caption = .ReadProperty("Caption", "Caption")
        Set Me.Font = .ReadProperty("Font", Ambient.Font)
        Me.ForeColor = .ReadProperty("ForeColor", vbWhite)
        Me.ShowBackGround = .ReadProperty("ShowBackGround", True)
        Me.ShowTitle = .ReadProperty("ShowTitle", True)
        Me.TextPosition = .ReadProperty("TextPosition", Text_Center)
        Me.TitleColor1 = .ReadProperty("TitleColor1", vbBlue)
        Me.TitleColor2 = .ReadProperty("TitleColor2", vbWhite)
        Me.TitleGradient = .ReadProperty("TitleGradient", Vertical)
        Me.TitleHeight = .ReadProperty("TitleHeight", 300)
        Me.Enabled = .ReadProperty("Enabled", True)
        Me.BorderColor = .ReadProperty("BorderColor", &HFF8080)
        Me.DisplayBorder = .ReadProperty("DisplayBorder", True)
        Me.BreakCorner = .ReadProperty("BreakCorner", True)
        Set Me.Picture = .ReadProperty("Picture", Nothing)
        Me.BorderWidth = .ReadProperty("BorderWidth", 1)
        Me.RoundAngle = .ReadProperty("RoundAngle", 7)
        Me.PictureAligment = .ReadProperty("PictureAligment", [Left Justify])
        Me.DisplayPicture = .ReadProperty("DisplayPicture", True)
        Me.PictureOffsetX = .ReadProperty("PictureOffsetX", 0)
        Me.PictureOffsetY = .ReadProperty("PictureOffsetY", 0)
        Me.GrayPictureWhenDisabled = .ReadProperty("GrayPictureWhenDisabled", True)
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
Public Property Get TextPosition() As TextPositionConstants: TextPosition = lTextPos: End Property
Public Property Let TextPosition(TextPosition As TextPositionConstants): lTextPos = TextPosition: bNotOk = False: UserControl_Paint: End Property
Public Property Get Caption() As String: Caption = sCaption: End Property
Public Property Let Caption(Caption As String): sCaption = Caption: bNotOk = False: UserControl_Paint: bNotOk = True: End Property
Public Property Get ShowTitle() As Boolean: ShowTitle = bShowTitle: End Property
Public Property Let ShowTitle(ShowTitle As Boolean): bShowTitle = ShowTitle: bNotOk = False: UserControl_Paint: End Property
Public Property Get TitleHeight() As Long: TitleHeight = lTitleHeight: End Property
Public Property Let TitleHeight(TitleHeight As Long): lTitleHeight = TitleHeight: bNotOk = False: UserControl_Paint: End Property
Public Property Get ForeColor() As OLE_COLOR: ForeColor = lForeColor: End Property
Public Property Let ForeColor(ForeColor As OLE_COLOR): lForeColor = ForeColor: UserControl.ForeColor = ForeColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get TitleColor1() As OLE_COLOR: TitleColor1 = tCol1: End Property
Public Property Let TitleColor1(TitleColor1 As OLE_COLOR): tCol1 = TitleColor1: bNotOk = False: UserControl_Paint: End Property
Public Property Get TitleColor2() As OLE_COLOR: TitleColor2 = tCol2: End Property
Public Property Let TitleColor2(TitleColor2 As OLE_COLOR): tCol2 = TitleColor2: bNotOk = False: UserControl_Paint: End Property
Public Property Get BackColor1() As OLE_COLOR: BackColor1 = bCol1: End Property
Public Property Let BackColor1(BackColor1 As OLE_COLOR): bCol1 = BackColor1: bNotOk = False: UserControl_Paint: End Property
Public Property Get BackColor2() As OLE_COLOR: BackColor2 = bCol2: End Property
Public Property Let BackColor2(BackColor2 As OLE_COLOR): bCol2 = BackColor2: bNotOk = False: UserControl_Paint: End Property
Public Property Get Font() As StdFont: Set Font = UserControl.Font: End Property
Public Property Set Font(Font As StdFont): Set UserControl.Font = Font: bNotOk = False: UserControl_Paint: End Property
Public Property Get ShowBackGround() As Boolean: ShowBackGround = bShowBackGround: End Property
Public Property Let ShowBackGround(ShowBackGround As Boolean): bShowBackGround = ShowBackGround: bNotOk = False: UserControl_Paint: End Property
Public Property Get TitleGradient() As GradientConstants: TitleGradient = lTitleGradient: End Property
Public Property Let TitleGradient(TitleGradient As GradientConstants): lTitleGradient = TitleGradient: bNotOk = False: UserControl_Paint: End Property
Public Property Get BackGradient() As GradientConstants: BackGradient = lBackGradient: End Property
Public Property Let BackGradient(BackGradient As GradientConstants): lBackGradient = BackGradient: bNotOk = False: UserControl_Paint: End Property
Public Property Get Enabled() As Boolean: Enabled = bEnable: End Property
Public Property Let Enabled(Enabled As Boolean)
bEnable = Enabled: bNotOk = False: UserControl_Paint: EnableControls
End Property
Public Property Get BorderColor() As OLE_COLOR: BorderColor = lBorderColor: End Property
Public Property Let BorderColor(BorderColor As OLE_COLOR): lBorderColor = BorderColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get DisplayBorder() As Boolean: DisplayBorder = bDisplayBorder: End Property
Public Property Let DisplayBorder(DisplayBorder As Boolean): bDisplayBorder = DisplayBorder: bNotOk = False: UserControl_Paint: End Property
Public Property Get BreakCorner() As Boolean: BreakCorner = bBreakCorner: End Property
Public Property Let BreakCorner(BreakCorner As Boolean): bBreakCorner = BreakCorner: bNotOk = False: UserControl_Paint: End Property
Public Property Get Picture() As Picture: Set Picture = PCTcolor.Picture: End Property
Public Property Set Picture(NewPic As Picture)
Set PCTcolor.Picture = NewPic
Set pctG.Picture = NewPic
If Not (NewPic Is Nothing) Then Call GrayScale(pctG)
PCTgray.Picture = pctG.Image
bNotOk = False: UserControl_Paint
End Property
Public Property Get RoundAngle() As Long: RoundAngle = lCornerSize: End Property
Public Property Let RoundAngle(RoundAngle As Long): lCornerSize = RoundAngle: bNotOk = False: UserControl_Paint: End Property
Public Property Get BorderWidth() As Long: BorderWidth = lBWidth: End Property
Public Property Let BorderWidth(BorderWidth As Long): lBWidth = BorderWidth: bNotOk = False: UserControl_Paint: End Property
Public Property Get PictureAligment() As PictureAligment: PictureAligment = pctAlign: End Property
Public Property Let PictureAligment(PictureAligment As PictureAligment): pctAlign = PictureAligment: bNotOk = False: UserControl_Paint: End Property
Public Property Get DisplayPicture() As Boolean: DisplayPicture = bPic: End Property
Public Property Let DisplayPicture(DisplayPicture As Boolean): bPic = DisplayPicture: bNotOk = False: UserControl_Paint: End Property
Public Property Get PictureOffsetX() As Long: PictureOffsetX = lOffsetX: End Property
Public Property Let PictureOffsetX(PictureOffsetX As Long): lOffsetX = PictureOffsetX: bNotOk = False: UserControl_Paint: End Property
Public Property Get PictureOffsetY() As Long: PictureOffsetY = lOffsetY: End Property
Public Property Let PictureOffsetY(PictureOffsetY As Long): lOffsetY = PictureOffsetY: bNotOk = False: UserControl_Paint: End Property
Public Property Get GrayPictureWhenDisabled() As Boolean: GrayPictureWhenDisabled = bGray: End Property
Public Property Let GrayPictureWhenDisabled(GrayPictureWhenDisabled As Boolean): bGray = GrayPictureWhenDisabled: bNotOk = False: UserControl_Paint: End Property


Private Sub UserControl_Paint()

    If bNotOk Or bNotOk2 Then Exit Sub     'pas prêt à peindre
    
    Call Refresh    'on refresh
End Sub




'=======================================================
'PRIVATE SUBS
'=======================================================
'=======================================================
'applique un gradient de couleur sur un objet de gauche à droite
'il doit être "en autoredraw=true" (si c'est une form, picturebox...)
'=======================================================
Private Sub FillGradientW(LeftColor As RGB_COLOR, _
    RightColor As RGB_COLOR, ByVal Width As Long, ByVal Height As Long, _
    Optional ByVal Dep As Long)
    
Dim rAverageColorPerSizeUnit As Double
Dim gAverageColorPerSizeUnit As Double
Dim bAverageColorPerSizeUnit As Double
Dim lWidth As Long
Dim x As Long
Dim lHeight As Long
Dim lSigne As Long

    With UserControl
        
        'récupère la largeur de l'objet
        lWidth = Width / 15
        lHeight = Height / 15
        
        'récupère la moyenne de couleur par unité de longueur
        rAverageColorPerSizeUnit = Abs((RightColor.R - LeftColor.R) / lWidth)
        gAverageColorPerSizeUnit = Abs((RightColor.G - LeftColor.G) / lWidth)
        bAverageColorPerSizeUnit = Abs((RightColor.B - LeftColor.B) / lWidth)
        
        'on change le signe (sens) au cas où
        If CLng(RGB(LeftColor.R, LeftColor.G, LeftColor.B)) <= _
            CLng(RGB(RightColor.R, RightColor.G, RightColor.B)) Then
            
            lSigne = 1
        Else
            lSigne = -1
        End If
        
        'se positionne tout à gauche de l'objet ==> balayera vers la droite
        Call MoveToEx(.hdc, 0, Dep, 0&)
        
        'pour chaque 'colonne' constituée par une ligne verticale, on trace une
        'ligne en récupérant la couleur correspondante
        For x = 0 To lWidth
            
            'change le ForeColor qui détermine la couleur de la Line
            'multiplie la largeur actuelle par la couleur par unité de longueur
            .ForeColor = RGB(LeftColor.R + x * rAverageColorPerSizeUnit * lSigne, LeftColor.G + x * _
                gAverageColorPerSizeUnit * lSigne, LeftColor.B + x * bAverageColorPerSizeUnit * lSigne)
               
            'trace une ligne
            Call LineTo(.hdc, x, lHeight)
            
            'bouge 'd'une colonne' vers la droite
            Call MoveToEx(.hdc, x, Dep, 0&)
        
        Next x
        
        'on refresh l'objet
        Call .Refresh
    End With

End Sub

'=======================================================
'applique un gradient de couleur sur un objet de gauche à droite
'il doit être "en autoredraw=true" (si c'est une form, picturebox...)
'=======================================================
Private Sub FillGradientH(LeftColor As RGB_COLOR, _
    RightColor As RGB_COLOR, ByVal Width As Long, ByVal Height As Long, _
    Optional ByVal Dep As Long)
    
Dim rAverageColorPerSizeUnit As Double
Dim gAverageColorPerSizeUnit As Double
Dim bAverageColorPerSizeUnit As Double
Dim lHeight As Long
Dim x As Long
Dim lSigne As Long

    With UserControl
        
        'récupère la hateur de l'objet
        lHeight = Height / 15
        
        'récupère la moyenne de couleur par unité de longueur
        rAverageColorPerSizeUnit = Abs((RightColor.R - LeftColor.R) / lHeight)
        gAverageColorPerSizeUnit = Abs((RightColor.G - LeftColor.G) / lHeight)
        bAverageColorPerSizeUnit = Abs((RightColor.B - LeftColor.B) / lHeight)

        'on change le signe (sens) au cas où
        If CLng(RGB(LeftColor.R, LeftColor.G, LeftColor.B)) <= _
            CLng(RGB(RightColor.R, RightColor.G, RightColor.B)) Then
            
            lSigne = 1
        Else
            lSigne = -1
        End If
        
        'se positionne tout à gauche de l'objet ==> balayera vers le bas
        Call MoveToEx(.hdc, 0, Dep, 0&)
        
        'pour chaque 'colonne' constituée par une ligne verticale, on trace une
        'ligne en récupérant la couleur correspondante
        For x = Dep To lHeight
            
            'change le ForeColor qui détermine la couleur de la Line
            'multiplie la largeur actuelle par la couleur par unité de longueur
            .ForeColor = RGB(LeftColor.R + x * rAverageColorPerSizeUnit * lSigne, LeftColor.G + x * _
                gAverageColorPerSizeUnit * lSigne, LeftColor.B + x * bAverageColorPerSizeUnit * lSigne)
               
            'trace une ligne
            Call LineTo(.hdc, Width, x)
            
            'bouge 'd'une colonne' vers la droite
            Call MoveToEx(.hdc, 0, x, 0&)
        
        Next x
        
        'on refresh l'objet
        Call .Refresh
    End With

End Sub

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
'convertit une couleur en long vers RGB
'=======================================================
Private Sub ToRGB(ByVal Color As Long, ByRef RGB As RGB_COLOR)
    With RGB
        .R = Color And &HFF&
        .G = (Color And &HFF00&) \ &H100&
        .B = Color \ &H10000
    End With
End Sub

'=======================================================
'récupère la hauteur d'un caractère
'=======================================================
Private Function GetCharHeight() As Long
Dim Res As Long
    Res = GetTabbedTextExtent(UserControl.hdc, "A", 1, 0, 0)
    GetCharHeight = (Res And &HFFFF0000) \ &H10000
End Function

''=======================================================
''renvoie le min des composantes
''=======================================================
'Private Function RGB_Min(RGB1 As RGB_COLOR, RGB2 As RGB_COLOR) As RGB_COLOR
'    With RGB_Min
'        If RGB1.R < RGB2.R Then .R = RGB1.R Else .R = RGB2.R
'        If RGB1.B < RGB2.B Then .B = RGB1.B Else .B = RGB2.B
'        If RGB1.G < RGB2.G Then .G = RGB1.G Else .G = RGB2.G
'    End With
'End Function
'
''=======================================================
''renvoie le max des composantes
''=======================================================
'Private Function RGB_Max(RGB1 As RGB_COLOR, RGB2 As RGB_COLOR) As RGB_COLOR
'    With RGB_Max
'        If RGB1.R > RGB2.R Then .R = RGB1.R Else .R = RGB2.R
'        If RGB1.B > RGB2.B Then .B = RGB1.B Else .B = RGB2.B
'        If RGB1.G > RGB2.G Then .G = RGB1.G Else .G = RGB2.G
'    End With
'End Function



'=======================================================
'PUBLIC SUB
'=======================================================

'=======================================================
'enable ou pas les controles qui sont dans le frame
'=======================================================
Private Sub EnableControls()
Dim ctr As Control
    
    On Error Resume Next
    
    For Each ctr In UserControl.ContainedControls
        ctr.Enabled = bEnable
    Next ctr
    
End Sub

'=======================================================
'on dessine tout
'=======================================================
Public Sub Refresh()
Dim x As Long
Dim RGB1 As RGB_COLOR
Dim RGB2 As RGB_COLOR
Dim Dep As Long
Dim R As RECT
Dim hBrush As Long
Dim Rec As Long
Dim hRgn As Long
Dim W As Long
Dim H As Long
    
    '//on locke le controle
    Call LockWindowUpdate(UserControl.hWnd)
    
    '//on efface et on vire le maskpicture
    Call UserControl.Cls
    UserControl.Picture = Nothing
    UserControl.MaskPicture = Nothing
    
    '//on convertir les différentes couleurs si couleurs système
    Call OleTranslateColor(lBorderColor, 0, lBorderColor)
    Call OleTranslateColor(lForeColor, 0, lForeColor)
    Call OleTranslateColor(tCol1, 0, tCol1)
    Call OleTranslateColor(tCol2, 0, tCol2)
    Call OleTranslateColor(bCol1, 0, bCol1)
    Call OleTranslateColor(bCol2, 0, bCol2)
    
    
    '//on commence par créer le Title
    If bShowTitle Then
        
'        If lBackStyle = Transparent And bBreakCorner = True Then
'            'alors c'est transparent et avec arrondi
'            'alors on créé la zone perso dès maintenant
'            'pour ne pas pouvoir tracer le titre dans les coins
'            'haut-gauche et haut-droite
'
'            'on définit une zone rectangulaire à bords arrondi
'            hRgn = CreateRoundRectRgn(0, 0, UserControl.ScaleWidth / 15, _
'                UserControl.ScaleHeight / 15, lCornerSize, lCornerSize)
'
'            'on défini la zone rectangulaire arrondi comme nouvelle fenêtre
'            Call SetWindowRgn(UserControl.hWnd, hRgn, True)
'
'            'on détruit le brush et la zone
'            Call DeleteObject(hRgn)
'        End If
            
            
        'récupère les 3 composantes des deux couleurs
        Call ToRGB(tCol1, RGB1)
        Call ToRGB(tCol2, RGB2)
        
        If lTitleGradient = None Then
            'pas de gradient
            'on dessine alors un rectangle
            UserControl.Line (0, 0)-(UserControl.ScaleWidth, lTitleHeight), _
                tCol1, BF
        ElseIf lTitleGradient = Horizontal Then
            'gradient horizontal
            Call FillGradientH(RGB1, RGB2, UserControl.ScaleWidth, lTitleHeight)
        Else
            'gradient vertical
            Call FillGradientW(RGB1, RGB2, UserControl.ScaleWidth, lTitleHeight)
        End If
    End If
    
    '//créé un rectangle
    Call SetRect(R, 0, (lTitleHeight / 15 - GetCharHeight) / 2, UserControl.Width / 15, lTitleHeight / 15)
    
    'on affiche le caption
    UserControl.ForeColor = lForeColor
    If lTextPos = Text_Center Then
        'au centre
        Call DrawText(UserControl.hdc, sCaption, Len(sCaption), R, DT_CENTER)
    ElseIf lTextPos = Text_Right Then
        'à droite
        Call DrawText(UserControl.hdc, sCaption, Len(sCaption), R, DT_RIGHT)
    Else
        'à gauche
        Call DrawText(UserControl.hdc, sCaption, Len(sCaption), R, DT_LEFT)
    End If
    
    If lBackStyle = Transparent Then
        'alors on créé une image dans le maskpicture
        
        'on dessine donc maintenant le bord
        If bDisplayBorder Then
            If bBreakCorner = False Then
                'alors c'est un rectangle
                
                'on défini un brush
                hBrush = CreateSolidBrush(lBorderColor)
                
                'on définit une zone rectangulaire à bords arrondi
                hRgn = CreateRectRgn(0, 0, UserControl.ScaleWidth / 15, _
                    UserControl.ScaleHeight / 15)
                
                'on dessine le contour
                Call FrameRgn(UserControl.hdc, hRgn, hBrush, lBWidth, lBWidth)
    
                'on détruit le brush et la zone
                Call DeleteObject(hBrush)
                Call DeleteObject(hRgn)
                
            Else
                'alors c'est un arrondi
                
                'on défini un brush
                hBrush = CreateSolidBrush(lBorderColor)
                
                'on définit une zone rectangulaire à bords arrondi
                hRgn = CreateRoundRectRgn(0, 0, UserControl.ScaleWidth / 15, _
                    UserControl.ScaleHeight / 15, lCornerSize, lCornerSize)
                
                'on dessine le contour
                Call FrameRgn(UserControl.hdc, hRgn, hBrush, lBWidth, lBWidth)
                
                'on défini la zone rectangulaire arrondi comme nouvelle fenêtre
                Call SetWindowRgn(UserControl.hWnd, hRgn, True)
    
                'on détruit le brush et la zone
                Call DeleteObject(hBrush)
                Call DeleteObject(hRgn)
    
            End If
        End If
        
        UserControl.MaskPicture = UserControl.Image
    End If

    '//créé le gradient du reste du contrôle
    If lBackStyle = Opaque Then
        
        UserControl.BackStyle = 1
        
        If bShowBackGround Then
    
            'récupère les 3 composantes des deux couleurs
            Call ToRGB(bCol1, RGB1)
            Call ToRGB(bCol2, RGB2)
            
            If bShowTitle Then
                Dep = lTitleHeight / 15 'commence le gradient de pas tout en haut
            Else
                Dep = 0 'commence de tout en haut le gradient
            End If
            
            If lBackGradient = None Then
                'pas de gradient
                'on dessine alors un rectangle
                UserControl.Line (15, lTitleHeight + 1)-(UserControl.ScaleWidth - 30, _
                    UserControl.ScaleHeight - 30), bCol1, BF
            ElseIf lBackGradient = Horizontal Then
                'gradient horizontal
                Call FillGradientH(RGB1, RGB2, UserControl.ScaleWidth, _
                    UserControl.ScaleHeight, Dep)
            Else
                'gradient vertical
                Call FillGradientW(RGB1, RGB2, UserControl.ScaleWidth, _
                    UserControl.ScaleHeight, Dep)
            End If
        End If
'    Else
'        'on récupère le maskpicture et on met le controle en transparent
'
'        UserControl.BackStyle = 0
'        UserControl.Picture = UserControl.MaskPicture
'
    End If
    
    
    
    '//on va se tracer la bitmapt maintenant ^^
    If bPic Then
        'on montre l'image présente
        
        'on choisit celle à afficher (Grise ou pas)
        If bEnable Or Not (bGray) And bShowTitle Then
            With PCTcolor
                'on move IMG en fonction du choix de Aligment
                'on centre l'image dans la barre de titre
                If pctAlign = [Left Justify] Then
                    'à gauche
                    
                    W = (lTitleHeight - .Height) / 2 + lOffsetY
                    H = 30 + lCornerSize * 2 + lOffsetX
                    If W < 0 Then W = 0
                    If H < 0 Then H = 0
                    
                    .Top = W
                    .Left = H '2 pxls + arrondi
                    
                Else
                    'à droite
                    
                    W = (lTitleHeight - .Height) / 2 + lOffsetY
                    H = UserControl.Width - 30 - .Width - lCornerSize * 2 - lOffsetX
                    If H < 0 Then H = 0
                    If W < 0 Then W = 0
        
                    .Top = W
                    .Left = H '2 pxls + arrondi
                    
                End If
                    
                .Visible = True
                PCTgray.Visible = False
            End With
        ElseIf bShowTitle Then
            With PCTgray
                'on move IMG en fonction du choix de Aligment
                'on centre l'image dans la barre de titre
                If pctAlign = [Left Justify] Then
                    'à gauche
                    
                    W = (lTitleHeight - .Height) / 2 + lOffsetY
                    H = 30 + lCornerSize * 2 + lOffsetX
                    If W < 0 Then W = 0
                    If H < 0 Then H = 0
                    
                    .Top = W
                    .Left = H '2 pxls + arrondi
                    
                Else
                    'à droite
                    
                    W = (lTitleHeight - .Height) / 2 + lOffsetY
                    H = UserControl.Width - 30 - .Width - lCornerSize * 2 - lOffsetX
                    If H < 0 Then H = 0
                    If W < 0 Then W = 0
        
                    .Top = W
                    .Left = H '2 pxls + arrondi
                    
                End If
                    
                .Visible = True
                PCTcolor.Visible = False
            End With
        End If
    Else
        'on masque l'image
        PCTgray.Visible = False
        PCTcolor.Visible = False
    End If
    
    
    '//on trace la bordure si opaque
    If lBackStyle = Opaque And bDisplayBorder Then
    
        If bBreakCorner = False Then
            'alors c'est un rectangle
            
            'on défini un brush
            hBrush = CreateSolidBrush(lBorderColor)
            
            'on définit une zone rectangulaire à bords arrondi
            hRgn = CreateRectRgn(0, 0, UserControl.ScaleWidth / 15, _
                UserControl.ScaleHeight / 15)
            
            'on dessine le contour
            Call FrameRgn(UserControl.hdc, hRgn, hBrush, lBWidth, lBWidth)

            'on détruit le brush et la zone
            Call DeleteObject(hBrush)
            Call DeleteObject(hRgn)
        
        Else
            'alors c'est un arrondi
            
            'on défini un brush
            hBrush = CreateSolidBrush(lBorderColor)
            
            'on définit une zone rectangulaire à bords arrondi
            hRgn = CreateRoundRectRgn(0, 0, UserControl.ScaleWidth / 15, _
                UserControl.ScaleHeight / 15, lCornerSize, lCornerSize)
            
            'on dessine le contour
            Call FrameRgn(UserControl.hdc, hRgn, hBrush, lBWidth, lBWidth)
            
            'on défini la zone rectangulaire arrondi comme nouvelle fenêtre
            Call SetWindowRgn(UserControl.hWnd, hRgn, True)

            'on détruit le brush et la zone
            Call DeleteObject(hBrush)
            Call DeleteObject(hRgn)
            
        End If
        
    End If
    
    bNotOk = True

    '//on délocke le controle --> a évité les clignotements
    Call LockWindowUpdate(0)
    
    'permet de refresh la bordure
    Call UserControl.Refresh
End Sub
