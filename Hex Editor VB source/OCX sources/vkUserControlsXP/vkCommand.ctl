VERSION 5.00
Begin VB.UserControl vkCommand 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   PropertyPages   =   "vkCommand.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "vkCommand.ctx":004D
   Begin VB.PictureBox pctG 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1800
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image PCTgray 
      Height          =   240
      Left            =   1920
      Top             =   1800
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image PCTcolor 
      Height          =   240
      Left            =   2280
      Top             =   1800
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "vkCommand"
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
'VARIABLES PRIVEES
'=======================================================
Private mAsm(63) As Byte    'contient le code ASM
Private OldProc As Long     'adresse de l'ancienne window proc
Private objHwnd As Long     'handle de l'objet concerné
Private ET As TRACKMOUSEEVENTTYPE   'type pour le mouse_hover et le mouse_leave
Private IsMouseIn As Boolean    'si la souris est dans le controle

Private bPushed As Boolean
Private lTextPos As AlignmentConstants
Private lForeColor As OLE_COLOR
Private bCol1 As OLE_COLOR
Private bCol2 As OLE_COLOR
Private tCol1 As OLE_COLOR
Private tCol2 As OLE_COLOR
Private sCaption As String
Private lGradient As GradientConstants
Private bNotOk As Boolean
Private bNotOk2 As Boolean
Private bEnable As Boolean
Private lBorderColor As OLE_COLOR
Private bBreakCorner As Boolean
Private pctAlign As PictureAlignment
Private bPic As Boolean
Private lOffsetX As Long
Private lOffsetY As Long
Private bGray As Boolean
Private bDrawFocus As Boolean
Private bDrawMouseInRect As Boolean
Private bHasFocus As Boolean
Private lNotEnabledColor As OLE_COLOR
Private bUnRefreshControl As Boolean
Private bHasLeftOneTime As Boolean


'=======================================================
'EVENTS
'=======================================================
Public Event Click()
Attribute Click.VB_Description = "Happens when control gets a click (leftbutton)"
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Happens when a key is down"
Public Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Happens when a key is pressed"
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Happens when a key is up"
Public Event MouseHover()
Attribute MouseHover.VB_Description = "Happens when mouse enters control"
Public Event MouseLeave()
Attribute MouseLeave.VB_Description = "Happens when mouse leaves control"
Public Event MouseWheel(Sens As Wheel_Sens)
Attribute MouseWheel.VB_Description = "Happens when control gets a wheel"
Public Event MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
Attribute MouseDown.VB_Description = "Happens when control gets a click"
Public Event MouseUp(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
Attribute MouseUp.VB_Description = "Happens when control gets a mouseup"
Public Event MouseDblClick(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
Attribute MouseDblClick.VB_Description = "Happens when control gets a dblclick"
Public Event MouseMove(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
Attribute MouseMove.VB_Description = "Happens when mouse moves on control"




'=======================================================
'USERCONTROL SUBS
'=======================================================
'=======================================================
' /!\ NE PAS DEPLACER CETTE FONCTION /!\ '
'=======================================================
' Cette fonction doit rester la premiere '
' fonction "public" du module de classe  '
'=======================================================
Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Attribute WindowProc.VB_Description = "Internal proc for subclassing"
Attribute WindowProc.VB_MemberFlags = "40"
Dim iControl As Integer
Dim iShift As Integer
Dim z As Long
Dim x As Long
Dim y As Long

    Select Case uMsg
        
        Case WM_LBUTTONDBLCLK
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * 15
                y = HiWord(lParam) * 15
                
                bPushed = True: Refresh: RaiseEvent Click
                RaiseEvent MouseDblClick(vbLeftButton, iShift, iControl, x, y)
        Case WM_LBUTTONDOWN
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * 15
                y = HiWord(lParam) * 15
                
                bPushed = True: Refresh: RaiseEvent Click
                RaiseEvent MouseDown(vbLeftButton, iShift, iControl, x, y)
        Case WM_LBUTTONUP
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * 15
                y = HiWord(lParam) * 15
                
                bPushed = False: Refresh
                If bPushed And bEnable Then RaiseEvent Click
                Call DrawMouseEnterRect
                RaiseEvent MouseUp(vbLeftButton, iShift, iControl, x, y)
        Case WM_MBUTTONDBLCLK
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * 15
                y = HiWord(lParam) * 15
                
                RaiseEvent MouseDblClick(vbMiddleButton, iShift, iControl, x, y)
        Case WM_MBUTTONDOWN
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * 15
                y = HiWord(lParam) * 15
                
                RaiseEvent MouseDown(vbMiddleButton, iShift, iControl, x, y)
        Case WM_MBUTTONUP
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * 15
                y = HiWord(lParam) * 15
                
                RaiseEvent MouseUp(vbMiddleButton, iShift, iControl, x, y)
        Case WM_MOUSEHOVER
            If IsMouseIn = False Then
                Call DrawMouseEnterRect
                RaiseEvent MouseHover
                IsMouseIn = True
            End If
        Case WM_MOUSELEAVE
            RaiseEvent MouseLeave
            bPushed = False
            IsMouseIn = False
            bNotOk = False
            If bHasLeftOneTime Then
                Call Refresh(False): Call DrawFocusRects
            Else
                bHasLeftOneTime = True
            End If
        Case WM_MOUSEMOVE
            Call TrackMouseEvent(ET)
            
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * 15
                y = HiWord(lParam) * 15
    
                If (wParam And MK_LBUTTON) = MK_LBUTTON Then z = vbLeftButton
                If (wParam And MK_RBUTTON) = MK_RBUTTON Then z = vbRightButton
                If (wParam And MK_MBUTTON) = MK_MBUTTON Then z = vbMiddleButton
                RaiseEvent MouseMove(z, iShift, iControl, x, y)
        Case WM_RBUTTONDBLCLK
                        iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * 15
                y = HiWord(lParam) * 15
                
                RaiseEvent MouseDblClick(vbRightButton, iShift, iControl, x, y)
        Case WM_RBUTTONDOWN
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * 15
                y = HiWord(lParam) * 15
                
                RaiseEvent MouseDown(vbRightButton, iShift, iControl, x, y)
        Case WM_RBUTTONUP
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * 15
                y = HiWord(lParam) * 15
                
                RaiseEvent MouseUp(vbRightButton, iShift, iControl, x, y)
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
        .BackColor1 = &HFBFBFB       '
        .BackColor2 = &HDCDCDC    '
        .BackColorPushed1 = &HC8C8C8          '
        .BackColorPushed2 = &HEBEBEB       '
        .BackGradient = Horizontal '
        .Caption = "Caption" '
        .Font = Ambient.Font '
        .ForeColor = 7552000 '
        .TextPosition = vbCenter '
        .Enabled = True '
        .BorderColor = 7552000    '
        .BreakCorner = True '
        Set .Picture = Nothing
        .PictureAlignment = [Left Justify]
        .DisplayPicture = True
        .PictureOffsetX = 0
        .PictureOffsetY = 0
        .GrayPictureWhenDisabled = True
        .DrawFocus = True
        .DrawMouseInRect = True
        .DisabledBackColor = 15198183
        .UnRefreshControl = False
    End With
    bNotOk2 = False
    Call UserControl_Paint  'refresh
End Sub

Private Sub UserControl_GotFocus()

    If bEnable = False Then
        'on ne garde pas le focus
        Call SendKeys("{Tab}")
        Exit Sub
    End If
    
    'on a alors le focus
    bHasFocus = True
    
    'trace le rectangle de focus
    Call DrawFocusRects
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp, vbKeyLeft:
            Call SendKeys("+{Tab}")
        Case vbKeyDown, vbKeyRight:
            Call SendKeys("{Tab}")
        Case vbKeySpace
            If bPushed = False Then
                bPushed = True: Refresh: RaiseEvent Click
            End If
        Case vbKeyReturn
            RaiseEvent Click
    End Select
    
    'Call Refresh
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
    If KeyCode = vbKeySpace Then bPushed = False: Refresh: DrawMouseEnterRect
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
        Call .WriteProperty("UnRefreshControl", Me.UnRefreshControl, False)
        Call .WriteProperty("BackColor1", Me.BackColor1, &HFBFBFB)
        Call .WriteProperty("BackColor2", Me.BackColor2, &HDCDCDC)
        Call .WriteProperty("BackColorPushed1", Me.BackColor1, &HC8C8C8)
        Call .WriteProperty("BackColorPushed2", Me.BackColor2, &HEBEBEB)
        Call .WriteProperty("BackGradient", Me.BackGradient, Horizontal)
        Call .WriteProperty("Caption", Me.Caption, "Caption")
        Call .WriteProperty("Font", Me.Font, Ambient.Font)
        Call .WriteProperty("ForeColor", Me.ForeColor, 7552000)
        Call .WriteProperty("TextPosition", Me.TextPosition, vbCenter)
        Call .WriteProperty("Enabled", Me.Enabled, True)
        Call .WriteProperty("BorderColor", Me.BorderColor, 7552000)
        Call .WriteProperty("BreakCorner", Me.BreakCorner, True)
        Call .WriteProperty("Picture", Me.Picture, Nothing)
        Call .WriteProperty("PictureAlignment", Me.PictureAlignment, [Left Justify])
        Call .WriteProperty("DisplayPicture", Me.DisplayPicture, True)
        Call .WriteProperty("PictureOffsetX", Me.PictureOffsetX, 0)
        Call .WriteProperty("PictureOffsetY", Me.PictureOffsetY, 0)
        Call .WriteProperty("GrayPictureWhenDisabled", Me.GrayPictureWhenDisabled, True)
        Call .WriteProperty("DrawFocus", Me.DrawFocus, True)
        Call .WriteProperty("DrawMouseInRect", Me.DrawMouseInRect, True)
        Call .WriteProperty("DisabledBackColor", Me.DisabledBackColor, 15198183)
    End With
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    bNotOk2 = True
    With PropBag
        Me.BackColor1 = .ReadProperty("BackColor1", &HFBFBFB)
        Me.BackColor2 = .ReadProperty("BackColor2", &HDCDCDC)
        Me.BackColorPushed1 = .ReadProperty("BackColor1", &HC8C8C8)
        Me.BackColorPushed2 = .ReadProperty("BackColor2", &HEBEBEB)
        Me.BackGradient = .ReadProperty("BackGradient", Horizontal)
        Me.Caption = .ReadProperty("Caption", "Caption")
        Set Me.Font = .ReadProperty("Font", Ambient.Font)
        Me.ForeColor = .ReadProperty("ForeColor", 7552000)
        Me.TextPosition = .ReadProperty("TextPosition", vbCenter)
        Me.Enabled = .ReadProperty("Enabled", True)
        Me.BorderColor = .ReadProperty("BorderColor", 7552000)
        Me.BreakCorner = .ReadProperty("BreakCorner", True)
        Set Me.Picture = .ReadProperty("Picture", Nothing)
        Me.PictureAlignment = .ReadProperty("PictureAlignment", [Left Justify])
        Me.DisplayPicture = .ReadProperty("DisplayPicture", True)
        Me.PictureOffsetX = .ReadProperty("PictureOffsetX", 0)
        Me.PictureOffsetY = .ReadProperty("PictureOffsetY", 0)
        Me.GrayPictureWhenDisabled = .ReadProperty("GrayPictureWhenDisabled", True)
        Me.DrawFocus = .ReadProperty("DrawFocus", True)
        Me.DrawMouseInRect = .ReadProperty("DrawMouseInRect", True)
        Me.DisabledBackColor = .ReadProperty("DisabledBackColor", 15198183)
        Me.UnRefreshControl = .ReadProperty("UnRefreshControl", False)
    End With
    bNotOk2 = False
    'Call UserControl_Paint  'refresh
    
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
Private Sub LaunchKeyMouseEvents()
                
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
Public Property Get hDc() As Long: hDc = UserControl.hDc: End Property
Attribute hDc.VB_Description = "Get the control hDc"
Public Property Get hWnd() As Long: hWnd = UserControl.hWnd: End Property
Attribute hWnd.VB_Description = "Handle of the control"
Public Property Get TextPosition() As AlignmentConstants: TextPosition = lTextPos: End Property
Attribute TextPosition.VB_Description = "Text position"
Public Property Let TextPosition(TextPosition As AlignmentConstants): lTextPos = TextPosition: bNotOk = False: UserControl_Paint: End Property
Public Property Get Caption() As String: Caption = sCaption: End Property
Attribute Caption.VB_Description = "Text to display"
Public Property Let Caption(Caption As String): sCaption = Caption: bNotOk = False: UserControl_Paint: bNotOk = True: End Property
Public Property Get ForeColor() As OLE_COLOR: ForeColor = lForeColor: End Property
Attribute ForeColor.VB_Description = "Text color"
Public Property Let ForeColor(ForeColor As OLE_COLOR): lForeColor = ForeColor: UserControl.ForeColor = ForeColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get BackColor1() As OLE_COLOR: BackColor1 = bCol1: End Property
Attribute BackColor1.VB_Description = "Color1 of the back gradient"
Public Property Let BackColor1(BackColor1 As OLE_COLOR): bCol1 = BackColor1: bNotOk = False: UserControl_Paint: End Property
Public Property Get BackColor2() As OLE_COLOR: BackColor2 = bCol2: End Property
Attribute BackColor2.VB_Description = "Color2 of the back gradient"
Public Property Let BackColor2(BackColor2 As OLE_COLOR): bCol2 = BackColor2: bNotOk = False: UserControl_Paint: End Property
Public Property Get BackColorPushed1() As OLE_COLOR: BackColorPushed1 = tCol1: End Property
Attribute BackColorPushed1.VB_Description = "Color1 of the back gradient when command is pushed"
Public Property Let BackColorPushed1(BackColorPushed1 As OLE_COLOR): tCol1 = BackColorPushed1: bNotOk = False: UserControl_Paint: End Property
Public Property Get BackColorPushed2() As OLE_COLOR: BackColorPushed2 = tCol2: End Property
Attribute BackColorPushed2.VB_Description = "Color2 of the back gradient when command is pushed"
Public Property Let BackColorPushed2(BackColorPushed2 As OLE_COLOR): tCol2 = BackColorPushed2: bNotOk = False: UserControl_Paint: End Property
Public Property Get Font() As StdFont: Set Font = UserControl.Font: End Property
Attribute Font.VB_Description = "Text font"
Public Property Set Font(Font As StdFont): Set UserControl.Font = Font: bNotOk = False: UserControl_Paint: End Property
Public Property Get BackGradient() As GradientConstants: BackGradient = lGradient: End Property
Attribute BackGradient.VB_Description = "Type of back gradient"
Public Property Let BackGradient(BackGradient As GradientConstants): lGradient = BackGradient: bNotOk = False: UserControl_Paint: End Property
Public Property Get Enabled() As Boolean: Enabled = bEnable: End Property
Attribute Enabled.VB_Description = "Enable or not control"
Public Property Let Enabled(Enabled As Boolean)
bEnable = Enabled: bNotOk = False: UserControl_Paint
End Property
Public Property Get BorderColor() As OLE_COLOR: BorderColor = lBorderColor: End Property
Attribute BorderColor.VB_Description = "Color of the border"
Public Property Let BorderColor(BorderColor As OLE_COLOR): lBorderColor = BorderColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get BreakCorner() As Boolean: BreakCorner = bBreakCorner: End Property
Attribute BreakCorner.VB_Description = "Use round corner or not"
Public Property Let BreakCorner(BreakCorner As Boolean): bBreakCorner = BreakCorner: bNotOk = False: UserControl_Paint: End Property
Public Property Get Picture() As Picture: Set Picture = PCTcolor.Picture: End Property
Attribute Picture.VB_Description = "Picture to display"
Public Property Set Picture(NewPic As Picture)
Set PCTcolor.Picture = NewPic
Set pctG.Picture = NewPic
If Not (NewPic Is Nothing) Then
    pctG.Width = PCTcolor.Width
    pctG.Height = PCTcolor.Height
    Call GrayScale(pctG)
End If
PCTgray.Picture = pctG.Image
bNotOk = False: UserControl_Paint
End Property
Public Property Get PictureAlignment() As PictureAlignment: PictureAlignment = pctAlign: End Property
Attribute PictureAlignment.VB_Description = "Alignment of the picture"
Public Property Let PictureAlignment(PictureAlignment As PictureAlignment): pctAlign = PictureAlignment: bNotOk = False: UserControl_Paint: End Property
Public Property Get DisplayPicture() As Boolean: DisplayPicture = bPic: End Property
Attribute DisplayPicture.VB_Description = "Display a picture or not"
Public Property Let DisplayPicture(DisplayPicture As Boolean): bPic = DisplayPicture: bNotOk = False: UserControl_Paint: End Property
Public Property Get PictureOffsetX() As Long: PictureOffsetX = lOffsetX: End Property
Attribute PictureOffsetX.VB_Description = "Offset (twips) of picture"
Public Property Let PictureOffsetX(PictureOffsetX As Long): lOffsetX = PictureOffsetX: bNotOk = False: UserControl_Paint: End Property
Public Property Get PictureOffsetY() As Long: PictureOffsetY = lOffsetY: End Property
Attribute PictureOffsetY.VB_Description = "Offset (twips) of picture"
Public Property Let PictureOffsetY(PictureOffsetY As Long): lOffsetY = PictureOffsetY: bNotOk = False: UserControl_Paint: End Property
Public Property Get GrayPictureWhenDisabled() As Boolean: GrayPictureWhenDisabled = bGray: End Property
Attribute GrayPictureWhenDisabled.VB_Description = "Use a gray picture (or color picture) when control is not enabled"
Public Property Let GrayPictureWhenDisabled(GrayPictureWhenDisabled As Boolean): bGray = GrayPictureWhenDisabled: bNotOk = False: UserControl_Paint: End Property
Public Property Get DrawFocus() As Boolean: DrawFocus = bDrawFocus: End Property
Attribute DrawFocus.VB_Description = "Draw focus when control has focus"
Public Property Let DrawFocus(DrawFocus As Boolean): bDrawFocus = DrawFocus: bNotOk = False: UserControl_Paint: End Property
Public Property Get DrawMouseInRect() As Boolean: DrawMouseInRect = bDrawMouseInRect: End Property
Attribute DrawMouseInRect.VB_Description = "Draw focus when mouse is hover control"
Public Property Let DrawMouseInRect(DrawMouseInRect As Boolean): bDrawMouseInRect = DrawMouseInRect: bNotOk = False: UserControl_Paint: End Property
Public Property Get DisabledBackColor() As OLE_COLOR: DisabledBackColor = lNotEnabledColor: End Property
Attribute DisabledBackColor.VB_Description = "Color of then back color when control is not enabled"
Public Property Let DisabledBackColor(DisabledBackColor As OLE_COLOR): lNotEnabledColor = DisabledBackColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get UnRefreshControl() As Boolean: UnRefreshControl = bUnRefreshControl: End Property
Attribute UnRefreshControl.VB_Description = "Prevent to refresh control"
Public Property Let UnRefreshControl(UnRefreshControl As Boolean): bUnRefreshControl = UnRefreshControl: End Property


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
        Call MoveToEx(.hDc, 0, Dep, 0&)
        
        'pour chaque 'colonne' constituée par une ligne verticale, on trace une
        'ligne en récupérant la couleur correspondante
        For x = 0 To lWidth
            
            'change le ForeColor qui détermine la couleur de la Line
            'multiplie la largeur actuelle par la couleur par unité de longueur
            .ForeColor = RGB(LeftColor.R + x * rAverageColorPerSizeUnit * lSigne, LeftColor.G + x * _
                gAverageColorPerSizeUnit * lSigne, LeftColor.B + x * bAverageColorPerSizeUnit * lSigne)
               
            'trace une ligne
            Call LineTo(.hDc, x, lHeight)
            
            'bouge 'd'une colonne' vers la droite
            Call MoveToEx(.hDc, x, Dep, 0&)
        
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
        Call MoveToEx(.hDc, 0, Dep, 0&)
        
        'pour chaque 'colonne' constituée par une ligne verticale, on trace une
        'ligne en récupérant la couleur correspondante
        For x = Dep To lHeight
            
            'change le ForeColor qui détermine la couleur de la Line
            'multiplie la largeur actuelle par la couleur par unité de longueur
            .ForeColor = RGB(LeftColor.R + x * rAverageColorPerSizeUnit * lSigne, LeftColor.G + x * _
                gAverageColorPerSizeUnit * lSigne, LeftColor.B + x * bAverageColorPerSizeUnit * lSigne)
               
            'trace une ligne
            Call LineTo(.hDc, Width, x)
            
            'bouge 'd'une colonne' vers la droite
            Call MoveToEx(.hDc, 0, x, 0&)
        
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
    Res = GetTabbedTextExtent(UserControl.hDc, "A", 1, 0, 0)
    GetCharHeight = (Res And &HFFFF0000) \ &H10000
End Function


'=======================================================
'PUBLIC SUB
'=======================================================
'=======================================================
'on dessine tout
'=======================================================
Public Sub Refresh(Optional ByVal ShowFocusRects As Boolean = True)
Attribute Refresh.VB_Description = "Refresh the control"
Dim x As Long
Dim RGB1 As RGB_COLOR
Dim RGB2 As RGB_COLOR
Dim R As RECT
Dim hBrush As Long
Dim Rec As Long
Dim hRgn As Long
Dim W As Long
Dim H As Long
    
    If bUnRefreshControl Then Exit Sub
    
    
    '//on efface et on vire le maskpicture
    Call UserControl.Cls
    UserControl.Picture = Nothing
    UserControl.MaskPicture = Nothing
    
    '//on convertir les différentes couleurs si couleurs système
    Call OleTranslateColor(lBorderColor, 0, lBorderColor)
    Call OleTranslateColor(lForeColor, 0, lForeColor)
    Call OleTranslateColor(bCol1, 0, bCol1)
    Call OleTranslateColor(bCol2, 0, bCol2)
    
    
    '//on va tracer le rectangle de focus si on a le focus
    If bHasFocus And ShowFocusRects Then Call UserControl_GotFocus
    
    
    '//créé le gradient du contrôle
    If bEnable Then
        
        If bPushed = False Then
            'pas appuyé
            
            'récupère les 3 composantes des deux couleurs
            Call OleTranslateColor(bCol1, 0, bCol1)
            Call OleTranslateColor(bCol2, 0, bCol2)
            Call ToRGB(bCol1, RGB1)
            Call ToRGB(bCol2, RGB2)
            
            If lGradient = None Then
                'pas de gradient
                'on dessine alors un rectangle
                Line (15, 0)-(ScaleWidth - 30, _
                    ScaleHeight - 30), bCol1, BF
            ElseIf lGradient = Horizontal Then
                'gradient horizontal
                Call FillGradientH(RGB1, RGB2, ScaleWidth, _
                    ScaleHeight, 0)
            Else
                'gradient vertical
                Call FillGradientW(RGB1, RGB2, ScaleWidth, _
                    ScaleHeight, 0)
            End If
        Else
            'appuyé sur le bouton
            
            'récupère les 3 composantes des deux couleurs
            Call OleTranslateColor(tCol1, 0, tCol1)
            Call OleTranslateColor(tCol2, 0, tCol2)
            Call ToRGB(tCol2, RGB1)
            Call ToRGB(tCol1, RGB2)
            
            If lGradient = None Then
                'pas de gradient
                'on dessine alors un rectangle
                Line (15, 0)-(ScaleWidth - 30, _
                    ScaleHeight - 30), bCol2, BF
            ElseIf lGradient = Horizontal Then
                'gradient horizontal
                Call FillGradientH(RGB2, RGB1, ScaleWidth, _
                    ScaleHeight, 0)
            Else
                'gradient vertical
                Call FillGradientW(RGB2, RGB1, ScaleWidth, _
                    ScaleHeight, 0)
            End If
        End If
    Else
        'controle non enabled
        UserControl.BackColor = lNotEnabledColor
    End If
    

    '//créé un rectangle
    If bPic And PCTcolor.Picture Then
        'alors on affiche la picture ==> décale le texte
        If pctAlign = [Left Justify] Then
            Call SetRect(R, 8 + PCTcolor.Width / 30, (Height / 15 - GetCharHeight) / 2, Width / 15 - 4, Height / 15)
        Else
            Call SetRect(R, 4, (Height / 15 - GetCharHeight) / 2, Width / 15 - 8 - PCTcolor.Width / 30, Height / 15)
        End If
    Else
        'pas de picture ==> pas de décalage
        Call SetRect(R, 4, (Height / 15 - GetCharHeight) / 2, Width / 15 - 4, Height / 15)
    End If
    
    'on affiche le caption
    If bEnable Then
        UserControl.ForeColor = lForeColor
    Else
        UserControl.ForeColor = 9934743
    End If
    If lTextPos = vbCenter Then
        'au centre
        Call DrawText(UserControl.hDc, sCaption, Len(sCaption), R, DT_CENTER)
    ElseIf lTextPos = vbRightJustify Then
        'à droite
        Call DrawText(UserControl.hDc, sCaption, Len(sCaption), R, DT_RIGHT)
    Else
        'à gauche
        Call DrawText(UserControl.hDc, sCaption, Len(sCaption), R, DT_LEFT)
    End If
    
    
    
    '//on va se tracer la bitmap maintenant ^^
    If bPic And PCTcolor.Picture Then
        'on montre l'image présente
        
        'on choisit celle à afficher (Grise ou pas)
        If bEnable Or Not (bGray) Then
            With PCTcolor
                'on move IMG en fonction du choix de Aligment
                'on centre l'image dans la barre de titre
                If pctAlign = [Left Justify] Then
                    'à gauche
                    
                    W = (Height - .Height) / 2 + lOffsetY
                    H = (Width - TextWidth(sCaption)) / 2 - .Width + lOffsetX - 4
                    If W < 0 Then W = 0
                    If H < 0 Then H = 0
                    
                    .Top = W
                    .Left = H '2 pxls + arrondi
                    
                Else
                    'à droite
                    
                    W = (Height - .Height) / 2 + lOffsetY
                    H = (Width - TextWidth(sCaption)) / 2 + TextWidth(sCaption) / 2 + .Width + lOffsetX + 4
                    If H < 0 Then H = 0
                    If W < 0 Then W = 0
        
                    .Top = W
                    .Left = H '2 pxls + arrondi
                    
                End If
                    
                .Visible = True
                PCTgray.Visible = False
            End With
        Else
            With PCTgray
                'on move IMG en fonction du choix de Aligment
                'on centre l'image dans la barre de titre
                If pctAlign = [Left Justify] Then
                    'à gauche
                    
                    W = (Height - .Height) / 2 + lOffsetY
                    H = (Width - TextWidth(sCaption)) / 2 - .Width + lOffsetX
                    If W < 0 Then W = 0
                    If H < 0 Then H = 0
                    
                    .Top = W
                    .Left = H '2 pxls + arrondi
                    
                Else
                    'à droite
                    
                    W = (Height - .Height) / 2 + lOffsetY
                    H = (Width - TextWidth(sCaption)) / 2 + .Width + lOffsetX
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
    
    
    
    '//dessine le contour
    If bBreakCorner = False Then
        'alors c'est un rectangle
        
        'on défini un brush
        hBrush = CreateSolidBrush(lBorderColor)
        
        'on définit une zone rectangulaire à bords arrondi
        hRgn = CreateRectRgn(0, 0, ScaleWidth / 15, _
            ScaleHeight / 15)
        
        'on dessine le contour
        Call FrameRgn(UserControl.hDc, hRgn, hBrush, 1, 1)

        'on détruit le brush et la zone
        Call DeleteObject(hBrush)
        Call DeleteObject(hRgn)
    
    Else
        'alors c'est un arrondi
        
        'on défini un brush
        hBrush = CreateSolidBrush(lBorderColor)
        
        'on définit une zone rectangulaire à bords arrondi
        hRgn = CreateRoundRectRgn(0, 0, ScaleWidth / 15, _
            ScaleHeight / 15, 7, 7)
        
        'on dessine le contour
        Call FrameRgn(UserControl.hDc, hRgn, hBrush, 1, 1)
        
        'on défini la zone rectangulaire arrondi comme nouvelle fenêtre
        Call SetWindowRgn(UserControl.hWnd, hRgn, True)

        'on détruit le brush et la zone
        Call DeleteObject(hBrush)
        Call DeleteObject(hRgn)
        
    End If

    
    bNotOk = True

    '//on délocke le controle --> a évité les clignotements
   ' Call LockWindowUpdate(0)
    
    'permet de refresh la bordure
    Call UserControl.Refresh
End Sub

'=======================================================
'trace le rectangle d'entrée sur le controle
'=======================================================
Private Sub DrawMouseEnterRect()
Dim x As Long
Dim W As Long
Dim H As Long
    
    'couleur orange foncée : 38631 et 3257087
    'couleur orange clair : 7064575
    'couleur orange très clair : 9033981,9889535
    
    If bDrawMouseInRect = False Or bEnable = False Then Exit Sub
    
    With UserControl
        H = .Height
        W = .Width
    End With
    
    
    '//trace les lignes orange foncé
    UserControl.ForeColor = 3257087
    Line (15, 60)-(15, H - 60)
    Line (W - 45, 60)-(W - 45, H - 60)
    Line (30, H - 60)-(W - 45, H - 60)
    Line (45, H - 45)-(W - 60, H - 45), 38631
    
    
    '//trace les lignes orange clair
    UserControl.ForeColor = 7064575
    Line (30, 30)-(30, H - 60)
    Line (W - 60, 30)-(W - 60, H - 60)
        

    '//trace la ligne orange très claire
    Line (30, 30)-(W - 45, 30), 9033981
    Line (60, 15)-(W - 75, 15), 9889535
    
    Call UserControl.Refresh

End Sub

'=======================================================
'trace le rectangle de focus
'=======================================================
Private Sub DrawFocusRects()
Dim x As Long
Dim W As Long
Dim H As Long
    
    'couleur bleu foncée : 15696491 et 15183500
    'couleur bleu clair : 15782325
    'couleur bleu très clair : 16242621,16771022
    
    If bDrawFocus = False Or IsMouseIn Or bHasFocus = False Or bEnable = False Then Exit Sub
    
    
    With UserControl
        H = .Height
        W = .Width
    End With
    
    
    '//trace les lignes bleu foncé
    UserControl.ForeColor = 15183500
    Line (15, 60)-(15, H - 60)
    Line (W - 45, 60)-(W - 45, H - 60)
    Line (30, H - 60)-(W - 45, H - 60)
    Line (45, H - 45)-(W - 60, H - 45), 15696491
    
    
    '//trace les lignes bleu clair
    UserControl.ForeColor = 15782325
    Line (30, 30)-(30, H - 60)
    Line (W - 60, 30)-(W - 60, H - 60)
        

    '//trace la ligne bleu très claire
    Line (30, 30)-(W - 45, 30), 16242621
    Line (60, 15)-(W - 75, 15), 16771022
    
    Call UserControl.Refresh

End Sub

'=======================================================
'renvoie l'objet extender de ce usercontrol (pour les propertypages)
'=======================================================
Friend Property Get MyExtender() As Object
    Set MyExtender = UserControl.Extender
End Property
Friend Property Let MyExtender(MyExtender As Object)
    Set UserControl.Extender = MyExtender
End Property
