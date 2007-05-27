VERSION 5.00
Begin VB.UserControl vkListBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3780
   PropertyPages   =   "vkListBox.ctx":0000
   ScaleHeight     =   3210
   ScaleWidth      =   3780
   ToolboxBitmap   =   "vkListBox.ctx":002D
   Begin vkUserContolsXP.vkVScrollPrivate VS 
      Height          =   2295
      Left            =   2400
      TabIndex        =   6
      Top             =   480
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   4048
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   5
      Left            =   600
      Picture         =   "vkListBox.ctx":033F
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   4
      Left            =   240
      Picture         =   "vkListBox.ctx":0589
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   3
      Left            =   1320
      Picture         =   "vkListBox.ctx":07D3
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   2
      Left            =   960
      Picture         =   "vkListBox.ctx":0A1D
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   1
      Left            =   600
      Picture         =   "vkListBox.ctx":0C67
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   0
      Left            =   240
      Picture         =   "vkListBox.ctx":0EB1
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "vkListBox"
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

Private bDisplayBorder As Boolean
Private lBackColor As OLE_COLOR
Private bEnable As Boolean
Private lListCount As Long
Private lListIndex As Long
Private lHeight() As Long
Private bSelected() As Boolean
Private bChecked() As Boolean
Private bMultiSelect As Boolean
Private lNewIndex As Long
Private lSelCount As Long
Private bStyleCheckBox As Boolean
Private lTopIndex As Long
Private bNotOk As Boolean
Private bNotOk2 As Boolean
Private bUnRefreshControl As Boolean
Private lForeColor As OLE_COLOR
Private lBorderColor As OLE_COLOR
Private bSorted As Boolean
Private lCheckCount As Long
Private tAlig As AlignmentConstants
Private lSelColor As Long
Private lPrevSel As Long
Private vsPushed As Boolean
Private MouseItemIndex As Long
Private lFullRowSelect As Boolean
Private lBorderSelColor  As OLE_COLOR
Private tmpMouseItemIndex As Long
Private Col As clsFastCollection
Private zNumber As Long
Private bVSvisible As Boolean


'=======================================================
'EVENTS
'=======================================================
Public Event ItemClick(Item As vkListItem)
Public Event ItemChek(Item As vkListItem)
Public Event Scroll()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseHover()
Public Event MouseLeave()
Public Event MouseWheel(Sens As Wheel_Sens)
Public Event MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
Public Event MouseUp(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
Public Event MouseDblClick(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
Public Event MouseMove(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)




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
                
            RaiseEvent MouseDblClick(vbLeftButton, iShift, iControl, x, y)
        Case WM_LBUTTONDOWN
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * 15
                y = HiWord(lParam) * 15
                
                RaiseEvent MouseDown(vbLeftButton, iShift, iControl, x, y)
        Case WM_LBUTTONUP
                iShift = Abs((wParam And MK_SHIFT) = MK_SHIFT)
                iControl = Abs((wParam And MK_CONTROL) = MK_CONTROL)
                x = LoWord(lParam) * 15
                y = HiWord(lParam) * 15
                
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
                RaiseEvent MouseHover
                IsMouseIn = True
            End If
        Case WM_MOUSELEAVE
            RaiseEvent MouseLeave
            IsMouseIn = False
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

        Case WM_PAINT
            bNotOk = True  'évite le clignotement lors du survol de la souris
    End Select
    
    'appel de la routine standard pour les autres messages
    WindowProc = CallWindowProc(OldProc, hWnd, uMsg, wParam, lParam)
    
End Function

Private Sub UserControl_Initialize()
Dim Ofs As Long
Dim Ptr As Long
    
    ReDim bSelected(1)
    ReDim bChecked(1)
    ReDim lHeight(1)
    lListCount = 1
    
    'instancie la collection
    Set Col = New clsFastCollection
    
    'créé un controle dynamique
    'Set VS = Controls.Add("vkUserContolsXP.vkVScroll", "VS")
    'VS.MyExtender.Visible = True
    
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
    
    'initialise le VS
    With VS
        .Max = 2
        .Min = 1
        .Value = 1
        .Visible = True
    End With
End Sub

Private Sub UserControl_InitProperties()
    'valeurs par défaut
    bNotOk2 = True
    With Me
        .BackColor = &HFFFFFF
        .BorderColor = 12937777
        .DisplayBorder = True
        .Enabled = True
        .ForeColor = vbBlack
        .MultiSelect = True
        .Sorted = False
        .StyleCheckBox = False
        .UnRefreshControl = False
        .ListIndex = -1
        .DisplayVScroll = True
        .Alignment = vbLeftJustify
        .SelColor = 16768444
        .Font = Ambient.Font
        .FullRowSelect = True
        .BorderSelColor = 16419097
        .TopIndex = 1
    End With
    bNotOk2 = False
    Call UserControl_Paint  'refresh
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
Private Sub vkVScroll1_Scroll()
    RaiseEvent Scroll
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'sélection d'un item
Dim z As Long
Dim s As Long
Dim e As Long

    On Error Resume Next
    
    'si dans la zone gauche et que style=checkboxes ==> on checke
    If bStyleCheckBox Then
        If x < 255 Then
            bChecked(MouseItemIndex) = Not (bChecked(MouseItemIndex))
        End If
    End If

    'détermine quel Item est sélectionné
    s = 0   'hauteur temporaire
    For z = lTopIndex To lListCount
        s = s + lHeight(z)
        If s > y Then e = z: Exit For
    Next z
        
    If bMultiSelect = False Then
        'déselectionne tout
        Call UnSelectAll(False)
    Else
        'alors on teste en fonction du Shift
        If (Shift And vbShiftMask) = vbShiftMask Then
            'on sélectionne tout entre lPrevSel et e-1
            Dim o As Boolean
            If e - 1 > lPrevSel Then
                o = bSelected(lPrevSel)
            End If
            For s = e To lPrevSel Step IIf(e - 1 < lPrevSel, 1, -1)
                'Col.Item(s).Selected = True
                bSelected(s) = True
            Next s
            If e - 1 > lPrevSel Then
                'on supprime le premier(terme correctif)
                'Col.Item(lPrevSel).Selected = o
                bSelected(lPrevSel) = o
            End If
        ElseIf (Shift And vbCtrlMask) = vbCtrlMask Then
            'on permute le sélectionné et on touche pas au reste
            'Col.Item(e).Selected = Not (Col.Item(e).Selected)
            bSelected(e) = Not (bSelected(e))
            Call Refresh
            Exit Sub    'évite de revenir à selected(e)=true
        Else
            'déselectionne tout
            Call UnSelectAll(False)
        End If
    End If
        
    
    'alors si un élément est sélectionné
    If e Then
        'Col.Item(e).Selected = True
        bSelected(e) = True
        Call Refresh
    End If
    
    'sauvegarde le dernier Item sauvegardé
    lPrevSel = e - 1
        
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim z As Long
Dim z2 As Long
Dim z3 As Long
Dim e As Long
Dim m As Long
Dim f As Long
    
    If bStyleCheckBox = False Then Exit Sub
    
    On Error Resume Next

    'on détermine quel Item est survolé
    z2 = -1
    For f = lTopIndex To lListCount
        e = e + lHeight(f)
        If e > y Then
            If z2 = -1 Then z2 = z: m = e - lHeight(f)
        End If
        If e >= Height - 50 Then
            z3 = z
            Exit For
        End If
        z = z + 1
    Next f
    
    'si pas suffisemment d'items pour remplir la vue, alors le nombre d'affichés = listcount
    If z3 = 0 Then z3 = ListCount
    
    'récupère l'Item survolé
    MouseItemIndex = lTopIndex + z2
    
    'redessine les images si nécessaire (item survolé différent)
    If MouseItemIndex <> tmpMouseItemIndex Then Call SplitIMGandShow(z3)
    
    'sauvegarde les bornes (en height) de l'item survolé
    tmpMouseItemIndex = MouseItemIndex
    
End Sub

Private Sub UserControl_Terminate()
    'vire le subclassing
    If OldProc Then Call SetWindowLong(UserControl.hWnd, GWL_WNDPROC, OldProc)
    'on clear la collection
    Call Col.Clear
    Set Col = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Font", Me.Font, Ambient.Font)
        Call .WriteProperty("BackColor", Me.BackColor, &HFFFFFF)
        Call .WriteProperty("BorderColor", Me.BorderColor, 12937777)
        Call .WriteProperty("DisplayBorder", Me.DisplayBorder, True)
        Call .WriteProperty("Enabled", Me.Enabled, True)
        Call .WriteProperty("ForeColor", Me.ForeColor, vbBlack)
        Call .WriteProperty("MultiSelect", Me.MultiSelect, True)
        Call .WriteProperty("Sorted", Me.Sorted, True)
        Call .WriteProperty("StyleCheckBox", Me.StyleCheckBox, False)
        Call .WriteProperty("UnRefreshControl", Me.UnRefreshControl, False)
        Call .WriteProperty("ListIndex", Me.ListIndex, -1)
        Call .WriteProperty("DisplayVScroll", Me.DisplayVScroll, True)
        Call .WriteProperty("Alignment", Me.Alignment, vbLeftJustify)
        Call .WriteProperty("SelColor", Me.SelColor, 16768444)
        Call .WriteProperty("FullRowSelect", Me.FullRowSelect, True)
        Call .WriteProperty("BorderSelColor", Me.BorderSelColor, 16419097)
        Call .WriteProperty("TopIndex", Me.TopIndex, 1)
    End With
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    bNotOk2 = True
    With PropBag
        Me.BackColor = .ReadProperty("BackColor", &HFFFFFF)
        Me.BorderColor = .ReadProperty("BorderColor", 12937777)
        Me.DisplayBorder = .ReadProperty("DisplayBorder", True)
        Me.Enabled = .ReadProperty("Enabled", True)
        Me.ForeColor = .ReadProperty("ForeColor", vbBlack)
        Me.MultiSelect = .ReadProperty("MultiSelect", True)
        Me.Sorted = .ReadProperty("Sorted", True)
        Me.StyleCheckBox = .ReadProperty("StyleCheckBox", False)
        Set Me.Font = .ReadProperty("Font", Ambient.Font)
        Me.UnRefreshControl = .ReadProperty("UnRefreshControl", False)
        Me.ListIndex = .ReadProperty("ListIndex", -1)
        Me.DisplayVScroll = .ReadProperty("DisplayVScroll", True)
        Me.Alignment = .ReadProperty("Alignment", vbLeftJustify)
        Me.SelColor = .ReadProperty("SelColor", 16768444)
        Me.FullRowSelect = .ReadProperty("FullRowSelect", True)
        Me.BorderSelColor = .ReadProperty("BorderSelColor", 16419097)
        Me.TopIndex = .ReadProperty("TopIndex", 1)
    End With
    bNotOk2 = False
    'Call UserControl_Paint  'refresh
    
    'le bon endroit pour lancer le subclassing
    Call LaunchKeyMouseEvents
    If Ambient.UserMode Then
        Call VS.LaunchKeyMouseEvents    'subclasse également le VS
    End If
End Sub
Private Sub UserControl_Resize()
    If Height < 800 Then Height = 800
    With VS
        .Width = 255
        .Top = 0
        .Left = Width - 255
        .Height = Height
    End With
    Call Refresh  'refresh
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
Public Property Get hdc() As Long: hdc = UserControl.hdc: End Property
Public Property Get hWnd() As Long: hWnd = UserControl.hWnd: End Property
Public Property Get SelColor() As OLE_COLOR: SelColor = lSelColor: End Property
Public Property Let SelColor(SelColor As OLE_COLOR): lSelColor = SelColor: UserControl.ForeColor = ForeColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get ForeColor() As OLE_COLOR: ForeColor = lForeColor: End Property
Public Property Let ForeColor(ForeColor As OLE_COLOR): lForeColor = ForeColor: UserControl.ForeColor = ForeColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get BackColor() As OLE_COLOR: BackColor = lBackColor: End Property
Public Property Let BackColor(BackColor As OLE_COLOR): UserControl.BackColor = BackColor: lBackColor = BackColor: End Property
Public Property Get Font() As StdFont: Set Font = UserControl.Font: End Property
Public Property Set Font(Font As StdFont): Set UserControl.Font = Font: bNotOk = False: UserControl_Paint: End Property
Public Property Get Enabled() As Boolean: Enabled = bEnable: End Property
Public Property Let Enabled(Enabled As Boolean): bEnable = Enabled: bNotOk = False: UserControl_Paint: End Property
Public Property Get DisplayBorder() As Boolean: DisplayBorder = bDisplayBorder: End Property
Public Property Let DisplayBorder(DisplayBorder As Boolean): bDisplayBorder = DisplayBorder: bNotOk = False: UserControl_Paint: End Property
Public Property Get BorderColor() As OLE_COLOR: BorderColor = lBorderColor: End Property
Public Property Let BorderColor(BorderColor As OLE_COLOR): lBorderColor = BorderColor: bNotOk = False: UserControl_Paint: End Property
Public Property Get List(Index As Long) As String: On Error Resume Next: List = Col.Item(Index).Text: End Property
Public Property Let List(Index As Long, List As String): On Error Resume Next: Col.Item(Index).Text = List: bNotOk = False: UserControl_Paint: End Property
Public Property Get ListCount() As Long: ListCount = lListCount - 1: End Property
Public Property Get ListIndex() As Long: ListIndex = lListIndex: End Property
Public Property Let ListIndex(ListIndex As Long): lListIndex = ListIndex: bNotOk = False: UserControl_Paint: End Property
Public Property Get MultiSelect() As Boolean: MultiSelect = bMultiSelect: End Property
Public Property Let MultiSelect(MultiSelect As Boolean): bMultiSelect = MultiSelect: End Property
Public Property Get NewIndex() As Long: NewIndex = lNewIndex: End Property
Public Property Get SelCount() As Long: SelCount = lSelCount: End Property
Public Property Get Selected(Index As Long) As Boolean: On Error Resume Next: Selected = bSelected(Index): End Property
Public Property Get Checked(Index As Long) As Boolean: On Error Resume Next: Checked = bChecked(Index): End Property
Public Property Let Selected(Index As Long, Selected As Boolean): On Error Resume Next: bSelected(Index) = Selected: End Property
Public Property Let Checked(Index As Long, Checked As Boolean): On Error Resume Next: bChecked(Index) = Checked: End Property
Public Property Get Sorted() As Boolean: Sorted = bSorted: End Property
Public Property Let Sorted(Sorted As Boolean): bSorted = Sorted: bNotOk = False: UserControl_Paint: End Property
Public Property Get TopIndex() As Long: TopIndex = lTopIndex: End Property
Public Property Let TopIndex(TopIndex As Long): lTopIndex = TopIndex: bNotOk = False: UserControl_Paint: End Property
Public Property Get StyleCheckBox() As Boolean: StyleCheckBox = bStyleCheckBox: End Property
Public Property Let StyleCheckBox(StyleCheckBox As Boolean): bStyleCheckBox = StyleCheckBox: bNotOk = False: UserControl_Paint: End Property
Public Property Get Item(Index As Long) As vkListItem
On Error Resume Next: Set Item = Col.Item(Index)
Item.Checked = bChecked(Index)
Item.Selected = bSelected(Index)
End Property
Public Property Let Item(Index As Long, Item As vkListItem)
On Error Resume Next
Set Col.Item(Index) = Item
bSelected(Index) = Item.Selected
bChecked(Index) = Item.Checked
lHeight(Index) = Item.Height
bNotOk = False: UserControl_Paint
End Property
Public Property Get UnRefreshControl() As Boolean: UnRefreshControl = bUnRefreshControl: End Property
Public Property Let UnRefreshControl(UnRefreshControl As Boolean): bUnRefreshControl = UnRefreshControl: End Property
Public Property Get VScroll() As vkVScroll: On Error Resume Next: Set VScroll = VS: End Property
Public Property Let VScroll(VScroll As vkVScroll): On Error Resume Next: Set VS = VScroll: Call UserControl_Resize: End Property
Public Property Get DisplayVScroll() As Boolean: DisplayVScroll = bVSvisible: End Property
Public Property Let DisplayVScroll(DisplayVScroll As Boolean)
bVSvisible = DisplayVScroll
VS.Visible = bVSvisible
bNotOk = False: UserControl_Paint
End Property
Public Property Get CheckCount() As Long: CheckCount = lCheckCount: End Property
Public Property Get Alignment() As AlignmentConstants: Alignment = tAlig: End Property
Public Property Let Alignment(Alignment As AlignmentConstants): tAlig = Alignment: bNotOk = False: UserControl_Paint: End Property
Public Property Get ListItems() As vkListItems: On Error Resume Next: Set ListItems = Col: End Property
'Public Property Let ListItems(ListItems As vkListItems): On Error Resume Next: Set Col = ListItems: bNotOk = False: UserControl_Paint: End Property
Public Property Get FullRowSelect() As Boolean: FullRowSelect = lFullRowSelect: End Property
Public Property Let FullRowSelect(FullRowSelect As Boolean): lFullRowSelect = FullRowSelect: bNotOk = False: UserControl_Paint: End Property
Public Property Get BorderSelColor() As OLE_COLOR: BorderSelColor = lBorderSelColor: End Property
Public Property Let BorderSelColor(BorderSelColor As OLE_COLOR): lBorderSelColor = BorderSelColor: UserControl.ForeColor = ForeColor: bNotOk = False: UserControl_Paint: End Property


Private Sub UserControl_Paint()

    If bNotOk Or bNotOk2 Or bUnRefreshControl Then Exit Sub     'pas prêt à peindre
    
    Call Refresh    'on refresh
End Sub








'=======================================================
'PUBLIC SUBS
'=======================================================
'=======================================================
'ajoute un objet à la liste des objets
'=======================================================
Public Sub AddItem(Optional ByVal Caption As String, Optional ByVal Item As _
    vkListItem, Optional ByVal Key As String, Optional ByVal Index As Long = -1)
    
Dim tIt As vkListItem
    
    lListCount = lListCount + 1
        
    'redimensionne les tableaux avec le nombre d'items de la liste
    ReDim Preserve bChecked(lListCount - 1)
    ReDim Preserve bSelected(lListCount - 1)
    ReDim Preserve lHeight(lListCount - 1)
    
    If Item Is Nothing Then
        'alors on créé un nouvel Item dont on définit les prop par défaut
        Set tIt = New vkListItem
        With tIt
            .BackColor = lBackColor
            .Checked = False
            .Font = UserControl.Font
            .ForeColor = lForeColor
            .Key = Key
            .Selected = False
            .Text = Caption
            .Height = TextHeight(.Text) + 50
            .Alignment = tAlig
            .SelColor = lSelColor
            .BorderSelColor = lBorderSelColor
        End With
        
        If Index = -1 Then
            tIt.Index = Col.Count + 1
            lHeight(lListCount - 1) = tIt.Height
            Call Col.Add(tIt)
        Else
            tIt.Index = Index
            lHeight(Index) = tIt.Height
            Call Col.Add(tIt, Index)
        End If
        
    Else
        'on ajoute l'item passé en paramètre
        If Index = -1 Then
            Item.Index = lListCount - 1
            bSelected(Item.Index) = Item.Selected
            bChecked(Item.Index) = Item.Checked
            lHeight(lListCount - 1) = Item.Height
            Call Col.Add(Item)
        Else
            bSelected(Index) = Item.Selected
            bChecked(Index) = Item.Checked
            lHeight(Index) = Item.Height
            Call Col.Add(Item, Index)
        End If
   
    End If
    
    With VS
        .UnRefreshControl = True
        .Max = lListCount
        .UnRefreshControl = False
    End With
    
    'on refresh
    Call Refresh
End Sub

'=======================================================
'efface tous les objets de la liste
'=======================================================
Public Sub Clear()
Dim x As Long
    
    'efface les tableau
    ReDim bSelected(1)
    ReDim bChecked(1)
    ReDim lHeight(1)
    
    'on vide la collection...
    Call Col.Clear
    
    lListCount = 1
    lSelCount = 0
    lCheckCount = 0
    VS.Max = 1
    
    'refresh
    Call Refresh
End Sub

'=======================================================
'inverse la sélection
'=======================================================
Public Sub InvertSelection()
Dim x As Long
Dim y As Long

    'inverse le contenu du tableau
    For x = 1 To lListCount - 1
        bSelected(x) = Not (bSelected(x))
        If bSelected(x) Then y = y + 1
    Next x
    
    lSelCount = y
    
    'refresh
    Call Refresh
    
End Sub

'=======================================================
'inverse les cases cochées
'=======================================================
Public Sub InvertChecks()
Dim x As Long
Dim y As Long

    'inverse le contenu du tableau
    For x = 1 To lListCount - 1
        bChecked(x) = Not (bChecked(x))
        If bChecked(x) Then y = y + 1
    Next x
    
    lCheckCount = y
    
    'refresh
    Call Refresh
    
End Sub

'=======================================================
'enlève un item de la liste
'=======================================================
Public Sub RemoveItem(ByVal Index As Long)
    
    'vire l'item
    If bChecked(Index) Then
        lCheckCount = lCheckCount - 1
    End If
    If bSelected(Index) Then
        lSelCount = lSelCount - 1
    End If
    
    Call Col.Remove(Index)
    
    'on redimensionne les tableaux
    lListCount = lListCount - 1
    If lListCount < 1 Then lListCount = 1
    ReDim Preserve bChecked(lListCount)
    ReDim Preserve bSelected(lListCount)
    ReDim Preserve lHeight(lListCount)
    
    VS.Max = lListCount - 1
    
    'refresh
    Call Refresh
    
End Sub

'=======================================================
'sélectionne tout
'=======================================================
Public Sub SelectAll()
Dim x As Long

    'remplit le contenu du tableau
    For x = 1 To lListCount - 1
        bSelected(x) = True
    Next x
    
    lSelCount = lListCount
    
    'refresh
    Call Refresh
    
End Sub

'=======================================================
'ne sélectionne rien
'=======================================================
Public Sub UnSelectAll(Optional ByVal RefreshControl As Boolean = True)
Dim x As Long
    
    'remplit le contenu du tableau
    ReDim bSelected(lListCount)
    
    lSelCount = 0
    
    'refresh
    If RefreshControl Then Call Refresh
    
End Sub

'=======================================================
'checke tout
'=======================================================
Public Sub CheckAll()
Dim x As Long

    'remplit le contenu du tableau
    For x = 1 To lListCount - 1
        bChecked(x) = True
    Next x
    
    lCheckCount = lListCount
    
    'refresh
    Call Refresh
    
End Sub

'=======================================================
'ne check rien
'=======================================================
Public Sub UnCheckAll()
Dim x As Long
    
    'sélectionne tout
    ReDim bChecked(lListCount)
    
    lCheckCount = 0
    
    'refresh
    Call Refresh
    
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
Dim st As Long
Dim y As Long
Dim z As Long
Dim hRgn As Long
Dim x As Long
Dim hBrush As Long
Dim e As Long



Static ni As Long
Dim op As Long
    ni = ni + 1
    
    
    
    
    If bUnRefreshControl Then Exit Sub
    
    On Error Resume Next

    '//on efface
    Call UserControl.Cls

    If bEnable Then
        UserControl.ForeColor = lForeColor
    Else
        'couleur de enabled=false
        UserControl.ForeColor = 10070188
    End If

    '//convertit les couleurs
    Call OleTranslateColor(lBackColor, 0, lBackColor)
    Call OleTranslateColor(lForeColor, 0, lForeColor)
    Call OleTranslateColor(lBorderColor, 0, lBorderColor)
    
    
    '//on trace chaque élément de la liste
    
    'calcule le nombre d'items qui seront affichés
    x = 0 'contient la hauteur des composants affichés
    z = 0 'contient le nombre d'items à afficher
    For y = lTopIndex To lListCount - 1
        x = x + lHeight(y)
        If x >= Height - 30 Then Exit For
        z = z + 1
    Next y
    
    'limite le Max
    If lListCount <= z + TopIndex Then VS.Max = lListCount - z
    zNumber = z 'sauvegarde le nombre d'Items affichés
        
    If z < lListCount - 1 Then VS.Enabled = True Else VS.Enabled = False

    'on affiche maintenant chaque controle
    y = 1 'contient la hauteur temporaire
    st = 0
    
    For x = lTopIndex To lTopIndex + z

        'trace le texte
        Call DrawItem(Col.Item(x), y, x)

        'trace l'icone si présente
        If Not (Col.Item(x).Icon = 0) Then
            Call DrawItemIcon(Col.Item(x), y, x)
        End If

        'update la hauteur temporaire
        y = y + lHeight(x)
    Next x


    '//on trace le contour
    If bDisplayBorder Then
        'on défini un brush
        hBrush = CreateSolidBrush(lBorderColor)

        'on définit une zone rectangulaire à bords arrondi
        hRgn = CreateRectRgn(0, 0, ScaleWidth / 15, _
            ScaleHeight / 15)

        'on dessine le contour
        Call FrameRgn(UserControl.hdc, hRgn, hBrush, 1, 1)

        'on détruit le brush et la zone
        Call DeleteObject(hBrush)
        Call DeleteObject(hRgn)
    End If
    
    
    '//affiche les checkboxes
    If bStyleCheckBox Then Call SplitIMGandShow(z)
    
    
    '//on refresh le control
    Call UserControl.Refresh
    
    bNotOk = True

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

'=======================================================
'dessine un item sur le control
'=======================================================
Private Sub DrawItem(Item As vkListItem, lTop As Long, Index As Long)
Dim R As RECT
Dim st As Long
Dim tF As StdFont
Dim o As Long
Dim o2 As Long
Dim f As Long
Dim e As Long
Dim H As Long

    If bVSvisible Then
        'alors on décale le rect pour l'alignement à droite
        o = 19
    Else
        o2 = 17
    End If
    
    'décalage vers la droite si picture de checkboxes
    If bStyleCheckBox Then
        e = 15
    End If
    If Item.Icon Then
        e = e + Item.pxlIconWidth + 2
    End If
    
    If lFullRowSelect = False Then
        'alors on effectue un décalage si Check
        If bStyleCheckBox Then
            f = 230
        End If
        If Item.Icon Then
            f = f + 15 * Item.pxlIconWidth + 80
        End If
    End If
    
    
    'définit la fonte de l'item sur le controle
    Set tF = UserControl.Font
    Set UserControl.Font = Item.Font
    
    'récupère la hauteur du texte à afficher
    H = (lHeight(Index) - TextHeight(Item.Text)) / 30
    
    'définit une zone pour le texte
    Call SetRect(R, 7 + e, 1 + lTop / 15 + H, ScaleWidth / 15 - 1 - o, _
        1 + lTop / 15 + H + TextHeight(Item.Text) / 15) 'lTop + _
        (ScaleHeight - lTop - H / 2) / 15 + 1)
    
    'dessine un rectangle (backcolor ou selection) dans cette zone
    If bSelected(Index) = False Then
        'backcolor
        Line (15, lTop + 30)-(Width - 255 - 30 + o2 * 15, lTop + lHeight(Index) + 15), Item.BackColor, BF
    Else
        'sélection
        If f Then
            'alors on décale ==> on doit quand même faire le backColor
            Line (15, lTop + 30)-(Width - 255 - 30 + o2 * 15, lTop + lHeight(Index) + 15), Item.BackColor, BF
        End If
        
        'fond de la sélection
        Line (15 + f, lTop + 30)-(Width - 255 - 30 + o2 * 15, lTop + lHeight(Index) + 15), Item.SelColor, BF
        'bordure de la sélection
        Line (15 + f, lTop + 15)-(Width - 255 - 30 + o2 * 15, lTop + 15), Item.BorderSelColor
        Line (Width - 255 - 30 + o2 * 15, lTop + 30)-(Width - 255 - 30 + o2 * 15, lTop + lHeight(Index) + 15), Item.BorderSelColor
        Line (Width - 255 - 30 + o2 * 15, lTop + lHeight(Index) + 15)-(15 + f, lTop + lHeight(Index) + 15), Item.BorderSelColor
        Line (15 + f, lTop + lHeight(Index) + 15)-(15 + f, lTop + 15), Item.BorderSelColor
    End If
        
    
    'prépare l'alignement du texte
    If Item.Alignment = vbLeftJustify Then
        st = DT_LEFT
    ElseIf Item.Alignment = vbCenter Then
        st = DT_CENTER
    Else
        st = DT_RIGHT
    End If
    
    'définit la ForeColor et trace le texte
    UserControl.ForeColor = Item.ForeColor
    Call DrawText(UserControl.hdc, Item.Text, Len(Item.Text), R, st)
    Set UserControl.Font = tF 'restaure la fonte d'origine
End Sub

'=======================================================
'dessine l'icone d'un item
'=======================================================
Private Sub DrawItemIcon(Item As vkListItem, lTop As Long, Index As Long)
Dim y As Long
Dim SrcDC As Long
Dim SrcObj As Long
Dim e As Long
    
    'calcule le décalage en haut
    y = 1 + lTop / 15 + lHeight(Index) / 30 - Item.pxlIconHeight / 2

    SrcDC = CreateCompatibleDC(UserControl.hdc)
    SrcObj = SelectObject(SrcDC, Item.Icon)
    
    'décalage vers la droite si picture de checkboxes
    If bStyleCheckBox Then
        e = 15
    End If

    Call BitBlt(UserControl.hdc, 4 + e, y, Item.pxlIconWidth, _
        Item.pxlIconHeight, SrcDC, 0, 0, SRCCOPY)

    Call DeleteDC(SrcDC)
    Call DeleteObject(SrcObj)

End Sub

Private Sub VS_Change(Value As Currency)
Static lngOldValue
    
    'limite le Max
    If lListCount <= zNumber + TopIndex + 1 Then VS.Max = lListCount - zNumber
    
    lTopIndex = CLng(Value)
    
    'on en refresh QUE si on a changé de value entre temps
    If lngOldValue <> CLng(Value) Then Call Refresh
    lngOldValue = Value
End Sub

Private Sub VS_Scroll()
    lTopIndex = CLng(VS.Value)
    Call Refresh
End Sub

'=======================================================
'remplit la liste depuis un fichier
'=======================================================
Public Sub FillByFile(ByVal File As String)
Dim lFile As Long
Dim x As Long
Dim s As String
Dim t() As String
    
    On Error Resume Next
    
    'récupère le contenu du fichier
    lFile = FreeFile
    Open File For Binary Access Read As #lFile
    s = Space$(FileLen(File))
    Get #lFile, , s
    Close lFile
    
    'sépare chaque ligne
    ReDim t(0)
    t() = Split(s, vbNewLine, , vbBinaryCompare)
    
    'ajoute tous les items
    bUnRefreshControl = True
    For x = 0 To UBound(t())
        Call Me.AddItem(t(x))
    Next x
    bUnRefreshControl = False
    Call Refresh

End Sub

'=======================================================
'sauve la liste vers un fichier
'=======================================================
Public Sub SaveToFile(ByVal File As String)
Dim lFile As Long
Dim x As Long
Dim s As String
    
    On Error Resume Next
    
    'créé une string depuis les items
    lFile = FreeFile
    Open File For Binary Access Write As #lFile
    For x = 1 To lListCount
        s = Col.Item(x).Text
        If x < lListCount Then
             s = s & vbNewLine
        End If
        Put #lFile, , s
    Next x
    Close lFile

End Sub

'=======================================================
'affiche une des 6 images en la découpant depuis l'image complète
'=======================================================
Private Sub SplitIMGandShow(ByVal z As Long)
Dim hBrush As Long
Dim hRgn As Long
Dim x As Long
Dim y As Single
Dim lIMG As Long
Dim tVal As Boolean
Dim e As Long
    
    Debug.Print "SplitIMGandShow"
    '0 rien
    '1 survol
    '2 enabled=false
    '3 value enable
    '4 value survol enable
    '5 enable=false OR gray

'    SrcDC = CreateCompatibleDC(UserControl.hdc)
'    SrcObj = SelectObject(SrcDC, CreateCompatibleBitmap(UserControl.hdc, _
'        78, 13))

    'là, on va tracer un rectangle de la couleur BackColor pour effacer les pictures
    'Line (15, 15)-(230, Height - 30), lBackColor, BF
        
    'on découpe l'image correspondant à lIMG depuis Image1 et on blit
    'sur l'usercontrol
    
    If Col.Item(1) Is Nothing Then Exit Sub
    
    For x = lTopIndex To lTopIndex + z
    
        'Top de l'image
        e = y + lHeight(x) / 2 - 100
        
        If bChecked(x) Then
            If MouseItemIndex = x Then
                'checké et survolé
                lIMG = 4
            Else
                'checké sans survol
                lIMG = 3
            End If
        Else
            If MouseItemIndex = x Then
                'pas checké mais survol
                lIMG = 1
            Else
                'pas checké, pas survol
                lIMG = 0
            End If
        End If
        

        'trace l'image
        Call BitBlt(UserControl.hdc, 2, e / 15, 13, 13, pic(lIMG).hdc, _
              0, 0, SRCCOPY)
        
        'update la hauteur temporaire
        y = y + lHeight(x)
    Next x

    '//on trace le contour
    If bDisplayBorder Then
        'on défini un brush
        hBrush = CreateSolidBrush(lBorderColor)

        'on définit une zone rectangulaire à bords arrondi
        hRgn = CreateRectRgn(0, 0, ScaleWidth / 15, _
            ScaleHeight / 15)

        'on dessine le contour
        Call FrameRgn(UserControl.hdc, hRgn, hBrush, 1, 1)

        'on détruit le brush et la zone
        Call DeleteObject(hBrush)
        Call DeleteObject(hRgn)
    End If
    
    Call UserControl.Refresh
    Debug.Print Rnd
    'libère
'    Call DeleteDC(SrcDC)
'    Call DeleteObject(SrcObj)

End Sub
