VERSION 5.00
Begin VB.UserControl ExtendedVScrollBar 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.VScrollBar VS 
      Height          =   2295
      Left            =   240
      SmallChange     =   10
      TabIndex        =   0
      Tag             =   "BE CAREFUL /!\ DO NOT MODIFY SMALLCHANGE AND LARGECHANGE VALUES IN THIS CONTROL"
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "ExtendedVScrollBar"
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
'//SCROLLBAR PERMETTANT D'ALLER A PLUS DE 2^15 (922337203685477)
'
'/!\ NE PAS MODIFIER LES VALEURS SMALLCHANGE ET LARGECHANGE
'DU CONTROLE SCROLLBAR POSE SUR LE USERCONTROL
'
'La vérification de la cohérence des valeurs Min, Max et Value
'est primaire. Si vous utilisez ce contrôle dans votre propre
'contexte, il sera nécessaire d'effectuer des vérifications
'plus poussées dans les Property Let des propriétés Min, Max et Value
'pour prévenir tout bug de la part d'un utilisateur
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
Private Const WM_EXITSIZEMOVE           As Long = &H232
Private Const WM_MOUSEMOVE              As Long = &H200
Private Const WM_MOVING                 As Long = &H216
Private Const WM_SIZING                 As Long = &H214
Private Const WM_MOUSELEAVE             As Long = &H2A3
Private Const WM_RBUTTONDBLCLK          As Long = &H206
Private Const WM_RBUTTONDOWN            As Long = &H204
Private Const WM_RBUTTONUP              As Long = &H205
Private Const WM_LBUTTONDBLCLK          As Long = &H203
Private Const WM_LBUTTONDOWN            As Long = &H201
Private Const WM_MBUTTONDBLCLK          As Long = &H209
Private Const WM_MBUTTONDOWN            As Long = &H207
Private Const WM_MBUTTONUP              As Long = &H208


'=======================================================
'APIs
'=======================================================
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
Private Enum eMsgWhen                                                       'When to callback
  MSG_BEFORE = 1                                                            'Callback before the original WndProc
  MSG_AFTER = 2                                                             'Callback after the original WndProc
  MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                                'Callback before and after the original WndProc
End Enum
Public Enum MOUSE_ACTION
    WHEEL_UP
    WHEEL_DOWN
    LEFT_UP
    LEFT_DBLCLICK
    RIGHT_UP
    RIGHT_DBLCLICK
    RIGHT_CLICK
    MIDDLE_UP
    MIDDLE_DBLCLICK
    MIDDLE_CLICK
    MOUSE_LEAVE
    MOUSE_MOVE
    MOUSE_ENTER
End Enum


'=======================================================
'SUBCLASS VARIABLE DECLARATION
'=======================================================
Private z_ScMem             As Long                                         'Thunk base address
Private z_Sc(64)            As Long                                         'Thunk machine-code initialised here
Private z_Funk              As Collection                                   'hWnd/thunk-address collection
Private bTrack                As Boolean
Private bTrackUser32          As Boolean
Private bInCtrl               As Boolean

'=======================================================
'EVENTS
'=======================================================
Public Event MouseAction(ByVal lngMouseAction As MOUSE_ACTION)


'=======================================================
'VARIABLES PRIVEES
'=======================================================
Private lMin As Currency
Private lMax As Currency
Private lValue As Currency
Private lSmallChange As Currency
Private lLargeChange As Currency
Private lOldValue As Currency
Public Event Change(Value As Currency)
Private bRecursive As Boolean   'pour éviter des boucles lors de l'update


'=======================================================
'USERCONTROL SUBS
'=======================================================
Private Sub UserControl_InitProperties()
    'valeurs par défaut
    Me.Min = 0
    Me.Max = 100
    Me.Value = 50
    Me.LargeChange = 10
    Me.SmallChange = 1
End Sub

Private Sub UserControl_Terminate()
'enleve le hook
    sc_Terminate
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Min", Me.Min, 1)
    Call PropBag.WriteProperty("Value", Me.Value, 50)
    Call PropBag.WriteProperty("LargeChange", Me.LargeChange, 10)
    Call PropBag.WriteProperty("SmallChange", Me.SmallChange, 1)
    Call PropBag.WriteProperty("Max", Me.Max, 100)
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Me.Min = PropBag.ReadProperty("Min", 1)
    Me.Value = PropBag.ReadProperty("Value", 50)
    Me.Max = PropBag.ReadProperty("Max", 100)
    Me.LargeChange = PropBag.ReadProperty("LargeChange", 100)
    Me.SmallChange = PropBag.ReadProperty("SmallChange", 100)
    lOldValue = VS.Value
    RefreshVS
    
    'c'est la bonne place pour commencer à subclasser
    If Ambient.UserMode Then  'If we're not in design mode
        bTrack = True
        bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")

        If Not bTrackUser32 Then
            If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then bTrack = False
        End If
      
      If bTrack Then
        'OS supports mouse leave, so let's subclass for it
      With VS
        'Subclass the UserControl
        
        sc_Subclass .hWnd
        
        sc_AddMsg .hWnd, WM_MOUSEWHEEL
        sc_AddMsg .hWnd, WM_MOUSEMOVE, MSG_AFTER
        sc_AddMsg .hWnd, WM_MOUSELEAVE, MSG_AFTER
        

        sc_AddMsg .hWnd, WM_MBUTTONDBLCLK
        sc_AddMsg .hWnd, WM_MBUTTONDOWN
        sc_AddMsg .hWnd, WM_MBUTTONUP
        sc_AddMsg .hWnd, WM_RBUTTONDBLCLK
        sc_AddMsg .hWnd, WM_RBUTTONDOWN
        sc_AddMsg .hWnd, WM_RBUTTONUP
        sc_AddMsg .hWnd, WM_LBUTTONDBLCLK
        sc_AddMsg .hWnd, WM_LBUTTONDOWN     'UP
        
      End With
    End If
    
    'Subclass the parent form
    With UserControl.Parent
      sc_Subclass .hWnd
      sc_AddMsg .hWnd, WM_MOVING, MSG_AFTER
      sc_AddMsg .hWnd, WM_SIZING, MSG_AFTER
      sc_AddMsg .hWnd, WM_EXITSIZEMOVE, MSG_AFTER
    End With
    End If
End Sub
Private Sub UserControl_Resize()
    VS.Height = UserControl.Height
    VS.Width = UserControl.Width
    VS.Left = 0
    VS.Top = 0
End Sub


'=======================================================
'PROPERTIES
'=======================================================
Public Property Get SmallChange() As Currency: SmallChange = lSmallChange: End Property
Public Property Let SmallChange(SmallChange As Currency): lSmallChange = SmallChange: RefreshVS: End Property
Public Property Get LargeChange() As Currency: LargeChange = lLargeChange: End Property
Public Property Let LargeChange(LargeChange As Currency): lLargeChange = LargeChange: RefreshVS: End Property
Public Property Get Min() As Currency: Min = lMin: End Property
Public Property Let Min(Min As Currency): lMin = Min: RefreshVS: End Property
Public Property Get Max() As Currency: Max = lMax: End Property
Public Property Let Max(Max As Currency): lMax = Max: RefreshVS: End Property
Public Property Get Value() As Currency: Value = lValue: End Property
Public Property Let Value(Value As Currency): lValue = Value: RefreshVS: lOldValue = VS.Value: End Property



'=======================================================
'rafraichit le VRAI scrollbar
'=======================================================
Private Sub RefreshVS()
'rafraichit le VS posé dans le UserControl
Dim lPercent As Double
Dim RealRange As Currency
Dim VirtualRange As Currency

    'calcule tout d'abord les intervalles réelles et virtuelles
    
    CheckValues 'vérifie que les valeurs sont compatibles

    RealRange = VS.Max - VS.Min
    VirtualRange = lMax - lMin
    
    'calcule maintenant le pourcentage du VS (vrituel ou réel, c'est la même chose)
    If VirtualRange Then lPercent = (lValue - lMin) / VirtualRange Else lPercent = 0
    
    'affecte la nouvelle value au VRAI VS
    bRecursive = True   'évite de faire une boucle
    VS.Value = VS.Min + lPercent * RealRange
    bRecursive = False
    
    'libère l'event
    RaiseEvent Change(lValue)
End Sub

'=======================================================
'calcule les nouvelles valeurs virtuelles
'=======================================================
Private Sub VS_Change()
Dim lPercent As Double
Dim RealRange As Currency
Dim VirtualRange As Currency
Dim lEcart As Currency
Dim lDelta As Currency
Dim l As Currency

    CheckValues 'vérifie que les valeurs sont compatibles

    If bRecursive Then Exit Sub
    
    'alors on recalcule les valeurs virtuelles
    
    'teste si l'on a appuyé sur les flèches (smallchange) ou
    'sur la zone de largechange, ou bien si l'on a utilisé Scroll (ou directement changement de value)
    lEcart = lOldValue - VS.Value 'différence entre l'état d'avant et l'état actuel
    
    If Abs(lEcart) = VS.SmallChange Or Abs(lEcart) = VS.LargeChange Then
        'alors c'est un smallchange/largechange
        If Abs(lEcart) = VS.SmallChange Then lDelta = Sgn(lEcart) * lSmallChange
        If Abs(lEcart) = VS.LargeChange Then lDelta = Sgn(lEcart) * lLargeChange

        'delta représente donc l'écart VIRTUEL entre avant et maintenant
        'ajoute le lDelta à la valeur virtuelle
        lValue = lValue - lDelta
        
        'calcule les range et le percentage
        RealRange = VS.Max - VS.Min
        VirtualRange = lMax - lMin
        If VirtualRange Then lPercent = (lValue - lMin) / VirtualRange Else lPercent = 0
                
        'affecte les VRAIES valeurs
        l = lPercent * RealRange
        bRecursive = True   'évite les boucles
        VS.Value = VS.Min + l
        bRecursive = False
                        
    Else
        'scroll ou changement de value par code
        
        'calcule les valeurs range (identique)
        RealRange = VS.Max - VS.Min
        VirtualRange = lMax - lMin
        If VirtualRange Then lPercent = (VS.Value - VS.Min) / RealRange Else lPercent = 0    'pourcentage NOUVEAU
        
        'affecte les valeurs VIRTUELLES
        lValue = Round(lMin + lPercent * VirtualRange)  'arrondi, car le currency gère les décimales
    End If
    
    'libère l'event
    RaiseEvent Change(lValue)
   
    lOldValue = VS.Value    'sauvegarde la position actuelle du VRAI VS

End Sub

Private Sub VS_Scroll()

    DoEvents    '/!\ IMPORTANT : DO NOT REMOVE
    'it allows to refresh correctly the HW control
    
    Call VS_Change
End Sub

'=======================================================
'vérfie que les valeurs du usercontrol sont acceptables
'=======================================================
Private Sub CheckValues()
Dim l As Currency

    '/!\ Vérifications PRIMAIRES qui doivent aussi être faites dans les Property Let
    'du usercontrol

    If Me.Min > Me.Max Then
        l = Me.Min
        Me.Min = Me.Max
        Me.Max = l
    End If
    If Me.Value > Me.Max Then
        Me.Value = Me.Max
    ElseIf Me.Value < Me.Min Then
        Me.Value = Me.Min
    End If
End Sub



























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

'Track the mouse leaving the indicated window
Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
  Dim tme As TRACKMOUSEEVENT_STRUCT
  
  If bTrack Then
    With tme
      .cbSize = Len(tme)
      .dwFlags = TME_LEAVE
      .hwndTrack = lng_hWnd
    End With

    If bTrackUser32 Then
      Call TrackMouseEvent(tme)
    Else
      Call TrackMouseEventComCtl(tme)
    End If
  End If
End Sub

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

    'libère un event en fonction du message
    Select Case uMsg
        
        Case WM_MOUSEWHEEL
            RaiseEvent MouseAction(IIf(wParam < 0, WHEEL_DOWN, WHEEL_UP))
    
        Case WM_MOUSEMOVE
            If Not bInCtrl Then
                bInCtrl = True
                Call TrackMouseLeave(lng_hWnd)
                RaiseEvent MouseAction(MOUSE_ENTER)
            Else
                RaiseEvent MouseAction(MOUSE_MOVE)
            End If
        
        Case WM_MOUSELEAVE
            bInCtrl = False
            RaiseEvent MouseAction(MOUSE_LEAVE)
    
        Case WM_LBUTTONDOWN
            RaiseEvent MouseAction(LEFT_UP)
    
        Case WM_RBUTTONDOWN
            RaiseEvent MouseAction(RIGHT_CLICK)
    
        Case WM_MBUTTONDOWN
            RaiseEvent MouseAction(MIDDLE_CLICK)
        
        Case WM_RBUTTONUP
            RaiseEvent MouseAction(RIGHT_UP)
    
        Case WM_MBUTTONUP
            RaiseEvent MouseAction(MIDDLE_UP)
    
        Case WM_RBUTTONDBLCLK
            RaiseEvent MouseAction(RIGHT_DBLCLICK)
    
        Case WM_LBUTTONDBLCLK
            RaiseEvent MouseAction(LEFT_DBLCLICK)
    
        Case WM_MBUTTONDBLCLK
            RaiseEvent MouseAction(MIDDLE_DBLCLICK)
    End Select
    
End Sub

