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
' -----------------------------------------------
'
' Hex Editor VB
' Coded by violent_ken (Alain Descotes)
'
' -----------------------------------------------
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
' -----------------------------------------------


Option Explicit

'-------------------------------------------------------
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
'-------------------------------------------------------


'-------------------------------------------------------
'VARIABLES PRIVEES
'-------------------------------------------------------
Private lMin As Currency
Private lMax As Currency
Private lValue As Currency
Private lSmallChange As Currency
Private lLargeChange As Currency
Private lOldValue As Currency
Public Event Change(Value As Currency)
Private bRecursive As Boolean   'pour éviter des boucles lors de l'update


'-------------------------------------------------------
'USERCONTROL SUBS
'-------------------------------------------------------
Private Sub UserControl_InitProperties()
    'valeurs par défaut
    Me.Min = 0
    Me.Max = 100
    Me.Value = 50
    Me.LargeChange = 10
    Me.SmallChange = 1
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
End Sub
Private Sub UserControl_Resize()
    VS.Height = UserControl.Height
    VS.Width = UserControl.Width
    VS.Left = 0
    VS.Top = 0
End Sub


'-------------------------------------------------------
'PROPERTIES
'-------------------------------------------------------
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



'-------------------------------------------------------
'rafraichit le VRAI scrollbar
'-------------------------------------------------------
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

'-------------------------------------------------------
'calcule les nouvelles valeurs virtuelles
'-------------------------------------------------------
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

'-------------------------------------------------------
'vérfie que les valeurs du usercontrol sont acceptables
'-------------------------------------------------------
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


