VERSION 5.00
Begin VB.UserControl BGraphe 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   FillColor       =   &H000000FF&
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox pct 
      Height          =   615
      Left            =   480
      ScaleHeight     =   555
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "BGraphe"
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
' A complete hexadecimal editor for Windows �
' (Editeur hexad�cimal complet pour Windows �)
'
' Copyright � 2006-2007 by Alain Descotes.
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
'//GRAPHE EN "BARRES" POUR VISUALISER LES OCCURENCES
'D'APPARITION DES BYTES
'=======================================================

Private m(255) As Long
Private lBackColor As OLE_COLOR
Private lBarreColor1 As OLE_COLOR
Private lBarreColor2 As OLE_COLOR
Public Event KeyDown(KeyCode As Integer, Shift As Integer)

Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(bByteX As Byte, lOccurence As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'=======================================================
'USERCONTROL INITIALISATION
'=======================================================
Private Sub UserControl_InitProperties()
'valeurs par d�faut
    With Me
        .BackColor = vbWhite
        .BarreColor1 = vbRed
        .BarreColor2 = vbRed
    End With
End Sub

'=======================================================
'EVENTS
'=======================================================
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'calcule la barre sur laquelle on est
Dim dPxlPerX

    On Error Resume Next
    
    'en X
    dPxlPerX = (UserControl.Width / 256)
    
    RaiseEvent MouseMove(Round((x - 30) / dPxlPerX), m(Round((x - 30) / _
        dPxlPerX)), Button, Shift, x, y)
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'=======================================================
'USERCONTROL PROPERTIES
'=======================================================
Private Sub UserControl_Resize()
    Call TraceGraph
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("BackColor", Me.BackColor, vbWhite)
        Call .WriteProperty("BarreColor1", Me.BarreColor1, vbRed)
        Call .WriteProperty("BarreColor2", Me.BarreColor2, vbRed)
    End With
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Me.BackColor = .ReadProperty("BackColor", vbWhite)
        Me.BarreColor1 = .ReadProperty("BarreColor1", vbRed)
        Me.BarreColor2 = .ReadProperty("BarreColor2", vbRed)
    End With
End Sub
Public Property Get BackColor() As OLE_COLOR
    BackColor = lBackColor
End Property
Public Property Let BackColor(BackColor As OLE_COLOR)
    lBackColor = BackColor
    Call TraceGraph
End Property
Public Property Get BarreColor1() As OLE_COLOR
    BarreColor1 = lBarreColor1
End Property
Public Property Let BarreColor1(BarreColor1 As OLE_COLOR)
    lBarreColor1 = BarreColor1
    Call TraceGraph
End Property
Public Property Get BarreColor2() As OLE_COLOR
    BarreColor2 = lBarreColor2
End Property
Public Property Let BarreColor2(BarreColor2 As OLE_COLOR)
    lBarreColor2 = BarreColor2
    Call TraceGraph
End Property




'=======================================================
'Trace le graphique
'=======================================================
Public Sub TraceGraph()
Dim lMaxVal As Long
Dim x As Long
Dim dPxlPerX As Double
Dim dPxlPerY As Double
Dim lColorR As Double
Dim lColorG As Double
Dim lColorB As Double
Dim lColorRGB As Long
Dim lR1 As Long
Dim lG1 As Long
Dim lB1 As Long
Dim lR2 As Long
Dim lG2 As Long
Dim lB2 As Long


    Call ClearGraphe
    
    'calcule la valeur maximale (pour l'�chelle)
    For x = 0 To 255
        If m(x) > lMaxVal Then lMaxVal = m(x)
    Next x
    
    If lMaxVal = 0 Then Exit Sub
    
    'calcule l'�chelle
    'en X
    dPxlPerX = (UserControl.Width / 256)
    'en Y
    dPxlPerY = (UserControl.Height / lMaxVal)
    
    'peut commencer � tracer
    UserControl.BackColor = Me.BackColor
    
    'd�termien les incr�mentations des couleurs
    LongToRGB Me.BarreColor2, lR1, lG1, lB1
    LongToRGB Me.BarreColor1, lR2, lG2, lB2
    lColorR = (lR1 - lR2) / 255
    lColorG = (lG1 - lG2) / 255
    lColorB = (lB1 - lB2) / 255
    
    For x = 0 To 255
        'd�termine la couleur
        lColorRGB = RGB(lR2 + lColorR * x, lG2 + lColorG * x, lB2 + lColorB * x)
        
        'trace la barre
        UserControl.Line (dPxlPerX * x, UserControl.Height)-(dPxlPerX * (x + 1), UserControl.Height - dPxlPerY * m(x)), lColorRGB, BF
    Next x
    
End Sub

'=======================================================
'transforme une couleur long en RGB
'=======================================================
Private Function LongToRGB(ByVal lLong As Long, ByRef lRed As Long, ByRef lGreen As Long, ByRef lBlue As Long)
    lBlue = CLng(Int(lLong / 65536))
    lGreen = CLng(Int((lLong - CLng(lBlue) * 65536) / 256))
    lRed = CLng(lLong - CLng(lBlue) * 65536 - CLng(lGreen) * 256)
End Function

'=======================================================
'obtient une valeur � mettre dans le tableau
'=======================================================
Public Sub AddValue(ByVal bByte As Byte, ByVal lOccurence As Long)
    
    On Error GoTo ErrGestion
    
    m(bByte) = lOccurence
    
ErrGestion:
End Sub

'=======================================================
'obtient une valeur � mettre dans le tableau
'=======================================================
Public Function GetValue(ByVal bByte As Byte) As Long
    
    On Error GoTo ErrGestion
    
    GetValue = m(bByte)
    
ErrGestion:
End Function

'=======================================================
'efface les valeurs
'=======================================================
Public Sub ClearValues()
Dim x As Long

    For x = 0 To 255
        m(x) = 0
    Next x
End Sub

'=======================================================
'efface le graphe
'=======================================================
Public Sub ClearGraphe()
    Call UserControl.Cls
End Sub

'=======================================================
'sauvegarder en bitmap
'=======================================================
Public Sub SaveBMP(ByVal sFile As String, Optional ByVal lWidth As Long = 640, Optional ByVal lHeight As Long = 480)
    With pct
        .ScaleMode = 3 'pixels
        .Width = lWidth
        .Height = lHeight
        .Picture = UserControl.Image
        Call SavePicture(.Picture, sFile)
    End With
End Sub

