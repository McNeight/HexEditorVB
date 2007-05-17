VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEnums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
'ENUMS
'=======================================================
Public Enum BackStyleConstants
    Transparent = 0
    Opaque = 1
End Enum
Public Enum Label_Mode
    No = 0
    PercentageMode = 1
    ValueMode = 2
    Steps = 3
End Enum
Public Enum Border
    NoBorder = 0
    FixedSingle = 1
End Enum
Public Enum Mode_Degrade
    OnlyValue = 1
    AllLengh = 2
End Enum
Public Enum Button_Type
    NoButton = 0
    LeftButton = 1
    RightButton = 2
    MiddleButton = 4
End Enum
Public Enum Text_Alignment
    TopLeft = 1
    TopCenter = 2
    TopRight = 3
    MiddleLeft = 4
    MiddleCenter = 5
    MiddleRight = 6
    BottomLeft = 7
    BottomCenter = 8
    BottomRight = 9
End Enum
Public Enum TextPositionConstants
    [Text_Left] = 0
    [Text_Center] = 1
    [Text_Right] = 2
End Enum
Public Enum Wheel_Sens
    WHEEL_UP
    WHEEL_DOWN
End Enum
Public Enum GradientConstants
    None = 0
    Vertical = 1
    Horizontal = 2
End Enum
Public Enum PictureAlignment
    [Left Justify]
    [Right Justify]
End Enum
Public Enum Direction
    [Left_Right]
    [Up_Down]
End Enum
