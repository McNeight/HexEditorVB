VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'CLASSE POUR LA CONCATENATION RAPIDE DE STRINGS
'=======================================================

'=======================================================
'API
'=======================================================
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDst As _
    Any, ByVal pSrc As Any, ByVal ByteLen As Long)

'=======================================================
'CONSTANTE
'=======================================================
Private Const BUFFER_SIZE                   As Long = 1000000

Private StrLen As Long
Private Buffer As String


Public Property Get Value() As String: Value = Left$(Buffer, StrLen): End Property
Public Property Let Value(s As String): Buffer = vbNullString: StrLen = 0: Call Append(s): End Property


'=======================================================
'concatene Value et s
'=======================================================
Public Sub Append(s As String)
Dim l As Long

   l = StrLen + Len(s)

   Do While l > Len(Buffer)
      Buffer = Buffer & Space$(BUFFER_SIZE)
   Loop

   CopyMemory StrPtr(Buffer) + (StrLen * 2), StrPtr(s), LenB(s)
   
   StrLen = l

End Sub
