VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUndoSubItem"
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
'//CLASSE POUR LA GESTION DU UNDO/REDO
'TYPE PERSO CONTENANT L'HISTORIQUE (SUB ITEM)
'=======================================================

'=======================================================
'TYPES & ENUMS
'=======================================================
Public Enum UNDO_TYPE   'contient les types de Undo possibles
    actByteWritten = 0 '�criture d'un byte de type string
    actRestArea = 2 'restriction de visualisation
End Enum

'=======================================================
'PUBLIC VARIABLES (composant du type perso que constitue la classe)
'=======================================================
'variables qui d�finissent chaque entr�e de l'historique
Public tUndoType As UNDO_TYPE     'type de l'item qui consitue cet �l�ment
Public sData1 As String     'donn�e string 1 (g�n�ralement l'ancienne string)
Public sData2 As String     'donn�e string 2
Public curData1 As Currency   'donn�e currency 1 (g�n�ralement l'offset de l'ancienne string)
Public curData2 As Currency     'donn�e currency 2 (g�n�ralement l'offset max de la restriction de visualisation, et dans ce cas curData1 et l'offset min)
Public bytData1 As Byte     'donn�e byte 1 (g�n�ralement la composant de gauche du byte modifi�)
Public bytData2 As Byte     'donn�e byte 1 (g�n�ralement la composant de droite du byte modifi�)
Public lngData1 As Long     'donn�e long 1 (g�n�ralement le PID du processus)
