VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUndoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'//CLASSE POUR LA GESTION DU UNDO/REDO
'TYPE PERSO CONTENANT L'HISTORIQUE
'-------------------------------------------------------

'-------------------------------------------------------
'ENUMS & TYPES
'-------------------------------------------------------
Public Enum EDITION_TYPE   'contient le type d'édition à annuler
    edtFile 'fichier
    edtProcess  'mémoire
    edtDisk 'disque
End Enum

'-------------------------------------------------------
'PUBLIC VARIABLES (composant du type perso que constitue la classe)
'-------------------------------------------------------

'variables générales à la classe
Public lRang As Long        'valeur du composant actuel (position dans l'historique)
Public tEditType As EDITION_TYPE    'type d'édition de l'historique
Public Frm As Form  'contient la form gérée
Public lvHisto As ListView  'contient le listview de l'historique
