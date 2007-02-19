VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProcessModuleItem"
Attribute VB_GlobalNameSpace = True
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
'//PERSONNAL TYPE
'RENSEIGNEMENTS SIMPLES SUR UN MODULE DE PROCESSUS
'=======================================================

Public dwSize As Long                     'taille de cette structure (à initialiser avant l'appel à Module32First ou Module32Next)
Public th32ModuleID As Long               'ID du module
Public th32ProcessID As Long              'ID du processus qui utilise le module
Public GlblcntUsage As Long               'compteur d'usage global pour toutes les instances présentes dans la mémoire du système
Public ProccntUsage As Long               'commteur d'usage du module pour le processus
Public modBaseAddr As Long                'adresse de début du module en mémoire
Public modBaseSize As Long                'taille du module en mémoire
Public hModule As Long                    'HMODULE du module
Public szModule As String                 'nom du module
Public szExeFile As String                'path du module
