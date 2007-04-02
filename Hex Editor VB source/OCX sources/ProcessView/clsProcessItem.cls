VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProcessItem"
Attribute VB_GlobalNameSpace = True
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
'//PERSONNAL TYPE
'RENSEIGNEMENTS SIMPLES SUR UN PROCESSUS
'=======================================================


'=======================================================
'ENUMS & TYPES
'=======================================================
'enum contenant les valeurs des priorités
Public Enum PRIORITY_TYPE
    HIGH_PRIORITY = &H80
    IDLE_PRIORITY = &H40
    NORMAL_PRIORITY = &H20
    REALTIME_PRIORITY = &H100
    ABOVE_NORMAL_PRIORITY = 44444
    BELOW_NORMAL_PRIORITY = -128
End Enum

'type concernant les informations mémoire des processus
Private Type PROCESS_MEMORY_COUNTERS
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
End Type

Private mdl() As ProcessModuleItem  'passage par variant pour le tableau
 
'=======================================================
'variables de la structure clsProcessItem
'=======================================================
Public dwSize As Long
Public cntUsage As Long
Public th32ProcessID As Long
Public th32DefaultHeapID As Long
Public th32ModuleID As Long
Public cntThreads As Long
Public th32ParentProcessID As Long
Public pcPriClassBase As Long
Public dwFlags As Long
Public szExeFile As String
Public childWnd As Integer
Public szImagePath As String
Public procParentProcess As ProcessItem
Public procModulesCount As Long
Public procMemory As ProcessMemoryItem
Public IsProc As Boolean
Public Property Get procModules() As Variant: procModules = mdl(): End Property
Public Property Let procModules(procModules As Variant): mdl() = procModules: End Property
