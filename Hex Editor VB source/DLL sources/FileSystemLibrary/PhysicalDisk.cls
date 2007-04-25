VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PhysicalDisk"
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


Private clsFile As FileSystem
Private MyDisk As ClassDisk


'=======================================================
'Constantes
'=======================================================
Private Const GENERIC_READ                  As Long = &H80000000
Private Const FILE_SHARE_READ               As Long = &H1
Private Const FILE_SHARE_WRITE              As Long = &H2
Private Const OPEN_EXISTING                 As Long = 3
Private Const GENERIC_WRITE                 As Long = &H40000000
'=======================================================
'APIs
'=======================================================
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


'=======================================================
'variables de la structure clsDriveInfo
'=======================================================
Private Type ClassDisk
    Cylinders As Currency
    MediaType As MediaType
    TracksPerCylinder As Long
    SectorsPerTrack As Long
    TotalSpace As Currency
    BytesPerSector As Long
    TotalPhysicalSectors As Currency
    strMediaType As String
    DiskName As String
    DiskNumber As Byte
End Type

'=======================================================
'Définit le path de cet objet
'=======================================================
Friend Function SetDiskNumber(DiskNumber As Byte, Optional ByVal RefreshInformations As Boolean = True)
    'on définit le path du fichier
    MyDisk.DiskNumber = DiskNumber
    If RefreshInformations Then Call RefreshInfos
End Function

'=======================================================
'//Propriétés de l'objet
'=======================================================
Public Property Get Cylinders() As Currency: Cylinders = MyDisk.Cylinders: End Property
Public Property Get MediaType() As MediaType: MediaType = MyDisk.MediaType: End Property
Public Property Get TracksPerCylinder() As Long: TracksPerCylinder = MyDisk.TracksPerCylinder: End Property
Public Property Get SectorsPerTrack() As Long: SectorsPerTrack = MyDisk.SectorsPerTrack: End Property
Public Property Get TotalSpace() As Currency: TotalSpace = MyDisk.TotalSpace: End Property
Public Property Get BytesPerSector() As Long: BytesPerSector = MyDisk.BytesPerSector: End Property
Public Property Get TotalPhysicalSectors() As Currency: TotalPhysicalSectors = MyDisk.TotalPhysicalSectors: End Property
Public Property Get strMediaType() As String: strMediaType = MyDisk.strMediaType: End Property
Public Property Get DiskNumber() As Byte: DiskNumber = MyDisk.DiskNumber: End Property
Public Property Get DiskName() As String: DiskName = Trim$(MyDisk.DiskName): End Property



'=======================================================
'//Méthodes
'=======================================================

'=======================================================
'Rafraichit les infos
'=======================================================
Public Sub RefreshInfos()
Dim tGeom As DiskGeometry
Dim hDrive As Long
    
    'récupère un handle valide sur ce drive physique
    hDrive = CreateFile("\\.\PHYSICALDRIVE" & CStr(Me.DiskNumber), GENERIC_READ Or _
        GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0, 0)

    'récupère la géométrie du disque
    tGeom = clsFile.GetPhysicalDiskGeometry_HANDLE(DiskNumber)

    'récupère le nom du disque
    MyDisk.DiskName = clsFile.GetPhysicalDiskName_HANDLE(DiskNumber, Me.DiskNumber)
    
    'referme le handle
    Call CloseHandle(hDrive)
    
    With MyDisk
        .BytesPerSector = tGeom.BytesPerSector
        .Cylinders = tGeom.Cylinders
        .MediaType = tGeom.MediaType
        .SectorsPerTrack = tGeom.SectorsPerTrack
        .TracksPerCylinder = tGeom.TracksPerCylinder
        .strMediaType = tGeom.strMediaType
    
        'calculs diverses
        .TotalSpace = CCur(.SectorsPerTrack * .TracksPerCylinder * _
            .Cylinders * .BytesPerSector)
        .TotalPhysicalSectors = CCur(.SectorsPerTrack * _
            .TracksPerCylinder * .Cylinders)
    End With
   
End Sub

'=======================================================
'Lecture dans le disque
'=======================================================
Public Function ReadPhysicalDiskString(ByVal StartingSector As Currency, ByVal _
    nBytes As Long) As String
    
Dim bpsec As Long

    'détermine le nombre de bytes par secteur
    If Me.BytesPerSector = 0 Then bpsec = _
        clsFile.GetPhysicalDiskGeometry(Me.DiskNumber).BytesPerSector Else _
        bpsec = Me.BytesPerSector
    
    ReadPhysicalDiskString = clsFile.ReadPhysicalDiskString(Me.DiskNumber, StartingSector, _
        nBytes, bpsec)

End Function

'=======================================================
'Ecriture dans le disque
'=======================================================
Public Function WritePhysicalDiskString(ByVal StartingSector As Currency, ByVal _
    nBytes As Long, ByVal StringToWrite As String) As Long
    
Dim bpsec As Long

    'détermine le nombre de bytes par secteur
    If Me.BytesPerSector = 0 Then bpsec = _
        clsFile.GetPhysicalDiskGeometry(Me.DiskNumber).BytesPerSector Else _
        bpsec = Me.BytesPerSector
    
    WritePhysicalDiskString = clsFile.WritePhysicalDiskString(Me.DiskNumber, _
        StartingSector, nBytes, bpsec, StringToWrite)
    
End Function


'=======================================================
'Subs de la classe
'=======================================================
Private Sub Class_Initialize()
    'instancie la classe clsFile
    Set clsFile = New FileSystem
End Sub

Private Sub Class_Terminate()
    'libère la classe clsFile
    Set clsFile = Nothing
End Sub
