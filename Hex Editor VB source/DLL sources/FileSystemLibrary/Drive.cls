VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Drive"
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
Private MyDrive As ClassDrive

'=======================================================
'Constantes
'=======================================================
Private Const GENERIC_READ                  As Long = &H80000000
Private Const FILE_SHARE_READ               As Long = &H1
Private Const FILE_SHARE_WRITE              As Long = &H2
Private Const OPEN_EXISTING                 As Long = 3
Private Const INVALID_HANDLE_VALUE          As Long = -1
'=======================================================
'APIs
'=======================================================
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


'=======================================================
'Type privé contenant les infos sur ce fichier
'=======================================================
Private Type ClassDrive
    Cylinders As Currency
    TracksPerCylinder As Long
    SectorsPerTrack As Long
    PercentageFree As Double
    FreeSpace As Currency
    UsedSpace As Currency
    UsedClusters As Currency
    StartingOffset As Currency
    TotalSpace As Currency
    SectorPerCluster As Long
    BytesPerSector As Long
    FreeClusters As Long
    TotalClusters As Long
    TotalLogicalSectors As Currency
    TotalPhysicalSectors As Currency
    VolumeLetter As String
    BytesPerCluster As Long
    HiddenSectors As Long
    PartitionLength As Currency
    VolumeName As String
    VolumeSerialNumber As Long
    FileSystemName As String
    strMediaType As String
    DriveType As Byte
    strDriveType As String
End Type

'=======================================================
'Définit le path de cet objet
'=======================================================
Friend Function SetVolumeLetter(VolumeLetter As String, Optional ByVal RefreshInformations As Boolean = True)
    'on définit le path du fichier
    MyDrive.VolumeLetter = VolumeLetter
    If RefreshInformations Then Call RefreshInfos
End Function

'=======================================================
'//Propriétés de l'objet
'=======================================================
Public Property Get Cylinders() As Currency: Cylinders = MyDrive.Cylinders: End Property
Public Property Get TracksPerCylinder() As Long: TracksPerCylinder = MyDrive.TracksPerCylinder: End Property
Public Property Get SectorsPerTrack() As Long: SectorsPerTrack = MyDrive.SectorsPerTrack: End Property
Public Property Get PercentageFree() As Double: PercentageFree = MyDrive.PercentageFree: End Property
Public Property Get FreeSpace() As Currency: FreeSpace = MyDrive.FreeSpace: End Property
Public Property Get UsedSpace() As Currency: UsedSpace = MyDrive.UsedSpace: End Property
Public Property Get UsedClusters() As Currency: UsedClusters = MyDrive.UsedClusters: End Property
Public Property Get StartingOffset() As Currency: StartingOffset = MyDrive.StartingOffset: End Property
Public Property Get TotalSpace() As Currency: TotalSpace = MyDrive.TotalSpace: End Property
Public Property Get SectorPerCluster() As Long: SectorPerCluster = MyDrive.SectorPerCluster: End Property
Public Property Get BytesPerSector() As Long: BytesPerSector = MyDrive.BytesPerSector: End Property
Public Property Get FreeClusters() As Long: FreeClusters = MyDrive.FreeClusters: End Property
Public Property Get TotalClusters() As Long: TotalClusters = MyDrive.TotalClusters: End Property
Public Property Get TotalLogicalSectors() As Currency: TotalLogicalSectors = MyDrive.TotalLogicalSectors: End Property
Public Property Get TotalPhysicalSectors() As Currency: TotalPhysicalSectors = MyDrive.TotalPhysicalSectors: End Property
Public Property Get VolumeLetter() As String: VolumeLetter = MyDrive.VolumeLetter: End Property
Public Property Get BytesPerCluster() As Long: BytesPerCluster = MyDrive.BytesPerCluster: End Property
Public Property Get HiddenSectors() As Long: HiddenSectors = MyDrive.HiddenSectors: End Property
Public Property Get PartitionLength() As Currency: PartitionLength = MyDrive.PartitionLength: End Property
Public Property Get VolumeName() As String: VolumeName = MyDrive.VolumeName: End Property
Public Property Get VolumeSerialNumber() As Long: VolumeSerialNumber = MyDrive.VolumeSerialNumber: End Property
Public Property Get FileSystemName() As String: FileSystemName = MyDrive.FileSystemName: End Property
Public Property Get strMediaType() As String: strMediaType = MyDrive.strMediaType: End Property
Public Property Get DriveType() As Byte: DriveType = MyDrive.DriveType: End Property
Public Property Get strDriveType() As String: strDriveType = MyDrive.strDriveType: End Property



'=======================================================
'//Méthodes
'=======================================================

'=======================================================
'Rafraichit les infos
'=======================================================
Public Sub RefreshInfos()
Dim tGeom As DiskGeometry
Dim tSizes As DriveSizes
Dim tVol As VolumeInfo
Dim tPart As PartitionInfo
Dim hDevice As Long

    'ouvre le drive
    hDevice = CreateFile("\\.\" & UCase$(Me.VolumeLetter) & ":", GENERIC_READ, _
        FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)

    If hDevice = INVALID_HANDLE_VALUE Then Exit Sub
    
    'récupère les infos diverses
    With clsFile
        tGeom = .GetDriveGeometry_HANDLE(hDevice)
        If tGeom.BytesPerSector = 0 Then Exit Sub
        tSizes = .GetDriveSizes_HANDLE(hDevice)
        tPart = .GetDrivePartitionInfo_HANDLE(hDevice)
        tVol = .GetDriveVolumeInfo(Me.VolumeLetter)
    End With
    
    'referme le handle
    Call CloseHandle(hDevice)
    
    With MyDrive
        .BytesPerCluster = tSizes.BytesPerCluster
        .BytesPerSector = tGeom.BytesPerSector
        .Cylinders = tGeom.Cylinders
        .FileSystemName = tVol.FileSystemName
        .FreeClusters = tSizes.FreeClusters
        .DriveType = clsFile.GetDriveType(Me.VolumeLetter)
        .FreeSpace = tSizes.FreeSpace
        .HiddenSectors = tPart.HiddenSectors
        .PartitionLength = tPart.PartitionLength
        .PercentageFree = (tSizes.UsedSpace / tSizes.TotalSpace) * 100
        .SectorPerCluster = tSizes.BytesPerCluster / tSizes.BytesPerSector
        .SectorsPerTrack = tGeom.SectorsPerTrack
        .StartingOffset = tPart.StartingOffset
        .strDriveType = clsFile.GetStringDriveType(Me.DriveType)
        .strMediaType = tGeom.strMediaType
        .TotalClusters = tSizes.TotalClusters
        .TotalSpace = tSizes.TotalSpace
        If .FileSystemName = "NTFS" Then
            .TotalLogicalSectors = .TotalSpace / .BytesPerSector
            .TotalPhysicalSectors = .TotalLogicalSectors + .HiddenSectors
        Else
            .TotalLogicalSectors = .PartitionLength / .BytesPerSector
            .TotalPhysicalSectors = .TotalLogicalSectors + .HiddenSectors
        End If
        .TracksPerCylinder = tGeom.TracksPerCylinder
        .UsedClusters = tSizes.UsedClusters
        .UsedSpace = tSizes.UsedSpace
        .VolumeName = tVol.VolumeName
        .VolumeSerialNumber = tVol.VolumeSerialNumber
    End With

End Sub

'=======================================================
'Renvoie true si accessible
'=======================================================
Public Function IsDriveAvailable() As Boolean
    IsDriveAvailable = clsFile.IsDriveAvailable(Me.VolumeLetter)
End Function

'=======================================================
'Affiche la boite de dialogue propriétés
'=======================================================
Public Function ShowPropertyBox(ByVal hWnd As Long) As Long
    ShowPropertyBox = clsFile.ShowDriveProperty(Me.VolumeLetter, hWnd)
End Function

'=======================================================
'Lecture dans le disque
'=======================================================
Public Function ReadDriveString(ByVal StartingSector As Currency, ByVal nBytes As _
    Long) As String
    
Dim bpsec As Long

    'détermine le nombre de bytes par secteur
    If Me.BytesPerSector = 0 Then bpsec = _
        clsFile.GetDriveSizes(Me.VolumeLetter).BytesPerSector Else _
        bpsec = Me.BytesPerSector
    
    ReadDriveString = clsFile.ReadDriveString(Me.VolumeLetter, StartingSector, _
        nBytes, bpsec)

End Function

'=======================================================
'Ecriture dans le disque
'=======================================================
Public Function WriteDriveString(ByVal StartingSector As Currency, ByVal nBytes As _
    Long, ByVal StringToWrite As String) As Long
    
Dim bpsec As Long

    'détermine le nombre de bytes par secteur
    If Me.BytesPerSector = 0 Then bpsec = _
        clsFile.GetDriveSizes(Me.VolumeLetter).BytesPerSector Else _
        bpsec = Me.BytesPerSector
    
    WriteDriveString = clsFile.WriteDriveString(Me.VolumeLetter, StartingSector, _
        nBytes, bpsec, StringToWrite)
    
End Function

'=======================================================
'Récupère l'icone du fichier
'=======================================================
Public Function GetIcon(ByVal Size As IconSize) As IPictureDisp
    Set GetIcon = clsFile.GetIcon(Me.VolumeLetter & ":\", Size)
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

