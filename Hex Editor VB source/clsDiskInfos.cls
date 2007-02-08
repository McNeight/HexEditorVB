VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDiskInfos"
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
'//MODULE DE CLASSE PERMETTANT DE RECUPERER LES INFOS SUR LES DRIVES
'// by violent_ken
'-------------------------------------------------------



'-------------------------------------------------------
'CONSTANTES
'-------------------------------------------------------
Private Const GENERIC_READ                  As Long = &H80000000
Private Const GENERIC_WRITE                 As Long = &H40000000
Private Const FILE_SHARE_READ               As Long = &H1
Private Const FILE_SHARE_WRITE              As Long = &H2
Private Const OPEN_EXISTING                 As Long = 3
Private Const INVALID_HANDLE_VALUE          As Long = -1&
Private Const FILE_BEGIN                    As Long = 0
Private Const IOCTL_DISK_GET_DRIVE_GEOMETRY As Long = &H70000
Private Const IOCTL_DISK_GET_PARTITION_INFO As Long = &H74004


'-------------------------------------------------------
'APIS
'-------------------------------------------------------
Private Declare Function GetLogicalDriveStrings Lib "kernel32.dll" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function DeviceIoControl Lib "kernel32.dll" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, lpOverlapped As OVERLAPPED) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long


'-------------------------------------------------------
'TYPES & ENUMS
'-------------------------------------------------------

Private Type DISK_GEOMETRY
   Cylinders         As Currency
   MediaType         As Long
   TracksPerCylinder As Long
   SectorsPerTrack   As Long
   BytesPerSector    As Long
End Type
Private Type OVERLAPPED
    ternal As Long
    ternalHigh As Long
    Offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type
Private Type PARTITION_INFORMATION
  StartingOffset As Currency
  PartitionLength As Currency
  HiddenSectors As Long
  PartitionNumber As Currency
  PartitionType As PARTITION_TYPE
  BootIndicator As Byte
  RecognizedPartition As Byte
  RewritePartition As Byte
End Type
Private Type DRIVE_LAYOUT_INFORMATION
  PartitionCount As Long
  Signature As Long
  PartitionEntry(1) As PARTITION_INFORMATION
End Type




'-------------------------------------------------------
'FUNCTIONS AND PROCEDURE
'-------------------------------------------------------

'-------------------------------------------------------
'obtient le drive (string) logique
'-------------------------------------------------------
Public Function GetLogicalDrive(ByVal Drive As String) As clsDrive
Dim hDrive As Long
Dim bytesReturned As Long
Dim strDrive As String
Dim tDrive As clsDrive
'temp variables
Dim a As Long, b As Long, c As Long, e As Long
Dim S As String, s2 As String, d As Long
Dim ovrTemp As OVERLAPPED

    On Error GoTo DriveNonDispo

    Set tDrive = New clsDrive

    'obtient un path valide de drive
    If Len(Drive) <> 6 Then strDrive = BuildDrive(Drive) Else strDrive = Drive

    tDrive.VolumeLetter = Left$(Right$(strDrive, 2), 1)
    
    '//obtient le handle du drive
    hDrive = CreateFile(strDrive, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    
    If hDrive = INVALID_HANDLE_VALUE Then
        CloseHandle hDrive
        Set GetLogicalDrive = tDrive
        Exit Function
    End If
    
    '//obtient les infos sur la géométrie du disque
        Dim tGeom As DISK_GEOMETRY
        
        'obtient les infos souhaitées par l'API DeviceIoControl ; ControlCode=Geometry
        DeviceIoControl hDrive, IOCTL_DISK_GET_DRIVE_GEOMETRY, ByVal 0&, 0&, tGeom, _
        Len(tGeom), bytesReturned, ovrTemp
        
        With tGeom
            tDrive.BytesPerSector = .BytesPerSector
            tDrive.Cylinders = 10000 * .Cylinders
            tDrive.MediaType = .MediaType
            tDrive.strMediaType = GetMediaType(.MediaType)
            tDrive.TracksPerCylinder = .TracksPerCylinder
            tDrive.SectorsPerTrack = .SectorsPerTrack
        End With
        
        '//tailles diverses
        GetDiskFreeSpace Right$(strDrive, 2) & "\", a, e, b, c
        tDrive.SectorPerCluster = a: tDrive.FreeClusters = b: tDrive.TotalClusters = c
        tDrive.BytesPerSector = e
        
        'taille = nbre de clusters * nbre de secteurs par cluster * nbre de bytes par secteur
        tDrive.TotalSpace = CCur(CCur(tDrive.BytesPerSector) * CCur(tDrive.SectorPerCluster) * CCur(tDrive.TotalClusters))
        tDrive.UsedClusters = tDrive.TotalClusters - tDrive.FreeClusters
        
        tDrive.BytesPerCluster = CLng(tDrive.BytesPerSector * tDrive.SectorPerCluster)



    '//obtient les infos PARTITION_INFO
        Dim tPart As PARTITION_INFORMATION
        
        'ControlCode=IOCTL_DISK_GET_PARTITION_INFO
        DeviceIoControl hDrive, IOCTL_DISK_GET_PARTITION_INFO, ByVal 0&, 0&, tPart, _
        Len(tPart), bytesReturned, ovrTemp
        
        With tPart
            tDrive.PartitionLength = 10000 * .PartitionLength
            tDrive.HiddenSectors = IIf(tDrive.TotalSpace <> 0, .HiddenSectors, 0)
            'tDrive.PartitionNumber = 10000 * .PartitionNumber
            'tDrive.PartitionType = .PartitionType
            tDrive.StartingOffset = IIf(tDrive.TotalSpace <> 0, .StartingOffset * 10000, 0)
        End With
                
    '//API GetVolumeInformation
        'buffer
        S = String$(255, 0)
        s2 = S
        
        'get info
        GetVolumeInformation Right$(strDrive, 2) & "\", S, 255, d, 0, 0, s2, 255
        tDrive.VolumeName = S
        tDrive.VolumeSerialNumber = d
        tDrive.FileSystemName = s2
        
        'formate string
        tDrive.VolumeName = Left$(tDrive.VolumeName, InStr(1, tDrive.VolumeName, vbNullChar) - 1)
        tDrive.FileSystemName = Left$(tDrive.FileSystemName, InStr(1, tDrive.FileSystemName, vbNullChar) - 1)
    
    If tDrive.FileSystemName = "NTFS" Then
        tDrive.TotalLogicalSectors = tDrive.TotalSpace / tDrive.BytesPerSector
        tDrive.TotalPhysicalSectors = tDrive.TotalLogicalSectors + tDrive.HiddenSectors
    Else
        tDrive.TotalLogicalSectors = tDrive.PartitionLength / tDrive.BytesPerSector
        tDrive.TotalPhysicalSectors = tDrive.TotalLogicalSectors + tDrive.HiddenSectors
    End If

    
    '//disk space
        tDrive.FreeSpace = CCur(((tDrive.FreeClusters / 10000) * tDrive.BytesPerCluster) * 10000) 'évite le dépassement de capacité
        tDrive.PercentageFree = CDbl((tDrive.FreeSpace / tDrive.TotalSpace) * 100)
        tDrive.UsedSpace = CCur(tDrive.TotalSpace - tDrive.FreeSpace)
    
    tDrive.DriveType = GetDriveType(Right$(strDrive, 2) & "\")
    tDrive.strDriveType = GetStringDriveType(tDrive.DriveType)
    
DriveNonDispo:

    CloseHandle hDrive
    Set GetLogicalDrive = tDrive
    
End Function

'-------------------------------------------------------
'obtient le drive (byte) physique
'-------------------------------------------------------
Public Function GetPhysicalDrive(ByVal Drive As Byte) As clsDrive
Dim hDrive As Long
Dim bytesReturned As Long
Dim strDrive As String
Dim tDrive As clsDrive
'temp variables
Dim a As Long, b As Long, c As Long, e As Long
Dim S As String, s2 As String, d As Long
Dim tGeom As DISK_GEOMETRY


    Set tDrive = New clsDrive

    'obtient un path valide de drive
    strDrive = "\\.\PHYSICALDRIVE" & CStr(Drive)

    
    '//obtient le handle du drive
    hDrive = CreateFile(strDrive, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    
    If hDrive = INVALID_HANDLE_VALUE Then
        CloseHandle hDrive
        Set GetPhysicalDrive = tDrive
        Exit Function
    End If
    
    
    '//obtient les infos sur la géométrie du disque
        Dim ovrTemp As OVERLAPPED
        
        'obtient les infos souhaitées par l'API DeviceIoControl ; ControlCode=Geometry
        DeviceIoControl hDrive, IOCTL_DISK_GET_DRIVE_GEOMETRY, ByVal 0&, 0&, tGeom, _
        Len(tGeom), bytesReturned, ovrTemp
        
        With tGeom
            tDrive.BytesPerSector = .BytesPerSector
            tDrive.Cylinders = 10000 * .Cylinders
            tDrive.MediaType = .MediaType
            tDrive.strMediaType = GetMediaType(.MediaType)
            tDrive.TracksPerCylinder = .TracksPerCylinder
            tDrive.SectorsPerTrack = .SectorsPerTrack
        End With

    '//tailles
        tDrive.TotalSpace = tDrive.SectorsPerTrack * tDrive.TracksPerCylinder * tDrive.Cylinders * tDrive.BytesPerSector
    
    CloseHandle hDrive
    
    Set GetPhysicalDrive = tDrive
End Function

'-------------------------------------------------------
'renvoie un drive compatible avec l'api CreateFile
'-------------------------------------------------------
Private Function BuildDrive(ByVal sDrive As String) As String
    BuildDrive = "\\.\" & UCase$(Left$(sDrive, 2))
End Function

'-------------------------------------------------------
'renvoie un type de drive en String correspondant au type en long
'-------------------------------------------------------
Private Function GetMediaType(ByVal bytType As Byte) As String

   Select Case bytType
      Case 0:               GetMediaType = "Format unknown"
      Case F5_1Pt2_512:     GetMediaType = "5.25, 1.2MB, 512 bytes/sector"
      Case F3_1Pt44_512:    GetMediaType = "3.5, 1.44MB, 512 bytes/sector"
      Case F3_2Pt88_512:    GetMediaType = "3.5, 2.88MB, 512 bytes/sector"
      Case F3_20Pt8_512:    GetMediaType = "3.5, 20.8MB, 512 bytes/sector"
      Case F3_720_512:      GetMediaType = "3.5, 720KB, 512 bytes/sector"
      Case F5_360_512:      GetMediaType = "5.25, 360KB, 512 bytes/sector"
      Case F5_320_512:      GetMediaType = "5.25, 320KB, 512 bytes/sector"
      Case F5_320_1024:     GetMediaType = "5.25, 320KB, 1024 bytes/sector"
      Case F5_180_512:      GetMediaType = "5.25, 180KB, 512 bytes/sector"
      Case F5_160_512:      GetMediaType = "5.25, 160KB, 512 bytes/sector"
      Case Removable:       GetMediaType = "Removable media other than floppy"
      Case FixedMedia:      GetMediaType = "Fixed hard disk"
      Case F3_120M_512:     GetMediaType = "3.5, 120M Floppy"
      Case F3_640_512:      GetMediaType = "3.5, 640KB, 512 bytes/sector"
      Case F5_640_512:      GetMediaType = "5.25, 640KB, 512 bytes/sector"
      Case F5_720_512:      GetMediaType = "5.25, 720KB, 512 bytes/sector"
      Case F3_1Pt2_512:     GetMediaType = "3.5, 1.2Mb, 512 bytes/sector"
      Case F3_1Pt23_1024:   GetMediaType = "3.5, 1.23Mb, 1024 bytes/sector"
      Case F5_1Pt23_1024:   GetMediaType = "5.25, 1.23MB, 1024 bytes/sector"
      Case F3_128Mb_512:    GetMediaType = "3.5 MO 128Mb, 512 bytes/sector"
      Case F3_230Mb_512:    GetMediaType = "3.5 MO 230Mb, 512 bytes/sector"
      Case F8_256_128:      GetMediaType = "8, 256KB, 128 bytes/sector"
      Case F3_200Mb_512:    GetMediaType = "3.5, 200M Floppy (HiFD)"
      Case F3_240M_512:     GetMediaType = "3.5, 240Mb Floppy (HiFD)"
      Case F3_32M_512:      GetMediaType = "3.5, 32Mb Floppy"
   End Select

End Function

'-------------------------------------------------------
'renvoie un type de drive en String correspondant au type en long
'-------------------------------------------------------
Private Function GetStringDriveType(ByVal bytType As Byte) As String

    Select Case bytType
        Case 1: GetStringDriveType = "No Root Directory"
        Case 2: GetStringDriveType = "Removeable Disk"
        Case 3: GetStringDriveType = "Drive fixed"
        Case 4: GetStringDriveType = "Remote"
        Case 5: GetStringDriveType = "CD-ROM"
        Case 6: GetStringDriveType = "RamDisk"
        Case Else: GetStringDriveType = "Unknown"
    End Select
    
End Function
'-------------------------------------------------------
'renvoie le nombre de disques physiques
'-------------------------------------------------------
Public Property Get NumberOfPhysicalDrives() As Byte
Dim bytVar As Byte
Dim hDrive As Long

    bytVar = 0: hDrive = 1

    While hDrive <> INVALID_HANDLE_VALUE
        'tente d'ouvrir les Drives 0,1,2,3,.....
        's'arrête à la première erreur (hdrive=-1)
        hDrive = CreateFile("\\.\PHYSICALDRIVE" & CStr(bytVar), GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
        CloseHandle hDrive
        bytVar = bytVar + 1
    Wend
    
    NumberOfPhysicalDrives = bytVar - 1
End Property

'-------------------------------------------------------
'renvoie la liste des drives physiques
'de 0 à Ubound-1
'-------------------------------------------------------
Public Function GetLogicalDrivesList(ByRef strDrives() As String) As Byte
Dim strBuffer As String
Dim x As Byte, lngLen As Long
Dim S() As String

    'créé un buffer qui recevra les noms de partition
    strBuffer = String$(255, 0)
    
    'obtient les noms
    lngLen = GetLogicalDriveStrings(255, strBuffer)
    
    'formate la string
    strBuffer = Left$(strBuffer, lngLen)
    
    'redimensionne le tableau
    ReDim S(Len(strBuffer) / 4)
    
    For x = 0 To Len(strBuffer) Step 4
        S(x / 4) = Mid$(strBuffer, x + 1, 4)
    Next x
    
    strDrives = S
    GetLogicalDrivesList = CByte(x / 4 - 1)
End Function


'-------------------------------------------------------
'vérifie que la partition est accessible
'-------------------------------------------------------
Public Function IsLogicalDriveAccessible(ByVal Drive As String) As Boolean
Dim hDrive As Long
Dim strDrive As String
'temp variables
Dim a As Long, b As Long, c As Long, e As Long

    On Error GoTo DriveNonDispo
    
    IsLogicalDriveAccessible = False

    'obtient un path valide de drive
    If Len(Drive) <> 6 Then strDrive = BuildDrive(Drive) Else strDrive = Drive
    
    '//obtient le handle du drive
    hDrive = CreateFile(strDrive, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    
    'récupère des infos sur la taille
    GetDiskFreeSpace Right$(strDrive, 2) & "\", a, e, b, c
    
    If c = 0 Or e = 0 Or a = 0 Or hDrive = INVALID_HANDLE_VALUE Then GoTo DriveNonDispo

    IsLogicalDriveAccessible = True

DriveNonDispo:
    CloseHandle hDrive
End Function

'-------------------------------------------------------
'vérifie que la partition est accessible
'-------------------------------------------------------
Public Function IsPhysicalDriveAccessible(ByVal Drive As Byte) As Boolean
Dim hDrive As Long
Dim bytesReturned As Long
Dim strDrive As String
Dim tGeom As DISK_GEOMETRY
Dim ovrTemp As OVERLAPPED

    'obtient un path valide de drive
    strDrive = "\\.\PHYSICALDRIVE" & CStr(Drive)

    '//obtient le handle du drive
    hDrive = CreateFile(strDrive, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)

    'obtient les infos souhaitées par l'API DeviceIoControl ; ControlCode=Geometry
    DeviceIoControl hDrive, IOCTL_DISK_GET_DRIVE_GEOMETRY, ByVal 0&, 0&, tGeom, _
    Len(tGeom), bytesReturned, ovrTemp
    
    If bytesReturned = 0 Or hDrive = INVALID_HANDLE_VALUE Or tGeom.BytesPerSector = 0 Or tGeom.Cylinders = 0 Or tGeom.SectorsPerTrack = 0 Or tGeom.TracksPerCylinder = 0 Then GoTo DriveNonDispo

    IsPhysicalDriveAccessible = True
        
DriveNonDispo:
    CloseHandle hDrive
End Function


