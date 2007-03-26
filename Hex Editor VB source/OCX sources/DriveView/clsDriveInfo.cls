VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrive"
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
'INFOS SUR UN DRIVE
'=======================================================


'=======================================================
'ENUMS
'=======================================================
Public Enum PARTITION_TYPE
    PARTITION_ENTRY_UNUSED = &H0
    PARTITION_FAT_12 = &H1
    PARTITION_XENIX_1 = &H2
    PARTITION_XENIX_2 = &H3
    PARTITION_FAT_16 = &H4
    PARTITION_EXTENDED = &H5
    PARTITION_HUGE = &H6
    PARTITION_IFS = &H7
    PARTITION_OS2BOOTMGR = &HA
    PARTITION_FAT32 = &HB
    PARTITION_FAT32_XINT13 = &HC
    PARTITION_XINT13 = &HE
    PARTITION_XINT13_EXTENDED = &HF
    PARTITION_PREP = &H41
    PARTITION_LDM = &H42
    PARTITION_UNIX = &H63
End Enum
Public Enum MEDIA_TYPE
    Unknown = 0
    F5_1Pt2_512 = 1
    F3_1Pt44_512 = 2
    F3_2Pt88_512 = 3
    F3_20Pt8_512 = 4
    F3_720_512 = 5
    F5_360_512 = 6
    F5_320_512 = 7
    F5_320_1024 = 8
    F5_180_512 = 9
    F5_160_512 = 10
    Removable = 11
    FixedMedia = 12
    F3_120M_512 = 13
    F3_640_512 = 14
    F5_640_512 = 15
    F5_720_512 = 16
    F3_1Pt2_512 = 17
    F3_1Pt23_1024 = 18
    F5_1Pt23_1024 = 19
    F3_128Mb_512 = 20
    F3_230Mb_512 = 21
    F8_256_128 = 22
    F3_200Mb_512 = 23
    F3_240M_512 = 24
    F3_32M_512 = 25
End Enum


'=======================================================
'variables de la structure clsDriveInfo
'=======================================================
Public Cylinders As Currency
Public MediaType As MEDIA_TYPE
Public TracksPerCylinder As Long
Public SectorsPerTrack As Long
Public PercentageFree As Double
Public FreeSpace As Currency
Public UsedSpace As Currency
Public UsedClusters As Currency
Public StartingOffset As Currency
Public PartitionNumber As Currency
Public TotalSpace As Currency
Public SectorPerCluster As Long
Public BytesPerSector As Long
Public FreeClusters As Long
Public TotalClusters As Long
Public TotalLogicalSectors As Currency
Public TotalPhysicalSectors As Currency
Public VolumeLetter As String
Public BytesPerCluster As Long
Public HiddenSectors As Long
Public PartitionLength As Currency
Public VolumeName As String
Public VolumeSerialNumber As Long
Public FileSystemName As String
Public PartitionType As PARTITION_TYPE
Public strMediaType As String
Public DriveType As Byte
Public strDriveType As String

