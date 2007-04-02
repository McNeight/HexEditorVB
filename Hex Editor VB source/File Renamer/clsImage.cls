VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =======================================================
'
' File Renamer VB (part of Hex Editor VB)
' Coded by violent_ken (Alain Descotes)
'
' =======================================================
'
' A Windows utility which allows to rename lots of file (part of Hex Editor VB)
'
' Copyright (c) 2006-2007 by Alain Descotes.
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

'=======================================================--------
'//CLASSE IMAGE
'=======================================================--------


'=======================================================--------
'VARIABLES PUBLIQUES
'=======================================================--------
Public sType As String    'PNG/JPEG/BMP/GIF
Public lHeight As Long
Public lWidth As Long
Public sFile As String


'=======================================================--------
'TYPES & ENUMS
'=======================================================--------
Private Enum IMAGE_EXTENSION_TYPE
    lUNKNOWN = 0
    lGIF = 1
    lJPG = 2
    lPNG = 3
    lBMP = 4
End Enum
Private Type WordBytes
    byte1 As Byte
    byte2 As Byte
End Type
Private Type DWordBytes
    byte1 As Byte
    byte2 As Byte
End Type
Private Type WordWrapper
    Value As Integer
End Type

'=======================================================--------
'VARIABLES PRIVEES
'=======================================================--------
Private bytBuffer() As Byte


'=======================================================--------
'retourne un type IMAGE_INFO contenant des infos sur un fichier image
'=======================================================--------
Public Sub GetImageInfos()
Dim lFile As Long
Dim lngWidth As Long
Dim lngHeight As Long
Dim tType As IMAGE_EXTENSION_TYPE
Dim lngSize As Long

    On Error GoTo ErrGestion

    'par défaut
    lngWidth = 0
    lngHeight = 0
    tType = lUNKNOWN
    
    'ouverture du fichier en lecture binaire
    lFile = FreeFile
    Open sFile For Binary Access Read As #lFile
    
    lngSize = LOF(lFile)
    ReDim bytBuffer(lngSize)  'dimensionne à la taille du fichier
    
    'récupère le buffer
    Get #lFile, 1, bytBuffer()
    
    'vérifie le header ==> PNG
    If bytBuffer(0) = 137 And bytBuffer(1) = 80 And bytBuffer(2) = 78 Then
        tType = lPNG
        lngWidth = BEWord(18)
        lngHeight = BEWord(22)
    End If
    'vérifie le header ==> GIF
    If bytBuffer(0) = 71 And bytBuffer(1) = 73 And bytBuffer(2) = 70 Then
        tType = lGIF
        lngWidth = LEWord(6)
        lngHeight = LEWord(8)
    End If
    'vérifie le header ==> BMP
    If bytBuffer(0) = 66 And bytBuffer(1) = 77 Then
        tType = lBMP
        lngWidth = LEWord(18)
        lngHeight = LEWord(22)
    End If
    'vérifie le header ==> JPG
    If tType = lUNKNOWN Then
        'rien de particulier à dire, on tritouille les bytes à la recherche des infos
        'bloc de fonctions trouvé sur le net, mais je ne sais plus où -__-'
        Dim lPos As Long
        Do
            If (bytBuffer(lPos) = &HFF And bytBuffer(lPos + 1) = &HD8 And bytBuffer(lPos + 2) = &HFF) _
            Or (lPos >= lngSize - 10) Then Exit Do
            lPos = lPos + 1
        Loop
        lPos = lPos + 2
        If lPos >= lngSize - 10 Then Exit Sub
        Do
            Do
                If bytBuffer(lPos) = &HFF And bytBuffer(lPos + 1) <> &HFF Then Exit Do
                lPos = lPos + 1
                If lPos >= lngSize - 10 Then Exit Sub
            Loop
            lPos = lPos + 1
            If (bytBuffer(lPos) >= &HC0) And (bytBuffer(lPos) <= &HC3) Then Exit Do
            lPos = lPos + BEWord(lPos + 1)
            If lPos >= lngSize - 10 Then Exit Sub
        Loop
        tType = lJPG
        lngHeight = BEWord(lPos + 4)
        lngWidth = BEWord(lPos + 6)
    End If
    
    'affectation des valeurs aux variables publiques de la classe
    Me.lHeight = lngHeight
    Me.lWidth = lngWidth
    Me.sType = ImageType2String(tType)
    
ErrGestion:

    'vide le tableau (buffer)
    ReDim bytBuffer(0)
    'referme le fichier
    Close #lFile
End Sub


'=======================================================--------
'FONCTIONS PRIVEES
'=======================================================--------

'=======================================================--------
'retourne une string qui renvoie le type du fichier
'=======================================================--------
Private Function ImageType2String(ByVal lType As IMAGE_EXTENSION_TYPE) As String
    Select Case lType
        Case lUNKNOWN
            ImageType2String = "UNKNOWN"
        Case lGIF
            ImageType2String = "GIF"
        Case lJPG
            ImageType2String = "JPG"
        Case lPNG
            ImageType2String = "PNG"
        Case lBMP
            ImageType2String = "BMP"
    End Select
End Function

'=======================================================--------
'2 fonctions de conversion
'/!\ ces fonctions ont été trouvées sur le net, me rappelle plus vraiment d'où
'attention à l'affectation LSet qui affecte dans un type perso un AUTRE type perso
'les données binaires de x2 sont copiées dans l'espace mémoire de x1
'=======================================================--------
Private Function LEWord(position As Long) As Long
Dim x1 As WordBytes
Dim x2 As WordWrapper
    x1.byte1 = bytBuffer(position)
    x1.byte2 = bytBuffer(position + 1)
    LSet x2 = x1
    LEWord = x2.Value
End Function
Private Function BEWord(position As Long) As Long
Dim x1 As WordBytes
Dim x2 As WordWrapper
    x1.byte1 = bytBuffer(position + 1)
    x1.byte2 = bytBuffer(position)
    LSet x2 = x1
    BEWord = x2.Value
End Function



