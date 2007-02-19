VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTagInfo"
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
' An Windows utility which allows to rename lots of file (part of Hex Editor VB)
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

'=======================================================
'//CLASSE DE RECUPERATION DES TAGS D'UN MP3
'=======================================================

'=======================================================--------
'PUBLIC PROCEDURES & FUNCTIONS
'=======================================================--------

'=======================================================--------
'récupération des tags d'un fichier mp3 v1
'=======================================================--------
Public Function GetTagsV1(ByVal mp3File As String) As clsTag
Dim AllTag As String * 127
Dim lFile As Byte
Dim bytGenre As Byte
Dim bytTrack As Byte
Dim sAlbum As String
Dim sArtist As String
Dim sTitle As String
Dim iTrack As Integer
Dim tGenre As GENRE_OF_MUSIC
Dim sComment As String
Dim sYear As String

    'instancie la classe
    '/!\ pas de classe temporaire, affectation directe à la fonction
    Set GetTagsV1 = New clsTag
    
    'initialisation des buffers
    sArtist = vbNullString
    sAlbum = vbNullString
    sTitle = vbNullString
    sYear = vbNullString
    sComment = vbNullString
    tGenre = [No Genre]
    iTrack = 0
    
    'ouvre le fichier et choppe remplit des buffers
    'ouverture sans API ReadFile/CreatFile pour alléger le code
    lFile = FreeFile
    Open mp3File For Binary Access Read As #lFile
    
        Get #lFile, LOF(lFile) - 1, bytTrack   'piste
        Get #lFile, LOF(lFile) - 127, AllTag    'autre tags (string 127 bytes)
        Get #lFile, LOF(lFile) - 0, bytGenre   'genre
    
    'referme le fichier
    Close #lFile
    
    If LCase(Left(AllTag, 3)) <> "tag" Then
        'pas conforme ==> on quitte
        Exit Function
    End If
    
    'récupère les sous chaines depuis le buffer
    AllTag = Right$(AllTag, Len(AllTag) - 3)
    sTitle = Mid$(AllTag, 1, 30)
    sArtist = Mid$(AllTag, 31, 30)
    sAlbum = Mid$(AllTag, 61, 30)
    sYear = Mid$(AllTag, 91, 4)
    sComment = Mid$(AllTag, 95, 28)
    If bytGenre <= 147 Then tGenre = bytGenre
    
    'attribution des valeurs à la fonction
    With GetTagsV1
        .Album = FormatedString(sAlbum)
        .Artist = FormatedString(sArtist)
        .Comment = FormatedString(sComment)
        .Genre = bytGenre
        .strGenre = Genre2String(bytGenre)  'genre converti en string
        .Title = FormatedString(sTitle)
        .TrackV1 = bytTrack
        .Year = sYear
    End With
    
End Function


'=======================================================--------
'PRIVATE PROCEDURES & FUNCTIONS
'=======================================================--------

'=======================================================--------
'formatage de string
'=======================================================--------
Private Function FormatedString(ByVal sString As String) As String

    If InStr(sString, vbNullChar) Then _
    FormatedString = Left$(sString, InStr(sString, vbNullChar) - 1) Else _
    FormatedString = sString
    
    FormatedString = Trim$(FormatedString)
End Function

'=======================================================--------
'convertit un genre (byte) en string
'=======================================================--------
Private Function Genre2String(ByVal bytGenre As Byte) As String
    Select Case bytGenre
        Case 0
            Genre2String = "Blues"
        Case 1
            Genre2String = "Classic Rock"
        Case 2
            Genre2String = "Country"
        Case 3
            Genre2String = "Dance"
        Case 4
            Genre2String = "Disco"
        Case 5
            Genre2String = "Funk"
        Case 6
            Genre2String = "Grunge"
        Case 7
            Genre2String = "Hip-Hop"
        Case 8
            Genre2String = "Jazz"
        Case 9
            Genre2String = "Metal"
        Case 10
            Genre2String = "New Age"
        Case 11
            Genre2String = "Oldies"
        Case 12
            Genre2String = "Other"
        Case 13
            Genre2String = "Pop"
        Case 14
            Genre2String = "R&B"
        Case 15
            Genre2String = "Rap"
        Case 16
            Genre2String = "Reggae"
        Case 17
            Genre2String = "Rock"
        Case 18
            Genre2String = "Techno"
        Case 19
            Genre2String = "Industrial"
        Case 20
            Genre2String = "Alternative"
        Case 21
            Genre2String = "Ska"
        Case 22
            Genre2String = "Death Metal"
        Case 23
            Genre2String = "Pranks"
        Case 24
            Genre2String = "Soundtrack"
        Case 25
            Genre2String = "Euro-Techno"
        Case 26
            Genre2String = "Ambient"
        Case 27
            Genre2String = "Trip-Hop"
        Case 28
            Genre2String = "Vocal"
        Case 29
            Genre2String = "Jazz+Funk"
        Case 30
            Genre2String = "Fusion"
        Case 31
            Genre2String = "Trance"
        Case 32
            Genre2String = "Classical"
        Case 33
            Genre2String = "Instrumental"
        Case 34
            Genre2String = "Acid"
        Case 35
            Genre2String = "House"
        Case 36
            Genre2String = "Game"
        Case 37
            Genre2String = "Sound Clip"
        Case 38
            Genre2String = "Gospel"
        Case 39
            Genre2String = "Noise"
        Case 40
            Genre2String = "Alternative Rock"
        Case 41
            Genre2String = "Bass"
        Case 42
            Genre2String = "Soul"
        Case 43
            Genre2String = "Punk"
        Case 44
            Genre2String = "Space"
        Case 45
            Genre2String = "Meditative"
        Case 46
            Genre2String = "Instrumental Pop"
        Case 47
            Genre2String = "Instrumental Rock"
        Case 48
            Genre2String = "Ethnic"
        Case 49
            Genre2String = "Gothic"
        Case 50
            Genre2String = "Darkwave"
        Case 51
            Genre2String = "Techno-Industrial"
        Case 52
            Genre2String = "Electronic"
        Case 53
            Genre2String = "Pop-Folk"
        Case 54
            Genre2String = "Eurodance"
        Case 55
            Genre2String = "Dream"
        Case 56
            Genre2String = "Southern Rock"
        Case 57
            Genre2String = "Comedy"
        Case 58
            Genre2String = "Cult"
        Case 59
            Genre2String = "Gangsta"
        Case 60
            Genre2String = "Top 40"
        Case 61
            Genre2String = "Christian Rap"
        Case 62
            Genre2String = "Pop/Funk"
        Case 63
            Genre2String = "Jungle"
        Case 64
            Genre2String = "Native US"
        Case 65
            Genre2String = "Cabaret"
        Case 66
            Genre2String = "New Wave"
        Case 67
            Genre2String = "Psychadelic"
        Case 68
            Genre2String = "Rave"
        Case 69
            Genre2String = "Showtunes"
        Case 70
            Genre2String = "Trailer"
        Case 71
            Genre2String = "Lo-Fi"
        Case 72
            Genre2String = "Tribal"
        Case 73
            Genre2String = "Acid Punk"
        Case 74
            Genre2String = "Acid Jazz"
        Case 75
            Genre2String = "Polka"
        Case 76
            Genre2String = "Retro"
        Case 77
            Genre2String = "Musical"
        Case 78
            Genre2String = "Rock & Roll"
        Case 79
            Genre2String = "Hard Rock"
        Case 80
            Genre2String = "Folk"
        Case 81
            Genre2String = "Folk-Rock"
        Case 82
            Genre2String = "National Folk"
        Case 83
            Genre2String = "Swing"
        Case 84
            Genre2String = "Fast Fusion"
        Case 85
            Genre2String = "Bebob"
        Case 86
            Genre2String = "Latin"
        Case 87
            Genre2String = "Revival"
        Case 88
            Genre2String = "Celtic"
        Case 89
            Genre2String = "Bluegrass"
        Case 90
            Genre2String = "Avantgarde"
        Case 91
            Genre2String = "Gothic Rock"
        Case 92
            Genre2String = "Progressive Rock"
        Case 93
            Genre2String = "Psychedelic Rock"
        Case 94
            Genre2String = "Symphonic Rock"
        Case 95
            Genre2String = "Slow Rock"
        Case 96
            Genre2String = "Big Band"
        Case 97
            Genre2String = "Chorus"
        Case 98
            Genre2String = "Easy Listening"
        Case 99
            Genre2String = "Acoustic"
        Case 100
            Genre2String = "Humour"
        Case 101
            Genre2String = "Speech"
        Case 102
            Genre2String = "Chanson"
        Case 103
            Genre2String = "Opera"
        Case 104
            Genre2String = "Chamber Music"
        Case 105
            Genre2String = "Sonata"
        Case 106
            Genre2String = "Symphony"
        Case 107
            Genre2String = "Booty Bass"
        Case 108
            Genre2String = "Primus"
        Case 109
            Genre2String = "Porn Groove"
        Case 110
            Genre2String = "Satire"
        Case 111
            Genre2String = "Slow Jam"
        Case 112
            Genre2String = "Club"
        Case 113
            Genre2String = "Tango"
        Case 114
            Genre2String = "Samba"
        Case 115
            Genre2String = "Folklore"
        Case 116
            Genre2String = "Ballad"
        Case 117
            Genre2String = "Power Ballad"
        Case 118
            Genre2String = "Rhytmic Soul"
        Case 119
            Genre2String = "Freestyle"
        Case 120
            Genre2String = "Duet"
        Case 121
            Genre2String = "Punk Rock"
        Case 122
            Genre2String = "Drum Solo"
        Case 123
            Genre2String = "Acapella"
        Case 124
            Genre2String = "Euro-House"
        Case 125
            Genre2String = "Dance Hall"
        Case 126
            Genre2String = "Goa"
        Case 127
            Genre2String = "Drum & Bass"
        Case 128
            Genre2String = "Club-House"
        Case 129
            Genre2String = "Hardcore"
        Case 130
            Genre2String = "Terror"
        Case 131
            Genre2String = "Indie"
        Case 132
            Genre2String = "BritPop"
        Case 133
            Genre2String = "Negerpunk"
        Case 134
            Genre2String = "Polsk Punk"
        Case 135
            Genre2String = "Beat"
        Case 136
            Genre2String = "Christian Gangsta"
        Case 137
            Genre2String = "Heavy Metal"
        Case 138
            Genre2String = "Black Metal"
        Case 139
            Genre2String = "Crossover"
        Case 140
            Genre2String = "Contemporary C"
        Case 141
            Genre2String = "Christian Rock"
        Case 142
            Genre2String = "Merengue"
        Case 143
            Genre2String = "Salsa"
        Case 144
            Genre2String = "Thrash Metal"
        Case 145
            Genre2String = "Anime"
        Case 146
            Genre2String = "JPop"
        Case 147
            Genre2String = "SynthPop"
        Case Else
            Genre2String = "No genre"
    End Select
End Function
