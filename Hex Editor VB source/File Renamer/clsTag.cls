VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -----------------------------------------------
'
' File Renamer VB (part of Hex Editor VB)
' Coded by violent_ken (Alain Descotes)
'
' -----------------------------------------------
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
' -----------------------------------------------


Option Explicit

'-----------------------------------------------
'//OBJET TAG
'-----------------------------------------------

'-----------------------------------------------
'//VARIABLES DU TYPE PERSO
'-----------------------------------------------
Public Album As String
Public Artist As String
Public Comment As String
'Public Composer As String
'Public Copyright As String
'Public Encode As String
Public Genre As GENRE_OF_MUSIC
Public strGenre As String
'Public OriginalArtist As String
Public Title As String
Public TrackV1 As Integer   'IDv1 ==> integer
'Public TrackV2 As String    'IDv2 ==> string
'Public Url As String
Public Year As String


'-----------------------------------------------
'//ENUMS
'-----------------------------------------------
Public Enum GENRE_OF_MUSIC
    [Genre - Blues] = 0
    [Genre - Classic Rock] = 1
    [Genre - Country] = 2
    [Genre - Dance] = 3
    [Genre - Disco] = 4
    [Genre - Funk] = 5
    [Genre - Grunge] = 6
    [Genre - Hip-Hop] = 7
    [Genre - Jazz] = 8
    [Genre - Metal] = 9
    [Genre - New Age] = 10
    [Genre - Oldies] = 11
    [Genre - Other] = 12
    [Genre - Pop] = 13
    [Genre - R&B] = 14
    [Genre - Rap] = 15
    [Genre - Reggae] = 16
    [Genre - Rock] = 17
    [Genre - Techno] = 18
    [Genre - Industrial] = 19
    [Genre - Alternative] = 20
    [Genre - Ska] = 21
    [Genre - Death Metal] = 22
    [Genre - Pranks] = 23
    [Genre - Soundtrack] = 24
    [Genre - Euro-Techno] = 25
    [Genre - Ambient] = 26
    [Genre - Trip-Hop] = 27
    [Genre - Vocal] = 28
    [Genre - Jazz+Funk] = 29
    [Genre - Fusion] = 30
    [Genre - Trance] = 31
    [Genre - Classical] = 32
    [Genre - Instrumental] = 33
    [Genre - Acid] = 34
    [Genre - House] = 35
    [Genre - Game] = 36
    [Genre - Sound Clip] = 37
    [Genre - Gospel] = 38
    [Genre - Noise] = 39
    [Genre - Alternative Rock] = 40
    [Genre - Bass] = 41
    [Genre - Punk] = 42
    [Genre - Space] = 43
    [Genre - Meditative] = 44
    [Genre - Instrumental Pop] = 45
    [Genre - Instrumental Rock] = 46
    [Genre - Ethnic] = 47
    [Genre - Gothic] = 48
    [Genre - Darkwave] = 49
    [Genre - Techno-Industrial] = 50
    [Genre - Electronic] = 51
    [Genre - Pop-Folk] = 52
    [Genre - Eurodance] = 53
    [Genre - Dream] = 55
    [Genre - Southern Rock] = 56
    [Genre - Comedy] = 57
    [Genre - Cult] = 58
    [Genre - Gangsta] = 59
    [Genre - Top 40] = 60
    [Genre - Christian Rap] = 61
    [Genre - Pop/Funk] = 62
    [Genre - Jungle] = 63
    [Genre - Native US] = 64
    [Genre - Cabaret] = 65
    [Genre - New Wave] = 66
    [Genre - Psychadelic] = 67
    [Genre - Rave] = 68
    [Genre - Showtunes] = 69
    [Genre - Trailer] = 70
    [Genre - Lo-Fi] = 71
    [Genre - Tribal] = 72
    [Genre - Acid Punk] = 73
    [Genre - Acid Jazz] = 74
    [Genre - Polka] = 75
    [Genre - Retro] = 76
    [Genre - Musical] = 77
    [Genre - Rock & Roll] = 78
    [Genre - Hard Rock] = 79
    [Genre - Folk] = 80
    [Genre - Folk-Rock] = 81
    [Genre - National Folk] = 82
    [Genre - Swing] = 83
    [Genre - Fast Fusion] = 84
    [Genre - Bebob] = 85
    [Genre - Latin] = 86
    [Genre - Revival] = 87
    [Genre - Celtic] = 88
    [Genre - Bluegrass] = 89
    [Genre - Avantgarde] = 90
    [Genre - Gothic Rock] = 91
    [Genre - Progressive Rock] = 92
    [Genre - Psychedelic Rock] = 93
    [Genre - Symphonic Rock] = 94
    [Genre - Slow Rock] = 95
    [Genre - Big Band] = 96
    [Genre - Chorus] = 97
    [Genre - Easy Listening] = 98
    [Genre - Acoustic] = 99
    [Genre - Humour] = 100
    [Genre - Speech] = 101
    [Genre - Chanson] = 102
    [Genre - Opera] = 103
    [Genre - Chamber Music] = 104
    [Genre - Sonata] = 105
    [Genre - Symphony] = 106
    [Genre - Booty Bass] = 107
    [Genre - Primus] = 108
    [Genre - Porn Groove] = 109
    [Genre - Satire] = 110
    [Genre - Slow Jam] = 111
    [Genre - Club] = 112
    [Genre - Tango] = 113
    [Genre - Samba] = 114
    [Genre - Folklore] = 115
    [Genre - Ballad] = 116
    [Genre - Power Ballad] = 117
    [Genre - Rhytmic Soul] = 118
    [Genre - Freestyle] = 119
    [Genre - Duet] = 120
    [Genre - Punk Rock] = 121
    [Genre - Drum Solo] = 121
    [Genre - Acapella] = 123
    [Genre - Euro-House] = 124
    [Genre - Dance Hall] = 125
    [Genre - Goa] = 126
    [Genre - Drum & Bass] = 127
    [Genre - Club-House] = 128
    [Genre - Hardcore] = 129
    [Genre - Terror] = 130
    [Genre - Indie] = 131
    [Genre - BritPop] = 132
    [Genre - Negerpunk] = 133
    [Genre - Polsk Punk] = 134
    [Genre - Beat] = 135
    [Genre - Christian Gangsta] = 136
    [Genre - Heavy Metal] = 137
    [Genre - Black Metal] = 138
    [Genre - Crossover] = 139
    [Genre - Contemporary C] = 140
    [Genre - Christian Rock] = 141
    [Genre - Merengue] = 142
    [Genre - Salsa] = 143
    [Genre - Thrash Metal] = 144
    [Genre - Anime] = 145
    [Genre - JPop] = 146
    [Genre - SynthPop] = 147
    [No Genre] = 255
End Enum
