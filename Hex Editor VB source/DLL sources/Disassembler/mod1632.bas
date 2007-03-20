Attribute VB_Name = "mod1632"
Option Explicit

Public bOperandSizeOverride As Byte
Public bAddressSizeOverride As Byte
Public dwInitESP As Long
'Public dwAddressSizeBytes As Byte
'Public dwAddressSizeBits As Byte

Public Sub Set16BitsDecode()
bOperandSizeOverride = &H0
bAddressSizeOverride = &H0
dwInitESP = 2
End Sub

Public Sub Set32BitsDecode()
bOperandSizeOverride = &H66
bAddressSizeOverride = &H67
dwInitESP = 4
End Sub

