VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAsmProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'   O    O    O                                 ¯/ /¯/ /¯/ ¯/
'    \__/ \__/                                /¯  / / / /  /
'    /_||_||_\         /|¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯ ¯   ¯   ¯    |
'   // ||_||          | | ::..         Demo           ..:: |
'   \\_/\ #\ __       | |+--------------------------------+|
'   /_(  \  )__}      | |       Coded by EBArtSoft@        |
'  //  \O_\/          | |        Copyright © 2007          |
'  \\__||_||___       | |    http://www.ebartcodes.info    |
'   \__||_||__/        \|__________________________________|
'   ===========
'  ==== E.B ====  --------------------------------------------------------'
'                      <<< ALL RIGHTS RESERVED >>>                        '
' ------------------------------------------------------------------------'
' Permission  to  use,  copy,  modify,  and  distribute this software for '
' any  purpose and  without  fee  is  hereby  granted,  provided that the '
' above copyright notice appear in all copies and that both the copyright '
' notice  and  this permission notice appear in supporting documentation. '
' ------------------------------------------------------------------------'
' THE  MATERIAL  EMBODIED  ON  THIS  SOFTWARE IS  PROVIDED TO YOU "AS-IS" '
' AND  WITHOUT  WARRANTY  OF  ANY  KIND,  EXPRESS,  IMPLIED OR OTHERWISE, '
' INCLUDING  WITHOUT  LIMITATION,  ANY  WARRANTY  OF   MERCHANTABILITY OR '
' FITNESS  FOR  A  PARTICULAR  PURPOSE.  IN  NO EVENT  SHALL WE BE LIABLE '
' TO  YOU  OR  ANYONE ELSE FOR  ANY DIRECT, SPECIAL, INCIDENTAL, INDIRECT '
' OR  CONSEQUENTIAL  DAMAGES  OF  ANY  KIND,  OR ANY  DAMAGES WHATSOEVER, '
' INCLUDING  WITHOUT  LIMITATION,  LOSS  OF PROFIT, LOSS  OF USE, SAVINGS '
' OR REVENUE, OR THE CLAIMS OF THIRD PARTIES, WHETHER OR  NOT WE HAS BEEN '
' ADVISED  OF  THE  POSSIBILITY  OF  SUCH  LOSS,  HOWEVER  CAUSED AND  ON '
' ANY  THEORY  OF  LIABILITY,  ARISING  OUT OF OR IN CONNECTION  WITH THE '
' POSSESSION, USE OR PERFORMANCE OF THIS SOFTWARE.                        '
' ------------------------------------------------------------------------'

'========================================================================
' Start
'========================================================================
Private Declare Sub MemoryCopy Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private mCode()  As Byte
Private mProc()  As Long
Private mPtr     As Long
Private mSize    As Long
Private mCount   As Long

'===================================================================
' START OF ASM FUNCTIONS
'===================================================================

Private Sub Class_Initialize()
    
    '-----------------------------------
    ' bnAlloc2MoAlea
    '-----------------------------------
    AsmBegin
    Asm &H8B, &H44, &H24, &H8       '    mov         eax,dword ptr [esp+8]
    Asm &H85, &HC0                  '    test        eax,eax
    Asm &H74, &H7C                  '    je          0040108B
    Asm &H57                        '    push        edi
    Asm &H50                         '    push        eax
    Asm &H53                        '    push        ebx
    Asm &H8B, &HF8                  '    mov         edi,eax
    Asm &HBB, &H0, &H0, &H2, &H0    '    mov         ebx,20000h
    Asm &HF, &H31                   '    rdtsc
    Asm &H8B, &HD0                  '    mov         edx,eax
    Asm &HB9, &HFD, &H43, &H3, &H0  '    mov         ecx,343FDh
    Asm &H83, &HEF, &H4             '    sub         edi,4
    Asm &H8B, &HC2                  '    mov         eax,edx
    Asm &HF7, &HE1                  '    mul         eax,ecx
    Asm &H5, &HC3, &H9E, &H26, &H0  '    add         eax,269EC3h
    Asm &H83, &HC7, &H4             '    add         edi,4
    Asm &H8B, &HD0                  '    mov         edx,eax
    Asm &HC1, &HC8, &H10            '    ror         eax,10h
    Asm &H5, &HDF, &HC3, &H79, &H0  '    add         eax,79C3DFh
    Asm &H89, &H7                   '    mov         dword ptr [edi],eax
    Asm &H8B, &HC2                  '    mov         eax,edx
    Asm &HF7, &HE1                  '    mul         eax,ecx
    Asm &H5, &HC3, &H9E, &H26, &H0  '    add         eax,269EC3h
    Asm &H83, &HC7, &H4             '    add         edi,4
    Asm &H8B, &HD0                  '    mov         edx,eax
    Asm &HC1, &HC8, &H10            '    ror         eax,10h
    Asm &H5, &HDF, &HC3, &H79, &H0  '    add         eax,79C3DFh
    Asm &H89, &H7                   '    mov         dword ptr [edi],eax
    Asm &H8B, &HC2                  '    mov         eax,edx
    Asm &HF7, &HE1                  '    mul         eax,ecx
    Asm &H5, &HC3, &H9E, &H26, &H0  '    add         eax,269EC3h
    Asm &H83, &HC7, &H4             '    add         edi,4
    Asm &H8B, &HD0                  '    mov         edx,eax
    Asm &HC1, &HC8, &H10            '    ror         eax,10h
    Asm &H5, &HDF, &HC3, &H79, &H0  '    add         eax,79C3DFh
    Asm &H89, &H7                   '    mov         dword ptr [edi],eax
    Asm &H8B, &HC2                  '    mov         eax,edx
    Asm &HF7, &HE1                  '    mul         eax,ecx
    Asm &H5, &HC3, &H9E, &H26, &H0  '    add         eax,269EC3h
    Asm &H83, &HC7, &H4             '    add         edi,4
    Asm &H8B, &HD0                  '    mov         edx,eax
    Asm &HC1, &HC8, &H10            '    ror         eax,10h
    Asm &H5, &HDF, &HC3, &H79, &H0  '    add         eax,79C3DFh
    Asm &H4B                        '    dec         ebx
    Asm &H89, &H7                   '    mov         dword ptr [edi],eax
    Asm &H75, &H9D                  '    jne         00401025
    Asm &H5B                        '    pop         ebx
    Asm &H58                        '    pop         eax
    Asm &H5F                        '    pop         edi
    Asm &HC2, &H4, &H0              '    ret         4
    AsmEnd
        
    If (mCount = 0) Then Exit Sub
    
    Dim adr() As Long
    Dim ptr   As Long
    
    ReDim adr(mCount - 1)
    
    For ptr = 0 To UBound(mProc)
        adr(ptr) = VarPtr(mCode(mProc(ptr)))
    Next
    
    MemoryCopy ptr, ByVal ObjPtr(Me), 4
    MemoryCopy ByVal (ptr + 28), adr(0), 4 * mCount
    
End Sub

'===================================================================
' START OF ASM MANAGER
'===================================================================

Private Sub AsmBegin()
    If mCount Then ReDim Preserve mProc(mCount) Else ReDim mProc(0)
    mProc(mCount) = mPtr
    mCount = mCount + 1
End Sub

Private Sub AsmEnd()
    Dim i As Long
    For i = 1 To (4 - (mPtr Mod 4))
        Asm &HCC
    Next
End Sub

Private Sub Asm(ParamArray b() As Variant)
    Dim i As Long
    For i = 0 To UBound(b)
        If ((mPtr + i) >= mSize) Then
            If mSize Then ReDim Preserve mCode(mSize + 31) Else ReDim mCode(31)
            mSize = mSize + 32
        End If
        mCode(mPtr) = b(i)
        mPtr = mPtr + 1
    Next
End Sub

'===================================================================
' START OF EXPORTED FUNCTIONS
'===================================================================

Public Sub bnAlloc2MoAlea(ByRef Table As Byte): End Sub
