__declspec(naked) long* __stdcall bnAlloc2MoAlea()
{
  __asm {
    push    PAGE_READWRITE
    push    MEM_COMMIT or MEM_RESERVE or MEM_TOP_DOWN
    push    2097152 ;// 2 Mo
    push    0
    call    dword ptr VirtualAlloc
    test    eax, eax
    je      short allocEXIT
    push    edi
    push    eax
    push    ebx
    mov     edi, eax          ;// EDI = *pdt
    mov     ebx, 131072       ;// NBR TOURS DE 16 OCTETS
    rdtsc
    mov     edx, eax          ;// EDX = seedRand
    mov     ecx, 214013
    sub     edi, 4
  nextNBR:
    mov     eax, edx
    mul     ecx
    add     eax, 2531011
    add     edi, 4
    mov     edx, eax          ;// seedRand
    ror     eax, 16
    add     eax, 7979999
    mov     [edi], eax
    
    mov     eax, edx
    mul     ecx
    add     eax, 2531011
    add     edi, 4
    mov     edx, eax          ;// seedRand
    ror     eax, 16
    add     eax, 7979999
    mov     [edi], eax
    
    mov     eax, edx
    mul     ecx
    add     eax, 2531011
    add     edi, 4
    mov     edx, eax          ;// seedRand
    ror     eax, 16
    add     eax, 7979999
    mov     [edi], eax
    
    mov     eax, edx
    mul     ecx
    add     eax, 2531011
    add     edi, 4
    mov     edx, eax          ;// seedRand
    ror     eax, 16
    add     eax, 7979999
    
    dec     ebx
    mov     [edi], eax
    jnz     short nextNBR
    pop     ebx
    pop     eax               ;// POINTEUR ORIGINAL
    pop     edi
allocEXIT:
    ret     0
  }
}