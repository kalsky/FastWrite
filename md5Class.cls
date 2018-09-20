VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "md5Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'******************************************************************************
'
'  MD5 Message-Digest Algorithm
'
'  Copyright (C) 2004 - J?r?me Frossard, PTAHSOFT GMBH
'
'  VB6 implementation based on the reference implementation found in the
'  document: "RFC 1321 - The MD5 Message-Digest Algorithm"
'
'  You can choose between two implementation. By default, the unsigned
'  math are done by the module MD5Helper.bas and are implemented in
'  assembler. To use it you need the module MD5Helper.bas and a reference
'  to the file MD5HelperLib.tlb. If you want pure VB6 implementation
'  without any depedency, you can use the conditional compilation variable
'  MD5_PUREVB = 1.
'
'  You can use this software under the terms of the original license.
'
'******************************************************************************
'
'  Copyright (C) 1991-2, RSA Data Security, Inc. Created 1991. All
'  rights reserved.
'
'  License to copy and use this software is granted provided that it
'  is identified as the "RSA Data Security, Inc. MD5 Message-Digest
'  Algorithm " in all material mentioning or referencing this software"
'  or this function.
'
'  License is also granted to make and use derivative works provided
'  that such works are identified as "derived from the RSA Data
'  Security, Inc. MD5 Message-Digest Algorithm" in all material
'  mentioning or referencing the derived work.
'
'  RSA Data Security, Inc. makes no representations concerning either
'  the merchantability of this software or the suitability of this
'  software for any particular purpose. It is provided "as is"
'  without express or implied warranty of any kind.
'
'  These notices must be retained in any copies of any part of this
'  documentation and/or software.
'
'******************************************************************************

#Const MD5_PUREVB = 1

'
' Constants for MD5Transform routine.
'
Private Const S11 As Long = 7
Private Const S12 As Long = 12
Private Const S13 As Long = 17
Private Const S14 As Long = 22
Private Const S21 As Long = 5
Private Const S22 As Long = 9
Private Const S23 As Long = 14
Private Const S24 As Long = 20
Private Const S31 As Long = 4
Private Const S32 As Long = 11
Private Const S33 As Long = 16
Private Const S34 As Long = 23
Private Const S41 As Long = 6
Private Const S42 As Long = 10
Private Const S43 As Long = 15
Private Const S44 As Long = 21

'
' MD5 context.
'
Private Type TMD5Context
    State(3)     As Long
    Count(1)     As Long
    Buffer(63) As Byte
End Type

Private m_context As TMD5Context



'//////////////////////////////////////////////////////////////////////////////
'//
'// External Procedure and Function
'//

Private Declare Sub MD5_memcpy Lib "kernel32.dll" _
            Alias "RtlMoveMemory" ( _
            hpvDest As Any, _
            hpvSource As Any, _
            ByVal cbCopy As Long)

Private Declare Sub MD5_zeromem Lib "kernel32.dll" _
            Alias "RtlZeroMemory" ( _
            hpvDest As Any, _
            ByVal cbSize As Long)
            

'//////////////////////////////////////////////////////////////////////////////
'//
'// Class Outgoing Interface Implementation
'//

'------------------------------------------------------------------------------
' Class_Initialize
'------------------------------------------------------------------------------
Private Sub Class_Initialize()
#If MD5_PUREVB <> 1 Then
    Call InitMD5Helper
#End If
'Debug.Print "md5 class Initialized"
End Sub


'//////////////////////////////////////////////////////////////////////////////
'//
'// Unsigned arithmetic helpers
'//

#If MD5_PUREVB = 1 Then

'------------------------------------------------------------------------------
' URol
'------------------------------------------------------------------------------
Friend Function URol(ByVal val As Long, ByVal numofBits As Long) As Long

    Dim leftVal As Long
    Dim bitCounter As Long
    Dim endfor As Long
    
    endfor = (numofBits Mod 32)
    For bitCounter = 1 To endfor

        'get the bit 30 and bit 31
        leftVal = val And &HC0000000

        'shift all other bits one bit to the left
        val = (val And &H3FFFFFFF) * 2

        'move the bit 31 (sign bit) to bit 1
        If leftVal < 0 Then
            val = val Or &H1
        End If

        'move the bit 30 to the bit 31 (sign bit)
        If (leftVal And &H40000000) = &H40000000 Then
            val = val Or &H80000000
        End If

    Next

    URol = val

End Function

'------------------------------------------------------------------------------
' UShr
'------------------------------------------------------------------------------
Friend Function UShr(ByVal val As Long, ByVal numofBits As Integer) As Long

    Dim leftVal As Long
    Dim bitCounter As Long
    Dim endfor As Long
    
    endfor = (numofBits Mod 32)
    For bitCounter = 1 To endfor

        'get the bit 31
        leftVal = val And &H80000000

        'shift all other bits one bit to the right
        val = (val And &H7FFFFFFF) \ 2

        'move the bit 31 to the bit 30
        If (leftVal < 0) Then
            val = val Or &H40000000
        End If

    Next

    UShr = val

End Function

'------------------------------------------------------------------------------
' UShl
'------------------------------------------------------------------------------
Friend Function UShl(ByVal val As Long, ByVal numofBits As Integer) As Long

    Dim leftVal As Long
    Dim bitCounter As Long
    Dim endfor As Long
    
    endfor = (numofBits Mod 32)
    For bitCounter = 1 To endfor

        'get the bit 31 and the bit 30
        leftVal = val And &HC0000000

        'shift all other bits one bit to the right
        val = (val And &H3FFFFFFF) * 2

        'move the bit 30 to the bit 31 (sign bit)
        If (leftVal And &H40000000) = &H40000000 Then
            val = val Or &H80000000
        End If

    Next

    UShl = val

End Function

'------------------------------------------------------------------------------
' UAdd
'------------------------------------------------------------------------------
Friend Function UAdd( _
                    ByVal val1 As Long, _
                    ByVal val2 As Long) As Long

    Dim lowWord As Long
    Dim highWord As Long
    Dim carry As Long

    lowWord = (val1 And &HFFFF&) + (val2 And &HFFFF&)
    carry = UShr(lowWord, 16)
    highWord = UShr(val1, 16) + UShr(val2, 16) + carry

    UAdd = UShl(highWord, 16) Or (lowWord And &HFFFF&)

End Function

'------------------------------------------------------------------------------
' UAdd4
'------------------------------------------------------------------------------
Friend Function UAdd4( _
                    ByVal val1 As Long, _
                    ByVal val2 As Long, _
                    ByVal val3 As Long, _
                    ByVal val4 As Long) As Long

    Dim highWord As Long
    Dim lowWord As Long
    Dim carry   As Long

    lowWord = (val1 And &HFFFF&) + (val2 And &HFFFF&) + (val3 And &HFFFF&) + (val4 And &HFFFF&)
    carry = UShr(lowWord, 16)
    highWord = UShr(val1, 16) + UShr(val2, 16) + UShr(val3, 16) + UShr(val4, 16) + carry

    UAdd4 = UShl(highWord, 16) Or (lowWord And &HFFFF&)

End Function

#End If 'MD5_PUREVB = 1



'//////////////////////////////////////////////////////////////////////////////
'//
'// MD5 routines
'//

#If MD5_PUREVB = 1 Then

'------------------------------------------------------------------------------
' FF, GG, HH, and II transformations for rounds 1, 2, 3, and 4.
'------------------------------------------------------------------------------
Friend Sub FF(a As Long, b As Long, c As Long, d As Long, X As Long, s As Long, ac As Long)
    'F(b, c, d) = bc v not(b) d
    a = UAdd4(a, (b And c) Or (Not (b) And d), X, ac)
    a = URol(a, s)
    a = UAdd(a, b)
End Sub
Friend Sub GG(a As Long, b As Long, c As Long, d As Long, X As Long, s As Long, ac As Long)
    'G(b, c, d) = bd v c not(d)
    a = UAdd4(a, (b And d) Or (c And Not (d)), X, ac)
    a = URol(a, s)
    a = UAdd(a, b)
End Sub
Friend Sub HH(a As Long, b As Long, c As Long, d As Long, X As Long, s As Long, ac As Long)
    'H(b, c, d) = b xor c xor d
    a = UAdd4(a, b Xor c Xor d, X, ac)
    a = URol(a, s)
    a = UAdd(a, b)
End Sub
Friend Sub II(a As Long, b As Long, c As Long, d As Long, X As Long, s As Long, ac As Long)
    'I(b, c, d) = c xor (b v not(d))
    a = UAdd4(a, c Xor (b Or Not (d)), X, ac)
    a = URol(a, s)
    a = UAdd(a, b)
End Sub

#End If 'MD5_PUREVB = 1


'------------------------------------------------------------------------------
' MD5 initialization. Begins an MD5 operation, writing a new context.
'------------------------------------------------------------------------------
Private Sub MD5Init()
    
    m_context.Count(0) = 0
    m_context.Count(1) = 0
    
    'Load magic initialization constants.
    m_context.State(0) = &H67452301
    m_context.State(1) = &HEFCDAB89
    m_context.State(2) = &H98BADCFE
    m_context.State(3) = &H10325476

End Sub

'------------------------------------------------------------------------------
' MD5 block update operation. Continues an MD5 message-digest operation,
' processing another message block, and updating the context.
'------------------------------------------------------------------------------
Private Sub MD5Update(ByRef inputData() As Byte, ByVal inputLen As Long)
    
    Dim i As Integer
    Dim index As Long
    Dim partLen As Long
    Dim block(63) As Byte
    Dim inputLen64(1) As Long
    
    '
    ' Compute number of bytes mod 64
    '
#If MD5_PUREVB = 1 Then
    index = (m_context.Count(0) \ 8) Mod 64
#Else
    index = MD5Math.UShr(m_context.Count(0), 3) And &H3F
#End If
    
    '
    ' Update number of bits
    '
#If MD5_PUREVB = 1 Then
    '(TODO : use 64 bit integer)
    m_context.Count(0) = UAdd(m_context.Count(0), inputLen * 8)
#Else
    inputLen64(0) = MD5Math.UShl(inputLen, 3)
    inputLen64(1) = MD5Math.UShr(inputLen, 29)
    Call MD5Math.UAdd64(m_context.Count(0), inputLen64(0), m_context.Count(0))
#End If

    
    partLen = 64 - index
    
    '
    ' Transform as many times as possible.
    '
    If inputLen >= partLen Then
    
        Call MD5_memcpy(m_context.Buffer(index), inputData(0), partLen)
        Call MD5Transform(m_context.Buffer)
        
        i = partLen
        Do While i + 63 < inputLen
            Call MD5_memcpy(block(0), inputData(i), 64)
            Call MD5Transform(block)
            i = i + 64
        Loop
        
        index = 0
    
    Else
        
        i = 0
    
    End If
    
    '
    ' Buffer remaining input
    '
    If inputLen - i > 0 Then
        Call MD5_memcpy(m_context.Buffer(index), inputData(i), inputLen - i)
    End If

End Sub

'------------------------------------------------------------------------------
' MD5 finalization. Ends an MD5 message-digest operation, writing the
' the message digest and zeroizing the context.
'------------------------------------------------------------------------------
Private Sub MD5Final(ByRef digest() As Byte)
    
    Dim padding(63) As Byte
    Dim bits(7) As Byte
    Dim index As Long
    Dim padLen As Long
    'Dim i As Long
    
    padding(0) = &H80
    
    '
    ' Save number of bits
    '
    Call Encode(bits, m_context.Count, 8)
    
    '
    ' Pad out to 56 mod 64.
    '
#If MD5_PUREVB = 1 Then
    index = (m_context.Count(0) \ 8) Mod 64
#Else
    index = MD5Math.UShr(m_context.Count(0), 3) And &H3F
#End If

    If index < 56 Then
        padLen = 56 - index
    Else
        padLen = 120 - index
    End If
    
    Call MD5Update(padding, padLen)
    
    '
    ' Append length (before padding)
    '
    Call MD5Update(bits, 8)
    
    '
    ' Store state in digest
    '
    ReDim digest(15)
    Call Encode(digest, m_context.State, 16)

    '
    ' Zeroize sensitive information
    '

    Call MD5_zeromem(m_context, 98)

End Sub

'------------------------------------------------------------------------------
' MD5 basic transformation. Transforms state based on block.
'------------------------------------------------------------------------------
Private Sub MD5Transform(block() As Byte)
    
    Dim X(15) As Long
    Dim a As Long
    Dim b As Long
    Dim c As Long
    Dim d As Long
    'Dim i As Long
    
    a = m_context.State(0)
    b = m_context.State(1)
    c = m_context.State(2)
    d = m_context.State(3)
    
    Call Decode(X, block, 64)

#If MD5_PUREVB = 1 Then
    With Me
#Else
    With MD5Math
#End If
        
        '
        ' Round 1
        '
        Call .FF(a, b, c, d, X(0), S11, &HD76AA478) '1
        Call .FF(d, a, b, c, X(1), S12, &HE8C7B756) '2
        Call .FF(c, d, a, b, X(2), S13, &H242070DB) '3
        Call .FF(b, c, d, a, X(3), S14, &HC1BDCEEE) '4
        Call .FF(a, b, c, d, X(4), S11, &HF57C0FAF) '5
        Call .FF(d, a, b, c, X(5), S12, &H4787C62A) '6
        Call .FF(c, d, a, b, X(6), S13, &HA8304613) '7
        Call .FF(b, c, d, a, X(7), S14, &HFD469501) '8
        Call .FF(a, b, c, d, X(8), S11, &H698098D8) '9
        Call .FF(d, a, b, c, X(9), S12, &H8B44F7AF) '10
        Call .FF(c, d, a, b, X(10), S13, &HFFFF5BB1) '11
        Call .FF(b, c, d, a, X(11), S14, &H895CD7BE) '12
        Call .FF(a, b, c, d, X(12), S11, &H6B901122) '13
        Call .FF(d, a, b, c, X(13), S12, &HFD987193) '14
        Call .FF(c, d, a, b, X(14), S13, &HA679438E) '15
        Call .FF(b, c, d, a, X(15), S14, &H49B40821) '16
    
        '
        ' Round 2
        '
        Call .GG(a, b, c, d, X(1), S21, &HF61E2562) '17
        Call .GG(d, a, b, c, X(6), S22, &HC040B340) '18
        Call .GG(c, d, a, b, X(11), S23, &H265E5A51) '19
        Call .GG(b, c, d, a, X(0), S24, &HE9B6C7AA) '20
        Call .GG(a, b, c, d, X(5), S21, &HD62F105D) '21
        Call .GG(d, a, b, c, X(10), S22, &H2441453) '22
        Call .GG(c, d, a, b, X(15), S23, &HD8A1E681) '23
        Call .GG(b, c, d, a, X(4), S24, &HE7D3FBC8) '24
        Call .GG(a, b, c, d, X(9), S21, &H21E1CDE6) '25
        Call .GG(d, a, b, c, X(14), S22, &HC33707D6) '26
        Call .GG(c, d, a, b, X(3), S23, &HF4D50D87) '27
        Call .GG(b, c, d, a, X(8), S24, &H455A14ED) '28
        Call .GG(a, b, c, d, X(13), S21, &HA9E3E905) '29
        Call .GG(d, a, b, c, X(2), S22, &HFCEFA3F8) '30
        Call .GG(c, d, a, b, X(7), S23, &H676F02D9) '31
        Call .GG(b, c, d, a, X(12), S24, &H8D2A4C8A) '32
        
        '
        ' Round 3
        '
        Call .HH(a, b, c, d, X(5), S31, &HFFFA3942) '33
        Call .HH(d, a, b, c, X(8), S32, &H8771F681) '34
        Call .HH(c, d, a, b, X(11), S33, &H6D9D6122) '35
        Call .HH(b, c, d, a, X(14), S34, &HFDE5380C) '36
        Call .HH(a, b, c, d, X(1), S31, &HA4BEEA44) '37
        Call .HH(d, a, b, c, X(4), S32, &H4BDECFA9) '38
        Call .HH(c, d, a, b, X(7), S33, &HF6BB4B60) '39
        Call .HH(b, c, d, a, X(10), S34, &HBEBFBC70) '40
        Call .HH(a, b, c, d, X(13), S31, &H289B7EC6) '41
        Call .HH(d, a, b, c, X(0), S32, &HEAA127FA) '42
        Call .HH(c, d, a, b, X(3), S33, &HD4EF3085) '43
        Call .HH(b, c, d, a, X(6), S34, &H4881D05)   '44
        Call .HH(a, b, c, d, X(9), S31, &HD9D4D039) '45
        Call .HH(d, a, b, c, X(12), S32, &HE6DB99E5) '46
        Call .HH(c, d, a, b, X(15), S33, &H1FA27CF8) '47
        Call .HH(b, c, d, a, X(2), S34, &HC4AC5665) '48
    
        '
        ' Round 4
        '
        Call .II(a, b, c, d, X(0), S41, &HF4292244) '49
        Call .II(d, a, b, c, X(7), S42, &H432AFF97) '50
        Call .II(c, d, a, b, X(14), S43, &HAB9423A7) '51
        Call .II(b, c, d, a, X(5), S44, &HFC93A039) '52
        Call .II(a, b, c, d, X(12), S41, &H655B59C3) '53
        Call .II(d, a, b, c, X(3), S42, &H8F0CCC92) '54
        Call .II(c, d, a, b, X(10), S43, &HFFEFF47D) '55
        Call .II(b, c, d, a, X(1), S44, &H85845DD1) '56
        Call .II(a, b, c, d, X(8), S41, &H6FA87E4F) '57
        Call .II(d, a, b, c, X(15), S42, &HFE2CE6E0) '58
        Call .II(c, d, a, b, X(6), S43, &HA3014314) '59
        Call .II(b, c, d, a, X(13), S44, &H4E0811A1) '60
        Call .II(a, b, c, d, X(4), S41, &HF7537E82) '61
        Call .II(d, a, b, c, X(11), S42, &HBD3AF235) '62
        Call .II(c, d, a, b, X(2), S43, &H2AD7D2BB) '63
        Call .II(b, c, d, a, X(9), S44, &HEB86D391) '64
    
        m_context.State(0) = .UAdd(m_context.State(0), a)
        m_context.State(1) = .UAdd(m_context.State(1), b)
        m_context.State(2) = .UAdd(m_context.State(2), c)
        m_context.State(3) = .UAdd(m_context.State(3), d)

    End With
    
End Sub

'------------------------------------------------------------------------------
' Encodes input (UINT4) into output (unsigned char). Assumes len is
' a multiple of 4.
'------------------------------------------------------------------------------
Private Sub Encode( _
                outputData() As Byte, _
                inputData() As Long, _
                ByVal inputLen As Integer)
    
    Dim i As Long
    Dim j As Long
    
#If MD5_PUREVB = 1 Then
    With Me
#Else
    With MD5Math
#End If
        
        Do While j < inputLen
            outputData(j) = inputData(i) And &HFF&
            outputData(j + 1) = .UShr(inputData(i), 8) And &HFF&
            outputData(j + 2) = .UShr(inputData(i), 16) And &HFF&
            outputData(j + 3) = .UShr(inputData(i), 24) And &HFF&
            
            i = i + 1
            j = j + 4
        Loop
    
    End With

End Sub

'------------------------------------------------------------------------------
' Decodes input (unsigned char) into output (UINT4). Assumes len is
' a multiple of 4.
'------------------------------------------------------------------------------
Private Sub Decode( _
                outputData() As Long, _
                inputData() As Byte, _
                ByVal inputLen As Integer)
    
    Dim i As Long
    Dim j As Long
    
#If MD5_PUREVB = 1 Then
    With Me
#Else
    With MD5Math
#End If

        Do While j < inputLen
            outputData(i) = inputData(j) Or _
                            .UShl(inputData(j + 1), 8) Or _
                            .UShl(inputData(j + 2), 16) Or _
                            .UShl(inputData(j + 3), 24)
            
            i = i + 1
            j = j + 4
        Loop
    
    End With

End Sub


'//////////////////////////////////////////////////////////////////////////////
'//
'// Default Interface Implementation
'//

'------------------------------------------------------------------------------
' CreateHash
'------------------------------------------------------------------------------
Public Function CreateHash(inputData() As Byte) As Byte()
    Dim digest() As Byte
    Call MD5Init
    Call MD5Update(inputData, UBound(inputData) - LBound(inputData) + 1)
    Call MD5Final(digest)
    CreateHash = digest
End Function

