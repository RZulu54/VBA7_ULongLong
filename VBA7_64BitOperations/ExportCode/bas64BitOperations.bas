Attribute VB_Name = "bas64BitOperations"
'============================================================================================
'= 64 bit operations with
'= data type LongLong =   64 bit variables for 64 bit MS-Office (VBA7)
'= data type Long * 2 = 32*2 bit variables for 32 bit MS-Office or Visual Bsic 6
'= (by Roger Zuehlsdorf 2024 / email:rogzuehlsdorf@yahoo.de)
'============================================================================================
' List of Function in module 'bas64BitOperations'
' ------------------------------------------------
' 32-bit mode     64-bit mode           remarks
' ==========================================================================
' call Init32Bit first/call Init64Bit first
' Clear64         LL=0
' ClearBit64      ClearBit64LL
' SetBit64        SetBit64LL
' IsBitSet64      IsBitSet64LL
' Let64           L1 = L2
' AND64           LL = LL1 AND LL2
' OR64            LL = LL1 OR LL2
' XOr64           LL = LL1 XOR LL2
' NOT64           LL = NOT LL1
' ANDNOT64        LL = LL1 AND NOT LL2  for speed only (example)
' Equal64         (LL1 = LL2)
' IsNotEmpty64    LL1 <> 0
' ShiftLeft64     ShiftLeft64LL         shift left  1 bit
' ShiftRight64    ShiftRight64LL        shift right  1 bit
' BitsShiftLeft64   BitsShiftLeftLL     shift left  1 to 63 bits
' BitsShiftRight64  BitsShiftRightLL    shift right  1 to 63 bits
' PopCnt64        PopCnt64LL            returns number of bits set
' Lsb64           Lsb64LL               returns position of left most bit
' Rsb64           Rsb64LL               returns position of right most bit
' PopLsb64        PopLsb64LL            returns position of left most bit and removes it
' MoreThanOne64   PopCnt64LL>1          for speed only (example)
' SetAND64        LL = LL AND LL1       for speed only (example)
' SetOR64         LL = LL OR LL1        for speed only (example)
'============================================================================================


Option Explicit

Public Const MIN_INTEGER  As Integer = -32768 ' max 16 bit integer
Public Const MAX_INTEGER  As Integer = 32767  ' min 16 bit integer

'--- bit values as constants for bits 0-31
Public Const BitL_0 As Long = &H1&
Public Const BitL_1 As Long = &H2&
Public Const BitL_2 As Long = &H4&
Public Const BitL_3 As Long = &H8&
Public Const BitL_4 As Long = &H10&
Public Const BitL_5 As Long = &H20&
Public Const BitL_6 As Long = &H40&
Public Const BitL_7 As Long = &H80&
Public Const BitL_8 As Long = &H100&
Public Const BitL_9 As Long = &H200&
Public Const BitL_10 As Long = &H400&
Public Const BitL_11 As Long = &H800&
Public Const BitL_12 As Long = &H1000&
Public Const BitL_13 As Long = &H2000&
Public Const BitL_14 As Long = &H4000&
Public Const BitL_15 As Long = &H8000&
Public Const BitL_16 As Long = &H10000
Public Const BitL_17 As Long = &H20000
Public Const BitL_18 As Long = &H40000
Public Const BitL_19 As Long = &H80000
Public Const BitL_20 As Long = &H100000
Public Const BitL_21 As Long = &H200000
Public Const BitL_22 As Long = &H400000
Public Const BitL_23 As Long = &H800000
Public Const BitL_24 As Long = &H1000000
Public Const BitL_25 As Long = &H2000000
Public Const BitL_26 As Long = &H4000000
Public Const BitL_27 As Long = &H8000000
Public Const BitL_28 As Long = &H10000000
Public Const BitL_29 As Long = &H20000000
Public Const BitL_30 As Long = &H40000000
Public Const BitL_31 As Long = &H80000000


Public Type TBit64  ' emulate 64 bit, use 2x32 bit long
  i0 As Long
  i1 As Long
End Type

Public Type TInt16x4 ' emulate 64 bit, use 4x16 bit integer, needed for PopCount helper array
  i0 As Integer
  i1 As Integer
  i2 As Integer
  i3 As Integer
End Type

Public Bit32Value(31) As Long ' values for each bit in a 32 bit Long (0..31)
Public Pop16Cnt(MIN_INTEGER To MAX_INTEGER) As Long ' Max -int to +int(-32768...32767)  / long faster than byte
Public PopCntL(0 To 65535) As Long  ' Bit count as long for 16 bit positive value only (0...65535) / long faster than byte
Public Int16x4 As TInt16x4 ' 4x16 bit integer, needed for PopCount helper array

Public LSB16(MIN_INTEGER To MAX_INTEGER) As Integer  ' Lookup table for left most bit set
Public RSB16(MIN_INTEGER To MAX_INTEGER) As Integer  ' Lookup table for right most bit set

#If VBA7 And Win64 Then
  '--- LongLong data type 64 bit
  ' the ^ character identifies a LongLong data type
  
  ' constants needed for fast operations (special cases)
  'Public Const Bit63LL As LongLong = -2 ^ 63 ' negative value needed to set sign bit > slow?
  ' ^ = LongLong identifier / Strange: this shows an error ?!?  -9223372036854775808^
  Public Const Bit63LL As LongLong = -9223372036854775807^ - 1 ' avoid overflow with -9223372036854775808^
  
  'Public Const Bit62LL As LongLong = 2 ^ 62  'slow?
  Public Const Bit62LL As LongLong = 4611686018427387904^   ' to get this value: debug.Print CLngLng(2 ^62)

  Public Type TBit64LL ' Needed for LSET that maps data to 16 bit structure for LSB/RSB/PopCnt functions
    Bit64LL As LongLong
  End Type
  
  Public Bit64ValueLL(63) As LongLong ' 64 bit value lookup array for a bit at position 0-63
  
#End If ' VBA7 And Win64 <<<<<

'============================================================================================
'=
'=                32-bit implementation to simulate 64 bit operations
'=
'============================================================================================

'=========================================================================
'=  32 bit init functions
'=========================================================================
Public Sub Init32Bit()
  Dim i As Long, j As Long, k As Long, SqBB As Long
  
  For i = 0 To 31: Bit32Value(i) = BitMask32(i): Next
  
  For j = MIN_INTEGER To MAX_INTEGER
    Pop16Cnt(j) = Pop16CountFkt(j)
    LSB16(j) = -1
    For i = 0 To 15
       If CBool(j And Bit32Value(i)) Then LSB16(j) = i: Exit For
    Next
    RSB16(j) = -1
    For i = 15 To 0 Step -1
       If CBool(j And Bit32Value(i)) Then RSB16(j) = i: Exit For
    Next
  Next
  
  For j = 0 To 65535 ' PopCount for all positive with type long
    PopCntL(j) = Pop16CountLng(j)
  Next
  
End Sub

Function BitMask32(ByVal BitPos As Long) As Long ' 32 bit
  ' compute bit value 0..31 bit, special case for 31 bit = sign (2 ^32 > overflow)
  'If BitPos < 0 Or BitPos > 31 Then Err.Raise 6 ' overflow
  If BitPos < 31 Then
    BitMask32 = 2 ^ BitPos
  Else
    BitMask32 = BitL_31
  End If
End Function

Public Function Pop16CountFkt(ByVal x As Long) As Long
  ' fill PopCnt (number of bits set) lookup table for for Max -int to +int(-32768...32767)
  Pop16CountFkt = 0: If x = 0 Then Exit Function
  If x < 0 Then Pop16CountFkt = Pop16CountFkt + 1: x = x And Not BitL_15
  Do While x > 0
    Pop16CountFkt = Pop16CountFkt + 1: x = x And (x - 1)
  Loop
End Function

Public Function Pop16CountLng(ByVal x As Long) As Long
  ' fill PopCnt (number of bits set) lookup table for for positive values only
  Pop16CountLng = 0: If x = 0 Then Exit Function
  Do While x > 0
    Pop16CountLng = Pop16CountLng + 1: x = x And (x - 1)
  Loop
End Function


'============================================================================================
'= 32 bit operations
'= simulate unsigned 64bit using two signed long variables (problem: handle sign bit)
'= see http://www.xbeat.net/vbspeed/index.htm  for more functions and other implementations
'============================================================================================
Public Sub Clear64(op1 As TBit64)
  '--- clear both long variables of op1
  op1.i0 = 0: op1.i1 = 0
End Sub

Public Sub ClearBit64(op1 As TBit64, ByVal BitPos As Long)
  '--- clear bit in op1 at BitPos 0 to 63
  'Debug.Assert BitPos >= 0 And BitPos <= 63
  If BitPos < 32 Then
    op1.i0 = op1.i0 And Not Bit32Value(BitPos)
  Else
    op1.i1 = op1.i1 And Not Bit32Value(BitPos - 32)
  End If
End Sub

Public Sub SetBit64(op1 As TBit64, ByVal BitPos As Long)
  '--- set bit in op1 at BitPos 0 to 63
  'Debug.Assert BitPos >= 0 And BitPos <= 63
  If BitPos < 32 Then op1.i0 = op1.i0 Or Bit32Value(BitPos) Else op1.i1 = op1.i1 Or Bit32Value(BitPos - 32)
End Sub

Public Function IsBitSet64(op1 As TBit64, ByVal BitPos As Long) As Boolean
  '--- return  IsBitSet64 = (Is bit set at BitPos 0 to 63) ?
  'Debug.Assert BitPos >= 0 And BitPos <= 63
  If BitPos < 32 Then IsBitSet64 = CBool(op1.i0 And Bit32Value(BitPos)) Else IsBitSet64 = CBool(op1.i1 And Bit32Value(BitPos - 32))
End Function

Public Sub AND64(Result As TBit64, op1 As TBit64, op2 As TBit64)
  '--- return result = op1 AND op2
  Result.i0 = op1.i0 And op2.i0: Result.i1 = op1.i1 And op2.i1
End Sub

Public Sub OR64(Result As TBit64, op1 As TBit64, op2 As TBit64)
  '--- return result = op1 OR op2
  Result.i0 = op1.i0 Or op2.i0: Result.i1 = op1.i1 Or op2.i1
End Sub

Public Sub XOr64(Result As TBit64, op1 As TBit64, op2 As TBit64)
  '--- return result = op1 XOR op2
  Result.i0 = op1.i0 Xor op2.i0: Result.i1 = op1.i1 Xor op2.i1
End Sub

Public Sub NOT64(Result As TBit64, op1 As TBit64)
  '--- return result = NOT op1
  Result.i0 = Not op1.i0: Result.i1 = Not op1.i1
End Sub

Public Function Equal64(op1 As TBit64, op2 As TBit64) As Boolean
  '--- return Equal64 = (op1 = op2)  > is op1 equal op2 ?
  If op1.i0 = op2.i0 Then
    If op1.i1 = op2.i1 Then Equal64 = True Else Equal64 = False
  Else
    Equal64 = False
  End If
End Function

Public Sub Let64(Result As TBit64, op1 As TBit64)
  '---  Result = op1        >  assign op1 to Result
  Result.i0 = op1.i0: Result.i1 = op1.i1  ' much faster then  Result=Op1 !!!! UDT data type assigns are mem copies
End Sub

Public Function IsNotEmpty64(op1 As TBit64) As Boolean
  '--- return IsNotEmpty64 = (op1 <> 0)
  If op1.i0 <> 0 Then IsNotEmpty64 = True: Exit Function
  If op1.i1 <> 0 Then IsNotEmpty64 = True: Exit Function
  IsNotEmpty64 = False
End Function

Public Function ShiftLeft64(op1 As TBit64) As TBit64
  '--- return ShiftLeft64 = shift op1 1 bit left (order 0...63)
  ShiftLeft64.i0 = (op1.i0 And Not BitL_31) \ &H2&
  If (op1.i0 And BitL_31) Then ShiftLeft64.i0 = ShiftLeft64.i0 Or BitL_30
  If (op1.i1 And BitL_0) Then ShiftLeft64.i0 = ShiftLeft64.i0 Or BitL_31
  ShiftLeft64.i1 = (op1.i1 And Not BitL_31) \ &H2&
  If (op1.i0 And BitL_31) Then ShiftLeft64.i1 = ShiftLeft64.i1 Or BitL_30
End Function

Public Function BitsShiftLeft64(op1 As TBit64, ByVal Shift As Long) As TBit64
  '--- return BitsShiftLeft64 = op1 moved <shift> bit left (order 0...63) (shift values allowed: 0-64)
    Dim Mask As Long, SignBit As Long, CopyBits As Long, CopyMask As Long, CopyShift As Long, RightShifted As Long
    ' Shift = 31    ': op1.i0 = 0: op1.i1 = BitL_31 '# Debug
    
    BitsShiftLeft64.i0 = op1.i0: BitsShiftLeft64.i1 = op1.i1 ' much faster then BitsShiftLeft64 = op1
    If Shift = 0 Then Exit Function ' nothing to do
    
    If Shift < 0 Or Shift >= 64 Then BitsShiftLeft64.i0 = 0: BitsShiftLeft64.i1 = 0: Exit Function ' invalid Shift
 
    If Shift >= 32 Then ' special case, I1 always zero
      BitsShiftLeft64.i0 = op1.i1: BitsShiftLeft64.i1 = 0 ' just move the long variable I1 to I0 and shift I0 later
      Shift = Shift - 32: If Shift = 0 Then Exit Function 'finished if shift was  32
      If Shift = 31 Then BitsShiftLeft64.i0 = op1.i1 And BitL_31: Exit Function
      Mask = Not (Bit32Value(Shift) - 1)
    Else
      '---------------- shift I1 32-bit long of TBit64 -------------------------------
      ' create a mask with 1's for the digits that will be preserved
      If Shift = 31 Then ' special case
         BitsShiftLeft64.i0 = 0: BitsShiftLeft64.i1 = 0
         If op1.i1 And BitL_31 Then BitsShiftLeft64.i1 = BitL_0 'move I1 sign bit to bit 0
         If op1.i0 And BitL_31 Then BitsShiftLeft64.i0 = BitL_0 'move I0 sign bit to bit 0
         If op1.i1 And BitL_30 Then SignBit = BitL_31 ' move I1 bit 30 to I1 bit 31
         BitsShiftLeft64.i0 = BitsShiftLeft64.i0 Or ((op1.i1 And Not BitL_30 And Not BitL_31) * 2) Or SignBit
         Exit Function
      Else
        ' bits needed to copy from I1 to I0
        CopyBits = op1.i1 And (Bit32Value(Shift) - 1)
        CopyShift = 32 - Shift ' Shift this bits to the right edge
        CopyMask = Bit32Value(31 - CopyShift)
        If CopyBits And CopyMask Then
          RightShifted = (CopyBits And (CopyMask - 1)) * Bit32Value(CopyShift) Or BitL_31
        Else
          RightShifted = (CopyBits And (CopyMask - 1)) * Bit32Value(CopyShift)
        End If
        ' ShowDebugBit32L RightShifted '#Debug
        
        ' clear all the digits that will be discarded, and also clear the sign bit
        SignBit = (op1.i1 < 0) And Bit32Value(31 - Shift)
        Mask = Not (Bit32Value(Shift) - 1)
        BitsShiftLeft64.i1 = (op1.i1 And Not BitL_31) And Mask
        ' do the shift, without overflow, then add the sign bit
        BitsShiftLeft64.i1 = (BitsShiftLeft64.i1 \ Bit32Value(Shift)) Or SignBit
      End If
    End If
    
    '---------------- shift I0 32-bit long from TBit64 -------------------------------
    ' clear all the digits that will be discarded, and also clear the sign bit
    SignBit = (BitsShiftLeft64.i0 < 0) And Bit32Value(31 - Shift)
    BitsShiftLeft64.i0 = (BitsShiftLeft64.i0 And Not BitL_31) And Mask
    ' do the shift, without any problem, add the sign bit and add RightShifted from I1
    BitsShiftLeft64.i0 = (BitsShiftLeft64.i0 \ Bit32Value(Shift)) Or SignBit Or RightShifted

End Function

Public Function ShiftRight64(op1 As TBit64) As TBit64
  '--- return ShiftRight64 = shift op1 1 bit right (order 0...63)
  If op1.i0 And BitL_30 Then
    ShiftRight64.i0 = ((op1.i0 And Not BitL_30 And Not BitL_31) * &H2&) Or BitL_31 ' move bit30 to bit 31 to avoid overflow
  Else
    ShiftRight64.i0 = (op1.i0 And &H3FFFFFFF) * &H2&
  End If
  If op1.i1 And BitL_30 Then
    ShiftRight64.i1 = ((op1.i1 And Not BitL_30 And Not BitL_31) * &H2&) Or BitL_31
  Else
    ShiftRight64.i1 = (op1.i1 And &H3FFFFFFF) * &H2&
  End If
  If (op1.i0 And BitL_31) Then ShiftRight64.i1 = ShiftRight64.i1 Or BitL_0
End Function

Public Function BitsShiftRight64(op1 As TBit64, ByVal Shift As Long) As TBit64
  '--- return BitsShiftLeft64 = op1 moved <shift> bit right (order 0...63) (shift values allowed: 0-64)
  Dim Mask As Long, SignBit As Long, CopyBits As Long, CopyMask As Long, CopyShift As Long, LeftShifted As Long
    
  BitsShiftRight64.i0 = op1.i0: BitsShiftRight64.i1 = op1.i1 ' much faster then BitsShiftRight64 = op1
  If Shift = 0 Then Exit Function ' nothing to do
  If Shift < 0 Or Shift >= 64 Then BitsShiftRight64.i0 = 0: BitsShiftRight64.i1 = 0: Exit Function ' invalid Shift
 
  If Shift >= 32 Then
    BitsShiftRight64.i1 = op1.i0: BitsShiftRight64.i0 = 0 ' just move the long variable I0 to I1 and shift I1 later
    Shift = Shift - 32: If Shift = 0 Then Exit Function 'finished if shift was  32
    If Shift = 31 Then
     If BitsShiftRight64.i1 And BitL_0 Then BitsShiftRight64.i1 = BitL_31: Exit Function
    End If
    Mask = Bit32Value(31 - Shift)
  Else
    ' bits needed to copy from I0 to I1
    CopyBits = op1.i0
    CopyShift = 32 - Shift ' Shift this bits to the left edge
    SignBit = (CopyBits < 0) And Bit32Value(31 - CopyShift)
    Mask = Not (Bit32Value(CopyShift - 1) - 1)
    CopyBits = (CopyBits And Not BitL_31) And Mask
    ' do the shift, without any problem, add the sign bit
    CopyBits = (CopyBits \ Bit32Value(CopyShift)) Or SignBit
    
    ' move I0 right
    Mask = Bit32Value(31 - Shift)
    If op1.i0 And Mask Then
      BitsShiftRight64.i0 = (op1.i0 And (Mask - 1)) * Bit32Value(Shift) Or BitL_31
    Else
      BitsShiftRight64.i0 = (op1.i0 And (Mask - 1)) * Bit32Value(Shift)
    End If
    ' ShowDebugBit32L CopyBits
  End If
 
  ' move I1 right and add copy bits from I0
  'Mask = Bit32Value(31 - Shift)
  If BitsShiftRight64.i1 <> 0 Then
    If BitsShiftRight64.i1 And Mask Then
      BitsShiftRight64.i1 = (op1.i1 And (Mask - 1)) * Bit32Value(Shift) Or BitL_31 Or CopyBits
    Else
      BitsShiftRight64.i1 = (op1.i1 And (Mask - 1)) * Bit32Value(Shift) Or CopyBits
    End If
  End If
End Function

Public Function PopCnt64_V1(op1 As TBit64) As Long ' > LSET is slow! PopCnt64 is faster
  ' returns number of bits set
  If Not CBool(op1.i0 Or op1.i1) Then PopCnt64_V1 = 0: Exit Function
  LSet Int16x4 = op1 ' assign to 4x16 bit structure
  PopCnt64_V1 = Pop16Cnt(Int16x4.i0) + Pop16Cnt(Int16x4.i1) + Pop16Cnt(Int16x4.i2) + Pop16Cnt(Int16x4.i3)
End Function

Public Function PopCnt64(op1 As TBit64) As Long ' faster version without slow LSET
  ' returns number of bits set
  If op1.i0 = 0 Then If op1.i1 = 0 Then PopCnt64 = 0: Exit Function
  If op1.i0 <> 0 Then ' split into 16 bit parts
    PopCnt64 = PopCntL(op1.i0 And &HFFFF&) + Pop16Cnt((op1.i0 And &HFFFF0000) \ &H10000)
  End If
  If op1.i1 <> 0 Then
    PopCnt64 = PopCnt64 + PopCntL(op1.i1 And &HFFFF&) + Pop16Cnt((op1.i1 And &HFFFF0000) \ &H10000)
  End If
End Function

Public Function Lsb64(op1 As TBit64) As Long
 ' returns position of first bit set, order 0...63 (-1 if no bit set)
 Lsb64 = -1
 If op1.i0 <> 0 Then
   LSet Int16x4 = op1
   Lsb64 = LSB16(Int16x4.i0): If Lsb64 >= 0 Then Exit Function
   Lsb64 = LSB16(Int16x4.i1): If Lsb64 >= 0 Then Lsb64 = Lsb64 + 16: Exit Function
 ElseIf op1.i1 <> 0 Then
   LSet Int16x4 = op1
   Lsb64 = LSB16(Int16x4.i2): If Lsb64 >= 0 Then Lsb64 = Lsb64 + 32: Exit Function
   Lsb64 = LSB16(Int16x4.i3): If Lsb64 >= 0 Then Lsb64 = Lsb64 + 48
 End If
End Function

Public Function Rsb64(op1 As TBit64) As Long
 ' returns position of last bit set on right side , order 0...63 (-1 if no bit set)
 Rsb64 = -1
 If op1.i1 <> 0 Then
   LSet Int16x4 = op1
   Rsb64 = RSB16(Int16x4.i3): If Rsb64 >= 0 Then Rsb64 = Rsb64 + 48: Exit Function
   Rsb64 = RSB16(Int16x4.i2): If Rsb64 >= 0 Then Rsb64 = Rsb64 + 32: Exit Function
 ElseIf op1.i0 <> 0 Then
   LSet Int16x4 = op1
   Rsb64 = RSB16(Int16x4.i1): If Rsb64 >= 0 Then Rsb64 = Rsb64 + 16: Exit Function
   Rsb64 = RSB16(Int16x4.i0)
 End If
End Function

'
'--- additional 32 bit based functions
'
Public Function PopLsb64(op1 As TBit64) As Long
  ' returns position of left most bit and removes this bit
  PopLsb64 = Lsb64(op1)
  If PopLsb64 >= 0 Then ' clear bit
    If PopLsb64 < 32 Then op1.i0 = op1.i0 And Not Bit32Value(PopLsb64) Else op1.i1 = op1.i1 And Not Bit32Value(PopLsb64 - 32)
  End If
End Function

'
'--- additional functions to speedup some cases
'
Public Function MoreThanOne64(op1BB As TBit64) As Boolean
  ' returns MoreThanOne = (number of bits set in op1 > 1)
  MoreThanOne64 = CBool(PopCnt64(op1BB) > 1) ' more than one bit set
End Function

Public Sub SetAND64(op1 As TBit64, op2 As TBit64)
   ' returns op1 = op1 AND op2
  op1.i0 = op1.i0 And op2.i0: op1.i1 = op1.i1 And op2.i1
End Sub

Public Sub SetOR64(op1 As TBit64, op2 As TBit64)
   ' returns op1 = op1 OR op2
  op1.i0 = op1.i0 Or op2.i0: op1.i1 = op1.i1 Or op2.i1
End Sub

Public Sub ANDNOT64(Result As TBit64, op1 As TBit64, op2 As TBit64)
   ' returns op1 = op1 AND NOT op2
  Result.i0 = op1.i0 And Not op2.i0: Result.i1 = op1.i1 And Not op2.i1
End Sub

Public Function GetBitList64L(In64L As TBit64) As String
 ' Returns a string that shows the bits set in TBit64 structure (for debug)
 Dim i As Long, s As String
 GetBitList64L = ""
 For i = 0 To 63
   If IsBitSet64(In64L, i) Then GetBitList64L = GetBitList64L & "X" Else GetBitList64L = GetBitList64L & "."
   If i = 31 Then GetBitList64L = GetBitList64L & ":"
 Next
End Function


'============================================================================================
'=
'=            64 bit MS-Office implementation with 64 bit data type LongLong
'=
'============================================================================================

#If VBA7 And Win64 Then 'Note: Win64 = Office64 bit (not Windows 64 bit)

Public Sub Init64Bit()
  '--- init for lookup tables
  Dim i As Long, j As Long, k As Long, SqBB As Long, BitLL As LongLong
  
  For i = 0 To 31: Bit32Value(i) = BitMask32(i): Next
  For i = 0 To 63: Bit64ValueLL(i) = BitMask64(i): Next
  
  For j = MIN_INTEGER To MAX_INTEGER
    Pop16Cnt(j) = Pop16CountFkt(j)
    LSB16(j) = -1
    For i = 0 To 15
       If CBool(j And Bit32Value(i)) Then LSB16(j) = i: Exit For
    Next
    RSB16(j) = -1
    For i = 15 To 0 Step -1
       If CBool(j And Bit32Value(i)) Then RSB16(j) = i: Exit For
    Next
  Next
  
  For j = 0 To 65535 ' Popcount for Long data type and positive 16 bit values
    PopCntL(j) = Pop16CountLng(j)
  Next
  
End Sub
  

Function BitMask64(ByVal BitPos As Long) As LongLong
  ' compute bit value 0..63 bit, special case for 63 bit = sign (2 ^63 > overflow)
  If BitPos < 0 Or BitPos > 63 Then Err.Raise 6 ' overflow
  If BitPos < 63 Then
    BitMask64 = 2 ^ BitPos
  Else ' Bitpos 63 : 2 ^63 > overflow error
    BitMask64 = -9223372036854775807^ - 1  '  ^ = LongLong identifier / Strange: this shows an error ?!?  -9223372036854775808^
  End If
End Function
  
Public Sub ClearBit64LL(ValueLL As LongLong, ByVal BitPos As Long)
  '--- clear bit in op1 at BitPos 0 to 63
  'Debug.Assert BitPos >= 0 And BitPos <= 63
  ValueLL = ValueLL And Not Bit64ValueLL(BitPos)
End Sub

Public Sub SetBit64LL(ValueLL As LongLong, ByVal BitPos As Long)
  '--- set bit in op1 at BitPos 0 to 63
  'Debug.Assert BitPos >= 0 And BitPos <= 63
  ValueLL = ValueLL Or Bit64ValueLL(BitPos)
End Sub

Public Function IsBitSet64LL(ValueLL As LongLong, ByVal BitPos As Long) As Boolean
  '--- return  IsBitSet64 = (Is bit set at BitPos 0 to 63) ?
  'Debug.Assert BitPos >= 0 And BitPos <= 63
  IsBitSet64LL = CBool(ValueLL And Bit64ValueLL(BitPos))
End Function
  
  
Public Function BitsShiftLeftLL(ByVal ValueLL As LongLong, ByVal ShiftCount As Long) As LongLong
  '--- return BitsShiftLeftLL = shift BitsShiftLeftLL ShiftCount bits left (order 0...63)
 Select Case ShiftCount
  Case 1 To 63
    If ValueLL And Bit63LL Then ' minus sign = bit 63
      BitsShiftLeftLL = (((ValueLL And Not Bit63LL) \ 2) Or Bit62LL) \ Bit64ValueLL(ShiftCount - 1)
    Else
      BitsShiftLeftLL = ValueLL \ Bit64ValueLL(ShiftCount)
    End If
  Case 0
    BitsShiftLeftLL = ValueLL
  End Select
End Function

Public Function ShiftLeft64LL(ByVal ValueLL As LongLong) As LongLong
  '--- return ShiftLeft64LL = shift ShiftLeft64LL 1 bit left (order 0...63)
  If ValueLL And Bit63LL Then ' minus sign = bit 63
    ShiftLeft64LL = ((ValueLL And Not Bit63LL) \ 2) Or Bit62LL
  Else
    ShiftLeft64LL = ValueLL \ 2
  End If
End Function

Public Static Function BitsShiftRightLL(ByVal ValueLL As LongLong, ByVal ShiftCount As Long) As LongLong
  '--- return ShiftRight64 = shift oValueLL ShiftCount bits right (order 0...63)
  Dim MaskLL As LongLong
  Select Case ShiftCount
  Case 1 To 63
    MaskLL = Bit64ValueLL(63 - ShiftCount)
    If ValueLL And MaskLL Then
      BitsShiftRightLL = (ValueLL And (MaskLL - 1)) * Bit64ValueLL(ShiftCount) Or Bit63LL
    Else
      BitsShiftRightLL = (ValueLL And (MaskLL - 1)) * Bit64ValueLL(ShiftCount)
    End If
  Case 0
    BitsShiftRightLL = ValueLL
  End Select
End Function

Public Static Function ShiftRightLL(ByVal ValueLL As LongLong) As LongLong
  '--- return ShiftRight64 = shift oValueLL ShiftCount bits right (order 0...63)
  If ValueLL And Bit62LL Then
    ShiftRightLL = (ValueLL And (Bit62LL - 1)) * 2 Or Bit63LL
  Else
    ShiftRightLL = (ValueLL And (Bit62LL - 1)) * 2
  End If
End Function


Public Function PopCnt64LL(ValueLL As LongLong) As Long
  ' returns PopCnt64LL = number of bits set
  Dim Bit64TLL As TBit64LL
  If ValueLL = 0 Then PopCnt64LL = 0: Exit Function
  Bit64TLL.Bit64LL = ValueLL ' User defined type needed for LSET!!!
  LSet Int16x4 = Bit64TLL ' assign to 4x16 bit structure and add bit count 4 times
  PopCnt64LL = Pop16Cnt(Int16x4.i0) + Pop16Cnt(Int16x4.i1) + Pop16Cnt(Int16x4.i2) + Pop16Cnt(Int16x4.i3)
End Function

Public Function Lsb64LL(ValueLL As LongLong) As Long
 ' returns Lsb64LL = position of first bit set (order 0..63)
  Dim Bit64TLL As TBit64LL
  If ValueLL = 0 Then Lsb64LL = -1: Exit Function
  Bit64TLL.Bit64LL = ValueLL ' User defined type needed for LSET!!!
  LSet Int16x4 = Bit64TLL  ' assign to 4x16 bit structure
  Lsb64LL = LSB16(Int16x4.i0): If Lsb64LL >= 0 Then Exit Function
  Lsb64LL = LSB16(Int16x4.i1): If Lsb64LL >= 0 Then Lsb64LL = Lsb64LL + 16: Exit Function
  Lsb64LL = LSB16(Int16x4.i2): If Lsb64LL >= 0 Then Lsb64LL = Lsb64LL + 32: Exit Function
  Lsb64LL = LSB16(Int16x4.i3): If Lsb64LL >= 0 Then Lsb64LL = Lsb64LL + 48
End Function

Public Function Rsb64LL(ValueLL As LongLong) As Long
 ' returns Rsb64LL = position of last bit set (order 0..63)
  Dim Bit64TLL As TBit64LL
  If ValueLL = 0 Then Rsb64LL = -1: Exit Function
  Bit64TLL.Bit64LL = ValueLL ' User defined type needed for LSET!!!
  LSet Int16x4 = Bit64TLL  ' assign to 4x16 bit structure
  Rsb64LL = RSB16(Int16x4.i3): If Rsb64LL >= 0 Then Rsb64LL = Rsb64LL + 48: Exit Function
  Rsb64LL = RSB16(Int16x4.i2): If Rsb64LL >= 0 Then Rsb64LL = Rsb64LL + 32: Exit Function
  Rsb64LL = RSB16(Int16x4.i1): If Rsb64LL >= 0 Then Rsb64LL = Rsb64LL + 16: Exit Function
  Rsb64LL = RSB16(Int16x4.i0)
End Function

Public Function PopLsb64LL(ValueLL As LongLong) As Long
 ' returns PopLsb64LL = position of first bit set (order 0..63)  and removes this bit from ValueLL
  If ValueLL = 0 Then PopLsb64LL = -1: Exit Function
  PopLsb64LL = Lsb64LL(ValueLL)
  ValueLL = ValueLL And Not Bit64ValueLL(PopLsb64LL)
End Function

Public Function MoreThanOne64LL(ValueLL As LongLong) As Boolean
  ' returns MoreThanOne64LL = (number of bits set in op1 > 1)
  MoreThanOne64LL = CBool(PopCnt64LL(ValueLL) > 1) ' more than one bit set
End Function

Public Function GetBitList64LL(In64LL As LongLong) As String
 ' Returns a string that shows the bits set in LongLong (for debug)
 Dim i As Long, s As String
 GetBitList64LL = ""
 For i = 0 To 63
   If IsBitSet64LL(In64LL, i) Then GetBitList64LL = GetBitList64LL & "X" Else GetBitList64LL = GetBitList64LL & "."
 Next
End Function


'================================================================================================
'= Obsolete functions for LongLong:
'= no functions needed because LongLong 64 bit=> write operations directly in source
'================================================================================================
'Public Sub Clear64LL(ValueLL As LongLong)
'  ValueLL = 0
'End Sub
'
'Public Sub AND64LL(ResultLL As LongLong, op1LL As LongLong, op2LL As LongLong)
'  '--- return resultLL = op1LL AND op2LL
'  ResultLL = op1LL And op2LL
'End Sub
'
'Public Sub OR64LL(ResultLL As LongLong, op1LL As LongLong, op2LL As LongLong)
'  '--- return resultLL = op1LL OR op2LL
'  ResultLL = op1LL Or op2LL
'End Sub
'
'Public Sub XOr64LL(ResultLL As LongLong, op1LL As LongLong, op2LL As LongLong)
'  '--- return resultLL = op1 XOR op2
'  ResultLL = op1LL Or op2LL
'End Sub
'
'Public Sub NOT64LL(ResultLL As LongLong, op1LL As LongLong)
'  '--- return resultLL = NOT op1
'  ResultLL = Not op1LL
'End Sub
'
'Public Function Equal64LL(op1LL As LongLong, op2LL As LongLong) As Boolean
'  '--- return Equal64LL = (op1 = op2)  > is op1 equal op2 ?
'  Equal64LL = CBool(op1LL = op2LL)
'End Function
'
#End If ' VBA7 And Win64

