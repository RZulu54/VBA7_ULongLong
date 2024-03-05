Attribute VB_Name = "bas64BitTest"
'=========================================================================
'= Test functions for 64bit operations
'=========================================================================
Option Explicit

Public psSetting As String  ' Switch 64 / 32 bit functions

#If VBA7 And Win64 Then 'Note: Win64 = Office64 bit (not Windows 64 bit)
  '-----------------------------
  '--- 64 bit  LongLong        -
  '-----------------------------
  Public Sub Test_ShiftLeft1()
    Test64 "SHIFTLEFT1"
  End Sub
  
  Public Sub Test_ShiftLeft2()
    Test64 "SHIFTLEFT2"
  End Sub
  
  Public Sub Test_ShiftRight1()
    Test64 "SHIFTRIGHT1"
  End Sub
  
  Public Sub Test_ShiftRight2()
    Test64 "SHIFTRIGHT2"
  End Sub
  
  Public Sub Test_PopCnt()
    Test64 "POPCNT"
  End Sub
  
Public Sub Test64All()
  Dim i As Long, j As Long
  Dim Bit64LL As LongLong, BitAct64LL As LongLong, BitNew64LL As LongLong, BitShifts As Long
  
  ClearOutBox ' clear the test output box
  
  Init64Bit
  
  For BitShifts = 0 To 64
    For j = 0 To 2
      Bit64LL = 0
      For i = 0 To 63
        If (i Mod 2 = j) Or (j = 2) Then
          SetBit64LL Bit64LL, i
        End If
      Next
      
      ' shift left 0-64 bitshiftsl
      If BitShifts = 0 Then WriteTestOutput "testing 64-bit LongLong: BitsShiftLeft64  with pattern: " & Left(GetBitList64LL(Bit64LL), 16)
      BitAct64LL = Bit64LL
      For i = 0 To 63
        BitNew64LL = BitsShiftLeftLL(BitAct64LL, BitShifts)
        VerifyBitShiftLL BitAct64LL, BitNew64LL, -BitShifts ' minus for left shift
        BitAct64LL = BitNew64LL '--- next shift
      Next
      
      ' single shift left
      If BitShifts = 0 Then WriteTestOutput "testing 64-bit LongLong: ShiftLeft64      with pattern: " & Left(GetBitList64LL(Bit64LL), 16)
      If BitShifts = 1 Then
        BitAct64LL = Bit64LL
        For i = 0 To 63
          BitNew64LL = ShiftLeft64LL(BitAct64LL)
          VerifyBitShiftLL BitAct64LL, BitNew64LL, -BitShifts ' minus for left shift
          BitAct64LL = BitNew64LL '--- next shift
        Next
      End If
          
      ' shift right 0-64 bitshifts
      If BitShifts = 0 Then WriteTestOutput "testing 64-bit LongLong: BitsShiftRight64 with pattern: " & Left(GetBitList64LL(Bit64LL), 16)
      BitAct64LL = Bit64LL
      For i = 0 To 63
        BitNew64LL = BitsShiftRightLL(BitAct64LL, BitShifts)
        VerifyBitShiftLL BitAct64LL, BitNew64LL, BitShifts
        BitAct64LL = BitNew64LL '--- next shift
      Next
      
      ' single shift right
      If BitShifts = 0 Then WriteTestOutput "testing 64-bit LongLong: ShiftRight64     with pattern: " & Left(GetBitList64LL(Bit64LL), 16)
      If BitShifts = 1 Then
        BitAct64LL = Bit64LL
        For i = 0 To 63
          BitNew64LL = ShiftRightLL(BitAct64LL)
          VerifyBitShiftLL BitAct64LL, BitNew64LL, BitShifts
          BitAct64LL = BitNew64LL '--- next shift
        Next
      End If
   Next j
 Next BitShifts
 DoEvents
 MsgBox "64-bit Test OK!"
End Sub
    
#Else

  '-----------------------------
  '--- 32 bit  2xLong          -
  '-----------------------------
  
  Public Sub Test_ShiftLeft1()
    Test32 "SHIFTLEFT1"
  End Sub
  
  Public Sub Test_ShiftLeft2()
    Test32 "SHIFTLEFT2"
  End Sub

  Public Sub Test_ShiftRight1()
    Test32 "SHIFTRIGHT1"
  End Sub

  Public Sub Test_ShiftRight2()
    Test32 "SHIFTRIGHT2"
  End Sub
  
#End If  '---- Win64
  
  
  Public Sub Test_SwitchTo64Bit()
   #If VBA7 And Win64 Then 'Note: Win64 = Office64 bit (not Windows 64 bit)
     psSetting = "64bit"
     Worksheets("Test").Range("A2") = "Test setting: 64 bit functions active"
   #Else
     MsgBox "Not possible: 64Bit MsOffice needed!"
     Test_SwitchTo32Bit
   #End If
  End Sub
  
  Public Sub Test_SwitchTo32Bit()
    psSetting = "32bit"
    Worksheets("Test").Range("A2") = "Test setting: 32 bit functions active"
  End Sub
  
  Public Sub Test_FullTest64Bit()
   #If VBA7 And Win64 Then 'Note: Win64 = Office64 bit (not Windows 64 bit)
     psSetting = "64bit"
     Worksheets("Test").Range("A2") = "Test setting: 64 bit functions active"
     Test64All
   #Else
     MsgBox "Not possible: 64Bit MsOffice needed!"
   #End If
  End Sub



Public Sub ClearOutBox()
 ' clear test output and write header
 '                                 0123456789012345678901234567890123456789012345678901234567890123
 Worksheets("Test").txtOut.Text = "0_______8_______16______24______32______40______48______56_____63 (Bit Count  LeftPos RightPos)"
End Sub

#If VBA7 And Win64 Then 'Note: Win64 = Office64 bit (not Windows 64 bit)

Public Sub ShowBit64LL(bbInLL As LongLong)
 Dim i As Long, s As String, sLine As String
 sLine = ""
 For i = 0 To 63
   If bbInLL And Bit64ValueLL(i) Then
    sLine = sLine & "X"
   Else
    sLine = sLine & "."
   End If
 Next
 sLine = sLine & "  (PopCnt:" & Right$("  " & PopCnt64LL(bbInLL), 2) & "  LSB:" & Right$("  " & Lsb64LL(bbInLL), 2) & "  RSB:" & Right$("  " & Rsb64LL(bbInLL), 2) & ")"
 WriteTestOutput sLine
End Sub


Public Sub ShowDebugBit64LL(bbInLL As LongLong)
 Dim i As Long, s As String, sLine As String
 sLine = ""
 For i = 0 To 63
   If bbInLL And Bit64ValueLL(i) Then
    sLine = sLine & "X"
   Else
    sLine = sLine & "."
   End If
 Next
 Debug.Print sLine
End Sub


Public Sub Test64(ByVal sTestCase As String)
  Dim i As Long, j As Long
  Dim Bit64LL As LongLong, Bit64aLL As LongLong, BitShifts As Long
  
  ClearOutBox ' clear the test output box
  
  BitShifts = Val(Worksheets("Test").cboBits)
  
  Init64Bit
  
  Select Case sTestCase
  Case "SHIFTLEFT1", "SHIFTLEFT2", "SHIFTRIGHT1", "SHIFTRIGHT2"
    '--- Test ShiftLeft , ShiftRight
    If Right$(sTestCase, 1) = "1" Then j = 0 Else j = 1 ' start with bit 0 or bit 1, end with bit 62 or bit 32
    Bit64LL = 0
    For i = 0 To 63
     If i Mod 2 = j Then
       Bit64LL = Bit64LL Or Bit64ValueLL(i)
     End If
    Next
    ShowBit64LL Bit64LL
    
    Bit64aLL = Bit64LL
    For i = 0 To 63
      If Left$(sTestCase, 9) = "SHIFTLEFT" Then
        Bit64aLL = BitsShiftLeftLL(Bit64aLL, BitShifts): ShowBit64LL Bit64aLL
      ElseIf Left$(sTestCase, 10) = "SHIFTRIGHT" Then
       Bit64aLL = BitsShiftRightLL(Bit64aLL, BitShifts): ShowBit64LL Bit64aLL
      End If
    Next i
    'Debug.Print String(90, "=")
  
  Case "POPCNT"
    '--- Test POPCNT
    Bit64LL = 0: ShowBit64LL Bit64LL
    For i = 0 To 63
      Bit64LL = Bit64LL Or Bit64ValueLL(i): ShowBit64LL Bit64LL
    Next
    Debug.Print String(90, "=")
  
    Bit64LL = 0: ShowBit64LL Bit64LL
    For i = 63 To 0 Step -1
       Bit64LL = Bit64LL Or Bit64ValueLL(i): ShowBit64LL Bit64LL
    Next
    'Debug.Print String(90, "=")
  
  End Select

End Sub


Public Sub VerifyBitShiftLL(Old64LL As LongLong, New64LL As LongLong, ByVal BitShifts As Long)
  Dim OldPos As Long, NewPos As Long
  
  For OldPos = 0 To 63
    If OldPos <= 31 Then
       NewPos = OldPos + BitShifts ' Bitshifts negative for left shift, positive for right shift
       If NewPos >= 0 And NewPos <= 63 Then
         If IsBitSet64LL(Old64LL, OldPos) <> IsBitSet64LL(New64LL, NewPos) Then MsgBox "Error": Stop
       End If
    End If
  Next OldPos
End Sub


#End If  '<<<<<<<<  VBA7 And Win64


'-----------------------------
'--- 32 bit  2x32 Long       -
'-----------------------------

Public Sub ShowBit64L(bbInL As TBit64)
 Dim sLine As String
 sLine = GetBitList64L(bbInL) & "  (PopCnt:" & Right$("  " & PopCnt64(bbInL), 2) & "  LSB:" & Right$("  " & Lsb64(bbInL), 2) & "  RSB:" & Right$("  " & Rsb64(bbInL), 2) & ")"
 WriteTestOutput sLine
End Sub

Public Sub ShowDebugBit64L(bbInL As TBit64)
 Dim sLine As String
 sLine = GetBitList64L(bbInL) & "  (PopCnt:" & Right$("  " & PopCnt64(bbInL), 2) & "  LSB:" & Right$("  " & Lsb64(bbInL), 2) & "  RSB:" & Right$("  " & Rsb64(bbInL), 2) & ")"
 Debug.Print sLine
End Sub

Public Sub WriteTestOutput(isLine As String)
 Worksheets("Test").txtOut.Text = Worksheets("Test").txtOut.Text & vbNewLine & isLine
End Sub

Public Sub ShowDebugBit32L(bbInL As Long)
 Dim i As Long, s As String, sLine As String
 sLine = ""
 For i = 0 To 31
   If bbInL And Bit32Value(i) Then
    sLine = sLine & "X"
   Else
    sLine = sLine & "."
   End If
   If i = 31 Then sLine = sLine & ":"
 Next
 Debug.Print sLine
End Sub

Public Sub Test32(ByVal sTestCase As String)
  Dim i As Long, j As Long
  Dim Bit64L As TBit64, Bit64aL As TBit64, BitShifts As Long
  
  ClearOutBox ' clear the test output box
  
  'Debug.Print String(255, vbNewLine) ' Clear Debug
  
  BitShifts = Val(Worksheets("Test").cboBits)
  
  Init32Bit
  
  Select Case sTestCase
  Case "SHIFTLEFT1", "SHIFTLEFT2", "SHIFTRIGHT1", "SHIFTRIGHT2"
    '--- Test ShiftLeft , ShiftRight> fill start pattern
    If Right$(sTestCase, 1) = "1" Then j = 0 Else j = 1 ' start with bit 0 or bit 1, end with bit 62 or bit 32
    Clear64 Bit64L
    For i = 0 To 63
      If i Mod 2 = j Then
        SetBit64 Bit64L, i
      End If
    Next
    ShowBit64L Bit64L
    
    Bit64aL = Bit64L
    For i = 0 To 63
      If Left$(sTestCase, 9) = "SHIFTLEFT" Then
        Bit64aL = BitsShiftLeft64(Bit64aL, BitShifts): ShowBit64L Bit64aL
      '  Bit64aL = ShiftLeft64(Bit64aL): ShowBit64L Bit64aL
      ElseIf Left$(sTestCase, 10) = "SHIFTRIGHT" Then
        Bit64aL = BitsShiftRight64(Bit64aL, BitShifts): ShowBit64L Bit64aL
       ' Bit64aL = ShiftRight64(Bit64aL): ShowBit64L Bit64aL
      End If
    Next i
    'Debug.Print String(90, "=")
  
 
  End Select

End Sub

Public Sub Test32All()
  Dim i As Long, j As Long
  Dim Bit64L As TBit64, BitAct64L As TBit64, BitNew64L As TBit64, BitShifts As Long
  
  ClearOutBox ' clear the test output box
  'Debug.Print String(255, vbNewLine) ' Clear Debug
  
  
  Init32Bit
  
  For BitShifts = 0 To 64
    For j = 0 To 2
      'Debug.Print "Testing " & BitShifts & " V" & j
        Clear64 Bit64L
        For i = 0 To 63
          If (i Mod 2 = j) Or (j = 2) Then
            SetBit64 Bit64L, i
          End If
        Next
        
        ' shift left 0-64 bitshifts
        If BitShifts = 0 Then WriteTestOutput "testing 2x32-bit Long: BitsShiftLeft64  with pattern: " & Left(GetBitList64L(Bit64L), 16)
        BitAct64L = Bit64L
        For i = 0 To 63
          BitNew64L = BitsShiftLeft64(BitAct64L, BitShifts)
          VerifyBitShift BitAct64L, BitNew64L, -BitShifts ' minus for left shift
          BitAct64L = BitNew64L '--- next shift
        Next
        
        ' single shift left
        If BitShifts = 0 Then WriteTestOutput "testing 2x32-bit Long: ShiftLeft64      with pattern: " & Left(GetBitList64L(Bit64L), 16)
        If BitShifts = 1 Then
          BitAct64L = Bit64L
          For i = 0 To 63
             BitNew64L = ShiftLeft64(BitAct64L)
             VerifyBitShift BitAct64L, BitNew64L, -BitShifts ' minus for left shift
             BitAct64L = BitNew64L '--- next shift
          Next
        End If
            
        ' shift right 0-64 bitshifts
        If BitShifts = 0 Then WriteTestOutput "testing 2x32-bit Long: BitsShiftRight64 with pattern: " & Left(GetBitList64L(Bit64L), 16)
        BitAct64L = Bit64L
        For i = 0 To 63
          BitNew64L = BitsShiftRight64(BitAct64L, BitShifts)
          VerifyBitShift BitAct64L, BitNew64L, BitShifts
          BitAct64L = BitNew64L '--- next shift
        Next
        
        ' single shift right
        If BitShifts = 0 Then WriteTestOutput "testing 2x32-bit Long: ShiftRight64     with pattern: " & Left(GetBitList64L(Bit64L), 16)
        If BitShifts = 1 Then
          BitAct64L = Bit64L
          For i = 0 To 63
             BitNew64L = ShiftRight64(BitAct64L)
             VerifyBitShift BitAct64L, BitNew64L, BitShifts
             BitAct64L = BitNew64L '--- next shift
          Next
        End If
   Next j
 Next BitShifts
 DoEvents
 MsgBox "32-bit Test OK!"
End Sub

Public Sub VerifyBitShift(Old64L As TBit64, New64L As TBit64, ByVal BitShifts As Long)
  Dim OldPos As Long, NewPos As Long
  
  For OldPos = 0 To 63
    If OldPos <= 31 Then
       NewPos = OldPos + BitShifts ' Bitshifts negative for left shift, positive for right shift
       If NewPos >= 0 And NewPos <= 63 Then
         If IsBitSet64(Old64L, OldPos) <> IsBitSet64(New64L, NewPos) Then MsgBox "Error": Stop
       End If
    End If
  Next OldPos
End Sub

Public Sub TestPopcnt() ' 32 bit
  Dim t As TBit64, i As Long, j As Long, x As Long, Cnt As Long
  '
  Init32Bit
  
 For i = 0 To 63
  For j = 0 To 63
   If i <> j Then
    Clear64 t
    SetBit64 t, i
    SetBit64 t, j
    Cnt = Cnt + 1
    If PopCnt64(t) <> PopCnt64_V1(t) Then Stop
   End If
  Next
 Next
 Debug.Print "OK: " & Cnt
End Sub

Public Sub TestDecimal()
'--- Arithmetic operation with very large values> Variant as type Decimal

'Public Type DecimalStructure ' (when sitting in a Variant)
'    '
'    ' +, -, *, /    Decimal is at the top of the hierarchy.
'    '               The hierarchy is: Byte, Integer, Long, Single, Double, Currency, Decimal.
'    ' \, Mod        Results is Long, not Decimal.
'    ' ^             Results is Double, not Decimal.
'    ' Int()         Works on Decimals.
'    ' Fix()         Works on Decimals.
'    ' Abs()         Works on Decimals.
'    ' Sgn()         Works on Decimals, although result is an Integer.
'    ' Sqr(), etc.   Most of these functions (including trig) return Double, not Decimal.
'    ' AND,OR,etc.   These convert to Long, which may cause overflow. <<<<<<<<<<<<<<<<<<<<<<<<<<<<################
'    '
'    ' Largest Decimal:      +/- 79228162514264337593543950335.  2^96-1 (sign bit handled separately)
'    ' Smallest Decimal:     +/- 0.0000000000000000000000000001  Notice that both largest and smallest are same width.
'    '
'    VariantType As Integer  ' Reserved, to act as the Variant Type when sitting in a 16-Byte-Variant.  Equals vbDecimal(14) when it's a Decimal type.
'    Base10NegExp As Byte    ' Base 10 exponent (0 to 28), moving decimal to right (smaller numbers) as this value goes higher.  Top three bits are never used.
'    sign As Byte            ' Sign bit only.  Other bits aren't used.
'    Hi32 As Long            ' Mantissa.
'    Lo32 As Long            ' Mantissa.
'    Mid32 As Long           ' Mantissa.
'End Type


Dim V1 As Variant
Dim V2 As Variant
Dim V3 As Variant

V1 = CDec("-9223372036854775808")
V2 = CDec(2 ^ 63)

V3 = V2 / 2#
Debug.Print V1, V2, V3

V3 = V1 * 2#
Debug.Print V1, V2, V3

V3 = V2 + 2#
Debug.Print V1, V2, V3
 ' Debug.Print V2 And 63#  >> overflow!
End Sub
