# VBA7_ULongLong
Emulate ULongLong 64 bit unsigned integers with VBA7 (64/32 bit MsOffice) or VB6 (Visual Basic 6)

-------------------------------------------------------------------------------------------------

The problem with VBA/VB6 is that there are no 64-bit unsigned integers (ULongLong) available.

The module 'bas64BitOperations.bas' contains functions to emulate bit operations for 64-bit unsigned integers.

For 32-bit MsOffice VBA or VB6 we can use two Long 32-bit signed integers combined in a new datatype TBit64.
 - this requires functions for basic bit operations like AND, OR, XOR, NOT.

For 64-bit MsOffice VBA7 we can use the new 64-bit signed integer datatype LongLong.
 - basic bit operations like AND, OR, XOR, NOT do NOT need extra functions but can be used directly.

Functions available for both variants are:

 - SetBit64, ClearBit64, ShiftLeft64, ShiftRight64, PopCnt64 (number of bits set), 

   Lsb64 (position of left most bit), Rsb64 (position of left most bit) and some more.
   

The Ms-Excel file VBA7.xlsm shows how to use these functions for 32/64-bit MsOffice.

For 64-bit Excel there is also a simple chess demonstration how to use 64-bit chessboard logic.

(for the best Excel/VB6 chess program available please see my project at github.com/RZulu54/ChessBrainVB)


