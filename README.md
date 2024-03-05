# VBA7_ULongLong
Emulate ULongLong 64 bit unsigned integers with VBA7 (64/32 bit MsOffice) or VB6 

List of functions in module 'bas64BitOperations'	

Emulate ULongLong 64 bit unsigned integer operations	

call Init32Bit before	call Init64Bit before	
32-bit mode		64-bit mode	    Some logic needed to handle the sign bit
================================================================================
Clear64			  LL=0	
ClearBit64		ClearBit64LL	
SetBit64		  SetBit64LL	
IsBitSet64		IsBitSet64LL	
Let64			    L1 = L2	
AND64			    LL = LL1 AND LL2	
OR64			    LL = LL1 OR LL2	
XOr64			    LL = LL1 XOR LL2	
NOT64			    LL = NOT LL1	
ANDNOT64		  LL = LL1 AND NOT LL2	for speed only (example)
Equal64			  (LL1 = LL2)	
IsNotEmpty64	LL1 <> 0	
ShiftLeft64		ShiftLeft64LL		      shift left  1 bit
ShiftRight64	ShiftRight64LL		    shift right  1 bit
BitsShiftLeft64		BitsShiftLeftLL		shift left  1 to 63 bits
BitsShiftRight64	BitsShiftRightLL	shift right  1 to 63 bits
PopCnt64		  PopCnt64LL		returns number of bits set
Lsb64			    Lsb64LL			returns position of left most bit
Rsb64			    Rsb64LL			returns position of right most bit
PopLsb64		  PopLsb64LL		returns position of left most bit and removes it
MoreThanOne64	PopCnt64LL>1		for speed only (example)
SetAND64		  LL = LL AND LL1		for speed only (example)
SetOR64			  LL = LL OR LL1		for speed only (example)

