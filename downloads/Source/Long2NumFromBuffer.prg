Function Long2NumFromBuffer(tnPointer)
	Local lnNum
	lnNum = 0
	= RtlP2PL(@lnNum, tnPointer, 4)
	Return lnNum
Endfunc

Function RtlP2PL(tnDest, tnSrc, tnLen)
	Declare RtlMoveMemory In WIN32API As RtlP2PL Long @Dest, Long Source, Long Length
	Return 	RtlP2PL(@tnDest, tnSrc, tnLen)
EndFunc 
