Lparameters lcText

Local lcInnerText, lcNextChar, lcThisChar, lnCount, lnEndPos, lnStartPos
lcText = Chrtran(lcText, '()', '[]')
Do While .T.
	lnStartPos = Min (AtNew(['], lcText), AtNew(["], lcText), AtNew(']', lcText))
	If lnStartPos > Len(lcText)
		Return lcText
	Endif

	lcThisChar = Substr (lcText, lnStartPos, 1)
	If lcThisChar $ ['"]
		lnEndPos = AtNew(lcThisChar, lcText, 2)
		lcText	 = Left (lcText, lnStartPos - 1) + ' ' + Substr (lcText, lnEndPos + 1)
	Else
		lnEndPos   = lnStartPos
		lnStartPos = Rat('[', lcText)
		lcText	   = Left (lcText, lnStartPos - 1) + '( ' + Substr (lcText, lnEndPos + 1)
	Endif
Enddo && While .T.


Procedure AtNew
	Lparameters lcSearchfor, lcSearched, lnOccurrence

	Return Evl (At (lcSearchfor, lcSearched, Evl(lnOccurrence, 1)), 1 + Len (lcSearched))
