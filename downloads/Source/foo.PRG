Lparameters lcText

Local lcChar, lnEndPos, lnPos, lnStartPos
Do While .T.
	lnEndPos = Max (Rat (['], lcText), Rat (["], lcText), Rat (')', lcText), Rat (']', lcText))
	If lnEndPos = 0
		Return lcText
	Endif
	lcChar = Substr (lcText, lnEndPos, 1)
	Do Case
		Case lcChar $ ['"]
			lnStartPos = Rat (lcChar, lcText, 2)
		Case lcChar $ ')'
			lnStartPos = Rat ('(', lcText)
		Case lcChar $ ']'
			lnStartPos = Rat ('[', lcText)
	Endcase
	lcText = Left (lcText, lnStartPos - 1) + ' ' + Substr (lcText, lnEndPos + 1)
Enddo && While .T.
