LPARAMETERS toParameter

LOCAL lnSelect, lcCode, llReturn, lScriptHandled

TRY
	* First try FoxCode lookup for Type="M" records
	lnSelect = SELECT()
	SELECT 0
	USE (_FOXCODE) AGAIN SHARE ORDER 1
	IF SEEK('M' + PADR(UPPER(toParameter.MenuItem), LEN(ABBREV)))
		lcCode = DATA
	ENDIF
	USE
	SELECT (lnSelect)
	IF NOT EMPTY(lcCode)
		llReturn = EXECSCRIPT(lcCode, toParameter)
		lScriptHandled=.T.
	ENDIF

	* Handle by passing to external routine as specified in Tip field
	IF !lScriptHandled
		lcProgram = ALLTRIM(toParameter.Tip)
		IF FILE(lcProgram)
			DO (lcProgram) WITH toParameter,llReturn
		ENDIF
	ENDIF

	* Custom script successful so let's disable native behavior
	IF llReturn
		toParameter.ValueType = 'V'
	ENDIF
CATCH
ENDTRY

RETURN llReturn
