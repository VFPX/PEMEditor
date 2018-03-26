#include foxpro.h

	#Define ccTab  	Chr(9)
	#Define ccLF	Chr(10)
	#Define ccCR	Chr(13)
	#Define ccCRLF	Chr(13) + Chr(10)

Lparameters tcCode

* Returns a collection indicating the beginning of each procedure / class / etc
* Each member in the collection has these properties:
*   .Type == 'Procedure' (Procedures and Functions)
*         == 'Class'     (Class Definition)
*         == 'End Class' (End of Class Definition)
*         == 'Method'    (Procedures and Functions within a class)
*         == or any of the following:
*					DO CASE, ENDCASE, DO WHILE, ENDDO, IF, ELSE, ENDIF, FOR, ENDFOR, SCAN, ENDSCAN, TRY, CATCH, ENDTRY, WITH, ENDWITH
*   .LineNumber == starts at zero
*   .Name
*   .ClassName
*   .Text      == text of the line
*   .ProcedureBoundary == .T. for Procedures, Function, Classes, etc.
*   .StartStructure    == .T. for beginning of control structures (IF, DO CASE, WITH, etc.)
*   .EndStructure      == .T. for ending of control structures (ENDIF, ENDCASE, ENDWITH, etc.)

****************************************************************

loResult = Createobject ('Collection')
****************************************************************

loObject = This.GetStructurePositionObject()
With loObject
	.LineNumber		   = -1
	.Type			   = 'Procedure'
	.ProcedureBoundary = .T.
Endwith
lcName = 'Top Of Code'
loResult.Add (loObject, lcName)
****************************************************************

loRegExp = Createobject ('VBScript.RegExp')
With loRegExp
	.IgnoreCase	= .T.
	.Global		= .T.
	.MultiLine	= .T.
Endwith

****************************************************************

lcCodeWithLFs = tcCode
Do While .T.
	loRegExp.Pattern = '\x0D[^\x0A]' && CRs not followed by LFs
	loMatches = loRegExp.Execute(lcCodeWithLFs)

	If loMatches.Count = 0
		Exit
	Else
		For lnI = loMatches.Count To 1 Step -1
			lnStartByte = loMatches.Item(lnI - 1).FirstIndex
			lcCodeWithLFs = Stuff(lcCodeWithLFs, lnStartByte + 1, 1, Chr(13) + Chr(10))
		Endfor
	Endif
Enddo
****************************************************************
lcPattern = This.GetControlStructurePattern()

With loRegExp
	.Pattern	= '^\s*(' + lcPattern + ')'
Endwith

loMatches = loRegExp.Execute (lcCodeWithLFs)
****************************************************************

llClassDef	  = .F. && currently within a class?
llTextEndText = .F. && currently within a Text/EndText block?
lcClassName	  = ''

lcMatchStartWords = ' TEXT DOCASE DOWHILE IF FOR SCAN TRY WITH'
lcMatchEndWords	  = ' ENDTEXT ENDCASE ENDDO ENDIF ENDFOR NEXT ENDSCAN ENDTRY ENDWITH ENDPROC ENDFUNC ENDDEFINE'
lcMatchOtherWords = ' ELSE CATCH CASE OTHERWISE'

For lnI = 1 To loMatches.Count

	* .Value
	* .FirstIndex
	* .Length
	With loMatches.Item (lnI - 1)

		* ignore leading CRLF's
		lcLeading 	   = Left( .Value, At(Getwordnum( Chrtran(.Value, CR + LF + Tab, '   '), 1), .Value) - 1) && text before first word
		lnLeadingCRLFs = Max (Rat (CR, lcLeading), Rat (LF, lcLeading))

		lnLineNumber	= Getwordcount(Left(lcCodeWithLFs, .FirstIndex + lnLeadingCRLFs), LF)
		lcOrigMatch = Substr(.Value, lnLeadingCRLFs + 1)
		lcMatch		= Chrtran (lcOrigMatch, CR + LF + Tab, '   ')
		lcName		= Getwordnum (lcMatch, Getwordcount (lcMatch))
		lcWord1		= Upper (Getwordnum (lcMatch, Max(1, Getwordcount (lcMatch) - 1)))

		lcKeyWord = Upper (Getwordnum (lcMatch, 1))
		lcKeyWord = lcKeyWord + IIf (lcKeyWord = 'DO', Upper (Getwordnum (lcMatch, 2)), '')
	Endwith

	llProcedureBoundary	= .F.
	llStartStructure	= .F.
	llEndStructure		= .F.
	llClosingExtraText  = .F.

	Do Case
		Case llTextEndText
			If 'ENDTEXT' = lcKeyWord
				lcName		   = ''
				lcType		   = 'ENDTEXT'
				llEndStructure = .T.
				llTextEndText  = .F.
			Else
				Loop
			Endif

		Case 'TEXT' = lcKeyWord
			lcName			 = ''
			lcType			 = 'TEXT'
			llStartStructure = .T.
			llTextEndText	 = .T.
			llClosingExtraText  = .T.

		Case 'ENDDEFINE' = lcWord1
				llClassDef	= .F.
				lcType		= 'End Class'
				lcName		= lcClassName + '.-EndDefine'
				lcClassName	= ''
			llEndStructure = .T.
			llProcedureBoundary	= .T.

		Case 'CLASS' = lcWord1
			llClassDef			= .T.
			lcType				= 'Class'
			lcClassName			= lcName
			llProcedureBoundary	= .T.
			llStartStructure = .T.

		*!* * Removed 06/23/2011 
		*!* Case llClassDef
		*!* 	If 'ENDDEFINE' = lcWord1
		*!* 		llClassDef	= .F.
		*!* 		lcType		= 'End Class'
		*!* 		lcName		= lcClassName + '.-EndDefine'
		*!* 		lcClassName	= ''
		*!* 	Else
		*!* 		lcType = 'Method'
		*!* 		lcName = lcClassName + '.' + lcName
		*!* 	Endif
		*!* 	llProcedureBoundary	= .T.

		*!* Case 'CLASS' = lcWord1
		*!* 	llClassDef			= .T.
		*!* 	lcType				= 'Class'
		*!* 	lcClassName			= lcName
		*!* 	llProcedureBoundary	= .T.

		Case (' ' + lcKeyWord) $ lcMatchStartWords
			lcName			 = ''
			lcType			 = Getwordnum (Substr (lcMatchStartWords, At (' ' + lcKeyWord, lcMatchStartWords) + 1), 1)
			llStartStructure = .T.
			llClosingExtraText  = Inlist(lcType, 'DO', 'IF', 'FOR', 'SCAN', 'TRY', 'WITH')

		Case (' ' + lcKeyWord) $ lcMatchEndWords
			lcName		   = ''
			lcType		   = Getwordnum (Substr (lcMatchEndWords, At (' ' + lcKeyWord, lcMatchEndWords) + 1), 1)
			llEndStructure = .T.

		Case (' ' + lcKeyWord) $ lcMatchOtherWords
			lcName = ''
			lcType = Getwordnum (Substr (lcMatchOtherWords, At (' ' + lcKeyWord, lcMatchOtherWords) + 1), 1)

		Case Upper(lcWord1) = 'PROC'
			lcType				= 'Procedure'
			llProcedureBoundary	= .T.
			llEndStructure = .T.

		Otherwise
			lcType				= 'Function'
			llProcedureBoundary	= .T.
			llEndStructure = .T.

	Endcase

	loObject = This.GetStructurePositionObject()

	With loObject
		.LineNumber		   = lnLineNumber
		.Type			   = lcType
		.Name			   = lcName
		.ClassName		   = lcClassName
		.Text			   = lcOrigMatch
		.ProcedureBoundary = llProcedureBoundary
		.StartStructure	   = llStartStructure
		.EndStructure	   = llEndStructure
		.ClosingExtraText  = llClosingExtraText
	Endwith

	If Empty (lcName)
		lcName = 'Line ' + Transform (lnLineNumber)
	Endif

	Try
		loResult.Add (loObject, lcName)
	Catch To loException When loException.ErrorNo = 2062
		loResult.Add (loObject, lcName + ' ' + Transform (lnLineNumber))
	Catch To loException
		This.ShowErrorMsg (loException)
	Endtry

Endfor

Return loResult


Procedure GetStructurePositionObject
Local loObject As 'Empty'

loObject = Createobject ('Empty')
AddProperty (loObject, 'Type')
AddProperty (loObject, 'LineNumber')
AddProperty (loObject, 'Name')
AddProperty (loObject, 'ClassName')
AddProperty (loObject, 'Text')
AddProperty (loObject, 'ProcedureBoundary')
AddProperty (loObject, 'StartStructure')
AddProperty (loObject, 'EndStructure')
AddProperty (loObject, 'ClosingExtraText')

Return loObject
EndProc 


Procedure ShowErrorMsg
Lparameters loException

Messagebox ('Error: ' + Transform (loException.ErrorNo) 	+ ccCRLF +							;
	  'Message: ' + loException.Message 					+ ccCRLF +							;
	  'Procedure: ' + loException.Procedure + ccCRLF + ;
	  'Line: ' + Transform (loException.Lineno) 			+ ccCRLF +							;
	  'Code: ' + loException.LineContents 							;
	  , MB_OK + MB_ICONEXCLAMATION, 'Error')
