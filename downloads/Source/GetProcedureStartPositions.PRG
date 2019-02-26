Lparameters tcCode, tcName

Local loObject As 'Empty'
Local loRegExp As 'VBScript.RegExp'
Local loResult As 'Collection'
Local lcClassName, lcMatch, lcName, lcPattern, lcType, lcWord1, llClassDef, llTextEndText, lnI
Local lnStartByte, loException, loMatches

* Returns a collection indicating the beginning of each procedure / class / etc
* Each member in the collection has these properties:
*   .Type == 'Procedure' (Procedures and Functions)
*         == 'Class'     (Class Definition)
*         == 'End Class' (End of Class Definition)
*         == 'Method'    (Procedures and Functions within a class)
*   .StartByte == starts at zero; thus, # of chars preceding start position
*   .Name
*   .ClassName

****************************************************************
loRegExp = Createobject ('VBScript.RegExp')
With loRegExp
	.IgnoreCase	= .T.
	.Global		= .T.
	.MultiLine	= .T.
Endwith

lcPattern = 'PROC(|E|ED|EDU|EDUR|EDURE)\s+(\w|\.)+'
lcPattern = lcPattern + '|' + 'FUNC(|T|TI|TIO|TION)\s+(\w|\.)+'
lcPattern = lcPattern + '|' + 'DEFINE\s+CLASS\s+\w+'
lcPattern = lcPattern + '|' + 'DEFI\s+CLAS\s+\w+'
lcPattern = lcPattern + '|' + 'ENDD(E|EF|EFI|EFIN|EFINE)'
lcPattern = lcPattern + '|' + 'PROT(|E|EC|ECT|ECTE|ECTED)\s+\w+\s+\w+'
lcPattern = lcPattern + '|' + 'HIDD(|E|EN)\s+\w+\s+\w+'

With loRegExp
	.Pattern	= '^\s*(' + lcPattern + ')'
Endwith

loMatches = loRegExp.Execute (tcCode)
****************************************************************

loResult = Createobject ('Collection')

llClassDef	  = .F. && currently within a class?
llTextEndText = .F. && currently within a Text/EndText block?
lcClassName	  = ''

For lnI = 1 To loMatches.Count

	* .Value
	* .FirstIndex
	* .Length
	With loMatches.Item (lnI - 1)
		lnStartByte	= .FirstIndex
		lcMatch		= Chrtran (.Value, CR + LF, '  ')
		lcName		= Getwordnum (lcMatch, Getwordcount (lcMatch))
		lcWord1		= Upper (Getwordnum (lcMatch, Max(1, Getwordcount (lcMatch) - 1)))
	Endwith

	Do Case
		Case llTextEndText
			If 'ENDTEXT' = lcWord1
				llTextEndText = .F.
			Endif
			Loop

		Case llClassDef
			If 'ENDDEFINE' = lcWord1
				llClassDef	= .F.
				lcType		= 'End Class'
				lcName		= lcClassName + '.-EndDefine'
				lcClassName	= ''
			Else
				lcType = 'Method'
				lcName = lcClassName + '.' + lcName
			Endif

		Case 'CLASS' = lcWord1
			llClassDef	= .T.
			lcType		= 'Class'
			lcClassName	= lcName

		Otherwise
			lcType = 'Procedure'

	Endcase

	* ignore leading CRLF's
	Do While Substr (tcCode, lnStartByte + 1, 1) $ CR + LF
		lnStartByte = lnStartByte + 1
	Enddo

	loObject = Createobject ('Empty')
	AddProperty (loObject, 'Type')
	AddProperty (loObject, 'StartByte')
	AddProperty (loObject, 'Name')
	AddProperty (loObject, 'ClassName')

	With loObject
		.StartByte = lnStartByte
		.Type	   = lcType
		.Name	   = lcName
		.ClassName = lcClassName
	Endwith

	Try
		loResult.Add (loObject, lcName)
	Catch To loException When loException.ErrorNo = 2062
		loResult.Add (loObject, lcName + ' ' + Transform (lnStartByte))
	Catch To loException
		This.ShowErrorMsg(loException)
	Endtry


Endfor

Return loResult
