*==============================================================================
* Procedure:		CreateLocalsStatements
* Purpose:			Returns a character string of LOCALs statements, formatted as desired
* Author:			Jim Nelson
* Parameters:
*	laLocals  	 =	An array of items to be created as locals, one per Line, in order of occurrence:
*          Col 1 = Name
*          Col 2 = AS phrase (from calls To Createobject And Newobject)
*          Col 3 = Array?
*          Col 4 = Has AS phrase?
*          Col 5 = Begins with lowercase 'l'
*   	   Col 6 = Already found in LOCALs
*          Col 7 = Array parameters
*          Col 8 = Actual assignment?
*          Col 9 = <Unused>
*         Col 10 = <Unused>

*   Note:  an item is an orphan if Col 6 is True and Col 8 is False

* Returns:			The header text To be inserted into the method
* Added:			06/14/2010
*==============================================================================
* Notes:
*	This is a sample for creating LOCAls statements.  It can be modified as needed.
*
*	For PEM Editor to use this PRG, it should be moved into the folder "Live Plug-Ins"
*

Lparameters laLocals

#Define LINEEND Chr(13) + Chr(10)
#Define Tab		Chr(9)

Local lcASPhrase, lcArrayDims, lcLine, lcLocalName, lcLocalPhrase, lcName, lcResult, lcVar
Local llAlreadyLocal, llArray, llCommasBeforeLocals, llContinuation, llHasASPhrase, llStartsWithL
Local llUseAsPhrase, llUseTabs, lnMaxWidth, lnRow, lnSelectionType, lnType

****************************************************************
*  Current preferences ....

lcLocalName	= 'Local '
llUseTabs	= .T.

With _oPEMEditor.oPrefs
	llContinuation		 = Not .lLocalsMultPerLine
	lnSelectionType		 = .nLocalsSelectionType
	lnMaxWidth			 = .nLocalsLineWidth
	llUseAsPhrase		 = .lLocalsUseAS
	llCommasBeforeLocals = .lCommasBeforeLocals
Endwith

****************************************************************


Asort (laLocals, 1, -1, 0, 0)

lcResult = ''
For lnType = 1 To 4
	lcLine = ''
	For lnRow = 1 To Alen (laLocals, 1)

		lcName		   = laLocals (lnRow, 1)
		lcASPhrase	   = laLocals (lnRow, 2)
		llArray		   = laLocals (lnRow, 3)
		llHasASPhrase  = laLocals (lnRow, 4)
		llStartsWithL  = laLocals (lnRow, 5)
		llAlreadyLocal = laLocals (lnRow, 6)
		lcArrayDims	   = laLocals (lnRow, 7)

		lcName = lcName + IIf (llArray, '[' + lcArrayDims + ']', '')

		Do Case

				****************************************************************
			Case llUseAsPhrase And llStartsWithL
				If lnType = 1
					lcVar		  = AddASPhrase (lcName, lcASPhrase)
					lcLocalPhrase = lcLocalName
				Else
					Loop
				Endif

			Case llUseAsPhrase And lnType = 1 And lnSelectionType = 4
				lcVar		  = AddASPhrase (lcName, lcASPhrase)
				lcLocalPhrase = lcLocalName

			Case llUseAsPhrase And lnType = 2 And lnSelectionType = 3
				lcVar		  = AddASPhrase (lcName, lcASPhrase)
				lcLocalPhrase = lcLocalName

			Case llUseAsPhrase And lnType = 3 And lnSelectionType = 2
				lcVar		  = AddASPhrase (lcName, lcASPhrase)
				lcLocalPhrase = '*:Global '

			Case llUseAsPhrase
				Loop

				****************************************************************
			Case lnSelectionType = 1 And Not llStartsWithL
				Loop

			Case llHasASPhrase And lnType = 1
				lcVar		  = lcName + ' as ' + lcASPhrase
				lcLocalPhrase = lcLocalName
			Case llHasASPhrase
				Loop

			Case llStartsWithL And lnType = 2
				lcVar		  = lcName
				lcLocalPhrase = lcLocalName
			Case llStartsWithL
				Loop

			Case lnSelectionType = 4 And lnType = 2
				lcVar		  = lcName
				lcLocalPhrase = lcLocalName

			Case lnSelectionType = 2 And lnType = 3 And llAlreadyLocal
				lcVar		  = lcName
				lcLocalPhrase = lcLocalName

			Case lnSelectionType = 2 And lnType = 4 And Not llAlreadyLocal
				lcVar		  = lcName
				lcLocalPhrase = '*:Global '
			Case lnSelectionType = 3 And lnType = 4
				lcVar		  = lcName
				lcLocalPhrase = lcLocalName

			Otherwise
				Loop
		Endcase

		Do Case
			Case Empty (lcLine)
				lcLine = lcLocalPhrase  + lcVar
			Case llContinuation Or llUseAsPhrase
				If llCommasBeforeLocals
					lcLine = lcLine + ';' + LINEEND + IIf (llUseTabs, Tab, '') + ', ' + lcVar
				Else
					lcLine = lcLine + ', ;' + LINEEND + IIf (llUseTabs, Tab, '') + lcVar
				Endif
			Case lnType # 1 And (Len (lcLine) + 2 + Len (lcVar)) <= lnMaxWidth
				lcLine = lcLine + ', ' + lcVar
			Otherwise
				lcResult = lcResult + lcLine + LINEEND
				lcLine	 = lcLocalPhrase  + lcVar
		Endcase
	Endfor

	If Not Empty (lcLine)
		lcResult = lcResult + lcLine + LINEEND
	Endif
Endfor


Return Left (lcResult, Len (lcResult) - 2)


Function AddASPhrase
	Lparameters lcName, lcASPhrase
	Local lcSuffix, lcVartype

	lcVartype = Upper (Substr (lcName, 2, 1))

	Do Case
		Case Not Empty (lcASPhrase)
			Return lcName + ' as ' + lcASPhrase
		Case lcName # 'l'
			Return lcName
		Case lcVartype = 'N'
			lcSuffix = ' as Number'
		Case lcVartype = 'I'
			lcSuffix = ' as Integer'
		Case lcVartype = 'O'
			lcSuffix = ' as Object'
		Case lcVartype = 'L'
			lcSuffix = ' as Boolean'
		Case lcVartype = 'C'
			lcSuffix = ' as String'
		Case lcVartype = 'U'
			lcSuffix = ' as Variant'
		Case lcVartype = 'D'
			lcSuffix = ' as Date'
		Case lcVartype = 'T'
			lcSuffix = ' as DateTime'
		Case lcVartype = 'Y'
			lcSuffix = ' as Currency'
		Otherwise
			lcSuffix = ''
	Endcase

	Return lcName + lcSuffix

Endfunc
