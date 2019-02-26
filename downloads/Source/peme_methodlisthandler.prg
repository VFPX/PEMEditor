
#include MemberDataEditor.H

Define Class PEMEditor_Server As Session

	DataSessionId 	= 2
	nKeywordLen 	= 0
	lIsNotCurrent	= .F.
	loDom			= .Null.

	Function Init

		Try
			Use Home() + 'WIZARDS\FDKEYWRD' Order TOKEN Again Shared In 0 Alias Keywords

			Select TOKEN From Keywords Into Cursor VFP_Keywords Readwrite
			Index On Upper(TOKEN) Tag TOKEN
			Use In Keywords

			This.nKeywordLen = Len(VFP_Keywords.TOKEN)
		Catch

		Endtry

		This.CreateCursors

		With This
			.loDom = Createobject(ccXML_DOM_CLASS)
			.loDom.Async = .F.
		Endwith

		Set Deleted On

	Endfunc


	Function Reset
		This.CreateCursors()
	Endfunc


	Function CreateCursors
		Create Cursor csrClassList (	 	;
			PK			 	I       Autoinc ,;
			cClassName 		C(100)			,;
			cOLEClassName 	C(100)			,;
			cClassLoc   	C(240)			 ;
			)
		Index On cClassName Tag Order

		Create Cursor csrClassParents (	 ;
			PK		 	I       	    ,;
			cClassLoc   C(240)			,;
			dFileTS     T				 ;
			)
		Index On PK Tag Order

		Create Cursor csrMembersList (	 ;
			PK		 	I       	    ,;
			cName	    C(60)			,;
			cType       C(1)			 ;
			)
		Index On PK Tag Order
	Endfunc


	Function Destroy
		Close Tables
	Endfunc


	Function MarkNotCurrent
		This.lIsNotCurrent	= .T.
	Endfunc

	Function MakeCurrent
		Local lAny, lChanged

		If This.lIsNotCurrent
			Select Distinct cClassLoc, dFileTS, .F. As lChanged	;
				from csrClassParents				;
				into Cursor cClassLocs Readwrite

			Scan
				If dFileTS # This.GetFileTimeStamp (cClassLoc)
					Replace lChanged With .T.
					lAny = .T.
				Endif
			Endscan

			If lAny
				Select PK From csrClassParents					;
					where cClassLoc In (Select cClassLoc From cClassLocs Where lChanged)	;
					into Cursor cPKs
				Index On PK Tag PG

				Select csrClassList
				Set Relation To PK Into cPKs
				Delete For Found("cPKs")
				Set Relation To

				Select * From csrClassParents 					;
					where PK In (Select PK From csrClassList)	;
					into Cursor csrClassParents Readwrite
				Index On PK Tag Order

				Select * From csrMembersList					;
					where PK In (Select PK From csrClassList)	;
					into Cursor csrMembersList Readwrite
				Index On PK Tag Order

			Endif
			This.lIsNotCurrent = .F.
		Endif
	Endfunc

	Function GetMethodList (toObject, tlClassDesigner)

		* 	Parameters:
		*		toObject	- object for which we want methods list
		*		tlClassDesigner	- working on class?
		*
		*	Result: an object containing array Members:
		*		Members - column 1 - Member Names
		*			    - column 2 - Type -- 'P', 'E', or

		****************************************************************

		Local laPK[1]
		Local lcClass, lcLibrary, loResult, lcOLEName

		loResult = Createobject('Empty')
		AddProperty(loResult, 'MethodList(1)')

		lcClass   	= Lower (Iif(tlClassDesigner, toObject.ParentClass, toObject.Class))
		lcLibrary 	= Lower (toObject.ClassLibrary)
		lcOLEName	= Lower (Iif(Upper(toObject.BaseClass) = 'OLECONTROL', toObject.OleClass, ''))
		Select PK From csrClassList 				;
			where cClassName = Lower(toObject.Class);
			and cOLEClassName = lcOLEName			;
			and cClassLoc = lcLibrary				;
			into Array laPK
		If _Tally > 0
			Select cName, cType From csrMembersList ;
				where PK = laPK(1)					;
				order By cName						;
				Into Array loResult.MethodList
		Else
			This.CreateMethodList( toObject, tlClassDesigner)
			Select cName, cType From cMethodList 	;
				order By cName						;
				Into Array loResult.MethodList
		Endif

		Return loResult
	Endfunc


	Function GetFileTimeStamp (lcLibrary)
		Local lcVCT
		lcVCT = Forceext(lcLibrary, 'VCT')
		If File(lcVCT)
			Return Max( Fdate (lcLibrary, 1), Fdate (lcVCT,1))
		Else
			Return Fdate (lcLibrary, 1)
		Endif
	Endfunc


	Function CreateMethodList (toObject, tlClassDesigner)

		* 	Parameters:
		*		toObject	- object for which we want methods list
		*		tlClassDesigner	- working on class?
		*
		*	Result: an object containing array Members:
		*		Members - column 1 - Member Names
		*			    - column 2 - Type -- 'P', 'E', or

		Local laObjectMembers(1), loResult
		Local lcName, llFixed, lcType, lnRow, lnPK
		Local Array laActiveXMembers[1]
		Local llActiveX, lnPos, lcNewName, lcMemberData, lcClass, lcLibrary, lcPrompt, lnLen, ltDateTime

		Create Cursor cMethodList (cName C(60), cType C(1), lFixed L)

		Amembers(laObjectMembers, toObject, 1)
		llActiveX = Lower(toObject.BaseClass) = 'olecontrol'
		If llActiveX
			Amembers(laActiveXMembers, toObject, 3)
		Endif llActiveX

		For lnRow = 1 To Alen(laObjectMembers, 1)
			If Inlist(laObjectMembers(lnRow,2), 'Method', 'Event')
				lcName = laObjectMembers(lnRow,1)
				lcType = Left (laObjectMembers(lnRow,2), 1)
				llFixed = .F.

				If llActiveX
					lnPos = Ascan(laActiveXMembers, lcName, -1, -1, 1, 15)
					If lnPos > 0
						lcName = laActiveXMembers[lnPos, 1]
						llFixed = .T.
					Endif lnPos > 0
				Endif llActiveX

				If Not llFixed
					lcNewName = This.SpellNative(lcName, lcName)
					If lcNewName # lcName
						lcName = lcNewName
						llFixed = .T.
					Else
						lcName = Lower(lcName)
					Endif
				Endif
				Insert Into cMethodList Values (lcName, lcType, llFixed)
			Endif
		Next lnRow

		****************************************************************
		* Post local MemberData, if any
		Try
			lcMemberData = Getpem(toObject, '_MemberData')
		Catch
			lcMemberData = ''
		Endtry
		If Not Isnull(lcMemberData)
			This.PostMemberData(lcMemberData)
		Endif

		****************************************************************
		*	Get names up the parent classes ...
		*	Most of this lifted from MemberDataEngine and then modified

		lcClass   = Lower( Iif(tlClassDesigner, toObject.ParentClass, toObject.Class))
		lcLibrary = toObject.ClassLibrary

		Insert Into csrClassList 					;
			(cClassName, cClassLoc,	cOLEClassName)	;
			Values									;
			(Lower(toObject.Class), Lower(lcLibrary), Lower(Iif(Upper(toObject.BaseClass) = 'OLECONTROL', toObject.OleClass, '')))
		lnPK = csrClassList.PK

		Do While Not Empty(lcLibrary)
			If Not File (lcLibrary)
				lcPrompt = Strtran (ccLOC_File_Not_Found, '<FileName>', Justfname(lcLibrary))
				Messagebox (lcPrompt, 0, ccLOC_MEM_CAP_TITLE)
				lcLibrary = ''
				Exit
			Endif

			Insert Into csrClassParents Values (lnPK, Lower(lcLibrary), This.GetFileTimeStamp (lcLibrary))

			Select 0
			Use (lcLibrary) Again Shared Alias VCX
			Locate For OBJNAME == lcClass And Lower(RESERVED1) = 'class'

			lnPos = At('_memberdata = ', PROPERTIES)
			If lnPos > 0
				lnPos = lnPos + 14

				* We have to handle properties with more than 255 characters in the value
				* differently.
				If Substr(PROPERTIES, lnPos, 1) = ccPROPERTIES_PADDING_CHAR
					lnLen        = Val(Alltrim(Substr(PROPERTIES, 			;
						lnPos + cnPROPERTIES_PADDING_SIZE, 					;
						cnPROPERTIES_LEN_SIZE)))
					lcMemberData = Substr(PROPERTIES, lnPos + 				;
						cnPROPERTIES_PADDING_SIZE + cnPROPERTIES_LEN_SIZE, 	;
						lnLen)
				Else
					lcMemberData = Strextract(Substr(PROPERTIES, lnPos), 	;
						'', ccCR)
				Endif Substr(PROPERTIES, lnPos, 1) = ccPROPERTIES_PADDING_CHAR

				This.PostMemberData(lcMemberData)

			Endif lnPos > 0

			Select VCX
			lcClass   = Lower(Class)
			If Not Empty(CLASSLOC)
				lcLibrary = Fullpath(CLASSLOC, Addbs(Justpath(lcLibrary)))
				If Not File( m.lcLibrary )
					lcLibrary = Fullpath( CLASSLOC )
				Endif
			Else
				lcLibrary = ''
			Endif Not Empty(CLASSLOC)

			Use In VCX
		Enddo While Not Empty(lcLibrary)

		Insert Into csrMembersList Select lnPK, cName, cType From cMethodList

	Endfunc


	Procedure PostMemberData(lcMemberData)
		Local loNodes, lcPEM, lcDisplay, loNode
		If Not Empty(lcMemberData)
			Select cMethodList
			This.loDom.LoadXML(lcMemberData)
			If This.loDom.parseError.errorCode = 0
				loNodes = This.loDom.selectNodes('//' + ccMEMBER_DATA_XML_ELEMENT)
				For Each loNode In loNodes
					lcPEM     = Upper(loNode.getAttribute('name'))
					lcDisplay = loNode.getAttribute('display')
					If Not Isnull(lcDisplay)
						Locate For Upper(cName) = lcPEM + ' ' And Not lFixed
						If Found()
							Replace cName With lcDisplay, lFixed With .T.
						Endif
					Endif
				Endfor
			Endif
		Endif Not Empty(lcMemberData)
	Endproc


	Function SpellNative (tcPem, tcName)
		Return Iif(Used('VFP_Keywords') And ;
			seek(Upper(Padr(tcPem, This.nKeywordLen)), 'VFP_Keywords'), ;
			trim(VFP_Keywords.TOKEN), tcName)
	Endfunc

Enddefine

