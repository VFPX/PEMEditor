#include MemberDataEditor.H

Define Class PEMEditor_Server As Session

	DataSessionId 	= 2
	nKeywordLen 	= 0
	lIsNotCurrent	= .F.
	loDom			= .Null.

	cCopyClass		= ''
	cCopyClassLibrary = ''
	cCopyClassBaseClass = ''

	Dimension 		aCopiedProperties(1)
	Dimension 		aCompareClassProperties(1)
	Dimension 		aCompareProperties(1)

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

		Set Ansi On
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
			PK		 	 I       	    ,;
			cName	     C(128)			,;
			cType        C(1)			,;
			lAccess      L				,;
			lAssign      L				,;
			nVisibility  N(2)			,;
			lNative      L				,;
			lInherited   L				,;
			nFavorites   N(2)		    ,;
			lReadonly    L				,;
			cScript      M				 ;
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

	Function GetMethodList (toObject, tlClassDesigner, tlTopOfClass)
		Local loResult As 'Empty'
		Local lcTypes, lnPK

		loResult = Createobject('Empty')
		AddProperty(loResult, 'MethodList(1)')

		lnPK = This.FetchPEMList (toObject, tlClassDesigner, tlTopOfClass)
		lcTypes = Evl (lcTypes, 'PEM')
		Select cName, cType 					;
			From csrMembersList 				;
			where PK = lnPK 					;
			and cType $ 'ME'					;
			order By cName						;
			Into Array loResult.MethodList

		Return loResult
	Endfunc


	Function GetPEMList (toObject, tlClassDesigner, tlTopOfClass)
		Local loResult As 'Empty'
		Local lcTypes, lnPK

		loResult = Createobject('Empty')
		AddProperty(loResult, 'MethodList(1)')

		lnPK = This.FetchPEMList (toObject, tlClassDesigner, tlTopOfClass)
		lcTypes = Evl (lcTypes, 'PEM')
		Select cName, cType, lAccess, lAssign, nVisibility, lNative, lInherited, IIf(nFavorites = 1, 2, 1), lReadonly, cScript ;
			From csrMembersList 				;
			where PK = lnPK 					;
			and cType $ 'PEM'					;
			order By cName						;
			Into Array loResult.MethodList

		Return loResult
	Endfunc


	Function FetchPEMList (toObject, tlClassDesigner, tlTopOfClass)
		lcClass   	= Lower (IIf(tlTopOfClass and tlClassDesigner, toObject.ParentClass, toObject.Class))
		lcLibrary 	= Lower (toObject.ClassLibrary)
		lcOLEName	= Lower (IIf(pemstatus(toObject, 'BaseClass', 5) and Upper(toObject.BaseClass) = 'OLECONTROL', toObject.OleClass, ''))
		Select PK From csrClassList 				;
			where cClassName = Lower(toObject.Class);
			and cOLEClassName = lcOLEName			;
			and cClassLoc = lcLibrary				;
			into Array laPK

		If _Tally > 0
			lnPK = laPK[1]
		Else
			lnPK = This.CreateMethodList( toObject, tlClassDesigner, tlTopOfClass)
		Endif

		Return lnPK
	Endfunc


	Function CreateMethodList (toObject, tlClassDesigner, tlTopOfClass)
		Local laActiveXMembers(1), laObjectMembers(1), lcClass, lcLibrary, lcLibraryNoPath, lcMemberData
		Local lcName, lcNewName, lcPrompt, lcType, llAccess, llActiveX, llAssign, llFixed, llInherited
		Local llNative, llReadOnly, lnLen, lnPK, lnPos, lnRow, lnVisibility

		Create Cursor cMethodList (	 	 ;
			PK		 	 I       	    ,;
			cName	     C(128)			,;
			cType        C(1)			,;
			lAccess      L				,;
			lAssign      L				,;
			nVisibility  N(2)			,;
			lNative      L				,;
			lInherited   L				,;
			nFavorites   N(2)		    ,;
			lReadonly    L				,;
			cScript      M				,;
			lFixed		 L				 ;
			)

		Amembers(laObjectMembers, toObject, 1, 'PHG#')
		llActiveX = pemstatus(toObject, 'BaseClass', 5) and Lower(toObject.BaseClass) = 'olecontrol'
		If llActiveX
			Amembers(laActiveXMembers, toObject, 3)
		Endif llActiveX

		For lnRow = 1 To Alen(laObjectMembers, 1)
			If Inlist(laObjectMembers(lnRow,2), 'Method', 'Event', 'Property')

				lcName = laObjectMembers(lnRow,1)
				****************************************************************
				llNative = 'N' $ laObjectMembers[lnRow, 3]
				llInherited = 'I' $ laObjectMembers[lnRow, 3]

				llAssign = 0 # Ascan(laObjectMembers, lcName + [_Assign], -1, -1, 1, 15)
				llAccess = 0 # Ascan(laObjectMembers, lcName + [_Access], -1, -1, 1, 15)

				lnVisibility = IIf( [G] $laObjectMembers[lnRow, 3], 1, IIf([P] $laObjectMembers[lnRow, 3], 2, 3))
				llReadOnly = 'R' $ laObjectMembers[lnRow, 3]

				Do Case
					Case lnVisibility = 1 && Public
					Case Not (llNative Or llInherited) && Custom
					Case tlTopOfClass && class -- thisform
					Case lnVisibility = 2 && Protected
						llReadOnly = .T.
					Case lnVisibility = 3 && Hidden
						Loop
				Endcase
				****************************************************************

				If llNative And Inlist( Lower( Alltrim( lcName ) ), [activecontrol], [activeform])
					Loop
				Endif

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

				****************************************************************

				Insert Into cMethodList	;
					(cName,	cType, lAccess, lAssign, nVisibility, lNative, lInherited, nFavorites, lReadonly, cScript, lFixed) ;
					Values ;
					(lcName, lcType, llAccess, llAssign, lnVisibility, llNative, llInherited, -1, llReadOnly, cScript, llFixed)
			Endif
		Next lnRow

		****************************************************************
		* Post local MemberData, if any
		*** JRN 2010-09-21 : NO!  Never called from TopOfForm, so _MemberData here ignored!
		If PemStatus(toObject, '_MemberData', 5) and PemStatus(toObject, '_MemberData', 0)
			Try
				lcMemberData = Getpem(toObject, '_MemberData')
			Catch
				lcMemberData = ''
			Endtry
			If Not Isnull(lcMemberData)
				This.PostMemberData(lcMemberData)
			Endif
		Endif

		****************************************************************
		*	Get names up the parent classes ...
		*	Most of this lifted from MemberDataEngine and then modified

		lcClass   = Lower( IIf(tlTopOfClass and tlClassDesigner, toObject.ParentClass, toObject.Class))
		lcLibrary = toObject.ClassLibrary

		Insert Into csrClassList 					;
			(cClassName, cClassLoc,	cOLEClassName)	;
			Values									;
			(Lower(toObject.Class), Lower(lcLibrary), Lower(IIf(pemstatus(toObject, 'BaseClass', 5) and Upper(toObject.BaseClass) = 'OLECONTROL', toObject.OleClass, '')))
		lnPK = csrClassList.PK

		Do While Not Empty(lcLibrary)

			* can not process classes based on PRGs
			If Inlist (Upper(Justext(lcLibrary)), 'PRG', 'FXP')
				Exit
			Endif

			If Not File (lcLibrary)
				* look for file along path
				lcLibraryNoPath = Justfname(lcLibrary)
				If File	(lcLibraryNoPath)
					lcLibrary = lcLibraryNoPath
				Else
					lcPrompt = Strtran (ccLOC_File_Not_Found, '<FileName>', lcLibraryNoPath)
					*	Messagebox (lcPrompt, 0, ccLOC_MEM_CAP_TITLE)
					Wait (lcPrompt) Window Nowait
					lcLibrary = ''
					Exit
				Endif
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

		Insert Into csrMembersList ;
			(PK,         cName,	cType, lAccess, lAssign, nVisibility, lNative, lInherited, nFavorites, lReadonly, cScript) ;
			Select lnPK, cName,	cType, lAccess, lAssign, nVisibility, lNative, lInherited, nFavorites, lReadonly, cScript ;
			From cMethodList

		Return lnPK

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
						Locate For Upper(cName) = lcPEM + ' '
						If Found() And Not lFixed
							Replace cName With lcDisplay, lFixed With .T.
						Endif
					Endif

					lcScript = loNode.getAttribute('script')
					If Not Isnull(lcScript)
						Locate For Upper(cName) = lcPEM + ' '
						If Found() And 0 = Len (cScript)
							Replace cScript With lcScript + ' '
						Endif
					Endif

					lcFavorites = loNode.getAttribute('favorites')
					If Not Isnull(lcFavorites)
						Locate For Upper(cName) = lcPEM + ' '
						If Found() And nFavorites < 0
							Replace nFavorites With IIf(lcFavorites == ccXML_TRUE, 1, 0)
						Endif
					Endif

				Endfor
			Endif
		Endif Not Empty(lcMemberData)
	Endproc


	Function SpellNative (tcPem, tcName)
		Return IIf(Used('VFP_Keywords') And ;
			seek(Upper(Padr(tcPem, This.nKeywordLen)), 'VFP_Keywords'), ;
			trim(VFP_Keywords.TOKEN), tcName)
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


Enddefine

