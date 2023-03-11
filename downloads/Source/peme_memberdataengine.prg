*==============================================================================
* Program:			MemberDataEngine.PRG
* Purpose:			MemberData engine classes
* Last revision:	11/12/2004
*==============================================================================

#include MemberDataEditor.H

* The MemberData Editor engine class.
***********************************************************************
*** Changed By.: Marcia G. Akins on 05 January 2007
*** Reason.....: Need the cursor in the edit property/method class, so base on custom instead of session
***********************************************************************
Define Class MemberDataEngine As Custom
	oObject = .Null.
	* Reference to the object we're working on MemberData for
	cMemberData = ''
	* Existing MemberData for the object
	Dimension aObjectMembers[1]
	* An array of members of the object
	lHadMemberData = .T.
	* .T. if the object already had MemberData
	oXMLDOM = .Null.
	* A reference to an XML DOM object
	cSelectedMember = ''
	* The name of the member passed to the editor by Builder.APP
	lCreateGetMemberDataRecord = .F.
	* .T. to create an _GetMemberData record in FOXCODE on a "setup editor"
	* call
	cErrorMessage = ''
	* The text of any error that occurred
	oMemberDataCollection = .Null.
	* A collection of objects containing information about the MemberData
	* for the object
	lPropertiesWindowOpen = .F.
	* .T. if the Properties window was open when we started
	Dimension aDockWindow[1]
	* An array of window dock states
	lClassDesigner = .F.
	* .T. if the object is a class
	lNoObjectMemberData = .F.
	* .T. if we can't create MemberData for the object
	nKeywordLen		= 0
	* Length of Native Keywords

	* Clean up upon exit.

	Function Destroy
		This.oObject               = .Null.
		This.oMemberDataCollection = .Null.
		This.oXMLDOM               = .Null.
	Endfunc

	Function Init
		* Open FDKEYWRD because it contains the correct case for all VFP PEMs. If we
		* can't, we won't worry about it.
		If Used('VFP_Keywords')
			This.nKeywordLen = Len(VFP_Keywords.Token)
		Endif
	Endfunc


	Function SpellNative (tcPem, tcName)
		Return IIf(Used('VFP_Keywords') And ;
			seek(Upper(Padr(tcPem, This.nKeywordLen)), 'VFP_Keywords'), ;
			trim(VFP_Keywords.Token), tcName)
	Endfunc

	* Set up properties based on the object we're working with.
	Function SetupEngine(tlAddToFavorites)
		Local laObjects[1], ;
			lnObjects, ;
			loObject

		With This

			* If we don't already have it, get a reference to the selected object.

			If Vartype(.oObject) <> 'O'
				lnObjects = Aselobj(laObjects)
				If lnObjects = 0
					lnObjects = Aselobj(laObjects, 1)
				Endif lnObjects = 0
				If lnObjects > 0
					.oObject = laObjects[1]
				Endif lnObjects > 0
			Endif Vartype(.oObject) <> 'O'
			If Vartype(.oObject) = 'O'

				* See if the selected object has a _MemberData property. If not and it's a
				* member object, tell the user we'll handle the class/form and then find it.

				.lHadMemberData = Pemstatus(.oObject, '_MemberData', 5)
				*!*	loObject = PEME_FindTopmostParent(.oObject, .lClassDesigner)
				*!*	If Not .lHadMemberData And Not Compobj(.oObject, loObject)
				*!*		***********************************************************************
				*!*		*** Changed By.: Marcia G. Akins on 05 January 2007
				*!*		*** Reason.....: do not want this messagebox if we are being called from edit property/method
				*!*		***********************************************************************
				*!*		*!*  					if tlAddToFavorites
				*!*		*!*  						messagebox(ccLOC_CANT_ADD_FAVORITES, ;
				*!*		*!*  							MB_OK + MB_ICONINFORMATION, ;
				*!*		*!*  							ccLOC_CAP_MEMBER_DATA_EDITOR)
				*!*		*!*  					else
				*!*		*!*  						messagebox(ccLOC_USE_PARENT, ;
				*!*		*!*  							MB_OK + MB_ICONINFORMATION, ;
				*!*		*!*  							ccLOC_CAP_MEMBER_DATA_EDITOR)
				*!*		*!*  					endif tlAddToFavorites
				*!*		.lNoObjectMemberData = .T.
				*!*	Endif Not .lHadMemberData ...

				* Get the object's existing MemberData (if any) and an array of its members.

				If .lHadMemberData
					.cMemberData = .oObject._MemberData
					If Vartype(.cMemberData) <> 'C'
						.cMemberData = ''
					Endif Vartype(.cMemberData) <> 'C'
				Else
					.cMemberData = ''
				Endif .lHadMemberData
				Amembers(.aObjectMembers, .oObject, 1, 'PHG#')

				* Create an XML DOM object and load the MemberData XML. If we have a problem,
				* bug out.

				.oXMLDOM = Createobject(ccXML_DOM_CLASS)
				.oXMLDOM.Async = .F.
				If Not Empty(.cMemberData)
					.oXMLDOM.LoadXML(.cMemberData)
					If .oXMLDOM.parseError.errorCode <> 0
						.cErrorMessage = ccLOC_INVALID_XML
					Endif .oXMLDOM.parseError.errorCode <> 0
				Endif Not Empty(.cMemberData)
			Endif Vartype(.oObject) = 'O'
		Endwith
	Endfunc

	* Set the MemberData for the object so the specified member appears on the
	* Favorites tab of the Properties window.

	Function AddMemberToFavorites
		Local loNode, ;
			lnMember, ;
			lcType, ;
			loRoot, ;
			lcMemberData
		With This

			* Get the node for the member (if it exists).

			loNode = .oXMLDOM.selectSingleNode('//' + ;
				ccMEMBER_DATA_XML_ELEMENT + '[@name = "' + .cSelectedMember + ;
				'"]')

			* If it doesn't exist, we'll have to create it. First, determine what type of
			* member we have (Property, Event, or Method).

			If Vartype(loNode) <> 'O'
				lnMember = Ascan(.aObjectMembers, .cSelectedMember, -1, -1, ;
					1, 15)
				lcType   = .aObjectMembers[lnMember, 2]

				* Create a MemberData node for this member.

				loNode = .oXMLDOM.createElement(ccMEMBER_DATA_XML_ELEMENT)
				loNode.setAttribute('name', .cSelectedMember)
				loNode.setAttribute('type', lcType)

				* Get the root node of the MemberData XML. If it doesn't exist, create it.

				loRoot = .oXMLDOM.selectSingleNode('/' + ccXML_ROOT_NODE)
				If Vartype(loRoot) <> 'O'
					loRoot = .oXMLDOM.createElement(ccXML_ROOT_NODE)
					.oXMLDOM.appendChild(loRoot)
				Endif Vartype(loRoot) <> 'O'

				* Add the new node to the root.

				loRoot.appendChild(loNode)
			Endif Vartype(loNode) <> 'O'

			* Set the favorites attribute to True and update the _Memberdata property of
			* the object.

			loNode.setAttribute('favorites', ccXML_TRUE)
			.cMemberData = .oXMLDOM.XML
			.WriteMemberData()
		Endwith
	Endfunc

	* Update the _Memberdata property of the object (create it if necessary).

	Function WriteMemberData
		With This
			If .lHadMemberData
				.oObject._MemberData = .cMemberData
			Else
				.oObject.AddProperty('_MemberData', .cMemberData)
			Endif .lHadMemberData
		Endwith
	Endfunc

	* Fill oMemberDataCollection with a collection of the MemberData for the
	* members of the object.

	Function CreateMemberDataCollection (tcCursorName, tlTopOfClass, tlAvoidBufferOverrun, tlMethodsOnly, tlNoErrorMsgOnMissingVCX)
		Local lnKeywordLen, ;
			lnLen, ;
			lnMember, ;
			llMemberDataInherited, ;
			llMemberDataChanged, ;
			llActiveX, ;
			laMembers[1], ;
			lnI, ;
			lcPEM, ;
			lcType, ;
			llNative, ;
			lnPos, ;
			loPEMMemberData, ;
			loGlobalMemberData, ;
			llMemberData, ;
			llInherited, ;
			lcName, ;
			lcClass, ;
			lcLibrary, ;
			lnLen, ;
			lcMemberData, ;
			loPEM, ;
			loDOM, ;
			loNodes, ;
			loNode, ;
			lcDisplay, ;
			loParent, ;
			llFavorites, ;
			llGlobalFavorites, ;
			llAssign, ;
			llAccess, ;
			lcTypeAbbrev, ;
			lnVisibility, ;
			lcFile, ;
			lnSelect , ;
			lnLines, ;
			laLines[ 1 ], ;
			lnI, ;
			lcDescription, ;
			lcDescr, ;
			lnRow, ;
			loErr, ;
			llOK2Add, ;
			llInherited2, ;
			llReadOnly,		;
			llChanged, 		;
			llNonDefault,	;
			llLocalCode, 	;
			llHasCode,		;
			lcLibraryNoPath, ;
			lcPrompt,		;
			lcDesc

		Local lcPrevClass, lcPrevLibrary
		***********************************************************************
		*** Changed By.: Marcia G. Akins on 05 January 2007
		*** Reason.....: Since I am going to display this data in a grid, it is more useful for
		*** ...........: me to have it in a cursor instead of a collection
		***********************************************************************

		*** JRN 2009-10-02 : The call to AMembers below can crash VFP with 'Buffer Overrun'

		Dimension laLines[ 1 ]
		If Not tlAvoidBufferOverrun
			*** JRN 04/08/2009 : oops .. this fails on grids in Version 8
			Try
				***********************************************************************
				*** Changed By.: Marcia G. Akins on 06 January 2007
				*** Reason.....: Use AMEMBERS() to get the description from column 4
				***********************************************************************
				Amembers( laLines, This.oObject, 3 )
			Catch
			Endtry
		Endif

		* Create a cursor of MemberData records from the FOXCODE table so we can look
		* for global PEMs. Bug out if we can't open FOXCODE.

		Try
			Use (_Foxcode) Again Shared In 0 Alias Foxcode
		Catch
		Endtry
		If Not Used('FOXCODE')
			Create Cursor FoxCode (Type C(1), Abbrev C(24), Tip C(1))
		Endif Not Used('FOXCODE')
		Select ABBREV From FoxCode ;
			where Type = 'P' And Not Deleted() ;
			into Cursor _PROPERTIES
		Index On Upper(ABBREV) Tag ABBREV
		Select ABBREV, TIP From FoxCode ;
			where Type = ccGLOBAL_MEMBER_DATA_TYPE And Not Deleted() ;
			into Cursor Global
		Index On Upper(ABBREV) Tag ABBREV
		Use In Foxcode
		lnLen = Len(ABBREV)

		* Instantiate an MS XML DOM object.

		loDOM = Createobject(ccXML_DOM_CLASS)
		loDOM.Async = .F.

		* Put all the object's PEMs into a collection. For ActiveX controls, we'll use
		* AMEMBERS again to get the correct case for members.

		With This
			.oMemberDataCollection = Createobject('Collection')
			lnMember = Ascan(.aObjectMembers, '_MemberData', -1, -1, 1, 15)
			If lnMember > 0
				llMemberDataInherited = 'I' $ .aObjectMembers[lnMember, 3]
				llMemberDataChanged   = 'C' $ .aObjectMembers[lnMember, 3]
			Endif lnMember > 0
			llActiveX = pemstatus(.oObject, 'BaseClass', 5) and Lower(.oObject.BaseClass) = 'olecontrol'
			If llActiveX
				Amembers(laMembers, .oObject, 3)
			Endif llActiveX

			For lnI = 1 To Alen(.aObjectMembers, 1)
				lcPEM    = .aObjectMembers[lnI, 1]
				lcType   = Lower(.aObjectMembers[lnI, 2])

				If tlMethodsOnly And Not Inlist (lcType, 'method', 'event')
					Loop
				Endif

				llNative = 'N' $ .aObjectMembers[lnI, 3]
				llInherited2 = 'I' $ .aObjectMembers[lnI, 3]
				* For each PEM that doesn't hold an object, get the information about its
				* MemberData.

				If lcType <> 'object'
					llMemberData = .HasMemberData(lcPEM)

					* This member has inherited member data if the member data was changed and this
					* member doesn't appear in it or the member data wasn't changed and this member
					* does appear (because we've just inherited the member data).

					llInherited = ((llMemberDataChanged And Not llMemberData) Or ;
						(Not llMemberDataChanged And llMemberData)) And ;
						llMemberDataInherited And 'I' $ .aObjectMembers[lnI, 3]

					* If we can find the member in FOXCODE, we'll use the abbreviation stored there
					* since it may have the correct case.

					lcName = IIf(Seek(Padr('.' + lcPEM, lnLen), '_PROPERTIES'), ;
						trim(Substr(_PROPERTIES.ABBREV, 2)), lcPEM)

					lcName = This.SpellNative(lcPEM, lcName)

					* Get the correct case for members of ActiveX controls.

					If llActiveX
						lnPos = Ascan(laMembers, lcPEM, -1, -1, 1, 15)
						If lnPos > 0
							lcName = laMembers[lnPos, 1]
						Endif lnPos > 0
					Endif llActiveX

					* If the name is still upper-cased and it's a native or ActiveX control
					* property, use PROPER() on it.

					If Upper(lcName) == lcName
						If llNative && Or llActiveX
							lcName = Proper(lcName)
							***********************************************************************
							*** Changed By.: Marcia G. Akins on 05 January 2007
							*** Reason.....: I want custom properties to shake out at the bottom of the list
							***********************************************************************
						Else
							lcName = Lower( lcName )
						Endif
					Endif Upper(lcName) == lcName ...

					* Create an object to hold information about the MemberData.

					loPEMMemberData = Createobject('MemberDataObject')
					With loPEMMemberData
						.HasMemberData = llMemberData
						.Inherited     = llInherited
						If llMemberData
							.Display   = This.FindAttributeForMember(lcPEM, ;
								'Display')
							.Favorites = This.FindAttributeForMember(lcPEM, ;
								'Favorites')
							.Override  = This.FindAttributeForMember(lcPEM, ;
								'Override')
							.Script    = This.FindAttributeForMember(lcPEM, ;
								'Script')
							.OriginalDisplay   = .Display
							.OriginalFavorites = .Favorites
							.OriginalOverride  = .Override
							.OriginalScript    = .Script
							lcName             = Evl(.Display, lcName)
						Endif llMemberData
						.CustomAttributes = This.GetCustomAttributesForMember(lcPEM, ;
							This.oXMLDOM)
					Endwith

					***********************************************************************
					*** Changed By.: Marcia G. Akins on 05 January 2007
					*** Reason.....: Add a record to the cursor
					***********************************************************************
					lcDescription = []
					*   If Type( [laLines], 1 ) = [A]
					If Alen (laLines) >= 4
						lnRow = Ascan( laLines, lcName, 1, -1, 1, 15 )
						If lnRow > 0
							lcDescription = laLines[ lnRow, 4 ]
						Endif
					Endif

					* Now get any global MemberData for the PEM.

					loGlobalMemberData = Createobject('MemberDataObject')
					With loGlobalMemberData
						.HasMemberData = This.HasGlobalMemberData(lcPEM)
						If .HasMemberData
							loDOM.LoadXML(Global.TIP)
							If loDOM.parseError.errorCode = 0
								.Display           = This.FindAttributeForMember(lcPEM, ;
									'Display',   loDOM)
								.Favorites         = This.FindAttributeForMember(lcPEM, ;
									'Favorites', loDOM)
								.Override          = This.FindAttributeForMember(lcPEM, ;
									'Override',  loDOM)
								.Script            = This.FindAttributeForMember(lcPEM, ;
									'Script',    loDOM)
								.OriginalDisplay   = .Display
								.OriginalFavorites = .Favorites
								.OriginalOverride  = .Override
								.OriginalScript    = .Script
								lcName             = IIf(Upper(lcName) == lcName, ;
									evl(.Display, lcName), lcName)

								lcDesc				= This.FindAttributeForMember(lcPEM, ;
									'Description',    loDOM)
								If (Not Isnull(lcDesc)) And Not Empty(lcDesc)
									lcDescription = lcDesc
								Endif

								* Get any custom attributes.

								.CustomAttributes = This.GetCustomAttributesForMember(lcPEM, ;
									loDOM)
							Endif loDOM.parseError.errorCode = 0
						Endif .HasMemberData
					Endwith

					lcDisplay = Evl(loPEMMemberData.Display, loGlobalMemberData.Display)

					*** JRN 02/03/2009 : Can't use PEMStatus(5) on OLEs

					*** DougHennig 05/23/2007: fixed bug: the name of the property is contained in lcName, not "cName"
					***              llAssign = PEMSTATUS( This.oObject, [cName] + [_Assign], 5 )
					***              llAccess = PEMSTATUS( This.oObject, [cName] + [_Access], 5 )
					*** llAssign = Pemstatus( This.oObject, lcName + [_Assign], 5 )
					*** llAccess = Pemstatus( This.oObject, lcName + [_Access], 5 )
					llAssign = 0 # Ascan(.aObjectMembers, lcName + [_Assign], -1, -1, 1, 15)
					llAccess = 0 # Ascan(.aObjectMembers, lcName + [_Access], -1, -1, 1, 15)

					llFavorites = loPEMMemberData.Favorites == ccXML_TRUE
					llGlobalFavorites = loGlobalMemberData.Favorites == ccXML_TRUE
					lcTypeAbbrev = Upper( Left( This.aObjectMembers[lnI, 2], 1 ) )
					lnVisibility = IIf( [G] $This.aObjectMembers[lnI, 3], 1, IIf([P] $This.aObjectMembers[lnI, 3], 2, 3))
					llReadOnly = 'R' $ This.aObjectMembers[lnI, 3]
					Do Case
						Case lnVisibility = 1 && Public
						Case Not (llNative Or llInherited2) && Custom
						Case tlTopOfClass && class -- thisform
						Case lnVisibility = 2 && Protected
							llReadOnly = .T.
						Case lnVisibility = 3 && Hidden
							Loop
					Endcase

					*** JRN - 11/20/2008 -- capture which PEMs are non-default
					llChanged = 'C' $ This.aObjectMembers[lnI, 3]
					Do Case
						Case lcTypeAbbrev $ 'P'
							llNonDefault = llChanged Or Not (llNative Or llInherited2) && Custom
							llLocalCode = .F.
							llHasCode	= .F.
						Case llChanged
							llNonDefault = Not Empty (This.oObject.ReadMethod (lcName))
							llLocalCode = llNonDefault
							llHasCode	= llNonDefault Or 'N' $ This.aObjectMembers[lnI, 3] Or 'I' $ This.aObjectMembers[lnI, 3]
						Otherwise
							llNonDefault = .F.
							llLocalCode = .F.
							llHasCode	= .F.
					Endcase

					***********************************************************************
					*** Changed By.: Marcia G. Akins on 09 January 2007
					*** Reason.....: only add native properties and then, only if it is not activeform or activecontrol
					***********************************************************************
					llOK2Add = ( llNative And lcTypeAbbrev == [P] And Not Inlist( Lower( Alltrim( lcName ) ), [activecontrol], [activeform] ) ) Or Not llNative
					llOK2Add = ( llNative And Not Inlist( Lower( Alltrim( lcName ) ), [activecontrol], [activeform] ) ) Or Not llNative
					If Not llOK2Add
						Loop
					Endif

					Insert Into (tcCursorName) ;
						( cName, lNative, lInherited, ;
						cType, lAccess, lAssign, nFavorites, nVisibility, ;
						lNonDefault, lHasCode, lReadOnly, cDescript ) ;
						VALUES ;
						( IIf( Not Empty( lcDisplay ), lcDisplay, lcName ), llNative, llInherited2, ;
						lcTypeAbbrev, llAccess, llAssign, 1 + IIf(llFavorites, 1, 0) + IIf(llGlobalFavorites, 2, 0), lnVisibility, ;
						llNonDefault, llHasCode, llReadOnly, lcDescription  )
					***********************************************************************
					* Add the PEM object to the collection. Due to a bug in AMEMBERS() that
					* sometimes includes the same PEM more than once in the array, we'll trap for
					* an duplicate item error when we add to the collection.

					loPEM = Createobject('MemberObject')
					With loPEM
						.Name             = lcName
						.NativePEM        = llNative
						.Type             = lcType
						.ClassMemberData  = loPEMMemberData
						.GlobalMemberData = loGlobalMemberData
						.Display          = Evl(loPEMMemberData.Display, loGlobalMemberData.Display)
						.Script           = Evl(loPEMMemberData.Script, loGlobalMemberData.Script)
					Endwith
					Try
						.oMemberDataCollection.Add(loPEM, lcPEM)
					Catch
					Endtry
				Endif lcType <> 'object'

				* update cName ... with array dimensions if appropriate
				*  If "A" = Type(".oObject." + lcPEM, 1)
				If "U" # Type("ALen(.oObject." + lcPEM + ")")
					Replace cName With Trim(cName) + '[' + Transform (Alen (.oObject.&lcPEM,1));
						+ IIf( 0 # Alen (.oObject.&lcPEM,2),									;
						"," + Transform (Alen (.oObject.&lcPEM,2)) , '') + ']' In (tcCursorName)
				Endif
			Next lnI

			* Now get member data from all classes up the class hierarchy so it's displayed
			* properly.

			***				*** DougHennig 05/23/2007: set lClassDesigner. If we're modifying a form, Parent.BaseClass
			***				***		is "Formset", while it's "Form" for a class
			***				.lClassDesigner = .oObject.Parent.BaseClass <> 'Formset'
			try
				lcClass   = IIf(tlTopOfClass and .lClassDesigner, Lower(.oObject.ParentClass), ;
					lower(.oObject.Class))
				lcLibrary = .oObject.ClassLibrary
			catch
				lcLibrary = ''
			endtry
			
			lcPrevClass = lcClass
			lcPrevLibrary = lcLibrary

			Do While Not Empty(lcLibrary)
				* can bot process classes based on PRGs
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
						*!* ******************** Removed 11/19/2014 *****************
						*!* If tlNoErrorMsgOnMissingVCX
						*!* 	Wait (lcPrompt) Window Nowait
						*!* Else
						*!* 	Messagebox (lcPrompt, 0, ccLOC_MEM_CAP_TITLE)
						*!* Endif
						lcLibrary = ''
						Exit
					Endif
				Endif

				lcPrevClass = lcClass
				lcPrevLibrary = lcLibrary

				Select 0
				Use (lcLibrary) Again Shared
				Locate For OBJNAME == lcClass And Lower(RESERVED1) = 'class'

				lnPos = At('_memberdata = ', PROPERTIES)
				If lnPos > 0
					lnPos = lnPos + 14

					* We have to handle properties with more than 255 characters in the value
					* differently.
					If Substr(PROPERTIES, lnPos, 1) = ccPROPERTIES_PADDING_CHAR
						lnLen        = Val(Alltrim(Substr(PROPERTIES, ;
							lnPos + cnPROPERTIES_PADDING_SIZE, ;
							cnPROPERTIES_LEN_SIZE)))
						lcMemberData = Substr(PROPERTIES, lnPos + ;
							cnPROPERTIES_PADDING_SIZE + cnPROPERTIES_LEN_SIZE, ;
							lnLen)
					Else
						lcMemberData = Strextract(Substr(PROPERTIES, lnPos), ;
							'', ccCR)
					Endif Substr(PROPERTIES, lnPos, 1) = ccPROPERTIES_PADDING_CHAR
					If Not Empty(lcMemberData)
						loDOM.LoadXML(lcMemberData)
						If loDOM.parseError.errorCode = 0
							loNodes = loDOM.selectNodes('//' + ;
								ccMEMBER_DATA_XML_ELEMENT)
							*** DougHennig 05/23/2007: save the current work area and select the csrMembers cursor
							lnSelect = Select()
							Select (tcCursorName)
							For Each loNode In loNodes
								lcPEM     = Upper(loNode.getAttribute('name'))
								lcDisplay = Nvl(loNode.getAttribute('display'), '')
								lcScript  = Nvl(loNode.getAttribute('script'), '')
								Try
									loPEM           = .oMemberDataCollection.Item(lcPEM)
									loPEM.Display   = Evl(loPEM.Display, ;
										lcDisplay)
									loPEM.Name      = Evl(loPEM.Display, ;
										loPEM.Name)
									loPEMMemberData = Createobject('MemberDataObject')
									With loPEMMemberData
										.HasMemberData = .T.
										.Location      = lcClass + ccLOC_OF + ;
											justfname(lcLibrary)
										.Display       = lcDisplay
										.Favorites     = Nvl(loNode.getAttribute('favorites'), ;
											'')
										.Override      = Nvl(loNode.getAttribute('override'), ;
											'')
										.Script        = Nvl(loNode.getAttribute('script'), ;
											'')
										.OriginalDisplay   = .Display
										.OriginalFavorites = .Favorites
										.OriginalOverride  = .Override
										.OriginalScript    = .Script
										.CustomAttributes  = This.GetCustomAttributesForMember(lcPEM, ;
											loDOM)
									Endwith

									* Store the MemberData object in the ParentMemberData property of the
									* appropriate object.

									Do While Vartype(loPEM.ParentMemberData) = 'O'
										loPEM = loPEM.ParentMemberData
									Enddo While Vartype(loPEM.ParentMemberData) = 'O'
									loPEM.ParentMemberData = loPEMMemberData
									*** DougHennig 05/23/2007: update the csrMembers cursor
									Locate For Upper(Trim(cName)) == Upper(lcPEM)
									*** DougHennig 11/21/2008: find the correct name for arrays
									* If Not Found() And Type('.oObject.' + lcPEM, 1) = 'A'
									If Not Found() And "U" # Type("Alen(.oObject." + lcPEM + ")")
										Locate For Upper(Trim(cName)) = Upper(lcPEM) + '['
										If Found()
											lcDisplay = Strtran(cName, lcPEM, lcDisplay, -1, -1, 1)
										Endif Found()
									Endif Not Found() ...
									*** DougHennig 11/21/2008: end of new code
									****************************************************************
									* Modified 10/25/2008 by Jim Nelson
									* Uses EVL for unexplained case where lcDisplay is blank
									Replace ;
										cName 		With Evl(lcDisplay, cName), ;
										nFavorites  With 1 + IIf(loPEMMemberData.Favorites = ccXML_TRUE, 1, 0) + IIf(nFavorites >= 3, 2, 0) ;
										cScript     With Evl (cScript, lcScript)
								Catch
								Endtry
							Next loNode
							*** DougHennig 05/23/2007: reselect the former work area
							Select (lnSelect)
						Endif loDOM.parseError.errorCode = 0
					Endif Not Empty(.cMemberData)
				Endif lnPos > 0
				lcClass   = Lower(Class)

				If Not Empty(CLASSLOC)
					lcLibrary = Fullpath(CLASSLOC, Addbs(Justpath(lcLibrary)))
					If Not File( m.lcLibrary )
						lcLibrary = Fullpath( CLASSLOC )
					Endif
				Else
					lcLibrary = ''
				Endif Not Empty(CLASSLOC)

				Use
			Enddo While Not Empty(lcLibrary)

			*** JRN 06/01/09 : Per Doug H
			*!*	* Handle MemberData in the containership hierarchy.

			*!*	loParent = Iif(Type('.oObject.Parent.Name') = 'C', ;
			*!*		.oObject.Parent, .Null.)
			*!*	Do While Vartype(loParent) = 'O'
			*!*		If Pemstatus(loParent, '_MemberData', 5)
			*!*			loDOM.LoadXML(loParent._MemberData)
			*!*			If loDOM.parseError.errorCode = 0
			*!*				loNodes = loDOM.selectNodes('//' + ;
			*!*					ccMEMBER_DATA_XML_ELEMENT)
			*!*				For Each loNode In loNodes
			*!*					lcPEM     = Upper(loNode.getAttribute('name'))
			*!*					lcDisplay = Nvl(loNode.getAttribute('display'), '')
			*!*					Try
			*!*						loPEM = .oMemberDataCollection.Item(lcPEM)
			*!*						.SetupContainershipHierarchy(loPEM)
			*!*						loPEM.Display   = Evl(loPEM.Display, ;
			*!*							lcDisplay)
			*!*						loPEM.Name      = Evl(loPEM.Display, ;
			*!*							loPEM.Name)
			*!*						loPEMMemberData = .FindMemberDataForContainer(loPEM, ;
			*!*							loParent.Name)
			*!*						With loPEMMemberData
			*!*							.HasMemberData = .T.
			*!*							.Location      = loParent.Name
			*!*							.Display       = lcDisplay
			*!*							.Favorites     = Nvl(loNode.getAttribute('favorites'), ;
			*!*								'')
			*!*							.Override      = Nvl(loNode.getAttribute('override'), ;
			*!*								'')
			*!*							.Script        = Nvl(loNode.getAttribute('script'), ;
			*!*								'')
			*!*							.OriginalDisplay   = .Display
			*!*							.OriginalFavorites = .Favorites
			*!*							.OriginalOverride  = .Override
			*!*							.OriginalScript    = .Script
			*!*							.CustomAttributes  = This.GetCustomAttributesForMember(lcPEM, ;
			*!*								loDOM)
			*!*						Endwith
			*!*					Catch
			*!*					Endtry
			*!*				Next loNode
			*!*			Endif loDOM.parseError.errorCode = 0
			*!*		Endif Pemstatus(loParent, '_MemberData', 5)
			*!*		Try
			*!*			loParent = loParent.Parent
			*!*		Catch
			*!*			loParent = .Null.
			*!*		Endtry
			*!*	Enddo While Vartype(loParent) = 'O'
		Endwith

		* Close the cursors.

		Use In _PROPERTIES
		Use In Global
	Endfunc

	* Return the custom property editor script, if any, for a member
	Function GetMemberScript (tcMember)
		Local lcResult
		Try
			lcResult = This.oMemberDataCollection.Item (Upper (Trim (tcMember))).Script
		Catch
			lcResult = ''
		Endtry
		Return lcResult
	Endfunc

	* Determine if the specified MemberData entry has a setting for the specified
	* member.

	Function HasMemberData(tcMember)
		Return Not Empty(Strextract(This.cMemberData, 'memberdata name="' + ;
			tcMember + '"', '/>', 1, 1))
	Endfunc

	* Determine if the specified member has global member data.

	Function HasGlobalMemberData(tcMember)
		Return Seek(Padr(Upper(tcMember), Len(Global.ABBREV)), 'GLOBAL')
	Endfunc

	* Find the specified attribute for the specified member.

	Function FindAttributeForMember(tcMember, tcAttribute, toDOM)
		Local loDOM, ;
			lcValue, ;
			loNode
		If Vartype(toDOM) = 'O'
			loDOM = toDOM
		Else
			loDOM = This.oXMLDOM
		Endif Vartype(toDOM) = 'O'
		lcValue = ''
		loNode  = loDOM.selectSingleNode('//' + ccMEMBER_DATA_XML_ELEMENT + ;
			'[@name = "' + Lower(tcMember) + '"]')
		If Vartype(loNode) = 'O'
			lcValue = Nvl(loNode.getAttribute(Lower(tcAttribute)), '')
		Endif Vartype(loNode) = 'O'
		Return lcValue
	Endfunc

	* Get any custom attribute for the specified member.

	Function GetCustomAttributesForMember(tcMember, toDOM)
		Local loAttributes, ;
			loNode, ;
			lcCurrExact, ;
			loAttribute
		loAttributes = Createobject('Collection')
		loNode       = toDOM.selectSingleNode('//' + ;
			ccMEMBER_DATA_XML_ELEMENT + '[@name = "' + Lower(tcMember) + '"]')
		If Vartype(loNode) = 'O'
			lcCurrExact = Set('EXACT')
			Set Exact On
			For Each loAttribute In loNode.Attributes
				If Not Inlist(loAttribute.Name, 'name', 'type', 'display', ;
						'favorites', 'override', 'script')
					loAttributes.Add(loAttribute, loAttribute.Name)
				Endif Not Inlist(loAttribute.Name ...
			Next loAttribute
			If lcCurrExact = 'OFF'
				Set Exact Off
			Endif lcCurrExact = 'OFF'
		Endif Vartype(loNode) = 'O'
		Return loAttributes
	Endfunc

	* Create a hierarchy of member data objects for the containers of the selected
	* object. We'll grab any parent if it has a _MemberData property or if it's the
	* topmost parent.

	Function SetupContainershipHierarchy(toPEM)
		Local loParent, ;
			loTopObject, ;
			llHaveFirstParent, ;
			loMemberData
		loParent          = IIf(Type('This.oObject.Parent.Name') = 'C', ;
			This.oObject.Parent, .Null.)
		loTopObject       = PEME_FindTopmostParent(This.oObject, ;
			This.lClassDesigner)
		llHaveFirstParent = .F.
		loMemberData      = toPEM
		Do While Vartype(loParent) = 'O'
			If Pemstatus(loParent, '_MemberData', 5) Or ;
					compobj(loParent, loTopObject)
				If llHaveFirstParent
					loMemberData.ParentMemberData = Createobject('MemberDataObject')
					loMemberData.ParentMemberData.Location  = loParent.Name
					loMemberData = loMemberData.ParentMemberData
				Else
					loMemberData.ContainerMemberData = Createobject('MemberDataObject')
					loMemberData.ContainerMemberData.Location  = loParent.Name
					loMemberData = loMemberData.ContainerMemberData
					llHaveFirstParent = .T.
				Endif llHaveFirstParent
			Endif Pemstatus(loParent, '_MemberData', 5) ...
			Try
				loParent = loParent.Parent
			Catch
				loParent = .Null.
			Endtry
		Enddo While Vartype(loParent) = 'O'
		Return
	Endfunc

	* Finds the appropriate member data object for the specified container.

	Function FindMemberDataForContainer(toPEM, tcContainer)
		Local loMemberData
		loMemberData = toPEM.ContainerMemberData
		Do While Vartype(loMemberData) = 'O' And ;
				loMemberData.Location <> tcContainer
			loMemberData = loMemberData.ParentMemberData
		Enddo While Vartype(loMemberData) = 'O' ...
		Return loMemberData
	Endfunc

	* Create a MemberData node for the specified member and member data set.

	Function CreateMemberDataForMember(toPEM, toMemberData, ;
			tlIgnoreInheritance)
		Local loNode
		loNode = This.oXMLDOM.createElement(ccMEMBER_DATA_XML_ELEMENT)
		loNode.setAttribute('name', Lower(toPEM.Name))
		loNode.setAttribute('type', toPEM.Type)
		This.UpdateMemberDataForMember(toPEM, toMemberData, loNode, ;
			tlIgnoreInheritance)
		Return loNode
	Endfunc

	Function UpdateMemberDataForMember(toPEM, toMemberData, toNode, ;
			tlIgnoreInheritance)
		Local llInherited, ;
			lnI, ;
			lcAttribute, ;
			loAttribute
		With toMemberData
			llInherited = .Inherited And Not tlIgnoreInheritance

			* If we have a display value and either this isn't a subclass or it is but the
			* display value is different, create the display attribute. If we don't have a
			* display value but we have a former display atrribute, remove it.

			Do Case
				Case Not Empty(.Display) And (Not llInherited Or ;
						not .Display == .OriginalDisplay)
					toNode.setAttribute('display', .Display)
				Case Empty(.Display) And ;
						not Empty(Nvl(toNode.getAttribute('display'), ''))
					toNode.removeAttribute('display')
			Endcase

			* If this is supposed to be a favorite and this isn't a subclass, or it is a
			* subclass but the member's favorite status is different in the subclass,
			* create the favorite attribute. If we don't have a favorite value but we have
			* a former favorite atrribute, remove it.

			Do Case
				Case Not Empty(.Favorites) And (Not llInherited Or ;
						.Favorites <> .OriginalFavorites)
					toNode.setAttribute('favorites', .Favorites)
				Case Empty(.Favorites) And ;
						not Empty(Nvl(toNode.getAttribute('favorites'), ''))
					toNode.removeAttribute('favorites')
			Endcase

			* If this is supposed to be an override and this isn't a subclass, or it is a
			* subclass but the member's override status is different in the subclass,
			* create the override attribute. If we don't have a override value but we have
			* a former override atrribute, remove it.

			Do Case
				Case Not Empty(.Override) And (Not llInherited Or ;
						.Override <> .OriginalOverride)
					toNode.setAttribute('override', .Override)
				Case Empty(.Override) And ;
						not Empty(Nvl(toNode.getAttribute('override'), ''))
					toNode.removeAttribute('override')
			Endcase

			* If we have a script and either this isn't a subclass or it is but the script
			* is different, create the script attribute. If we don't have a script but we
			* have a former script atrribute, remove it.

			Do Case
				Case Not Empty(.Script) And (Not llInherited Or ;
						not .Script == .OriginalScript)
					toNode.setAttribute('script', .Script)
				Case Empty(.Script) And ;
						not Empty(Nvl(toNode.getAttribute('script'), ''))
					toNode.removeAttribute('script')
			Endcase

			* Remove all custom sttributes (we'll add any we're supposed to save next).

			For lnI = toNode.Attributes.Length - 1 To 0 Step -1
				lcAttribute = toNode.Attributes(lnI).Name
				If Not Inlist(Lower(lcAttribute), 'name', 'type', 'display', ;
						'favorites', 'override', 'script')
					toNode.removeAttribute(lcAttribute)
				Endif Not Inlist(Lower(lcAttribute) ...
			Next lnI

			* Handle custom attributes.

			If Vartype(.CustomAttributes) = 'O'
				For Each loAttribute In .CustomAttributes
					toNode.setAttribute(loAttribute.Name, loAttribute.Value)
				Next loAttribute
			Endif Vartype(.CustomAttributes) = 'O'
		Endwith
	Endfunc

	* Hide the Properties window if it's open.

	Function HidePropertiesWindow
		If Wvisible('Properties')
			Adockstate(This.aDockWindow)
			Hide Window PROPERTIES
			This.lPropertiesWindowOpen = .T.
		Endif Wvisible('Properties')
	Endfunc

	* Redisplay the Properties window if necessary; if it was docked, we may need
	* to redock it (it's "before" docking state may not match the docking state
	* when we re-activate the window.

	Function ShowPropertiesWindow
		Local laDockWindow[1], ;
			lnCurrRow, ;
			lnRow, ;
			lnTabRow, ;
			lnPosition, ;
			lcWindow
		With This
			If .lPropertiesWindowOpen
				Activate Window PROPERTIES
				Adockstate(laDockWindow)
				lnCurrRow = Ascan(laDockWindow, 'PROPERTIES', -1, -1, 1, 15)
				lnRow     = Ascan(.aDockWindow, 'PROPERTIES', -1, -1, 1, 15)
				If lnRow > 0 And .aDockWindow[lnRow, 2] = 1 And ;
						(.aDockWindow[lnRow, 3] <> laDockWindow[lnCurrRow, 3] Or ;
						.aDockWindow[lnRow, 4] <> laDockWindow[lnCurrRow, 4] )
					lnTabRow = Ascan(.aDockWindow, 'PROPERTIES', -1, -1, 4, 15)
					If lnTabRow > 0
						lnPosition = 4
						lcWindow   = .aDockWindow[lnTabRow, 1]
					Else
						lnPosition = .aDockWindow[lnRow, 3]
						lcWindow   = .aDockWindow[lnRow, 4]
					Endif lnTabRow > 0
					If Not Inlist(lcWindow, 'VIEW', 'TRACE', 'WATCH', ;
							'COMMAND', 'DOCUMENT', 'LOCALS', 'DEBUG OUTPUT', ;
							'CALL STACK')
						lcWindow = ''
					Endif Not Inlist(lcWindow ...
					lcWindow = IIf(Empty(lcWindow), '', 'window "' + ;
						lcWindow + '"')
					Dock Window PROPERTIES position lnPosition &lcWindow
				Endif lnRow > 0 ...
			Endif .lPropertiesWindowOpen
		Endwith
	Endfunc

	* Add _GetMemberData script to FOXCODE.

	Function RegisterGetMemberDataScript(tlRegister)
		Local lcCode

		* Create the code for the _GetMemberData entry in FOXCODE.

		TEXT to lcCode noshow
lparameters toFoxcode
local laObjects[1], ;
	loObject
if aselobj(laObjects) > 0 or aselobj(laObjects, 1) > 0
	loObject = __FindTopMostParent(laObjects[1])
	if vartype(loObject) = 'O' and not pemstatus(loObject, '_memberdata', 5)
		loObject.AddProperty('_memberdata', '')
	endif vartype(loObject) <> = 'O' ...
endif aselobj(laObjects) > 0 ...
return ''

* Find the top-most parent object for the specified object. The complication is
* that the top level container is a Formset in the Form and Class Designers,
* and even worse, one level down from the Formset is a Form for non-form
* classes in the Class Designer. So, we need to be careful about finding the
* top-most parent object.

function __FindTopMostParent(toObject)
local laObjects[1], ;
	llClass, ;
	loObject1, ;
	loObject2, ;
	llDone
if aselobj(laObjects, 3) <> 3
	return
endif aselobj(laObjects, 3) <> 3
if vartype(laObjects[2]) <> 'C'
	return
endif vartype(laObjects[2]) <> 'C'
llClass = upper(justext(laObjects[2])) = 'VCX'
store toObject to loObject1, loObject2
llDone = .F.
do while not llDone
	if not empty(loObject1.ClassLibrary) or (not llClass and ;
		not upper(loObject1.Class) == 'FORMSET')
		loObject2 = loObject1
	endif not empty(loObject1.ClassLibrary) ...
	if type('loObject1.Parent.Name') = 'C'
		loObject1 = loObject1.Parent
	else
		llDone = .T.
	endif type('loObject1.Parent.Name') = 'C'
enddo while not llDone
return loObject2
		ENDTEXT

		* Add or remove the _GetMemberData record in FOXCODE.

		Try
			Use (_Foxcode) Again Shared
			Locate For Type = ccGLOBAL_MEMBER_DATA_TYPE And ;
				upper(ABBREV) = Upper(ccGETMEMBERDATA_ABBREV) And Not Deleted()
			Do Case
				Case Not tlRegister And Found()
					Delete
				Case Not tlRegister
				Case Found()
					Replace Data With lcCode
				Otherwise
					Insert Into (_Foxcode) (Type, ABBREV, Data) ;
						values (ccGLOBAL_MEMBER_DATA_TYPE, ;
						ccGETMEMBERDATA_ABBREV, lcCode)
			Endcase
			Use
		Catch
		Endtry
	Endfunc

	* Add information about MemberDataEditor to FOXCODE.

	Function RegisterMemberDataEditor(tcPath)
		Local lcXML

		* Create the XML for the _MemberData entry in FOXCODE.

		TEXT to lcXML noshow textmerge pretext 2
		<<ccXML_DECLARATION>>
		<VFPData><memberdata name="_memberdata" type="property" display="_MemberData"
		script="do [<<tcPath>>MemberDataEditor.app]"/></VFPData>
		ENDTEXT

		* Add the _MemberData record to FOXCODE.

		Try
			Use (_Foxcode) Again Shared
			Locate For Type = ccGLOBAL_MEMBER_DATA_TYPE And ;
				upper(ABBREV) = '_MEMBERDATA' And Not Deleted()
			If Found()
				Replace TIP With lcXML
			Else
				Insert Into (_Foxcode) (Type, ABBREV, TIP) ;
					values (ccGLOBAL_MEMBER_DATA_TYPE, '_memberdata', lcXML)
			Endif Found()
			Use
		Catch
		Endtry
	Endfunc

	Function CreateMemberMethodList (tcCursorName, tlTopOfClass, tlMethodsOnly, tlNoErrorMsgOnMissingVCX)
		Local lnKeywordLen, ;
			lnLen, ;
			lnMember, ;
			llMemberDataInherited, ;
			llMemberDataChanged, ;
			llActiveX, ;
			laMembers[1], ;
			lnI, ;
			lcPEM, ;
			lcType, ;
			llNative, ;
			lnPos, ;
			loPEMMemberData, ;
			loGlobalMemberData, ;
			llMemberData, ;
			llInherited, ;
			lcName, ;
			lcClass, ;
			lcLibrary, ;
			lnLen, ;
			lcMemberData, ;
			loPEM, ;
			loDOM, ;
			loNodes, ;
			loNode, ;
			lcDisplay, ;
			loParent, ;
			llFavorites, ;
			llGlobalFavorites, ;
			llAssign, ;
			llAccess, ;
			lcTypeAbbrev, ;
			lnVisibility, ;
			lcFile, ;
			lnSelect , ;
			lnLines, ;
			laLines[ 1 ], ;
			lnI, ;
			lcDescription, ;
			lcDescr, ;
			lnRow, ;
			loErr, ;
			llOK2Add, ;
			llInherited2, ;
			llReadOnly,		;
			llChanged, 		;
			llNonDefault,	;
			llLocalCode, 	;
			llHasCode,		;
			lcPrompt

		***********************************************************************
		*** Changed By.: Marcia G. Akins on 05 January 2007
		*** Reason.....: Since I am going to display this data in a grid, it is more useful for
		*** ...........: me to have it in a cursor instead of a collection
		***********************************************************************

		Dimension laLines[ 1 ]

		* Create a cursor of MemberData records from the FOXCODE table so we can look
		* for global PEMs. Bug out if we can't open FOXCODE.

		Try
			Use (_Foxcode) Again Shared In 0 Alias Foxcode
		Catch
		Endtry
		If Not Used('FOXCODE')
			Create Cursor FoxCode (Type C(1), Abbrev C(24), Tip C(1))
		Endif Not Used('FOXCODE')
		Select ABBREV From FoxCode ;
			where Type = 'P' And Not Deleted() ;
			into Cursor _PROPERTIES
		Index On Upper(ABBREV) Tag ABBREV
		Select ABBREV, TIP From FoxCode ;
			where Type = ccGLOBAL_MEMBER_DATA_TYPE And Not Deleted() ;
			into Cursor Global
		Index On Upper(ABBREV) Tag ABBREV
		Use In Foxcode
		lnLen = Len(ABBREV)

		* Instantiate an MS XML DOM object.

		loDOM = Createobject(ccXML_DOM_CLASS)
		loDOM.Async = .F.

		* Put all the object's PEMs into a collection. For ActiveX controls, we'll use
		* AMEMBERS again to get the correct case for members.

		With This
			.oMemberDataCollection = Createobject('Collection')
			lnMember = Ascan(.aObjectMembers, '_MemberData', -1, -1, 1, 15)
			If lnMember > 0
				llMemberDataInherited = 'I' $ .aObjectMembers[lnMember, 3]
				llMemberDataChanged   = 'C' $ .aObjectMembers[lnMember, 3]
			Endif lnMember > 0
			llActiveX = pemstatus(.oObject, 'BaseClass', 5) and Lower(.oObject.BaseClass) = 'olecontrol'
			If llActiveX
				Amembers(laMembers, .oObject, 3)
			Endif llActiveX

			For lnI = 1 To Alen(.aObjectMembers, 1)
				lcPEM    = .aObjectMembers[lnI, 1]
				lcType   = Lower(.aObjectMembers[lnI, 2])

				If tlMethodsOnly And Not Inlist (lcType, 'method', 'event')
					Loop
				Endif

				llNative = 'N' $ .aObjectMembers[lnI, 3]
				llInherited2 = 'I' $ .aObjectMembers[lnI, 3]
				* For each PEM that doesn't hold an object, get the information about its
				* MemberData.

				If lcType <> 'object'
					llMemberData = .HasMemberData(lcPEM)

					* This member has inherited member data if the member data was changed and this
					* member doesn't appear in it or the member data wasn't changed and this member
					* does appear (because we've just inherited the member data).

					llInherited = ((llMemberDataChanged And Not llMemberData) Or ;
						(Not llMemberDataChanged And llMemberData)) And ;
						llMemberDataInherited And 'I' $ .aObjectMembers[lnI, 3]

					* If we can find the member in FOXCODE, we'll use the abbreviation stored there
					* since it may have the correct case.

					lcName = IIf(Seek(Padr('.' + lcPEM, lnLen), '_PROPERTIES'), ;
						trim(Substr(_PROPERTIES.ABBREV, 2)), lcPEM)

					lcName = This.SpellNative(lcPEM, lcName)

					* Get the correct case for members of ActiveX controls.

					If llActiveX
						lnPos = Ascan(laMembers, lcPEM, -1, -1, 1, 15)
						If lnPos > 0
							lcName = laMembers[lnPos, 1]
						Endif lnPos > 0
					Endif llActiveX

					* If the name is still upper-cased and it's a native or ActiveX control
					* property, use PROPER() on it.

					If Upper(lcName) == lcName
						If llNative && Or llActiveX
							lcName = Proper(lcName)
							***********************************************************************
							*** Changed By.: Marcia G. Akins on 05 January 2007
							*** Reason.....: I want custom properties to shake out at the bottom of the list
							***********************************************************************
						Else
							lcName = Lower( lcName )
						Endif
					Endif Upper(lcName) == lcName ...

					* Create an object to hold information about the MemberData.

					loPEMMemberData = Createobject('MemberDataObject')
					With loPEMMemberData
						.HasMemberData = llMemberData
						.Inherited     = llInherited
						If llMemberData
							.Display   = This.FindAttributeForMember(lcPEM, ;
								'Display')
							.Favorites = This.FindAttributeForMember(lcPEM, ;
								'Favorites')
							.Override  = This.FindAttributeForMember(lcPEM, ;
								'Override')
							.Script    = This.FindAttributeForMember(lcPEM, ;
								'Script')
							.OriginalDisplay   = .Display
							.OriginalFavorites = .Favorites
							.OriginalOverride  = .Override
							.OriginalScript    = .Script
							lcName             = Evl(.Display, lcName)
						Endif llMemberData
						.CustomAttributes = This.GetCustomAttributesForMember(lcPEM, ;
							This.oXMLDOM)
					Endwith

					***********************************************************************
					*** Changed By.: Marcia G. Akins on 05 January 2007
					*** Reason.....: Add a record to the cursor
					***********************************************************************
					lcDescription = []
					*   If Type( [laLines], 1 ) = [A]
					If Alen (laLines) >= 4
						lnRow = Ascan( laLines, lcName, 1, -1, 1, 15 )
						If lnRow > 0
							lcDescription = laLines[ lnRow, 4 ]
						Endif
					Endif

					* Now get any global MemberData for the PEM.

					loGlobalMemberData = Createobject('MemberDataObject')
					With loGlobalMemberData
						.HasMemberData = This.HasGlobalMemberData(lcPEM)
						If .HasMemberData
							loDOM.LoadXML(Global.TIP)
							If loDOM.parseError.errorCode = 0
								.Display           = This.FindAttributeForMember(lcPEM, ;
									'Display',   loDOM)
								.Favorites         = This.FindAttributeForMember(lcPEM, ;
									'Favorites', loDOM)
								.Override          = This.FindAttributeForMember(lcPEM, ;
									'Override',  loDOM)
								.Script            = This.FindAttributeForMember(lcPEM, ;
									'Script',    loDOM)
								.OriginalDisplay   = .Display
								.OriginalFavorites = .Favorites
								.OriginalOverride  = .Override
								.OriginalScript    = .Script
								lcName             = IIf(Upper(lcName) == lcName, ;
									evl(.Display, lcName), lcName)

								lcDesc				= This.FindAttributeForMember(lcPEM, ;
									'Description',    loDOM)
								If Not Isnull(lcDesc)
									lcDescription = lcDesc
								Endif

								* Get any custom attributes.

								.CustomAttributes = This.GetCustomAttributesForMember(lcPEM, ;
									loDOM)
							Endif loDOM.parseError.errorCode = 0
						Endif .HasMemberData
					Endwith

					lcDisplay = Evl(loPEMMemberData.Display, loGlobalMemberData.Display)

					*** JRN 02/03/2009 : Can't use PEMStatus(5) on OLEs

					*** DougHennig 05/23/2007: fixed bug: the name of the property is contained in lcName, not "cName"
					***              llAssign = PEMSTATUS( This.oObject, [cName] + [_Assign], 5 )
					***              llAccess = PEMSTATUS( This.oObject, [cName] + [_Access], 5 )
					*** llAssign = Pemstatus( This.oObject, lcName + [_Assign], 5 )
					*** llAccess = Pemstatus( This.oObject, lcName + [_Access], 5 )
					llAssign = 0 # Ascan(.aObjectMembers, lcName + [_Assign], -1, -1, 1, 15)
					llAccess = 0 # Ascan(.aObjectMembers, lcName + [_Access], -1, -1, 1, 15)

					llFavorites = loPEMMemberData.Favorites == ccXML_TRUE
					llGlobalFavorites = loGlobalMemberData.Favorites == ccXML_TRUE
					lcTypeAbbrev = Upper( Left( This.aObjectMembers[lnI, 2], 1 ) )
					lnVisibility = IIf( [G] $This.aObjectMembers[lnI, 3], 1, IIf([P] $This.aObjectMembers[lnI, 3], 2, 3))
					llReadOnly = 'R' $ This.aObjectMembers[lnI, 3]
					Do Case
						Case lnVisibility = 1 && Public
						Case Not (llNative Or llInherited2) && Custom
						Case tlTopOfClass && class -- thisform
						Case lnVisibility = 2 && Protected
							llReadOnly = .T.
						Case lnVisibility = 3 && Hidden
							Loop
					Endcase

					*** JRN - 11/20/2008 -- capture which PEMs are non-default
					llChanged = 'C' $ This.aObjectMembers[lnI, 3]
					Do Case
						Case lcTypeAbbrev $ 'P'
							llNonDefault = llChanged
							llLocalCode = .F.
							llHasCode	= .F.
						Case llChanged
							llNonDefault = Not Empty (This.oObject.ReadMethod (lcName))
							llLocalCode = llNonDefault
							llHasCode	= llNonDefault Or 'N' $ This.aObjectMembers[lnI, 3] Or 'I' $ This.aObjectMembers[lnI, 3]
						Otherwise
							llNonDefault = .F.
							llLocalCode = .F.
							llHasCode	= .F.
					Endcase

					***********************************************************************
					*** Changed By.: Marcia G. Akins on 09 January 2007
					*** Reason.....: only add native properties and then, only if it is not activeform or activecontrol
					***********************************************************************
					llOK2Add = ( llNative And lcTypeAbbrev == [P] And Not Inlist( Lower( Alltrim( lcName ) ), [activecontrol], [activeform] ) ) Or Not llNative
					llOK2Add = ( llNative And Not Inlist( Lower( Alltrim( lcName ) ), [activecontrol], [activeform] ) ) Or Not llNative
					If Not llOK2Add
						Loop
					Endif

					Insert Into (tcCursorName) ;
						( cName, lNative, lInherited, ;
						cType, lAccess, lAssign, nFavorites, nVisibility, ;
						lNonDefault, lHasCode, lReadOnly, cDescript ) ;
						VALUES ;
						( IIf( Not Empty( lcDisplay ), lcDisplay, lcName ), llNative, llInherited2, ;
						lcTypeAbbrev, llAccess, llAssign, 1 + IIf(llFavorites, 1, 0) + IIf(llGlobalFavorites, 2, 0), lnVisibility, ;
						llNonDefault, llHasCode, llReadOnly, lcDescription  )
					***********************************************************************
					* Add the PEM object to the collection. Due to a bug in AMEMBERS() that
					* sometimes includes the same PEM more than once in the array, we'll trap for
					* an duplicate item error when we add to the collection.

					loPEM = Createobject('MemberObject')
					With loPEM
						.Name             = lcName
						.NativePEM        = llNative
						.Type             = lcType
						.ClassMemberData  = loPEMMemberData
						.GlobalMemberData = loGlobalMemberData
						.Display          = Evl(loPEMMemberData.Display, loGlobalMemberData.Display)
						.Script           = Evl(loPEMMemberData.Script, loGlobalMemberData.Script)
					Endwith
					Try
						.oMemberDataCollection.Add(loPEM, lcPEM)
					Catch
					Endtry
				Endif lcType <> 'object'

				* update cName ... with array dimensions if appropriate
				*  If "A" = Type(".oObject." + lcPEM, 1)
				If "U" # Type("ALen(.oObject." + lcPEM + ")")
					Replace cName With Trim(cName) + '[' + Transform (Alen (.oObject.&lcPEM,1));
						+ IIf( 0 # Alen (.oObject.&lcPEM,2),									;
						"," + Transform (Alen (.oObject.&lcPEM,2)) , '') + ']' In (tcCursorName)
				Endif
			Next lnI

			* Now get member data from all classes up the class hierarchy so it's displayed
			* properly.

			***				*** DougHennig 05/23/2007: set lClassDesigner. If we're modifying a form, Parent.BaseClass
			***				***		is "Formset", while it's "Form" for a class
			***				.lClassDesigner = .oObject.Parent.BaseClass <> 'Formset'
			try
				lcClass   = IIf(tlTopOfClass and .lClassDesigner, Lower(.oObject.ParentClass), ;
					lower(.oObject.Class))
				lcLibrary = .oObject.ClassLibrary
			catch
				lcLibrary = ''
			endtry

			Do While Not Empty(lcLibrary)
				If Not File (lcLibrary)
					lcPrompt = Strtran (ccLOC_File_Not_Found, '<FileName>', Justfname(lcLibrary))
					*!* ******************** Removed 11/19/2014 *****************
					*!* If tlNoErrorMsgOnMissingVCX
					*!* 	Wait (lcPrompt) Window Nowait
					*!* Else
					*!* 	Messagebox (lcPrompt, 0, ccLOC_MEM_CAP_TITLE)
					*!* Endif
					lcLibrary = ''
					Exit
				Endif

				Select 0
				Use (lcLibrary) Again Shared
				Locate For OBJNAME == lcClass And Lower(RESERVED1) = 'class'

				lnPos = At('_memberdata = ', PROPERTIES)
				If lnPos > 0
					lnPos = lnPos + 14

					* We have to handle properties with more than 255 characters in the value
					* differently.
					If Substr(PROPERTIES, lnPos, 1) = ccPROPERTIES_PADDING_CHAR
						lnLen        = Val(Alltrim(Substr(PROPERTIES, ;
							lnPos + cnPROPERTIES_PADDING_SIZE, ;
							cnPROPERTIES_LEN_SIZE)))
						lcMemberData = Substr(PROPERTIES, lnPos + ;
							cnPROPERTIES_PADDING_SIZE + cnPROPERTIES_LEN_SIZE, ;
							lnLen)
					Else
						lcMemberData = Strextract(Substr(PROPERTIES, lnPos), ;
							'', ccCR)
					Endif Substr(PROPERTIES, lnPos, 1) = ccPROPERTIES_PADDING_CHAR
					If Not Empty(lcMemberData)
						loDOM.LoadXML(lcMemberData)
						If loDOM.parseError.errorCode = 0
							loNodes = loDOM.selectNodes('//' + ;
								ccMEMBER_DATA_XML_ELEMENT)
							*** DougHennig 05/23/2007: save the current work area and select the csrMembers cursor
							lnSelect = Select()
							Select (tcCursorName)
							For Each loNode In loNodes
								lcPEM     = Upper(loNode.getAttribute('name'))
								lcDisplay = Nvl(loNode.getAttribute('display'), '')
								Try
									loPEM           = .oMemberDataCollection.Item(lcPEM)
									loPEM.Display   = Evl(loPEM.Display, ;
										lcDisplay)
									loPEM.Name      = Evl(loPEM.Display, ;
										loPEM.Name)
									loPEMMemberData = Createobject('MemberDataObject')
									With loPEMMemberData
										.HasMemberData = .T.
										.Location      = lcClass + ccLOC_OF + ;
											justfname(lcLibrary)
										.Display       = lcDisplay
										.Favorites     = Nvl(loNode.getAttribute('favorites'), ;
											'')
										.Override      = Nvl(loNode.getAttribute('override'), ;
											'')
										.Script        = Nvl(loNode.getAttribute('script'), ;
											'')
										.OriginalDisplay   = .Display
										.OriginalFavorites = .Favorites
										.OriginalOverride  = .Override
										.OriginalScript    = .Script
										.CustomAttributes  = This.GetCustomAttributesForMember(lcPEM, ;
											loDOM)
									Endwith

									* Store the MemberData object in the ParentMemberData property of the
									* appropriate object.

									Do While Vartype(loPEM.ParentMemberData) = 'O'
										loPEM = loPEM.ParentMemberData
									Enddo While Vartype(loPEM.ParentMemberData) = 'O'
									loPEM.ParentMemberData = loPEMMemberData
									*** DougHennig 05/23/2007: update the csrMembers cursor
									Locate For Upper(Trim(cName)) == Upper(lcPEM)
									*** DougHennig 11/21/2008: find the correct name for arrays
									* If Not Found() And Type('.oObject.' + lcPEM, 1) = 'A'
									If Not Found() And "U" # Type("Alen(.oObject." + lcPEM + ")")
										Locate For Upper(Trim(cName)) = Upper(lcPEM) + '['
										If Found()
											lcDisplay = Strtran(cName, lcPEM, lcDisplay, -1, -1, 1)
										Endif Found()
									Endif Not Found() ...
									*** DougHennig 11/21/2008: end of new code
									****************************************************************
									* Modified 10/25/2008 by Jim Nelson
									* Uses EVL for unexplained case where lcDisplay is blank
									Replace cName With Evl(lcDisplay, cName), ;
										nFavorites With 1 + IIf(loPEMMemberData.Favorites = ccXML_TRUE, 1, 0) + IIf(nFavorites >= 3, 2, 0)
								Catch
								Endtry
							Next loNode
							*** DougHennig 05/23/2007: reselect the former work area
							Select (lnSelect)
						Endif loDOM.parseError.errorCode = 0
					Endif Not Empty(.cMemberData)
				Endif lnPos > 0
				lcClass   = Lower(Class)

				If Not Empty(CLASSLOC)
					lcLibrary = Fullpath(CLASSLOC, Addbs(Justpath(lcLibrary)))
					If Not File( m.lcLibrary )
						lcLibrary = Fullpath( CLASSLOC )
					Endif
				Else
					lcLibrary = ''
				Endif Not Empty(CLASSLOC)

				Use
			Enddo While Not Empty(lcLibrary)

			*** JRN 06/01/09 : Per Doug H
			*!*	* Handle MemberData in the containership hierarchy.

			*!*	loParent = Iif(Type('.oObject.Parent.Name') = 'C', ;
			*!*		.oObject.Parent, .Null.)
			*!*	Do While Vartype(loParent) = 'O'
			*!*		If Pemstatus(loParent, '_MemberData', 5)
			*!*			loDOM.LoadXML(loParent._MemberData)
			*!*			If loDOM.parseError.errorCode = 0
			*!*				loNodes = loDOM.selectNodes('//' + ;
			*!*					ccMEMBER_DATA_XML_ELEMENT)
			*!*				For Each loNode In loNodes
			*!*					lcPEM     = Upper(loNode.getAttribute('name'))
			*!*					lcDisplay = Nvl(loNode.getAttribute('display'), '')
			*!*					Try
			*!*						loPEM = .oMemberDataCollection.Item(lcPEM)
			*!*						.SetupContainershipHierarchy(loPEM)
			*!*						loPEM.Display   = Evl(loPEM.Display, ;
			*!*							lcDisplay)
			*!*						loPEM.Name      = Evl(loPEM.Display, ;
			*!*							loPEM.Name)
			*!*						loPEMMemberData = .FindMemberDataForContainer(loPEM, ;
			*!*							loParent.Name)
			*!*						With loPEMMemberData
			*!*							.HasMemberData = .T.
			*!*							.Location      = loParent.Name
			*!*							.Display       = lcDisplay
			*!*							.Favorites     = Nvl(loNode.getAttribute('favorites'), ;
			*!*								'')
			*!*							.Override      = Nvl(loNode.getAttribute('override'), ;
			*!*								'')
			*!*							.Script        = Nvl(loNode.getAttribute('script'), ;
			*!*								'')
			*!*							.OriginalDisplay   = .Display
			*!*							.OriginalFavorites = .Favorites
			*!*							.OriginalOverride  = .Override
			*!*							.OriginalScript    = .Script
			*!*							.CustomAttributes  = This.GetCustomAttributesForMember(lcPEM, ;
			*!*								loDOM)
			*!*						Endwith
			*!*					Catch
			*!*					Endtry
			*!*				Next loNode
			*!*			Endif loDOM.parseError.errorCode = 0
			*!*		Endif Pemstatus(loParent, '_MemberData', 5)
			*!*		Try
			*!*			loParent = loParent.Parent
			*!*		Catch
			*!*			loParent = .Null.
			*!*		Endtry
			*!*	Enddo While Vartype(loParent) = 'O'
		Endwith

		* Close the cursors.

		Use In _PROPERTIES
		Use In Global
	Endfunc

	Function CreateMemberDataNew (tcCursorName, tlTopOfClass, tlAvoidBufferOverrun, tlMethodsOnly, tlNoErrorMsgOnMissingVCX)
		Local loDOM As ccXML_DOM_CLASS
		Local loGlobalMemberData As 'MemberDataObject'
		Local loNode As Object
		Local loNodes As Object
		Local loPEM As 'MemberObject'
		Local loPEMMemberData As 'MemberDataObject'
		Local laLines(1), laMembers(1), lcClass, lcDesc, lcDescription, lcDisplay, lcLibrary, lcLibraryNoPath
		Local lcMemberData, lcName, lcPEM, lcPrevClass, lcPrevLibrary, lcPrompt, lcType, lcTypeAbbrev
		Local llAccess, llActiveX, llAssign, llChanged, llFavorites, llGlobalFavorites, llHasCode, llInherited
		Local llInherited2, llLocalCode, llMemberData, llMemberDataChanged, llMemberDataInherited, llNative
		Local llNonDefault, llOK2Add, llReadOnly, lnI, lnLen, lnMember, lnPos, lnRow, lnSelect, lnVisibility

		***********************************************************************
		*** Changed By.: Marcia G. Akins on 05 January 2007
		*** Reason.....: Since I am going to display this data in a grid, it is more useful for
		*** ...........: me to have it in a cursor instead of a collection
		***********************************************************************

		*** JRN 2009-10-02 : The call to AMembers below can crash VFP with 'Buffer Overrun'

		Dimension laLines[ 1 ]
		If Not tlAvoidBufferOverrun
			*** JRN 04/08/2009 : oops .. this fails on grids in Version 8
			Try
				***********************************************************************
				*** Changed By.: Marcia G. Akins on 06 January 2007
				*** Reason.....: Use AMEMBERS() to get the description from column 4
				***********************************************************************
				Amembers( laLines, This.oObject, 3 )
			Catch
			Endtry
		Endif

		* Create a cursor of MemberData records from the FOXCODE table so we can look
		* for global PEMs. Bug out if we can't open FOXCODE.

		Try
			Use (_Foxcode) Again Shared In 0 Alias Foxcode
		Catch
		Endtry
		If Not Used('FOXCODE')
			This.cErrorMessage = ccLOC_CANT_OPEN_FOXCODE
			Return
		Endif Not Used('FOXCODE')
		Select ABBREV From (_Foxcode) ;
			where Type = 'P' And Not Deleted() ;
			into Cursor _PROPERTIES
		Index On Upper(ABBREV) Tag ABBREV
		Select ABBREV, TIP From (_Foxcode) ;
			where Type = ccGLOBAL_MEMBER_DATA_TYPE And Not Deleted() ;
			into Cursor Global
		Index On Upper(ABBREV) Tag ABBREV
		Use In Foxcode
		lnLen = Len(ABBREV)

		* Instantiate an MS XML DOM object.

		loDOM = Createobject(ccXML_DOM_CLASS)
		loDOM.Async = .F.

		* Put all the object's PEMs into a collection. For ActiveX controls, we'll use
		* AMEMBERS again to get the correct case for members.

		With This
			.oMemberDataCollection = Createobject('Collection')
			lnMember = Ascan(.aObjectMembers, '_MemberData', -1, -1, 1, 15)
			If lnMember > 0
				llMemberDataInherited = 'I' $ .aObjectMembers[lnMember, 3]
				llMemberDataChanged   = 'C' $ .aObjectMembers[lnMember, 3]
			Endif lnMember > 0
			llActiveX = pemstatus(.oObject, 'BaseClass', 5) and Lower(.oObject.BaseClass) = 'olecontrol'
			If llActiveX
				Amembers(laMembers, .oObject, 3)
			Endif llActiveX

			For lnI = 1 To Alen(.aObjectMembers, 1)
				lcPEM    = .aObjectMembers[lnI, 1]
				lcType   = Lower(.aObjectMembers[lnI, 2])

				If tlMethodsOnly And Not Inlist (lcType, 'method', 'event')
					Loop
				Endif

				llNative = 'N' $ .aObjectMembers[lnI, 3]
				llInherited2 = 'I' $ .aObjectMembers[lnI, 3]
				* For each PEM that doesn't hold an object, get the information about its
				* MemberData.

				If lcType <> 'object'
					llMemberData = .HasMemberData(lcPEM)

					* This member has inherited member data if the member data was changed and this
					* member doesn't appear in it or the member data wasn't changed and this member
					* does appear (because we've just inherited the member data).

					llInherited = ((llMemberDataChanged And Not llMemberData) Or ;
						(Not llMemberDataChanged And llMemberData)) And ;
						llMemberDataInherited And 'I' $ .aObjectMembers[lnI, 3]

					* If we can find the member in FOXCODE, we'll use the abbreviation stored there
					* since it may have the correct case.

					lcName = IIf(Seek(Padr('.' + lcPEM, lnLen), '_PROPERTIES'), ;
						trim(Substr(_PROPERTIES.ABBREV, 2)), lcPEM)

					lcName = This.SpellNative(lcPEM, lcName)

					* Get the correct case for members of ActiveX controls.

					If llActiveX
						lnPos = Ascan(laMembers, lcPEM, -1, -1, 1, 15)
						If lnPos > 0
							lcName = laMembers[lnPos, 1]
						Endif lnPos > 0
					Endif llActiveX

					* If the name is still upper-cased and it's a native or ActiveX control
					* property, use PROPER() on it.

					If Upper(lcName) == lcName
						If llNative && Or llActiveX
							lcName = Proper(lcName)
							***********************************************************************
							*** Changed By.: Marcia G. Akins on 05 January 2007
							*** Reason.....: I want custom properties to shake out at the bottom of the list
							***********************************************************************
						Else
							lcName = Lower( lcName )
						Endif
					Endif Upper(lcName) == lcName ...

					* Create an object to hold information about the MemberData.

					loPEMMemberData = Createobject('MemberDataObject')
					With loPEMMemberData
						.HasMemberData = llMemberData
						.Inherited     = llInherited
						If llMemberData
							.Display   = This.FindAttributeForMember(lcPEM, ;
								'Display')
							.Favorites = This.FindAttributeForMember(lcPEM, ;
								'Favorites')
							.Override  = This.FindAttributeForMember(lcPEM, ;
								'Override')
							.Script    = This.FindAttributeForMember(lcPEM, ;
								'Script')
							.OriginalDisplay   = .Display
							.OriginalFavorites = .Favorites
							.OriginalOverride  = .Override
							.OriginalScript    = .Script
							lcName             = Evl(.Display, lcName)
						Endif llMemberData
						.CustomAttributes = This.GetCustomAttributesForMember(lcPEM, ;
							This.oXMLDOM)
					Endwith

					***********************************************************************
					*** Changed By.: Marcia G. Akins on 05 January 2007
					*** Reason.....: Add a record to the cursor
					***********************************************************************
					lcDescription = []
					*   If Type( [laLines], 1 ) = [A]
					If Alen (laLines) >= 4
						lnRow = Ascan( laLines, lcName, 1, -1, 1, 15 )
						If lnRow > 0
							lcDescription = laLines[ lnRow, 4 ]
						Endif
					Endif

					* Now get any global MemberData for the PEM.

					loGlobalMemberData = Createobject('MemberDataObject')
					With loGlobalMemberData
						.HasMemberData = This.HasGlobalMemberData(lcPEM)
						If .HasMemberData
							loDOM.LoadXML(Global.TIP)
							If loDOM.parseError.errorCode = 0
								.Display           = This.FindAttributeForMember(lcPEM, ;
									'Display',   loDOM)
								.Favorites         = This.FindAttributeForMember(lcPEM, ;
									'Favorites', loDOM)
								.Override          = This.FindAttributeForMember(lcPEM, ;
									'Override',  loDOM)
								.Script            = This.FindAttributeForMember(lcPEM, ;
									'Script',    loDOM)
								.OriginalDisplay   = .Display
								.OriginalFavorites = .Favorites
								.OriginalOverride  = .Override
								.OriginalScript    = .Script
								lcName             = IIf(Upper(lcName) == lcName, ;
									evl(.Display, lcName), lcName)

								lcDesc				= This.FindAttributeForMember(lcPEM, ;
									'Description',    loDOM)
								If (Not Isnull(lcDesc)) And Not Empty(lcDesc)
									lcDescription = lcDesc
								Endif

								* Get any custom attributes.

								.CustomAttributes = This.GetCustomAttributesForMember(lcPEM, ;
									loDOM)
							Endif loDOM.parseError.errorCode = 0
						Endif .HasMemberData
					Endwith

					lcDisplay = Evl(loPEMMemberData.Display, loGlobalMemberData.Display)

					*** JRN 02/03/2009 : Can't use PEMStatus(5) on OLEs

					*** DougHennig 05/23/2007: fixed bug: the name of the property is contained in lcName, not "cName"
					***              llAssign = PEMSTATUS( This.oObject, [cName] + [_Assign], 5 )
					***              llAccess = PEMSTATUS( This.oObject, [cName] + [_Access], 5 )
					*** llAssign = Pemstatus( This.oObject, lcName + [_Assign], 5 )
					*** llAccess = Pemstatus( This.oObject, lcName + [_Access], 5 )
					llAssign = 0 # Ascan(.aObjectMembers, lcName + [_Assign], -1, -1, 1, 15)
					llAccess = 0 # Ascan(.aObjectMembers, lcName + [_Access], -1, -1, 1, 15)

					llFavorites = loPEMMemberData.Favorites == ccXML_TRUE
					llGlobalFavorites = loGlobalMemberData.Favorites == ccXML_TRUE
					lcTypeAbbrev = Upper( Left( This.aObjectMembers[lnI, 2], 1 ) )
					lnVisibility = IIf( [G] $This.aObjectMembers[lnI, 3], 1, IIf([P] $This.aObjectMembers[lnI, 3], 2, 3))
					llReadOnly = 'R' $ This.aObjectMembers[lnI, 3]
					Do Case
						Case lnVisibility = 1 && Public
						Case Not (llNative Or llInherited2) && Custom
						Case tlTopOfClass && class -- thisform
						Case lnVisibility = 2 && Protected
							llReadOnly = .T.
						Case lnVisibility = 3 && Hidden
							Loop
					Endcase

					*** JRN - 11/20/2008 -- capture which PEMs are non-default
					llChanged = 'C' $ This.aObjectMembers[lnI, 3]
					Do Case
						Case lcTypeAbbrev $ 'P'
							llNonDefault = llChanged Or Not (llNative Or llInherited2) && Custom
							llLocalCode = .F.
							llHasCode	= .F.
						Case llChanged
							llNonDefault = Not Empty (This.oObject.ReadMethod (lcName))
							llLocalCode = llNonDefault
							llHasCode	= llNonDefault Or 'N' $ This.aObjectMembers[lnI, 3] Or 'I' $ This.aObjectMembers[lnI, 3]
						Otherwise
							llNonDefault = .F.
							llLocalCode = .F.
							llHasCode	= .F.
					Endcase

					***********************************************************************
					*** Changed By.: Marcia G. Akins on 09 January 2007
					*** Reason.....: only add native properties and then, only if it is not activeform or activecontrol
					***********************************************************************
					llOK2Add = ( llNative And lcTypeAbbrev == [P] And Not Inlist( Lower( Alltrim( lcName ) ), [activecontrol], [activeform] ) ) Or Not llNative
					llOK2Add = ( llNative And Not Inlist( Lower( Alltrim( lcName ) ), [activecontrol], [activeform] ) ) Or Not llNative
					If Not llOK2Add
						Loop
					Endif

					Insert Into (tcCursorName) ;
						( cName, lNative, lInherited, ;
						cType, lAccess, lAssign, nFavorites, nVisibility, ;
						lNonDefault, lHasCode, lReadOnly, cDescript ) ;
						VALUES ;
						( IIf( Not Empty( lcDisplay ), lcDisplay, lcName ), llNative, llInherited2, ;
						lcTypeAbbrev, llAccess, llAssign, 1 + IIf(llFavorites, 1, 0) + IIf(llGlobalFavorites, 2, 0), lnVisibility, ;
						llNonDefault, llHasCode, llReadOnly, lcDescription  )
					***********************************************************************
					* Add the PEM object to the collection. Due to a bug in AMEMBERS() that
					* sometimes includes the same PEM more than once in the array, we'll trap for
					* an duplicate item error when we add to the collection.

					loPEM = Createobject('MemberObject')
					With loPEM
						.Name             = lcName
						.NativePEM        = llNative
						.Type             = lcType
						.ClassMemberData  = loPEMMemberData
						.GlobalMemberData = loGlobalMemberData
						.Display          = Evl(loPEMMemberData.Display, loGlobalMemberData.Display)
						.Script           = Evl(loPEMMemberData.Script, loGlobalMemberData.Script)
					Endwith
					Try
						.oMemberDataCollection.Add(loPEM, lcPEM)
					Catch
					Endtry
				Endif lcType <> 'object'

				* update cName ... with array dimensions if appropriate
				*  If "A" = Type(".oObject." + lcPEM, 1)
				If "U" # Type("ALen(.oObject." + lcPEM + ")")
					Replace cName With Trim(cName) + '[' + Transform (Alen (.oObject.&lcPEM,1));
						+ IIf( 0 # Alen (.oObject.&lcPEM,2),									;
						"," + Transform (Alen (.oObject.&lcPEM,2)) , '') + ']' In (tcCursorName)
				Endif
			Next lnI

			* Now get member data from all classes up the class hierarchy so it's displayed
			* properly.

			***				*** DougHennig 05/23/2007: set lClassDesigner. If we're modifying a form, Parent.BaseClass
			***				***		is "Formset", while it's "Form" for a class
			***				.lClassDesigner = .oObject.Parent.BaseClass <> 'Formset'
			try
				lcClass   = IIf(tlTopOfClass and .lClassDesigner, Lower(.oObject.ParentClass), ;
					lower(.oObject.Class))
				lcLibrary = .oObject.ClassLibrary
			catch
				lcLibrary = ''
			endtry

			lcPrevClass = lcClass
			lcPrevLibrary = lcLibrary

			Do While Not Empty(lcLibrary)
				* can bot process classes based on PRGs
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
						*!* ******************** Removed 11/19/2014 *****************
						*!* If tlNoErrorMsgOnMissingVCX
						*!* 	Wait (lcPrompt) Window Nowait
						*!* Else
						*!* 	Messagebox (lcPrompt, 0, ccLOC_MEM_CAP_TITLE)
						*!* Endif
						lcLibrary = ''
						Exit
					Endif
				Endif

				lcPrevClass = lcClass
				lcPrevLibrary = lcLibrary

				Select 0
				Use (lcLibrary) Again Shared
				Locate For OBJNAME == lcClass And Lower(RESERVED1) = 'class'

				lnPos = At('_memberdata = ', PROPERTIES)
				If lnPos > 0
					lnPos = lnPos + 14

					* We have to handle properties with more than 255 characters in the value
					* differently.
					If Substr(PROPERTIES, lnPos, 1) = ccPROPERTIES_PADDING_CHAR
						lnLen        = Val(Alltrim(Substr(PROPERTIES, ;
							lnPos + cnPROPERTIES_PADDING_SIZE, ;
							cnPROPERTIES_LEN_SIZE)))
						lcMemberData = Substr(PROPERTIES, lnPos + ;
							cnPROPERTIES_PADDING_SIZE + cnPROPERTIES_LEN_SIZE, ;
							lnLen)
					Else
						lcMemberData = Strextract(Substr(PROPERTIES, lnPos), ;
							'', ccCR)
					Endif Substr(PROPERTIES, lnPos, 1) = ccPROPERTIES_PADDING_CHAR
					If Not Empty(lcMemberData)
						loDOM.LoadXML(lcMemberData)
						If loDOM.parseError.errorCode = 0
							loNodes = loDOM.selectNodes('//' + ;
								ccMEMBER_DATA_XML_ELEMENT)
							*** DougHennig 05/23/2007: save the current work area and select the csrMembers cursor
							lnSelect = Select()
							Select (tcCursorName)
							For Each loNode In loNodes
								lcPEM     = Upper(loNode.getAttribute('name'))
								lcDisplay = Nvl(loNode.getAttribute('display'), '')
								lcScript  = Nvl(loNode.getAttribute('script'), '')
								Try
									loPEM           = .oMemberDataCollection.Item(lcPEM)
									loPEM.Display   = Evl(loPEM.Display, ;
										lcDisplay)
									loPEM.Name      = Evl(loPEM.Display, ;
										loPEM.Name)
									loPEMMemberData = Createobject('MemberDataObject')
									With loPEMMemberData
										.HasMemberData = .T.
										.Location      = lcClass + ccLOC_OF + ;
											justfname(lcLibrary)
										.Display       = lcDisplay
										.Favorites     = Nvl(loNode.getAttribute('favorites'), ;
											'')
										.Override      = Nvl(loNode.getAttribute('override'), ;
											'')
										.Script        = Nvl(loNode.getAttribute('script'), ;
											'')
										.OriginalDisplay   = .Display
										.OriginalFavorites = .Favorites
										.OriginalOverride  = .Override
										.OriginalScript    = .Script
										.CustomAttributes  = This.GetCustomAttributesForMember(lcPEM, ;
											loDOM)
									Endwith

									* Store the MemberData object in the ParentMemberData property of the
									* appropriate object.

									Do While Vartype(loPEM.ParentMemberData) = 'O'
										loPEM = loPEM.ParentMemberData
									Enddo While Vartype(loPEM.ParentMemberData) = 'O'
									loPEM.ParentMemberData = loPEMMemberData
									*** DougHennig 05/23/2007: update the csrMembers cursor
									Locate For Upper(Trim(cName)) == Upper(lcPEM)
									*** DougHennig 11/21/2008: find the correct name for arrays
									* If Not Found() And Type('.oObject.' + lcPEM, 1) = 'A'
									If Not Found() And "U" # Type("Alen(.oObject." + lcPEM + ")")
										Locate For Upper(Trim(cName)) = Upper(lcPEM) + '['
										If Found()
											lcDisplay = Strtran(cName, lcPEM, lcDisplay, -1, -1, 1)
										Endif Found()
									Endif Not Found() ...
									*** DougHennig 11/21/2008: end of new code
									****************************************************************
									* Modified 10/25/2008 by Jim Nelson
									* Uses EVL for unexplained case where lcDisplay is blank
									Replace cName With Evl(lcDisplay, cName), ;
										nFavorites With 1 + IIf(loPEMMemberData.Favorites = ccXML_TRUE, 1, 0) + IIf(nFavorites >= 3, 2, 0)
									If not Empty (lcScript)
										? 1 / ''
									EndIf 
								Catch
								Endtry
							Next loNode
							*** DougHennig 05/23/2007: reselect the former work area
							Select (lnSelect)
						Endif loDOM.parseError.errorCode = 0
					Endif Not Empty(.cMemberData)
				Endif lnPos > 0
				lcClass   = Lower(Class)

				If Not Empty(CLASSLOC)
					lcLibrary = Fullpath(CLASSLOC, Addbs(Justpath(lcLibrary)))
					If Not File( m.lcLibrary )
						lcLibrary = Fullpath( CLASSLOC )
					Endif
				Else
					lcLibrary = ''
				Endif Not Empty(CLASSLOC)

				Use
			Enddo While Not Empty(lcLibrary)

			*** JRN 06/01/09 : Per Doug H
			*!*	* Handle MemberData in the containership hierarchy.

			*!*	loParent = Iif(Type('.oObject.Parent.Name') = 'C', ;
			*!*		.oObject.Parent, .Null.)
			*!*	Do While Vartype(loParent) = 'O'
			*!*		If Pemstatus(loParent, '_MemberData', 5)
			*!*			loDOM.LoadXML(loParent._MemberData)
			*!*			If loDOM.parseError.errorCode = 0
			*!*				loNodes = loDOM.selectNodes('//' + ;
			*!*					ccMEMBER_DATA_XML_ELEMENT)
			*!*				For Each loNode In loNodes
			*!*					lcPEM     = Upper(loNode.getAttribute('name'))
			*!*					lcDisplay = Nvl(loNode.getAttribute('display'), '')
			*!*					Try
			*!*						loPEM = .oMemberDataCollection.Item(lcPEM)
			*!*						.SetupContainershipHierarchy(loPEM)
			*!*						loPEM.Display   = Evl(loPEM.Display, ;
			*!*							lcDisplay)
			*!*						loPEM.Name      = Evl(loPEM.Display, ;
			*!*							loPEM.Name)
			*!*						loPEMMemberData = .FindMemberDataForContainer(loPEM, ;
			*!*							loParent.Name)
			*!*						With loPEMMemberData
			*!*							.HasMemberData = .T.
			*!*							.Location      = loParent.Name
			*!*							.Display       = lcDisplay
			*!*							.Favorites     = Nvl(loNode.getAttribute('favorites'), ;
			*!*								'')
			*!*							.Override      = Nvl(loNode.getAttribute('override'), ;
			*!*								'')
			*!*							.Script        = Nvl(loNode.getAttribute('script'), ;
			*!*								'')
			*!*							.OriginalDisplay   = .Display
			*!*							.OriginalFavorites = .Favorites
			*!*							.OriginalOverride  = .Override
			*!*							.OriginalScript    = .Script
			*!*							.CustomAttributes  = This.GetCustomAttributesForMember(lcPEM, ;
			*!*								loDOM)
			*!*						Endwith
			*!*					Catch
			*!*					Endtry
			*!*				Next loNode
			*!*			Endif loDOM.parseError.errorCode = 0
			*!*		Endif Pemstatus(loParent, '_MemberData', 5)
			*!*		Try
			*!*			loParent = loParent.Parent
			*!*		Catch
			*!*			loParent = .Null.
			*!*		Endtry
			*!*	Enddo While Vartype(loParent) = 'O'
		Endwith

		* Close the cursors.

		Use In _PROPERTIES
		Use In Global
	Endfunc

Enddefine

* A class to represent a member. Each *MemberData property contains a
* MemberDataObject object with the MemberData attributes from the specified
* location.

Define Class MemberObject As Custom
	ClassMemberData     = .Null.
	ParentMemberData    = .Null.
	ContainerMemberData = .Null.
	GlobalMemberData    = .Null.
	NativePEM           = .F.
	Type                = ''
	Display             = ''
	Script	            = ''
Enddefine

* A class to hold attribute values for a member.

Define Class MemberDataObject As Custom
	HasMemberData     = .F.
	Inherited         = .F.
	Display           = ''
	Location          = ''
	OriginalDisplay   = ''
	Favorites         = ''
	OriginalFavorites = ''
	Override          = ''
	OriginalOverride  = ''
	Script            = ''
	OriginalScript    = ''
	ParentMemberData  = .Null.
	CustomAttributes  = .Null.
Enddefine
