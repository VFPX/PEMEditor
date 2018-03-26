*==============================================================================
* Program:			MDCleaner.PRG
* Purpose:			Cleans up _MemberData for the selected object
* Author:			Doug Hennig
* Last Revision:	04/02/2009
* Parameters:		tlClass - .T. if the object is a class, .F. if it's a form
* Returns:			.T.
* Environment in:	a class is open in the Class Designer or a form is open in
*						the Form Designer
* Environment out:	_MemberData for the object has been cleaned up:
*						- records for members which don't exist are removed
*						- records for inherited members are removed
*						- the type attribute is removed
*						- the favorites attribute is removed if the member
*						  isn't a favorite
*						- CR and tabs are stripped out
*==============================================================================

Lparameters tlClass, tlMemberDataKeepType
Local laObjects[1], ;
	loObject, ;
	loXML, ;
	lcMessage, ;
	laMembers[1], ;
	loRoot, ;
	loNodes, ;
	loNode, ;
	lcName, ;
	lnMember, ;
	lcFavorites, ;
	lcXML,			;
	lcResult

Local loParentXML, loParentNode, lcAttrName, lcAttrValue, lcParentAttrValue, lnI, lnSelect

* We need some constants in MemberDataEditor.H.

#include MemberDataEditor.H

* Localization constants.

#Define ccTITLE					'MemberData Cleaner'
#Define ccMSG_NO_OBJECT			'Please open a class in the Class Designer or form in the Form Designer, then run this program again.'
#Define ccMSG_NO_MEMBERDATA		'The selected object has no member data.'
#Define ccMSG_XML_LOAD_ERROR	'An error occurred loading the XML in _MemberData at position <Insert1> of line <Insert2>. The error is:'

* Get a reference to the form or class. Bug out if there isn't one.

If Aselobj(laObjects, 1) = 0
	Messagebox(ccMSG_NO_OBJECT, 16, ccTITLE)
	Return
Endif Aselobj(laObjects, 1) = 0
loObject = PEME_FindTopMostParent(laObjects[1], tlClass)

* Bug out if the object doesn't have a _MemberData property or it's empty.

If Not Pemstatus(loObject, '_MemberData', 5) Or Empty(loObject._MemberData)
	Messagebox(ccMSG_NO_MEMBERDATA, 16, ccTITLE)
	Return
Endif Not Pemstatus(loObject, '_MemberData', 5) ...

* Load the XML in _MemberData. Exit if we can't.

lnSelect = Select()
loXML = Createobject(ccXML_DOM_CLASS)
loXML.Async = .F.
loXML.LoadXML(loObject._MemberData)
If loXML.parseError.errorCode <> 0
	lcMessage = Strtran(ccMSG_XML_LOAD_ERROR, '<Insert1>', ;
		transform(loXML.parseError.linepos))
	lcMessage = Strtran(lcMessage, '<Insert2>', ;
		transform(loXML.parseError.Line))
	Messagebox(lcMessage + ccCR + ccCR + loXML.parseError.reason, 16, ccTITLE)
	Return
Endif loXML.parseError.errorCode <> 0

*** JRN 06/30/09 : Get the inherited members of the object.
loParentXML = GetParentMemberData(loObject, tlClass)

Amembers(laMembers, loObject, 1)

* Process each node in the XML. If the specified member doesn't exist (meaning
* it was likely removed and its _MemberData entry wasn't purged) or is
* inherited (meaning this is a subclass so we don't need the _MemberData entry
* for this member), remove the node from the XML. If we're keeping it, remove
* the "type" attribute and if the member isn't a favorite, the "favorites"
* attribute.

loRoot  = loXML.selectSingleNode('/' + ccXML_ROOT_NODE)
loNodes = loXML.selectNodes('/' + ccXML_ROOT_NODE + '/' + ccMEMBER_DATA_XML_ELEMENT)
For Each loNode In loNodes
	lcName   = loNode.getAttribute('name')
	lnMember = Ascan(laMembers, lcName, -1, -1, 1, 15)
	If lnMember = 0 && if no longer a member, this is a dead branch
		loRoot.removeChild(loNode)
	Else
		*** JRN 06/30/2009 : remove entries that duplicate inherited values
		If Vartype(loParentXML) = 'O'
			loParentNode = loParentXML.selectSingleNode('//memberdata[@name = "' + Lower(lcName ) + '"]')
			If Vartype(loParentNode) = 'O'
				For lnI = loNode.Attributes.Length - 1 To 1 Step -1
					lcAttrName = loNode.Attributes(lnI).nodename
					lcAttrValue = loNode.getAttribute(lcAttrName)
					lcParentAttrValue = loParentNode.getAttribute(lcAttrName)
					If lcAttrValue == lcParentAttrValue
						loNode.removeAttribute(lcAttrName )
					Endif
				Next lnI
			Endif
		Endif

		*** JRN 06/01/09 : remove attribute 'Type' unless it's a property AND we're keeping it
		If tlMemberDataKeepType And (								;
				Pemstatus(loObject, lcName, 3) = 'Property'			;
				or Not Empty(Nvl(loNode.getAttribute('script'), ''));
				)
			loNode.setAttribute('type', 'property')
		Else
			loNode.removeAttribute('type')  && removed 05/20/2009 ... JRN
		Endif

		If Nvl (loNode.getAttribute('display'), Lower (lcName)) == Lower (lcName)
			loNode.removeAttribute('display')
		Endif

		If Not Nvl (loNode.getAttribute('favorites'), ccXML_TRUE) == ccXML_TRUE
			loNode.removeAttribute('favorites')
		Endif

		If Empty (Nvl (loNode.getAttribute('script'), 'not empty'))
			loNode.removeAttribute('script')
		Endif

		If 1 = loNode.Attributes.Length
			loRoot.removeChild(loNode)
		Endif

		laMembers (lnMember, 1) = '?' && if there are duplicate entries, this will mean only the first will be kept

	Endif lnMember = 0 ...
Next loNode

* Write the XML back to _MemberData. Note that if there's nothing left but the
* root node, we'll blank it altogether.

If loRoot.hasChildNodes
	lcXML = loXML.XML
	lcXML = Strtran(lcXML, ccCRLF)
	lcXML = Strtran(lcXML, ccCR)
	lcXML = Strtran(lcXML, ccTAB)
	lcResult = lcXML
Else
	lcResult = ''
Endif loRoot.hasChildNodes

Select (lnSelect)
Return lcResult



Procedure GetParentMemberData (toObject, tlClass)

	Local laMemberDatas(1)
	Local loChildXML As ccXML_DOM_CLASS
	Local loNode As Object
	Local loNodes As Object
	Local loParentXML As ccXML_DOM_CLASS
	Local loParentNode As Object
	Local lcAttrName, lcAttrValue, lcClass, lcLibrary, lcMemberData, lcNodeName, lcPrompt, lnI, lnJ, lnLen
	Local lnNFound, lnPos, loRoot, lnSelect

	With toObject

		lcClass   = Iif(tlClass, Lower(.ParentClass), Lower(.Class))
		lcLibrary = .ClassLibrary

		lnNFound = 0
		lnSelect = Select()
		Do While Not Empty(lcLibrary)
			If Not File (lcLibrary)
				lcPrompt = Strtran (ccLOC_File_Not_Found, '<FileName>', Justfname(lcLibrary))
				Messagebox (lcPrompt, 0, ccLOC_MEM_CAP_TITLE)
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
					lnNFound = lnNFound + 1
					Dimension laMemberDatas (lnNFound)
					laMemberDatas (lnNFound) = lcMemberData
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

		If lnNFound = 0
			Select (lnSelect)
			Return ''
		Endif

		loParentXML = Createobject(ccXML_DOM_CLASS)
		loParentXML.Async = .F.
		loParentXML.LoadXML(laMemberDatas (lnNFound))
		If loParentXML.parseError.errorCode # 0
			Select (lnSelect)
			Return ''
		Endif

		For lnI = lnNFound - 1 To 1 Step -1
			loChildXML = Createobject(ccXML_DOM_CLASS)
			loChildXML.Async = .F.
			loChildXML.LoadXML(laMemberDatas (lnI))
			If loChildXML.parseError.errorCode = 0
				loNodes = loChildXML.selectNodes('/' + ccXML_ROOT_NODE + '/' + ccMEMBER_DATA_XML_ELEMENT)
				For Each loNode In loNodes
					lcNodeName = loNode.getAttribute("name")
					loParentNode = loParentXML.selectSingleNode('//memberdata[@name = "' + Lower(lcNodeName ) + '"]')

					If Vartype(loParentNode) # 'O'
						loParentNode = loParentXML.createElement(ccMEMBER_DATA_XML_ELEMENT)
						loParentNode.setAttribute('name', lcNodeName)
						loRoot = loParentXML.selectSingleNode('/' + ccXML_ROOT_NODE)
						loRoot.appendChild(loParentNode)
					Endif

					For lnJ = 1 To loNode.Attributes.Length - 1
						lcAttrName = loNode.Attributes(lnJ).nodename
						lcAttrValue = loNode.Attributes(lnJ).nodevalue
						loParentNode.setAttribute(lcAttrName, lcAttrValue)
					Next lnJ

				Endfor
			Endif

		Next lnI

		Select (lnSelect)
		Return loParentXML
	Endwith

Endproc
