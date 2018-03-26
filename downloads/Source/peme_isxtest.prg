#Define cnMaxColumn2Width 300

#Define NOMATCHFOUND ' <no matches>'
* Characters which terminate popup and show current selection
#DEFINE POPUP_TERMINATION_CHARS " )[.=+-*/%,]"
* Characters which also have a space before them
#DEFINE EXTRA_SPACE_LIST		'+-*$/%<>=!^;'

#Define ccTab chr[9]

*========================================================================================
*
* Version: 2010-02Feb-20
*
*========================================================================================
*
* This program implements partial IntelliSense in VFP 6-9. To enable 
* IntelliSenseX, simply execute this program at any time when using
* Visual FoxPro or put it into your startup program.
*
* To configure ISX please see the section just below the comment block.
*
* To stop IntelliSenseX run this program again and pass "QUIT" as a
* parameter. Alternatively, you can simply remove the ON KEY LABEL
* macros for the ALT+I and the "." key.
*
* Currently only IntelliSense for variable names is implemented. This
* means that whenever you enter "m." in a MODIFY COMMAND window or
* in a Method edit window, you get a list of all variables declared
* in the current procedure. ISX doesn't parse the entire sourcecode 
* for memory variables, but only the current procedure or method and
* only those variables listed in a LOCAL, PRIVATE, PUBLIC, LPARAMETER
* and PARAMETER statement.  ALT+I can be used to trigger this list.
*
* ALT+RIGHTARROW triggers a universal autocomplete function that tries to determine the
* type of what you have entered and offers a list of all possible values.
*
* Please note that I haven't written this program as an excercise for
* good coding styles <g>, rather as an excercise to see if 
* IntelliSense is possible within Visual FoxPro itself. Therefore
* you won't find the Assertions you would otherwise find in my code.
*
*========================================================================================
*
* Acknowledgements
*
* Thanks to George Tasker for his really helpful documentation on the
* FoxTools.Fll. You can download his ToolHelp.Hlp file from the
* UniversalThread and the CompuServe MSDEVAPP forum. George also made
* some suggestions to improve this program.
*
* Also thanks to Ken Levy, who couldn't implement an inline Intelli-
* Sense feature in his SuperCls and thereby convinced me that there 
* must be a way to do it, even only for the purpose of doing 
* something that Ken Levy couldn't do. <bg>
*
* Thanks to all the folks that posted me bug reports, especially
* Frank Cazabon. Thanks to Gerry Hughes for correcting the typos in 
* my comments.
*
* Louis D. Zelus added a nifty feature to my version to make ISX 
* even more useful. Thanks for that! The code based on his work is
* marked with "LDZ:".
*
* Sietse Wijnkler added a lot of new cool features: He added the
* ability to distinguish different types that all are triggered by
* a period and the code to display variables, object properties and
* field names. Code based on his work is marked with "SW:".
*
* Jürgen "wOOdy" Wondzinski pointed out that special characters like
* "ö" are valid variable names and IsAlpha() returns .T. for them.
* Therefore any of these characters is detected by ISX, as well.
*
* Tamar E. Granor and Peter Steinke, both requested the list DEFINE 
* features which is why I finally added it.
*
* Thanks to Eddy Maue for his contributions:
*
*   Ce qu'ile fait de plus maintenant
*    -  Alt-Q pour arrêter Isx
*    -  Alt-Q pour redemarrer Isx
*    - Ouvre automatiquements :
*            -Les tables présentes dans les répertoires courants et de recherches
*             (set path to)
*            -Les vues présentes dans le projet actif
*            -Les query présents dans les répertoires courants et de recherches
*             (set path to)
*              Petit point à ne pas négliger. Le curseur produit par le fichier 
*              MyQuery.qpr doit être du même nom que le fichier
*
* In English:
* 
*    - ALT+Q enables/disables ISX
*    - files are opened automatically:
*        - tables available in the current directory or the search path (SET PATH TO)
*        - Views available in the current project
*        - Queries available in the current directory or the search path (SET PATH TO)
*          Minor, but important restriction: The cursor created by the query program
*          must have the same alias as the filename.

* Mike Yearwood added supported for maximized editing windows which caused a lot
* of flickering everytime the popup came up.
* 
* Thanks to all those who pointed out bugs in ISX's releases: 
* 
*  - Nina Schwanzer
*  - Del Lee
*  - Pamela Thalacker
*  - Christophe Chenavier
*  - Aragorn Rockstroh
*  - Claude Hebert 
*  - Jens Kippnich
*  - Stefan Wübbe
*
*========================================================================================
*
* This program has been written in 1999-2005 by Christof Wollenhaupt
* and is placed into Public Domain. You can use the entire
* code or parts of it as you like in any private or commercial
* application. None of the contributors to this programm can be hold 
* liable for any damage or problems, using this program may cause.
*
* If you added a new feature, please let me know. If you want I add
* your feature to my master copy of ISX to let others use your 
* feature, as well. Please note that since the entire program is 
* placed into Public Domain, this places your code into Public 
* Domain, as well. Of course, your contributions are acknlowdeged in
* the comment at the beginning of this file.
*
*========================================================================================
*
* Known problems:
*
* - So far ISX has not been tested with different Display appearance
*   settings, like wider scrollbars or form borders, large fonts and
*   the like. Some values are hardcoded and might be wrong for non-
*   standard Windows settings.
*
* - When you enter a period into a textbox, the cursor is set to the first character of
*   the textbox and then the period entered. If SelectOnEntry is true, everything is
*   replaced by the period. This is caused by a bug in VFP that makes all ON KEY LABEL
*   behave this way. You can disable this behavior by commenting out the lines starting 
*   with "ON KEY LABEL .". In this case, you must use ALT+I or ALT+RIGHTARROW do expand
*   the variable.
*
*========================================================================================


*========================================================================================
* Configuration. 
*
* Over the time I got many enhanced versions of ISX, many of which include new hotkeys.
* To give everyone control over the hotkey assignment and to disable/enable particular
* features, I added the following configuration section. By commenting out a #DEFINE, you
* disable a particular feature. Changing the value changes the hotkey.
*
*========================================================================================


*========================================================================================
* Main program
*========================================================================================
Lparameters toEditBox, toThisform, toThis, tcProviders, tcInvocation, tlForceThor

	AutoComplete(m.toEditBox, m.toThisform, m.toThis, m.tcProviders, m.tcInvocation, m.tlForceThor)
	
Return 

*!* * Removed 8/1/2012 / JRN
*!* Lparameters tcParam, tcParam2, tlParam3

*!* 	*!* * Removed 7/27/2012 / JRN ... moved to surround the call to the popup form
*!* 	*!* so that if errors occur during the collection of the names in the popup
*!* 	*!* the ON KEY LABEL definitions are not lost
*!* 	*!* Push Key Clear
*!* 	AutoComplete( m.tcParam, m.tcParam2, tlParam3 )

*!* Return


*========================================================================================
* Provides a generic autocomplete function. AutoComplete checks all content providers 
* if they have something to add to the global list and displays the list as a popup
*========================================================================================
 Procedure AutoComplete
Lparameters toEditBox, toThisform, toThis, tcProviders, tcInvocation, tlForceThor
	
	*--------------------------------------------------------------------------------------
	* The list of providers can be limited. This speeds up program execution if one knows 
	* from the context that only few content providers actually fit.
	*--------------------------------------------------------------------------------------
	
	Local lcProviders
	If Empty(m.tcProviders)
		m.lcProviders = "VAR,DEFINE,TABLE,OBJ,THOR"
	Else
		m.lcProviders = Upper(m.tcProviders)
	EndIf 
	
	*-----------------------------------------------------------------
	* Make sure, FoxTools.Fll is loaded.
	*-----------------------------------------------------------------
	If not "FOXTOOLS.FLL" $ Upper(Set("Library"))
		Set Library to (Home()+"FoxTools.Fll") Additive
	Endif

	*!* * Removed 8/1/2012 / JRN
	*!* *-----------------------------------------------------------------
	*!* * Get the current window and verify that it is a valid window.
	*!* *-----------------------------------------------------------------
	*!* Local lnWHandle
	*!* lnWHandle = GetCurrentWindow()
	*!* If lnWHandle == 0
	*!* 	If not Empty(m.tcInvocation)
	*!* 		Clear TypeAhead
	*!* 		Keyboard "{Mouse}{Mouse}"+m.tcInvocation Plain
	*!* 	Endif
	*!* 	Return
	*!* Endif

	*!* *-----------------------------------------------------------------
	*!* * Verify that the current window is indeed an edit window.
	*!* *-----------------------------------------------------------------
	*!* Local lnEditSource 
	*!* lnEditSource = GetEditSource(m.lnWHandle)
	*!* If not InList( m.lnEditSource, 1, 8, 10, 12 )
	*!* 	If not Empty(m.tcInvocation)
	*!* 		Clear TypeAhead
	*!* 		Keyboard "{Mouse}{Mouse}"+m.tcInvocation Plain
	*!* 	Endif
	*!* 	Return
	*!* EndIf

	*--------------------------------------------------------------------------------------
	* Fill an object with details about the current context. We determine what the user 
	* has entered so far and what's left from that Position. 
	*--------------------------------------------------------------------------------------
	Local loISX
	m.loISX = CreateObject("Relation")
	m.loISX.AddProperty("aList[1]")
	m.loISX.AddProperty("nCount",0)
	m.loISX.AddProperty("cTextLeft",GetLineLeftFromCursor(m.toEditBox))
	If 'O' = Vartype(m.toThisForm) and Lower(m.toThisForm.BaseClass) = 'form'
		m.loISX.AddProperty("oThisForm",m.toThisForm)
	EndIf 
	m.loISX.AddProperty("oThis",m.toThis)
	m.loISX.AddProperty("oEditBox",m.toEditBox)
	m.loISX.AddProperty("cName","")
	m.loISX.AddProperty("cEntity","")
	m.loISX.AddProperty("cInvocation",m.tcInvocation)
	m.loISX.AddProperty("lForceThor",m.tlForceThor)
	m.loISX.AddProperty("lSortList",.T.)
	m.loISX.AddProperty("ISXOptions",GetISXOptions())
	
	*--------------------------------------------------------------------------------------
	* Determine the part of the name that has been entered so far. This code has been 
	* kindly provided by Louis D. Zelus.
	*--------------------------------------------------------------------------------------
	Local lcLine, lcChar
	If Empty(m.tcInvocation)
		Do While Len(m.loISX.cTextLeft) > 0
			m.lcChar = Right( m.loISX.cTextLeft, 1 )
			If IsAlpha(m.lcChar) or IsDigit(m.lcChar) or m.lcChar == "_"
				m.loISX.cTextLeft = Left( m.loISX.cTextLeft, Len(m.loISX.cTextLeft)-1 )
				m.loISX.cName = m.lcChar + m.loISX.cName
			Else
				Exit
			Endif
		Enddo
	EndIf 

	*--------------------------------------------------------------------------------------
	* Determines the name of the entity. This code is courtesy of Sietse Wijnkler. 
	*--------------------------------------------------------------------------------------
	Do While Len(m.loISX.cTextLeft) > 0
		m.lcChar = Right( m.loISX.cTextLeft, 1 )
		If IsAlpha(m.lcChar) or IsDigit(m.lcChar) or m.lcChar == "_" or m.lcChar == "."
			m.loISX.cTextLeft = Left( m.loISX.cTextLeft, Len(m.loISX.cTextLeft)-1 )
			m.loISX.cEntity = m.lcChar + m.loISX.cEntity
		Else
			Exit
		Endif
	EndDo
	If Right(m.loISX.cEntity,1) == "."
		m.loISX.cEntity = Left( m.loISX.cEntity, Len(m.loISX.cEntity)-1 )
	EndIf 
	
	*--------------------------------------------------------------------------------------
	* This array lists all the providers
	*--------------------------------------------------------------------------------------
	Local loReadProviders 
	m.loReadProviders = CreateObject('ReadProviders')
	m.loReadProviders.Run(m.loISX, m.lcProviders)

	*!* * Removed 8/2/2012 / JRN
	*!* Local laProvider[5,2]
	*!* laProvider = ""
	*!* laProvider[1,1] = "VAR"
	*!* laProvider[1,2] = "CP_Variables"
	*!* laProvider[2,1] = "DEFINE"
	*!* laProvider[2,2] = "CP_Defines"
	*!* laProvider[3,1] = "TABLE"
	*!* laProvider[3,2] = "CP_Tables"
	*!* laProvider[4,1] = "OBJ"
	*!* laProvider[4,2] = "CP_Objects"
	*!* laProvider[5,1] = "THOR"
	*!* laProvider[5,2] = "CP_THOROBJECTS"

	*!* *--------------------------------------------------------------------------------------
	*!* * Get data from each provider and merge it into the list
	*!* *--------------------------------------------------------------------------------------
	*!* Local laAll[1], lnAll, laRequest[1], lnRequest, lnProvider, lnDataSession
	*!* lnDataSession = Set("Datasession")
	*!* Set Datasession to 1
	*!* lnAll = 0
	*!* For lnRequest=1 to ALines(laRequest,Chrtran(m.lcProviders,",",Chr(13)+Chr(10)),.T.) 
	*!* 	For lnProvider=1 to Alen(laProvider,1)
	*!* 		If Upper(Alltrim(laRequest[m.lnRequest])) == laProvider[m.lnProvider,1] 
	*!* 			loISX.nCount = 0
	*!* 			Dimension loISX.aList[1]
	*!* 			loISX.aList = ""
	*!* 			&laProvider[m.lnProvider,2](m.loISX)
	*!* 			If m.loISX.nCount > 0
	*!* 				Dimension laAll[m.lnAll+m.loISX.nCount]
	*!* 				Acopy(m.loISX.aList,laAll,1,m.loISX.nCount, m.lnAll+1)
	*!* 				lnAll = m.lnAll + m.loISX.nCount
	*!* 			EndIf 
	*!* 		EndIf 
	*!* 	EndFor 
	*!* 	If m.lnAll = 0
	*!* 		Exit
	*!* 	EndIf 
	*!* EndFor 
	*!* Set Datasession to (lnDataSession)

	*--------------------------------------------------------------------------------------
	* If there's anything in the list, display the popup
	*--------------------------------------------------------------------------------------
	*!* * Removed 8/1/2012 / JRN
	*!* If m.lnAll == 0
	*!* 	If not Empty(m.tcInvocation)
	*!* 		Clear TypeAhead
	*!* 		Keyboard "{Mouse}{Mouse}"+m.tcInvocation Plain
	*!* 	Endif
	*!* Else 
	*!* 	If not Empty(m.tcInvocation)
	*!* 		InsertText( m.lnWHandle, m.tcInvocation )
	*!* 	EndIf
	If m.loISX.nCount # 0 
		Push Key Clear
		DisplayPopup(m.loISX)
		Pop Key 
	EndIf
	
EndProc


*========================================================================================
* Determines all include files that fit in the current situation and adds them to the
* list.
*========================================================================================
 Procedure CP_Defines
Lparameters toISX

	Local loFile
	If Type("_VFP.ActiveProject") == "O"
		For each m.loFile in _VFP.ActiveProject.Files
			If Upper(JustExt(m.loFile.Name)) == "H"
				ReadDefines(m.toISX,m.loFile.Name)
			EndIf 
		EndFor 
	Else
		ReadDefines(m.toISX,Home()+"FoxPro.H")
	EndIf 

EndProc 


*========================================================================================
* Adds all constants from an include file to the array.
*========================================================================================
 Procedure ReadDefines
LParameter m.toISX, m.tcFile

	*--------------------------------------------------------------------------------------
	* File must exist.
	*--------------------------------------------------------------------------------------
	If not File(m.tcFile)
		Return 
	EndIf 
	
	*--------------------------------------------------------------------------------------
	* To increase performance, we cache files if possible.
	*--------------------------------------------------------------------------------------
	Local laDefine[1], lnItem, lnCount
	If not IsInCache( "DEFINE", m.toISX, m.tcFile )
		If Version(4) >= "07.00"
			m.lnCount = AProcInfo(laDefine,m.tcFile)
		Else
			m.lnCount = X6_AProcInfo(@m.laDefine,m.tcFile)
		EndIf 
		For m.lnItem=1 to m.lnCount
			If m.laDefine[m.lnItem,3] == "Define"
				m.toISX.nCount = m.toISX.nCount + 1
				Dimension m.toISX.aList[m.toISX.nCount]
				m.toISX.aList[m.toISX.nCount] = m.laDefine[m.lnItem,1]
			EndIf 
		EndFor 
		AddToCache( "DEFINE", m.toISX, m.tcFile )
	EndIf 

EndProc


*========================================================================================
* The cache is an array in _SCREEN that holds the name of the file, the time stamp, the
* provider ID and the contents of the array.
*========================================================================================
 Procedure IsInCache
LParameter m.tcProvider, m.toISX, m.tcFile

	If Type("_Screen.ISXCache[1,1]") == "U"
		Return .F.
	EndIf

	Local lnLine
	If Version(4) >= "07.00"
		m.lnLine = Ascan( _Screen.ISXCache, m.tcFile+"?"+m.tcProvider, -1, -1, 1, 1+2+4+8 )
	Else
		Local lnCurLine
		m.lnLine = 0
		For m.lnCurLine=1 to Alen(_Screen.ISXCache,1)
			If Type(_Screen.ISXCache[m.lnCurLine]) == "C"
				If Upper(m.tcFile+"?"+m.tcProvider) == Upper(_Screen.ISXCache[m.lnCurLine])
					m.lnLine = m.lnCurLine
					Exit
				EndIf 
			EndIf 
		EndFor 
	EndIf 
	If m.lnLine == 0
		Return .F.
	EndIf 
	
	If Fdate(m.tcFile,1) # _Screen.ISXCache[m.lnLine,2]
		Return .F.
	EndIf
	
	m.toISX.nCount = _Screen.ISXCache[m.lnLine,3]
	ALines( toISX.aList, _Screen.ISXCache[m.lnLine,4] )

Return .T.


*========================================================================================
* Adds the current entry to the cache.
*========================================================================================
 Procedure AddToCache
LParameter m.tcProvider, m.toISX, m.tcFile

	If Type("_Screen.ISXCache[1,1]") == "U"
		_Screen.AddProperty("ISXCache[1,4]")
	EndIf

	Local lnLine
	If Version(4) >= "07.00"
		m.lnLine = Ascan( _Screen.ISXCache, m.tcFile+"?"+m.tcProvider, -1, -1, 1, 1+2+4+8 )
	Else
		Local lnCurLine
		m.lnLine = 0
		For m.lnCurLine=1 to Alen(_Screen.ISXCache)
			If Upper(m.tcFile+"?"+m.tcProvider) == Upper(_Screen.ISXCache[m.lnCurLine])
				m.lnLine = m.lnCurLine
				Exit
			EndIf 
		EndFor 
	EndIf 
	If m.lnLine == 0
		m.lnLine = Alen(_Screen.ISXCache,1) + 1
		Dimension _Screen.ISXCache[m.lnLine,Alen(_Screen.ISXCache,2)]
	EndIf 	

	Local lnItem
	_Screen.ISXCache[m.lnLine,1] = m.tcFile+"?"+m.tcProvider
	_Screen.ISXCache[m.lnLine,2] = Fdate(m.tcFile,1)
	_Screen.ISXCache[m.lnLine,3] = m.toISX.nCount
	_Screen.ISXCache[m.lnLine,4] = ""
	For m.lnItem=1 to m.toISX.nCount
		_Screen.ISXCache[m.lnLine,4] = _Screen.ISXCache[m.lnLine,4] + ;
			m.toISX.aList[m.lnItem] + Chr(13)+Chr(10)
	EndFor 

EndProc
	

*====================================================================
* SW: Fills an array with all PEMs for the objectname typed in
* Returns the number of PEMs. The object has to exist to work
*====================================================================
*** JRN 8/1/2012 : Modified to look only for Thisform. and This.
* and also to only show objects or non-native properties

 Procedure CP_Objects
	Lparameters toISX

	Local lcEntity, lcName, lnCount, lnNewItems, loEntity, loPEMs, lnAdded
	m.lcEntity = 'toISX.o' + m.toISX.cEntity
	If Type (m.lcEntity) = [O]
		*!* * Removed 8/1/2012 / JRN
		*!* If Version(4) >= '07.00'
		*!* 	If    Upper (lcEntity) == '_SCREEN' Or Upper (lcEntity) == '_VFP'		;
		*!* 			Or Upper (lcEntity) = '_VFP.'
		*!* 		Return
		*!* 	Endif
		*!* Endif
		m.loEntity = Evaluate (m.lcEntity)
		m.loPEMs	 = _oPEMEditor.oTools.GetPEMList (m.loEntity, 'P')

		m.toISX.nCount = m.loPEMs.Count
		Dimension m.toISX.aList[m.toISX.nCount]
		m.lnAdded = 0
		For m.lnCount = 1 To m.toISX.nCount
			m.lcName = m.loPEMs.Item[m.lnCount].cName
			If '[' $ m.lcName
				m.lcName = Left(m.lcName, At('[', m.lcName) - 1) 
			EndIf 
			m.lcName = Alltrim(m.lcName)
			If PemStatus(m.loEntity, m.lcName, 4)
				m.lnAdded = m.lnAdded + 1
				m.toISX.aList[m.lnAdded] = m.lcName
			Endif
		Next
		m.toISX.nCount = m.lnAdded
		Dimension m.toISX.aList[Max(1,m.lnAdded)]
	Endif
Endproc


*====================================================================
*** JRN 7/26/2012 : Calls a Thor PROC that can interpret the supplied
* name, which is some presumed to be some kind of object reference
* to a table, and adds the field names from the table
* Returns the number of PEMs. The object has to exist to work
*====================================================================
 Procedure CP_ThorObjects
Lparameters toISX
	
	Local lnCount, loCollection
	If Type('_Screen.cThorDispatcher') = 'C'
		m.loCollection = Execscript(_Screen.cThorDispatcher, 'Thor_Proc_GetFieldsFromObjectName', m.toISX)
		Do Case
			Case 'C' = Type("loCollection.aList[1]")
				Dimension m.toISX.aList[1]
				Acopy(loCollection.aList, m.toISX.aList)
				m.toISX.nCount = Alen(m.loCollection.aList, 1)

			Case 'O' = Vartype(m.loCollection)
				m.toISX.lSortList = .F.
				m.toISX.nCount = m.loCollection.Count
				Dimension m.toISX.aList[m.loCollection.Count]
				For m.lnCount = 1 To m.loCollection.Count
					m.toISX.aList[m.lnCount] = m.loCollection.Item[m.lnCount]
				EndFor 
				Return m.loCollection.Count

			Otherwise 
				Return 0
		Endcase
	Else
		Return 0
	Endif
		
EndProc 


*====================================================================
* SW: Fills an array with all Fields for the cursor typed in.
* Returns the number of Fields. The cursor has to be open to work
*====================================================================
 Procedure CP_Tables
Lparameters toISX

	Local laFields[1], lcName, lnCount, laFields
	m.lcName = JustStem(m.toISX.cEntity)
	* November 11, 2004 Modified by Eddy Maue 
	*** JRN 7/26/2012 : added call to OpenThorTable, which uses standard
	* Thor proc to open tables (beyond that offered by OpenTable)
	If Occurs(".",m.toISX.cEntity)==0 And !"m." == LOWER(m.toISX.cEntity+".") AND ;
         IIF(Used(m.lcName),.t.,;
         IIF(File(m.lcName+".dbf"),OpenTable(m.lcName),;
         IIF(File(m.lcName+".qpr"),ExecQuery(m.lcName),;
         IIF(OpenView(m.lcName), .T., OpenThorTable(m.lcName))))) 
 
		m.toISX.lSortList = .F.		
		m.toISX.nCount	= Afields(laFields, m.lcName)
		DIMENSION m.toISX.aList[m.toISX.nCount]
		For m.lnCount = 1 To m.toISX.nCount
			m.lcField = FixFieldNameCase(m.laFields[m.lnCount, 1], m.toISX.Isxoptions.Fieldnamescase)
			m.lcField = m.lcField + ccTab + m.laFields[m.lnCount, 2]
			Do Case
				Case m.laFields[m.lnCount, 2] = 'C'
					m.lcField = m.lcField + '(' + Transform(m.laFields[m.lnCount, 3]) + ')'
				Case m.laFields[m.lnCount, 2] = 'N' And m.laFields[m.lnCount, 4] = 0
					m.lcField = m.lcField + '(' + Transform(m.laFields[m.lnCount, 3]) + ')'
				Case m.laFields[m.lnCount, 2] = 'N'
					m.lcField = m.lcField + '(' + Transform(m.laFields[m.lnCount, 3]) + ',' + Transform(m.laFields[m.lnCount, 4]) + ')'
				Case m.laFields[m.lnCount, 2] = 'B'
					m.lcField = m.lcField + '(' + Transform(m.laFields[m.lnCount, 4]) + ')'
			Endcase

			m.toISX.aList[m.lnCount] = m.lcField
		NEXT
	ENDIF

EndProc 


 Procedure FixFieldNameCase (tcName, tnFieldNamesCase)
	Local lcField
	m.lcField = m.tcName
	Do Case
		Case m.tnFieldNamesCase = 1
			m.lcField = Lower (m.lcField)
		Case m.tnFieldNamesCase = 2
			m.lcField = Upper (m.lcField)
		Case m.tnFieldNamesCase = 3
			m.lcField = Proper (m.lcField)
		Case m.tnFieldNamesCase = 4
			m.lcField = Lower (Left (m.lcField, 1)) + Substr (m.lcField, 2)
	Endcase
	Return m.lcField
Endproc

*====================================================================
* Open the table
* Eddy Maue
* November 11, 2004
*====================================================================
 Procedure OpenTable
   Lparameters lcName
   Use (m.lcName) In 0
   Return Used(m.lcName)
ENDPROC

*====================================================================
* Open a query
*====================================================================
* Eddy Maue
* November 11, 2004
*====================================================================
 Procedure ExecQuery
   Lparameters lcName
   Do (m.lcName+".qpr")
   Return Used(m.lcName)
ENDPROC

*====================================================================
* Open a view
*====================================================================
* Eddy Maue
* November 11, 2004
*====================================================================
 Procedure OpenView
   Lparameters lcName,lcSafety,lcConsol
   If Type("_vfp.ActiveProject")="U" .OR. EMPTY(DBC())
      Return .F.
   ENDIF
   m.lcSafety = "Set Safety "+Set("safety")
   Set Safety Off
   List Views To FILE _view.tmp NOCONSOLE 
   If ":"+ALLTRIM(Lower(m.lcName))+"(" $ STRTRAN(Lower(Filetostr("_view.tmp"))," ","")
      Use (m.lcName) In 0
   Endif
   &lcSafety
   RETURN USED(m.lcName)

*====================================================================
* Call standard Thor Proc to open a Thor table
* JRN: Jim Nelson
* July 25, 2012
*====================================================================
 Procedure OpenThorTable
   Lparameters lcName
   Do Case
   	   Case Empty(m.lcName)
   	   	    Return .F.
	   Case Type("_Screen.cThorDispatcher") = 'C'
	   	    m.lcResult = ExecScript(_Screen.cThorDispatcher, 'PEME_OpenTable', m.lcName)
	   	    Return (not IsNull(m.lcResult)) and not Empty(m.lcResult) 
	   Otherwise 
	   		Return .f.
   Endcase
ENDPROC


*========================================================================================
* Displays a popup with all the values from taList, lets the user incrementally approach
* the desired item and inserts it into the editor.
*========================================================================================
 Procedure DisplayPopup
	Lparameter m.toISX

	Local lcISXForm, lnDataSession, loEditBoxForm, loPopupForm
	If m.toISX.nCount > 0

		*-----------------------------------------------------------------
		* Determine the cursor position in _SCREEN coordinates
		*-----------------------------------------------------------------
		m.loEditBoxForm = m.toISX.oEditBox
		Do While m.loEditBoxForm.BaseClass # 'Form'
			m.loEditBoxForm = m.loEditBoxForm.Parent
		Enddo
		Do Case
			Case m.loEditBoxForm.Desktop
				m.lcISXForm = 'ISXFormDeskTop'

			Case m.loEditBoxForm.Dockable # 1
				m.lcISXForm = 'ISXForm'

			Case m.loEditBoxForm.Docked
				m.lcISXForm = 'ISXFormDeskTop'

			Otherwise
				m.lcISXForm = 'ISXFormDeskTop'

		Endcase

		m.loPopupForm = Createobject ( m.lcISXForm, m.toISX)
		If Vartype (m.loPopupForm) == 'O'
			m.loPopupForm.Show()
		Endif
		m.loPopupForm = Null

	Endif
	Clear Class isxForm

Endproc


*====================================================================
* Determines the source of the window identified by the passed 
* WHandle. It returns the following values:
*
* -1     The window is not an edit window
*  0     Command Window
*  1     MODIFY COMMAND window
*  2     MODIFY FILE window
*  8     Menu Designer code window
* 10     Method Edit Window in Class or Form Designer
* 12     MODIFY PROCEDURE window
*
* This procedure uses _EdGetEnv() from the FoxTools.Fll to determine
* the edit source. Passing an invalid handle causes an exception in
* VFP 5 and VFP 3. In VFP 6 this raises an error 2028 (API function
* caused an exception). Therefore we return -1 in this case, too.
*====================================================================
 Procedure GetEditSource
LParameter m.tnWHandle

	Local laEnv[25], lnSource, lnOK, lcError
	m.lcError = On( "Error" )
	On Error m.lnOK = 0
	m.lnOK = _EdGetEnv( m.tnWHandle, @m.laEnv )
	On Error &lcError
	If m.lnOK == 0
		m.lnSource = -1
	Else
		m.lnSource = m.laEnv[25]
	Endif
	
Return m.lnSource


*====================================================================
* Returns the WHandle of the current edit window or 0, if no edit
* window is available.
*====================================================================
 Procedure GetCurrentWindow

	Local lnWindowOnTop
	m.lnWindowOnTop = _WOnTop()
	If m.lnWindowOnTop <= 0
		Return 0
	Endif
	If GetEditSource( m.lnWindowOnTop ) == -1
		m.lnWindowOnTop = 0
	Endif
	
Return m.lnWindowOnTop


*====================================================================
* Returns the current cursor position in the edit window identified
* by the WHandle. On error -1 is returned.
*====================================================================
 Procedure GetFileCursorPos
Lparameters tnWHandle

	Local lnCursorPos
	m.lnCursorPos = _EdGetPos( m.tnWHandle )
	
Return m.lnCursorPos


*====================================================================
* Changes the current cursor position in the edit window identified
* by the WHandle.
*====================================================================
 Procedure SetFileCursorPos
LParameter m.toEditBox, m.tnPosition

	m.toEditBox.SelStart = m.tnPosition
	m.toEditBox.SelLength =0 

EndProc


*====================================================================
* Returns the current line of the edit window identified by the
* WHandle. The line number is zero based. On Error -1 is returned.
*====================================================================
 Procedure GetCurrentLine
LParameters tnWHandle

	Local lnCursorPos, lnLineNo
	m.lnCursorPos = GetFileCursorPos( m.tnWHandle )
	If m.lnCursorPos < 0
		m.lnLineNo = -1
	Else
		m.lnLineNo = _EdGetLNum( m.tnWhandle, m.lnCursorPos )
	Endif
	
Return m.lnLineNo


*====================================================================
* Returns the cursor position within the current line of the edit
* window identified by the WHandle. The cursor position is 0 based.
* On error -1 is returned.
*====================================================================
 Procedure GetCurrentCol
Lparameters tnWHandle

	Local lnCursorPos, lnLineNo, lnColumn, lnLineStart
	m.lnCursorPos = GetFileCursorPos( m.tnWHandle )
	If m.lnCursorPos < 0
		Return -1
	Endif
	m.lnLineNo = GetCurrentLine( m.tnWHandle )
	If m.lnLineNo < 0
		Return -1
	Endif
	m.lnLineStart = GetLineStart( m.tnWHandle, m.lnLineNo )
	m.lnColumn = m.lnCursorPos - m.lnLineStart

Return m.lnColumn


*====================================================================
* Returns the beginning of the specific line in the edit window
* identified by WHandle. Returns -1 on error.
*====================================================================
 Procedure GetLineStart
LParameter m.tnWHandle, m.tnLineNo

	Local lnLineStart
	m.lnLineStart = _EdGetLPos( m.tnWHandle, m.tnLineNo )
	
Return m.lnLineStart


*====================================================================
* Returns the text of the specified line in the edit window 
* identified by the WHandle. A terminating carriage return is 
* removed. Returns an empty string on error. The line must be zero
* based.
*====================================================================
 Procedure GetLine
Lparameters tnWHandle, tnLine

	Local lnStartPos, lnEndPos, lcString
	m.lnStartPos = GetLineStart( m.tnWHandle, m.tnLine )
	m.lnEndPos = GetLineStart( m.tnWHandle, m.tnLine+1 )
	If m.lnStartPos == m.lnEndPos
		m.lcString = ""
	Else
		m.lnEndPos = m.lnEndPos - 1
		m.lcString = _EdGetStr( m.tnWHandle, m.lnStartPos, m.lnEndPos )
		m.lcString = Chrtran( m.lcString, Chr(13), "" )
	Endif

Return m.lcString


*====================================================================
* Returns the text in the current line that is to the left of the 
* cursor in the edit window identified by the WHandle. Returns "" on
* error.
*====================================================================
 Procedure GetLineLeftFromCursor
Lparameters toEditBox
Return Left(m.toEditBox.Value, m.toEditBox.SelStart)

*!* * Removed 8/1/2012 / JRN
*!* Lparameters tnWHandle

*!* 	Local lnCurLine, lnCurCol, lcLine
*!* 	lnCurLine = GetCurrentLine( m.tnWHandle )
*!* 	If m.lnCurLine < 0
*!* 		Return ""
*!* 	Endif
*!* 	lnCurCol = GetCurrentCol( m.tnWHandle )
*!* 	If m.lnCurCol < 0
*!* 		Return ""
*!* 	Endif
*!* 	If m.lnCurCol == 0
*!* 		lcLine = ""
*!* 	Else
*!* 		lcLine = GetLine( m.tnWHandle, m.lnCurLine )
*!* 		lcLine = Left( m.lcLine, m.lnCurCol )
*!* 	Endif

*!* Return m.lcLine


*====================================================================
* Inserts text in the edit window identified by WHandle. The text is
* stored in tcText, the position is optional. tcOptions can contains
* a combination of the following values:
*
*  R   The current selection is replaced
*  B   The cursor is positioned at the beginning of the inserted
*      text.
*  E   (default) The cursor is positioned at the end of the inserted 
*      text.
*  H   The inserted text is highlighted.
*====================================================================
*** JRN 8/1/2012 : <tnPosition> is never used

 Procedure InsertText
Lparameters toEditBox, tcText, tnPosition, tcOptions

	*-----------------------------------------------------------------
	* Normalize options
	*-----------------------------------------------------------------
	Local lcOptions
	If Vartype(m.tcOptions) == "C"
		m.lcOptions = Upper( Alltrim(m.tcOptions) )
	Else
		m.lcOptions = ""
	Endif
	
	*-----------------------------------------------------------------
	* If a position is passed, Change the current cursor position
	* accordingly.
	*-----------------------------------------------------------------
	*!* * Removed 8/2/2012 / JRN
	*!* If Vartype(m.tnPosition) == "N"
	*!* 	SetFileCursorPos( m.tnWHandle, m.tnPosition )
	*!* Endif
	
	*-----------------------------------------------------------------
	* Insert the Text at the current position. If the "R" option is
	* used, delete the current selection.
	*-----------------------------------------------------------------
	Local lnSelLength, lnSelStart

	m.lnSelStart	= m.toEditBox.SelStart
	m.lnSelLength	= m.toEditBox.SelLength
	If 'R' $ m.lcOptions
			m.toEditBox.Value		= Left (m.toEditBox.Value, m.lnSelStart)		;
				+ Substr (m.toEditBox.Value, m.lnSelStart + m.lnSelLength + 1)
			m.toEditBox.SelLength	= 0
			m.lnSelLength			= m.toEditBox.SelLength
	Endif
	m.toEditBox.Value		= Left (m.toEditBox.Value, m.lnSelStart)		;
		+ m.tcText													;
		+ Substr (m.toEditBox.Value, m.lnSelStart + m.lnSelLength + 1)
	m.toEditBox.SelLength	= m.toEditBox.SelLength + Len (m.tcText)
	m.lnSelLength			= m.toEditBox.SelLength
		
	*!* * Removed 8/2/2012 / JRN
	*!* *-----------------------------------------------------------------
	*!* * Set the cursor accordingly. "E" is the default of VFP. We don't
	*!* * need any action for that.
	*!* *-----------------------------------------------------------------
	*!* Do Case
	*!* Case "B" $ m.lcOptions
	*!* 	SetFileCursorPos( m.tnWHandle, m.lnStartPosition )
	*!* Case "H" $ m.lcOptions
	*!* 	_EdSelect( m.tnWHandle, m.lnStartPosition, m.lnEndPosition )
	*!* Endcase

EndProc


*========================================================================================
* Fills an array with all variable declarations in the current procedure of the edit 
* window identified by the WHandle. Variable declarations are only searched backward from
* the current position. Returns the number of variables.
*
*! 2004-10Oct-19  ChrisW 
*      Added support for variables with non-english characters such as "ö".
*      In VFP 9 the array limitation has been lifted.
*========================================================================================
 Procedure CP_Variables
Lparameters toISX
    Local loVariables AS Object
 
	*--------------------------------------------------------------------------------------
	* Check if the current entity is a variable
	*--------------------------------------------------------------------------------------
	Local llIsVariable
	DO Case
	Case Upper(m.toISX.cEntity)=="M"
		m.llIsVariable = .T.
	Case Empty(m.toISX.cEntity)
		If Empty(m.toISX.cInvocation)
			m.llIsVariable = .T.
		Else
			m.llIsVariable = .F.
		EndIf 
	Otherwise 
		m.llIsVariable = .F.
	EndCase 
	If not m.llIsVariable 
		Return
	EndIf
	
	*-----------------------------------------------------------------
	* Get the current line as a starting point. We start with the line
	* before that line. 
	*-----------------------------------------------------------------
	Local lnEnd
	m.lnEnd = GetCurrentLine( m.toISX.nWHandle )
	If m.lnEnd <= 0
		Return
	Else
		m.lnEnd = m.lnEnd - 1
	Endif

	*-----------------------------------------------------------------
	* Because GetLine() is quite slow with large program files, we
	* read the entire program up to the line before the current line
	* into an array and parse that. Since an array can only contain
	* up to 65000 lines, we make sure that we don't read more than 
	* that into the laText array.
	*-----------------------------------------------------------------
	Local lnLineCount, laText[1], lnStart
	If m.lnEnd >= 65000 and Version(4) < "09.00"
		m.lnStart = m.lnEnd - 65000
	Else
		m.lnStart = 0
	Endif
	m.lnLineCount = AGetLines(m.toISX.nWHandle,@m.laText,m.lnStart,m.lnEnd)
		
	*--------------------------------------------------------------------------------------
	*** JRN 7/26/2012 :
	* If PEM Editor / IDE Tools is available, use its method GetVariablesList
	*   to get the list of all variables, including those assigned but not localized
	* Parse all line backward until we get to PROCEDURE or FUNCTION,
	*   then let GetVariablesList do its work on the intervening text
	*--------------------------------------------------------------------------------------
	Local lnCurrentLine, lcLine, lnPos, lcCommand, lcValidCmds, lcCodeBlock, lcVar
	If 'O' = Type('_oPEMEditor.oUtils.oBeautifyX')
		m.lcCodeBlock = ''
		For m.lnCurrentLine = m.lnLineCount to 1 Step -1
			m.lcLine = NormalizeLine( m.laText[m.lnCurrentLine] )
			m.lcCodeBlock = m.lcLine + Chr(13) + m.lcCodeBlock
			If Len(m.lcLine) < 4
				Loop
			EndIf
			m.lcCommand = GetWordNum(m.lcLine,1)
			If IsFoxProCommand( m.lcCommand, "PROCEDURE,FUNCTION,PROTECTED,HIDDEN" )
				Exit
			Endif
		EndFor
		m.loVariables = _oPEMEditor.oUtils.oBeautifyX.GetVariablesList(m.lcCodeBlock)
		For each m.lcVar in m.loVariables foxobject 
			m.toISX.nCount = m.toISX.nCount + 1
			Dimension m.toISX.aList[m.toISX.nCount]
			m.toISX.aList[m.toISX.nCount] = m.lcVar			
		EndFor 	
		Return
	EndIf	
	*--------------------------------------------------------------------------------------
	* Parse all lines backwards for the following keywords: LOCAL,
	* PUBLIC, PROCEDURE, FUNCTION. We add all variables in the
	* LOCAL and PUBLIC lines and stop parsing when we find PROCEDURE
	* or FUNCTION.
	*--------------------------------------------------------------------------------------
	For m.lnCurrentLine = m.lnLineCount to 1 Step -1
		m.lcLine = NormalizeLine( m.laText[m.lnCurrentLine] )
		If Len(m.lcLine) < 4
			Loop
		EndIf
		If Version(4) >= "07.00"
			m.lcCommand = GetWordNum(m.lcLine,2)
		Else
			m.lcCommand = X6_GetWordNum(m.lcLine,2)
		EndIf 
		If m.lcCommand == "="
			Loop
		EndIf 
		If Version(4) >= "07.00"
			m.lcCommand = GetWordNum(m.lcLine,1)
		Else
			m.lcCommand = X6_GetWordNum(m.lcLine,1)
		EndIf 
		m.lcValidCmds = ;
			"LOCAL,PUBLIC,LPARAMETERS,PARAMETERS,PRIVATE,PROCEDURE,FUNCTION,PROTECTED," + ;
			"HIDDEN"
		If not IsFoxProCommand(m.lcCommand,m.lcValidCmds)
			Loop
		EndIf
		m.lnPos = At( " ", m.lcLine )
		If m.lnPos == 0 or m.lnPos == Len(m.lcLine)
			Loop
		Endif
		m.lcLine = Alltrim( Substr(m.lcLine,m.lnPos) )
	  If IsFoxProCommand(m.lcCommand,"LOCAL")
			If Version(4) >= "07.00"
				m.lcCommand = GetWordNum(m.lcLine,1)
			Else
				m.lcCommand = X6_GetWordNum(m.lcLine,1)
			EndIf 
			If IsFoxProCommand(m.lcCommand,"ARRAY")
				m.lnPos = At( " ", m.lcLine )
				If m.lnPos == 0 or m.lnPos == Len(m.lcLine)
					Loop
				Endif
				m.lcLine = Alltrim( Substr(m.lcLine,m.lnPos) )
			EndIf 
	  EndIf
		If IsFoxProCommand( m.lcCommand, "PROCEDURE,FUNCTION,PROTECTED,HIDDEN" )
			m.lnPos = At( "(", m.lcLine )
			If m.lnPos == 0 or m.lnPos == Len(m.lcLine)
				Exit
			EndIf
			m.lcLine = Substr(m.lcLine,m.lnPos+1)
		EndIf
	  m.lnCurrentLine = m.lnCurrentLine - ;
	  	CP_VariablesAdd( m.toISX, m.lcLine, m.lnCurrentLine, @m.laText )
		If IsFoxProCommand( m.lcCommand, "PROCEDURE,FUNCTION,PROTECTED,HIDDEN" )
			Exit
		Endif
	Endfor
	
EndProc 


*========================================================================================
* 
*========================================================================================
 Procedure CP_VariablesAdd
LParameter m.toISX, m.tcLine, m.tnCurrentLine, m.taText
External array taText 
	Local lcLine, lnLineOffset, lnCurrentVar, laDeclarations[1], lcCurrentVar, ;
		lnPosInVar, lcChar, lnPos
	m.lcLine = m.tcLine
	m.lnLineOffset = 0
	
	Do While .T.
		m.lcLine = Chrtran( m.lcLine, ",", Chr(13) )
		For m.lnCurrentVar = 1 to ALines( laDeclarations, m.lcLine )
			m.lcCurrentVar = Alltrim( m.laDeclarations[m.lnCurrentVar] )
			If Empty( m.lcCurrentVar )
				Loop
			Endif
			If     not IsAlpha( m.lcCurrentVar ) ;
			   and not Left(m.lcCurrentVar,1) == "_"
				Loop
			Endif
			m.lnPos = At( " ", m.lcCurrentVar )
			If m.lnPos == 0
				m.lnPos = Len( m.lcCurrentVar )
			Else
				m.lnPos = m.lnPos - 1
			Endif
			m.lcCurrentVar = Left( m.lcCurrentVar, m.lnPos )
			If LEFT(LOWER(m.lcCurrentVar),2)=='m.'
				m.lcCurrentVar = SUBSTR(m.lcCurrentVar,3)
			EndIf
			For m.lnPosInVar = 2 to Len(m.lcCurrentVar)
				m.lcChar = SubStr(m.lcCurrentVar,m.lnPosInVar,1)
				If not (IsAlpha(m.lcChar) or IsDigit(m.lcChar) or m.lcChar="_")
					m.lcCurrentVar = Left( m.lcCurrentVar, m.lnPosInVar-1 )
					Exit
				Endif
			Endfor
			m.toISX.nCount = m.toISX.nCount + 1
			Dimension m.toISX.aList[m.toISX.nCount]
			m.toISX.aList[m.toISX.nCount] = m.lcCurrentVar
		Endfor
		If Right(m.lcLine,1) # ";"
			Exit
		Endif
		m.lnLineOffset = m.lnLineOffset + 1
		If m.tnCurrentLine + m.lnLineOffset > Alen(m.taText,1)
			Exit
		Endif
		m.lcLine = NormalizeLine( ;
			m.taText[m.tnCurrentLine+m.lnLineOffset] ;
		)
	Enddo

Return m.lnLineOffset


*========================================================================================
* Returns .T., when the first string is a FoxPro command.
*========================================================================================
 Procedure IsFoxProCommand
LParameter m.tcCommand, m.tcCommandList

	Local laList[1], lnLine, llFound
	
	m.llFound = .F.
	For m.lnLine=1 to ALines(laList,Chrtran(m.tcCommandList,",",Chr(13)+Chr(10)))
		If Left(Upper(m.laList[m.lnLine]),Len(m.tcCommand)) == Upper(m.tcCommand)
			m.llFound = .T.
			Exit
		Endif
	EndFor 

Return m.llFound


*====================================================================
* Normalizes a line. This means: All tabs are converted to single
* blanks, leading or trailing blanks are removed. Comments starting
* with && are removed.
*====================================================================
 Procedure NormalizeLine
Lparameters tcLine

	Local lcLine, lnPos
	m.lcLine = Chrtran( m.tcLine, Chr(9), " " )
	If "&"+"&" $ m.lcLine
		m.lnPos = At( "&"+"&", m.lcLine )
		m.lcLine = Left( m.lcLine, m.lnPos-1 )
	Endif 
	m.lcLine = Alltrim(m.lcLine)

Return m.lcLine


*====================================================================
* GetKeyLabel takes the parameters passed to the KeyPress event and
* returns the label name that can be used for KEYBOARD or ON KEY
* LABEL, etc.
*====================================================================
 Procedure GetKeyLabel
LParameter m.tnKeyCode, m.tnSAC

	Local lcLabel
	Do Case
	Case Between(m.tnKeyCode,33,126)
		m.lcLabel = Chr(m.tnKeyCode)
	Case Between(m.tnKeyCode,128,255)
		m.lcLabel = Chr(m.tnKeyCode)
	Case m.tnSAC == 2 and Between(m.tnKeyCode,1,26)
		Do Case
		Case m.tnKeyCode == 2
			m.lcLabel = "CTRL+RIGHTARROW"
		Case m.tnKeyCode == 8
			m.lcLabel = ""
		Case m.tnKeyCode == 10
			m.lcLabel = "CTRL+ENTER"
		Case m.tnKeyCode == 23
			m.lcLabel = "CTRL+END"
		Case m.tnKeyCode == 26
			m.lcLabel = "CTRL+LEFTARROW"
		Otherwise
			m.lcLabel = "CTRL+" + Chr(m.tnKeyCode+64)
		Endcase
	Case m.tnSAC == 0 and m.tnKeyCode < 0
		m.lcLabel = "F" + Alltrim(Str(Abs(m.tnKeyCode)+1))
	Case m.tnSAC == 0 and m.tnKeyCode == 22
		m.lcLabel = "INS"
	Case m.tnSAC == 1 and m.tnKeyCode == 22
		m.lcLabel = "SHIFT+INS"
	Case m.tnSAC == 0 and m.tnKeyCode == 1
		m.lcLabel = "HOME"
	Case m.tnSAC == 0 and m.tnKeyCode == 7
		m.lcLabel = "DEL"
	Case m.tnSAC == 0 and m.tnKeyCode == 28
		m.lcLabel = "F1"
	Case m.tnSAC == 0 and m.tnKeyCode == 6
		m.lcLabel = "END"
	Case m.tnSAC == 0 and m.tnKeyCode == 18
		m.lcLabel = "PGUP"
	Case m.tnSAC == 0 and m.tnKeyCode == 3
		m.lcLabel = "PGDN"
	Case m.tnSAC == 0 and m.tnKeyCode == 5
		m.lcLabel = "UPARROW"
	Case m.tnSAC == 0 and m.tnKeyCode == 28
		m.lcLabel = "F1"
	Case m.tnSAC == 0 and m.tnKeyCode == 24
		m.lcLabel = "DNARROW"
	Case m.tnSAC == 0 and m.tnKeyCode == 4
		m.lcLabel = "RIGHTARROW"
	Case m.tnSAC == 0 and m.tnKeyCode == 19
		m.lcLabel = "LEFTARROW"
	Case m.tnSAC == 0 and m.tnKeyCode == 27
		m.lcLabel = "ESC"
	Case m.tnSAC == 0 and m.tnKeyCode == 13
		m.lcLabel = "ENTER"
	Case m.tnSAC == 0 and m.tnKeyCode == 127
		m.lcLabel = "BACKSPACE"
	Case m.tnSAC == 0 and m.tnKeyCode == 9
		m.lcLabel = "TAB"
	Case m.tnSAC == 0 and m.tnKeyCode == 32
		m.lcLabel = "SPACEBAR"
	Case m.tnSAC == 1 and m.tnKeyCode == 13
		m.lcLabel = "SHIFT+ENTER"
	Case m.tnSAC == 1 and m.tnKeyCode == 127
		m.lcLabel = "SHIFT+BACKSPACE"
	Case m.tnSAC == 1 and m.tnKeyCode == 15
		m.lcLabel = "SHIFT+TAB"
	Case m.tnSAC == 1 and m.tnKeyCode == 32
		m.lcLabel = "SHIFT+SPACEBAR"
	Case m.tnSAC == 2 and m.tnKeyCode == 29
		m.lcLabel = "CTRL+HOME"
	Case m.tnSAC == 2 and m.tnKeyCode == 31
		m.lcLabel = "CTRL+PGUP"
	Case m.tnSAC == 2 and m.tnKeyCode == 30
		m.lcLabel = "CTRL+PGDN"
	Case m.tnSAC == 2 and m.tnKeyCode == 128
		m.lcLabel = "CTRL+BACKSPACE"
	Case m.tnSAC == 2 and m.tnKeyCode == 32
		m.lcLabel = "CTRL+SPACEBAR"
	Otherwise
		m.lcLabel = ""
	Endcase

Return m.lcLabel


*====================================================================
* Fills an array with all lines between nStart and nEnd. 
*====================================================================
 Procedure AGetLines
LParameter m.tnWHandle, m.raText, m.tnStart, m.tnEnd

	*-----------------------------------------------------------------
	* Copy the text between nStart and nEnd into a string variable.
	*-----------------------------------------------------------------
	Local lnStartPos, lnEndPos, lcString
	m.lnStartPos = GetLineStart( m.tnWHandle, m.tnStart )
	m.lnEndPos = GetLineStart( m.tnWHandle, m.tnEnd+1 ) - 1
	m.lcString = _EdGetStr( m.tnWHandle, m.lnStartPos, m.lnEndPos )

	*-----------------------------------------------------------------
	* And parse this into an array
	*-----------------------------------------------------------------
	Local lnCount
	m.lnCount = ALines( raText, m.lcString )
	
Return m.lnCount


*====================================================================
* The FoxTools function _AGetEnv() doesn't return proper font infor-
* mation. Instead it claims that "MS Sans Serif", 8 pt. is the 
* current font. This function returns font information for the speci-
* fied window by accessing the GDI.
*====================================================================
 Procedure WGetFontInfo
LParameter m.tnWHandle, m.rcFontName, m.rnFontSize, m.rnStyle

	*-----------------------------------------------------------------
	* In addition to the window handle of this window we also need
	* the HWND of the child window that contains the actual editor.
	* The GetClientWindow() function retrieves this window handle.
	*-----------------------------------------------------------------
	Local lnHWND
	m.lnHWND = GetClientWindow( m.tnWHandle )
	If m.lnHWND == 0
		Return .F.
	Endif
	
	*-----------------------------------------------------------------
	* Using this HWND we can then get a Device Context. 
	*-----------------------------------------------------------------
	Local lnHWND, lnHDC
	Declare LONG GetDC in Win32API LONG
	m.lnHDC = GetDC( m.lnHWND )
	If m.lnHDC == 0
		Return .F.
	Endif
	
	*-----------------------------------------------------------------
	* With this device context we can now get an object handle to the
	* currently selected font.
	*-----------------------------------------------------------------
	Local lnHFONT
	Declare LONG GetCurrentObject in Win32API LONG, LONG
	m.lnHFONT = GetCurrentObject( m.lnHDC, 6 )  && OBJ_FONT
	If m.lnHFONT == 0
		Return .F.
	Endif
	
	*-----------------------------------------------------------------
	* The HFONT handle to the current font can be used to obtain more
	* detailled information about the selected font. We need to rename
	* the API function GetObject(), because it interferes with VFP's
	* GETOBJECT() function
	*-----------------------------------------------------------------
	Local lcLogFont
	Declare Integer GetObject in Win32API as GDI_GetObject ;
		LONG, Integer, String@
	m.lcLogFont = Replicate( Chr(0), 1024 )
	If GDI_GetObject( m.lnHFONT, 1024, @m.lcLogFont ) == 0
		Return .F.
	Endif
	
	*-----------------------------------------------------------------
	* Now to extract the font information from the LOGFONT structure.
	*-----------------------------------------------------------------
	Local lnSize, lcName, lnStyle
	m.lnSize = Abs( FromInt(Left(m.lcLogFont,4)) - 2^32 )
	m.lcName = SubStr( m.lcLogFont, 29 )
	m.lcName = Left( m.lcName, At(Chr(0),m.lcName)-1 )
	m.lnStyle = 0
	If FromInt(SubStr(m.lcLogFont,17,4)) == 700
		m.lnStyle = m.lnStyle + 1
	Endif
	If FromInt(SubStr(m.lcLogFont,21,4)) # 0
		m.lnStyle = m.lnStyle + 2
	Endif
	
	*-----------------------------------------------------------------
	* We now have the height of the font in pixels but what we need 
	* are points.
	*-----------------------------------------------------------------
	Local lnResolution
	Declare Integer GetDeviceCaps in Win32API Integer, Integer
	m.lnResolution = GetDeviceCaps( m.lnHDC, 90 ) && LOGPIXELSY
	m.lnSize = m.lnSize / m.lnResolution * 72
	m.lnSize = Round( m.lnSize, 0 )
		
	*-----------------------------------------------------------------
	* Finally release the device context
	*-----------------------------------------------------------------
	Declare Integer ReleaseDC In Win32API LONG, LONG
	ReleaseDC( m.lnHWND, m.lnHDC )
	
	*-----------------------------------------------------------------
	* And pass the values pack as parameters
	*-----------------------------------------------------------------
	m.rcFontName = m.lcName
	m.rnFontSize = m.lnSize
	m.rnStyle = m.lnStyle
	
Return .T.


*====================================================================
* The editor only works on the editor window and you can only get the
* HWND of this window using the Window Handle. For many Windows ope-
* rations, however, you need the HWND of the child window that con-
* tains the actual editor area. This function returns the HWND of 
* this window. It's not that easy, because Method snippet windows
* actually have two child windows, one for the text editor and one
* with the method and object dropdown combos.
*====================================================================
 Procedure GetClientWindow
LParameter m.tnWHandle
	
	*-----------------------------------------------------------------
	* Convert the Window Handle into a HWND
	*-----------------------------------------------------------------
	Local lnHWND
	m.lnHWND = _WhToHWND( m.tnWHandle )

	*-----------------------------------------------------------------
	* FindWindowEx returns all child windows of a given parent window.
	* We use it to find a child of the edit window that doesn't have
	* another child window, because method edit windows have a second 
	* which we can identify since it has another child window.
	*-----------------------------------------------------------------
	Local lnChild
	Declare Integer FindWindowEx in Win32API ;
		Integer, Integer, String, String
	m.lnChild = 0
	Do While .T.
		m.lnChild = FindWindowEx( m.lnHWND, m.lnChild, NULL, NULL )
		If m.lnChild == 0
			Exit
		Endif
		If FindWindowEx( m.lnChild, 0, NULL, NULL ) == 0
			Exit
		Endif
	Enddo
		
Return m.lnChild


*====================================================================
* Returns the position of the text cursor (caret) in _SCREEN coordi-
* nates. If the window identified by the passed window handle doesn't
* have the focus, or the position can't be determined, this function
* returns .F.
*====================================================================
 Procedure GetCaretPosition
LParameter m.tnWHandle, m.rnTop, m.rnLeft

	*-----------------------------------------------------------------
	* Check whether this window has got the focus.
	*-----------------------------------------------------------------
	Declare Integer GetFocus in Win32API
	If GetFocus() # _WhToHWND( m.tnWHandle )
		Return .F.
	Endif
	
	*-----------------------------------------------------------------
	* Determine the cursor position. This position is relative to the
	** OK
	* client area of the editing subwindow of the actual editing win-
	* dow.
	*-----------------------------------------------------------------
	Local lnLeft, lnTop, lcPOINT
	Declare Integer GetCaretPos in Win32API String@
	m.lcPOINT = Space(8)
	If GetCaretPos( @m.lcPOINT ) == 0
		m.lnLeft = MCol(3)
		m.lnTop = MRow(3)
	Else
		m.lnLeft = Asc(Left(m.lcPOINT,1))+256*Asc(SubSTr(m.lcPOINT,2,1))
		m.lnTop = Asc(SubSTr(m.lcPOINT,5,1))+256*Asc(SubStr(m.lcPOINT,6,1))
	Endif
	
	*-----------------------------------------------------------------
	* To convert this postion to _SCREEN coordinates, we have to 
	* determine the position of the client window relative to the 
	* desktop window and correlate this with the absolute position of
	* the _SCREEN window. Hence, we need first the HWNDs of both 
	* windows.
	*-----------------------------------------------------------------
	Local lnChild, lnScreen
	Declare Integer GetParent in Win32API Integer
	m.lnChild = GetClientWindow( m.tnWHandle )
	If m.lnChild == 0
		Return .F.
	Endif
	m.lnScreen = GetParent( _WhToHWND(m.tnWHandle) )
	If m.lnScreen == 0
		Return .F.
	Endif

	*-----------------------------------------------------------------
	* Now we can determine the position of both windows.
	*-----------------------------------------------------------------
	Local lnChildTop, lnChildLeft, lnScreenTop, lnScreenLeft, lcRect
	m.lcRect = Replicate( Chr(0), 16 )
	Declare Integer GetWindowRect in Win32API Long, String@
	GetWindowRect( m.lnChild, @m.lcRect )
	m.lnChildLeft = FromInt( Left(m.lcRect,4) )
	m.lnChildTop = FromInt( SubSTr(m.lcRect,5,4) )
	GetWindowRect( m.lnScreen, @m.lcRect )
	m.lnScreenLeft = FromInt( Left(m.lcRect,4) )
	m.lnScreenTop = FromInt( SubSTr(m.lcRect,5,4) )
	
	*-----------------------------------------------------------------
	* Now combine the position of the edit window and the cursor
	* position.
	*-----------------------------------------------------------------
	m.rnLeft = m.lnLeft + m.lnChildLeft - m.lnScreenLeft
	m.rnTop = m.lnTop + m.lnChildTop - m.lnScreenTop

EndProc
 	
	
Procedure FromInt
Parameter tcString
  Private m.nValue, m.nT
  m.nValue =0
  For m.nT = 1 to Len(tcString)
    m.nValue = m.nValue + Asc(SubStr(tcString,m.nT,1))*256^(m.nT-1)
  Endfor
Return m.nValue


*====================================================================
* The following class displays a popup window at the current cursor
* position and lets the user continue to type.
*
* The characters a-z, A-Z, 0-9 and _ are inserted into the active
* edit window as the user types. The previous position is saved in
* order to restore the text if necessary.
*
* ESC terminates the popup and doesn't change the text.
* 
* TAB inserts the current selection and terminates the popup.
*
* SPACEBAR inserts the current selection, adds a blank and terminates
* the popup.
*
* Any other key terminates the popup and is repeated so it is handled
* properly by VFP. If the user enters the first character that 
* doesn't match an item in the list, or entered a full item where 
* none exists that has the same name, but additional characters, the
* list is terminated as well.
*
*====================================================================
 Define CLASS isxFormDeskTop as ISXForm
	DeskTop = .T.
EndDefine

Define CLASS isxForm as Form

	AlwaysOnTop = .T.
	WindowType = 1
	TitleBar = 0
	BorderStyle = 0

	nWHandle = 0
	nCurrentPos = 0
	cSearchString = ""
	cVarString = ""
	Dimension aItems[1,3]
	Dimension aItemsSource[1,3]
	lScrolled = .F.
	*Mike Yearwood - these support reducing screen caption flicker
	cScreenCaption = ""
	cWindowCaption = ""
	lMaximized = .F.
	cOnKeyLabelDocActivation = ''
	oISXOptions = .null.
	oEditBox = .null.
	
	Add Object isxList as Listbox with ;
		ColumnCount = 2, ;
		ColumnLines = .F., ;
		IncrementalSearch = .F., ;
		AutoHideScrollBar = 1

PROCEDURE Load
this.lMaximized = wmaximum()
IF THIS.lMaximized
	THIS.cWindowCaption = LOWER(WTITLE())
	THIS.cScreenCaption = _screen.Caption
ENDIF
RETURN DODEFAULT()
ENDPROC

 PROCEDURE Show
*====================================================================
* Mike Yearwood
* When the edit window is maximized, the screen caption reads
* currentedit.prg * - current vfp system window caption
* When this window goes active, the screen caption changes
* which causes a flicker. To stop that flicker, set the screen
* caption to what it was before.
*====================================================================

IF THIS.lMaximized
	_Screen.Caption = this.cWindowCaption + " * - " + this.cScreenCaption
ENDIF
	*** JRN 9/11/2011 : remove dot activation while this form is active
	#IFDEF DOT_ACTIVATION
		This.cOnKeyLabelDocActivation = On('Key', [DOT_ACTIVATION])
		On Key Label DOT_ACTIVATION 
	#ENDIF
ENDPROC

 PROCEDURE Destroy
*Mike Yearwood
*Prevent screen caption flicker.
IF THIS.lMaximized
	_Screen.Caption = this.cScreenCaption
ENDIF
	*** JRN 9/11/2011 : restore dot activation 
	This.oEditBox = .null.

	#IFDEF DOT_ACTIVATION
		Local lcOnKeyLabelDocActivation
		m.lcOnKeyLabelDocActivation = This.cOnKeyLabelDocActivation
		On Key Label DOT_ACTIVATION  &lcOnKeyLabelDocActivation 
	#ENDIF
ENDPROC

*====================================================================
* When the form is initialized, we have to determine its position
* and get a handle to the current edit window. Pass an array to this
* form that contains all possible values the user can enter.
*====================================================================
 Procedure Init
LParameter m.toISX
With This

	*-----------------------------------------------------------------
	* Get the handle for the current window. 
	*-----------------------------------------------------------------
	*!* * Removed 8/1/2012 / JRN
	*!* .nWHandle = toISX.nWHandle
	*!* .nCurrentPos = GetFileCursorPos( .nWHandle )
	.oEditBox = m.toISX.oEditBox 
	
		*-----------------------------------------------------------------
		* Copy the array and sort it case-insensitive
		*-----------------------------------------------------------------
		If 0 = Alen(m.toISX.aList, 2)
			Dimension m.toISX.aList[Alen(m.toISX.aList, 1), 1]
		Endif

		If m.toISX.lSortList Or m.toISX.ISXOptions.AlphaSort
			*!* * Removed 8/25/2012 / JRN
			*!* If Version(4) >= "07.00"
			Asort( m.toISX.aList, -1, -1, 0, 1 )
			*!* * Removed 8/25/2012 / JRN
			*!* Else
			*!* 	Dimension laValues[toISX.nCount,2]
			*!* 	For lnValue = 1 to toISX.nCount
			*!* 		laValues[m.lnValue,1] = Upper(toISX.aList[m.lnValue])
			*!* 		laValues[m.lnValue,2] = m.lnValue
			*!* 	EndFor
			*!* 	Asort( laValues, 1 )
			*!* EndIf 
		Else

		Endif
		
		*--------------------------------------------------------------------------------------
		* Fill the listbox with all possible values.
		*--------------------------------------------------------------------------------------
		With .isxList
			.FontName = m.toISX.ISXOptions.FontName
			.FontSize = m.toISX.ISXOptions.FontSize
		Endwith
		.oISXOptions = m.toISX.ISXOptions
		*	.lSearchSecondColumn = m.toISX.lSearchSecondColumn

		m.lnMaxWidth		   = 0
		m.lnMaxFieldDefWidth = 0
		m.lcVarString		   = ''
		Dimension .aItems[m.toISX.nCount, 6]
		.aItems		   = ''
		m.lnColumnCount  = Min(5, Alen(m.toISX.aList, 2))
		m.lnAvgCharWidth = Fontmetric(6, .isxList.FontName, .isxList.FontSize)
		For m.lnValue = 1 To m.toISX.nCount
			*!* * Removed 8/26/2012 / JRN
			*!* If Version(4) >= "07.00"
			m.lcValue = Alltrim(m.toISX.aList[m.lnValue, 1])
			*!* * Removed 8/26/2012 / JRN
			*!* Else
			*!* 	lcValue = toISX.aList[laValues[m.lnValue,2]]
			*!* EndIf 

			If ccTAB $ m.lcValue
				.aItems[m.lnValue, 3] = Getwordnum(m.lcValue, 2, ccTAB)
				m.lcValue				  = Getwordnum(m.lcValue, 1, ccTAB)
			Endif

			.aItems[m.lnValue, 1] = Upper(m.lcValue)
			.aItems[m.lnValue, 2] = m.lcValue
			m.lcVarString			  = m.lcVarString + ':' + Padr(Upper(m.lcValue), 128)
			m.lnWidth				  = Txtwidth(m.lcValue, .isxList.FontName, .isxList.FontSize) * m.lnAvgCharWidth
			m.lnMaxWidth			  = Max( m.lnMaxWidth, m.lnWidth )
			For m.lnCol = 2 To m.lnColumnCount
				.aItems[m.lnValue, m.lnCol + 1] = Evl(m.toISX.aList[m.lnValue, m.lnCol], '')
			EndFor
			If not Empty(.aItems[m.lnValue, 3])
				m.lnWidth				  = Txtwidth(.aItems[m.lnValue, 3], .isxList.FontName, .isxList.FontSize) * m.lnAvgCharWidth
				m.lnMaxFieldDefWidth	  = Min(cnMaxColumn2Width, Max( m.lnMaxFieldDefWidth, m.lnWidth )) 
			EndIf 
		Endfor

		Dimension .aItemsSource(Alen(.aItems, 1), Alen(.aItems, 2))
		Acopy(.aItems, .aItemsSource)
		.cVarString = m.lcVarString
	*** JRN 9/11/2011 : Add a little more to width
	m.lnMaxWidth = m.lnMaxWidth + 30
	If m.toISX.ISXOptions.ShowFieldDefs and m.lnMaxFieldDefWidth >0
		m.lnMaxFieldDefWidth = m.lnMaxFieldDefWidth + 20
	Else
		m.lnMaxFieldDefWidth = 0
	EndIf 
	
		With .isxList
			.ColumnWidths  = '0,' + Alltrim(Str(m.lnMaxWidth)) + ',' + Alltrim(Str(m.lnMaxFieldDefWidth))
			.ColumnCount   = 3
			.RowSource	   = 'Thisform.aItems'
			.RowSourceType = 5
			.Requery()
			*** JRN 9/11/2011 : Compute height, up to 24 at a time
			m.lnTotalWidth = m.lnMaxWidth + m.lnMaxFieldDefWidth + Iif(.ListCount > m.toISX.ISXOptions.DisplayCount, Sysmetric(5), 0)
			m.lnHeight	 = Min(.ListCount, m.toISX.ISXOptions.DisplayCount) *(Fontmetric(1, .FontName, .FontSize) + 2) + 4
			.Move( 0, 0, m.lnTotalWidth, m.lnHeight )
			If .ListCount < 6
				.Height = m.lnHeight
			Endif

		Endwith
		.Width	= m.lnTotalWidth
		.Height	= .isxList.Height

	*-----------------------------------------------------------------
	* LDZ: Try to find this variable name in the list of variables we
	* assembled above. If we find it, we select this entry and save
	* what has been entered so far.
	*-----------------------------------------------------------------
	Local lnIndex
	m.lcInput = m.toISX.cName
	.cSearchString = m.lcInput
	.isxlist.ListIndex = 1
	Do Case
		Case Len(m.lcInput) = 0
		Case This.oISXOptions.Filtering
			This.SetFilter(m.lcInput)
		Otherwise 
		 	m.lnIndex = At( ":"+Upper(m.lcInput), .cVarString )
			If m.lnIndex == 0
				.isxlist.ListIndex = 0
			Else
				.isxlist.ListIndex = (m.lnIndex/129) + 1
			Endif
	Endcase

	*-----------------------------------------------------------------
	* The original version of the following few code blocks has been 
	* kindly provided by Louis D. Zelus. I've modified it to match the
	* rest of the code here. The purpose is to simulate a behavior
	* in VB. If the variable is inserted via ALT+I, everything already
	* typed is used to position the list and if the already entered
	* parts are sufficient to uniquely identify the variablem it's
	* inserted without displaying the popup at all. All blocks based
	* on his code start with LDZ.
	*-----------------------------------------------------------------
	
	*-----------------------------------------------------------------
	* LDZ: If a variable name has been entered, we highlight it in the
	* edit window.
	*-----------------------------------------------------------------
	Local lnStartPos, lnEndPos
	If Len(m.lcInput) > 0
		m.lnEndPos = GetFileCursorPos( .nWHandle )
		m.lnStartPos = m.lnEndPos - Len(m.lcInput)
		_EdSelect( .nWHandle, m.lnStartPos, m.lnEndPos )
	Endif

	*-----------------------------------------------------------------
	* LDZ: If there's no second instance of this start, accept it 
	* immediately without displaying the popup. The full variable name
	* is inserted with the proper case at the current position 
	* replacing the selection.
	*-----------------------------------------------------------------
	Do Case
		Case Len (m.lcInput) = 0
	
		Case Not m.toISX.ISXOptions.SelectIfOne
			This.lScrolled = .T.
	
		Case m.toISX.ISXOptions.Filtering
			If Alen (This.aItems, 1) = 1
				InsertText( .oEditBox, "", , "R" )
				InsertText( .oEditBox, .isxList.List[.isxList.ListIndex,2] )
				Return .F.
			Else
				This.lScrolled = .T.
			Endif
	
		Case At ( ':' + Upper (m.lcInput), .cVarString, 2 ) == 0		;
				And Not m.lnIndex == 0
			InsertText( .oEditBox, "", , "R" )
			InsertText( .oEditBox, .isxList.List[.isxList.ListIndex,2] )
			Return .F.

		Otherwise
			This.lScrolled = .T.
	Endcase



	*!* * Removed 8/1/2012 / JRN
	*!* Local lnLeft, lnTop
	*!* If not GetCaretPosition( .nWHandle, @lnTop, @lnLeft )
	*!* 	Return .F.
	*!* Endif

	*!* *-----------------------------------------------------------------
	*!* * As we position the popup BELOW the current line, we need to 
	*!* * know the height of this line in pixels.
	*!* *-----------------------------------------------------------------
	*!* Local lnLineHeight, lnAvgCharWidth, lcFontName, lnFontSize
	*!* If not WGetFontInfo( .nWHAndle, @lcFontName, @lnFontSize )
	*!* 	Return .F.
	*!* Endif
	*!* lnLineHeight = FontMetric( 1, m.lcFontName, m.lnFontSize )
	*!* lnAvgCharWidth = FontMetric(6,m.lcFontName,m.lnFontSize)

	*!* *-----------------------------------------------------------------
	*!* * We make sure that the popup doesn't move below the VFP window to
	*!* * keep it visible all the time. If it doesn't fit into the area 
	*!* * below the cursor, we move it upwards.
	*!* *-----------------------------------------------------------------
	*!* If m.lnTop + .Height + m.lnLineHeight > _Screen.Height
	*!* 	lnTop = m.lnTop - .Height
	*!* Else
	*!* 	lnTop = m.lnTop + m.lnLineHeight
	*!* Endif
	*!* .Top = m.lnTop

	*!* *------------------------------------------------------------------
	*!* * As for the height of the VFP window, we do the same for the
	*!* * width. If the popup won't fit into the VFP _Screen, we flip
	*!* * it horizontally.
	*!* *------------------------------------------------------------------
	*!* If m.lnLeft + .Width + lnAvgCharWidth > _Screen.Width
	*!* 	lnLeft = m.lnLeft - .Width
	*!* Else
	*!* 	lnLeft = m.lnLeft + lnAvgCharWidth
	*!* EndIf
	*!* .Left = m.lnLeft

	Local lnLeft, lnTop, loEditBoxForm
	m.loEditBoxForm = m.toISX.oEditBox
	Do While m.loEditBoxForm.BaseClass # 'Form'
		m.loEditBoxForm = m.loEditBoxForm.Parent
	Enddo	
	
	Do Case
		Case This.Desktop and m.loEditBoxForm.Dockable # 0
			m.lnTop  = Objtoclient (m.toISX.oEditBox, 1) + Objtoclient (m.loEditBoxForm, 1)
			If m.lnTop + 22 + m.lnHeight <= _vfp.Height - Sysmetric(9) - 12
				m.lnTop = m.lnTop + 22
			Else
				m.lnTop = Max (0, m.lnTop - .Height)
			Endif
			m.lnLeft = Objtoclient (m.toISX.oEditBox, 2) + Objtoclient (m.loEditBoxForm, 2)
		
		Case This.Desktop  and m.loEditBoxForm.Dockable = 0
			m.lnTop  = Objtoclient (m.toISX.oEditBox, 1)  + Objtoclient (m.loEditBoxForm, 1) + Sysmetric(9) + Sysmetric(4)
			If m.lnTop + 22 + m.lnHeight <= _vfp.Height - Sysmetric(9) - 12
				m.lnTop = m.lnTop + 22
			Else
				m.lnTop = Max (0, m.lnTop - .Height)
			Endif
			m.lnLeft = Objtoclient (m.toISX.oEditBox, 2) + Objtoclient (m.loEditBoxForm, 2)		;
				+ Sysmetric(3)
		
		Otherwise 
			m.lnTop  = m.loEditBoxForm.Top  + Objtoclient (m.toISX.oEditBox, 1) + Sysmetric(9)
			If m.lnTop + 22 + m.lnHeight <= _Screen.Height
				m.lnTop = m.lnTop + 22
			Else
				m.lnTop = Max (0, m.lnTop - m.lnHeight)
			Endif
			m.lnLeft = m.loEditBoxForm.Left + Objtoclient (m.toISX.oEditBox, 2)		;
				+ Sysmetric(3)
	Endcase
	
	.Left = m.lnLeft
	.Top  = m.lnTop
	
	Endwith
EndProc


*========================================================================================
* Filter the listbox
*========================================================================================
 Procedure SetFilter(tcInput, tlSearchSecondColumn)
 	Local lcMatch, lcNewVarString, llLeadingMatch, llMatch, lnCol, lnColCount, lnListIndex
 	Local lnMatchRowCount, lnPos, lnRow, lnRowCount, lnStartIndex
 	With This
 		*!* * Removed 9/22/2012 / JRN
 		*!* m.lnStartIndex = .ISXList.ListIndex
 		*!* If m.lnStartIndex > 0
 		*!* 	m.lcMatch =  .ISXList.List[m.lnStartIndex]
 		*!* Endif
 
 		m.lnRowCount = Alen(.aItemsSource, 1)
 		m.lnColCount = Alen(.aItemsSource, 2)
 		Dimension .aItems(m.lnRowCount, m.lnColCount)
 		.aItems			  = ''
 		.aItems[1, 2]	  = NOMATCHFOUND
		.ISXList.ListIndex = 1
 		m.lnMatchRowCount = 0
 		m.lcNewVarString  = ''
 		m.lnListIndex	  = 0
 		For m.lnRow = 1 To m.lnRowCount
 			If Empty(m.tcInput)
 				m.llMatch = .T.
				m.llLeadingMatch = .T.
 			Else
 				m.lnPos = Atc(m.tcInput, This.aItemsSource[m.lnRow, 2])
 				Do Case
 					Case m.lnPos = 1
 						m.llMatch		 = .T.
 						m.llLeadingMatch = .T.
 					Case m.lnPos > 0 And This.oISXOptions.FilterType = 2
 						m.llMatch = .T.
 					Case tlSearchSecondColumn
 						m.lnPos = Atc(m.tcInput, Evl(This.aItemsSource[m.lnRow, 3], ''))
 						If m.lnPos = 1 Or(m.lnPos > 0 And This.oISXOptions.FilterType = 2)
 							m.llMatch = .T.
 						Else
 							m.llMatch = .F.
 						Endif
 					Otherwise
 						m.llMatch = .F.
 				Endcase
 			Endif && Empty(m.tcInput)
 			If m.llMatch
 				m.lnMatchRowCount					  = m.lnMatchRowCount + 1
 				For m.lnCol = 1 To m.lnColCount
 					This.aItems[m.lnMatchRowCount, m.lnCol] = This.aItemsSource[m.lnRow, m.lnCol]
 				Endfor
 				If m.llLeadingMatch And m.lnListIndex = 0
 					m.lnListIndex = m.lnMatchRowCount
 					.ISXList.ListIndex = m.lnListIndex
 				Endif
 			Endif
 
 			*!* * Removed 8/29/2012 / JRN
 			*!* lnPos = Atc(m.tcInput, This.aItemsSource[lnRow, 1])
 			*!* If Empty(m.tcInput) Or lnPos = 1 Or (lnPos > 0 And This.oISXOptions.FilterType = 2)
 			*!* 	lnMatchRowCount					  = lnMatchRowCount + 1
 			*!* 	For lnCol = 1 To lnColCount
 			*!* 		This.aItems[m.lnMatchRowCount, lnCol] = This.aItemsSource[m.lnRow, lnCol]
 			*!* 	Endfor
 			*!* Endif
 		Endfor
 		m.lnMatchRowCount = Max(m.lnMatchRowCount, 1)
 		Dimension This.aItems[m.lnMatchRowCount, m.lnColCount ]
 		.ISXList.Requery()
 		.cSearchString	   = m.tcInput
 	Endwith
Endproc



*========================================================================================
* If we don't hide the popup before releasing it, the focus might not go back to the 
* edit window. This happens when we have a Data Session window docked on one side and 
* a code editing window maximized. In this case the focus switches to the datasession
* window and Aliases listbox disappears.
*========================================================================================
 Procedure Release
	This.Hide()
EndProc

 Procedure isxList.KeyPress
LParameter m.tnKeyCode, m.tnSAC
With This


	*-----------------------------------------------------------------
	* If FilterToggled key has been pressed
	*-----------------------------------------------------------------
	Local loThor As Thor_Engine Of 'C:\VISUAL FOXPRO\PROGRAMS\9.0\COMMON\Thor\Source\Thor.vcx'
	If m.tnKeyCode = Asc (This.Parent.oISXOptions.FilterToggle)
		This.Parent.oISXOptions.FilterType = Iif (This.Parent.oISXOptions.FilterType = 1, 2, 1)
		This.Parent.SetFilter (This.Parent.cSearchString)
	
		* Main Thor Engine
		m.loThor        = Execscript (_Screen.cThorDispatcher, 'Thor Engine=')
		m.loThor.SetOption ('FilterToggle', 'IntellisenseX', This.Parent.oISXOptions.FilterToggle)
	
		Nodefault
		Return
	Endif
		
	*-----------------------------------------------------------------
	* If the Up or Down Arrow has been pressed, we do nothing, but 
	* remember that the user scrolled in the list, because this acti-
	* vates the enter key.
	*-----------------------------------------------------------------
	Local llScrolled
	*** JRN 9/11/2011 : Added 3 (PgDn) and 18 (PgUp)
	If m.tnSAC == 0 and InList( m.tnKeyCode, 5, 24, 3, 18 )
		.Parent.lScrolled = .T.
		Return
	Endif
	m.llScrolled = .Parent.lScrolled
	.Parent.lScrolled = .F.

	*-----------------------------------------------------------------
	* Determines whether a name qualifier has been entered.
	*-----------------------------------------------------------------
	Local llQualifier
	m.llQualifier = .F.
	If m.tnSAC == 0 and Between(m.tnKeyCode,Asc("a"),Asc("z"))
		m.llQualifier = .T.
	Endif	
	If m.tnSAC == 1 and Between(m.tnKeyCode,Asc("A"),Asc("Z"))
		m.llQualifier = .T.
	Endif	
	If m.tnSAC == 0 and Between(m.tnKeyCode,Asc("0"),Asc("9"))
		m.llQualifier = .T.
	Endif	
	If m.tnSAC == 1 and m.tnKeyCode == Asc("_")
		m.llQualifier = .T.
	Endif	
	
	*-----------------------------------------------------------------
	* If a qualifier has been entered, we insert the character into
	* the current edit window. We also perform an incremental search
	* on the Text being inserted.
	*-----------------------------------------------------------------
	Local lcSearch, lnIndex
	If m.llQualifier
		m.lcSearch = .Parent.cSearchString + Chr(m.tnKeyCode)
	Endif
	
	*-----------------------------------------------------------------
	* BACKSPACE deletes the last character.
	*-----------------------------------------------------------------
	If m.tnSAC == 0 and m.tnKeyCode == 127
		If Len(.Parent.cSearchString) > 0
			m.lcSearch = .Parent.cSearchString
			m.lcSearch = Left( m.lcSearch, Len(m.lcSearch)-1 )
			m.llQualifier = .T.
		Endif
	Endif

	*-----------------------------------------------------------------
	* Now that we handled BACKSPACE, we can update the variable name
	* in the edit window.
	*-----------------------------------------------------------------
	If m.llQualifier
		InsertText( .Parent.oEditBox, m.lcSearch, , "RH" )
		If This.Parent.oISXOptions.Filtering
			This.Parent.SetFilter(m.lcSearch)
		else
		m.lnIndex = At( ":"+Upper(m.lcSearch), .Parent.cVarString )
		If m.lnIndex == 0
			.ListIndex = 0
		Else
			.ListIndex = (m.lnIndex/129) + 1
		Endif
		EndIf 
		.Parent.cSearchString = m.lcSearch
		NoDefault
		Return
		
	Endif
	
	*-----------------------------------------------------------------
	* The following flags determine how to procede. 
	*-----------------------------------------------------------------
	Local lcTextToInsert, llResendKey, llClearInput
	m.lcTextToInsert = ""
	m.llResendKey = .T.
	m.llClearInput = .F.
	Do Case
	
	*-----------------------------------------------------------------
	* If TAB has been pressed, insert the current selection and 
	* release the popup
	*-----------------------------------------------------------------
	Case m.tnSAC == 0 and m.tnKeyCode == 9 and .ListIndex > 0
		m.lcTextToInsert = .List[.ListIndex,2]
		m.llResendKey = .F.
		m.llClearInput = .T.
	
	*-----------------------------------------------------------------
	* If ENTER has been pressed after the user made a selection with
	* the arrow keys, we insert the current selection and release the 
	* popup, because after scrolling the user has the feeling of using
	* a plain listbox where enter performs a selection.
	*-----------------------------------------------------------------
	Case     m.tnSAC == 0 ;
	     and m.tnKeyCode == 13 ;
	     and .ListIndex > 0
		m.lcTextToInsert = .List[.ListIndex,2]
		m.llResendKey = .F.
		m.llClearInput = .T.
	
	*-----------------------------------------------------------------
	* Several keys insert the current selection plus the typed 
	* character and release the popup. These are usually keys that 
	* directly follow a variable name.
	*-----------------------------------------------------------------
	*** JRN 7/26/2012 : Created separate list of keystrokes for which
	* a space is to be added after
	Case Between(m.tnKeyCode,0,255) ;
			and Chr(m.tnKeyCode) $ POPUP_TERMINATION_CHARS + EXTRA_SPACE_LIST ;
			and .ListIndex > 0
		m.lcTextToInsert = .List[.ListIndex,2]
		*** JRN 7/26/2012 : add the extra space
		If Chr(m.tnKeyCode) $ EXTRA_SPACE_LIST 
			m.lcTextToInsert = m.lcTextToInsert + ' '
		EndIf 
		m.llClearInput = .T.
	
	*-----------------------------------------------------------------
	* If ESC has been pressed, the text is unselected.
	*-----------------------------------------------------------------
	Case m.tnSAC == 0 and m.tnKeyCode == 27
		m.llResendKey = .F.
		m.llClearInput = .T.
	
	*-----------------------------------------------------------------
	* terminate the popup for any other key and leave the text.
	*-----------------------------------------------------------------
	Otherwise
	Endcase
	
	*-----------------------------------------------------------------
	* If the currently entered Text should be deleted, insert an empty
	* string using the replace option. Insert text afterwards.
	*-----------------------------------------------------------------
	If m.llClearInput
		InsertText( .Parent.oEditBox, "", , "R" )
	Else
		SetFileCursorPos(.Parent.oEditBox, .Parent.oEditBox.SelStart)
	Endif
	Do Case
		Case Empty( m.lcTextToInsert )
			
		Case  m.lcTextToInsert == NOMATCHFOUND
			InsertText( .Parent.oEditBox, This.Parent.cSearchString )
		Otherwise
			InsertText( .Parent.oEditBox, m.lcTextToInsert )
	Endcase	
	*-----------------------------------------------------------------
	* Close the form.
	*-----------------------------------------------------------------
	NoDefault
	Thisform.Release()
	
	*-----------------------------------------------------------------
	* And repeat the keystroke if necessary
	*-----------------------------------------------------------------
	Local lcKey
	
	Local lcKey
	If m.llResendKey
		m.lcKey = GetKeyLabel ( m.tnKeyCode, m.tnSAC )
		If Not Empty (m.lcKey)
			If Len (m.lcKey) # 1
				m.lcKey = '{' + m.lcKey + '}'
			Endif
		Endif
	Endif
	Clear Typeahead
	Keyboard '{END}' + Evl(m.lcKey, '')
		
Endwith
EndProc


*====================================================================
* Double-clicking is the same as TAB.
*====================================================================
 Procedure isxList.DblClick

	Clear TypeAhead
	Keyboard "{Tab}" Plain

EndProc


 EndDefine



*========================================================================================
* VFP 6: Returns a specific word in a string
*========================================================================================
Function X6_GetWordNum
LParameter m.tcString, m.tnWord, m.tcDelimiter

	Local lcString, lcDelimiter, lnWord, laWords[1], lnFound, lcWord
	
	If Vartype(m.tcDelimiter) == "C"
		m.lcDelimiter = m.tcDelimiter
	Else
		m.lcDelimiter = Chr(9)+Chr(32)
	EndIf 
	m.lcString = Chrtran(m.tcString,m.lcDelimiter,Replicate(Chr(13),Len(m.lcDelimiter)))
	m.lnFound = 0
	m.lcWord = ""
	For m.lnWord = 1 to ALines(laWords,m.lcString)
		If not Empty(m.laWords[m.lnWord])
			m.lnFound = m.lnFound + 1
			If m.lnFound == m.tnWord
				m.lcWord = m.laWords[m.lnWord]
				Exit
			EndIf 
		EndIf 
	EndFor 

Return m.lcWord


*========================================================================================
* VFP 6: Returns a list of all defines
*========================================================================================
 Procedure X6_AProcInfo
LParameter m.taArray, m.tcFile

	Local laLines[1], lnLine, lnFound
	
	m.lnFound = 0
	For m.lnLine = 1 to ALines(laLines,FileToStr(m.tcFile))
		If Upper(X6_GetWordNum(m.laLines[m.lnLine],1)) == "#DEFINE"
			m.lnFound = m.lnFound + 1
			Dimension m.taArray[m.lnFound,3]
			m.taArray[m.lnFound,1] = X6_GetWordNum(m.laLines[m.lnLine],2)
			m.taArray[m.lnFound,3] = "Define"
		EndIf 
	EndFor 

Return m.lnFound


*========================================================================================
* Get object with Thor Options
*========================================================================================
 Procedure GetISXOptions
	Return Execscript (_Screen.cThorDispatcher, 'Thor_Proc_GetISXOptions') 
Endproc


 Define Class ReadProviders As Session

	Procedure Init
		Set DataSession To 1
	Endproc

 	Procedure Run (toISX, tcProviders)
		Local laProvider[5, 2], lnProvider, lnRequest

		m.laProvider		 = ''
		m.laProvider[1, 1] = 'VAR'
		m.laProvider[1, 2] = 'CP_Variables'
		m.laProvider[2, 1] = 'DEFINE'
		m.laProvider[2, 2] = 'CP_Defines'
		m.laProvider[3, 1] = 'TABLE'
		m.laProvider[3, 2] = 'CP_Tables'
		m.laProvider[4, 1] = 'OBJ'
		m.laProvider[4, 2] = 'CP_Objects'
		m.laProvider[5, 1] = 'THOR'
		m.laProvider[5, 2] = 'CP_THOROBJECTS'
	
		For m.lnRequest = 1 To Alines (laRequest, Chrtran (m.tcProviders, ',', Chr(13) + Chr(10)), .T.)
			For m.lnProvider = 1 To Alen (m.laProvider, 1)
				If Upper (Alltrim (laRequest[m.lnRequest])) == m.laProvider[m.lnProvider, 1]
					m.toISX.nCount = 0
					Dimension m.toISX.aList[1]
					m.toISX.aList = ''
					&laProvider[m.lnProvider, 2] (m.toISX)
					If m.toISX.nCount > 0
						Return
					Endif
				Endif
			Endfor
		Endfor
		Return
	Endproc
 	
Enddefine
