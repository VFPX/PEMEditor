#Define InternalSnippets 'Case, New, DD, F, L, LG, Sel, UC, Edit, From, DC, IT, CT'
#Define ccCR Chr(13)
#Define ccLF Chr(10)

Define Class DoubleHash As Session

	cSnippetPath 		  = []
	nRows		 		  = 0
	cOriginalPath 		  = ''
	oPrefs				  = .Null.

	Dimension aPMDList(1)

	Procedure Init
		Lparameters loPrefs, lcSnippetPath
		With This
			.oPrefs = loPrefs
			.cSnippetPath = lcSnippetPath
		Endwith
		Set DataSession To 1
	Endproc

	Procedure Release
		This.oPrefs = .Null.
	Endproc

	Procedure Destroy
		This.Release()
	Endproc


	Procedure HandleDoubleHash
		Lparameters lcKeyWord, lcParameters
		Local lcResult, lcSnippetPRG 

		If [?] = lcKeyWord Or 'HELP' == lcKeyWord
			Return This.Hash_Help()
		Endif

		lcSnippetPRG = 'PEME_Snippet_' + lcKeyWord

		lcResult = .F.

		If _Screen.lThorDebugMode
			lcResult = Execscript(_Screen.cThorDispatcher, lcSnippetPRG, lcParameters, lcKeyWord)
		Else
			Try
				lcResult = Execscript(_Screen.cThorDispatcher, lcSnippetPRG, lcParameters, lcKeyWord)
			Catch To loException
				This.ShowErrorMsg(loException, lcSnippetPRG)
			Endtry
		Endif

		* 	lcResult = Nvl(lcResult, '')
		Return lcResult

	Endproc


	Procedure Hash_Help

		#Define CR   Chr(13)
		#Define LF   Chr(10)
		#Define CRLF Chr(13) + Chr(10)
		#Define Tab  Chr(9)

		Local loForm As [PEME_FRMBROWSER]
		Local laFiles[1], laFolders[1], lcDetail, lcFileName, lcHTML, lcHTMLHeader, lcInsertText, lcLastName
		Local lcPlugIns, lcResult, lcSamples, lcSummary, lcSyntax, lcThisRow, lnCount, lnCounter, lnI
		Local laInternals(1), lcInternal, lcToolFolder

		This.nRows = 0

		*** JRN 2010-10-28 : away we go
		lcHTMLHeader = _oPEMEditor.oUtils.oIDEx.GetHTMLHeader()
		TEXT To lcHTML Noshow Textmerge
<<lcHTMLHeader>>
</head>
<body >
<table border="1">
	<tr>
		<th>Snippet</th>
		<th>Description</th>
	</tr>
		ENDTEXT

		* User defined Plug-Ins
		This.AddFolderSnippets (Curdir())
		lnCount = Alines (laFolders, Set ('Path'), .T., ';', ',')
		For lnI = 1 To lnCount
			If Not Empty (laFolders (lnI))
				This.AddFolderSnippets (laFolders (lnI))
			Endif
		Endfor
		lcToolFolder = Execscript (_Screen.cThorDispatcher, 'Tool Folder=')
		This.AddFolderSnippets (lcToolFolder + 'My Tools\')
		This.AddFolderSnippets (lcToolFolder)

		Alines(laInternals, InternalSnippets, .T., ',')
		For lnI = 1 To Alen(laInternals)
			lcInternal = Alltrim(laInternals(lnI))
			This.AddRow (PEME_Snippet_&lcInternal._Internal (), '', .T.)
		Endfor

		Asort (This.aPMDList, 8, -1, 0, 1)
		lcLastName = This.aPMDList(1, 1)
		For lnI = 2 To This.nRows
			This.aPMDList (lnI, 7) = Upper (lcLastName) == Upper (This.aPMDList (lnI, 1))
			lcLastName = This.aPMDList (lnI, 1)
		Endfor

		For lnI = 1 To This.nRows
			If This.aPMDList (lnI, 7)
				Loop
			Endif

			*!* * Removed 12/21/2011 
			*!* lcSyntax  = '<b>' + This.oPrefs.cDynamicSnippetMarker			;
			*!* 	+ ' ' + This.aPMDList (lnI, 1) + '</b>'			;
			*!* 	+ ' ' + This.aPMDList (lnI, 2)
			lcSyntax  = '<b> ' + This.aPMDList (lnI, 1) + '</b>'			;
			 	+ ' ' + This.aPMDList (lnI, 2)
			lcSummary = This.aPMDList (lnI, 3)

			lcDetail = This.aPMDList (lnI, 4)
			If Not Empty (lcDetail)
				lcSyntax = '<a href="#PMD' + Transform (lnI) + '">' + lcSyntax + '</a> '
			Endif
			TEXT To lcThisRow Noshow Textmerge
	<tr>
		<td> <<lcSyntax>> </td>
		<td> <<lcSummary>> </td>
	</tr>
			ENDTEXT
			lcHTML = lcHTML + lcThisRow + CRLF
		Endfor

		lcHTML = lcHTML + '</table>' + CRLF

		lcSummary = '<a href="http://pemeditor.blogspot.com/">Go to the PEM Editor blog</a>'
		lcHTML = lcHTML + '<br/>' + lcSummary + '<br/>' + CRLF

		lcHTML = M.lcHTML + "<table border=0>"+CRLF
		For lnI = 1 To This.nRows
			If This.aPMDList (lnI, 7)
				Loop
			Endif
			*!* * Removed 12/21/2011 
			*!* lcName  = '<b>' + This.oPrefs.cDynamicSnippetMarker ;
			*!* 	+ '&nbsp' + This.aPMDList (lnI, 1) + '</b>'
			lcName  = '<b> ' + This.aPMDList (lnI, 1) + '</b>'
			lcSyntax  = lcName + ' ' + This.aPMDList (lnI, 2)
			*	lcSyntax = lcSyntax + Iif (This.aPMDList (lnI, 6), ' (Built-In)', ' (See '  + Lower (This.aPMDList (lnI, 5)) + ')')

			lcDetail = This.aPMDList (lnI, 4)
			If Not Empty (lcDetail)
				lcName = '<a name="PMD' + Transform (lnI) + '">' + M.lcName + '</a> '
				TEXT To lcThisRow Noshow Textmerge
	<tr><td valign="top"><div class="bookmark">
		<<lcName>>
	</div></td>
	<td><div class="bookmark-content">
		<b>Syntax: <<lcSyntax>></b><br/>
		<<lcDetail>>
	</div></td></tr>
				ENDTEXT
				lcHTML = lcHTML + lcThisRow + CRLF
			Endif
		Endfor
		lcHTML = lcHTML + "</table>"+ CRLF

		lcHTML = lcHTML + '</body>' + CRLF

		Do Form PEME_Help_Internal With [Dynamic Snippets], lcHTML, 600, 500

		Return '' && Handled it .. replace PMD with blank

	Endproc


	Procedure AddFolderSnippets (lcFolder)
		Local laFiles[1], lcFileName, lnCount, lnCounter, loResult
		If Directory(lcFolder,1)
			lnCount	= Adir (laFiles, Addbs (lcFolder) + [PEME_Snippet*.PRG])
			If lnCount > 0
				Asort (laFiles, 1, -1, 0, 1)
				For lnCounter = 1 To lnCount
					lcFileName = laFiles (lnCounter, 1)
					loResult   = Execscript(_Screen.cThorDispatcher, lcFolder + lcFileName)
					This.AddRow (loResult, lcFileName)
				Endfor
			Endif
		Endif
	Endproc

	Procedure AddRow (loRow, lcFileName, llBuiltIn)
		Local lnRow

		If 'O' # Vartype (loRow)
			Return
		Endif

		lnRow	   = This.nRows + 1
		This.nRows = lnRow

		Dimension This.aPMDList (lnRow, 8)
		With loRow
			This.aPMDList (lnRow, 1) = .Name
			This.aPMDList (lnRow, 2) = .Syntax
			This.aPMDList (lnRow, 3) = .Summary
			This.aPMDList (lnRow, 4) = .Detail
			This.aPMDList (lnRow, 5) = lcFileName
			This.aPMDList (lnRow, 6) = llBuiltIn
			This.aPMDList (lnRow, 8) = .Name + Str (lnRow, 4, 0)
		Endwith
	Endproc


	Procedure InsertFileName (lcFileName)
		Return Strtran (lcFileName, ' ', '%20')
	Endproc


	Procedure GoToDefinition(lcPRG, lcText, lcTextBefore, lcTextAfter)
		Return ExecScript(FileToStr(lcPRG), lcText, lcTextBefore, lcTextAfter)
		* Do (lcPRG) With lcText, lcTextBefore, lcTextAfter to llResult
	Endproc

	Procedure ShowErrorMsg
		Lparameters loException, lcSnippetFileName

		Messagebox ('Error: ' + Transform (loException.ErrorNo)	+ ccCR + ccLF +	 ;
			'Message: ' + loException.Message 					+ ccCR + ccLF +	 ;
			'Procedure: ' + loException.Procedure 				+ ccCR + ccLF +	 ;
			'Line: ' + Transform (loException.Lineno) 			+ ccCR + ccLF +	 ;
			'Code: ' + loException.LineContents									 ;
			, 0        + 48, 'Dynamic Snippet Error: ' + Justfname(lcSnippetFileName))
	Endproc


Enddefine
