****************************************************************
*	First Section: Compile-time constants -- modify as needed  *
****************************************************************

* snippet-keyword: case insensitive, NOT in quotes
#Define Snippet_Keyword 	Edit

* delimiter between parameters, IN QUOTES; if empty, only one parameter
#Define Delimiter_Char	 	''

* minimum number of parameters to be accepted
#Define Min_Parameters      1

* maximum number of parameters to be accepted
#Define Max_Parameters      1




*****************************************************************
*	 Middle Section: Setup and cleanup code:  DO NOT CHANGE!!!  *
*****************************************************************

Lparameters lcParameters, lcKeyWord

Local loParams As Collection
Local lcParams, lnI

Do Case
		* if no parameters passed, this is a request for Help
	Case Pcount() = 0
		Return PublishHelp()

		* Only process our keyword 
	Case Pcount() = 2 And Not Upper ([Snippet_Keyword]) == lcKeyWord
		Return .F. && not mine!

	Otherwise
		lcParams = _oPEMEditor.ExtractSnippetParameters (lcParameters, Delimiter_Char, [Snippet_Keyword], Min_Parameters, Max_Parameters)
		If 'C' = Vartype (lcParams)
			Return Process (&lcParams)
		Endif
Endcase


Function CreateHelpResult (lcSyntax, lcSummaryHTML, lcDetailHTML)
	Local loResult As 'Empty'
	loResult = Createobject ('Empty')
	AddProperty (loResult, 'Name', [Snippet_Keyword])
	AddProperty (loResult, 'Syntax', Evl (lcSyntax, ''))
	AddProperty (loResult, 'Summary', Evl (lcSummaryHTML, ''))
	AddProperty (loResult, 'Detail', Evl (lcDetailHTML, ''))
	Return loResult
Endproc




*****************************************************************
****** Last Section: Custom code for this snippet				*
*****************************************************************

#Define CR   Chr(13)
#Define LF   Chr(10)
#Define CRLF Chr(13) + Chr(10)
#Define Tab  Chr(9)

* Put your code here to process the parameters passed; make sure to set parameters appropriately
* Result is one of following:
*   character string -- string to insert into edit window, replacing the snippet there (usual case, by far)
*                       if the characters '^^^' are found, they indicate where the cursor is to be placed
*   .T.              -- handled, but edit window already updated (advanced use)
*   .F.              -- not handled

Function Process
	Lparameters lcSnippetName

	Local lcFileName, lcFullPath, lcSampleFile, lnResult
	lcFileName = 'PEME_Snippet_' + lcSnippetName + '.Prg'
	lcFullPath = Fullpath (lcFileName)

	If Not File (lcFullPath)
		lcFullPath = Addbs (_oPEMEditor.oUtils.cSnippetPath) + lcFileName
		If Not File (lcFullPath)
			lcSampleFile = Addbs (_oPEMEditor.oUtils.cSnippetPath) + 'Snippet Samples\' + lcFileName
			If File (lcSampleFile)
				lnResult = Messagebox ('Sample snippet file found for ' + lcSnippetName + CRLF + CRLF + 'Open it for viewing?' ;
					  , 4 + 32)
				If lnResult = 6
					_oPEMEditor.oTools.AddMRUFile (lcSampleFile)
					Modify Command (lcSampleFile) Nowait
					Messagebox ('To activate this sample, save it into its parent folder' + CRLF + 'or into any folder in your path.', 64)
					Return ''
				Endif
			Endif
			Messagebox ('Snippet "' + lcSnippetName + '" not found', 16)
		Endif
	Endif

	_oPEMEditor.oTools.AddMRUFile (lcFullPath)
	Modify Command (lcFullPath) Nowait

	Return ''
Endfunc



***************************************************************************************************
* Publish the help for this snippet: calls function CreateHelpResult with three parameters:
*    Syntax
*    Summary
*    Full description

* Note that all have these may contain HTML tags
***************************************************************************************************

Function PublishHelp
	Local lcDetailHTML, lcSummaryHTML, lcSyntax

	lcSyntax = 'Snippet_Keyword'

	Text To lcSummaryHTML Noshow
        <b>Edits the PRG for a snippet</b>
	Endtext

	Return CreateHelpResult (lcSyntax, lcSummaryHTML, lcDetailHTML)
Endproc

