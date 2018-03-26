*==============================================================================
* Procedure:		GoToDefinition
* Purpose:			Customized handling of GoToDefinition.  This is called if 
*                      the normal processing of Go To Definition found no matches.
* Author:			Jim Nelson
* Parameters:		
*	tcText	  		= the text being searched for
*	tcTextBefore  	= the text (on the same line) preceding the text being search for
*	tcTextAfter  	= the text (on the same line) following the text being search for
* Returns:			.T. if matches found
* Added:			12/20/2009
*
*==============================================================================
* Notes:
*	This is a sample PRG for customized uses of GoToDefinition.  It does two things:
*     [1] - If the search text is the alias of a table/cursor, or is the name of
*           a table that can be opened, that table/cursor is displayed, showing
*           the list of fields and a grid for browsing the table
*     [2] - Else, if the search text is the name of a class, that class is opened
*           (using one of the methods from _oPEMEditor.oTools
*
* Other files needed for this example (which, of course, are also customizable)
*    PEME_Table.PRG
*    PEME_LS.SC*
*    PEME_mxBrowser.VC*

Lparameters tcText, tcTextBefore, tcTextAfter

#Define CR   Chr(13)
#Define LF   Chr(10)
#Define CRLF Chr(13) + Chr(10)
#Define Tab  Chr(9)

Local lcFileName, lcNewFileName, lcOpenFilePRG

lcFileName = tcText
If Not Used (lcFileName)
	lcOpenFilePRG = [PEME_OpenTable.PRG]
	If File (lcOpenFilePRG)
		lcNewFileName = PEME_OpenTable (lcFileName) && may return a different alias
		If 'C' = Vartype (lcNewFileName)
			lcFileName = lcNewFileName
		Endif
	Endif
Endif

If Used (lcFileName)
	ShowTable (lcFileName)
	Return .T.
Else
	Return .F.
Endif


Procedure ShowTable (lcFileName)

	Local lcToolFolder, lnSelect, loFP_Form
	lnSelect = Select()
	Select (lcFileName)
	lcToolFolder		 = Execscript (_Screen.cThorDispatcher, 'Tool Folder=')
	loFP_Form			 = Newobject ('FP_FORM', lcToolFolder + 'PEME_SNIPPET_FP.VCX', '', lcFileName)
	loFP_Form.WindowType = 0
	loFP_Form.Show()
	_Screen.AddProperty (Sys(2015), loFP_Form)
	Select (lnSelect)

	Return


Endproc