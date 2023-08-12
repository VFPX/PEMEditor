#Define CR Chr(13)
#Define LF Chr(10)

Lparameters lcThorFolder, lcApplicationPath, lcVersion

Local loTool As Object
Local lcPRG, lcText

Scan For (Not Empty (prgName)) and not ThorRepo
	Scatter Name loTool Memo
	If Empty(loTool.Category)
		* Main Thor Engine
		loThor        = Execscript(_Screen.cThorDispatcher, "Thor Engine=")
		loThor.RemoveTool('Thor_Tool_PEME_' + loTool.prgName)
	Else
		DefinePRG (loTool, lcApplicationPath, lcVersion)
	Endif
Endscan

* PEMEditor_StartIDETools
lcText = Get_Setup_IDETools (lcApplicationPath)
lcPRG  = lcThorFolder + '\PEMEditor_StartIDETools'
Erase (lcPRG + '.*')
Strtofile (lcText, lcPRG + '.PRG',  0)

*!* * Removed 02/23/12 
*!* * PEMEditor_Cut
*!* lcText = Get_Cut (lcApplicationPath)
*!* lcPRG  = lcThorFolder + '\PEMEditor_Cut'
*!* Erase (lcPRG + '.*')
*!* Strtofile (lcText, lcPRG + '.PRG',  0)

*!* * PEMEditor_Copy
*!* lcText = Get_Copy (lcApplicationPath)
*!* lcPRG  = lcThorFolder + '\PEMEditor_Copy'
*!* Erase (lcPRG + '.*')
*!* Strtofile (lcText, lcPRG + '.PRG',  0)

*!* * PEMEditor_Paste
*!* lcText = Get_Paste (lcApplicationPath)
*!* lcPRG  = lcThorFolder + '\PEMEditor_Paste'
*!* Erase (lcPRG + '.*')
*!* Strtofile (lcText, lcPRG + '.PRG',  0)

Endproc



Procedure DefinePRG (loTool, lcPath, lcVersion)

	Local lcPRGText
	loThorInfo = Execscript(_Screen.cThorDispatcher,'Thor Register=')

	With loThorInfo
		* Required
		*	.Prompt		   = loTool.PrgPrompt
		.AppID         = 'PEMEditor'
		.Prompt		   = loTool.ToolPrmpt
		.Description   = Evl(loTool.Descript, loTool.Summary)
		.prgName       = 'Thor_Tool_PEME_' + loTool.prgName
		.AppName       = 'PEMEditor.APP'
		.Classes       = 'loTools = Tools of PEME_Tools|https://github.com/VFPX/Thor/blob/master/Docs/Thor_tools_object.md';
			+ ', loEditorWin = EditorWin of PEME_EditorWin|https://github.com/VFPX/Thor/blob/master/Docs/Thor_editorwindow_object.md'

		TEXT to .Code noshow textmerge
If Type('_Screen.cThorDispatcher') = 'C'
	Execscript (_Screen.cThorDispatcher, 'PEMEditor_StartIDETools')
	_oPEMEditor.oUtils.<<Alltrim(loTool.Code)>>
Else
	Messagebox("Thor is not active; this tool requires Thor", 16, "Thor is not active", 0)
Endif
		ENDTEXT

		* Optional
		.StatusBarText = Evl(loTool.StatusBar, loTool.Summary)
		.CanRunAtStartUp  = loTool.StartUp
		.Summary       = loTool.Summary

		* For public tools, such as PEM Editor, etc.
		.Source		   = 'IDE Tools'
		.Category      = Trim(loTool.Category)
		*	.SubCategory   = Trim(loTool.SubCat)
		.Version	   = lcVersion
		.Sort		   = loTool.SortOrder && the sort order for all items from the same .Source
		*	.MinorSort	   = loTool.MinorSort && the sort order for all items from the same .Source
		If Not Empty(loTool.WebPage)
*** DH 2018-02-07: commented out because simple concatenation link this may not work
***			.Link      = 'http://vfpx.codeplex.com/wikipage?title=' + Strtran(Trim(loTool.WebPage), ' ', '%20')
		Endif

		If loTool.Optioncls
			.OptionClasses = GetOptionClassNames()
			.ClassDefinitions = GetOptionClasses()		
		EndIf 
		
		.OptionTool =  Trim(loTool.OptionTool)
		
		.Plugins = loTool.PlugIns

		
		.Register()

	Endwith

Endproc


Procedure Get_Setup_IDETools (lcApplicationPath)

	Local lcPRGText
	TEXT To lcPRGText Noshow Textmerge
Local loPEMEditor

If 'O' # Vartype (_oPEMEditor)
	Release _oPEMEditor
	Public 	_oPEMEditor

	loPEMEditor = CreateObject('PEMEditor')
	_oPEMEditor = loPEMEditor.Start()
Endif

Return

EndProc


Define Class PEMEditor as Session

	Procedure Start
		Return Newobject('PEMEditor_Main', 'PEME_Main.VCX', '<<Addbs(lcApplicationPath)>>PEMEditor.APP')
	EndProc

EndDefine
	ENDTEXT
	Return lcPRGText

Endproc


Procedure Get_Cut (lcApplicationPath)

	Local lcPRGText
	TEXT To lcPRGText Noshow Textmerge
ExecScript(_Screen.cThorDispatcher, 'PEMEditor_StartIDETools')
_oPEMEditor.oEditorWin.FindWindow()
_oPEMEditor.oEditorWin.Cut()
	ENDTEXT
	Return lcPRGText

Endproc


Procedure Get_Copy (lcApplicationPath)

	Local lcPRGText
	TEXT To lcPRGText Noshow Textmerge
ExecScript(_Screen.cThorDispatcher, 'PEMEditor_StartIDETools')
_oPEMEditor.oEditorWin.FindWindow()
_oPEMEditor.oEditorWin.Copy()
	ENDTEXT
	Return lcPRGText

Endproc


Procedure Get_Paste (lcApplicationPath)

	Local lcPRGText
	TEXT To lcPRGText Noshow Textmerge
ExecScript(_Screen.cThorDispatcher, 'PEMEditor_StartIDETools')
_oPEMEditor.oEditorWin.FindWindow()
_oPEMEditor.oEditorWin.Paste()
	ENDTEXT
	Return lcPRGText

Endproc


Procedure GetOptionClassNames
	Return 'clsSourceControl, clsChangeClassVFPDialog, clsResizeWindow, clsCreateLocals, clsBeautifyX'
EndProc 

Procedure GetOptionClasses
	Local lcText
	Text to lcText NoShow
Define Class clsSourceControl As Custom

	Tool		  = 'Source Control'
	Key			  = 'Source Control'
	Value		  = .F.
	EditClassName = 'clsSourceControl of Thor_Options_SourceControl.VCX'

Enddefine


Define Class clsChangeClassVFPDialog As Custom

	Tool		  = 'ReDefine Parent Class'
	Key			  = 'ReDefine Parent Class'
	Value		  = .F.
	EditClassName = 'clsChangeClassVFPDialog of Thor_options_ChanceClassVFPDialog.VCX'

Enddefine


Define Class clsResizeWindow As Custom

	Tool		  = 'Move/Resize Code Window'
	Key			  = 'Move/Resize Code Window'
	Value		  = ''
	EditClassName = 'clsResizeWindow of Thor_Options_ResizeWindow.VCX'

Enddefine


Define Class clsCreateLocals As Custom

	Tool		  = 'Create Locals'
	Key			  = 'Create Locals'
	Value		  = ''
	EditClassName = 'clsCreateLocals of Thor_Options_CreateLocals.VCX'

Enddefine


Define Class clsBeautifyX As Custom

	Tool		  = 'BeautifyX'
	Key			  = 'BeautifyX'
	Value		  = ''
	EditClassName = 'clsBeautifyX of Thor_Options_BeautifyX.VCX'

Enddefine
	
	EndText
	Return lcText
EndProc	
