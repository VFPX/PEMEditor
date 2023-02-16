*** DougHennig 12/23/2008: removed all "new" since classes, VCXs and APP were renamed

********************************************************************
*** Name.....: PROGRAM2
*** Author...: Marcia G. Akins
*** Date.....: 01/15/2007
*** Notice...: Copyright (c) 2007 Tightline Computers, Inc
*** Compiler.: Visual FoxPro 09.00.0000.3504 for Windows
*** Function.: Modeless Replacement for the edit property/method dialog -
*** ...................: also handles removing member data when you remove the member
*** Returns..: Logical
********************************************************************

#include EditPropertyDialog.H

Lparameters lcThor

Local lcPath, ;
	lcMenuContextCode, ;
	lcAppPath, ;
	lcPropertyWindowContextMenuCode, ;
	lcEditPropertyMenuCode, ;
	lcGetStartPemEditorCode , ;
	lnSelect,					;
	lcRegisteredMessage

lcPath = Addbs(Justpath(Sys(16)))
lcAppPath = Forcepath('PEMEditor.APP', lcPath)

If Evl(lcThor, '') = 'Thor'
	RegisterWithThor (lcAppPath, lcPath, .T.)
	Return
EndIf

If PEME_IsVersion9()
	ExecScript(_Screen.cThorDispatcher, 'Thor_Tool_PEME_ManagePEMEditorAccess')
	lcRegisteredMessage = ccLOC_DIALOG_REGISTERED
Else
	lcRegisteredMessage = ccLOC_DIALOG_REGISTERED_V8
Endif

****************************************************************
*** JRN 12/11/2008 : Create StartPEMEditor.PRG
Erase ("StartPEMEditor.*")
lcGetStartPemEditorCode = GetStartPemEditorCode(lcAppPath)
Strtofile(lcGetStartPemEditorCode , "StartPEMEditor.PRG")
****************************************************************

Messagebox(lcRegisteredMessage , MB_OK + MB_ICONINFORMATION)

RegisterWithThor (lcAppPath, lcPath)
Return


****************************************************************
****************************************************************


Procedure GetStartPemEditorCode(lcAppPath)

	Local lcCode

	TEXT to lcCode noshow textmerge
Lparameters toParameter

* StartPEMEditor() - brings up the normal PEMEditor form.
* StartPEMEditor(2) - brings up the Document TreeView form.
* StartPEMEditor(.Null.) - instantiates the PEM Editor object (_oPEMEditor)
*                 , enabling the IDE features of PEM Editor, without opening 
*                 either of the forms (This is a new feature in version 6.50.)

Local loPEMEditor

If 'O' # Vartype (_oPEMEditor)
	Release _oPEMEditor
	Public 	_oPEMEditor

	loPEMEditor = CreateObject('PEMEditor')
	_oPEMEditor = loPEMEditor.Start()
Endif

_oPEMEditor.oUtils.ShowForm(toParameter)



Define Class PEMEditor as Session

	Procedure Start
		Return Newobject('PEMEditor_Main', 'PEME_Main.VCX', '<<lcAppPath>>')
	EndProc
	
EndDefine 
	ENDTEXT

	Return lcCode

Endproc

****************************************************************
****************************************************************

Procedure RegisterWithThor
	Lparameters tcApp, tcPath, tlQuiet
	Local loPEMEditor, loPEMEditorSession

	If 'C' # Type('_Screen.cThorDispatcher')
		Return
	Endif

	loPEMEditorSession = Createobject('PEMEditor')
	loPEMEditor = loPEMEditorSession.Start(tcApp)
	loPEMEditor.oUtils.RegisterWithThor(tlQuiet)
	
	Release Pad _PEMEditor Of _Msysmenu
	With loPEMEditor.oPrefs
		.nToolsPlacement = 4
		.SaveValues()
	EndWith 

	lcFile = tcPath + 'PEMEditorVersion.txt'
	Erase (lcFile)
	StrToFile(loPEMEditor.oUtils.oPrefs.cFullVersion, lcFile, 0)

	Return



Define Class PEMEditor As Session

	Procedure Start (tcApp)
		Return Newobject('PEMEditor_Main', 'PEME_Main.VCX', tcApp)
	Endproc

Enddefine

