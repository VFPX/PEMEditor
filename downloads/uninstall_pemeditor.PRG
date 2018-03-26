ExecScript(_Screen.cThorDispatcher, 'Thor_Tool_Peme_ManagePEMEditorAccess')
Return

*** JRN 05/01/2012


*** JRN 03/08/2009 : extracted code from "editproperty" to perform uninstallation.
* un-install all parts of PEMEditor installed

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

#include Source\EditPropertyDialog.H
Local lcPath, ;
	lcMenuContextCode, ;
	lcAppPath, ;
	lcEditPropertyCode, ;
	lcEditPropertyMenuCode, ;
	lnSelect


* Add the records to FOXCODE.

lnSelect = Select()
Select 0
Use (_Foxcode) Again Shared Alias Foxcode Order 1

If Seek('M24460')								;
		and "PEMEDITOR.APP" $ Upper (Data)
	Delete
Endif Not Seek('M24460')

If Seek('MEDIT PROPERTY/METHOD...')				;
		And "PEMEDITOR.APP" $ Upper (Data)
	Delete
Endif Not Seek('MEDIT PROPERTY/METHOD...')

If Seek('M' + Upper(ccLOC_MENU_NEW_PROPERTY))	;
		And "PEMEDITOR.APP" $ Upper (Data)
	Delete
Endif Not Seek('M' ...

If Seek('M' + Upper(ccLOC_MENU_NEW_METHOD))		;
		And "PEMEDITOR.APP" $ Upper (Data)
	Delete
Endif Not Seek('M' ...

Use

_oPEMEditor = Null
_oDocTreeView = Null
Release Pad _PEMEditor Of _Msysmenu

****************************************************************
Select (lnSelect)
Messagebox(ccLOC_DIALOG_UNINSTALLED, MB_OK + MB_ICONINFORMATION)

Return

