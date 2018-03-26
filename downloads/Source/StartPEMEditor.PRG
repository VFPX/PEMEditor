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
		Return Newobject('PEMEditor_Main', 'PEME_Main.VCX', 'C:\VISUAL FOXPRO\PROGRAMS\TOOLS\PEM EDITOR\CURRENT RELEASE\PEMEditor.APP')
	EndProc
	
EndDefine 