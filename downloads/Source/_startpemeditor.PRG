lparameters tcPath, toParameter
Local loPEMEditor

If 'O' # Vartype (_oPEMEditor)
	Release _oPEMEditor
	Public 	_oPEMEditor

	loPEMEditor = CreateObject('PEMEditor')
	_oPEMEditor = loPEMEditor.Start(tcPath)
Endif

_oPEMEditor.oUtils.ShowForm(toParameter)



Define Class PEMEditor as Session

	Procedure Start (tcPath)
		Return Newobject('PEMEditor_Main', 'PEME_Main.VCX', tcPath + 'PEMEditor.APP')
	EndProc
	
EndDefine 