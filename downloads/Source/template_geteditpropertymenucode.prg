Lparameters toParameter

Local llReturn

Try
	If 'O' # Vartype (_oPEMEditor)
		Release _oPEMEditor
		Public 	_oPEMEditor

		loPEMEditor = Createobject('PEMEditor')
		_oPEMEditor = loPEMEditor.Start()
	Endif
	_oPEMEditor.oUtils.ShowForm()
	llReturn = .T.
Catch

Endtry
****************************************************************


Return llReturn



Define Class PEMEditor As Session

	Procedure Start
		Return Newobject('PEMEditor_Main', 'PEME_Main.VCX', '<<lcAppPath>>')
	Endproc

Enddefine
