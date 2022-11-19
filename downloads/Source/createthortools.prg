Lparameters lcThorFolder, lcApplicationPath, lcVersion

Local lcPRG, lcText, loTool
Scan
	Scatter Name loTool
	lcText = DefinePRG (loTool, lcApplicationPath, lcVersion)
	lcPRG  = lcThorFolder + '\Thor_Tool_' + Trim (prgname)
	Erase (lcPRG + '.*')
	Strtofile (lcText, lcPRG + '.PRG',  0)
Endscan

Endproc



Procedure DefinePRG (loTool, lcPath, lcVersion)

	Local lcPRGText
	Text To lcPRGText Noshow Textmerge
Lparameters lxParam1

****************************************************************
****************************************************************
* Standard prefix for all tools for Thor, allowing this tool to
*   tell Thor about itself.

If Pcount() = 1						  ;
		And 'O' = Vartype (lxParam1)  ;
		And 'thorinfo' = Lower (lxParam1.Class)

	With lxParam1
		.Prompt		 = '<<Trim (loTool.Prompt)>>'
		.Description = '<<Trim (loTool.Descript)>>'
		.Source		 = 'PEM Editor'
		.Version	 = '<<lcVersion>>'
		.Sort		 = <<loTool.SortOrder>>
		.SubMenu	 = <<loTool.SubMenu>>
		.Link        = 'http://pemeditor.blogspot.com/'
		.OptionTool =  '<<Trim(loTool.OptionTool)>>'
	Endwith

	Return lxParam1
Endif

****************************************************************
****************************************************************
* Normal processing for this tool begins here.    

Local loPEMEditor

If 'O' # Vartype (_oPEMEditor)
	Release _oPEMEditor
	Public 	_oPEMEditor

	loPEMEditor = CreateObject('PEMEditor')
	_oPEMEditor = loPEMEditor.Start()
Endif

_oPEMEditor.oUtils.<<Exec>>

Return

EndProc


Define Class PEMEditor as Session

	Procedure Start
		Return Newobject('PEMEditor_Main', 'PEME_Main.VCX', '<<Addbs(lcPath)>>PEMEditor.APP')
	EndProc
	
EndDefine 

	Endtext

	Return lcPRGText

Endproc