* Release references to PEMEditor.app since we'll be rebuilding it.

if type('_oPEMEditor') = 'O'
	_oPEMEditor.Release()
endif type('_oPEMEditor') = 'O'
release _oPEMEditor

*** TODO: why is this needed?

use Downloads\Source\editproperty.vcx Exclusive
pack
use

* Update the version information in PEMEditorVersion.h.

text to lcVersion noshow textmerge
#define	ccPEMEVERSION [PEM Editor - <<pcVersion>> - <<cmonth(pdVersionDate)>> <<day(pdVersionDate)>>, <<year(pdVersionDate)>> - <<strtran(pcVersionDate, '-')>>]
endtext
strtofile(lcVersion, 'downloads\source\pemeditorversion.h')

* Recreate EditProperty.prg from EditProperty_Template.prg and Template_StartPEMEditor.prg.

*** TODO: is this needed? If not, dump two template files
*erase editproperty.prg
*erase editproperty.fxp
*InsertPRGTextFromTemplateFiles('EditProperty_Template.prg', 'EditProperty.prg')

return


*** TODO: dump this
procedure InsertPRGTextFromTemplateFiles(tcFromPRG, tcToPRG)
local lcNewText, lcSourceCode, lcText
lcSourceCode = filetostr(forceext(tcFromPRG, 'PRG'))
do while .T.
	lcText = strextract(lcSourceCode, '<<<<', '>>>>', 1, 4)
	if empty(lcText)
		exit 
	endif empty(lcText)
	lcNewText	 = evaluate(substr(lcText, 5, len(lcText) - 8))
	lcSourceCode = stuff(lcSourceCode, at(lcText, lcSourceCode), len(lcText), lcNewText)
enddo while .T.
strTofile(lcSourceCode, forceext(tcToPRG, 'PRG')) 
