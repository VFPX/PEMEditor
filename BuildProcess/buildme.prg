* Release references to PEMEditor.app since we'll be rebuilding it.

if type('_oPEMEditor') = 'O'
	_oPEMEditor.Release()
endif type('_oPEMEditor') = 'O'
release _oPEMEditor

* Update the version information in PEMEditorVersion.h.

text to lcVersion noshow textmerge
#define	ccPEMEVERSION [PEM Editor - <<pcVersion>> - <<cmonth(pdVersionDate)>> <<day(pdVersionDate)>>, <<year(pdVersionDate)>> - <<strtran(pcVersionDate, '-')>>]
endtext
strtofile(lcVersion, 'downloads\source\pemeditorversion.h')

return
