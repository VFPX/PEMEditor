Lparameters tcPRGFileName

Local lcFileName, lcPRGFileName

lcPRGFileName = tcPRGFileName + IIf ('.' $ tcPRGFileName, '', '.PRG')

lcFileName = FullPath('PEME_' + lcPRGFileName)
If File (lcFileName)
	Return (lcFileName)
Endif

lcFileName = This.cPlugInPath + lcPRGFileName
If File (lcFileName)
	Return (lcFileName)
Endif

lcFileName = This.cApplicationPath + lcPRGFileName
If File (lcFileName)
	Return (lcFileName)
Endif

Return ExecScript(_Screen.cThorDispatcher, 'Full Path=PEME_' + lcPRGFileName)
