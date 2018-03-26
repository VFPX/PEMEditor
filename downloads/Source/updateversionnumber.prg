#Define VersionFile         	'PemEditorVersion.h'
#Define UpdaterVersionFile  	'..\Installation Folder\_PEMEditorVersionFile.txt'
#Define UpdaterVersionFileBeta  '..\Installation Folder - Beta\_PEMEditorVersionFileBeta.txt'

#Define CRLF        Chr(13) + Chr(10)

Local laLines[1], lcNewText, lcNewVersion, lcOldVersion, lcURL, lcUpdateVersionFile, lcVersion
Local lnPos

Alines(laLines, Filetostr(VersionFile))
lcOldVersion = Substr(laLines(1), 3)
If Occurs('.', lcOldVersion) = 3
	lnPos		 = At('.', lcOldVersion, 3)
	lcOldVersion = Left(lcOldVersion, lnPos) + Transform(Val(Substr(lcOldVersion, lnPos + 1)) + 1, '@L 99') + ' Beta'
Endif

lcVersion = Inputbox('New Version: ', '', lcOldVersion)
If Empty(lcVersion)
	Return
Endif

*********************************************************
*!* * Removed 7/31/2012 / JRN
*!* lcUrl = 'http://bit.ly/PEMEditorZip' && old
*!* lcUrl = 'http://vfpxrepository.com/dl/thorupdate/Tools/PEM_Editor/PEMEditor.zip'

If Upper('Beta') $ Upper(lcVersion)
	lcUpdateVersionFile	= UpdaterVersionFileBeta
	lcURL				= 'http://vfpxrepository.com/dl/thorupdate/Tools/PEM_Editor/Beta/PEMEditor.zip'
Else
	lcUpdateVersionFile	= UpdaterVersionFile
	lcURL				= 'http://vfpxrepository.com/dl/thorupdate/Tools/PEM_Editor/PEMEditor.zip'
Endif

*********************************************************
*!* * Removed 8/12/2012 / JRN
*!* lcNewVersion = CreateCloudVersionFile([PEM Editor w/IDE Tools], lcVersion, lcURL, lcUpdateVersionFile, 'PEM Editor Home Page', 'http://vfpx.codeplex.com/wikipage?title=PEM%20Editor%207%20with%20IDE%20Tools')
lcNewVersion = CreateCloudVersionFile([PEM Editor], lcVersion, lcURL, lcUpdateVersionFile, 'PEM Editor Home Page', 'http://vfpx.codeplex.com/wikipage?title=PEM%20Editor%207%20with%20IDE%20Tools')

* and create .H file with version information
lcNewText = '* ' + lcVersion
lcNewText = lcNewText + CRLF + ''
lcNewText = lcNewText + CRLF + '#Define cnVersion         ' + Alltrim(lcVersion)
lcNewText = lcNewText + CRLF + '#Define cdVersionDate     ' + SpellDate()
lcNewText = lcNewText + CRLF + '#Define	ccPEMEVERSION     [' + lcNewVersion + ']'

lcNewText			= lcNewText + CRLF + '#Define	ccThorVERSIONFILE [ThorVersion.txt]'
lcUpdateVersionFile	= UpdaterVersionFile

Erase(VersionFile)
Strtofile(lcNewText, VersionFile, 0)

*********************************************************

Return lcVersion
