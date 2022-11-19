Lparameters lcPrefix, lcVersion, lcSourceURL, lcUpdateVersionFile, lcLinkPrompt, lcLink

*********************************************************
Local lcNewVersion, lcText
lcNewVersion = lcPrefix + ' - ' + Alltrim (lcVersion) + [ - ] + SpellDate() + ' - ' + Dtoc(Date(),1)

Text To lcText Noshow Textmerge
Lparameters toUpdateInfo

AddProperty(toUpdateInfo, 'AvailableVersion',	'<<lcNewVersion>>')
AddProperty(toUpdateInfo, 'SourceFileUrl', 		'<<lcSourceURL>>')
AddProperty(toUpdateInfo, 'LinkPrompt', 		'<<Evl(lcLinkPrompt, "")>>')
AddProperty(toUpdateInfo, 'Link', 				'<<Evl(lcLink, "")>>')

Execscript (_Screen.cThorDispatcher, 'Result=', toUpdateInfo)
Return toUpdateInfo

Endtext

Erase (lcUpdateVersionFile)
Strtofile (lcText, lcUpdateVersionFile, 0)
*********************************************************

Return lcNewVersion
