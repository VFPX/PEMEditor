Lparameters loPrefs

Local loKey As Object
Local loThor As Object
Local lcKey, lcPrefName, lcPRG

loThor = Execscript (_Screen.cthordispatcher, 'Thor Engine=')

Select  *								;
	From ThorSource						;
	Where (Not Empty (PRGName))			;
	And (Not Empty (Prefname))		;
	Into Cursor crsr_ThorTools Readwrite

Scan
	lcPrefName = Trim (Prefname)
	lcKey	   = Evaluate ('loPrefs.' + lcPrefName)
	Do Case
		Case Empty (lcKey)

		Case '*' $ lcKey

		*	Case '+' $ lcKey And 0 # Val(Substr(lcKey, At('+', lcKey) + 2))

		Otherwise
			loKey = loThor.GetKeyValueFromKeyName (lcKey, .T.)
			lcPRG = Forceext ('Thor_Tool_PEME_' + Trim (PRGName), 'PRG')
			loThor.AddToolHotKey (lcPRG, loKey.KeyCode, loKey.nShifts, .T.)
	Endcase
Endscan

Return
