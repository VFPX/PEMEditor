Procedure RestartAfterClearAll (loThisForm, lcRestartText)

	#Define NewLine 			Chr(13) + Chr(10)
	#Define CloseButtonWidth 	Sysmetric( Iif(Version(5) >= 900 And loThisForm.Dockable # 0, 33, 28))

	* Invoked from QueryUnload to re-open a form after "CLEAR ALL"
	* loThisForm	= This
	* lcRestartText	= text to execute to re-open the form

	* Sample:  RestartAfterClearAll (This, "Do Form 'Form1'")

	* Jim Nelson 05/23/2009
	* w/modifications by Bernard Bout 5/25/2009

	If "O" = Vartype(loThisForm)
		* Wrong release type?
		If loThisForm.ReleaseType # 1
			Return
		Endif

		* Mouse over Close 'X' Button?
		If Between (Mcol (loThisForm.Name,3), loThisForm.Width - CloseButtonWidth, loThisForm.Width) And Mrow (loThisForm.Name,3) = -1
			* close clicked
			Return
		Endif
	Endif

	* looks like Clear All was entered
	lcStem = Sys(2023) + "\_RestartAfterClearAll"
	lcFlagFile = lcStem + ".TXT"
	lcPRGFile = lcStem + ".PRG"

	If Not File (lcFlagFile )
		Erase (ForceExt(lcStem, 'TXT'))
		Erase (ForceExt(lcStem, 'PRG'))
		Strtofile ("Filler", lcFlagFile , .F.)
		Strtofile ("Erase '" + lcFlagFile + "'" + NewLine, lcPRGFile, .T.)
	Endif

	Strtofile (lcRestartText + NewLine, lcPRGFile, .T.)

	_Shell = "Do '" + lcPRGFile + "'"

Endproc
