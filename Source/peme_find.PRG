Parameters goPEMEditor, goObject, lcPhrase, FullPathName
Local llMatch, loException, lxValue
With goObject
	llMatch = .F.
	Try
		lxValue = Evaluate (lcPhrase)
		llMatch = (Not Empty (lxValue)) and (Not IsNull (lxValue))
	Catch To loException

	Endtry
Endwith
Return llMatch


****************************************************************
Function EQ (lcLeft, lcRight)
	Return Lower (lcLeft) = Lower (lcRight)
EndFunc


Function Contains (lcLeft, lcRight)
	Return Lower (lcLeft) $ Lower (lcRight)
EndFunc


Function Exists (lcName)
	Return Pemstatus (goObject, lcName, 5)
EndFunc


Function HasCode (lcName)
	Return Pemstatus (goObject, lcName, 5) And Pemstatus (goObject, lcName, 0)
EndFunc


Function NonDefault (lcName)
	If Pemstatus (goObject, lcName, 5) And Pemstatus (goObject, lcName, 0)
		If Inlist (Pemstatus (goObject, lcName, 3), 'Method', 'Event')
			Return Not Empty (goObject.ReadMethod (lcName)) 
		Else
			Return .T.
		Endif
	Else
		Return .F.
	Endif
EndFunc


Function FullObjectName
	Return goPEMEditor.getobjectpath (goObject)
EndFunc
