*** 2010-11-09 : Tore Bleken

Lparameters  lcFileName

Local lcFileToFind, lcResult, lcReturn, lnFileCount, lnX, loDBF

lcResult	 = ''
lcFileToFind = Alltrim (Lower (Juststem (lcFileName)))
lcReturn	 = ''
Try
	Use (lcFileToFind) Alias (Juststem (lcFileToFind)) In 0
	lcReturn = Juststem (lcFileToFind)
Catch
	loDBF		= _oPEMEditor.oTools.GetMRUList ('dbf')
	lnFileCount	= loDBF.Count
	If lnFileCount > 0
		For lnX = 1 To lnFileCount
			Do Case
				Case Empty (Justpath (lcFileToFind)) And lcFileToFind == Lower (Juststem (loDBF.Item (lnX))) And File (loDBF.Item (lnX))
					lcResult = loDBF.Item (lnX)
					Exit
				Case Not Empty (Justpath (lcFileToFind)) And lcFileToFind $ Lower (loDBF.Item (lnX)) And File (loDBF.Item (lnX))
					lcResult = loDBF.Item (lnX)
					Exit
				Otherwise
			Endcase
		Endfor
	Endif
	Release loDBF
	If Not Empty (lcResult)
		Try
			Use (lcResult) In 0 Alias (Juststem (lcResult))
			lcReturn = Juststem (lcFileName)
			lcResult = lcReturn
			_oPEMEditor.oTools.AddMRUFile (Dbf (lcResult))
		Endtry
	Endif
	If Empty (lcResult)
		Try
			Use (lcFileName) In 0 Alias (Juststem (lcFileName))
			lcReturn = Juststem (lcFileName)
			_oPEMEditor.oTools.AddMRUFile (Dbf (lcFileName))
		Catch
			Messagebox (lcFileName + ' not open!!', 0, 'PEME_OpenTable.PRG')
		Endtry
	Endif
Endtry
Return lcReturn


