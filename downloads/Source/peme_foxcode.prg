Define Class clsFoxCode As Custom

	cError = ''

	Procedure Init
		This.Setup
	Endproc


	Procedure Setup
		Local lcFoxCode, llFailure
		lcFoxCode = (_Foxcode)
		llFailure = .F.
		Try
			Use (lcFoxCode) Again Shared Alias PEME_FoxCode
		Catch
			llFailure = .T.
		Endtry
		If llFailure
			This.cError = 'Unable to open FoxCode'
			Return .F.
		Endif
	Endproc


	Procedure Destroy
		Use In (Select ('PEME_FoxCode')) && Close 'PEME_FoxCode'
	Endproc


	Procedure GetStatus(lcType, lcAbbrev, lcData)
		Local lnResult, lnSelect
		lnSelect = Select()
		Select PEME_FoxCode
		lcAbbrev = Padr(lcAbbrev, Len(Abbrev))
		Locate For Type = lcType And Upper(Abbrev) == Upper(lcAbbrev) And Not Deleted()
		Do Case
			Case Not Found()
				lnResult = 1
			Case Upper(Data) == Upper(lcData)
				lnResult = 2
			Case 'PEMEDITOR.APP' $ Upper(Data)
				lnResult = 3
			Otherwise
				lnResult = 4
		Endcase
		Select (lnSelect)
		Return lnResult
	Endproc

	Procedure AddItem(lcType, lcAbbrev, lcData)
		Local lnSelect, lnStatus
		lnStatus = This.GetStatus(lcType, lcAbbrev, lcData)
		Do Case
			Case lnStatus = 1
				Insert Into PEME_FoxCode(Type, Abbrev, Data, Timestamp) Values (lcType, lcAbbrev, lcData, Datetime())
			Case lnStatus = 2

			Otherwise
				Replace Data With lcData, Timestamp With Datetime() In PEME_FoxCode
		Endcase
	Endproc

	Procedure RemoveItem(lcType, lcAbbrev, lcData)
		Local lnSelect, lnStatus
		lnStatus = This.GetStatus(lcType, lcAbbrev, lcData)
		If lnStatus # 1
			Delete In PEME_FoxCode
		Endif
	Endproc

Enddefine
