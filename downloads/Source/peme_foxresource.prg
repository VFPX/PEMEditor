* Abstract:
*   Class for add/retrieving values
*	from FoxUser resource file.
*

Define Class FoxResource As Custom
	Protected oCollection

	oCollection  = .Null.

	ResourceType = "PREFW"
	ResourceFile = ''
	lDefaultResourceFile = .T.
	cVersion	 = 2

	Procedure Init()
		This.oCollection = Createobject("Collection")
		This.ResourceFile = Sys(2005)
	Endproc

	Procedure Destroy()
		This.CloseResource()
	Endproc

	* Clear out all options
	Function Clear()
		This.oCollection.Remove(-1)
	Endfunc

	Function Set(cOption, xValue)
		Try
			This.oCollection.Remove(Upper(m.cOption))
		Catch
		Endtry
		* doesn't exist yet, so add
		Return This.oCollection.Add(m.xValue, Upper(m.cOption))
	Endfunc

	Function Get(cOption)
		Local xValue

		Try
			m.xValue = This.oCollection.Item(Upper(m.cOption))
		Catch
			m.xValue = .Null.
		Endtry

		Return m.xValue
	Endfunc

	Function OpenResource()
		If This.lDefaultResourceFile And !(Set("RESOURCE") == "ON")
			Return .F.
		Endif

		If !Used("FoxResource")
			If File(This.ResourceFile)
				Try
					Use (This.ResourceFile) Alias FoxResource In 0 Shared Again
				Catch
				Endtry
			Endif
		Endif
		Return Used("FoxResource")
	Endfunc

	Function CloseResource()
		If Used("FoxResource")
			Use In FoxResource
		Endif
	Endfunc

	Procedure Save(cID, cName)
		Local nSelect
		Local cType
		Local i
		Local Array aOptions[1]

		If Vartype(m.cName) <> 'C'
			m.cName = ''
		Endif
		If This.OpenResource()
			m.nSelect = Select()

			m.cType = Padr(This.ResourceType, Len(FoxResource.Type))
			m.cID   = Padr(m.cID, Len(FoxResource.Id))

			Select FoxResource
			Locate For Type == m.cType And Id == m.cID And Name == m.cName
			If !Found()
				Append Blank In FoxResource
				Replace ;
					Type With m.cType, ;
					Name With m.cName, ;
					ID With m.cID, ;
					ReadOnly With .F. ;
					IN FoxResource
			Endif

			If !FoxResource.ReadOnly
				If This.oCollection.Count > 0
					Dimension aOptions[THIS.oCollection.Count, 2]
					For m.i = 1 To This.oCollection.Count
						aOptions[m.i, 1] = This.oCollection.GetKey(m.i)
						aOptions[m.i, 2] = This.oCollection.Item(m.i)
					Endfor
					Save To Memo Data All Like aOptions
				Else
					Blank Fields Data In FoxResource
				Endif

				Replace ;
					Updated With Date(), ;
					ckval With Val(Sys(2007, FoxResource.Data)) ;
					IN FoxResource
			Endif

			This.CloseResource()

			Select (m.nSelect)
		Endif
	Endproc

	Procedure Load(cID, cName)
		Local nSelect
		Local cType
		Local i
		Local nCnt
		Local Array aOptions[1]

		If Vartype(m.cName) <> 'C'
			m.cName = ''
		Endif

		* THIS.Clear()
		If This.OpenResource()
			m.nSelect = Select()

			m.cType = Padr(This.ResourceType, Len(FoxResource.Type))
			m.cID   = Padr(m.cID, Len(FoxResource.Id))

			Select FoxResource
			Locate For Type == m.cType And Id == m.cID And Name == m.cName
			If Found() And !Empty(Data) And ckval == Val(Sys(2007, Data))
				Restore From Memo Data Additive
				If Vartype(aOptions[1,1]) == 'C'
					m.nCnt = Alen(aOptions, 1)
					For m.i = 1 To m.nCnt
						This.Set(aOptions[m.i, 1], aOptions[m.i, 2])
					Endfor
				Endif
			Endif

			This.CloseResource()

			Select (m.nSelect)
		Endif
	Endproc

	Function GetData(cID, cName)
		Local cData
		Local nSelect
		Local cType

		If Vartype(m.cName) <> 'C'
			m.cName = ''
		Endif

		m.cData = .Null.
		If This.OpenResource()
			m.nSelect = Select()

			m.cType = Padr(This.ResourceType, Len(FoxResource.Type))
			m.cID   = Padr(m.cID, Len(FoxResource.Id))

			Select FoxResource
			Locate For Type == m.cType And Id == m.cID And Name == m.cName
			If Found() And !Empty(Data) && AND ckval == VAL(SYS(2007, Data))
				m.cData = FoxResource.Data
			Endif

			This.CloseResource()

			Select (m.nSelect)
		Endif

		Return m.cData
	Endfunc

	Function SetData(cID, cName, cSetData, nCkVal)
		Local nSelect
		Local cType
		Local lSuccess

		If Vartype(m.cName) <> 'C'
			m.cName = ''
		Endif

		m.lSuccess = .F.
		If This.OpenResource()
			m.nSelect = Select()

			m.cType = Padr(This.ResourceType, Len(FoxResource.Type))
			m.cID   = Padr(m.cID, Len(FoxResource.Id))

			Select FoxResource
			Locate For Type == m.cType And Id == m.cID And Name == m.cName
			If !Found()
				Append Blank In FoxResource
				Replace ;
					Type With m.cType, ;
					Name With m.cName, ;
					ID With m.cID, ;
					ReadOnly With .F. ;
					IN FoxResource
			Endif

			If !FoxResource.ReadOnly
				Replace Data With m.cSetData In FoxResource

				If Vartype(nCkVal) <> 'N'
					nCkVal = Val(Sys(2007, FoxResource.Data))
				Endif

				Replace ;
					Updated With Date(), ;
					ckval With nCkVal ;
					IN FoxResource

				m.lSuccess = .T.
			Endif


			This.CloseResource()

			Select (m.nSelect)
		Endif

		Return m.lSuccess
	Endfunc

	* Add an item to an MRU list
	Function AddToMRU(cMRUName, cMRUItem)
		Local cMRUData
		Local nSelect
		Local i
		Local nCnt
		Local Array aMRUList[1]

		If This.OpenResource()
			m.nSelect = Select()

			m.cType    = Padr(This.ResourceType, Len(FoxResource.Type))
			m.cMRUName = Padr(m.cMRUName, Len(FoxResource.Id))

			Select FoxResource
			Locate For Type == m.cType And Id == m.cMRUName And Empty(Name)
			If !Found()
				Append Blank In FoxResource
				Replace ;
					Type With m.cType, ;
					Name With '', ;
					ID With m.cMRUName, ;
					ReadOnly With .F., ;
					Data With Chr(4) + Chr(0) + Chr(0) ;
					IN FoxResource
			Endif
			If !FoxResource.ReadOnly
				m.nCnt = Alines(aMRUList, Substr(FoxResource.Data, 3), .F., Chr(0))

				m.cMRUList = m.cMRUItem + Chr(0)
				For m.i = 1 To m.nCnt
					If !(Upper(m.cMRUItem) == Upper(aMRUList[m.i]))
						m.cMRUList = m.cMRUList + aMRUList[m.i] + Chr(0)
					Endif
				Endfor
				m.cMRUList = m.cMRUList + Chr(0)

				Replace ;
					Data With Chr(4) + Chr(0) + m.cMRUList, ;
					ckval With Val(Sys(2007, m.cMRUList)), ;
					Updated With Date() ;
					IN FoxResource

				m.lSuccess = .T.
			Endif

			This.CloseResource()

			Select (m.nSelect)
		Endif
	Endfunc

	* save to a specific fieldname
	Function SaveTo(cField, cAlias)
		Local i
		Local nSelect
		Local lSuccess
		Local Array aOptions[1]

		If Vartype(m.cAlias) <> 'C'
			m.cAlias = Alias()
		Endif

		If Used(m.cAlias)
			m.nSelect = Select()
			Select (m.cAlias)

			If This.oCollection.Count > 0
				Dimension aOptions[THIS.oCollection.Count, 2]
				For m.i = 1 To This.oCollection.Count
					aOptions[m.i, 1] = This.oCollection.GetKey(m.i)
					aOptions[m.i, 2] = This.oCollection.Item(m.i)
				Endfor
				Save To Memo &cField All Like aOptions
			Else
				Blank Fields &cField In FoxResource
			Endif
			Select (m.nSelect)
			m.lSuccess = .T.
		Else
			m.lSuccess = .F.
		Endif

		Return m.lSuccess
	Endfunc


	Function RestoreFrom(cField, cAlias)
		Local i
		Local nSelect
		Local lSuccess
		Local Array aOptions[1]

		If Vartype(m.cAlias) <> 'C'
			m.cAlias = Alias()
		Endif

		If Used(m.cAlias)
			m.nSelect = Select()
			Select (m.cAlias)

			Restore From Memo &cField Additive
			If Vartype(aOptions[1,1]) == 'C'
				m.nCnt = Alen(aOptions, 1)
				For m.i = 1 To m.nCnt
					This.Set(aOptions[m.i, 1], aOptions[m.i, 2])
				Endfor
			Endif

			Select (m.nSelect)
			m.lSuccess = .T.
		Else
			m.lSuccess = .F.
		Endif

		Return m.lSuccess
	Endfunc

	*** JRN 02/10/2009 : Added to allow for alternative resource files
	Function UseResourceFile (tcFileName)
		This.CloseResource()
		This.ResourceFile = Fullpath(tcFileName)
		This.lDefaultResourceFile = .F.
	Endfunc

	Function CreateResourceFile (tcFileName)
		This.CloseResource()
		Create Table (tcFileName) Free (	;
			TYPE            C(12)	;
			, Id            C(12)	;
			, Name          M		;
			, ReadOnly      L		;
			, ckval         N(6)	;
			, Data          M		;
			, Updated       D		;
			)
		Use

		This.ResourceFile = Fullpath(tcFileName)
		lDefaultResourceFile = .F.
	Endfunc

Enddefine

