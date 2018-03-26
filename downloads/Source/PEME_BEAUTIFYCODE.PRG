#include 'beautify.h'

Define Class CBeautify As Session
	DataSession = 2		&& private
	Visible = .F.
	Name = 'Beautify'
	outfile = ''
	LibName		= ''

	Procedure Init
		Local LibName, mdbf
		If Not Used("fdkeywrd")
			mdbf = This.whereis ('FDKEYWRD.DBF')
			If Not File (m.mdbf)
				Messagebox (E_STRING1_LOC, 0)
				Return .F.
			Endif
			Use (m.mdbf) Order token Alias fdkeywrd In 0
		Endif

		LibName = This.whereis ('fd3.fll')
		If Not File (m.LibName)
			Messagebox (E_STRING2_LOC, 0)
			Return .F.
		Endif
		This.LibName = m.LibName

	Endproc

	Procedure Destroy
		Try
			Release Library (This.LibName)
		Catch

		Endtry
	Endproc

	Procedure Execute (tcCode, tcOptions, tnRetryAttempt, tcExplText)
		*		tcCode should be passed by ref

		Local lcCodeBlock, lnExecuteCounter, lnRetryAttempt
		Local llSuccess, LibName, xrefname, mdbf
		Local errLogFile, mOldLogErrors
		Local nWindowHandle
		Local nStartPos
		Local nEndPos
		Local nStartLine
		Local nEndLine
		Local nRetCode
		Local nPos
		Local cCodeBlock
		Local cTempInFile
		Local cIndentText
		Local i, nCnt
		Local lSelection
		Local nNewLen
		Local cFoxToolsLibrary
		Local Array aEdEnv[25]
		Local Array aCodeLines[1]

		Set Talk Off
		Set Safety Off		&& scoped to datasession

		* 		These variables are needed by the FD3.FLL library
		symbol = ''
		winname = 0
		winpos = 0
		File = ''
		filetype = ''
		done = 0
		Flags = ''
		sniplineno = 0
		classname = ''
		BaseClass = ''
		mtemp = ''
		temp = ''
		fpoutfile = -1
		mout = ''
		totallines = 0
		lnRetryAttempt = IIf (Empty (tnRetryAttempt), 3, tnRetryAttempt)

		Set Library To (This.LibName) Additive
		Set Message To ''
		Select fdkeywrd

		*	 Generate a temp file in the temp file directory
		outfile = Addbs (Sys(2023)) + Substr (Sys(2015), 3, 10) + '.TMP'

		*	 If usercase mode, create xref table for user symbols
		If (Substr (tcOptions, 1, 1) = Chr(3))
			xrefname = 'FDXREF'
			Create Cursor (m.xrefname) (									    ;
				symbol C(65),												    ;
				ProcName C(40),												    ;
				FLAG C(1),													    ;
				LINENO N(5),												    ;
				SnipRecNo N(5),												    ;
				SnipFld C(10),												    ;
				sniplineno N(5),											    ;
				ADJUST N(5),												    ;
				Filename C(161)												    ;
				)
			Index On Flag Tag Flag 	&& for rushmore
			Index On Upper (symbol) + Flag Tag symbol
		Endif

		cTempInFile = Addbs (Sys(2023)) + Substr (Sys(2015), 3, 10) + '.TMP'

		If Strtofile (tcCode, m.cTempInFile) > 0
			outfile = Addbs (Sys(2023)) + Substr (Sys(2015), 3, 10) + '.TMP'
			lnExecuteCounter = lnRetryAttempt
			llSuccess = .F.
			Do While lnExecuteCounter > 0 And Not llSuccess
				*Wait Window 'Calling beautify. Try ' + Transform (lnRetryAttempt - lnExecuteCounter + 1) +;
				' of ' + Transform (lnRetryAttempt)  + ' ... ' Timeout 0.1
				lnExecuteCounter = lnExecuteCounter - 1
				DoEvents Force

				llSuccess = Beautify ((cTempInFile), (outfile), (tcOptions))

				*** JRN 05/20/2010 : Courtesy Gregory L Reichert
				If Substr(m.tcoptions,5,1)=Chr(3) and Substr(m.tcoptions,1,1) $ Chr(3) + Chr[4]
					This.ForceBeautifyAllLines (m.outfile, m.tcOptions)
				EndIf 

				*Wait Window 'Calling beautify. Try ' + Transform (lnRetryAttempt - lnExecuteCounter + 1) +;
				' of ' + Transform (lnRetryAttempt)  + ' ... ' Timeout 0.1

				If m.llSuccess
					* 	read back from the file that the Beautify function created
					lcCodeBlock = Filetostr (m.outfile)
				Else  && m.llSuccess
					If lnExecuteCounter = 0
						Strtofile (Transform (Datetime()) + ' error by executing beautify ' +;
							tcExplText + Chr(13) + Chr(10),					    ;
							'BeautifyError.txt', 1)
					Endif
				Endif && m.llSuccess
			Enddo
		Endif	&& STRTOFILE(tcCode, m.cTempInFile

		If (Substr (tcOptions, 1, 1) = Chr(3))
			Select fdxref
			Use
			*			Delete FILE (m.xrefname)
		Endif

		tcCode = lcCodeBlock
		Return .T.
	Endproc && Execute


	*------------------------------------------------------------
	* Description: corrent case when set to mixed
	* Parameters:
	* Return:
	* Use:
	*------------------------------------------------------------
	* Id Date		By		Description
	*  1 05/12/2010	GLR		Initial Creation
	*
	*------------------------------------------------------------
	Procedure CorrectMixedCase(lcFile, options )

		*-- if mixed case on keywords
		If Substr(m.options,5,1)=Chr(3)
			Local laCode[1], lcCode2 As String, ii
			Local tmpFile1 As String
			Local tmpFile2 As String
			Local llDoLastToo As Boolean, llSkip As Boolean
			Local lcTestLine As String

			*-- temporary files
			tmpFile1 = This.tempfile()
			tmpFile2 = This.tempfile()

			*-- force indent change off
			m.options = Stuff(m.options,13,1,Chr(3))

			llSkip = .F.
			llDoLastToo = .F.
			lcCode2 = []
			lcTestLine = []
			For ii= 1 To Alines(laCode, Filetostr(m.lcFile), 2)
				lcTestLine = Alltrim(Chrtran(laCode[ii],Tab,[ ]))
				Do Case
					Case Empty(lcTestLine )
						*-- handle blank lines
						lcCode2 = lcCode2 + crlf

					Case Inlist(Getwordnum(Upper(lcTestLine ),1,[ ]), ;
							[TEXT] )
						*-- keyword block, skip
						*-- beautify each continuation line at a time
						Strtofile(laCode[ii],m.tmpFile1,0)
						If Not Beautify(m.tmpFile1, m.tmpFile2, m.options )
							*-- failed, rollback
							Return .F.
						Endif
						*-- collect line
						lcCode2 = lcCode2 + Filetostr(m.tmpFile2) &&+ crlf
						llSkip = .T.
					Case Inlist(Getwordnum(Upper(lcTestLine),1,[ ]), ;
							[ENDTEXT] )
						*-- end of keyword block
						*-- beautify each continuation line at a time
						Strtofile(laCode[ii],m.tmpFile1,0)
						If Not Beautify(m.tmpFile1, m.tmpFile2, m.options )
							*-- failed, rollback
							Return .F.
						Endif
						*-- collect line
						lcCode2 = lcCode2 + Filetostr(m.tmpFile2) &&+ crlf
						llSkip = .F.

					Case Atc([*# beautify keyword_nochange],lcTestLine )>0
						*-- skip nochange keywords
						lcCode2 = lcCode2 + laCode[ii] + crlf
						llSkip = .T.
					Case [*# BEAUTIFY]==Upper(lcTestLine )
						*-- resume
						lcCode2 = lcCode2 + laCode[ii] + crlf
						llSkip = .F.

					Case llSkip
						*-- skip beautifying
						lcCode2 = lcCode2 + laCode[ii] + crlf

					Case Left(lcTestLine,1)==[*]
						*-- skip comment lines
						lcCode2 = lcCode2 + laCode[ii] + crlf

						*--CASE (AT([;], GETWORDNUM(laCode[ii],1,[&]+[&])>0)
						*-- remove trailing comments. look for ending ';'
					Case Right(Rtrim(Getwordnum(lcTestLine ,1,[&]+[&])),1)==[;]
						*-- beautify each continuation line at a time
						Strtofile(laCode[ii],m.tmpFile1,0)
						If Not Beautify(m.tmpFile1, m.tmpFile2, m.options )
							*-- failed, rollback
							Return .F.
						Endif
						*-- collect line
						lcCode2 = lcCode2 + Filetostr(m.tmpFile2) &&+ crlf
						llDoLastToo = .T.

					Case llDoLastToo
						*-- include the last line of the continuation
						Strtofile(laCode[ii],m.tmpFile1,0)
						If Not Beautify(m.tmpFile1, m.tmpFile2, m.options )
							Return .F.
						Endif
						lcCode2 = lcCode2 + Filetostr(m.tmpFile2) &&+ crlf
						llDoLastToo = .F.

					Otherwise
						*-- all other lines
						lcCode2 = lcCode2 + laCode[ii] + crlf
				Endcase
			Next
			Strtofile(m.lcCode2, m.lcFile, 0)
			Erase (m.tmpFile1)
			Erase (m.tmpFile2)
		Endif
	Endproc


	*------------------------------------------------------------
	* Description:
	* Parameters:  <para>, <req/opt>, D=<def>, <desc>
	* Return:
	* Use:
	*------------------------------------------------------------
	* Id Date		By		Description
	*  1 05/06/2010	GLR		Initial Creation
	*
	*------------------------------------------------------------
	Procedure tempfile()
		Return Addbs(Sys(2023)) + Substr(Sys(2015), 3, 10) + ".TMP"
	Endproc


	*------------------------------------------------------------
	* Description: correct case when set to mixed, new version
	*------------------------------------------------------------
	* Id Date		By		Description
	*  1 05/21/2010	JRN		Initial Creation
	*
	*------------------------------------------------------------
	* Borrowing on GLR's concept above, but different take:
	* double space all the code, so that no code is on a continuation line
	* then apply beautify, then remove the double spacing

	Procedure CorrectMixedCaseNew (lcFile, options )

		*-- if mixed case on keywords
		If Substr(m.options,5,1)=Chr(3)
			Local laCode[1], lcCode2 As String
			Local tmpFile1 As String
			Local tmpFile2 As String

			*-- temporary files
			tmpFile1 = This.tempfile()
			tmpFile2 = This.tempfile()

			*-- force indent change off
			m.options = Stuff(m.options,13,1,Chr(3))

			* double space all lines
			lcCode = Strtran (Filetostr(m.lcFile), crlf, crlf + crlf)

			* apply beautify
			Strtofile(lcCode, m.tmpFile1,0)
			If Not Beautify(m.tmpFile1, m.tmpFile2, m.options )
				Return .F.
			Endif

			* and remove the doubling
			lcCode2 = Strtran (Filetostr(m.tmpFile2), crlf + crlf, crlf)

			*
			Strtofile(m.lcCode2, m.lcFile, 0)
			Erase (m.tmpFile1)
			Erase (m.tmpFile2)
		Endif
	Endproc


	*------------------------------------------------------------
	* Id Date		By		Description
	*  1 06/07/2010	JRN		Initial Creation
	*
	*------------------------------------------------------------
	* Borrowing on GLR's concept above, of passing thru Beautify again ...
	* double space all the code, so that no code is on a continuation line
	* also adding a '!' at the beginning of each line
	* then apply beautify, then remove the double spacing and '!'
	* this applies beautify correctly to ALL lines, including within TEXT/ENDTEXT
	* (so the calling program must know when to reverse this)

	Procedure ForceBeautifyAllLines (lcFile, options )

		*-- if mixed case on keywords
		Local laCode[1], lcCode2, replacement As String
		Local tmpFile1 As String
		Local tmpFile2 As String

		*-- temporary files
		tmpFile1 = This.tempfile()
		tmpFile2 = This.tempfile()

		*-- force indent change off
		m.options = Stuff(m.options,13,1,Chr(3))
		m.replacement = crlf + crlf + '!'

		* double space all lines
		lcCode = Strtran (Filetostr(m.lcFile), crlf, m.replacement)

		* apply beautify
		Strtofile(lcCode, m.tmpFile1,0)
		If Not Beautify(m.tmpFile1, m.tmpFile2, m.options )
			Return .F.
		Endif

		* and remove the doubling
		lcCode2 = Strtran (Filetostr(m.tmpFile2), m.replacement, crlf)

		*
		Strtofile(m.lcCode2, m.lcFile, 0)
		Erase (m.tmpFile1)
		Erase (m.tmpFile2)
	Endproc


	Protected Procedure whereis (mfile)
		Local mtemp
		If File (m.mfile)
			Return m.mfile
		Endif
		mtemp = 'fd3fll\' + m.mfile
		If File (m.mtemp)
			Return m.mtemp
		Endif
		mtemp = Sys(2004) + m.mfile
		If File (m.mtemp)
			Return m.mtemp
		Endif
		mtemp = Sys(2004) + 'wizards\' + m.mfile
		If File (m.mtemp)
			Return m.mtemp
		Endif
		Return ''
	Endproc


Enddefine
