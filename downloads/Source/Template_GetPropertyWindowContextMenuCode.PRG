	Lparameter oParm

	Local lcPEM, laObjs, laDock, lnPos, lnRow, lnCol, lnDockPos, lcWin, llHandle, llProperty, llMethod, llCustom, llInherited, llReadOnly

	lcPEM = oParm.UserTyped
	lnDockPos = 0

	Dimension laObjs[1]

	Do Case
		Case Aselobj(laObjs) # 0
			loObject = laObjs[1]
			llHandle = .T.
		Case Aselobj(laObjs,1) # 0
			loObject = laObjs[1]
			*** This is the case we want to handle
			llHandle = .T.
		Otherwise
			loObject=_Screen
	Endcase

	If llHandle

		Local lcPoint, lnx, lny, lnBar

		m.lcPoint = 0h0000000000000000

		apiGetCursorPos_pemeditor(@m.lcPoint)
		apiScreenToClient_pemeditor(_Screen.HWnd, @m.lcPoint)

		m.lnx = CToBin(Substr(m.lcPoint, 1, 4), "4rs")
		m.lny = CToBin(Substr(m.lcPoint, 5, 4), "4rs")

		m.lnCol = ScreenPixelsToCols(m.lnx)
		m.lnRow = ScreenPixelsToRows(m.lny)
		*!* calloatti

		Private plReset, plZoom, plDoForm, plFavorites, plMemberEd, plExpression, plHelp, plRemove

		plReset			= .F.
		plZoom 			= .F.
		plExpression	= .F.
		plDoForm 		= .F.
		plFavorites 	= .F.
		plMemberEd  	= .F.
		plHelp 		 	= .F.
		plRemove	 	= .F.

		lnBar = 0
		Define Popup myPopup From lnRow, lnCol SHORTCUT

		Try
			*** See if it is user defined and whether it
			*** is a property or a method to decide how to construct the menu
			Try
				llReadOnly = Pemstatus( loObject, lcPEM, 1)
			Catch
				llReadOnly = .F.
			EndTry
			llProperty = Lower( Alltrim( Pemstatus( loObject, lcPEM, 3 ) ) ) = [property]
			llMethod = Inlist( Lower( Alltrim( Pemstatus( loObject, lcPEM, 3 ) ) ), [method], [event]  )
			llCustom = Pemstatus( loObject, lcPEM, 4 )
			llInherited = Pemstatus( loObject, lcPEM, 6 )
		Catch
			llCustom = .T.
		EndTry

		If not llReadonly
			lnBar = lnBar + 1
			Define Bar (lnBar) Of myPopup Prompt <<ccLOC_CM_ResetToDefault>>
			On Selection Bar (lnBar) Of myPopup plReset = .t.

			If llProperty
				lnBar = lnBar + 1
				Define Bar (lnBar) Of myPopup Prompt <<ccLOC_CM_ZOOM>>
				On Selection Bar (lnBar) Of myPopup plZoom = .T.

				lnBar = lnBar + 1
				Define Bar (lnBar) Of myPopup Prompt <<ccLOC_CM_ExpressionBuilder>>
				On Selection Bar (lnBar) Of myPopup plExpression = .T.
			EndIf

			If llCustom and not llInherited
				lnBar = lnBar + 1
				Define Bar (lnBar) Of myPopup Prompt <<ccLOC_CM_Remove>>
				On Selection Bar (lnBar) Of myPopup plRemove = .T.
			EndIf

			lnBar = lnBar + 1
			Define Bar (lnBar) Of myPopup Prompt "\-"

		EndIf

		lnBar = lnBar + 1
		Define Bar (lnBar) Of myPopup Prompt <<ccLOC_CM_PEM_Editor>>
		On Selection Bar (lnBar) Of myPopup plDoForm = .T.

		lnBar = lnBar + 1
		Define Bar (lnBar) Of myPopup Prompt <<ccLOC_CM_MemberData>>
		On Selection Bar (lnBar) Of myPopup plMemberEd = .T.

		lnBar = lnBar + 1
		Define Bar (lnBar) Of myPopup Prompt "\-"

		lnBar = lnBar + 1
		Define Bar (lnBar) Of myPopup Prompt <<ccLOC_CM_AddToFavorites>>
		On Selection Bar (lnBar) Of myPopup plFavorites = .T.

		If not llCustom
			lnBar = lnBar + 1
			Define Bar (lnBar) Of myPopup Prompt <<ccLOC_CM_Help>>
			On Selection Bar (lnBar) Of myPopup plHelp = .T.
		EndIf

		Activate Popup myPopup
	Endif


	Do Case
		Case plReset
			ResetToDefault (loObject, lcPEM)
			Return llHandle

		Case plZoom
			Release _oZoom
			Public _oZoom
			_oZoom = Newobject('ZoomDialog', 'Source\EditProperty.vcx', ;
				'<<lcAppPath>>', loObject, lcPEM )
			_oZoom.Show()
			Return llHandle

		Case plExpression
			ExpressionBuilder( loObject, lcPEM)
			Return llHandle

		Case plDoForm
			If 'O' # Vartype (_oPEMEditor)
				Release _oPEMEditor
				Public 	_oPEMEditor
				_oPEMEditor = Newobject('PEMEditor_Main', 'PEME_Main.VCX', '<<lcAppPath>>')
			Endif

			_oPEMEditor.oUtils.ShowForm()
			_oPEMEditor.oUtils.oPEMEditor.LocatePEM(lcPEM)
			Return llHandle

		Case plFavorites
			Do (_Builder) With loObject, "MemberData", 11, lcPEM
			Return llHandle

		Case plMemberEd
			Do (_Builder) With loObject, "MemberData", 1, lcPEM
			Return llHandle

		Case plHelp
			GetHelp (lcPEM)
			Return llHandle

		Case plRemove
			Try
				lcName = Trim(lcPEM) + '_assign'
				If PemStatus (loObject, lcName, 5)
					Removeproperty(loObject, lcName)
				Endif
				lcName = Trim(lcPEM) + '_access'
				If PemStatus (loObject, lcName, 5)
					Removeproperty(loObject, lcName)
				Endif
				Removeproperty(loObject, lcPEM)
			Catch

			EndTry
			Return llHandle

		Otherwise
			Return llHandle

	Endcase

	Function apiGetCursorPos_pemeditor
		Lparameters lpPoint
		Declare Integer GetCursorPos In win32api As apiGetCursorPos_pemeditor;
			String  @lpPoint
		Return apiGetCursorPos_pemeditor(@m.lpPoint)
	Endfunc

	Function apiScreenToClient_pemeditor
		Lparameters nHwnd, lpPoint
		Declare Integer ScreenToClient In win32api As apiScreenToClient_pemeditor ;
			Integer nhWnd, ;
			String  @lpPoint
		Return apiScreenToClient_pemeditor(m.nHwnd, @m.lpPoint)
	Endfunc

	Function ScreenPixelsToRows
		Lparameters tnPixels
		Return ScreenPixelsToFoxels(m.tnPixels, .T.)
	Endfunc

	Function ScreenPixelsToCols
		Lparameters tnPixels
		Return ScreenPixelsToFoxels(m.tnPixels, .F.)
	Endfunc

	Function ScreenPixelsToFoxels
		Lparameter tnPixels, tlVertical

		Local lnFoxels, lcFontStyle
		m.lcFontStyle = ""

		If _Screen.FontBold = .T. Then
			m.lcFontStyle = m.lcFontStyle + "B"
		Endif

		If _Screen.FontItalic = .T. Then
			m.lcFontStyle = m.lcFontStyle + "I"
		Endif

		m.lnFoxels = m.tnPixels/Fontmetric(Iif(m.tlVertical, 1, 6), _Screen.FontName, _Screen.FontSize, m.lcFontStyle)

		Return m.lnFoxels
	EndFunc

	****************************************************************
	Procedure ResetToDefault

		****************************************************************
		********************************************************************
		*** Name.....: Reset2Default
		*** Author...: Marcia G. Akins
		*** Date.....: 05/13/2007
		*** Notice...: Copyright (c) 2007 Tightline Computers, Inc
		*** Compiler.: Visual FoxPro 09.00.0000.3504 for Windows
		*** Function.: Call the resetToDefault method of all the selected objects
		*** Returns..: Logical
		********************************************************************
		Lparameters toObject, tcPEM
		Local lnSelected, laSelected[ 1 ], lnI

		*** Issue a resetToDefault() for all the selected objects
		lnSelected = Aselobj( laSelected )
		If lnSelected > 0
			For lnI = 1 To lnSelected
				laSelected[ lnI ].ResetToDefault( tcPEM )
			Endfor
		Else
			toObject.ResetToDefault( tcPEM )
		Endif

	Endproc


	****************************************************************
	Procedure ExpressionBuilder( loObject, lcPEM)
		Local lcExpression, lcNewExpression, loException

		lcExpression = Substr( loObject.ReadExpression(lcPEM), 2)

		Getexpr (lcPEM) To lcNewExpression Default (lcExpression)

		Try
			loObject.WriteExpression(lcPEM, lcNewExpression)
		Catch To loException

		Endtry

	Endproc

	****************************************************************

	Procedure GetHelp(lcPEM)
		Help &lcPEM
	Endproc

