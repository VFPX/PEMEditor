Define Class BaseColumn As Column

	lClassesOnly	 = .F.

Enddefine


Define Class ObjectColumn As BaseColumn

	ReadOnly 		= .T.

	lFixedColumn    = .F.
	cPEMName		= ''
	
	nCurrentObjNumber		= 0
	
	lChanged        = .T.
	lExists			= .F.
	lHasCode		= .F.
	lNonDefault		= .F.
	lReadOnly		= .F.

	cValue			= ' '
	cOrigin			= ' '
	cType			= ' '
	
	Procedure Init
		This.Header1.FontBold = .T.
	EndProc 	

	Procedure SetPEM (lcPEM, lnWidth)
		lcPem = Alltrim (lcPEM)
		This.cPEMName = lcPEM
		This.Header1.Caption = ' ' + lcPEM
		If not Empty (lnWidth)
			This.Width = lnWidth
		EndIf 

		This.ControlSource = '(ThisForm.DynValue("' + This.Name + '"))'
		This.DynamicBackColor = 'This.DynBackColor(This.' + This.Name + ')'
		This.DynamicFontBold = 'This.DynFontBold(This.' + This.Name + ')'
		This.DynamicFontItalic = 'This.DynFontItalic(This.' + This.Name + ')'

		BindEvent (This.Header1, 'RightClick', This.Parent, 'HeaderRightClick')

		BindEvent (This.Text1, 'DblClick', This.Parent, 'HandlePropertyDblClick')
		BindEvent (This.Text1, 'RightClick', This.Parent, 'HandlePropertyRightClick')

	EndProc 	
	
Enddefine
