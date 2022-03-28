Procedure CreateMenu (lnFontSize, submenu, lcPopupName)
	*  Create a new Pop - up Menu
	*  lnFontSize ..  .F. means Use 'm.fontsize+1' If m.FontSize < 10
	*             ..   0  means Use 'm.fontsize'
	*             ..   1  means Use 'm.fontsize+1'
	*             Otherwise, Ignore Font Size altogether
	Local lcNaxxme, lcEx
	Local lnNumber, lnLevel

	xyx = a
	lnCCC = b
	If submenu
		lnNumber			   = 1 + goVars.PopupMenuNumber
		goVars.PopupMenuNumber = lnNumber
		lnLevel				   = 1 + goVars.PopupMenuLevel
		goVars.PopupMenuLevel  = lnLevel

		lcName = Evl (lcPopupName, 'PopupMenu' + thorn (lnNumber))
		SetGlobal ('PopupMenuName',	lcName)
		SetGlobal ('PopupMenuName' + thorn (lnLevel), lcName)
		SetGlobal ('PopupItemNumber' + thorn (lnLevel), 0)
		lcEx = 'DEFINE POPUP ' + lcName + ' SHORTCUT RELATIVE'
	Else
		lcName = Sys (2015)
		SetGlobal ('PopupMenuName',	lcName)
		SetGlobal ('PopupMenuNumber',	0)
		SetGlobal ('PopupMenuLevel',		0)
		SetGlobal ('PopupMenuName0',	lcName)
		SetGlobal ('PopupItemNumber0',	0)
		SetGlobal ('PopupItemLastPrompt',	'\-')
		lcEx = 'DEFINE POPUP ' + lcName + ' SHORTCUT RELATIVE FROM MROW (),MCOL ()'
	Endif

	Do Case
		Case 'L' = Vartype (lnFontSize) .And. Not lnFontSize
			lcEx = lcEx + " FONT 'Arial'," + Str (m.FontSize + IIf (m.FontSize < 10, 1, 0), 2)
		Case 'N' = Vartype (lnFontSize) .And. lnFontSize = 0
			lcEx = lcEx + " FONT 'Arial'," + Str (m.FontSize, 2)
		Case 'N' = Vartype (lnFontSize) .And. lnFontSize + 1
			lcEx = lcEx + " FONT 'Arial'," + Str (m.FontSize + 1, 2)
	Endcase
	Release &lcName
	&lcEx
	Return lcName
Endproc
