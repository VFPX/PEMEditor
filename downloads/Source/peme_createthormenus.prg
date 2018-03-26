#Define CR Chr(13)
#Define LF Chr(10)

Local lcPopupName, lcPrompt, llTopLevel, lnMyMenuID, lnThorMenuID, loThor

loThor = Execscript (_Screen.cthordispatcher, 'Thor Engine=')

Select  *						;
	  , 0000    As  ThorMenuID	;
	From ThorMenus				;
	Into Cursor crsr_ThorMenus Readwrite

Scan
	*  Create or find this menu
	lcPopupName	= Trim (PopupName)
	lcPrompt	= Trim (Prompt)
	llTopLevel	= TopLevel
	lnMyMenuID	= MenuID

	lnThorMenuID = loThor.GetMenuByName (lcPopupName)
	If lnThorMenuID = 0
		lnThorMenuID = loThor.AddMenu (lcPrompt, llTopLevel, .F., , , , , lcPopupName)
	Else
		Try
			loThor.SetMenuPrompt (lnThorMenuID, lcPrompt)
		Catch
		Endtry
	Endif

	Replace									;
			ThorMenuID	With  lnThorMenuID	;
		In crsr_ThorMenus
Endscan

Select crsr_ThorMenus

Scan
	CreateMenuBars (loThor, ThorMenuID, MenuID)
	Select crsr_ThorMenus
Endscan

loThor.PackTables()

Return


Procedure CreateMenuBars (loThor, lnThorMenuID, lnMyMenuID)

	Local laThorSubMenuID[1], laTools[1], lcPRG, lcPrompt, lcStatusBar, lnCount, lnI, lnID, lnSubMenuID

	* Eliminate previous entries
	lnCount = loThor.GetMenuItems (lnThorMenuID, @laTools)

	For lnI = 1 To lnCount
		lnID = laTools (lnI, 1)
		loThor.RemoveItemFromMenu (lnID)
	Endfor

	* Now re-add
	Select  *							;
		From ThorSource					;
		Where MenuID = lnMyMenuID		;
		Order By sortorder				;
		Into Cursor crsr_SubMenu
	Scan
		lcPrompt	= ToolPrmpt
		lcStatusBar	= Strtran (Trim (Evl (StatusBar, Summary)), CR + LF, ' ')
		Do Case
			Case submenu
				lnSubMenuID = SubMenuID
				Select ThorMenuID From crsr_ThorMenus Where MenuID = lnSubMenuID Into Array laThorSubMenuID
				loThor.AddSubMenu (lnThorMenuID, laThorSubMenuID(1), lcPrompt, , lcStatusBar)

			Case Separator
				loThor.AddItemToMenu (lnThorMenuID, , , .T.)
				
			Case Empty(Category)

			Otherwise
				lcPRG		= Forceext ('Thor_Tool_PEME_' + Trim (prgName), 'PRG')

				*!* * Removed 8/11/2011 ... IDE Tools must keep their own prompts!????
				*!* For lnI = 1 To lnCount
				*!* 	If Upper (laTools (lnI, 2)) = Upper(lcPRG)
				*!* 		lcPrompt = laTools (lnI, 3)
				*!* 	Endif
				*!* Endfor

				loThor.AddItemToMenu (lnThorMenuID, lcPRG, lcPrompt, .F., , lcStatusBar)

		Endcase
	Endscan

Endproc
