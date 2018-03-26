*==============================================================================
* Procedure:  SetGridHighlighting
* Purpose:    Sets grid highlighting (colors, etc.)
* Author:     Jim Nelson
*             Bhavbhuti Nathwani
* Parameters:
*        loGrid      = the grid
*        lcObject    = character string which can be used in Dynamic* properties
*                            to refer to the currently select objecgt
*        lnViewType  = 1 for Normal View
*                      2 for Method View
* Returns:    (ignored)
* Added:      02/12/2009
* Updated:    Bhavbhuti Nathwani - 2009/06/10
* Updated:    BN - 2009/06/13
*==============================================================================
* Notes:
*   This is a sample PRG to set the set the Dynamic* properties for the main grid.
*   It actually can be used to modify ANY properties of the grid.
****************************************************************

*   It can be modified as needed to fit your own needs.
*
*   This example marks the names for read-only properties in italics.
*
*	For PEM Editor to use this PRG, it must be moved into the folder "Live Plug-Ins"
*
* When this PRG is called, the following properties have already been set
* for all columns in the grid:
*                DynamicBackColor
*                DynamicForeColor
*                DynamicFontBold
*                DynamicFontItalic
*
* Grid Columns:
*   .Columns(1) ... ColName
*   .Columns(2) ... ColType
*   .Columns(3) ... ColAccess
*   .Columns(4) ... ColAssign
*   .Columns(5) ... ColVisibility
*   .Columns(6) ... Hierarchy
*   .Columns(7) ... ColFavorite
*   .Columns(8) ... NonDefault
*   .Columns(9) ... ColTypeIcon
*   .Columns(10) .. ColValue
*   .Columns(11) .. ColCode
*
* Cursor csrMembers:
*        cName           C(60)
*        cType           C(1)
*        lAccess         L
*        lAssign         L
*        nVisibility     N(2)
*        lNative         L
*        lInherited      L
*        lNonDefault     L
*        lHasCode        L
*        nFavorites      N(1)
*        lReadOnly       L
*        cDescript       M

Lparameters loGrid, lcObject, lnViewType

Local lcBgColor, lcDynamic, lcFgColor, lnExpr, lnI

If lnViewType = 1
	With loGrid
		lcDynamic = 'lReadOnly'
		.ColName.DynamicFontItalic = lcDynamic
	Endwith
Endif


*** BN applying a custom DynamicBackColor to all the columns
*** BN natives are in grays
*** BN inherited ones in pinks
*** BN added ones in yellows
*** first two IIF statements take care of when the grid is disabled and current record highlighing
lcBgColor = Transform (GetSysColor ( 13 ))
lcFgColor = Transform (GetSysColor ( 14 ))

lnExpr = 'IIF(!THIS.ENABLED, RGB(212,208,200), '								;
	+ 'IIF(Recno() = This.nCurrentRecno and This.HighlightStyle # 2, ' + lcBgColor + ', ';
	+ "IIF(csrMembers.lNative, IIF(csrMembers.cType = 'P', RGB(215,215,215), RGB(225,225,225)), ";
	+ "IIF(csrMembers.lInherited, IIF(csrMembers.cType = 'P', RGB(255,180,200), RGB(255,190,210)), ";
	+ "IIF(csrMembers.cType = 'P', RGB(255,180,100), RGB(255,200,120))))))"

For lnI = 1 To 11
	loGrid.Columns[lnI].DynamicBackColor = lnExpr
Endfor


*** BN applying a custom DynamicForeColor to all the columns
*** BN non-defaults in red (bold too due to the default DynamicFontBold setting)
*** BN defaults in black
*** first two IIF statements take care of when the grid is disabled and current record highlighing
*** if disabled, all gray
lnExpr = 'IIF(!This.Enabled, Rgb(128, 128, 128), '								;
	+ 'IIF(Recno() = This.nCurrentRecno and This.HighlightStyle # 2, ' + lcFgColor + ', ';
	+ 'IIF(csrMembers.lNonDefault, RGB(255,0,0), RGB(0,0,0))))'

For lnI = 1 To 11
	loGrid.Columns[lnI].DynamicForeColor = lnExpr
Endfor
