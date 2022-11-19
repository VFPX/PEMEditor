*==============================================================================
* Procedure:     SetGridSortOrder
* Purpose:       Sets grid highlighting (colors, etc.)
* Author:        Jim Nelson
*                Bhavbhuti Nathwani
* Parameters:
*   tcCurrenSortOrder  = "Current" (Normal) sort order
*   lnViewType         = 1 for Normal View
*                        2 for Method View
* Returns:       Character string to be used as the sort order ... Index on
* Added:         03/16/2009
* Updated:       Bhavbhuti Nathwani - 2009/06/22
*==============================================================================
* Notes:
*   This is a sample PRG to use an alternate sort order than can be created
*   by clicking on column headings
****************************************************************

*   It can be modified as needed to fit your own needs.
*
*	For PEM Editor to use this PRG, it must be moved into the folder "Live Plug-Ins"
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

LPARAMETERS tcCurrenSortOrder, lnViewType

*-- IF lnViewType = 1
*--         RETURN [+ICase(lNonDefault, 'A', nFavorites # 1, 'C', 'D')+Icase(csrMembers.lNative,'C', csrMembers.lInherited,'B', 'A')+CName]
*-- ELSE
*--         RETURN [+ICase(lNonDefault, 'A', lHasCode, 'B', nFavorites # 1, 'C', 'D')+Icase(csrMembers.lNative,'C', csrMembers.lInherited,'B', 'A')+CName]
*-- ENDIF
*--

RETURN [+ICASE(csrMembers.lNative, 'A', csrMembers.lInherited, 'B', !csrMembers.lInherited, 'C', 'D')] ;
	+ [+ICASE(csrMembers.cType = 'P', 'B', 'A')] ;
	+ [+cName]

*[+ICASE(csrMembers.lNonDefault, 'A', 'B')] ;
