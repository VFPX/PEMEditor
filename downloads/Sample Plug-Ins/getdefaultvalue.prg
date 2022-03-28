*==============================================================================
* Procedure:		GetDefaultValue
* Purpose:			Returns the default value to be used for a new property
*						based on the name of the property
* Author:			Jim Nelson
* Parameters:		
*	lcName  	  = Name of the new property
* Returns:			The default value
* Added:			02/12/2009
*==============================================================================
* Notes:
*	This is a sample PRG to set the default value for a property based on the
*	name of the property.  It can be modified as needed to fit your own needs.
*	
*   This example uses the second character in the name of the property, as an
*	alternative to the PEM Editor default which uses the first letter.
*	
*	For PEM Editor to use this PRG, it must be moved into the folder "Live Plug-Ins"
*

Lparameters lcName

Local lcDefault, lcSecond

lcSecond  = Substr (lcName, 2, 1)
Do Case
	Case Not Islower (lcSecond)
		lcDefault = .F.
	Case lcSecond = 'c'
		lcDefault = []
	Case lcSecond = 'l'
		lcDefault = .F.
	Case Inlist (lcSecond, 'n', 'i', 'b', 'y')
		lcDefault = 0
	Case lcSecond = 'd'
		lcDefault = {}
	Case lcSecond = 't'
		lcDefault = {/:}
	Case lcSecond = 'o' Or lcSecond = 'u'
		lcDefault = [.Null.]
	Otherwise
		lcDefault = .F.
Endcase

Return lcDefault