*==============================================================================
* Procedure:		GetNewMethodHeader
* Purpose:			Returns a method header which is used for any new method.
* Author:			Jim Nelson
* Parameters:		
*	lcMethodName  = Name of the new method
*	lcDescription = Its description
*   lnVisibility  = 1 for Public, 2 for Protected, 3 for Hidden
* Returns:			The header text to be inserted into the method
* Added:			02/12/2009
*==============================================================================
* Notes:
*	This is a sample PRG to set a method header for any new method created 
*	by PEM Editor.  It can be modified as needed to fit your own needs.
*	
*	For PEM Editor to use this PRG, it should be moved into the folder
*	where PEM Editor is installed.
*	

Lparameters lcMethodName, lcDescription, lnVisibility  

Local lbNew
SET TEXTMERGE ON
Set Century ON
TEXT TO lcResult NOSHOW
* Added: <<DTOC(Date())>>
*========================

ENDTEXT
SET TEXTMERGE OFF

Return lcResult