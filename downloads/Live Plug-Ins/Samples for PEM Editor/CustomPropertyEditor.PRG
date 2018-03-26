*==============================================================================
* Procedure:		CustomPropertyEditor
* Purpose:			Invokes custom property editors
* Author:			Jim Nelson
* Parameters:
*	lcAppPath	  = Path to default folder for property editors
*	lcProperty	  = Name of Property
*   lxValue		  = Current Value of the Property (see also notes)
*	lcVarType	  = VarType of the value (see also notes)
*   llSingleValue = See Notes
*
* Returns:			- An object with following values:
*					  loResult.Cancelled = .T.  if editor invoked, but cancelled
*					  loResult.Value     = value return by editor, to be applied to all properties
*
* Added:			03/24/2009
*==============================================================================
* Notes:
*	This is a sample PRG invoke custom property editors.  The samples provided have already
*   been implemented as part of PEMEditor, on the Preferences form, so there's no reason to use
*   them -- except to demonstrate how new custom property editors could be implemented.  
*
* This example performs three different editors
*    1) Toggling of properties that are logical-valued
*    2) Demo of invoking native VFP anchor editor.
*    2) Demo of invoking df Anchor Builder.
*
* Note that both of the anchor editors provided here have been modified from the original, so
* that they can accept the parameters provided and return a result, rather than actuall updating
* and properties themselves.  
*
* The property being edited may be applied to a single control or multiple controls
* If multiple controls,
*		<lxValue> 	is the value of the property for the first control.
*    	<lcVarType> is a list of all the VarType(s) across all properties,
*		<llSingleValue> is .T. if all the values are identical
*
*   It can be modified as needed to fit your own needs.
*
*	For PEM Editor to use this PRG, it must be moved into the folder "Live Plug-Ins"
*

Lparameters lcAppPath, lcProperty, lxValue, lcVarType, llSingleValue
Local loAnchorEditor As 'AnchorEditorPEME' Of  lcAppPath + 'AnchorEditor\AnchorEditor.vcx'
Local loResult As 'Empty'
Local lnAnchor

loResult = Createobject ('Empty')
AddProperty (loResult, 'Cancelled', .F.)
AddProperty (loResult, 'Assigned', .T.)
AddProperty (loResult, 'Value')

Do Case

		* toggle logical values
	Case lcVarType == 'L'
		If llSingleValue
			loResult.Value = Not lxValue
			Return loResult
		Else
			loResult.Value = .F. && if mixed values, turn 'em all off
			Return loResult
		Endif

		* Anchor editor (VFP Native)
		* 	Invoked form must be modal, return result of the new value.
		* 	Return this value if it is meant to be used; else fall thru
	Case Proper (lcProperty) = 'Anchor'
		loAnchorEditor = Newobject ('AnchorEditorPEME', lcAppPath + 'AnchorEditor\AnchorEditor.vcx', .Null., lxValue)
		loAnchorEditor.Show()
		If Vartype (loAnchorEditor) = 'O'
			loResult.Value = loAnchorEditor.uValue && negative on 'abort'
			loAnchorEditor.Release()
		Else
			loResult.Cancelled = .T.
		Endif
		Return loResult

		* df Anchor builder
	Case Proper (lcProperty) = 'Anchor'
		Do Form (lcAppPath + 'BuildAnchor\BuildAnchorPEME.scx') With lxValue To lnAnchor
		If lnAnchor >= 0
			loResult.Value = lnAnchor
		Else
			loResult.Cancelled = .T.
		Endif
		Return loResult

Endcase

loResult.Assigned = .F.
Return loResult
