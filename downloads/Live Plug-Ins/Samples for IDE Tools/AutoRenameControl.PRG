*==============================================================================
* Procedure:        AutoRenameControl
* Purpose:          Returns the new name to be uses when auto-renaming controls
* Author:           Jim Nelson
* Parameters:       
*   loControl     = the control being renamed
*   lcNewName     = the default name that PEM Editor would assign to this control
* Returns:          The name that will be used for this control (empty means "change nothing")
* Added:            03/05/2009
*==============================================================================
* Notes:
*   This is a sample PRG to assign some different default names for some controls.
*
*   Important Note:  Since there is always the possibility of duplicate names,
*   PEM Editor, like VFP will attach a numeric suffix to this name if necessary.
*
*   Normally, the names created will be txt, txt1, txt2, etc ... but you can force
*   the values txt1, txt2, txt3 by including '#' in the name (see below)
*   

Lparameters loControl, lcNewName

*** JRN 03/05/2009 : Default behavior
Local lcBaseClass
lcBaseClass = Lower (loControl.BaseClass)

Do Case

    Case lcBaseClass = 'grid'
        lcNewName = 'grdGrid'

    Case lcBaseClass = 'column'
        lcNewName = 'col#'

    Case lcBaseClass = 'header'
        lcNewName = 'hdr'

    Case lcBaseClass = 'pageframe'
        lcNewName = 'Frame'

    Case lcBaseClass = 'page'
        lcNewName = 'Pg#'

Endcase


Return lcNewName