*==============================================================================
* Procedure:        OpenVCXFile
* Purpose:          Opens a VCX file for Browsing (e.g., using Class Browser)  
* Author:           Jim Nelson
* Parameters:       
*   tcFileName    = Full name of the VCX file to be opened
*
* Returns:          The default value
* Added:            08/17/2011
*==============================================================================
* Notes:
*   This is a sample PRG to open a VCX for browsing.  
*   It can be modified as needed to fit your own needs.
*   

Lparameters tcFileName
Do(_Browser) With (tcFileName)
