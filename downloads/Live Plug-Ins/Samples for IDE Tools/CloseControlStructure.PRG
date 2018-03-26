*==============================================================================
* Procedure:        CloseControlStructure
* Purpose:          Returns the text to be used as the closing line of code
*                      
* Author:           Jim Nelson
* Parameters:       
*   lcType        = First word or control structure ('IF', 'TRY', 'FOR', etc.)
*   lcNewText     = Proposed closing word ('EndIf', 'EndTry', 'EndFor', etc.)
*   lcStartLine   = The remainder of the first line of the control structure
*   lcAllText     = The entire text of the control structure
*
* Returns:          The default value
* Added:            04/06/2011
*==============================================================================
* Notes:
*   This is a sample PRG to set the text to be used as the closing line of code
*   in a control structure.  It is called by 'Close Control Structure'.  
*   It can be modified as needed to fit your own needs.  (See the defined constants.)
*   


* minimum # of lines in the control structure to even bother with
#Define     MinLineCount    3

* For structure  'For lnValue ='  ... Use Next lnValue ??
#Define     UseNext         .F.

* For all other structures, put remainder of line in comments?
#Define     UseComments     .F.



Lparameters lcType, lcNewText, lcStartLine, lcAllText

Local laLines[1], lnLineCount
lnLineCount = Alines (laLines, lcAllText)

Do Case

        * Don't bother if this control structure is too short
    Case lnLineCount < MinLineCount
        Return lcNewText

        * TRY and DO CASE
    Case InList(lcType, 'TRY', 'DOCASE')
        Return lcNewText

        * For lnValue =  ... Use Next lnValue
    Case UseNext and lcType = 'FOR' And Upper (Getwordnum (lcStartLine, 1)) # 'EACH'
        Return 'Next ' + Getwordnum (lcStartLine, 1)

        * With / EndWith requires a comment
    Case lcType = 'WITH'
        Return lcNewText + ' &' + '& ' + lcStartLine

        * Use comments?
    Case UseComments
        Return lcNewText + ' &' + '& ' + lcStartLine

        * Or leave as code
    Otherwise
        Return lcNewText + ' ' + lcStartLine

Endcase
