### IDE Tools: Plug-Ins

A number of plug-ins are available to alter the behavior of PEM Editor. Samples of each of them are provided in the sub-folder "**Sample Plug-Ins**." 

Any plug-in can be enabled by either:

*   Copying it into the sub-folder "Live Plug-Ins". 
*   Adding _PEME__ to the name of the plug-in (_PEME_GoToDefinition_, for instance) and saving it in your path.

Each sample PRG is amply commented, describing parameters and results.

#### AutoRenameControl: 
Determines names for objects when they are automatically renamed.  All or part of the default behavior may be over-ridden.
#### CloseControlStructure
Determines the text to be used when closing a control structure (EndIF, EndWith, etc.) May be used to include the text from the opening of the control structure, if desired:
```foxpro
If This.lProperty  
  blablabla  
EndIf This.lProperty
```
#### CreateLocalsStatements

Determines which variables assigned in a procedure are to be included in the LOCAL statements and also the format of the LOCAL statement.  Called with an array of all possible locals.

#### GoToDefinition

Called ***after*** all other searches in GoToDefinition have failed, this allows further customization, such as trying to open the table for the name under the cursor.