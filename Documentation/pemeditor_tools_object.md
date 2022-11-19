The **PEM Editor Tools Object**, documented here, contains a number of useful tools of value in building tools in Thor.

It is always available by executing the following line of code (which uses the Thor Dispatcher).  Note that the PEM Editor form need not be open for this to work.

```foxpro
loTools = ExecScript(_Screen.cThorDispatcher, 'class= tools from pemeditor')
```
Methods (Parameters)|Description
---|---
AddMRUFile (tcFileName, tcClassName)|Adds a file to its appropriate MRU list (in the FoxPro resource file). If the file is a class library, but no class name is supplied, adds the file to the MRU list for class libraries (unique to PEM Editor)
CloseForms()|Close the PEM Editor and Document TreeView forms, if open
CreateNewPEM (tcType, tcName, txValue)|Creates a new property or method: tcType = 'P' for Property, 'M' for Method tcName = Name for new PEM (_MemberData if appropriate txValue = Value (for properties) or method code (for methods)
DiskFileName (tcFileName)|Returns the file name as it is stored on disk (that is, with current upper/lower case).
EditMethod (toObject, tcMethodName)|Opens a method (or event) for editing. {toObject} may be an object reference, .T. for the current form or class, or empty for the current object.
EditSourceX (tcFileName, tcClass, tnStartRange, tnEndRange)|Opens any file for editing (as **EditSource** does), with additional capabilities: The file is added to its appropriate MRU library. The file is opened with the correct case for the file name, so that when it is saved the case for the file name will not be changed. If the file is a class library, and no class name is supplied, the class browser is opened. You will be asked whether you want to check out the file from Source Code Control (if you use SCC and you have marked the appropriate item in the Preferences file.)
FindObjects (tcSearchText)|Finds all objects matching the search criteria in {tcSearchText}. The search criteria are the same as are specified by the 'Find' (binoculars) search button. The result is a collection, where each item in the collection is an object with two properties: Object - a reference to the object FullObjectName - the full path name to the object _For example, **.FindObjects ('Exists ("ControlSource")' )** returns a collection of all objects having a ControlSource._
FindProcedure (tcName)|Finds a PRG named {tcName}, or a procedure or function named {tcName} within a PRG, or a constant named {tcName}, opens the file for editing, and highlights the searched-for name.
GetCurrentObject (tlTopOfForm)|Returns the currently selected object or form/class
GetPEMList (toObject, tcTypes)|Returns a collection of the names of PEMs for an object. {toObject} may be an object reference, .T. for the current form or class, or empty for the current object. {tcTypes} may be one or more of 'P' (for properties), 'E' (for Events), or 'M' (for Methods); or, if empty or missing, the collection will contain all PEMs.
GetMRUList (tcName)|Returns a collection of file names in a MRU list. {tcName} may be a file name, a file extension, or the actual MRU-ID (if you know it)
UseHighlightedTable (lnDataSession)|Selects the table name that is currently highlighted (or, if nothing is highlighted, the name under the cursor). If that alias is not in use, calls **PEME_OpenTable.PRG**, if it exists, to open the table. A sample of **PEME_OpenTable.PRG** is found in the 'Sample Plug-ins' folder of PEM Editor.  To make this PRG available, copy it anywhere on your path, or into the 'Live Plug-Ins' folder of PEM Editor.
