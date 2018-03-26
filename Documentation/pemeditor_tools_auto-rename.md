### IDE Tools: Auto-Rename

Applies to tools:

*   _Auto-rename current object_
*   _Auto-rename current object and all its child objects_                    

In the VFP Help File, the section entitled 'Object Naming Conventions' provides a recommended format for naming objects - a three letter prefix, based on the type of object, following by some meaningful name. Curiously, when adding new objects to a form or class, VFP does not use this convention by using the suggested three letter prefix.

This tool provides the capability to rename objects so that they use this naming convention (that is, to rename all those objects that still have the default name created for them by VFP). The name created will have the recommended prefix, followed by all or part of the Caption, ControlSource, or RecordSource (if they exist), or else a number; in any case, the new names will be uniquely assigned.

Notes:

*   For tool  _Auto-rename current object and all its child objects_,  objects are only renamed if they still have the default name created for them by VFP.  (There is a pop-up that notifies you of any objects that were not renamed).
*   This option does <span style="text-decoration: underline;">not</span> replace references to objects in any method code. Any such changes would have to be made manually. Thus, it is recommended that this option be used very early, before there are any references to the objects in method code.

**Customization:** There is a plug-In PRG that allows you to modify the behavior of this tool to fit your own needs.  The plug-in is called with the suggested new name, which can be modified if desired.  See [Plug-Ins](pemeditor_tools_plug-in_prgs.md).