### IDE Tools: Copy properties and methods

This feature copies properties and methods from the currently selected object into a "clipboard" for later pasting or comparing.

The copied properties and methods can be used a number of different ways:

*   By selecting a different object and then using [Compare with copied object](pemeditor_tools_compare_objects.md)
*   By selecting a different object and then using [Paste properties and method code](pemeditor_tools_paste_properties.md)
*   By using [Paste Object](pemeditor_tools_paste_object.md) to create a new object.

The "different object" can be another object on the same form or class, in a different form or class, or in the parent class.

This feature is also used as a final step when copying properties and method code from an object into its parent class during re-factoring.  See [Paste properties and method code](pemeditor_tools_paste_properties.md).

A few notes of interest:

*   This tool allows you to copy any object, even if VFP cannot copy it (the familiar message "Cannot copy objects because some are members of a parent class").  In this case, the tool actually digs into the parent class that the object is a member of, and extracts all non-default properties and methods from there as well. 
*   You can copy an entire class when editing to simplify later pasting:
    *   Use this tool on the class
    *   Close the class.
    *   Open the form or class to paste into.
    *   Use [Paste Object](pemeditor_tools_paste_object.md).
*   The "clipboard" that the properties and method code gets pasted into survives even if you close PEM Editor; they will survive as long as 'Clear All' is not used, which removes the object containing them from memory.
*   Unlike VFP Copy, this tool can only copy a single object.
