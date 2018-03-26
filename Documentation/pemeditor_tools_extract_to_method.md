### IDE Tools: Extract To Method

This tool is used to extract highlighted text into a new method in the form or class being edited, as follows:

*   Highlight the desired text.
*   Run this tool.
*   You will be prompted for the name of the method to be created. You may use mixed case names; _MemberData will be updated appropriately.  This will cause the following:
    *   The new method will be created.
    *   The new method will be opened, populated with the highlighted text.
    *   In the original method, the highlighted text will be replaced with the appropriate reference to the new method, one of:
        *   Thisform.NewMethod()
        *   This.NewMethod()
        *   This.Parent.Parent … .NewMethod()