All of the IDE features of PEM Editor use the functions from **FoxTools.fll** to access and modify the currently open editing window.

These tools can be of value in creating tools that are accessible through Thor. The methods in this object serve two purposes:

1.  As wrappers for the function in FoxTools
2.  To simplify the task of identifying the currently open editing window.

Editing windows are identified by a numeric value called a handle. All of the IDE features in PEM Editor first call **oEditorWindow.FindWindow()**, which determines the handle for the currently open window and saves it.

Thereafter, all of the other methods of this object refer to this saved window handle so that it need not be explicitly referenced in their parameters.

Some further notes:

1.  Character positions and line counts start at 0, not 1. (i.e., be careful)
2.  There is full documentation for the functions in FoxTools.fll available, in a sub-folder of PEM Editor's Source folder.

It is always available by executing the following line of code (which uses the Thor Dispatcher).  Note that the PEM Editor form need not be open for this to work.


```foxpro
loEditorWin = ExecScript(_Screen.cThorDispatcher, 'class= editorwin from pemeditor')
```

## Window manipulation: handles, size, position, title, etc.

Methods (Parameters)|Description
---|---
CloseWindow()|Close the current window
FindLastWindow()|Returns the handle of the most recently used window which is either of a PRG or method code from the Form or Class Designer.
FindWindow()|Saves the handle for the currently active window, and returns its window type:  
"|0 – Command Window, Form and Class Designers, other FoxPro windows  
"|1 – Program file (MODIFY COMMAND)  
"|2 – Text Editor (MODIFY FILE)  
"|8 – Menu code edit window  
"|10 – Method code edit window of the Class or Form Designer  
"|12 – Stored procedure in a DBC (MODIFY PROCEDURE)  
"|-1 - None
GetHeight()|Returns the height of the editing window, in pixels.
GetLeft()|Returns the left position of the editing window, in pixels.
GetOpenWindows()|Returns a collection of the handles of all open windows, most recently used first.
GetTitle()|Returns the title for the current window
GetTop()|Returns the top position of the editing window, in pixels.
GetWidth()|Returns the width of the editing window, in pixels.
GetWindowHandle()|Returns the handle of the current editing window
MoveWindow (tnLeft, tnTop)|Moves the editing window to position {tnLeft}, {tnTop}. Both are in pixels.
ResizeWindow (tnWidth, tnHeight)|Resizes the editing window to {tnWidth} by {tHeight}. Both are in pixels.
SelectWindow (tnHandle)|Selects (brings to the foreground) window with handle {tnhandle}
SetHandle (tnHandle)|Sets the handle (used to indicate the window being referenced in most of these commands/)
SetTitle (tcNewTitle)|Sets the title for the editing window to {tcNewTitle}


## Text manipulation
Methods (Parameters)|Description  
---|---
Copy()|Copies the currently highlighted text into the clipboard
Cut()|Cuts the currently highlighted text
EnsureVisible (tnPosition, tlScroll)|Ensures that the character at position {tnPosition} is visible in the editing window. If {tlScroll} is true, it is brought to the mid-point of the editing window
GetCharacter (tnPosition)|Returns the character at position {tnPosition}
GetEnvironment {tnIndex}|Returns a single environment setting. {tnIndex} takes values from 1 to 25. See _EdGetEnv in the help for FoxTools for a description of all the settings. Frequently used values are:  
"|2 – File Size  
"|17 – Selection start  
"|18 – Selection end  
"|25 – Window Type (see above)
GetFileSize()|Returns the file size
GetLineNumber (tnSelStart)|Returns the line number for the character at position {tnPosition}
GetLineStart (tnSelStart, tnLineOffset)|Determines the line number for the character at position {tnPosition} and returns the position for the character at the beginning of that line. If {tnLineOffset} is supplied, it first offsets the line number by that amount. Thus .GetLineStart (tnSelStart, 1) gives the start position of the next line after the line for {tnSelStart}
GetSelEnd()|Returns the position for the end of the currently highlighted text
GetSelStart()|Returns the position for the start of the currently highlighted text
GetString (tnSelStart, tnSelEnd)|Returns the string of characters from position {tnSelStart} through {tnSelEnd}
Paste()|Pastes the contents of the clipboard into the editing window
Select (tnSelStart, tnSelEnd)|Selects (highlights) the string of characters from position {tnSelStart} through {tnSelEnd}
SetInsertionPoint (tnPosition)|Sets the insertion point to {tnPosition}
