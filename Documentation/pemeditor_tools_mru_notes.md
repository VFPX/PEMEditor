### Managing MRU lists and Source Code Control

This may seem like an unlikely pairing of features.  We’ll get to that in a bit.

### MRU Lists

MRU lists are the long-forgotten step-child of FoxPro.  Effective tools when all work was done from the command window, they have fallen into disuse as reliance on the command window waned – the MRU lists are only accessible from the command window and are only updated for files opened from the command window.

This began to change with the introduction of PEM Editor.  All forms and classes that are opened when PEM Editor is open are automatically added to their appropriate MRU Lists.  PEM Editor even introduced a number of methods (now Thor tools) that either make the MRU lists available without using the command window ([MRUs](pemeditor_tools_mrus.md)), or provide alternate ways to open files or classes which, as a side benefit, also automatically update the MRU lists ([Open Files](pemeditor_tools_open_files.md) and [Favorites](pemeditor_tools_favorites.md)).

**Source Code Control**

There are two distinct issues which may be of interest to those who use Source Code Control (SCC)

1.  They ***may*** wish to be asked whether they want to check out files from their SCC Provider whenever the files are being opened.
2.  And some SCC Providers are case-sensitive with regards to file names, and FoxPro is notorious for being very flippant about changing the case of file names.

**To the rescue …**

All files opened by PEM Editor and [GoFish](https://github.com/mattslay/GoFish) use the same method to open files, and this method performs three additional tasks:

1.  It adds the file to the appropriate MRU list (and, if the file is a class library, maintains a separate MRU list only for class libraries).
2.  It opens the file with the same case as it exists on disk, which implies that the case of the name will not be changed when it is saved.
3.  If the project uses Source Code Control and if the user desires, the user can be prompted as to whether to also check out the file from SCC.

This is not enough, though.  Many people open their files using the project manager.  These same capabilities listed above can also be made available for any file opened by the Project Manager by placing the following code in the QueryModifyFile Event of the ProjectHook class for the project:

```foxpro
* Set the following to .T. if you want to be prompted to check out files from SCC  
#Define clCheckOutFromSourceCodeControl .F.

Lparameters oFile, cClassName
Local llSuccess, loTools

Try
	* see	https://github.com/VFPX/Thor/blob/master/Docs/Thor_tools_object.md
	loTools	  = Execscript (_Screen.cThorDispatcher, 'class= tools from pemeditor')
	llSuccess = loTools.EditSourceX (oFile.Name, cClassName, , , , clCheckOutFromSourceCodeControl)
Catch
	llSuccess = .F.
Endtry

If llSuccess
	Nodefault
Endif
```
