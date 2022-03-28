*** JIM: remove the second parameter and use "This" instead of toPEM everywhere it occurs
lparameters tcName, ;
	toPEM
local laObjects[1], ;
	lcObjectFile, ;
	lcInclude, ;
	loInclude, ;
	lcMethodCode, ;
	lcName, ;
	laLines[1], ;
	lnLines, ;
	lnI, ;
	lcIncludeFile, ;
	llFound

* Get the include file for the form/class being edited.

if aselobj(laObjects, 3) = 0 or empty(laObjects[3])
	return .F.
endif aselobj(laObjects, 3) = 0 ...
lcObjectFile = laObjects[2]
lcInclude    = laObjects[3]

* Create an object to hold return values.

loInclude = createobject('Empty')
addproperty(loInclude, 'File', '')
addproperty(loInclude, 'Line', 0)

* Start with the current method.

if toPEM.oEditorWin.FindWindow() < 0
	return .F.
endif toPEM.oEditorWin.FindWindow() < 0
*** JIM: this gives me no code
lcMethodCode = upper(toPEM.oEditorWin.Getstring(0, 10000000))
lcName       = upper(tcName)
do case

* If the current method contains #DEFINE <constant>, we're done.

	case '#DEFINE ' + lcName $ lcMethodCode
		lnLines = alines(laLines, lcMethodCode)
		for lnI = 1 to lnLines
			if laLines[lnI] = '#DEFINE ' + lcName
				loInclude.Line = lnI
				exit
			endif laLines[lnI] = '#DEFINE ' + lcName
		next lnI

* If the current method contains #INCLUDE, process those files.

	case '#INCLUDE' $ lcMethodCode
		for lnI = 1 to occurs('#INCLUDE', lcMethodCode)
			lcIncludeFile = strextract(lcMethodCode, '#INCLUDE ', chr(13), ;
				lnI, 2)
			lcIncludeFile = fullpath(lcIncludeFile, lcObjectFile)
*** JIM: make this a method call
			ProcessInclude(lcIncludeFile, lcName, loInclude)
			if loInclude.Line > 0
				exit
			endif loInclude.Line > 0
		next lnI
endcase

* If it wasn't found, process the include file for the form/class.

if loInclude.Line = 0 and not empty(lcInclude)
*** JIM: make this a method call
	ProcessInclude(lcInclude, lcName, loInclude)
endif loInclude.Line = 0 ...
do case

* We didn't find it, so see if we can find a program. If so, edit it.

	case loInclude.Line = 0
*** JIM: make this a method call
		FindProgram(lcName, loInclude)
		if not empty(loInclude.File)
			llFound = .T.
			editsource(loInclude.File, loInclude.Line)
		endif not empty(loInclude.File)

* We found it in an include file, so edit it.

	case not empty(loInclude.File)
		llFound = .T.
		editsource(loInclude.File, loInclude.Line)

* We found it in the current method.

	otherwise
		llFound = .T.
*** JIM: remove next line and handle this case
messagebox('Found at ' + transform(loInclude.Line))
endcase
return llFound


* Process the specified include file.

function ProcessInclude(tcIncludeFile, tcName, toInclude)
*** 11/18/2010: change
*!*	local lcInclude, ;
*!*		laLines[1], ;
*!*		lnLines, ;
*!*		lnI, ;
*!*		lcIncludeFile
local lcInclude, ;
	lnLine, ;
	lnI, ;
	lcIncludeFile
*** 11/18/2010: end of change

* If the specified file exists, read its contents.

if file(tcIncludeFile)
	lcInclude = upper(filetostr(tcIncludeFile))
*** 11/18/2010: new code
	lnLine    = LocateLine(lcInclude, tcName, 1)
*** 11/18/2010: end of new code
	do case

* If the specified constant exists, find the line it's on and save that to
* the return object.

*** 11/18/2010: change
*!*		case '#DEFINE ' + tcName $ lcInclude
*!*			lnLines = alines(laLines, lcInclude)
*!*			for lnI = 1 to lnLines
*!*				if laLines[lnI] = '#DEFINE ' + tcName
*!*					toInclude.File = tcIncludeFile
*!*					toInclude.Line = lnI
*!*					exit
*!*				endif laLines[lnI] = '#DEFINE ' + tcName
*!*			next lnI
		case lnLine > 0
			toInclude.File = tcIncludeFile
			toInclude.Line = lnLine
*** 11/18/2010: end of change

* If this include file #INCLUDEs one of more others, process each of them. If
* one of them have the constant, we're done.

		case '#INCLUDE' $ lcInclude
			for lnI = 1 to occurs('#INCLUDE', lcInclude)
				lcIncludeFile = strextract(lcInclude, '#INCLUDE ', chr(13), ;
					lnI, 2)
				lcIncludeFile = fullpath(lcIncludeFile, tcIncludeFile)
*** JIM: make this a method call
				ProcessInclude(lcIncludeFile, tcName, toInclude)
				if toInclude.Line > 0
					exit
				endif toInclude.Line > 0
			next lnI
	endcase
endif file(tcIncludeFile)
return


* See if we can find a PRG or procedure of the specified name.

function FindProgram(tcName, toInclude)
local lcFile, ;
	llFound, ;
	loFile, ;
	lcProcedures, ;
	laLines[1], ;
	lnLines, ;
	lnI

* First check for a PRG in the path.

lcFile = tcName + '.PRG'
if file(lcFile)
	toInclude.File = lcFile
	llFound = .T.
endif file(lcFile)

* Next check all PRGs in the active project. We may have to look in procedure
* files.

if not llFound and vartype(_vfp.ActiveProject) = 'O'
	for each loFile in _vfp.ActiveProject.Files
		if loFile.Type = 'P'
*** JIM: make this a method call
			ProcessPRG(loFile.Name, tcName, toInclude)
			if not empty(toInclude.File)
				llFound = .T.
				exit
			endif not empty(toInclude.File)
		endif loFile.Type = 'P'
	next loFile
endif not llFound ...

* Finally, check all open procedure files.

if not llFound
	lcProcedures = set('Procedure')
	lcProcedures = strtran(lcProcedures, ';', ',')
	lnLines      = alines(laLines, lcProcedures, 4, ',')
	for lnI = 1 to lnLines
*** JIM: make this a method call
		ProcessPRG(laLines[lnI], tcName, toInclude)
		if not empty(toInclude.File)
			exit
		endif not empty(toInclude.File)
	next lnI
endif not llFound
return


* See if the specified PRG either has the name we're looking for or
* contains a function or procedure of that name.

function ProcessPRG(tcPRG, tcName, toInclude)
*** 11/18/2010: change
*!*	local lcFile, ;
*!*		lcCode, ;
*!*		laLines[1], ;
*!*		lnLines, ;
*!*		lnI, ;
*!*		lcLine
local lcFile, ;
	lcCode, ;
	lnLine
*** 11/18/2010: end of change
lcFile = tcName + '.PRG'
do case
	case not file(tcPRG)
	case upper(justfname(tcPRG)) = lcFile
		toInclude.File = tcPRG
	otherwise
		lcCode = upper(filetostr(tcPRG))
*** 11/18/2010: change
*!*			if 'PROCEDURE ' + tcName $ lcCode or ;
*!*				'PROC ' + tcName $ lcCode or ;
*!*				'FUNCTION ' + tcName $ lcCode or ;
*!*				'FUNC ' + tcName $ lcCode
*!*				lnLines = alines(laLines, lcCode)
*!*				for lnI = 1 to lnLines
*!*					lcLine = laLines[lnI]
*!*					if lcLine = 'PROCEDURE ' + tcName or ;
*!*						lcLine = 'PROC ' + tcName or ;
*!*						lcLine = 'FUNCTION ' + tcName or ;
*!*						lcLine = 'FUNC ' + tcName
*!*						toInclude.File = tcPRG
*!*						toInclude.Line = lnI
*!*						exit
*!*					endif lcLine = 'PROCEDURE ' + tcName ...
*!*				next lnI
*!*			endif 'PROCEDURE ' + tcName $ lcCode ...
		lnLine = LocateLine(lcCode, tcName, 2)
		if lnLine = 0
			lnLine = LocateLine(lcCode, tcName, 3)
		endif lnLine = 0
		if lnLine > 0
			toInclude.File = tcPRG
			toInclude.Line = lnLine
		endif lnLine > 0
*** 11/18/2010: end of change
endcase
return

*** 11/18/2010: new code
function LocateLine(tcCode, tcName, tnPattern)
local loRegExp, ;
	loMatches, ;
	lnReturn

* Create and set up a RegExp object.

loRegExp = createobject('VBScript.RegExp')
loRegExp.IgnoreCase = .T.
loRegExp.Global = .T.

* Decide whether we're looking for #DEFINE, PROC, or FUNC based on tnPattern:
*	1 = #DEFINE
*	2 = PROCEDURE
*	3 = FUNCTION

do case
	case tnPattern = 1
		loRegExp.Pattern = '#DEFINE\s+' + tcName
	case tnPattern = 2
		loRegExp.Pattern = 'PROC(EDURE)?\s+' + tcName
	case tnPattern = 3
		loRegExp.Pattern = 'FUNC(TION)?\s+' + tcName
endcase

* Try to find matches. If not, return 0. Otherwise, figure out what line it's
* on by counting CRLF.
  
loMatches = loRegExp.Execute(tcCode)
if loMatches.Count = 0
	lnReturn = 0
else
	lnReturn = occurs(chr(13), substr(tcCode, 1, loMatches.Item(0).FirstIndex)) + 1
endif loMatches.Count = 0
return lnReturn
*** 11/18/2010: end of new code
