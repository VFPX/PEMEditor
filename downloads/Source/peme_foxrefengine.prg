*** JRN 03/24/2010 : Unabashedly lifted from Code References
*   With modifications as needed (and noted)

* Abstract...:
*	Primary class for Code References application.
*
* Changes....:
*
#include "foxpro.h"
#include "PEME_foxref.h"

Define Class FoxRef As Session
	Protected lIgnoreErrors As Boolean
	Protected lRefreshMode
	Protected cProgressForm
	Protected lCancel
	Protected tTimeStamp
	Protected lIgnoreErrors

	Name = "FoxRef"

	* search match engines
	MatchClass		   = "MatchDefault"
	MatchClassLib      = "PEME_foxmatch.prg"
	WildMatchClass     = "MatchClass"
	WildMatchClassLib  = "PEME_foxmatch.prg"

	* default search engine for Open Window
	FindWindowClass    = "RefSearchWindow"
	FindWindowClassLib = "FoxRefSearch_Window.prg"

	oSearchEngine   = .Null.

	WindowHandle    = -1
	WindowFilename  = ''
	WindowLineNo    = 0


	Comments        = COMMENTS_INCLUDE
	MatchCase       = .F.
	WholeWordsOnly  = .F.
	ProjectHomeDir  = .F.  && True to search only files in Project's Home Directory or below

	SubFolders      = .F.
	Wildcards       = .F.
	Quiet           = .F.  && quiet mode -- don't display search progress
	ShowProgress    = .F.  && show a progress form

	Errors          = .Null.

	FileTypes       = '|'
	ReportFile      = REPORT_FILE

	XSLTemplate     = "foxref.xsl"

	* MRU array lists
	Dimension aLookForMRU[10]
	Dimension aClassMRU[10]
	Dimension aReplaceMRU[10]
	Dimension aFolderMRU[10]
	Dimension aFileTypesMRU[10]
	Dimension aDefaultFileTypes[1]

	aLookForMRU     = '?'
	aReplaceMRU     = ''
	aFolderMRU      = ''
	aFileTypesMRU   = ''

	SearchType			= ''
	Pattern             = ''
	OverwritePrior      = .T.
	ConfirmReplace      = .T.  && confirm each replacement
	BackupOnReplace     = .T.  && false to not backup when doing global replace
	DisplayReplaceLog   = .T.  && create activity log for replacements
	PreserveCase        = .F.  && preserve case during a Replace operation
	FileNameExactMatch	= .F.  && exact match to file name
	ClassSearch			= ''
	ClassExactMatch		= .F.
	SearchBaseClass		= ''
	HomeDir				= ''

	FoxRefDirectory = ''
	RefTable        = ''
	DefTable        = ''
	FileTable       = ''
	AddInTable      = ''
	ProjectFile     = ''
	FileDirectory   = ''
	Templates 		= ''

	ActivityLog     = ''

	* The following are set by the Options dialog
	* There should be a corresponding entry in FoxRefOption.DBF
	* (except for BackupStyle & FontString)
	IncludeDefTable     = .T.  && create Definition table when searching
	CodeOnly            = .F.  && search only source code & expressions (not names and other none-code items)
	FormProperties      = .T.  && search form/class property names & values
	AutoProjectHomeDir  = .F.  && True to search only files in Project's Home Directory or below when doing definitions automatically
	ShowRefsPerLine     = .F.  && display a column in search results that depicts number of references found on the line
	ShowFileTypeHistory	= .F.  && True to keep filetype history in addition to showing common filetypes in search dialog
	ShowDistinctMethodLine = .F. && True to show columns for method/line apart from Class
	SortMostRecentFirst = .F.

	BackupStyle         = 1    && 1 = "filename.ext.bak"   2 = "Backup of filename.ext"
	FontString          = FONT_DEFAULT

	* XML Export Options
	XMLFormat          = XMLFORMAT_ELEMENTS
	XMLSchema          = .T.

	* This is the SetID for the last Replacement Log after it's saved
	ReplaceLogSetID    = ''

	* properties used internally
	cSetID              = ''
	lRefreshMode        = .F.
	lIgnoreErrors       = .F.
	oProgressForm       = .Null.
	lCancel             = .F.
	tTimeStamp          = .Null.
	lDefinitionsOnly    = .F.

	* collection of files we've backed up in this session
	oBackupCollection   = .Null.

	oFileCollection     = .Null.
	oSearchCollection   = .Null.
	oProcessedCollection = .Null.

	oEngineCollection   = .Null.
	oReportCollection   = .Null.

	oFileTypeCollection = .Null.

	cTalk           = ''
	nLangOpt        = 0
	cEscapeState    = ''
	cSYS3054        = ''
	cSaveUDFParms   = ''
	cSaveLib        = ''
	cExclusive      = ''
	cCompatible     = ''


	oOptions        = .Null.
	lInitError      = .F.

	oProjectFileRef = .Null.

	cResourceFile	= ''
	cResultsTable	= ''


	Procedure Init (lcResourceFile)

		Local nSelect
		Local oException
		Local cAddInType
		Local nMemoWidth

		This.cTalk = Set("TALK")
		Set Talk Off
		Set Deleted On


		This.cCompatible = Set("COMPATIBLE")
		Set Compatible Off

		This.cExclusive = Set("EXCLUSIVE")
		Set Exclusive Off

		This.nLangOpt = _vfp.LanguageOptions
		_vfp.LanguageOptions = 0

		This.cEscapeState = Set("ESCAPE")
		Set Escape Off

		This.cSYS3054 = Sys(3054)
		Sys(3054,0)

		This.cSaveLib      = Set("LIBRARY")

		This.cSaveUDFParms = Set("UDFPARMS")
		Set Udfparms To Value

		Set Exact Off

		*** JRN 03/24/2010 : Resource File Name comes as parameter
		This.cResourceFile = lcResourceFile

		This.RestorePrefs()

		This.CreateResultsTable()

	Endfunc


	Procedure Destroy()
		Local cCompatible

		Try
			Erase (This.cResultsTable)
		Catch

		Endtry

		If This.cEscapeState = "ON"
			Set Escape On
		Endif
		If This.cTalk = "ON"
			Set Talk On
		Endif
		If This.cExclusive = "ON"
			Set Exclusive On
		Endif
		Sys(3054,Int(Val(This.cSYS3054)))

		_vfp.LanguageOptions = This.nLangOpt

		If This.cSaveUDFParms = "REFERENCE"
			Set Udfparms To Reference
		Endif

		m.cCompatible = This.cCompatible
		Set Compatible &cCompatible
	Endfunc


	Function CreateResultsTable()
		This.cResultsTable = Addbs(Sys(2023)) + Sys(2015) + ".DBF"
		Select 0

		Create Table (This.cResultsTable) Free (	;
			SearchType		C(10),					;
			FullName		M,						;
			Class			C(128),					;
			Folder			C(120),					;
			Filename		C(120),					;
			Parent			C(128),					;
			ClassLoc		C(120),					;
			BaseClass		C(15),					;
			TimeStamp		T						;
			)

		Use

	Endfunc


	Function OpenResultsTable
		Use (This.cResultsTable) Exclusive Alias ResultsTable
		Delete For SearchType = This.SearchType
		Pack
	Endfunc


	* Initialize all of the search engines
	* with our search options
	Procedure SearchInit()
		Local i
		Local oException
		Local nSelect
		Local lSuccess

		nSelect = Select()

		m.oException = .Null.
		m.lSuccess = .T.

		Try
			If This.SearchType = 'Class'
				This.oSearchEngine = Newobject(This.WildMatchClass, This.WildMatchClassLib)
			Else
				This.oSearchEngine = Newobject(This.MatchClass, This.MatchClassLib)
			Endif

			This.oSearchEngine.SearchType 		= This.SearchType
			This.oSearchEngine.SearchBaseClass 	= This.SearchBaseClass
			This.oSearchEngine.ClassSearch		= This.ClassSearch
			This.oSearchEngine.ClassExactMatch	= This.ClassExactMatch
			This.oSearchEngine.SearchBaseClass	= This.SearchBaseClass

			*	m.lSuccess = This.oSearchEngine.InitEngine()
		Catch To oException
			m.lSuccess = .F.
			Messagebox(m.oException.Message, MB_ICONEXCLAMATION, APPNAME_LOC)
		Endtry

		Return m.lSuccess
	Endproc


	Function UpdateLookForMRU(cPattern)
		Local nRow

		nRow = Ascan(This.aLookForMRU, cPattern, -1, -1, 1, 15)
		If nRow > 0
			=Adel(This.aLookForMRU, nRow)
		Endif
		=Ains(This.aLookForMRU, 1)
		This.aLookForMRU[1] = cPattern
	Endfunc

	Function UpdateFolderMRU(cFolder)
		Local nRow

		nRow = Ascan(This.aFolderMRU, cFolder, -1, -1, 1, 15)
		If nRow > 0
			=Adel(This.aFolderMRU, nRow)
		Endif
		=Ains(This.aFolderMRU, 1)
		This.aFolderMRU[1] = cFolder
	Endfunc


	Function UpdateFileTypesMRU(cFileTypes)
		Local nRow

		nRow = Ascan(This.aFileTypesMRU, cFileTypes, -1, -1, 1, 15)
		If nRow > 0
			=Adel(This.aFileTypesMRU, nRow)
		Endif
		=Ains(This.aFileTypesMRU, 1)
		This.aFileTypesMRU[1] = cFileTypes
	Endfunc


	* Abstract:
	*   Set a specific project to display result sets for.
	*	Pass an empty string or "global" to display result sets
	*	that are not associated with a project.
	*
	* Parameters:
	*   [cProject]
	Function SetProject(cProjectFile, lOverwrite)
		Local lSuccess
		Local cRefTable
		Local i
		Local lFoundProject
		Local lOpened
		Local oFileRef
		Local oErr

		lSuccess = .F.

		If Vartype(cProjectFile) <> 'C'
			cProjectFile = This.ProjectFile
		Endif

		This.oProjectFileRef = .Null.
		If Empty(cProjectFile)
			* use the active project if a project name is not passsed
			If Application.Projects.Count > 0
				cProjectFile = Application.ActiveProject.Name
			Else
				cProjectFile = PROJECT_GLOBAL
			Endif

			lSuccess = .T.
		Else
			* make sure Project specified is open
			If cProjectFile == PROJECT_GLOBAL
				lSuccess = .T.
			Else
				cProjectFile = Upper(Forceext(Fullpath(cProjectFile), "PJX"))

				For i = 1 To Application.Projects.Count
					If Upper(Application.Projects(i).Name) == cProjectFile
						cProjectFile = Application.Projects(i).Name
						lSuccess = .T.
						Exit
					Endif
				Endfor
				If !lSuccess
					If File(cProjectFile)
						lOpened = .T.
						* open the project
						Try
							Modify Project (cProjectFile) Nowait

						Catch To oErr
							Messagebox(oErr.Message, MB_ICONEXCLAMATION, APPNAME_LOC)
							lOpened = .F.
						Endtry

						If lOpened
							* search again to find where in the Projects collection it is
							For i = 1 To Application.Projects.Count
								If Upper(Application.Projects(i).Name) == cProjectFile
									cProjectFile = Application.Projects(i).Name
									lSuccess = .T.
									Exit
								Endif
							Endfor
						Endif
					Endif
				Endif

			Endif
		Endif

		If lSuccess
			If Empty(cProjectFile) Or cProjectFile == PROJECT_GLOBAL
				This.ProjectFile = PROJECT_GLOBAL
			Else
				This.ProjectFile = Upper(cProjectFile)
				This.oProjectFileRef = Application.ActiveProject
			Endif

		Endif
		Return lSuccess
	Endfunc


	****************************************************************

	Function Search(cPattern, lShowDialog)
		Local cExtension, cFileTypes, lcTemplate, lSuccess

		This.OpenResultsTable()

		cPattern = Alltrim(This.Pattern)
		cExtension = JustExt(cPattern)
		If not Empty(cExtension)
			cPattern = JustStem(cPattern)
		EndIf 

		Do Case
			Case Empty(cPattern)
				lcTemplate = '*.'
			Case This.FileNameExactMatch
				lcTemplate = cPattern + '.'
			Otherwise
				lcTemplate = '*' + cPattern + '*.'
		Endcase

		Do Case
			Case This.SearchType = 'Class'
				cFileTypes = '|VCX'
			Case This.SearchType = 'Form'
				cFileTypes = '|SCX'
			Case Empty(This.FileTypes)
				cFileTypes = '|*'
			Otherwise 
				cFileTypes = Left (This.FileTypes, Max(0, Len(This.FileTypes) - 1))
		Endcase

		This.Templates = Substr(Strtran (cFileTypes, '|', '|' + lcTemplate),2)

		If This.SetProject(.Null., This.OverwritePrior)
			If This.ProjectFile == PROJECT_GLOBAL Or Empty(This.ProjectFile)
				lSuccess = This.FolderSearch(cPattern)
			Else
				lSuccess = This.ProjectSearch(cPattern)
			Endif

		Endif

		Use In ResultsTable

		Return lSuccess
	Endfunc


	* -- Search a Folder
	Function FolderSearch(cPattern, cFileDir)
		Local i, j
		Local cFileDir
		Local nFileTypesCnt
		Local cFileTypes
		Local lAutoYield
		Local lSuccess
		Local Array aFileList[1]
		Local Array aFileTypes[1]

		If Vartype(cFileDir) <> 'C' Or Empty(cFileDir)
			cFileDir = Addbs(This.FileDirectory)
		Else
			cFileDir = Addbs(cFileDir)
		Endif

		If Empty(cFileDir) Or !Directory(cFileDir)
			Return .F.
		Endif

		If !This.SearchInit()
			Return .F.
		Endif

		lAutoYield = _vfp.AutoYield
		_vfp.AutoYield = .T.

		nFileTypesCnt = Alines(aFileTypes, Evl(This.Templates, '*'), .T., '|')
		lSuccess = This.ProcessFolder(cFileDir, cPattern, @aFileTypes, nFileTypesCnt)

		_vfp.AutoYield = lAutoYield

		This.UpdateLookForMRU(cPattern)
		This.UpdateFolderMRU(cFileDir)
		This.UpdateFileTypesMRU(cFileTypes)

		Return lSuccess
	Endfunc


	* used in conjuction with FolderSearch() for
	* when we're searching subfolders
	Function ProcessFolder(cFileDir, cPattern, aFileTypes, nFileTypesCnt)
		Local nFolderCnt
		Local cFilename
		Local i, j
		Local nFileCnt
		Local nProgress
		Local lSuccess
		Local Array aFileList[1]
		Local Array aFolderList[1]

		cFileDir = Addbs(cFileDir)

		lSuccess = .T.

		For i = 1 To nFileTypesCnt
			If This.lCancel
				Exit
			Endif

			Try
				nFileCnt = Adir(aFileList, cFileDir + aFileTypes[i], '', 1)
			Catch
				nFileCnt = 0
			Endtry
			For j = 1 To nFileCnt
				If This.lCancel
					Exit
				Endif

				cFilename = aFileList[j, 1]
				If !This.FileSearch(cFileDir + cFilename, cPattern)
					Exit
				Endif
			Endfor
		Endfor

		* Process any sub-directories
		If !This.lCancel
			If This.SubFolders
				Try
					nFolderCnt = Adir(aFolderList, cFileDir + "*.*", 'D', 1)
				Catch
					nFolderCnt = 0
				Endtry
				For i = 1 To nFolderCnt
					If !aFolderList[i, 1] == '.' And !aFolderList[i, 1] == '..' And 'D'$aFolderList[i, 5] And Directory(cFileDir + aFolderList[i, 1])
						This.ProcessFolder(cFileDir + aFolderList[i, 1], cPattern, @aFileTypes, nFileTypesCnt)
					Endif
					If This.lCancel
						Exit
					Endif
				Endfor
			Endif
		Endif

		Return lSuccess
	Endfunc



	* -- Search files in a Project
	* -- Pass an empty cPattern to only collect definitions
	Function ProjectSearch(cPattern, cProjectFile)
		Local nFileIndex
		Local nProjectIndex
		Local oProjectRef
		Local oFileRef
		Local cFileTypes
		Local nFileTypesCnt
		Local lAutoYield
		Local lSuccess
		Local nFileCnt
		Local lSuccess
		Local i
		Local cHomeDir
		Local oDefFileTypes
		Local cFilePath
		Local oMatchFileCollection
		Local Array aFileTypes[1]
		Local Array aFileList[1]
		Local cExt, cFile, cFilename, oFileTypesCollection

		If Vartype(cPattern) <> 'C'
			cPattern = This.Pattern
		Endif

		If Vartype(cProjectFile) <> 'C' Or Empty(cProjectFile)
			cProjectFile = This.ProjectFile
		Endif

		If !This.SearchInit()
			Return .F.
		Endif

		lSuccess = .T.

		Do Case
			Case This.SearchType = 'Class'
				cFileTypes = '|VCX|'
			Case This.SearchType = 'Form'
				cFileTypes = '|SCX|'
			Otherwise
				cFileTypes = This.FileTypes
				If 1 >= Len (cFileTypes)
					cFileTypes = '|'
				Endif
		Endcase

		cPattern   = Upper (Alltrim(This.Pattern))

		lAutoYield = _vfp.AutoYield
		_vfp.AutoYield = .T.

		For Each oProjectRef In Application.Projects
			If Upper(oProjectRef.Name) == This.ProjectFile
				cHomeDir = Addbs(Upper(oProjectRef.HomeDir))

				* now process each file in the project that matches our filetypes
				i = 0
				For Each oFileRef In oProjectRef.Files
					If (!This.ProjectHomeDir Or Addbs(Upper(Justpath(oFileRef.Name))) = cHomeDir)
						cFilename = Upper(oFileRef.Name)
						cFile = Juststem (cFilename)
						cExt  = Justext (cFilename)
						If (1 = Len (cFileTypes) Or ('|' + cExt + '|') $ cFileTypes)		;
								and (Empty (cPattern) Or cPattern $ cFile)					;
								and (cPattern == cFile Or Not This.FileNameExactMatch)
							This.FileSearch(oFileRef.Name, cPattern)
						Endif
					Endif
					If This.lCancel
						lSuccess = .F.
						Exit
					Endif
				Endfor

				Exit
			Endif
		Endfor

		_vfp.AutoYield = lAutoYield

		This.UpdateLookForMRU(cPattern)
		This.UpdateFileTypesMRU(cFileTypes)

		Return lSuccess
	Endfunc


	* Search a file
	* -- Pass an empty cPattern to only collect definitions
	Function FileSearch(cFilename, cPattern)
		Local nSelect
		Local cFileFind
		Local cFolderFind
		Local nSelect
		Local lDefinitions
		Local lSearch
		Local cFileID
		Local cFileAction
		Local oEngine
		Local lSuccess
		Local nIndex
		Local cLowerFilename
		Local Array aFileList[1]

		If This.lCancel
			Return .F.
		Endif

		m.nSelect = Select()
		m.lSuccess = .T.

		With This.oSearchEngine
			.Filename = cFilename
			.SearchFor()
		Endwith
		Select (m.nSelect)

		Return m.lSuccess
	Endfunc

	****************************************************************
	* Save preferences to FoxPro Resource file
	Function SavePrefs()
		Local nSelect
		Local lSuccess
		Local nMemoWidth
		Local nCnt
		Local cData
		Local i
		Local oOptionCollection
		Local Array aFileList[1]
		Local Array FOXREF_OPTIONS[1]
		Local Array FOXREF_LOOKFOR_MRU[10]
		Local Array FOXREF_FOLDER_MRU[10]
		Local Array FOXREF_FILETYPES_MRU[10]

		=Acopy(This.aLookForMRU, FOXREF_LOOKFOR_MRU)
		=Acopy(This.aReplaceMRU, FOXREF_REPLACE_MRU)
		=Acopy(This.aFolderMRU, FOXREF_FOLDER_MRU)
		=Acopy(This.aFileTypesMRU, FOXREF_FILETYPES_MRU)

		oOptionCollection = Createobject("Collection")
		* Add any properties you want to save to
		* the resource file to this collection
		With oOptionCollection
			.Add(This.Comments, 			"Comments")
			.Add(This.MatchCase, 			"MatchCase")
			.Add(This.WholeWordsOnly, 		"WholeWordsOnly")
			.Add(This.Wildcards, 			"Wildcards")
			.Add(This.ProjectHomeDir, 		"ProjectHomeDir")
			.Add(This.SubFolders, 			"SubFolders")
			.Add(This.OverwritePrior, 		"OverwritePrior")
			.Add(This.FileTypes,			"FileTypes")
			.Add(This.IncludeDefTable, 		"IncludeDefTable")
			.Add(This.CodeOnly, 			"CodeOnly")
			.Add(This.FormProperties, 		"FormProperties")
			.Add(This.AutoProjectHomeDir, 	"AutoProjectHomeDir")
			.Add(This.ConfirmReplace, 		"ConfirmReplace")
			.Add(This.BackupOnReplace, 		"BackupOnReplace")
			.Add(This.DisplayReplaceLog, 	"DisplayReplaceLog")
			.Add(This.PreserveCase, 		"PreserveCase")
			.Add(This.BackupStyle, 			"BackupStyle")
			.Add(This.ShowRefsPerLine, 		"ShowRefsPerLine")
			.Add(This.ShowFileTypeHistory, 	"ShowFileTypeHistory")
			.Add(This.ShowDistinctMethodLine, "ShowDistinctMethodLine")
			.Add(This.SortMostRecentFirst, 	"SortMostRecentFirst")
			.Add(This.FontString, 			"FontString")
			.Add(This.FoxRefDirectory, 		"FoxRefDirectory")

			.Add(This.FileNameExactMatch, 	"FileNameExactMatch")
			.Add(This.ClassSearch, 			"ClassSearch")
			.Add(This.ClassExactMatch, 		"ClassExactMatch")
			.Add(This.SearchBaseClass, 		"SearchBaseClass")
			.Add(This.HomeDir, 				"HomeDir")
		Endwith

		Dimension FOXREF_OPTIONS[MAX(oOptionCollection.Count, 1), 2]
		For i = 1 To oOptionCollection.Count
			FOXREF_OPTIONS[m.i, 1] = oOptionCollection.GetKey(m.i)
			FOXREF_OPTIONS[m.i, 2] = oOptionCollection.Item(m.i)
		Endfor


		nSelect = Select()

		lSuccess = .F.

		* make sure Resource file exists and is not read-only
		Try
			*** JRN 03/24/2010 : Use This.cResourceFile instead of sys(2005)
			nCnt = Adir(aFileList, This.cResourceFile)
		Catch
			nCnt = 0
		Endtry

		If nCnt > 0 And Atcc('R', aFileList[1, 5]) == 0
			If !Used("FoxResource")
				Use (This.cResourceFile) In 0 Shared Again Alias FoxResource
			Endif
			If Used("FoxResource") And !Isreadonly("FoxResource")
				nMemoWidth = Set('MEMOWIDTH')
				Set Memowidth To 255

				Select FoxResource
				Locate For Upper(Alltrim(Type)) == "PREFW" And Upper(Alltrim(Id)) == RESOURCE_ID And Empty(Name)
				If !Found()
					Append Blank In FoxResource
					Replace ;
						Type With "PREFW", ;
						ID With RESOURCE_ID, ;
						ReadOnly With .F. ;
						IN FoxResource
				Endif

				If !FoxResource.ReadOnly
					Save To Memo Data All Like FOXREF_*

					Replace ;
						Updated With Date(), ;
						ckval With Val(Sys(2007, FoxResource.Data)) ;
						IN FoxResource

					lSuccess = .T.
				Endif
				Set Memowidth To (nMemoWidth)

				Use In FoxResource
			Endif
		Endif

		Select (nSelect)

		Return lSuccess
	Endfunc


	* retrieve preferences from the FoxPro Resource file
	Function RestorePrefs()
		Local nSelect
		Local lSuccess
		Local nMemoWidth
		Local i
		Local nCnt

		Local Array FOXREF_LOOKFOR_MRU[10]
		Local Array FOXREF_REPLACE_MRU[10]
		Local Array FOXREF_FOLDER_MRU[10]
		Local Array FOXREF_FILETYPES_MRU[10]
		Local Array FOXREF_OPTIONS[1]

		m.nSelect = Select()

		m.lSuccess = .F.

		*** JRN 03/24/2010 : Use This.cResourceFile instead of sys(2005)
		If File(This.cResourceFile)    && resource file not found.
			Use (This.cResourceFile) In 0 Shared Again Alias FoxResource
			If Used("FoxResource")
				m.nMemoWidth = Set('MEMOWIDTH')
				Set Memowidth To 255

				Select FoxResource
				Locate For Upper(Alltrim(Type)) == "PREFW" And ;
					UPPER(Alltrim(Id)) == RESOURCE_ID And ;
					EMPTY(Name) And ;
					!Deleted()

				If Found() And !Empty(Data) And ckval == Val(Sys(2007, Data))
					Restore From Memo Data Additive

					If Type("FOXREF_LOOKFOR_MRU") == 'C'
						=Acopy(FOXREF_LOOKFOR_MRU, This.aLookForMRU)
					Endif
					If Type("FOXREF_REPLACE_MRU") == 'C'
						=Acopy(FOXREF_REPLACE_MRU, This.aReplaceMRU)
					Endif
					If Type("FOXREF_FOLDER_MRU") == 'C'
						=Acopy(FOXREF_FOLDER_MRU, This.aFolderMRU)
					Endif
					If Type("FOXREF_FILETYPES_MRU") == 'C'
						=Acopy(FOXREF_FILETYPES_MRU, This.aFileTypesMRU)
					Endif

					If Type("FOXREF_OPTIONS") == 'C'
						m.nCnt = Alen(FOXREF_OPTIONS, 1)
						For m.i = 1 To m.nCnt
							If Vartype(FOXREF_OPTIONS[m.i, 1]) == 'C' And Pemstatus(This, FOXREF_OPTIONS[m.i, 1], 5)
								Store FOXREF_OPTIONS[m.i, 2] To ("THIS." + FOXREF_OPTIONS[m.i, 1])
							Endif
						Endfor
					Endif

					m.lSuccess = .T.
				Endif

				Set Memowidth To (m.nMemoWidth)

				Use In FoxResource
			Endif
		Endif

		Select (m.nSelect)

		Return m.lSuccess
	Endfunc


	*!*		****************************************************************
	*!*		****************************************************************

	*!*		Function FoxRefDirectory_Assign(cFoxRefDirectory)
	*!*			m.cFoxRefDirectory = Addbs(m.cFoxRefDirectory)
	*!*			If Empty(m.cFoxRefDirectory) Or !Directory(m.cFoxRefDirectory)
	*!*				m.cFoxRefDirectory = Addbs(Home(7))
	*!*				If !Directory(m.cFoxRefDirectory)
	*!*					m.cFoxRefDirectory = Addbs(Home())
	*!*				Endif
	*!*			Endif
	*!*			This.FoxRefDirectory = m.cFoxRefDirectory

	*!*			This.DefTable   = This.FoxRefDirectory + DEF_TABLE
	*!*			This.FileTable  = This.FoxRefDirectory + FILE_TABLE
	*!*			This.AddInTable = This.FoxRefDirectory + ADDIN_TABLE

	*!*			This.InitAddIns()
	*!*			This.OpenTables()
	*!*		Endfunc

	*!*		* Open the Add In table which contains our filetypes to process
	*!*		Function InitAddIns(lExclusive)
	*!*			Local nSelect
	*!*			Local nMemoWidth
	*!*			Local cAddInType

	*!*			m.nSelect = Select()

	*!*			* if we don't find the AddIn table on disk, then copy out
	*!*			* our project version of it
	*!*			If !File(Forceext(This.AddInTable, "DBF"))
	*!*				Try
	*!*					Use FoxRefAddin In 0 Shared Again
	*!*					Select FoxRefAddin
	*!*					Copy To (This.AddInTable) With Production
	*!*				Catch
	*!*				Finally
	*!*					If Used("FoxRefAddin")
	*!*						Use In FoxRefAddin
	*!*					Endif
	*!*				Endtry
	*!*			Endif

	*!*			Try
	*!*				If m.lExclusive
	*!*					Use (This.AddInTable) Alias AddInCursor In 0 Exclusive
	*!*				Else
	*!*					Use (This.AddInTable) Alias AddInCursor In 0 Shared Again
	*!*				Endif
	*!*			Catch
	*!*			Endtry

	*!*			If !Used("AddInCursor")
	*!*				* we didn't find the Add-In table on disk, so use our built-in version
	*!*				Try
	*!*					Use FoxRefAddin Alias AddInCursor In 0 Shared Again
	*!*				Catch
	*!*				Endtry
	*!*			Endif

	*!*			* process AddIn table
	*!*			This.oEngineCollection = Createobject("Collection")
	*!*			This.oReportCollection = Createobject("Collection")
	*!*			This.oFileTypeCollection = Createobject("Collection") && default filetypes from FoxRefAddin

	*!*			This.AddFileType("*.*", FILETYPE_CLASS_DEFAULT, FILETYPE_LIBRARY_DEFAULT)

	*!*			If Used("AddInCursor")
	*!*				m.nMemoWidth = Set("MEMOWIDTH")
	*!*				Set Memowidth To 1000

	*!*				* setup file types
	*!*				Select AddInCursor
	*!*				Scan All
	*!*					m.cAddInType = Rtrim(AddInCursor.Type)

	*!*					Do Case
	*!*						Case m.cAddInType == ADDINTYPE_FINDFILE
	*!*							This.AddFileType(AddInCursor.Data, AddInCursor.ClassName, AddInCursor.Classlib)

	*!*						Case m.cAddInType == ADDINTYPE_IGNOREFILE
	*!*							This.AddFileType(AddInCursor.Data)

	*!*						Case m.cAddInType == ADDINTYPE_FINDWINDOW
	*!*							If !Empty(AddInCursor.ClassName)
	*!*								This.FindWindowClass    = AddInCursor.ClassName
	*!*								This.FindWindowClassLib = AddInCursor.Classlib
	*!*							Endif

	*!*						Case m.cAddInType == ADDINTYPE_MATCH
	*!*							If !Empty(AddInCursor.ClassName)
	*!*								This.MatchClass    = AddInCursor.ClassName
	*!*								This.MatchClassLib = AddInCursor.Classlib
	*!*							Endif

	*!*						Case m.cAddInType == ADDINTYPE_WILDMATCH
	*!*							If !Empty(AddInCursor.ClassName)
	*!*								This.WildMatchClass    = AddInCursor.ClassName
	*!*								This.WildMatchClassLib = AddInCursor.Classlib
	*!*							Endif

	*!*						Case m.cAddInType == ADDINTYPE_REPORT
	*!*							This.AddReport(AddInCursor.Data, AddInCursor.Classlib, AddInCursor.ClassName, AddInCursor.Method, AddInCursor.Filename)

	*!*						Case m.cAddInType == ADDINTYPE_FILETYPE
	*!*							This.oFileTypeCollection.Add(AddInCursor.Data)
	*!*					Endcase
	*!*				Endscan
	*!*				Set Memowidth To (m.nMemoWidth)

	*!*			Endif
	*!*			Select (m.nSelect)
	*!*		Endfunc

	*!*		Procedure FontString_Access
	*!*			If Empty(This.FontString)
	*!*				Return FONT_DEFAULT
	*!*			Else
	*!*				Return This.FontString
	*!*			Endif
	*!*		Endproc


	*!*		* Add another file to process definitions for
	*!*		* These are files that aren't in our normal
	*!*		* scope processing, but are discovered along the
	*!*		* way, such as #include
	*!*		Function AddFileToProcess(cFilename)
	*!*			If File(m.cFilename)
	*!*				This.oFileCollection.AddNoDupe(Lower(Fullpath(m.cFilename)))
	*!*			Endif
	*!*		Endfunc

	*!*		* Add a file to search -- these are files that aren't
	*!*		* in our normal scope processing, but we search anyhow,
	*!*		* such as a Table within a DBC when searching a project
	*!*		Function AddFileToSearch(cFilename)
	*!*			If File(m.cFilename)
	*!*				This.oSearchCollection.AddNoDupe(Lower(Fullpath(m.cFilename)))
	*!*			Endif
	*!*		Endfunc


	*!*		* Add a report from the Add-Ins
	*!*		Function AddReport(cReportName, cClassLib, cClassName, cMethod, cFilename)
	*!*			Local oReportAddIn

	*!*			oReportAddIn = Newobject("ReportAddIn")
	*!*			oReportAddIn.ReportName      = m.cReportName
	*!*			oReportAddIn.RptClassLibrary = m.cClassLib
	*!*			oReportAddIn.RptClassName    = m.cClassName
	*!*			oReportAddIn.RptMethod       = m.cMethod
	*!*			oReportAddIn.RptFilename     = m.cFilename

	*!*			Try
	*!*				This.oReportCollection.Add(oReportAddIn, m.cReportName)
	*!*			Catch
	*!*			Endtry
	*!*		Endfunc

	*!*		* Add a filetype search
	*!*		Function AddFileType(cFileSkeleton, cClassName, cClassLibrary)
	*!*			Local nIndex
	*!*			Local lSuccess
	*!*			Local oEngine

	*!*			If Vartype(m.cClassName) <> 'C'
	*!*				m.cClassName = ''
	*!*			Endif

	*!*			m.cFileSkeleton = Upper(Alltrim(m.cFileSkeleton))
	*!*			If Empty(m.cFileSkeleton)
	*!*				m.cFileSkeleton = "*.*"
	*!*			Endif

	*!*			For m.i = 1 To This.oEngineCollection.Count
	*!*				If This.oEngineCollection.GetKey(m.i) == m.cFileSkeleton
	*!*					This.oEngineCollection.Remove(m.i)
	*!*					Exit
	*!*				Endif
	*!*			Endfor

	*!*			If Empty(m.cClassName)
	*!*				This.oEngineCollection.Add(.Null., m.cFileSkeleton)
	*!*			Else
	*!*				Try
	*!*					m.oEngine = Newobject(m.cClassName, m.cClassLibrary)
	*!*				Catch
	*!*					m.oEngine = .Null.
	*!*				Finally
	*!*				Endtry
	*!*				If !Isnull(oEngine)
	*!*					If m.cFileSkeleton = "*.*" And This.oEngineCollection.Count > 0
	*!*						This.oEngineCollection.Add(m.oEngine, m.cFileSkeleton, 1)
	*!*					Else
	*!*						This.oEngineCollection.Add(m.oEngine, m.cFileSkeleton)
	*!*					Endif
	*!*				Endif
	*!*			Endif

	*!*			Return m.lSuccess
	*!*		Endfunc





	*!*		* Methods for updating Errors collection
	*!*		Procedure ClearErrors()
	*!*			This.Errors.Remove(-1)
	*!*		Endproc

	*!*		Procedure AddError(cErrorMsg)
	*!*			This.Errors.Add(cErrorMsg)
	*!*		Endproc

	*!*		Procedure AddLog(cLog)
	*!*			If Pcount() == 0
	*!*				If !Empty(This.ActivityLog)
	*!*					This.ActivityLog = This.ActivityLog + Chr(13) + Chr(10)
	*!*				Endif
	*!*			Else
	*!*				If !Empty(m.cLog)
	*!*					This.ActivityLog = This.ActivityLog + m.cLog + Chr(13) + Chr(10)
	*!*				Endif
	*!*			Endif
	*!*		Endproc


	*!*		* Save replacement log
	*!*		Function SaveLog(cReplaceText)
	*!*			If Vartype(m.cReplaceText) <> 'C'
	*!*				m.cReplaceText = ''
	*!*			Endif

	*!*			This.ReplaceLogSetID = Sys(2015)
	*!*			Insert Into FoxRefCursor ( ;
	*!*				UniqueID, ;
	*!*				SetID, ;
	*!*				RefID, ;
	*!*				RefType, ;
	*!*				FileID, ;
	*!*				Symbol, ;
	*!*				ClassName, ;
	*!*				ProcName, ;
	*!*				ProcLineNo, ;
	*!*				LineNo, ;
	*!*				ColPos, ;
	*!*				MatchLen, ;
	*!*				Abstract, ;
	*!*				RecordID, ;
	*!*				UpdField, ;
	*!*				Checked, ;
	*!*				NoReplace, ;
	*!*				Timestamp, ;
	*!*				Inactive ;
	*!*				) Values ( ;
	*!*				SYS(2015), ;
	*!*				THIS.ReplaceLogSetID, ;
	*!*				SYS(2015), ;
	*!*				REFTYPE_LOG, ;
	*!*				'', ;
	*!*				M.cReplaceText, ;
	*!*				'', ;
	*!*				'', ;
	*!*				0, ;
	*!*				0, ;
	*!*				0, ;
	*!*				0, ;
	*!*				THIS.ActivityLog, ;
	*!*				'', ;
	*!*				'', ;
	*!*				.F., ;
	*!*				.F., ;
	*!*				DATETIME(), ;
	*!*				.F. ;
	*!*				)
	*!*		Endfunc




	*!*		* Make sure the Filetypes list is delimited with spaces
	*!*		Procedure FileTypes_Assign(cFileTypes)
	*!*			*** JRN 2010-03-14 : Don't remove spaces
	*!*			This.FileTypes = cFileTypes
	*!*			*	THIS.FileTypes = CHRTRAN(cFileTypes, ',;', '  ')
	*!*		Endfunc


	*!*		* ---
	*!*		* --- Definition Table methods
	*!*		* ---
	*!*		Function CreateFileTable()
	*!*			Local lSuccess
	*!*			Local cSafety

	*!*			m.lSuccess = .T.

	*!*			m.cSafety = Set("SAFETY")
	*!*			Set Safety Off
	*!*			If Used(Juststem(This.FileTable))
	*!*				Use In (This.FileTable)
	*!*			Endif


	*!*			Try
	*!*				Create Table (This.FileTable) Free ( ;
	*!*					UniqueID C(10), ;
	*!*					Folder M, ;
	*!*					Filename C(100), ;
	*!*					FileAction C(1), ;
	*!*					Timestamp T Null ;
	*!*					)
	*!*			Catch
	*!*				m.lSuccess = .F.
	*!*				Messagebox(ERROR_CREATEFILETABLE_LOC + Chr(10) + Chr(10) + Forceext(This.FileTable, "DBF"), MB_ICONSTOP, APPNAME_LOC)
	*!*			Endtry

	*!*			If m.lSuccess
	*!*				* insert a record that represents open windows
	*!*				Insert Into (This.FileTable) ( ;
	*!*					UniqueID, ;
	*!*					Filename, ;
	*!*					Folder, ;
	*!*					Timestamp, ;
	*!*					FileAction ;
	*!*					) Values ( ;
	*!*					"WINDOW", ;
	*!*					OPENWINDOW_LOC, ;
	*!*					'', ;
	*!*					DATETIME(), ;
	*!*					FILEACTION_NODEFINITIONS ;
	*!*					)

	*!*				Index On UniqueID Tag UniqueID
	*!*				Index On Filename Tag Filename
	*!*				Index On FileAction Tag FileAction

	*!*				Use In (Juststem(This.FileTable))
	*!*			Endif

	*!*			Set Safety &cSafety

	*!*			Return m.lSuccess
	*!*		Endfunc

	*!*		Function CreateDefTable()
	*!*			Local lSuccess
	*!*			Local cSafety

	*!*			m.lSuccess = .T.

	*!*			m.cSafety = Set("SAFETY")
	*!*			Set Safety Off
	*!*			If Used(Juststem(This.DefTable))
	*!*				Use In (This.DefTable)
	*!*			Endif


	*!*			Try
	*!*				Create Table (This.DefTable) Free ( ;
	*!*					UniqueID C(10), ;
	*!*					DefType C(1), ;
	*!*					FileID C(10), ;
	*!*					Symbol M, ;
	*!*					ClassName M, ;
	*!*					ProcName M, ;
	*!*					ProcLineNo i, ;
	*!*					LineNo i, ;
	*!*					Abstract M, ;
	*!*					Inactive L ;
	*!*					)

	*!*			Catch
	*!*				m.lSuccess = .F.
	*!*				Messagebox(ERROR_CREATEDEFTABLE_LOC + Chr(10) + Chr(10) + Forceext(This.DefTable, "DBF"), MB_ICONSTOP, APPNAME_LOC)
	*!*			Endtry

	*!*			If m.lSuccess
	*!*				Index On UniqueID Tag UniqueID
	*!*				Index On DefType Tag DefType
	*!*				Index On Inactive Tag Inactive
	*!*				Index On FileID Tag FileID

	*!*				Use In (Juststem(This.DefTable))
	*!*			Endif

	*!*			Set Safety &cSafety

	*!*			Return m.lSuccess
	*!*		Endfunc


	*!*		* Open a File & Definition tables
	*!*		Function OpenTables(lExclusive, lQuiet)
	*!*			Local lSuccess

	*!*			This.CloseTables()

	*!*			m.lSuccess = .T.

	*!*			If !File(Forceext(This.FileTable, "DBF"))
	*!*				m.lSuccess = This.CreateFileTable()
	*!*			Endif
	*!*			If m.lSuccess And !File(Forceext(This.DefTable, "DBF"))
	*!*				m.lSuccess = This.CreateDefTable()
	*!*			Endif

	*!*			If m.lSuccess
	*!*				* Open the table of files processed
	*!*				Try
	*!*					If m.lExclusive
	*!*						Use (This.FileTable) Alias FileCursor In 0 Exclusive
	*!*					Else
	*!*						Use (This.FileTable) Alias FileCursor In 0 Shared Again
	*!*					Endif
	*!*				Catch
	*!*				Endtry

	*!*				* Open the table of Definitions
	*!*				Try
	*!*					If m.lExclusive
	*!*						Use (This.DefTable) Alias FoxDefCursor In 0 Exclusive
	*!*					Else
	*!*						Use (This.DefTable) Alias FoxDefCursor In 0 Shared Again
	*!*					Endif
	*!*				Catch
	*!*				Endtry



	*!*				If Used("FileCursor")
	*!*					If Type("FileCursor.UniqueID") <> 'C' Or Type("FileCursor.Filename") <> 'C'
	*!*						m.lSuccess = .F.
	*!*						If !m.lQuiet
	*!*							Messagebox(ERROR_BADFILETABLE_LOC + Chr(10) + Chr(10) + Forceext(This.FileTable, "DBF"), MB_ICONSTOP, APPNAME_LOC)
	*!*						Endif
	*!*					Endif
	*!*				Else
	*!*					If !m.lQuiet
	*!*						Messagebox(ERROR_OPENFILETABLE_LOC + Chr(10) + Chr(10) + Forceext(This.FileTable, "DBF"), MB_ICONSTOP, APPNAME_LOC)
	*!*					Endif
	*!*					m.lSuccess = .F.
	*!*				Endif

	*!*				If m.lSuccess
	*!*					If Used("FoxDefCursor")
	*!*						If Type("FoxDefCursor.DefType") <> 'C'
	*!*							m.lSuccess = .F.
	*!*							If !m.lQuiet
	*!*								Messagebox(ERROR_BADDEFTABLE_LOC + Chr(10) + Chr(10) + Forceext(This.DefTable, "DBF"), MB_ICONSTOP, APPNAME_LOC)
	*!*							Endif
	*!*						Endif
	*!*					Else
	*!*						If !m.lQuiet
	*!*							Messagebox(ERROR_OPENDEFTABLE_LOC + Chr(10) + Chr(10) + Forceext(This.DefTable, "DBF"), MB_ICONSTOP, APPNAME_LOC)
	*!*						Endif
	*!*						m.lSuccess = .F.
	*!*					Endif
	*!*				Endif

	*!*			Endif

	*!*			Return m.lSuccess
	*!*		Endfunc

	*!*		*!*		Function GetAvailableOptions()
	*!*		*!*			Local oOptions
	*!*		*!*			Local nSelect
	*!*		*!*			Local oException
	*!*		*!*			Local oRefOption

	*!*		*!*			nSelect = Select()

	*!*		*!*			oOptions = .Null.
	*!*		*!*			oException = .Null.
	*!*		*!*			Try
	*!*		*!*				Select * From ;
	*!*		*!*					FoxRefOption ;
	*!*		*!*					ORDER By DisplayOrd ;
	*!*		*!*					INTO Cursor FoxRefOptionCursor
	*!*		*!*			Catch To oException
	*!*		*!*			Endtry

	*!*		*!*			If Isnull(oException)
	*!*		*!*				oOptions = Newobject("Collection")
	*!*		*!*				Select FoxRefOptionCursor
	*!*		*!*				Scan All
	*!*		*!*					oRefOption = Newobject("RefOption", "foxrefengine.prg")
	*!*		*!*					With oRefOption
	*!*		*!*						.OptionName   = Rtrim(FoxRefOptionCursor.OptionName)
	*!*		*!*						.Description  = FoxRefOptionCursor.Descrip
	*!*		*!*						.PropertyName = Rtrim(FoxRefOptionCursor.OptionProp)
	*!*		*!*						Try
	*!*		*!*							.OptionValue  = Eval("THIS." + .PropertyName)
	*!*		*!*						Catch
	*!*		*!*						Endtry
	*!*		*!*					Endwith
	*!*		*!*					oOptions.Add(oRefOption, oRefOption.PropertyName)
	*!*		*!*				Endscan
	*!*		*!*			Else
	*!*		*!*				Messagebox(m.oException.Message, MB_ICONSTOP, APPNAME_LOC)
	*!*		*!*			Endif

	*!*		*!*			If Used("FoxRefOption")
	*!*		*!*				Use In FoxRefOption
	*!*		*!*			Endif
	*!*		*!*			If Used("FoxRefOptionCursor")
	*!*		*!*				Use In FoxRefOptionCursor
	*!*		*!*			Endif

	*!*		*!*			Select (nSelect)

	*!*		*!*			Return oOptions
	*!*		*!*		Endfunc


	*!*		Function SaveOptions(oOptionsCollection As Collection)
	*!*			Local oOptions
	*!*			Local oException
	*!*			Local oRefOption


	*!*			For Each oRefOption In oOptionsCollection
	*!*				If !Isnull(oRefOption.OptionValue)
	*!*					Store (oRefOption.OptionValue) To ("THIS." + oRefOption.PropertyName)
	*!*				Endif
	*!*			Endfor
	*!*		Endfunc


	*!*		Function CloseTables()
	*!*			If Used("FoxDefCursor")
	*!*				Use In FoxDefCursor
	*!*			Endif
	*!*			If Used("FileCursor")
	*!*				Use In FileCursor
	*!*			Endif
	*!*			If Used("AddInCursor")
	*!*				Use In AddInCursor
	*!*			Endif
	*!*		Endfunc

	*!*		* Get the name of the first valid Reference table
	*!*		* We check the structure here of existing files
	*!*		* before we overwrite or open a reference table.
	*!*		* Increment a file count as necessary, so instead
	*!*		* of opening "giftrap_ref.dbf", we might end up
	*!*		* opening "giftrap_ref3.dbf" (if giftrap_ref, giftrap_ref1, and
	*!*		* giftrap_ref2 are not of the correct type)
	*!*		Function GetRefTableName(cRefTable)
	*!*			Local lTableIsOkay
	*!*			Local nFileCnt
	*!*			Local cNewRefTable
	*!*			Local oException
	*!*			Local cFilename

	*!*			m.cFilename = Chrtran(Justfname(m.cRefTable), INVALID_ALIAS_CHARS, Replicate('_', Lenc(INVALID_ALIAS_CHARS)))
	*!*			m.cRefTable = Addbs(Justpath(m.cRefTable)) + m.cFilename

	*!*			m.cNewRefTable = Forceext(m.cRefTable, "DBF")

	*!*			m.nFileCnt = 0
	*!*			m.lTableIsOkay = .F.
	*!*			Do While !lTableIsOkay And m.nFileCnt < 100
	*!*				Try
	*!*					If File(m.cNewRefTable)
	*!*						Use (m.cNewRefTable) Alias CheckRefTable In 0 Shared Again
	*!*						m.lTableIsOkay = (Type("CheckRefTable.RefType") == 'C')
	*!*						Use In CheckRefTable
	*!*					Else
	*!*						m.lTableIsOkay = .T.
	*!*					Endif
	*!*				Catch
	*!*				Endtry

	*!*				If !m.lTableIsOkay
	*!*					m.nFileCnt = m.nFileCnt + 1
	*!*					m.cNewRefTable = Addbs(Justpath(m.cRefTable)) + Juststem(m.cRefTable) + Transform(m.nFileCnt) + ".DBF"
	*!*				Endif
	*!*			Enddo
	*!*			If m.nFileCnt > 0
	*!*				m.cRefTable = Addbs(Justpath(m.cRefTable)) + Juststem(m.cRefTable) + Transform(m.nFileCnt) + ".DBF"
	*!*			Endif

	*!*			Return m.cRefTable
	*!*		Endfunc

	*!*		* ---
	*!*		* --- Reference Table methods
	*!*		* ---


	*!*		Function CreateRefTable(cRefTable)
	*!*			Local lSuccess
	*!*			Local cSafety
	*!*			Local oException

	*!*			m.lSuccess = .T.

	*!*			This.RefTable = ''

	*!*			m.cRefTable = This.GetRefTableName(m.cRefTable)

	*!*			m.cSafety = Set("SAFETY")
	*!*			Set Safety Off

	*!*			If Used(Juststem(m.cRefTable))
	*!*				Use In (Juststem(m.cRefTable))
	*!*			Endif

	*!*			Try
	*!*				Create Table (m.cRefTable) Free ( ;
	*!*					UniqueID C(10), ;
	*!*					SetID C(10), ;
	*!*					RefID C(10), ;
	*!*					RefType C(1), ;
	*!*					FindType C(1), ;
	*!*					FileID C(10), ;
	*!*					Symbol M, ;
	*!*					ClassName M, ;
	*!*					ProcName M, ;
	*!*					ProcLineNo i, ;
	*!*					LineNo i, ;
	*!*					ColPos i, ;
	*!*					MatchLen i, ;
	*!*					Abstract M, ;
	*!*					RecordID C(10), ;
	*!*					UpdField C(15), ;
	*!*					Checked L Null, ;
	*!*					NoReplace L, ;
	*!*					TimeStamp T Null, ;
	*!*					oTimeStamp N(10), ;
	*!*					Inactive L ;
	*!*					)
	*!*			Catch To oException
	*!*				m.lSuccess = .F.
	*!*				Messagebox(ERROR_CREATEREFTABLE_LOC + " (" + m.oException.Message + "):" + Chr(10) + Chr(10) + Forceext(m.cRefTable, "DBF"), MB_ICONSTOP, APPNAME_LOC)
	*!*			Endtry

	*!*			If m.lSuccess
	*!*				Index On RefType Tag RefType
	*!*				Index On SetID Tag SetID
	*!*				Index On RefID Tag RefID
	*!*				Index On UniqueID Tag UniqueID
	*!*				Index On FileID Tag FileID
	*!*				Index On Checked Tag Checked
	*!*				Index On Inactive Tag Inactive


	*!*				* add the record that holds our results window search position & other options
	*!*				Insert Into (m.cRefTable) ( ;
	*!*					UniqueID, ;
	*!*					SetID, ;
	*!*					RefType, ;
	*!*					FindType, ;
	*!*					FileID, ;
	*!*					Symbol, ;
	*!*					ClassName, ;
	*!*					ProcName, ;
	*!*					ProcLineNo, ;
	*!*					LineNo, ;
	*!*					ColPos, ;
	*!*					MatchLen, ;
	*!*					Abstract, ;
	*!*					RecordID, ;
	*!*					UpdField, ;
	*!*					Timestamp, ;
	*!*					Checked, ;
	*!*					NoReplace, ;
	*!*					Inactive ;
	*!*					) Values ( ;
	*!*					SYS(2015), ;
	*!*					'', ;
	*!*					REFTYPE_INIT, ;
	*!*					'', ;
	*!*					'', ;
	*!*					THIS.ProjectFile, ;
	*!*					'', ;
	*!*					'', ;
	*!*					0, ;
	*!*					0, ;
	*!*					0, ;
	*!*					0, ;
	*!*					'', ;
	*!*					'', ;
	*!*					'', ;
	*!*					DATETIME(), ;
	*!*					.F., ;
	*!*					.F., ;
	*!*					.F. ;
	*!*					)

	*!*				This.RefTable = m.cRefTable

	*!*				Use
	*!*			Endif

	*!*			Set Safety &cSafety

	*!*			Return m.lSuccess
	*!*		Endfunc

	*!*		* Open a FoxRef table
	*!*		* Return TRUE if table exists and it's in the correct format
	*!*		* [lCreate]    = True to create table if it doesn't exist
	*!*		* [lExclusive] = True to open for exclusive use
	*!*		Function OpenRefTable(cRefTable, lExclusive)
	*!*			Local lSuccess
	*!*			Local oException

	*!*			If Used("FoxRefCursor")
	*!*				Use In FoxRefCursor
	*!*			Endif
	*!*			This.RefTable = ''

	*!*			m.lSuccess = .T.

	*!*			m.cRefTable = This.GetRefTableName(m.cRefTable)

	*!*			If !File(Forceext(m.cRefTable, "DBF"))
	*!*				m.lSuccess = This.CreateRefTable(m.cRefTable)
	*!*			Endif

	*!*			If m.lSuccess
	*!*				Try
	*!*					If m.lExclusive
	*!*						Use (m.cRefTable) Alias FoxRefCursor In 0 Exclusive
	*!*					Else
	*!*						Use (m.cRefTable) Alias FoxRefCursor In 0 Shared Again
	*!*					Endif
	*!*				Catch To oException
	*!*					Messagebox(ERROR_OPENREFTABLE_LOC + " (" + m.oException.Message + "):" + Chr(10) + Chr(10) + Forceext(m.cRefTable, "DBF"), MB_ICONSTOP, APPNAME_LOC)
	*!*					m.lSuccess = .F.
	*!*				Endtry

	*!*				If m.lSuccess
	*!*					If Type("FoxRefCursor.RefType") == 'C'
	*!*						This.RefTable = m.cRefTable
	*!*						*** JRN 2010-03-20 : Add field oTimeStamp
	*!*						If Type("FoxRefCursor.oTimeStamp") # 'N'
	*!*							Local lnSelect
	*!*							lnSelect = Select()

	*!*							Select FoxRefCursor
	*!*							Use (cRefTable) Exclusive Alias FoxRefCursor
	*!*							Alter Table FoxRefCursor Add Column oTimeStamp N(10)

	*!*							Use (cRefTable) Alias FoxRefCursor Shared Again
	*!*							Select(lnSelect)
	*!*						Endif
	*!*						***
	*!*					Else
	*!*						m.lSuccess = .F.
	*!*						Messagebox(ERROR_BADREFTABLE_LOC + Chr(10) + Chr(10) + Forceext(m.cRefTable, "DBF"), MB_ICONSTOP, APPNAME_LOC)
	*!*					Endif
	*!*				Endif
	*!*			Endif

	*!*			Return m.lSuccess
	*!*		Endfunc

	*!*		* return number of search sets in the current FoxRef table
	*!*		Function SearchCount()
	*!*			Local nSelect
	*!*			Local nSearchCnt
	*!*			Local Array aSearchCnt[1]

	*!*			m.nSelect = Select()
	*!*			Select Cnt(*) ;
	*!*				FROM (This.RefTable) ;
	*!*				WHERE ;
	*!*				RefType == REFTYPE_SEARCH And !Inactive ;
	*!*				INTO Array aSearchCnt
	*!*			If _Tally > 0
	*!*				m.nSearchCnt = aSearchCnt[1]
	*!*			Else
	*!*				m.nSearchCnt = 0
	*!*			Endif

	*!*			Select (m.nSelect)

	*!*			Return m.nSearchCnt
	*!*		Endfunc

	*!*		* Returns TRUE if there is only 1 record in the FoxRef table,
	*!*		* indicating that no searches have been done yet
	*!*		* (the first record is initialization information)
	*!*		Function FirstSearch()
	*!*			Return Used("FoxRefCursor") And Reccount("FoxRefCursor") <= 1
	*!*		Endfunc

	*!*		* Add to cursor of available project files.
	*!*		* This cursor is all files in current project,
	*!*		* plus any files we encounter along the way
	*!*		* that are #include or SET PROCEDURE TO, SET CLASSLIB TO
	*!*		Function AddFileToProjectCursor(cFilename)
	*!*			Local nSelect
	*!*			Local i
	*!*			Local nCnt
	*!*			Local Array aFileList[1]

	*!*			m.nSelect = Select()

	*!*			m.cFilename = Lower(Fullpath(m.cFilename))
	*!*			m.cFolder   = Justpath(m.cFilename)

	*!*			m.cFilename = Padr(Justfname(m.cFilename), 100)

	*!*			Select ProjectFilesCursor
	*!*			Locate For Filename == m.cFilename And Folder == m.cFolder
	*!*			If !Found()
	*!*				Insert Into ProjectFilesCursor ( ;
	*!*					Filename, ;
	*!*					Folder ;
	*!*					) Values ( ;
	*!*					M.cFilename, ;
	*!*					M.cFolder ;
	*!*					)
	*!*			Endif

	*!*			Select (m.nSelect)
	*!*		Endfunc

	*!*		* Grabs all Include files for this project
	*!*		* and adds them to our project list cursor
	*!*		Function UpdateProjectFiles()
	*!*			Local nSelect
	*!*			Local i
	*!*			Local nCnt
	*!*			Local oFileRef
	*!*			Local Array aFileList[1]

	*!*			m.nSelect = Select()

	*!*			If Used("ProjectFilesCursor")
	*!*				Use In ProjectFilesCursor
	*!*			Endif
	*!*			Create Cursor ProjectFilesCursor ( ;
	*!*				Folder M, ;
	*!*				Filename C(100) ;
	*!*				)
	*!*			Index On Filename Tag Filename

	*!*			* Add in all files that are in currently in the project
	*!*			If Vartype(This.oProjectFileRef) == 'O'
	*!*				For Each m.oFileRef In This.oProjectFileRef.Files
	*!*					This.AddFileToProjectCursor(oFileRef.Name)
	*!*				Endfor
	*!*			Endif

	*!*			* add in all dependencies
	*!*			Select ;
	*!*				DISTINCT Padr(Leftc(Symbol, 254), 254) As IncludeFile ;
	*!*				FROM (This.DefTable) DefTable ;
	*!*				WHERE ;
	*!*				DefTable.DefType == DEFTYPE_INCLUDEFILE And ;
	*!*				!DefTable.Inactive ;
	*!*				INTO Array aFileList
	*!*			m.nCnt = _Tally
	*!*			For m.i = 1 To m.nCnt
	*!*				Try
	*!*					If File(Rtrim(aFileList[m.i, 1]))  && make sure we can find the file along our path somewhere
	*!*						This.AddFileToProjectCursor(aFileList[m.i, 1])
	*!*					Endif
	*!*				Catch
	*!*				Endtry
	*!*			Endfor

	*!*			Select (m.nSelect)
	*!*		Endfunc

	*!*		* return project files as a collection
	*!*		Function GetProjectFiles()
	*!*			Local nSelect
	*!*			Local oProjectFiles

	*!*			nSelect = Select()

	*!*			oProjectFiles = Createobject("Collection")

	*!*			If Used("ProjectFilesCursor")
	*!*				Select ProjectFilesCursor
	*!*				Scan All
	*!*					oProjectFiles.Add(Addbs(Rtrim(ProjectFilesCursor.Folder)) + Rtrim(ProjectFilesCursor.Filename))
	*!*				Endscan
	*!*			Endif
	*!*			Select (nSelect)

	*!*			Return oProjectFiles
	*!*		Endfunc

	*!*		* collect all definitions for a project/folder and currently open window
	*!*		* without doing an actual search
	*!*		* [lLocalOnly] = search open window for LOCALS & PARAMETERS
	*!*		Function CollectDefinitions(lLocalOnly)
	*!*			Local nSelect
	*!*			Local cOpenFile
	*!*			Local nCnt
	*!*			Local i
	*!*			Local lSuccess
	*!*			Local lDefinitionsOnly
	*!*			Local lOverwritePrior
	*!*			Local lCodeOnly
	*!*			Local cFileTypes
	*!*			Local lProjectHomeDir
	*!*			Local oException


	*!*			If !This.SearchInit()
	*!*				Return .F.
	*!*			Endif

	*!*			lDefinitionsOnly  = This.lDefinitionsOnly
	*!*			lOverwritePrior   = This.OverwritePrior
	*!*			lCodeOnly         = This.CodeOnly
	*!*			cFileTypes        = This.FileTypes
	*!*			lProjectHomeDir   = This.ProjectHomeDir

	*!*			This.lDefinitionsOnly = .T.
	*!*			This.OverwritePrior   = .F.
	*!*			This.CodeOnly         = .F.
	*!*			This.FileTypes        = FILETYPES_DEFINITIONS
	*!*			This.ProjectHomeDir   = This.AutoProjectHomeDir

	*!*			m.nSelect = Select()

	*!*			* make sure we have the latest definitions for the open window
	*!*			Try
	*!*				If m.lLocalOnly
	*!*					Update FoxDefCursor ;
	*!*						SET Inactive = .T. ;
	*!*						WHERE FileID = "WINDOW"
	*!*					If This.WindowHandle >= 0 And Vartype(This.oWindowEngine) == 'O'
	*!*						This.oWindowEngine.WindowHandle = This.WindowHandle
	*!*						This.oWindowEngine.ProcessDefinitions(This)

	*!*						Update FileCursor Set ;
	*!*							Filename = Justfname(This.WindowFilename), ;
	*!*							Folder = Justpath(This.WindowFilename), ;
	*!*							FileAction = FILEACTION_DEFINITIONS ;
	*!*							WHERE UniqueID = "WINDOW"
	*!*					Endif
	*!*				Else
	*!*					If This.SetProject(.Null., .F.)
	*!*						If This.ProjectFile == PROJECT_GLOBAL Or Empty(This.ProjectFile)
	*!*							m.lSuccess = This.FolderSearch('')
	*!*						Else
	*!*							m.lSuccess = This.ProjectSearch('')
	*!*						Endif
	*!*					Endif

	*!*					* Process definitions for files that
	*!*					* were #included
	*!*					m.nCnt = This.oFileCollection.Count
	*!*					For m.i = 1 To m.nCnt
	*!*						m.lSuccess = This.FileSearch(This.oFileCollection.Item(m.i), '')
	*!*						If !m.lSuccess
	*!*							Exit
	*!*						Endif
	*!*					Endfor
	*!*					This.oFileCollection.Remove(-1)
	*!*					This.oProcessedCollection.Remove(-1)

	*!*					This.UpdateProjectFiles()
	*!*				Endif
	*!*			Catch To oException
	*!*				Messagebox(oException.Message)
	*!*			Endtry

	*!*			This.lDefinitionsOnly   = lDefinitionsOnly
	*!*			This.OverwritePrior     = lOverwritePrior
	*!*			This.CodeOnly           = lCodeOnly
	*!*			This.FileTypes          = cFileTypes
	*!*			This.ProjectHomeDir     = lProjectHomeDir

	*!*			Select (m.nSelect)
	*!*		Endfunc



	*!*		* Determine if the Reference table we want to open is
	*!*		* actually one of ours.  If we're overwriting or a reference
	*!*		* table doesn't exist for this project, then create a new
	*!*		* Reference Table.
	*!*		*
	*!*		* Once we have a reference table, then we add a new record
	*!*		* that represents the search criteria for this particular
	*!*		* search.
	*!*		Function UpdateRefTable(cScope, cPattern, cProjectOrDir)
	*!*			Local nSelect
	*!*			Local cSafety
	*!*			Local cSearchOptions
	*!*			Local cRefTable
	*!*			Local i

	*!*			m.nSelect = Select()

	*!*			If Vartype(cRefTable) <> 'C' Or Empty(cRefTable)
	*!*				cRefTable = This.RefTable
	*!*			Endif

	*!*			If Empty(cRefTable)
	*!*				Return .F.
	*!*			Endif

	*!*			If Used("FoxRefCursor")
	*!*				Use In FoxRefCursor
	*!*			Endif

	*!*			If !This.OpenRefTable(cRefTable)
	*!*				Return .F.
	*!*			Endif


	*!*			This.tTimeStamp = Datetime()


	*!*			* Since we're only doing definitions and there
	*!*			* is no search in progress, then don't create
	*!*			* a Search Set record in the FoxRef table
	*!*			If This.lDefinitionsOnly
	*!*				Return .T.
	*!*			Endif

	*!*			* build a string representing the search options that
	*!*			* we can store to the FoxRef cursor
	*!*			cSearchOptions = Iif(This.Comments == COMMENTS_EXCLUDE, 'X', '') + ;
	*!*				IIF(This.Comments == COMMENTS_ONLY, 'C', '') + ;
	*!*				IIF(This.MatchCase, 'M', '') + ;
	*!*				IIF(This.WholeWordsOnly, 'W', '') + ;
	*!*				IIF(This.ProjectHomeDir, 'H', '') + ;
	*!*				IIF(This.FormProperties, 'P', '') + ;
	*!*				IIF(This.SubFolders, 'S', '') + ;
	*!*				IIF(This.Wildcards, 'Z', '') + ;
	*!*				';' + Alltrim(This.FileTypes)


	*!*			* if we've already searched for this same exact symbol
	*!*			* with the same exact criteria in the same exact project/folder,
	*!*			* then simply update what we have
	*!*			If This.lRefreshMode
	*!*				Select FoxRefCursor
	*!*				Locate For SetID == This.cSetID And RefType == REFTYPE_SEARCH And !Inactive
	*!*				This.lRefreshMode = Found()
	*!*			Endif

	*!*			If !This.lRefreshMode
	*!*				Select FoxRefCursor
	*!*				Locate For RefType == REFTYPE_SEARCH And ClassName == cProjectOrDir And Symbol == cPattern And Abstract == cSearchOptions And !Inactive
	*!*				This.lRefreshMode = Found()
	*!*			Endif

	*!*			If This.lRefreshMode
	*!*				This.tTimeStamp = FoxRefCursor.Timestamp

	*!*				This.cSetID = FoxRefCursor.SetID
	*!*				Update FoxRefCursor ;
	*!*					SET Inactive = .T. ;
	*!*					WHERE ;
	*!*					SetID == This.cSetID And ;
	*!*					(RefType == REFTYPE_RESULT Or RefType == REFTYPE_ERROR Or RefType == REFTYPE_NOMATCH)
	*!*			Else
	*!*				This.cSetID = Sys(2015)

	*!*				* add the record that specifies the search criteria, etc
	*!*				Insert Into FoxRefCursor ( ;
	*!*					UniqueID, ;
	*!*					SetID, ;
	*!*					RefType, ;
	*!*					FindType, ;
	*!*					Symbol, ;
	*!*					ClassName, ;
	*!*					ProcName, ;
	*!*					ProcLineNo, ;
	*!*					LineNo, ;
	*!*					ColPos, ;
	*!*					MatchLen, ;
	*!*					Abstract, ;
	*!*					RecordID, ;
	*!*					UpdField, ;
	*!*					Checked, ;
	*!*					NoReplace, ;
	*!*					Timestamp, ;
	*!*					Inactive ;
	*!*					) Values ( ;
	*!*					SYS(2015), ;
	*!*					THIS.cSetID, ;
	*!*					REFTYPE_SEARCH, ;
	*!*					'', ;
	*!*					cPattern, ;
	*!*					cProjectOrDir, ;
	*!*					'', ;
	*!*					0, ;
	*!*					0, ;
	*!*					0, ;
	*!*					0, ;
	*!*					cSearchOptions, ;
	*!*					'', ;
	*!*					'', ;
	*!*					.F., ;
	*!*					.F., ;
	*!*					DATETIME(), ;
	*!*					.F. ;
	*!*					)
	*!*			Endif

	*!*			* update each of the search engines with the new SetID
	*!*			If Vartype(This.oWindowEngine) == 'O'
	*!*				This.oWindowEngine.SetID = This.cSetID
	*!*			Endif
	*!*			For m.i = 1 To This.oEngineCollection.Count
	*!*				If Vartype(This.oEngineCollection.Item(m.i)) == 'O'
	*!*					With This.oEngineCollection.Item(m.i)
	*!*						.SetID = This.cSetID
	*!*					Endwith
	*!*				Endif
	*!*			Endfor

	*!*			Select Distinct SetID, FileID, Timestamp ;
	*!*				FROM FoxRefCursor ;
	*!*				WHERE (RefType == REFTYPE_RESULT Or RefType == REFTYPE_NOMATCH) And !Inactive ;
	*!*				INTO Cursor FoxRefSearchedCursor


	*!*			Select (m.nSelect)

	*!*			Return .T.
	*!*		Endfunc



	*!*		* Return Search Engine to use based upon filetype
	*!*		Function GetEngine(m.cFilename)
	*!*			Local nEngineIndex
	*!*			Local i

	*!*			If This.oEngineCollection.Count == 0
	*!*				Return .Null.
	*!*			Endif

	*!*			m.cFilename = Upper(m.cFilename)

	*!*			* determine which search engine to use based upon the filetype
	*!*			m.nEngineIndex = 1
	*!*			For m.i = 2 To This.oEngineCollection.Count
	*!*				If Vartype(This.oEngineCollection.Item(m.i)) == 'O'
	*!*					If This.WildcardMatch(This.oEngineCollection.GetKey(m.i), m.cFilename)
	*!*						m.nEngineIndex = m.i
	*!*						Exit
	*!*					Endif
	*!*				Endif
	*!*			Endfor

	*!*			* if we're still using the default search engine,
	*!*			* then make sure this file type isn't set to
	*!*			* be excluded
	*!*			If m.nEngineIndex == 1
	*!*				For m.i = 2 To This.oEngineCollection.Count
	*!*					If Vartype(This.oEngineCollection.Item(m.i)) <> 'O'
	*!*						If This.WildcardMatch(This.oEngineCollection.GetKey(m.i), m.cFilename)
	*!*							m.nEngineIndex = m.i
	*!*							Exit
	*!*						Endif
	*!*					Endif
	*!*				Endfor
	*!*			Endif

	*!*			Return This.oEngineCollection.Item(m.nEngineIndex)
	*!*		Endfunc


	*!*		* Search a file
	*!*		* -- Pass an empty cPattern to only collect definitions
	*!*		Function FileSearch(cFilename, cPattern)
	*!*			Local nSelect
	*!*			Local cFileFind
	*!*			Local cFolderFind
	*!*			Local nSelect
	*!*			Local lDefinitions
	*!*			Local lSearch
	*!*			Local cFileID
	*!*			Local cFileAction
	*!*			Local oEngine
	*!*			Local lSuccess
	*!*			Local nIndex
	*!*			Local cLowerFilename
	*!*			Local Array aFileList[1]

	*!*			If This.lCancel
	*!*				Return .F.
	*!*			Endif

	*!*			m.nSelect = Select()
	*!*			m.lSuccess = .T.

	*!*			With This.oSearchEngine
	*!*				.Filename = cFilename
	*!*				.SearchFor()
	*!*			Endwith
	*!*			Select (m.nSelect)

	*!*			Return m.lSuccess
	*!*		Endfunc



	*!*		* refresh results for all Sets in the Ref table or a single set
	*!*		Function RefreshResults(cSetID)
	*!*			Local nSelect
	*!*			Local i
	*!*			Local nCnt
	*!*			Local lSuccess
	*!*			Local Array aRefList[1]

	*!*			m.nSelect = Select()

	*!*			m.lSuccess = .T.

	*!*			If Vartype(m.cSetID) == 'C' And !Empty(m.cSetID)
	*!*				This.RefreshResultSet(m.cSetID)
	*!*			Else
	*!*				If File(Forceext(This.RefTable, "dbf"))
	*!*					m.nSelect = Select()

	*!*					m.lSuccess = This.OpenRefTable(This.RefTable)
	*!*					If m.lSuccess
	*!*						Select SetID ;
	*!*							FROM FoxRefCursor ;
	*!*							WHERE RefType == REFTYPE_SEARCH And !Inactive ;
	*!*							INTO Array aRefList
	*!*						m.nCnt = _Tally

	*!*						For m.i = 1 To m.nCnt
	*!*							This.RefreshResultSet(aRefList[i])
	*!*							If This.lCancel
	*!*								Exit
	*!*							Endif
	*!*						Endfor
	*!*					Endif


	*!*					Select (m.nSelect)
	*!*				Endif
	*!*				This.cSetID = ''
	*!*			Endif

	*!*			Return m.lSuccess
	*!*		Endfunc

	*!*		* refresh an existing search set
	*!*		Function RefreshResultSet(cSetID)
	*!*			Local nSelect
	*!*			Local lSuccess
	*!*			Local cScope
	*!*			Local cFolder
	*!*			Local cProject
	*!*			Local cPattern
	*!*			Local cSearchOptions

	*!*			m.lSuccess = .F.

	*!*			If File(Forceext(This.RefTable, "dbf"))
	*!*				m.nSelect = Select()

	*!*				If !This.OpenRefTable(This.RefTable)
	*!*					Return .F.
	*!*				Endif

	*!*				This.cSetID = m.cSetID

	*!*				Select FoxRefCursor
	*!*				Locate For RefType == REFTYPE_SEARCH And SetID == m.cSetID And !Inactive
	*!*				lSuccess = Found()
	*!*				If lSuccess
	*!*					cSearchOptions = Leftc(FoxRefCursor.Abstract, At_c(';', FoxRefCursor.Abstract) - 1)
	*!*					If 'X'$cSearchOptions
	*!*						This.Comments = COMMENTS_EXCLUDE
	*!*					Endif
	*!*					If 'C'$cSearchOptions
	*!*						This.Comments = COMMENTS_ONLY
	*!*					Endif
	*!*					This.MatchCase      = 'M' $ cSearchOptions
	*!*					This.WholeWordsOnly = 'W' $ cSearchOptions
	*!*					This.FormProperties = 'P' $ cSearchOptions
	*!*					This.ProjectHomeDir = 'H' $ cSearchOptions
	*!*					This.SubFolders     = 'S' $ cSearchOptions
	*!*					This.Wildcards      = 'Z' $ cSearchOptions

	*!*					This.OverwritePrior = .F.

	*!*					This.FileTypes = Alltrim(Substrc(FoxRefCursor.Abstract, At_c(';', FoxRefCursor.Abstract) + 1))

	*!*					cFolder  = Rtrim(FoxRefCursor.ClassName)
	*!*					cProject = ''

	*!*					If Upper(Justext(cFolder)) == "PJX"
	*!*						cScope = SCOPE_PROJECT
	*!*						cProject = cFolder
	*!*					Else
	*!*						cScope = SCOPE_FOLDER
	*!*					Endif

	*!*					cPattern = FoxRefCursor.Symbol
	*!*				Endif


	*!*				If lSuccess
	*!*					Do Case
	*!*						Case cScope == SCOPE_FOLDER
	*!*							lSuccess = This.FolderSearch(cPattern, cFolder)
	*!*						Case cScope == SCOPE_PROJECT
	*!*							lSuccess = This.ProjectSearch(cPattern, cProject)
	*!*						Otherwise
	*!*							lSuccess = .F.
	*!*					Endcase
	*!*				Endif

	*!*				Select (m.nSelect)
	*!*			Endif

	*!*			Return m.lSuccess
	*!*		Endfunc

	*!*		Function SetChecked(cUniqueID, lChecked)
	*!*			If Pcount() < 2
	*!*				lChecked = .T.
	*!*			Endif
	*!*			If Used("FoxRefCursor") And Seek(cUniqueID, "FoxRefCursor", "UniqueID")
	*!*				If !Isnull(FoxRefCursor.Checked)
	*!*					Replace Checked With lChecked In FoxRefCursor
	*!*				Endif
	*!*			Endif
	*!*		Endfunc




	*!*		* given a filename, return the FileID from RefFile
	*!*		Function GetFileID(cFilename)
	*!*			Local cFName
	*!*			Local cFolder
	*!*			Local nSelect
	*!*			Local cFileID
	*!*			Local Array aFileList[1]

	*!*			m.nSelect = Select()

	*!*			m.cFilename = Lower(m.cFilename)
	*!*			cFName  = Padr(Lower(Justfname(m.cFilename)), 100)
	*!*			* cFolder = PADR(LOWER(JUSTPATH(m.cFilename)), 240)
	*!*			cFolder = Lower(Justpath(m.cFilename))

	*!*			Select UniqueID ;
	*!*				FROM (This.FileTable) FileTable ;
	*!*				WHERE ;
	*!*				FileTable.Filename == cFName And ;
	*!*				FileTable.Folder == cFolder And ;
	*!*				FileTable.FileAction <> FILEACTION_INACTIVE ;
	*!*				INTO Array aFileList
	*!*			If _Tally > 0
	*!*				m.cFileID = aFileList[1]
	*!*			Else
	*!*				m.cFileID = ''
	*!*			Endif

	*!*			Select (m.nSelect)

	*!*			Return m.cFileID
	*!*		Endfunc


	*!*		* Show a progress form while searching
	*!*		Function ProgressInit(cDescription, nMax)
	*!*			If This.ShowProgress
	*!*				If Vartype(This.oProgressForm) <> 'O'
	*!*					This.UpdateProgress()
	*!*				Endif
	*!*				If Vartype(This.oProgressForm) == 'O'
	*!*					This.oProgressForm.SetMax(m.nMax)
	*!*					This.oProgressForm.SetDescription(m.cDescription)
	*!*				Endif
	*!*				DoEvents
	*!*			Endif
	*!*		Endfunc

	*!*		Function UpdateProgress(nValue, cMsg, lFilename)
	*!*			If This.ShowProgress
	*!*				If Vartype(This.oProgressForm) <> 'O'
	*!*					This.lCancel = .F.
	*!*					*	This.oProgressForm = Newobject("CProgressForm", "FoxRef.vcx")
	*!*					This.oProgressForm.Show()
	*!*				Endif

	*!*				If m.lFilename And !Empty(Justpath(m.cMsg))
	*!*					* truncate filenames so they fit
	*!*					m.cMsg = Displaypath(m.cMsg, 60)
	*!*				Endif

	*!*				If !This.oProgressForm.SetProgress(m.nValue, m.cMsg)  && FALSE is returned if Cancel button is pressed
	*!*					If Messagebox(SEARCH_CANCEL_LOC, MB_ICONQUESTION + MB_YESNO, APPNAME_LOC) == IDYES
	*!*						This.lCancel = .T.
	*!*					Else
	*!*						This.oProgressForm.lCancel = .F.
	*!*					Endif
	*!*				Endif
	*!*				DoEvents
	*!*			Endif
	*!*		Endfunc

	*!*		Function CloseProgress()
	*!*			If Vartype(This.oProgressForm) == 'O'
	*!*				This.oProgressForm.Release()
	*!*			Endif
	*!*		Endfunc



	*!*		* Clear a specified Result Set or all results
	*!*		* from the Reference table
	*!*		Function ClearResults(cSetID, cFileID)
	*!*			Local nSelect
	*!*			Local cAlias

	*!*			m.nSelect = Select()

	*!*			Do Case
	*!*				Case Vartype(m.cFileID) == 'C' And !Empty(m.cFileID)
	*!*					* Clear specified file
	*!*					Delete From (This.RefTable) Where SetID == m.cSetID And FileID == m.cFileID

	*!*				Case Vartype(m.cSetID) == 'C' And !Empty(m.cSetID)
	*!*					* Clear specified result set
	*!*					Delete From (This.RefTable) Where SetID == m.cSetID

	*!*				Otherwise
	*!*					* Clear all results
	*!*					Delete From (This.RefTable) ;
	*!*						WHERE ;
	*!*						RefType == REFTYPE_RESULT Or ;
	*!*						RefType == REFTYPE_SEARCH Or ;
	*!*						RefType == REFTYPE_NOMATCH Or ;
	*!*						RefType == REFTYPE_ERROR Or ;
	*!*						RefType == REFTYPE_LOG
	*!*			Endcase


	*!*			* if we can get this table open exclusive, then
	*!*			* we should pack it
	*!*			If This.OpenRefTable(This.RefTable, .T.)
	*!*				Select FoxRefCursor
	*!*				Try
	*!*					Pack In FoxRefCursor
	*!*				Catch
	*!*					* no big deal that we can't pack the table -- just ignore the error
	*!*				Endtry
	*!*			Endif


	*!*			* open it again shared
	*!*			This.OpenRefTable(This.RefTable)


	*!*			Select (m.nSelect)

	*!*			Return .T.
	*!*		Endfunc



	*!*		Function ProjectMatch(cFileTypes, oMatchFileCollection, cFilename)
	*!*			Local i
	*!*			Local lMatch

	*!*			lMatch = .F.
	*!*			For i = 1 To oMatchFileCollection.Count
	*!*				If Empty(Justpath(oMatchFileCollection.Item(i))) Or Upper(Justpath(oMatchFileCollection.Item(i))) == Upper(Justpath(cFilename))
	*!*					If This.FileMatch(Justfname(cFilename), Justfname(oMatchFileCollection.Item(i)))
	*!*						lMatch = .T.
	*!*						Exit
	*!*					Endif
	*!*				Endif
	*!*			Endfor
	*!*			Return lMatch
	*!*		Endfunc

	*!*		* borrowed from Class Browser
	*!*		Function WildcardMatch(tcMatchExpList, tcExpressionSearched, tlMatchAsIs)
	*!*			Local lcMatchExpList,lcExpressionSearched,llMatchAsIs,lcMatchExpList2
	*!*			Local lnMatchLen,lnExpressionLen,lnMatchCount,lnCount,lnCount2,lnSpaceCount
	*!*			Local lcMatchExp,lcMatchType,lnMatchType,lnAtPos,lnAtPos2
	*!*			Local llMatch,llMatch2

	*!*			If Alltrim(tcMatchExpList) == "*.*"
	*!*				Return .T.
	*!*			Endif

	*!*			If Empty(tcExpressionSearched)
	*!*				If Empty(tcMatchExpList) Or Alltrim(tcMatchExpList) == "*"
	*!*					Return .T.
	*!*				Endif
	*!*				Return .F.
	*!*			Endif
	*!*			lcMatchExpList=Lower(Alltrim(Strtran(tcMatchExpList,Tab," ")))
	*!*			lcExpressionSearched=Lower(Alltrim(Strtran(tcExpressionSearched,Tab," ")))
	*!*			lnExpressionLen=Lenc(lcExpressionSearched)
	*!*			If lcExpressionSearched==lcMatchExpList
	*!*				Return .T.
	*!*			Endif
	*!*			llMatchAsIs=tlMatchAsIs
	*!*			If Leftc(lcMatchExpList,1)==["] And Rightc(lcMatchExpList,1)==["]
	*!*				llMatchAsIs=.T.
	*!*				lcMatchExpList=Alltrim(Substrc(lcMatchExpList,2,Lenc(lcMatchExpList)-2))
	*!*			Endif
	*!*			If Not llMatchAsIs And " "$lcMatchExpList
	*!*				llMatch=.F.
	*!*				lnSpaceCount=Occurs(" ",lcMatchExpList)
	*!*				lcMatchExpList2=lcMatchExpList
	*!*				lnCount=0
	*!*				Do While .T.
	*!*					lnAtPos=At_c(" ",lcMatchExpList2)
	*!*					If lnAtPos=0
	*!*						lcMatchExp=Alltrim(lcMatchExpList2)
	*!*						lcMatchExpList2=""
	*!*					Else
	*!*						lnAtPos2=At_c(["],lcMatchExpList2)
	*!*						If lnAtPos2<lnAtPos
	*!*							lnAtPos2=At_c(["],lcMatchExpList2,2)
	*!*							If lnAtPos2>lnAtPos
	*!*								lnAtPos=lnAtPos2
	*!*							Endif
	*!*						Endif
	*!*						lcMatchExp=Alltrim(Leftc(lcMatchExpList2,lnAtPos))
	*!*						lcMatchExpList2=Alltrim(Substrc(lcMatchExpList2,lnAtPos+1))
	*!*					Endif
	*!*					If Empty(lcMatchExp)
	*!*						Exit
	*!*					Endif
	*!*					lcMatchType=Leftc(lcMatchExp,1)
	*!*					Do Case
	*!*						Case lcMatchType=="+"
	*!*							lnMatchType=1
	*!*						Case lcMatchType=="-"
	*!*							lnMatchType=-1
	*!*						Otherwise
	*!*							lnMatchType=0
	*!*					Endcase
	*!*					If lnMatchType#0
	*!*						lcMatchExp=Alltrim(Substrc(lcMatchExp,2))
	*!*					Endif
	*!*					llMatch2=This.WildcardMatch(lcMatchExp,lcExpressionSearched, .T.)
	*!*					If (lnMatchType=1 And Not llMatch2) Or (lnMatchType=-1 And llMatch2)
	*!*						Return .F.
	*!*					Endif
	*!*					llMatch=(llMatch Or llMatch2)
	*!*					If lnAtPos=0
	*!*						Exit
	*!*					Endif
	*!*				Enddo
	*!*				Return llMatch
	*!*			Else
	*!*				If Leftc(lcMatchExpList,1)=="~"
	*!*					Return (Difference(Alltrim(Substrc(lcMatchExpList,2)),lcExpressionSearched)>=3)
	*!*				Endif
	*!*			Endif
	*!*			lnMatchCount=Occurs(",",lcMatchExpList)+1
	*!*			If lnMatchCount>1
	*!*				lcMatchExpList=","+Alltrim(lcMatchExpList)+","
	*!*			Endif
	*!*			For lnCount = 1 To lnMatchCount
	*!*				If lnMatchCount=1
	*!*					lcMatchExp=Lower(Alltrim(lcMatchExpList))
	*!*					lnMatchLen=Lenc(lcMatchExp)
	*!*				Else
	*!*					lnAtPos=At_c(",",lcMatchExpList,lnCount)
	*!*					lnMatchLen=At_c(",",lcMatchExpList,lnCount+1)-lnAtPos-1
	*!*					lcMatchExp=Lower(Alltrim(Substrc(lcMatchExpList,lnAtPos+1,lnMatchLen)))
	*!*				Endif
	*!*				For lnCount2 = 1 To Occurs("?",lcMatchExp)
	*!*					lnAtPos=At_c("?",lcMatchExp)
	*!*					If lnAtPos>lnExpressionLen
	*!*						If (lnAtPos-1)=lnExpressionLen
	*!*							lcExpressionSearched=lcExpressionSearched+"?"
	*!*						Endif
	*!*						Exit
	*!*					Endif
	*!*					lcMatchExp=Stuff(lcMatchExp,lnAtPos,1,Substrc(lcExpressionSearched,lnAtPos,1))
	*!*				Endfor
	*!*				If Empty(lcMatchExp) Or lcExpressionSearched==lcMatchExp Or ;
	*!*						lcMatchExp=="*" Or lcMatchExp=="?" Or lcMatchExp=="%%"
	*!*					Return .T.
	*!*				Endif
	*!*				If Leftc(lcMatchExp,1)=="*"
	*!*					Return (Substrc(lcMatchExp,2)==Rightc(lcExpressionSearched,Lenc(lcMatchExp)-1))
	*!*				Endif
	*!*				If Leftc(lcMatchExp,1)=="%" And Rightc(lcMatchExp,1)=="%" And ;
	*!*						SUBSTRC(lcMatchExp,2,lnMatchLen-2)$lcExpressionSearched
	*!*					Return .T.
	*!*				Endif
	*!*				lnAtPos=At_c("*",lcMatchExp)
	*!*				If lnAtPos>0 And (lnAtPos-1)<=lnExpressionLen And ;
	*!*						LEFTC(lcExpressionSearched,lnAtPos-1)==Leftc(lcMatchExp,lnAtPos-1)
	*!*					Return .T.
	*!*				Endif
	*!*			Endfor
	*!*			Return .F.
	*!*		Endfunc

	*!*		* used for comparing filenames against a wildcard
	*!*		* For folder searches we can use ADIR(), but for project
	*!*		* searches we need to use this function
	*!*		Function FileMatch(cText As String, cPattern As String)
	*!*			Local i, j, k, cPattern, nPatternLen, nTextLen, ch

	*!*			If Pcount() < 2
	*!*				cPattern = This.cPattern
	*!*			Endif

	*!*			nPatternLen = Lenc(cPattern)
	*!*			nTextLen = Lenc(cText)

	*!*			If nPatternLen == 0
	*!*				Return .T.
	*!*			Endif
	*!*			If nTextLen == 0
	*!*				Return .F.
	*!*			Endif

	*!*			j = 1
	*!*			For i = 1 To nPatternLen
	*!*				If j > Lenc(cText)
	*!*					Return .F.
	*!*				Else
	*!*					ch = Substrc(cPattern, i, 1)
	*!*					If ch == FILEMATCH_ANY
	*!*						j = j + 1
	*!*					Else
	*!*						If ch == FILEMATCH_MORE
	*!*							For k = j To nTextLen
	*!*								If This.FileMatch(Substrc(cText, k), Substrc(cPattern, i + 1))
	*!*									Return .T.
	*!*								Endif
	*!*							Endfor
	*!*							Return .F.
	*!*						Else
	*!*							If j <= nTextLen And ch <> Substrc(cText, j, 1)
	*!*								Return .F.
	*!*							Else
	*!*								j = j + 1
	*!*							Endif
	*!*						Endif
	*!*					Endif
	*!*				Endif
	*!*			Endfor

	*!*			* RETURN j > nTextLen
	*!*			Return .T.
	*!*		Endfunc


	*!*		* -- Parse out what's in Abstract field to return
	*!*		* -- either the search criteria or the file types searched
	*!*		Function ParseAbstract(cAbstract, cParseWhat)
	*!*			Local cDisplayText
	*!*			Local cSearchOptions

	*!*			cDisplayText = ''

	*!*			Do Case
	*!*				Case cParseWhat == "CRITERIA"
	*!*					cSearchOptions = Leftc(cAbstract, At_c(';', cAbstract) - 1)

	*!*					If 'X' $ cSearchOptions
	*!*						cDisplayText = cDisplayText + Iif(Empty(cDisplayText), '', ", ") + COMMENTS_EXCLUDE_LOC
	*!*					Endif
	*!*					If 'C' $ cSearchOptions
	*!*						cDisplayText = cDisplayText + Iif(Empty(cDisplayText), '', ", ") + COMMENTS_ONLY_LOC
	*!*					Endif
	*!*					If 'M' $ cSearchOptions
	*!*						cDisplayText = cDisplayText + Iif(Empty(cDisplayText), '', ", ") + CRITERIA_MATCHCASE_LOC
	*!*					Endif
	*!*					If 'W' $ cSearchOptions
	*!*						cDisplayText = cDisplayText + Iif(Empty(cDisplayText), '', ", ") + CRITERIA_WHOLEWORDS_LOC
	*!*					Endif
	*!*					If 'P' $ cSearchOptions
	*!*						cDisplayText = cDisplayText + Iif(Empty(cDisplayText), '', ", ") + CRITERIA_FORMPROPERTIES_LOC
	*!*					Endif
	*!*					If 'H' $ cSearchOptions And This.ProjectFile <> PROJECT_GLOBAL
	*!*						cDisplayText = cDisplayText + Iif(Empty(cDisplayText), '', ", ") + CRITERIA_PROJECTHOMEDIR_LOC
	*!*					Endif
	*!*					If 'S' $ cSearchOptions And This.ProjectFile == PROJECT_GLOBAL
	*!*						cDisplayText = cDisplayText + Iif(Empty(cDisplayText), '', ", ") + CRITERIA_SUBFOLDERS_LOC
	*!*					Endif
	*!*					If 'Z' $ cSearchOptions
	*!*						cDisplayText = cDisplayText + Iif(Empty(cDisplayText), '', ", ") + CRITERIA_WILDCARDS_LOC
	*!*					Endif

	*!*				Case cParseWhat == "FILETYPES"
	*!*					cDisplayText = Alltrim(Substrc(Rtrim(cAbstract), At_c(';', cAbstract) + 1))

	*!*			Endcase

	*!*			Return cDisplayText
	*!*		Endfunc


	*!*		* -- Cleanup tables used by FoxRef -- removing
	*!*		* -- any references to files we can't find and
	*!*		* -- packing the files
	*!*		Function CleanupTables()
	*!*			Local nSelect
	*!*			Local cFilename
	*!*			Local lSuccess

	*!*			m.nSelect = Select()

	*!*			m.lSuccess = .F.

	*!*			If Used(Juststem(This.RefTable))
	*!*				Use In (Juststem(This.RefTable))
	*!*			Endif
	*!*			If Used(Juststem(This.DefTable))
	*!*				Use In (Juststem(This.DefTable))
	*!*			Endif
	*!*			If Used(Juststem(This.FileTable))
	*!*				Use In (Juststem(This.FileTable))
	*!*			Endif

	*!*			If This.OpenTables(.T., .T.)
	*!*				Select FileCursor
	*!*				Delete All For (FileAction == FILEACTION_ERROR Or FileAction == FILEACTION_INACTIVE) And UniqueID <> "WINDOW"
	*!*				Scan All For !Empty(Filename)
	*!*					m.cFilename = Addbs(Rtrim(FileCursor.Folder)) + Rtrim(FileCursor.Filename)
	*!*					If !File(m.cFilename)
	*!*						Delete In FileCursor

	*!*						Select FoxDefCursor
	*!*						Delete All For FileID == FileCursor.UniqueID In FoxDefCursor
	*!*					Endif
	*!*				Endscan

	*!*				Select FoxDefCursor
	*!*				Delete All For Inactive In FoxDefCursor

	*!*				Select FileCursor
	*!*				Pack In FileCursor

	*!*				Select FoxDefCursor
	*!*				Pack In FoxDefCursor

	*!*				m.lSuccess = .T.
	*!*			Else
	*!*				Messagebox(ERROR_EXCLUSIVE_LOC, MB_ICONEXCLAMATION, APPNAME_LOC)
	*!*			Endif


	*!*			* re-open the tables shared
	*!*			This.OpenTables()

	*!*			Select (m.nSelect)

	*!*			Return m.lSuccess
	*!*		Endfunc

	*!*		Function PrintReport(cReportName, lPreview)
	*!*			Local cFilename
	*!*			Local cExt
	*!*			Local oReportAddIn

	*!*			Try
	*!*				oReportAddIn = This.oReportAddIn.Item(m.cReportName)
	*!*			Catch
	*!*				oReportAddIn = .Null.
	*!*			Endtry

	*!*			If !Isnull(oReportAddIn)
	*!*				With oReportAddIn
	*!*					Do Case
	*!*						Case !Empty(.RptFilename)
	*!*							m.cExt = Upper(Justext(.RptFilename))
	*!*							Do Case
	*!*								Case m.cExt == "PRG" Or m.cExt == "FXP"
	*!*									Do (.RptFilename) With m.lPreview, This

	*!*								Case m.cExt == "SCX"
	*!*									Do Form (.RptFilename) With m.lPreview, This

	*!*								Case m.cExt == "FRX"
	*!*									Report Form (.RptFilename)
	*!*							Endcase

	*!*						Case !Empty(This.RptClassName)
	*!*					Endcase
	*!*				Endwith
	*!*			Endif
	*!*		Endfunc

	*!*		* Abstract:
	*!*		*   This program will shell out to specified file,
	*!*		*	which can be a URL (e.g. http://www.microsoft.com),
	*!*		*	a filename, etc
	*!*		*
	*!*		* Parameters:
	*!*		*	<cFile>
	*!*		*	[cParameters]
	*!*		Function ShellTo(cFile, cParameters)
	*!*			Local cRun
	*!*			Local cSysDir
	*!*			Local nRetValue

	*!*			*-- GetDesktopWindow gives us a window handle to
	*!*			*-- pass to ShellExecute.
	*!*			Declare Integer GetDesktopWindow In user32.Dll
	*!*			Declare Integer GetSystemDirectory In kernel32.Dll ;
	*!*				STRING @cBuffer, ;
	*!*				INTEGER liSize

	*!*			Declare Integer ShellExecute In shell32.Dll ;
	*!*				INTEGER, ;
	*!*				STRING @cOperation, ;
	*!*				STRING @cFile, ;
	*!*				STRING @cParameters, ;
	*!*				STRING @cDirectory, ;
	*!*				INTEGER nShowCmd

	*!*			If Vartype(m.cParameters) <> 'C'
	*!*				m.cParameters = ''
	*!*			Endif

	*!*			m.cOperation = "open"
	*!*			m.nRetValue = ShellExecute(GetDesktopWindow(), @m.cOperation, @m.cFile, @m.cParameters, '', SW_SHOWNORMAL)
	*!*			If m.nRetValue = SE_ERR_NOASSOC && No association exists
	*!*				m.cSysDir = Space(260)  && MAX_PATH, the maximum path length

	*!*				*-- Get the system directory so that we know where Rundll32.exe resides.
	*!*				m.nRetValue = GetSystemDirectory(@m.cSysDir, Lenc(m.cSysDir))
	*!*				m.cSysDir = Substrc(m.cSysDir, 1, m.nRetValue)
	*!*				m.cRun = "RUNDLL32.EXE"
	*!*				cParameters = "shell32.dll,OpenAs_RunDLL "
	*!*				m.nRetValue = ShellExecute(GetDesktopWindow(), "open", m.cRun, m.cParameters + m.cFile, m.cSysDir, SW_SHOWNORMAL)
	*!*			Endif

	*!*			Return m.nRetValue
	*!*		Endfunc

	*!*		****************************************************************


Enddefine



*!*	Define Class ReportAddIn As Custom
*!*		Name = "ReportAddIn"

*!*		ReportName      = ''
*!*		RptFilename     = ''
*!*		RptClassLibrary = ''
*!*		RptClassName    = ''
*!*		RptMethod       = ''
*!*	Enddefine

*!*	Define Class RefOption As Custom
*!*		Name = "RefOption"

*!*		OptionName   = ''
*!*		Description  = ''
*!*		OptionValue  = .Null.
*!*		PropertyName = ''
*!*	Enddefine




*!*	Procedure _EdGetLPos
*!*	Procedure _EdSelect
*!*	Procedure _EdSToPos
*!*	Procedure _edgetenv
*!*	Procedure _wfindtitl
*!*	Procedure _wselect

