*** JRN 2010-07-24 : Borrowed from FoxTabs (VFPx)
*** to be able to obtain list of all open wHandles


* _________________________________________________________________
*
* FoxTabs Manager Class
*
* [TODO] Describe class
* _________________________________________________________________

#Include "PEME_FoxTabs.h"

* _________________________________________________________________
*
* FoxTab Manager class
*
Define Class FoxTabsManager As Custom

	* _____________________________________________________________
	*
	* Define class properties

	Add Object FoxTabs As Collection
	*!!!*		Add Object WindowsEvents As FoxTabsEventHandler

	* _____________________________________________________________
	*
	* Define class members

	Function Init()

		* Declare Win32 API functions
		Declare Integer FindWindowEx In Win32API Integer hWndParent, Integer hwndChildAfter, String lpszClass, String lpszWindow
		Declare Integer GetWindowInfo In Win32API Integer HWnd, String @ pwindowinfo
		Declare Integer GetWindowText In Win32API Integer HWnd, String @szText, Integer nLen
		Declare Integer GetAncestor In Win32API Integer HWnd, Integer gaFlags

	Endfunc

	Function SetBindings()

		*!!!*	* Bind to all WM_CREATE messages
		*!!!*	BindEvent(0, WM_CREATE, This.WindowsEvents, "WMEventHandler")

		*!!!*	* Setup event bindings to our Windows message event handler
		*!!!*	BindEvent(This.WindowsEvents, "WindowShowEvent", This, "WindowShow")
		*!!!*	BindEvent(This.WindowsEvents, "WindowSetTextEvent", This, "WindowSetText")
		*!!!*	BindEvent(This.WindowsEvents, "WindowSetFocusEvent", This, "WindowSetFocus")
		*!!!*	BindEvent(This.WindowsEvents, "WindowDestroyEvent", This, "WindowDestroy")

	Endfunc

	Function LoadWindows (HWnd As Integer)

		Local oException As Exception
		Local lnChildHWnd As Integer

		Try
			* Check if the window is a VFP IDE window
			If This.IsIDEWindow (HWnd)
				* Add a new FoxTab object to our collection
				This.NewFoxTab (HWnd)
			Endif

			* Check all child windows
			lnChildHWnd = 0
			Do While .T.

				* Look for a child window of the current hWnd
				lnChildHWnd = FindWindowEx (HWnd, lnChildHWnd, 0, 0)
				If Empty (lnChildHWnd)
					* No child exists
					Exit
				Endif

				* Call this method recursively until all child windows have been processed
				This.LoadWindows (lnChildHWnd)

			Enddo

		Catch To oException
			*!!!*		* Raise error event
			*!!!*		Raiseevent (This.Parent, 'LogError', oException, 'Exception caught while loading existing IDE windows.')

		Endtry

	Endfunc

	Function GetWindows
		This.LoadWindows(Application.HWnd)
		Return This.FoxTabs
	Endfunc

	Function NewFoxTab (HWnd As Integer)

		Local oException As Exception
		Local oFoxTab As FoxTab
		Local Array laEnv [25]
		Local lnResult
		Try
			* Check if a FoxTab already exists for this handle
			lnWHandle = Sys(2326, HWnd)
			If Not Empty (This.FoxTabs.GetKey (Transform (lnWHandle, '@0x')))
				Exit
			Endif

			* Create a new instance of our FoxTab class
			oFoxTab = Newobject ('FoxTab')

			* Set windows handle and window name properties
			oFoxTab.HWnd = HWnd
			oFoxTab.WindowName = This.getWindowTitle (HWnd)
			*** JRN 2010-07-24 :
			oFoxTab.nHWnd			= HWnd
			oFoxTab.nWHandle		= lnWHandle
			If 0 # _EdGetEnv (lnWHandle, @laEnv )
				oFoxTab.nWHandleType	= laEnv[25] && Edit session Env[25]
			Endif

			* Add the class to our collection
			This.FoxTabs.Add (oFoxTab, Transform (lnWHandle, '@0x'))

			* Raise the add FoxTab event
			*!!!*	Raiseevent (This, 'AddFoxTabEvent', oFoxTab)

			* Setup event bindings for this window
			*!!!*			This.WindowsEvents.SetBindings (HWnd)

		Catch To oException When oException.ErrorNo = 2071		&& user thrown
			* Throw to caller
			Throw oException

		Catch To oException
			* Set exception header details and throw to caller
			oException.UserValue = 'Unhandled exception caught while adding a new FoxTab object to the collection.'
			Throw oException

		Endtry

	Endfunc

	Function getWindowTitle (HWnd As Integer) As String

		Local lcBuffer As String

		* Get the title using the WIN32API function
		lcBuffer = Space(200)

		GetWindowText (HWnd, @lcBuffer, Len (lcBuffer))

		* Strip out nulls and trim
		lcBuffer = Alltrim (Chrtran (lcBuffer, Chr(0), ''))

		Return lcBuffer

	Endfunc

	Function IsIDEWindow (HWnd As Integer) As Boolean

		Local oException As Exception
		Local lbReturn As Boolean, lcWindowTitle As String, lnParentHWnd As Integer

		Local laToolbars[1]
		laToolbars = ''
		= Adockstate (laToolbars, 2)

		Try
			* Gather the windows title
			lcWindowTitle = Nvl (This.getWindowTitle (HWnd), '')

			* Check the window title exclusion conditions
			*JAL*				lbReturn = Application.hWnd # hWnd									;					&& Ignore the main FoxPro window
			*JAL*							And This.Parent.FoxTabsToolbar.hWnd # hWnd				;	&& Ignore foxtabs toolbar
			*JAL*							And Not Empty(lcWindowTitle)							;				&& Ignore windows with no titles
			*JAL*							And Not InList(Lower(lcWindowTitle), "compile", "command", "properties", "document view", "data session", "debugger", "watch", "locals", "trace", "call stack", "debug output", "parentclass browser") ;
			*JAL*							And Not InList(Lower(lcWindowTitle), "expression builder", "expression builder options") ;
			*JAL*							And Ascan(laToolbars,lcWindowTitle,-1,-1,1,7) = 0

			*!!!*	And This.Parent.FoxTabsToolbar.HWnd # HWnd		;	&& Ignore foxtabs toolbar

			*JAL* Now allow tracking of IDE windows except debugger
			lbReturn = Application.HWnd # HWnd														;	&& Ignore the main FoxPro window
			And Not Empty (lcWindowTitle)														;	&& Ignore windows with no titles
			And Not Inlist (Lower (lcWindowTitle), 'compile', 'debugger', 'watch', 'locals', 'trace', 'call stack', 'debug output', 'parentclass browser') ;
				And Not Inlist (Lower (lcWindowTitle), 'expression builder', 'expression builder options') ;
				And Ascan (laToolbars, lcWindowTitle, -1, -1, 1, 7) = 0
			*Wait window nowait Transform(lcWindowTitle)

			* Gather the hWnd of this windows parent
			lnParentHWnd = GetAncestor (HWnd, GA_PARENT)

			* Only allow children of _Screen
			lbReturn = lbReturn And lnParentHWnd = _Screen.HWnd

			* Check the border style
			lbReturn = lbReturn And This.HasBorder (HWnd)

		Catch To oException
			* Set details and throw to caller
			oException.UserValue = 'Unhandled exception caught checking if the hWnd is an IDE Window.'
			Throw oException

		Endtry

		Return lbReturn

	Endfunc

	Function HasBorder (HWnd As Integer) As Boolean

		Local lbReturn As Boolean, lcBuffer As String, lcStyle As String

		* Check the border style using the GetWindowInfo API
		lcBuffer = Space(60)

		GetWindowInfo (HWnd, @lcBuffer)

		* Parse WINDOWINFO struct
		lcStyle = CToBin (Substr (lcBuffer, (4 * 9) + 1, 4), '4rs')
		If Bitand (lcStyle, WS_BORDER) = WS_BORDER
			lbReturn = .T.
		Endif

		Return lbReturn

	Endfunc

	Function RefreshWindows (HWnd As Integer)

		Local oException As Exception
		Local lnChildHWnd As Integer

		Local i As Integer
		Local lnTabHWnd As Integer

		Try
			* Go through and refresh all the Windows
			For i = 1 To This.FoxTabs.Count
				* refresh all the FoxTabs titles
				* convert the hex back to integer first though
				lnTabHWnd	= Int (Val (This.FoxTabs.Item[i].HWnd))
				This.FoxTabs.Item[i].WindowName = This.getWindowTitle (lnTabHWnd)
			Endfor

			This.LoadWindows (HWnd)

		Catch To oException
			*!!!*				* Raise error event
			*!!!*				Raiseevent (This.Parent, 'LogError', oException, 'Exception caught while loading existing IDE windows.')

		Endtry

	Endfunc

	* _____________________________________________________________
	*
	* Define event handlers

	Function WindowShow (HWnd As Integer)

		Local oFoxTab As Object
		Local oException As Exception

		Try
			* Check if a valid hWnd was passed
			If Empty (HWnd)
				Exit
			Endif

			* Check if a FoxTab object exists for this hWnd in our collection
			If Not Empty (This.FoxTabs.GetKey (Transform (HWnd, '@0x')))
				* Check the window is still a child of _Screen.
				If _Screen.HWnd	# GetAncestor (HWnd, GA_PARENT)
					* Remove FoxTab from our collection
					This.WindowDestroy (HWnd)
				Endif
				* Exit the try...catch block
				Exit
			Endif

			* Check if the window is a VFP IDE window we are intersted in
			If Not This.IsIDEWindow (HWnd)
				* Release WM_SETTEXT event binding to this window
				*!!!*	UnBindEvents(hWnd, WM_SETTEXT)
				* Exit the try...catch block
				Exit
			Endif

			* Add a new FoxTab object to our collection
			This.NewFoxTab (HWnd)

		Catch To oException When oException.ErrorNo = 2071		&& user thrown
			* Throw to caller
			Throw oException

		Catch To oException
			* Set exception header details and throw to caller
			oException.UserValue = 'Unhandled exception caught while adding a new FoxTab object to the collection.'
			Throw oException

		Endtry

	Endfunc

	Function WindowSetText (HWnd As Integer)

		Local oFoxTab As Object
		Local oException As Exception
		Local cWindowCaption As String

		Try
			* Check if a valid hWnd was passed
			If Empty (HWnd)
				Exit
			Endif

			* Sometimes the WM_SETEXT event occurs before the WM_SHOWWINDOW (eg when VFP opens a window maximised)
			*	so we check if we need to add a FoxTab object here as well.
			This.WindowShow (HWnd)

			* Check if a FoxTab object exists for this hWnd in our collection
			If Empty (This.FoxTabs.GetKey (Transform (HWnd, '@0x')))
				* Exit the try...catch block
				Exit
			Endif

			* Check the window is still a child of _Screen.
			* 	I have noticed VFP will create a child of _Screen, then
			*	SetParent to a descendent of _Screen later.
			If _Screen.HWnd	# GetAncestor (HWnd, GA_PARENT)
				* Remove FoxTab from our collection
				This.WindowDestroy (HWnd)
				* Exit the try...catch block
				Exit
			Endif

			* Set the window name of the corresponding FoxTab object
			oFoxTab = This.FoxTabs.Item (Transform (HWnd, '@0x'))
			cWindowCaption = This.getWindowTitle (HWnd)
			If Not oFoxTab.WindowName == cWindowCaption
				oFoxTab.WindowName = cWindowCaption

				* Raise the OnChange event
				*!!!*	Raiseevent (This, 'OnChangeEvent', oFoxTab)
			Endif

			* Set the window name of the corresponding FoxTab object
			*oFoxTab = This.FoxTabs.Item(Transform(hWnd, "@0x"))
			*oFoxTab.WindowName = This.getWindowTitle(hWnd)

			* Raise the OnChange event
			*RaiseEvent(This, "OnChangeEvent", oFoxTab)

		Catch To oException When oException.ErrorNo = 2071		&& user thrown
			* Throw to caller
			Throw oException

		Catch To oException
			* Set exception header details and throw to caller
			oException.UserValue = 'Unhandled exception caught while adding a new FoxTab object to the collection.'
			Throw oException

		Endtry

	Endfunc

	Function WindowSetFocus (HWnd As Integer)

		Local oFoxTab As Object
		Local oException As Exception

		Try
			* Check the window is still a child of _Screen.
			* 	I have noticed VFP will create a child of _Screen, then
			*	SetParent to a descendent of _Screen later.
			If _Screen.HWnd	# GetAncestor (HWnd, GA_PARENT)
				* Remove FoxTab from our collection
				This.WindowDestroy (HWnd)
				* Exit the try...catch block
				Exit
			Endif

			* Obtain a reference to the FoxTab object to pass to event delegate
			oFoxTab = This.FoxTabs.Item (Transform (HWnd, '@0x'))

			* Update the window name
			oFoxTab.WindowName = This.getWindowTitle (HWnd)

			* Raise the got focus event
			*!!!*	Raiseevent (This, 'GotFocusEvent', oFoxTab)

		Catch To oException When oException.ErrorNo = 2071		&& user thrown
			* Throw to caller
			Throw oException

		Catch To oException
			* Set exception header details and throw to caller
			oException.UserValue = 'Unhandled exception caught while updating the window name of the active FoxTab.'
			Throw oException

		Endtry

	Endfunc

	Function WindowDestroy (HWnd As Integer)

		Local oFoxTab As Object
		Local oException As Exception

		Try
			* Check if a FoxTab object exists for this hWnd in our collection
			If Empty (This.FoxTabs.GetKey (Transform (HWnd, '@0x')))
				* Not interested
				Exit
			Endif

			* Obtain a reference to the FoxTab object to pass to event delegate
			oFoxTab = This.FoxTabs.Item (Transform (HWnd, '@0x'))

			* Raise the remove FoxTab event
			*!!!*	Raiseevent (This, 'RemoveFoxTabEvent', oFoxTab)

			* Release FoxTab reference so we can remove it from the collection
			oFoxTab = Null

			* Remove the FoxTab from the collection
			This.FoxTabs.Remove (Transform (HWnd, '@0x'))

			* Release event bindings to this window
			*!!!*	UnBindEvents(hWnd)

		Catch To oException When oException.ErrorNo = 2071		&& user thrown
			* Throw to caller
			Throw oException

		Catch To oException
			* Set exception header details and throw to caller
			oException.UserValue = 'Unhandled exception caught while removing the active FoxTab from our collection.'
			Throw oException

		Endtry

	Endfunc

	* Remove all bindings
	* Pause function when user wants to avoid any interaction while testing their app
	Function RemoveBindings()

		Do While Type ('This.FoxTabs[1]') = 'O'
			This.WindowDestroy (Val (This.FoxTabs[1].HWnd))
		Enddo

		*!!!*	UnBindEvents(0)

	Endfunc

	* _____________________________________________________________
	*
	* Define event delegates

	Function AddFoxTabEvent (oFoxTab As Object)
		* Event interface
	Endfunc

	Function RemoveFoxTabEvent (oFoxTab As Object)
		* Event interface
	Endfunc

	Function GotFocusEvent (oFoxTab As Object)
		* Event interface
	Endfunc

	Function OnChangeEvent (oFoxTab As Object)
		* Event interface
	Endfunc

Enddefine

* _________________________________________________________________
*
* FoxTab class
*
Define Class FoxTab As Custom

	* Define class properties
	HWnd 			= ''
	WindowName 		= ''
	ContentType		= ''
	AssociatedIcon	= ''

	*** JRN 2010-07-24 :
	nHWnd			= 0
	nWHandle		= 0
	nWHandleType	= 0 && Edit session Env[25]
	lMatched		= .F.

	* Define class members
	Function WindowName_Assign (lpcWindowName As String)

		* Set the content type derived from the window name
		This.ContentType = This.GetContentType (lpcWindowName)

		* Finally, set the property value
		This.WindowName = lpcWindowName

	Endfunc

	Function ContentType_Assign (lpcContentType As String)

		* Set the associated icon that corresponds to this content type
		This.AssociatedIcon = This.GetAssociatedIcon (lpcContentType)

		* Finally, set the property value
		This.ContentType = lpcContentType

	Endfunc

	Function GetContentType (lpcWindowName As String) As String

		Local lcContentType As String

		* [TODO] Complete remaining content types
		Do Case
			Case Empty (lpcWindowName)	&& should not occur
				lcContentType = 'VFP'

			Case 'project manager' $ Lower (lpcWindowName)
				lcContentType = 'PJX'

			Case 'form designer' $ Lower (lpcWindowName)
				lcContentType = 'SCX'

			Case 'class designer' $ Lower (lpcWindowName)		;
					Or 'class browser' $ Lower (lpcWindowName)
				lcContentType = 'VCX'

			Case 'database designer' $ Lower (lpcWindowName)
				lcContentType = 'DBC'

			Case 'table designer' $ Lower (lpcWindowName)
				lcContentType = 'DBF'

			Case 'report designer' $ Lower (lpcWindowName)
				lcContentType = 'FRX'

			Case 'menu designer' $ Lower (lpcWindowName)
				lcContentType = 'MNX'

			Case 'label designer' $ Lower (lpcWindowName)
				lcContentType = 'LBX'

			Case 'query designer' $ Lower (lpcWindowName)
				lcContentType = 'QPX'

			Case Used (lpcWindowName)
				lcContentType = 'DBF'

			Otherwise
				* Check if the window name contains the name of the file
				lpcWindowName = Strtran (lpcWindowName, ' ', '')
				lpcWindowName = Strtran (lpcWindowName, '*', '')

				If At ('.', lpcWindowName) > 0 And Len (Justext (lpcWindowName)) <= 3
					lcContentType = Upper (Justext (lpcWindowName))
				Else
					* Default to VFP type
					lcContentType = 'VFP'
				Endif

		Endcase

		Return lcContentType

	Endfunc

	Function GetAssociatedIcon (lpcContentType As String) As String

		Local lcAssociatedIcon As String

		* Set the associated icon that corresponds to this content type
		Do Case
			Case lpcContentType = 'VCX'
				lcAssociatedIcon = 'icoClass'

			Case lpcContentType = 'CUR'
				lcAssociatedIcon = 'icoCursor'

			Case lpcContentType = 'DBC'
				lcAssociatedIcon = 'icoDatabase'

			Case lpcContentType = 'SCX'
				lcAssociatedIcon = 'icoForm'

			Case lpcContentType = 'MNX'
				lcAssociatedIcon = 'icoMenu'

			Case lpcContentType = 'PRG'
				lcAssociatedIcon = 'icoProgram'

			Case lpcContentType = 'PJX'
				lcAssociatedIcon = 'icoProject'

			Case lpcContentType = 'FRX'
				lcAssociatedIcon = 'icoReport'

			Case lpcContentType = 'SLX'
				lcAssociatedIcon = 'icoSolution'

			Case lpcContentType = 'DBF'
				lcAssociatedIcon = 'icoTable'

			Case Inlist (lpcContentType, 'TXT', 'H', 'INI', 'LOG')
				lcAssociatedIcon = 'icoText'

			Case Inlist (lpcContentType, 'XML', 'XSD')
				lcAssociatedIcon = 'icoXml'

			Case Inlist (lpcContentType, 'XSL', 'XSLT')
				lcAssociatedIcon = 'icoXsl'

			Otherwise
				lcAssociatedIcon = 'icoVfp'

		Endcase

		Return lcAssociatedIcon

	Endfunc

Enddefine

*!!!*	* _________________________________________________________________
*!!!*	*
*!!!*	* Windows Message Events Handler
*!!!*	*
*!!!*	Define Class FoxTabsEventHandler As Custom

*!!!*		* _____________________________________________________________
*!!!*		*
*!!!*		* Define class properties

*!!!*		*!!!*	PrevWndFunc	= 0
*!!!*		LastHWnd = 0

*!!!*		* _____________________________________________________________
*!!!*		*
*!!!*		* Define class members
*!!!*		Function Init()

*!!!*			* Declare Win32 API functions
*!!!*			Declare Integer CallWindowProc In Win32API Integer lpPrevWndFunc, Integer HWnd, Integer Msg, Integer wParam, Integer Lparam
*!!!*			Declare Integer GetWindowLong In Win32API Integer HWnd, Integer nIndex

*!!!*			* Store handle for use in CallWindowProc
*!!!*			*!!!*	This.PrevWndFunc = GetWindowLong (_vfp.HWnd, GWL_WNDPROC)

*!!!*		Endfunc

*!!!*		Function Destroy()

*!!!*			* Unbind all windows message events
*!!!*			*!!!*	UnBindEvents(0)

*!!!*		Endfunc

*!!!*		*!!!*	Function WMEventHandler(hWnd As Integer, Msg As Integer, wParam As Integer, lParam As Integer)

*!!!*		*!!!*		Local oException As Exception
*!!!*		*!!!*		Local lnReturn As Integer

*!!!*		*!!!*		lnReturn = 0

*!!!*		*!!!*		* In some scenarios, the hWnd and/or Msg variables do not exist.
*!!!*		*!!!*		* I don't see how that's possible, given that they are defined above, but it happens.
*!!!*		*!!!*		* See http://www.codeplex.com/VFPX/WorkItem/View.aspx?WorkItemId=20989 for an example.
*!!!*		*!!!*		* My hope is that this code will prevent other scenarios that cause C5s.
*!!!*		*!!!*		If Type("hWnd") <> "N" or Type("Msg") <> "N"
*!!!*		*!!!*			Return 0
*!!!*		*!!!*		EndIf

*!!!*		*!!!*		Try
*!!!*		*!!!*			* Handle each windows message case
*!!!*		*!!!*			Do Case
*!!!*		*!!!*				*CASE INLIST(lParam, 1235848, 1236368, 1236268, 1238376, 1238848, 1234504, 1235024, 1237032, 1236884, 1235540, 1234504)
*!!!*		*!!!*					* Don't do anything if these paramaters are passed, they refer to compile dialog events
*!!!*		*!!!*					* This was the nasty workaround - we've left the code in here for laughs for now
*!!!*		*!!!*
*!!!*		*!!!*				Case Type('This.Parent.Parent') <> 'O'
*!!!*		*!!!*					* Don't do anything	- the simple Compile error work around
*!!!*		*!!!*					* If we can't reference our parent's parent, then something weird is happening
*!!!*		*!!!*					* eg we are in the middle of a compile and VFP has totally shelled out to a
*!!!*		*!!!*					* different memory space which can't even recognise our parent
*!!!*		*!!!*					* so just let it fall through with out trying to bind to the events
*!!!*		*!!!*
*!!!*		*!!!*				Case Msg = WM_CREATE
*!!!*		*!!!*
*!!!*		*!!!*					*JAL* VFP crashes if bound to WM_CREATE when Call Stack window is opened manually after
*!!!*		*!!!*					*	a program is suspended. This code temporarily unbinds WM_CREATE and turns it back on
*!!!*		*!!!*					*	in the WM_ACTIVATE event below.
*!!!*		*!!!*					* See http://www.codeplex.com/VFPX/WorkItem/View.aspx?WorkItemId=19307
*!!!*		*!!!*					If Lower(This.Parent.GetWindowTitle(hWnd)) =  "call stack" and Program(0) <> "FOXTABS"
*!!!*		*!!!*						UnBindEvents(0, WM_CREATE)
*!!!*		*!!!*						BindEvent(0, WM_ACTIVATE, This, "WMEventHandler")
*!!!*		*!!!*					EndIf
*!!!*		*!!!*
*!!!*		*!!!*					* Raise the window create event
*!!!*		*!!!*					RaiseEvent(This, "WindowCreateEvent", hWnd)
*!!!*		*!!!*
*!!!*		*!!!*					* Bind to these events so we can add it to our collection
*!!!*		*!!!*					BindEvent(hWnd, WM_SHOWWINDOW, This, "WMEventHandler", 4)
*!!!*		*!!!*					BindEvent(hWnd, WM_SETTEXT, This, "WMEventHandler", 4)
*!!!*		*!!!*
*!!!*		*!!!*				Case Msg = WM_ACTIVATE
*!!!*		*!!!*					* Rebind CM_CREATE event after unbound above
*!!!*		*!!!*					UnBindEvents(0, WM_ACTIVATE)
*!!!*		*!!!*					BindEvent(0, WM_CREATE, This, "WMEventHandler")

*!!!*		*!!!*				Case Msg = WM_SHOWWINDOW

*!!!*		*!!!*					If wParam # 0
*!!!*		*!!!*						* Raise the window show event
*!!!*		*!!!*						RaiseEvent(This, "WindowShowEvent", hWnd)
*!!!*		*!!!*					EndIf
*!!!*		*!!!*
*!!!*		*!!!*					* Unbind to the this event as we do not require it any more
*!!!*		*!!!*					UnBindEvents(hWnd, WM_SHOWWINDOW)
*!!!*		*!!!*
*!!!*		*!!!*				Case InList(Msg, WM_SETFOCUS, WM_WINDOWPOSCHANGED, WM_CHILDACTIVATE)
*!!!*		*!!!*					* Raise the window set focus event
*!!!*		*!!!*					* Only raise when window is different than last time, to avoid flicker in toolbar
*!!!*		*!!!*					If hWnd <> This.LastHWnd
*!!!*		*!!!*						This.LastHWnd = hWnd
*!!!*		*!!!*						RaiseEvent(This, "WindowSetFocusEvent", hWnd)
*!!!*		*!!!*					EndIf

*!!!*		*!!!*				Case Msg = WM_SETTEXT
*!!!*		*!!!*					* Raise the window set text event
*!!!*		*!!!*					RaiseEvent(This, "WindowSetTextEvent", hWnd)
*!!!*		*!!!*
*!!!*		*!!!*				Case Msg = WM_DESTROY
*!!!*		*!!!*					* Raise the window destroy event
*!!!*		*!!!*					RaiseEvent(This, "WindowDestroyEvent", hWnd)
*!!!*		*!!!*
*!!!*		*!!!*			EndCase
*!!!*		*!!!*
*!!!*		*!!!*		Catch To oException
*!!!*		*!!!*			* Raise error event
*!!!*		*!!!*
*!!!*		*!!!*			* log the event call - used for debugging the compile dialog error
*!!!*		*!!!*			* Temporary code for catching the compile crash lParam
*!!!*		*!!!*			* Again, just left in for amusement value for now
*!!!*		*!!!*			* lcLog = CHR(13) + TRANSFORM(DATETIME()) + ' - ' + TRANSFORM(hWnd) + ' - ' + TRANSFORM(Msg) + ' - ' + TRANSFORM(wParam) + ' - ' + TRANSFORM(lParam) + ' - Crash '
*!!!*		*!!!*			* = STRTOFILE( lcLog, 'c:\foxtabscrashlog.txt', 1)

*!!!*		*!!!*			If Type('This.Parent.Parent') = 'O' and PemStatus(This.Parent.Parent, "LogError", 5)
*!!!*		*!!!*				RaiseEvent(This.Parent.Parent, "LogError", oException, "Exception caught while handling windows message event." + Chr(13) + " hWnd:" + Transform(hWnd, "@x0") + " Msg:" + Transform(Msg, "@x0") + " wParam:" + Transform(wParam, "@x0") + " lParam:" + Transform(lParam, "@x0"))
*!!!*		*!!!*			EndIf
*!!!*		*!!!*
*!!!*		*!!!*		Finally
*!!!*		*!!!*			* Must pass the message on
*!!!*		*!!!*			lnReturn = CallWindowProc(This.PrevWndFunc, hWnd, Msg, wParam, lParam)
*!!!*		*!!!*
*!!!*		*!!!*		EndTry

*!!!*		*!!!*		Return lnReturn
*!!!*		*!!!*
*!!!*		*!!!*	EndFunc

*!!!*		Function SetBindings (HWnd As Integer)

*!!!*			*!!!*	* Setup event bindings for this window
*!!!*			*!!!*	BindEvent(hWnd, WM_DESTROY, This, "WMEventHandler", 4)
*!!!*			*!!!*	BindEvent(hWnd, WM_SETTEXT, This, "WMEventHandler", 4)
*!!!*			*!!!*	BindEvent(hWnd, WM_SETFOCUS, This, "WMEventHandler", 4)
*!!!*			*!!!*	* Some IDE Windows don't have WM_SETFOCUS event, so use other events
*!!!*			*!!!*	If InList(Lower(This.Parent.GetWindowTitle(hWnd)), "project manager", "properties", "class designer", "form designer")
*!!!*			*!!!*		BindEvent(hWnd, WM_WINDOWPOSCHANGED, This, "WMEventHandler", 4)
*!!!*			*!!!*	EndIf

*!!!*		Endfunc

*!!!*		* _____________________________________________________________
*!!!*		*
*!!!*		* Define event delegates

*!!!*		Function WindowCreateEvent (HWnd As Integer)
*!!!*			* Event interface
*!!!*		Endfunc

*!!!*		Function WindowShowEvent (HWnd As Integer)
*!!!*			* Event interface
*!!!*		Endfunc

*!!!*		Function WindowDestroyEvent (HWnd As Integer)
*!!!*			* Event interface
*!!!*		Endfunc

*!!!*		Function WindowSetFocusEvent (HWnd As Integer)
*!!!*			* Event interface
*!!!*		Endfunc

*!!!*		Function WindowSetTextEvent (HWnd As Integer)
*!!!*			* Event interface
*!!!*		Endfunc

*!!!*		* _____________________________________________________________

*!!!*	Enddefine

*!!!*	* _________________________________________________________________
