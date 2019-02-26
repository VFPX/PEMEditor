#include AnchorEditor.H
local lcPath, ;
	laObjects[1], ;
	lnObjects

* Get the path.

lcPath = addbs(justpath(sys(16)))

* Get a reference to the selected object. If there isn't one, this may be a
* "setup" call, so do the setup tasks.

lnObjects = aselobj(laObjects)
if lnObjects = 0
	do SetupAnchorEditor with lcPath
	messagebox(ccLOC_EDITOR_REGISTERED, MB_OK + MB_ICONINFORMATION)
	return .F.
endif lnObjects = 0

* Run the form.

release _oAnchorEditor
public _oAnchorEditor
_oAnchorEditor = newobject('AnchorEditor', 'AnchorEditor.vcx')
if vartype(_oAnchorEditor) = 'O'
	_oAnchorEditor.Show()
else
	release _oAnchorEditor
endif vartype(_oAnchorEditor) = 'O'


function SetupAnchorEditor(tcPath)
local lcXML, ;
	lnSelect

* Create the XML for the Anchor entry in FOXCODE.

text to lcXML noshow textmerge
<VFPData><memberdata name="anchor" type="property" favorites="True"
script="do [<<tcPath>>AnchorEditor.app]"/></VFPData>
endtext

* Add the Anchor record to FOXCODE.

lnSelect = select()
select 0
use (_foxcode) again shared
locate for TYPE = 'E' and upper(ABBREV) = 'ANCHOR'
if found()
	replace TIP with lcXML
else
	insert into (_foxcode) (TYPE, ABBREV, TIP) ;
		values ('E', 'ANCHOR', lcXML)
endif found()
use
select (lnSelect)
return
