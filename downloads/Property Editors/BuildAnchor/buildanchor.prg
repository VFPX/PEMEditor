local lcPath, laObjects[1], lnObjects, i

lcPath = addbs( justpath( sys( 16 ) ) )

lnObjects = aselobj( laObjects )
if ( lnObjects = 0 )
   * not launched from property sheet
	if ( SetupBuildAnchor( lcPath ) )
      messagebox( "BuildAnchor setup" )
   endif
endif

for i = 1 to _screen.FormCount
   if ( _screen.Forms[i].Name == "frmdfBuildAnchor" )
      * BuildAnchor is running
      _screen.Forms[i].cmdApply.SetFocus()
      exit
   endif
endfor

if ( i > _screen.FormCount )
   * form not already running so start it
   do form (lcPath + "BuildAnchor")
endif


function SetupBuildAnchor( tcPath )
local lcXML, lnSelect, llRetVal

llRetVal = .f.

* memberdata XML

text to lcXML noshow textmerge
<?xml version="1.0" encoding="Windows-1252" standalone="yes"?>
<VFPData>
<memberdata name="anchor" type="property" script="do '<<tcPath>>BuildAnchor.prg'"/>
</VFPData>
endtext

lnSelect = select()
select 0
use (_foxcode) again shared
locate for ( ( type = "E" ) and ( upper( abbrev ) = "ANCHOR" ) )
if found()
   if ( ! tip == lcXML )
      * update the entry
      replace tip with lcXML
      llRetVal = .t.
   endif
else
   * add the Anchor entry
   llRetVal = .t.
   insert into (_foxcode) ( type, abbrev, tip, timestamp, uniqueid ) ;
      values ( "E", "ANCHOR", lcXML, datetime(), sys( 2015 ) )
endif
use
select (lnSelect)

return llRetVal
