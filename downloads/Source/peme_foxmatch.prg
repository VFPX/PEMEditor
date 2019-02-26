#include "peme_foxref.h"
#include "foxpro.h"

* object we use to populate our collection of Matches
Define Class Match As Custom
	Name = "Match"
Enddefine

* Abtract Match Engine Class
Define Class MatchEngine As Custom
	Name 		= "MatchEngine"
	FileName 	= ''
	SearchType	= ''
	ClassSearch	= ''
	ClassExactMatch	= .F.
	SearchBaseClass	= ''

	Function SearchFor
		Return .T.
	Endfunc

	Function AddMatch (ttTimeStamp, tcClass, tcParentClass, tcClassLoc, tcBaseClass, tnLineNumber)
	
		ltTimeStamp 	= Evl (ttTimeStamp, Fdate (This.FileName, 1))
		lcClass			= Evl (tcClass, '')
		lcParentClass	= Evl (tcParentClass, '')
		If Empty (Evl (tcClassLoc, ''))
			lcClassLoc		= ''
		Else
			lcClassLoc		= This.GetClassFileName(tcClassLoc, This.FileName)
		Endif
		lcBaseClass		= Evl (tcBaseClass, '')

		Insert Into ResultsTable	(	;
			SearchType				,	;
			FullName				,	;
			Class					,	;
			Folder					,	;
			FileName				,	;
			Parent					,	;
			ClassLoc				,	;
			BaseClass				,	;
			TimeStamp				,	;
			LineNumber					;
			)     						;
			Values						;
			(							;
			This.SearchType			,	;
			This.FileName			,	;
			lcClass					,	;
			JustPath(This.FileName)	,	;
			JustFname(This.FileName),	;
			lcParentClass			,	;
			lcClassLoc				,	;
			lcBaseClass				,	;
			ltTimeStamp				,	;
			Evl(tnLineNumber, 0)        ;
			)
	Endfunc


	Function SearchBaseClass_Assign (lcNewValue)
		If lcNewValue = '<'
			This.SearchBaseClass = ''
		Else
			This.SearchBaseClass = Alltrim (Lower (lcNewValue))
		Endif
	Endfunc


	Function ClassSearch_Assign (lcNewValue)
		This.ClassSearch = Alltrim (Lower (lcNewValue))
	Endfunc


	Function DecodeTimeStamp (nTimestamp)

		Local nDate, nDay, nHr, nMin, nMonth, nSec, nTime, nYear

		nDate=Bitrshift(nTimestamp,16)
		nTime=Bitand(nTimestamp,2^16-1)

		nYear=Bitand(Bitrshift(nDate,9),2^8-1)+1980
		nMonth=Bitand(Bitrshift(nDate,5),2^4-1)
		nDay=Bitand(nDate,2^5-1)

		nHr=Bitand(Bitrshift(nTime,11),2^5-1)
		nMin=Bitand(Bitrshift(nTime,5),2^6-1)
		nSec=Bitand(nTime,2^5-1)

		Return Datetime(nYear,nMonth,nDay,nHr,nMin,nSec)

	Endfunc


	Function GetClassFileName
		Lparameters tcClassLoc, tcPath

		*** JRN 02/24/2009 : per Doug ... handles cases of absolute paths and relative to current directory

		Do Case
			Case ":" $ tcClassLoc
				Return tcClassLoc

			Case File( Fullpath( m.tcClassLoc, m.tcPath ) )
				Return Fullpath(tcClassLoc, tcPath)

			Otherwise
				Return Fullpath( m.tcClassLoc )
		Endcase

	Endfunc

Enddefine


* Standard Match Engine
Define Class MatchDefault As MatchEngine
	Name = "MatchDefault"

	Function SearchFor()
		This.AddMatch()
	Endfunc

Enddefine


* Match class names
Define Class MatchClass As MatchEngine
	Name = "MatchWildcard"

	Function SearchFor()
		Select (Select('VCX'))
		Use (This.FileName) Shared Again Alias VCX

		Scan For Lower(Reserved1) = 'class' And Not Deleted()
			If Empty(This.SearchBaseClass) Or Lower(BaseClass) = This.SearchBaseClass
				If Empty (This.ClassSearch) Or Lower (This.ClassSearch) $ Lower(objname)
					If (Not This.ClassExactMatch) Or Lower (This.ClassSearch) == Lower(objname)

						This.AddMatch (This.DecodeTimeStamp(Timestamp), objname, Class, ClassLoc, BaseClass)

					Endif
				Endif
			Endif
		Endscan

		Use

	Endfunc

Enddefine

