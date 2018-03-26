Lparameters tcStyle, tcPem, toObject

* lcStyle =
*     	'Code'
**		'Value'

loClassCollection = CreateObject('Collection')

*  tcPem = Alltrim(ThisForm.oOriginalMember.cname)
Do case
	case "O" = Vartype (toObject)
		loObject = toObject
	Case 0 # ASelObj(laObjects) or 0 # ASelObj(laObjects,1)
		loObject = laObjects(1)
	Otherwise
		Return loClassCollection
Endcase

If Empty (tcPem) Or PEMStatus (loObject, tcPem, 6)
	loSetupInfo = This.getparentcodesetupinfo (			;
		Lower(loObject.Class)							;
		, Iif(Thisform.lClass And Thisform.oTopOfForm = loObject, Thisform.cSourceFileName, loObject.ClassLibrary)	;
		, loObject										;
		)

	With loSetupInfo
		lcClass		= .Class
		lcClassLib	= .Classlib
		lcParent	= .Parent
	Endwith

	loDetailInfo = This.getparentcodedetailinfo (loClassCollection, lcClass, lcClassLib, lcParent, loObject.Name, lcStyle, tcPem)
	With loDetailInfo
		lcClass		= .Class
		lcClassLib	= .Classlib
		lcParent	= .Parent
	Endwith

	*** JRN 12/26/2008 : Handling of members (pages, optionbuttons, etc)
	Do Case
		Case Thisform.lIsTopOfForm && nothing to do

		Case Lower(lcClass) == Lower(loObject.BaseClass) && nothing here either

		Case Thisform.GetPEMStatus(loObject.Parent, "MemberClass", 5)
			loDetailInfo = This.getparentcodedetailinfo (lcStyle, loObject.Parent.MemberClass, loObject.Parent.MemberClassLibrary, '', tcPem, loObject)
			With loDetailInfo
				lcResult	= lcResult + Iif (Inlist (lcStyle, 'First Code', 'Inherited Code', 'Inherited Value') And Not Empty (lcResult), '', .Result)
				lcClass		= .Class
				lcClassLib	= .Classlib
				lcParent	= .Parent
			Endwith

		Case Thisform.GetPEMStatus(loObject.Parent, "HeaderClass", 5)
			loDetailInfo = This.getparentcodedetailinfo (lcStyle, loObject.Parent.HeaderClass, loObject.Parent.HeaderClassLibrary, '', tcPem, loObject)
			With loDetailInfo
				lcResult	= lcResult + Iif (Inlist (lcStyle, 'First Code', 'Inherited Code', 'Inherited Value') And Not Empty (lcResult), '', .Result)
				lcClass		= .Class
				lcClassLib	= .Classlib
				lcParent	= .Parent
			Endwith

	Endcase

	****************************************************************
	If 'L' # Vartype (lcClassLib)
		Do Case
			Case Inlist (lcStyle, 'Code', 'All Code')
				lcResult = lcResult + This.addparentcodetext(lcStyle, ccLOC_VFP_Base_Class + lcClass, ccLOC_VFP_Base_Class + lcClass, '', '')
			Case Inlist (lcStyle, 'Value', 'Inherited Value')
				Local laProps(2)
				Select * From DefaultNativeProperties 				;
					Where cpropname = Upper(tcPem) 				;
					And (Empty(cbaseclass) Or cbaseclass = Upper(loObject.BaseClass)) Into Array laProps
				Do Case
					Case _Tally = 0
						lcResult = lcResult + This.addparentcodetext(lcStyle, ccLOC_VFP_Base_Class + lcClass, ccLOC_VFP_Base_Class + lcClass, '', '')
					Case lcStyle = 'Value'
						lcResult = lcResult + This.addparentcodetext(lcStyle, ccLOC_VFP_Base_Class + lcClass, ccLOC_VFP_Base_Class + lcClass, '',		;
							tcPem + ' = ' + Alltrim(laProps(_Tally, 3)))
					Case Empty(lcResult)
						lcResult = lcResult + This.addparentcodetext(lcStyle, ccLOC_VFP_Base_Class + lcClass, ccLOC_VFP_Base_Class + lcClass, '',		;
							Alltrim(laProps(_Tally, 3)))
				Endcase
		Endcase
	Endif

Endif

Use
Select (lnSelect)
Return lcResult

Use

Select (lnSelect)

Return lcResult


Lparameters lcClass, lcClassLib, loObject
Local loResult, lcParent

loResult = Createobject("Empty")

Use In (Select("VCX"))
*** JRN 12/12/2008 : special case for classes when at top of form
* since class and classlib point to the current class, not the parent
If Thisform.lClass And loObject = Thisform.oTopOfForm
	Use (lcClassLib) Again Shared Alias VCX
	Locate For Lower(OBJNAME) == lcClass and Lower(RESERVED1) = 'class'

	If Found()
		lcClass = VCX.Class
		If Not Empty(VCX.ClassLoc)
			lcClassLib = Thisform.GetClassFileName(VCX.ClassLoc, Addbs(Justpath(lcClassLib)))
		Else
			lcClassLib = ''
		Endif
	Else
		lcClass = ''
		lcClassLib = ''
	Endif
Else
	* if this object was inherited as an object contained in another class,
	* get the full object name in that class
	* special handling for pages, columns, and buttons, since their names are not read-only
	* even when inherited

	Do While .T.

		Do Case
			Case loObject = Thisform.oTopOfForm
				Exit && obviously at top of class/form
			Case Thisform.Isreadonly(loObject, "Name")
				* if name is read-only, continue -- must be part of parent class
			Case Inlist(Lower(loObject.BaseClass) + '#', 'page#', 'column#', 'optionbutton#', 'header#')
				* any of these baseclasses must be part of parent class ... #'s because of pageframe match
			Otherwise && for all others, we must be done
				Exit
		Endcase

		lcParent = loObject.Parent.Name + Iif( Empty(lcParent), '', "." + lcParent)
		loObject = loObject.Parent
		lcClassLib = loObject.ClassLibrary
	Enddo

	* but the topmost name must be the name of the class, not the name of the object here
	
	If Thisform.lClass and loObject = ThisForm.oTopOfForm
		lcClassName = loObject.ParentClass
	Else
		lcClassName = loObject.Class
	EndIf
	
	Do Case
		Case Empty(lcParent)

		Case Not "." $ lcParent
			lcParent = lcClassName 

		Otherwise
			lcParent = lcClassName  + Substr(lcParent, At(".", lcParent))

	Endcase
Endif

Use In (Select("VCX"))
AddProperty (loResult, "Class", 	lcClass)
AddProperty (loResult, "ClassLib", 	lcClassLib)
AddProperty (loResult, "Parent", 	lcParent)
Return loResult
