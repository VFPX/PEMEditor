#include MemberDataEditorEnglish.H
#include FoxPro.H

#define cnMAX_MEMBER_DATA_SIZE			8192
	* maximum size of member data
#define ccGLOBAL_MEMBER_DATA_TYPE		'E'
	* value of TYPE in FOXCODE for global member data
#define ccMEMBER_DATA_XML_ELEMENT		'memberdata'
	* the member data element in the member data XML
#define ccXML_TRUE						'True'
	* the value for true in the member data XML
#define ccXML_FALSE						'False'
	* the value for false in the member data XML
#define ccXML_DECLARATION				''
	* the XML declaration at the start of the member data XML; formerly used '<?xml version="1.0" encoding="Windows-1252" standalone="yes"?>'
#define ccXML_ROOT_NODE					'VFPData'
	* the root node for the member data XML
#define ccXML_DOM_CLASS					'MSXML2.DOMDocument'
	* the class to use for the XML DOM object
#define cnCALLED_FROM_CLASS				0
	* the value passed to the editor by Builder.APP when called from the Class menu
#define cnCALLED_FROM_FORM				1
	* the value passed to the editor by Builder.APP when called from the Form menu
#define cnADD_TO_FAVORITES_CLASS		10
	* the value passed to the editor by Builder.APP for an "Add to Favorites" call from the Class Designer
#define cnADD_TO_FAVORITES_FORM			11
	* the value passed to the editor by Builder.APP for an "Add to Favorites" call from the Form Designer
#define clAUTO_CREATE_MEMBER_DATA		.T.
	* .T. to automatically create an _MemberData property for objects when
	* they're opened in the Form or Class Designer
#define ccGETMEMBERDATA_ABBREV			'_GetMemberData'
	* the value of FOXCODE.ABBREV for the _GetMemberData record
#define ccCRLF							chr(13) + chr(10)
	* carriage return + linefeed
#define ccCR							chr(13)
	* carriage return
#define ccTAB							chr(9)
	* tab
#define ccPROPERTIES_PADDING_CHAR		chr(1)
	* the padding character used for properties with values > 255 characters
#define cnPROPERTIES_PADDING_SIZE		517
	* the size of the padding area for properties with values > 255 characters
#define cnPROPERTIES_LEN_SIZE			8
	* the size of the length structure for properties with values > 255 characters
#define cnSCOPE_OBJECT					1
	* the value of the option group for the Object scope
#define cnSCOPE_GLOBAL					2
	* the value of the option group for the Global scope
#define cnSCOPE_CONTAINER				3
	* the value of the option group for the Container scope

				