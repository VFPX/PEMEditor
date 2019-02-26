#include EditPropertyDialogEnglish.H
#include Beautify.H
#include FoxPro.H

#Define CR 								Chr[13]
#Define LF								Chr(10)
#Define CRLF							CR + LF
#Define TAB								Chr(9)
#Define BLANKS							' ' + Tab
#Define ISABLANK 						$ BLANKS
#Define LINEEND							CR + LF
* list of characters that can't be part of a name assigned a value;
* note that period (.) is not in the list, intentionally
#Define NOTNAMECHARS					(' !"#$%&()*+,-/:;<=>?@[\]^`{|}~' + ['] + Tab)
#Define ide_LOC_PutCursorHere			'^^^'

#define cnMAX_MEMBER_DATA_SIZE			8192
	* maximum size of member data
#define ccGLOBAL_MEMBER_DATA_TYPE		'E'
	* value of TYPE in FOXCODE for global member data
#define ccMEMBER_DATA_XML_ELEMENT		'memberdata'
	* the member data element in the member data XML
#define ccXML_TRUE						'True'
	* the value for true in the member data XML
#define ccXML_ROOT_NODE					'VFPData'
	* the root node for the member data XML
#define ccXML_DOM_CLASS					'MSXML2.DOMDocument'
	* the class to use for the XML DOM object
	* the code to put into the Assign method
#define ccCR							chr(13)
	* carriage return
#define ccLF							chr(10)
	* linefeed
#define ccCRLF							chr(13) + chr(10)
	* carriage return + linefeed
#define ccTAB							chr(9)
	* tab
#define ccMaxDescriptionLength			254
	* maximum length for descriptions of PEMs

#define clLOWER_CASE_FIRST_LETTER		.F.

#define cnDEFAULT_COLORS				[None, 0, None, 0, None, 0, None, 0, None, 0, Bold, 0, None, 0, Bold, 0, None, 0, BackColor, 14737632, BackColor, 14737632 ]

#define cnSAMPLE_COLORS_2				[ForeColor, 16711680, ForeColor, 255, None, 0, None, 0, BackColor, 14737632, Bold, 0, None, 0, Bold, 0, BackColor, 16777215, BackColor, 14737632, BackColor, 14737632 ]

#define cnSAMPLE_COLORS_3				[ForeColor, 234705, ForeColor, 16744576, None, 0, None, 0, BackColor, 14737632, Bold, 0, None, 0, Bold, 0, BackColor, 16777215, BackColor, 14737632, BackColor, 14737632 ]

#define cnSAMPLE_COLORS_4				[None, 0, None, 0, BackColor, 15915974, BackColor, 9959417, BackColor, 16777215, Bold, 0, None, 0, BackColor, 13434828, BackColor, 11788021, BackColor, 16777215, BackColor, 12632256 ]

#define ccTopDividerLine				'*' + Replicate('=', 80) + ccCRLF
#define ccBottomDividerLine				'*' + Replicate('-', 80) + ccCRLF
#define ccXML_FALSE 					'False'
#define ccPROPERTIES_PADDING_CHAR		chr(1)
	* the padding character used for properties with values > 255 characters
#define cnPROPERTIES_PADDING_SIZE		517
	* the size of the padding area for properties with values > 255 characters
#define cnPROPERTIES_LEN_SIZE			8
	* the size of the length structure for properties with values > 255 characters
#Define SM_CMONITORS 80

#Define ccFontSizeName					'Tahoma'
#Define ccFontSizeLarge					9
#Define ccFontSizeMedium				8
#Define ccFontSizeSmall					7

#Define ccNameCol						1
#Define ccTypeCol						2
#Define ccAccessCol						3
#Define ccAssignCol						4
#Define ccVisibilityCol					5
#Define ccNativeCol						6
#Define ccInheritedCol					7
#Define ccNonDefaultCol					8
#Define ccHasCodeCol					9
#Define ccFavoritesCol					10
#Define ccReadOnlyXCol					11
#Define ccDescriptCol					12
#Define ccPasteDescriptCol 				12
#Define ccObjNumberCol					13
#Define ccPastecValueCol 				14
#Define ccScript						15

#Define ccPasteValueCol 				16
#Define ccPasteTypeCol 					17
#Define ccPasteSelectCol 				18
#Define ccPastelNewCol 					19

#Define ccHandlerPrefix					"_pemeditor_"

#Define ccTextFileExtensions			" ASP ASPX BAS BAT C CGI CP CPP CSS FPW FRA H HTM HTML INI LOG MPR PHP PS1 PY PYC PYO QPR RB SCA SQL TXT VCA XML "
#Define ccCharacterProperties			'CAPTION', 'COLUMNWIDTHS', 'COMMENT', 'DISPLAYVALUE', 'FORMAT', 'INPUTMASK', 'TAG', 'TOOLTIPTEXT'