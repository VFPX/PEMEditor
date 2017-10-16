#Define CR						CHR(13)

#Define DV_Key_Back				'* F7'
#Define DV_Key_Beautify			'* F6'
#Define DV_Key_CloseAllWindows	'* Shift+F8'
#Define DV_Key_CtrlC			''
#Define DV_Key_CtrlCAdditive	'* Alt+C'
#Define DV_Key_CtrlX			'* Ctrl+X'
#Define DV_Key_CtrlXAdditive	'* Alt+X'
#Define DV_Key_DoubleHash		'* F5'
#Define DV_Key_ExtractToMethod	'* Ctrl+F8'
#Define DV_Key_FormatMenu		'* Shift+F5'
#Define DV_Key_Forward			'* F8'
#Define DV_Key_GoToDef			'* F12'
#Define DV_Key_IDList			'* Ctrl+F6'
#Define DV_Key_Locals			'* Shift+F6'
#Define DV_Key_Launch			'* Alt+F12'
#Define DV_Key_LaunchDocTreeView 	'* Alt+F11'
#Define DV_Key_OpenMenu			'* Ctrl+0'
#Define DV_Key_MoveWindow		'* F11'

#Define DV_SortOrder1			'+cName'
#Define DV_SortOrder2			[+IIF(cType = 'P', '0', IIF(lNonDefault, '1', IIF(lHasCode, '2', IIF(Not lNative, '3', '4')))) + CName]

#define DV_DEFAULT_COLORS		[None, 0, None, 0, None, 0, None, 0, None, 0, Bold, 0, None, 0, Bold, 0, None, 0, BackColor, 14737632, BackColor, 14737632 ]

#define DV_SAMPLE_COLORS_2		[ForeColor, 16711680, ForeColor, 255, None, 0, None, 0, BackColor, 14737632, Bold, 0, None, 0, Bold, 0, BackColor, 16777215, BackColor, 14737632, BackColor, 14737632 ]

#define DV_SAMPLE_COLORS_3		[ForeColor, 234705, ForeColor, 16744576, None, 0, None, 0, BackColor, 14737632, Bold, 0, None, 0, Bold, 0, BackColor, 16777215, BackColor, 14737632, BackColor, 14737632 ]

#define DV_SAMPLE_COLORS_4		[None, 0, None, 0, BackColor, 15915974, BackColor, 9959417, BackColor, 16777215, Bold, 0, None, 0, BackColor, 13434828, BackColor, 11788021, BackColor, 16777215, BackColor, 12632256 ]

#Define DV_ACCESS_CODE_Scalar			'return This.PEM_Name_Place_Holder'

#Define DV_ACCESS_CODE_Array			'LParameters tnDim1, tnDim2'			 			+ CR + ;
	CR +										        	   ;
	'Do Case '								 			+ CR + ;
	CR +										        	   ;
	'    * Normal (not an array) ' 							+ CR + ;
	'    Case PCount() = 0' 							+ CR + ;
	'        return This.PEM_Name_Place_Holder' 		+ CR + ;
	CR +										        	   ;
	'    * Array, one dimension ' 			 			+ CR + ;
	'    Case PCount() = 1' 							+ CR + ;
	'        return This.PEM_Name_Place_Holder(tnDim1)' + CR + ;
	CR +										        	   ;
	'    * Array, two dimensions ' 						+ CR + ;
	'    Case PCount() = 2 ' 							+ CR + ;
	'        return This.PEM_Name_Place_Holder(tnDim1, tnDim2)' + CR + ;
	CR +										        	   ;
	'EndCase' + CR

#Define DV_ASSIGN_CODE_Scalar			'lparameters tPEM_Name_Place_Holder' + Chr(13) + 'This.PEM_Name_Place_Holder = tPEM_Name_Place_Holder'
#Define DV_Assign_CODE_Array			'LParameters tPEM_Name_Place_Holder, tnDim1, tnDim2'+ CR + ;
	CR +										        	   ;
	'Do Case '								 			+ CR + ;
	CR +										        	   ;
	'    * Normal (not an array) ' 							+ CR + ;
	'    Case PCount() = 1' 							+ CR + ;
	'        This.PEM_Name_Place_Holder = tPEM_Name_Place_Holder' 		+ CR + ;
	CR +										        	   ;
	'    * Array, one dimension ' 			 			+ CR + ;
	'    Case PCount() = 2' 							+ CR + ;
	'        This.PEM_Name_Place_Holder (tnDim1) = tPEM_Name_Place_Holder' + CR + ;
	CR +										        	   ;
	'    * Array, two dimensions ' 						+ CR + ;
	'    Case PCount() = 3 ' 							+ CR + ;
	'        This.PEM_Name_Place_Holder (tnDim1, tnDim2) = tPEM_Name_Place_Holder' + CR + ;
	CR +										        	   ;
	'EndCase' + CR
