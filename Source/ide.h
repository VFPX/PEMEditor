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

#Define PutCursorHere 					'^^^'

