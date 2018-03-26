lcBreakWords = '\Wselect\W|\Wfrom\W|\Wwhere\W|\Wgroup by\W|\Worder by\W|\Whaving\W|\Wunion\W|\Winto\W'

lcBreakWords = '\,|' + lcBreakWords

lcBreakWords = lcBreakWords + '|\Wleft join\W|\Wleft outer join\W|\Wright join\W|\Wright outer join\W|\Wfull join\W' ;
	+ '|\Wfull outer join\W|\Winner join\W|\Wjoin\W'

lcBreakWords = lcBreakWords + '|\Won\W'

lcBreakWords = lcBreakWords + '|\W.and.\W|\W.or.\W|\Wand\W|\Wor\W'

ALines(laLines, lcBreakwords, 0, '|')

lcxxx = '(\,|\Wselect\W|\Wfrom\W|\Won\W|\Wleft join\W|\Wleft outer join\W|\Wright join\W|\Wright outer join\W|\Wfull join\W' ;
	+ '|\Wfull outer join\W|\Wjoin\W|\Wwhere\W|\Worder by\W|\Wgroup by\W|\W.and.\W|\W.or.\W|\Wand\W|\Wor\W|\Whaving\W|\Wunion\W|\Winto\W)'

For lnI = 1 to Alen(laLines)
	? laLines[lnI], At(laLines[lnI], lcxxx)
EndFor 