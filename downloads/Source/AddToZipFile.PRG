Lparameters tcFilename, tcSource

lcVFPCompressFileName = 'vfpcompression.fll'
Set Library To (lcVFPCompressFileName) Additive
If ZipOpen (tcSource, .T.)
	ZipFile (tcFileName, .t.)
	ZipClose()
	lnResult = 1
Else
	? 1 / ' '
Endif
