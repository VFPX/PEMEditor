Local lcCurrentLine, lcNewText, lcText, lnPos, lnSelEnd, lnSelStart, lnStartOfLine, loEditorWin
loEditorWin = _oPEMEditor.oEditorWin

* Find current window; exit if none
If loEditorWin.FindWindow() < 0
	Return
Endif

* Current cursor position
lnSelStart = loEditorWin.GetSelStart()

* Start of current line
lnStartOfLine = loEditorWin.SkipLine (lnSelStart, 0)

* Text of current line, up to cursor
lcCurrentLine = loEditorWin.GetString (lnStartOfLine, lnSelStart)

lnPos = Rat ('##', lcCurrentLine)
If lnPos = 0
	Return
Endif

lnSelEnd   = lnSelStart
lnSelStart = lnStartOfLine + lnPos - 1
lcText	   = loEditorWin.GetString (lnSelStart, lnSelEnd)

* pass parameters to routine DOIT
lcNewText = DOIT (Substr (lcText, 3))

* Select our text
loEditorWin.Select (lnSelStart, lnSelEnd)

* and paste in the new text
_Cliptext = lcNewText

loEditorWin.Paste()


Procedure DOIT (lcText)
	Return '$$$$ Look at this: ' + lcText + ' !!!!'
Endproc
