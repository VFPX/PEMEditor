* editorwin home page = http://vfpx.codeplex.com/wikipage?title=thor%20editorwindow%20object
Local loEditorWin as Editorwin of "c:\visual foxpro\programs\9.0\common\thor\tools\apps\pem editor\source\peme_editorwin.vcx"
loEditorWin = ExecScript(_Screen.cThorDispatcher, "Class= editorwin from pemeditor")
?loEditorWin.GetEnvironment[25]