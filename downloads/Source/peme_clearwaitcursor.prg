*!* ClearWaitCursor() ... courtesy of Cesar Chalom
Local lcPoint As String, lnx As Integer, lny As Integer, loPosition

Declare Integer GetCursorPos In win32api As ClearWaitCursor_GetCursorPos String  @lpPoint
Declare Integer SetCursorPos In win32api As ClearWaitCursor_SetCursorPos Integer nX, Integer nY

m.lcPoint = 0h0000000000000000
ClearWaitCursor_GetCursorPos(@m.lcPoint)

loPosition = this.CalculateShortcutMenuPosition()

ClearWaitCursor_SetCursorPos(loPosition.Row, loPosition.Column)

Return
