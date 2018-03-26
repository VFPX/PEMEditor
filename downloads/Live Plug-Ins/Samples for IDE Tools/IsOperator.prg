* 2011.11.22 MarioPeschke
* This method checks for specific operator's that can be used to add a "space" around opoerator's
* The list should be completed

Lparameters tcline,tncharpos
Local llreturn
llreturn = .F.
Do Case
	Case Atc("(",tcline)>0 AND Atc("(",tcline)<tncharpos
		llreturn = .T.
	Case Atc("=",tcline)>0 AND Atc("=",tcline)<tncharpos
		llreturn = .T.
	Case Atc("Store",tcline)>0 AND Atc("Store",tcline)<tncharpos
		llreturn = .T.
	Case Atc("For",tcline)>0 AND Atc("For",tcline)<tncharpos
		llreturn = .T.
	Case Atc("If",tcline)>0 AND Atc("If",tcline)<tncharpos
		llreturn = .T.
	Case Atc("Case",tcline)>0 AND Atc("Case",tcline)<tncharpos
		llreturn = .T.
	OTHERWISE
		llreturn = .F.
Endcase
Return llreturn
