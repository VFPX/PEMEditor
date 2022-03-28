lparameters toObject, tcWhereFrom
local lcCaption
toObject.shpBox.Width  = toObject.Width
toObject.shpBox.Height = toObject.Height - toObject.shpBox.Top
lcCaption = inputbox('Label caption', 'BaseLabelledBox Builder', ;
	toObject.lblBox.Caption)
if not empty(lcCaption)
	toObject.lblBox.Caption = lcCaption
endif not empty(lcCaption)
