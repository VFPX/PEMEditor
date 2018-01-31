Lparameters loUpdateObject

With loUpdateObject
    .AppName              = 'PEMEditor.APP'
    .ApplicationName      = 'PEM Editor'
    .ToolName             = 'Thor_Tool_PEME_GoToDefinition'
    .VersionFileURL       = 'https://raw.githubusercontent.com/VFPX/PEMEditor/master/ThorUpdater/_PEMEditorVersionFile.txt'
    .VersionLocalFilename = 'PEM Editor Version.txt'
    .RegisterWithThor     = 'Do "##InstallFolder##PEMEditor.APP" with "Thor"'
    .ShowErrorMessage     = 'No'
    .SourceFileURL        = 'PEMEditor_Source.Zip = https://raw.githubusercontent.com/VFPX/PEMEditor/master/ThorUpdater/PEMEditor_Source.Zip'
    .Link                 = 'https://github.com/VFPX/PEMEditor'
    .LinkPrompt           = 'PEM Editor Home Page'
Endwith

AddProperty(loUpdateObject, 'UpdateNowIfNotInstalled', 'Yes')    

Return loUpdateObject

