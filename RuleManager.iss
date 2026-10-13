[Setup]
AppName=RuleManager
AppVersion=1.0.0
DefaultDirName={commonappdata}\Autodesk\Inventor Addins
DefaultGroupName=RuleManager
OutputDir=Output
OutputBaseFilename=RuleManagerSetup
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Dirs]
Name: "{commonappdata}\Autodesk\Inventor Addins"
Name: "D:\Bibliotecas\Ilogic"
Name: "D:\Bibliotecas\iFeatures"

[Files]
; Todo el AddIn
Source: "Release\*"; DestDir: "{commonappdata}\Autodesk\Inventor Addins"; Flags: ignoreversion recursesubdirs createallsubdirs overwritereadonly

; Biblioteca iLogic
Source: "D:\Bibliotecas\Ilogic\*"; DestDir: "D:\Bibliotecas\Ilogic"; Flags: ignoreversion recursesubdirs createallsubdirs overwritereadonly

; Biblioteca iFeatures
Source: "D:\Bibliotecas\iFeatures\*"; DestDir: "D:\Bibliotecas\iFeatures"; Flags: ignoreversion recursesubdirs createallsubdirs overwritereadonly

[Icons]
Name: "{group}\Desinstalar RuleManager"; Filename: "{uninstallexe}"