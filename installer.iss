[Setup]
AppName=DUA Automation
AppVersion=1.0.1
DefaultDirName={autopf}\DUA Automation
DefaultGroupName=DUA Automation
UninstallDisplayIcon={app}\DUA_Automation.exe
Compression=lzma2
SolidCompression=yes
OutputDir=Output
OutputBaseFilename=DUA_Automation_Setup

[Files]
Source: "dist\DUA_Automation.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "resources\*"; DestDir: "{app}\resources"; Flags: ignoreversion recursesubdirs

[Icons]
Name: "{group}\DUA Automation"; Filename: "{app}\DUA_Automation.exe"
Name: "{commondesktop}\DUA Automation"; Filename: "{app}\DUA_Automation.exe"

[Run]
Filename: "{app}\DUA_Automation.exe"; Description: "Iniciar DUA Automation agora"; Flags: nowait postinstall skipifsilent
