; Inno Setup Script for Agro DSR Professional v7.0
[Setup]
AppName=Agro DSR Professional
AppVersion=7.0
AppPublisher=Agro Thai - Rajnikant Joshi
DefaultDirName={pf}\Agro DSR Professional
DefaultGroupName=Agro DSR Professional
OutputDir=.
OutputBaseFilename=Agro_DSR_Pro_Setup_v7.0
SetupIconFile=assets\xtreme.ico
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
DisableProgramGroupPage=yes
AppMutex=AgroDSRProMutex
CloseApplications=yes
RestartApplications=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "dist\Agro_DSR_Pro.exe"; DestDir: "{app}"; Flags: ignoreversion
; staff_registry.json is created by the app on first run — not bundled

[Icons]
Name: "{group}\Agro DSR Professional"; Filename: "{app}\Agro_DSR_Pro.exe"
Name: "{commondesktop}\Agro DSR Professional"; Filename: "{app}\Agro_DSR_Pro.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Run]
Filename: "{app}\Agro_DSR_Pro.exe"; Description: "{cm:LaunchProgram,Agro DSR Professional}"; Flags: nowait postinstall skipifsilent
