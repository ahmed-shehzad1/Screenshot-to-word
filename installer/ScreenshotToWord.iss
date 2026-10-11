; Build the app first with PyInstaller, then compile this file with Inno Setup 6.
#define MyAppName "Screenshot to Word"
#define MyAppVersion "0.1.0"
#define MyAppPublisher "Screenshot-to-Word contributors"
#define MyAppExeName "ScreenshotToWord.exe"

[Setup]
AppId={{C78F8BC9-6D41-4C7A-A6B1-4E8C4F8D7A31}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\Screenshot to Word
DefaultGroupName={#MyAppName}
OutputDir=..\release
OutputBaseFilename=ScreenshotToWord-Setup-{#MyAppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=lowest

[Files]
Source: "..\dist\ScreenshotToWord\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Screenshot to Word"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\Screenshot to Word"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional icons:"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch Screenshot to Word"; Flags: postinstall nowait skipifsilent
