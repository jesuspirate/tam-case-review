; Inno Setup script: turns app\dist\TAM Case Review\ into a one-click installer.
; build_windows.bat runs this automatically when Inno Setup 6 is installed.
; Installs per user (no admin rights needed) with Start menu and optional desktop shortcuts.

#define AppName "TAM Case Review"
#define AppVersion "1.0"

[Setup]
AppId={{9CE7AD54-4C18-46DE-BB4E-97CF9AEB0105}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=TAM toolkit
DefaultDirName={localappdata}\Programs\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=dist
OutputBaseFilename={#AppName} Setup
SetupIconFile=assets\icon.ico
UninstallDisplayIcon={app}\{#AppName}.exe
Compression=lzma2
SolidCompression=yes
WizardStyle=modern

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Shortcuts:"

[Files]
Source: "dist\{#AppName}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppName}.exe"
Name: "{userdesktop}\{#AppName}"; Filename: "{app}\{#AppName}.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppName}.exe"; Description: "Open {#AppName} now"; Flags: nowait postinstall skipifsilent
