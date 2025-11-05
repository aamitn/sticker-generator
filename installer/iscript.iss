; Sticker Generator Installer Script

[Setup]
AppName=Sticker Generator
AppVersion=1.0
DefaultDirName={pf}\Sticker Generator
DefaultGroupName=Sticker Generator
UninstallDisplayIcon={app}\app.exe
OutputDir=output
OutputBaseFilename=StickerGeneratorSetup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64

; Use the main app icon (from project root)
SetupIconFile="..\icon.ico"

[Files]
; main app executable
Source: "..\dist\app.exe"; DestDir: "{app}"; Flags: ignoreversion

; sticker image (copied to app dir)
Source: "..\sticker.png"; DestDir: "{app}"; Flags: ignoreversion

; app icon for shortcuts
Source: "..\icon.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Start Menu shortcut
Name: "{group}\Sticker Generator"; Filename: "{app}\app.exe"; IconFilename: "{app}\icon.ico"

; Desktop shortcut
Name: "{commondesktop}\Sticker Generator"; Filename: "{app}\app.exe"; IconFilename: "{app}\icon.ico"; Tasks: desktopicon

[Tasks]
; Optional desktop shortcut
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Run]
; Option to launch app after installation
Filename: "{app}\app.exe"; Description: "Launch Sticker Generator"; Flags: nowait postinstall skipifsilent shellexec
