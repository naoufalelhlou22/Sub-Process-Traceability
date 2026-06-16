[Setup]
; App Information
AppName=Sub-Process Traceability System
AppVersion=1.0
AppPublisher=Your Company Name
DefaultDirName={autopf}\Sub-Process Traceability
DefaultGroupName=Sub-Process Traceability
; The setup file it will generate
OutputDir=installer
OutputBaseFilename=TraceabilitySystem_Setup
SetupIconFile=assets\icon.ico
Compression=lzma
SolidCompression=yes
; Require admin rights to install to Program Files
PrivilegesRequired=admin

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Copy the main executable
Source: "dist\Sub-Process Traceability.exe"; DestDir: "{app}"; Flags: ignoreversion
; Copy the quality app executable
Source: "dist\Quality App.exe"; DestDir: "{app}"; Flags: ignoreversion
; (Optional) Copy an assets folder if your portable exe needs external files
; Source: "assets\*"; DestDir: "{app}\assets"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; Create Start Menu Shortcuts
Name: "{group}\Sub-Process Traceability"; Filename: "{app}\Sub-Process Traceability.exe"
Name: "{group}\Quality App"; Filename: "{app}\Quality App.exe"
Name: "{group}\{cm:UninstallProgram,Sub-Process Traceability System}"; Filename: "{uninstallexe}"

; Create Desktop Shortcuts
Name: "{autodesktop}\Sub-Process Traceability"; Filename: "{app}\Sub-Process Traceability.exe"; Tasks: desktopicon
Name: "{autodesktop}\Quality App"; Filename: "{app}\Quality App.exe"; Tasks: desktopicon

[Run]
; Option to launch the app immediately after installation
Filename: "{app}\Sub-Process Traceability.exe"; Description: "{cm:LaunchProgram,Sub-Process Traceability System}"; Flags: nowait postinstall skipifsilent
