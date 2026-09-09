#define MyAppName "Word to PDF"
#define MyAppVersion GetEnv("RELEASE_VERSION")
#if MyAppVersion == ""
  #define MyAppVersion "1.0.0"
#endif
#define MyAppPublisher "Word to PDF"
#define MyAppExeName "WordToPDF.exe"

[Setup]
AppId={{8C5C0D4C-8B76-4D47-A4CF-6D0C7A7B91C8}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\Word to PDF
DefaultGroupName=Word to PDF
OutputDir=installer
OutputBaseFilename=WordToPDF-Setup-v{#MyAppVersion}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=admin
UninstallDisplayName=Word to PDF
Uninstallable=yes
SetupIconFile=

[Files]
Source: "publish\WordToPDF.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\Word to PDF"; Filename: "{app}\WordToPDF.exe"
Name: "{autodesktop}\Word to PDF"; Filename: "{app}\WordToPDF.exe"

[Run]
Filename: "{app}\WordToPDF.exe"; Description: "Launch Word to PDF"; Flags: nowait postinstall skipifsilent
