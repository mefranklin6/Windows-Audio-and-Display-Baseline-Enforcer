#ifndef AppVersion
  #define AppVersion "1.0.0"
#endif
#ifndef SourceDir
  #error SourceDir must point to the Flutter Windows release directory.
#endif
#ifndef RepoRoot
  #error RepoRoot must point to the repository root.
#endif
#ifndef OutputDir
  #define OutputDir "."
#endif

#define AppName "Windows Audio and Display Baseline Enforcer"
#define AppExeName "deployment_orchestrator_app.exe"

[Setup]
AppId={{5A949713-E789-42B4-8A6A-F912B88920AE}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=mefranklin6
AppPublisherURL=https://github.com/mefranklin6/Windows-Audio-and-Display-Baseline-Enforcer
AppSupportURL=https://github.com/mefranklin6/Windows-Audio-and-Display-Baseline-Enforcer/issues
AppUpdatesURL=https://github.com/mefranklin6/Windows-Audio-and-Display-Baseline-Enforcer/releases
DefaultDirName={localappdata}\Programs\Windows Audio and Display Baseline Enforcer
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir={#OutputDir}
OutputBaseFilename=Windows-Audio-and-Display-Baseline-Enforcer-{#AppVersion}-Setup
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#AppExeName}
CloseApplications=yes
RestartApplications=no

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#RepoRoot}\installer_scripts\*"; DestDir: "{app}\installer_scripts"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#RepoRoot}\utility_scripts\*"; DestDir: "{app}\utility_scripts"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#RepoRoot}\targets.txt.example"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#RepoRoot}\targets.txt.example"; DestDir: "{app}"; DestName: "targets.txt"; Flags: onlyifdoesntexist

[Dirs]
Name: "{app}\BGInfo"

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"

[Run]
Filename: "{app}\{#AppExeName}"; Description: "Launch {#AppName}"; Flags: nowait postinstall skipifsilent
