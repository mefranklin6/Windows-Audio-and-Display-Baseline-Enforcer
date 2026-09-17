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

#define AppName "Windows Audio and Display Baseline Enforcer Orchestrator"
#define AppExeName "Windows-Audio-and-Display-Baseline-Enforcer-Orchestrator.exe"
#define AppDataDir "{userappdata}\Windows Audio and Display Baseline Enforcer Orchestrator"
#define LegacyAppDataDir "{userappdata}\Windows Audio and Display Baseline Enforcer"
#define LegacyInstallDir "{localappdata}\Programs\Windows Audio and Display Baseline Enforcer"

[Setup]
AppId={{5A949713-E789-42B4-8A6A-F912B88920AE}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=mefranklin6
AppPublisherURL=https://github.com/mefranklin6/Windows-Audio-and-Display-Baseline-Enforcer
AppSupportURL=https://github.com/mefranklin6/Windows-Audio-and-Display-Baseline-Enforcer/issues
AppUpdatesURL=https://github.com/mefranklin6/Windows-Audio-and-Display-Baseline-Enforcer/releases
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
PrivilegesRequired=admin
UsePreviousAppDir=no
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir={#OutputDir}
OutputBaseFilename=Windows-Audio-and-Display-Baseline-Enforcer-Orchestrator-{#AppVersion}-Setup
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#AppExeName}
CloseApplications=yes
RestartApplications=no

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Excludes: "wade_orchestrator.exe"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#RepoRoot}\installer_scripts\*"; DestDir: "{#AppDataDir}\installer_scripts"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#RepoRoot}\utility_scripts\*"; DestDir: "{#AppDataDir}\utility_scripts"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#RepoRoot}\targets.txt.example"; DestDir: "{#AppDataDir}"; Flags: ignoreversion
Source: "{#LegacyAppDataDir}\settings.json"; DestDir: "{#AppDataDir}"; Flags: external skipifsourcedoesntexist onlyifdoesntexist
Source: "{#LegacyInstallDir}\targets.txt"; DestDir: "{#AppDataDir}"; Flags: external skipifsourcedoesntexist onlyifdoesntexist
Source: "{#LegacyInstallDir}\BGInfo\*"; DestDir: "{#AppDataDir}\BGInfo"; Flags: external skipifsourcedoesntexist onlyifdoesntexist recursesubdirs createallsubdirs
Source: "{#LegacyInstallDir}\logs\*"; DestDir: "{#AppDataDir}\logs"; Flags: external skipifsourcedoesntexist onlyifdoesntexist recursesubdirs createallsubdirs
Source: "{#RepoRoot}\targets.txt.example"; DestDir: "{#AppDataDir}"; DestName: "targets.txt"; Flags: onlyifdoesntexist

[Dirs]
Name: "{#AppDataDir}\BGInfo"

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExeName}"; WorkingDir: "{#AppDataDir}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; WorkingDir: "{#AppDataDir}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"

[Run]
Filename: "{app}\{#AppExeName}"; WorkingDir: "{#AppDataDir}"; Description: "Launch {#AppName}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}"
Type: filesandordirs; Name: "{#AppDataDir}"
Type: filesandordirs; Name: "{#LegacyAppDataDir}"
Type: filesandordirs; Name: "{#LegacyInstallDir}"
