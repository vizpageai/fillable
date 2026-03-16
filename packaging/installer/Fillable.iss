; FillableDOC Inno Setup installer for Microsoft Store (MSI/EXE) submission path.
; Build:
;   1) Build dist\FillableDOC.exe first.
;   2) Compile this file with ISCC.exe.

#define MyAppName "FillableDOC"
#define MyAppVersion "2026.02.27.1"
#define MyAppPublisher "VizpageAI"
#define MyAppExeName "FillableDOC.exe"
#ifndef StoreCapture
#define StoreCapture "0"
#endif

[Setup]
AppId={{E4A2E95F-155D-4F4F-9E6D-C5A4A634A7A8}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DisableProgramGroupPage=yes
OutputDir=..\..\dist\installer
OutputBaseFilename=FillableDOC-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=lowest

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

#if StoreCapture == "0"
[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
#endif

[Files]
Source: "..\..\dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
#if StoreCapture == "0"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
#endif

[Run]
Filename: "{app}\{#MyAppExeName}"; Parameters: "--install-context-menu"; Flags: runhidden

[UninstallRun]
Filename: "{app}\{#MyAppExeName}"; Parameters: "--uninstall-context-menu"; Flags: runhidden
