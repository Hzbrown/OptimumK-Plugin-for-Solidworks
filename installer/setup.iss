; OptimumK SolidWorks Plugin – Inno Setup installer script
; Compile with: ISCC.exe installer\setup.iss
; Requires Inno Setup 6+ (https://jrsoftware.org/isdl.php)

#define AppName      "OptimumK SolidWorks Plugin"
#define AppVersion   "1.0.0"
#define AppPublisher "OptimumK"
#define AppExeName   "OptimumK_SolidWorks_Plugin.exe"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={commonpf64}\OptimumK
DefaultGroupName={#AppName}
OutputDir=.
OutputBaseFilename=OptimumK_Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64
SetupIconFile=source\icon.ico
UninstallDisplayIcon={app}\icon.ico

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; Main Python GUI executable
Source: "source\{#AppExeName}";        DestDir: "{app}"; Flags: ignoreversion
; C# SolidWorks subprocess tool (interop types embedded — no SW DLLs needed)
Source: "source\SuspensionTools.exe";  DestDir: "{app}"; Flags: ignoreversion
; SolidWorks part file used for markers (placed next to exe)
Source: "source\Marker.SLDPRT";        DestDir: "{app}"; Flags: ignoreversion
; Help file and images
Source: "source\help.htm";             DestDir: "{app}"; Flags: ignoreversion isreadme
Source: "source\help_files\*";         DestDir: "{app}\help_files"; Flags: ignoreversion
; Application icon
Source: "source\icon.ico";             DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#AppName}";       Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\icon.ico"
Name: "{group}\Uninstall";        Filename: "{uninstallexe}"
Name: "{commondesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\icon.ico"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"
