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
; C# SolidWorks subprocess tool
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

[Code]
// ─────────────────────────────────────────────────────────────────────────────
// Wizard page variables
// ─────────────────────────────────────────────────────────────────────────────
var
  SwPage:       TWizardPage;
  SwPathEdit:   TEdit;
  SwBrowseBtn:  TButton;
  DetectedSwPath: String;

// ─────────────────────────────────────────────────────────────────────────────
// Registry helpers to detect the SolidWorks installation directory
// ─────────────────────────────────────────────────────────────────────────────
function DetectSolidWorksPath(): String;
var
  InstallDir: String;
begin
  Result := '';

  // Preferred key (modern SW installs)
  if RegQueryStringValue(HKLM,
      'SOFTWARE\SolidWorks\Applications\SolidWorks',
      'InstallDir', InstallDir) then
  begin
    Result := InstallDir;
    Exit;
  end;

  // Alternate key
  if RegQueryStringValue(HKLM,
      'SOFTWARE\SolidWorks\SOLIDWORKS',
      'InstallPath', InstallDir) then
  begin
    Result := InstallDir;
    Exit;
  end;

  // 32-bit registry view fallback
  if RegQueryStringValue(HKLM32,
      'SOFTWARE\SolidWorks\Applications\SolidWorks',
      'InstallDir', InstallDir) then
  begin
    Result := InstallDir;
  end;
end;

// Strip trailing backslash so path arithmetic is consistent
function StripSlash(const S: String): String;
begin
  Result := S;
  if (Length(Result) > 0) and (Result[Length(Result)] = '\') then
    Delete(Result, Length(Result), 1);
end;

// Return the api\redist subfolder of a SolidWorks install dir
function ApiRedistPath(const SwDir: String): String;
begin
  Result := StripSlash(SwDir) + '\api\redist';
end;

// Check whether the chosen path contains the two required interop DLLs
function PathIsValid(const ApiPath: String): Boolean;
begin
  Result := FileExists(ApiPath + '\SolidWorks.Interop.sldworks.dll') and
            FileExists(ApiPath + '\SolidWorks.Interop.swconst.dll');
end;

// ─────────────────────────────────────────────────────────────────────────────
// Browse button click – open a folder picker
// ─────────────────────────────────────────────────────────────────────────────
procedure BrowseBtnClick(Sender: TObject);
var
  Folder: String;
begin
  Folder := SwPathEdit.Text;
  if BrowseForFolder('Select your SolidWorks installation folder:', Folder, False) then
    SwPathEdit.Text := Folder;
end;

// ─────────────────────────────────────────────────────────────────────────────
// Create the custom SolidWorks path wizard page
// ─────────────────────────────────────────────────────────────────────────────
procedure InitializeWizard();
var
  LblTitle, LblDesc, LblPath: TLabel;
begin
  DetectedSwPath := DetectSolidWorksPath();

  // Insert after the "Select Destination Location" page
  SwPage := CreateCustomPage(wpSelectDir,
    'SolidWorks Installation',
    'Specify where SolidWorks is installed on this computer.');

  LblTitle := TLabel.Create(SwPage);
  LblTitle.Parent  := SwPage.Surface;
  LblTitle.Caption := 'SolidWorks installation folder:';
  LblTitle.Left    := 0;
  LblTitle.Top     := 8;
  LblTitle.Width   := SwPage.SurfaceWidth;

  LblDesc := TLabel.Create(SwPage);
  LblDesc.Parent  := SwPage.Surface;
  LblDesc.Caption :=
    'The installer will copy SolidWorks interop DLLs from this location.' + #13#10 +
    'The folder should contain an "api\redist" subfolder.';
  LblDesc.Left    := 0;
  LblDesc.Top     := 24;
  LblDesc.Width   := SwPage.SurfaceWidth;
  LblDesc.AutoSize := True;
  LblDesc.WordWrap := True;

  SwPathEdit := TEdit.Create(SwPage);
  SwPathEdit.Parent := SwPage.Surface;
  SwPathEdit.Left   := 0;
  SwPathEdit.Top    := 72;
  SwPathEdit.Width  := SwPage.SurfaceWidth - 90;

  if DetectedSwPath <> '' then
    SwPathEdit.Text := DetectedSwPath
  else
    SwPathEdit.Text := 'C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS';

  SwBrowseBtn := TButton.Create(SwPage);
  SwBrowseBtn.Parent  := SwPage.Surface;
  SwBrowseBtn.Caption := 'Browse...';
  SwBrowseBtn.Left    := SwPathEdit.Left + SwPathEdit.Width + 8;
  SwBrowseBtn.Top     := SwPathEdit.Top - 2;
  SwBrowseBtn.Width   := 80;
  SwBrowseBtn.OnClick := @BrowseBtnClick;

  LblPath := TLabel.Create(SwPage);
  LblPath.Parent  := SwPage.Surface;
  LblPath.Caption := '(e.g. C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS)';
  LblPath.Left    := 0;
  LblPath.Top     := SwPathEdit.Top + SwPathEdit.Height + 6;
  LblPath.Width   := SwPage.SurfaceWidth;
end;

// ─────────────────────────────────────────────────────────────────────────────
// Validate the page before allowing Next
// ─────────────────────────────────────────────────────────────────────────────
function NextButtonClick(CurPageID: Integer): Boolean;
var
  ApiPath: String;
begin
  Result := True;

  if CurPageID = SwPage.ID then
  begin
    ApiPath := ApiRedistPath(SwPathEdit.Text);
    if not PathIsValid(ApiPath) then
    begin
      MsgBox(
        'SolidWorks interop DLLs were not found at:' + #13#10 +
        '  ' + ApiPath + #13#10#13#10 +
        'Please choose the folder that contains your SolidWorks installation ' +
        '(the one with "api\redist" inside it).',
        mbError, MB_OK);
      Result := False;
    end;
  end;
end;

// ─────────────────────────────────────────────────────────────────────────────
// After the main files are installed, copy the two interop DLLs
// ─────────────────────────────────────────────────────────────────────────────
procedure CurStepChanged(CurStep: TSetupStep);
var
  ApiPath, DestPath: String;
begin
  if CurStep = ssPostInstall then
  begin
    ApiPath  := ApiRedistPath(SwPathEdit.Text);
    DestPath := ExpandConstant('{app}');

    if not FileCopy(ApiPath + '\SolidWorks.Interop.sldworks.dll',
                    DestPath + '\SolidWorks.Interop.sldworks.dll', False) then
      MsgBox('Warning: Could not copy SolidWorks.Interop.sldworks.dll', mbError, MB_OK);

    if not FileCopy(ApiPath + '\SolidWorks.Interop.swconst.dll',
                    DestPath + '\SolidWorks.Interop.swconst.dll', False) then
      MsgBox('Warning: Could not copy SolidWorks.Interop.swconst.dll', mbError, MB_OK);
  end;
end;
