; ═══════════════════════════════════════════════════════════════════════════════
;  OpenSheet.iss — Inno Setup 6 installer script
;  Builds a single OpenSheet-Setup.exe that installs on Windows 10/11 x64
; ═══════════════════════════════════════════════════════════════════════════════

#define AppName      "OpenSheet"
#define AppVersion   "1.0.0"
#define AppPublisher "OpenSheet"
#define AppURL       "https://opensheet.app"
#define AppExeName   "OpenSheet.exe"
#define AppIcon      "opensheet.ico"

[Setup]
AppId={{F4A2C8E1-3B5D-4F9A-8C2E-1D7B3A6F5E9C}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL={#AppURL}
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
AllowNoIcons=yes
LicenseFile=
InfoBeforeFile=
OutputDir=.
OutputBaseFilename=OpenSheet-Setup
SetupIconFile=
Compression=lzma2/ultra
SolidCompression=yes
WizardStyle=modern
WizardResizable=yes
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
MinVersion=10.0.17763
PrivilegesRequired=admin
UninstallDisplayIcon={app}\{#AppExeName}
UninstallDisplayName={#AppName} {#AppVersion}

; Modern wizard look
WizardImageFile=compiler:WizModernImage-IS.bmp
WizardSmallImageFile=compiler:WizModernSmallImage-IS.bmp

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon";     Description: "{cm:CreateDesktopIcon}";      GroupDescription: "{cm:AdditionalIcons}";
Name: "startmenuicon";   Description: "Create Start Menu shortcut";   GroupDescription: "{cm:AdditionalIcons}"; Flags: checkedonce
Name: "fileassoc_xlsx";  Description: "Associate .xlsx files with OpenSheet"; GroupDescription: "File Associations";
Name: "fileassoc_csv";   Description: "Associate .csv files with OpenSheet";  GroupDescription: "File Associations";

[Files]
; Main executable
Source: "OpenSheet.exe"; DestDir: "{app}"; Flags: ignoreversion

; Application DLLs
Source: "*.dll"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs

; Qt platform plugins
Source: "platforms\*"; DestDir: "{app}\platforms"; Flags: ignoreversion recursesubdirs
Source: "styles\*";    DestDir: "{app}\styles";    Flags: ignoreversion recursesubdirs createallsubdirs skipifsourcedoesntexist
Source: "sqldrivers\*"; DestDir: "{app}\sqldrivers"; Flags: ignoreversion recursesubdirs createallsubdirs skipifsourcedoesntexist
Source: "imageformats\*"; DestDir: "{app}\imageformats"; Flags: ignoreversion recursesubdirs createallsubdirs skipifsourcedoesntexist
Source: "iconengines\*";  DestDir: "{app}\iconengines";  Flags: ignoreversion recursesubdirs createallsubdirs skipifsourcedoesntexist

[Icons]
Name: "{group}\{#AppName}";            Filename: "{app}\{#AppExeName}"
Name: "{group}\Uninstall {#AppName}";  Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}";      Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Registry]
; .xlsx file association
Root: HKCR; Subkey: ".xlsx";                          ValueType: string; ValueData: "OpenSheet.Document"; Flags: uninsdeletevalue; Tasks: fileassoc_xlsx
Root: HKCR; Subkey: "OpenSheet.Document";              ValueType: string; ValueData: "OpenSheet Workbook"; Flags: uninsdeletekey; Tasks: fileassoc_xlsx
Root: HKCR; Subkey: "OpenSheet.Document\DefaultIcon"; ValueType: string; ValueData: "{app}\{#AppExeName},0"; Tasks: fileassoc_xlsx
Root: HKCR; Subkey: "OpenSheet.Document\shell\open\command"; ValueType: string; ValueData: """{app}\{#AppExeName}"" ""%1"""; Tasks: fileassoc_xlsx

; .csv file association
Root: HKCR; Subkey: ".csv";                          ValueType: string; ValueData: "OpenSheet.CSV"; Flags: uninsdeletevalue; Tasks: fileassoc_csv
Root: HKCR; Subkey: "OpenSheet.CSV";                 ValueType: string; ValueData: "CSV File"; Flags: uninsdeletekey; Tasks: fileassoc_csv
Root: HKCR; Subkey: "OpenSheet.CSV\DefaultIcon";     ValueType: string; ValueData: "{app}\{#AppExeName},0"; Tasks: fileassoc_csv
Root: HKCR; Subkey: "OpenSheet.CSV\shell\open\command"; ValueType: string; ValueData: """{app}\{#AppExeName}"" ""%1"""; Tasks: fileassoc_csv

; Add to "Open with" list for .xlsx
Root: HKCR; Subkey: ".xlsx\OpenWithProgids"; ValueType: string; ValueName: "OpenSheet.Document"; ValueData: ""; Flags: uninsdeletevalue

[Run]
Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(AppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Code]
// Kill running OpenSheet instances before install
function InitializeSetup(): Boolean;
var
  ResultCode: Integer;
begin
  Exec(ExpandConstant('taskkill'), '/f /im OpenSheet.exe', '',
       SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Result := True;
end;
