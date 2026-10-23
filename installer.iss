; DocPreview installer script for Inno Setup 6.
; Build with:  ISCC.exe installer.iss
; Expects dist\DocPreview.exe to exist (run build.bat first).

#define AppName        "DocPreview"
; AppVersion can be overridden from the command line: ISCC /DAppVersion=1.2.3
#ifndef AppVersion
  #define AppVersion     "1.0.0"
#endif
#define AppPublisher   "Church Zheng"
#define AppURL         "https://github.com/xsnj-9527/rea-word"
#define AppExeName     "DocPreview.exe"

[Setup]
AppId={{8E4C1B7A-2F3D-4A6E-9C51-7B2E8D4A1F03}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}/issues
AppUpdatesURL={#AppURL}/releases
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
LicenseFile=LICENSE
OutputDir=installer_out
OutputBaseFilename=DocPreview-{#AppVersion}-setup
SetupIconFile=src\icon.ico
UninstallDisplayIcon={app}\{#AppExeName}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
; Per-user install by default: no admin prompt, and "Open with" entries land
; in HKCU where the running user can actually read them.
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog

[Languages]
Name: "chinesesimplified"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "associate"; Description: "把 DocPreview 加入 .docx 的右键「打开方式」（不修改默认程序）"; GroupDescription: "集成"; Flags: checkedonce

[Files]
Source: "dist\{#AppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "LICENSE"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{group}\卸载 {#AppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Registry]
; Only the "Open with" list is touched; the default .docx handler is left
; exactly as it was. Every value written here is removed on uninstall.
Root: HKCU; Subkey: "Software\Classes\{#AppName}.docx"; ValueType: string; \
    ValueName: ""; ValueData: "DocPreview 只读预览"; Flags: uninsdeletekey; Tasks: associate
Root: HKCU; Subkey: "Software\Classes\{#AppName}.docx\DefaultIcon"; ValueType: string; \
    ValueName: ""; ValueData: "{app}\{#AppExeName},0"; Tasks: associate
Root: HKCU; Subkey: "Software\Classes\{#AppName}.docx\shell\open"; ValueType: string; \
    ValueName: "FriendlyAppName"; ValueData: "{#AppName}"; Tasks: associate
Root: HKCU; Subkey: "Software\Classes\{#AppName}.docx\shell\open\command"; ValueType: string; \
    ValueName: ""; ValueData: """{app}\{#AppExeName}"" ""%1"""; Tasks: associate
Root: HKCU; Subkey: "Software\Classes\.docx\OpenWithProgids"; ValueType: none; \
    ValueName: "{#AppName}.docx"; Flags: uninsdeletevalue; Tasks: associate
Root: HKCU; Subkey: "Software\Classes\.docm\OpenWithProgids"; ValueType: none; \
    ValueName: "{#AppName}.docx"; Flags: uninsdeletevalue; Tasks: associate

[Run]
Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchProgram,{#AppName}}"; \
    Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Nothing outside {app} is deleted here: the registry entries above carry
; uninsdeletekey / uninsdeletevalue, so WPS/Word associations stay intact.
Type: filesandordirs; Name: "{app}"
