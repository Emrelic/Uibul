; UIBUL - Universal UI Element Inspector
; Inno Setup Script
; Bu script ile setup.exe olusturmak icin:
; 1. Inno Setup'i indirin: https://jrsoftware.org/isdl.php
; 2. Bu .iss dosyasini Inno Setup Compiler ile acin
; 3. Compile (Ctrl+F9) butonuna basin
; 4. Output klasorunde UIBUL_Setup.exe olusacaktir

#define MyAppName "UIBUL - UI Element Inspector"
#define MyAppVersion "3.0"
#define MyAppPublisher "UIBUL"
#define MyAppURL "https://github.com/uibul"
#define MyAppExeName "UIElementInspector.exe"
#define MyAppPublishDir "..\UIElementInspector\publish"

[Setup]
AppId={{E7A3B2C1-4D5F-6E78-9A0B-C1D2E3F4A5B6}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
DefaultDirName={autopf}\UIBUL
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
; Output dosyasi ayarlari
OutputDir=.\Output
OutputBaseFilename=UIBUL_v3_Setup
; Sıkıştırma
Compression=lzma2/ultra64
SolidCompression=yes
; Görünüm
WizardStyle=modern
SetupIconFile={#MyAppPublishDir}\Resources\app.ico
UninstallDisplayIcon={app}\Resources\app.ico
; Yetki
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
; Platform
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
; Bilgi
LicenseFile=
InfoBeforeFile=
InfoAfterFile=

[Languages]
Name: "turkish"; MessagesFile: "compiler:Languages\Turkish.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 6.1; Check: not IsAdminInstallMode

[Files]
; Tum publish klasorunu kopyala (self-contained, .NET runtime dahil)
Source: "{#MyAppPublishDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\Resources\app.ico"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\Resources\app.ico"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\Resources\app.ico"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}\Archive"
Type: filesandordirs; Name: "{app}\Logs"
