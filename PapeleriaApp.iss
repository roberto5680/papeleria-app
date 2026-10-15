#define MyAppName "PapeleriaApp"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "PapeleriaApp"
#define MyAppExeName "PapeleriaApp.exe"

[Setup]
AppId={{A8F2C6B1-5D73-4A91-9E42-7C8B1F4D6320}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\PapeleriaApp
DefaultGroupName={#MyAppName}
OutputDir=.\instalador
OutputBaseFilename=PapeleriaApp_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin

[Files]
Source: ".\dist\PapeleriaApp\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

Source: "C:\abarrotes_app\DATOS_PAPELERIA.db"; DestDir: "C:\abarrotes_app"; Flags: ignoreversion

[Dirs]
Name: "{app}\static\uploads"; Permissions: users-modify
Name: "{app}\respaldos"; Permissions: users-modify
Name: "{app}\tickets_generados"; Permissions: users-modify
Name: "C:\abarrotes_app"; Permissions: users-modify

[Icons]
Name: "{autodesktop}\PapeleriaApp"; Filename: "{app}\PapeleriaApp.exe"
Name: "{group}\PapeleriaApp"; Filename: "{app}\PapeleriaApp.exe"

[Run]
Filename: "{app}\PapeleriaApp.exe"; Description: "Iniciar PapeleriaApp"; Flags: nowait postinstall skipifsilent

