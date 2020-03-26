; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "kool-aide"
#define MyAppVersion "0.9.0"
#define MyAppPublisher "Rxs-labs"
#define MyAppURL "https://rsx-labs.github.io/"
#define MyAppExeName "kool-aide.exe"

[Setup]
; NOTE: The value of AppId uniquely identifies this application. Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{D98C3A87-8280-4892-8B8D-FBD013DDE7DD}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DisableDirPage=yes
DefaultGroupName={#MyAppName}
LicenseFile=C:\Dev\codes\kool-aide\LICENSE
InfoBeforeFile=C:\Dev\codes\kool-aide\docs\install_intro.rtf
InfoAfterFile=C:\Dev\codes\kool-aide\docs\install_exit.rtf
; Uncomment the following line to run in non administrative install mode (install for current user only.)
;PrivilegesRequired=lowest
OutputDir=C:\Temp\kool-aide-build
OutputBaseFilename=kool-aide_v0.9.1_public_release_setup
SetupIconFile=C:\Dev\codes\kool-aide\src\kool_aide\assets\images\kool-aide-install.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "C:\Dev\codes\kool-aide\build\kool-aide_v0.9.1_public\kool-aide.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Dev\codes\kool-aide\build\kool-aide_v0.9.1_public\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"

[Dirs]
Name: {app}; Permissions: users-full

