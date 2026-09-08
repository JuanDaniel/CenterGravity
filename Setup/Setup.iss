; Inno Setup script for Center Gravity.
; Run build.ps1 first to produce Setup\Output\CenterGravity.bundle, then compile this file.

#define MyAppName "Center Gravity"
#define MyAppVersion "1.2.0"
#define MyAppPublisher "Juan Daniel SANTANA"
#define MyAppURL "https://github.com/JuanDaniel/CenterGravity"

[Setup]
; AppId uniquely identifies this application - do not reuse it for other products.
AppId={{48BD08DC-F6B7-400B-85CE-70071034E058}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
CreateAppDir=no
DisableProgramGroupPage=yes
OutputBaseFilename=CenterGravity-{#MyAppVersion}-Setup
SetupIconFile=.\icon.ico
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={uninstallexe}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; Deploy the whole bundle. Revit loads the folder that matches each installed release.
Source: ".\Output\CenterGravity.bundle\*"; \
    DestDir: "{commonappdata}\Autodesk\ApplicationPlugins\CenterGravity.bundle"; \
    Flags: recursesubdirs createallsubdirs ignoreversion

[UninstallDelete]
Type: filesandordirs; Name: "{commonappdata}\Autodesk\ApplicationPlugins\CenterGravity.bundle"
