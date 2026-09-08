#define MyAppName "SiPP"
#ifndef MyAppVersion
  #define MyAppVersion "2.3.0"
#endif

[Setup]
AppId={{B3B6E1C8-6D4A-4A7B-9F22-5F9E3B7D1C40}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher=SiPP
AppPublisherURL=https://github.com/dsfenton14-svg/Sipp-Aplicacion
AppSupportURL=https://github.com/dsfenton14-svg/Sipp-Aplicacion
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=admin
OutputDir=..\dist\installer
OutputBaseFilename=SiPP-Setup-v{#MyAppVersion}
SetupIconFile=..\imagenes\icono1.ico
UninstallDisplayIcon={app}\SiPP.exe
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Crear un acceso directo en el escritorio"; GroupDescription: "Accesos directos:"

[Files]
Source: "..\dist\SiPP\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\SiPP"; Filename: "{app}\SiPP.exe"; WorkingDir: "{app}"
Name: "{autodesktop}\SiPP"; Filename: "{app}\SiPP.exe"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\SiPP.exe"; Description: "Iniciar SiPP"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}"
