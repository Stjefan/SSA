; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

[Setup]
AppId={{5F0F9DD6-79C2-4007-BD38-53257856C226}}
#define MyDateTimeString GetDateTimeString('yyyymmdd', '', '');
AppName=7Label 2018
AppVerName=7Label V20180824
AppPublisher=Stratenschulte Messtechnik
AppPublisherURL=http://www.stratenschulte-messtechnik.de/
AppSupportURL=http://www.stratenschulte-messtechnik.de/
AppUpdatesURL=http://www.stratenschulte-messtechnik.de/
DefaultDirName={pf}\7Label 2018
DefaultGroupName=7Label 2018
AllowNoIcons=yes
LicenseFile=\\kuf-nas-01\Auftrag\10\105\1050\10500\zo Bearbeitung\SSA_20171219\Lizenzbestimmungen.txt
OutputDir=\\kuf-nas-01\Auftrag\10\105\1050\10500\zo Bearbeitung\SSA_20171219
OutputBaseFilename=setup_Schallschutzausweis_7Label_V{#MyDateTimeString} 
;20141020
Password=k34GvGTFV42jr
SetupIconFile=\\kuf-nas-01\Auftrag\10\105\1050\10500\zo Bearbeitung\SSA_20171219\SSA_Logo.ico
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin

[Languages]
Name: "german"; MessagesFile: "compiler:Languages\German.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "\\kuf-nas-01\Auftrag\10\105\1050\10500\zo Bearbeitung\SSA_20171219\SSA_20171219\bin\Debug\SSA_20171219.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "\\kuf-nas-01\Auftrag\10\105\1050\10500\zo Bearbeitung\SSA_20171219\SSA_20171219\bin\Debug\DEGA_Empfehlung_103.pdf"; DestDir: "{app}"; Flags: ignoreversion
Source: "\\kuf-nas-01\Auftrag\10\105\1050\10500\zo Bearbeitung\SSA_20171219\SSA_20171219\bin\Debug\Handbuch.pdf"; DestDir: "{app}"; Flags: ignoreversion
Source: "\\kuf-nas-01\Auftrag\10\105\1050\10500\zo Bearbeitung\SSA_20171219\Interop.Excel.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "\\kuf-nas-01\Auftrag\10\105\1050\10500\zo Bearbeitung\SSA_20171219\Interop.Office.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "\\kuf-nas-01\Auftrag\10\105\1050\10500\zo Bearbeitung\SSA_20171219\Interop.VBIDE.dll"; DestDir: "{app}"; Flags: ignoreversion
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: "{group}\7Label"; Filename: "{app}\SSA_20171219.exe"
Name: "{group}\{cm:ProgramOnTheWeb,7Label}"; Filename: "http://www.stratenschulte-messtechnik.de/"
Name: "{group}\{cm:UninstallProgram,7Label}"; Filename: "{uninstallexe}"
Name: "{commondesktop}\7Label"; Filename: "{app}\SSA_20171219.exe"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\7Label"; Filename: "{app}\SSA_20171219.exe"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\SSA_20171219.exe"; Description: "{cm:LaunchProgram,7Label}"; Flags: nowait postinstall skipifsilent
