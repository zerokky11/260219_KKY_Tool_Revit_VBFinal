; KKY Family Browser recovered host installer.

#ifndef MyAppName
  #define MyAppName "KKY Family Browser RevitHost"
#endif
#ifndef MyAppVersion
  #define MyAppVersion "1.0.1"
#endif
#ifndef MyOutputBaseName
  #define MyOutputBaseName "KKY_FamilyBrowser_RevitHost(2019,21,23,25,27)_v" + MyAppVersion + "_Setup"
#endif
#ifndef MyOutputDir
  #define MyOutputDir "..\artifacts\family-browser\installers"
#endif
#ifndef MyStageRoot
  #define MyStageRoot "..\artifacts\family-browser\stage"
#endif

#define MyAppPublisher "Kyeongyeon Kim"
#define MyAppURL "kkykiki89@nate.com"

[Setup]
AppId={{1D5C814F-F8F3-4DF5-91C8-5B61F4DB0FB7}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
UninstallDisplayName={#MyAppName}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}

DefaultDirName=C:\ProgramData\Autodesk\Revit\Addins\2025
DisableDirPage=yes
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=admin
UseSetupLdr=no
CloseApplications=yes
RestartApplications=no

OutputDir={#MyOutputDir}
OutputBaseFilename={#MyOutputBaseName}
SetupIconFile=..\Compile\KKY_Tool_Revit_Installer.ico
SolidCompression=yes
WizardStyle=modern dark windows11

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "korean";  MessagesFile: "compiler:Languages\Korean.isl"

[Tasks]
Name: "revit2019"; Description: "Revit 2019 install"; GroupDescription: "Select Revit versions to install:"
Name: "revit2021"; Description: "Revit 2021 install"; GroupDescription: "Select Revit versions to install:"
Name: "revit2023"; Description: "Revit 2023 install"; GroupDescription: "Select Revit versions to install:"
Name: "revit2025"; Description: "Revit 2025 install"; GroupDescription: "Select Revit versions to install:"
Name: "revit2027"; Description: "Revit 2027 install"; GroupDescription: "Select Revit versions to install:"

[InstallDelete]
Type: filesandordirs; Name: "{commonappdata}\Autodesk\Revit\Addins\2019\KKY_FamilyBrowser"; Tasks: revit2019
Type: files; Name: "{commonappdata}\Autodesk\Revit\Addins\2019\KKY_FamilyBrowser_RevitHost.addin"; Tasks: revit2019
Type: files; Name: "{commonappdata}\Autodesk\Revit\Addins\2019\KKY_FamilyBrowser_RevitHost_2025.addin"; Tasks: revit2019
Type: files; Name: "{commonappdata}\Autodesk\Revit\Addins\2019\KKY_FamilyBrowser_RevitHost_2027.addin"; Tasks: revit2019
Type: filesandordirs; Name: "{commonappdata}\Autodesk\Revit\Addins\2021\KKY_FamilyBrowser"; Tasks: revit2021
Type: files; Name: "{commonappdata}\Autodesk\Revit\Addins\2021\KKY_FamilyBrowser_RevitHost.addin"; Tasks: revit2021
Type: files; Name: "{commonappdata}\Autodesk\Revit\Addins\2021\KKY_FamilyBrowser_RevitHost_2025.addin"; Tasks: revit2021
Type: files; Name: "{commonappdata}\Autodesk\Revit\Addins\2021\KKY_FamilyBrowser_RevitHost_2027.addin"; Tasks: revit2021
Type: filesandordirs; Name: "{commonappdata}\Autodesk\Revit\Addins\2023\KKY_FamilyBrowser"; Tasks: revit2023
Type: files; Name: "{commonappdata}\Autodesk\Revit\Addins\2023\KKY_FamilyBrowser_RevitHost.addin"; Tasks: revit2023
Type: files; Name: "{commonappdata}\Autodesk\Revit\Addins\2023\KKY_FamilyBrowser_RevitHost_2025.addin"; Tasks: revit2023
Type: files; Name: "{commonappdata}\Autodesk\Revit\Addins\2023\KKY_FamilyBrowser_RevitHost_2027.addin"; Tasks: revit2023
Type: filesandordirs; Name: "{commonappdata}\Autodesk\Revit\Addins\2025\KKY_FamilyBrowser"; Tasks: revit2025
Type: files; Name: "{commonappdata}\Autodesk\Revit\Addins\2025\KKY_FamilyBrowser_RevitHost.addin"; Tasks: revit2025
Type: files; Name: "{commonappdata}\Autodesk\Revit\Addins\2025\KKY_FamilyBrowser_RevitHost_2025.addin"; Tasks: revit2025
Type: files; Name: "{commonappdata}\Autodesk\Revit\Addins\2025\KKY_FamilyBrowser_RevitHost_2027.addin"; Tasks: revit2025
Type: filesandordirs; Name: "{commonappdata}\Autodesk\Revit\Addins\2027\KKY_FamilyBrowser"; Tasks: revit2027
Type: files; Name: "{commonappdata}\Autodesk\Revit\Addins\2027\KKY_FamilyBrowser_RevitHost.addin"; Tasks: revit2027
Type: files; Name: "{commonappdata}\Autodesk\Revit\Addins\2027\KKY_FamilyBrowser_RevitHost_2025.addin"; Tasks: revit2027
Type: files; Name: "{commonappdata}\Autodesk\Revit\Addins\2027\KKY_FamilyBrowser_RevitHost_2027.addin"; Tasks: revit2027

[Files]
Source: "{#MyStageRoot}\Rvt2019\KKY_FamilyBrowser_RevitHost.addin"; \
    DestDir: "{commonappdata}\Autodesk\Revit\Addins\2019"; \
    Flags: ignoreversion; \
    Tasks: revit2019
Source: "{#MyStageRoot}\Rvt2019\KKY_FamilyBrowser\*"; \
    DestDir: "{commonappdata}\Autodesk\Revit\Addins\2019\KKY_FamilyBrowser"; \
    Flags: ignoreversion recursesubdirs createallsubdirs; \
    Tasks: revit2019

Source: "{#MyStageRoot}\Rvt2021\KKY_FamilyBrowser_RevitHost.addin"; \
    DestDir: "{commonappdata}\Autodesk\Revit\Addins\2021"; \
    Flags: ignoreversion; \
    Tasks: revit2021
Source: "{#MyStageRoot}\Rvt2021\KKY_FamilyBrowser\*"; \
    DestDir: "{commonappdata}\Autodesk\Revit\Addins\2021\KKY_FamilyBrowser"; \
    Flags: ignoreversion recursesubdirs createallsubdirs; \
    Tasks: revit2021

Source: "{#MyStageRoot}\Rvt2023\KKY_FamilyBrowser_RevitHost.addin"; \
    DestDir: "{commonappdata}\Autodesk\Revit\Addins\2023"; \
    Flags: ignoreversion; \
    Tasks: revit2023
Source: "{#MyStageRoot}\Rvt2023\KKY_FamilyBrowser\*"; \
    DestDir: "{commonappdata}\Autodesk\Revit\Addins\2023\KKY_FamilyBrowser"; \
    Flags: ignoreversion recursesubdirs createallsubdirs; \
    Tasks: revit2023

Source: "{#MyStageRoot}\Rvt2025\KKY_FamilyBrowser_RevitHost_2025.addin"; \
    DestDir: "{commonappdata}\Autodesk\Revit\Addins\2025"; \
    Flags: ignoreversion; \
    Tasks: revit2025
Source: "{#MyStageRoot}\Rvt2025\KKY_FamilyBrowser\*"; \
    DestDir: "{commonappdata}\Autodesk\Revit\Addins\2025\KKY_FamilyBrowser"; \
    Flags: ignoreversion recursesubdirs createallsubdirs; \
    Tasks: revit2025

Source: "{#MyStageRoot}\Rvt2027\KKY_FamilyBrowser_RevitHost_2027.addin"; \
    DestDir: "{commonappdata}\Autodesk\Revit\Addins\2027"; \
    Flags: ignoreversion; \
    Tasks: revit2027
Source: "{#MyStageRoot}\Rvt2027\KKY_FamilyBrowser\*"; \
    DestDir: "{commonappdata}\Autodesk\Revit\Addins\2027\KKY_FamilyBrowser"; \
    Flags: ignoreversion recursesubdirs createallsubdirs; \
    Tasks: revit2027

[Code]
function IsRevitInstalled(Version: string): Boolean;
var
  AddinPath, RevitPath64, RevitPath32: string;
  RegistryPath: string;
begin
  AddinPath := ExpandConstant('{commonappdata}\Autodesk\Revit\Addins\' + Version);
  RevitPath64 := ExpandConstant('{commonpf}\Autodesk\Revit ' + Version);
  RevitPath32 := ExpandConstant('{commonpf32}\Autodesk\Revit ' + Version);
  RegistryPath := 'SOFTWARE\Autodesk\Revit\' + Version;

  Result :=
    DirExists(AddinPath) or
    DirExists(RevitPath64) or
    DirExists(RevitPath32) or
    RegKeyExists(HKLM, RegistryPath) or
    RegKeyExists(HKCU, RegistryPath);
end;

function MakeCaption(Year: string; Installed: Boolean): string;
begin
  if Installed then
    Result := Format('Revit %s install (detected)', [Year])
  else
    Result := Format('Revit %s install (manual selection)', [Year]);
end;

procedure UpdateTask(CaptionToken: string; Year: string; Installed: Boolean);
var
  i: Integer;
  cap: string;
begin
  for i := 0 to WizardForm.TasksList.Items.Count - 1 do
  begin
    cap := WizardForm.TasksList.ItemCaption[i];
    if Pos(CaptionToken, cap) > 0 then
    begin
      WizardForm.TasksList.ItemCaption[i] := MakeCaption(Year, Installed);
      WizardForm.TasksList.Checked[i] := Installed;
      WizardForm.TasksList.ItemEnabled[i] := True;
    end;
  end;
end;

procedure CurPageChanged(CurPageID: Integer);
begin
  if CurPageID <> wpSelectTasks then
    Exit;

  UpdateTask('Revit 2019', '2019', IsRevitInstalled('2019'));
  UpdateTask('Revit 2021', '2021', IsRevitInstalled('2021'));
  UpdateTask('Revit 2023', '2023', IsRevitInstalled('2023'));
  UpdateTask('Revit 2025', '2025', IsRevitInstalled('2025'));
  UpdateTask('Revit 2027', '2027', IsRevitInstalled('2027'));
end;
