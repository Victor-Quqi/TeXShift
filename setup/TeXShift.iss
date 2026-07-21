#ifndef AppVersion
  #error AppVersion is required
#endif
#ifndef SourceDir
  #error SourceDir is required
#endif
#ifndef OutputDir
  #error OutputDir is required
#endif
#ifndef AssemblyVersion
  #error AssemblyVersion is required
#endif
#ifndef PublicKeyToken
  #error PublicKeyToken is required
#endif
#define AppName "TeXShift"
#define Publisher "Victor-Quqi"
#define ReleaseClsid "{{1EE8F914-ECBD-4709-92C0-E770851C4966}"
#define ReleaseProgId "TeXShift.AddIn.Connect"
#define AddInClass "TeXShift.AddIn.Connect"
#define AssemblyFullName "TeXShift.AddIn, Version=" + AssemblyVersion + ", Culture=neutral, PublicKeyToken=" + PublicKeyToken
#define LegacyUpgradeCode "{AC31220A-C96E-4A51-9AC5-647B9877502C}"

[Setup]
AppId={{861A99E6-1888-40DC-9593-3779EDACB74B}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#Publisher}
AppPublisherURL=https://github.com/Victor-Quqi/TeXShift
AppSupportURL=https://github.com/Victor-Quqi/TeXShift/issues
AppUpdatesURL=https://github.com/Victor-Quqi/TeXShift/releases
VersionInfoVersion={#AppVersion}.0
DefaultDirName={autopf}\TeXShift
DefaultGroupName=TeXShift
DisableProgramGroupPage=yes
DisableWelcomePage=no
PrivilegesRequired=admin
SetupArchitecture=x64
MinVersion=10.0
OutputDir={#OutputDir}
OutputBaseFilename=TeXShift-{#AppVersion}-x64-Setup
SetupIconFile={#SourcePath}TeXShiftSetup\Assets\T&MtoN_256.ico
LicenseFile={#SourcePath}..\LICENSE
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
RestartApplications=no
AppMutex=Local\TeXShift.AddIn.1EE8F914-ECBD-4709-92C0-E770851C4966
UninstallDisplayName=TeXShift
UninstallDisplayIcon={app}\TeXShift.ico

[Languages]
Name: "en"; MessagesFile: "compiler:Default.isl,Languages\TeXShift.en.isl"
Name: "zhcn"; MessagesFile: "compiler:Languages\ChineseSimplified.isl,Languages\TeXShift.zh-CN.isl"

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#SourcePath}TeXShiftSetup\Assets\T&MtoN_256.ico"; DestDir: "{app}"; DestName: "TeXShift.ico"; Flags: ignoreversion

[Registry]
Root: HKLM64; Subkey: "Software\Classes\CLSID\{{1EE8F914-ECBD-4709-92C0-E770851C4966}"; ValueType: string; ValueName: ""; ValueData: "{#AddInClass}"; Flags: uninsdeletekey
Root: HKLM64; Subkey: "Software\Classes\CLSID\{{1EE8F914-ECBD-4709-92C0-E770851C4966}"; ValueType: string; ValueName: "AppID"; ValueData: "{#ReleaseClsid}"
Root: HKLM64; Subkey: "Software\Classes\CLSID\{{1EE8F914-ECBD-4709-92C0-E770851C4966}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"
Root: HKLM64; Subkey: "Software\Classes\CLSID\{{1EE8F914-ECBD-4709-92C0-E770851C4966}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKLM64; Subkey: "Software\Classes\CLSID\{{1EE8F914-ECBD-4709-92C0-E770851C4966}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "{#AddInClass}"
Root: HKLM64; Subkey: "Software\Classes\CLSID\{{1EE8F914-ECBD-4709-92C0-E770851C4966}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyFullName}"
Root: HKLM64; Subkey: "Software\Classes\CLSID\{{1EE8F914-ECBD-4709-92C0-E770851C4966}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKLM64; Subkey: "Software\Classes\CLSID\{{1EE8F914-ECBD-4709-92C0-E770851C4966}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAddInCodeBase}"
Root: HKLM64; Subkey: "Software\Classes\CLSID\{{1EE8F914-ECBD-4709-92C0-E770851C4966}\InprocServer32\{#AssemblyVersion}"; ValueType: string; ValueName: "Class"; ValueData: "{#AddInClass}"
Root: HKLM64; Subkey: "Software\Classes\CLSID\{{1EE8F914-ECBD-4709-92C0-E770851C4966}\InprocServer32\{#AssemblyVersion}"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyFullName}"
Root: HKLM64; Subkey: "Software\Classes\CLSID\{{1EE8F914-ECBD-4709-92C0-E770851C4966}\InprocServer32\{#AssemblyVersion}"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKLM64; Subkey: "Software\Classes\CLSID\{{1EE8F914-ECBD-4709-92C0-E770851C4966}\InprocServer32\{#AssemblyVersion}"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAddInCodeBase}"
Root: HKLM64; Subkey: "Software\Classes\CLSID\{{1EE8F914-ECBD-4709-92C0-E770851C4966}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#ReleaseProgId}"
Root: HKLM64; Subkey: "Software\Classes\CLSID\{{1EE8F914-ECBD-4709-92C0-E770851C4966}\Implemented Categories\{{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"; ValueType: none; Flags: uninsdeletekeyifempty
Root: HKLM64; Subkey: "Software\Classes\{#ReleaseProgId}"; ValueType: string; ValueName: ""; ValueData: "{#AddInClass}"; Flags: uninsdeletekey
Root: HKLM64; Subkey: "Software\Classes\{#ReleaseProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#ReleaseClsid}"
Root: HKLM64; Subkey: "Software\Classes\AppID\{{1EE8F914-ECBD-4709-92C0-E770851C4966}"; ValueType: string; ValueName: "DllSurrogate"; ValueData: ""; Flags: uninsdeletekey
Root: HKLM64; Subkey: "Software\Microsoft\Office\OneNote\Addins\{#ReleaseProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "TeXShift"; Flags: uninsdeletekey
Root: HKLM64; Subkey: "Software\Microsoft\Office\OneNote\Addins\{#ReleaseProgId}"; ValueType: string; ValueName: "Description"; ValueData: "TeXShift Add-in"
Root: HKLM64; Subkey: "Software\Microsoft\Office\OneNote\Addins\{#ReleaseProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"
Root: HKLM64; Subkey: "Software\Microsoft\Office\OneNote\Addins\{#ReleaseProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"
Root: HKLM64; Subkey: "Software\TeXShift\Installer"; ValueType: none; Flags: uninsdeletekey
Root: HKLM64; Subkey: "Software\TeXShift\Installer"; ValueType: string; ValueName: "OwnerSid"; ValueData: "{code:GetCapturedOwnerSid}"
Root: HKLM64; Subkey: "Software\TeXShift\Installer"; ValueType: string; ValueName: "AppDataRoot"; ValueData: "{code:GetCapturedAppDataRoot}"
Root: HKLM64; Subkey: "Software\TeXShift\Installer"; ValueType: string; ValueName: "LocalAppDataRoot"; ValueData: "{code:GetCapturedLocalAppDataRoot}"
Root: HKLM64; Subkey: "Software\TeXShift\Installer"; ValueType: string; ValueName: "TempRoot"; ValueData: "{code:GetCapturedTempRoot}"

[Code]
const
  SCS_64BIT_BINARY = 6;
  LegacyDetectedExitCode = 10;
  InstallerContextSubkey = 'Software\TeXShift\InstallerContext';
  InstallerStateSubkey = 'Software\TeXShift\Installer';
  UninstallRegistrySubkey = 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{861A99E6-1888-40DC-9593-3779EDACB74B}_is1';

var
  DeleteUserDataRequested: Boolean;
  RecordedUserDataAvailable: Boolean;
  OriginalUserLegacyInstalled: Boolean;
  LegacyUpgradeDetected: Boolean;
  LegacyUpgradePage: TOutputMsgWizardPage;
  LegacyUpgradeOpenAppsButton: TNewButton;
  LegacyNextCaptionApplied: Boolean;
  CapturedOwnerSid: String;
  CapturedAppDataRoot: String;
  CapturedLocalAppDataRoot: String;
  CapturedTempRoot: String;
  RecordedAppDataRoot: String;
  RecordedLocalAppDataRoot: String;
  RecordedTempRoot: String;

function GetBinaryType(lpApplicationName: String; var lpBinaryType: Cardinal): Boolean;
  external 'GetBinaryTypeW@kernel32.dll stdcall';

function GetAddInCodeBase(Param: String): String;
var
  Path: String;
begin
  Path := ExpandConstant('{app}\TeXShift.AddIn.dll');
  StringChangeEx(Path, '\', '/', True);
  Result := 'file:///' + Path;
end;

function GetCapturedOwnerSid(Param: String): String;
begin
  Result := CapturedOwnerSid;
end;

function GetCapturedAppDataRoot(Param: String): String;
begin
  Result := CapturedAppDataRoot;
end;

function GetCapturedLocalAppDataRoot(Param: String): String;
begin
  Result := CapturedLocalAppDataRoot;
end;

function GetCapturedTempRoot(Param: String): String;
begin
  Result := CapturedTempRoot;
end;

function GetOneNotePath(var Path: String): Boolean;
begin
  Result := RegQueryStringValue(
    HKLM64,
    'Software\Microsoft\Windows\CurrentVersion\App Paths\ONENOTE.EXE',
    '',
    Path);
end;

procedure SetCurrentUserContextFallback;
begin
  CapturedOwnerSid := '';
  CapturedAppDataRoot := GetEnv('APPDATA');
  CapturedLocalAppDataRoot := GetEnv('LOCALAPPDATA');
  CapturedTempRoot := GetEnv('TEMP');
end;

function LoadOriginalUserContext(Token: String): Boolean;
var
  UserSids: TArrayOfString;
  Index: Integer;
  ContextKey: String;
  StoredToken: String;
  OwnerSid: String;
  AppDataRoot: String;
  LocalAppDataRoot: String;
  TempRoot: String;
begin
  Result := False;
  if not RegGetSubkeyNames(HKU, '', UserSids) then
    Exit;

  for Index := 0 to GetArrayLength(UserSids) - 1 do
  begin
    ContextKey := UserSids[Index] + '\' + InstallerContextSubkey;
    if RegQueryStringValue(HKU, ContextKey, 'Token', StoredToken) and
       (StoredToken = Token) then
    begin
      OwnerSid := UserSids[Index];
      Result :=
        RegQueryStringValue(HKU, ContextKey, 'AppDataRoot', AppDataRoot) and
        RegQueryStringValue(HKU, ContextKey, 'LocalAppDataRoot', LocalAppDataRoot) and
        RegQueryStringValue(HKU, ContextKey, 'TempRoot', TempRoot) and
        (AppDataRoot <> '') and (LocalAppDataRoot <> '') and (TempRoot <> '');
      RegDeleteKeyIncludingSubkeys(HKU, ContextKey);
      if Result then
      begin
        CapturedOwnerSid := OwnerSid;
        CapturedAppDataRoot := AppDataRoot;
        CapturedLocalAppDataRoot := LocalAppDataRoot;
        CapturedTempRoot := TempRoot;
      end;
      Exit;
    end;
  end;
end;

procedure CaptureOriginalUserContext;
var
  Token: String;
  Command: String;
  ResultCode: Integer;
  Started: Boolean;
begin
  SetCurrentUserContextFallback;
  OriginalUserLegacyInstalled := False;
  Token := GetSHA256OfUnicodeString(
    ExpandConstant('{srcexe}') + GetDateTimeString('yyyymmddhhnnsszzz', '-', ':'));
  Command :=
    '$ErrorActionPreference = ''Stop''; ' +
    '$key = ''HKCU:\Software\TeXShift\InstallerContext''; ' +
    'New-Item -Path $key -Force | Out-Null; ' +
    'New-ItemProperty -Path $key -Name Token -Value ''' + Token + ''' -PropertyType String -Force | Out-Null; ' +
    'New-ItemProperty -Path $key -Name AppDataRoot -Value $env:APPDATA -PropertyType String -Force | Out-Null; ' +
    'New-ItemProperty -Path $key -Name LocalAppDataRoot -Value $env:LOCALAPPDATA -PropertyType String -Force | Out-Null; ' +
    'New-ItemProperty -Path $key -Name TempRoot -Value $env:TEMP -PropertyType String -Force | Out-Null; ' +
    '$installer = New-Object -ComObject WindowsInstaller.Installer; ' +
    'if (@($installer.RelatedProducts(''{#LegacyUpgradeCode}'')).Count -gt 0) { exit ' +
      IntToStr(LegacyDetectedExitCode) + ' }';
  Started := ExecAsOriginalUser(
    ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe'),
    '-NoProfile -NonInteractive -ExecutionPolicy Bypass -Command "' + Command + '"',
    '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  if Started then
  begin
    LoadOriginalUserContext(Token);
    OriginalUserLegacyInstalled := ResultCode = LegacyDetectedExitCode;
  end;
end;

function LoadRecordedUserDataRoots: Boolean;
begin
  Result :=
    RegQueryStringValue(HKLM64, InstallerStateSubkey, 'AppDataRoot', RecordedAppDataRoot) and
    RegQueryStringValue(HKLM64, InstallerStateSubkey, 'LocalAppDataRoot', RecordedLocalAppDataRoot) and
    RegQueryStringValue(HKLM64, InstallerStateSubkey, 'TempRoot', RecordedTempRoot) and
    (RecordedAppDataRoot <> '') and
    (RecordedLocalAppDataRoot <> '') and
    (RecordedTempRoot <> '');
end;

function IsCommandLineParameterPresent(Parameter: String): Boolean;
var
  Index: Integer;
begin
  Result := False;
  for Index := 1 to ParamCount do
  begin
    if CompareText(ParamStr(Index), Parameter) = 0 then
    begin
      Result := True;
      Exit;
    end;
  end;
end;

function ShouldDeleteUserData: Boolean;
begin
  Result := RecordedUserDataAvailable and
    (IsCommandLineParameterPresent('/DELETEUSERDATA') or DeleteUserDataRequested);
end;

function ShowInteractiveUninstallDialog: Boolean;
var
  Form: TSetupForm;
  PromptLabel: TNewStaticText;
  DeleteUserDataCheckBox: TNewCheckBox;
  DeleteUserDataDetailsLabel: TNewStaticText;
  Bevel: TBevel;
  UninstallButton: TNewButton;
  CancelButton: TNewButton;
  ButtonWidth: Integer;
begin
  Result := False;
  Form := CreateCustomForm(ScaleX(440), ScaleY(205), False, True);
  try
    Form.Caption := CustomMessage('UninstallTitle');

    PromptLabel := TNewStaticText.Create(Form);
    PromptLabel.Parent := Form;
    PromptLabel.Left := ScaleX(24);
    PromptLabel.Top := ScaleY(24);
    PromptLabel.Width := Form.ClientWidth - ScaleX(48);
    PromptLabel.AutoSize := False;
    PromptLabel.WordWrap := True;
    PromptLabel.Font.Style := [fsBold];
    PromptLabel.Caption := CustomMessage('UninstallPrompt');
    PromptLabel.AdjustHeight;

    DeleteUserDataCheckBox := TNewCheckBox.Create(Form);
    DeleteUserDataCheckBox.Parent := Form;
    DeleteUserDataCheckBox.Left := PromptLabel.Left;
    DeleteUserDataCheckBox.Top := PromptLabel.Top + PromptLabel.Height + ScaleY(24);
    DeleteUserDataCheckBox.Width := PromptLabel.Width;
    DeleteUserDataCheckBox.Height := ScaleY(20);
    DeleteUserDataCheckBox.Caption := CustomMessage('DeleteUserDataOption');
    DeleteUserDataCheckBox.Checked := IsCommandLineParameterPresent('/DELETEUSERDATA');
    DeleteUserDataCheckBox.Visible := RecordedUserDataAvailable;

    DeleteUserDataDetailsLabel := TNewStaticText.Create(Form);
    DeleteUserDataDetailsLabel.Parent := Form;
    DeleteUserDataDetailsLabel.Left := DeleteUserDataCheckBox.Left + ScaleX(20);
    DeleteUserDataDetailsLabel.Top := DeleteUserDataCheckBox.Top +
      DeleteUserDataCheckBox.Height + ScaleY(4);
    DeleteUserDataDetailsLabel.Width := PromptLabel.Width - ScaleX(20);
    DeleteUserDataDetailsLabel.AutoSize := False;
    DeleteUserDataDetailsLabel.WordWrap := True;
    DeleteUserDataDetailsLabel.Caption := CustomMessage('DeleteUserDataDetails');
    DeleteUserDataDetailsLabel.AdjustHeight;
    DeleteUserDataDetailsLabel.Visible := RecordedUserDataAvailable;

    Bevel := TBevel.Create(Form);
    Bevel.Parent := Form;
    Bevel.Left := 0;
    Bevel.Top := Form.ClientHeight - ScaleY(50);
    Bevel.Width := Form.ClientWidth;
    Bevel.Height := ScaleY(2);
    Bevel.Shape := bsTopLine;

    CancelButton := TNewButton.Create(Form);
    CancelButton.Parent := Form;
    CancelButton.Caption := SetupMessage(msgButtonCancel);
    CancelButton.Top := Form.ClientHeight - ScaleY(38);
    CancelButton.Height := ScaleY(23);
    CancelButton.ModalResult := mrCancel;
    CancelButton.Cancel := True;

    UninstallButton := TNewButton.Create(Form);
    UninstallButton.Parent := Form;
    UninstallButton.Caption := CustomMessage('UninstallButton');
    UninstallButton.Top := CancelButton.Top;
    UninstallButton.Height := CancelButton.Height;
    UninstallButton.ModalResult := mrOk;
    UninstallButton.Default := True;

    ButtonWidth := Form.CalculateButtonWidth([UninstallButton.Caption, CancelButton.Caption]);
    CancelButton.Width := ButtonWidth;
    CancelButton.Left := Form.ClientWidth - ButtonWidth - ScaleX(16);
    UninstallButton.Width := ButtonWidth;
    UninstallButton.Left := CancelButton.Left - ButtonWidth - ScaleX(8);
    Form.ActiveControl := UninstallButton;

    Result := Form.ShowModal = mrOk;
    if Result and RecordedUserDataAvailable then
      DeleteUserDataRequested := DeleteUserDataCheckBox.Checked;
  finally
    Form.Free;
  end;
end;

function InitializeUninstall: Boolean;
begin
  DeleteUserDataRequested := False;
  RecordedUserDataAvailable := LoadRecordedUserDataRoots;
  if IsCommandLineParameterPresent('/INTERACTIVE') then
    Result := ShowInteractiveUninstallDialog
  else
    Result := True;
end;

procedure DeleteDataDirectory(BasePath: String; DirectoryName: String);
var
  Path: String;
begin
  if BasePath = '' then
    Exit;

  Path := ExpandFileName(AddBackslash(BasePath) + DirectoryName);
  if DirExists(Path) and (not DelTree(Path, True, True, True)) then
    Log('Unable to delete user data directory: ' + Path);
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if (CurUninstallStep = usPostUninstall) and ShouldDeleteUserData then
  begin
    DeleteDataDirectory(RecordedAppDataRoot, 'TeXShift');
    DeleteDataDirectory(RecordedLocalAppDataRoot, 'TeXShift');
    DeleteDataDirectory(RecordedTempRoot, 'TeXShift_WebView2');
    DeleteDataDirectory(RecordedTempRoot, 'TeXShift_Mermaid_WebView2');
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  InteractiveUninstallCommand: String;
begin
  if CurStep = ssPostInstall then
  begin
    InteractiveUninstallCommand := '"' + ExpandConstant('{uninstallexe}') +
      '" /SILENT /INTERACTIVE';
    if not RegWriteStringValue(
      HKLM64, UninstallRegistrySubkey, 'UninstallString', InteractiveUninstallCommand) then
      Log('Unable to configure the interactive uninstall command.');
  end;
end;

function InitializeSetup: Boolean;
var
  OneNotePath: String;
  BinaryType: Cardinal;
begin
  Result := False;

  if not IsDotNetInstalled(net48, 0) then
  begin
    TaskDialogMsgBox(
      CustomMessage('DotNetRequired'), '', mbError,
      MB_OK, [CustomMessage('DialogButtonOK')], 0);
    Exit;
  end;

  if not GetOneNotePath(OneNotePath) then
  begin
    TaskDialogMsgBox(
      CustomMessage('OneNoteNotFound'), '', mbError,
      MB_OK, [CustomMessage('DialogButtonOK')], 0);
    Exit;
  end;

  if (not GetBinaryType(OneNotePath, BinaryType)) or (BinaryType <> SCS_64BIT_BINARY) then
  begin
    TaskDialogMsgBox(
      CustomMessage('OneNoteX64Required'), '', mbError,
      MB_OK, [CustomMessage('DialogButtonOK')], 0);
    Exit;
  end;

  CaptureOriginalUserContext;
  LegacyUpgradeDetected := OriginalUserLegacyInstalled or
    IsMsiProductInstalled('{#LegacyUpgradeCode}', 0);
  Result := True;
end;

procedure OpenInstalledAppsButtonClick(Sender: TObject);
var
  ErrorCode: Integer;
begin
  ShellExecAsOriginalUser(
    'open', 'ms-settings:appsfeatures', '', '', SW_SHOWNORMAL, ewNoWait, ErrorCode);
end;

procedure InitializeWizard;
begin
  WizardForm.WelcomeLabel2.Caption :=
    CustomMessage('LegacyWelcomeWarning') + #13#10#13#10 +
    WizardForm.WelcomeLabel2.Caption;

  if LegacyUpgradeDetected then
  begin
    LegacyUpgradePage := CreateOutputMsgPage(
      wpWelcome,
      CustomMessage('LegacyUpgradePageTitle'),
      CustomMessage('LegacyUpgradeDetected'),
      CustomMessage('LegacyUpgradeAction'));
    LegacyUpgradeOpenAppsButton := TNewButton.Create(LegacyUpgradePage);
    LegacyUpgradeOpenAppsButton.Parent := LegacyUpgradePage.Surface;
    LegacyUpgradeOpenAppsButton.Caption := CustomMessage('LegacyUpgradeOpenApps');
    LegacyUpgradeOpenAppsButton.Width := WizardForm.CalculateButtonWidth([LegacyUpgradeOpenAppsButton.Caption]);
    LegacyUpgradeOpenAppsButton.Height := WizardForm.NextButton.Height;
    LegacyUpgradeOpenAppsButton.Top := LegacyUpgradePage.MsgLabel.Top +
      LegacyUpgradePage.MsgLabel.Height + ScaleY(16);
    LegacyUpgradeOpenAppsButton.OnClick := @OpenInstalledAppsButtonClick;
  end;
end;

procedure CurPageChanged(CurPageID: Integer);
begin
  if Assigned(LegacyUpgradePage) and (CurPageID = LegacyUpgradePage.ID) then
  begin
    WizardForm.NextButton.Caption := CustomMessage('LegacyUpgradeContinue');
    LegacyNextCaptionApplied := True;
  end
  else if LegacyNextCaptionApplied then
  begin
    WizardForm.NextButton.Caption := SetupMessage(msgButtonNext);
    LegacyNextCaptionApplied := False;
  end;
end;
