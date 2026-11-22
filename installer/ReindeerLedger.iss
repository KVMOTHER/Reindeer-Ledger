#define MyAppName "Reindeer Ledger"
#define MyAppVersion "0.6.0"
#define MyAppPublisher "KV/RaMIS Project Group"
#define MyAppURL "https://kvramispg.vercel.app"
#define MyAppExeName "ReindeerLedger.exe"

[Setup]
AppId={{A1F2E4B5-93E7-4A9B-86C2-DA0A7F0BC8D2}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} v{#MyAppVersion} RC
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={localappdata}\Programs\Reindeer Ledger
DefaultGroupName=Reindeer Ledger
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=..\release
OutputBaseFilename=ReindeerLedger_Setup_v0.6.0_RC
SetupIconFile=..\assets\app.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
InfoBeforeFile=welcome_ru.txt
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

[Tasks]
Name: "desktopicon"; Description: "Создать ярлык на рабочем столе"; GroupDescription: "Дополнительные задачи:"; Flags: checkedonce

[Files]
Source: "..\dist\ReindeerLedger\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Reindeer Ledger"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\Reindeer Ledger"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Запустить Reindeer Ledger"; Flags: nowait postinstall skipifsilent

[Code]
var
  SiteLabel: TNewStaticText;
  MailLabel: TNewStaticText;

procedure OpenSite(Sender: TObject);
var ResultCode: Integer;
begin
  ShellExec('open', 'https://kvramispg.vercel.app', '', '', SW_SHOWNORMAL, ewNoWait, ResultCode);
end;

procedure OpenMail(Sender: TObject);
var ResultCode: Integer;
begin
  ShellExec('open', 'mailto:kvramisprojectgroup@gmail.com', '', '', SW_SHOWNORMAL, ewNoWait, ResultCode);
end;

procedure InitializeWizard();
begin
  SiteLabel := TNewStaticText.Create(WizardForm);
  SiteLabel.Parent := WizardForm.WelcomePage;
  SiteLabel.Caption := 'Сайт: kvramispg.vercel.app';
  SiteLabel.Cursor := crHand;
  SiteLabel.Font.Color := clBlue;
  SiteLabel.OnClick := @OpenSite;
  SiteLabel.Left := WizardForm.WelcomeLabel2.Left;
  SiteLabel.Top := WizardForm.WelcomeLabel2.Top + WizardForm.WelcomeLabel2.Height + ScaleY(14);
  SiteLabel.AutoSize := True;

  MailLabel := TNewStaticText.Create(WizardForm);
  MailLabel.Parent := WizardForm.WelcomePage;
  MailLabel.Caption := 'Поддержка: kvramisprojectgroup@gmail.com';
  MailLabel.Cursor := crHand;
  MailLabel.Font.Color := clBlue;
  MailLabel.OnClick := @OpenMail;
  MailLabel.Left := SiteLabel.Left;
  MailLabel.Top := SiteLabel.Top + ScaleY(22);
  MailLabel.AutoSize := True;
end;
