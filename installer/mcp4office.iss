; Installatore per utente di mcp4office (nessun diritto di amministratore richiesto).
; Se gia' installato, scrive il marcatore .upgrading prima di disinstallare i vecchi binari
; così le configurazioni MCP e gli args utente (es. --sola-lettura, --pdf-roots) vengono preservati.

#define AppName "mcp4office"
#ifndef AppVersion
  #define AppVersion "0.5.7"
#endif
#define AppId "{7F4D9B21-6C38-4E2A-B8E1-9D3A5F2C1084}"
#define BuildDir "..\bin\Release\net48"

[Setup]
AppId={{7F4D9B21-6C38-4E2A-B8E1-9D3A5F2C1084}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=mcp4office (locale)
DefaultDirName={localappdata}\Programs\{#AppName}
DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=..\dist
OutputBaseFilename=mcp4office-Setup-{#AppVersion}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
UninstallDisplayName={#AppName} {#AppVersion}
UninstallDisplayIcon={app}\mcp4office.exe
ArchitecturesAllowed=x64compatible x86compatible
CloseApplications=yes

[Languages]
Name: "italian"; MessagesFile: "compiler:Languages\Italian.isl"

[Files]
Source: "{#BuildDir}\mcp4office.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildDir}\mcp4office.exe.config"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "{#BuildDir}\*.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion

[UninstallDelete]
Type: files; Name: "{app}\install-report.txt"
Type: files; Name: "{app}\*.tmp"
Type: dirifempty; Name: "{app}"

[Code]
const
  UninstallRegKey = 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4D9B21-6C38-4E2A-B8E1-9D3A5F2C1084}_is1';

function IsMcp4OfficeRunning(): Boolean;
var
  rc: Integer;
begin
  Result := Exec('cmd.exe', '/C tasklist /FI "IMAGENAME eq mcp4office.exe" 2>NUL | find /I "mcp4office.exe" >NUL', '', SW_HIDE, ewWaitUntilTerminated, rc) and (rc = 0);
end;

function GetExistingUninstaller(): String;
var
  s: String;
begin
  Result := '';
  if RegQueryStringValue(HKCU, UninstallRegKey, 'QuietUninstallString', s) and (s <> '') then
    Result := s
  else if RegQueryStringValue(HKCU, UninstallRegKey, 'UninstallString', s) and (s <> '') then
    Result := s;
end;

function LoadUtf8Report(const Path: String): String;
var
  lines: TArrayOfString;
  i: Integer;
begin
  Result := '';
  if FileExists(Path) and LoadStringsFromFile(Path, lines) then
  begin
    for i := 0 to GetArrayLength(lines) - 1 do
      Result := Result + lines[i] + #13#10;
  end;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  uninst, exePath, params, markerPath: String;
  rc: Integer;
begin
  Result := '';

  if IsMcp4OfficeRunning() then
  begin
    Result := 'mcp4office.exe e'' attualmente in esecuzione (aperto da Antigravity, Claude Desktop o Codex).' + #13#10 +
              'Chiudi i client AI aperti e rilancia il Setup.';
    Exit;
  end;

  uninst := GetExistingUninstaller();
  if uninst = '' then
    Exit;

  exePath := RemoveQuotes(uninst);
  if not FileExists(exePath) then
  begin
    if (Length(uninst) > 1) and (uninst[1] = '"') then
    begin
      Delete(uninst, 1, 1);
      rc := Pos('"', uninst);
      if rc > 1 then
        exePath := Copy(uninst, 1, rc - 1);
    end;
  end;

  if FileExists(exePath) then
  begin
    // Scrive marcatore .upgrading cosi' la disinstallazione rimuove i vecchi binari ma PRESERVA config MCP, args utente e install-manifest.json!
    markerPath := AddBackslash(ExtractFileDir(exePath)) + '.upgrading';
    SaveStringToFile(markerPath, 'upgrade', False);

    params := '/VERYSILENT /NORESTART /SUPPRESSMSGBOXES';
    if not Exec(exePath, params, '', SW_HIDE, ewWaitUntilTerminated, rc) then
    begin
      DeleteFile(markerPath);
      Result := 'Impossibile disinstallare automaticamente la versione precedente di mcp4office. Chiudi i client AI e riprova.';
      Exit;
    end;
    if rc <> 0 then
    begin
      DeleteFile(markerPath);
      Result := 'La disinstallazione della versione precedente ha restituito codice ' + IntToStr(rc) + '. Chiudi i client AI aperti e riprova.';
      Exit;
    end;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  rc: Integer;
  reportFile, reportText: String;
begin
  if CurStep = ssPostInstall then
  begin
    reportFile := ExpandConstant('{app}\install-report.txt');
    rc := -1;
    Exec(ExpandConstant('{app}\mcp4office.exe'),
         '--configure-clients --report "' + reportFile + '"',
         '', SW_HIDE, ewWaitUntilTerminated, rc);
    // Rimuove in ogni caso l'eventuale marcatore .upgrading qualunque sia stato l'esito
    DeleteFile(ExpandConstant('{app}\.upgrading'));

    reportText := LoadUtf8Report(reportFile);

    if not WizardSilent then
    begin
      if rc = 0 then
        MsgBox('Installazione di {#AppName} {#AppVersion} completata!' + #13#10 + #13#10 +
               'Esito configurazione client MCP:' + #13#10 + reportText + #13#10 +
               'Riavvia i client AI aperti per caricare i tool di mcp4office.',
               mbInformation, MB_OK)
      else
        MsgBox('{#AppName} {#AppVersion} installato, ma uno o piu'' client MCP hanno segnalato errori (codice ' + IntToStr(rc) + '):' + #13#10 + #13#10 +
               reportText + #13#10 +
               'Verifica i file indicati o configura manualmente seguendo README.md.',
               mbError, MB_OK);
    end;
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  rc: Integer;
  exeFile, reportFile, reportText: String;
begin
  if CurUninstallStep = usUninstall then
  begin
    exeFile := ExpandConstant('{app}\mcp4office.exe');
    reportFile := ExpandConstant('{app}\uninstall-report.txt');
    if FileExists(exeFile) then
    begin
      rc := -1;
      Exec(exeFile, '--unconfigure-clients --report "' + reportFile + '"', '', SW_HIDE, ewWaitUntilTerminated, rc);
      if rc <> 0 then
      begin
        // In caso di errore, preserva install-manifest.json e avvisa l'utente
        reportText := LoadUtf8Report(reportFile);
        if not UninstallSilent then
          MsgBox('Attenzione: rimozione delle voci MCP incompleta (codice ' + IntToStr(rc) + ').' + #13#10 +
                 'install-manifest.json e'' stato preservato per il recupero.' + #13#10 + #13#10 + reportText,
                 mbError, MB_OK);
      end
      else
      begin
        DeleteFile(reportFile);
      end;
    end;
  end;
end;
