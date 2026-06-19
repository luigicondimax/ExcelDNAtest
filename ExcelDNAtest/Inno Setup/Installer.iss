; ONLY FOR TEST THE SCRIPT IN INNO SETUP COMPILER
;#define IsAOT false
;#define ReleasePath "C:\Users\luigi\source\repos\ExcelDNAtest\ExcelDNAtest\bin\ReleaseAOT\net10.0-windows"
;#define ReleasePath "C:\Users\luigi\source\repos\ExcelDNAtest\ExcelDNAtest\bin\Release\net10.0-windows"
;#define XllName "ExcelDNAtest"
#define ResourcePath "C:\Users\luigi\source\repos\ExcelDNAtest\ExcelDNAtest\Inno Setup\Resources\"
;#define AppName "ExcelDNAtest"
; END TEST 

#if IsAOT == "true"
  #define ReleasePath32 ReleasePath + "\win-x86\publish"
  #define ReleasePath64 ReleasePath + "\win-x64\publish"

  #define AppDll32 XllName + "-AddIn64.xll" ; TO CHECK
  #define AppDll64 XllName + "-AddIn64.xll"
#else
  #define ReleasePath32 ReleasePath + "\publish"
  #define ReleasePath64 ReleasePath32

  #define AppDll32 XllName + "-AddIn-packed.xll"
  #define AppDll64 XllName + "-AddIn64-packed.xll"
#endif

#define OutputPath ReleasePath + "\Installer"

#define AppVersion GetVersionNumbersString(ReleasePath64 + "\" + AppDll64)
#define AppPublisher "Luigi Serio"
#define AppURL "https://www.sss.com"
#define AppId "{{1BF7AF30-CCF1-4D09-9267-7E70F239A172}}"  ; CHANGE TO UNIQUE GUID
#define NomeLibreria XllName
#define NETVersion "10.0"

[Languages]
Name: "en"; MessagesFile: "compiler:Default.isl"
Name: "it"; MessagesFile: "compiler:Languages\Italian.isl"

[CustomMessages]
it.ExcelRunningError=Microsoft Excel è attualmente aperto.%n%nChiudi Excel e clicca su Riprova per procedere.
en.ExcelRunningError=Microsoft Excel is currently open.%n%nPlease close Excel and click Retry to proceed.
it.ExcelNotFound=Excel non rilevato. L'installazione verrà interrotta.
en.ExcelNotFound=Excel not detected. Installation will be interrupted.
it.NetDesktopRequired=Per procedere con l'installazione è necessario il runtime .NET {#NETVersion} Desktop.%n%n \
Procedo con l'installazione?
en.NetDesktopRequired=To proceed with the installation, you need the .NET {#NETVersion} Desktop runtime.%n%n \
Do you want to proceed with the installation?
it.NetDesktopNotInstalled=Non ho potuto installare il runtime .NET {#NETVersion} Desktop.%n%n \
Scaricare dal sito https://dotnet.microsoft.com/it-it/download/dotnet/{#NETVersion} prima di riavviare l'installazione
en.NetDesktopNotInstalled=I couldn't install the .NET {#NETVersion} Desktop runtime.%n%n \
Download it from https://dotnet.microsoft.com/it-it/download/dotnet/{#NETVersion} before restarting the installation.

[Setup]
AppId={#AppId}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
UsePreviousAppDir=yes
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
WizardStyle=modern dynamic
DisableDirPage=auto
AllowNoIcons=yes
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
CloseApplications=yes

; Icona e Grafica
SetupIconFile={#ResourcePath}company.ico
UninstallDisplayIcon={#ResourcePath}company.ico
LicenseFile={#ResourcePath}product_License.rtf
WizardSmallImageFile={#ResourcePath}company2.ico

; Output
OutputDir={#OutputPath}
OutputBaseFilename=ExcelToolsetInstaller
Compression=lzma
SolidCompression=yes

; Architettura (permette installazione su sistemi a 64 bit)
ArchitecturesInstallIn64BitMode=x64compatible
ArchitecturesAllowed=x64compatible and x86compatible

//[Components]
//Name: "addin32"; Description: "Add-in per Excel (32-bit)"; Flags: exclusive;
//Name: "addin64"; Description: "Add-in per Excel (64-bit)"; Flags: exclusive;

[Files]
; File comuni
;Source: "{#ReleasePath}ExcelToolset.dll"; DestDir: "{app}"; Flags: ignoreversion
;Source: "{#ReleasePath}ExcelDna.IntelliSense.dll"; DestDir: "{app}"; Flags: ignoreversion
;Source: "{#ReleasePath}ExcelToolset-AddIn.deps.json"; DestDir: "{app}"; Flags: ignoreversion

; Risorse localizzate
;Source: "{#ReleasePath}it\ExcelToolset.resources.dll"; DestDir: "{app}\it"; Flags: ignoreversion
;Source: "{#ReleasePath}en\ExcelToolset.resources.dll"; DestDir: "{app}\en"; Flags: ignoreversion

; File condizionali (Installati solo in base ai Bit di Excel rilevati nel [Code])
; Components: addin64
;Source: "{#ReleasePath}{#NomeLibreria}.xll"; DestDir: "{app}"; Flags: ignoreversion; Check: IsExcel32
;Source: "{#ReleasePath}{#NomeLibreria}.dna"; DestDir: "{app}"; Flags: ignoreversion; Check: IsExcel32
;Source: "{#ReleasePath}{#NomeLibreria}64.xll"; DestDir: "{app}"; Flags: ignoreversion; Check: IsExcel64
;Source: "{#ReleasePath}{#NomeLibreria}64.dna"; DestDir: "{app}"; Flags: ignoreversion; Check: IsExcel64

#if IsAOT != "true"
  Source: "{#ReleasePath32}\{#AppDll32}"; DestDir: "{app}"; Flags: ignoreversion; Check: IsExcel32
  
  Source: "{#ReleasePath64}\{#AppDll64}"; DestDir: "{app}"; Flags: ignoreversion; Check: IsExcel64
#else
  Source: "{#ReleasePath32}\{#AppDll32}"; DestDir: "{app}"; Flags: ignoreversion; Check: IsExcel32
  Source: "{#ReleasePath32}\libHarfBuzzSharp.dll"; DestDir: "{app}"; Flags: ignoreversion; Check: IsExcel32
  Source: "{#ReleasePath32}\libSkiaSharp.dll"; DestDir: "{app}"; Flags: ignoreversion; Check: IsExcel32
  Source: "{#ReleasePath32}\av_libglesv2.dll"; DestDir: "{app}"; Flags: ignoreversion; Check: IsExcel32

  Source: "{#ReleasePath64}\{#AppDll64}"; DestDir: "{app}"; Flags: ignoreversion; Check: IsExcel64
  Source: "{#ReleasePath64}\libHarfBuzzSharp.dll"; DestDir: "{app}"; Flags: ignoreversion; Check: IsExcel64
  Source: "{#ReleasePath64}\libSkiaSharp.dll"; DestDir: "{app}"; Flags: ignoreversion; Check: IsExcel64
  Source: "{#ReleasePath64}\av_libglesv2.dll"; DestDir: "{app}"; Flags: ignoreversion; Check: IsExcel64
#endif

[Icons]
Name: "{group}\My Program"; Filename: "{app}\MyProg.exe"

[Code]
var
  ExcelVersion: String;
  ExcelBit: String;
  RestartRequired: Boolean;

#include "utility.pas"
#include "logic.pas"

function InitializeSetup(): Boolean;
var
  Response: Integer;
  Msg:String;
#if IsAOT != "true"
  ErrorCode: Integer;
#endif

begin
  Result := False; // Di base non procediamo finché non siamo pronti

  // 1. Loop di controllo Excel aperto
  while IsExcelRunning() do
  begin
    Msg := CustomMessage('ExcelRunningError');
    Response := MsgBox(Msg, mbError, MB_RETRYCANCEL);
    
    if Response = IDCANCEL then
    begin
      // L'utente ha annullato la disinstallazione
      Exit;
    end;
  end;

  GetExcelVersionAndBit();

  if ExcelBit = '0' then begin
    Msg := CustomMessage('ExcelNotFound');

    MsgBox(Msg, mbError, MB_OK);
    Result := False;
    Exit;
  end;

#if IsAOT != "true"

  if not IsDotNetDesktopRuntimeInstalled() then begin  
    Msg := CustomMessage('NetDesktopRequired');
    
    if MsgBox(msg, mbConfirmation, MB_YESNO) = IDYES then begin
      Result := InstallDotNetViaWinget();
      if  Result=False then
        begin
          Msg := CustomMessage('NetDesktopNotInstalled');
          MsgBox(msg, mbConfirmation, MB_OK);
          ShellExec('open', 'https://dotnet.microsoft.com/en-us/download/dotnet/{#NETVersion}', '', '', SW_SHOWNORMAL, ewNoWait, ErrorCode);
        end;
      Exit
      end
    else
      begin
        Msg := CustomMessage('NetDesktopNotInstalled');
        MsgBox(msg, mbConfirmation, MB_OK);
        ShellExec('open', 'https://dotnet.microsoft.com/en-us/download/dotnet/{#NETVersion}', '', '', SW_SHOWNORMAL, ewNoWait, ErrorCode);
        Result := False; // Blocca installazione come il tuo Launch Condition
        Exit
      end
  end;
#endif
    
  Result :=True;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  // ssPostInstall avviene dopo che i file sono stati copiati sul disco
  if CurStep = ssPostInstall then
  begin
    if ExcelVersion <> '0' then
    begin
      AddExcelKey();
    end;
  end;
end;

function InitializeUninstall(): Boolean;
var
  Response: Integer;
  Msg:String;
begin
  Result := False; // Di base non procediamo finché non siamo pronti

  // 1. Loop di controllo Excel aperto
  while IsExcelRunning() do
  begin
    Msg := CustomMessage('ExcelRunningError');
    Response := MsgBox(Msg, mbError, MB_RETRYCANCEL);
    
    if Response = IDCANCEL then
    begin
      // L'utente ha annullato la disinstallazione
      Exit;
    end;
  end;

  Result := True; // Ritorna True per procedere con la disinstallazione
end;

procedure CurUninstallStepChanged(UninstallStep: TUninstallStep);
begin
  // usUninstall è il passo che avviene DOPO che l'utente ha confermato 
  if UninstallStep = usUninstall then
  begin
    Log('Start to remove excel key...');
    
    // Rileviamo la versione (le variabili globali ExcelVersion/Bit non sono persistenti)
    GetExcelVersionAndBit();
    
    if ExcelVersion <> '0' then
    begin
      RemoveExcelKey();
    end;
  end;
end;

function NeedRestart(): Boolean;
begin
  Result := RestartRequired;
end;