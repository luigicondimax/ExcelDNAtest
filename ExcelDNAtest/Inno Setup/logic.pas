// --- COSTANTI E API WIN32 ---
const
  SCS_32BIT_BINARY = 0;
  SCS_64BIT_BINARY = 6;

// Importazione dell'API per controllare se un EXE è a 32 o 64 bit
function GetBinaryType(lpApplicationName: String; var lpBinaryType: Integer): Boolean;
external 'GetBinaryTypeW@kernel32.dll stdcall';
// --- 1. RILEVAMENTO VERSIONE E BITNESS ---

// --- FUNZIONE DI SUPPORTO: ESTRAE VERSIONE E BIT DA UN EXE ---
procedure OttieniVersionAndBitByExe(FilePath: String; var Versione: String; var Bitness: String);
var
  MS, LS: Cardinal;
  BinType: Integer;
begin
  if FileExists(FilePath) then
  begin
    // Estrae la versione del file (Major Part)
    if GetVersionNumbers(FilePath, MS, LS) then
      Versione := IntToStr(MS shr 16)
    else
      Versione := '0';

    // Determina il bitness (32 o 64)
    if GetBinaryType(FilePath, BinType) then
    begin
      case BinType of
        SCS_32BIT_BINARY: Bitness := '32';
        SCS_64BIT_BINARY: Bitness := '64';
      else
        Bitness := '0';
      end;
    end;
  end;
end;

// --- 1. OUTLOOK KEY ---
procedure GetVersionAndBitByOutlookKey(var Versione: String; var Bitness: String);
var
  CurVer, TempBitness: String;
  DotPos: Integer;
begin
  Log('Tentativo 1: Outlook Registry');
  if RegQueryStringValue(HKEY_CLASSES_ROOT, 'Outlook.Application\CurVer', '', CurVer) then
  begin
    DotPos := LastPos('.', CurVer);
    if DotPos > 0 then
    begin
      Versione := Copy(CurVer, DotPos + 1, Length(CurVer));
      
    // Controlla Bitness (prova sia vista 64 che 32)
    if not RegQueryStringValue(HKEY_LOCAL_MACHINE, 'SOFTWARE\Microsoft\Office\' + Versione + '.0\Outlook', 'Bitness', TempBitness) then
       RegQueryStringValue(HKEY_LOCAL_MACHINE, 'SOFTWARE\WOW6432Node\Microsoft\Office\' + Versione + '.0\Outlook', 'Bitness', TempBitness);

    if TempBitness = 'x64' then Bitness := '64' 
    else if TempBitness = 'x86' then Bitness := '32'
    else Bitness := '0';
  end;
end;
end;

// --- 2. APP PATHS ---
procedure GetVersionAndBitByAppPath(var Versione: String; var Bitness: String);
var
  Path: String;
begin
  Log('Tentativo 2: App Paths');
  if not RegQueryStringValue(HKEY_LOCAL_MACHINE, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe', '', Path) then
     RegQueryStringValue(HKEY_LOCAL_MACHINE, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe', '', Path);

  if Path <> '' then
    begin
      // Rimuovi eventuali virgolette dal percorso
      StringChangeEx(Path, '"', '', True);
      OttieniVersionAndBitByExe(Path, Versione, Bitness);
    end;
  end;

// --- 3. CLICK TO RUN ---
procedure GetVersionAndBitByClickToRunKey(var Versione: String; var Bitness: String);
var
  VerFull, Platform: String;
begin
  Log('Tentativo 3: ClickToRun');
  if RegQueryStringValue(HKEY_LOCAL_MACHINE, 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration', 'ClientVersionToReport', VerFull) then
  begin
    Versione := Copy(VerFull, 1, Pos('.', VerFull) - 1);
    if RegQueryStringValue(HKEY_LOCAL_MACHINE, 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration', 'Platform', Platform) then
    begin
      if Platform = 'x64' then Bitness := '64' else Bitness := '32';
    end;
  end;
end;

// --- 4. EXCEL COM ---
procedure GetVersionAndBitByExcelCOM(var Versione: String; var Bitness: String);
var
  ExcelApp: Variant;
  OSInfo: String;
begin
  Log('Tentativo 4: Excel COM');
  try
    ExcelApp := CreateOleObject('Excel.Application');
    Versione := ExcelApp.Version;
    if Pos('.', Versione) > 0 then Versione := Copy(Versione, 1, Pos('.', Versione) - 1);
    
    OSInfo := ExcelApp.OperatingSystem;
    if Pos('64', OSInfo) > 0 then Bitness := '64' else Bitness := '32';
    
    ExcelApp.Quit;
  except
    Log('Excel COM non disponibile');
  end;
end;

// --- 5. TYPELIB / CLSID ---
procedure GetVersionAndBitByTypeLibKey(var Versione: String; var Bitness: String);
var
  CLSIDKey: String;
  SubKeys: TArrayOfString;
  I: Integer;
  Value, Path: String;
begin
  Log('Tentativo 5: TypeLib Scan');
  CLSIDKey := 'SOFTWARE\Classes\CLSID'; // Inno gestisce HKCR qui
  if RegGetSubkeyNames(HKEY_LOCAL_MACHINE, CLSIDKey, SubKeys) then
  begin
    for I := 0 to GetArrayLength(SubKeys) - 1 do
    begin
      if RegQueryStringValue(HKEY_LOCAL_MACHINE, CLSIDKey + '\' + SubKeys[I] + '\TypeLib', '', Value) then
      begin
        if Value = '{00020813-0000-0000-C000-000000000046}' then
        begin
          if RegQueryStringValue(HKEY_LOCAL_MACHINE, CLSIDKey + '\' + SubKeys[I] + '\LocalServer32', '', Path) then
          begin
            StringChangeEx(Path, '"', '', True);
            OttieniVersionAndBitByExe(Path, Versione, Bitness);
            if Versione <> '0' then Exit;
          end;
        end;
      end;
    end;
  end;
end;

// --- FUNZIONE PRINCIPALE DI ORCHESTRAZIONE ---
procedure GetExcelVersionAndBit();
begin
  ExcelVersion := '0';
  ExcelBit := '0';

  // 1. Outlook
  GetVersionAndBitByOutlookKey(ExcelVersion, ExcelBit);
  if (ExcelVersion <> '0') and (ExcelBit <> '0') then Exit;

  // 2. App Path
  GetVersionAndBitByAppPath(ExcelVersion, ExcelBit);
  if (ExcelVersion <> '0') and (ExcelBit <> '0') then Exit;

  // 3. ClickToRun
  GetVersionAndBitByClickToRunKey(ExcelVersion, ExcelBit);
  if (ExcelVersion <> '0') and (ExcelBit <> '0') then Exit;

  // 4. COM
  GetVersionAndBitByExcelCOM(ExcelVersion, ExcelBit);
  if (ExcelVersion <> '0') and (ExcelBit <> '0') then Exit;

  // 5. TypeLib
  GetVersionAndBitByTypeLibKey(ExcelVersion, ExcelBit);
end;


// --- 2. CONTROLLO .NET DESKTOP RUNTIME ---
function IsDotNetDesktopRuntimeInstalled(): Boolean;
var
  NetArchKey, SharedPath: String;
  BaseKey: String;
  Versions: TArrayOfString;
  I: Integer;
  FindData: TFindRec;
begin
  Result := False;

  // 1. Determiniamo il percorso in base ai bit di Excel
  // Se Excel è 64bit -> C:\Program Files\dotnet
  // Se Excel è 32bit -> C:\Program Files (x86)\dotnet
  if ExcelBit = '64' then
    SharedPath := ExpandConstant('{pf}\dotnet\shared\Microsoft.WindowsDesktop.App\')
  else
    SharedPath := ExpandConstant('{pf32}\dotnet\shared\Microsoft.WindowsDesktop.App\');

  Log('Ricerca .NET Desktop Runtime su disco in: ' + SharedPath);

  // 2. SCANSIONE DELLE CARTELLE (Metodo più affidabile)
  if DirExists(SharedPath) then
  begin
    // Cerchiamo tutte le sottocartelle (ogni cartella è una versione)
    if FindFirst(SharedPath + '*', FindData) then
    begin
      try
        repeat
          if (FindData.Attributes and FILE_ATTRIBUTE_DIRECTORY <> 0) and 
             (FindData.Name <> '.') and (FindData.Name <> '..') then
          begin
            Log('Trovata cartella versione: ' + FindData.Name);
            if CompareVersions(FindData.Name, '{#NETVersion}') >= 0 then
            begin
              Log('La versione ' + FindData.Name + ' soddisfa il requisito.');
              Result := True;
              Break;
            end;
          end;
        until not FindNext(FindData);
      finally
        FindClose(FindData);
      end;
    end;
  end;

  // 3. FALLBACK: REGISTRO (Se non abbiamo trovato nulla su disco)
  if not Result then
  begin
    Log('Nessuna cartella valida trovata, provo fallback su registro...');
    if ExcelBit = '64' then NetArchKey := 'x64' else NetArchKey := 'x86';
    BaseKey := 'SOFTWARE\dotnet\Setup\InstalledVersions\' + NetArchKey + '\sharedfx\Microsoft.WindowsDesktop.App';
    
    if RegGetSubkeyNames(HKEY_LOCAL_MACHINE, BaseKey, Versions) then
    begin
      for I := 0 to GetArrayLength(Versions) - 1 do
      begin
        if CompareVersions(Versions[I], '{#NETVersion}') >= 0 then
        begin
          Result := True;
          Break;
        end;
      end;
    end;
  end;
end;


// --- 3. LOGICA DI REGISTRAZIONE EXCEL (OPEN KEY) ---

procedure RemoveExcelKey();
var
  BaseKey, KeyValue, TargetFile: String;
  ValueNames: TArrayOfString;
  I: Integer;
  CurrentData: String;
begin
  // 1. Costruzione del percorso come nel tuo C# (/R "percorso")
  // Assumiamo che ExcelBit sia popolato da GetExcelVersionAndBit
  if ExcelBit = '32' then
    TargetFile := ExpandConstant('{app}\') + '{#AppDll32}' // Corrisponde a PercorsoLibreria32
  else if ExcelBit = '64' then
    TargetFile := ExpandConstant('{app}\') + '{#AppDll64}' // Corrisponde a PercorsoLibreria64
  else
    Exit;

  // Formattazione identica al tuo C#: /R "C:\Path\To\Addin.xll"
  KeyValue := '/R "' + TargetFile + '"';
  
  // Computer\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\{versione}.0\Excel\Options
  BaseKey := 'Software\Microsoft\Office\' + ExcelVersion + '.0\Excel\Options';

  Log('RemoveExcelKey: Cerco ' + KeyValue + ' in ' + BaseKey);

  // 2. Ottieni tutti i nomi dei valori nella chiave Options
  if RegGetValueNames(HKEY_CURRENT_USER, BaseKey, ValueNames) then
  begin
    for I := 0 to GetArrayLength(ValueNames) - 1 do
    begin
      // Filtro: deve iniziare con "OPEN" (equivalente al tuo .Where(s => s.StartsWith("OPEN")))
      if Pos('OPEN', ValueNames[I]) = 1 then
      begin
        // Leggiamo il dato del valore attuale
        if RegQueryStringValue(HKEY_CURRENT_USER, BaseKey, ValueNames[I], CurrentData) then
        begin
          // Confronto con il percorso calcolato
          if CurrentData = KeyValue then
          begin
            if RegDeleteValue(HKEY_CURRENT_USER, BaseKey, ValueNames[I]) then
              Log('Delete: ' + ValueNames[I])
            else
              Log('Errore durante la cancellazione di: ' + ValueNames[I]);
          end;
        end;
      end;
    end;
  end;
end;

procedure AddExcelKey();
var
  BaseKey, KeyName, KeyValue, ExistingValue: String;
  I: Integer;
  FullAddinPath: String;
begin
  BaseKey := 'Software\Microsoft\Office\' + ExcelVersion + '.0\Excel\Options';
  if ExcelBit = '64' then 
    FullAddinPath := ExpandConstant('{app}\{#AppDll64}')
  else 
    FullAddinPath := ExpandConstant('{app}\{#AppDll32}');
  
  KeyValue := '/R "' + FullAddinPath + '"';

  Log('Tentativo di registrazione Excel: ' + KeyValue);
  
  KeyName := 'OPEN';
  I := 1;
  while RegValueExists(HKEY_CURRENT_USER, BaseKey, KeyName) do begin
    RegQueryStringValue(HKEY_CURRENT_USER, BaseKey, KeyName, ExistingValue);
    if ExistingValue = KeyValue then
    begin
      Log('Add-in già presente in: ' + KeyName);
      Exit; 
    end;
    KeyName := 'OPEN' + IntToStr(I);
    I := I + 1;
  end;
  
  // Scrittura del valore nel primo slot libero trovato
  if RegWriteStringValue(HKEY_CURRENT_USER, BaseKey, KeyName, KeyValue) then
    Log('Registrato con successo in: ' + KeyName)
  else
    Log('Errore durante la scrittura nel registro.');
    
end;

