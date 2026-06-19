// Funzione per trovare l'ultima occorrenza di un carattere in una stringa
function LastPos(SubStr: String; S: String): Integer;
var
  i: Integer;
begin
  Result := 0;
  for i := Length(S) downto 1 do
  begin
    if Copy(S, i, Length(SubStr)) = SubStr then
    begin
      Result := i;
      Break;
    end;
  end;
end;


// Funzioni di controllo per la sezione [Files]
function IsExcel32(): Boolean; begin Result := (ExcelBit = '32'); end;
function IsExcel64(): Boolean; begin Result := (ExcelBit = '64'); end;


// --- FUNZIONE DI SUPPORTO: CONFRONTO VERSIONI (es. "6.0.33" vs "6.0.10") ---
// Ritorna: 1 se V1 > V2, 0 se V1 = V2, -1 se V1 < V2
function CompareVersions(V1, V2: String): Integer;
var
  P, V1Part, V2Part: Integer;
begin
  Result := 0;
  while (Result = 0) and ((V1 <> '') or (V2 <> '')) do
  begin
    P := Pos('.', V1);
    if P > 0 then begin
      V1Part := StrToIntDef(Copy(V1, 1, P - 1), 0);
      Delete(V1, 1, P);
    end else begin
      V1Part := StrToIntDef(V1, 0);
      V1 := '';
    end;

    P := Pos('.', V2);
    if P > 0 then begin
      V2Part := StrToIntDef(Copy(V2, 1, P - 1), 0);
      Delete(V2, 1, P);
    end else begin
      V2Part := StrToIntDef(V2, 0);
      V2 := '';
    end;

    if V1Part > V2Part then Result := 1
    else if V1Part < V2Part then Result := -1;
  end;
end;


function PreferArm64Files: Boolean;
begin
  Result := IsArm64;
end;

function PreferX64Files: Boolean;
begin
  Result := not PreferArm64Files and IsX64Compatible;
end;

function PreferX86Files: Boolean;
begin
  Result := not PreferArm64Files and not PreferX64Files;
end;

// --- 4. EVENTI INSTALLAZIONE ---
function IsExcelRunning(): Boolean;
var
  Locator, WmiService, ObjectsList: Variant;
begin
  Result := False;
  try
    // 1. Crea il Locator
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    
    // 2. Connetti al server (spezzando la chiamata l'identificatore viene riconosciuto)
    WmiService := Locator.ConnectServer('.', 'root\CIMV2');
    
    // 3. Esegui la query
    ObjectsList := WmiService.ExecQuery('SELECT Name FROM Win32_Process WHERE Name = "excel.exe"');
    
    // 4. Verifica il conteggio dei risultati
    if not VarIsClear(ObjectsList) then
    begin
      Result := (ObjectsList.Count > 0);
    end;
  except
    Log('WMI Error: ' + GetExceptionMessage);
    Result := False;
  end;
end;

function GetNetMajorVersion(): String;
var
  FullVersion: String;
  DotPos: Integer;
begin
  // Recuperiamo la stringa "11.0" definita sopra
  FullVersion := '{#NETVersion}'; 
  
  // Cerchiamo la posizione del punto
  DotPos := Pos('.', FullVersion);
  
  if DotPos > 0 then
    // Estraiamo tutto ciò che c'è prima del punto
    Result := Copy(FullVersion, 1, DotPos - 1)
  else
    // Se non c'è un punto, usiamo l'intera stringa come fallback
    Result := FullVersion;
end;

function InstallDotNetViaWinget(): Boolean;
var
  ResultCode: Integer;
  Params,MajorVersion: String;
begin
  MajorVersion := GetNetMajorVersion();
  // Verifichiamo se l'architettura è 64 bit per specificare il parametro corretto
  // Se ExcelBit è 64, forziamo l'architettura x64, altrimenti x86
  Params := Format('install Microsoft.DotNet.DesktopRuntime.%s --silent --accept-package-agreements --accept-source-agreements',[MajorVersion]);
  
  if ExcelBit = '64' then
    Params := Params + ' --architecture x64'
  else
    Params := Params + ' --architecture x86';

  // Eseguiamo winget.exe (usiamo cmd /c per sicurezza nel path)
  Log('Esecuzione comando: winget ' + Params);
  
  // Utilizziamo ewWaitUntilTerminated per assicurarci che finisca prima di proseguire
  if Exec(ExpandConstant('{cmd}'), '/C winget ' + Params, '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
    begin
      Log('Winget terminato con codice: ' + IntToStr(ResultCode));
      // 0 = Successo
case ResultCode of
      0: begin
        Log('Installazione completata con successo.');
        Result := True;
      end;
      -1978335189: begin // 0x8A15002B
        Log('Winget segnala che il pacchetto è già installato. Proseguo.');
        Result := True; 
      end;
      3010: begin
        Log('Installazione riuscita, riavvio richiesto.');
        // Inno Setup non riavvia subito automaticamente qui, 
        // ma segna che è andato tutto bene.
        RestartRequired := True;
        Result := True; 
      end;
    else
      begin
        Log('Errore Winget non gestito: ' + IntToStr(ResultCode));
        Result := False;
      end;
    end;
  end
  else
    begin
      Log('Impossibile eseguire cmd/winget.');
      Result := False;
    end;
end;