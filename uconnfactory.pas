unit uconnfactory;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, SQLDB, SQLDBLib, DB, MSSQLConn, PQConnection, OracleConnection, ODBCConn, SQLite3Conn, IBConnection,
  mysql80conn, Contnrs, Dialogs, Generics.Collections;

type

{ TDBType }

TDBType = (
  dbMSSQL = 1,
  dbFirebird = 2,
  dbPostgres = 3,
  dbMySQL = 4,
  dbMariaDB = 5,
  dbSQLite3 = 6,
  dbOracle = 7,
  dbODBC = 8
);

{ TConnectionInfo }

TConnectionInfo = class
  public
    aCodConn, aCodType, aPort: Integer;
    aName, aHost, aDatabase, aLibrary, aCharset, aUser, aPassword: String;
end;

{ Dictionary }
TDict = specialize TDictionary<String, TCustomConnection>;

{ TConnectionManager }

TConnectionManager = class
private
  FConnections: TFPHashObjectList;
  FActiveConnections: TDict;
public
  constructor Create;
  destructor Destroy; override;

  function GetOrCreateConnection(Info: TConnectionInfo): TCustomConnection;
  function DisconnectByName(ConnectionName: String): Boolean;
  function GetActiveConnection(ConnectionName: String): TCustomConnection;
  function IsConnected(ConnectionName: String): Boolean;
  procedure CloseAll;

  property DictActiveConnections: TDict read FActiveConnections;
end;

function GenerateUniqueKey(aName: String): String;

var
  GlobalConnManager: TConnectionManager;

implementation

uses
  utils;

function GenerateUniqueKey(aName: String): String;
var
  GUID: TGUID;
begin
  CreateGUID(GUID);
  Result := aName + GUIDToString(GUID);
end;

{ TConnectionManager }

constructor TConnectionManager.Create;
begin
  inherited;
  FConnections := TFPHashObjectList.Create(True);
  FActiveConnections := TDict.Create;
end;

destructor TConnectionManager.Destroy;
begin
  if Assigned(FConnections) then
  begin
    CloseAll;
    FreeAndNil(FActiveConnections);
    FreeAndNil(FConnections);
  end;
  inherited Destroy;
end;

function TConnectionManager.GetOrCreateConnection(Info: TConnectionInfo
  ): TCustomConnection;
var
  Key: String;
  ExistingConn: TObject;
begin
  if not Assigned(FConnections) then
    raise Exception.Create('Gerenciador de conexões não inicializado');

  Key := Info.aName;

  if FActiveConnections.ContainsKey(Key) then
    Exit(FActiveConnections.Items[Key]);

  ExistingConn := FConnections.Find(Key);

  if Assigned(ExistingConn) then
  begin
    if ExistingConn is TCustomConnection then
      Exit(TCustomConnection(ExistingConn))
    else
      raise Exception.Create('Objeto de conexão inválido encontrado');
  end;

  case Info.aCodType of
    Ord(dbMSSQL):
      Result := TMSSQLConnection.Create(nil);
    Ord(dbFirebird):
      Result := TIBConnection.Create(nil);
    Ord(dbPostgres):
      Result := TPQConnection.Create(nil);
    Ord(dbMySQL), Ord(dbMariaDB): // MariaDB também usa TMySQL80Connection
      Result := TMySQL80Connection.Create(nil);
    Ord(dbSQLite3):
      Result := TSQLite3Connection.Create(nil);
    Ord(dbOracle):
      Result := TOracleConnection.Create(nil);
    Ord(dbODBC):
      Result := TODBCConnection.Create(nil);
    else
      raise Exception.Create('Tipo de conexão não suportado.');
  end;

  if Result is TMSSQLConnection then
  begin
    { TMSSQLConnection não expõe uma propriedade Port própria; a porta é embutida
      no HostName no formato 'servidor:porta' (suportado nativamente pelo driver). }
    if Info.aPort > 0 then
      TMSSQLConnection(Result).HostName := Info.aHost + ':' + IntToStr(Info.aPort)
    else
      TMSSQLConnection(Result).HostName := Info.aHost;
    TMSSQLConnection(Result).DatabaseName := Info.aDatabase;
    TMSSQLConnection(Result).UserName := Info.aUser;
    TMSSQLConnection(Result).Password := DecryptPassword(Info.aPassword);
    if Info.aCharset <> '' then
      TMSSQLConnection(Result).CharSet := Info.aCharset;
  end
  else if Result is TIBConnection then
  begin
    TIBConnection(Result).HostName := Info.aHost;
    TIBConnection(Result).DatabaseName := Info.aDatabase;
    TIBConnection(Result).UserName := Info.aUser;
    TIBConnection(Result).Password := DecryptPassword(Info.aPassword);
    TIBConnection(Result).Port := Info.aPort;
    if Info.aCharset <> '' then
      TIBConnection(Result).CharSet := Info.aCharset;
  end
  else if Result is TPQConnection then
  begin
    { TPQConnection não expõe Port como propriedade pública; a porta é passada
      via Params, que é repassado como parâmetro de conexão bruto ao libpq. }
    TPQConnection(Result).HostName := Info.aHost;
    TPQConnection(Result).DatabaseName := Info.aDatabase;
    TPQConnection(Result).UserName := Info.aUser;
    TPQConnection(Result).Password := DecryptPassword(Info.aPassword);
    if Info.aCharset <> '' then
      TPQConnection(Result).CharSet := Info.aCharset;
    if Info.aPort > 0 then
      TPQConnection(Result).Params.Add('port=' + IntToStr(Info.aPort));
  end
  else if Result is TMySQL80Connection then
  begin
    TMySQL80Connection(Result).HostName := Info.aHost;
    TMySQL80Connection(Result).DatabaseName := Info.aDatabase;
    TMySQL80Connection(Result).UserName := Info.aUser;
    TMySQL80Connection(Result).Password := DecryptPassword(Info.aPassword);
    if Info.aCharset <> '' then
      TMySQL80Connection(Result).CharSet := Info.aCharset;
    if Info.aPort > 0 then
      TMySQL80Connection(Result).Port := Info.aPort;
  end
  else if Result is TOracleConnection then
  begin
    { TOracleConnection monta a EZ connect string como '//HostName/DatabaseName';
      a porta, quando informada, entra no HostName como 'servidor:porta'. }
    if Info.aPort > 0 then
      TOracleConnection(Result).HostName := Info.aHost + ':' + IntToStr(Info.aPort)
    else
      TOracleConnection(Result).HostName := Info.aHost;
    TOracleConnection(Result).DatabaseName := Info.aDatabase;
    TOracleConnection(Result).UserName := Info.aUser;
    TOracleConnection(Result).Password := DecryptPassword(Info.aPassword);
  end
  else if Result is TODBCConnection then
  begin
    { TODBCConnection ignora HostName/Port: DatabaseName é o nome do DSN configurado
      no sistema (ODBC Data Source Administrator). }
    TODBCConnection(Result).DatabaseName := Info.aDatabase;
    TODBCConnection(Result).UserName := Info.aUser;
    TODBCConnection(Result).Password := DecryptPassword(Info.aPassword);
  end
  else if Result is TSQLite3Connection then
  begin
    TSQLite3Connection(Result).DatabaseName := Info.aDatabase;
    if Info.aCharset <> '' then
      TSQLite3Connection(Result).CharSet := Info.aCharset;
  end;

  try
    FConnections.Add(Key, Result);
    FActiveConnections.Add(Key, Result);
  except
    on E: Exception do
    begin
      Result.Free;
      raise Exception.Create('Falha ao adicionar conexão: ' + E.Message);
    end;
  end;
end;

function TConnectionManager.DisconnectByName(ConnectionName: String): Boolean;
var
  Conn: TCustomConnection;
  i: Integer;
begin
  Result := False;
  if FActiveConnections.TryGetValue(ConnectionName, Conn) then
  begin
    if Conn.Connected then
      Conn.Connected := False;
    FActiveConnections.Remove(ConnectionName);

    i := FConnections.FindIndexOf(ConnectionName);
    if i >= 0 then
      FConnections.Delete(i);

    Result := True;
  end;
end;

function TConnectionManager.GetActiveConnection(ConnectionName: String
  ): TCustomConnection;
begin
  if not FActiveConnections.TryGetValue(ConnectionName, Result) then
    Result := nil;
end;

function TConnectionManager.IsConnected(ConnectionName: String): Boolean;
var
  Conn: TCustomConnection;
begin
  Conn := GetActiveConnection(ConnectionName);
  Result := Assigned(Conn) and Conn.Connected;
end;

procedure TConnectionManager.CloseAll;
var
  i: Integer;
begin
  if Assigned(FConnections) then
  begin
    for i := 0 to FConnections.Count - 1 do
    begin
      if FConnections.Items[i] is TCustomConnection then
      begin
        if TCustomConnection(FConnections.Items[i]).Connected then
          TCustomConnection(FConnections.Items[i]).Connected := False;
      end;
    end;
    FConnections.Clear; // Remove todas as conexões
  end;
end;

end.
