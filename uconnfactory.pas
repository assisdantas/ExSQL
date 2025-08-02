unit uconnfactory;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, SQLDB, SQLDBLib, DB, MSSQLConn, PQConnection, OracleConnection, ODBCConn, SQLite3Conn, IBConnection,
  mysql80conn, Contnrs, Dialogs;

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

{ TConnectionManager }

TConnectionManager = class
private
  FConnections: TFPHashObjectList;
  class var FInstance: TConnectionManager;  // Variável estática para a instância única
public
  constructor Create;
  destructor Destroy; override;
  function GetOrCreateConnection(Info: TConnectionInfo): TCustomConnection;
  procedure CloseAll;
end;

function GenerateUniqueKey(aName: String): String;

implementation

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
end;

destructor TConnectionManager.Destroy;
begin
  if Assigned(FConnections) then
  begin
    CloseAll;
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

  Key := GenerateUniqueKey(Info.aName);
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

  end
  else if Result is TIBConnection then
  begin
    TIBConnection(Result).HostName := Info.aHost;
    TIBConnection(Result).DatabaseName := Info.aDatabase;
    TIBConnection(Result).UserName := Info.aUser;
    TIBConnection(Result).Password := Info.aPassword;
    TIBConnection(Result).Port := Info.aPort;
  end
  else if Result is TPQConnection then
  begin

  end
  else if Result is TMySQL80Connection then
  begin

  end
  else if Result is TOracleConnection then
  begin

  end
  else if Result is TODBCConnection then
  begin

  end;

  try
    FConnections.Add(Key, Result);
  except
    on E: Exception do
    begin
      Result.Free;
      raise Exception.Create('Falha ao adicionar conexão: ' + E.Message);
    end;
  end;
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

