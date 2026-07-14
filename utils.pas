unit utils;

{$mode ObjFPC}{$H+}

interface

uses
  Forms, Classes, SysUtils, Dialogs, LCLType, StdCtrls, ComCtrls, SQLDB, SQLDBLib, SQLite3Conn, StrUtils,
  DCPrijndael, DCPsha256;

const
  { Chave usada apenas para ofuscar a senha guardada em connections.db contra leitura casual do
    arquivo. Não é segredo de nível cofre: quem tem acesso ao binário tem acesso à chave. }
  PASSWORD_CIPHER_KEY = 'ExSQL-Connection-Password-Key-v1';

procedure CreateMainConnection;
procedure CreateMainLibLoader;
procedure CreateMainTransaction;
procedure DestroyObject(aObject: TObject);
procedure LoadDatabaseTypes(aList: TComboBox);
procedure LoadConnections;
procedure ClearConnections;
function GetCurrentDir: String;
function GetPlatform: String;
function GetDateTime: String;
function EncryptPassword(aPassword: String): String;
function DecryptPassword(aPassword: String): String;
procedure LogError(const aMessage: String);
procedure EnsureHistoryTable;
procedure AddHistoryEntry(const aConnName, aSQLText: String);
procedure LoadHistoryEntries(const aConnName: String; aList: TStrings; aMaxRows: Integer = 100);
function ExtractSQLFromHistoryItem(const aItem: String): String;

type
  { Application.OnException exige "procedure(...) of object"; um procedure
    solto não é compatível, daí este objeto-wrapper só para expor o método. }
  TAppExceptionHandler = class
    procedure Handle(Sender: TObject; E: Exception);
  end;

var
  MainConn: TSQLite3Connection;
  LibLoader: TSQLDBLibraryLoader;
  MainTrans: TSQLTransaction;
  AppExceptionHandler: TAppExceptionHandler;

implementation

uses
  uconnfactory, umain;

procedure CreateMainConnection;
begin
  CreateMainLibLoader;
  CreateMainTransaction;

  MainConn := TSQLite3Connection.Create(nil);

  with MainConn do
  begin
    MainConn.CharSet := 'UTF8';
    MainConn.DatabaseName := GetCurrentDir + 'data\connections.db';

    try
      MainTrans.DataBase := MainConn;
      MainConn.Connected := True;
      EnsureHistoryTable;
    except
      on E: Exception do
      begin
        MessageDlg('Obter dados',
          'Ocorreu um erro ao obter dados da aplicação."'+E.Message+'"',
          mtError, [mbOk], 0, mbOk);
      end;
    end;
  end;
end;

procedure EnsureHistoryTable;
var
  Query: TSQLQuery;
begin
  Query := TSQLQuery.Create(nil);
  Query.DataBase := MainConn;
  Query.Transaction := MainTrans;
  try
    Query.SQL.Text := 'CREATE TABLE IF NOT EXISTS HISTORY (' +
      'CODHIST INTEGER PRIMARY KEY AUTOINCREMENT, ' +
      'CONNNAME TEXT, ' +
      'SQLTEXT TEXT, ' +
      'EXECDATE TEXT)';
    Query.ExecSQL;
    MainTrans.CommitRetaining;
  finally
    FreeAndNil(Query);
  end;
end;

procedure AddHistoryEntry(const aConnName, aSQLText: String);
var
  Query: TSQLQuery;
begin
  { Histórico é conveniência, não deve nunca impedir a execução da query em
    si — qualquer falha aqui (ex.: tabela ainda não criada em um DB antigo)
    é apenas registrada em log, nunca propagada. }
  try
    Query := TSQLQuery.Create(nil);
    Query.DataBase := MainConn;
    Query.Transaction := MainTrans;
    try
      Query.SQL.Text := 'INSERT INTO HISTORY (CONNNAME, SQLTEXT, EXECDATE) VALUES (:CONNNAME, :SQLTEXT, :EXECDATE)';
      Query.ParamByName('CONNNAME').AsString := aConnName;
      Query.ParamByName('SQLTEXT').AsString := aSQLText;
      Query.ParamByName('EXECDATE').AsString := GetDateTime;
      Query.ExecSQL;
      MainTrans.CommitRetaining;
    finally
      FreeAndNil(Query);
    end;
  except
    on E: Exception do
      LogError('Falha ao gravar histórico de query: ' + E.Message);
  end;
end;

procedure LoadHistoryEntries(const aConnName: String; aList: TStrings; aMaxRows: Integer);
var
  Query: TSQLQuery;
  SingleLineSQL: String;
begin
  { Cada item de aList é "<SQL numa linha só>    (<data/hora>)" — a marcação de
    data no final é só para exibição; TfrmHistory remove esse sufixo antes de
    inserir o texto de volta no editor (ver ExtractSQLFromHistoryItem). }
  aList.Clear;
  Query := TSQLQuery.Create(nil);
  Query.DataBase := MainConn;
  Query.Transaction := MainTrans;
  try
    Query.SQL.Text := 'SELECT SQLTEXT, EXECDATE FROM HISTORY WHERE CONNNAME = :CONNNAME ' +
      'ORDER BY CODHIST DESC LIMIT :MAXROWS';
    Query.ParamByName('CONNNAME').AsString := aConnName;
    Query.ParamByName('MAXROWS').AsInteger := aMaxRows;
    Query.Open;
    try
      while not Query.EOF do
      begin
        SingleLineSQL := StringReplace(Trim(Query.FieldByName('SQLTEXT').AsString), LineEnding, ' ', [rfReplaceAll]);
        SingleLineSQL := StringReplace(SingleLineSQL, #10, ' ', [rfReplaceAll]);
        SingleLineSQL := StringReplace(SingleLineSQL, #13, ' ', [rfReplaceAll]);
        aList.Add(SingleLineSQL + '    (' + Query.FieldByName('EXECDATE').AsString + ')');
        Query.Next;
      end;
    finally
      Query.Close;
    end;
  finally
    FreeAndNil(Query);
  end;
end;

function ExtractSQLFromHistoryItem(const aItem: String): String;
var
  ParenPos: Integer;
begin
  ParenPos := LastDelimiter('(', aItem);
  if ParenPos > 0 then
    Result := Trim(Copy(aItem, 1, ParenPos - 1))
  else
    Result := aItem;
end;

procedure CreateMainLibLoader;
var
  CurrentDir, PlatformCPU: String;
begin
  CurrentDir := GetCurrentDir;
  PlatformCPU := GetPlatform;
  LibLoader := TSQLDBLibraryLoader.Create(nil);

  try
    LibLoader.ConnectionType := 'sqlite3';
    LibLoader.LibraryName := CurrentDir + PlatformCPU + 'sqlite3.dll';
    LibLoader.Enabled := True;
    LibLoader.LoadLibrary;
  except
    on E: Exception do
    begin
      MessageDlg('ExSQL',
        'Não foi possível carregar a biblioteca de link dinâmico do SQLite. A reinstalação do programa pode resolver '+
        'o problema. "'+E.Message+'"', mtError, [mbOk], 0, mbok);
      Application.Terminate;
      Exit;
    end;
  end;
end;

procedure CreateMainTransaction;
begin
  MainTrans := TSQLTransaction.Create(nil);
end;

procedure DestroyObject(aObject: TObject);
begin
  FreeAndNil(aObject);
end;

procedure LoadDatabaseTypes(aList: TComboBox);
var
  Query: TSQLQuery;
begin
  Query := TSQLQuery.Create(nil);
  Query.DataBase := MainConn;
  Query.Transaction := MainTrans;

  try
    with Query do
    begin
      Close;
      SQL.Clear;
      SQL.Text := 'SELECT CODTYPE, NAME FROM TYPE';

      try
        Open;

        aList.Items.Clear;
        if (not IsEmpty) then
        begin
          while (not EOF) do
          begin
            aList.Items.AddObject(
              FieldByName('NAME').AsString,
              TObject(PtrInt(FieldByName('CODTYPE').AsInteger))
            );
            Next;
          end;
        end;
      except
        on E: Exception do
        begin
          MessageDlg('Obter dados',
            'Ocorreu um erro ao obter dados do aplicativo. "'+E.Message+'"',
            mtError, [mbOk], 0, mbOk);
        end;
      end;
    end;
  finally
    FreeAndNil(Query);
  end;
end;

procedure LoadConnections;
var
  Query: TSQLQuery;
  Node: TTreeNode;
  Info: TConnectionInfo;
begin
  ClearConnections;

  Query := TSQLQuery.Create(nil);
  Query.DataBase := MainConn;
  Query.Transaction := MainTrans;

  try
    try
      with Query do
      begin
        SQL.Text := ('SELECT * FROM CONNECTION');

        Open;

        if IsEmpty then Exit;

        while (not EOF) do
        begin
          Info := TConnectionInfo.Create;

          Info.aCodConn := FieldByName('CODCONN').AsInteger;
          Info.aName := FieldByName('NAME').AsString;
          Info.aDatabase := FieldByName('DATABASE').AsString;
          Info.aHost := FieldByName('HOST').AsString;
          Info.aPort := FieldByName('PORT').AsInteger;
          Info.aLibrary := FieldByName('LIBRARY').AsString;
          Info.aCharset := FieldByName('CHARSET').AsString;
          Info.aUser := FieldByName('USER').AsString;
          Info.aPassword := FieldByName('PASSWORD').AsString;
          Info.aCodType := FieldByName('CODTYPE').AsInteger;

          Node := frmMain.tvwConnection.Items.Add(nil, Info.aName);
          Node.Data := Info;

          Next;
        end;

        Close;
      end;
    except
      on E: Exception do
      begin
        MessageDlg('ExSQL',
          'Ocorreu um erro ao tentar obter as conexões configuradas. "'+E.Message+'"',
          mtError, [mbOk], 0, mbOK);
      end;
    end;
  finally
    FreeAndNil(Query);
  end;
end;

procedure ClearConnections;
var
  i: Integer;
begin
  { Libera o TConnectionInfo (Data) anexado a cada nó; os nós em si são
    destruídos de uma vez por Items.Clear logo abaixo. Fazer Free em
    Items[i] diretamente (ao invés de Items[i].Data) destruiria os nós um a
    um durante a iteração, encolhendo Items.Count e estourando o índice. }
  for i := 0 to frmMain.tvwConnection.Items.Count - 1 do
    TObject(frmMain.tvwConnection.Items[i].Data).Free;
  frmMain.tvwConnection.Items.Clear;
end;

function GetCurrentDir: String;
begin
  Result := ExtractFilePath(Application.ExeName);
end;

function GetPlatform: String;
var
  Platform: String;
begin
  Platform := StrUtils.IfThen((SizeOf(Pointer) = 8), 'x64', 'x86');
  Result := 'sqlite\'+PlatForm+'\';

  frmMain.Caption := 'ExSQL '+PlatForm;
end;

function GetDateTime: String;
begin
  Result := FormatDateTime('dd-mm-yyyy hh:mm:ss', Now);
end;

function EncryptPassword(aPassword: String): String;
var
  Cipher: TDCP_rijndael;
begin
  if aPassword = '' then Exit('');

  Cipher := TDCP_rijndael.Create(nil);
  try
    Cipher.InitStr(PASSWORD_CIPHER_KEY, TDCP_sha256);
    Result := Cipher.EncryptString(aPassword);
  finally
    Cipher.Burn;
    Cipher.Free;
  end;
end;

function DecryptPassword(aPassword: String): String;
var
  Cipher: TDCP_rijndael;
begin
  if aPassword = '' then Exit('');

  Cipher := TDCP_rijndael.Create(nil);
  try
    Cipher.InitStr(PASSWORD_CIPHER_KEY, TDCP_sha256);
    Result := Cipher.DecryptString(aPassword);
  finally
    Cipher.Burn;
    Cipher.Free;
  end;
end;

procedure LogError(const aMessage: String);
var
  LogFile: TextFile;
  LogPath: String;
begin
  LogPath := GetCurrentDir + 'logs';
  if not DirectoryExists(LogPath) then
    ForceDirectories(LogPath);

  AssignFile(LogFile, LogPath + '\error.log');
  try
    if FileExists(LogPath + '\error.log') then
      Append(LogFile)
    else
      Rewrite(LogFile);

    WriteLn(LogFile, '[' + GetDateTime + '] ' + aMessage);
  finally
    CloseFile(LogFile);
  end;
end;

procedure TAppExceptionHandler.Handle(Sender: TObject; E: Exception);
begin
  { Registra em arquivo antes de exibir ao usuário, para permitir diagnóstico
    pós-morte de erros que o usuário não saiba/consiga descrever. }
  try
    LogError('Exceção não tratada: ' + E.ClassName + ': ' + E.Message);
  except
    { Se nem o log funcionar (ex.: disco cheio, sem permissão), não deixar
      isso silenciar a mensagem de erro original para o usuário. }
  end;

  MessageDlg('ExSQL',
    'Ocorreu um erro inesperado: "' + E.Message + '".' + LineEnding +
    'Detalhes foram registrados em logs\error.log.',
    mtError, [mbOk], 0, mbOk);
end;

initialization
  AppExceptionHandler := TAppExceptionHandler.Create;

finalization
  AppExceptionHandler.Free;

end.

