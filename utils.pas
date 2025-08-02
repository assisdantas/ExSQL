unit utils;

{$mode ObjFPC}{$H+}

interface

uses
  Forms, Classes, SysUtils, Dialogs, LCLType, StdCtrls, ComCtrls, SQLDB, SQLDBLib, SQLite3Conn, StrUtils, BCrypt;

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
function CompareHashPassword(aPassword, aHash: String): Boolean;

var
  MainConn: TSQLite3Connection;
  LibLoader: TSQLDBLibraryLoader;
  MainTrans: TSQLTransaction;

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
  for i := 0 to frmMain.tvwConnection.Items.Count - 1 do
    TObject(frmMain.tvwConnection.Items[i]).Free;
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
begin
  Result := TBCrypt.GenerateHash(aPassword);
end;

function CompareHashPassword(aPassword, aHash: String): Boolean;
begin
  Result := TBcrypt.CompareHash(aPassword, aHash);
end;


end.

