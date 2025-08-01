unit utils;

{$mode ObjFPC}{$H+}

interface

uses
  Forms, Classes, SysUtils, Dialogs, LCLType, StdCtrls, SQLDB, SQLDBLib, SQLite3Conn;

procedure CreateMainConnection;
procedure CreateMainLibLoader;
procedure CreateMainTransaction;
procedure DestroyObject(aObject: TObject);
procedure LoadDatabaseTypes(aList: TComboBox);
function GetCurrentDir: String;

var
  MainConn: TSQLite3Connection;
  LibLoader: TSQLDBLibraryLoader;
  MainTrans: TSQLTransaction;

implementation

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
  CurrentDir: String;
begin
  CurrentDir := GetCurrentDir;
  LibLoader := TSQLDBLibraryLoader.Create(nil);

  LibLoader.ConnectionType := 'sqlite3';
  LibLoader.LibraryName := CurrentDir + 'sqlite3.dll';
  LibLoader.Enabled := True;
  LibLoader.LoadLibrary;
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
  i: Integer;
  Query: TSQLQuery;
begin
  Query := TSQLQuery.Create(nil);
  Query.DataBase := MainConn;
  Query.Transaction := MainTrans;

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
            TObject(Pointer(FieldByName('CODTYPE').AsInteger))
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
end;

function GetCurrentDir: String;
begin
  Result := ExtractFilePath(Application.ExeName);
end;

end.

