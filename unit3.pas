unit Unit3;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, SQLDB, SQLDBLib, SQLite3Conn;

procedure CreateMainConnection;

var
  MainConn: TSQLite3Connection;
  LibLoader: TSQLDBLibraryLoader;

implementation

procedure CreateMainConnection;
begin
  LibLoader := TSQLDBLibraryLoader.Create(nil);
  MainConn := TSQLite3Connection.Create(nil);
end;

end.

