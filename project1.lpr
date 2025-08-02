program project1;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}
  cthreads,
  {$ENDIF}
  {$IFDEF HASAMIGA}
  athreads,
  {$ENDIF}
  Interfaces, SysUtils, // this includes the LCL widgetset
  Forms, umain, ucreateconn, utils, usplash, uconnfactory
  { you can add units after this };

{$R *.res}

var
  SplashScreen: TfrmSplash;

begin
  RequireDerivedFormResource := True;
  Application.Scaled := True;
  {$PUSH}{$WARN 5044 OFF}
  Application.MainFormOnTaskbar := True;
  {$POP}
  Application.Initialize;

  SplashScreen := TfrmSplash.Create(Application);
  SplashScreen.Show;

  { Criar formulários }
  Application.CreateForm(TfrmMain, frmMain);
  Application.CreateForm(TfrmCreateConn, frmCreateConn);
  //Application.CreateForm(TfrmSplash, frmSplash);

  { Executar funções e procedimentos básicos }
  CreateMainConnection;
  LoadConnections;
  LoadDatabaseTypes(frmCreateConn.cbbType);

  SplashScreen.Refresh;
  SplashScreen.Hide;
  FreeAndNil(SplashScreen);

  Application.Run;
end.

