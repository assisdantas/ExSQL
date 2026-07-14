program project1;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}
  cthreads,
  {$ENDIF}
  {$IFDEF HASAMIGA}
  athreads,
  {$ENDIF}
  Interfaces, SysUtils, Dialogs, Classes, PairSplitter, // this includes the LCL widgetset
  Forms, umain, ucreateconn, utils, usplash, uconnfactory, uabout
  { you can add units after this };

{$R *.res}

var
  SplashScreen: TfrmSplash;

begin
  try
    { TPairSplitterSide é usada pelo PairSplitter da tela principal e persistida no .lfm,
      mas só é registrada para streaming em tempo de design (dentro de "procedure Register"
      do pacote LCL). Sem este registro explícito, o carregamento do formulário falha em
      tempo de execução com "Class TPairSplitterSide not found". }
    RegisterClass(TPairSplitterSide);
  RequireDerivedFormResource := True;
  Application.Scaled := True;
  {$PUSH}{$WARN 5044 OFF}
  Application.MainFormOnTaskbar := True;
  {$POP}
  Application.Initialize;
  Application.OnException := @AppExceptionHandler.Handle;

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
  except
    on E: Exception do
      MessageDlg('ExSQL',
        'Ocorreu um erro inesperado ao iniciar a aplicação: "'+E.Message+'".',
        mtError, [mbOk], 0, mbOk);
  end;
end.

