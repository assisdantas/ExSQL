unit umain;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, SQLDBLib, SQLDB, IBConnection, SQLite3Conn, DB,
  Forms, Controls, Graphics, Dialogs, StdCtrls, DBGrids,
  ComCtrls, ExtCtrls, PairSplitter, Buttons, lazutf8, SynEdit,
  SynHighlighterSQL, StrUtils, RegExpr, Windows, SynEditTypes, SynEditKeyCmds,
  SynCompletion, LCLType, Menus, ActnList, Types, MSSQLConn, PQConnection,
  OracleConnection, ODBCConn, mysql80conn, Grids, uconnfactory, IniFiles, TypInfo;

type

  TSQLExecThread = class;

  { TfrmMain }

  TfrmMain = class(TForm)
    actExecuteSQL: TAction;
    actCommit: TAction;
    actConnect: TAction;
    actDisconnect: TAction;
    actEditSQL: TAction;
    actScriptExecute: TAction;
    actOpen: TAction;
    actCopyConn: TAction;
    actSave: TAction;
    actSaveAs: TAction;
    actClose: TAction;
    actNew: TAction;
    actPaste: TAction;
    actCopy: TAction;
    actCut: TAction;
    actRedo: TAction;
    actUndo: TAction;
    actRollback: TAction;
    actNewConn: TAction;
    actEditConn: TAction;
    actDeleteConn: TAction;
    actQueryHistory: TAction;
    actExportCSV: TAction;
    actCancelExecution: TAction;
    TimerExec: TTimer;
    ActionList1: TActionList;
    FontDialog1: TFontDialog;
    OpenDialog1: TOpenDialog;
    SaveDialog1: TSaveDialog;
    ImageList1: TImageList;
    ImageList2: TImageList;
    MainMenu1: TMainMenu;
    Memo1: TMemo;
    MenuItem1: TMenuItem;
    MenuItem10: TMenuItem;
    MenuItem11: TMenuItem;
    MenuItem12: TMenuItem;
    MenuItem13: TMenuItem;
    MenuItem14: TMenuItem;
    MenuItem15: TMenuItem;
    MenuItem16: TMenuItem;
    MenuItem17: TMenuItem;
    MenuItem18: TMenuItem;
    MenuItem19: TMenuItem;
    MenuItem20: TMenuItem;
    MenuItem21: TMenuItem;
    MenuItem22: TMenuItem;
    MenuItem23: TMenuItem;
    MenuItem24: TMenuItem;
    MenuItem25: TMenuItem;
    MenuItem26: TMenuItem;
    MenuItem27: TMenuItem;
    MenuItem28: TMenuItem;
    MenuItem29: TMenuItem;
    MenuItem30: TMenuItem;
    MenuItem31: TMenuItem;
    MenuItem32: TMenuItem;
    PageControl1: TPageControl;
    Panel1: TPanel;
    Panel3: TPanel;
    Separator6: TMenuItem;
    PopupMenu2: TPopupMenu;
    Separator5: TMenuItem;
    Separator4: TMenuItem;
    Separator3: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem4: TMenuItem;
    MenuItem5: TMenuItem;
    MenuItem6: TMenuItem;
    MenuItem7: TMenuItem;
    MenuItem8: TMenuItem;
    MenuItem9: TMenuItem;
    Separator2: TMenuItem;
    Separator1: TMenuItem;
    PairSplitter2: TPairSplitter;
    Panel4: TPanel;
    PopupMenu1: TPopupMenu;
    Separator7: TMenuItem;
    StatusBar1: TStatusBar;
    SynCompletion1: TSynCompletion;
    SynSQLSyn1: TSynSQLSyn;
    Timer1: TTimer;
    Timer2: TTimer;
    ToolBar1: TToolBar;
    ToolBar2: TToolBar;
    ToolButton1: TToolButton;
    ToolButton10: TToolButton;
    ToolButton11: TToolButton;
    ToolButton12: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    ToolButton6: TToolButton;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    btnConnect: TToolButton;
    ToolButton9: TToolButton;
    tvwConnection: TTreeView;
    procedure actCommitExecute(Sender: TObject);
    procedure actConnectExecute(Sender: TObject);
    procedure actCopyConnExecute(Sender: TObject);
    procedure actDeleteConnExecute(Sender: TObject);
    procedure actDisconnectExecute(Sender: TObject);
    procedure actEditConnExecute(Sender: TObject);
    procedure actEditSQLExecute(Sender: TObject);
    procedure actExecuteSQLExecute(Sender: TObject);
    procedure actNewConnExecute(Sender: TObject);
    procedure actRollbackExecute(Sender: TObject);
    procedure actScriptExecuteExecute(Sender: TObject);
    procedure actNewExecute(Sender: TObject);
    procedure actCloseExecute(Sender: TObject);
    procedure actOpenExecute(Sender: TObject);
    procedure actSaveExecute(Sender: TObject);
    procedure actSaveAsExecute(Sender: TObject);
    procedure actUndoExecute(Sender: TObject);
    procedure actRedoExecute(Sender: TObject);
    procedure actCopyExecute(Sender: TObject);
    procedure actPasteExecute(Sender: TObject);
    procedure actCutExecute(Sender: TObject);
    procedure actQueryHistoryExecute(Sender: TObject);
    procedure actExportCSVExecute(Sender: TObject);
    procedure actCancelExecutionExecute(Sender: TObject);
    procedure TimerExecTimer(Sender: TObject);
    procedure tvwConnectionExpanding(Sender: TObject; Node: TTreeNode; var AllowExpansion: Boolean);
    procedure DBGrid1TitleClick(Column: TColumn);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure MenuItem14Click(Sender: TObject);
    procedure MenuItem15Click(Sender: TObject);
    procedure SQLQuery1AfterOpen(DataSet: TDataSet);
    procedure SQLQuery1BeforeOpen(DataSet: TDataSet);
    procedure SynEdit1Click(Sender: TObject);
    procedure SynEdit1CommandProcessed(Sender: TObject; var Command: TSynEditorCommand; var AChar: TUTF8Char;
      Data: pointer);
    procedure SynEdit1KeyPress(Sender: TObject; var Key: char);
    procedure SynEdit1StatusChange(Sender: TObject; Changes: TSynStatusChanges);
    procedure Timer1Timer(Sender: TObject);
    procedure Timer2Timer(Sender: TObject);
    procedure tvwConnectionDblClick(Sender: TObject);
    procedure tvwConnectionSelectionChanged(Sender: TObject);
    procedure PageControl1Change(Sender: TObject);
    // outras procedures...
    procedure ExPairSplitterEditorRiseze(Sender: Tobject);
  private
    LastWord: string;
    DoOpen: Boolean;
    FBaseCompletionWords: TStringList;
    { Estado da execução em background (ver TSQLExecThread) — só uma execução
      por vez em todo o app, já que Memo1/Panel3/StatusBar1 são compartilhados
      entre abas (não são per-tab), então mesmo antes desta mudança nunca
      houve suporte real a duas execuções simultâneas visíveis. }
    FExecThread: TSQLExecThread;
    FExecTab: TTabSheet;
    FExecQuery: TSQLQuery;
    FExecEditor: TSynEdit;
    FExecConnName: String;
    FExecConn: TDatabase;
    FExecStartCount: Int64;
    FExecCancelRequested: Boolean;
    procedure RefreshSchemaAutocomplete;
    procedure PopulateSchemaNodes(ConnNode: TTreeNode; Info: TConnectionInfo);
    function GetLastStatementFromScript(const Script: String): String;
    procedure PrepareExecSQL(Which: String);
    procedure ExecutionThreadDone;
    function FriendlyFieldTypeName(aType: TFieldType): String;
    procedure SetColumnTitlesWithTypes(aTab: TTabSheet; aQuery: TSQLQuery);
    procedure SelectSQLBlock(aSynEdit: TSynEdit);
    procedure TitleOrderUpdate(aGrid: TDBGrid; aField, aDirection: String);
    procedure CreateTabEdit(ConnName: String; ConnType: Integer; Conn: TCustomConnection);
    procedure VerifyConnStatus;
    procedure UpdateActionsForActiveTab;
    function FindConnectionInfoByName(const aName: String): TConnectionInfo;
    function IsSelectSQL(const SQLText: string): Boolean;
    function GetSQLBlockAtCursor(SynEdit: TSynEdit): string;
    function GetCurrentWordAtCursor(aSynEdit: TSynEdit): string;
    function GetActiveSynEdit: TSynEdit;
    function OpenExec(Query: TSQLQuery; aSQL: String): String;
    function RemoveOrderBy(const aSQL: String): String;
    function GetCurrentOrderBy(const aSQL: String; out AField, ADirection: String): Boolean;
    function ConType(aCodType: Integer): Integer;
    procedure SaveWindowLayout;
    procedure LoadWindowLayout;
  public

  end;

  { TSQLExecThread }

  { Executa a query/script em background para permitir cancelar e mostrar um
    cronômetro sem travar a UI (antes disso, Query.Open/ExecuteScript rodavam
    direto na thread principal, bloqueando a aplicação inteira até terminar).
    Só toca em objetos de banco de dados (TSQLQuery/TSQLScript/TSQLTransaction)
    aqui — nenhum controle visual é acessado fora de Synchronize/da thread
    principal, já que LCL não é thread-safe. }
  TSQLExecThread = class(TThread)
  private
    FForm: TfrmMain;
    FQuery: TSQLQuery;
    FScript: TSQLScript;
    FTrans: TSQLTransaction;
    FWhich: String;
    FSQL: String;
    FScriptText: String;
    FHistoryText: String;
    FErrorMessage: String;
    FRowsAffected: Integer;
    FHasError: Boolean;
    FElapsedNs: Int64;
  protected
    procedure Execute; override;
  public
    constructor Create(aForm: TfrmMain; aQuery: TSQLQuery; aScript: TSQLScript; aTrans: TSQLTransaction;
      const aWhich, aSQL, aScriptText: String);
  end;

var
  frmMain: TfrmMain;
  UsuarioWT, SenhaDB, AliasDB, UsuarioDB, CodRotina: String;

implementation

{$R *.lfm}

uses
  ucreateconn, utils, uabout, uhistory;

constructor TSQLExecThread.Create(aForm: TfrmMain; aQuery: TSQLQuery; aScript: TSQLScript; aTrans: TSQLTransaction;
  const aWhich, aSQL, aScriptText: String);
begin
  inherited Create(True);
  FreeOnTerminate := True;
  FForm := aForm;
  FQuery := aQuery;
  FScript := aScript;
  FTrans := aTrans;
  FWhich := aWhich;
  FSQL := aSQL;
  FScriptText := aScriptText;
  if FWhich = 'query' then
    FHistoryText := aSQL
  else
    FHistoryText := aScriptText;
end;

procedure TSQLExecThread.Execute;
var
  StartCount, EndCount, Frequency: Int64;
  LastStatement: String;
begin
  FHasError := False;
  QueryPerformanceFrequency(Frequency);
  QueryPerformanceCounter(StartCount);
  try
    try
      if FWhich = 'query' then
      begin
        if FTrans.Active then
          FTrans.EndTransaction;
        FTrans.StartTransaction;

        FQuery.SQL.Text := FSQL;
        FForm.OpenExec(FQuery, FSQL);
        FRowsAffected := FQuery.RowsAffected;
      end
      else
      begin
        if FTrans.Active then
          FTrans.EndTransaction;
        FTrans.StartTransaction;

        FScript.Script.Text := FScriptText;
        FScript.ExecuteScript;
        FRowsAffected := -1;

        { Convenção comum em clientes SQL: se a última instrução do script for
          um SELECT, mostra o resultado na grid (reaproveitando a mesma
          transação já aberta pelo script, sem commitar/descartar nada). }
        LastStatement := FForm.GetLastStatementFromScript(FScriptText);
        if (LastStatement <> '') and FForm.IsSelectSQL(LastStatement) then
        begin
          FQuery.Close;
          FQuery.SQL.Text := LastStatement;
          FQuery.Open;
        end;
      end;
    except
      on E: Exception do
      begin
        FHasError := True;
        FErrorMessage := E.Message;
      end;
    end;
  finally
    QueryPerformanceCounter(EndCount);
    FElapsedNs := ((EndCount - StartCount) * 1000000000) div Frequency;
    Synchronize(@FForm.ExecutionThreadDone);
  end;
end;

{ TfrmMain }

procedure TfrmMain.FormResize(Sender: TObject);
begin
  { Panel3 (mensagens) usa Align = alBottom e PageControl1 usa Align = alClient;
    o LCL já redimensiona ambos automaticamente, sem necessidade de código aqui. }
end;

procedure TfrmMain.FormShow(Sender: TObject);
begin
  // Nada a fazer aqui: cada aba de editor recebe FontDialog1.Font ao ser criada em CreateTabEdit.
end;

procedure TfrmMain.MenuItem14Click(Sender: TObject);
var
  ActiveEditor: TSynEdit;
begin
  if not FontDialog1.Execute then Exit;

  ActiveEditor := GetActiveSynEdit;
  if Assigned(ActiveEditor) then
    ActiveEditor.Font := FontDialog1.Font;
end;

procedure TfrmMain.MenuItem15Click(Sender: TObject);
begin
  with TfrmAbout.Create(nil) do
  try
    ShowModal;
  finally
    Free;
  end;
end;

function TfrmMain.GetActiveSynEdit: TSynEdit;
begin
  Result := nil;
  if not Assigned(PageControl1.ActivePage) then Exit;
  Result := TSynEdit(PageControl1.ActivePage.FindComponent('TabSQLEditor'));
end;

procedure TfrmMain.SQLQuery1AfterOpen(DataSet: TDataSet);
begin
  Screen.Cursor := crDefault;
end;

procedure TfrmMain.SQLQuery1BeforeOpen(DataSet: TDataSet);
begin
  Screen.Cursor := crSQLWait;
end;

procedure TfrmMain.SynEdit1Click(Sender: TObject);
begin
  if Trim(TSynEdit(Sender).Text) <> '' then
  begin
    Timer1.Enabled := False;
    Timer1.Enabled := True;
  end;
end;

procedure TfrmMain.SynEdit1CommandProcessed(Sender: TObject; var Command: TSynEditorCommand; var AChar: TUTF8Char;
  Data: pointer);
begin
  {Timer1.Enabled := False;
  Timer1.Enabled := True;}
end;

procedure TfrmMain.SynEdit1KeyPress(Sender: TObject; var Key: char);
var
  CurrentWord: string;
begin
  // Ignora teclas que não formam palavras
  if not (Key in ['A'..'Z', 'a'..'z', '0'..'9', '_']) then
  begin
    Timer2.Enabled := False;
    Exit;
  end;

  // Obtém a palavra atual
  CurrentWord := GetCurrentWordAtCursor(TSynEdit(Sender));

  // Ativa o timer se a palavra tiver pelo menos 2 caracteres
  if Length(CurrentWord) >= 2 then
  begin
    LastWord := CurrentWord;
    Timer2.Enabled := False;
    Timer2.Enabled := True;
  end
  else
    Timer2.Enabled := False;
end;

procedure TfrmMain.SynEdit1StatusChange(Sender: TObject; Changes: TSynStatusChanges);
begin
  {if (scCaretX in Changes) or (scCaretY in Changes) then
  begin
    Timer1.Enabled := False;
    Timer1.Enabled := True;
  end;}
end;

procedure TfrmMain.Timer1Timer(Sender: TObject);
var
  ActiveEditor: TSynEdit;
begin
  Timer1.Enabled := False; // desativa para não repetir

  ActiveEditor := GetActiveSynEdit;
  if Assigned(ActiveEditor) then
    SelectSQLBlock(ActiveEditor);
end;

procedure TfrmMain.Timer2Timer(Sender: TObject);
var
  pt: TPoint;
  tokenRect: TRect;
  CurrentWord: string;
  ActiveEditor: TSynEdit;
begin
  Timer2.Enabled := False;

  ActiveEditor := GetActiveSynEdit;
  if not Assigned(ActiveEditor) then Exit;

  // Obtém a palavra atual novamente para garantir que ainda é relevante
  CurrentWord := GetCurrentWordAtCursor(ActiveEditor);

  if Length(CurrentWord) < 2 then Exit; // Mínimo de 2 caracteres

  // Obtém posição do cursor em pixels relativos ao SynEdit
  pt := ActiveEditor.RowColumnToPixels(ActiveEditor.CaretXY);

  // Converte para coordenadas da tela
  pt := ActiveEditor.ClientToScreen(pt);

  // Define o retângulo onde a janela de sugestões será exibida
  tokenRect.Left := pt.X;
  tokenRect.Top := pt.Y;
  tokenRect.Right := pt.X + 1; // largura mínima
  tokenRect.Bottom := pt.Y + ActiveEditor.LineHeight;

  // Executa a sugestão com a palavra atual
  SynCompletion1.Execute(CurrentWord, tokenRect);
end;

procedure TfrmMain.tvwConnectionDblClick(Sender: TObject);
begin
  btnConnect.Click;
end;

procedure TfrmMain.tvwConnectionSelectionChanged(Sender: TObject);
begin
  if not Assigned(tvwConnection.Selected) then
    Exit;

  if not Assigned(tvwConnection.Selected.Data) then
    Exit;
end;

procedure TfrmMain.PopulateSchemaNodes(ConnNode: TTreeNode; Info: TConnectionInfo);
var
  Conn: TCustomConnection;
  SQLConn: TSQLConnection;
  TempTrans: TSQLTransaction;
  OwnsTransaction: Boolean;
  Tables: TStringList;
  i: Integer;
  TableNode: TTreeNode;
begin
  Conn := GlobalConnManager.GetActiveConnection(Info.aName);
  if not Assigned(Conn) or not (Conn is TSQLConnection) then Exit;
  SQLConn := TSQLConnection(Conn);

  { SQLDB proíbe reatribuir Transaction enquanto a atual estiver ativa (evita
    órfãos de transações pendentes) — então nunca trocamos a Transaction de
    uma conexão que já tenha uma (ex.: a aba de editor já configurou uma
    permanente em CreateTabEdit); só criamos uma própria quando a conexão
    ainda não tem nenhuma (conectado mas nenhuma aba aberta ainda). Rodar a
    consulta de metadados dentro de uma transação já ativa (com trabalho
    pendente do usuário) é seguro — é só um SELECT de leitura, não afeta o
    que está pendente. }
  OwnsTransaction := not Assigned(SQLConn.Transaction);
  if OwnsTransaction then
  begin
    TempTrans := TSQLTransaction.Create(nil);
    TempTrans.DataBase := SQLConn;
    SQLConn.Transaction := TempTrans;
  end;

  try
    Tables := TStringList.Create;
    try
      try
        SQLConn.GetTableNames(Tables, False);
        ConnNode.DeleteChildren;
        for i := 0 to Tables.Count - 1 do
        begin
          TableNode := tvwConnection.Items.AddChild(ConnNode, Tables[i]);
          { nó placeholder só para exibir a seta de expandir; as colunas de
            fato são carregadas sob demanda em tvwConnectionExpanding, para
            não disparar N consultas de metadados logo ao conectar. }
          tvwConnection.Items.AddChild(TableNode, '...');
        end;
        ConnNode.Expand(False);
      except
        on E: Exception do
          LogError('Falha ao carregar tabelas da conexão ' + Info.aName + ': ' + E.Message);
      end;
    finally
      Tables.Free;
    end;
  finally
    if OwnsTransaction then
    begin
      if SQLConn.Transaction.Active then
        SQLConn.Transaction.Commit;
      SQLConn.Transaction := nil;
      TempTrans.Free;
    end;
  end;
end;

procedure TfrmMain.tvwConnectionExpanding(Sender: TObject; Node: TTreeNode; var AllowExpansion: Boolean);
var
  ConnNode: TTreeNode;
  Info: TConnectionInfo;
  Conn: TCustomConnection;
  SQLConn: TSQLConnection;
  TempTrans, UseTrans: TSQLTransaction;
  TempQuery: TSQLQuery;
  OwnsTransaction: Boolean;
  i: Integer;
begin
  AllowExpansion := True;

  { Só nós de tabela (nível 1, com um único filho placeholder "...") precisam
    de carregamento sob demanda das colunas. }
  if (Node.Level <> 1) or (Node.Count <> 1) or (Node.Items[0].Text <> '...') then
    Exit;

  ConnNode := Node.Parent;
  if not Assigned(ConnNode) or not Assigned(ConnNode.Data) then Exit;
  Info := TConnectionInfo(ConnNode.Data);

  Conn := GlobalConnManager.GetActiveConnection(Info.aName);
  if not Assigned(Conn) or not (Conn is TSQLConnection) then Exit;
  SQLConn := TSQLConnection(Conn);

  { A maioria dos engines só permite UMA transação ativa por conexão física
    por vez — criar uma TSQLTransaction própria enquanto a aba já tem uma
    ativa nessa mesma conexão falha com "cannot start a transaction within a
    transaction" (não é só sobre a propriedade Connection.Transaction, é uma
    limitação da conexão real subjacente). Por isso reaproveitamos a
    Transaction já existente da conexão quando há uma; só criamos uma própria
    se a conexão ainda não tiver nenhuma (conectado, mas nenhuma aba aberta). }
  OwnsTransaction := not Assigned(SQLConn.Transaction);
  if OwnsTransaction then
  begin
    TempTrans := TSQLTransaction.Create(nil);
    TempTrans.DataBase := SQLConn;
    UseTrans := TempTrans;
  end
  else
  begin
    TempTrans := nil;
    UseTrans := SQLConn.Transaction;
  end;

  TempQuery := TSQLQuery.Create(nil);
  try
    TempQuery.DataBase := SQLConn;
    TempQuery.Transaction := UseTrans;

    { "WHERE 1=0" é uma forma portável entre os 7 engines suportados de pedir
      só os metadados das colunas (nome + tipo) sem trazer nenhuma linha —
      GetFieldNames só devolve nomes, sem tipo. }
    try
      TempQuery.SQL.Text := 'SELECT * FROM ' + Node.Text + ' WHERE 1=0';
      TempQuery.Open;

      Node.DeleteChildren;
      for i := 0 to TempQuery.FieldCount - 1 do
        tvwConnection.Items.AddChild(Node,
          TempQuery.Fields[i].FieldName + ' : ' + FriendlyFieldTypeName(TempQuery.Fields[i].DataType));

      TempQuery.Close;
    except
      on E: Exception do
        LogError('Falha ao carregar colunas da tabela ' + Node.Text + ': ' + E.Message);
    end;
  finally
    TempQuery.Free;
    if OwnsTransaction then
    begin
      if TempTrans.Active then
        TempTrans.Commit;
      TempTrans.Free;
    end;
  end;
end;

procedure TfrmMain.ExPairSplitterEditorRiseze(Sender: Tobject);
begin
  TPairSplitter(Sender).Position := TPairSplitter(Sender).Parent.ClientWidth div 2;
end;

function TfrmMain.FriendlyFieldTypeName(aType: TFieldType): String;
begin
  { GetEnumName cobre os ~30 valores de TFieldType de uma vez, sem precisar
    manter um "case" manual em paralelo à enumeração da RTL. }
  Result := GetEnumName(TypeInfo(TFieldType), Ord(aType));
  if (Length(Result) > 2) and (LowerCase(Copy(Result, 1, 2)) = 'ft') then
    Result := Copy(Result, 3, Length(Result));
end;

procedure TfrmMain.SetColumnTitlesWithTypes(aTab: TTabSheet; aQuery: TSQLQuery);
var
  Grid: TDBGrid;
  i: Integer;
begin
  if not Assigned(aTab) then Exit;

  Grid := aTab.FindComponent('TabGrid') as TDBGrid;
  if not Assigned(Grid) then Exit;

  for i := 0 to Grid.Columns.Count - 1 do
    if Assigned(Grid.Columns[i].Field) then
      Grid.Columns[i].Title.Caption := Grid.Columns[i].Field.FieldName + ' : ' +
        FriendlyFieldTypeName(Grid.Columns[i].Field.DataType);
end;

procedure TfrmMain.ExecutionThreadDone;
var
  Hours, Minutes, Seconds, Nanoseconds, ElapsedNs: Int64;
  FormatedTime: String;
  ErrorLine: Integer;
  Reg: TRegExpr;
  TDS: TDataSource;
begin
  { Chamado via Synchronize a partir de TSQLExecThread.Execute — roda na
    thread principal, então já pode tocar em controles visuais livremente. }
  TimerExec.Enabled := False;
  StatusBar1.Panels[0].Text := '';

  { A grid foi desvinculada da query antes de iniciar a thread (TDS.DataSet :=
    nil), já que reabrir a query dentro da thread de background dispararia
    notificações do DataSource para a grid fora da thread principal — proibido
    em LCL. Revincular aqui faz a grid repaginar com segurança. }
  if Assigned(FExecTab) then
  begin
    TDS := FExecTab.FindComponent('TabDataSource') as TDataSource;
    if Assigned(TDS) and Assigned(FExecQuery) then
      TDS.DataSet := FExecQuery;
  end;

  ElapsedNs := FExecThread.FElapsedNs;
  Hours := ElapsedNs div 3600000000000;
  ElapsedNs := ElapsedNs mod 3600000000000;
  Minutes := ElapsedNs div 60000000000;
  ElapsedNs := ElapsedNs mod 60000000000;
  Seconds := ElapsedNs div 1000000000;
  ElapsedNs := ElapsedNs mod 1000000000;
  Nanoseconds := ElapsedNs;
  FormatedTime := Format('%.2d:%.2d:%.2d:%.9d', [Hours, Minutes, Seconds, Nanoseconds]);

  Panel3.Visible := True;

  if FExecThread.FHasError then
  begin
    ToolButton4.Enabled := False;
    ToolButton5.Enabled := False;

    if FExecCancelRequested then
      Memo1.Lines.Add(GetDateTime + ': Execução cancelada pelo usuário (conexão fechada à força).')
    else
      Memo1.Lines.Add(FExecThread.FErrorMessage);

    Memo1.SelStart := Length(Memo1.Text);

    if not FExecCancelRequested and Assigned(FExecEditor) then
    begin
      Memo1.SetFocus;

      // Expressão para capturar linha do erro em mensagens do tipo "error at line 3"
      ErrorLine := -1;
      Reg := TRegExpr.Create;
      try
        Reg.Expression := 'line\s+(\d+)';
        if Reg.Exec(FExecThread.FErrorMessage) then
          ErrorLine := StrToIntDef(Reg.Match[1], -1);
      finally
        Reg.Free;
      end;

      if (ErrorLine > 0) and (ErrorLine <= FExecEditor.Lines.Count) then
      begin
        FExecEditor.CaretY := ErrorLine;
        FExecEditor.SetFocus;
        FExecEditor.CaretX := 1;
        FExecEditor.SelectLine;
      end;
    end;
  end
  else
  begin
    ToolButton4.Enabled := True;
    ToolButton5.Enabled := True;

    Memo1.Lines.Add(GetDateTime + ': Alterações anteriores descartadas.');
    Memo1.Lines.Add(GetDateTime + ': Linhas afetadas: ' + IntToStr(FExecThread.FRowsAffected));
    Memo1.Lines.Add(GetDateTime + ': Informações da execução:');
    Memo1.Lines.Add('    Tempo de execução: ' + FormatedTime);

    if Assigned(FExecQuery) and FExecQuery.Active then
      SetColumnTitlesWithTypes(FExecTab, FExecQuery);

    AddHistoryEntry(FExecConnName, FExecThread.FHistoryText);
  end;

  UpdateActionsForActiveTab;
  actExecuteSQL.Enabled := True;
  actScriptExecute.Enabled := True;
  actCancelExecution.Enabled := False;

  FExecThread := nil;
  FExecTab := nil;
  FExecQuery := nil;
  FExecEditor := nil;
  FExecConn := nil;
  FExecCancelRequested := False;
end;

procedure TfrmMain.PrepareExecSQL(Which: String);
var
  TTab: TTabSheet;
  TQuery: TSQLQuery;
  TScript: TSQLScript;
  TTrans: TSQLTransaction;
  TEditor: TSynEdit;
  TDS: TDataSource;
  DBConnected: Boolean;
  aSQL, aScriptText: String;
begin
  if Assigned(FExecThread) then
  begin
    MessageDlg('ExSQL',
      'Já existe uma consulta em execução. Aguarde terminar ou clique em Cancelar antes de iniciar outra.',
      mtWarning, [mbOk], 0, mbOk);
    Exit;
  end;

  TTab := PageControl1.ActivePage;

  if not Assigned(TTab) then Exit;

  TQuery := TTab.FindComponent('TabQuery') as TSQLQuery;
  TScript := TTab.FindComponent('TabScript') as TSQLScript;
  TTrans := TTab.FindComponent('TabTransac') as TSQLTransaction;
  TEditor := TTab.FindComponent('TabSQLEditor') as TSynEdit;
  TDS := TTab.FindComponent('TabDataSource') as TDataSource;

  if not Assigned(TEditor) then Exit;

  if Which = 'script' then
    DBConnected := TScript.DataBase.Connected
  else
    DBConnected := TQuery.DataBase.Connected;

  if not DBConnected then
  begin
    MessageDlg('ExSQL',
      'A conexão com o banco de dados está fechada.',
      mtWarning, [mbOk], 0, mbOk);
    Exit;
  end;

  if Trim(TEditor.Text) = '' then
  begin
    MessageDlg('Query',
      'Nenhuma query informada.',
      mtInformation, [mbOk], 0, mbOk);
    Exit;
  end;

  if Which = 'query' then
  begin
    if Trim(TEditor.SelText) <> '' then
      aSQL := TEditor.SelText
    else
      aSQL := TEditor.Text;

    aSQL := Trim(aSQL);
    if aSQL.EndsWith(';') then
      Delete(aSQL, Length(aSQL), 1);
  end
  else
    aScriptText := TEditor.Text;

  Panel3.Visible := False;
  Memo1.Clear;

  FExecTab := TTab;
  FExecQuery := TQuery;
  FExecEditor := TEditor;
  FExecConnName := TTab.Hint;
  FExecConn := TQuery.DataBase;
  FExecCancelRequested := False;

  { A grid é desvinculada da query antes de disparar a thread: reabrir a query
    (Query.Open) dentro dela notificaria o DataSource/grid fora da thread
    principal, o que é proibido em LCL. TDS é revinculado em
    ExecutionThreadDone, já de volta na thread principal. }
  if Assigned(TDS) then
    TDS.DataSet := nil;

  QueryPerformanceCounter(FExecStartCount);
  StatusBar1.Panels[0].Text := 'Executando... 00:00:00';
  TimerExec.Enabled := True;

  actExecuteSQL.Enabled := False;
  actScriptExecute.Enabled := False;
  actCancelExecution.Enabled := True;

  FExecThread := TSQLExecThread.Create(Self, TQuery, TScript, TTrans, Which, aSQL, aScriptText);
  FExecThread.Start;
end;

procedure TfrmMain.SelectSQLBlock(aSynEdit: TSynEdit);
var
  CurLine: Integer;
  StartLine, EndLine: Integer;
begin
  CurLine := aSynEdit.CaretY;

  // Subir até linha em branco ou topo
  StartLine := CurLine;
  while (StartLine > 1) and (Trim(aSynEdit.Lines[StartLine - 1]) <> '') do
    Dec(StartLine);

  // Descer até linha em branco ou fim
  EndLine := CurLine;
  while (EndLine < aSynEdit.Lines.Count) and (Trim(aSynEdit.Lines[EndLine]) <> '') do
    Inc(EndLine);

  aSynEdit.BlockBegin := Point(1, StartLine);
  aSynEdit.BlockEnd   := Point(Length(aSynEdit.Lines[EndLine - 1]) + 1, EndLine);
end;

function TfrmMain.IsSelectSQL(const SQLText: string): Boolean;
var
  Trimmed: string;
begin
  Trimmed := Trim(SQLText);
  Result := AnsiStartsText('SELECT', UpperCase(Trimmed)) or
            AnsiStartsText('WITH', UpperCase(Trimmed));
end;

function TfrmMain.GetSQLBlockAtCursor(SynEdit: TSynEdit): string;
var
  StartLine, EndLine, i: Integer;
begin
  StartLine := SynEdit.CaretY;
  EndLine := SynEdit.CaretY;

  while (StartLine > 1) and (Pos(';', SynEdit.Lines[StartLine - 1]) = 0) do
    Dec(StartLine);

  while (EndLine < SynEdit.Lines.Count) and (Pos(';', SynEdit.Lines[EndLine - 1]) = 0) do
    Inc(EndLine);

  Result := '';
  for i := StartLine - 1 to EndLine - 1 do
    Result := Result + SynEdit.Lines[i] + LineEnding;
end;

function TfrmMain.GetCurrentWordAtCursor(aSynEdit: TSynEdit): string;
var
  Line: string;
  Col, StartPos: Integer;
begin
  Result := '';

  if aSynEdit.CaretY < 1 then Exit;
  Line := aSynEdit.Lines[aSynEdit.CaretY - 1];
  Col := aSynEdit.CaretX;

  if (Col < 1) or (Length(Line) = 0) then Exit;

  // Encontra o início da palavra
  StartPos := Col;
  while (StartPos > 1) and (Line[StartPos-1] in ['A'..'Z', 'a'..'z', '0'..'9', '_']) do
    Dec(StartPos);

  // Extrai a palavra completa
  while (StartPos <= Length(Line)) and (Line[StartPos] in ['A'..'Z', 'a'..'z', '0'..'9', '_']) do
  begin
    Result := Result + Line[StartPos];
    Inc(StartPos);
  end;
end;

function TfrmMain.OpenExec(Query: TSQLQuery; aSQL: String): String;
var
  Info: TSQLStatementInfo;
begin
  Info := Query.SQLConnection.GetStatementInfo(aSQL);

  case Info.StatementType of
    stSelect, stSelectForUpd:
    begin
      Query.Open;  // SELECT precisa abrir o dataset
      Result := 'Linhas afetadas: '; // INSERT, UPDATE, DELETE, etc
      DoOpen := True;
    end
    else
    begin
      Query.ExecSQL; // INSERT, UPDATE, DELETE, etc
      Result := 'Linhas afetadas: ';
      DoOpen := False;
    end;
  end;
end;

function TfrmMain.RemoveOrderBy(const aSQL: String): String;
var
  PosOrder: Integer;
begin
  { LastDelimiter trata o 1º argumento como um conjunto de caracteres, não uma
    substring — usar Pos evita cortar a SQL no último O/R/D/E/B/Y/espaço que
    aparecer nela (ex.: truncava "...FROM people" para "...FROM peopl", pois
    o 'E' final de "people" batia com um dos caracteres de "ORDER BY"). }
  PosOrder := Pos('ORDER BY', UpperCase(aSQL));
  if PosOrder > 0 then
    Result := Trim(Copy(aSQL, 1, PosOrder - 1))
  else
    Result := aSQL;
end;

function TfrmMain.GetLastStatementFromScript(const Script: String): String;
var
  Parts: TStringList;
  i: Integer;
begin
  { Split ingênuo por ';' (mesmo nível de simplicidade já usado em RemoveOrderBy/
    SelectSQLBlock nesta base de código) — não trata ';' dentro de strings
    literais, aceitável para o uso típico de scripts deste app. }
  Result := '';
  Parts := TStringList.Create;
  try
    Parts.StrictDelimiter := True;
    Parts.Delimiter := ';';
    Parts.DelimitedText := Script;
    for i := Parts.Count - 1 downto 0 do
    begin
      if Trim(Parts[i]) <> '' then
      begin
        Result := Trim(Parts[i]);
        Break;
      end;
    end;
  finally
    Parts.Free;
  end;
end;

function TfrmMain.GetCurrentOrderBy(const aSQL: String; out AField, ADirection: String): Boolean;
var
  PosOrder, SpacePos: Integer;
  Suffix: String;
begin
  Result := False;
  AField := '';
  ADirection := 'ASC';

  PosOrder := Pos('ORDER BY', UpperCase(aSQL));
  if PosOrder = 0 then Exit;

  Suffix := Trim(Copy(aSQL, PosOrder + Length('ORDER BY'), MaxInt));
  if Suffix = '' then Exit;

  SpacePos := Pos(' ', Suffix);
  if SpacePos > 0 then
  begin
    AField := Copy(Suffix, 1, SpacePos - 1);
    if UpperCase(Trim(Copy(Suffix, SpacePos + 1, MaxInt))) = 'DESC' then
      ADirection := 'DESC';
  end
  else
    AField := Suffix;

  Result := True;
end;

function TfrmMain.ConType(aCodType: Integer): Integer;
var
  i: Integer;
begin
  for i := 0 to frmCreateConn.cbbType.Items.Count - 1 do
  begin
    if PtrInt(frmCreateConn.cbbType.Items.Objects[i]) = aCodType then
    begin
      Result := i;
      Break;
    end;
  end;
end;

function TfrmMain.FindConnectionInfoByName(const aName: String): TConnectionInfo;
var
  i: Integer;
begin
  Result := nil;
  for i := 0 to tvwConnection.Items.Count - 1 do
  begin
    if Assigned(tvwConnection.Items[i].Data) and
       (TConnectionInfo(tvwConnection.Items[i].Data).aName = aName) then
    begin
      Result := TConnectionInfo(tvwConnection.Items[i].Data);
      Break;
    end;
  end;
end;

procedure TfrmMain.TitleOrderUpdate(aGrid: TDBGrid; aField, aDirection: String);
var
  i: integer;
  ColumnTitle: string;
begin
  for i:= 0 to aGrid.Columns.Count - 1 do
  begin
    ColumnTitle := StringReplace(aGrid.Columns[i].Title.Caption, '↑', '', [rfReplaceAll, rfIgnoreCase]);
    ColumnTitle := StringReplace(ColumnTitle, '↓', '', [rfReplaceAll, rfIgnoreCase]);
    if aGrid.Columns[i].FieldName = aField then
      aGrid.Columns[i].Title.Caption := ColumnTitle + '' + StrUtils.IfThen(aDirection = 'ASC', '↑', '↓')
    else
      aGrid.Columns[i].Title.Caption := ColumnTitle;
  end;
end;

procedure TfrmMain.CreateTabEdit(ConnName: String; ConnType: Integer; Conn: TCustomConnection);
var
  Tab: TTabSheet;
  TabPairSplitterEditor: TPairSplitter;
  TabSQLEditor: TSynEdit;
  TabDataGrid: TDBGrid;
  TabSQLQuery: TSQLQuery;
  TabSQLScript: TSQLScript;
  TabTrans: TSQLTransaction;
  TabDS: TDataSource;
  TabDB: TDatabase;
begin
  Tab := TTabSheet.Create(PageControl1);
  Tab.PageControl := PageControl1;
  Tab.Caption := 'Editor <'+ConnName+'>';
  Tab.Hint := ConnName; // usado por actNewExecute/actCloseExecute para identificar a conexão da aba

  TabPairSplitterEditor := TPairSplitter.Create(Tab);
  TabPairSplitterEditor.Parent := Tab;
  TabPairSplitterEditor.Align := alClient;
  TabPairSplitterEditor.SplitterType := pstHorizontal;
  TabPairSplitterEditor.Cursor := crHSplit;
  TabPairSplitterEditor.Position := Tab.ClientWidth div 2;
  TabPairSplitterEditor.OnResize := @ExPairSplitterEditorRiseze;

  { Owner é Tab (não TabPairSplitterEditor.Sides[0]) para que Tab.FindComponent('TabSQLEditor')
    consiga localizar o editor mais tarde (FindComponent só busca entre os componentes cujo
    Owner direto é o próprio TTabSheet); o Parent visual continua sendo o painel do splitter. }
  TabSQLEditor := TSynEdit.Create(Tab);
  TabSQLEditor.Name := 'TabSQLEditor';
  { TCustomSynEdit.SetName ecoa o Name para Text quando ambos estão vazios no momento da
    atribuição (comportamento de conveniência de design-time do SynEdit) — limpa aqui para
    garantir que a aba sempre comece com o editor vazio. }
  TabSQLEditor.Clear;
  TabSQLEditor.Parent := TabPairSplitterEditor.Sides[0];
  TabSQLEditor.Align := alClient;
  TabSQLEditor.Highlighter := SynSQLSyn1;
  TabSQLEditor.Font := FontDialog1.Font;
  TabSQLEditor.OnClick := @SynEdit1Click;
  TabSQLEditor.OnKeyPress := @SynEdit1KeyPress;
  TabSQLEditor.OnStatusChange := @SynEdit1StatusChange;
  TabSQLEditor.OnCommandProcessed := @SynEdit1CommandProcessed;

  case ConnType of
    Ord(dbMSSQL):
      TabDB := TMSSQLConnection(Conn);
    Ord(dbFirebird):
      TabDB := TIBConnection(Conn);
    Ord(dbPostgres):
      TabDB := TPQConnection(Conn);
    Ord(dbMySQL), Ord(dbMariaDB): // MariaDB também usa TMySQL80Connection
      TabDB := TMySQL80Connection(Conn);
    Ord(dbSQLite3):
      TabDB := TSQLite3Connection(Conn);
    Ord(dbOracle):
      TabDB := TOracleConnection(Conn);
    Ord(dbODBC):
      TabDB := TODBCConnection(Conn);
    else
      raise Exception.Create('Tipo de conexão não suportado.');
  end;

  TabTrans := TSQLTransaction.Create(Tab);
  TabTrans.DataBase := TabDB;
  TabTrans.Name := 'TabTransac';

  { GetTableNames/GetFieldNames (usados no autocomplete com schema real) exigem
    que a conexão tenha uma Transaction própria configurada, além da associada
    à query/script da aba. Todos os 7 tipos de conexão suportados descendem de
    TSQLConnection, que publica essa propriedade (TDatabase, o tipo de TabDB,
    não a publica). }
  if Conn is TSQLConnection then
    TSQLConnection(Conn).Transaction := TabTrans;

  TabSQLQuery := TSQLQuery.Create(Tab);
  TabSQLQuery.DataBase := TabDB;
  TabSQLQuery.Transaction := TabTrans;
  TabSQLQuery.Name := 'TabQuery';

  TabSQLScript := TSQLScript.Create(Tab);
  TabSQLScript.DataBase := TabDB;
  TabSQLScript.Transaction := TabTrans;
  TabSQLScript.Name := 'TabScript';

  TabDS := TDataSource.Create(Tab);
  TabDS.Name := 'TabDataSource';
  TabDS.DataSet := TabSQLQuery;

  { Owner é Tab (mesmo motivo do TabSQLEditor: FindComponent só busca entre
    filhos diretos do Owner) para que SetColumnTitlesWithTypes/ExecutionThreadDone
    consigam achar a grid via Tab.FindComponent('TabGrid'); o Parent visual
    continua sendo o painel do splitter. }
  TabDataGrid := TDBGrid.Create(Tab);
  TabDataGrid.Name := 'TabGrid';
  TabDataGrid.Parent := TabPairSplitterEditor.Sides[1];
  TabDataGrid.TitleStyle := tsNative;
  TabDataGrid.Align := alClient;
  TabDataGrid.DataSource := TabDS;
  TabDataGrid.OnTitleClick := @DBGrid1TitleClick;
  { dgEditing vem habilitado por padrão no TDBGrid; como a grid é só para
    visualizar/ordenar resultado de query, edição inline acidental de célula
    (sem aviso, sem indicação visual) é desligada — mudanças reais devem ser
    feitas via SQL explícito (UPDATE/INSERT), não por clique na grid. }
  TabDataGrid.Options := TabDataGrid.Options - [dgEditing];

  PageControl1.ActivePage := Tab;
  UpdateActionsForActiveTab;
  RefreshSchemaAutocomplete;
end;

procedure TfrmMain.RefreshSchemaAutocomplete;
var
  TTab: TTabSheet;
  TabQuery: TSQLQuery;
  TabConn: TSQLConnection;
  Tables, Fields: TStringList;
  i: Integer;
begin
  SynCompletion1.ItemList.Assign(FBaseCompletionWords);

  TTab := PageControl1.ActivePage;
  if not Assigned(TTab) then Exit;

  TabQuery := TTab.FindComponent('TabQuery') as TSQLQuery;
  if not Assigned(TabQuery) or not (TabQuery.DataBase is TSQLConnection) then Exit;

  TabConn := TSQLConnection(TabQuery.DataBase);
  if not TabConn.Connected then Exit;

  { Falha ao consultar metadados (engine sem suporte completo a GetTableNames/
    GetFieldNames, timeout, etc.) não deve impedir o uso do editor — o
    autocomplete apenas volta a ter só as palavras-chave estáticas. }
  try
    Tables := TStringList.Create;
    Fields := TStringList.Create;
    try
      TabConn.GetTableNames(Tables, False);
      SynCompletion1.ItemList.AddStrings(Tables);

      for i := 0 to Tables.Count - 1 do
      begin
        Fields.Clear;
        TabConn.GetFieldNames(Tables[i], Fields);
        SynCompletion1.ItemList.AddStrings(Fields);
      end;
    finally
      Tables.Free;
      Fields.Free;
    end;
  except
    on E: Exception do
      LogError('Falha ao carregar schema para autocomplete: ' + E.Message);
  end;
end;

procedure TfrmMain.UpdateActionsForActiveTab;
var
  TTab: TTabSheet;
  TTrans: TSQLTransaction;
  TransActive: Boolean;
begin
  TTab := PageControl1.ActivePage;
  TransActive := False;

  if Assigned(TTab) then
  begin
    TTrans := TTab.FindComponent('TabTransac') as TSQLTransaction;
    TransActive := Assigned(TTrans) and TTrans.Active;
  end;

  actCommit.Enabled := TransActive;
  actRollback.Enabled := TransActive;
end;

procedure TfrmMain.PageControl1Change(Sender: TObject);
begin
  UpdateActionsForActiveTab;
  RefreshSchemaAutocomplete;
end;

procedure TfrmMain.VerifyConnStatus;
var
  Info: TConnectionInfo;
begin
  Info := TConnectionInfo(tvwConnection.Selected.Data);

  if Assigned(Info) then
  begin
    if GlobalConnManager.IsConnected(Info.aName) then
    begin
      actDisconnect.Enabled := True;
      actConnect.Enabled := False;
    end else
    begin
      actDisconnect.Enabled := False;
      actConnect.Enabled := True;
    end;
  end;
end;

procedure TfrmMain.actNewConnExecute(Sender: TObject);
begin
  frmCreateConn.Editing := False;
  frmCreateConn.Show;
end;

procedure TfrmMain.actRollbackExecute(Sender: TObject);
var
  TTab: TTabSheet;
  TTrans: TSQLTransaction;
begin
  TTab := PageControl1.ActivePage;
  if not Assigned(TTab) then Exit;

  TTrans := TTab.FindComponent('TabTransac') as TSQLTransaction;
  if not (Assigned(TTrans) and TTrans.Active) then Exit;

  if MessageDlg('Rollback',
    'Deseja desfazer as alterações feitas?',
    mtConfirmation, [mbYes, mbNo], 0, mbNo) = mrYes then
  begin
    TTrans.Rollback;
    Memo1.Lines.Add(GetDateTime+': Alterações anteriores descartadas.');
    UpdateActionsForActiveTab;
  end;
end;

procedure TfrmMain.actScriptExecuteExecute(Sender: TObject);
begin
  PrepareExecSQL('script');
end;

procedure TfrmMain.actCommitExecute(Sender: TObject);
var
  TTab: TTabSheet;
  TTrans: TSQLTransaction;
begin
  TTab := PageControl1.ActivePage;
  if not Assigned(TTab) then Exit;

  TTrans := TTab.FindComponent('TabTransac') as TSQLTransaction;
  if not (Assigned(TTrans) and TTrans.Active) then Exit;

  if MessageDlg('Commit',
    'Deseja aplicar as alterações feitas?',
    mtConfirmation, [mbYes, mbNo], 0, mbNo) = mrYes then
  begin
    TTrans.Commit;
    Memo1.Lines.Add(GetDateTime+': As alterações foram aplicadas.');
    UpdateActionsForActiveTab;
  end;
end;

procedure TfrmMain.actConnectExecute(Sender: TObject);
var
  ConnectionInfo: TConnectionInfo;
  Conn: TCustomConnection;
begin
  if Assigned(tvwConnection.Selected) and Assigned(tvwConnection.Selected.Data) then
  begin
    ConnectionInfo := TConnectionInfo(tvwConnection.Selected.Data);
    Conn := GlobalConnManager.GetOrCreateConnection(ConnectionInfo);
    Conn.Connected := True;
    tvwConnection.Selected.ImageIndex := 0;
    actConnect.Enabled := False;
    actDisconnect.Enabled := True;
    actEditSQL.Enabled := True;
    PopulateSchemaNodes(tvwConnection.Selected, ConnectionInfo);
  end;
end;

procedure TfrmMain.actDisconnectExecute(Sender: TObject);
begin
  if Assigned(tvwConnection.Selected) then
  begin
    GlobalConnManager.DisconnectByName(tvwConnection.Selected.Text);
    actDisconnect.Enabled := False;
    actConnect.Enabled := True;
    actEditSQL.Enabled := False;
    tvwConnection.Selected.DeleteChildren;
  end;
end;

procedure TfrmMain.actEditConnExecute(Sender: TObject);
var
  ConnectionInfo: TConnectionInfo;
begin
  if not (Assigned(tvwConnection.Selected) and Assigned(tvwConnection.Selected.Data)) then
    Exit;

  ConnectionInfo := TConnectionInfo(tvwConnection.Selected.Data);

  with frmCreateConn do
  begin
    Editing := True;
    EditingCodConn := ConnectionInfo.aCodConn;
    cbbType.ItemIndex := ConType(ConnectionInfo.aCodType);
    edtLibrary.Text := ConnectionInfo.aLibrary;
    edtCharset.Text := ConnectionInfo.aCharset;
    edtName.Text := ConnectionInfo.aName;
    edtHost.Text := ConnectionInfo.aHost;
    if ConnectionInfo.aPort > 0 then
      edtPort.Text := IntToStr(ConnectionInfo.aPort)
    else
      edtPort.Text := '';
    edtDatabase.Text := ConnectionInfo.aDatabase;
    edtUser.Text := ConnectionInfo.aUser;
    edtPassword.Text := '';
    Show;
  end;
end;

procedure TfrmMain.actDeleteConnExecute(Sender: TObject);
var
  ConnectionInfo: TConnectionInfo;
  Query: TSQLQuery;
begin
  if not (Assigned(tvwConnection.Selected) and Assigned(tvwConnection.Selected.Data)) then
    Exit;

  ConnectionInfo := TConnectionInfo(tvwConnection.Selected.Data);

  if MessageDlg('ExSQL',
    'Deseja realmente excluir a conexão "'+ConnectionInfo.aName+'"?',
    mtConfirmation, [mbYes, mbNo], 0, mbNo) <> mrYes then
    Exit;

  GlobalConnManager.DisconnectByName(ConnectionInfo.aName);

  Query := TSQLQuery.Create(nil);
  Query.DataBase := MainConn;
  Query.Transaction := MainTrans;

  try
    try
      with Query do
      begin
        Close;
        SQL.Clear;
        SQL.Text := 'DELETE FROM CONNECTION WHERE CODCONN = :CODCONN';
        ParamByName('CODCONN').AsInteger := ConnectionInfo.aCodConn;

        if MainTrans.Active then MainTrans.EndTransaction;
        MainTrans.StartTransaction;

        ExecSQL;
        MainTrans.Commit;

        LoadConnections;
      end;
    except
      on E: Exception do
      begin
        MainTrans.Rollback;
        MessageDlg('ExSQL',
          'Ocorreu um erro ao excluir a conexão "'+E.Message+'".',
          mtError, [mbOk], 0, mbOk);
      end;
    end;
  finally
    FreeAndNil(Query);
  end;
end;

procedure TfrmMain.actCopyConnExecute(Sender: TObject);
var
  ConnectionInfo: TConnectionInfo;
  Query: TSQLQuery;
  NewName: String;
begin
  if not (Assigned(tvwConnection.Selected) and Assigned(tvwConnection.Selected.Data)) then
    Exit;

  ConnectionInfo := TConnectionInfo(tvwConnection.Selected.Data);

  NewName := ConnectionInfo.aName + ' - cópia';
  if not InputQuery('ExSQL', 'Nome da nova conexão:', NewName) then
    Exit;

  if Trim(NewName) = '' then
  begin
    MessageDlg('ExSQL', 'Informe um nome válido para a conexão.', mtWarning, [mbOk], 0, mbOk);
    Exit;
  end;

  Query := TSQLQuery.Create(nil);
  Query.DataBase := MainConn;
  Query.Transaction := MainTrans;

  try
    try
      with Query do
      begin
        Close;
        SQL.Clear;
        { A senha é copiada já criptografada (aPassword vem direto do banco), sem
          precisar decodificá-la e recriptografá-la. }
        SQL.Text :=
          'INSERT INTO CONNECTION (NAME, DATABASE, HOST, PORT, LIBRARY, CHARSET, USER, PASSWORD, CODTYPE) ' +
          'VALUES (:NAME, :DATABASE, :HOST, :PORT, :LIBRARY, :CHARSET, :USER, :PASSWORD, :CODTYPE)';

        ParamByName('NAME').AsString := NewName;
        ParamByName('DATABASE').AsString := ConnectionInfo.aDatabase;
        if ConnectionInfo.aHost <> '' then ParamByName('HOST').AsString := ConnectionInfo.aHost else ParamByName('HOST').Clear;
        if ConnectionInfo.aPort > 0 then ParamByName('PORT').AsInteger := ConnectionInfo.aPort else ParamByName('PORT').Clear;
        ParamByName('LIBRARY').AsString := ConnectionInfo.aLibrary;
        if ConnectionInfo.aCharset <> '' then ParamByName('CHARSET').AsString := ConnectionInfo.aCharset else ParamByName('CHARSET').Clear;
        if ConnectionInfo.aUser <> '' then ParamByName('USER').AsString := ConnectionInfo.aUser else ParamByName('USER').Clear;
        if ConnectionInfo.aPassword <> '' then ParamByName('PASSWORD').AsString := ConnectionInfo.aPassword else ParamByName('PASSWORD').Clear;
        ParamByName('CODTYPE').AsInteger := ConnectionInfo.aCodType;

        if MainTrans.Active then MainTrans.EndTransaction;
        MainTrans.StartTransaction;

        ExecSQL;
        MainTrans.Commit;

        LoadConnections;
      end;
    except
      on E: Exception do
      begin
        MainTrans.Rollback;
        MessageDlg('ExSQL',
          'Ocorreu um erro ao copiar a conexão "'+E.Message+'".',
          mtError, [mbOk], 0, mbOk);
      end;
    end;
  finally
    FreeAndNil(Query);
  end;
end;

procedure TfrmMain.actEditSQLExecute(Sender: TObject);
var
  ConnectionInfo: TConnectionInfo;
begin
  if Assigned(tvwConnection.Selected) and Assigned(tvwConnection.Selected.Data) then
  begin
    ConnectionInfo := TConnectionInfo(tvwConnection.Selected.Data);
    CreateTabEdit(ConnectionInfo.aName, ConnectionInfo.aCodType, GlobalConnManager.GetActiveConnection(ConnectionInfo.aName));
  end;
end;

procedure TfrmMain.actExecuteSQLExecute(Sender: TObject);
begin
  PrepareExecSQL('query');
end;

procedure TfrmMain.actNewExecute(Sender: TObject);
var
  ActiveTab: TTabSheet;
  Info: TConnectionInfo;
  ConnName: String;
begin
  ActiveTab := PageControl1.ActivePage;
  ConnName := '';

  if Assigned(ActiveTab) then
    ConnName := ActiveTab.Hint;

  if ConnName = '' then
  begin
    if Assigned(tvwConnection.Selected) and Assigned(tvwConnection.Selected.Data) then
      ConnName := TConnectionInfo(tvwConnection.Selected.Data).aName;
  end;

  if ConnName = '' then Exit;

  Info := FindConnectionInfoByName(ConnName);
  if not Assigned(Info) then Exit;

  if not GlobalConnManager.IsConnected(Info.aName) then
  begin
    MessageDlg('ExSQL',
      'A conexão "'+Info.aName+'" não está aberta.',
      mtWarning, [mbOk], 0, mbOk);
    Exit;
  end;

  CreateTabEdit(Info.aName, Info.aCodType, GlobalConnManager.GetActiveConnection(Info.aName));
end;

procedure TfrmMain.actCloseExecute(Sender: TObject);
var
  ActiveTab: TTabSheet;
  TTrans: TSQLTransaction;
begin
  ActiveTab := PageControl1.ActivePage;
  if not Assigned(ActiveTab) then Exit;

  if Assigned(FExecThread) and (FExecTab = ActiveTab) then
  begin
    MessageDlg('ExSQL',
      'Esta aba tem uma consulta em execução. Cancele-a antes de fechar a aba.',
      mtWarning, [mbOk], 0, mbOk);
    Exit;
  end;

  TTrans := ActiveTab.FindComponent('TabTransac') as TSQLTransaction;
  if Assigned(TTrans) and TTrans.Active then
  begin
    if MessageDlg('ExSQL',
      'Há alterações não confirmadas nesta aba. Deseja descartá-las e fechar mesmo assim?',
      mtConfirmation, [mbYes, mbNo], 0, mbNo) <> mrYes then
      Exit;
    TTrans.Rollback;
  end;

  ActiveTab.Free;
  UpdateActionsForActiveTab;
  RefreshSchemaAutocomplete;
end;

procedure TfrmMain.actOpenExecute(Sender: TObject);
var
  ActiveEditor: TSynEdit;
begin
  ActiveEditor := GetActiveSynEdit;
  if not Assigned(ActiveEditor) then Exit;

  if OpenDialog1.Execute then
    ActiveEditor.Lines.LoadFromFile(OpenDialog1.FileName);
end;

procedure TfrmMain.actSaveAsExecute(Sender: TObject);
var
  ActiveEditor: TSynEdit;
begin
  ActiveEditor := GetActiveSynEdit;
  if not Assigned(ActiveEditor) then Exit;

  if SaveDialog1.Execute then
    ActiveEditor.Lines.SaveToFile(SaveDialog1.FileName);
end;

procedure TfrmMain.actSaveExecute(Sender: TObject);
begin
  actSaveAsExecute(Sender);
end;

procedure TfrmMain.actUndoExecute(Sender: TObject);
var
  ActiveEditor: TSynEdit;
begin
  ActiveEditor := GetActiveSynEdit;
  if Assigned(ActiveEditor) then
    ActiveEditor.Undo;
end;

procedure TfrmMain.actRedoExecute(Sender: TObject);
var
  ActiveEditor: TSynEdit;
begin
  ActiveEditor := GetActiveSynEdit;
  if Assigned(ActiveEditor) then
    ActiveEditor.Redo;
end;

procedure TfrmMain.actCopyExecute(Sender: TObject);
var
  ActiveEditor: TSynEdit;
begin
  ActiveEditor := GetActiveSynEdit;
  if Assigned(ActiveEditor) then
    ActiveEditor.CopyToClipboard;
end;

procedure TfrmMain.actPasteExecute(Sender: TObject);
var
  ActiveEditor: TSynEdit;
begin
  ActiveEditor := GetActiveSynEdit;
  if Assigned(ActiveEditor) then
    ActiveEditor.PasteFromClipboard;
end;

procedure TfrmMain.actCutExecute(Sender: TObject);
var
  ActiveEditor: TSynEdit;
begin
  ActiveEditor := GetActiveSynEdit;
  if Assigned(ActiveEditor) then
    ActiveEditor.CutToClipboard;
end;

procedure TfrmMain.actQueryHistoryExecute(Sender: TObject);
var
  TTab: TTabSheet;
  ActiveEditor: TSynEdit;
  ChosenSQL: String;
begin
  TTab := PageControl1.ActivePage;
  if not Assigned(TTab) or (Trim(TTab.Hint) = '') then
  begin
    MessageDlg('ExSQL', 'Abra uma aba de editor conectada antes de consultar o histórico.',
      mtInformation, [mbOk], 0, mbOk);
    Exit;
  end;

  ActiveEditor := GetActiveSynEdit;
  if not Assigned(ActiveEditor) then Exit;

  frmHistory := TfrmHistory.Create(nil);
  try
    LoadHistoryEntries(TTab.Hint, frmHistory.History.Items);
    if frmHistory.History.Items.Count = 0 then
    begin
      MessageDlg('Histórico', 'Nenhuma consulta registrada ainda para "' + TTab.Hint + '".',
        mtInformation, [mbOk], 0, mbOk);
      Exit;
    end;

    frmHistory.History.ItemIndex := 0;
    if frmHistory.ShowModal = mrOK then
    begin
      if frmHistory.History.ItemIndex >= 0 then
      begin
        ChosenSQL := ExtractSQLFromHistoryItem(frmHistory.History.Items[frmHistory.History.ItemIndex]);
        ActiveEditor.SelText := ChosenSQL;
        ActiveEditor.SetFocus;
      end;
    end;
  finally
    FreeAndNil(frmHistory);
  end;
end;

procedure TfrmMain.actExportCSVExecute(Sender: TObject);
const
  CSV_SEPARATOR = ',';
var
  TTab: TTabSheet;
  TQuery: TSQLQuery;
  CSVFile: TextFile;
  i: Integer;
  Line, FieldValue: String;

  function CSVEscape(const aValue: String): String;
  begin
    Result := aValue;
    if (Pos(CSV_SEPARATOR, Result) > 0) or (Pos('"', Result) > 0) or
       (Pos(#13, Result) > 0) or (Pos(#10, Result) > 0) then
      Result := '"' + StringReplace(Result, '"', '""', [rfReplaceAll]) + '"';
  end;

begin
  TTab := PageControl1.ActivePage;
  if not Assigned(TTab) then Exit;

  TQuery := TTab.FindComponent('TabQuery') as TSQLQuery;
  if not Assigned(TQuery) or not TQuery.Active or TQuery.IsEmpty then
  begin
    MessageDlg('Exportar CSV', 'Não há resultado de query aberto nesta aba para exportar.',
      mtInformation, [mbOk], 0, mbOk);
    Exit;
  end;

  SaveDialog1.Filter := 'Arquivo CSV (*.csv)|*.csv';
  SaveDialog1.DefaultExt := 'csv';
  if not SaveDialog1.Execute then Exit;

  AssignFile(CSVFile, SaveDialog1.FileName);
  try
    Rewrite(CSVFile);

    Line := '';
    for i := 0 to TQuery.FieldCount - 1 do
    begin
      if i > 0 then Line := Line + CSV_SEPARATOR;
      Line := Line + CSVEscape(TQuery.Fields[i].FieldName);
    end;
    WriteLn(CSVFile, Line);

    TQuery.DisableControls;
    try
      TQuery.First;
      while not TQuery.EOF do
      begin
        Line := '';
        for i := 0 to TQuery.FieldCount - 1 do
        begin
          if i > 0 then Line := Line + CSV_SEPARATOR;
          if TQuery.Fields[i].IsNull then
            FieldValue := ''
          else
            FieldValue := TQuery.Fields[i].AsString;
          Line := Line + CSVEscape(FieldValue);
        end;
        WriteLn(CSVFile, Line);
        TQuery.Next;
      end;
    finally
      TQuery.EnableControls;
    end;
  finally
    CloseFile(CSVFile);
  end;

  MessageDlg('Exportar CSV', 'Resultado exportado para "' + SaveDialog1.FileName + '".',
    mtInformation, [mbOk], 0, mbOk);
end;

procedure TfrmMain.TimerExecTimer(Sender: TObject);
var
  Frequency, CurrentCount, ElapsedNs, TotalSeconds: Int64;
begin
  QueryPerformanceFrequency(Frequency);
  QueryPerformanceCounter(CurrentCount);
  ElapsedNs := ((CurrentCount - FExecStartCount) * 1000000000) div Frequency;
  TotalSeconds := ElapsedNs div 1000000000;

  StatusBar1.Panels[0].Text := Format('Executando... %.2d:%.2d:%.2d',
    [TotalSeconds div 3600, (TotalSeconds div 60) mod 60, TotalSeconds mod 60]);
end;

procedure TfrmMain.actCancelExecutionExecute(Sender: TObject);
begin
  if not Assigned(FExecThread) then Exit;

  if MessageDlg('Cancelar execução',
    'Tentar cancelar a consulta em execução?' + LineEnding + LineEnding +
    'Não há uma forma segura e genérica de interromper uma consulta em andamento em ' +
    'todos os bancos suportados por este app, então o cancelamento força o fechamento ' +
    'da conexão desta aba. Isso normalmente interrompe a consulta, mas a conexão precisará ' +
    'ser reaberta depois.',
    mtConfirmation, [mbYes, mbNo], 0, mbNo) <> mrYes then
    Exit;

  FExecCancelRequested := True;
  StatusBar1.Panels[0].Text := 'Cancelando...';
  actCancelExecution.Enabled := False;

  { Melhor esforço: forçar o fechamento da conexão desta aba, tentando romper
    a chamada bloqueante que a thread de execução está fazendo. Sem uma API de
    cancelamento genuína e portável entre os 7 engines suportados (cada driver
    exige um mecanismo próprio — libpq tem PQcancel, Firebird tem
    isc_cancel_operation, etc.), esta é a única forma prática de tentar
    interromper a consulta a partir daqui. }
  try
    if Assigned(FExecConn) then
      FExecConn.Connected := False;
  except
    { Ignorado de propósito: fechar a conexão enquanto a thread ainda está no
      meio de uma chamada bloqueante é uma condição de corrida inerente a essa
      técnica — se der erro aqui, a thread ainda vai reportar seu próprio erro
      via ExecutionThreadDone normalmente. }
  end;
end;

procedure TfrmMain.DBGrid1TitleClick(Column: TColumn);
var
  Grid: TDBGrid;
  Qry: TSQLQuery;
  ClickedFieldName: String;
  CurrentField, CurrentDirection, NewDirection, OriginSQL: String;
begin
  { Cada aba tem sua própria grid/query dinâmicas (CreateTabEdit); resolve-se
    a query correta a partir do próprio Column clicado (Column.Grid/Field),
    em vez de depender de qual aba está ativa no momento. }
  if not (Column.Grid is TDBGrid) then Exit;
  Grid := TDBGrid(Column.Grid);

  if not Assigned(Column.Field) or not (Column.Field.DataSet is TSQLQuery) then Exit;
  Qry := TSQLQuery(Column.Field.DataSet);

  { Qry.Active só é True depois de um SELECT bem-sucedido (Open); um DML via
    ExecSQL nunca deixa a query Active, então isso já garante que só se pode
    ordenar o resultado de uma consulta real, sem depender do DoOpen global
    (que reflete apenas a última execução em qualquer aba, não desta). }
  if not Qry.Active or Qry.IsEmpty then Exit;

  { Capturado antes de qualquer Close: os TField auto-criados de Qry (e por
    tabela Column.Field) são destruídos quando a query fecha, então ler
    Column.Field.FieldName depois do Close devolve um objeto/placeholder
    diferente com FieldName vazio — gerando "ORDER BY  ASC" (sem coluna). }
  ClickedFieldName := Column.Field.FieldName;

  OriginSQL := RemoveOrderBy(Qry.SQL.Text);

  { A direção de ordenação é lida do próprio SQL atual (não de um estado
    global), para que alternar ASC/DESC em uma aba não seja afetado pelo
    último clique em outra aba. }
  if GetCurrentOrderBy(Qry.SQL.Text, CurrentField, CurrentDirection) and
     (CompareText(CurrentField, ClickedFieldName) = 0) then
    NewDirection := StrUtils.IfThen(CurrentDirection = 'ASC', 'DESC', 'ASC')
  else
    NewDirection := 'ASC';

  Qry.Close;
  Qry.SQL.Text := OriginSQL;
  Qry.SQL.Add('ORDER BY ' + ClickedFieldName + ' ' + NewDirection);
  Qry.Open;

  TitleOrderUpdate(Grid, ClickedFieldName, NewDirection);
end;

procedure TfrmMain.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  if Assigned(FExecThread) then
  begin
    MessageDlg('ExSQL',
      'Há uma consulta em execução. Cancele-a antes de fechar o ExSQL.',
      mtWarning, [mbOk], 0, mbOk);
    CloseAction := caNone;
    Exit;
  end;

  SaveWindowLayout;
  ClearConnections;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
begin
  GlobalConnManager := TConnectionManager.Create;
  LoadWindowLayout;

  { Guarda a lista estática de palavras-chave SQL do .lfm antes de qualquer
    conexão adicionar nomes de tabela/coluna, para poder restaurá-la como
    base ao trocar de aba/conexão (RefreshSchemaAutocomplete). }
  FBaseCompletionWords := TStringList.Create;
  FBaseCompletionWords.Assign(SynCompletion1.ItemList);

  { Atalhos de teclado — antes só existia Ctrl+Space (autocomplete); nenhuma
    ação de arquivo/execução tinha atalho. Definidos em código (em vez de no
    .lfm) para usar Menus.ShortCut() em vez de calcular o inteiro codificado
    manualmente. }
  actExecuteSQL.ShortCut := Menus.ShortCut(VK_F5, []);
  actScriptExecute.ShortCut := Menus.ShortCut(VK_F5, [ssShift]);
  actNew.ShortCut := Menus.ShortCut(VK_N, [ssCtrl]);
  actOpen.ShortCut := Menus.ShortCut(VK_O, [ssCtrl]);
  actSave.ShortCut := Menus.ShortCut(VK_S, [ssCtrl]);
  actSaveAs.ShortCut := Menus.ShortCut(VK_S, [ssCtrl, ssShift]);
  actClose.ShortCut := Menus.ShortCut(VK_F4, [ssCtrl]);
end;

procedure TfrmMain.SaveWindowLayout;
var
  Ini: TIniFile;
begin
  Ini := TIniFile.Create(GetCurrentDir + 'data\layout.ini');
  try
    Ini.WriteBool('MainWindow', 'Maximized', WindowState = wsMaximized);

    { Se estiver maximizada, RestoredLeft/Top/Width/Height preservam o
      tamanho/posição "normal" para restaurar na próxima abertura, em vez de
      gravar as coordenadas da janela maximizada. }
    if WindowState = wsMaximized then
    begin
      Ini.WriteInteger('MainWindow', 'Left', RestoredLeft);
      Ini.WriteInteger('MainWindow', 'Top', RestoredTop);
      Ini.WriteInteger('MainWindow', 'Width', RestoredWidth);
      Ini.WriteInteger('MainWindow', 'Height', RestoredHeight);
    end
    else
    begin
      Ini.WriteInteger('MainWindow', 'Left', Left);
      Ini.WriteInteger('MainWindow', 'Top', Top);
      Ini.WriteInteger('MainWindow', 'Width', Width);
      Ini.WriteInteger('MainWindow', 'Height', Height);
    end;

    Ini.WriteInteger('MainWindow', 'SplitterPosition', PairSplitter2.Position);
  finally
    Ini.Free;
  end;
end;

procedure TfrmMain.LoadWindowLayout;
var
  Ini: TIniFile;
begin
  if not FileExists(GetCurrentDir + 'data\layout.ini') then Exit;

  Ini := TIniFile.Create(GetCurrentDir + 'data\layout.ini');
  try
    { Position=poScreenCenter (definido no .lfm) recentraliza a janela ao
      exibi-la, ignorando Left/Top atribuídos aqui, a menos que seja trocado
      para poDesigned antes. }
    Position := poDesigned;
    Left := Ini.ReadInteger('MainWindow', 'Left', Left);
    Top := Ini.ReadInteger('MainWindow', 'Top', Top);
    Width := Ini.ReadInteger('MainWindow', 'Width', Width);
    Height := Ini.ReadInteger('MainWindow', 'Height', Height);
    PairSplitter2.Position := Ini.ReadInteger('MainWindow', 'SplitterPosition', PairSplitter2.Position);

    if Ini.ReadBool('MainWindow', 'Maximized', False) then
      WindowState := wsMaximized;
  finally
    Ini.Free;
  end;
end;

procedure TfrmMain.FormDestroy(Sender: TObject);
begin
  GlobalConnManager.Free;
  FBaseCompletionWords.Free;
end;

end.
