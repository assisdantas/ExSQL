unit umain;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, SQLDBLib, SQLDB, IBConnection, SQLite3Conn, DB,
  Forms, Controls, Graphics, Dialogs, StdCtrls, DBGrids,
  ComCtrls, ExtCtrls, PairSplitter, Buttons, lazutf8, SynEdit,
  SynHighlighterSQL, StrUtils, RegExpr, Windows, SynEditTypes, SynEditKeyCmds,
  SynCompletion, LCLType, Menus, ActnList, Types, MSSQLConn, PQConnection,
  OracleConnection, ODBCConn, mysql80conn, Grids;

type

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
    ActionList1: TActionList;
    DBGrid1: TDBGrid;
    FontDialog1: TFontDialog;
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
    PairSplitter1: TPairSplitter;
    PairSplitterSide1: TPairSplitterSide;
    PairSplitterSide2: TPairSplitterSide;
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
    SynEdit1: TSynEdit;
    SynSQLSyn1: TSynSQLSyn;
    TabSheet1: TTabSheet;
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
    procedure actDisconnectExecute(Sender: TObject);
    procedure actEditConnExecute(Sender: TObject);
    procedure actEditSQLExecute(Sender: TObject);
    procedure actExecuteSQLExecute(Sender: TObject);
    procedure actNewConnExecute(Sender: TObject);
    procedure actRollbackExecute(Sender: TObject);
    procedure DBGrid1TitleClick(Column: TColumn);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure MenuItem14Click(Sender: TObject);
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
    // outras procedures...
    procedure ExPairSplitterEditorRiseze(Sender: Tobject);
  private
    LastWord, OrderColumn, DirectionColumn: string;
    DoOpen: Boolean;
    procedure ExecuteSQL;
    procedure SelectSQLBlock;
    procedure TitleOrderUpdate(aGrid: TDBGrid; aField, aDirection: String);
    procedure CreateTabEdit(ConnName: String; ConnType: Integer; Conn: TCustomConnection);
    procedure VerifyConnStatus;
    function IsSelectSQL(const SQLText: string): Boolean;
    function GetSQLBlockAtCursor(SynEdit: TSynEdit): string;
    function GetCurrentWordAtCursor: string;
    function OpenExec(Query: TSQLQuery; Connection: TOracleConnection; aSQL: String): String;
    function RemoveOrderBy(const aSQL: String): String;
    function ConType(aCodType: Integer): Integer;
  public

  end;

var
  frmMain: TfrmMain;
  UsuarioWT, SenhaDB, AliasDB, UsuarioDB, CodRotina: String;

implementation

{$R *.lfm}

uses
  uconnfactory, ucreateconn, utils;

{ TfrmMain }

procedure TfrmMain.FormResize(Sender: TObject);
begin
  PairSplitter1.Sides[0].Height := Panel1.Height div 2;
end;

procedure TfrmMain.FormShow(Sender: TObject);
begin
  FontDialog1.Font := SynEdit1.Font;
end;

procedure TfrmMain.MenuItem14Click(Sender: TObject);
begin
  if FontDialog1.Execute then
    SynEdit1.Font := FontDialog1.Font;
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
  if Trim(SynEdit1.Text) <> '' then
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
  CurrentWord := GetCurrentWordAtCursor;

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
begin
  Timer1.Enabled := False; // desativa para não repetir
  SelectSQLBlock;
end;

procedure TfrmMain.Timer2Timer(Sender: TObject);
var
  pt: TPoint;
  tokenRect: TRect;
  CurrentWord: string;
begin
  Timer2.Enabled := False;

  // Obtém a palavra atual novamente para garantir que ainda é relevante
  CurrentWord := GetCurrentWordAtCursor;

  if Length(CurrentWord) < 2 then Exit; // Mínimo de 2 caracteres

  // Obtém posição do cursor em pixels relativos ao SynEdit
  pt := SynEdit1.RowColumnToPixels(SynEdit1.CaretXY);

  // Converte para coordenadas da tela
  pt := SynEdit1.ClientToScreen(pt);

  // Define o retângulo onde a janela de sugestões será exibida
  tokenRect.Left := pt.X;
  tokenRect.Top := pt.Y;
  tokenRect.Right := pt.X + 1; // largura mínima
  tokenRect.Bottom := pt.Y + SynEdit1.LineHeight;

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

procedure TfrmMain.ExPairSplitterEditorRiseze(Sender: Tobject);
begin
  TPairSplitter(Sender).Position := TPairSplitter(Sender).Parent.ClientWidth div 2;
end;

procedure TfrmMain.ExecuteSQL;
var
  ErrorLine: Integer;
  StartCount, EndCount, Frequency: Int64;
  ElapsedNs: Int64;
  Hours, Minutes, Seconds, Nanoseconds: Int64;
  Reg: TRegExpr;
  FormatedTime, aSQL, MsgReturn: String;
begin
  Panel3.Visible := False;
  Memo1.Clear;

  if Trim(SynEdit1.SelText) <> '' then
    aSQL := SynEdit1.SelText
  else
    aSQL := SynEdit1.Text;

  aSQL := Trim(aSQL);

  if aSQL.EndsWith(';') then
    Delete(aSQL, Length(aSQL), 1);

  {with SQLQuery1 do
  begin
    Close;
    Clear;

    try
      QueryPerformanceFrequency(Frequency);
      QueryPerformanceCounter(StartCount);

      if SQLTransaction1.Active then
        SQLTransaction1.EndTransaction;

      Memo1.Lines.Add(GetDateTime+': Alterações anteriores descartadas.');

      SQLTransaction1.StartTransaction;

      SQL.Text := aSQL;

      MsgReturn := OpenExec(SQLQuery1, OracleConnection1, aSQL);

      QueryPerformanceCounter(EndCount);

      ElapsedNs := ((EndCount - StartCount) * 1000000000) div Frequency;

      Hours := ElapsedNs div 3600000000000;
      ElapsedNs := ElapsedNs mod 3600000000000;

      Minutes := ElapsedNs div 60000000000;
      ElapsedNs := ElapsedNs mod 60000000000;

      Seconds := ElapsedNs div 1000000000;
      ElapsedNs := ElapsedNs mod 1000000000;

      Nanoseconds := ElapsedNs;

      FormatedTime := Format('%.2d:%.2d:%.2d:%.9d', [Hours, Minutes, Seconds, Nanoseconds]);
      Memo1.Lines.Add(GetDateTime+': '+MsgReturn+' '+IntToStr(RowsAffected));
      Memo1.Lines.Add(GetDateTime+': Informações da execução:');
      Memo1.Lines.Add('    Tempo de execução: '+FormatedTime);

      Panel3.Visible := True;

      ToolButton4.Enabled := True;
      ToolButton5.Enabled := True;
    except
      on E: Exception do
      begin
        Screen.Cursor := crDefault;
        ErrorLine := -1;

        ToolButton4.Enabled := False;
        ToolButton5.Enabled := False;

        // Expressão para capturar linha do erro em mensagens do tipo "error at line 3"
        Reg := TRegExpr.Create;
        try
          Reg.Expression := 'line\s+(\d+)';
          if Reg.Exec(E.Message) then
          begin
            ErrorLine := StrToIntDef(Reg.Match[1], -1);
          end;
        finally
          Reg.Free;
        end;

        // Mostra mensagem de erro
        Panel3.Visible := True;
        Memo1.Lines.Add(E.Message);
        Memo1.SelStart := Length(Memo1.Text);
        Memo1.SetFocus;

        // Se conseguiu identificar linha do erro
        if (ErrorLine > 0) and (ErrorLine <= SynEdit1.Lines.Count) then
        begin
          SynEdit1.CaretY := ErrorLine;
          SynEdit1.SetFocus;
          // Se quiser também selecionar a linha inteira:
          SynEdit1.CaretX := 1;
          SynEdit1.SelectLine;
        end;
      end;
    end;
  end;}
end;

procedure TfrmMain.SelectSQLBlock;
var
  CurLine: Integer;
  StartLine, EndLine: Integer;
begin
  CurLine := SynEdit1.CaretY;

  // Subir até linha em branco ou topo
  StartLine := CurLine;
  while (StartLine > 1) and (Trim(SynEdit1.Lines[StartLine - 1]) <> '') do
    Dec(StartLine);

  // Descer até linha em branco ou fim
  EndLine := CurLine;
  while (EndLine < SynEdit1.Lines.Count) and (Trim(SynEdit1.Lines[EndLine]) <> '') do
    Inc(EndLine);

  SynEdit1.BlockBegin := Point(1, StartLine);
  SynEdit1.BlockEnd   := Point(Length(SynEdit1.Lines[EndLine - 1]) + 1, EndLine);
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

function TfrmMain.GetCurrentWordAtCursor: string;
var
  Line: string;
  Col, StartPos: Integer;
begin
  Result := '';

  if SynEdit1.CaretY < 1 then Exit;
  Line := SynEdit1.Lines[SynEdit1.CaretY - 1];
  Col := SynEdit1.CaretX;

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

function TfrmMain.OpenExec(Query: TSQLQuery; Connection: TOracleConnection; aSQL: String): String;
var
  Info: TSQLStatementInfo;
begin
  Info := Connection.GetStatementInfo(aSQL);

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
  PosOrder := LastDelimiter('ORDER BY', UpperCase(aSQL));
  if PosOrder > 0 then
    Result := Trim(Copy(aSQL, 1, PosOrder - 1))
  else
    Result := aSQL;
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
  PairSplitterEditor: TPairSplitter;
  SQLEditor: TSynEdit;
  DataGrid: TDBGrid;
  SQLQuery: TSQLQuery;
  SQLScript: TSQLScript;
  Trans: TSQLTransaction;
  DS: TDataSource;
  DB: TDatabase;
begin
  Tab := TTabSheet.Create(PageControl1);
  Tab.PageControl := PageControl1;
  Tab.Caption := 'Editor <'+ConnName+'>';

  PairSplitterEditor := TPairSplitter.Create(Tab);
  PairSplitterEditor.Parent := Tab;
  PairSplitterEditor.Align := alClient;
  PairSplitterEditor.SplitterType := pstVertical;
  PairSplitterEditor.Cursor := crVSplit;
  PairSplitterEditor.Position := Tab.ClientWidth div 2;
  PairSplitterEditor.OnResize := @ExPairSplitterEditorRiseze;

  SQLEditor := TSynEdit.Create(PairSplitterEditor.Sides[0]);
  SQLEditor.Parent := PairSplitterEditor.Sides[0];
  SQLEditor.Align := alClient;
  SQLEditor.Highlighter := SynSQLSyn1;
  SQLEditor.Font := FontDialog1.Font;
  // colocar eventos

  case ConnType of
    Ord(dbMSSQL):
      DB := TMSSQLConnection(Conn);
    Ord(dbFirebird):
      DB := TIBConnection(Conn);
    Ord(dbPostgres):
      DB := TPQConnection(Conn);
    Ord(dbMySQL), Ord(dbMariaDB): // MariaDB também usa TMySQL80Connection
      DB := TMySQL80Connection(Conn);
    Ord(dbSQLite3):
      DB := TSQLite3Connection(Conn);
    Ord(dbOracle):
      DB := TOracleConnection(Conn);
    Ord(dbODBC):
      DB := TODBCConnection(Conn);
    else
      raise Exception.Create('Tipo de conexão não suportado.');
  end;

  Trans := TSQLTransaction.Create(Tab);
  Trans.DataBase := DB;

  SQLQuery := TSQLQuery.Create(Tab);
  SQLQuery.DataBase := DB;
  SQLQuery.Transaction := Trans;

  DS := TDataSource.Create(nil);
  DS.DataSet := SQLQuery;

  DataGrid := TDBGrid.Create(PairSplitterEditor.Sides[1]);
  DataGrid.Parent := PairSplitterEditor.Sides[1];
  DataGrid.TitleStyle := tsNative;
  DataGrid.Align := alClient;
  DataGrid.DataSource := DS;
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

procedure TfrmMain.actExecuteSQLExecute(Sender: TObject);
begin
  {if not OracleConnection1.Connected then
  begin
    MessageDlg('Conexão',
      'Não conectado ao banco.',
      mtWarning, [mbOk], 0, mbOk);
    Exit;
  end;}


  if (Trim(SynEdit1.Text) <> '') then
    ExecuteSQL
  else
    MessageDlg('Query',
      'Nenhuma query informada.',
      mtInformation, [mbOk], 0, mbOk);
end;

procedure TfrmMain.actNewConnExecute(Sender: TObject);
begin
  frmCreateConn.Editing := False;
  frmCreateConn.Show;
end;

procedure TfrmMain.actRollbackExecute(Sender: TObject);
begin
  {if MessageDlg('Rollback',
    'Deseja desfazer alterações feitas?',
    mtConfirmation, [mbYes, mbNo], 0, mbNo) = mrYes then
  begin
    SQLTransaction1.Rollback;
    Memo1.Lines.Add(GetDateTime+': Alterações anteriores descartadas.');
    actRollback.Enabled := False;
    actCommit.Enabled := False;
  end;}
end;

procedure TfrmMain.actCommitExecute(Sender: TObject);
begin
  {if MessageDlg('Commit',
    'Deseja aplicar alterações feitas?',
    mtConfirmation, [mbYes, mbNo], 0, mbNo) = mrYes then
  begin
    SQLTransaction1.Commit;
    Memo1.Lines.Add(GetDateTime+': As alterações foram aplicadas.');
    actRollback.Enabled := False;
    actCommit.Enabled := False;
  end;}
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
  end;
end;

procedure TfrmMain.actEditConnExecute(Sender: TObject);
var
  ConnectionInfo: TConnectionInfo;
begin
  if Assigned(tvwConnection.Selected) and Assigned(tvwConnection.Selected.Data) then
  begin
    ConnectionInfo := TConnectionInfo(tvwConnection.Selected.Data);
  end;

  with frmCreateConn do
  begin
    Editing := True;
    cbbType.ItemIndex := ConType(ConnectionInfo.aCodType);
    edtLibrary.Text := ConnectionInfo.aLibrary;
    edtCharset.Text := ConnectionInfo.aCharset;
    edtName.Text := ConnectionInfo.aName;
    edtHost.Text := ConnectionInfo.aHost;
    edtPort.Text := IntToStr(ConnectionInfo.aPort);
    edtDatabase.Text := ConnectionInfo.aDatabase;
    edtuser.Text := ConnectionInfo.aDatabase;
    edtPassword.Text := ConnectionInfo.aPassword;
    Show;
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

procedure TfrmMain.DBGrid1TitleClick(Column: TColumn);
var
  OriginSQL: String;
begin
  {if (SQLQuery1.Active) and (not SQLQuery1.IsEmpty) and (DoOpen) then
  begin
    if OrderColumn = Column.FieldName then
      DirectionColumn := StrUtils.IfThen(DirectionColumn = 'ASC', 'DESC', 'ASC')
    else
    begin
      OrderColumn := Column.FieldName;
      DirectionColumn := 'ASC';
    end;

    SQLQuery1.Close;

    OriginSQL := SQLQuery1.SQL.Text;
    OriginSQL := RemoveOrderBy(OriginSQL);

    SQLQuery1.SQL.Text := OriginSQL;

    SQLQuery1.SQL.Add('ORDER BY '+OrderColumn+' '+DirectionColumn);

    SQLQuery1.Open;
    TitleOrderUpdate(DBGrid1, OrderColumn, DirectionColumn);
  end;}
end;

procedure TfrmMain.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  ClearConnections;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
begin
  GlobalConnManager := TConnectionManager.Create;
end;

procedure TfrmMain.FormDestroy(Sender: TObject);
begin
  GlobalConnManager.Free;
end;

end.
