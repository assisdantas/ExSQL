unit umain;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, SQLDBLib, SQLDB, IBConnection, SQLite3Conn, DB,
  Forms, Controls, Graphics, Dialogs, StdCtrls, DBGrids,
  ComCtrls, ExtCtrls, PairSplitter, Buttons, lazutf8, SynEdit,
  SynHighlighterSQL, StrUtils, RegExpr, Windows, SynEditTypes, SynEditKeyCmds,
  SynCompletion, LCLType, Menus, ActnList, Types, MSSQLConn, PQConnection,
  OracleConnection, ODBCConn, mysql80conn, Grids, uconnfactory, IniFiles;

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
    procedure RefreshSchemaAutocomplete;
    procedure ExecuteSQL(aSynEdit: TSynEdit; aQuery: TSQLQuery; aScript: TSQLScript; aTrans: TSQLTransaction; aWhich: String);
    procedure PrepareExecSQL(Which: String);
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

var
  frmMain: TfrmMain;
  UsuarioWT, SenhaDB, AliasDB, UsuarioDB, CodRotina: String;

implementation

{$R *.lfm}

uses
  ucreateconn, utils, uabout;

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

procedure TfrmMain.ExPairSplitterEditorRiseze(Sender: Tobject);
begin
  TPairSplitter(Sender).Position := TPairSplitter(Sender).Parent.ClientWidth div 2;
end;

procedure TfrmMain.ExecuteSQL(aSynEdit: TSynEdit; aQuery: TSQLQuery; aScript: TSQLScript; aTrans: TSQLTransaction;
  aWhich: String);
var
  ErrorLine, RowsAffect: Integer;
  StartCount, EndCount, Frequency: Int64;
  ElapsedNs: Int64;
  Hours, Minutes, Seconds, Nanoseconds: Int64;
  Reg: TRegExpr;
  FormatedTime, aSQL, MsgReturn: String;
begin
  Panel3.Visible := False;
  Memo1.Clear;

  try
    QueryPerformanceFrequency(Frequency);
    QueryPerformanceCounter(StartCount);

    if aWhich = 'query' then
    begin
      if Trim(aSynEdit.SelText) <> '' then
        aSQL := aSynEdit.SelText
      else
        aSQL := aSynEdit.Text;

      aSQL := Trim(aSQL);

      if aSQL.EndsWith(';') then
        Delete(aSQL, Length(aSQL), 1);

      if aTrans.Active then
        aTrans.EndTransaction;

      Memo1.Lines.Add(GetDateTime+': Alterações anteriores descartadas.');

      aTrans.StartTransaction;

      aQuery.SQL.Text := aSQL;

      MsgReturn := OpenExec(aQuery, aSQL);

      RowsAffect := aQuery.RowsAffected;
    end
    else if aWhich = 'script' then
    begin
      if aTrans.Active then
        aTrans.EndTransaction;

      Memo1.Lines.Add(GetDateTime+': Alterações anteriores descartadas.');

      aTrans.StartTransaction;

      aScript.Script := aSynEdit.Lines;
      aScript.ExecuteScript;

      MsgReturn := 'Linhas afetadas: ';

      RowsAffect := -1;
    end;

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
    Memo1.Lines.Add(GetDateTime+': '+MsgReturn+' '+IntToStr(RowsAffect));
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
      if (ErrorLine > 0) and (ErrorLine <= aSynEdit.Lines.Count) then
      begin
        aSynEdit.CaretY := ErrorLine;
        aSynEdit.SetFocus;
        // Se quiser também selecionar a linha inteira:
        aSynEdit.CaretX := 1;
        aSynEdit.SelectLine;
      end;
    end;
  end;
end;

procedure TfrmMain.PrepareExecSQL(Which: String);
var
  TTab: TTabSheet;
  TQuery: TSQLQuery;
  TScript: TSQLScript;
  TTrans: TSQLTransaction;
  TEditor: TSynEdit;
  DBConnected: Boolean;
begin
  TTab := PageControl1.ActivePage;

  if not Assigned(TTab) then Exit;

  TQuery := TTab.FindComponent('TabQuery') as TSQLQuery;
  TScript := TTab.FindComponent('TabScript') as TSQLScript;
  TTrans := TTab.FindComponent('TabTransac') as TSQLTransaction;
  TEditor := TTab.FindComponent('TabSQLEditor') as TSynEdit;

  if not Assigned(TEditor) then Exit;

  if Which = 'script' then
  begin
    DBConnected := TScript.DataBase.Connected;
  end
  else if Which = 'query' then
  begin
    DBConnected := TQuery.DataBase.Connected;
  end;

  if not DBConnected then
  begin
    MessageDlg('ExSQL',
      'A conexão com o banco de dados está fechada.',
      mtWarning, [mbOk], 0, mbOk);
    Exit;
  end;

  if (Trim(TEditor.Text) <> '') then
  begin
    ExecuteSQL(TEditor, TQuery, TScript, TTrans, Which);
    UpdateActionsForActiveTab;
  end
  else
    MessageDlg('Query',
      'Nenhuma query informada.',
      mtInformation, [mbOk], 0, mbOk);
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
  TabDS.DataSet := TabSQLQuery;

  TabDataGrid := TDBGrid.Create(TabPairSplitterEditor.Sides[1]);
  TabDataGrid.Parent := TabPairSplitterEditor.Sides[1];
  TabDataGrid.TitleStyle := tsNative;
  TabDataGrid.Align := alClient;
  TabDataGrid.DataSource := TabDS;
  TabDataGrid.OnTitleClick := @DBGrid1TitleClick;

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
