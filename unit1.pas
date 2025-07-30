unit Unit1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, SQLDBLib, SQLDB, DB, oracleconnection, Forms, Controls, Graphics, Dialogs, StdCtrls, DBGrids,
  ComCtrls, ExtCtrls, PairSplitter, Buttons, lazutf8, SynEdit, SynHighlighterSQL, StrUtils, RegExpr,
  Windows, SynEditTypes, SynEditKeyCmds, SynCompletion, LCLType, Types;

type

  { TForm1 }

  TForm1 = class(TForm)
    Button1: TBitBtn;
    DataSource1: TDataSource;
    DBGrid1: TDBGrid;
    EditBanco: TEdit;
    EditHost: TEdit;
    EditPorta: TEdit;
    EditSenha: TEdit;
    EditUsuario: TEdit;
    FontDialog1: TFontDialog;
    ImageList1: TImageList;
    Memo1: TMemo;
    OracleConnection1: TOracleConnection;
    PairSplitter1: TPairSplitter;
    PairSplitterSide1: TPairSplitterSide;
    PairSplitterSide2: TPairSplitterSide;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    SQLDBLibraryLoader1: TSQLDBLibraryLoader;
    SQLQuery1: TSQLQuery;
    SQLTransaction1: TSQLTransaction;
    StatusBar1: TStatusBar;
    SynCompletion1: TSynCompletion;
    SynEdit1: TSynEdit;
    SynSQLSyn1: TSynSQLSyn;
    Timer1: TTimer;
    Timer2: TTimer;
    ToolBar1: TToolBar;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    procedure Button1Click(Sender: TObject);
    procedure DBGrid1TitleClick(Column: TColumn);
    procedure FormCreate(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SQLQuery1AfterOpen(DataSet: TDataSet);
    procedure SQLQuery1BeforeOpen(DataSet: TDataSet);
    procedure SynEdit1Click(Sender: TObject);
    procedure SynEdit1CommandProcessed(Sender: TObject; var Command: TSynEditorCommand; var AChar: TUTF8Char;
      Data: pointer);
    procedure SynEdit1KeyPress(Sender: TObject; var Key: char);
    procedure SynEdit1StatusChange(Sender: TObject; Changes: TSynStatusChanges);
    procedure Timer1Timer(Sender: TObject);
    procedure Timer2Timer(Sender: TObject);
    procedure ToolButton1Click(Sender: TObject);
    procedure ToolButton2Click(Sender: TObject);
    procedure ToolButton4Click(Sender: TObject);
    procedure ToolButton5Click(Sender: TObject);
  private
    LastWord, OrderColumn, DirectionColumn: string;
    DoOpen: Boolean;
    procedure ExecuteSQL;
    procedure SelectSQLBlock;
    function GetDateTime: String;
    function IsSelectSQL(const SQLText: string): Boolean;
    function GetSQLBlockAtCursor(SynEdit: TSynEdit): string;
    function GetCurrentWordAtCursor: string;
    function OpenExec(Query: TSQLQuery; Connection: TOracleConnection; aSQL: String): String;
    procedure AtualizarTitulo(aGrid: TDBGrid; aCampo, aDirecao: String);
  public

  end;

var
  Form1: TForm1;
  UsuarioWT, SenhaDB, AliasDB, UsuarioDB, CodRotina: String;

implementation

{$R *.lfm}

{ TForm1 }

procedure TForm1.FormCreate(Sender: TObject);
begin
  UsuarioWT := ParamStrUTF8(1);
  SenhaDB := ParamStrUTF8(2);
  AliasDB := ParamStrUTF8(3);
  UsuarioDB := ParamStrUTF8(4);
  CodRotina := ParamStrUTF8(5);
end;

procedure TForm1.FormResize(Sender: TObject);
begin
  PairSplitterSide1.Height := Panel2.Height div 2;
end;

procedure TForm1.FormShow(Sender: TObject);
begin
  FontDialog1.Font := SynEdit1.Font;
end;

procedure TForm1.SQLQuery1AfterOpen(DataSet: TDataSet);
begin
  Screen.Cursor := crDefault;
end;

procedure TForm1.SQLQuery1BeforeOpen(DataSet: TDataSet);
begin
  Screen.Cursor := crSQLWait;
end;

procedure TForm1.SynEdit1Click(Sender: TObject);
begin
  if Trim(SynEdit1.Text) <> '' then
  begin
    Timer1.Enabled := False;
    Timer1.Enabled := True;
  end;
end;

procedure TForm1.SynEdit1CommandProcessed(Sender: TObject; var Command: TSynEditorCommand; var AChar: TUTF8Char;
  Data: pointer);
begin
  {Timer1.Enabled := False;
  Timer1.Enabled := True;}
end;

procedure TForm1.SynEdit1KeyPress(Sender: TObject; var Key: char);
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

procedure TForm1.SynEdit1StatusChange(Sender: TObject; Changes: TSynStatusChanges);
begin
  {if (scCaretX in Changes) or (scCaretY in Changes) then
  begin
    Timer1.Enabled := False;
    Timer1.Enabled := True;
  end;}
end;

procedure TForm1.Timer1Timer(Sender: TObject);
begin
  Timer1.Enabled := False; // desativa para não repetir
  SelectSQLBlock;
end;

procedure TForm1.Timer2Timer(Sender: TObject);
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

procedure TForm1.ToolButton1Click(Sender: TObject);
begin
  if FontDialog1.Execute then
    SynEdit1.Font := FontDialog1.Font;
end;

procedure TForm1.ToolButton2Click(Sender: TObject);
begin
  if not OracleConnection1.Connected then
  begin
    MessageDlg('Conexão',
      'Não conectado ao banco.',
      mtWarning, [mbOk], 0, mbOk);
    Exit;
  end;


  if (Trim(SynEdit1.Text) <> '') then
    ExecuteSQL
  else
    MessageDlg('Query',
      'Nenhuma query informada.',
      mtInformation, [mbOk], 0, mbOk);
end;

procedure TForm1.ToolButton4Click(Sender: TObject);
begin
  if MessageDlg('Commit',
    'Deseja aplicar alterações feitas?',
    mtConfirmation, [mbYes, mbNo], 0, mbNo) = mrYes then
  begin
    SQLTransaction1.Commit;
    Memo1.Lines.Add(GetDateTime+': As alterações foram aplicadas.');
    ToolButton4.Enabled := False;
    ToolButton5.Enabled := False;
  end;
end;

procedure TForm1.ToolButton5Click(Sender: TObject);
begin
  if MessageDlg('Rollback',
    'Deseja desfazer alterações feitas?',
    mtConfirmation, [mbYes, mbNo], 0, mbNo) = mrYes then
  begin
    SQLTransaction1.Rollback;
    Memo1.Lines.Add(GetDateTime+': Alterações anteriores descartadas.');
    ToolButton4.Enabled := False;
    ToolButton5.Enabled := False;
  end;
end;

procedure TForm1.ExecuteSQL;
var
  ErrorLine: Integer;
  StartCount, EndCount, Frequency: Int64;
  ElapsedNs: Int64;
  Hours, Minutes, Seconds, Nanoseconds: Int64;
  Reg: TRegExpr;
  StartTime, EndTime, DiffTime: TDateTime;
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

  with SQLQuery1 do
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
  end;
end;

procedure TForm1.SelectSQLBlock;
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

function TForm1.GetDateTime: String;
begin
  Result := FormatDateTime('dd-mm-yyyy hh:mm:ss', Now);
end;

function TForm1.IsSelectSQL(const SQLText: string): Boolean;
var
  Trimmed: string;
begin
  Trimmed := Trim(SQLText);
  Result := AnsiStartsText('SELECT', UpperCase(Trimmed)) or
            AnsiStartsText('WITH', UpperCase(Trimmed));
end;

function TForm1.GetSQLBlockAtCursor(SynEdit: TSynEdit): string;
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

function TForm1.GetCurrentWordAtCursor: string;
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

function TForm1.OpenExec(Query: TSQLQuery; Connection: TOracleConnection; aSQL: String): String;
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

procedure TForm1.AtualizarTitulo(aGrid: TDBGrid; aCampo, aDirecao: String);
var
  i: integer;
  Titulo: string;
begin
  for i:= 0 to aGrid.Columns.Count - 1 do
  begin
    Titulo := StringReplace(aGrid.Columns[i].Title.Caption, '↑', '', [rfReplaceAll, rfIgnoreCase]);
    Titulo := StringReplace(Titulo, '↓', '', [rfReplaceAll, rfIgnoreCase]);
    if aGrid.Columns[i].FieldName = aCampo then
      aGrid.Columns[i].Title.Caption := Titulo + '' + StrUtils.IfThen(aDirecao = 'ASC', '↑', '↓')
    else
      aGrid.Columns[i].Title.Caption := Titulo;
  end;
end;

procedure TForm1.Button1Click(Sender: TObject);
var
  i: Integer;
  TextBox: TControl;
begin
  for i := 0 to Panel1.ControlCount -1 do
  begin
    if Panel1.Controls[i] is TEdit then
      TextBox := Panel1.Controls[i];

    if (Trim(TEdit(TextBox).Text) = '') then
    begin
      MessageDlg('Informação',
        'Informe o "'+TEdit(TextBox).TextHint+'" na caixa.',
        mtWarning, [mbOk], 0, mbOk);
      TEdit(TextBox).SetFocus;
      Exit;
    end;
  end;

  with OracleConnection1 do
  begin
    if Button1.Caption = 'Conectar' then
    begin
      if SQLDBLibraryLoader1.Enabled then
      begin
        SQLDBLibraryLoader1.UnloadLibrary;
        SQLDBLibraryLoader1.Enabled := False;
      end;

      Connected := False;
      HostName := EditHost.Text + StrUtils.IfThen((Trim(EditPorta.Text) = ''), '', ':'+EditPorta.Text);
      DatabaseName := EditBanco.Text;
      Password := EditSenha.Text;
      UserName := EditUsuario.Text;

      try
        SQLDBLibraryLoader1.Enabled := True;
        SQLDBLibraryLoader1.LoadLibrary;
        Connected := True;
        Button1.Caption := 'Desconectar';
        StatusBar1.Panels[0].Text := EditUsuario.Text;
        StatusBar1.Panels[1].Text := EditBanco.Text+'@'+EditHost.Text+':'+EditPorta.Text;
      except
        on E: Exception do
        begin
          Button1.Caption := 'Conectar';
          MessageDlg('Erro', 'Erro ao conectar: "' + E.Message + '"', mtError, [mbOk], 0, mbOk);
        end;
      end;
    end else
    begin
      if SQLDBLibraryLoader1.Enabled then
      begin
        SQLDBLibraryLoader1.UnloadLibrary;
        SQLDBLibraryLoader1.Enabled := False;
        Connected := False;
        StatusBar1.Panels[0].Text := '';
        StatusBar1.Panels[1].Text := 'Desconectado';
      end;
    end;
  end;
end;

procedure TForm1.DBGrid1TitleClick(Column: TColumn);
begin
  if (SQLQuery1.Active) and (not SQLQuery1.IsEmpty) and (DoOpen) then
  begin
    if OrderColumn = Column.FieldName then
      DirectionColumn := StrUtils.IfThen(DirectionColumn = 'ASC', 'DESC', 'ASC')
    else
    begin
      OrderColumn := Column.FieldName;
      DirectionColumn := 'ASC';
    end;

    SQLQuery1.Close;

    SQLQuery1.SQL.Add('ORDER BY '+OrderColumn+' '+DirectionColumn);

    SQLQuery1.Open;
    AtualizarTitulo(DBGrid1, OrderColumn, DirectionColumn);
  end;
end;

end.

