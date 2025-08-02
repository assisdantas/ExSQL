unit ucreateconn;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, StrUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls,
  StdCtrls, Buttons, SQLDB;

type

  { TfrmCreateConn }

  TfrmCreateConn = class(TForm)
    Button1: TButton;
    Button2: TButton;
    cbbType: TComboBox;
    grpbDatabase: TGroupBox;
    grpbRegister: TGroupBox;
    grpbCredentials: TGroupBox;
    edtName: TLabeledEdit;
    edtHost: TLabeledEdit;
    edtPort: TLabeledEdit;
    edtDatabase: TLabeledEdit;
    edtCharset: TLabeledEdit;
    edtUser: TLabeledEdit;
    edtPassword: TLabeledEdit;
    edtLibrary: TLabeledEdit;
    OpenDialog1: TOpenDialog;
    btnOpenLibrary: TSpeedButton;
    btnOpenDatabase: TSpeedButton;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure btnOpenLibraryClick(Sender: TObject);
    procedure btnOpenDatabaseClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    procedure OpenFile(EditBox: TLabeledEdit);
    procedure InsertNewConn;
    procedure UpdateConn;
    procedure FormClearFields(Container: TGroupBox);
    function HasEmptyField(Container: TGroupBox): Boolean;
  public
    Editing: Boolean;
  end;

var
  frmCreateConn: TfrmCreateConn;

implementation

{$R *.lfm}

uses
  utils;

{ TfrmCreateConn }

procedure TfrmCreateConn.btnOpenLibraryClick(Sender: TObject);
begin
  OpenFile(edtLibrary);
end;

procedure TfrmCreateConn.Button2Click(Sender: TObject);
begin
  Close;
end;

procedure TfrmCreateConn.Button1Click(Sender: TObject);
var
  i: Integer;
begin
  if cbbType.ItemIndex < 0 then
  begin
    MessageDlg('ExSQL',
      'Informe o tipo de banco de dados.', mtWarning, [mbOk], 0, mbOk);
    Exit;
  end;

  for i := 0 to frmCreateConn.ControlCount - 1 do
  begin
    if frmCreateConn.Controls[i] is TGroupBox then
    begin
      if HasEmptyField(TGroupBox(frmCreateConn.Controls[i])) then
      Exit;
    end;
  end;
  if Editing then
    UpdateConn
  else
    InsertNewConn;
end;

procedure TfrmCreateConn.btnOpenDatabaseClick(Sender: TObject);
begin
  OpenFile(edtDatabase);
end;

procedure TfrmCreateConn.FormShow(Sender: TObject);
begin
  if Editing then
    cbbType.Enabled := False
  else
    cbbType.Enabled := True;
end;

procedure TfrmCreateConn.OpenFile(EditBox: TLabeledEdit);
begin
  if OpenDialog1.Execute then
    EditBox.Text := OpenDialog1.FileName;
end;

procedure TfrmCreateConn.InsertNewConn;
var
  i: Integer;
  Query: TSQLQuery;
begin
  Query := TSQLQuery.Create(nil);
  Query.DataBase := MainConn;
  Query.Transaction := MainTrans;

  try
    try
      with Query do
      begin
        Close;
        SQL.Clear;
        SQL.Add('INSERT INTO CONNECTION');
        SQL.Add('(');
        SQL.Add('NAME, ');
        SQL.Add('DATABASE, ');
        if (Trim(edtHost.Text) <> '') then SQL.Add('HOST, ');
        if (Trim(edtPort.Text) <> '') then SQL.Add('PORT, ');
        SQL.Add('LIBRARY, ');
        if (Trim(edtCharset.Text) <> '') then SQL.Add('CHARSET, ');
        if (Trim(edtUser.Text) <> '') then SQL.Add('USER, ');
        if (Trim(edtPassword.Text) <> '') then SQL.Add('PASSWORD, ');
        SQL.Add('CODTYPE ');
        SQL.Add(')');
        SQL.Add('VALUES');
        SQL.Add('(');
        SQL.Add(':NAME, ');
        SQL.Add(':DATABASE, ');
        if (Trim(edtHost.Text) <> '') then SQL.Add(':HOST, ');
        if (Trim(edtPort.Text) <> '') then SQL.Add(':PORT, ');
        SQL.Add(':LIBRARY, ');
        if (Trim(edtCharset.Text) <> '') then SQL.Add(':CHARSET, ');
        if (Trim(edtUser.Text) <> '') then SQL.Add(':USER, ');
        if (Trim(edtPassword.Text) <> '') then SQL.Add(':PASSWORD, ');
        SQL.Add(':CODTYPE ');
        SQL.Add(')');

        ParamByName('NAME').AsString := edtName.Text;
        ParamByName('DATABASE').AsString := edtDatabase.Text;
        if (Trim(edtHost.Text) <> '') then ParamByName('HOST').AsString := edtHost.Text;
        if (Trim(edtPort.Text) <> '') then ParamByName('PORT').AsInteger := StrToInt(edtPort.Text);
        ParamByName('LIBRARY').AsString := edtLibrary.Text;
        if (Trim(edtCharset.Text) <> '') then ParamByName('CHARSET').AsString := edtCharset.Text;
        if (Trim(edtUser.Text) <> '') then ParamByName('USER').AsString := edtUser.Text;
        if (Trim(edtPassword.Text) <> '') then ParamByName('PASSWORD').AsString := EncryptPassword(edtPassword.Text);
        ParamByName('CODTYPE').AsInteger := PtrInt(cbbType.Items.Objects[cbbType.ItemIndex]);

        if MainTrans.Active then MainTrans.EndTransaction;
        MainTrans.StartTransaction;

        ExecSQL;
        MainTrans.Commit;

        MessageDlg('ExSQL',
          'A conexão foi criada.',
          mtInformation, [mbOk], 0, mbOk);

        for i := 0 to frmCreateConn.ControlCount - 1 do
        begin
          if frmCreateConn.Controls[i] is TGroupBox then
          begin
            FormClearFields(TGroupBox(frmCreateConn.Controls[i]));
          end;
        end;

        Close;
      end;
    except
      on E: Exception do
      begin
        MainTrans.Rollback;
        MessageDlg('ExSQL',
          'Ocorreu um erro ao criar a nova conexão "'+E.Message+'".',
          mtError, [mbOk], 0, mbOk);
      end;
    end;
  finally
    FreeAndNil(Query);
  end;
end;

procedure TfrmCreateConn.UpdateConn;
begin

end;

procedure TfrmCreateConn.FormClearFields(Container: TGroupBox);
var
  i: Integer;
  TextBox: TControl;
begin
  for i := 0 to Container.ControlCount - 1 do
  begin
    if Container.Controls[i] is TLabeledEdit then
    begin
      TextBox := Container.Controls[i];

      if (Trim(TLabeledEdit(TextBox).Text) = '') then
      begin
        TLabeledEdit(TextBox).Clear;
      end;
    end;
  end;
end;

function TfrmCreateConn.HasEmptyField(Container: TGroupBox): Boolean;
const
  IgnoredFields: array[0..4] of string = ('edtCharset', 'edtPort', 'edtHost', 'edtUser', 'edtPassword');
var
  i: Integer;
  TextBox: TControl;
  BoxName: String;
begin
  Result := False;

  for i := 0 to Container.ControlCount - 1 do
  begin
    if Container.Controls[i] is TLabeledEdit then
    begin
      TextBox := Container.Controls[i];
      BoxName := TLabeledEdit(TextBox).Name;

      if StrUtils.MatchText(BoxName, IgnoredFields) then
        Continue;

      if (Trim(TLabeledEdit(TextBox).Text) = '') then
      begin
        MessageDlg('ExSQL',
          'Informe o "' + TLabeledEdit(TextBox).EditLabel.Caption + '" na caixa.',
          mtWarning, [mbOk], 0, mbOk);
        TLabeledEdit(TextBox).SetFocus;
        Result := True;
        Exit;
      end;
    end;
  end;
end;

end.

