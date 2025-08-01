unit ucreateconn;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls, Buttons, SQLDB;

type

  { TfrmCreateConn }

  TfrmCreateConn = class(TForm)
    Button1: TButton;
    Button2: TButton;
    cbbType: TComboBox;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    GroupBox3: TGroupBox;
    LabeledEdit1: TLabeledEdit;
    LabeledEdit2: TLabeledEdit;
    LabeledEdit3: TLabeledEdit;
    edtDatabase: TLabeledEdit;
    LabeledEdit4: TLabeledEdit;
    LabeledEdit5: TLabeledEdit;
    LabeledEdit6: TLabeledEdit;
    edtLibrary: TLabeledEdit;
    OpenDialog1: TOpenDialog;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    procedure Button2Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
  private
    procedure OpenFile(EditBox: TLabeledEdit);
  public

  end;

var
  frmCreateConn: TfrmCreateConn;

implementation

{$R *.lfm}

uses
  utils;

{ TfrmCreateConn }

procedure TfrmCreateConn.SpeedButton1Click(Sender: TObject);
begin
  OpenFile(edtLibrary);
end;

procedure TfrmCreateConn.Button2Click(Sender: TObject);
begin
  Close;
end;

procedure TfrmCreateConn.FormShow(Sender: TObject);
begin
  LoadDatabaseTypes(cbbType);
end;

procedure TfrmCreateConn.SpeedButton2Click(Sender: TObject);
begin
  OpenFile(edtDatabase);
end;

procedure TfrmCreateConn.OpenFile(EditBox: TLabeledEdit);
begin
  if OpenDialog1.Execute then
    EditBox.Text := OpenDialog1.FileName;
end;

end.

