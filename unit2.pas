unit Unit2;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls, Buttons, SQLDB;

type

  { TForm2 }

  TForm2 = class(TForm)
    Button1: TButton;
    Button2: TButton;
    ComboBox1: TComboBox;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    GroupBox3: TGroupBox;
    LabeledEdit1: TLabeledEdit;
    LabeledEdit2: TLabeledEdit;
    LabeledEdit3: TLabeledEdit;
    edtDatabase: TLabeledEdit;
    LabeledEdit5: TLabeledEdit;
    LabeledEdit6: TLabeledEdit;
    edtLibrary: TLabeledEdit;
    OpenDialog1: TOpenDialog;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    procedure Button2Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
  private
    procedure OpenFile(EditBox: TLabeledEdit);
  public

  end;

var
  Form2: TForm2;

implementation

{$R *.lfm}

{ TForm2 }

procedure TForm2.SpeedButton1Click(Sender: TObject);
begin
  OpenFile(edtLibrary);
end;

procedure TForm2.Button2Click(Sender: TObject);
begin
  Close;
end;

procedure TForm2.SpeedButton2Click(Sender: TObject);
begin
  OpenFile(edtDatabase);
end;

procedure TForm2.OpenFile(EditBox: TLabeledEdit);
begin
  if OpenDialog1.Execute then
    EditBox.Text := OpenDialog1.FileName;
end;

end.

