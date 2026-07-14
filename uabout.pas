unit uabout;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, ExtCtrls,
  Buttons;

type

  { TfrmAbout }

  TfrmAbout = class(TForm)
    btnOK: TBitBtn;
    Image1: TImage;
    lblTitle: TLabel;
    lblDescription: TLabel;
    lblVersion: TLabel;
    lblAuthor: TLabel;
  private

  public

  end;

var
  frmAbout: TfrmAbout;

implementation

{$R *.lfm}

end.
