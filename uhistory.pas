unit uhistory;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Dialogs, StdCtrls, Buttons, ExtCtrls;

type

  { TfrmHistory }

  TfrmHistory = class(TForm)
    btnCancel: TBitBtn;
    btnInsert: TBitBtn;
    lbxHistory: TListBox;
    lblHint: TLabel;
    procedure lbxHistoryDblClick(Sender: TObject);
  private
  public
    { Preenchido pelo chamador (umain.pas) antes de ShowModal; contém as
      linhas já formatadas por utils.LoadHistoryEntries. }
    property History: TListBox read lbxHistory;
  end;

var
  frmHistory: TfrmHistory;

implementation

{$R *.lfm}

{ TfrmHistory }

procedure TfrmHistory.lbxHistoryDblClick(Sender: TObject);
begin
  if lbxHistory.ItemIndex >= 0 then
    ModalResult := mrOK;
end;

end.
