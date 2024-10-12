unit fFormulaRequester;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls,
  Buttons, Menus, ActnList, Types,
  uPodcastFormulaRelated;
const
  KEYWORD_PUBDATE = '{PubDate}';
  KEYWORD_EPISODE = '{Episode}';


type

  { TfrmFormulaRequester }

  TfrmFormulaRequester = class(TForm)
    actAddPubDate: TAction;
    actAddEpisode: TAction;
    ActionList1: TActionList;
    btnCancel: TBitBtn;
    btnOk: TBitBtn;
    cbFormula: TComboBox;
    lblFormula: TLabel;
    MainMenu1: TMainMenu;
    MenuItem1: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItem3: TMenuItem;
    procedure AddKeywordExecute(Sender: TObject);
    procedure cbFormulaCloseUp(Sender: TObject);
    procedure cbFormulaDropDown(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { private declarations }
    FPodcasFormulasList: TPodcasFormulasList;
  public
    { public declarations }
    property PodcasFormulasList: TPodcasFormulasList read FPodcasFormulasList write FPodcasFormulasList;
  end;

var
  frmFormulaRequester: TfrmFormulaRequester;

implementation

{$R *.lfm}

{ TfrmFormulaRequester }

procedure TfrmFormulaRequester.AddKeywordExecute(Sender: TObject);
var
  sTextToAdd: string = '';
begin
  case TComponent(Sender).Tag of
    1: sTextToAdd := KEYWORD_PUBDATE;
    2: sTextToAdd := KEYWORD_EPISODE;
  end;

  if sTextToAdd <> '' then
    cbFormula.SelText := sTextToAdd;
end;

{ TfrmFormulaRequester.cbFormulaCloseUp }
procedure TfrmFormulaRequester.cbFormulaCloseUp(Sender: TObject);
var
  iPosSeparator: integer;
  sTextToWrite: string;
begin
  if TComboBox(Sender).ItemIndex <> -1 then
  begin
    sTextToWrite := TComboBox(Sender).Items.Strings[TComboBox(Sender).ItemIndex];
    iPosSeparator := pos('|', sTextToWrite);
    if iPosSeparator <> 0 then sTextToWrite := copy(sTextToWrite, succ(iPosSeparator));
    TComboBox(Sender).Text := sTextToWrite;
  end;
end;

{ TfrmFormulaRequester.cbFormulaDropDown }
procedure TfrmFormulaRequester.cbFormulaDropDown(Sender: TObject);
var
  iMaybeComboIndex: integer;
begin
  iMaybeComboIndex := PodcasFormulasList.GetFormulaIndexForPodcastFormula(TComboBox(Sender).Text);
  if iMaybeComboIndex <> -1 then TComboBox(Sender).ItemIndex := iMaybeComboIndex;
end;

procedure TfrmFormulaRequester.FormCreate(Sender: TObject);
begin
  Caption := Application.MainForm.Caption + ' - Enter your formula';
end;

end.






