unit uPodcastFormulaRelated;

{$mode objfpc}{$H+}

interface

uses
  IniFiles, Classes, SysUtils, StdCtrls;
type
  TPodcastFormulaEntry = class(TObject)
  private
    FPodcastName: string;
    FPodcastFormula: string;
  public
    constructor Create(const paramPodcastName, paramPodcastFormula: string);
    function CombineForDropList: string;
    property PodcastName: string read FPodcastName write FPodcastName;
    property PodcastFormula: string read FPodcastFormula write FPodcastFormula;
  end;

  TPodcasFormulasList = class(TList)
  private
    function GetPodcastFormulaEntry(const paramIndex: integer): TPodcastFormulaEntry;
  public
    function GetFormulaIndexForPodcastName(const paramPodcastName: string): integer;
    function GetFormulaIndexForPodcastFormula(const paramPodcastFormula: string): integer;
    function GetBestFormulaForThisPodcastName(const paramPodcastName: string): string;
    procedure AddNewOrUpdateFormulaEntries(const paramPodcastName, paramPodcastFormula: string; bSortAfterAddition: boolean = True);
    procedure SaveToConfiguration(const ConfigIni: TIniFile);
    procedure LoadFromConfiguration(const ConfigIni: TIniFile);
    procedure FeedComboWithPossibleFormulas(const paramComboBox: TComboBox);
    procedure SortByPodcastName;
    property PodcastFormulaEntry[paramIndex: integer]: TPodcastFormulaEntry read GetPodcastFormulaEntry;
  end;


implementation

const
  sSECTION_FORMULAS: string = 'Formulas';
  s_CONFIG_PARAM_PODNAME: string = 'PodName';
  s_CONFIG_PARAM_PODFORMULA: string = 'PodFormula';
  s_CONFIG_LAST_ENTRY = 'LastEntry';
  sMYSEPARATOR = '|';

{ TPodcastFormulaEntry.Create }
constructor TPodcastFormulaEntry.Create(const paramPodcastName, paramPodcastFormula: string);
begin
  inherited Create;
  self.FPodcastName := paramPodcastName;
  self.FPodcastFormula := paramPodcastFormula;
end;

{ TPodcastFormulaEntry.CombineForDropList }
function TPodcastFormulaEntry.CombineForDropList: string;
begin
  Result := self.PodcastName + sMYSEPARATOR + self.PodcastFormula;
end;

{ TPodcasFormulasList.GetPodcastFormulaEntry }
function TPodcasFormulasList.GetPodcastFormulaEntry(const paramIndex: integer): TPodcastFormulaEntry;
begin
  Result := TPodcastFormulaEntry(Items[paramIndex]);
end;

{ TPodcasFormulasList.GetFormulaIndexForPodcastName }
function TPodcasFormulasList.GetFormulaIndexForPodcastName(const paramPodcastName: string): integer;
var
  iSeekIndex: integer = 0;
begin
  Result := -1;
  while (Result = -1) and (iSeekIndex < self.Count) do
    if SameText(PodcastFormulaEntry[iSeekIndex].PodcastName, paramPodcastName) then
      Result := iSeekIndex
    else
      Inc(iSeekIndex);
end;

{ TPodcasFormulasList.GetFormulaIndexForPodcastFormula }
function TPodcasFormulasList.GetFormulaIndexForPodcastFormula(const paramPodcastFormula: string): integer;
var
  iSeekIndex: integer = 0;
begin
  Result := -1;
  while (Result = -1) and (iSeekIndex < self.Count) do
    if SameText(PodcastFormulaEntry[iSeekIndex].PodcastFormula, paramPodcastFormula) then
      Result := iSeekIndex
    else
      Inc(iSeekIndex);
end;

{ TPodcasFormulasList.GetBestFormulaForThisPodcastName }
function TPodcasFormulasList.GetBestFormulaForThisPodcastName(const paramPodcastName: string): string;
var
  iMaybeIndex: integer;
begin
  iMaybeIndex := self.GetFormulaIndexForPodcastName(paramPodcastName);
  if iMaybeIndex <> -1 then
    Result := self.PodcastFormulaEntry[iMaybeIndex].PodcastFormula
  else if self.Count > 0 then
    Result := self.PodcastFormulaEntry[0].PodcastFormula
  else
    Result := '';
end;

{ TPodcasFormulasList.AddNewOrUpdateFormulaEntries }
procedure TPodcasFormulasList.AddNewOrUpdateFormulaEntries(const paramPodcastName, paramPodcastFormula: string; bSortAfterAddition: boolean = True);
var
  iMaybeIndex: integer;
  APodcastFormulaEntry: TPodcastFormulaEntry;
begin
  iMaybeIndex := self.GetFormulaIndexForPodcastName(paramPodcastName);
  if iMaybeIndex <> -1 then
  begin
    PodcastFormulaEntry[iMaybeIndex].PodcastFormula := paramPodcastFormula;
  end
  else
  begin
    APodcastFormulaEntry := TPodcastFormulaEntry.Create(paramPodcastName, paramPodcastFormula);
    Self.Add(APodcastFormulaEntry);
    if bSortAfterAddition then SortByPodcastName;
  end;
end;

{ TPodcasFormulasList.LoadFromConfiguration }
procedure TPodcasFormulasList.LoadFromConfiguration(const ConfigIni: TIniFile);
var
  iFormulaIndex: integer;
  sPodcastName, sPodcastFormula: string;
begin
  iFormulaIndex := 0;
  repeat
    sPodcastName := ConfigIni.ReadString(sSECTION_FORMULAS, Format('%s%3.3d', [s_CONFIG_PARAM_PODNAME, iformulaIndex]), s_CONFIG_LAST_ENTRY);
    sPodcastFormula := ConfigIni.ReadString(sSECTION_FORMULAS, Format('%s%3.3d', [s_CONFIG_PARAM_PODFORMULA, iformulaIndex]), s_CONFIG_LAST_ENTRY);
    if sPodcastName <> s_CONFIG_LAST_ENTRY then  self.AddNewOrUpdateFormulaEntries(sPodcastName, sPodcastFormula, False);
    Inc(iFormulaIndex);
  until (iFormulaIndex = 1000) or (sPodcastName = s_CONFIG_LAST_ENTRY);
end;

{ TPodcasFormulasList.SaveToConfiguration }
procedure TPodcasFormulasList.SaveToConfiguration(const ConfigIni: TIniFile);
var
  iFormulaIndex: integer;
begin
  for iFormulaIndex := 0 to pred(self.Count) do
  begin
    ConfigIni.WriteString(sSECTION_FORMULAS, Format('%s%3.3d', [s_CONFIG_PARAM_PODNAME, iFormulaIndex]), self.PodcastFormulaEntry[iFormulaIndex].PodcastName);
    ConfigIni.WriteString(sSECTION_FORMULAS, Format('%s%3.3d', [s_CONFIG_PARAM_PODFORMULA, iFormulaIndex]), self.PodcastFormulaEntry[iFormulaIndex].PodcastFormula);
  end;
  ConfigIni.WriteString(sSECTION_FORMULAS, Format('%s%3.3d', [s_CONFIG_PARAM_PODNAME, self.Count]), s_CONFIG_LAST_ENTRY);
  ConfigIni.WriteString(sSECTION_FORMULAS, Format('%s%3.3d', [s_CONFIG_PARAM_PODFORMULA, self.Count]), s_CONFIG_LAST_ENTRY);
end;

{ TPodcasFormulasList.FeedComboWithPossibleFormulas }
procedure TPodcasFormulasList.FeedComboWithPossibleFormulas(const paramComboBox: TComboBox);
var
  iFormulaIndex: integer;
begin
  paramComboBox.Items.BeginUpdate;
  try
    paramComboBox.Items.Clear;
    for iFormulaIndex := 0 to pred(Count) do
      paramComboBox.Items.Add(self.PodcastFormulaEntry[iFormulaIndex].CombineForDropList);
  finally
    paramComboBox.Items.EndUpdate;
  end;
end;

{ ComparePodcastEntryByPodcastName }
function ComparePodcastEntryByPodcastName(Item1, Item2: Pointer): integer;
begin
  Result := CompareText(TPodcastFormulaEntry(Item1).PodcastName, TPodcastFormulaEntry(Item2).PodcastName);
end;

{ TPodcasFormulasList.SortByPodcastName }
procedure TPodcasFormulasList.SortByPodcastName;
begin
  self.Sort(@ComparePodcastEntryByPodcastName);
end;

end.



