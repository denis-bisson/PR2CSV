unit fpPR2CSV;

{$mode objfpc}{$H+}

{-$define DEBUG}

interface

uses
  Classes, SysUtils, sqlite3conn, IBConnection, sqldb, DB, FileUtil, Forms,
  Controls, Graphics, Dialogs, Menus, ActnList, ComCtrls, DBGrids, StdCtrls,
  ExtCtrls, EditBtn, Grids, IniFiles,
  uPodcastFormulaRelated;
type
  { TfrmPR2CSV }
  TfrmPR2CSV = class(TForm)
    actBuildCSV: TAction;
    actExit: TAction;
    actAddFormulasForKnownPodcasts: TAction;
    actEditManually: TAction;
    actRenameFiles: TAction;
    actSelectFromSamePodcast: TAction;
    actSelectAll: TAction;
    actValidateRenameCouldBeDone: TAction;
    actSetFormula: TAction;
    actRemoveFormula: TAction;
    alMainActionList: TActionList;
    cbSeparator: TComboBox;
    dePodcastRepublicDatabase: TDirectoryEdit;
    dePodcastRepublicMP3Files: TDirectoryEdit;
    deTargetDirectoryForCSV: TDirectoryEdit;
    ilMainImageList: TImageList;
    Label1: TLabel;
    lbldePodcastRepublicDatabase: TLabel;
    lbldePodcastRepublicMP3Files: TLabel;
    lbldeTargetDirectoryForCSV: TLabel;
    lblLog: TLabel;
    MainMenu1: TMainMenu;
    memoCSV: TMemo;
    MenuItem1: TMenuItem;
    MenuItem10: TMenuItem;
    MenuItem11: TMenuItem;
    MenuItem12: TMenuItem;
    MenuItem13: TMenuItem;
    MenuItem14: TMenuItem;
    MenuItem15: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem4: TMenuItem;
    MenuItem5: TMenuItem;
    MenuItem6: TMenuItem;
    MenuItem7: TMenuItem;
    MenuItem8: TMenuItem;
    MenuItem9: TMenuItem;
    miAction: TMenuItem;
    mySQLite3Connection: TSQLite3Connection;
    pcMainPageControl: TPageControl;
    PopupMenu1: TPopupMenu;
    sbMainStatusBar: TStatusBar;
    sdDialog: TSelectDirectoryDialog;
    SQLQueryGeneric: TSQLQuery;
    SQLTransaction1: TSQLTransaction;
    sfFileGrid: TStringGrid;
    tbExit1: TToolButton;
    tbSetForumla: TToolButton;
    tbConfig: TToolBar;
    tbBuildCSV: TToolButton;
    tbFileGrid: TToolBar;
    ToolButton1: TToolButton;
    tbExit2: TToolButton;
    tbRemoveformula: TToolButton;
    tbValidateRenameCouldBeDone: TToolButton;
    tbSelectAll: TToolButton;
    tbRenameFile: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    tsConfiguration: TTabSheet;
    tsLog: TTabSheet;
    tsFileGrid: TTabSheet;
    procedure actAddFormulasForKnownPodcastsExecute(Sender: TObject);
    procedure actBuildCSVExecute(Sender: TObject);
    procedure actEditManuallyExecute(Sender: TObject);
    procedure actExitExecute(Sender: TObject);
    procedure actRemoveFormulaExecute(Sender: TObject);
    procedure actRenameFilesExecute(Sender: TObject);
    procedure actSelectAllExecute(Sender: TObject);
    procedure actSelectFromSamePodcastExecute(Sender: TObject);
    procedure actSetFormulaExecute(Sender: TObject);
    procedure actValidateRenameCouldBeDoneExecute(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure sfFileGridDblClick(Sender: TObject);
  private
    { private declarations }
    sDatabaseFilename: string;
    myEnglishFormat: TFormatSettings;
    sSeparator, sOutputCSVFilename: string;
    sCSVResult: TStringList;
    sLastFormula: string;
    FPodcasFormulasList: TPodcasFormulasList;
    function AttemptToGetDatabaseFilename: boolean;
    function isValidDatabaseFilename(paramFilename: string): boolean;
    procedure AddCSVLine(paramPodCast, paramPubDate, paramEpisode, paramOriginalFilename: string; paramJustToCSVFile: boolean = False);
    procedure LoadConfiguration;
    procedure SaveConfiguration;
    function GetOutputCSVFilename: string;
    procedure SaveCSVFile;
    procedure ArrangeFileGrid;
    procedure InitFileGrid;
    procedure InitMyEnglishFormat;
    procedure RefreshOutputNameForSelection;
    function GetTitleBarCaption: string;
    function isMP3FilePresent(sFilename: string): boolean;
    function isOkToDoRenaming: boolean;
    function ValidateBasicPathExists: boolean;
    function ValidateAllSourceFileExists: boolean;
    function ValidateAllDestinationFilenameAreDifferents: boolean;
    function ValidateAllDestinationFilenameDontExists: boolean;
    function DoActualRenamingJob: boolean;
    function IncludeRepetitionInFormula(sFormula: string; iIndexRepetition: integer): string;
    procedure GenerateAllFilenamesFromFormulas;
  public
    { public declarations }
    property PodcasFormulasList: TPodcasFormulasList read FPodcasFormulasList write FPodcasFormulasList;
  end;

var
  frmPR2CSV: TfrmPR2CSV;

implementation

{$R *.lfm}

uses
  // Lazarus
  FileInfo, Registry, dateutils,

  // PR2CSV
  uMyFileRename, fFormulaRequester;
const
  //  REGISTRY_LOCATION = 'HKEY_CURRENT_USER\SOFTWARE\PR2CSV'; // By default, it will be saved to "HKEY_CURRENT_USER\SOFTWARE"
  sMAINSECTION = 'MainSection';

  COL_FileNumber = 0;
  COL_Podcast = 1;
  COL_PubDate = 2;
  COL_EpisodeName = 3;
  COL_OriginalFilename = 4;
  COL_FormulaForOutputName = 5;
  COL_OutputName = 6;

  sEPISODE_KEYWORD = 'Episode_R6';
  sDOWNLOAD_KEYWORD = 'Download_R5';
  sPOD_KEYWORD = 'Pod_R8';

{ TfrmPR2CSV }

{ TfrmPR2CSV.AddCSVLine }
// Will add to the memo a CSV line from the provided fields.
// Routine should make sure to use the appropriate separator, quotes if necessary, etc.
procedure TfrmPR2CSV.AddCSVLine(paramPodCast, paramPubDate, paramEpisode, paramOriginalFilename: string; paramJustToCSVFile: boolean = False);
var
  sLineToAdd: string;
  iWorkingRow: integer;
begin
  sLineToAdd := paramPodCast;
  sLineToAdd := sLineToAdd + sSeparator + paramPubDate;
  sLineToAdd := sLineToAdd + sSeparator + paramEpisode;
  sLineToAdd := sLineToAdd + sSeparator + paramOriginalFilename;
  sCSVResult.Add(sLineToAdd);

  if not paramJustToCSVFile then
  begin
    iWorkingRow := sfFileGrid.RowCount;
    sfFileGrid.RowCount := sfFileGrid.RowCount + 1;
    sfFileGrid.Cells[COL_FileNumber, iWorkingRow] := Format('%4.4d', [iWorkingRow]);
    sfFileGrid.Cells[COL_Podcast, iWorkingRow] := paramPodCast;
    sfFileGrid.Cells[COL_PubDate, iWorkingRow] := paramPubDate;
    sfFileGrid.Cells[COL_EpisodeName, iWorkingRow] := paramEpisode;
    sfFileGrid.Cells[COL_OriginalFilename, iWorkingRow] := paramOriginalFilename;
    sfFileGrid.Cells[COL_FormulaForOutputName, iWorkingRow] := '';
    sfFileGrid.Cells[COL_OutputName, iWorkingRow] := '';
  end;
end;

{ TfrmPR2CSV.isValidDatabaseFilename }
// Will attempt to open the file as a SQLite file and see if we can get at least one epidose from it.
// If so, will return TRUE and will return FALSE otherwise.
function TfrmPR2CSV.isValidDatabaseFilename(paramFilename: string): boolean;
begin
  memoCSV.Lines.Add(Format('Trying to read file "%s" and see if we have table "%s"...', [paramFilename, sEPISODE_KEYWORD]));
  mySQLite3Connection.Close; // Ensure the connection is closed when we start
  mySQLite3Connection.DatabaseName := paramFilename;

  mySQLite3Connection.Open;
  SQLTransaction1.Active := True;

  SQLQueryGeneric.PacketRecords := -1;
  SQLQueryGeneric.SQL.Clear;
  SQLQueryGeneric.SQL.Text := 'select * from "' + sEPISODE_KEYWORD + '"';
  SQLQueryGeneric.Open;

  memoCSV.Lines.Add(Format('Number of records found: %d', [SQLQueryGeneric.RecordCount]));
  if SQLQueryGeneric.RecordCount > 0 then
  begin
    memoCSV.Lines.Add('That''s it!');
    Result := True;
  end
  else
  begin
    memoCSV.Lines.Add('That''s not it!');
  end;

  SQLQueryGeneric.Close;

  SQLTransaction1.Active := False;
  mySQLite3Connection.Close(True);
end;

{ TfrmPR2CSV.AttemptToGetDatabaseFilename }
// Will scan the files in the specified folder and will return the one that is probably the actual database.
// If none is found, we will return an empty string}
function TfrmPR2CSV.AttemptToGetDatabaseFilename: boolean;
var
  mySearchRec: TSearchRec;
  iKeepGoing: integer;
begin
  Result := False;
  sDatabaseFilename := '';

  iKeepGoing := FindFirst(IncludeTrailingPathDelimiter(dePodcastRepublicDatabase.Text) + '*.*', faAnyFile, mySearchRec);
  while (iKeepGoing = 0) and (sDatabaseFilename = '') do
  begin
    if fileExists(IncludeTrailingPathDelimiter(dePodcastRepublicDatabase.Text) + mySearchRec.Name) then
      if isValidDatabaseFilename(IncludeTrailingPathDelimiter(dePodcastRepublicDatabase.Text) + mySearchRec.Name) then
        sDatabaseFilename := IncludeTrailingPathDelimiter(dePodcastRepublicDatabase.Text) + mySearchRec.Name;
    iKeepGoing := FindNext(mySearchRec);
  end;
  SysUtils.FindClose(mySearchRec);
  Result := (sDatabaseFilename <> '');
end;

{ TfrmPR2CSV.actBuildCSVExecute }
procedure TfrmPR2CSV.actBuildCSVExecute(Sender: TObject);
var
  iRecord: integer;
  ftitle, fuuid, fpodItunesUrl, fPubDate, fpodUUID, fsavedFileName: TField;
  slEpisodeTitle, slEpisodeuuid, slEpisodepodItunesUrl: TStringList;
  slEpisodepubdate, slDownloaduuid, slDownloadSavedFilename: TStringList;
  slPodR6podUUID, slPodR6title: TStringList;

  dtPubDate: TDateTime;
  MyYear, MyMonth, MyDay: word;

  iCurrentImportCount, iTotalToImport, iEpisodeIndex, iPodcastIndex: integer;
  iFileIndex: integer;
  bKeepGoing: boolean;
  myGridRect: TGridRect;
begin
  memoCsv.Clear;
  pcMainPageControl.ActivePage := tsLog;

  bKeepGoing := ValidateBasicPathExists;
  if bKeepGoing then
    bKeepGoing := AttemptToGetDatabaseFilename;

  if bKeepGoing then
  begin
    sCSVResult.Clear;

    slEpisodeTitle := TStringList.Create;
    slEpisodeuuid := TStringList.Create;
    slEpisodepodItunesUrl := TStringList.Create;
    slEpisodepubdate := TStringList.Create;
    slDownloaduuid := TStringList.Create;
    slDownloadSavedFilename := TStringList.Create;
    slPodR6podUUID := TStringList.Create;
    slPodR6title := TStringList.Create;

    try
      mySQLite3Connection.Close; // Ensure the connection is closed when we start
      mySQLite3Connection.DatabaseName := sDatabaseFilename;

      mySQLite3Connection.Open;
      SQLTransaction1.Active := True;

      // 1. Fill our Episode related information
      memoCSV.Lines.Add('Importing information from "episode"...');
      iCurrentImportCount := 0;
      SQLQueryGeneric.PacketRecords := -1;
      SQLQueryGeneric.SQL.Clear;
      SQLQueryGeneric.SQL.Text := 'select * from "' + sEPISODE_KEYWORD + '"';
      SQLQueryGeneric.Open;
      try
        iTotalToImport := SQLQueryGeneric.RecordCount;
        for iRecord := 0 to pred(iTotalToImport) do
        begin
          ftitle := SQLQueryGeneric.FieldByName('episodeTitle');
          fuuid := SQLQueryGeneric.FieldByName('episodeUUID');
          fpodItunesUrl := SQLQueryGeneric.FieldByName('podUUID');
          fPubDate := SQLQueryGeneric.FieldByName('pubDate');
          if (ftitle <> nil) and (fPubDate <> nil) and (fpodItunesUrl <> nil) and
            (fuuid <> nil) then
          begin
            dtPubDate := ScanDateTime('ddd, dd mmm yyyy hh:nn:ss',
              fPubDate.AsString, myEnglishFormat, 1);
            DecodeDate(dtPubDate, MyYear, MyMonth, MyDay);
            slEpisodeTitle.Add(ftitle.AsString);
            slEpisodeuuid.Add(fuuid.AsString);
            slEpisodepodItunesUrl.Add(fpodItunesUrl.AsString);
            slEpisodepubdate.Add(Format('%4.4d-%2.2d-%2.2d', [MyYear, MyMonth, MyDay]));
            Inc(iCurrentImportCount);
          end
          else
          begin
            bKeepGoing := False;
            memoCSV.Lines.Add(Format('ERROR! Problem with record %d', [iRecord]));
          end;
          SQLQueryGeneric.Next;
        end;
      finally
        SQLQueryGeneric.Close;
      end;
      memoCSV.Lines.Add(Format('Information from "episode" imported (%d/%d)',
        [iCurrentImportCount, iTotalToImport]));

      // 2. Fill our Download related information
      memoCSV.Lines.Add('Importing information from "downloadV2"...');
      iCurrentImportCount := 0;
      SQLQueryGeneric.PacketRecords := -1;
      SQLQueryGeneric.SQL.Clear;
      SQLQueryGeneric.SQL.Text := 'select * from "' + sDOWNLOAD_KEYWORD + '"';
      SQLQueryGeneric.Open;
      try
        iTotalToImport := SQLQueryGeneric.RecordCount;
        for iRecord := 0 to pred(iTotalToImport) do
        begin
          fuuid := SQLQueryGeneric.FieldByName('episodeUUID');
          fsavedFileName := SQLQueryGeneric.FieldByName('savedFileName');
          if (fuuid <> nil) and (fsavedFileName <> nil) then
          begin
            slDownloaduuid.Add(fuuid.AsString);
            slDownloadSavedFilename.Add(fsavedFileName.AsString);
            Inc(iCurrentImportCount);
          end
          else
          begin
            bKeepGoing := False;
            memoCSV.Lines.Add(Format('ERROR! Problem with record %d', [iRecord]));
          end;
          SQLQueryGeneric.Next;
        end;
      finally
        SQLQueryGeneric.Close;
      end;
      memoCSV.Lines.Add(Format('Information from "download" imported (%d/%d)',
        [iCurrentImportCount, iTotalToImport]));

      // 3. Fill our PodR6 related information
      memoCSV.Lines.Add('Importing information from "' + sPOD_KEYWORD + '"...');
      iCurrentImportCount := 0;
      SQLQueryGeneric.PacketRecords := -1;
      SQLQueryGeneric.SQL.Clear;
      SQLQueryGeneric.SQL.Text := 'select * from "' + sPOD_KEYWORD + '"';
      SQLQueryGeneric.Open;
      try
        iTotalToImport := SQLQueryGeneric.RecordCount;
        for iRecord := 0 to pred(iTotalToImport) do
        begin
          fpodUUID := SQLQueryGeneric.FieldByName('podUUID');
          ftitle := SQLQueryGeneric.FieldByName('podName');
          if (fpodUUID <> nil) and (ftitle <> nil) then
          begin
            slPodR6podUUID.Add(fpodUUID.AsString);
            slPodR6title.Add(ftitle.AsString);
            Inc(iCurrentImportCount);
          end
          else
          begin
            bKeepGoing := False;
            memoCSV.Lines.Add(Format('ERROR! Problem with record %d', [iRecord]));
          end;
          SQLQueryGeneric.Next;
        end;
      finally
        SQLQueryGeneric.Close;
      end;
      memoCSV.Lines.Add(Format('Information from "podR6" imported (%d/%d)',
        [iCurrentImportCount, iTotalToImport]));

      {$IFDEF DEBUG}
      slEpisodeTitle.SaveToFile(IncludeTrailingPathDelimiter(
        deTargetDirectoryForCSV.Directory) + 'slEpisodeTitle.txt');
      slEpisodeuuid.SaveToFile(IncludeTrailingPathDelimiter(
        deTargetDirectoryForCSV.Directory) + 'slEpisodeuuid.txt');
      slEpisodepodItunesUrl.SaveToFile(IncludeTrailingPathDelimiter(
        deTargetDirectoryForCSV.Directory) + 'slEpisodepodItunesUrl.txt');
      slEpisodepubdate.SaveToFile(IncludeTrailingPathDelimiter(
        deTargetDirectoryForCSV.Directory) + 'slEpisodepubdate.txt');
      slDownloaduuid.SaveToFile(IncludeTrailingPathDelimiter(
        deTargetDirectoryForCSV.Directory) + 'slDownloaduuid.txt');
      slDownloadSavedFilename.SaveToFile(IncludeTrailingPathDelimiter(
        deTargetDirectoryForCSV.Directory) + 'slDownloadSavedFilename.txt');
      slPodR6podUUID.SaveToFile(IncludeTrailingPathDelimiter(
        deTargetDirectoryForCSV.Directory) + 'slPodV3podUUID.txt');
      slPodR6title.SaveToFile(IncludeTrailingPathDelimiter(
        deTargetDirectoryForCSV.Directory) + 'slPodR6title.txt');
      {$ENDIF}

      //slEpisodeuuid.SaveToFile('slEpisodeuuid.lst');
      //slDownloaduuid.SaveToFile('slDownloaduuid.lst');
      //slEpisodepodItunesUrl.SaveToFile('slEpisodepodItunesUrl.lst');
      //slPodR6podUUID.SaveToFile('slPodR6podUUID.lst');

      // 4. Output the actual data.
      InitFileGrid;
      AddCSVLine('Podcast', 'PubDate', 'Show', 'Filename', True);
      for iFileIndex := 0 to pred(slDownloadSavedFilename.Count) do
      begin
        iEpisodeIndex := slEpisodeuuid.IndexOf(slDownloaduuid.Strings[iFileIndex]);
        if iEpisodeIndex <> -1 then
        begin
          iPodcastIndex := slPodR6podUUID.IndexOf(
            slEpisodepodItunesUrl.Strings[iEpisodeIndex]);
          if iPodcastIndex <> -1 then
          begin
            AddCSVLine(slPodR6title.Strings[iPodcastIndex],
              slEpisodepubdate.Strings[iEpisodeIndex],
              slEpisodeTitle.Strings[iEpisodeIndex],
              slDownloadSavedFilename[iFileIndex]);
          end
          else
          begin
            memoCSV.Lines.Add('Problem finding podcast index for  ' +
              slDownloaduuid.Strings[iFileIndex]);
          end;
        end
        else
        begin
          memoCSV.Lines.Add('Problem finding pubDate of ' +
            slDownloaduuid.Strings[iFileIndex]);
        end;
      end;

    finally
      // 5. Release memory
      FreeAndNil(slEpisodeTitle);
      FreeAndNil(slEpisodeuuid);
      FreeAndNil(slEpisodepubdate);
      FreeAndNil(slDownloaduuid);
      FreeAndNil(slDownloadSavedFilename);
      FreeAndNil(slPodR6podUUID);
      FreeAndNil(slPodR6title);
    end;

    // 6. Saving to file result
    SaveCSVFile;
    sfFileGrid.Invalidate;
    sfFileGrid.SortOrder := soAscending;
    sfFileGrid.SortColRow(True, COL_Podcast);
    ArrangeFileGrid;

    if sfFileGrid.RowCount > 1 then
      sfFileGrid.Row := 1;

    if bKeepGoing then
    begin
      pcMainPageControl.ActivePage := tsFileGrid;
      Application.ProcessMessages;
      if sfFileGrid.CanFocus then self.FocusControl(sfFileGrid);
    end;
    ShowMessage('Done!');
  end;
end;

function TfrmPR2CSV.ValidateBasicPathExists: boolean;
begin
  Result := False;
  memoCSV.Lines.Add('Validating location of MP3 file...');
  dePodcastRepublicMP3Files.Text :=
    ExcludeTrailingPathDelimiter(dePodcastRepublicMP3Files.Text);
  if DirectoryExists(dePodcastRepublicMP3Files.Text) then
  begin
    memoCSV.Lines.Add('Good! Location for MP3 validated with success!');

    memoCSV.Lines.Add('Validating location for target CSV file...');
    deTargetDirectoryForCSV.Text :=
      ExcludeTrailingPathDelimiter(deTargetDirectoryForCSV.Text);
    if DirectoryExists(deTargetDirectoryForCSV.Text) then
    begin
      memoCSV.Lines.Add('Good! Location for target CSV validated with success!');
      Result := True;
    end
    else
    begin
      memoCSV.Lines.Add(
        'ERROR: Can''t find the directory specified for the target CSV file...');
      memoCSV.Lines.Add('       ' + deTargetDirectoryForCSV.Text);
    end;
  end
  else
  begin
    memoCSV.Lines.Add('ERROR: Can''t find the directory specified for the MP3 file...');
    memoCSV.Lines.Add('       ' + dePodcastRepublicMP3Files.Text);
  end;
end;

procedure TfrmPR2CSV.actExitExecute(Sender: TObject);
begin
  Close;
end;

{ TfrmPR2CSV.actRemoveFormulaExecute }
procedure TfrmPR2CSV.actRemoveFormulaExecute(Sender: TObject);
var
  iSelectionRange, iRow: integer;
begin
  for iSelectionRange := 0 to pred(sfFileGrid.SelectedRangeCount) do
    for iRow := sfFileGrid.SelectedRange[iSelectionRange].Top to
      sfFileGrid.SelectedRange[iSelectionRange].Bottom do
    begin
      sfFileGrid.Cells[COL_FormulaForOutputName, iRow] := '';
      sfFileGrid.Cells[COL_OutputName, iRow] := '';
    end;
  sfFileGrid.AutoSizeColumn(COL_FormulaForOutputName);
  sfFileGrid.AutoSizeColumn(COL_OutputName);
end;

{ TfrmPR2CSV.actRenameFilesExecute }
procedure TfrmPR2CSV.actRenameFilesExecute(Sender: TObject);
begin
  memoCSV.Clear;
  if isOkToDoRenaming then
  begin
    if DoActualRenamingJob then
    begin
      memoCSV.Lines.Add('Job was done successfully!');
      MessageDlg('There was no error!', mtInformation, [mbOK], 0);
    end
    else
    begin
      memoCSV.Lines.Add('Job was not completed successfully...');
      pcMainPageControl.ActivePage := tsLog;
      MessageDlg('ERROR: See the log for the error...', mtError, [mbOK], 0);
    end;
  end;
end;

procedure TfrmPR2CSV.actSelectAllExecute(Sender: TObject);
var
  myGridRect: TGridRect;
begin
  myGridRect.Top := 1;
  myGridRect.Left := 0;
  myGridRect.Bottom := pred(sfFileGrid.RowCount);
  myGridRect.Right := pred(sfFileGrid.ColCount);
  sfFileGrid.BeginUpdate;
  sfFileGrid.Selection := myGridRect;
  sfFileGrid.EndUpdate(True);
end;

{ TfrmPR2CSV.actSelectFromSamePodcastExecute }
procedure TfrmPR2CSV.actSelectFromSamePodcastExecute(Sender: TObject);
var
  iSelectedRow, iRow, iFirstMatchingRow, iLastMatchingRow: integer;
  sCurrentPodCast: string;
  SelectionGridRect: TGridRect;
begin
  iSelectedRow := sfFileGrid.Row;
  if iSelectedRow <> -1 then
  begin
    sCurrentPodCast := sfFileGrid.Cells[COL_Podcast, iSelectedRow];
    if sCurrentPodCast <> '' then
    begin
      iRow := 0;
      iFirstMatchingRow := -1;
      while (iRow < sfFileGrid.Rowcount) and (iFirstMatchingRow = -1) do
        if sfFileGrid.Cells[COL_Podcast, iRow] = sCurrentPodCast then
          iFirstMatchingRow := iRow
        else
          Inc(iRow);
      iLastMatchingRow := -1;
      while (iRow < sfFileGrid.Rowcount) and (iLastMatchingRow = -1) do
        if sfFileGrid.Cells[COL_Podcast, iRow] <> sCurrentPodCast then
          iLastMatchingRow := pred(iRow)
        else
          Inc(iRow);
      if iLastMatchingRow = -1 then
        iLastMatchingRow := pred(iRow);

      SelectionGridRect.Top := iFirstMatchingRow;
      SelectionGridRect.Left := 0;
      SelectionGridRect.Bottom := iLastMatchingRow;
      SelectionGridRect.Right := pred(sfFileGrid.ColCount);
      sfFileGrid.Selection := SelectionGridRect;

      sfFileGrid.TopRow := iFirstMatchingRow;
    end;
  end;
end;

{ TfrmPR2CSV.FormCreate }
procedure TfrmPR2CSV.FormCreate(Sender: TObject);
begin
  Caption := GetTitleBarCaption;
  sSeparator := ',';
  FPodcasFormulasList := TPodcasFormulasList.Create;
  LoadConfiguration;
  sCSVResult := TStringList.Create;
  InitFileGrid;
  InitMyEnglishFormat;
end;

{ TfrmPR2CSV.FormDestroy }
procedure TfrmPR2CSV.FormDestroy(Sender: TObject);
begin
  FreeAndNil(sCSVResult);
  SaveConfiguration;
end;

{ TfrmPR2CSV.LoadConfiguration }
procedure TfrmPR2CSV.LoadConfiguration;
var
  wsMaybeWindowStatus: TWindowState;
  sIniFilename, SectionScreenSize: string;
  AIniConfigFile: TIniFile;
begin
  sIniFilename := ExtractFilePath(ParamStr(0)) + 'PR2CSV.ini';
  SectionScreenSize := Format('W%dH%d', [Screen.Width, Screen.Height]);
  AIniConfigFile := TIniFile.Create(sIniFilename);
  try
    dePodcastRepublicDatabase.Directory := AIniConfigFile.ReadString(sMAINSECTION, 'PodcastRepublicDatabase', '');
    dePodcastRepublicMP3Files.Directory := AIniConfigFile.ReadString(sMAINSECTION, 'PodcastRepublicMP3Files', '');
    deTargetDirectoryForCSV.Directory := AIniConfigFile.ReadString(sMAINSECTION, 'TargetDirectoryForCSV', '');
    sLastFormula := AIniConfigFile.ReadString(sMAINSECTION, 'sLastFormula', '');
    pcMainPageControl.ActivePageIndex := AIniConfigFile.ReadIntegeR(sMAINSECTION, 'pcMainPageControl', 0);
    wsMaybeWindowStatus := TWindowState(AIniConfigFile.ReadInteger(SectionScreenSize, 'WindowSate', Ord(wsNormal)));
    self.FPodcasFormulasList.LoadFromConfiguration(AIniConfigFile);
    self.FPodcasFormulasList.SortByPodcastName;
    if wsMaybeWindowStatus = wsMaximized then
    begin
      Self.WindowState := wsMaybeWindowStatus;
    end
    else
    begin
      Self.WindowState := wsNormal;
      Self.Left := AIniConfigFile.ReadInteger(SectionScreenSize, 'WindowLeft', ((Screen.Width - Constraints.MinWidth) div 2));
      Self.Top := AIniConfigFile.ReadInteger(SectionScreenSize, 'WindowTop', ((Screen.Height - Constraints.MinHeight) div 2));
      Self.Width := AIniConfigFile.ReadInteger(SectionScreenSize, 'WindowWidth', Constraints.MinWidth);
      Self.Height := AIniConfigFile.ReadInteger(SectionScreenSize, 'WindowHeight', Constraints.MinHeight);
    end;
    cbSeparator.ItemIndex := AIniConfigFile.ReadInteger(sMAINSECTION, 'cbSeparator', 0);
  finally
    AIniConfigFile.Free;
  end;
end;

{ TfrmPR2CSV.SaveConfiguration }
procedure TfrmPR2CSV.SaveConfiguration;
var
  AIniConfigFile: TIniFile;
  sIniFilename, SectionScreenSize: string;
begin
  SectionScreenSize := Format('W%dH%d', [Screen.Width, Screen.Height]);
  sIniFilename := ExtractFilePath(ParamStr(0)) + 'PR2CSV.ini';
  AIniConfigFile := TIniFile.Create(sIniFilename);
  try
    AIniConfigFile.WriteString(sMAINSECTION, 'PodcastRepublicDatabase', dePodcastRepublicDatabase.Directory);
    AIniConfigFile.WriteString(sMAINSECTION, 'PodcastRepublicMP3Files', dePodcastRepublicMP3Files.Directory);
    AIniConfigFile.WriteString(sMAINSECTION, 'TargetDirectoryForCSV', deTargetDirectoryForCSV.Directory);
    AIniConfigFile.WriteString(sMAINSECTION, 'sLastFormula', sLastFormula);
    AIniConfigFile.WriteInteger(sMAINSECTION, 'pcMainPageControl', pcMainPageControl.ActivePageIndex);
    AIniConfigFile.WriteInteger(SectionScreenSize, 'WindowSate', Ord(Self.WindowState));
    if Self.WindowState = wsNormal then
    begin
      AIniConfigFile.WriteInteger(SectionScreenSize, 'WindowLeft', Left);
      AIniConfigFile.WriteInteger(SectionScreenSize, 'WindowTop', Top);
      AIniConfigFile.WriteInteger(SectionScreenSize, 'WindowWidth', Width);
      AIniConfigFile.WriteInteger(SectionScreenSize, 'WindowHeight', Height);
    end;
    AIniConfigFile.WriteInteger(sMAINSECTION, 'cbSeparator', cbSeparator.ItemIndex);

    self.FPodcasFormulasList.SaveToConfiguration(AIniConfigFile);
  finally
    AIniConfigFile.Free;
  end;
end;

{ TfrmPR2CSV.GetOutputCSVFilename }
function TfrmPR2CSV.GetOutputCSVFilename: string;
var
  FreezeTime: TDatetime;
  MyYear, MyMonth, MyDay, MyHour, MyMin, MySec, MyMillisec: word;
begin
  FreezeTime := now;
  DecodeDate(FreezeTime, MyYear, MyMonth, MyDay);
  DecodeTime(FreezeTime, MyHour, MyMin, MySec, MyMillisec);
  Result := IncludeTrailingPathDelimiter(deTargetDirectoryForCSV.Directory) +
    'PR2CSV_Output_File_' + Format('%4.4d-%2.2d-%2.2d@2.2d-%2.2d-%2.2d-%3.3d',
    [MyYear, MyMonth, MyDay, MyHour, MyMin, MySec, MyMillisec]) + '.csv';
end;

{ TfrmPR2CSV.SaveCSVFile }
procedure TfrmPR2CSV.SaveCSVFile;
begin
  sOutputCSVFilename := GetOutputCSVFilename;
  sCSVResult.SaveToFile(sOutputCSVFilename);
end;

procedure TfrmPR2CSV.InitMyEnglishFormat;
begin
  myEnglishFormat := FormatSettings;
  myEnglishFormat.ShortDayNames[1] := 'Sun';
  myEnglishFormat.ShortDayNames[2] := 'Mon';
  myEnglishFormat.ShortDayNames[3] := 'Tue';
  myEnglishFormat.ShortDayNames[4] := 'Wed';
  myEnglishFormat.ShortDayNames[5] := 'Thu';
  myEnglishFormat.ShortDayNames[6] := 'Fri';
  myEnglishFormat.ShortDayNames[7] := 'Sat';

  myEnglishFormat.ShortMonthNames[1] := 'Jan';
  myEnglishFormat.ShortMonthNames[2] := 'Feb';
  myEnglishFormat.ShortMonthNames[3] := 'Mar';
  myEnglishFormat.ShortMonthNames[4] := 'Apr';
  myEnglishFormat.ShortMonthNames[5] := 'May';
  myEnglishFormat.ShortMonthNames[6] := 'Jun';
  myEnglishFormat.ShortMonthNames[7] := 'Jul';
  myEnglishFormat.ShortMonthNames[8] := 'Aug';
  myEnglishFormat.ShortMonthNames[9] := 'Sep';
  myEnglishFormat.ShortMonthNames[10] := 'Oct';
  myEnglishFormat.ShortMonthNames[11] := 'Nov';
  myEnglishFormat.ShortMonthNames[12] := 'Dec';
end;

{ TfrmPR2CSV.InitFileGrid }
procedure TfrmPR2CSV.InitFileGrid;
begin
  sfFileGrid.Clean;
  sfFileGrid.RowCount := 1;
  sfFileGrid.Cells[COL_FileNumber, 0] := '#';
  sfFileGrid.Cells[COL_OriginalFilename, 0] := 'Original filename';
  sfFileGrid.Cells[COL_Podcast, 0] := 'Podcast';
  sfFileGrid.Cells[COL_EpisodeName, 0] := 'Episode name';
  sfFileGrid.Cells[COL_PubDate, 0] := 'Pub Date';
  sfFileGrid.Cells[COL_FormulaForOutputName, 0] := 'Formula for rename';
  sfFileGrid.Cells[COL_OutputName, 0] := 'New name';
  sfFileGrid.AutoSizeColumns;
end;

{ TfrmPR2CSV.ArrangeFileGrid }
procedure TfrmPR2CSV.ArrangeFileGrid;
begin
  sfFileGrid.AutoSizeColumns;
  //sfFileGrid.AutoSizeColumn(COL_FileNumber);
  //sfFileGrid.AutoSizeColumn(COL_PubDate);
end;

{ TfrmPR2CSV.IncludeRepetitionInFormula }
function TfrmPR2CSV.IncludeRepetitionInFormula(sFormula: string; iIndexRepetition: integer): string;
var
  iIndex, iPeriod: integer;
begin
  iPeriod := 0;

  iIndex := length(sFormula);
  if iIndex > 0 then
  begin
    while (iPeriod = 0) and (iIndex > 0) do
      if sFormula[iIndex] = '.' then
        iPeriod := iIndex
      else
        Dec(iIndex);
  end;

  if iPeriod <> 0 then
    Result := copy(sFormula, 1, pred(iPeriod)) + '(' + IntToStr(iIndexRepetition) + ')' + copy(sFormula, iPeriod, succ(length(sFormula) - iPeriod))
  else
    Result := sFormula + '(' + IntToStr(iIndexRepetition) + ')';
end;

{ EpureFilename }
//Invalid char in a filename: \ / : * ? " < > |
//                                    x
function EpureFilename(sFilename: string): string;
begin
  Result := StringReplace(sFilename, '\', '⧵', [rfReplaceAll]);
  Result := StringReplace(Result, '/', '∕', [rfReplaceAll]);
  Result := StringReplace(Result, ':', '꞉', [rfReplaceAll]);
  Result := StringReplace(Result, '*', '', [rfReplaceAll]);
  Result := StringReplace(Result, '?', '？', [rfReplaceAll]);
  Result := StringReplace(Result, '"', 'ʺ', [rfReplaceAll]);
  Result := StringReplace(Result, '<', '﹤', [rfReplaceAll]);
  Result := StringReplace(Result, '>', '﹥', [rfReplaceAll]);
  Result := StringReplace(Result, '|', '⏐', [rfReplaceAll]);
  Result := StringReplace(Result, '.', '．', [rfReplaceAll]);
end;

{ TfrmPR2CSV.actSetFormulaExecute }
procedure TfrmPR2CSV.actSetFormulaExecute(Sender: TObject);
var
  iSelectionRange, iRow, iIndexRepetition: integer;
  sMaybeFormula, sMaybeOutputFilename: string;

  slExistingOutputFilenames: TStringList;
begin
  slExistingOutputFilenames := TStringList.Create;
  try
    FPodcasFormulasList.FeedComboWithPossibleFormulas(frmFormulaRequester.cbFormula);
    if SameText(sfFileGrid.Cells[COL_FormulaForOutputName, sfFileGrid.Selection.Top], '') then
      frmFormulaRequester.cbFormula.Text := self.FPodcasFormulasList.GetBestFormulaForThisPodcastName(sfFileGrid.Cells[COL_Podcast, sfFileGrid.Selection.Top])
    else
      frmFormulaRequester.cbFormula.Text := sfFileGrid.Cells[COL_FormulaForOutputName, sfFileGrid.Selection.Top];
    frmFormulaRequester.Left := Left + ((Width - frmFormulaRequester.Width) div 2);
    frmFormulaRequester.Top := Top + ((Height - frmFormulaRequester.Height) div 2);
    if frmFormulaRequester.cbFormula.Text = '' then frmFormulaRequester.cbFormula.Text := sLastFormula;
    frmFormulaRequester.ActiveControl := frmFormulaRequester.cbFormula;
    if frmFormulaRequester.ShowModal = mrOk then
    begin
      FPodcasFormulasList.AddNewOrUpdateFormulaEntries(sfFileGrid.Cells[COL_Podcast, sfFileGrid.Selection.Top], frmFormulaRequester.cbFormula.Text);
      sLastFormula := frmFormulaRequester.cbFormula.Text;
      for iSelectionRange := 0 to pred(sfFileGrid.SelectedRangeCount) do
        for iRow := sfFileGrid.SelectedRange[iSelectionRange].Top to sfFileGrid.SelectedRange[iSelectionRange].Bottom do
        begin
          iIndexRepetition := 0;
          repeat
            sMaybeFormula := frmFormulaRequester.cbFormula.Text;
            if iIndexRepetition <> 0 then  sMaybeFormula := IncludeRepetitionInFormula(sMaybeFormula, iIndexRepetition);
            Inc(iIndexRepetition);
            sMaybeOutputFilename := sMaybeFormula;
            if sMaybeOutputFilename <> '' then
            begin
              sMaybeOutputFilename := StringReplace(sMaybeOutputFilename, KEYWORD_PUBDATE, sfFileGrid.Cells[COL_PubDate, iRow], [rfReplaceAll, rfIgnoreCase]);
              sMaybeOutputFilename := StringReplace(sMaybeOutputFilename, KEYWORD_EPISODE, EpureFilename(sfFileGrid.Cells[COL_EpisodeName, iRow]), [rfReplaceAll, rfIgnoreCase]);
            end;
          until (slExistingOutputFilenames.IndexOf(sMaybeOutputFilename) = -1);

          sfFileGrid.Cells[COL_FormulaForOutputName, iRow] := sMaybeFormula;
          sfFileGrid.Cells[COL_OutputName, iRow] := sMaybeOutputFilename;
          slExistingOutputFilenames.add(sMaybeOutputFilename);
        end;
      sfFileGrid.AutoSizeColumn(COL_FormulaForOutputName);
      sfFileGrid.AutoSizeColumn(COL_OutputName);
    end;
  finally
    FreeAndNil(slExistingOutputFilenames);
  end;
end;

{ TfrmPR2CSV.RefreshOutputNameForSelection }
procedure TfrmPR2CSV.RefreshOutputNameForSelection;
var
  iSelectionRange, iRow: integer;
  sResult: string;
begin
  for iSelectionRange := 0 to pred(sfFileGrid.SelectedRangeCount) do
    for iRow := sfFileGrid.SelectedRange[iSelectionRange].Top to
      sfFileGrid.SelectedRange[iSelectionRange].Bottom do
    begin
      sResult := sfFileGrid.Cells[COL_FormulaForOutputName, sfFileGrid.Selection.Top];
      if sResult <> '' then
      begin
        sResult := StringReplace(sResult, KEYWORD_PUBDATE,
          sfFileGrid.Cells[COL_PubDate, iRow], [rfReplaceAll, rfIgnoreCase]);
        sResult := StringReplace(sResult, KEYWORD_EPISODE,
          sfFileGrid.Cells[COL_EpisodeName, iRow], [rfReplaceAll, rfIgnoreCase]);
      end;
      sfFileGrid.Cells[COL_OutputName, iRow] := sResult;
    end;
end;

{ TfrmPR2CSV.GetTitleBarCaption }
// Code taken and modified from: http://forum.lazarus.freepascal.org/index.php?topic=12435.0
// [0] = Major version, [1] = Minor ver, [2] = Revision, [3] = Build Number
// The above values can be found in the menu: Project > Project Options > Version Info
function TfrmPR2CSV.GetTitleBarCaption: string;
var
  Info: TVersionInfo;
begin
  Info := TVersionInfo.Create;
  try
    Info.Load(HINSTANCE);
    Result := Format('%s v%d.%d (build %d)', [Application.Title,
      Info.FixedInfo.FileVersion[0], Info.FixedInfo.FileVersion[1],
      Info.FixedInfo.FileVersion[3]]);
  finally
    Info.Free;
  end;
end;

{ TfrmPR2CSV.isMP3FilePresent }
function TfrmPR2CSV.isMP3FilePresent(sFilename: string): boolean;
begin
  Result := fileexists(IncludeTrailingPathDelimiter(
    dePodcastRepublicMP3Files.Directory) + sFilename);
end;

{ TfrmPR2CSV.actValidateRenameCouldBeDoneExecute }
procedure TfrmPR2CSV.actValidateRenameCouldBeDoneExecute(Sender: TObject);
begin
  memoCSV.Clear;
  if isOkToDoRenaming then
  begin
    MessageDlg('There is no error, you may proceed!', mtInformation, [mbOK], 0);
  end
  else
  begin
    pcMainPageControl.ActivePage := tsLog;
    MessageDlg('ERROR: See the log for the error...', mtError, [mbOK], 0);
  end;
end;

{ TfrmPR2CSV.isOkToDoRenaming }
function TfrmPR2CSV.isOkToDoRenaming: boolean;
begin
  Result := False;
  if ValidateAllSourceFileExists then
  begin
    if ValidateAllDestinationFilenameAreDifferents then
    begin
      if ValidateAllDestinationFilenameDontExists then
      begin
        memoCSV.Lines.Add('We may proceed, there is no error!');
        Result := True;
      end;
    end;
  end;
end;

{ TfrmPR2CSV.ValidateAllSourceFileExists }
function TfrmPR2CSV.ValidateAllSourceFileExists: boolean;
const
  MAXTOLERABLEERROR = 20;
var
  iNumberOfMissing, iVerified, iRow: integer;
begin
  memoCSV.Lines.Add('Let''s verify all source file exist...');
  iNumberOfMissing := 0;
  iVerified := 0;
  iRow := 1;
  while (iNumberOfMissing < MAXTOLERABLEERROR) and (iRow < sfFileGrid.RowCount) do
  begin
    if (sfFileGrid.Cells[COL_FormulaForOutputName, iRow] <> '') or
      (sfFileGrid.Cells[COL_OutputName, iRow] <> '') then
    begin
      Inc(iVerified);
      if not (isMP3FilePresent(sfFileGrid.Cells[COL_OriginalFilename, iRow])) then
      begin
        memoCSV.Lines.Add('ERROR: For row with #' +
          sfFileGrid.Cells[COL_FileNumber, iRow] +
          ', we can''t see this file so we can''t rename it: ');
        memoCSV.Lines.Add('       ' + sfFileGrid.Cells[COL_OriginalFilename, iRow]);
        memoCSV.Lines.Add('It was expected to be in this directory: ' +
          dePodcastRepublicMP3Files.Directory);
        Inc(iNumberOfMissing);
        if iNumberOfMissing >= MAXTOLERABLEERROR then
          memoCSV.Lines.Add('Too many errors, we give up checking...');
      end;
    end;
    Inc(iRow);
  end;
  memoCSV.Lines.Add(Format('Number of file about to be renamed: %d', [iVerified]));
  memoCSV.Lines.Add(Format('     Number of source file missing: %d',
    [iNumberOfMissing]));

  if (iVerified = 0) and (iNumberOfMissing = 0) then
    memoCSV.Lines.Add('ERROR: It looks like there is now file to rename...');

  Result := ((iVerified > 0) and (iNumberOfMissing = 0));
end;

{ TfrmPR2CSV.ValidateAllDestinationFilenameAreDifferents }
function TfrmPR2CSV.ValidateAllDestinationFilenameAreDifferents: boolean;
const
  MAXTOLERABLEERROR = 20;
var
  slSourceRowNumber, slTargetFilename: TStringList;
  iMaybeAlreadyPresent, iNumberOfConflicting, iVerified, iRow: integer;

begin
  slSourceRowNumber := TStringList.Create;
  slTargetFilename := TStringList.Create;
  try
    memoCSV.Lines.Add('Let''s verify all destination filename are different...');
    iNumberOfConflicting := 0;
    iVerified := 0;
    iRow := 1;
    while (iNumberOfConflicting < MAXTOLERABLEERROR) and (iRow < sfFileGrid.RowCount) do
    begin
      if sfFileGrid.Cells[COL_OutputName, iRow] <> '' then
      begin
        Inc(iVerified);
        iMaybeAlreadyPresent :=
          slTargetFilename.indexof(sfFileGrid.Cells[COL_OutputName, iRow]);
        if iMaybeAlreadyPresent = -1 then
        begin
          slTargetFilename.add(sfFileGrid.Cells[COL_OutputName, iRow]);
          slSourceRowNumber.add(sfFileGrid.Cells[COL_FileNumber, iRow]);
        end
        else
        begin
          memoCSV.Lines.Add('ERROR: For row with #' +
            sfFileGrid.Cells[COL_FileNumber, iRow] + ', the resulting filename already exists!');
          memoCSV.Lines.Add('       This filename: ' +
            sfFileGrid.Cells[COL_OutputName, iRow]);
          memoCSV.Lines.Add(' Was already used for row #: ' +
            slSourceRowNumber.Strings[iMaybeAlreadyPresent]);
          Inc(iNumberOfConflicting);
          if iNumberOfConflicting >= MAXTOLERABLEERROR then
            memoCSV.Lines.Add('Too many errors, we give up checking...');
        end;
      end;
      Inc(iRow);
    end;
    memoCSV.Lines.Add(Format('Number of file about to be renamed: %d', [iVerified]));
    memoCSV.Lines.Add(Format('Number of target filename used at least twice: %d',
      [iNumberOfConflicting]));

    if (iVerified = 0) and (iNumberOfConflicting = 0) then
      memoCSV.Lines.Add('ERROR: It looks like there is now file to rename...');

    Result := ((iVerified > 0) and (iNumberOfConflicting = 0));

  finally
    slTargetFilename.Free;
    slSourceRowNumber.Free;
  end;
end;

{ TfrmPR2CSV.ValidateAllDestinationFilenameDontExists }
function TfrmPR2CSV.ValidateAllDestinationFilenameDontExists: boolean;
const
  MAXTOLERABLEERROR = 20;
var
  iNumberOfAlreadyThere, iVerified, iRow: integer;
begin
  memoCSV.Lines.Add('Let''s verify all destination file doesn''t exist...');
  iNumberOfAlreadyThere := 0;
  iVerified := 0;
  iRow := 1;
  while (iNumberOfAlreadyThere < MAXTOLERABLEERROR) and (iRow < sfFileGrid.RowCount) do
  begin
    if (sfFileGrid.Cells[COL_FormulaForOutputName, iRow] <> '') or
      (sfFileGrid.Cells[COL_OutputName, iRow] <> '') then
    begin
      Inc(iVerified);
      if (isMP3FilePresent(sfFileGrid.Cells[COL_OutputName, iRow])) then
      begin
        memoCSV.Lines.Add('ERROR: For row with #' +
          sfFileGrid.Cells[COL_FileNumber, iRow] +
          ', the destination filename already exist so we can''t rename it: ');
        memoCSV.Lines.Add('       ' + sfFileGrid.Cells[COL_OriginalFilename, iRow]);
        memoCSV.Lines.Add('Destination file that is already existing: ');
        memoCSV.Lines.Add('       ' + IncludeTrailingPathDelimiter(
          dePodcastRepublicMP3Files.Directory) + sfFileGrid.Cells[COL_OutputName, iRow]);
        Inc(iNumberOfAlreadyThere);
        if iNumberOfAlreadyThere >= MAXTOLERABLEERROR then
          memoCSV.Lines.Add('Too many errors, we give up checking...');
      end;
    end;
    Inc(iRow);
  end;
  memoCSV.Lines.Add(Format('Number of file about to be renamed: %d', [iVerified]));
  memoCSV.Lines.Add(Format('Number of destination that already exist: %d',
    [iNumberOfAlreadyThere]));

  if (iVerified = 0) and (iNumberOfAlreadyThere = 0) then
    memoCSV.Lines.Add('ERROR: It looks like there is now file to rename...');

  Result := ((iVerified > 0) and (iNumberOfAlreadyThere = 0));
end;

{ TfrmPR2CSV.DoActualRenamingJob }
function TfrmPR2CSV.DoActualRenamingJob: boolean;
var
  iNumberOfError, iNumberOfAttempt, iRow: integer;
  sFilenameIn, sFilenameOut: string;
begin
  memoCSV.Lines.Add('About to attempt to rename file...');
  iNumberOfError := 0;
  iNumberOfAttempt := 0;
  iRow := 1;
  while (iRow < sfFileGrid.RowCount) do
  begin
    if (sfFileGrid.Cells[COL_OutputName, iRow] <> '') then
    begin
      Inc(iNumberOfAttempt);
      sFilenameIn := IncludeTrailingPathDelimiter(dePodcastRepublicMP3Files.Directory) +
        sfFileGrid.Cells[COL_OriginalFilename, iRow];
      sFilenameOut := IncludeTrailingPathDelimiter(dePodcastRepublicMP3Files.Directory) +
        sfFileGrid.Cells[COL_OutputName, iRow];
      memoCSV.Lines.Add('About to rename file "' + sFilenameIn +
        '" by "' + sFilenameOut + '"...');
      if mbRenameFile(sFilenameIn, sFilenameOut) then
      begin
        memoCSV.Lines.Add('It worked!');
      end
      else
      begin
        Inc(iNumberOfError);
        memoCSV.Lines.Add('ERROR! It failed...');
      end;
    end;
    Inc(iRow);
  end;
  memoCSV.Lines.Add(Format('Number of rename attempt: %d', [iNumberOfAttempt]));
  memoCSV.Lines.Add(Format('Number of error: %d', [iNumberOfError]));

  if (iNumberOfAttempt = 0) and (iNumberOfError = 0) then
    memoCSV.Lines.Add('ERROR: It looks like there is now file to rename...');

  Result := ((iNumberOfAttempt > 0) and (iNumberOfError = 0));
end;

{ TfrmPR2CSV.sfFileGridDblClick }
procedure TfrmPR2CSV.sfFileGridDblClick(Sender: TObject);
begin
  actSelectFromSamePodcastExecute(actSelectFromSamePodcast);
  actSetFormulaExecute(actSetFormula);
end;

{ TfrmPR2CSV.actAddFormulasForKnownPodcastsExecute }
procedure TfrmPR2CSV.actAddFormulasForKnownPodcastsExecute(Sender: TObject);
var
  iRowIndex, iMaybeFormulaIndex, iFormulaAreadyPresent, iFormulaAdded, iFormulaMissing: integer;
begin
  iFormulaAreadyPresent := 0;
  iFormulaAdded := 0;
  iFormulaMissing := 0;

  for iRowIndex := 1 to pred(sfFileGrid.RowCount) do
  begin
    if SameText(sfFileGrid.Cells[COL_FormulaForOutputName, iRowIndex], '') then
    begin
      iMaybeFormulaIndex := PodcasFormulasList.GetFormulaIndexForPodcastName(sfFileGrid.Cells[COL_Podcast, iRowIndex]);
      if iMaybeFormulaIndex <> -1 then
      begin
        sfFileGrid.Cells[COL_FormulaForOutputName, iRowIndex] := PodcasFormulasList.PodcastFormulaEntry[iMaybeFormulaIndex].PodcastFormula;
        Inc(iFormulaAdded);
      end
      else
      begin
        Inc(iFormulaMissing);
      end;
    end
    else
    begin
      Inc(iFormulaAreadyPresent);
    end;
  end;

  if iFormulaAdded > 0 then
  begin
    GenerateAllFilenamesFromFormulas;
    sfFileGrid.AutoSizeColumn(COL_FormulaForOutputName);
    sfFileGrid.AutoSizeColumn(COL_OutputName);
  end;

  MessageDlg(Format('Already present: %d' + #$0A + 'Just added: %d' + #$0A + 'Still missing: %d', [iFormulaAreadyPresent, iFormulaAdded, iFormulaMissing]), mtInformation, [mbOK], 0);
end;

{ TfrmPR2CSV.GenerateAllFilenamesFromFormulas }
procedure TfrmPR2CSV.GenerateAllFilenamesFromFormulas;
var
  iSelectionRange, iRowIndex, iIndexRepetition: integer;
  sMaybeFormula, sMaybeOutputFilename: string;

  slExistingOutputFilenames: TStringList;
begin
  slExistingOutputFilenames := TStringList.Create;
  try
    for iRowIndex := 1 to pred(sfFileGrid.RowCount) do
    begin
      iIndexRepetition := 0;

      if not SameText(sfFileGrid.Cells[COL_OriginalFilename, iRowIndex], '') then
      begin
        if not SameText(sfFileGrid.Cells[COL_FormulaForOutputName, iRowIndex], '') then
        begin
          repeat
            sMaybeFormula := sfFileGrid.Cells[COL_FormulaForOutputName, iRowIndex];
            if iIndexRepetition <> 0 then  sMaybeFormula := IncludeRepetitionInFormula(sMaybeFormula, iIndexRepetition);
            Inc(iIndexRepetition);
            sMaybeOutputFilename := sMaybeFormula;
            if sMaybeOutputFilename <> '' then
            begin
              sMaybeOutputFilename := StringReplace(sMaybeOutputFilename, KEYWORD_PUBDATE, sfFileGrid.Cells[COL_PubDate, iRowIndex], [rfReplaceAll, rfIgnoreCase]);
              sMaybeOutputFilename := StringReplace(sMaybeOutputFilename, KEYWORD_EPISODE, EpureFilename(sfFileGrid.Cells[COL_EpisodeName, iRowIndex]), [rfReplaceAll, rfIgnoreCase]);
            end;
          until (slExistingOutputFilenames.IndexOf(sMaybeOutputFilename) = -1);

          sfFileGrid.Cells[COL_FormulaForOutputName, iRowIndex] := sMaybeFormula;
          sfFileGrid.Cells[COL_OutputName, iRowIndex] := sMaybeOutputFilename;
          slExistingOutputFilenames.add(sMaybeOutputFilename);
        end;
      end;
    end;
  finally
    FreeAndNil(slExistingOutputFilenames);
  end;
end;

{ TfrmPR2CSV.actEditManuallyExecute }
procedure TfrmPR2CSV.actEditManuallyExecute(Sender: TObject);
var
  sCurrentFilename: string;
begin
  if sfFileGrid.RowCount > 1 then
  begin
    if sfFileGrid.Row > 0 then
    begin
      sCurrentFilename := sfFileGrid.Cells[COL_OutputName, sfFileGrid.Row];
      if InputQuery('New wanted filename', 'Enter wanted filename', sCurrentFilename) then
      begin
        sfFileGrid.Cells[COL_OutputName, sfFileGrid.Row] := sCurrentFilename;
        sfFileGrid.AutoSizeColumn(COL_FormulaForOutputName);
        sfFileGrid.AutoSizeColumn(COL_OutputName);
      end;
    end;
  end;

end;



end.
