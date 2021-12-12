program PR2CSV;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}{$ENDIF}
  Interfaces, // this includes the LCL widgetset
  Forms, fpPR2CSV, fFormulaRequester, uMyFileRename, uPodcastFormulaRelated;

{$R *.res}

begin
  RequireDerivedFormResource:=True;
  Application.Initialize;
  Application.CreateForm(TfrmPR2CSV, frmPR2CSV);
  Application.CreateForm(TfrmFormulaRequester, frmFormulaRequester);
  frmFormulaRequester.PodcasFormulasList := frmPR2CSV.PodcasFormulasList;
  Application.Run;
end.

