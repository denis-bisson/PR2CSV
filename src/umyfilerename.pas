unit uMyFileRename;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils;

function mbRenameFile(const OldName: String; NewName: String): Boolean;

implementation

uses
  Windows, LazUTF8;

// The following code is coming directly from source code of Double Commander.
// https://sourceforge.net/p/doublecmd/code/HEAD/tree/
// ..\components\doublecmd\dcwindows.pas
function UTF16LongName(const FileName: String): UnicodeString;
var
  Temp: PWideChar;
begin
  if Pos('\\', FileName) = 0 then
    Result := '\\?\' + UTF8Decode(FileName)
  else begin
    Result := '\\?\UNC\' + UTF8Decode(Copy(FileName, 3, MaxInt));
  end;
  Temp := Pointer(Result) + 4;
  while Temp^ <> #0 do
  begin
    if Temp^ = '/' then Temp^:= '\';
    Inc(Temp);
  end;
  if ((Temp - 1)^ = DriveSeparator) then Result:= Result + '\';
end;

// The following code is coming directly from source code of Double Commander.
// https://sourceforge.net/p/doublecmd/code/HEAD/tree/
// ..\components\doublecmd\dcosutils.pas
function mbRenameFile(const OldName: String; NewName: String): Boolean;
var
  wOldName,
  wNewName: UnicodeString;
begin
  wNewName:= UTF16LongName(NewName);
  wOldName:= UTF16LongName(OldName);
  Result:= MoveFileExW(PWChar(wOldName), PWChar(wNewName), MOVEFILE_REPLACE_EXISTING);
end;

end.

