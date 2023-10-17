// Program was compiled with Delphi 7.0 (Build 8.1)
// Program uses Turbo Power Systools library (available on SourceForge)
program ADOBrowser;

uses
  Forms,
  uBrowseMemo in 'uBrowseMemo.pas' {frmBrowseMemo},
  uBrowser in 'uBrowser.pas' {Form_Browser},
  OpenADOTableDialog in 'OpenADOTableDialog.pas' {frmOpenADOTable};

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TForm_Browser, Form_Browser);
  Application.Run;
end.
