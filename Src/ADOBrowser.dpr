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
