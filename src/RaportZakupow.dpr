program RaportZakupow;

uses
  Vcl.Forms,
  MainForm in 'MainForm.pas' {frmMain};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := 'Raport zakupów i sprzedaży';
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
