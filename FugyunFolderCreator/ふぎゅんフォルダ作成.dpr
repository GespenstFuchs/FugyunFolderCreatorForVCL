program �ӂ����t�H���_�쐬;

uses
  Vcl.Forms,
  MainUnit in 'MainUnit.pas' {MainForm},
  TMemoHistoryUnit in 'TMemoHistoryUnit.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
