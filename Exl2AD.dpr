program Exl2AD;

uses
  Forms,
  Main in 'Main.pas' {Exl2ADfm},
  ActiveDs_TLB in 'C:\Users\User\Documents\Embarcadero\Studio\19.0\Imports\ActiveDs_TLB.pas',
  uLogManager in 'uLogManager.pas',
  uSingletonTemplate in 'uSingletonTemplate.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TExl2ADfm, Exl2ADfm);
  Application.Run;
end.
