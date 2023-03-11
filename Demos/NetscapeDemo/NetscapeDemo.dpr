program NetscapeDemo;

{%File 'readme.txt'}

uses
  Forms,
  Main in 'Main.pas' {FormMain};

{$R *.RES}

begin
  Application.Initialize;
  Application.Title := 'Netscape target demo';
  Application.CreateForm(TFormMain, FormMain);
  Application.Run;
end.
