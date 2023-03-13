// Note: If the Delphi IDE inserts a "{$R *.TLB}" here, just delete it again.
// This project does not use a type library.
library DragDropHandlerShellExt;

{%File 'readme.txt'}

uses
  ComServ,
  DragDropHandlerMain in 'DragDropHandlerMain.pas' {DataModuleDragDropHandler: TDataModule};

{$R *.res}

exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer,
  DllInstall;

begin
end.
