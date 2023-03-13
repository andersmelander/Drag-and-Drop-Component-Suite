// Note: If the Delphi IDE inserts a "{$R *.TLB}" here, just delete it again.
// This project does not use a type library.
library ContextMenuHandlerShellExt;

{$R 'About.res' 'About.rc'}
{$R *.res}

uses
  ComServ,
  ContextMenuHandlerMain in 'ContextMenuHandlerMain.pas' {DataModuleContextMenuHandler: TDataModule};

exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer,
  DllInstall;

begin
end.
