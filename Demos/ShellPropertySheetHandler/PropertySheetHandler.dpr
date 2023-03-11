library PropertySheetHandler;


{$R 'PropertySheetHandlerDialogTemplate.res' 'PropertySheetHandlerDialogTemplate.rc'}
{$R *.RES}

uses
  System.Win.ComServ,
  FormPropertySheet in 'FormPropertySheet.pas' {PropertySheetForm},
  ModulePropertySheetHandler in 'ModulePropertySheetHandler.pas' {DataModulePropertySheetHandler: TDataModule};

exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer,
  DllInstall;

begin
end.
