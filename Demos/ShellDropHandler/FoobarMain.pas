unit FoobarMain;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Menus, StdCtrls, ExtCtrls;

type
  TFormFileList = class(TForm)
    Panel1: TPanel;
    MemoFileList: TMemo;
    MainMenu1: TMainMenu;
    MenuFile: TMenuItem;
    MenuFileOpen: TMenuItem;
    MenuFileSave: TMenuItem;
    N1: TMenuItem;
    MenuFileExit: TMenuItem;
    MenuSetup: TMenuItem;
    MenuSetupRegister: TMenuItem;
    MenuSetupUnregister: TMenuItem;
    Memo1: TMemo;
    OpenDialog1: TOpenDialog;
    SaveDialog1: TSaveDialog;
    MenuFileSaveAs: TMenuItem;
    MenuFileNew: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure MenuSetupRegisterClick(Sender: TObject);
    procedure MenuFileExitClick(Sender: TObject);
    procedure MenuSetupUnregisterClick(Sender: TObject);
    procedure MenuFileOpenClick(Sender: TObject);
    procedure MenuFileSaveClick(Sender: TObject);
    procedure MenuFileSaveAsClick(Sender: TObject);
    procedure MemoFileListChange(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure MenuFileNewClick(Sender: TObject);
  private
    FDirty: boolean;
    FFileName: string;
    procedure SetDirty(const Value: boolean);
    procedure SetFileName(const Value: string);
  public
    property FileName: string read FFileName write SetFileName;
    property Dirty: boolean read FDirty write SetDirty;
    procedure LoadFile(const AFilename: string);
    function SaveFile(const AFilename: string): boolean;
    function Clear: boolean;
  end;

var
  FormFileList: TFormFileList;

implementation

{$R *.DFM}

uses
  ComObj;

const
  RegistryRoot: HKEY = HKEY_CURRENT_USER;
  RegistryPrefix: string = 'SOFTWARE\Classes\';

resourcestring
  sFileClass = 'FoobarFile';
  sFileType = 'Foobar File List';
  sFileExtension = '.foobar';

  sTitle = 'Foobar List Editor - %s';
  sNewFile = 'new list';
  sSaveMods = 'Your modifications has not been saved.'+#13+'Save now?';
  sUnregisterNotice = 'Remember to also unregister the drop handler DLL';
  sRegisterNotice = 'Remember to also register the drop handler DLL';

procedure TFormFileList.FormCreate(Sender: TObject);

  procedure LoadFileList(const List: string);
  var
    Files: TStringList;
  begin
    Files := TStringList.Create;
    try
      Files.LoadFromFile(List, TEncoding.Unicode);
      MemoFileList.Lines.AddStrings(Files);
    finally
      Files.Free;
    end;
  end;

var
  i: integer;
  ParamFileName: string;
begin
  FileName := '';

  // Display command line (for debug purposes).
  Memo1.Lines.Text := GetCommandLine;

  if (ParamCount > 0) then
  begin
    // First parameter is file list.
    LoadFile(ParamStr(1));

    // Additional parameters are file names which should be added to the list.
    // If a filename starts with @ it indicates that the file contains a list of
    // file names which should be added to the list.
    for i := 2 to ParamCount do
    begin
      ParamFileName := ParamStr(i);
      if (Copy(ParamFileName, 1, 1) = '@') then
        LoadFileList(Copy(ParamFileName, 2, MaxInt))
      else
        MemoFileList.Lines.Add(ParamFileName);
    end;
  end;

  // Determine if the file association has already been registered and modify
  // the register menu items accordingly.
  MenuSetupRegister.Enabled := (GetRegStringValue(RegistryPrefix+sFileClass+'\DefaultIcon', '', RegistryRoot) = '');
  MenuSetupUnregister.Enabled := not MenuSetupRegister.Enabled;
end;

procedure TFormFileList.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  CanClose := Clear;
end;

procedure TFormFileList.MenuSetupRegisterClick(Sender: TObject);
begin
  // Register file association.
  CreateRegKey(RegistryPrefix+sFileExtension, '', sFileClass, RegistryRoot);
  CreateRegKey(RegistryPrefix+sFileExtension+'\ShellNew', 'NullFile', '', RegistryRoot);
  CreateRegKey(RegistryPrefix+sFileClass, '', sFileType, RegistryRoot);
  CreateRegKey(RegistryPrefix+sFileClass+'\shell\open\command', '', Application.ExeName+' "%1"', RegistryRoot);
  CreateRegKey(RegistryPrefix+sFileClass+'\DefaultIcon', '', Application.ExeName+',0', RegistryRoot);
  MenuSetupRegister.Enabled := False;
  MenuSetupUnregister.Enabled := True;
  if (GetRegStringValue(RegistryPrefix+sFileClass+'\shellex\DropHandler', '', RegistryRoot) = '') then
    ShowMessage(sRegisterNotice);
end;

procedure TFormFileList.MenuSetupUnregisterClick(Sender: TObject);
begin
  // Unregister file association.
  DeleteRegKey(RegistryPrefix+sFileClass+'\DefaultIcon', RegistryRoot);
  DeleteRegKey(RegistryPrefix+sFileClass+'\shell\open\command', RegistryRoot);
  DeleteRegKey(RegistryPrefix+sFileClass+'\shell\open', RegistryRoot);
  DeleteRegKey(RegistryPrefix+sFileClass+'\shell', RegistryRoot);
  DeleteRegKey(RegistryPrefix+sFileClass, RegistryRoot);
  DeleteRegKey(RegistryPrefix+sFileExtension+'\ShellNew', RegistryRoot);
  DeleteRegKey(RegistryPrefix+sFileExtension, RegistryRoot);
  MenuSetupRegister.Enabled := True;
  MenuSetupUnregister.Enabled := False;
  if (GetRegStringValue(RegistryPrefix+sFileClass+'\shellex\DropHandler', '', RegistryRoot) <> '') then
    ShowMessage(sUnregisterNotice);
end;

procedure TFormFileList.MenuFileExitClick(Sender: TObject);
begin
  Close;
end;

procedure TFormFileList.MenuFileNewClick(Sender: TObject);
begin
  Clear;
end;

procedure TFormFileList.MenuFileOpenClick(Sender: TObject);
begin
  if (Clear) then
  begin
    OpenDialog1.Filename := FileName;
    if (OpenDialog1.Execute) then
      LoadFile(OpenDialog1.Filename);
  end;
end;

procedure TFormFileList.MenuFileSaveAsClick(Sender: TObject);
begin
  SaveFile('');
end;

procedure TFormFileList.MenuFileSaveClick(Sender: TObject);
begin
  SaveFile(FileName);
end;

procedure TFormFileList.LoadFile(const AFilename: string);
begin
  MemoFileList.Lines.LoadFromFile(AFilename);
  FileName := AFilename;
  Dirty := False;
end;

function TFormFileList.SaveFile(const AFilename: string): boolean;
begin
  Result := True;
  if (AFilename = '') then
  begin
    SaveDialog1.Filename := FileName;
    if (SaveDialog1.Execute) then
      FileName := SaveDialog1.Filename
    else
      Result := False;
  end else
    FileName := AFilename;

  if (Result) then
  begin
    MemoFileList.Lines.SaveToFile(Filename);
    Dirty := False;
  end;
end;

function TFormFileList.Clear: boolean;
var
  Answer: word;
begin
  Result := True;
  // Check for unsaved changes and prompt.
  if (Dirty) then
  begin
    Answer := MessageDlg(sSaveMods, mtConfirmation, [mbYes, mbNo, mbCancel], 0);
    case Answer of
      mrYes:
        Result := SaveFile(FileName);
      mrCancel:
        Result := False;
    end;
  end;

  if (Result) then
  begin
    MemoFileList.Lines.Clear;
    FileName := '';
    Dirty := False;
  end;
end;

procedure TFormFileList.MemoFileListChange(Sender: TObject);
begin
  Dirty := True;
end;

procedure TFormFileList.SetDirty(const Value: boolean);
begin
  // Enable the "Save" menu item if the file has been modified and we have a
  // file name for it.
  FDirty := Value;
  MenuFileSave.Enabled := FDirty and (FileName <> '');
end;

procedure TFormFileList.SetFileName(const Value: string);
begin
  FFileName := Value;
  if (FFileName <> '') then
  begin
    Caption := Format(sTitle, [FFileName]);
    MenuFileSave.Enabled := Dirty;
  end else
  begin
    Caption := Format(sTitle, [sNewFile]);
    MenuFileSave.Enabled := False;
  end;
end;

end.

