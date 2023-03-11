unit ModulePropertySheetHandler;
(*
** Portions copyright © 2023 François Piette.
**
** This demo is based, in part, on François Piette's property sheet handler experiment.
*)
interface

uses
  Vcl.Controls,
  WinApi.Windows, WinApi.Messages, WinApi.ActiveX,
  WinApi.ShlObj, WinApi.CommCtrl,
  System.Win.ComObj, System.Win.Registry,
  System.Types, System.Classes, System.SysUtils;

type
  TDataModulePropertySheetHandler = class(TDataModule,
    IUnknown,
    IShellExtInit,
    IShellPropSheetExt
  )
  strict private
    FDataObject: IDataObject;
    FFiles: TStrings;
    FFolderPIDL: PItemIDList;
  private
    // IShellExtInit
    function ShellExtInit_Initialize(pidlfolder: PItemIDList; lpdobj: IDataObject; Hkeyprogid : HKEY): HResult; stdcall;
    function IShellExtInit.Initialize = ShellExtInit_Initialize;
    // IShellPropSheetExt
    function AddPages(lpfnAddPage: TFNAddPropSheetPage; lParam: LPARAM): HResult; stdcall;
    function ReplacePage(uPageID: TEXPPS; lpfnReplaceWith: TFNAddPropSheetPage; lParam: LPARAM): HResult; stdcall;
  protected
    procedure SetFolder(Folder: PItemIDList);
    function GetFolder: string;
    function CreatePropertySheetForm(AParent: HWND; AFormClass: TWinControlClass): TWinControl;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    property DataObject: IDataObject read FDataObject;
    property Files: TStrings read FFiles;
    property FolderPIDL: PItemIDList read FFolderPIDL;
    property Folder: string read GetFolder;
  end;

implementation

{%CLASSGROUP 'Vcl.Controls.TControl'}

{$R *.dfm}

uses
  System.StrUtils, System.Win.ComServ, WinApi.ShellAPI,

  DragDrop,
  DragDropFile,
  DragDropPIDL,
  DragDropComObj,

  FormPropertySheet;

const
  // CLSID for this shell extension.
  // Modify this for your own shell extensions (press [Ctrl]+[Shift]+G in
  // the IDE editor to gererate a new CLSID).
  CLSID_PropertySheetHandler: TGUID = '{1067C264-8B1F-4B22-919F-DB5191C359CB}';
  sFileClass = 'pasfile';
  sFileExtension = '.pas';
  sClassName = 'DelphiPropSheetShellExt';

resourcestring
  // Description of our shell extension.
  sDescription = 'Drag and Drop Component Suite property sheet demo';

////////////////////////////////////////////////////////////////////////////////

constructor TDataModulePropertySheetHandler.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FFiles := TStringList.Create;
end;

destructor TDataModulePropertySheetHandler.Destroy;
begin
  SetFolder(nil);
  FFiles.Free;
  inherited Destroy;
end;

////////////////////////////////////////////////////////////////////////////////

function TDataModulePropertySheetHandler.CreatePropertySheetForm(AParent: HWND; AFormClass: TWinControlClass): TWinControl;
begin
  Result := AFormClass.CreateParented(AParent);
  try

    // Save a pointer to the property sheet form into the parent dialog user
    // data so we can get it later.
    SetWindowLongPtr(AParent, GWLP_USERDATA, NativeInt(Result));

    Result.Top := 0;
    Result.Left := 0;
  except
    Result.Free;
    raise;
  end;
end;

////////////////////////////////////////////////////////////////////////////////

function TDataModulePropertySheetHandler.ShellExtInit_Initialize(pidlfolder: PItemIDList; lpdobj: IDataObject; Hkeyprogid: HKEY): HResult;
begin
  Result := NOERROR;

  FFiles.Clear;
  SetFolder(pidlFolder);

  // Save a reference to the source data object.
  FDataObject := lpdobj;
  try

    // Extract source file names and store them in a string list.
    // Note that not all shell objects provide us with a IDataObject (e.g. the
    // Directory\Background object).
    if (DataObject <> nil) then
      with TFileDataFormat.Create(dfdConsumer) do
        try
          if GetData(DataObject) then
            FFiles.Assign(Files);
        finally
          Free;
        end;

  finally
    FDataObject := nil;
  end;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TDataModulePropertySheetHandler.SetFolder(Folder: PItemIDList);
begin
  if (FFolderPIDL <> Folder) then
  begin
    if (FFolderPIDL <> nil) then
      coTaskMemFree(FFolderPIDL);

    FFolderPIDL := nil;

    if (Folder <> nil) then
      FFolderPIDL := ILClone(Folder);
  end;
end;

function TDataModulePropertySheetHandler.GetFolder: string;
begin
  Result := GetFullPathFromPIDL(FolderPIDL);
end;

////////////////////////////////////////////////////////////////////////////////

function PropertySheetDlgProc(hDlg: HWND; uMessage: UINT; wParam: WPARAM; lParam: LPARAM): Boolean; stdcall;
var
  PropSheetPage: PPropSheetPage;
  PropertySheetHandler: TDataModulePropertySheetHandler;
  PropertySheetForm: TPropertySheetForm;
  UserData: NativeInt;
  Rect: TRect;
begin
  Result := TRUE;

  case uMessage of
    WM_INITDIALOG:
      begin
        // Must return TRUE to direct the system to set the keyboard
        // focus to the control specified by wParam
        PropSheetPage := PPropSheetPage(lParam);
        PropertySheetHandler := TDataModulePropertySheetHandler(PropSheetPage.lParam);

        // Create a Delphi form and parent it to the dialog
        PropertySheetForm := PropertySheetHandler.CreatePropertySheetForm(hDlg, TPropertySheetForm) as TPropertySheetForm;
        PropertySheetForm.SetFiles(PropertySheetHandler.Files);
        PropertySheetForm.Visible := True;
      end;

    WM_SIZE:
      begin
        UserData := GetWindowLongPtr(hDLG, GWLP_USERDATA);
        PropertySheetForm := TPropertySheetForm(UserData);

        GetWindowRect(hDlg, Rect);
        PropertySheetForm.Width := Rect.Width;
        PropertySheetForm.Height := Rect.Height;
      end;
  else
    Result := FALSE;
  end;
end;

////////////////////////////////////////////////////////////////////////////////


function PropertySheetCallback(hWnd: HWND; uMessage: UINT; var PSP : TPropSheetPage): UINT; stdcall;
begin
  case uMessage of
    PSPCB_RELEASE:
      if PSP.lParam <> 0 then
        // Allow the class to be released.
        TDataModulePropertySheetHandler(PSP.lParam)._Release;
  end;
  Result := 1;
end;

////////////////////////////////////////////////////////////////////////////////

function TDataModulePropertySheetHandler.AddPages(lpfnAddPage: TFNAddPropSheetPage; lParam: LPARAM): HResult;
var
//  PageTitle: string;
  PropSheetPage: TPropSheetPage;
  hPage: HPropSheetPage;
resourcestring
  sPageTitle = 'Delphi Property Sheet';  // Displayed on page tab
begin
  Result := NOERROR;

//  PageTitle := sPageTitle;

  PropSheetPage := Default(TPropSheetPage);
  PropSheetPage.dwSize      := SizeOf(PropSheetPage);
  PropSheetPage.dwFlags     := PSP_USETITLE or PSP_USECALLBACK;
  PropSheetPage.hInstance   := HInstance;
  PropSheetPage.pszTemplate := 'IDD_PROPERTY_PAGE';
  PropSheetPage.pszTitle    := PChar(sPageTitle);
  PropSheetPage.pfnDlgProc  := @PropertySheetDlgProc;
  PropSheetPage.pfnCallback := @PropertySheetCallback;
  PropSheetPage.lParam      := IntPtr(Self);   // points to ourself

  hPage := CreatePropertySheetPage(PropSheetPage);

  if (hPage <> nil) then
  begin
    if not lpfnAddPage(hPage, lParam) then
      DestroyPropertySheetPage(hPage);
  end else
    Result := E_UNEXPECTED;

  // Prevent the class from being destroyed before the COM server is destroyed.
  _AddRef;
end;


////////////////////////////////////////////////////////////////////////////////

function TDataModulePropertySheetHandler.ReplacePage(uPageID: TEXPPS; lpfnReplaceWith: TFNAddPropSheetPage; lParam: LPARAM): HResult;
begin
  Result := E_NOTIMPL;
end;


////////////////////////////////////////////////////////////////////////////////

type
  TPropertySheetHandlerFactory = class(TDropContextMenuFactory)
  protected
    function HandlerRegSubKey: string; override;
  end;

function TPropertySheetHandlerFactory.HandlerRegSubKey: string;
begin
  Result := 'PropertySheetHandlers';
end;

////////////////////////////////////////////////////////////////////////////////

initialization
  TPropertySheetHandlerFactory.Create(ComServer, TDataModulePropertySheetHandler,
    CLSID_PropertySheetHandler, sClassName, sDescription, sFileClass,
    sFileExtension, ciMultiInstance);
end.

