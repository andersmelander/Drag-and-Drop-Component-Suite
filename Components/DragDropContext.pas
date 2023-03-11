unit DragDropContext;

// -----------------------------------------------------------------------------
// Project:         Drag and Drop Component Suite.
// Module:          DragDropContext
// Description:     Implements Context Menu Handler Shell Extensions.
// Version:         4.0
// Date:            18-MAY-2001
// Target:          Win32, Delphi 5-6
// Authors:         Anders Melander, anders@melander.dk, http://www.melander.dk
// Copyright        © 1997-2001 Angus Johnson & Anders Melander
// -----------------------------------------------------------------------------
interface

uses
  DragDrop,
  DragDropComObj,
  Menus,
  ShlObj,
  ActiveX,
  Windows,
  Classes;

{$include DragDrop.inc}

type
////////////////////////////////////////////////////////////////////////////////
//
//		TDropContextMenu
//
////////////////////////////////////////////////////////////////////////////////
// Partially based on Borland's ShellExt demo.
////////////////////////////////////////////////////////////////////////////////
// A typical shell context menu handler session goes like this:
// 1. User selects one or more files and right clicks on them.
//    The files must of a file type which has a context menu handler registered.
// 2. The shell loads the context menu handler module.
// 3. The shell instantiates the registered context menu handler object as an
//    in-process COM server.
// 4. The IShellExtInit.Initialize method is called with a data object which
//    contains the dragged data.
// 5. The IContextMenu.QueryContextMenu method is called to populate the popup
//    menu.
//    TDropContextMenu uses the PopupMenu property to populate the shell context
//    menu.
// 6. If the user chooses one of the context menu menu items we have supplied,
//    the IContextMenu.InvokeCommand method is called.
//    TDropContextMenu locates the corresponding TMenuItem and fires the menu
//    items OnClick event.
// 7. The shell unloads the context menu handler module (usually after a few
//    seconds).
////////////////////////////////////////////////////////////////////////////////
  TDropContextMenu = class(TInterfacedComponent, IShellExtInit, IContextMenu)
  private
    FContextMenu: TPopupMenu;
    FMenuOffset: integer;
    FDataObject: IDataObject;
    FOnPopup: TNotifyEvent;
    FFiles: TStrings;
    procedure SetContextMenu(const Value: TPopupMenu);
  protected
    procedure Notification(AComponent: TComponent;
      Operation: TOperation); override;
    { IShellExtInit }
     function Initialize(pidlFolder: PItemIDList; lpdobj: IDataObject;
      hKeyProgID: HKEY): HResult; stdcall;
    { IContextMenu }
    function QueryContextMenu(Menu: HMENU; indexMenu, idCmdFirst, idCmdLast,
      uFlags: UINT): HResult; stdcall;
    function InvokeCommand(var lpici: TCMInvokeCommandInfo): HResult; stdcall;
    function GetCommandString(idCmd, uType: UINT; pwReserved: PUINT;
      pszName: LPSTR; cchMax: UINT): HResult; stdcall;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    property DataObject: IDataObject read FDataObject;
    property Files: TStrings read FFiles;
  published
    property ContextMenu: TPopupMenu read FContextMenu write SetContextMenu;
    property OnPopup: TNotifyEvent read FOnPopup write FOnPopup;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TDropContextMenuFactory
//
////////////////////////////////////////////////////////////////////////////////
// COM Class factory for TDropContextMenu.
////////////////////////////////////////////////////////////////////////////////
  TDropContextMenuFactory = class(TShellExtFactory)
  protected
    function HandlerRegSubKey: string; virtual;
  public
    procedure UpdateRegistry(Register: Boolean); override;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		Component registration
//
////////////////////////////////////////////////////////////////////////////////
procedure Register;


////////////////////////////////////////////////////////////////////////////////
//
//		Misc.
//
////////////////////////////////////////////////////////////////////////////////


////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
//
//			IMPLEMENTATION
//
////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
implementation

uses
  DragDropFile,
  DragDropPIDL,
  Registry,
  ComObj,
  SysUtils;

////////////////////////////////////////////////////////////////////////////////
//
//		Component registration
//
////////////////////////////////////////////////////////////////////////////////

procedure Register;
begin
  RegisterComponents(DragDropComponentPalettePage, [TDropContextMenu]);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		Utilities
//
////////////////////////////////////////////////////////////////////////////////


////////////////////////////////////////////////////////////////////////////////
//
//		TDropContextMenu
//
////////////////////////////////////////////////////////////////////////////////
constructor TDropContextMenu.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FFiles := TStringList.Create;
end;

destructor TDropContextMenu.Destroy;
begin
  FFiles.Free;
  inherited Destroy;
end;

function TDropContextMenu.GetCommandString(idCmd, uType: UINT;
  pwReserved: PUINT; pszName: LPSTR; cchMax: UINT): HResult;
var
  ItemIndex: integer;
begin
  ItemIndex := integer(idCmd);
  // Make sure we aren't being passed an invalid argument number
  if (ItemIndex >= 0) and (ItemIndex < FContextMenu.Items.Count) then
  begin
    if (uType = GCS_HELPTEXT) then
      // return help string for menu item.
      StrLCopy(pszName, PChar(FContextMenu.Items[ItemIndex].Hint), cchMax);
    Result := NOERROR;
  end else
    Result := E_INVALIDARG;
end;

function TDropContextMenu.InvokeCommand(var lpici: TCMInvokeCommandInfo): HResult;
var
  ItemIndex: integer;
begin
  Result := E_FAIL;

  // Make sure we are not being called by an application
  if (FContextMenu = nil) or (HiWord(Integer(lpici.lpVerb)) <> 0) then
    Exit;

  ItemIndex := LoWord(lpici.lpVerb);
  // Make sure we aren't being passed an invalid argument number
  if (ItemIndex < 0) or (ItemIndex >= FContextMenu.Items.Count) then
  begin
    Result := E_INVALIDARG;
    Exit;
  end;

  // Execute the menu item specified by lpici.lpVerb.
  try
    try
      FContextMenu.Items[ItemIndex].Click;
      Result := NOERROR;
    except
      on E: Exception do
      begin
        Windows.MessageBox(0, PChar(E.Message), 'Error',
          MB_OK or MB_ICONEXCLAMATION or MB_SYSTEMMODAL);
        Result := E_UNEXPECTED;
      end;
    end;
  finally
    FDataObject := nil;
    FFiles.Clear;
  end;
end;

function TDropContextMenu.QueryContextMenu(Menu: HMENU; indexMenu, idCmdFirst,
  idCmdLast, uFlags: UINT): HResult;
var
  i: integer;
  Last: integer;
  Flags: UINT;

  function IsLine(Item: TMenuItem): boolean;
  begin
  {$ifdef VER13_PLUS}
    Result := Item.IsLine;
  {$else}
    Result := Item.Caption = '-';
  {$endif}
  end;

begin
  Last := 0;

  if (FContextMenu <> nil) and (((uFlags and $0000000F) = CMF_NORMAL) or
     ((uFlags and CMF_EXPLORE) <> 0)) then
  begin
    FMenuOffset := idCmdFirst;
    for i := 0 to FContextMenu.Items.Count-1 do
      if (FContextMenu.Items[i].Visible) then
      begin
        Flags := MF_STRING or MF_BYPOSITION;
        if (not FContextMenu.Items[i].Enabled) then
          Flags := Flags or MF_GRAYED;
        if (IsLine(FContextMenu.Items[i])) then
          Flags := Flags or MF_SEPARATOR;
        // Add one menu item to context menu
        InsertMenu(Menu, indexMenu, Flags, FMenuOffset+i,
          PChar(FContextMenu.Items[i].Caption));
        inc(indexMenu);
        Last := i+1;
      end;
  end else
    FMenuOffset := 0;

  // Return number of menu items added
  Result := MakeResult(SEVERITY_SUCCESS, FACILITY_NULL, Last)
end;

function TDropContextMenu.Initialize(pidlFolder: PItemIDList;
  lpdobj: IDataObject; hKeyProgID: HKEY): HResult;
begin
  FFiles.Clear;

  if (lpdobj = nil) then
  begin
    Result := E_INVALIDARG;
    Exit;
  end;

  // Save a reference to the source data object.
  FDataObject := lpdobj;

  // Extract source file names and store them in a string list.
  with TFileDataFormat.Create(nil) do
    try
      if GetData(DataObject) then
        FFiles.Assign(Files);
    finally
      Free;
    end;

  if (Assigned(FOnPopup)) then
    FOnPopup(Self);

  Result := NOERROR;
end;

procedure TDropContextMenu.SetContextMenu(const Value: TPopupMenu);
begin
  if (Value <> FContextMenu) then
  begin
    if (FContextMenu <> nil) then
      FContextMenu.RemoveFreeNotification(Self);
    FContextMenu := Value;
    if (Value <> nil) then
      Value.FreeNotification(Self);
  end;
end;

procedure TDropContextMenu.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  if (Operation = opRemove) and (AComponent = FContextMenu) then
    FContextMenu := nil;
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
//
//		TDropContextMenuFactory
//
////////////////////////////////////////////////////////////////////////////////
function TDropContextMenuFactory.HandlerRegSubKey: string;
begin
  Result := 'ContextMenuHandlers';
end;

procedure TDropContextMenuFactory.UpdateRegistry(Register: Boolean);
var
  ClassIDStr: string;
begin
  ClassIDStr := GUIDToString(ClassID);
  if Register then
  begin
    inherited UpdateRegistry(Register);
    CreateRegKey(FileClass+'\shellex\'+HandlerRegSubKey+'\'+ClassName, '', ClassIDStr);

    if (Win32Platform = VER_PLATFORM_WIN32_NT) then
      with TRegistry.Create do
        try
          RootKey := HKEY_LOCAL_MACHINE;
          OpenKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions', True);
          OpenKey('Approved', True);
          WriteString(ClassIDStr, Description);
        finally
          Free;
        end;
  end else
  begin
    if (Win32Platform = VER_PLATFORM_WIN32_NT) then
      with TRegistry.Create do
        try
          RootKey := HKEY_LOCAL_MACHINE;
          OpenKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions', True);
          OpenKey('Approved', True);
          DeleteKey(ClassIDStr);
        finally
          Free;
        end;
    DeleteRegKey(FileClass+'\shellex\'+HandlerRegSubKey+'\'+ClassName);
    inherited UpdateRegistry(Register);
  end;
end;

end.
