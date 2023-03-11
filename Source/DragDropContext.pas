unit DragDropContext;
// -----------------------------------------------------------------------------
// Project:         Drag and Drop Component Suite.
// Module:          DragDropContext
// Description:     Implements Context Menu Handler Shell Extensions.
// Version:         4.2
// Date:            05-APR-2008
// Target:          Win32, Delphi 5-2007
// Authors:         Anders Melander, anders@melander.dk, http://melander.dk
// Copyright        © 1997-1999 Angus Johnson & Anders Melander
//                  © 2000-2008 Anders Melander
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
  TDropContextMenu = class(TInterfacedComponent, IShellExtInit, IContextMenu,
    IContextMenu2, IContextMenu3)
  private
    FContextMenu: TPopupMenu;
    FMenuOffset: integer;
    FDataObject: IDataObject;
    FOnPopup: TNotifyEvent;
    FFiles: TStrings;
    FMenuHandle: HMenu;
    FOwnerDraw: boolean;
  protected
    procedure SetContextMenu(const Value: TPopupMenu);
    procedure Notification(AComponent: TComponent;
      Operation: TOperation); override;
    function GetMenuItem(Index: integer): TMenuItem;
    procedure DrawMenuItem(var DrawItemStruct: TDrawItemStruct);
    procedure MeasureItem(var MeasureItemStruct: TMeasureItemStruct);
    property MenuHandle: HMenu read FMenuHandle;
    property OwnerDraw: boolean read FOwnerDraw write FOwnerDraw;
    { IShellExtInit }
     function Initialize(pidlFolder: PItemIDList; lpdobj: IDataObject;
      hKeyProgID: HKEY): HResult; stdcall;
    { IContextMenu }
    function QueryContextMenu(Menu: HMENU; indexMenu, idCmdFirst, idCmdLast,
      uFlags: UINT): HResult; stdcall;
    function InvokeCommand(var lpici: TCMInvokeCommandInfo): HResult; stdcall;
    function GetCommandString(idCmd, uType: UINT; pwReserved: PUINT;
      pszName: LPSTR; cchMax: UINT): HResult; stdcall;
    { IContextMenu2 }
    function HandleMenuMsg(uMsg: UINT; WParam, LParam: Integer): HResult; stdcall;
    { IContextMenu3 }
    function HandleMenuMsg2(uMsg: UINT; wParam, lParam: Integer;
      var lpResult: Integer): HResult; stdcall;
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
  Dialogs,
  DragDropFile,
  DragDropPIDL,
  Registry,
  ComObj,
  SysUtils,
  Messages,
  Graphics, // TCanvas
  Controls, // TControlCanvas
  Forms; // Screen


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
  FOwnerDraw := True;
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
  MenuItem: TMenuItem;
begin
  ItemIndex := integer(idCmd);
  MenuItem := GetMenuItem(ItemIndex);
  // Make sure we aren't being passed an invalid argument number
  if (MenuItem <> nil) then
  begin
    case uType of
      GCS_HELPTEXTA:
        // return ANSI help string for menu item.
        StrLCopy(pszName, PChar(MenuItem.Hint), cchMax);
      GCS_HELPTEXTW:
        // return UNICODE help string for menu item.
        StringToWideChar(MenuItem.Hint, PWideChar(pszName),
          cchMax);
      GCS_VERBA:
        pszName^ := #0;
      GCS_VERBW:
        PWideChar(pszName)^ := #0;
    end;
    Result := NOERROR;
  end else
    Result := E_INVALIDARG;
end;

function TDropContextMenu.InvokeCommand(var lpici: TCMInvokeCommandInfo): HResult;
var
  ItemIndex: integer;
  MenuItem: TMenuItem;
begin
  Result := E_FAIL;

  // Make sure we are not being called by an application
  if (FContextMenu = nil) or (HiWord(Integer(lpici.lpVerb)) <> 0) then
    Exit;

  ItemIndex := LoWord(lpici.lpVerb);
  MenuItem := GetMenuItem(ItemIndex);
  // Make sure we aren't being passed an invalid argument number
  if (MenuItem <> nil) then
  begin
    // Execute the menu item specified by lpici.lpVerb.
    try
      try
        MenuItem.Click;
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
      FFiles.Clear;
    end;
  end else
    Result := E_INVALIDARG;
end;

function TDropContextMenu.QueryContextMenu(Menu: HMENU; indexMenu, idCmdFirst,
  idCmdLast, uFlags: UINT): HResult;
var
  i: integer;
  MenuID, NextMenuID: integer;
  NewItem: UINT;
  Flags: UINT;

  function IsLine(Item: TMenuItem): boolean;
  begin
  {$ifdef VER13_PLUS}
    Result := Item.IsLine;
  {$else}
    Result := Item.Caption = '-';
  {$endif}
  end;

  function SetMenuID(Handle: HMENU; MenuIndex: integer; MenuID: UINT): boolean;
  var
    MenuItemInfo: TMenuItemInfo;
    Buffer: array[0..79] of Char;
  begin
    Result := False;
    MenuItemInfo.cbSize := 44; // Required for Windows 95
    MenuItemInfo.fMask := MIIM_ID;
    MenuItemInfo.dwTypeData := Buffer;
    MenuItemInfo.cch := SizeOf(Buffer);
    if (GetMenuItemInfo(Handle, MenuIndex, True, MenuItemInfo)) then
    begin
      MenuItemInfo.fMask := MIIM_ID;
      MenuItemInfo.wID := MenuID;
      Result := SetMenuItemInfo(Handle, MenuIndex, True, MenuItemInfo);
    end;
  end;

  function SetMenuItemID(MenuItem: TMenuItem; var MenuID: integer): UINT;
  var
    i: integer;
  begin
    if (SetMenuID(MenuItem.Parent.Handle, MenuItem.MenuIndex, MenuID)) and (OwnerDraw) then
    begin
      Result := MenuID;
      inc(MenuID);
      for i := 0 to MenuItem.Count-1 do
        SetMenuItemID(MenuItem.Items[i], MenuID);
    end else
      Result := 0;
  end;

begin
  if (FContextMenu <> nil) and (((uFlags and $0000000F) = CMF_NORMAL) or
     ((uFlags and CMF_EXPLORE) <> 0)) then
  begin
    FMenuOffset := idCmdFirst;
    MenuID := FMenuOffset;
    NextMenuID := idCmdFirst;
    for i := 0 to FContextMenu.Items.Count-1 do
    begin
      MenuID := SetMenuItemID(FContextMenu.Items[i], NextMenuID);
      if (FContextMenu.Items[i].Visible) then
      begin
        Flags := MF_BYPOSITION;
        if (OwnerDraw) and (FContextMenu.Items[i].Count > 0) then
        begin
          NewItem := FContextMenu.Items[i].Handle;
          Flags := Flags or MF_POPUP;
          // Note: Apparently top level popup items doesn't support owner draw.
          // If MF_OWNERDRAW is specified together with MF_POPUP, the menu item
          // will not be drawn.
          // Flags := Flags or MF_OWNERDRAW;
        end else
        begin
          NewItem := MenuID;
          Flags := Flags or MF_STRING;
          if (OwnerDraw) and (not IsLine(FContextMenu.Items[i])) then
            Flags := Flags or MF_OWNERDRAW;
        end;
        if (not FContextMenu.Items[i].Enabled) then
          Flags := Flags or MF_GRAYED;
        if (IsLine(FContextMenu.Items[i])) then
          Flags := Flags or MF_SEPARATOR;

        // Add one menu item to context menu
        InsertMenu(Menu, integer(indexMenu), Flags, NewItem,
          PChar(FContextMenu.Items[i].Caption));
        // Set menu ID - required to work around problem with duplicate entries
        // in IE4+ Explorer file menu.
        SetMenuID(Menu, indexMenu, MenuID);

        // Special handling of menu bitmap for top level popup menu item.
        if ((not OwnerDraw) or (Flags and MF_POPUP <> 0)) and
          ((not FContextMenu.Items[i].Bitmap.Empty) or
            ((FContextMenu.Images <> nil) and
            (FContextMenu.Items[i].ImageIndex <> -1))) then
        begin
          // We have to use SetMenuItemBitmaps for top level menu items. If we
          // rely on Delphi's default handling of menu bitmap images, the image
          // will be drawn to far to the right.
          if (FContextMenu.Items[i].Bitmap.Empty) then
            FContextMenu.Images.GetBitmap(FContextMenu.Items[i].ImageIndex,
              FContextMenu.Items[i].Bitmap);
          SetMenuItemBitmaps(Menu, indexMenu, MF_BYPOSITION,
            FContextMenu.Items[i].Bitmap.Handle, 0);
        end;
        inc(indexMenu);
      end;
    end;
  end else
  begin
    FMenuOffset := 0;
    MenuID := 0;
  end;

  // Return number of menu items added
  Result := MakeResult(SEVERITY_SUCCESS, FACILITY_NULL,
    MenuID-FMenuOffset);
end;

function TDropContextMenu.HandleMenuMsg(uMsg: UINT; WParam,
  LParam: Integer): HResult;
begin
  Result := HandleMenuMsg2(uMsg, WParam, LParam, PInteger(nil)^);
end;

function TDropContextMenu.HandleMenuMsg2(uMsg: UINT; wParam,
  lParam: Integer; var lpResult: Integer): HResult;
begin
  case uMsg of
    WM_INITMENUPOPUP:
      FMenuHandle := wParam;
    WM_DRAWITEM:
      DrawMenuItem(PDrawItemStruct(lParam)^);
    WM_MEASUREITEM:
      MeasureItem(PMeasureItemStruct(lParam)^);
  end;
  Result := S_OK;
end;

function TDropContextMenu.Initialize(pidlFolder: PItemIDList;
  lpdobj: IDataObject; hKeyProgID: HKEY): HResult;
begin
  FFiles.Clear;

  // Save a reference to the source data object.
  FDataObject := lpdobj;
  try

    // TODO : We probably need an event to customize the initialization.

    // Extract source file names and store them in a string list.
    // Note that not all shell objects provide us with a IDataObject (i.e. the
    // Directory\Background object).
    if (DataObject <> nil) then
      with TFileDataFormat.Create(nil) do
        try
          if GetData(DataObject) then
            FFiles.Assign(Files);
        finally
          Free;
        end;

    if (Assigned(FOnPopup)) then
      FOnPopup(Self);

  finally
    FDataObject := nil;
  end;

  Result := NOERROR;
end;

{$ifndef VER13_PLUS}
type
  TComponentCracker = class(TComponent);
{$endif}

procedure TDropContextMenu.SetContextMenu(const Value: TPopupMenu);
begin
  if (Value <> FContextMenu) then
  begin
{$ifdef VER13_PLUS}
    if (FContextMenu <> nil) then
      FContextMenu.RemoveFreeNotification(Self);
{$else}
    if (FContextMenu <> nil) then
      TComponentCracker(FContextMenu).Notification(Self, opRemove);
{$endif}
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

function TDropContextMenu.GetMenuItem(Index: integer): TMenuItem;

  function DoGetMenuItem(Menu: TMenuItem): TMenuItem;
  var
    i: integer;
  begin
    i := 0;
    Result := nil;
    while (Result = nil) and (i < Menu.Count) do
    begin
      if (Index = 0) then
        Result := Menu.Items[i];
      Dec(Index);
      if (Result = nil) and (Menu.Items[i].Count > 0) and (OwnerDraw) then
        Result := DoGetMenuItem(Menu.Items[i]);
      inc(i);
    end;
  end;

begin
  if (FContextMenu <> nil) then
    Result := DoGetMenuItem(FContextMenu.Items)
  else
    Result := nil;
end;

type
  TMenuItemCracker = class(TMenuItem);
{$ifndef VER13_PLUS}
  TOwnerDrawState = set of (odSelected, odGrayed, odDisabled, odChecked,
    odFocused, odDefault, odHotLight, odInactive, odNoAccel, odNoFocusRect,
    odReserved1, odReserved2, odComboBoxEdit);
{$endif}

{$ifndef VER13_PLUS}
function GetMenuFont: HFONT;
var
  NonClientMetrics: TNonClientMetrics;
begin
  NonClientMetrics.cbSize := sizeof(NonClientMetrics);
  if SystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0, @NonClientMetrics, 0) then
    Result := CreateFontIndirect(NonClientMetrics.lfMenuFont)
  else
    Result := GetStockObject(SYSTEM_FONT);
end;
{$endif}

procedure TDropContextMenu.DrawMenuItem(var DrawItemStruct: TDrawItemStruct);
var
  ItemIndex: integer;
  MenuItem: TMenuItem;
  Canvas: TCanvas;
  SaveIndex: Integer;
  Win98Plus: Boolean;
  State: TOwnerDrawState;
begin
  // Make sure context is valid.
  if (FContextMenu = nil) or (DrawItemStruct.CtlType <> ODT_MENU) then
    Exit;

  ItemIndex := integer(DrawItemStruct.itemID)-FMenuOffset;
  MenuItem := GetMenuItem(ItemIndex);
  // Make sure we aren't being passed an invalid item ID.
  if (MenuItem <> nil) then
  begin
    Canvas := TControlCanvas.Create;
    try
      SaveIndex := SaveDC(DrawItemStruct.hDC);
      try
        Canvas.Handle := DrawItemStruct.hDC;
{$ifdef VER13_PLUS}
        Canvas.Font := Screen.MenuFont;
{$else}
        Canvas.Font.Handle := GetMenuFont;
{$endif}

        State := TOwnerDrawState(LongRec(DrawItemStruct.itemState).Lo);

        Win98Plus := (Win32MajorVersion > 4) or
          ((Win32MajorVersion = 4) and (Win32MinorVersion > 0));

        (*
        ** The Following code works around two problems :
        ** 1) The shell context menu always draws text with a horizontal offset
        **    of 16. Since it is impossible to specify the horizontal offset of
        **    a VCL menu item without resorting to a completely ownerdrawn item,
        **    we have to fool the VCL into drawing the menu item at the correct
        **    position. We do this by shifting the draw rect 8 pixels to the
        **    right for top level menu items. When this happens...
        ** 2) ...the menu item background will also be shifted 8 pixels to the
        **    right. Because of this we have draw the background manually.
        ** 3) The VCLs drawing of selected menu bitmaps interferes with 1 & 2,
        **    so we have to fool the VCL into believing that the menu item isn't
        **    selected even if it is.
        **
        ** Normally we would just call Menus.DrawMenuItem and let it set up the
        ** canvas for us, but because we have to disable TMenuItem's drawing of
        ** the selected state, we must do it manually here and instead call
        ** TMenuItem.AdvancedDrawItem.
        **
        ** We could get away with using Menus.DrawMenuItem for sub menu items,
        ** but its easier to use the same code for all items. A side effect of
        ** this is that our selected menu bitmaps will look like other shell
        ** context menu bitmaps, instead of Delphi's button image look.
        *)
        if (odSelected in State) then
        begin
          Canvas.Brush.Color := clHighlight;
          Canvas.Font.Color := clHighlightText;
          Exclude(State, odSelected);
        end else
        if Win98Plus and (odInactive in State) then
        begin
          Canvas.Brush.Color := clMenu;
          Canvas.Font.Color := clGrayText;
        end else
        begin
          Canvas.Brush.Color := clMenu;
          Canvas.Font.Color := clMenuText;
        end;

        if ((MenuItem.Parent <> nil) and (MenuItem.Parent.Parent = nil)) and
          not((MenuItem.GetParentMenu <> nil) and
           (MenuItem.GetParentMenu.OwnerDraw or (MenuItem.GetParentMenu.Images <> nil)) and
{$ifdef VER13_PLUS}
           (Assigned(MenuItem.OnAdvancedDrawItem) or Assigned(MenuItem.OnDrawItem))) then
{$else}
           (Assigned(MenuItem.OnDrawItem))) then
{$endif}
        begin
          Canvas.FillRect(DrawItemStruct.rcItem);
          if (MenuItem.GetParentMenu.Images = nil) or (MenuItem.ImageIndex = -1) then // Work around: ImageList images are drawn with different rules...
            Inc(DrawItemStruct.rcItem.Left, 8);
        end;

        // TODO : Unless menu item is ownerdraw we should handle the draw internally instead of relying on TMenuItem's draw code.
{$ifdef VER13_PLUS}
        TMenuItemCracker(MenuItem).AdvancedDrawItem(Canvas,
          DrawItemStruct.rcItem, State, False);
{$else}
        TMenuItemCracker(MenuItem).DrawItem(Canvas,
          DrawItemStruct.rcItem, (odSelected in State));
{$endif}
        // Menus.DrawMenuItem(MenuItem, Canvas, DrawItemStruct.rcItem, State);
      finally
        Canvas.Handle := 0;
        RestoreDC(DrawItemStruct.hDC, SaveIndex);
      end;
    finally
      Canvas.Free;
    end;
  end;
end;

procedure TDropContextMenu.MeasureItem(var MeasureItemStruct: TMeasureItemStruct);
var
  ItemIndex: integer;
  MenuItem: TMenuItem;
  Canvas: TCanvas;
  SaveIndex: Integer;
  DC: HDC;
begin
  // Make sure context is valid.
  if (FContextMenu = nil) or (MeasureItemStruct.CtlType <> ODT_MENU) then
    Exit;

  ItemIndex := integer(MeasureItemStruct.itemID)-FMenuOffset;
  MenuItem := GetMenuItem(ItemIndex);
  // Make sure we aren't being passed an invalid item ID.
  if (MenuItem <> nil) then
  begin
    DC := GetWindowDC(GetForegroundWindow);
    try
      Canvas := TControlCanvas.Create;
      try
        SaveIndex := SaveDC(DC);
        try
          Canvas.Handle := DC;
{$ifdef VER13_PLUS}
          Canvas.Font := Screen.MenuFont;
{$else}
          Canvas.Font.Handle := GetMenuFont;
{$endif}
          TMenuItemCracker(MenuItem).MeasureItem(Canvas, Integer(MeasureItemStruct.itemWidth),
            Integer(MeasureItemStruct.itemHeight));
        finally
          Canvas.Handle := 0;
          RestoreDC(DC, SaveIndex);
        end;
      finally
        Canvas.Free;
      end;
    finally
      ReleaseDC(GetForegroundWindow, DC);
    end;
  end else
  begin
    MeasureItemStruct.itemWidth := 150;
    MeasureItemStruct.itemHeight := 20;
  end;
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
          if OpenKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved',
            False) then
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
          if OpenKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved',
            False) then
            DeleteKey(ClassIDStr);
        finally
          Free;
        end;

    DeleteDefaultRegValue(FileClass+'\shellex\'+HandlerRegSubKey+'\'+ClassName);
    DeleteEmptyRegKey(FileClass+'\shellex\'+HandlerRegSubKey+'\'+ClassName, True);
    inherited UpdateRegistry(Register);
  end;
end;

end.

