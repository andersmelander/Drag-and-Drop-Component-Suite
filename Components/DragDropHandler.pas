unit DragDropHandler;

// -----------------------------------------------------------------------------
// Project:         Drag and Drop Component Suite.
// Module:          DragDropHandler
// Description:     Implements Drop and Drop Context Menu Shell Extenxions
//                  (a.k.a. drag-and-drop handlers).
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
  DragDropContext,
  Menus,
  ShlObj,
  ActiveX,
  Windows,
  Classes;

{$include DragDrop.inc}

type
////////////////////////////////////////////////////////////////////////////////
//
//		TDragDropHandler
//
////////////////////////////////////////////////////////////////////////////////
// A typical drag-and-drop handler session goes like this:
// 1. User right-drags (drags with the right mouse button) and drops one or more
//    source files which has a registered drag-and-drop handler.
// 2. The shell loads the drag-and-drop handler module.
// 3. The shell instantiates the registered drag drop handler object as an
//    in-process COM server.
// 4. The IShellExtInit.Initialize method is called with the name of the target
//    folder and a data object which contains the dragged data.
//    The target folder name is stored in the TDragDropHandler.TargetFolder
//    property as a string and in the TargetPIDL property as a PIDL.
// 5. The IContextMenu.QueryContextMenu method is called to populate the popup
//    menu.
//    TDragDropHandler uses the PopupMenu property to populate the drag-and-drop
//    context menu.
// 6. If the user chooses one of the context menu items we have supplied, the
//    IContextMenu.InvokeCommand method is called.
//    TDragDropHandler locates the corresponding TMenuItem and fires the menu
//    items OnClick event.
// 7. The shell unloads the drag-and-drop handler module (usually after a few
//    seconds).
////////////////////////////////////////////////////////////////////////////////
  TDragDropHandler = class(TDropContextMenu, IShellExtInit, IContextMenu)
  private
    FFolderPIDL: pItemIDList;
  protected
    function GetFolder: string;
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
    destructor Destroy; override;
    function GetFolderPIDL: pItemIDList; // Caller must free PIDL!
    property Folder: string read GetFolder;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TDragDropHandlerFactory
//
////////////////////////////////////////////////////////////////////////////////
// COM Class factory for TDragDropHandler.
////////////////////////////////////////////////////////////////////////////////
  TDragDropHandlerFactory = class(TDropContextMenuFactory)
  protected
    function HandlerRegSubKey: string; override;
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
  RegisterComponents(DragDropComponentPalettePage, [TDragDropHandler]);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		Utilities
//
////////////////////////////////////////////////////////////////////////////////


////////////////////////////////////////////////////////////////////////////////
//
//		TDragDropHandler
//
////////////////////////////////////////////////////////////////////////////////
destructor TDragDropHandler.Destroy;
begin
  if (FFolderPIDL <> nil) then
    ShellMalloc.Free(FFolderPIDL);
  inherited Destroy;
end;

function TDragDropHandler.GetCommandString(idCmd, uType: UINT;
  pwReserved: PUINT; pszName: LPSTR; cchMax: UINT): HResult;
begin
  Result := inherited GetCommandString(idCmd, uType, pwReserved, pszName, cchMax);
end;

function TDragDropHandler.GetFolder: string;
begin
  Result := GetFullPathFromPIDL(FFolderPIDL);
end;

function TDragDropHandler.GetFolderPIDL: pItemIDList;
begin
  Result := CopyPIDL(FFolderPIDL);
end;

function TDragDropHandler.InvokeCommand(var lpici: TCMInvokeCommandInfo): HResult;
begin
  Result := E_FAIL;
  try
    Result := inherited InvokeCommand(lpici);
  finally
    if (Result <> E_FAIL) then
    begin
      ShellMalloc.Free(FFolderPIDL);
      FFolderPIDL := nil;
    end;
  end;
end;

function TDragDropHandler.QueryContextMenu(Menu: HMENU; indexMenu, idCmdFirst,
  idCmdLast, uFlags: UINT): HResult;
begin
  Result := inherited QueryContextMenu(Menu, indexMenu, idCmdFirst,
    idCmdLast, uFlags);
end;

function TDragDropHandler.Initialize(pidlFolder: PItemIDList;
  lpdobj: IDataObject; hKeyProgID: HKEY): HResult;
begin
  if (pidlFolder <> nil) then
  begin
    // Copy target folder PIDL.
    FFolderPIDL := CopyPIDL(pidlFolder);
    Result := inherited Initialize(pidlFolder, lpdobj, hKeyProgID);
  end else
    Result := E_INVALIDARG;
end;

////////////////////////////////////////////////////////////////////////////////
//
//		TDragDropHandlerFactory
//
////////////////////////////////////////////////////////////////////////////////
function TDragDropHandlerFactory.HandlerRegSubKey: string;
begin
  Result := 'DragDropHandlers';
end;

end.
