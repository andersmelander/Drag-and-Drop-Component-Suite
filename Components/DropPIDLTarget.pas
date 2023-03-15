unit DropPIDLTarget;

// -----------------------------------------------------------------------------
// Project:         Drag and Drop Target Components
// Component Names: TDropPIDLTarget
// Module:          DropPIDLTarget
// Description:     Implements Dragging & Dropping of PIDLs
//                  TO your application from another.
// Version:	       3.5
// Date:            30-MAR-1999
// Target:          Win32, Delphi3, Delphi4, C++ Builder 3, C++ Builder 4
// Authors:         Angus Johnson,   ajohnson@rpi.net.au
//                  Anders Melander, anders@melander.dk
//                                   http://www.melander.dk
//                  Graham Wideman,  graham@sdsu.edu
//                                   http://www.wideman-one.com
// Copyright        ©1997-99 Angus Johnson, Anders Melander & Graham Wideman
// -----------------------------------------------------------------------------


interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  DropSource, DropTarget, ActiveX, ShlObj, ClipBrd, DropPIDLSource;

type
  TDropPIDLTarget = class(TDropTarget)
  private
    PIDLFormatEtc: TFormatEtc;
    fShellMalloc: IMalloc;
    fPIDLs: TStrings; //Used internally to store PIDLs. I use strings to simplify cleanup.
    fFiles: TStrings; //List of filenames (paths)
    function GetPidlCount: integer;
  protected
    procedure ClearData; override;
    function DoGetData: boolean; override;
    function HasValidFormats: boolean; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; Override;

    function GetFolderPidl: pItemIdList;
    function GetRelativeFilePidl(index: integer): pItemIdList;
    function GetAbsoluteFilePidl(index: integer): pItemIdList;
    function PasteFromClipboard: longint; Override;
    property PidlCount: integer read GetPidlCount; //includes folder pidl in count
    //If you just want the filenames (not PIDLs) then use ...
    property Filenames: TStrings read fFiles;
  end;

procedure Register;

implementation

procedure Register;
begin
  RegisterComponents('DragDrop', [TDropPIDLTarget]);
end;

// -----------------------------------------------------------------------------
//			Miscellaneous Functions...
// -----------------------------------------------------------------------------

{-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=} 
function GetPidlsFromHGlobal(const HGlob: HGlobal; var Pidls: TStrings): boolean;
var
  i: integer;
  pInt: ^UINT;
  pCIDA: PIDA;
begin
  result := false;
  pCIDA := PIDA(GlobalLock(HGlob));
  try
    pInt := @(pCIDA^.aoffset[0]);
    for i := 0 to pCIDA^.cidl do
    begin
      Pidls.add(PidlToString(pointer(UINT(pCIDA)+ pInt^)));
      inc(pInt);
    end;
    if Pidls.count > 1 then result := true;
 finally
    GlobalUnlock(HGlob);
  end;
end;

{By implementing the following TStrings class, component processing is reduced.}
// -----------------------------------------------------------------------------
//			TPIDLTargetStrings
// -----------------------------------------------------------------------------

type 
  TPIDLTargetStrings = class(TStrings)
  private
    DropPIDLTarget: TDropPIDLTarget;
  protected
    function Get(Index: Integer): string; override;
    function GetCount: Integer; override;
    procedure Put(Index: Integer; const S: string); override;
    procedure PutObject(Index: Integer; AObject: TObject); override;
  public
    procedure Clear; override;
    procedure Delete(Index: Integer); override;
    procedure Insert(Index: Integer; const S: string); override;
  end;


function TPIDLTargetStrings.Get(Index: Integer): string; 
var
  PidlStr: string;
  buff: array [0..MAX_PATH] of char;
begin
  with DropPIDLTarget do
  begin
    PidlStr := JoinPidlStrings(fPIDLs[0], fPIDLs[Index+1]);
    SHGetPathFromIDList(PItemIDList(pChar(PidlStr)),buff);
    result := buff;
  end;
end;

function TPIDLTargetStrings.GetCount: Integer; 
begin
  with DropPIDLTarget do
    if fPIDLs.count < 2 then
      result := 0 else
      result := fPIDLs.count-1;
end;

procedure TPIDLTargetStrings.Put(Index: Integer; const S: string); 
begin
end;

procedure TPIDLTargetStrings.PutObject(Index: Integer; AObject: TObject); 
begin
end;

procedure TPIDLTargetStrings.Clear;
begin
end;

procedure TPIDLTargetStrings.Delete(Index: Integer); 
begin
end;

procedure TPIDLTargetStrings.Insert(Index: Integer; const S: string); 
begin
end;

// -----------------------------------------------------------------------------
//			TDropPIDLTarget
// -----------------------------------------------------------------------------

{-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=} 
constructor TDropPIDLTarget.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  fPIDLs := TStringList.create;
  fFiles := TPIDLTargetStrings.create;
  TPIDLTargetStrings(fFiles).DropPIDLTarget := self;
  SHGetMalloc(fShellMalloc);
  with PIDLFormatEtc do
  begin
    cfFormat := CF_IDLIST;
    ptd := nil;
    dwAspect := DVASPECT_CONTENT;
    lindex := -1;
    tymed := TYMED_HGLOBAL;
  end;
end;

{-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=} 
destructor TDropPIDLTarget.Destroy;
begin
  fPIDLs.free;
  fFiles.free;
  //fShellMalloc := nil;
  inherited Destroy;
end;

{-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=} 
function TDropPIDLTarget.HasValidFormats: boolean;
begin
  result := (DataObject.QueryGetData(PIDLFormatEtc) = S_OK);
end;

{-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=} 
procedure TDropPIDLTarget.ClearData;
begin
  fPIDLs.clear;
end;

{-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=} 
function TDropPIDLTarget.DoGetData: boolean;
var
  medium: TStgMedium;
begin
  fPIDLs.clear;
  result := false;
  if (DataObject.GetData(PIDLFormatEtc, medium) = S_OK) then
  begin
    try
      if (medium.tymed <> TYMED_HGLOBAL) then exit;
      result := GetPidlsFromHGlobal(medium.HGlobal,fPIDLs);
    finally
      ReleaseStgMedium(medium);
    end;
  end;
end;

//Note: It is the component user's responsibility to cleanup 
//the returned PIDLs from the following 3 methods. 
//Use - CoTaskMemFree() - to free the PIDLs.

{-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=} 
function TDropPIDLTarget.GetFolderPidl: pItemIdList;
begin
  result :=nil;
  if fPIDLs.count = 0 then exit;
  result := fShellMalloc.alloc(length(fPIDLs[0]));
  if result <> nil then
    move(pChar(fPIDLs[0])^,result^,length(fPIDLs[0]));
end;

{-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=} 
function TDropPIDLTarget.GetRelativeFilePidl(index: integer): pItemIdList;
begin
  result :=nil;
  if (index < 1) or (index >= fPIDLs.count) then exit;
  result := fShellMalloc.alloc(length(fPIDLs[index]));
  if result <> nil then
    move(pChar(fPIDLs[index])^,result^,length(fPIDLs[index]));
end;

{-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=} 
function TDropPIDLTarget.GetAbsoluteFilePidl(index: integer): pItemIdList;
var
  s: string;
begin
  result :=nil;
  if (index < 1) or (index >= fPIDLs.count) then exit;
  s := JoinPidlStrings(fPIDLs[0], fPIDLs[index]);
  result := fShellMalloc.alloc(length(s));
  if result <> nil then
    move(pChar(s)^,result^,length(s));
end;

{-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=} 
function TDropPIDLTarget.PasteFromClipboard: longint;
var
  Global: HGlobal;
  Preferred: longint;
begin
  result  := DROPEFFECT_NONE;
  if not ClipBoard.HasFormat(CF_IDLIST) then exit;
  Global := Clipboard.GetAsHandle(CF_IDLIST);
  fPIDLs.clear;
  if not GetPidlsFromHGlobal(Global,fPidls) then exit;
  Preferred := inherited PasteFromClipboard;
  //if no Preferred DropEffect then return copy else return Preferred ...
  if (Preferred = DROPEFFECT_NONE) then
    result := DROPEFFECT_COPY else
    result := Preferred;
end;

{-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=} 
function TDropPIDLTarget.GetPidlCount: integer;
begin
  result := fPidls.count; //Note: includes folder pidl in count!
end;

// -----------------------------------------------------------------------------
// -----------------------------------------------------------------------------

//initialization
//  CF_IDLIST is 'registered' in DropSource

end.
