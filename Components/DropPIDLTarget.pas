unit DropPIDLTarget;

// -----------------------------------------------------------------------------
// Project:         Drag and Drop Component Suite
// Component Names: TDropPIDLTarget
// Module:          DropPIDLTarget
// Description:     Implements Dragging & Dropping of PIDLs
//                  TO your application from another.
// Version:         3.7
// Date:            22-JUL-1999
// Target:          Win32, Delphi 3 - Delphi 5, C++ Builder 3, C++ Builder 4
// Authors:         Angus Johnson,   ajohnson@rpi.net.au
//                  Anders Melander, anders@melander.dk
//                                   http://www.melander.dk
// Copyright        © 1997-99 Angus Johnson & Anders Melander
// -----------------------------------------------------------------------------


interface

uses
  DropSource, DropTarget,
  Classes, ActiveX, ShlObj;

{$include DragDrop.inc}

type
  TDropPIDLTarget = class(TDropTarget)
  private
    PIDLFormatEtc,
    fFileNameMapFormatEtc,
    fFileNameMapWFormatEtc: TFormatEtc;
    fPIDLs: TStrings; // Used internally to store PIDLs. I use strings to simplify cleanup.
    fFiles: TStrings; // List of filenames (paths)
    fMappedNames: TStrings;
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
    //MappedNames is only needed if files need to be renamed after a drag op
    //eg dragging from 'Recycle Bin'.
    property MappedNames: TStrings read fMappedNames;
  end;

procedure Register;

implementation

uses
  DropPIDLSource,
  Windows,
  SysUtils,
  ClipBrd;

procedure Register;
begin
  RegisterComponents('DragDrop', [TDropPIDLTarget]);
end;

// -----------------------------------------------------------------------------
//			Miscellaneous Functions...
// -----------------------------------------------------------------------------

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

// -----------------------------------------------------------------------------

function TPIDLTargetStrings.Get(Index: Integer): string;
var
  PidlStr: string;
  buff: array [0..MAX_PATH] of char;
begin
  with DropPIDLTarget do
  begin
    if (Index < 0) or (Index > fPIDLs.count-2) then
      raise Exception.create('Filename index out of range');
    PidlStr := JoinPidlStrings(fPIDLs[0], fPIDLs[Index+1]);
    if SHGetPathFromIDList(PItemIDList(pChar(PidlStr)),buff) then
      result := buff else
      result := '';
  end;
end;
// -----------------------------------------------------------------------------

function TPIDLTargetStrings.GetCount: Integer;
begin
  with DropPIDLTarget do
    if fPIDLs.count < 2 then
      result := 0 else
      result := fPIDLs.count-1;
end;
// -----------------------------------------------------------------------------

//Overriden abstract methods which do not need implementation...

procedure TPIDLTargetStrings.Put(Index: Integer; const S: string);
begin
end;
// -----------------------------------------------------------------------------

procedure TPIDLTargetStrings.PutObject(Index: Integer; AObject: TObject);
begin
end;
// -----------------------------------------------------------------------------

procedure TPIDLTargetStrings.Clear;
begin
end;
// -----------------------------------------------------------------------------

procedure TPIDLTargetStrings.Delete(Index: Integer);
begin
end;
// -----------------------------------------------------------------------------

procedure TPIDLTargetStrings.Insert(Index: Integer; const S: string);
begin
end;

// -----------------------------------------------------------------------------
//			TDropPIDLTarget
// -----------------------------------------------------------------------------

constructor TDropPIDLTarget.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  fPIDLs := TStringList.create;
  fFiles := TPIDLTargetStrings.create;
  TPIDLTargetStrings(fFiles).DropPIDLTarget := self;
  fMappedNames := TStringList.Create;
  //SHGetMalloc(fShellMalloc);
  with PIDLFormatEtc do
  begin
    cfFormat := CF_IDLIST;
    ptd := nil;
    dwAspect := DVASPECT_CONTENT;
    lindex := -1;
    tymed := TYMED_HGLOBAL;
  end;
  with fFileNameMapFormatEtc do
  begin
    cfFormat := CF_FILENAMEMAP;
    ptd := nil;
    dwAspect := DVASPECT_CONTENT;
    lindex := -1;
    tymed := TYMED_HGLOBAL;
  end;
  with fFileNameMapWFormatEtc do
  begin
    cfFormat := CF_FILENAMEMAPW;
    ptd := nil;
    dwAspect := DVASPECT_CONTENT;
    lindex := -1;
    tymed := TYMED_HGLOBAL;
  end;
end;
// -----------------------------------------------------------------------------

destructor TDropPIDLTarget.Destroy;
begin
  fPIDLs.free;
  fFiles.free;
  fMappedNames.free;
  inherited Destroy;
end;
// ----------------------------------------------------------------------------- 

function TDropPIDLTarget.HasValidFormats: boolean;
begin
  result := (DataObject.QueryGetData(PIDLFormatEtc) = S_OK);
end;
// -----------------------------------------------------------------------------

procedure TDropPIDLTarget.ClearData;
begin
  fPIDLs.clear;
  fMappedNames.clear;
end;
// -----------------------------------------------------------------------------

function TDropPIDLTarget.DoGetData: boolean;
var
  medium: TStgMedium;
  pFilename: pChar;
  pFilenameW: PWideChar;
  sFilename: String;
begin
  ClearData;
  result := false;

  //--------------------------------------------------------------------------
  if (DataObject.GetData(PIDLFormatEtc, medium) = S_OK) then
  begin
    try
      if (medium.tymed <> TYMED_HGLOBAL) then exit;
      result := GetPidlsFromHGlobal(medium.HGlobal,fPIDLs);
    finally
      ReleaseStgMedium(medium);
    end;

    if not result then exit;
    //Now check for FileNameMapping as well ...
    //--------------------------------------------------------------------------
    if (DataObject.GetData(fFileNameMapFormatEtc, medium) = S_OK) then
    begin
      try
        if (medium.tymed = TYMED_HGLOBAL) then
        begin
          pFilename := GlobalLock(medium.HGlobal);
          try
            while true do
            begin
              sFilename := pFilename;
              if sFilename = '' then break;
              fMappedNames.add(sFilename);
              inc(pFilename, length(sFilename)+1);
            end;
            if Filenames.count <> fMappedNames.count then
              fMappedNames.clear;
          finally
            GlobalUnlock(medium.HGlobal);
          end;
        end;
      finally
        ReleaseStgMedium(medium);
      end;
    end
    //WideChar support for WinNT...
    else if (DataObject.GetData(fFileNameMapWFormatEtc, medium) = S_OK) then
    try
      if (medium.tymed = TYMED_HGLOBAL) then
      begin
        pFilenameW := GlobalLock(medium.HGlobal);
        try
          while true do
          begin
            sFilename := WideCharToString(pFilenameW);
            if sFilename = '' then break;
            fMappedNames.add(sFilename);
            inc(pFilenameW, length(sFilename)+1);
          end;
          if fFiles.count <> fMappedNames.count then
            fMappedNames.clear;
        finally
          GlobalUnlock(medium.HGlobal);
        end;
      end;
    finally
      ReleaseStgMedium(medium);
    end;

  end;
end;
// -----------------------------------------------------------------------------

//Note: It is the component user's responsibility to cleanup
//the returned PIDLs from the following 3 methods.
//Use - CoTaskMemFree() - to free the PIDLs.
function TDropPIDLTarget.GetFolderPidl: pItemIdList;
begin
  result :=nil;
  if fPIDLs.count = 0 then exit;
  result := ShellMalloc.alloc(length(fPIDLs[0]));
  if result <> nil then
    move(pChar(fPIDLs[0])^,result^,length(fPIDLs[0]));
end;
// -----------------------------------------------------------------------------

function TDropPIDLTarget.GetRelativeFilePidl(index: integer): pItemIdList;
begin
  result :=nil;
  if (index < 1) or (index >= fPIDLs.count) then exit;
  result := ShellMalloc.alloc(length(fPIDLs[index]));
  if result <> nil then
    move(pChar(fPIDLs[index])^,result^,length(fPIDLs[index]));
end;
// ----------------------------------------------------------------------------- 

function TDropPIDLTarget.GetAbsoluteFilePidl(index: integer): pItemIdList;
var
  s: string;
begin
  result :=nil;
  if (index < 1) or (index >= fPIDLs.count) then exit;
  s := JoinPidlStrings(fPIDLs[0], fPIDLs[index]);
  result := ShellMalloc.alloc(length(s));
  if result <> nil then
    move(pChar(s)^,result^,length(s));
end;
// ----------------------------------------------------------------------------- 

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
// ----------------------------------------------------------------------------- 

function TDropPIDLTarget.GetPidlCount: integer;
begin
  result := fPidls.count; //Note: includes folder pidl in count!
end;
// -----------------------------------------------------------------------------
// -----------------------------------------------------------------------------

end.
