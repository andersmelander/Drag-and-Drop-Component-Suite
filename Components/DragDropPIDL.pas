unit DragDropPIDL;
// -----------------------------------------------------------------------------
// Project:         Drag and Drop Component Suite
// Module:          DragDropPIDL
// Description:     Implements Dragging & Dropping of PIDLs (files and folders).
// Version:         4.1
// Date:            22-JAN-2002
// Target:          Win32, Delphi 4-6, C++Builder 4-6
// Authors:         Anders Melander, anders@melander.dk, http://www.melander.dk
// Copyright        © 1997-2002 Angus Johnson & Anders Melander
// -----------------------------------------------------------------------------

interface

uses
  DragDrop,
  DropTarget,
  DropSource,
  DragDropFormats,
  DragDropFile,
  Windows,
  ActiveX,
  Classes,
  ShlObj;

{$include DragDrop.inc}

type
////////////////////////////////////////////////////////////////////////////////
//
//		TPIDLClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
// Supports the 'Shell IDList Array' format.
////////////////////////////////////////////////////////////////////////////////
  TPIDLClipboardFormat = class(TCustomSimpleClipboardFormat)
  private
    FPIDLs: TStrings; // Used internally to store PIDLs. We use strings to simplify cleanup.
    FFilenames: TStrings;
  protected
    function ReadData(Value: pointer; Size: integer): boolean; override;
    function WriteData(Value: pointer; Size: integer): boolean; override;
    function GetSize: integer; override;
  public
    constructor Create; override;
    destructor Destroy; override;
    function GetClipboardFormat: TClipFormat; override;
    procedure Clear; override;
    function HasData: boolean; override;
    property PIDLs: TStrings read FPIDLs;
    property Filenames: TStrings read FFilenames;
  end;


type
////////////////////////////////////////////////////////////////////////////////
//
//		TPIDLDataFormat
//
////////////////////////////////////////////////////////////////////////////////
  TPIDLDataFormat = class(TCustomDataFormat)
  private
    FPIDLs: TStrings;
    FFilenames: TStrings;
  protected
  public
    constructor Create(AOwner: TDragDropComponent); override;
    destructor Destroy; override;
    function Assign(Source: TClipboardFormat): boolean; override;
    function AssignTo(Dest: TClipboardFormat): boolean; override;
    procedure Clear; override;
    function HasData: boolean; override;
    function NeedsData: boolean; override;
    property PIDLs: TStrings read FPIDLs;
    property Filenames: TStrings read FFilenames;
  end;


type
////////////////////////////////////////////////////////////////////////////////
//
//		TDropPIDLTarget
//
////////////////////////////////////////////////////////////////////////////////
  TDropPIDLTarget = class(TCustomDropMultiTarget)
  private
    FPIDLDataFormat: TPIDLDataFormat;
    FFileMapDataFormat: TFileMapDataFormat;
    function GetFilenames: TStrings;
  protected
    function GetPIDLs: TStrings;
    function GetPIDLCount: integer;
    function GetMappedNames: TStrings;
    property PIDLs: TStrings read GetPIDLs;
    function DoGetPIDL(Index: integer): pItemIdList;
    function GetPreferredDropEffect: LongInt; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; Override;

    // Note: It is the callers responsibility to cleanup
    // the returned PIDLs from the following 3 methods:
    // - GetFolderPidl
    // - GetRelativeFilePidl
    // - GetAbsoluteFilePidl
    // Use the CoTaskMemFree procedure to free the PIDLs.
    function GetFolderPIDL: pItemIdList;
    function GetRelativeFilePIDL(Index: integer): pItemIdList;
    function GetAbsoluteFilePIDL(Index: integer): pItemIdList;
    property PIDLCount: integer read GetPIDLCount; // Includes folder pidl in count

    // If you just want the filenames (not PIDLs) then use ...
    property Filenames: TStrings read GetFilenames;
    // MappedNames is only needed if files need to be renamed after a drag or
    // e.g. dragging from 'Recycle Bin'.
    property MappedNames: TStrings read GetMappedNames;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TDropPIDLSource
//
////////////////////////////////////////////////////////////////////////////////
  TDropPIDLSource = class(TCustomDropMultiSource)
  private
    FPIDLDataFormat: TPIDLDataFormat;
    FFileMapDataFormat: TFileMapDataFormat;
  protected
    function GetMappedNames: TStrings;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure CopyFolderPIDLToList(pidl: PItemIDList);
    procedure CopyFilePIDLToList(pidl: PItemIDList);
    property MappedNames: TStrings read GetMappedNames;
  end;


////////////////////////////////////////////////////////////////////////////////
//
//		PIDL utility functions
//
////////////////////////////////////////////////////////////////////////////////

//: GetPIDLsFromData extracts a PIDL list from a memory block and stores the
// PIDLs in a string list.
function GetPIDLsFromData(Data: pointer; Size: integer; PIDLs: TStrings): boolean;

//: GetPIDLsFromHGlobal extracts a PIDL list from a global memory block and
// stores the PIDLs in a string list.
function GetPIDLsFromHGlobal(const HGlob: HGlobal; PIDLs: TStrings): boolean;

//: GetPIDLsFromFilenames converts a list of files to PIDLs and stores the
// PIDLs in a string list. All the PIDLs are relative to a common root.
function GetPIDLsFromFilenames(const Files: TStrings; PIDLs: TStrings): boolean;

//: GetRootFolderPIDL finds the PIDL of the folder which is the parent of a list
// of files. The PIDl is returned as a string. If the files do not share a
// common root, an empty string is returnde.
function GetRootFolderPIDL(const Files: TStrings): string;

//: GetFullPIDLFromPath converts a path (filename and path) to a folder/filename
// PIDL pair.
function GetFullPIDLFromPath(Path: string): pItemIDList;

//: GetFullPathFromPIDL converts a folder/filename PIDL pair to a full path.
function GetFullPathFromPIDL(PIDL: pItemIDList): string;

//: PIDLToString converts a single PIDL to a string.
function PIDLToString(pidl: PItemIDList): string;

//: StringToPIDL converts a PIDL string to a PIDL.
function StringToPIDL(const PIDL: string): PItemIDList;

//: JoinPIDLStrings merges two PIDL strings into one.
function JoinPIDLStrings(pidl1, pidl2: string): string;

//: ConvertFilesToShellIDList converts a list of files to a PIDL list. The
// files are relative to the folder specified by the Path parameter. The PIDLs
// are returned as a global memory handle.
function ConvertFilesToShellIDList(Path: string; Files: TStrings): HGlobal;

//: GetSizeOfPIDL calculates the size of a PIDL list.
function GetSizeOfPIDL(PIDL: pItemIDList): integer;

//: CopyPIDL makes a copy of a PIDL.
// It is the callers responsibility to free the returned PIDL.               
function CopyPIDL(PIDL: pItemIDList): pItemIDList;

{$ifndef BCB}
// Undocumented PIDL utility functions...
// From http://www.geocities.com/SiliconValley/4942/
function ILCombine(pidl1,pidl2:PItemIDList): PItemIDList; stdcall;
function ILFindLastID(pidl: PItemIDList): PItemIDList; stdcall;
function ILClone(pidl: PItemIDList): PItemIDList; stdcall;
function ILRemoveLastID(pidl: PItemIDList): LongBool; stdcall;
function ILIsEqual(pidl1,pidl2: PItemIDList): LongBool; stdcall;
procedure ILFree(Buffer: PItemIDList); stdcall;

// Undocumented IMalloc utility functions...
function SHAlloc(BufferSize: ULONG): Pointer; stdcall;
procedure SHFree(Buffer: Pointer); stdcall;
{$endif}

////////////////////////////////////////////////////////////////////////////////
//
//		PIDL/IShellFolder utility functions
//
////////////////////////////////////////////////////////////////////////////////

//: GetShellFolderOfPath retrieves an IShellFolder interface which can be used
// to manage the specified folder.
function GetShellFolderOfPath(FolderPath: string): IShellFolder;

//: GetPIDLDisplayName retrieves the display name of the specified PIDL,
// relative to the specified folder.
function GetPIDLDisplayName(Folder: IShellFolder; PIDL: PItemIdList): string;

//: GetSubPIDL retrieves the PIDL of the specified file or folder to a PIDL.
// The PIDL is relative to the folder specified by the Folder parameter.
function GetSubPIDL(Folder: IShellFolder; Sub: string): pItemIDList;


////////////////////////////////////////////////////////////////////////////////
//
//		Component registration
//
////////////////////////////////////////////////////////////////////////////////
procedure Register;

implementation

uses
  ShellAPI,
  SysUtils;

resourcestring
  sNoFolderPIDL = 'Folder PIDL must be added first';

////////////////////////////////////////////////////////////////////////////////
//
//		Component registration
//
////////////////////////////////////////////////////////////////////////////////
procedure Register;
begin
  RegisterComponents(DragDropComponentPalettePage, [TDropPIDLTarget,
    TDropPIDLSource]);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		PIDL utility functions
//
////////////////////////////////////////////////////////////////////////////////
function GetPIDLsFromData(Data: pointer; Size: integer; PIDLs: TStrings): boolean;
var
  i: integer;
  pOffset: ^UINT;
  PIDL: PItemIDList;
begin
  PIDLs.Clear;

  Result := (Data <> nil) and
    (Size >= integer(PIDA(Data)^.cidl) * (SizeOf(UINT)+SizeOf(PItemIDList)) + SizeOf(UINT));
  if (not Result) then
    exit;

  pOffset := @(PIDA(Data)^.aoffset[0]);
  i := PIDA(Data)^.cidl; // Note: Count doesn't include folder PIDL
  while (i >= 0) do
  begin
    PIDL := PItemIDList(UINT(Data)+ pOffset^);
    PIDLs.Add(PIDLToString(PIDL));
    inc(pOffset);
    dec(i);
  end;
  Result := (PIDLs.Count > 1);
end;

function GetPIDLsFromHGlobal(const HGlob: HGlobal; PIDLs: TStrings): boolean;
var
  pCIDA: PIDA;
begin
  pCIDA := PIDA(GlobalLock(HGlob));
  try
    Result := GetPIDLsFromData(pCIDA, GlobalSize(HGlob), PIDLs);
  finally
    GlobalUnlock(HGlob);
  end;
end;

resourcestring
  sBadDesktop = 'Failed to get interface to Desktop';
  sBadFilename = 'Invalid filename: %s';

(*
** Find the folder which is the parent of all the files in a list.
*)
function GetRootFolderPIDL(const Files: TStrings): string;
var
  DeskTopFolder: IShellFolder;
  WidePath: WideString;
  PIDL: pItemIDList;
  PIDLs: TStrings;
  s: string;
  PIDL1, PIDL2: pItemIDList;
  Size, MaxSize: integer;
  i: integer;
begin
  Result := '';
  if (Files.Count = 0) then
    exit;

  if (SHGetDesktopFolder(DeskTopFolder) <> NOERROR) then
    raise Exception.Create(sBadDesktop);

  PIDLs := TStringList.Create;
  try
    // First convert all paths to PIDLs.
    for i := 0 to Files.Count-1 do
    begin
      WidePath := ExtractFilePath(Files[i]);
      if (DesktopFolder.ParseDisplayName(0, nil, PWideChar(WidePath), PULONG(nil)^,
        PIDL, PULONG(nil)^) <> NOERROR) then
        raise Exception.Create(sBadFilename);
      try
        PIDLs.Add(PIDLToString(PIDL));
      finally
        coTaskMemFree(PIDL);
      end;
    end;

    Result := PIDLs[0];
    MaxSize := Length(Result)-SizeOf(Word);
    PIDL := pItemIDList(PChar(Result));
    for i := 1 to PIDLs.Count-1 do
    begin
      s := PIDLs[1];
      PIDL1 := PIDL;
      PIDL2 := pItemIDList(PChar(s));
      Size := 0;
      while (Size < MaxSize) and (PIDL1^.mkid.cb <> 0) and (PIDL1^.mkid.cb = PIDL2^.mkid.cb) and (CompareMem(PIDL1, PIDL2, PIDL1^.mkid.cb)) do
      begin
        inc(Size, PIDL1^.mkid.cb);
        inc(integer(PIDL2), PIDL1^.mkid.cb);
        inc(integer(PIDL1), PIDL1^.mkid.cb);
      end;
      if (Size <> MaxSize) then
      begin
        MaxSize := Size;
        SetLength(Result, Size+SizeOf(Word));
        PIDL1^.mkid.cb := 0;
      end;
      if (Size = 0) then
        break;
    end;
  finally
    PIDLs.Free;
  end;
end;

function GetPIDLsFromFilenames(const Files: TStrings; PIDLs: TStrings): boolean;
var
  RootPIDL: string;
  i: integer;
  PIDL: pItemIdList;
  FilePIDL: string;
begin
  Result := False;
  PIDLs.Clear;
  if (Files.Count = 0) then
    exit;

  // Get the PIDL of the root folder...
  // All the file PIDLs will be relative to this PIDL
  RootPIDL := GetRootFolderPIDL(Files);
  if (RootPIDL = '') then
    exit;

  Result := True;

  PIDLS.Add(RootPIDL);
  // Add the file PIDLs (all relative to the root)...
  for i := 0 to Files.Count-1 do
  begin
    PIDL := GetFullPIDLFromPath(Files[i]);
    if (PIDL = nil) then
    begin
      Result := False;
      PIDLs.Clear;
      break;
    end;
    try
      FilePIDL := PIDLToString(PIDL);
    finally
      coTaskMemFree(PIDL);
    end;
    // Remove the root PIDL from the file PIDL making it relative to the root.
    PIDLS.Add(copy(FilePIDL, Length(RootPIDL)-SizeOf(Word)+1,
      Length(FilePIDL)-(Length(RootPIDL)-SizeOf(Word))));
  end;
end;

function GetSizeOfPIDL(PIDL: pItemIDList): integer;
var
  Size: integer;
begin
  if (PIDL <> nil) then
  begin
    Result := SizeOf(PIDL^.mkid.cb);
    repeat
      Size := PIDL^.mkid.cb;
      inc(Result, Size);
      inc(integer(PIDL), Size);
    until (Size = 0);
  end else
    Result := 0;
end;

function CopyPIDL(PIDL: pItemIDList): pItemIDList;
var
  Size: integer;
begin
  Size := GetSizeOfPIDL(PIDL);
  if (Size > 0) then
  begin
    Result := ShellMalloc.Alloc(Size);
    if (Result <> nil) then
      Move(PIDL^, Result^, Size);
  end else
    Result := nil;
end;

function GetFullPIDLFromPath(Path: string): pItemIDList;
var
  DeskTopFolder: IShellFolder;
  WidePath: WideString;
begin
  WidePath := Path;
  if (SHGetDesktopFolder(DeskTopFolder) = NOERROR) then
  begin
    if (DesktopFolder.ParseDisplayName(0, nil, PWideChar(WidePath), PULONG(nil)^,
      Result, PULONG(nil)^) <> NOERROR) then
      Result := nil;
  end else
    Result := nil;
end;

function GetFullPathFromPIDL(PIDL: pItemIDList): string;
var
  Path: array[0..MAX_PATH] of char;
begin
  if SHGetPathFromIDList(PIDL, Path) then
    Result := Path
  else
    Result := '';
end;

// See "Clipboard Formats for Shell Data Transfers" in Ole.hlp...
// (Needed to drag links (shortcuts).)
type
  POffsets = ^TOffsets;
  TOffsets = array[0..$FFFF] of UINT;

function ConvertFilesToShellIDList(Path: string; Files: TStrings): HGlobal;
var
  shf: IShellFolder;
  PathPidl, pidl: pItemIDList;
  Ida: PIDA;
  pOffset: POffsets;
  ptrByte: ^Byte;
  i, PathPidlSize, IdaSize, PreviousPidlSize: integer;
begin
  Result := 0;
  shf := GetShellFolderOfPath(path);
  if shf = nil then
    exit;
  // Calculate size of IDA structure ...
  // cidl: UINT ; Directory pidl
  // offset: UINT ; all file pidl offsets
  IdaSize := (Files.Count + 2) * SizeOf(UINT);

  PathPidl := GetFullPIDLFromPath(path);
  if PathPidl = nil then
    exit;
  try
    PathPidlSize := GetSizeOfPidl(PathPidl);

    //Add to IdaSize space for ALL pidls...
    IdaSize := IdaSize + PathPidlSize;
    for i := 0 to Files.Count-1 do
    begin
      pidl := GetSubPidl(shf, files[i]);
      try
        IdaSize := IdaSize + GetSizeOfPidl(Pidl);
      finally
        ShellMalloc.Free(pidl);
      end;
    end;

    //Allocate memory...
    Result := GlobalAlloc(GMEM_SHARE or GMEM_ZEROINIT, IdaSize);
    if (Result = 0) then
      exit;
    try
      Ida := GlobalLock(Result);
      try
        FillChar(Ida^, IdaSize, 0);

        //Fill in offset and pidl data...
        Ida^.cidl := Files.Count; //cidl = file count
        pOffset := POffsets(@(Ida^.aoffset));
        pOffset^[0] := (Files.Count+2) * sizeof(UINT); //offset of Path pidl

        ptrByte := pointer(Ida);
        inc(ptrByte, pOffset^[0]); //ptrByte now points to Path pidl
        Move(PathPidl^, ptrByte^, PathPidlSize); //copy path pidl

        PreviousPidlSize := PathPidlSize;
        for i := 1 to Files.Count do
        begin
          pidl := GetSubPidl(shf,files[i-1]);
          try
            pOffset^[i] := pOffset^[i-1] + UINT(PreviousPidlSize); //offset of pidl
            PreviousPidlSize := GetSizeOfPidl(Pidl);

            ptrByte := pointer(Ida);
            inc(ptrByte, pOffset^[i]); //ptrByte now points to current file pidl
            Move(Pidl^, ptrByte^, PreviousPidlSize); //copy file pidl
                                  //PreviousPidlSize = current pidl size here
          finally
            ShellMalloc.Free(pidl);
          end;
        end;
      finally
        GlobalUnLock(Result);
      end;
    except
      GlobalFree(Result);
      raise;
    end;
  finally
    ShellMalloc.Free(PathPidl);
  end;
end;

function PIDLToString(pidl: PItemIDList): String;
var
  PidlLength: integer;
begin
  PidlLength := GetSizeOfPidl(pidl);
  SetLength(Result, PidlLength);
  Move(pidl^, PChar(Result)^, PidlLength);
end;

function StringToPIDL(const PIDL: string): PItemIDList;
begin
  Result := ShellMalloc.Alloc(Length(PIDL));
  if (Result <> nil) then
    Move(PChar(PIDL)^, Result^, Length(PIDL));
end;

function JoinPIDLStrings(pidl1, pidl2: string): String;
var
  PidlLength: integer;
begin
  if Length(pidl1) <= 2 then
    PidlLength := 0
  else
    PidlLength := Length(pidl1)-2;
  SetLength(Result, PidlLength + Length(pidl2));
  if PidlLength > 0 then
    Move(PChar(pidl1)^, PChar(Result)^, PidlLength);
  Move(PChar(pidl2)^, Result[PidlLength+1], Length(pidl2));
end;

{$ifndef BCB}
// BCB appearantly doesn't support ordinal DLL imports in Delphi units. Strange!
// Use LoadLibrary and GetProcAddress if you need access to these functions from
// C++Builder.
function ILCombine(pidl1,pidl2:PItemIDList): PItemIDList; stdcall;
  external shell32 index 25;
function ILFindLastID(pidl: PItemIDList): PItemIDList; stdcall;
  external shell32 index 16;
function ILClone(pidl: PItemIDList): PItemIDList; stdcall;
  external shell32 index 18;
function ILRemoveLastID(pidl: PItemIDList): LongBool; stdcall;
  external shell32 index 17;
function ILIsEqual(pidl1,pidl2: PItemIDList): LongBool; stdcall;
  external shell32 index 21;
procedure ILFree(Buffer: PItemIDList); stdcall;
  external shell32 index 155;

function SHAlloc(BufferSize: ULONG): Pointer; stdcall;
  external shell32 index 196;
procedure SHFree(Buffer: Pointer); stdcall;
  external shell32 index 195;
{$endif}

////////////////////////////////////////////////////////////////////////////////
//
//		PIDL/IShellFolder utility functions
//
////////////////////////////////////////////////////////////////////////////////
function GetShellFolderOfPath(FolderPath: string): IShellFolder;
var
  DeskTopFolder: IShellFolder;
  PathPidl: pItemIDList;
  WidePath: WideString;
  pdwAttributes: ULONG;
begin
  Result := nil;
  WidePath := FolderPath;
  pdwAttributes := SFGAO_FOLDER;
  if (SHGetDesktopFolder(DeskTopFolder) <> NOERROR) then
    exit;
  if (DesktopFolder.ParseDisplayName(0, nil, PWideChar(WidePath), PULONG(nil)^,
    PathPidl, pdwAttributes) = NOERROR) then
    try
      if (pdwAttributes and SFGAO_FOLDER <> 0) then
        DesktopFolder.BindToObject(PathPidl, nil, IID_IShellFolder,
          // Note: For Delphi 4 and prior, the ppvOut parameter must be a pointer.
          pointer(Result));
    finally
      ShellMalloc.Free(PathPidl);
    end;
end;

function GetSubPIDL(Folder: IShellFolder; Sub: string): pItemIDList;
var
  WidePath: WideString;
begin
  WidePath := Sub;
  Folder.ParseDisplayName(0, nil, PWideChar(WidePath), PULONG(nil)^, Result,
    PULONG(nil)^);
end;

function GetPIDLDisplayName(Folder: IShellFolder; PIDL: PItemIdList): string;
var
  StrRet: TStrRet;
begin
  Result := '';
  Folder.GetDisplayNameOf(PIDL, 0, StrRet);
  case StrRet.uType of
    STRRET_WSTR: Result := WideCharToString(StrRet.pOleStr);
    STRRET_OFFSET: Result := PChar(UINT(PIDL)+StrRet.uOffset);
    STRRET_CSTR: Result := StrRet.cStr;
  end;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TPIDLsToFilenamesStrings
//
////////////////////////////////////////////////////////////////////////////////
// Used internally to convert PIDLs to filenames on-demand.
////////////////////////////////////////////////////////////////////////////////
type
  TPIDLsToFilenamesStrings = class(TStrings)
  private
    FPIDLs: TStrings;
  protected
    function Get(Index: Integer): string; override;
    function GetCount: Integer; override;
    procedure Put(Index: Integer; const S: string); override;
    procedure PutObject(Index: Integer; AObject: TObject); override;
  public
    constructor Create(APIDLs: TStrings);
    procedure Clear; override;
    procedure Delete(Index: Integer); override;
    procedure Insert(Index: Integer; const S: string); override;
    procedure Assign(Source: TPersistent); override;
  end;

constructor TPIDLsToFilenamesStrings.Create(APIDLs: TStrings);
begin
  inherited Create;
  FPIDLs := APIDLs;
end;

function TPIDLsToFilenamesStrings.Get(Index: Integer): string;
var
  PIDL: string;
  Path: array [0..MAX_PATH] of char;
begin
  if (Index < 0) or (Index > FPIDLs.Count-2) then
    raise Exception.create('Filename index out of range');
  PIDL := JoinPIDLStrings(FPIDLs[0], FPIDLs[Index+1]);
  if SHGetPathFromIDList(PItemIDList(PChar(PIDL)), Path) then
    Result := Path
  else
    Result := '';
end;

function TPIDLsToFilenamesStrings.GetCount: Integer;
begin
  if FPIDLs.Count < 2 then
    Result := 0
  else
    Result := FPIDLs.Count-1;
end;

procedure TPIDLsToFilenamesStrings.Assign(Source: TPersistent);
begin
  if Source is TStrings then
  begin
    BeginUpdate;
    try
      GetPIDLsFromFilenames(TStrings(Source), FPIDLs);
    finally
      EndUpdate;
    end;
  end else
    inherited Assign(Source);
end;

// Inherited abstract methods which do not need implementation...
procedure TPIDLsToFilenamesStrings.Put(Index: Integer; const S: string);
begin
end;

procedure TPIDLsToFilenamesStrings.PutObject(Index: Integer; AObject: TObject);
begin
end;

procedure TPIDLsToFilenamesStrings.Clear;
begin
end;

procedure TPIDLsToFilenamesStrings.Delete(Index: Integer);
begin
end;

procedure TPIDLsToFilenamesStrings.Insert(Index: Integer; const S: string);
begin
end;

////////////////////////////////////////////////////////////////////////////////
//
//		TPIDLClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
constructor TPIDLClipboardFormat.Create;
begin
  inherited Create;
  FPIDLs := TStringList.Create;
  FFilenames := TPIDLsToFilenamesStrings.Create(FPIDLs);
end;

destructor TPIDLClipboardFormat.Destroy;
begin
  FFilenames.Free;
  FPIDLs.Free;
  inherited Destroy;
end;

var
  CF_IDLIST: TClipFormat = 0;

function TPIDLClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  if (CF_IDLIST = 0) then
    CF_IDLIST := RegisterClipboardFormat(CFSTR_SHELLIDLIST);
  Result := CF_IDLIST;
end;

procedure TPIDLClipboardFormat.Clear;
begin
  FPIDLs.Clear;
end;

function TPIDLClipboardFormat.HasData: boolean;
begin
  Result := (FPIDLs.Count > 0);
end;

function TPIDLClipboardFormat.GetSize: integer;
var
  i: integer;
begin
  Result := (FPIDLs.Count+1) * SizeOf(UINT);
  for i := 0 to FPIDLs.Count-1 do
    inc(Result, Length(FPIDLs[i]));
end;

function TPIDLClipboardFormat.ReadData(Value: pointer;
  Size: integer): boolean;
begin
  Result := GetPIDLsFromData(Value, Size, FPIDLs);
end;

function TPIDLClipboardFormat.WriteData(Value: pointer;
  Size: integer): boolean;
var
  i: integer;
  pCIDA: PIDA;
  Offset: integer;
  pOffset: ^UINT;
  PIDL: PItemIDList;
begin
  pCIDA := PIDA(Value);
  pCIDA^.cidl := FPIDLs.Count-1; // Don't count folder PIDL
  pOffset := @(pCIDA^.aoffset[0]); // Points to aoffset[0]
  Offset := (FPIDLs.Count+1)*SizeOf(UINT); // Size of CIDA structure
  PIDL := PItemIDList(integer(pCIDA) + Offset); // PIDLs are stored after CIDA structure.

  for i := 0 to FPIDLs.Count-1 do
  begin
    pOffset^ := Offset; // Store relative offset of PIDL into aoffset[i]
    // Copy the PIDL
    Move(PChar(FPIDLs[i])^, PIDL^, length(FPIDLs[i]));
    // Move on to next PIDL
    inc(Offset, length(FPIDLs[i]));
    inc(pOffset);
    inc(integer(PIDL), length(FPIDLs[i]));
  end;

  Result := True;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TPIDLDataFormat
//
////////////////////////////////////////////////////////////////////////////////
constructor TPIDLDataFormat.Create(AOwner: TDragDropComponent);
begin
  inherited Create(AOwner);
  FPIDLs := TStringList.Create;
  TStringList(FPIDLs).OnChanging := DoOnChanging;
  FFilenames := TPIDLsToFilenamesStrings.Create(FPIDLs);
end;

destructor TPIDLDataFormat.Destroy;
begin
  FFilenames.Free;
  FPIDLs.Free;
  inherited Destroy;
end;

function TPIDLDataFormat.Assign(Source: TClipboardFormat): boolean;
begin
  Result := True;

  if (Source is TPIDLClipboardFormat) then
    FPIDLs.Assign(TPIDLClipboardFormat(Source).PIDLs)

  else if (Source is TFileClipboardFormat) then
    Result := GetPIDLsFromFilenames(TFileClipboardFormat(Source).Files, FPIDLs)

  else
    Result := inherited Assign(Source);
end;

function TPIDLDataFormat.AssignTo(Dest: TClipboardFormat): boolean;
begin
  Result := True;

  if (Dest is TPIDLClipboardFormat) then
    TPIDLClipboardFormat(Dest).PIDLs.Assign(FPIDLs)

  else if (Dest is TFileClipboardFormat) then
    TFileClipboardFormat(Dest).Files.Assign(Filenames)

  else
    Result := inherited Assign(Dest);
end;

procedure TPIDLDataFormat.Clear;
begin
  FPIDLs.Clear;
end;

function TPIDLDataFormat.HasData: boolean;
begin
  Result := (FPIDLs.Count > 0);
end;

function TPIDLDataFormat.NeedsData: boolean;
begin
  Result := (FPIDLs.Count = 0);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TDropPIDLTarget
//
////////////////////////////////////////////////////////////////////////////////
constructor TDropPIDLTarget.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FPIDLDataFormat := TPIDLDataFormat.Create(Self);
  FFileMapDataFormat := TFileMapDataFormat.Create(Self);
end;

destructor TDropPIDLTarget.Destroy;
begin
  FPIDLDataFormat.Free;
  FFileMapDataFormat.Free;
  inherited Destroy;
end;

function TDropPIDLTarget.GetPIDLs: TStrings;
begin
  Result := FPIDLDataFormat.PIDLs;
end;

function TDropPIDLTarget.DoGetPIDL(Index: integer): pItemIdList;
var
  PIDL: string;
begin
  PIDL := PIDLs[Index];
  Result := ShellMalloc.Alloc(Length(PIDL));
  if (Result <> nil) then
    Move(PChar(PIDL)^, Result^, Length(PIDL));
end;

function TDropPIDLTarget.GetFolderPidl: pItemIdList;
begin
  Result := DoGetPIDL(0);
end;

function TDropPIDLTarget.GetRelativeFilePidl(Index: integer): pItemIdList;
begin
  Result := nil;
  if (index < 1) then
    exit;
  Result := DoGetPIDL(Index);
end;

function TDropPIDLTarget.GetAbsoluteFilePidl(Index: integer): pItemIdList;
var
  PIDL: string;
begin
  Result := nil;
  if (index < 1) then
    exit;
  PIDL := JoinPIDLStrings(PIDLs[0], PIDLs[Index]);
  Result := ShellMalloc.Alloc(Length(PIDL));
  if (Result <> nil) then
    Move(PChar(PIDL)^, Result^, Length(PIDL));
end;

function TDropPIDLTarget.GetPIDLCount: integer;
begin
   // Note: Includes folder PIDL in count!
  Result := FPIDLDataFormat.PIDLs.Count;
end;

function TDropPIDLTarget.GetFilenames: TStrings;
begin
  Result := FPIDLDataFormat.Filenames;
end;

function TDropPIDLTarget.GetMappedNames: TStrings;
begin
  Result := FFileMapDataFormat.FileMaps;
end;

function TDropPIDLTarget.GetPreferredDropEffect: LongInt;
begin
  Result := inherited GetPreferredDropEffect;
  if (Result = DROPEFFECT_NONE) then
    Result := DROPEFFECT_COPY;
end;

////////////////////////////////////////////////////////////////////////////////
//
//		TDropPIDLSource
//
////////////////////////////////////////////////////////////////////////////////
constructor TDropPIDLSource.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FPIDLDataFormat := TPIDLDataFormat.Create(Self);
  FFileMapDataFormat := TFileMapDataFormat.Create(Self);
end;

destructor TDropPIDLSource.Destroy;
begin
  FPIDLDataFormat.Free;
  FFileMapDataFormat.Free;
  inherited Destroy;
end;

procedure TDropPIDLSource.CopyFolderPIDLToList(pidl: PItemIDList);
begin
  //Note: Once the PIDL has been copied into the list it can be 'freed'.
  FPIDLDataFormat.Clear;
  FFileMapDataFormat.Clear;
  FPIDLDataFormat.PIDLs.Add(PIDLToString(pidl));
end;

procedure TDropPIDLSource.CopyFilePIDLToList(pidl: PItemIDList);
begin
  // Note: Once the PIDL has been copied into the list it can be 'freed'.
  // Make sure that folder pidl has been added.
  if (FPIDLDataFormat.PIDLs.Count < 1) then
    raise Exception.Create(sNoFolderPIDL);
  FPIDLDataFormat.PIDLs.Add(PIDLToString(pidl));
end;

function TDropPIDLSource.GetMappedNames: TStrings;
begin
  Result := FFileMapDataFormat.FileMaps;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		Initialization/Finalization
//
////////////////////////////////////////////////////////////////////////////////

initialization
  // Data format registration
  TPIDLDataFormat.RegisterDataFormat;
  // Clipboard format registration
  TPIDLDataFormat.RegisterCompatibleFormat(TPIDLClipboardFormat, 0, csSourceTarget, [ddRead]);
  TPIDLDataFormat.RegisterCompatibleFormat(TFileClipboardFormat, 1, csSourceTarget, [ddRead]);

finalization
  TPIDLDataFormat.UnregisterDataFormat;

end.
