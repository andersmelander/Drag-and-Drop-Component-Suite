unit DragDropFile;
// -----------------------------------------------------------------------------
// Project:         Drag and Drop Component Suite.
// Module:          DragDropFile
// Description:     Implements Dragging and Dropping of files and folders.
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
  DropTarget,
  DropSource,
  DragDropFormats,
  ActiveX,
  Windows,
  Classes;

{$include DragDrop.inc}

{$ifdef DD_WIDESTRINGLIST}
type
  TWString = record
    WString: WideString;
  end;

  TWideStringsHelper = class helper for TStrings
  private
    function GetWideText: WideString;
  protected
    function GetWide(Index: Integer): WideString;
    procedure PutWide(Index: Integer; const S: WideString);
    function IsWide: boolean;
  published
  public
    function Add(const S: WideString): Integer; overload;
    function AddWide(const S: WideString): Integer;
    function AddObject(const S: WideString; AObject: TObject): Integer; overload;
    function AddObjectWide(const S: WideString; AObject: TObject): Integer;
    function IndexOf(const S: WideString): Integer; overload;
    function IndexOfWide(const S: WideString): Integer;
    procedure Insert(Index: Integer; const S: WideString); overload;
    procedure InsertWide(Index: Integer; const S: WideString);
    property WideStrings[Index: Integer]: WideString read GetWide write PutWide;
    property WideText: WideString read GetWideText;
    property Wide: boolean read IsWide;
  end;

  TWideStrings = class(TStrings)
  private
  protected
    function GetWide(Index: Integer): WideString; virtual; abstract;
    procedure PutWide(Index: Integer; const S: WideString); virtual; abstract;
  public
    function AddObject(const S: WideString; AObject: TObject): Integer; reintroduce; virtual;
    function Add(const S: WideString): Integer; reintroduce; virtual;
    function IndexOf(const S: WideString): Integer; reintroduce; virtual;
  end;

  TWideStringList = class(TWideStrings)
  private
    FWideStringList: TList;
  protected
    function GetWide(Index: Integer): WideString; override;
    procedure PutWide(Index: Integer; const S: WideString); override;
    function Get(Index: Integer): string; override;
    function GetCount: Integer; override;
  public
    constructor Create;
    destructor Destroy; override;
    procedure Clear; override;
    procedure AddStrings(Strings: TStrings); override;
    function Add(const S: WideString): Integer; override;
    function AddObject(const S: WideString; AObject: TObject): Integer; override;
    procedure Delete(Index: Integer); override;
    function IndexOf(const S: WideString): Integer; override;
    procedure Insert(Index: Integer; const S: string); overload; override;
    procedure Insert(Index: Integer; const S: WideString); overload;
  end;
{$else}
type
  TWideStrings = class(TStrings);
  TWideStringList = class(TStringList);
{$endif}


type
////////////////////////////////////////////////////////////////////////////////
//
//		TFileClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
  TFileClipboardFormat = class(TCustomSimpleClipboardFormat)
  private
    FFiles: TStrings;
    FWide: boolean;
  protected
    function ReadData(Value: pointer; Size: integer): boolean; override;
    function WriteData(Value: pointer; Size: integer): boolean; override;
    function GetSize: integer; override;
    property Wide: boolean read FWide;
  public
    constructor Create; override;
    destructor Destroy; override;
    function GetClipboardFormat: TClipFormat; override;
    function Assign(Source: TCustomDataFormat): boolean; override;
    function AssignTo(Dest: TCustomDataFormat): boolean; override;
    procedure Clear; override;
    function HasData: boolean; override;
    property Files: TStrings read FFiles;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TFilenameClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
  TFilenameClipboardFormat = class(TCustomTextClipboardFormat)
  public
    function GetClipboardFormat: TClipFormat; override;
    function Assign(Source: TCustomDataFormat): boolean; override;
    function AssignTo(Dest: TCustomDataFormat): boolean; override;
    property Filename: string read GetString write SetString;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TFilenameWClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
  TFilenameWClipboardFormat = class(TCustomWideTextClipboardFormat)
  public
    function GetClipboardFormat: TClipFormat; override;
    function Assign(Source: TCustomDataFormat): boolean; override;
    function AssignTo(Dest: TCustomDataFormat): boolean; override;
    property Filename: WideString read GetText write SetText;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TFilenameMapClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
  TFilenameMapClipboardFormat = class(TCustomSimpleClipboardFormat)
  private
    FFileMaps: TStrings;
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
    property FileMaps: TStrings read FFileMaps;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TFilenameMapWClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
  TFilenameMapWClipboardFormat = class(TCustomSimpleClipboardFormat)
  private
    FFileMaps: TStrings;
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
    property FileMaps: TStrings read FFileMaps;
  end;


////////////////////////////////////////////////////////////////////////////////
//
//		TFileMapDataFormat
//
////////////////////////////////////////////////////////////////////////////////
  TFileMapDataFormat = class(TCustomDataFormat)
  private
    FFileMaps: TStrings;
  public
    constructor Create(AOwner: TDragDropComponent); override;
    destructor Destroy; override;
    function Assign(Source: TClipboardFormat): boolean; override;
    function AssignTo(Dest: TClipboardFormat): boolean; override;
    procedure Clear; override;
    function HasData: boolean; override;
    function NeedsData: boolean; override;
    property FileMaps: TStrings read FFileMaps;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TFileDataFormat
//
////////////////////////////////////////////////////////////////////////////////
  TFileDataFormat = class(TCustomDataFormat)
  private
    FFiles: TStrings;
  protected
  public
    constructor Create(AOwner: TDragDropComponent); override;
    destructor Destroy; override;
    function Assign(Source: TClipboardFormat): boolean; override;
    function AssignTo(Dest: TClipboardFormat): boolean; override;
    procedure Clear; override;
    function HasData: boolean; override;
    function NeedsData: boolean; override;
    property Files: TStrings read FFiles;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TDropFileTarget
//
////////////////////////////////////////////////////////////////////////////////
  TDropFileTarget = class(TCustomDropMultiTarget)
  private
    FFileFormat: TFileDataFormat;
    FFileMapFormat: TFileMapDataFormat;
  protected
    function GetFiles: TStrings;
    function GetMappedNames: TStrings;
    function GetPreferredDropEffect: LongInt; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    property Files: TStrings read GetFiles;
    property MappedNames: TStrings read GetMappedNames;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TDropFileSource
//
////////////////////////////////////////////////////////////////////////////////
  TDropFileSource = class(TCustomDropMultiSource)
  private
    FFileFormat: TFileDataFormat;
    FFileMapFormat: TFileMapDataFormat;
    function GetFiles: TStrings;
    function GetMappedNames: TStrings;
  protected
    procedure SetFiles(AFiles: TStrings);
    procedure SetMappedNames(ANames: TStrings);
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  published
    property Files: TStrings read GetFiles write SetFiles;
    // MappedNames is only needed if files need to be renamed during a drag op.
    // E.g. dragging from 'Recycle Bin'.
    property MappedNames: TStrings read GetMappedNames write SetMappedNames;
  end;


////////////////////////////////////////////////////////////////////////////////
//
//		Misc.
//
////////////////////////////////////////////////////////////////////////////////
function ReadFilesFromHGlobal(const HGlob: HGlobal; Files: TStrings): boolean; // V4: renamed
function ReadFilesFromData(Data: pointer; Size: integer; Files: TStrings): boolean;
function ReadFilesFromZeroList(const Data: pointer; Size: integer;
  Wide: boolean; Files: TStrings): boolean;
function WriteFilesToZeroList(Data: pointer; Size: integer;
  Wide: boolean; const Files: TStrings): boolean;


////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
//
//			IMPLEMENTATION
//
////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
implementation

uses
{$ifdef VER14_PLUS}
  RTLConsts,
{$else}
  Consts,
{$endif}
  DragDropPIDL,
  SysUtils,
  ShlObj;

////////////////////////////////////////////////////////////////////////////////
//
//		Utilities
//
////////////////////////////////////////////////////////////////////////////////

function ReadFilesFromHGlobal(const HGlob: HGlobal; Files: TStrings): boolean;
var
  DropFiles: PDropFiles;
begin
  DropFiles := PDropFiles(GlobalLock(HGlob));
  try
    Result := ReadFilesFromData(DropFiles, GlobalSize(HGlob), Files)
  finally
    GlobalUnlock(HGlob);
  end;
end;

function ReadFilesFromData(Data: pointer; Size: integer; Files: TStrings): boolean;
var
  Wide: boolean;
begin
  Files.Clear;
  if (Data <> nil) then
  begin
    Wide := PDropFiles(Data)^.fWide;
    dec(Size, PDropFiles(Data)^.pFiles);
    inc(PChar(Data), PDropFiles(Data)^.pFiles);
    ReadFilesFromZeroList(Data, Size, Wide, Files);
  end;

  Result := (Files.Count > 0);
end;

function ReadFilesFromZeroList(const Data: pointer; Size: integer;
  Wide: boolean; Files: TStrings): boolean;
var
  p: PChar;
  StringSize: integer;
begin
  Result := False;
  if (Data <> nil) then
  begin
    p := Data;
    while (Size > 0) and (p^ <> #0) do
    begin
      if (Wide) then
      begin
{$ifdef DD_WIDESTRINGLIST}
        Files.AddWide(PWideChar(p));
{$else}
        Files.Add(PWideChar(p));
{$endif}
        StringSize := (Length(PWideChar(p)) + 1) * 2;
      end else
      begin
        Files.Add(p);
        StringSize := Length(p) + 1;
      end;
      inc(p, StringSize);
      dec(Size, StringSize);
      Result := True;
    end;
  end;
end;

function WriteFilesToZeroList(Data: pointer; Size: integer;
  Wide: boolean; const Files: TStrings): boolean;
var
{$ifdef DD_WIDESTRINGLIST}
  j: integer;
  pw: PWideChar;
  ws: WideString;
  pws: PWideChar;
{$endif}
  i: integer;
  p: PChar;
  StringSize: integer;
  s: string;
begin
  Result := False;
  if (Data <> nil) then
  begin
    p := Data;
    i := 0;
    dec(Size);
    while (Size > 0) and (i < Files.Count) do
    begin
      if (Wide) then
      begin
{$ifdef DD_WIDESTRINGLIST}
        if (Files.Wide) then
        begin
          ws := Files.WideStrings[i];
          pw := PWideChar(p);
          pws := PWideChar(ws);
          j := Size;
          while (j > 0) and (pws^ <> #0) do
          begin
            pw^ := pws^;
            inc(pw);
            inc(pws);
            dec(j, SizeOf(WideChar));
          end;
          pw^ := #0;
          StringSize := (Length(ws)+1)*2;
        end else
{$endif}
        begin
          s := Files[i];
          StringToWideChar(s, PWideChar(p), Size);
          StringSize := (Length(s)+1)*2;
        end;
      end else
      begin
        s := Files[i];
        StrPLCopy(p, s, Size);
        StringSize := Length(s)+1;
      end;
      inc(p, StringSize);
      dec(Size, StringSize);
      inc(i);
      Result := True;
    end;

    // Final teminating zero.
    if (Size >= 0) then
      p^ := #0;
  end;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TFileClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
constructor TFileClipboardFormat.Create;
begin
  inherited Create;
  FFiles := TWideStringList.Create;
  // Note: Setting dwAspect to DVASPECT_SHORT will request that the data source
  // returns the file names in short (8.3) format.
  // FFormatEtc.dwAspect := DVASPECT_SHORT;
  FWide := (Win32Platform = VER_PLATFORM_WIN32_NT);
end;

destructor TFileClipboardFormat.Destroy;
begin
  FFiles.Free;
  inherited Destroy;
end;

function TFileClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  Result := CF_HDROP;
end;

procedure TFileClipboardFormat.Clear;
begin
  FFiles.Clear;
end;

function TFileClipboardFormat.HasData: boolean;
begin
  Result := (FFiles.Count > 0);
end;

function TFileClipboardFormat.GetSize: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to FFiles.Count-1 do
    inc(Result, Length(FFiles[i])+1);
  if (Wide) then
    // Wide strings
    Inc(Result, Result);
  inc(Result, SizeOf(TDropFiles)+2);
end;

function TFileClipboardFormat.ReadData(Value: pointer;
  Size: integer): boolean;
begin
  Result := (Size > SizeOf(TDropFiles));
  if (not Result) then
    exit;

  Result := ReadFilesFromData(Value, Size, FFiles);
end;

function TFileClipboardFormat.WriteData(Value: pointer;
  Size: integer): boolean;
begin
  Result := (Size > SizeOf(TDropFiles));
  if (not Result) then
    exit;

  FillChar(Value^, Size, 0);
  PDropFiles(Value)^.pfiles := SizeOf(TDropFiles);
  PDropFiles(Value)^.fwide := BOOL(ord(Wide));
  inc(PChar(Value), SizeOf(TDropFiles));
  dec(Size, SizeOf(TDropFiles));

  WriteFilesToZeroList(Value, Size, Wide, FFiles);
end;

function TFileClipboardFormat.Assign(Source: TCustomDataFormat): boolean;
begin
  if (Source is TFileDataFormat) then
  begin
    FFiles.Assign(TFileDataFormat(Source).Files);
    Result := True;
  end else
    Result := inherited Assign(Source);
end;

function TFileClipboardFormat.AssignTo(Dest: TCustomDataFormat): boolean;
begin
  if (Dest is TFileDataFormat) then
  begin
    TFileDataFormat(Dest).Files.Assign(FFiles);
    Result := True;
  end else
    Result := inherited AssignTo(Dest);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TFilenameClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
var
  CF_FILENAMEA: TClipFormat = 0;

function TFilenameClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  if (CF_FILENAMEA = 0) then
    CF_FILENAMEA := RegisterClipboardFormat(CFSTR_FILENAMEA);
  Result := CF_FILENAMEA;
end;

function TFilenameClipboardFormat.Assign(Source: TCustomDataFormat): boolean;
begin
  if (Source is TFileDataFormat) then
  begin
    Result := (TFileDataFormat(Source).Files.Count > 0);
    if (Result) then
      Filename := TFileDataFormat(Source).Files[0];
  end else
    Result := inherited Assign(Source);
end;

function TFilenameClipboardFormat.AssignTo(Dest: TCustomDataFormat): boolean;
begin
  if (Dest is TFileDataFormat) then
  begin
    TFileDataFormat(Dest).Files.Add(Filename);
    Result := True;
  end else
    Result := inherited AssignTo(Dest);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TFilenameWClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
var
  CF_FILENAMEW: TClipFormat = 0;

function TFilenameWClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  if (CF_FILENAMEW = 0) then
    CF_FILENAMEW := RegisterClipboardFormat(CFSTR_FILENAMEW);
  Result := CF_FILENAMEW;
end;

function TFilenameWClipboardFormat.Assign(Source: TCustomDataFormat): boolean;
begin
  if (Source is TFileDataFormat) then
  begin
    Result := (TFileDataFormat(Source).Files.Count > 0);
    if (Result) then
      Filename := TFileDataFormat(Source).Files[0];
  end else
    Result := inherited Assign(Source);
end;

function TFilenameWClipboardFormat.AssignTo(Dest: TCustomDataFormat): boolean;
begin
  if (Dest is TFileDataFormat) then
  begin
    TFileDataFormat(Dest).Files.Add(Filename);
    Result := True;
  end else
    Result := inherited AssignTo(Dest);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TFilenameMapClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
var
  CF_FILENAMEMAP: TClipFormat = 0;

constructor TFilenameMapClipboardFormat.Create;
begin
  inherited Create;
  FFileMaps := TStringList.Create;
end;

destructor TFilenameMapClipboardFormat.Destroy;
begin
  FFileMaps.Free;
  inherited Destroy;
end;

function TFilenameMapClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  if (CF_FILENAMEMAP = 0) then
    CF_FILENAMEMAP := RegisterClipboardFormat(CFSTR_FILENAMEMAPA);
  Result := CF_FILENAMEMAP;
end;

procedure TFilenameMapClipboardFormat.Clear;
begin
  FFileMaps.Clear;
end;

function TFilenameMapClipboardFormat.HasData: boolean;
begin
  Result := (FFileMaps.Count > 0);
end;

function TFilenameMapClipboardFormat.GetSize: integer;
var
  i: integer;
begin
  Result := FFileMaps.Count + 1;
  for i := 0 to FFileMaps.Count-1 do
    inc(Result, Length(FFileMaps[i]));
end;

function TFilenameMapClipboardFormat.ReadData(Value: pointer;
  Size: integer): boolean;
begin
  Result := ReadFilesFromZeroList(Value, Size, False, FFileMaps);
end;

function TFilenameMapClipboardFormat.WriteData(Value: pointer;
  Size: integer): boolean;
begin
  Result := WriteFilesToZeroList(Value, Size, False, FFileMaps);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TFilenameMapWClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
var
  CF_FILENAMEMAPW: TClipFormat = 0;

constructor TFilenameMapWClipboardFormat.Create;
begin
  inherited Create;
  FFileMaps := TWideStringList.Create;
end;

destructor TFilenameMapWClipboardFormat.Destroy;
begin
  FFileMaps.Free;
  inherited Destroy;
end;

function TFilenameMapWClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  if (CF_FILENAMEMAPW = 0) then
    CF_FILENAMEMAPW := RegisterClipboardFormat(CFSTR_FILENAMEMAPW);
  Result := CF_FILENAMEMAPW;
end;

procedure TFilenameMapWClipboardFormat.Clear;
begin
  FFileMaps.Clear;
end;

function TFilenameMapWClipboardFormat.HasData: boolean;
begin
  Result := (FFileMaps.Count > 0);
end;

function TFilenameMapWClipboardFormat.GetSize: integer;
var
  i: integer;
begin
  Result := FFileMaps.Count + 1;
  for i := 0 to FFileMaps.Count-1 do
{$ifdef DD_WIDESTRINGLIST}
    inc(Result, Length(FFileMaps.WideStrings[i]));
{$else}
    inc(Result, Length(FFileMaps[i]));
{$endif}
  inc(Result, Result);
end;

function TFilenameMapWClipboardFormat.ReadData(Value: pointer;
  Size: integer): boolean;
begin
  Result := ReadFilesFromZeroList(Value, Size, True, FFileMaps);
end;

function TFilenameMapWClipboardFormat.WriteData(Value: pointer;
  Size: integer): boolean;
begin
  Result := WriteFilesToZeroList(Value, Size, True, FFileMaps);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TFileMapDataFormat
//
////////////////////////////////////////////////////////////////////////////////
constructor TFileMapDataFormat.Create(AOwner: TDragDropComponent);
begin
  inherited Create(AOwner);
  FFileMaps := TStringList.Create;
  TStringList(FFileMaps).OnChanging := DoOnChanging;
end;

destructor TFileMapDataFormat.Destroy;
begin
  FFileMaps.Free;
  inherited Destroy;
end;

function TFileMapDataFormat.Assign(Source: TClipboardFormat): boolean;
begin
  Result := True;

  if (Source is TFilenameMapClipboardFormat) then
    FFileMaps.Assign(TFilenameMapClipboardFormat(Source).FileMaps)

  else if (Source is TFilenameMapWClipboardFormat) then
    FFileMaps.Assign(TFilenameMapWClipboardFormat(Source).FileMaps)

  else
    Result := inherited Assign(Source);
end;

function TFileMapDataFormat.AssignTo(Dest: TClipboardFormat): boolean;
begin
  Result := True;

  if (Dest is TFilenameMapClipboardFormat) then
    TFilenameMapClipboardFormat(Dest).FileMaps.Assign(FFileMaps)

  else if (Dest is TFilenameMapWClipboardFormat) then
    TFilenameMapWClipboardFormat(Dest).FileMaps.Assign(FFileMaps)

  else
    Result := inherited AssignTo(Dest);
end;

procedure TFileMapDataFormat.Clear;
begin
  FFileMaps.Clear;
end;

function TFileMapDataFormat.HasData: boolean;
begin
  Result := (FFileMaps.Count > 0);
end;

function TFileMapDataFormat.NeedsData: boolean;
begin
  Result := (FFileMaps.Count = 0);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TFileDataFormat
//
////////////////////////////////////////////////////////////////////////////////
constructor TFileDataFormat.Create(AOwner: TDragDropComponent);
begin
  inherited Create(AOwner);
  FFiles := TWideStringList.Create;
  // FFiles := TStringList.Create;
  TStringList(FFiles).OnChanging := DoOnChanging;
end;

destructor TFileDataFormat.Destroy;
begin
  FFiles.Free;
  inherited Destroy;
end;

function TFileDataFormat.Assign(Source: TClipboardFormat): boolean;
begin
  Result := True;

  if (Source is TFileClipboardFormat) then
    FFiles.Assign(TFileClipboardFormat(Source).Files)

  else if (Source is TPIDLClipboardFormat) then
    FFiles.Assign(TPIDLClipboardFormat(Source).Filenames)

  else
    Result := inherited Assign(Source);
end;

function TFileDataFormat.AssignTo(Dest: TClipboardFormat): boolean;
begin
  Result := True;

  if (Dest is TFileClipboardFormat) then
    TFileClipboardFormat(Dest).Files.Assign(FFiles)

  else if (Dest is TPIDLClipboardFormat) then
    TPIDLClipboardFormat(Dest).Filenames.Assign(FFiles)

  else
    Result := inherited AssignTo(Dest);
end;

procedure TFileDataFormat.Clear;
begin
  FFiles.Clear;
end;

function TFileDataFormat.HasData: boolean;
begin
  Result := (FFiles.Count > 0);
end;

function TFileDataFormat.NeedsData: boolean;
begin
  Result := (FFiles.Count = 0);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TDropFileTarget
//
////////////////////////////////////////////////////////////////////////////////
constructor TDropFileTarget.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  OptimizedMove := True;

  FFileFormat := TFileDataFormat.Create(Self);
  FFileMapFormat := TFileMapDataFormat.Create(Self);
end;

destructor TDropFileTarget.Destroy;
begin
  FFileFormat.Free;
  FFileMapFormat.Free;
  inherited Destroy;
end;

function TDropFileTarget.GetFiles: TStrings;
begin
  Result := FFileFormat.Files;
end;

function TDropFileTarget.GetMappedNames: TStrings;
begin
  Result := FFileMapFormat.FileMaps;
end;

function TDropFileTarget.GetPreferredDropEffect: LongInt;
begin
  // TODO : Needs explanation of why this is nescessary.
  Result := inherited GetPreferredDropEffect;
  if (Result = DROPEFFECT_NONE) then
    Result := DROPEFFECT_COPY;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TDropFileSource
//
////////////////////////////////////////////////////////////////////////////////
constructor TDropFileSource.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);

  FFileFormat := TFileDataFormat.Create(Self);
  FFileMapFormat := TFileMapDataFormat.Create(Self);
end;

destructor TDropFileSource.Destroy;
begin
  FFileFormat.Free;
  FFileMapFormat.Free;
  inherited Destroy;
end;

function TDropFileSource.GetFiles: TStrings;
begin
  Result := FFileFormat.Files;
end;

function TDropFileSource.GetMappedNames: TStrings;
begin
  Result := FFileMapFormat.FileMaps;
end;

procedure TDropFileSource.SetFiles(AFiles: TStrings);
begin
  FFileFormat.Files.Assign(AFiles);
end;

procedure TDropFileSource.SetMappedNames(ANames: TStrings);
begin
  FFileMapFormat.FileMaps.Assign(ANames);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TWideStrings
//
////////////////////////////////////////////////////////////////////////////////
{$ifdef DD_WIDESTRINGLIST}
function TWideStrings.Add(const S: WideString): Integer;
begin
  Result := inherited Add(S);
end;

function TWideStrings.AddObject(const S: WideString; AObject: TObject): Integer;
begin
  Result := inherited AddObject(S, AObject);
end;
function TWideStrings.IndexOf(const S: WideString): Integer;
begin
  Result := inherited IndexOf(S);
end;

{$endif}

////////////////////////////////////////////////////////////////////////////////
//
//		TWideStringList
//
////////////////////////////////////////////////////////////////////////////////
{$ifdef DD_WIDESTRINGLIST}
function TWideStringList.Add(const S: WideString): Integer;
var
  PWStr: ^TWString;
begin
  New(PWStr);
  PWStr^.WString := S;
  Result := FWideStringList.Add(PWStr);
end;

function TWideStringList.AddObject(const S: WideString; AObject: TObject): Integer;
begin
  Result := Add(S);
  PutObject(Result, AObject);
end;

procedure TWideStringList.AddStrings(Strings: TStrings);
var
  I: Integer;
begin
  if (Strings.Wide) then
  begin
    BeginUpdate;
    try
      for I := 0 to Strings.Count - 1 do
        AddObjectWide(Strings.WideStrings[I], Strings.Objects[I]);
    finally
      EndUpdate;
    end;
  end else
    inherited AddStrings(Strings);
end;

procedure TWideStringList.Clear;
var
  Index: Integer;
  PWStr: ^TWString;
begin
  for Index := 0 to FWideStringList.Count-1 do
  begin
    PWStr := FWideStringList.Items[Index];
    if PWStr <> nil then
      Dispose(PWStr);
  end;
  FWideStringList.Clear;
end;

constructor TWideStringList.Create;
begin
  inherited Create;
  FWideStringList := TList.Create;
end;

procedure TWideStringList.Delete(Index: Integer);
var
  PWStr: ^TWString;
begin
  if (Index < 0) or (Index >= Count) then
    Error(SListIndexError, Index);
  PWStr := FWideStringList.Items[Index];
  if PWStr <> nil then
    Dispose(PWStr);
  FWideStringList.Delete(Index);
end;

destructor TWideStringList.Destroy;
begin
  Clear;
  FWideStringList.Free;

  inherited Destroy;
end;

function TWideStringList.Get(Index: Integer): string;
begin
  Result := GetWide(Index);
end;

function TWideStringList.GetCount: Integer;
begin
  Result := FWideStringList.Count;
end;

function TWideStringList.GetWide(Index: Integer): WideString;
var
  PWStr: ^TWString;
begin
  Result := '';
  if ( (Index >= 0) and (Index < FWideStringList.Count) ) then
  begin
    PWStr := FWideStringList.Items[Index];
    if PWStr <> nil then
      Result := PWStr^.WString;
  end;
end;

function TWideStringList.IndexOf(const S: WideString): Integer;
var
  Index: Integer;
  PWStr: ^TWString;
begin
  Result := -1;
  for Index := 0 to FWideStringList.Count -1 do
  begin
    PWStr := FWideStringList.Items[Index];
    if PWStr <> nil then
    begin
      if S = PWStr^.WString then
      begin
        Result := Index;
        break;
      end;
    end;
  end;
end;

procedure TWideStringList.Insert(Index: Integer; const S: string);
begin
  Insert(Index, WideString(S));
end;

procedure TWideStringList.Insert(Index: Integer; const S: WideString);
var
  PWStr: ^TWString;
begin
  if((Index < 0) or (Index > FWideStringList.Count)) then
    Error(SListIndexError, Index);
  if Index < FWideStringList.Count then
  begin
    PWStr := FWideStringList.Items[Index];
    if PWStr <> nil then
      PWStr.WString := S;
  end
  else
    Add(S);
end;

procedure TWideStringList.PutWide(Index: Integer; const S: WideString);
begin
  Insert(Index, S);
end;

function TWideStringsHelper.Add(const S: WideString): Integer;
begin
  if (Wide) then
    Result := AddWide(S)
  else
    Result := inherited Add(S);
end;

function TWideStringsHelper.AddObject(const S: WideString;
  AObject: TObject): Integer;
begin
  if (Wide) then
    Result := AddObjectWide(S, AObject)
  else
    Result := inherited AddObject(S, AObject);
end;

function TWideStringsHelper.AddObjectWide(const S: WideString;
  AObject: TObject): Integer;
begin
  if (Wide) then
    Result := TWideStrings(Self).AddObject(S, AObject)
  else
    Result := inherited AddObject(S, AObject);
end;

function TWideStringsHelper.AddWide(const S: WideString): Integer;
begin
  if (Wide) then
    Result := TWideStrings(Self).Add(S)
  else
    Result := inherited Add(S);
end;

function TWideStringsHelper.GetWide(Index: Integer): WideString;
begin
  if (Wide) then
    Result := TWideStrings(Self).GetWide(Index)
  else
    Result := Get(Index);
end;

function TWideStringsHelper.GetWideText: WideString;
var
  I, L, Size, Count: Integer;
  P: PWideChar;
  S, LB: WideString;
begin
  if (Wide) then
  begin
    Count := GetCount;
    Size := 0;
    LB := LineBreak;
    for I := 0 to Count - 1 do
      Inc(Size, Length(GetWide(I)) + Length(LB));

    SetLength(Result, Size);
    P := Pointer(Result);
    for I := 0 to Count - 1 do
    begin
      S := GetWide(I);
      L := Length(S);
      if L <> 0 then
      begin
        System.Move(Pointer(S)^, P^, L*2);
        Inc(P, L);
      end;
      L := Length(LB);
      if L <> 0 then
      begin
        System.Move(Pointer(LB)^, P^, L*2);
        Inc(P, L);
      end;
    end;
  end else
    Result := GetTextStr;
end;

function TWideStringsHelper.IndexOf(const S: WideString): Integer;
begin
  if (Wide) then
    Result := IndexOfWide(S)
  else
    Result := inherited IndexOf(S);
end;

function TWideStringsHelper.IndexOfWide(const S: WideString): Integer;
begin
  Result := TWideStrings(Self).IndexOf(S);
end;

procedure TWideStringsHelper.Insert(Index: Integer; const S: WideString);
begin
  if (Wide) then
    InsertWide(Index, S)
  else
    inherited Insert(Index, S);
end;

procedure TWideStringsHelper.InsertWide(Index: Integer; const S: WideString);
begin
  TWideStrings(Self).Insert(Index, S);
end;

function TWideStringsHelper.IsWide: boolean;
begin
  Result := (Self is TWideStrings);
end;

procedure TWideStringsHelper.PutWide(Index: Integer; const S: WideString);
begin
  if (Wide) then
    TWideStrings(Self).PutWide(Index, S)
  else
    Put(Index, S);
end;
{$endif}

////////////////////////////////////////////////////////////////////////////////
//
//		Initialization/Finalization
//
////////////////////////////////////////////////////////////////////////////////

initialization
  // Data format registration
  TFileDataFormat.RegisterDataFormat;
  TFileMapDataFormat.RegisterDataFormat;
  // Clipboard format registration
  TFileDataFormat.RegisterCompatibleFormat(TFileClipboardFormat, 0, csSourceTarget, [ddRead]);
  TFileDataFormat.RegisterCompatibleFormat(TPIDLClipboardFormat, 1, csSourceTarget, [ddRead]);
  TFileDataFormat.RegisterCompatibleFormat(TFilenameClipboardFormat, 2, csSourceTarget, [ddRead]);
  TFileDataFormat.RegisterCompatibleFormat(TFilenameWClipboardFormat, 2, csSourceTarget, [ddRead]);

  TFileMapDataFormat.RegisterCompatibleFormat(TFilenameMapClipboardFormat, 0, csSourceTarget, [ddRead]);
  TFileMapDataFormat.RegisterCompatibleFormat(TFilenameMapWClipboardFormat, 0, csSourceTarget, [ddRead]);

finalization
  // Data format unregistration
  TFileDataFormat.UnregisterDataFormat;
  TFileMapDataFormat.UnregisterDataFormat;

  // Clipboard format unregistration
  TFileClipboardFormat.UnregisterClipboardFormat;
  TFilenameClipboardFormat.UnregisterClipboardFormat;
  TFilenameWClipboardFormat.UnregisterClipboardFormat;
  TFilenameMapClipboardFormat.UnregisterClipboardFormat;
  TFilenameMapWClipboardFormat.UnregisterClipboardFormat;
end.
