unit DragDropFile;

// -----------------------------------------------------------------------------
// Project:         Drag and Drop Component Suite.
// Module:          DragDropFile
// Description:     Implements Dragging and Dropping of files and folders.
// Version:         4.0
// Date:            18-MAY-2001
// Target:          Win32, Delphi 5-6
// Authors:         Anders Melander, anders@melander.dk, http://www.melander.dk
// Copyright        © 1997-2001 Angus Johnson & Anders Melander
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

type
////////////////////////////////////////////////////////////////////////////////
//
//		TFileClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
  TFileClipboardFormat = class(TCustomSimpleClipboardFormat)
  private
    FFiles: TStrings;
  protected
    function ReadData(Value: pointer; Size: integer): boolean; override;
    function WriteData(Value: pointer; Size: integer): boolean; override;
    function GetSize: integer; override;
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
  // DONE -oanme -cStopShip : Rename TFilenameMapClipboardFormat to TFilenameMapClipboardFormat. Also wide version.
  TFilenameMapClipboardFormat = class(TCustomSimpleClipboardFormat)
  private
    FFileMaps		: TStrings;
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
    FFileMaps		: TStrings;
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
    FFileMaps		: TStrings;
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
    FFiles		: TStrings;
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
    FFileFormat		: TFileDataFormat;
    FFileMapFormat	: TFileMapDataFormat;
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
    FFileFormat		: TFileDataFormat;
    FFileMapFormat	: TFileMapDataFormat;
    function GetFiles: TStrings;
    function GetMappedNames: TStrings;
  protected
    procedure SetFiles(AFiles: TStrings);
    procedure SetMappedNames(ANames: TStrings);
  public
    constructor Create(aOwner: TComponent); override;
    destructor Destroy; override;
  published
    property Files: TStrings read GetFiles write SetFiles;
    // MappedNames is only needed if files need to be renamed during a drag op.
    // E.g. dragging from 'Recycle Bin'.
    property MappedNames: TStrings read GetMappedNames write SetMappedNames;
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
function ReadFilesFromHGlobal(const HGlob: HGlobal; Files: TStrings): boolean; // V4: renamed
function ReadFilesFromData(Data: pointer; Size: integer; Files: TStrings): boolean;
function ReadFilesFromZeroList(Data: pointer; Size: integer;
  Wide: boolean; Files: TStrings): boolean;
function WriteFilesToZeroList(Data: pointer; Size: integer;
  Wide: boolean; Files: TStrings): boolean;


////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
//
//			IMPLEMENTATION
//
////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
implementation

uses
  DragDropPIDL,
  SysUtils,
  ShlObj;

////////////////////////////////////////////////////////////////////////////////
//
//		Component registration
//
////////////////////////////////////////////////////////////////////////////////

procedure Register;
begin
  RegisterComponents(DragDropComponentPalettePage, [TDropFileTarget,
    TDropFileSource]);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		Utilities
//
////////////////////////////////////////////////////////////////////////////////

function ReadFilesFromHGlobal(const HGlob: HGlobal; Files: TStrings): boolean;
var
  DropFiles		: PDropFiles;
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
  Wide			: boolean;
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

function ReadFilesFromZeroList(Data: pointer; Size: integer;
  Wide: boolean; Files: TStrings): boolean;
var
  StringSize		: integer;
begin
  Result := False;
  if (Data <> nil) then
    while (Size > 0) and (PChar(Data)^ <> #0) do
    begin
      if (Wide) then
      begin
        Files.Add(PWideChar(Data));
        StringSize := (Length(PWideChar(Data)) + 1) * 2;
      end else
      begin
        Files.Add(PChar(Data));
        StringSize := Length(PChar(Data)) + 1;
      end;
      inc(PChar(Data), StringSize);
      dec(Size, StringSize);
      Result := True;
    end;
end;

function WriteFilesToZeroList(Data: pointer; Size: integer;
  Wide: boolean; Files: TStrings): boolean;
var
  i			: integer;
begin
  Result := False;
  if (Data <> nil) then
  begin
    i := 0;
    dec(Size);
    while (Size > 0) and (i < Files.Count) do
    begin
      if (Wide) then
      begin
        StringToWideChar(Files[i], Data, Size);
        dec(Size, (Length(Files[i])+1)*2);
      end else
      begin
        StrPLCopy(Data, Files[i], Size);
        dec(Size, Length(Files[i])+1);
      end;
      inc(PChar(Data), Length(Files[i])+1);
      inc(i);
      Result := True;
    end;

    // Final teminating zero.
    if (Size >= 0) then
      PChar(Data)^ := #0;
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
  FFiles := TStringList.Create;
  // Note: Setting dwAspect to DVASPECT_SHORT will request that the data source
  // returns the file names in short (8.3) format.
  // FFormatEtc.dwAspect := DVASPECT_SHORT;
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
  i			: integer;
begin
  Result := SizeOf(TDropFiles) + FFiles.Count + 1;
  for i := 0 to FFiles.Count-1 do
    inc(Result, Length(FFiles[i]));
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

  PDropFiles(Value)^.pfiles := SizeOf(TDropFiles);
  PDropFiles(Value)^.fwide := False;
  inc(PChar(Value), SizeOf(TDropFiles));
  dec(Size, SizeOf(TDropFiles));

  WriteFilesToZeroList(Value, Size, False, FFiles);
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
  i			: integer;
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
  FFileMaps := TStringList.Create;
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
  i			: integer;
begin
  Result := FFileMaps.Count + 1;
  for i := 0 to FFileMaps.Count-1 do
    inc(Result, Length(FFileMaps[i]));
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
  FFiles := TStringList.Create;
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
  Result := inherited GetPreferredDropEffect;
  if (Result = DROPEFFECT_NONE) then
    Result := DROPEFFECT_COPY;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TDropFileSource
//
////////////////////////////////////////////////////////////////////////////////
constructor TDropFileSource.Create(aOwner: TComponent);
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
