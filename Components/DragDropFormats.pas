unit DragDropFormats;

// -----------------------------------------------------------------------------
// Project:         Drag and Drop Component Suite.
// Module:          DragDropFormats
// Description:     Implements commonly used clipboard formats and base classes.
// Version:         4.0
// Date:            18-MAY-2001
// Target:          Win32, Delphi 5-6
// Authors:         Anders Melander, anders@melander.dk, http://www.melander.dk
// Copyright        © 1997-2001 Angus Johnson & Anders Melander
// -----------------------------------------------------------------------------

interface

uses
  DragDrop,
  Windows,
  Classes,
  ActiveX,
  ShlObj;

{$include DragDrop.inc}

type
////////////////////////////////////////////////////////////////////////////////
//
//		TStreamList
//
////////////////////////////////////////////////////////////////////////////////
// Utility class used by TFileContentsStreamClipboardFormat and
// TDataStreamDataFormat.
////////////////////////////////////////////////////////////////////////////////
  TStreamList = class(TObject)
  private
    FStreams		: TStrings;
    FOnChanging		: TNotifyEvent;
  protected
    function GetStream(Index: integer): TStream;
    function GetCount: integer;
    procedure Changing;
  public
    constructor Create;
    destructor Destroy; override;
    function Add(Stream: TStream): integer;
    function AddNamed(Stream: TStream; Name: string): integer;
    procedure Delete(Index: integer);
    procedure Clear;
    procedure Assign(Value: TStreamList);
    property Count: integer read GetCount;
    property Streams[Index: integer]: TStream read GetStream; default;
    property Names: TStrings read FStreams;
    property OnChanging: TNotifyEvent read FOnChanging write FOnChanging;
  end;


////////////////////////////////////////////////////////////////////////////////
//
//		TInterfaceList
//
////////////////////////////////////////////////////////////////////////////////
// List of named interfaces.
// Note: Delphi 5 also implements a TInterfaceList, but it can not be used
// because it doesn't support change notification and isn't extensible.
////////////////////////////////////////////////////////////////////////////////
// Utility class used by TFileContentsStorageClipboardFormat.
////////////////////////////////////////////////////////////////////////////////
  TInterfaceList = class(TObject)
  private
    FList		: TStrings;
    FOnChanging		: TNotifyEvent;
  protected
    function GetCount: integer;
    function GetName(Index: integer): string;
    function GetItem(Index: integer): IUnknown;
    procedure Changing;
  public
    constructor Create;
    destructor Destroy; override;
    function Add(Item: IUnknown): integer;
    function AddNamed(Item: IUnknown; Name: string): integer;
    procedure Delete(Index: integer);
    procedure Clear;
    procedure Assign(Value: TInterfaceList);
    property Items[Index: integer]: IUnknown read GetItem; default;
    property Names[Index: integer]: string read GetName;
    property Count: integer read GetCount;
    property OnChanging: TNotifyEvent read FOnChanging write FOnChanging;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TStorageInterfaceList
//
////////////////////////////////////////////////////////////////////////////////
// List of IStorage interfaces.
// Used by TFileContentsStorageClipboardFormat.
////////////////////////////////////////////////////////////////////////////////
  TStorageInterfaceList = class(TInterfaceList)
  private
  protected
    function GetStorage(Index: integer): IStorage;
  public
    property Storages[Index: integer]: IStorage read GetStorage; default;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TFixedStreamAdapter
//
////////////////////////////////////////////////////////////////////////////////
// TFixedStreamAdapter fixes several serious bugs in TStreamAdapter.CopyTo.
////////////////////////////////////////////////////////////////////////////////
  TFixedStreamAdapter = class(TStreamAdapter, IStream)
  public
    function CopyTo(stm: IStream; cb: Largeint; out cbRead: Largeint;
      out cbWritten: Largeint): HResult; override; stdcall;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TMemoryList
//
////////////////////////////////////////////////////////////////////////////////
// List which owns the memory blocks it points to.
////////////////////////////////////////////////////////////////////////////////
  TMemoryList = class(TObject)
  private
    FList: TList;
  protected
    function Get(Index: Integer): Pointer;
    function GetCount: Integer;
  public
    constructor Create;
    destructor Destroy; override;
    function Add(Item: Pointer): Integer;
    procedure Clear;
    procedure Delete(Index: Integer);
    property Count: Integer read GetCount;
    property Items[Index: Integer]: Pointer read Get; default;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TCustomSimpleClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
// Abstract base class for simple clipboard formats stored in global memory
// or a stream.
////////////////////////////////////////////////////////////////////////////////
//
// Two different methods of data transfer from the medium to the object are
// supported:
//
//   1) Descendant class reads data from a buffer provided by the base class.
//
//   2) Base class reads data from a buffer provided by the descendant class.
//
// Method #1 only requires that the descedant class implements the ReadData.
//
// Method #2 requires that the descedant class overrides the default
// DoGetDataSized method. The descedant DoGetDataSized method should allocate a
// buffer of the specified size and then call the ReadDataInto method to
// transfer data to the buffer. Even though the ReadData method will not be used
// in this scenario, it should be implemented as an empty method (to avoid
// abstract warnings).
//
// The WriteData method must be implemented regardless of which of the two
// approaches the class implements.
//
////////////////////////////////////////////////////////////////////////////////
  TCustomSimpleClipboardFormat = class(TClipboardFormat)
  private
  protected
    function DoGetData(ADataObject: IDataObject; const AMedium: TStgMedium): boolean; override;
    //: Transfer data from medium to a buffer of the specified size.
    function DoGetDataSized(ADataObject: IDataObject; const AMedium: TStgMedium;
      Size: integer): boolean; virtual;
    //: Transfer data from the specified buffer to the objects storage.
    function ReadData(Value: pointer; Size: integer): boolean; virtual; abstract;
    //: Transfer data from the medium to the specified buffer.
    function ReadDataInto(ADataObject: IDataObject; const AMedium: TStgMedium;
      Buffer: pointer; Size: integer): boolean; virtual;

    function DoSetData(const FormatEtcIn: TFormatEtc;
      var AMedium: TStgMedium): boolean; override;
    //: Transfer data from the objects storage to the specified buffer.
    function WriteData(Value: pointer; Size: integer): boolean; virtual; abstract;
    function GetSize: integer; virtual; abstract;
  public
    constructor Create; override;
  end;


////////////////////////////////////////////////////////////////////////////////
//
//		TCustomStringClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
// Abstract base class for simple clipboard formats.
// The data is stored in a string.
////////////////////////////////////////////////////////////////////////////////
  TCustomStringClipboardFormat = class(TCustomSimpleClipboardFormat)
  private
    FData: string;
    FTrimZeroes: boolean;
  protected
    function ReadData(Value: pointer; Size: integer): boolean; override;
    function WriteData(Value: pointer; Size: integer): boolean; override;
    function GetSize: integer; override;

    function GetString: string;
    procedure SetString(const Value: string);
    property Data: string read FData write FData; // DONE : Why is SetString used instead of FData?
  public
    procedure Clear; override;
    function HasData: boolean; override;
    property TrimZeroes: boolean read FTrimZeroes write FTrimZeroes;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TCustomStringListClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
// Abstract base class for simple cr/lf delimited string clipboard formats.
// The data is stored in a TStringList.
////////////////////////////////////////////////////////////////////////////////
  TCustomStringListClipboardFormat = class(TCustomSimpleClipboardFormat)
  private
    FLines		: TStrings;
  protected
    function ReadData(Value: pointer; Size: integer): boolean; override;
    function WriteData(Value: pointer; Size: integer): boolean; override;
    function GetSize: integer; override;

    function GetLines: TStrings;
    property Lines: TStrings read FLines;
  public
    constructor Create; override;
    destructor Destroy; override;
    procedure Clear; override;
    function HasData: boolean; override;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TCustomTextClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
// Abstract base class for simple text based clipboard formats.
////////////////////////////////////////////////////////////////////////////////
  TCustomTextClipboardFormat = class(TCustomStringClipboardFormat)
  private
  protected
    function GetSize: integer; override;
    property Text: string read GetString write SetString;
  public
    constructor Create; override;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TCustomWideTextClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
// Abstract base class for simple wide string clipboard formats storing the data
// in a wide string.
////////////////////////////////////////////////////////////////////////////////
  TCustomWideTextClipboardFormat = class(TCustomSimpleClipboardFormat)
  private
    FText		: WideString;
  protected
    function ReadData(Value: pointer; Size: integer): boolean; override;
    function WriteData(Value: pointer; Size: integer): boolean; override;
    function GetSize: integer; override;

    function GetText: WideString;
    procedure SetText(const Value: WideString);
    property Text: WideString read FText write FText;
  public
    procedure Clear; override;
    function HasData: boolean; override;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TTextClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
  TTextClipboardFormat = class(TCustomTextClipboardFormat)
  public
    function GetClipboardFormat: TClipFormat; override;
    property Text;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TCustomDWORDClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
  TCustomDWORDClipboardFormat = class(TCustomSimpleClipboardFormat)
  private
    FValue		: DWORD;
  protected
    function ReadData(Value: pointer; Size: integer): boolean; override;
    function WriteData(Value: pointer; Size: integer): boolean; override;
    function GetSize: integer; override;

    function GetValueDWORD: DWORD;
    procedure SetValueDWORD(Value: DWORD);
    function GetValueInteger: integer;
    procedure SetValueInteger(Value: integer);
    function GetValueLongInt: longInt;
    procedure SetValueLongInt(Value: longInt);
    function GetValueBoolean: boolean;
    procedure SetValueBoolean(Value: boolean);
  public
    procedure Clear; override;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TFileGroupDescritorClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
  TFileGroupDescritorClipboardFormat = class(TCustomSimpleClipboardFormat)
  private
    FFileGroupDescriptor	: PFileGroupDescriptor;
  protected
    function ReadData(Value: pointer; Size: integer): boolean; override;
    function WriteData(Value: pointer; Size: integer): boolean; override;
    function GetSize: integer; override;
  public
    function GetClipboardFormat: TClipFormat; override;
    destructor Destroy; override;
    procedure Clear; override;
    function HasData: boolean; override;
    property FileGroupDescriptor: PFileGroupDescriptor read FFileGroupDescriptor;
    procedure CopyFrom(AFileGroupDescriptor: PFileGroupDescriptor);
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TFileGroupDescritorWClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
  // Warning: TFileGroupDescriptorW has wrong declaration in ShlObj.pas!
  TFileGroupDescriptorW = record
    cItems: UINT;
    fgd: array[0..0] of TFileDescriptorW;
  end;

  PFileGroupDescriptorW = ^TFileGroupDescriptorW;

  TFileGroupDescritorWClipboardFormat = class(TCustomSimpleClipboardFormat)
  private
    FFileGroupDescriptor	: PFileGroupDescriptorW;
  protected
    function ReadData(Value: pointer; Size: integer): boolean; override;
    function WriteData(Value: pointer; Size: integer): boolean; override;
    function GetSize: integer; override;
  public
    function GetClipboardFormat: TClipFormat; override;
    destructor Destroy; override;
    procedure Clear; override;
    function HasData: boolean; override;
    property FileGroupDescriptor: PFileGroupDescriptorW read FFileGroupDescriptor;
    procedure CopyFrom(AFileGroupDescriptor: PFileGroupDescriptorW);
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TFileContentsClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
// Note: File contents must be zero terminated, so we descend from
// TCustomTextClipboardFormat instead of TCustomStringClipboardFormat.
////////////////////////////////////////////////////////////////////////////////
  TFileContentsClipboardFormat = class(TCustomTextClipboardFormat)
  public
    function GetClipboardFormat: TClipFormat; override;
    constructor Create; override;
    property Data;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TFileContentsStreamClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
  TFileContentsStreamClipboardFormat = class(TClipboardFormat)
  private
    FStreams: TStreamList;
  protected
  public
    constructor Create; override;
    destructor Destroy; override;
    function GetClipboardFormat: TClipFormat; override;
    function GetData(DataObject: IDataObject): boolean; override;
    procedure Clear; override;
    function HasData: boolean; override;
    function AssignTo(Dest: TCustomDataFormat): boolean; override;
    property Streams: TStreamList read FStreams;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TFileContentsStreamOnDemandClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
// Yeah, it's a long name, but I like my names descriptive.
////////////////////////////////////////////////////////////////////////////////
  TVirtualFileStreamDataFormat = class;
  TFileContentsStreamOnDemandClipboardFormat = class;

  TOnGetStreamEvent = procedure(Sender: TFileContentsStreamOnDemandClipboardFormat;
    Index: integer; out AStream: IStream) of object;

  TFileContentsStreamOnDemandClipboardFormat = class(TClipboardFormat)
  private
    FOnGetStream: TOnGetStreamEvent;
    FGotData: boolean;
    FDataRequested: boolean;
  protected
    function DoSetData(const FormatEtcIn: TFormatEtc;
      var AMedium: TStgMedium): boolean; override;
  public
    constructor Create; override;
    destructor Destroy; override;
    function GetClipboardFormat: TClipFormat; override;
    function GetData(DataObject: IDataObject): boolean; override;
    procedure Clear; override;
    function HasData: boolean; override;
    function AssignTo(Dest: TCustomDataFormat): boolean; override;
    function Assign(Source: TCustomDataFormat): boolean; override;

    function GetStream(Index: integer): IStream;

    property OnGetStream: TOnGetStreamEvent read FOnGetStream write FOnGetStream;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TFileContentsStorageClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
  TFileContentsStorageClipboardFormat = class(TClipboardFormat)
  private
    FStorages		: TStorageInterfaceList;
  protected
  public
    constructor Create; override;
    destructor Destroy; override;
    function GetClipboardFormat: TClipFormat; override;
    function GetData(DataObject: IDataObject): boolean; override;
    procedure Clear; override;
    function HasData: boolean; override;
    function AssignTo(Dest: TCustomDataFormat): boolean; override;
    property Storages: TStorageInterfaceList read FStorages;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TPreferredDropEffectClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
  TPreferredDropEffectClipboardFormat = class(TCustomDWORDClipboardFormat)
  public
    class function GetClassClipboardFormat: TClipFormat;
    function GetClipboardFormat: TClipFormat; override;
    function HasData: boolean; override;
    property Value: longInt read GetValueLongInt write SetValueLongInt;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TPerformedDropEffectClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
  TPerformedDropEffectClipboardFormat = class(TCustomDWORDClipboardFormat)
  public
    function GetClipboardFormat: TClipFormat; override;
    property Value: longInt read GetValueLongInt write SetValueLongInt;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TLogicalPerformedDropEffectClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
// Microsoft's latest (so far) "logical" solution to the never ending attempts
// of reporting back to the source which operation actually took place. Sigh!
////////////////////////////////////////////////////////////////////////////////
  TLogicalPerformedDropEffectClipboardFormat = class(TCustomDWORDClipboardFormat)
  public
    function GetClipboardFormat: TClipFormat; override;
    property Value: longInt read GetValueLongInt write SetValueLongInt;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TPasteSuccededClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
  TPasteSuccededClipboardFormat = class(TCustomDWORDClipboardFormat)
  public
    function GetClipboardFormat: TClipFormat; override;
    property Value: longInt read GetValueLongInt write SetValueLongInt;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TInDragLoopClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
  TInShellDragLoopClipboardFormat = class(TCustomDWORDClipboardFormat)
  public
    function GetClipboardFormat: TClipFormat; override;
    property InShellDragLoop: boolean read GetValueBoolean write SetValueBoolean;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TTargetCLSIDClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
  TTargetCLSIDClipboardFormat = class(TCustomSimpleClipboardFormat)
  private
    FCLSID: TCLSID;
  protected
    function ReadData(Value: pointer; Size: integer): boolean; override;
    function WriteData(Value: pointer; Size: integer): boolean; override;
    function GetSize: integer; override;
  public
    function GetClipboardFormat: TClipFormat; override;
    procedure Clear; override;
    function HasData: boolean; override;
    property CLSID: TCLSID read FCLSID write FCLSID;
  end;

////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////


////////////////////////////////////////////////////////////////////////////////
//
//		TTextDataFormat
//
////////////////////////////////////////////////////////////////////////////////
  TTextDataFormat = class(TCustomDataFormat)
  private
    FText		: string;
  protected
    procedure SetText(const Value: string);
  public
    function Assign(Source: TClipboardFormat): boolean; override;
    function AssignTo(Dest: TClipboardFormat): boolean; override;
    procedure Clear; override;
    function HasData: boolean; override;
    function NeedsData: boolean; override;
    property Text: string read FText write SetText;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TDataStreamDataFormat
//
////////////////////////////////////////////////////////////////////////////////
  TDataStreamDataFormat = class(TCustomDataFormat)
  private
    FStreams		: TStreamList;
  public
    constructor Create(AOwner: TDragDropComponent); override;
    destructor Destroy; override;
    procedure Clear; override;
    function HasData: boolean; override;
    function NeedsData: boolean; override;
    property Streams: TStreamList read FStreams;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TVirtualFileStreamDataFormat
//
////////////////////////////////////////////////////////////////////////////////
  TVirtualFileStreamDataFormat = class(TCustomDataFormat)
  private
    FFileDescriptors: TMemoryList;
    FFileNames: TStrings;
    FFileContentsClipboardFormat: TFileContentsStreamOnDemandClipboardFormat;
    FFileGroupDescritorClipboardFormat: TFileGroupDescritorClipboardFormat;
    FHasContents: boolean;
  protected
    procedure SetFileNames(const Value: TStrings);
    function GetOnGetStream: TOnGetStreamEvent;
    procedure SetOnGetStream(const Value: TOnGetStreamEvent);
  public
    constructor Create(AOwner: TDragDropComponent); override;
    destructor Destroy; override;

    function Assign(Source: TClipboardFormat): boolean; override;
    function AssignTo(Dest: TClipboardFormat): boolean; override;
    procedure Clear; override;
    function HasData: boolean; override;
    function NeedsData: boolean; override;
    property FileDescriptors: TMemoryList read FFileDescriptors;
    property FileNames: TStrings read FFileNames write SetFileNames;
    property FileContentsClipboardFormat: TFileContentsStreamOnDemandClipboardFormat
      read FFileContentsClipboardFormat;
    property FileGroupDescritorClipboardFormat: TFileGroupDescritorClipboardFormat
      read FFileGroupDescritorClipboardFormat;
    property OnGetStream: TOnGetStreamEvent read GetOnGetStream write SetOnGetStream;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TFeedbackDataFormat
//
////////////////////////////////////////////////////////////////////////////////
// Data used for communication between source and target.
// Only used by the drop source.
////////////////////////////////////////////////////////////////////////////////
  TFeedbackDataFormat = class(TCustomDataFormat)
  private
    FPreferredDropEffect: longInt;
    FPerformedDropEffect: longInt;
    FLogicalPerformedDropEffect: longInt;
    FPasteSucceded: longInt;
    FInShellDragLoop: boolean;
    FGotInShellDragLoop: boolean;
    FTargetCLSID: TCLSID;
  protected
    procedure SetInShellDragLoop(const Value: boolean);
    procedure SetPasteSucceded(const Value: longInt);
    procedure SetPerformedDropEffect(const Value: longInt);
    procedure SetPreferredDropEffect(const Value: longInt);
    procedure SetTargetCLSID(const Value: TCLSID);
    procedure SetLogicalPerformedDropEffect(const Value: Integer);
  public
    function Assign(Source: TClipboardFormat): boolean; override;
    function AssignTo(Dest: TClipboardFormat): boolean; override;
    procedure Clear; override;
    function HasData: boolean; override;
    function NeedsData: boolean; override;
    property PreferredDropEffect: longInt read FPreferredDropEffect
      write SetPreferredDropEffect;
    property PerformedDropEffect: longInt read FPerformedDropEffect
      write SetPerformedDropEffect;
    property LogicalPerformedDropEffect: longInt read FLogicalPerformedDropEffect
      write SetLogicalPerformedDropEffect;
    property PasteSucceded: longInt read FPasteSucceded write SetPasteSucceded;
    property InShellDragLoop: boolean read FInShellDragLoop
      write SetInShellDragLoop;
    property TargetCLSID: TCLSID read FTargetCLSID write SetTargetCLSID;
  end;


////////////////////////////////////////////////////////////////////////////////
//
//		TGenericClipboardFormat & TGenericDataFormat
//
////////////////////////////////////////////////////////////////////////////////
// TGenericDataFormat is not used internally by the library, but can be used to
// add support for new formats with a minimum of custom code.
// Even though TGenericDataFormat represents the data as a string, it can be
// used to transfer any kind of data.
// TGenericClipboardFormat is used internally by TGenericDataFormat but can also
// be used by other TCustomDataFormat descendants or as a base class for new
// clipboard formats.
// Note that you should not register TGenericClipboardFormat as compatible with
// TGenericDataFormat.
// To use TGenericDataFormat, all you need to do is instantiate it against
// the desired component and register your custom clipboard formats:
//
// var
//   MyCustomData: TGenericDataFormat;
//
//   MyCustomData := TGenericDataFormat.Create(DropTextTarget1);
//   MyCustomData.AddFormat('MyCustomFormat');
//
////////////////////////////////////////////////////////////////////////////////
  TGenericDataFormat = class(TCustomDataFormat)
  private
    FData		: string;
  protected
    function GetSize: integer;
    procedure DoSetData(const Value: string);
  public
    procedure Clear; override;
    function HasData: boolean; override;
    function NeedsData: boolean; override;
    procedure AddFormat(const AFormat: string);
    procedure SetDataHere(const AData; ASize: integer);
    function GetDataHere(var AData; ASize: integer): integer;
    property Data: string read FData write DoSetData;
    property Size: integer read GetSize;
  end;

  TGenericClipboardFormat = class(TCustomStringClipboardFormat)
  private
    FFormat: string;
  protected
    procedure SetClipboardFormatName(const Value: string); override;
    function GetClipboardFormatName: string; override;
    function GetClipboardFormat: TClipFormat; override;
  public
    function Assign(Source: TCustomDataFormat): boolean; override;
    function AssignTo(Dest: TCustomDataFormat): boolean; override;
    property Data;
  end;



////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
//
//			IMPLEMENTATION
//
////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
implementation

uses
  DropSource,
  DropTarget,
  SysUtils;

////////////////////////////////////////////////////////////////////////////////
//
//		TStreamList
//
////////////////////////////////////////////////////////////////////////////////
constructor TStreamList.Create;
begin
  inherited Create;
  FStreams := TStringList.Create;
end;

destructor TStreamList.Destroy;
begin
  Clear;
  FStreams.Free;
  inherited Destroy;
end;

procedure TStreamList.Changing;
begin
  if (Assigned(OnChanging)) then
    OnChanging(Self);
end;

function TStreamList.GetStream(Index: integer): TStream;
begin
  Result := TStream(FStreams.Objects[Index]);
end;

function TStreamList.Add(Stream: TStream): integer;
begin
  Result := AddNamed(Stream, '');
end;

function TStreamList.AddNamed(Stream: TStream; Name: string): integer;
begin
  Changing;
  Result := FStreams.AddObject(Name, Stream);
end;

function TStreamList.GetCount: integer;
begin
  Result := FStreams.Count;
end;

procedure TStreamList.Assign(Value: TStreamList);
begin
  Clear;
  FStreams.Assign(Value.Names);
  // Transfer ownership of objects
  Value.FStreams.Clear;
end;

procedure TStreamList.Delete(Index: integer);
begin
  Changing;
  FStreams.Delete(Index);
end;

procedure TStreamList.Clear;
var
  i			: integer;
begin
  Changing;
  for i := 0 to FStreams.Count-1 do
    if (FStreams.Objects[i] <> nil) then
      FStreams.Objects[i].Free;
  FStreams.Clear;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TInterfaceList
//
////////////////////////////////////////////////////////////////////////////////
constructor TInterfaceList.Create;
begin
  inherited Create;
  FList := TStringList.Create;
end;

destructor TInterfaceList.Destroy;
begin
  Clear;
  FList.Free;
  inherited Destroy;
end;

function TInterfaceList.Add(Item: IUnknown): integer;
begin
  Result := AddNamed(Item, '');
end;

function TInterfaceList.AddNamed(Item: IUnknown; Name: string): integer;
begin
  Changing;
  with FList do
  begin
    Result := AddObject(Name, nil);
    Objects[Result] := TObject(Item);
    Item._AddRef;
  end;
end;

procedure TInterfaceList.Changing;
begin
  if (Assigned(OnChanging)) then
    OnChanging(Self);
end;

procedure TInterfaceList.Clear;
var
  i			: Integer;
  p			: pointer;
begin
  Changing;
  with FList do
  begin
    for i := 0 to Count - 1 do
    begin
      p := Objects[i];
      IUnknown(p) := nil;
    end;
    Clear;
  end;
end;

procedure TInterfaceList.Assign(Value: TInterfaceList);
var
  i			: Integer;
begin
  Changing;
  for i := 0 to Value.Count - 1 do
    AddNamed(Value.Items[i], Value.Names[i]);
end;

procedure TInterfaceList.Delete(Index: integer);
var
  p			: pointer;
begin
  Changing;
  with FList do
  begin
    p := Objects[Index];
    IUnknown(p) := nil;
    Delete(Index);
  end;
end;

function TInterfaceList.GetCount: integer;
begin
  Result := FList.Count;
end;

function TInterfaceList.GetName(Index: integer): string;
begin
  Result := FList[Index];
end;

function TInterfaceList.GetItem(Index: integer): IUnknown;
var
  p			: pointer;
begin
  p := FList.Objects[Index];
  Result := IUnknown(p);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TStorageInterfaceList
//
////////////////////////////////////////////////////////////////////////////////
function TStorageInterfaceList.GetStorage(Index: integer): IStorage;
begin
  Result := IStorage(Items[Index]);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TMemoryList
//
////////////////////////////////////////////////////////////////////////////////
function TMemoryList.Add(Item: Pointer): Integer;
begin
  Result := FList.Add(Item);
end;

procedure TMemoryList.Clear;
var
  i: integer;
begin
  for i := FList.Count-1 downto 0 do
    Delete(i);
end;

constructor TMemoryList.Create;
begin
  inherited Create;
  FList := TList.Create;
end;

procedure TMemoryList.Delete(Index: Integer);
begin
  Freemem(FList[Index]);
  FList.Delete(Index);
end;

destructor TMemoryList.Destroy;
begin
  FList.Free;
  inherited Destroy;
end;

function TMemoryList.Get(Index: Integer): Pointer;
begin
  Result := FList[Index];
end;

function TMemoryList.GetCount: Integer;
begin
  Result := FList.Count;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TFixedStreamAdapter
//
////////////////////////////////////////////////////////////////////////////////
function TFixedStreamAdapter.CopyTo(stm: IStream; cb: Largeint; out cbRead: Largeint;
  out cbWritten: Largeint): HResult;
const
  MaxBufSize = 1024 * 1024;  // 1mb
var
  Buffer: Pointer;
  BufSize, BurstReadSize, BurstWriteSize: Integer;
  BytesRead, BytesWritten, BurstWritten: LongInt;
begin
  Result := S_OK;
  BytesRead := 0;
  BytesWritten := 0;
  try
    if (cb < 0) then
    begin
      // Note: The folowing is a workaround for a design bug in either explorer
      // or the clipboard. See comment in TCustomSimpleClipboardFormat.DoSetData
      // for an explanation.
      if (Stream.Position = Stream.Size) then
        Stream.Position := 0;

      cb := Stream.Size - Stream.Position;
    end;
    if cb > MaxBufSize then
      BufSize := MaxBufSize
    else
      BufSize := Integer(cb);
    GetMem(Buffer, BufSize);
    try
      while cb > 0 do
      begin
        if cb > BufSize then
          BurstReadSize := BufSize
        else
          BurstReadSize := cb;

        BurstWriteSize := Stream.Read(Buffer^, BurstReadSize);
        if (BurstWriteSize = 0) then
          break;
        Inc(BytesRead, BurstWriteSize);
        BurstWritten := 0;
        Result := stm.Write(Buffer, BurstWriteSize, @BurstWritten);
        Inc(BytesWritten, BurstWritten);
        if (Result = S_OK) and (Integer(BurstWritten) <> BurstWriteSize) then
          Result := E_FAIL;
        if Result <> S_OK then
          Exit;
        Dec(cb, BurstWritten);
      end;
    finally
      FreeMem(Buffer);
      if (@cbWritten <> nil) then
        cbWritten := BytesWritten;
      if (@cbRead <> nil) then
        cbRead := BytesRead;
    end;
  except
    Result := E_UNEXPECTED;
  end;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TCustomSimpleClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
constructor TCustomSimpleClipboardFormat.Create;
begin
  CreateFormat(TYMED_HGLOBAL or TYMED_ISTREAM);
end;

function TCustomSimpleClipboardFormat.DoGetData(ADataObject: IDataObject;
  const AMedium: TStgMedium): boolean;
var
  Stream		: IStream;
  StatStg		: TStatStg;
  Size			: integer;
begin
  // Get size from HGlobal.
  if (AMedium.tymed and TYMED_HGLOBAL <> 0) then
  begin
    Size := GlobalSize(AMedium.HGlobal);
    Result := True;
  end else
  // Get size from IStream.
  if (AMedium.tymed and TYMED_ISTREAM <> 0) then
  begin
    Stream := IStream(AMedium.stm);
    Result := (Stream <> nil) and (Stream.Stat(StatStg, STATFLAG_NONAME) = S_OK);
    Size := StatStg.cbSize;
    Stream := nil; // Not really nescessary.
  end else
  begin
    Size := 0;
    Result := False;
  end;

  if (Result) and (Size > 0) then
  begin
    // Read the given amount of data.
    Result := DoGetDataSized(ADataObject, AMedium, Size);
  end;
end;

function TCustomSimpleClipboardFormat.DoGetDataSized(ADataObject: IDataObject;
  const AMedium: TStgMedium; Size: integer): boolean;
var
  Buffer: pointer;
  Stream: IStream;
  Remaining: longInt;
  Chunk: longInt;
  pChunk: PChar;
begin
  if (Size > 0) then
  begin
    (*
    ** In this method we prefer TYMED_HGLOBAL over TYMED_ISTREAM and thus check
    ** for TYMED_HGLOBAL first.
    *)

    // Read data from HGlobal
    if (AMedium.tymed and TYMED_HGLOBAL <> 0) then
    begin
      // Use global memory as buffer
      Buffer := GlobalLock(AMedium.HGlobal);
      try
        // Read data from buffer into object
        Result := (Buffer <> nil) and (ReadData(Buffer, Size));
      finally
        GlobalUnlock(AMedium.HGlobal);
      end;
    end else
    // Read data from IStream
    if (AMedium.tymed and TYMED_ISTREAM <> 0) then
    begin
      // Allocate buffer
      GetMem(Buffer, Size);
      try
        // Read data from stream into buffer
        Stream := IStream(AMedium.stm);
        if (Stream <> nil) then
        begin
          Stream.Seek(0, STREAM_SEEK_SET, PLargeuint(nil)^);
          Result := True;
          Remaining := Size;
          pChunk := Buffer;
          while (Result) and (Remaining > 0) do
          begin
            Result := (Stream.Read(pChunk, Remaining, @Chunk) = S_OK);
            if (Chunk = 0) then
              break;
            inc(pChunk, Chunk);
            dec(Remaining, Chunk);
          end;
          Stream := nil; // Not really nescessary.
        end else
          Result := False;
        // Transfer data from buffer into object.
        Result := Result and (ReadData(Buffer, Size));
      finally
        FreeMem(Buffer);
      end;
    end else
      Result := False;
  end else
    Result := False;
end;

function TCustomSimpleClipboardFormat.ReadDataInto(ADataObject: IDataObject;
  const AMedium: TStgMedium; Buffer: pointer; Size: integer): boolean;
var
  Stream: IStream;
  p: pointer;
  Remaining: longInt;
  Chunk: longInt;
begin
  Result := (Buffer <> nil) and (Size > 0);
  if (Result) then
  begin
    // Read data from HGlobal
    if (AMedium.tymed and TYMED_HGLOBAL <> 0) then
    begin
      p := GlobalLock(AMedium.HGlobal);
      try
        Result := (p <> nil);
        if (Result) then
          Move(p^, Buffer^, Size);
      finally
        GlobalUnlock(AMedium.HGlobal);
      end;
    end else
    // Read data from IStream
    if (AMedium.tymed and TYMED_ISTREAM <> 0) then
    begin
      Stream := IStream(AMedium.stm);
      if (Stream <> nil) then
      begin
        Stream.Seek(0, STREAM_SEEK_SET, PLargeuint(nil)^);
        Remaining := Size;
        while (Result) and (Remaining > 0) do
        begin
          Result := (Stream.Read(Buffer, Remaining, @Chunk) = S_OK);
          if (Chunk = 0) then
            break;
          inc(PChar(Buffer), Chunk);
          dec(Remaining, Chunk);
        end;
      end else
        Result := False;
    end else
      Result := False;
  end;
end;

function TCustomSimpleClipboardFormat.DoSetData(const FormatEtcIn: TFormatEtc;
  var AMedium: TStgMedium): boolean;
var
  p: pointer;
  Size: integer;
  Stream: TMemoryStream;
  // Warning: TStreamAdapter.CopyTo is broken!
  StreamAdapter: TStreamAdapter;
begin
  Result := False;

  Size := GetSize;
  if (Size <= 0) then
    exit;

  if (FormatEtcIn.tymed and TYMED_ISTREAM <> 0) then
  begin

    Stream := TMemoryStream.Create;
    StreamAdapter := TFixedStreamAdapter.Create(Stream, soOwned);

    try
      Stream.Size := Size;
      Result := WriteData(Stream.Memory, Size);
      // Note: Conflicting information on which of the following two are correct:
      //
      //   1) Stream.Position := Size;
      //
      //   2) Stream.Position := 0;
      //
      // #1 is required for clipboard operations to succeed; The clipboard uses
      // a Seek(0, STREAM_SEEK_CUR) to determine the size of the stream.
      //
      // #2 is required for shell operations to succeed; The shell uses a
      // Read(-1) to read all of the stream.
      //
      // This library uses a Stream.Stat to determine the size of the stream and
      // then reads from start to end of stream.
      //
      // Since we use #1 (see below), we work around #2 in
      // TFixedStreamAdapter.CopyTo.
      if (Result) then
      begin
        Stream.Position := Size;
        IStream(AMedium.stm) := StreamAdapter as IStream;
      end;
    except
      Result := False;
    end;

    if (not Result) then
    begin
      StreamAdapter.Free;
      AMedium.stm := nil;
    end else
      AMedium.tymed := TYMED_ISTREAM;

  end else
  if (FormatEtcIn.tymed and TYMED_HGLOBAL <> 0) then
  begin

    AMedium.hGlobal := GlobalAlloc(GMEM_SHARE or GHND, Size);
    if (AMedium.hGlobal = 0) then
      exit;

    try
      p := GlobalLock(AMedium.hGlobal);
      try
        Result := (p <> nil) and WriteData(p, Size);
      finally
        GlobalUnlock(AMedium.hGlobal);
      end;
    except
      Result := False;
    end;

    if (not Result) then
    begin
      GlobalFree(AMedium.hGlobal);
      AMedium.hGlobal := 0;
    end else
      AMedium.tymed := TYMED_HGLOBAL;

  end else
    Result := False;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TCustomStringClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
procedure TCustomStringClipboardFormat.Clear;
begin
  FData := '';
end;

function TCustomStringClipboardFormat.HasData: boolean;
begin
  Result := (FData <> '');
end;


function TCustomStringClipboardFormat.ReadData(Value: pointer;
  Size: integer): boolean;
begin
  SetLength(FData, Size);
  Move(Value^, PChar(FData)^, Size);

  // IE adds a lot of trailing zeroes which is included in the string length.
  // To avoid confusion, we trim all trailing zeroes but the last (which is
  // managed automatically by Delphi).
  // Note that since this work around, if applied generally, would mean that we
  // couldn't use this class to handle arbitrary binary data (which might
  // include zeroes), we are required to explicitly enable it in the classes
  // where we need it (e.g. all TCustomTextClipboardFormat descedants).
  if (FTrimZeroes) then
    SetLength(FData, Length(PChar(FData)));

  Result := True;
end;

function TCustomStringClipboardFormat.WriteData(Value: pointer;
  Size: integer): boolean;
begin
  // Transfer string including terminating zero if requested.
  Result := (Size <= Length(FData)+1);
  if (Result) then
    Move(PChar(FData)^, Value^, Size);
end;

function TCustomStringClipboardFormat.GetSize: integer;
begin
  Result := Length(FData);
end;

function TCustomStringClipboardFormat.GetString: string;
begin
  Result := FData;
end;

procedure TCustomStringClipboardFormat.SetString(const Value: string);
begin
  FData := Value;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TCustomStringListClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
constructor TCustomStringListClipboardFormat.Create;
begin
  inherited Create;
  FLines := TStringList.Create
end;

destructor TCustomStringListClipboardFormat.Destroy;
begin
  FLines.Free;
  inherited Destroy;
end;

procedure TCustomStringListClipboardFormat.Clear;
begin
  FLines.Clear;
end;

function TCustomStringListClipboardFormat.HasData: boolean;
begin
  Result := (FLines.Count > 0);
end;

function TCustomStringListClipboardFormat.ReadData(Value: pointer;
  Size: integer): boolean;
var
  s			: string;
begin
  SetLength(s, Size+1);
  Move(Value^, PChar(s)^, Size);
  s[Size] := #0;
  FLines.Text := s;
  Result := True;
end;

function TCustomStringListClipboardFormat.WriteData(Value: pointer;
  Size: integer): boolean;
var
  s			: string;
begin
  s := FLines.Text;
  Result := (Size = Length(s)+1);
  if (Result) then
    Move(PChar(s)^, Value^, Size);
end;

function TCustomStringListClipboardFormat.GetSize: integer;
begin
  Result := Length(FLines.Text)+1;
end;

function TCustomStringListClipboardFormat.GetLines: TStrings;
begin
  Result := FLines;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TCustomTextClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
constructor TCustomTextClipboardFormat.Create;
begin
  inherited Create;
  TrimZeroes := True;
end;

function TCustomTextClipboardFormat.GetSize: integer;
begin
  Result := inherited GetSize;
  // Unless the data is already zero terminated, we add a byte to include
  // the string's implicit terminating zero.
  if (Data[Result] <> #0) then
    inc(Result);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TCustomWideTextClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
procedure TCustomWideTextClipboardFormat.Clear;
begin
  FText := '';
end;

function TCustomWideTextClipboardFormat.HasData: boolean;
begin
  Result := (FText <> '');
end;

function TCustomWideTextClipboardFormat.ReadData(Value: pointer;
  Size: integer): boolean;
begin
  SetLength(FText, Size div 2);
  Move(Value^, PWideChar(FText)^, Size);
  Result := True;
end;

function TCustomWideTextClipboardFormat.WriteData(Value: pointer;
  Size: integer): boolean;
begin
  Result := (Size <= (Length(FText)+1)*2);
  if (Result) then
    Move(PWideChar(FText)^, Value^, Size);
end;

function TCustomWideTextClipboardFormat.GetSize: integer;
begin
  Result := Length(FText)*2;
  // Unless the data is already zero terminated, we add two bytes to include
  // the string's implicit terminating zero.
  if (FText[Result] <> #0) then
    inc(Result, 2);
end;

function TCustomWideTextClipboardFormat.GetText: WideString;
begin
  Result := FText;
end;

procedure TCustomWideTextClipboardFormat.SetText(const Value: WideString);
begin
  FText := Value;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TTextClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
function TTextClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  Result := CF_TEXT;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TCustomDWORDClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
function TCustomDWORDClipboardFormat.ReadData(Value: pointer;
  Size: integer): boolean;
begin
  FValue := PDWORD(Value)^;
  Result := True;
end;

function TCustomDWORDClipboardFormat.WriteData(Value: pointer;
  Size: integer): boolean;
begin
  Result := (Size = SizeOf(DWORD));
  if (Result) then
    PDWORD(Value)^ := FValue;
end;

function TCustomDWORDClipboardFormat.GetSize: integer;
begin
  Result := SizeOf(DWORD);
end;

procedure TCustomDWORDClipboardFormat.Clear;
begin
  FValue := 0;
end;

function TCustomDWORDClipboardFormat.GetValueDWORD: DWORD;
begin
  Result := FValue;
end;

procedure TCustomDWORDClipboardFormat.SetValueDWORD(Value: DWORD);
begin
  FValue := Value;
end;

function TCustomDWORDClipboardFormat.GetValueInteger: integer;
begin
  Result := integer(FValue);
end;

procedure TCustomDWORDClipboardFormat.SetValueInteger(Value: integer);
begin
  FValue := DWORD(Value);
end;

function TCustomDWORDClipboardFormat.GetValueLongInt: longInt;
begin
  Result := longInt(FValue);
end;

procedure TCustomDWORDClipboardFormat.SetValueLongInt(Value: longInt);
begin
  FValue := DWORD(Value);
end;

function TCustomDWORDClipboardFormat.GetValueBoolean: boolean;
begin
  Result := (FValue <> 0);
end;

procedure TCustomDWORDClipboardFormat.SetValueBoolean(Value: boolean);
begin
  FValue := ord(Value);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TFileGroupDescritorClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
var
  CF_FILEGROUPDESCRIPTOR: TClipFormat = 0;

function TFileGroupDescritorClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  if (CF_FILEGROUPDESCRIPTOR = 0) then
    CF_FILEGROUPDESCRIPTOR := RegisterClipboardFormat(CFSTR_FILEDESCRIPTOR);
  Result := CF_FILEGROUPDESCRIPTOR;
end;

destructor TFileGroupDescritorClipboardFormat.Destroy;
begin
  Clear;
  inherited Destroy;
end;

procedure TFileGroupDescritorClipboardFormat.Clear;
begin
  if (FFileGroupDescriptor <> nil) then
  begin
    FreeMem(FFileGroupDescriptor);
    FFileGroupDescriptor := nil;
  end;
end;

function TFileGroupDescritorClipboardFormat.HasData: boolean;
begin
  Result := (FFileGroupDescriptor <> nil) and (FFileGroupDescriptor^.cItems <> 0);
end;

procedure TFileGroupDescritorClipboardFormat.CopyFrom(AFileGroupDescriptor: PFileGroupDescriptor);
var
  Size			: integer;
begin
  Clear;
  if (AFileGroupDescriptor <> nil) then
  begin
    Size := SizeOf(UINT) + AFileGroupDescriptor^.cItems * SizeOf(TFileDescriptor);
    GetMem(FFileGroupDescriptor, Size);
    Move(AFileGroupDescriptor^, FFileGroupDescriptor^, Size);
  end;
end;

function TFileGroupDescritorClipboardFormat.GetSize: integer;
begin
  if (FFileGroupDescriptor <> nil) then
    Result := SizeOf(UINT) + FFileGroupDescriptor^.cItems * SizeOf(TFileDescriptor)
  else
    Result := 0;
end;

function TFileGroupDescritorClipboardFormat.ReadData(Value: pointer;
  Size: integer): boolean;
begin
  // Validate size against count
  Result :=
    (Size - SizeOf(UINT)) DIV SizeOf(TFileDescriptor) = integer(PFileGroupDescriptor(Value)^.cItems);
  if (Result) then
    CopyFrom(PFileGroupDescriptor(Value));
end;

function TFileGroupDescritorClipboardFormat.WriteData(Value: pointer;
  Size: integer): boolean;
begin
  // Validate size against count
  Result := (FFileGroupDescriptor <> nil) and
    ((Size - SizeOf(UINT)) DIV SizeOf(TFileDescriptor) = integer(FFileGroupDescriptor^.cItems));

  if (Result) then
    Move(FFileGroupDescriptor^, Value^, Size);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TFileGroupDescritorWClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
var
  CF_FILEGROUPDESCRIPTORW: TClipFormat = 0;

function TFileGroupDescritorWClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  if (CF_FILEGROUPDESCRIPTORW = 0) then
    CF_FILEGROUPDESCRIPTORW := RegisterClipboardFormat(CFSTR_FILEDESCRIPTORW);
  Result := CF_FILEGROUPDESCRIPTORW;
end;

destructor TFileGroupDescritorWClipboardFormat.Destroy;
begin
  Clear;
  inherited Destroy;
end;

procedure TFileGroupDescritorWClipboardFormat.Clear;
begin
  if (FFileGroupDescriptor <> nil) then
  begin
    FreeMem(FFileGroupDescriptor);
    FFileGroupDescriptor := nil;
  end;
end;

function TFileGroupDescritorWClipboardFormat.HasData: boolean;
begin
  Result := (FFileGroupDescriptor <> nil) and (FFileGroupDescriptor^.cItems <> 0);
end;

procedure TFileGroupDescritorWClipboardFormat.CopyFrom(AFileGroupDescriptor: PFileGroupDescriptorW);
var
  Size			: integer;
begin
  Clear;
  if (AFileGroupDescriptor <> nil) then
  begin
    Size := SizeOf(UINT) + AFileGroupDescriptor^.cItems * SizeOf(TFileDescriptorW);
    GetMem(FFileGroupDescriptor, Size);
    Move(AFileGroupDescriptor^, FFileGroupDescriptor^, Size);
  end;
end;

function TFileGroupDescritorWClipboardFormat.GetSize: integer;
begin
  if (FFileGroupDescriptor <> nil) then
    Result := SizeOf(UINT) + FFileGroupDescriptor^.cItems * SizeOf(TFileDescriptorW)
  else
    Result := 0;
end;

function TFileGroupDescritorWClipboardFormat.ReadData(Value: pointer;
  Size: integer): boolean;
begin
  // Validate size against count
  Result :=
    (Size - SizeOf(UINT)) DIV SizeOf(TFileDescriptorW) = integer(PFileGroupDescriptor(Value)^.cItems);
  if (Result) then
    CopyFrom(PFileGroupDescriptorW(Value));
end;

function TFileGroupDescritorWClipboardFormat.WriteData(Value: pointer;
  Size: integer): boolean;
begin
  // Validate size against count
  Result := (FFileGroupDescriptor <> nil) and
    ((Size - SizeOf(UINT)) DIV SizeOf(TFileDescriptorW) = integer(FFileGroupDescriptor^.cItems));

  if (Result) then
    Move(FFileGroupDescriptor^, Value^, Size);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TFileContentsClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
var
  CF_FILECONTENTS: TClipFormat = 0;

constructor TFileContentsClipboardFormat.Create;
begin
  inherited Create;
  FFormatEtc.lindex := 0;
end;

function TFileContentsClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  if (CF_FILECONTENTS = 0) then
    CF_FILECONTENTS := RegisterClipboardFormat(CFSTR_FILECONTENTS);
  Result := CF_FILECONTENTS;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TFileContentsStreamClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
constructor TFileContentsStreamClipboardFormat.Create;
begin
  CreateFormat(TYMED_ISTREAM);
  FStreams := TStreamList.Create;
end;

destructor TFileContentsStreamClipboardFormat.Destroy;
begin
  Clear;
  FStreams.Free;
  inherited Destroy;
end;

function TFileContentsStreamClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  if (CF_FILECONTENTS = 0) then
    CF_FILECONTENTS := RegisterClipboardFormat(CFSTR_FILECONTENTS);
  Result := CF_FILECONTENTS;
end;

procedure TFileContentsStreamClipboardFormat.Clear;
begin
  FStreams.Clear;
end;

function TFileContentsStreamClipboardFormat.HasData: boolean;
begin
  Result := (FStreams.Count > 0);
end;

function TFileContentsStreamClipboardFormat.AssignTo(Dest: TCustomDataFormat): boolean;
begin
  Result := True;
  if (Dest is TDataStreamDataFormat) then
  begin
    TDataStreamDataFormat(Dest).Streams.Assign(Streams);
  end else
    Result := inherited AssignTo(Dest);
end;

{$IFOPT R+}
  {$DEFINE R_PLUS}
  {$RANGECHECKS OFF}
{$ENDIF}
function TFileContentsStreamClipboardFormat.GetData(DataObject: IDataObject): boolean;
var
  FGD: TFileGroupDescritorClipboardFormat;
  Count: integer;
  Medium: TStgMedium;
  Stream: IStream;
  Name: string;
  MemStream: TMemoryStream;
  StatStg: TStatStg;
  Size: longInt;
  Remaining: longInt;
  pChunk: PChar;
begin
  Result := False;

  Clear;
  FGD := TFileGroupDescritorClipboardFormat.Create;
  try
    if (FGD.GetData(DataObject)) then
    begin
      // Multiple objects, retrieve one at a time
      Count := FGD.FileGroupDescriptor^.cItems;
      FFormatEtc.lindex := 0;
    end else
    begin
      // Single object, retrieve "all" at once
      Count := 0;
      FFormatEtc.lindex := -1;
      Name := '';
    end;
    while (FFormatEtc.lindex < Count) do
    begin
      if (DataObject.GetData(FormatEtc, Medium) <> S_OK) then
        break;
      try
        inc(FFormatEtc.lindex);
        if (Medium.tymed <> TYMED_ISTREAM) then
          continue;
        Stream := IStream(Medium.stm);
        Stream.Stat(StatStg, STATFLAG_NONAME);
        MemStream := TMemoryStream.Create;
        try
          Remaining := StatStg.cbSize;
          MemStream.Size := Remaining;
          pChunk := MemStream.Memory;
          while (Remaining > 0) do
          begin
            if (Stream.Read(pChunk, Remaining, @Size) <> S_OK) or
              (Size = 0) then
              break;
            inc(pChunk, Size);
            dec(Remaining, Size);
          end;

          if (FFormatEtc.lindex > 0) then
            Name := FGD.FileGroupDescriptor^.fgd[FFormatEtc.lindex-1].cFileName;
          Streams.AddNamed(MemStream, Name);
        except
          MemStream.Free;
          raise;
        end;
        Stream := nil;
        Result := True;
      finally
        ReleaseStgMedium(Medium);
      end;
    end;
  finally
    FGD.Free;
  end;
end;
{$IFDEF R_PLUS}
  {$RANGECHECKS ON}
  {$UNDEF R_PLUS}
{$ENDIF}


////////////////////////////////////////////////////////////////////////////////
//
//		TFileContentsStreamOnDemandClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
constructor TFileContentsStreamOnDemandClipboardFormat.Create;
begin
  CreateFormat(TYMED_ISTREAM);
end;

destructor TFileContentsStreamOnDemandClipboardFormat.Destroy;
begin
  Clear;
  inherited Destroy;
end;

function TFileContentsStreamOnDemandClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  if (CF_FILECONTENTS = 0) then
    CF_FILECONTENTS := RegisterClipboardFormat(CFSTR_FILECONTENTS);
  Result := CF_FILECONTENTS;
end;

procedure TFileContentsStreamOnDemandClipboardFormat.Clear;
begin
  FGotData := False;
  FDataRequested := False;
end;

function TFileContentsStreamOnDemandClipboardFormat.HasData: boolean;
begin
  Result := FGotData or FDataRequested;
end;

function TFileContentsStreamOnDemandClipboardFormat.AssignTo(Dest: TCustomDataFormat): boolean;
begin
  if (Dest is TVirtualFileStreamDataFormat) then
  begin
    Result := True
  end else
    Result := inherited AssignTo(Dest);
end;

function TFileContentsStreamOnDemandClipboardFormat.Assign(
  Source: TCustomDataFormat): boolean;
begin
  if (Source is TVirtualFileStreamDataFormat) then
  begin
    // Acknowledge that we can offer the requested data, but defer the actual
    // data transfer.
    FDataRequested := True;
    Result := True
  end else
    Result := inherited Assign(Source);
end;

function TFileContentsStreamOnDemandClipboardFormat.DoSetData(
  const FormatEtcIn: TFormatEtc; var AMedium: TStgMedium): boolean;
var
  Stream		: IStream;
begin
  if (Assigned(FOnGetStream)) and (FormatEtcIn.tymed and TYMED_ISTREAM <> 0) and
    (FormatEtcIn.lindex <> -1) then
  begin
    FOnGetStream(Self, FormatEtcIn.lindex, Stream);

    if (Stream <> nil) then
    begin
      IStream(AMedium.stm) := Stream;
      AMedium.tymed := TYMED_ISTREAM;
      Result := True;
    end else
      Result := False;

  end else
    Result := False;
end;

function TFileContentsStreamOnDemandClipboardFormat.GetData(DataObject: IDataObject): boolean;
begin
  // Flag that data has been offered to us, but defer the actual data transfer.
  FGotData := True;
  Result := True;
end;

function TFileContentsStreamOnDemandClipboardFormat.GetStream(Index: integer): IStream;
var
  Medium		: TStgMedium;
begin
  Result := nil;
  FFormatEtc.lindex := Index;
  // Get an IStream interface from the source.
  if ((DataFormat.Owner as TCustomDroptarget).DataObject.GetData(FormatEtc,
    Medium) = S_OK) and (Medium.tymed = TYMED_ISTREAM) then
    try
      Result := IStream(Medium.stm);
    finally
      ReleaseStgMedium(Medium);
    end;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TFileContentsStorageClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
constructor TFileContentsStorageClipboardFormat.Create;
begin
  CreateFormat(TYMED_ISTORAGE);
  FStorages := TStorageInterfaceList.Create;
end;

destructor TFileContentsStorageClipboardFormat.Destroy;
begin
  Clear;
  FStorages.Free;
  inherited Destroy;
end;

function TFileContentsStorageClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  if (CF_FILECONTENTS = 0) then
    CF_FILECONTENTS := RegisterClipboardFormat(CFSTR_FILECONTENTS);
  Result := CF_FILECONTENTS;
end;

procedure TFileContentsStorageClipboardFormat.Clear;
begin
  FStorages.Clear;
end;

function TFileContentsStorageClipboardFormat.HasData: boolean;
begin
  Result := (FStorages.Count > 0);
end;

function TFileContentsStorageClipboardFormat.AssignTo(Dest: TCustomDataFormat): boolean;
begin
(*
  Result := True;
  if (Dest is TDataStreamDataFormat) then
  begin
    TDataStreamDataFormat(Dest).Streams.Assign(Streams);
  end else
*)
    Result := inherited AssignTo(Dest);
end;

{$IFOPT R+}
  {$DEFINE R_PLUS}
  {$RANGECHECKS OFF}
{$ENDIF}
function TFileContentsStorageClipboardFormat.GetData(DataObject: IDataObject): boolean;
var
  FGD			: TFileGroupDescritorClipboardFormat;
  Count			: integer;
  Medium		: TStgMedium;
  Storage		: IStorage;
  Name			: string;
begin
  Result := False;

  Clear;
  FGD := TFileGroupDescritorClipboardFormat.Create;
  try
    if (FGD.GetData(DataObject)) then
    begin
      // Multiple objects, retrieve one at a time
      Count := FGD.FileGroupDescriptor^.cItems;
      FFormatEtc.lindex := 0;
    end else
    begin
      // Single object, retrieve "all" at once
      Count := 0;
      FFormatEtc.lindex := -1;
      Name := '';
    end;
    while (FFormatEtc.lindex < Count) do
    begin
      if (DataObject.GetData(FormatEtc, Medium) <> S_OK) then
        break;
      try
        inc(FFormatEtc.lindex);
        if (Medium.tymed <> TYMED_ISTORAGE) then
          continue;
        Storage := IStorage(Medium.stg);
        if (FFormatEtc.lindex > 0) then
          Name := FGD.FileGroupDescriptor^.fgd[FFormatEtc.lindex-1].cFileName;
        Storages.AddNamed(Storage, Name);
        Storage := nil;
        Result := True;
      finally
        ReleaseStgMedium(Medium);
      end;
    end;
  finally
    FGD.Free;
  end;
end;
{$IFDEF R_PLUS}
  {$RANGECHECKS ON}
  {$UNDEF R_PLUS}
{$ENDIF}


////////////////////////////////////////////////////////////////////////////////
//
//		TPreferredDropEffectClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
var
  CF_PREFERREDDROPEFFECT: TClipFormat = 0;

// GetClassClipboardFormat is used by TCustomDropTarget.GetPreferredDropEffect 
class function TPreferredDropEffectClipboardFormat.GetClassClipboardFormat: TClipFormat;
begin
  if (CF_PREFERREDDROPEFFECT = 0) then
    CF_PREFERREDDROPEFFECT := RegisterClipboardFormat(CFSTR_PREFERREDDROPEFFECT);
  Result := CF_PREFERREDDROPEFFECT;
end;

function TPreferredDropEffectClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  Result := GetClassClipboardFormat;
end;

function TPreferredDropEffectClipboardFormat.HasData: boolean;
begin
  Result := True; //(Value <> DROPEFFECT_NONE);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TPerformedDropEffectClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
var
  CF_PERFORMEDDROPEFFECT: TClipFormat = 0;

function TPerformedDropEffectClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  if (CF_PERFORMEDDROPEFFECT = 0) then
    CF_PERFORMEDDROPEFFECT := RegisterClipboardFormat(CFSTR_PERFORMEDDROPEFFECT);
  Result := CF_PERFORMEDDROPEFFECT;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TLogicalPerformedDropEffectClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
var
  CF_LOGICALPERFORMEDDROPEFFECT: TClipFormat = 0;

function TLogicalPerformedDropEffectClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  if (CF_LOGICALPERFORMEDDROPEFFECT = 0) then
    CF_LOGICALPERFORMEDDROPEFFECT := RegisterClipboardFormat('Logical Performed DropEffect'); // *** DO NOT LOCALIZE ***
  Result := CF_LOGICALPERFORMEDDROPEFFECT;
end;



////////////////////////////////////////////////////////////////////////////////
//
//		TPasteSuccededClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
var
  CF_PASTESUCCEEDED: TClipFormat = 0;

function TPasteSuccededClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  if (CF_PASTESUCCEEDED = 0) then
    CF_PASTESUCCEEDED := RegisterClipboardFormat(CFSTR_PASTESUCCEEDED);
  Result := CF_PASTESUCCEEDED;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TInShellDragLoopClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
var
  CF_InDragLoop: TClipFormat = 0;

function TInShellDragLoopClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  if (CF_InDragLoop = 0) then
    CF_InDragLoop := RegisterClipboardFormat(CFSTR_InDragLoop);
  Result := CF_InDragLoop;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TTargetCLSIDClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
procedure TTargetCLSIDClipboardFormat.Clear;
begin
  FCLSID := GUID_NULL;
end;

var
  CF_TargetCLSID: TClipFormat = 0;

function TTargetCLSIDClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  if (CF_TargetCLSID = 0) then
    CF_TargetCLSID := RegisterClipboardFormat('TargetCLSID'); // *** DO NOT LOCALIZE ***
  Result := CF_TargetCLSID;
end;

function TTargetCLSIDClipboardFormat.GetSize: integer;
begin
  Result := SizeOf(TCLSID);
end;

function TTargetCLSIDClipboardFormat.HasData: boolean;
begin
  Result := not IsEqualCLSID(FCLSID, GUID_NULL);
end;

function TTargetCLSIDClipboardFormat.ReadData(Value: pointer;
  Size: integer): boolean;
begin
  // Validate size.
  Result := (Size = SizeOf(TCLSID));
  if (Result) then
    FCLSID := PCLSID(Value)^;
end;

function TTargetCLSIDClipboardFormat.WriteData(Value: pointer;
  Size: integer): boolean;
begin
  // Validate size.
  Result := (Size = SizeOf(TCLSID));
  if (Result) then
    PCLSID(Value)^ := FCLSID;
end;

////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////


////////////////////////////////////////////////////////////////////////////////
//
//		TTextDataFormat
//
////////////////////////////////////////////////////////////////////////////////
function TTextDataFormat.Assign(Source: TClipboardFormat): boolean;
begin
  Result := True;

  if (Source is TTextClipboardFormat) then
    FText := TTextClipboardFormat(Source).Text
  else if (Source is TFileContentsClipboardFormat) then
    FText := TFileContentsClipboardFormat(Source).Data
  else
    Result := inherited Assign(Source);
end;

function TTextDataFormat.AssignTo(Dest: TClipboardFormat): boolean;
var
  FGD: TFileGroupDescriptor;
  FGDW: TFileGroupDescriptorW;
resourcestring
  // Name of the text scrap file.
  sTextScrap = 'Text scrap.txt';
begin
  Result := True;

  if (Dest is TTextClipboardFormat) then
    TTextClipboardFormat(Dest).Text := FText
  else if (Dest is TFileContentsClipboardFormat) then
    TFileContentsClipboardFormat(Dest).Data := FText
  else if (Dest is TFileGroupDescritorClipboardFormat) then
  begin
    FillChar(FGD, SizeOf(FGD), 0);
    FGD.cItems := 1;
    StrPLCopy(FGD.fgd[0].cFileName, sTextScrap, SizeOf(FGD.fgd[0].cFileName));
    TFileGroupDescritorClipboardFormat(Dest).CopyFrom(@FGD);
  end else
  if (Dest is TFileGroupDescritorWClipboardFormat) then
  begin
    FillChar(FGDW, SizeOf(FGDW), 0);
    FGDW.cItems := 1;
    StringToWideChar(sTextScrap, PWideChar(@(FGDW.fgd[0].cFileName)), MAX_PATH);
    TFileGroupDescritorWClipboardFormat(Dest).CopyFrom(@FGDW);
  end else
    Result := inherited AssignTo(Dest);
end;

procedure TTextDataFormat.Clear;
begin
  Changing;
  FText := '';
end;

procedure TTextDataFormat.SetText(const Value: string);
begin
  Changing;
  FText := Value;
end;

function TTextDataFormat.HasData: boolean;
begin
  Result := (FText <> '');
end;

function TTextDataFormat.NeedsData: boolean;
begin
  Result := (FText = '');
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TDataStreamDataFormat
//
////////////////////////////////////////////////////////////////////////////////
constructor TDataStreamDataFormat.Create(AOwner: TDragDropComponent);
begin
  inherited Create(AOwner);
  FStreams := TStreamList.Create;
  FStreams.OnChanging := DoOnChanging;
end;

destructor TDataStreamDataFormat.Destroy;
begin
  Clear;
  FStreams.Free;
  inherited Destroy;
end;

procedure TDataStreamDataFormat.Clear;
begin
  Changing;
  FStreams.Clear;
end;

function TDataStreamDataFormat.HasData: boolean;
begin
  Result := (Streams.Count > 0);
end;

function TDataStreamDataFormat.NeedsData: boolean;
begin
  Result := (Streams.Count = 0);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TFileDescriptorToFilenameStrings
//
////////////////////////////////////////////////////////////////////////////////
// Used internally to convert between FileDescriptors and filenames on-demand.
////////////////////////////////////////////////////////////////////////////////
type
  TFileDescriptorToFilenameStrings = class(TStrings)
  private
    FFileDescriptors: TMemoryList;
  protected
    function Get(Index: Integer): string; override;
    function GetCount: Integer; override;
  public
    constructor Create(AFileDescriptors: TMemoryList);
    procedure Clear; override;
    procedure Delete(Index: Integer); override;
    procedure Insert(Index: Integer; const S: string); override;
    procedure Assign(Source: TPersistent); override;
  end;

constructor TFileDescriptorToFilenameStrings.Create(AFileDescriptors: TMemoryList);
begin
  inherited Create;
  FFileDescriptors := AFileDescriptors;
end;

function TFileDescriptorToFilenameStrings.Get(Index: Integer): string;
begin
  Result := PFileDescriptor(FFileDescriptors[Index]).cFileName;
end;

function TFileDescriptorToFilenameStrings.GetCount: Integer;
begin
  Result := FFileDescriptors.Count;
end;

procedure TFileDescriptorToFilenameStrings.Assign(Source: TPersistent);
var
  i: integer;
begin
  if Source is TStrings then
  begin
    BeginUpdate;
    try
      FFileDescriptors.Clear;
      for i := 0 to TStrings(Source).Count-1 do
        Add(TStrings(Source)[i]);
    finally
      EndUpdate;
    end;
  end else
    inherited Assign(Source);
end;

procedure TFileDescriptorToFilenameStrings.Clear;
begin
  FFileDescriptors.Clear;
end;

procedure TFileDescriptorToFilenameStrings.Delete(Index: Integer);
begin
  FFileDescriptors.Delete(Index);
end;

procedure TFileDescriptorToFilenameStrings.Insert(Index: Integer; const S: string);
var
  FD: PFileDescriptor;
begin
  if (Index = FFileDescriptors.Count) then
  begin
    GetMem(FD, SizeOf(TFileDescriptor));
    try
      FillChar(FD^, SizeOf(TFileDescriptor), 0);
      StrPLCopy(FD.cFileName, S, SizeOf(FD.cFileName));
      FFileDescriptors.Add(FD);
    except
      FreeMem(FD);
      raise;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
//
//		TVirtualFileStreamDataFormat
//
////////////////////////////////////////////////////////////////////////////////
constructor TVirtualFileStreamDataFormat.Create(AOwner: TDragDropComponent);
begin
  inherited Create(AOwner);
  FFileDescriptors := TMemoryList.Create;
  FFileNames := TFileDescriptorToFilenameStrings.Create(FFileDescriptors);

  // Add the "file group descriptor" and "file contents" clipboard formats to
  // the data format's list of compatible formats.
  // Note: This is normally done via TCustomDataFormat.RegisterCompatibleFormat,
  // but since this data format and the clipboard format class are specialized
  // to be used with each other, it is just as easy for us to add the formats
  // manually.
  FFileContentsClipboardFormat := TFileContentsStreamOnDemandClipboardFormat.Create;
  CompatibleFormats.Add(FFileContentsClipboardFormat);

  FFileGroupDescritorClipboardFormat := TFileGroupDescritorClipboardFormat.Create;

  // Normaly TFileGroupDescritorClipboardFormat supports both HGlobal and
  // IStream storage medium transfers, but for this demo we only use IStream.
//  FFileGroupDescritorClipboardFormat.FormatEtc.tymed := TYMED_ISTREAM;

  CompatibleFormats.Add(FFileGroupDescritorClipboardFormat);
end;

destructor TVirtualFileStreamDataFormat.Destroy;
begin
  FFileDescriptors.Free;
  FFileNames.Free;
  inherited Destroy;
end;

procedure TVirtualFileStreamDataFormat.SetFileNames(const Value: TStrings);
begin
  FFileNames.Assign(Value);
end;

{$IFOPT R+}
  {$DEFINE R_PLUS}
  {$RANGECHECKS OFF}
{$ENDIF}
function TVirtualFileStreamDataFormat.Assign(Source: TClipboardFormat): boolean;
var
  i: integer;
  FD: PFileDescriptor;
begin
  Result := True;

  (*
  ** TFileContentsStreamOnDemandClipboardFormat
  *)
  if (Source is TFileContentsStreamOnDemandClipboardFormat) then
  begin
    FHasContents := TFileContentsStreamOnDemandClipboardFormat(Source).HasData;
  end else
  (*
  ** TFileGroupDescritorClipboardFormat
  *)
  if (Source is TFileGroupDescritorClipboardFormat) then
  begin
    FFileDescriptors.Clear;
    for i := 0 to TFileGroupDescritorClipboardFormat(Source).FileGroupDescriptor^.cItems-1 do
    begin
      GetMem(FD, SizeOf(TFileDescriptor));
      try
        Move(TFileGroupDescritorClipboardFormat(Source).FileGroupDescriptor^.fgd[i],
          FD^, SizeOf(TFileDescriptor));
        FFileDescriptors.Add(FD);
      except
        FreeMem(FD);
        raise;
      end;
    end;
  end else
  (*
  ** None of the above...
  *)
    Result := inherited Assign(Source);
end;
{$IFDEF R_PLUS}
  {$RANGECHECKS ON}
  {$UNDEF R_PLUS}
{$ENDIF}

{$IFOPT R+}
  {$DEFINE R_PLUS}
  {$RANGECHECKS OFF}
{$ENDIF}
function TVirtualFileStreamDataFormat.AssignTo(Dest: TClipboardFormat): boolean;
var
  FGD: PFileGroupDescriptor;
  i: integer;
begin
  (*
  ** TFileContentsStreamOnDemandClipboardFormat
  *)
  if (Dest is TFileContentsStreamOnDemandClipboardFormat) then
  begin
    // Let the clipboard format handle the transfer.
    // No data is actually transferred, but TFileContentsStreamOnDemandClipboardFormat
    // needs to set a flag when data is requested.
    Result := Dest.Assign(Self);
  end else
  (*
  ** TFileGroupDescritorClipboardFormat
  *)
  if (Dest is TFileGroupDescritorClipboardFormat) then
  begin
    if (FFileDescriptors.Count > 0) then
    begin
      GetMem(FGD, SizeOf(UINT) + FFileDescriptors.Count * SizeOf(TFileDescriptor));
      try
        FGD.cItems := FFileDescriptors.Count;
        for i := 0 to FFileDescriptors.Count-1 do
          Move(FFileDescriptors[i]^, FGD.fgd[i], SizeOf(TFileDescriptor));

        TFileGroupDescritorClipboardFormat(Dest).CopyFrom(FGD);
      finally
        FreeMem(FGD);
      end;
      Result := True;
    end else
      Result := False;
  end else
  (*
  ** None of the above...
  *)
    Result := inherited AssignTo(Dest);
end;
{$IFDEF R_PLUS}
  {$RANGECHECKS ON}
  {$UNDEF R_PLUS}
{$ENDIF}

procedure TVirtualFileStreamDataFormat.Clear;
begin
  FFileDescriptors.Clear;
  FHasContents := False;
end;

function TVirtualFileStreamDataFormat.HasData: boolean;
begin
  Result := (FFileDescriptors.Count > 0) and
    ((FHasContents) or Assigned(FFileContentsClipboardFormat.OnGetStream));
end;

function TVirtualFileStreamDataFormat.NeedsData: boolean;
begin
  Result := (FFileDescriptors.Count = 0) or (not FHasContents);
end;

function TVirtualFileStreamDataFormat.GetOnGetStream: TOnGetStreamEvent;
begin
  Result := FFileContentsClipboardFormat.OnGetStream;
end;

procedure TVirtualFileStreamDataFormat.SetOnGetStream(const Value: TOnGetStreamEvent);
begin
  FFileContentsClipboardFormat.OnGetStream := Value;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TFeedbackDataFormat
//
////////////////////////////////////////////////////////////////////////////////
function TFeedbackDataFormat.Assign(Source: TClipboardFormat): boolean;
begin
  Result := True;

  if (Source is TPreferredDropEffectClipboardFormat) then
    FPreferredDropEffect := TPreferredDropEffectClipboardFormat(Source).Value

  else if (Source is TPerformedDropEffectClipboardFormat) then
    FPerformedDropEffect := TPerformedDropEffectClipboardFormat(Source).Value

  else if (Source is TLogicalPerformedDropEffectClipboardFormat) then
    FLogicalPerformedDropEffect := TLogicalPerformedDropEffectClipboardFormat(Source).Value

  else if (Source is TPasteSuccededClipboardFormat) then
    FPasteSucceded := TPasteSuccededClipboardFormat(Source).Value

  else if (Source is TTargetCLSIDClipboardFormat) then
    FTargetCLSID := TTargetCLSIDClipboardFormat(Source).CLSID

  else if (Source is TInShellDragLoopClipboardFormat) then
  begin
    FInShellDragLoop := TInShellDragLoopClipboardFormat(Source).InShellDragLoop;
    FGotInShellDragLoop := True;
  end else
    Result := inherited Assign(Source);
end;

function TFeedbackDataFormat.AssignTo(Dest: TClipboardFormat): boolean;
begin
  Result := True;

  if (Dest is TPreferredDropEffectClipboardFormat) then
    TPreferredDropEffectClipboardFormat(Dest).Value := FPreferredDropEffect

  else if (Dest is TPerformedDropEffectClipboardFormat) then
    TPerformedDropEffectClipboardFormat(Dest).Value := FPerformedDropEffect

  else if (Dest is TLogicalPerformedDropEffectClipboardFormat) then
    TLogicalPerformedDropEffectClipboardFormat(Dest).Value := FLogicalPerformedDropEffect

  else if (Dest is TPasteSuccededClipboardFormat) then
    TPasteSuccededClipboardFormat(Dest).Value := FPasteSucceded

  else if (Dest is TTargetCLSIDClipboardFormat) then
    TTargetCLSIDClipboardFormat(Dest).CLSID := FTargetCLSID

  else if (Dest is TInShellDragLoopClipboardFormat) then
    TInShellDragLoopClipboardFormat(Dest).InShellDragLoop := FInShellDragLoop

  else
    Result := inherited AssignTo(Dest);
end;

procedure TFeedbackDataFormat.Clear;
begin
  Changing;
  FPreferredDropEffect := DROPEFFECT_NONE;
  FPerformedDropEffect := DROPEFFECT_NONE;
  FInShellDragLoop := False;
  FGotInShellDragLoop := False;
end;

procedure TFeedbackDataFormat.SetInShellDragLoop(const Value: boolean);
begin
  Changing;
  FInShellDragLoop := Value;
end;

procedure TFeedbackDataFormat.SetPasteSucceded(const Value: longInt);
begin
  Changing;
  FPasteSucceded := Value;
end;

procedure TFeedbackDataFormat.SetPerformedDropEffect(
  const Value: longInt);
begin
  Changing;
  FPerformedDropEffect := Value;
end;

procedure TFeedbackDataFormat.SetLogicalPerformedDropEffect(
  const Value: longInt);
begin
  Changing;
  FLogicalPerformedDropEffect := Value;
end;

procedure TFeedbackDataFormat.SetPreferredDropEffect(
  const Value: longInt);
begin
  Changing;
  FPreferredDropEffect := Value;
end;

procedure TFeedbackDataFormat.SetTargetCLSID(const Value: TCLSID);
begin
  Changing;
  FTargetCLSID := Value;
end;

function TFeedbackDataFormat.HasData: boolean;
begin
  Result := (FPreferredDropEffect <> DROPEFFECT_NONE) or
    (FPerformedDropEffect <> DROPEFFECT_NONE) or
    (FPasteSucceded <> DROPEFFECT_NONE) or
    (FGotInShellDragLoop);
end;

function TFeedbackDataFormat.NeedsData: boolean;
begin
  Result := (FPreferredDropEffect = DROPEFFECT_NONE) or
    (FPerformedDropEffect = DROPEFFECT_NONE) or
    (FPasteSucceded = DROPEFFECT_NONE) or
    (not FGotInShellDragLoop);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TGenericClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
procedure TGenericClipboardFormat.SetClipboardFormatName(const Value: string);
begin
  FFormat := Value;
  if (FFormat <> '') then
    ClipboardFormat := RegisterClipboardFormat(PChar(FFormat));
end;

function TGenericClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  if (FFormatEtc.cfFormat = 0) and (FFormat <> '') then
    FFormatEtc.cfFormat := RegisterClipboardFormat(PChar(FFormat));
  Result := FFormatEtc.cfFormat;
end;

function TGenericClipboardFormat.GetClipboardFormatName: string;
begin
  Result := FFormat;
end;

function TGenericClipboardFormat.Assign(Source: TCustomDataFormat): boolean;
begin
  if (Source is TGenericDataFormat) then
  begin
    Data := TGenericDataFormat(Source).Data;
    Result := True;
  end else
    Result := inherited Assign(Source);
end;

function TGenericClipboardFormat.AssignTo(Dest: TCustomDataFormat): boolean;
begin
  if (Dest is TGenericDataFormat) then
  begin
    TGenericDataFormat(Dest).Data := Data;
    Result := True;
  end else
    Result := inherited AssignTo(Dest);
end;

////////////////////////////////////////////////////////////////////////////////
//
//		TGenericDataFormat
//
////////////////////////////////////////////////////////////////////////////////
procedure TGenericDataFormat.AddFormat(const AFormat: string);
var
  ClipboardFormat: TGenericClipboardFormat;
begin
  ClipboardFormat := TGenericClipboardFormat.Create;
  ClipboardFormat.ClipboardFormatName := AFormat;
  ClipboardFormat.DataDirections := [ddRead];
  CompatibleFormats.Add(ClipboardFormat);
end;

procedure TGenericDataFormat.Clear;
begin
  Changing;
  FData := '';
end;

function TGenericDataFormat.HasData: boolean;
begin
  Result := (FData <> '');
end;

function TGenericDataFormat.NeedsData: boolean;
begin
  Result := (FData = '');
end;

procedure TGenericDataFormat.DoSetData(const Value: string);
begin
  Changing;
  FData := Value;
end;

procedure TGenericDataFormat.SetDataHere(const AData; ASize: integer);
begin
  Changing;
  SetLength(FData, ASize);
  Move(AData, PChar(FData)^, ASize);
end;

function TGenericDataFormat.GetSize: integer;
begin
  Result := length(FData);
end;

function TGenericDataFormat.GetDataHere(var AData; ASize: integer): integer;
begin
  Result := Size;
  if (ASize < Result) then
    Result := ASize;
  Move(PChar(FData)^, AData, Result);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		Initialization/Finalization
//
////////////////////////////////////////////////////////////////////////////////

initialization
  // Data format registration
  TTextDataFormat.RegisterDataFormat;
  TDataStreamDataFormat.RegisterDataFormat;
  TVirtualFileStreamDataFormat.RegisterDataFormat;

  // Clipboard format registration
  TTextDataFormat.RegisterCompatibleFormat(TTextClipboardFormat, 0, csSourceTarget, [ddRead]);
  TTextDataFormat.RegisterCompatibleFormat(TFileContentsClipboardFormat, 1, csSourceTarget, [ddRead]);
  TTextDataFormat.RegisterCompatibleFormat(TFileGroupDescritorClipboardFormat, 1, [csSource], [ddRead]);
  TTextDataFormat.RegisterCompatibleFormat(TFileGroupDescritorWClipboardFormat, 1, [csSource], [ddRead]);
  TFeedbackDataFormat.RegisterCompatibleFormat(TPreferredDropEffectClipboardFormat, 0, csSourceTarget, [ddRead]);
  TFeedbackDataFormat.RegisterCompatibleFormat(TPerformedDropEffectClipboardFormat, 0, csSourceTarget, [ddWrite]);
  TFeedbackDataFormat.RegisterCompatibleFormat(TPasteSuccededClipboardFormat, 0, csSourceTarget, [ddWrite]);
  TFeedbackDataFormat.RegisterCompatibleFormat(TInShellDragLoopClipboardFormat, 0, csSourceTarget, [ddRead]);
  TFeedbackDataFormat.RegisterCompatibleFormat(TTargetCLSIDClipboardFormat, 0, csSourceTarget, [ddWrite]);
  TFeedbackDataFormat.RegisterCompatibleFormat(TLogicalPerformedDropEffectClipboardFormat, 0, csSourceTarget, [ddWrite]);
  TDataStreamDataFormat.RegisterCompatibleFormat(TFileContentsStreamClipboardFormat, 0, [csTarget], [ddRead]);

finalization
  TTextDataFormat.UnregisterDataFormat;
  TDataStreamDataFormat.UnregisterDataFormat;
  TFeedbackDataFormat.UnregisterDataFormat;
  TVirtualFileStreamDataFormat.UnregisterDataFormat;

  TTextClipboardFormat.UnregisterClipboardFormat;
  TFileGroupDescritorClipboardFormat.UnregisterClipboardFormat;
  TFileGroupDescritorWClipboardFormat.UnregisterClipboardFormat;
  TFileContentsClipboardFormat.UnregisterClipboardFormat;
  TFileContentsStreamClipboardFormat.UnregisterClipboardFormat;
  TPreferredDropEffectClipboardFormat.UnregisterClipboardFormat;
  TPerformedDropEffectClipboardFormat.UnregisterClipboardFormat;
  TPasteSuccededClipboardFormat.UnregisterClipboardFormat;
  TInShellDragLoopClipboardFormat.UnregisterClipboardFormat;
  TTargetCLSIDClipboardFormat.UnregisterClipboardFormat;
  TLogicalPerformedDropEffectClipboardFormat.UnregisterClipboardFormat;

end.


