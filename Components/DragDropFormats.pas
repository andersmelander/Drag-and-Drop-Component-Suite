unit DragDropFormats;
// -----------------------------------------------------------------------------
// Project:         Drag and Drop Component Suite.
// Module:          DragDropFormats
// Description:     Implements commonly used clipboard formats and base classes.
// Version:         4.1
// Date:            24-APR-2003
// Target:          Win32, Delphi 4-7, C++Builder 4-6
// Authors:         Anders Melander, anders@melander.dk, http://www.melander.dk
// Copyright        © 1997-2003 Angus Johnson & Anders Melander
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
    FStreams: TStrings;
    FOnChanging: TNotifyEvent;
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
//		TNamedInterfaceList
//
////////////////////////////////////////////////////////////////////////////////
// List of named interfaces.
// Note: Delphi 5 also implements a TInterfaceList, but it can not be used
// because it doesn't support change notification and isn't extensible.
////////////////////////////////////////////////////////////////////////////////
// Utility class used by TFileContentsStorageClipboardFormat.
////////////////////////////////////////////////////////////////////////////////
  TNamedInterfaceList = class(TObject)
  private
    FList: TStrings;
    FOnChanging: TNotifyEvent;
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
    procedure Assign(Value: TNamedInterfaceList);
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
// Used by TFileContentsStorageClipboardFormat and TStorageDataFormat.
////////////////////////////////////////////////////////////////////////////////
  TStorageInterfaceList = class(TNamedInterfaceList)
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
  private
    FHasSeeked: boolean;
  public
    function Seek(dlibMove: Largeint; dwOrigin: Longint;
      out libNewPosition: Largeint): HResult; override; stdcall;
    function Read(pv: Pointer; cb: Longint;
      pcbRead: PLongint): HResult; override; stdcall;
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
// DoGetDataSized method. The descendant DoGetDataSized method should allocate a
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
    FLines: TStrings;
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
    FText: WideString;
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
// Note: The system automatically synthesizes the CF_TEXT format from the
// CF_OEMTEXT format. On Windows NT/2000 the system also synthesizes from the
// CF_UNICODETEXT format.
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
    FValue: DWORD;
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

(*
////////////////////////////////////////////////////////////////////////////////
//
//		TFileGroupDescritorCustomClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
  TFileGroupDescritorCustomClipboardFormat = class(TCustomSimpleClipboardFormat)
  private
  protected
  public
    property Filenames[Index: integer]: string read GetFilename write SetFilename;
    property Count: integer read GetCount write SetCount;
  end;
*)

////////////////////////////////////////////////////////////////////////////////
//
//		TFileGroupDescritorClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
const
  // Missing declaration from shlobj.pas (D6 and earlier)
  FD_PROGRESSUI = $4000; // Show Progress UI w/Drag and Drop

type
  TFileGroupDescritorClipboardFormat = class(TCustomSimpleClipboardFormat)
  private
    FFileGroupDescriptor: PFileGroupDescriptor;
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
    FFileGroupDescriptor: PFileGroupDescriptorW;
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
// TODO : Why must it be zero terminated?
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
    FStorages: TStorageInterfaceList;
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
//		TPasteSucceededClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
  TPasteSucceededClipboardFormat = class(TCustomDWORDClipboardFormat)
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
//		TDataStreamDataFormat
//
////////////////////////////////////////////////////////////////////////////////
  TDataStreamDataFormat = class(TCustomDataFormat)
  private
    FStreams: TStreamList;
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
    FPasteSucceeded: longInt;
    FInShellDragLoop: boolean;
    FGotInShellDragLoop: boolean;
    FTargetCLSID: TCLSID;
  protected
    procedure SetInShellDragLoop(const Value: boolean);
    procedure SetPasteSucceeded(const Value: longInt);
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
    property PasteSucceeded: longInt read FPasteSucceeded write SetPasteSucceeded;
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
    FData: string;
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
//
//		Utilities
//
////////////////////////////////////////////////////////////////////////////////

// CreateIStreamFromIStorage stores a copy of an IStorage object on an IStream
// object and returns the IStream object.
// It is the callers resposibility to dispose of the IStream. Any modifications
// made to the IStream does not affect the original IStorage object.
//
// CreateIStreamFromIStorage and the work to integrate it into
// TFileContentsStreamClipboardFormat was funded by ThoughtShare Communications
// Inc.
function CreateIStreamFromIStorage(Storage: IStorage): IStream;
function CreateIStorageOnHGlobal(GlobalSource: HGLOBAL): IStorage;


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
  ComObj,
  AxCtrls,
  SysUtils;

////////////////////////////////////////////////////////////////////////////////
//
//		Utilities
//
////////////////////////////////////////////////////////////////////////////////
function CreateIStreamFromIStorage(Storage: IStorage): IStream;
var
  LockBytes: ILockBytes;
  HGlob: HGLOBAL;
  NewStorage: IStorage;
begin
  HGlob := GlobalAlloc(GMEM_SHARE or GMEM_ZEROINIT, 0);
  try
    // TODO : Check if the above memory is leaked.
    OleCheck(CreateILockBytesOnHGlobal(HGlob, False, LockBytes));
    OleCheck(StgCreateDocfileOnILockBytes(LockBytes,
      STGM_CREATE or STGM_READWRITE or STGM_SHARE_EXCLUSIVE{ or STGM_DIRECT},
      0, NewStorage));
    OleCheck(Storage.CopyTo(0, nil, nil, NewStorage));
    OleCheck(CreateStreamOnHGlobal(HGlob, True, Result));
  except
    // Eat exceptions since they wont work inside drag/drop anyway.
    GlobalFree(HGlob);
    Result := nil;
  end;
end;

function CreateIStorageOnHGlobal(GlobalSource: HGLOBAL): IStorage;
var
  LockBytesSource, LockBytesDest: ILockBytes;
  GlobalDest: HGLOBAL;
  StorageSource: IStorage;
begin
  // Alloc memory for new storage.
  GlobalDest := GlobalAlloc(GMEM_SHARE, 0);
  try
    // Open ILockBytes on allocated memory.
    try
      OleCheck(CreateILockBytesOnHGlobal(GlobalDest, True, LockBytesDest));
    except
      GlobalFree(GlobalDest);
      raise;
    end;
    // Open ILockBytes on source memory.
    OleCheck(CreateILockBytesOnHGlobal(GlobalSource, False, LockBytesSource));
    // Create IStorage on source.
    OleCheck(StgCreateDocfileOnILockBytes(LockBytesSource,
      STGM_CREATE or STGM_READWRITE or STGM_SHARE_EXCLUSIVE{ or STGM_DIRECT},
      0, StorageSource));
    // Create IStorage on dest.
    OleCheck(StgCreateDocfileOnILockBytes(LockBytesDest,
      STGM_CREATE or STGM_READWRITE or STGM_SHARE_EXCLUSIVE{ or STGM_DIRECT},
      0, Result));
    // copy source IStorage to dest IStorage.
    OleCheck(StorageSource.CopyTo(0, nil, nil, Result));
  except
    // Eat exceptions since they wont work inside drag/drop anyway.
    Result := nil;
  end;
end;

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
  i: integer;
begin
  Changing;
  for i := 0 to FStreams.Count-1 do
    if (FStreams.Objects[i] <> nil) then
      FStreams.Objects[i].Free;
  FStreams.Clear;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TNamedInterfaceList
//
////////////////////////////////////////////////////////////////////////////////
constructor TNamedInterfaceList.Create;
begin
  inherited Create;
  FList := TStringList.Create;
end;

destructor TNamedInterfaceList.Destroy;
begin
  Clear;
  FList.Free;
  inherited Destroy;
end;

function TNamedInterfaceList.Add(Item: IUnknown): integer;
begin
  Result := AddNamed(Item, '');
end;

function TNamedInterfaceList.AddNamed(Item: IUnknown; Name: string): integer;
begin
  Changing;
  with FList do
  begin
    Result := AddObject(Name, nil);
    Objects[Result] := TObject(Item);
    Item._AddRef;
  end;
end;

procedure TNamedInterfaceList.Changing;
begin
  if (Assigned(OnChanging)) then
    OnChanging(Self);
end;

procedure TNamedInterfaceList.Clear;
var
  i: Integer;
  p: pointer;
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

procedure TNamedInterfaceList.Assign(Value: TNamedInterfaceList);
var
  i: Integer;
begin
  Changing;
  for i := 0 to Value.Count - 1 do
    AddNamed(Value.Items[i], Value.Names[i]);
end;

procedure TNamedInterfaceList.Delete(Index: integer);
var
  p: pointer;
begin
  Changing;
  with FList do
  begin
    p := Objects[Index];
    IUnknown(p) := nil;
    Delete(Index);
  end;
end;

function TNamedInterfaceList.GetCount: integer;
begin
  Result := FList.Count;
end;

function TNamedInterfaceList.GetName(Index: integer): string;
begin
  Result := FList[Index];
end;

function TNamedInterfaceList.GetItem(Index: integer): IUnknown;
var
  p: pointer;
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
  Clear;
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
function TFixedStreamAdapter.Seek(dlibMove: Largeint; dwOrigin: Integer;
  out libNewPosition: Largeint): HResult;
begin
  Result := inherited Seek(dlibMove, dwOrigin, libNewPosition);
  FHasSeeked := True;
end;

function TFixedStreamAdapter.Read(pv: Pointer; cb: Integer;
  pcbRead: PLongint): HResult;
begin
  if (not FHasSeeked) then
    Seek(0, STREAM_SEEK_SET, PLargeuint(nil)^);
  Result := inherited Read(pv, cb, pcbRead);
end;

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
        // TODO : Add support for partial writes.
        Result := stm.Write(Buffer, BurstWriteSize, @BurstWritten);
        Inc(BytesWritten, BurstWritten);
        if (Succeeded(Result)) and (Integer(BurstWritten) <> BurstWriteSize) then
          Result := E_FAIL;
        if (Failed(Result)) then
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

  // Note: Don't specify TYMED_ISTORAGE here or simple clipboard operations
  // (e.g. copy/paste text) will fail. The clipboard apparently prefers
  // TYMED_ISTORAGE over other mediums and some applications can't handle that
  // medium.
end;

function TCustomSimpleClipboardFormat.DoGetData(ADataObject: IDataObject;
  const AMedium: TStgMedium): boolean;
var
  Stream: IStream;
  StatStg: TStatStg;
  Size: integer;
  Medium: TStgMedium;
begin
  // Get size from HGlobal.
  if (AMedium.tymed = TYMED_HGLOBAL) then
  begin
    Size := GlobalSize(AMedium.HGlobal);
    Result := True;
    // Read the given amount of data.
    if (Size > 0) then
      Result := DoGetDataSized(ADataObject, AMedium, Size);
  end else
  // Get size from IStream.
  if (AMedium.tymed = TYMED_ISTREAM) then
  begin
    Stream := IStream(AMedium.stm);
    Result := (Stream <> nil) and (Succeeded(Stream.Stat(StatStg, STATFLAG_NONAME)));
    Stream := nil;
    Size := StatStg.cbSize;
    // Read the given amount of data.
    if (Result) and (Size > 0) then
      Result := DoGetDataSized(ADataObject, AMedium, Size);
  end else
  // Get size and stream from IStorage.
  if (AMedium.tymed = TYMED_ISTORAGE) then
  begin
    Stream := CreateIStreamFromIStorage(IStorage(AMedium.stg));
    Result := (Stream <> nil) and (Succeeded(Stream.Stat(StatStg, STATFLAG_NONAME)));
    Size := StatStg.cbSize;
    Medium.tymed := TYMED_ISTREAM;
    Medium.unkForRelease := nil;
    Medium.stm := pointer(Stream);
    if (Result) and (Size > 0) then
      // Read the given amount of data.
      Result := DoGetDataSized(ADataObject, Medium, Size);
  end else
    Result := False;
end;

function TCustomSimpleClipboardFormat.DoGetDataSized(ADataObject: IDataObject;
  const AMedium: TStgMedium; Size: integer): boolean;
var
  Buffer: pointer;
  Stream: IStream;
  Remaining: longInt;
  Chunk: longInt;
  pChunk: PChar;
  HGlob: HGLOBAL;
  ChunkBuffer: pointer;
const
  MaxChunk = 1024*1024; // 1Mb.
begin
  if (Size > 0) then
  begin
    // Read data from HGlobal
    if (AMedium.tymed = TYMED_HGLOBAL) then
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
    if (AMedium.tymed = TYMED_ISTREAM) then
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

          // If we have to transfer large amounts of data it is much more
          // efficient to do so in small chunks in and to use global memory.
          // Memory allocated with GetMem is much too slow because it is paged
          // and thus causes trashing (excessive page faults) if we access a
          // large memory block sequentially.
          // Tests has shown that allocating a 10Mb buffer and trying to read
          // data into it in 1Kb chunks takes several minutes, while the same
          // data can be read into a 32Kb buffer in 1Kb chunks in seconds. The
          // Windows explorer uses a 1 Mb buffer.
          // The above tests were performed using the AsyncSource demo.
          HGlob := GlobalAlloc(GMEM_SHARE or GMEM_ZEROINIT, MaxChunk);
          if (HGlob = 0) then
          begin
            Result := False;
            exit;
          end;
          try
            ChunkBuffer := GlobalLock(HGlob);
            try
              if (ChunkBuffer = nil) then
              begin
                Result := False;
                exit;
              end;

              while (Result) and (Remaining > 0) do
              begin
                if (Remaining > MaxChunk) then
                  Chunk := MaxChunk
                else
                  Chunk := Remaining;
                // Result := (Stream.Read(pChunk, Chunk, @Chunk) = S_OK);
                Result := (Succeeded(Stream.Read(ChunkBuffer, Chunk, @Chunk)));
                if (Chunk = 0) then
                  break;
                Move(ChunkBuffer^, pChunk^, Chunk);
                inc(pChunk, Chunk);
                dec(Remaining, Chunk);
              end;
            finally
              GlobalUnlock(hGlob);
            end;
          finally
            GlobalFree(HGlob);
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
    if (AMedium.tymed = TYMED_HGLOBAL) then
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
    if (AMedium.tymed = TYMED_ISTREAM) then
    begin
      Stream := IStream(AMedium.stm);
      if (Stream <> nil) then
      begin
        Stream.Seek(0, STREAM_SEEK_SET, PLargeuint(nil)^);
        Remaining := Size;
        while (Result) and (Remaining > 0) do
        begin
          Result := (Succeeded(Stream.Read(Buffer, Remaining, @Chunk)));
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
  Global: HGLOBAL;
  Stream: IStream;
  OleStream: TStream;
begin
  Result := False;

  Size := GetSize;
  if (Size <= 0) then
    exit;

  (*
  ** In this method we prefer TYMED_ISTREAM over TYMED_HGLOBAL and thus check
  ** for TYMED_ISTREAM first.
  *)

  // (FormatEtcIn.tymed <> -1) is a work around for drop targets that specify
  // the tymed incorrectly. E.g. the Nero Express CD burner does this and thus
  // asks for more than it can handle. 
  if (FormatEtcIn.tymed <> -1) and
    (FormatEtc.tymed and FormatEtcIn.tymed and TYMED_ISTREAM <> 0) then
  begin

    // Problems related to position of cursor in returned stream:
    //
    //   1) In some situations (e.g. after OleFlushClipboard) the clipboard
    //      uses an IStream.Seek(0, STREAM_SEEK_CUR) to determine the size of
    //      the stream and thus requires that Stream.Position=Stream.Size.
    //
    //   2) On Windows NT 4 the shell (shell32.dll 4.71) uses an
    //      IStream.Read(-1) to read all of the stream and thus requires that
    //      Stream.Position=0.
    //
    //   3) On Windows 2K the shell (shell32.dll 5.0) uses a IStream.Read(16K)
    //      in a loop to sequentially read all of the stream and thus requires
    //      that Stream.Position=0.
    //
    // This library uses an IStream.Stat to determine the size of the stream,
    // then uses an IStream.Seek(0, STREAM_SEEK_SET) to position to start of
    // stream and finally reads sequentially to end of stream with a number of
    // IStream.Read().
    //
    // Since we have to satisfy #1 above in order to support the clipboard, we
    // work must around #2 in TFixedStreamAdapter.CopyTo.
    //
    // At present there is no satisfactory solution to problem #3 so Windows
    // 2000 might not be fully supported. One possible (but not fully tested)
    // solution would be to implement special handling of IStream.Read(16K).

//    Global := GlobalAlloc(GMEM_MOVEABLE, Size);
    Global := GlobalAlloc(GMEM_SHARE or GMEM_ZEROINIT, Size);
    if (Global = 0) then
      exit;
    try
      p := GlobalLock(Global);
      try
        Result := WriteData(p, Size);
      finally
        GlobalUnlock(Global);
      end;

      if (not Result) or (Failed(CreateStreamOnHGlobal(Global, True, Stream))) then
      begin
        GlobalFree(Global);
        exit;
      end;

      Stream.Seek(0, STREAM_SEEK_END, PLargeuint(nil)^);

      (*
      ** The following is a bit weird...
      ** In order to intercept the calls which the other end will make into our
      ** IStream object, we have to first wrap it in a TOleStream and then wrap
      ** the TOleStream in a TFixedStreamAdapter. The TFixedStreamAdapter will
      ** then be able to work around the problems mentioned above.
      **
      ** However...
      ** If you copy something to the clipboard and then close the source
      ** application, the clipboard will make a copy of the stream, release our
      ** TFixedStreamAdapter object and we are out of luck. If clipboard
      ** operations are of no importance to you, you can disable the two lines
      ** below which deals with TOLEStream and TFixedStreamAdapter and insert
      ** the following instead:
      **
      **   Stream.Seek(0, STREAM_SEEK_SET, LargeInt(nil^));
      *)
      OleStream := TOLEStream.Create(Stream);
      Stream := TFixedStreamAdapter.Create(OleStream, soOwned) as IStream;

      IStream(AMedium.stm) := Stream;
    except
      Result := False;
    end;

    if (not Result) then
      IStream(AMedium.stm) := nil
    else
      AMedium.tymed := TYMED_ISTREAM;

  end else
  if (FormatEtc.tymed and FormatEtcIn.tymed and TYMED_HGLOBAL <> 0) then
  begin

    AMedium.hGlobal := GlobalAlloc(GMEM_SHARE or GMEM_ZEROINIT, Size);
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
(*
  if (FormatEtcIn.tymed and TYMED_ISTORAGE <> 0) then
  begin

    Global := GlobalAlloc(GMEM_SHARE or GHND, Size);
    if (Global = 0) then
      exit;

    try
      try
        p := GlobalLock(Global);
        try
          Result := (p <> nil) and WriteData(p, Size);
        finally
          GlobalUnlock(Global);
        end;

        if (Result) then
        begin
          IStorage(AMedium.stg) := CreateIStorageOnHGlobal(Global);
          Result := (AMedium.stg <> nil);
        end;
      finally
        GlobalFree(Global);
      end;
      if (Result) then
        AMedium.tymed := TYMED_ISTORAGE;
    except
      IStorage(AMedium.stg) := nil;
      Result := False;
    end;

  end else
*)
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
  s: string;
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
  s: string;
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
  Size: integer;
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
  // Validate size against count.
  // Note: Some sources (e.g. Outlook) provides a larger buffer than is needed.
  Result :=
    (Size >= integer(PFileGroupDescriptor(Value)^.cItems * SizeOf(TFileDescriptor)+SizeOf(UINT)));
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
  Size: integer;
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
  CreateFormat(TYMED_ISTREAM or TYMED_ISTORAGE);
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
  AFormatEtc: TFormatEtc;
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
    // Make copy of original FormatEtc and work with the copy.
    // If we modify the original, we *must* change it back when we are done with
    // it.
    AFormatEtc := FormatEtc;
    if (FGD.GetData(DataObject)) then
    begin
      // Multiple objects, retrieve one at a time
      Count := FGD.FileGroupDescriptor^.cItems;
      AFormatEtc.lindex := 0;
    end else
    begin
      // Single object, retrieve "all" at once
      Count := 0;
      AFormatEtc.lindex := -1;
      Name := '';
    end;
    while (AFormatEtc.lindex < Count) do
    begin
      FillChar(Medium, SizeOf(Medium), 0);
      if (Failed(DataObject.GetData(AFormatEtc, Medium))) then
        break;
      try
        inc(AFormatEtc.lindex);

        if (Medium.tymed = TYMED_ISTORAGE) then
        begin
          Stream := CreateIStreamFromIStorage(IStorage(Medium.stg));
          if (Stream = nil) then
          begin
            Result := False;
            break;
          end;
        end else
        if (Medium.tymed = TYMED_ISTREAM) then
          Stream := IStream(Medium.stm)
        else
          continue;

        Stream.Stat(StatStg, STATFLAG_NONAME);
        MemStream := TMemoryStream.Create;
        try
          Remaining := StatStg.cbSize;
          MemStream.Size := Remaining;
          pChunk := MemStream.Memory;

          // Fix for Outlook attachment paste bug #1.
          // Some versions of Outlook doesn't reset the stream position after we
          // have read data from the stream, so the next time we ask Outlook for
          // the same stream (e.g. by pasting the same attachment twice), we get
          // a stream where the current position is at EOS.
          Stream.Seek(0, STREAM_SEEK_SET, PLargeuint(nil)^);

          while (Remaining > 0) do
          begin
            if (Failed(Stream.Read(pChunk, Remaining, @Size))) or
              (Size = 0) then
              break;
            inc(pChunk, Size);
            dec(Remaining, Size);
          end;
          // Fix for Outlook attachment paste bug  #2.
          // We reset the stream position here just to be nice to other
          // applications which might not have work arounds for this problem
          // (e.g. Windows Explorer).
          Stream.Seek(0, STREAM_SEEK_SET, PLargeuint(nil)^);

          if (AFormatEtc.lindex > 0) then
            Name := FGD.FileGroupDescriptor^.fgd[AFormatEtc.lindex-1].cFileName;
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
  // We also support TYMED_ISTORAGE for drop targets, but since we only support
  // TYMED_ISTREAM for both source and targets, we can't specify TYMED_ISTORAGE
  // here. See GetStream method.
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
    Result := True;
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
  Stream: IStream;
  Index: integer;
begin
  Index := FormatEtcIn.lindex;
  (*
  ** Warning:
  ** The meaning of the value -1 in FormatEtcIn.lindex is undocumented in this
  ** context (TYMED_ISTREAM), but can occur when pasting to the clipboard.
  ** Apparently the clipboard doesn't use the stream returned from a call with
  ** lindex = -1, but only uses it as a test to see if data is available.
  ** When the clipboard actually needs the data it will specify correct values
  ** for lindex.
  ** In version 4.0 we rejected the call if -1 was specified, but in order to
  ** support clipboard operations we now map -1 to 0.
  *)
  if (Index = -1) then
    Index := 0;

  if (Assigned(FOnGetStream)) and (FormatEtcIn.tymed and TYMED_ISTREAM <> 0) and
    (Index >= 0) then
  begin
    FOnGetStream(Self, Index, Stream);

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
  Medium: TStgMedium;
  AFormatEtc: TFormatEtc;
begin
  Result := nil;
  // Get an IStream interface from the source.
  AFormatEtc := FormatEtc;
  AFormatEtc.tymed := AFormatEtc.tymed or TYMED_ISTORAGE;
  AFormatEtc.lindex := Index;
  if (Succeeded((DataFormat.Owner as TCustomDroptarget).DataObject.GetData(AFormatEtc,
    Medium))) then
    try
      if (Medium.tymed = TYMED_ISTREAM) then
      begin
        Result := IStream(Medium.stm);
      end else
      if (Medium.tymed = TYMED_ISTORAGE) then
      begin
        Result := CreateIStreamFromIStorage(IStorage(Medium.stg));
      end;
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
  FGD: TFileGroupDescritorClipboardFormat;
  Count: integer;
  Medium: TStgMedium;
  Storage: IStorage;
  Name: string;
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
      if (Failed(DataObject.GetData(FormatEtc, Medium))) then
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
//		TPasteSucceededClipboardFormat
//
////////////////////////////////////////////////////////////////////////////////
var
  CF_PASTESUCCEEDED: TClipFormat = 0;

function TPasteSucceededClipboardFormat.GetClipboardFormat: TClipFormat;
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
    FObjects: TList;
  protected
    function Get(Index: Integer): string; override;
    function GetCount: Integer; override;
    procedure PutObject(Index: Integer; AObject: TObject); override;
    function GetObject(Index: Integer): TObject; override;
  public
    constructor Create(AFileDescriptors: TMemoryList);
    destructor Destroy; override;
    procedure Clear; override;
    procedure Delete(Index: Integer); override;
    procedure Insert(Index: Integer; const S: string); override;
    procedure Assign(Source: TPersistent); override;
  end;

constructor TFileDescriptorToFilenameStrings.Create(AFileDescriptors: TMemoryList);
begin
  inherited Create;
  FFileDescriptors := AFileDescriptors;
  FObjects := TList.Create;
end;

destructor TFileDescriptorToFilenameStrings.Destroy;
begin
  FObjects.Free;
  inherited Destroy;
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
        AddObject(TStrings(Source)[i], TStrings(Source).Objects[i]);
    finally
      EndUpdate;
    end;
  end else
    inherited Assign(Source);
end;

procedure TFileDescriptorToFilenameStrings.Clear;
begin
  FFileDescriptors.Clear;
  FObjects.Clear;
end;

procedure TFileDescriptorToFilenameStrings.Delete(Index: Integer);
begin
  FFileDescriptors.Delete(Index);
  FObjects.Delete(Index);
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
      FObjects.Add(nil);
    except
      FreeMem(FD);
      raise;
    end;
  end;
end;

procedure TFileDescriptorToFilenameStrings.PutObject(Index: Integer;
  AObject: TObject);
begin
  FObjects[Index] := AObject;
end;

function TFileDescriptorToFilenameStrings.GetObject(Index: Integer): TObject;
begin
  Result := FObjects[Index];
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

  else if (Source is TPasteSucceededClipboardFormat) then
    FPasteSucceeded := TPasteSucceededClipboardFormat(Source).Value

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

  else if (Dest is TPasteSucceededClipboardFormat) then
    TPasteSucceededClipboardFormat(Dest).Value := FPasteSucceeded

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

procedure TFeedbackDataFormat.SetPasteSucceeded(const Value: longInt);
begin
  Changing;
  FPasteSucceeded := Value;
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
    (FPasteSucceeded <> DROPEFFECT_NONE) or
    (FGotInShellDragLoop);
end;

function TFeedbackDataFormat.NeedsData: boolean;
begin
  Result := (FPreferredDropEffect = DROPEFFECT_NONE) or
    (FPerformedDropEffect = DROPEFFECT_NONE) or
    (FPasteSucceeded = DROPEFFECT_NONE) or
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
  TDataStreamDataFormat.RegisterDataFormat;
  TVirtualFileStreamDataFormat.RegisterDataFormat;

  // Clipboard format registration
  TFeedbackDataFormat.RegisterCompatibleFormat(TPreferredDropEffectClipboardFormat, 0, csSourceTarget, [ddRead]);
  TFeedbackDataFormat.RegisterCompatibleFormat(TPerformedDropEffectClipboardFormat, 0, csSourceTarget, [ddWrite]);
  TFeedbackDataFormat.RegisterCompatibleFormat(TPasteSucceededClipboardFormat, 0, csSourceTarget, [ddWrite]);
  TFeedbackDataFormat.RegisterCompatibleFormat(TInShellDragLoopClipboardFormat, 0, csSourceTarget, [ddRead]);
  TFeedbackDataFormat.RegisterCompatibleFormat(TTargetCLSIDClipboardFormat, 0, csSourceTarget, [ddWrite]);
  TFeedbackDataFormat.RegisterCompatibleFormat(TLogicalPerformedDropEffectClipboardFormat, 0, csSourceTarget, [ddWrite]);
  TDataStreamDataFormat.RegisterCompatibleFormat(TFileContentsStreamClipboardFormat, 0, [csTarget], [ddRead]);

finalization
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
  TPasteSucceededClipboardFormat.UnregisterClipboardFormat;
  TInShellDragLoopClipboardFormat.UnregisterClipboardFormat;
  TTargetCLSIDClipboardFormat.UnregisterClipboardFormat;
  TLogicalPerformedDropEffectClipboardFormat.UnregisterClipboardFormat;

end.

