unit DropSource;

// -----------------------------------------------------------------------------
// Project:         Drag and Drop Component Suite
// Module:          DropSource
// Description:     Implements Dragging & Dropping of data
//                  FROM your application to another.
// Version:         4.0
// Date:            18-MAY-2001
// Target:          Win32, Delphi 5-6
// Authors:         Anders Melander, anders@melander.dk, http://www.melander.dk
// Copyright        © 1997-2001 Angus Johnson & Anders Melander
// -----------------------------------------------------------------------------
// General changes:
// - Some component glyphs has changed.
//
// TDropSource changes:
// - CutToClipboard and CopyToClipboard now uses OleSetClipboard.
//   This means that descendant classes no longer needs to override the
//   CutOrCopyToClipboard method.
// - New OnGetData event.
// - Changed to use new V4 architecture:
//   * All clipboard format support has been removed from TDropSource, it has
//     been renamed to TCustomDropSource and the old TDropSource has been
//     modified to descend from TCustomDropSource and has moved to the
//     DropSource3 unit. TDropSource is now supported for backwards
//     compatibility only and will be removed in a future version.
//   * A new TCustomDropMultiSource, derived from TCustomDropSource, uses the
//     new architecture (with TClipboardFormat and TDataFormat) and is the new
//     base class for all the drop source components.
// - TInterfacedComponent moved to DragDrop unit.
// -----------------------------------------------------------------------------
// TODO -oanme -cCheckItOut : OleQueryLinkFromData
// TODO -oanme -cDocumentation : CutToClipboard and CopyToClipboard alters the value of PreferredDropEffect.
// TODO -oanme -cDocumentation : Clipboard must be flushed or emptied manually after CutToClipboard and CopyToClipboard. Automatic flush is not guaranteed.
// TODO -oanme -cDocumentation : Delete-on-paste. Why and How.
// TODO -oanme -cDocumentation : Optimized move. Why and How.
// TODO -oanme -cDocumentation : OnPaste event is only fired if target sets the "Paste Succeeded" clipboard format. Explorer does this for delete-on-paste move operations.
// TODO -oanme -cDocumentation : DragDetectPlus. Why and How.
// -----------------------------------------------------------------------------

interface

uses
  DragDrop,
  DragDropFormats,
  ActiveX,
  Controls,
  Windows,
  Classes;

{$include DragDrop.inc}

type
  TDragResult = (drDropCopy, drDropMove, drDropLink, drCancel,
    drOutMemory, drAsync, drUnknown);

  TDropEvent = procedure(Sender: TObject; DragType: TDragType;
    var ContinueDrop: Boolean) of object;

  //: TAfterDropEvent is fired after the target has finished processing a
  // successfull drop.
  // The Optimized parameter is True if the target either performed an operation
  // other than a move or performed an "optimized move". In either cases, the
  // source isn't required to delete the source data.
  // If the Optimized parameter is False, the target performed an "unoptimized
  // move" operation and the source is required to delete the source data to
  // complete the move operation.
  TAfterDropEvent = procedure(Sender: TObject; DragResult: TDragResult;
    Optimized: Boolean) of object;

  TFeedbackEvent = procedure(Sender: TObject; Effect: LongInt;
    var UseDefaultCursors: Boolean) of object;

  //: The TDropDataEvent event is fired when the target requests data from the
  // drop source or offers data to the drop source.
  // The Handled flag should be set if the event handler satisfied the request.
  TDropDataEvent = procedure(Sender: TObject; const FormatEtc: TFormatEtc;
    out Medium: TStgMedium; var Handled: Boolean) of object;

  //: TPasteEvent is fired when the target sends a "Paste Succeeded" value
  // back to the drop source after a clipboard transfer.
  // The DeleteOnPaste parameter is True if the source is required to delete
  // the source data. This will only occur after a CutToClipboard operation
  // (corresponds to a move drag/drop).
  TPasteEvent = procedure(Sender: TObject; Action: TDragResult;
    DeleteOnPaste: boolean) of object;


////////////////////////////////////////////////////////////////////////////////
//
//		TCustomDropSource
//
////////////////////////////////////////////////////////////////////////////////
// Abstract base class for all Drop Source components.
// Implements the IDropSource and IDataObject interfaces.
////////////////////////////////////////////////////////////////////////////////
  TCustomDropSource = class(TDragDropComponent, IDropSource, IDataObject,
    IAsyncOperation)
  private
    FDragTypes: TDragTypes;
    FFeedbackEffect: LongInt;
    // Events...
    FOnDrop: TDropEvent;
    FOnAfterDrop: TAfterDropEvent;
    FOnFeedback: TFeedBackEvent;
    FOnGetData: TDropDataEvent;
    FOnSetData: TDropDataEvent;
    FOnPaste: TPasteEvent;
    // Drag images...
    FImages: TImageList;
    FShowImage: boolean;
    FImageIndex: integer;
    FImageHotSpot: TPoint;
    FDragSourceHelper: IDragSourceHelper;
    // Async transfer...
    FAllowAsync: boolean;
    FRequestAsync: boolean;
    FIsAsync: boolean;

  protected
    property FeedbackEffect: LongInt read FFeedbackEffect write FFeedbackEffect;

    // IDropSource implementation
    function QueryContinueDrag(fEscapePressed: bool;
      grfKeyState: LongInt): HRESULT; stdcall;
    function GiveFeedback(dwEffect: LongInt): HRESULT; stdcall;

    // IDataObject implementation
    function GetData(const FormatEtcIn: TFormatEtc;
      out Medium: TStgMedium):HRESULT; stdcall;
    function GetDataHere(const FormatEtc: TFormatEtc;
      out Medium: TStgMedium):HRESULT; stdcall;
    function QueryGetData(const FormatEtc: TFormatEtc): HRESULT; stdcall;
    function GetCanonicalFormatEtc(const FormatEtc: TFormatEtc;
      out FormatEtcout: TFormatEtc): HRESULT; stdcall;
    function SetData(const FormatEtc: TFormatEtc; var Medium: TStgMedium;
      fRelease: Bool): HRESULT; stdcall;
    function EnumFormatEtc(dwDirection: LongInt;
      out EnumFormatEtc: IEnumFormatEtc): HRESULT; stdcall;
    function dAdvise(const FormatEtc: TFormatEtc; advf: LongInt;
      const advsink: IAdviseSink; out dwConnection: LongInt): HRESULT; stdcall;
    function dUnadvise(dwConnection: LongInt): HRESULT; stdcall;
    function EnumdAdvise(out EnumAdvise: IEnumStatData): HRESULT; stdcall;

    // IAsyncOperation implementation
    function EndOperation(hResult: HRESULT; const pbcReserved: IBindCtx;
      dwEffects: Cardinal): HRESULT; stdcall;
    function GetAsyncMode(out fDoOpAsync: LongBool): HRESULT; stdcall;
    function InOperation(out pfInAsyncOp: LongBool): HRESULT; stdcall;
    function SetAsyncMode(fDoOpAsync: LongBool): HRESULT; stdcall;
    function StartOperation(const pbcReserved: IBindCtx): HRESULT; stdcall;

    // Abstract methods
    function DoGetData(const FormatEtcIn: TFormatEtc;
      out Medium: TStgMedium): HRESULT; virtual; abstract;
    function DoSetData(const FormatEtc: TFormatEtc;
      var Medium: TStgMedium): HRESULT; virtual;
    function HasFormat(const FormatEtc: TFormatEtc): boolean; virtual; abstract;
    function GetEnumFormatEtc(dwDirection: LongInt): IEnumFormatEtc; virtual; abstract;

    // Data format event sink
    procedure DataChanging(Sender: TObject); virtual;

    // Clipboard
    function CutOrCopyToClipboard: boolean; virtual;
    procedure DoOnPaste(Action: TDragResult; DeleteOnPaste: boolean); virtual;

    // Property access
    procedure SetShowImage(Value: boolean);
    procedure SetImages(const Value: TImageList);
    procedure SetImageIndex(const Value: integer);
    procedure SetPoint(Index: integer; Value: integer);
    function GetPoint(Index: integer): integer;
    function GetPerformedDropEffect: longInt; virtual;
    function GetLogicalPerformedDropEffect: longInt; virtual;
    procedure SetPerformedDropEffect(const Value: longInt); virtual;
    function GetPreferredDropEffect: longInt; virtual;
    procedure SetPreferredDropEffect(const Value: longInt); virtual;
    function GetInShellDragLoop: boolean; virtual;
    function GetTargetCLSID: TCLSID; virtual;
    procedure SetInShellDragLoop(const Value: boolean); virtual;
    function GetLiveDataOnClipboard: boolean;
    procedure SetAllowAsync(const Value: boolean);

    // Component management
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;

    property DragSourceHelper: IDragSourceHelper read FDragSourceHelper;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    function Execute: TDragResult; virtual;
    function CutToClipboard: boolean; virtual;
    function CopyToClipboard: boolean; virtual;
    procedure FlushClipboard; virtual;
    procedure EmptyClipboard; virtual;

    property PreferredDropEffect: longInt read GetPreferredDropEffect
      write SetPreferredDropEffect;
    property PerformedDropEffect: longInt read GetPerformedDropEffect
      write SetPerformedDropEffect;
    property LogicalPerformedDropEffect: longInt read GetLogicalPerformedDropEffect;
    property InShellDragLoop: boolean read GetInShellDragLoop
      write SetInShellDragLoop;
    property TargetCLSID: TCLSID read GetTargetCLSID;
    property LiveDataOnClipboard: boolean read GetLiveDataOnClipboard;
    property AsyncTransfer: boolean read FIsAsync;

  published
    property DragTypes: TDragTypes read FDragTypes write FDragTypes;
    // Events
    property OnFeedback: TFeedbackEvent read FOnFeedback write FOnFeedback;
    property OnDrop: TDropEvent read FOnDrop write FOnDrop;
    property OnAfterDrop: TAfterDropEvent read FOnAfterDrop write FOnAfterDrop;
    property OnGetData: TDropDataEvent read FOnGetData write FOnGetData;
    property OnSetData: TDropDataEvent read FOnSetData write FOnSetData;
    property OnPaste: TPasteEvent read FOnPaste write FOnPaste;

    // Drag Images...
    property Images: TImageList read FImages write SetImages;
    property ImageIndex: integer read FImageIndex write SetImageIndex;
    property ShowImage: boolean read FShowImage write SetShowImage;
    property ImageHotSpotX: integer index 1 read GetPoint write SetPoint;
    property ImageHotSpotY: integer index 2 read GetPoint write SetPoint;
    // Async transfer...
    property AllowAsyncTransfer: boolean read FAllowAsync write SetAllowAsync;
  end;


////////////////////////////////////////////////////////////////////////////////
//
//		TCustomDropMultiSource
//
////////////////////////////////////////////////////////////////////////////////
// Drop target base class which can accept multiple formats.
////////////////////////////////////////////////////////////////////////////////
  TCustomDropMultiSource = class(TCustomDropSource)
  private
    FFeedbackDataFormat: TFeedbackDataFormat;
    FRawDataFormat: TRawDataFormat;

  protected
    function DoGetData(const FormatEtcIn: TFormatEtc;
      out Medium: TStgMedium):HRESULT; override;
    function DoSetData(const FormatEtc: TFormatEtc;
      var Medium: TStgMedium): HRESULT; override;
    function HasFormat(const FormatEtc: TFormatEtc): boolean; override;
    function GetEnumFormatEtc(dwDirection: LongInt): IEnumFormatEtc; override;

    function GetPerformedDropEffect: longInt; override;
    function GetLogicalPerformedDropEffect: longInt; override;
    function GetPreferredDropEffect: longInt; override;
    procedure SetPerformedDropEffect(const Value: longInt); override;
    procedure SetPreferredDropEffect(const Value: longInt); override;
    function GetInShellDragLoop: boolean; override;
    procedure SetInShellDragLoop(const Value: boolean); override;
    function GetTargetCLSID: TCLSID; override;

    procedure DoOnSetData(DataFormat: TCustomDataFormat;
      ClipboardFormat: TClipboardFormat);

  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    property DataFormats;
    // TODO : Add support for delayed rendering with OnRenderData event.
  published
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		TDropEmptySource
//
////////////////////////////////////////////////////////////////////////////////
// Do-nothing source for use with TDataFormatAdapter and such
////////////////////////////////////////////////////////////////////////////////
  TDropEmptySource = class(TCustomDropMultiSource);


////////////////////////////////////////////////////////////////////////////////
//
//		TDropSourceThread
//
////////////////////////////////////////////////////////////////////////////////
// Executes a drop source operation from a thread.
// TDropSourceThread is an alternative to the Windows 2000 Asynchronous Data
// Transfer support.
////////////////////////////////////////////////////////////////////////////////
type
  TDropSourceThread = class(TThread)
  private
    FDropSource: TCustomDropSource;
    FDragResult: TDragResult;
  protected
    procedure Execute; override;
  public
    constructor Create(ADropSource: TCustomDropSource; AFreeOnTerminate: Boolean);
    property DragResult: TDragResult read FDragResult;
    property Terminated;
  end;

////////////////////////////////////////////////////////////////////////////////
//
//		Utility functions
//
////////////////////////////////////////////////////////////////////////////////
  function DropEffectToDragResult(DropEffect: longInt): TDragResult;


////////////////////////////////////////////////////////////////////////////////
//
//		Component registration
//
////////////////////////////////////////////////////////////////////////////////
procedure Register;


(*******************************************************************************
**
**			IMPLEMENTATION
**
*******************************************************************************)
implementation

uses
  CommCtrl,
  ComObj,
  Graphics;


////////////////////////////////////////////////////////////////////////////////
//
//		Component registration
//
////////////////////////////////////////////////////////////////////////////////
procedure Register;
begin
  RegisterComponents(DragDropComponentPalettePage, [TDropEmptySource]);
end;


////////////////////////////////////////////////////////////////////////////////
//
//		Utility functions
//
////////////////////////////////////////////////////////////////////////////////
function DropEffectToDragResult(DropEffect: longInt): TDragResult;
begin
  case DropEffect of
    DROPEFFECT_NONE:
      Result := drCancel;
    DROPEFFECT_COPY:
      Result := drDropCopy;
    DROPEFFECT_MOVE:
      Result := drDropMove;
    DROPEFFECT_LINK:
      Result := drDropLink;
  else
    Result := drUnknown; // This is probably an error condition
  end;
end;

// -----------------------------------------------------------------------------
//			TCustomDropSource
// -----------------------------------------------------------------------------

constructor TCustomDropSource.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  DragTypes := [dtCopy]; //default to Copy.

  // Note: Normally we would call _AddRef or coLockObjectExternal(Self) here to
  // make sure that the component wasn't deleted prematurely (e.g. after a call
  // to RegisterDragDrop), but since our ancestor class TInterfacedComponent
  // disables reference counting, we do not need to do so.

  FImageHotSpot := Point(16,16);
  FImages := nil;
end;

destructor TCustomDropSource.Destroy;
begin
  // TODO -oanme -cImprovement : Maybe FlushClipboard would be more appropiate?
  EmptyClipboard;
  inherited Destroy;
end;

// -----------------------------------------------------------------------------

function TCustomDropSource.GetCanonicalFormatEtc(const FormatEtc: TFormatEtc;
  out FormatEtcout: TFormatEtc): HRESULT;
begin
  Result := DATA_S_SAMEFORMATETC;
end;
// -----------------------------------------------------------------------------

function TCustomDropSource.SetData(const FormatEtc: TFormatEtc;
  var Medium: TStgMedium; fRelease: Bool): HRESULT;
begin
  // Warning: Ordinarily it would be much more efficient to just call
  // HasFormat(FormatEtc) to determine if we support the given format, but
  // because we have to able to accept *all* data formats, even unknown ones, in
  // order to support the Windows 2000 drag helper functionality, we can't
  // reject any formats here. Instead we pass the request on to DoSetData and
  // let it worry about the details.

  // if (HasFormat(FormatEtc)) then
  // begin
    try
      Result := DoSetData(FormatEtc, Medium);
    finally
      if (fRelease) then
        ReleaseStgMedium(Medium);
    end;
  // end else
  //   Result:= DV_E_FORMATETC;
end;
// -----------------------------------------------------------------------------

function TCustomDropSource.DAdvise(const FormatEtc: TFormatEtc; advf: LongInt;
  const advSink: IAdviseSink; out dwConnection: LongInt): HRESULT;
begin
  Result := OLE_E_ADVISENOTSUPPORTED;
end;
// -----------------------------------------------------------------------------

function TCustomDropSource.DUnadvise(dwConnection: LongInt): HRESULT;
begin
  Result := OLE_E_ADVISENOTSUPPORTED;
end;
// -----------------------------------------------------------------------------

function TCustomDropSource.EnumDAdvise(out EnumAdvise: IEnumStatData): HRESULT;
begin
  Result := OLE_E_ADVISENOTSUPPORTED;
end;
// -----------------------------------------------------------------------------

function TCustomDropSource.GetData(const FormatEtcIn: TFormatEtc;
  out Medium: TStgMedium):HRESULT; stdcall;
var
  Handled: boolean;
begin
  Handled := False;
  if (Assigned(FOnGetData)) then
    // Fire event to ask user for data.
    FOnGetData(Self, FormatEtcIn, Medium, Handled);

  // If user provided data, there is no need to call descendant for it.
  if (Handled) then
    Result := S_OK
  else if (HasFormat(FormatEtcIn)) then
    // Call descendant class to get data.
    Result := DoGetData(FormatEtcIn, Medium)
  else
    Result:= DV_E_FORMATETC;
end;
// -----------------------------------------------------------------------------

function TCustomDropSource.GetDataHere(const FormatEtc: TFormatEtc;
  out Medium: TStgMedium):HRESULT; stdcall;
begin
  Result := E_NOTIMPL;
end;
// -----------------------------------------------------------------------------

function TCustomDropSource.QueryGetData(const FormatEtc: TFormatEtc): HRESULT; stdcall;
begin
  if (HasFormat(FormatEtc)) then
    Result:= S_OK
  else
    Result:= DV_E_FORMATETC;
end;
// -----------------------------------------------------------------------------

function TCustomDropSource.EnumFormatEtc(dwDirection: LongInt;
  out EnumFormatEtc:IEnumFormatEtc): HRESULT; stdcall;
begin
  EnumFormatEtc := GetEnumFormatEtc(dwDirection);
  if (EnumFormatEtc <> nil) then
    Result := S_OK
  else
    Result := E_NOTIMPL;
end;
// -----------------------------------------------------------------------------

// Implements IDropSource.QueryContinueDrag
function TCustomDropSource.QueryContinueDrag(fEscapePressed: bool;
  grfKeyState: LongInt): HRESULT; stdcall;
var
  ContinueDrop		: Boolean;
  DragType		: TDragType;
begin
  if FEscapePressed then
    Result := DRAGDROP_S_CANCEL
  // Allow drag and drop with either mouse buttons.
  else if (grfKeyState and (MK_LBUTTON or MK_RBUTTON) = 0) then
  begin
    ContinueDrop := DropEffectToDragType(FeedbackEffect, DragType) and
      (DragType in DragTypes);

    InShellDragLoop := False;

    // If a valid drop then do OnDrop event if assigned...
    if ContinueDrop and Assigned(OnDrop) then
      OnDrop(Self, DragType, ContinueDrop);

    if ContinueDrop then
      Result := DRAGDROP_S_DROP
    else
      Result := DRAGDROP_S_CANCEL;
  end else
    Result := S_OK;
end;
// -----------------------------------------------------------------------------

// Implements IDropSource.GiveFeedback
function TCustomDropSource.GiveFeedback(dwEffect: LongInt): HRESULT; stdcall;
var
  UseDefaultCursors: Boolean;
begin
  UseDefaultCursors := True;
  FeedbackEffect := dwEffect;
  if Assigned(OnFeedback) then
    OnFeedback(Self, dwEffect, UseDefaultCursors);
  if UseDefaultCursors then
    Result := DRAGDROP_S_USEDEFAULTCURSORS
  else
    Result := S_OK;
end;
// -----------------------------------------------------------------------------

function TCustomDropSource.DoSetData(const FormatEtc: TFormatEtc;
  var Medium: TStgMedium): HRESULT;
var
  Handled: boolean;
begin
  Result := E_NOTIMPL;
  if (Assigned(FOnSetData)) then
  begin
    Handled := False;
    // Fire event to ask user to handle data.
    FOnSetData(Self, FormatEtc, Medium, Handled);
    if (Handled) then
      Result := S_OK;
  end;
end;
// -----------------------------------------------------------------------------

procedure TCustomDropSource.SetAllowAsync(const Value: boolean);
begin
  if (FAllowAsync <> Value) then
  begin
    FAllowAsync := Value;
    if (not FAllowAsync) then
    begin
      FRequestAsync := False;
      FIsAsync := False;
    end;
  end;
end;

function TCustomDropSource.GetAsyncMode(out fDoOpAsync: LongBool): HRESULT;
begin
  fDoOpAsync := FRequestAsync;
  Result := S_OK;
end;

function TCustomDropSource.SetAsyncMode(fDoOpAsync: LongBool): HRESULT;
begin
  if (FAllowAsync) then
  begin
    FRequestAsync := fDoOpAsync;
    Result := S_OK;
  end else
    Result := E_NOTIMPL;
end;

function TCustomDropSource.InOperation(out pfInAsyncOp: LongBool): HRESULT;
begin
  pfInAsyncOp := FIsAsync;
  Result := S_OK;
end;

function TCustomDropSource.StartOperation(const pbcReserved: IBindCtx): HRESULT;
begin
  if (FRequestAsync) then
  begin
    FIsAsync := True;
    Result := S_OK;
  end else
    Result := E_NOTIMPL;
end;

function TCustomDropSource.EndOperation(hResult: HRESULT;
  const pbcReserved: IBindCtx; dwEffects: Cardinal): HRESULT;
var
  DropResult: TDragResult;
begin
  if (FIsAsync) then
  begin
    FIsAsync := False;
    if (Assigned(FOnAfterDrop)) then
    begin
      if (Succeeded(hResult)) then
        DropResult := DropEffectToDragResult(dwEffects and DragTypesToDropEffect(FDragTypes))
      else
        DropResult := drUnknown;
      FOnAfterDrop(Self, DropResult,
        (DropResult <> drDropMove) or (PerformedDropEffect <> DROPEFFECT_MOVE));
    end;
    Result := S_OK;
  end else
    Result := E_FAIL;
end;

function TCustomDropSource.Execute: TDragResult;

  function GetRGBColor(Value: TColor): DWORD;
  begin
    Result := ColorToRGB(Value);
    case Result of
      clNone: Result := CLR_NONE;
      clDefault: Result := CLR_DEFAULT;
    end;
  end;

var
  DropResult: HRESULT;
  AllowedEffects,
  DropEffect: longint;
  IsDraggingImage: boolean;
  shDragImage: TSHDRAGIMAGE;
  shDragBitmap: TBitmap;
begin
  shDragBitmap := nil;

  AllowedEffects := DragTypesToDropEffect(FDragTypes);

  // Reset the "Performed Drop Effect" value. If it is supported by the target,
  // the target will set it to the desired value when the drop occurs.
  PerformedDropEffect := -1;

  if (FShowImage) then
  begin
    // Attempt to create Drag Drop helper object.
    // At present this is only supported on Windows 2000. If the object can't be
    // created, we fall back to the old image list based method (which only
    // works within the application).
    CoCreateInstance(CLSID_DragDropHelper, nil, CLSCTX_INPROC_SERVER,
      IDragSourceHelper, FDragSourceHelper);

    // Display drag image.
    if (FDragSourceHelper <> nil) then
    begin
      IsDraggingImage := True;
      shDragBitmap := TBitmap.Create;
      shDragBitmap.PixelFormat := pfDevice;
      FImages.GetBitmap(ImageIndex, shDragBitmap);
      shDragImage.hbmpDragImage := shDragBitmap.Handle;
      shDragImage.sizeDragImage.cx := shDragBitmap.Width;
      shDragImage.sizeDragImage.cy := shDragBitmap.Height;
      shDragImage.crColorKey := GetRGBColor(FImages.BkColor);
      shDragImage.ptOffset.x := ImageHotSpotX;
      shDragImage.ptOffset.y := ImageHotSpotY;
      if Failed(FDragSourceHelper.InitializeFromBitmap(shDragImage, Self)) then
      begin
        FDragSourceHelper := nil;
        shDragBitmap.Free;
        shDragBitmap := nil;
      end;
    end else
      IsDraggingImage := False;

    // Fall back to image list drag image if platform doesn't support
    // IDragSourceHelper or if we "just" failed to initialize properly.
    if (FDragSourceHelper = nil) then
    begin
      IsDraggingImage := ImageList_BeginDrag(FImages.Handle, FImageIndex,
        FImageHotSpot.X, FImageHotSpot.Y);
    end;
  end else
    IsDraggingImage := False;

  if (AllowAsyncTransfer) then
    SetAsyncMode(True);

  try
    InShellDragLoop := True;
    try
      DropResult := DoDragDrop(Self, Self, AllowedEffects, DropEffect);
    finally
      // InShellDragLoop is also reset in TCustomDropSource.QueryContinueDrag.
      // This is just to make absolutely sure that it is reset (actually no big
      // deal if it isn't).
      InShellDragLoop := False;
    end;

  finally
    if IsDraggingImage then
    begin
      if (FDragSourceHelper <> nil) then
      begin
        FDragSourceHelper := nil;
        shDragBitmap.Free;
      end else
        ImageList_EndDrag;
    end;
  end;

  case DropResult of
    DRAGDROP_S_DROP:
      (*
      ** Special handling of "optimized move".
      ** If PerformedDropEffect has been set by the target to DROPEFFECT_MOVE
      ** and the drop effect returned from DoDragDrop is different from
      ** DROPEFFECT_MOVE, then an optimized move was performed.
      ** Note: This is different from how MSDN states that an optimized move is
      ** signalled, but matches how Windows 2000 signals an optimized move.
      **
      ** On Windows 2000 an optimized move is signalled by:
      ** 1) Returning DRAGDROP_S_DROP from DoDragDrop.
      ** 2) Setting drop effect to DROPEFFECT_NONE.
      ** 3) Setting the "Performed Dropeffect" format to DROPEFFECT_MOVE.
      **
      ** On previous version of Windows, an optimized move is signalled by:
      ** 1) Returning DRAGDROP_S_DROP from DoDragDrop.
      ** 2) Setting drop effect to DROPEFFECT_MOVE.
      ** 3) Setting the "Performed Dropeffect" format to DROPEFFECT_NONE.
      **
      ** The documentation states that an optimized move is signalled by:
      ** 1) Returning DRAGDROP_S_DROP from DoDragDrop.
      ** 2) Setting drop effect to DROPEFFECT_NONE or DROPEFFECT_COPY.
      ** 3) Setting the "Performed Dropeffect" format to DROPEFFECT_NONE.
      *)
      if (LogicalPerformedDropEffect = DROPEFFECT_MOVE) or
        ((DropEffect <> DROPEFFECT_MOVE) and (PerformedDropEffect = DROPEFFECT_MOVE)) then
        Result := drDropMove
      else
        Result := DropEffectToDragResult(DropEffect and AllowedEffects);
    DRAGDROP_S_CANCEL:
      Result := drCancel;
    E_OUTOFMEMORY:
      Result := drOutMemory;
    else
      // This should never happen!
      Result := drUnknown;
  end;

  // Reset PerformedDropEffect if the target didn't set it.
  if (PerformedDropEffect = -1) then
    PerformedDropEffect := DROPEFFECT_NONE;

  // Fire OnAfterDrop event unless we are in the middle of an async data
  // transfer.
  if (not AsyncTransfer) and (Assigned(FOnAfterDrop)) then
    FOnAfterDrop(Self, Result,
      (Result = drDropMove) and
      ((DropEffect <> DROPEFFECT_MOVE) or (PerformedDropEffect <> DROPEFFECT_MOVE)));

end;
// -----------------------------------------------------------------------------

function TCustomDropSource.GetPerformedDropEffect: longInt;
begin
  Result := DROPEFFECT_NONE;
end;

function TCustomDropSource.GetLogicalPerformedDropEffect: longInt;
begin
  Result := DROPEFFECT_NONE;
end;

procedure TCustomDropSource.SetPerformedDropEffect(const Value: longInt);
begin
  // Not implemented in base class
end;

function TCustomDropSource.GetPreferredDropEffect: longInt;
begin
  Result := DROPEFFECT_NONE;
end;

procedure TCustomDropSource.SetPreferredDropEffect(const Value: longInt);
begin
  // Not implemented in base class
end;

function TCustomDropSource.GetInShellDragLoop: boolean;
begin
  Result := False;
end;

function TCustomDropSource.GetTargetCLSID: TCLSID;
begin
  Result := GUID_NULL;
end;

procedure TCustomDropSource.SetInShellDragLoop(const Value: boolean);
begin
  // Not implemented in base class
end;

procedure TCustomDropSource.DataChanging(Sender: TObject);
begin
  // Data is changing - Flush clipboard to freeze the contents 
  FlushClipboard;
end;

procedure TCustomDropSource.FlushClipboard;
begin
  // If we have live data on the clipboard...
  if (LiveDataOnClipboard) then
    // ...we force the clipboard to make a static copy of the data
    // before the data changes.
    OleCheck(OleFlushClipboard);
end;

procedure TCustomDropSource.EmptyClipboard;
begin
  // If we have live data on the clipboard...
  if (LiveDataOnClipboard) then
    // ...we empty the clipboard.
    OleCheck(OleSetClipboard(nil));
end;

function TCustomDropSource.CutToClipboard: boolean;
begin
  PreferredDropEffect := DROPEFFECT_MOVE;
  // Copy data to clipboard
  Result := CutOrCopyToClipboard;
end;
// -----------------------------------------------------------------------------

function TCustomDropSource.CopyToClipboard: boolean;
begin
  PreferredDropEffect := DROPEFFECT_COPY;
  // Copy data to clipboard
  Result := CutOrCopyToClipboard;
end;
// -----------------------------------------------------------------------------

function TCustomDropSource.CutOrCopyToClipboard: boolean;
begin
  Result := (OleSetClipboard(Self as IDataObject) = S_OK);
end;

procedure TCustomDropSource.DoOnPaste(Action: TDragResult; DeleteOnPaste: boolean);
begin
  if (Assigned(FOnPaste)) then
    FOnPaste(Self, Action, DeleteOnPaste);
end;

function TCustomDropSource.GetLiveDataOnClipboard: boolean;
begin
  Result := (OleIsCurrentClipboard(Self as IDataObject) = S_OK);
end;

// -----------------------------------------------------------------------------

procedure TCustomDropSource.SetImages(const Value: TImageList);
begin
  if (FImages = Value) then
    exit;
  FImages := Value;
  if (csLoading in ComponentState) then
    exit;

  { DONE -oanme : Shouldn't FShowImage and FImageIndex only be reset if FImages = nil? }
  if (FImages = nil) or (FImageIndex >= FImages.Count) then
    FImageIndex := 0;
  FShowImage := FShowImage and (FImages <> nil) and (FImages.Count > 0);
end;
// -----------------------------------------------------------------------------

procedure TCustomDropSource.SetImageIndex(const Value: integer);
begin
  if (csLoading in ComponentState) then
  begin
    FImageIndex := Value;
    exit;
  end;

  if (Value < 0) or (FImages.Count = 0) or (FImages = nil) then
  begin
    FImageIndex := 0;
    FShowImage := False;
  end else
    if (Value < FImages.Count) then
      FImageIndex := Value;
end;
// -----------------------------------------------------------------------------

procedure TCustomDropSource.SetPoint(Index: integer; Value: integer);
begin
  if (Index = 1) then
    FImageHotSpot.x := Value
  else
    FImageHotSpot.y := Value;
end;
// -----------------------------------------------------------------------------

function TCustomDropSource.GetPoint(Index: integer): integer;
begin
  if (Index = 1) then
    Result := FImageHotSpot.x
  else
    Result := FImageHotSpot.y;
end;
// -----------------------------------------------------------------------------

procedure TCustomDropSource.SetShowImage(Value: boolean);
begin
  FShowImage := Value;
  if (csLoading in ComponentState) then
    exit;
  if (FImages = nil) then
    FShowImage := False;
end;
// -----------------------------------------------------------------------------

procedure TCustomDropSource.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
  if (Operation = opRemove) and (AComponent = FImages) then
    Images := nil;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TEnumFormatEtc
//
////////////////////////////////////////////////////////////////////////////////
// Format enumerator used by TCustomDropMultiTarget.
////////////////////////////////////////////////////////////////////////////////
type
  TEnumFormatEtc = class(TInterfacedObject, IEnumFormatEtc)
  private
    FFormats		: TClipboardFormats;
    FIndex		: integer;
  protected
    constructor CreateClone(AFormats: TClipboardFormats; AIndex: Integer);
  public
    constructor Create(AFormats: TDataFormats; Direction: TDataDirection);
    { IEnumFormatEtc implentation }
    function Next(Celt: LongInt; out Elt; pCeltFetched: pLongInt): HRESULT; stdcall;
    function Skip(Celt: LongInt): HRESULT; stdcall;
    function Reset: HRESULT; stdcall;
    function Clone(out Enum: IEnumFormatEtc): HRESULT; stdcall;
  end;

constructor TEnumFormatEtc.Create(AFormats: TDataFormats; Direction: TDataDirection);
var
  i, j			: integer;
begin
  inherited Create;
  FFormats := TClipboardFormats.Create(nil, False);
  FIndex := 0;
  for i := 0 to AFormats.Count-1 do
    for j := 0 to AFormats[i].CompatibleFormats.Count-1 do
      if (Direction in AFormats[i].CompatibleFormats[j].DataDirections) and
        (not FFormats.Contain(TClipboardFormatClass(AFormats[i].CompatibleFormats[j].ClassType))) then
        FFormats.Add(AFormats[i].CompatibleFormats[j]);
end;

constructor TEnumFormatEtc.CreateClone(AFormats: TClipboardFormats; AIndex: Integer);
var
  i			: integer;
begin
  inherited Create;
  FFormats := TClipboardFormats.Create(nil, False);
  FIndex := AIndex;
  for i := 0 to AFormats.Count-1 do
    FFormats.Add(AFormats[i]);
end;

function TEnumFormatEtc.Next(Celt: LongInt; out Elt;
  pCeltFetched: pLongInt): HRESULT;
var
  i			: integer;
  FormatEtc		: PFormatEtc;
begin
  i := 0;
  FormatEtc := PFormatEtc(@Elt);
  while (i < Celt) and (FIndex < FFormats.Count) do
  begin
    FormatEtc^ := FFormats[FIndex].FormatEtc;
    Inc(FormatEtc);
    Inc(i);
    Inc(FIndex);
  end;

  if (pCeltFetched <> nil) then
    pCeltFetched^ := i;

  if (i = Celt) then
    Result := S_OK
  else
    Result := S_FALSE;
end;

function TEnumFormatEtc.Skip(Celt: LongInt): HRESULT;
begin
  if (FIndex + Celt <= FFormats.Count) then
  begin
    inc(FIndex, Celt);
    Result := S_OK;
  end else
  begin
    FIndex := FFormats.Count;
    Result := S_FALSE;
  end;
end;

function TEnumFormatEtc.Reset: HRESULT;
begin
  FIndex := 0;
  Result := S_OK;
end;

function TEnumFormatEtc.Clone(out Enum: IEnumFormatEtc): HRESULT;
begin
  Enum := TEnumFormatEtc.CreateClone(FFormats, FIndex);
  Result := S_OK;
end;


////////////////////////////////////////////////////////////////////////////////
//
//		TCustomDropMultiSource
//
////////////////////////////////////////////////////////////////////////////////
type
  TSourceDataFormats = class(TDataFormats)
  public
    function Add(DataFormat: TCustomDataFormat): integer; override;
  end;

function TSourceDataFormats.Add(DataFormat: TCustomDataFormat): integer;
begin
  Result := inherited Add(DataFormat);
  // Set up change notification so drop source can flush clipboard if data changes.
  DataFormat.OnChanging := TCustomDropMultiSource(DataFormat.Owner).DataChanging;
end;

constructor TCustomDropMultiSource.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FDataFormats := TSourceDataFormats.Create;
  FFeedbackDataFormat := TFeedbackDataFormat.Create(Self);
  FRawDataFormat := TRawDataFormat.Create(Self);
end;

destructor TCustomDropMultiSource.Destroy;
var
  i			: integer;
begin
  EmptyClipboard;
  // Delete all target formats owned by the object
  for i := FDataFormats.Count-1 downto 0 do
    FDataFormats[i].Free;
  FDataFormats.Free;
  inherited Destroy;
end;

function TCustomDropMultiSource.DoGetData(const FormatEtcIn: TFormatEtc;
  out Medium: TStgMedium): HRESULT;
var
  i, j: integer;
  DF: TCustomDataFormat;
  CF: TClipboardFormat;
begin
  // TODO : Add support for delayed rendering with OnRenderData event.
  Medium.tymed := 0;
  Medium.UnkForRelease := nil;
  Medium.hGlobal := 0;

  Result := DV_E_FORMATETC;

  (*
  ** Loop through all data formats associated with this drop source to find one
  ** which can offer the clipboard format requested by the target.
  *)
  for i := 0 to DataFormats.Count-1 do
  begin
    DF := DataFormats[i];

    // Ignore empty data formats.
    if (not DF.HasData) then
      continue;

    (*
    ** Loop through all the data format's supported clipboard formats to find
    ** one which contains data and can provide it in the format requested by the
    ** target.
    *)
    for j := 0 to DF.CompatibleFormats.Count-1 do
    begin
      CF := DF.CompatibleFormats[j];
      (*
      ** 1) Determine if the clipboard format supports the format requested by
      **    the target.
      ** 2) Transfer data from the data format object to the clipboard format
      **    object.
      ** 3) Determine if the clipboard format object now has data to offer.
      ** 4) Transfer the data from the clipboard format object to the medium.
      *)
      if (CF.AcceptFormat(FormatEtcIn)) and
        (DataFormats[i].AssignTo(CF)) and
        (CF.HasData) and
        (CF.SetDataToMedium(FormatEtcIn, Medium)) then
      begin
        // Once data has been sucessfully transfered to the medium, we clear
        // the data in the TClipboardFormat object in order to conserve
        // resources.
        CF.Clear;
        Result := S_OK;
        exit;
      end;
    end;
  end;
end;

function TCustomDropMultiSource.DoSetData(const FormatEtc: TFormatEtc;
  var Medium: TStgMedium): HRESULT;
var
  i, j			: integer;
  GenericClipboardFormat: TRawClipboardFormat;
begin
  Result := E_NOTIMPL;

  // Get data for requested source format.
  for i := 0 to DataFormats.Count-1 do
    for j := 0 to DataFormats[i].CompatibleFormats.Count-1 do
      if (DataFormats[i].CompatibleFormats[j].AcceptFormat(FormatEtc)) and
        (DataFormats[i].CompatibleFormats[j].GetDataFromMedium(Self, Medium)) and
        (DataFormats[i].Assign(DataFormats[i].CompatibleFormats[j])) then
      begin
        DoOnSetData(DataFormats[i], DataFormats[i].CompatibleFormats[j]);
        // Once data has been sucessfully transfered to the medium, we clear
        // the data in the TClipboardFormat object in order to conserve
        // resources.
        DataFormats[i].CompatibleFormats[j].Clear;
        Result := S_OK;
        exit;
      end;

  // The requested data format wasn't supported by any of the registered
  // clipboard formats, but in order to support the Windows 2000 drag drop helper
  // object we have to accept any data which is written to the IDataObject.
  // To do this we create a new clipboard format object, initialize it with the
  // format information passed to us and copy the data.
  GenericClipboardFormat := TRawClipboardFormat.CreateFormatEtc(FormatEtc);
  FRawDataFormat.CompatibleFormats.Add(GenericClipboardFormat);
  if (GenericClipboardFormat.GetDataFromMedium(Self, Medium)) and
    (FRawDataFormat.Assign(GenericClipboardFormat)) then
    Result := S_OK;
end;

function TCustomDropMultiSource.GetEnumFormatEtc(dwDirection: Integer): IEnumFormatEtc;
begin
  if (dwDirection = DATADIR_GET) then
    Result := TEnumFormatEtc.Create(FDataFormats, ddRead)
  else if (dwDirection = DATADIR_SET) then
    Result := TEnumFormatEtc.Create(FDataFormats, ddWrite)
  else
    Result := nil;
end;

function TCustomDropMultiSource.HasFormat(const FormatEtc: TFormatEtc): boolean;
var
  i			,
  j			: integer;
begin
  Result := False;

  for i := 0 to DataFormats.Count-1 do
    for j := 0 to DataFormats[i].CompatibleFormats.Count-1 do
      if (DataFormats[i].CompatibleFormats[j].AcceptFormat(FormatEtc)) then
      begin
        Result := True;
        exit;
      end;
end;

function TCustomDropMultiSource.GetPerformedDropEffect: longInt;
begin
  Result := FFeedbackDataFormat.PerformedDropEffect;
end;

function TCustomDropMultiSource.GetLogicalPerformedDropEffect: longInt;
begin
  Result := FFeedbackDataFormat.LogicalPerformedDropEffect;
end;

function TCustomDropMultiSource.GetPreferredDropEffect: longInt;
begin
  Result := FFeedbackDataFormat.PreferredDropEffect;
end;

procedure TCustomDropMultiSource.SetPerformedDropEffect(const Value: longInt);
begin
  FFeedbackDataFormat.PerformedDropEffect := Value;
end;

procedure TCustomDropMultiSource.SetPreferredDropEffect(const Value: longInt);
begin
  FFeedbackDataFormat.PreferredDropEffect := Value;
end;

function TCustomDropMultiSource.GetInShellDragLoop: boolean;
begin
  Result := FFeedbackDataFormat.InShellDragLoop;
end;

procedure TCustomDropMultiSource.SetInShellDragLoop(const Value: boolean);
begin
  FFeedbackDataFormat.InShellDragLoop := Value;
end;

function TCustomDropMultiSource.GetTargetCLSID: TCLSID;
begin
  Result := FFeedbackDataFormat.TargetCLSID;
end;

procedure TCustomDropMultiSource.DoOnSetData(DataFormat: TCustomDataFormat;
  ClipboardFormat: TClipboardFormat);
var
  DropEffect		: longInt;
begin
  if (ClipboardFormat is TPasteSuccededClipboardFormat) then
  begin
    DropEffect := TPasteSuccededClipboardFormat(ClipboardFormat).Value;
    DoOnPaste(DropEffectToDragResult(DropEffect),
      (DropEffect = DROPEFFECT_MOVE) and (PerformedDropEffect = DROPEFFECT_MOVE));
  end;
end;

////////////////////////////////////////////////////////////////////////////////
//
//		TDropSourceThread
//
////////////////////////////////////////////////////////////////////////////////
constructor TDropSourceThread.Create(ADropSource: TCustomDropSource;
  AFreeOnTerminate: Boolean);
begin
  inherited Create(True);
  FreeOnTerminate := AFreeOnTerminate;
  FDropSource := ADropSource;
  FDragResult := drAsync;
end;

procedure TDropSourceThread.Execute;
var
  pt: TPoint;
  hwndAttach: HWND;
  dwAttachThreadID, dwCurrentThreadID : DWORD;
begin
  (*
  ** See Microsoft Knowledgebase Article Q139408 for an explanation of the
  ** AttachThreadInput stuff.
  **   http://support.microsoft.com/support/kb/articles/Q139/4/08.asp
  *)

  // Get handle of window under mouse-cursor.
  GetCursorPos(pt);
  hwndAttach := WindowFromPoint(pt);
  ASSERT(hwndAttach<>0, 'Can''t find window with drag-object');

  // Get thread IDs.
  dwAttachThreadID := GetWindowThreadProcessId(hwndAttach, nil);
  dwCurrentThreadID := GetCurrentThreadId();

  // Attach input queues if necessary.
  if (dwAttachThreadID <> dwCurrentThreadID) then
    AttachThreadInput(dwAttachThreadID, dwCurrentThreadID, True);
  try

    // Initialize OLE for this thread.
    OleInitialize(nil);
    try
      // Start drag & drop.
      FDragResult := FDropSource.Execute;
    finally
      OleUninitialize;
    end;

  finally
    // Restore input queue settings.
    if (dwAttachThreadID <> dwCurrentThreadID) then
      AttachThreadInput(dwAttachThreadID, dwCurrentThreadID, False);
    // Set Terminated flag so owner knows that drag has finished.
    Terminate;
  end;
end;

end.

