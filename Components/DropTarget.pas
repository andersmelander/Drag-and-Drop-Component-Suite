unit DropTarget;

// -----------------------------------------------------------------------------
//
//			*** NOT FOR RELEASE ***
//
// -----------------------------------------------------------------------------
// Project:         Drag and Drop Component Suite
// Component Names: TDropFileTarget, TDropTextTarget
// Module:          DropTarget
// Description:     Implements Dragging & Dropping of text and files
//                  INTO your application FROM another.
// Version:         3.7.1
// Date:            16-SEP-1999
// Target:          Win32, Delphi 3 - Delphi 5, C++ Builder 3, C++ Builder 4
// Authors:         Angus Johnson,   ajohnson@rpi.net.au
//                  Anders Melander, anders@melander.dk
//                                   http://www.melander.dk
// Copyright        © 1997-99 Angus Johnson & Anders Melander
// -----------------------------------------------------------------------------

interface

uses
  DropSource,
  Windows, ActiveX, Classes, Controls, CommCtrl, ExtCtrls;

{$include DragDrop.inc}

type

  TScrollDirection = (sdHorizontal, sdVertical);
  TScrollDirections = set of TScrollDirection;

  TDropTargetEvent = procedure(Sender: TObject;
    ShiftState: TShiftState; Point: TPoint; var Effect: Longint) of Object;

  TControlList = class(TObject)
  private
    FList		: TList;
    function GetControl(AIndex: integer): TWinControl;
    function GetCount: integer;
  public
    constructor Create;
    destructor Destroy; override;
    function Add(AControl: TWinControl): integer;
    procedure Remove(AControl: TWinControl);
    procedure Delete(AIndex: integer);
    function IndexOf(AControl: TWinControl): integer;
    property Count: integer read GetCount;
    property Controls[AIndex: integer]: TWinControl read GetControl; default;
  end;

  // Note: TInterfacedComponent is declared in DropSource.pas
  TDropTarget = class(TInterfacedComponent, IDropTarget)
  private
    FDataObj		: IDataObject;
    FDragTypes		: TDragTypes;
    FGetDataOnEnter	: boolean;
    FOnEnter		: TDropTargetEvent;
    FOnDragOver		: TDropTargetEvent;
    FOnLeave		: TNotifyEvent;
    FOnDrop		: TDropTargetEvent;
    FGetDropEffectEvent	: TDropTargetEvent;
    FTargets		: TControlList;
    FMultiTarget	: boolean;
    FTarget		: TWinControl;

    FImages		: TImageList;
    FDragImageHandle	: HImageList;
    FShowImage		: boolean;
    FImageHotSpot	: TPoint;
    FLastPoint		: TPoint;	// Point where DragImage was last painted (used internally)
    					// and paints any drag image 'cleanly'.
    FTargetScrollMargin	: integer;
    FScrollDirs		: TScrollDirections; //enables auto scrolling of target window during drags
    FScrollTimer	: TTimer;
    procedure DoTargetScroll(Sender: TObject);
    procedure SetShowImage(Show: boolean);
    function GetTarget: TWinControl;
  protected
    // IDropTarget methods...
    function DragEnter(const DataObj: IDataObject; grfKeyState: Longint;
      pt: TPoint; var dwEffect: Longint): HRESULT; StdCall;
    function DragOver(grfKeyState: Longint; pt: TPoint;
      var dwEffect: Longint): HRESULT; StdCall;
    function DragLeave: HRESULT; StdCall;
    function Drop(const dataObj: IDataObject; grfKeyState: Longint; pt: TPoint;
      var dwEffect: Longint): HRESULT; StdCall;

    procedure DoEnter(ShiftState: TShiftState; Point: TPoint; var Effect: Longint); virtual;
    procedure DoDragOver(ShiftState: TShiftState; Point: TPoint; var Effect: Longint); virtual;
    procedure DoDrop(ShiftState: TShiftState; Point: TPoint; var Effect: Longint); virtual;
    procedure DoLeave; virtual;

    function DoGetData: boolean; Virtual; Abstract;
    procedure ClearData; Virtual; Abstract;
    function HasValidFormats: boolean; Virtual; Abstract;
    function GetValidDropEffect(ShiftState: TShiftState;
      pt: TPoint; dwEffect: LongInt): LongInt; Virtual;
    procedure DoUnregister(ATarget: TWinControl);
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure Register(ATarget: TWinControl);
{$ifdef VER12_PLUS}
    procedure Unregister(ATarget: TWinControl = nil);
{$else}
    procedure Unregister;
{$endif}
    function FindTarget(p: TPoint): TWinControl;
    function FindNearestTarget(p: TPoint): TWinControl;
    function PasteFromClipboard: longint; Virtual;
    property DataObject: IDataObject read FDataObj;
    property Target: TWinControl read GetTarget;
    property Targets: TControlList read FTargets;
  published
    property Dragtypes: TDragTypes read FDragTypes write FDragTypes;
    property GetDataOnEnter: Boolean read FGetDataOnEnter write FGetDataOnEnter;
    //Events...
    property OnEnter: TDropTargetEvent read FOnEnter write FOnEnter;
    property OnDragOver: TDropTargetEvent read FOnDragOver write FOnDragOver;
    property OnLeave: TNotifyEvent read FOnLeave write FOnLeave;
    property OnDrop: TDropTargetEvent read FOnDrop write FOnDrop;
    property OnGetDropEffect: TDropTargetEvent
      read FGetDropEffectEvent write FGetDropEffectEvent;
    //Drag Images...
    property ShowImage: boolean read FShowImage write SetShowImage;
    property MultiTarget: boolean read FMultiTarget write FMultiTarget;
  end;

  TDropFileTarget = class(TDropTarget)
  private
    FFiles		: TStrings;
    FMappedNames	: TStrings;
    FFileNameMapFormatEtc,
    FFileNameMapWFormatEtc: TFormatEtc;
  protected
    procedure ClearData; override;
    function DoGetData: boolean; override;
    function HasValidFormats: boolean; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    function PasteFromClipboard: longint; Override;
    property Files: TStrings Read FFiles;
    //MappedNames is only needed if files need to be renamed after a drag op
    //eg dragging from 'Recycle Bin'.
    property MappedNames: TStrings read FMappedNames;
  end;

  TDropTextTarget = class(TDropTarget)
  private
    FText		: string;
  protected
    procedure ClearData; override;
    function DoGetData: boolean; override;
    function HasValidFormats: boolean; override;
  public
    function PasteFromClipboard: longint; Override;
    property Text: string Read FText Write FText;
  end;

  TDropDummy = class(TDropTarget)
  protected
    procedure ClearData; override;
    function DoGetData: boolean; override;
    function HasValidFormats: boolean; override;
  end;

const
  HDropFormatEtc: TFormatEtc = (cfFormat: CF_HDROP;
    ptd: nil; dwAspect: DVASPECT_CONTENT; lindex: -1; tymed: TYMED_HGLOBAL);
  TextFormatEtc: TFormatEtc = (cfFormat: CF_TEXT;
    ptd: nil; dwAspect: DVASPECT_CONTENT; lindex: -1; tymed: TYMED_HGLOBAL);

function GetFilesFromHGlobal(const HGlob: HGlobal; Files: TStrings): boolean;
function ClientPtToWindowPt(Handle: HWND; pt: TPoint): TPoint;

procedure Register;

implementation

uses
  SysUtils,
  Graphics,
  Messages,
  ShlObj,
  ClipBrd,
  Forms;

resourcestring
  sRegisterFailed	= 'Failed to register %s as a drop target';
  sUnregisterActiveTarget = 'Can''t unregister target while drag operation is in progress';

procedure Register;
begin
  RegisterComponents('DragDrop',[TDropFileTarget, TDropTextTarget, TDropDummy]);
end;

// TDummyWinControl is declared just to expose the protected property - Font -
// which is used to calculate the 'scroll margin' for the target window.
type
  TDummyWinControl = Class(TWinControl);

// -----------------------------------------------------------------------------
//			Miscellaneous functions ...
// -----------------------------------------------------------------------------

function ClientPtToWindowPt(Handle: HWND; pt: TPoint): TPoint;
var
  Rect: TRect;
begin
  ClientToScreen(Handle, pt);
  GetWindowRect(Handle, Rect);
  Result.X := pt.X - Rect.Left;
  Result.Y := pt.Y - Rect.Top;
end;
// -----------------------------------------------------------------------------

function GetFilesFromHGlobal(const HGlob: HGlobal; Files: TStrings): boolean;
var
  DropFiles		: PDropFiles;
  Filename		: PChar;
begin
  { TODO -oanme -cImprovement : Shouldn't there be a Files.Clear here? }
  DropFiles := PDropFiles(GlobalLock(HGlob));
  try
    Filename := PChar(DropFiles) + DropFiles^.pFiles;
    while (Filename^ <> #0) do
    begin
      if (DropFiles^.fWide) then // -> NT4 & Asian compatibility
      begin
        Files.Add(PWideChar(FileName));
        inc(Filename, (Length(PWideChar(FileName)) + 1) * 2);
      end else
      begin
        Files.Add(Filename);
        inc(Filename, Length(Filename) + 1);
      end;
    end;
  finally
    GlobalUnlock(HGlob);
  end;

  Result := (Files.Count > 0);
end;

// -----------------------------------------------------------------------------
//			TControlList
//	List of TWinControl objects.
//      Used for the TDroptarget.Targets property.
// -----------------------------------------------------------------------------

constructor TControlList.Create;
begin
  inherited Create;
  FList := TList.Create;
end;

destructor TControlList.Destroy;
begin
  FList.Free;
  inherited Destroy;
end;

function TControlList.Add(AControl: TWinControl): integer;
begin
  Result := FList.Add(AControl);
end;

procedure TControlList.Delete(AIndex: integer);
begin
  FList.Delete(AIndex);
end;

function TControlList.IndexOf(AControl: TWinControl): integer;
begin
  Result := FList.IndexOf(AControl);
end;

function TControlList.GetControl(AIndex: integer): TWinControl;
begin
  Result := TWinControl(FList[AIndex]);
end;

function TControlList.GetCount: integer;
begin
  Result := FList.Count;
end;

procedure TControlList.Remove(AControl: TWinControl);
begin
  FList.Remove(AControl);
end;

// -----------------------------------------------------------------------------
//			TDropTarget
// -----------------------------------------------------------------------------

constructor TDropTarget.Create( AOwner: TComponent );
var
  bm			: TBitmap;
begin
   inherited Create(AOwner);
   FScrollTimer := TTimer.create(Self);
   FScrollTimer.Interval := 100;
   FScrollTimer.Enabled := False;
   FScrollTimer.OnTimer := DoTargetScroll;
   _AddRef;
   FGetDataOnEnter := False;
   FTargets :=  TControlList.Create;

   FImages := TImageList.Create(Self);
   // Create a blank image for FImages which we will use to hide any cursor
   // 'embedded' in a drag image.
   // This avoids the possibility of two cursors showing.
   bm := TBitmap.Create;
   with bm do
     try
       Height := 32;
       Width := 32;
       Canvas.Brush.Color := clWindow;
       Canvas.FillRect(Canvas.ClipRect);
       FImages.AddMasked(bm, clWindow);
     finally
       Free;
     end;
   FDataObj := nil;
   ShowImage := True;
end;
// -----------------------------------------------------------------------------

destructor TDropTarget.Destroy;
begin
  Unregister;
  FImages.Free;
  FScrollTimer.Free;
  FTargets.Free;
  inherited Destroy;
end;
// -----------------------------------------------------------------------------

function TDropTarget.DragEnter(const dataObj: IDataObject; grfKeyState: Longint;
  pt: TPoint; var dwEffect: Longint): HRESULT;
var
  ShiftState		: TShiftState;
  TargetStyles		: longint;
begin
  ClearData;
  FDataObj := dataObj;
  { TODO -oanme -cBug : Why _AddRef here? Delphi should take care of this for us. }
  FDataObj._AddRef;
  Result := S_OK;

  // Find the target control
  // Note: We have to use FindNearestTarget instead of FindTarget, because the
  // drag point passed to us might be outside the target rect.
  FTarget := FindNearestTarget(pt);

  // Refuse drop if we couldn't find the target control (shouldn't happen).
  if (FTarget = nil) then
  begin
    FDataObj := nil;
    Result := DROPEFFECT_NONE;
    exit;
  end;

  pt := FTarget.ScreenToClient(pt);
  FLastPoint := pt;

  FDragImageHandle := 0;
  if ShowImage then
  begin
    FDragImageHandle := ImageList_GetDragImage(nil, @FImageHotSpot);
    if (FDragImageHandle <> 0) then
    begin
      //Currently we will just replace any 'embedded' cursor with our
      //blank (transparent) image otherwise we sometimes get 2 cursors ...
      ImageList_SetDragCursorImage(FImages.Handle, 0, FImageHotSpot.x, FImageHotSpot.y);
      with ClientPtToWindowPt(FTarget.Handle, pt) do
        ImageList_DragEnter(FTarget.handle, x, y);
    end;
  end;

  if not HasValidFormats then
  begin
    FDataObj := nil;
    dwEffect := DROPEFFECT_NONE;
    exit;
  end;

  FScrollDirs := [];

  // thanks to a suggestion by Praful Kapadia ...
  FTargetScrollMargin := abs(TDummyWinControl(FTarget).Font.Height);

  TargetStyles := GetWindowLong(FTarget.Handle, GWL_STYLE);
  if (TargetStyles and WS_HSCROLL <> 0) then
    FScrollDirs := FScrollDirs + [sdHorizontal];
  if (TargetStyles and WS_VSCROLL <> 0) then
    FScrollDirs := FScrollDirs + [sdVertical];
  //It's generally more efficient to get data only if a drop occurs
  //rather than on entering a potential target window.
  //However - sometimes there is a good reason to get it here.
  if FGetDataOnEnter then
    DoGetData;

  ShiftState := KeysToShiftState(grfKeyState);
  dwEffect := GetValidDropEffect(ShiftState, Pt, dwEffect);
  DoEnter(ShiftState, pt, dwEffect);

end;
// -----------------------------------------------------------------------------

procedure TDropTarget.DoEnter(ShiftState: TShiftState; Point: TPoint; var Effect: Longint);
begin
  if Assigned(FOnEnter) then
    FOnEnter(Self, ShiftState, Point, Effect);
end;
// -----------------------------------------------------------------------------

function TDropTarget.DragOver(grfKeyState: Longint;
  pt: TPoint; var dwEffect: Longint): HResult;
var
  ShiftState: TShiftState;
  IsScrolling: boolean;
begin
  // Refuse drop if we couldn't find the target control (shouldn't happen)
  if (FTarget = nil) then
  begin
    dwEffect := DROPEFFECT_NONE;
    Result := S_OK;
    exit;
  end;

  pt := FTarget.ScreenToClient(pt);

  if (FDataObj = nil) then
  begin
    //FDataObj = nil when no valid formats .... see DragEnter method.
    dwEffect := DROPEFFECT_NONE;
    IsScrolling := False;
  end else
  begin
    ShiftState := KeysToShiftState(grfKeyState);
    dwEffect := GetValidDropEffect(ShiftState, pt, dwEffect);
    DoDragOver(ShiftState, pt, dwEffect);

    if (FScrollDirs <> []) and (dwEffect and DROPEFFECT_SCROLL <> 0) then
    begin
      IsScrolling := True;
      FScrollTimer.Enabled := True
    end else
    begin
      IsScrolling := False;
      FScrollTimer.Enabled := False;
    end;
  end;

  if (FDragImageHandle <> 0) and
      ((FLastPoint.x <> pt.x) or (FLastPoint.y <> pt.y)) then
  begin
    FLastPoint := pt;
    if IsScrolling then
      //FScrollTimer.enabled := True
    else with ClientPtToWindowPt(FTarget.Handle, pt) do
      ImageList_DragMove(X, Y);
  end else
    FLastPoint := pt;

  Result := S_OK;
end;
// -----------------------------------------------------------------------------

procedure TDropTarget.DoDragOver(ShiftState: TShiftState; Point: TPoint; var Effect: Longint);
begin
  if Assigned(FOnDragOver) then
    FOnDragOver(Self, ShiftState, Point, Effect);
end;
// -----------------------------------------------------------------------------

function TDropTarget.DragLeave: HResult;
begin
  ClearData;
  FScrollTimer.Enabled := False;

  FDataObj := nil;

  if (FDragImageHandle <> 0) then
    ImageList_DragLeave(FTarget.Handle);

  DoLeave;
  FTarget := nil;
  Result := S_OK;
end;
// -----------------------------------------------------------------------------

procedure TDropTarget.DoLeave;
begin
  if Assigned(FOnLeave) then
    FOnLeave(Self);
end;
// -----------------------------------------------------------------------------

function TDropTarget.Drop(const dataObj: IDataObject; grfKeyState: Longint;
  pt: TPoint; var dwEffect: Longint): HResult;
var
  ShiftState		: TShiftState;
begin
  Result := S_OK;

  if FDataObj = nil then
  begin
    dwEffect := DROPEFFECT_NONE;
    exit;
  end;

  FScrollTimer.Enabled := False;

  if (FDragImageHandle <> 0) then
    ImageList_DragLeave(FTarget.Handle);

  ShiftState := KeysToShiftState(grfKeyState);
  pt := FTarget.ScreenToClient(pt);
  dwEffect := GetValidDropEffect(ShiftState, pt, dwEffect);
  if (not FGetDataOnEnter) and (not DoGetData) then
    dwEffect := DROPEFFECT_NONE
  else
    DoDrop(ShiftState, pt, dwEffect);

  // clean up!
  ClearData;
  FDataObj := nil;
  FTarget := nil;
end;
// -----------------------------------------------------------------------------

procedure TDropTarget.DoDrop(ShiftState: TShiftState; Point: TPoint; var Effect: Longint);
begin
  if Assigned(FOnDrop) then
    FOnDrop(Self, ShiftState, Point, Effect);
end;
// -----------------------------------------------------------------------------

procedure TDropTarget.Register(ATarget: TWinControl);
begin
  if (FTargets.IndexOf(ATarget) <> -1) then
    exit;

  if (not FMultiTarget) then
    Unregister;

  if (ATarget = nil) then
    exit;

  if not (RegisterDragDrop(ATarget.Handle, Self) = S_OK) then
      raise Exception.CreateFmt(sRegisterFailed, [ATarget.Name]);

  FTargets.Add(ATarget);
end;
// -----------------------------------------------------------------------------

{$ifdef VER12_PLUS}
procedure TDropTarget.Unregister(ATarget: TWinControl);
begin
  DoUnregister(ATarget);
end;
{$else}
procedure TDropTarget.Unregister;
begin
  DoUnregister(nil);
end;
{$endif}

procedure TDropTarget.DoUnregister(ATarget: TWinControl);
var
  i			: integer;
begin
  if (ATarget = nil) then
  begin
    for i := FTargets.Count-1 downto 0 do
      DoUnregister(FTargets[i]);
    exit;
  end;

  i := FTargets.IndexOf(ATarget);
  if (i = -1) then
    exit;

  if (ATarget = FTarget) then
    raise Exception.Create(sUnregisterActiveTarget);

  with Targets[i] do
    if (HandleAllocated) then
      // Ignore failed unregistrations - nothing to do about it anyway
      RevokeDragDrop(Handle);

  FTargets.Delete(i);
end;
// -----------------------------------------------------------------------------

function TDropTarget.FindTarget(p: TPoint): TWinControl;
var
  i			: integer;
  r			: TRect;
begin
  for i := 0 to Targets.Count-1 do
  begin
    Result := Targets[i];
    r := Result.ClientRect;
    inc(r.Right);
    inc(r.Bottom);
    if (PtInRect(r, Result.ScreenToClient(p))) then
      exit;
  end;
  Result := nil;
end;
// -----------------------------------------------------------------------------

function TDropTarget.FindNearestTarget(p: TPoint): TWinControl;
var
  i			: integer;
  r			: TRect;
  pc			: TPoint;
  Control		: TWinControl;
  Dist			,
  BestDist		: integer;

  function Distance(r: TRect; p: TPoint): integer;
  var
    dx			,
    dy			: integer;
  begin
    if (p.x < r.Left) then
      dx := r.Left - p.x
    else if (p.x > r.Right) then
      dx := r.Right - p.x
    else
      dx := 0;
    if (p.y < r.Top) then
      dy := r.Top - p.y
    else if (p.y > r.Bottom) then
      dy := r.Bottom - p.y
    else
      dy := 0;
    Result := dx*dx + dy*dy;
  end;

begin
  Result := nil;
  BestDist := high(integer);
  for i := 0 to Targets.Count-1 do
  begin
    Control := Targets[i];
    r := Control.ClientRect;
    inc(r.Right);
    inc(r.Bottom);
    pc := Control.ScreenToClient(p);
    if (PtInRect(r, p)) then
    begin
      Result := Control;
      exit;
    end;
    Dist := Distance(r, pc);
    if (Dist < BestDist) then
    begin
      Result := Control;
      BestDist := Dist;
    end;
  end;
end;

// -----------------------------------------------------------------------------

function TDropTarget.GetTarget: TWinControl;
begin
  Result := FTarget;
  if (Result = nil) then
  begin
    if (FTargets.Count > 0) then
      Result := TWinControl(FTargets[0])
    else
      Result := nil;
  end;
end;
// -----------------------------------------------------------------------------

procedure TDropTarget.SetShowImage(Show: boolean);
begin
  FShowImage := Show;
  if FDataObj <> nil then
    ImageList_DragShowNolock(FShowImage);
end;
// -----------------------------------------------------------------------------

function TDropTarget.GetValidDropEffect(ShiftState: TShiftState;
  pt: TPoint; dwEffect: LongInt): LongInt;
begin
  //dwEffect 'in' parameter = set of drop effects allowed by drop source.
  //Now filter out the effects disallowed by target...
  if not (dtCopy in FDragTypes) then
    dwEffect := dwEffect AND NOT DROPEFFECT_COPY;
  if not (dtMove in FDragTypes) then
    dwEffect := dwEffect AND NOT DROPEFFECT_MOVE;
  if not (dtLink in FDragTypes) then
    dwEffect := dwEffect AND NOT DROPEFFECT_LINK;
  Result := dwEffect;

  //'Default' behaviour can be overriden by assigning OnGetDropEffect.
  if Assigned(FGetDropEffectEvent) then
    FGetDropEffectEvent(self, ShiftState, pt, Result)
  else
  begin
    //As we're only interested in ssShift & ssCtrl here
    //mouse buttons states are screened out ...
    ShiftState := ([ssShift, ssCtrl] * ShiftState);

    if (ShiftState = [ssShift, ssCtrl]) and
      (dwEffect AND DROPEFFECT_LINK <> 0) then Result := DROPEFFECT_LINK
    else if (ShiftState = [ssShift]) and
      (dwEffect AND DROPEFFECT_MOVE <> 0) then Result := DROPEFFECT_MOVE
    else if (dwEffect AND DROPEFFECT_COPY <> 0) then Result := DROPEFFECT_COPY
    else if (dwEffect AND DROPEFFECT_MOVE <> 0) then Result := DROPEFFECT_MOVE
    else if (dwEffect AND DROPEFFECT_LINK <> 0) then Result := DROPEFFECT_LINK
    else Result := DROPEFFECT_NONE;
    //Add Scroll effect if necessary...
    if FScrollDirs = [] then
      Exit;
    if (sdHorizontal in FScrollDirs) and
      ((pt.x < FTargetScrollMargin) or (pt.x > fTarget.ClientWidth - FTargetScrollMargin)) then
      Result := Result OR integer(DROPEFFECT_SCROLL)
    else if (sdVertical in FScrollDirs) and
      ((pt.y < FTargetScrollMargin) or (pt.y > fTarget.ClientHeight - FTargetScrollMargin)) then
      Result := Result OR integer(DROPEFFECT_SCROLL);
  end;
end;
// -----------------------------------------------------------------------------

procedure TDropTarget.DoTargetScroll(Sender: TObject);
begin
  with FTarget, FLastPoint do
  begin
    if (FDragImageHandle <> 0) then
      ImageList_DragLeave(Handle);

    if (Y < FTargetScrollMargin) then
      Perform(WM_VSCROLL,SB_LINEUP, 0)
    else if (Y > ClientHeight - FTargetScrollMargin) then
      Perform(WM_VSCROLL,SB_LINEDOWN, 0);
    if (X < FTargetScrollMargin) then
      Perform(WM_HSCROLL,SB_LINEUP, 0)
    else if (X > ClientWidth - FTargetScrollMargin) then
      Perform(WM_HSCROLL,SB_LINEDOWN, 0);

    if (FDragImageHandle <> 0) then
      with ClientPtToWindowPt(Handle, FLastPoint) do
        ImageList_DragEnter(Handle, X, Y);
  end;
end;
// -----------------------------------------------------------------------------

function TDropTarget.PasteFromClipboard: longint;
var
  Global		: HGlobal;
  pEffect		: ^DWORD;
begin
  if not ClipBoard.HasFormat(CF_PREFERREDDROPEFFECT) then
    Result := DROPEFFECT_NONE
  else
  begin
    Global := Clipboard.GetAsHandle(CF_PREFERREDDROPEFFECT);
    try
      pEffect := pointer(GlobalLock(Global)); // DROPEFFECT_COPY, DROPEFFECT_MOVE ...
      Result := pEffect^;
    finally
      GlobalUnlock(Global);
    end;
  end;
end;

// -----------------------------------------------------------------------------
//			TDropFileTarget
// -----------------------------------------------------------------------------

constructor TDropFileTarget.Create( AOwner: TComponent );
begin
  inherited Create( AOwner );
  FFiles := TStringList.Create;
  FMappedNames := TStringList.Create;
  with FFileNameMapFormatEtc do
  begin
    cfFormat := CF_FILENAMEMAP;
    ptd := nil;
    dwAspect := DVASPECT_CONTENT;
    lindex := -1;
    tymed := TYMED_HGLOBAL;
  end;
  with FFileNameMapWFormatEtc do
  begin
    cfFormat := CF_FILENAMEMAPW;
    ptd := nil;
    dwAspect := DVASPECT_CONTENT;
    lindex := -1;
    tymed := TYMED_HGLOBAL;
  end;
end;
// -----------------------------------------------------------------------------

destructor TDropFileTarget.Destroy;
begin
  FFiles.Free;
  FMappedNames.Free;
  inherited Destroy;
end;
// -----------------------------------------------------------------------------

function TDropFileTarget.PasteFromClipboard: longint;
var
  Global		: HGlobal;
  Preferred		: longint;
begin
  Result  := DROPEFFECT_NONE;
  if not ClipBoard.HasFormat(CF_HDROP) then
    exit;
  Global := Clipboard.GetAsHandle(CF_HDROP);
  FFiles.Clear;
  if not GetFilesFromHGlobal(Global, FFiles) then
    exit;
  Preferred := inherited PasteFromClipboard;
  //if no Preferred DropEffect then return copy else return Preferred ...
  if (Preferred = DROPEFFECT_NONE) then
    Result := DROPEFFECT_COPY
  else
    Result := Preferred;
end;
// -----------------------------------------------------------------------------

function TDropFileTarget.HasValidFormats: boolean;
begin
  Result := (FDataObj.QueryGetData(HDropFormatEtc) = S_OK);
end;
// -----------------------------------------------------------------------------

procedure TDropFileTarget.ClearData;
begin
  FFiles.Clear;
  FMappedNames.Clear;
end;
// -----------------------------------------------------------------------------

function TDropFileTarget.DoGetData: boolean;
var
  medium		: TStgMedium;
  pFilename		: pChar;
  pFilenameW		: PWideChar;
  sFilename		: string;
begin
  ClearData;
  Result := False;

  if (FDataObj.GetData(HDropFormatEtc, medium) <> S_OK) then
    exit;

  try
    Result := (medium.tymed = TYMED_HGLOBAL) and
       GetFilesFromHGlobal(medium.HGlobal, FFiles);
  finally
    //Don't forget to clean-up!
    ReleaseStgMedium(medium);
  end;

  //OK, now see if file name mapping is also used ...
  if (FDataObj.GetData(FFileNameMapFormatEtc, medium) = S_OK) then
    try
      if (medium.tymed = TYMED_HGLOBAL) then
      begin
        pFilename := GlobalLock(medium.HGlobal);
        try
          while True do
          begin
            sFilename := pFilename;
            if sFilename = '' then
              break;
            FMappedNames.add(sFilename);
            inc(pFilename, length(sFilename)+1);
          end;
          if FFiles.Count <> FMappedNames.Count then
            FMappedNames.Clear;
        finally
          GlobalUnlock(medium.HGlobal);
        end;
      end;
    finally
      ReleaseStgMedium(medium);
    end
  else if (FDataObj.GetData(FFileNameMapWFormatEtc, medium) = S_OK) then
    try
      if (medium.tymed = TYMED_HGLOBAL) then
      begin
        pFilenameW := GlobalLock(medium.HGlobal);
        try
          while True do
          begin
            sFilename := WideCharToString(pFilenameW);
            if sFilename = '' then
              break;
            FMappedNames.Add(sFilename);
            inc(pFilenameW, length(sFilename)+1);
          end;
          if FFiles.Count <> FMappedNames.Count then
            FMappedNames.Clear;
        finally
          GlobalUnlock(medium.HGlobal);
        end;
      end;
    finally
      ReleaseStgMedium(medium);
    end;
end;

// -----------------------------------------------------------------------------
//			TDropTextTarget
// -----------------------------------------------------------------------------

function TDropTextTarget.PasteFromClipboard: longint;
var
  Global		: HGlobal;
  TextPtr		: pChar;
begin
  Result := DROPEFFECT_NONE;
  if not ClipBoard.HasFormat(CF_TEXT) then
    exit;
  Global := Clipboard.GetAsHandle(CF_TEXT);
  TextPtr := GlobalLock(Global);
  FText := TextPtr;
  GlobalUnlock(Global);
  Result := DROPEFFECT_COPY;
end;
// -----------------------------------------------------------------------------

function TDropTextTarget.HasValidFormats: boolean;
begin
  Result := (FDataObj.QueryGetData(TextFormatEtc) = S_OK);
end;
// -----------------------------------------------------------------------------

procedure TDropTextTarget.ClearData;
begin
  FText := '';
end;
// -----------------------------------------------------------------------------

function TDropTextTarget.DoGetData: boolean;
var
  medium		: TStgMedium;
  cText			: pchar;
begin
  Result := False;
  if FText <> '' then
    Result := True // already got it!
  else if (FDataObj.GetData(TextFormatEtc, medium) = S_OK) then
  begin
    try
      if (medium.tymed <> TYMED_HGLOBAL) then
        exit;
      cText := PChar(GlobalLock(medium.HGlobal));
      FText := cText;
      GlobalUnlock(medium.HGlobal);
      Result := True;
    finally
      ReleaseStgMedium(medium);
    end;
  end else
    Result := False;
end;

// -----------------------------------------------------------------------------
//			TDropDummy
//      This component is designed just to display drag images over the
//      registered TWincontrol but where no drop is desired (eg a TForm).
// -----------------------------------------------------------------------------

function TDropDummy.HasValidFormats: boolean;
begin
  Result := True;
end;
// -----------------------------------------------------------------------------

procedure TDropDummy.ClearData;
begin
  //abstract method override
end;
// -----------------------------------------------------------------------------

function TDropDummy.DoGetData: boolean;
begin
  Result := False;
end;
// -----------------------------------------------------------------------------
// -----------------------------------------------------------------------------

end.

