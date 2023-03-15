unit DropTarget;

// -----------------------------------------------------------------------------
// Project:         Drag and Drop Target Component
// Component Names: TDropFileTarget, TDropTextTarget
// Module:          DropTarget
// Description:     Implements Dragging & Dropping of text and files
//                  INTO your application FROM another.
// Version:	       3.6
// Date:            21-APR-1999
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
    Windows, Messages, ActiveX, Classes, Controls, ShlObj, ShellApi, SysUtils,
    ClipBrd, DropSource, Forms, CommCtrl, ExtCtrls, Graphics;

  type

  TScrollDirection = (sdHorizontal, sdVertical);
  TScrollDirections = set of TScrollDirection;
  
  TDropTargetEvent = procedure(Sender: TObject;
    ShiftState: TShiftState; Point: TPoint; var Effect: Longint) of Object;

  //Note: TInterfacedComponent is declared in DropSource.pas
  TDropTarget = class(TInterfacedComponent, IDropTarget)
  private
    fDataObj: IDataObject;
    fDragTypes: TDragTypes;
    fRegistered: boolean;
    fTarget: TWinControl;
    fGetDataOnEnter: boolean;
    fOnEnter: TDropTargetEvent;
    fOnDragOver: TDropTargetEvent;
    fOnLeave: TNotifyEvent;
    fOnDrop: TDropTargetEvent;
    fGetDropEffectEvent: TDropTargetEvent;

    fImages: TImageList;
    fDragImageHandle: HImageList;
    fShowImage: boolean;
    fImageHotSpot: TPoint;
    fLastPoint: TPoint; //Point where DragImage was last painted (used internally)

    fTargetScrollMargin: integer;
    fScrollDirs: TScrollDirections; //enables auto scrolling of target window during drags
    fScrollTimer: TTimer;   //and paints any drag image 'cleanly'.
    procedure DoTargetScroll(Sender: TObject);

    procedure SetTarget(Target: TWinControl);
  protected
    // IDropTarget methods...
    function DragEnter(const DataObj: IDataObject; grfKeyState: Longint;
      pt: TPoint; var dwEffect: Longint): HRESULT; StdCall;
    function DragOver(grfKeyState: Longint; pt: TPoint;
      var dwEffect: Longint): HRESULT; StdCall;
    function DragLeave: HRESULT; StdCall;
    function Drop(const dataObj: IDataObject; grfKeyState: Longint; pt: TPoint;
      var dwEffect: Longint): HRESULT; StdCall;

    function DoGetData: boolean; Virtual; Abstract;
    procedure ClearData; Virtual; Abstract;
    function HasValidFormats: boolean; Virtual; Abstract;
    function GetValidDropEffect(ShiftState: TShiftState;
      pt: TPoint; dwEffect: LongInt): LongInt; Virtual;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure Register(Target: TWinControl);
    procedure Unregister;
    function PasteFromClipboard: longint; Virtual;
    property DataObject: IDataObject read fDataObj;
    property Target: TWinControl read fTarget write SetTarget;
  published
    property Dragtypes: TDragTypes read fDragTypes write fDragTypes;
    property GetDataOnEnter: Boolean read fGetDataOnEnter write fGetDataOnEnter;
    //Events...
    property OnEnter: TDropTargetEvent read fOnEnter write fOnEnter;
    property OnDragOver: TDropTargetEvent read fOnDragOver write fOnDragOver;
    property OnLeave: TNotifyEvent read fOnLeave write fOnLeave;
    property OnDrop: TDropTargetEvent read fOnDrop write fOnDrop;
    property OnGetDropEffect: TDropTargetEvent
      read fGetDropEffectEvent write fGetDropEffectEvent;
    //Drag Images...
    property ShowImage: boolean read fShowImage write fShowImage;
  end;

  TDropFileTarget = class(TDropTarget)
  private
    fFiles: TStrings;
    fMappedNames: TStrings;
    fFileNameMapFormatEtc,
    fFileNameMapWFormatEtc: TFormatEtc;
  protected
    procedure ClearData; override;
    function DoGetData: boolean; override;
    function HasValidFormats: boolean; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    function PasteFromClipboard: longint; Override;
    property Files: TStrings Read fFiles;
    //MappedNames is only needed if files need to be renamed after a drag op
    //eg dragging from 'Recycle Bin'.
    property MappedNames: TStrings read fMappedNames;
  end;

  TDropTextTarget = class(TDropTarget)
  private
    fText: String;
  protected
    procedure ClearData; override;
    function DoGetData: boolean; override;
    function HasValidFormats: boolean; override;
  public
    function PasteFromClipboard: longint; Override;
    property Text: String Read fText Write fText;
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

function GetFilesFromHGlobal(const HGlob: HGlobal; var Files: TStrings): boolean;
function ClientPtToWindowPt(Handle: HWND; pt: TPoint): TPoint;

procedure Register;

implementation

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

function GetFilesFromHGlobal(const HGlob: HGlobal; var Files: TStrings): boolean;
var
  DropFiles: PDropFiles;
  Filename: PChar;
  s: string;
begin
  DropFiles := PDropFiles(GlobalLock(HGlob));
  try
    Filename := PChar(DropFiles) + DropFiles^.pFiles;
    while (Filename^ <> #0) do
    begin
      if (DropFiles^.fWide) then // -> NT4 compatibility
      begin
        s := PWideChar(FileName);
        inc(Filename, (Length(s) + 1) * 2);
      end else
      begin
        s := Filename;
        inc(Filename, Length(s) + 1);
      end;
      Files.Add(s);
    end;
  finally
    GlobalUnlock(HGlob);
  end;
  if Files.count > 0 then
    result := true else
    result := false;
end;

// -----------------------------------------------------------------------------
//			TDropTarget
// -----------------------------------------------------------------------------

function TDropTarget.DragEnter(const dataObj: IDataObject; grfKeyState: Longint;
  pt: TPoint; var dwEffect: Longint): HRESULT;
var
  ShiftState: TShiftState;
  TargetStyles: longint;
begin
  ClearData;
  fDataObj := dataObj;
  fDataObj._AddRef;
  result := S_OK;

  fDragImageHandle := 0;  
  if ShowImage then
  begin
    fDragImageHandle := ImageList_GetDragImage(nil,@fImageHotSpot);
    if (fDragImageHandle <> 0) then
    begin
      //Currently we will just replace any 'embedded' cursor with our
      //blank (transparent) image otherwise we sometimes get 2 cursors ...
      ImageList_SetDragCursorImage(fImages.Handle,0,fImageHotSpot.x,fImageHotSpot.y);
      with ClientPtToWindowPt(fTarget.handle,pt) do
        ImageList_DragEnter(fTarget.handle,x,y);
    end;
  end;
  
  if not HasValidFormats then
  begin
    fDataObj._Release;
    fDataObj := nil;
    dwEffect := DROPEFFECT_NONE;
    //result := E_FAIL;
    exit;
  end;

  fScrollDirs := [];

  //thanks to a suggestion by Praful Kapadia ...
  fTargetScrollMargin := abs(TDummyWinControl(fTarget).font.height);

  TargetStyles := GetWindowLong(fTarget.handle,GWL_STYLE);
  if (TargetStyles and WS_HSCROLL <> 0) then
    fScrollDirs := fScrollDirs + [sdHorizontal];
  if (TargetStyles and WS_VSCROLL <> 0) then
    fScrollDirs := fScrollDirs + [sdVertical];
  //It's generally more efficient to get data only if a drop occurs
  //rather than on entering a potential target window.
  //However - sometimes there is a good reason to get it here.
  if fGetDataOnEnter then DoGetData;

  ShiftState := KeysToShiftState(grfKeyState);
  pt := fTarget.ScreenToClient(pt);
  fLastPoint := pt;
  dwEffect := GetValidDropEffect(ShiftState,Pt,dwEffect);
  if Assigned(fOnEnter) then
    fOnEnter(self, ShiftState, pt, dwEffect);

end;
// -----------------------------------------------------------------------------

function TDropTarget.DragOver(grfKeyState: Longint;
  pt: TPoint; var dwEffect: Longint): HResult;
var
  ShiftState: TShiftState;
  IsScrolling: boolean;
begin

  pt := fTarget.ScreenToClient(pt);

  if fDataObj = nil then
  begin
    //fDataObj = nil when no valid formats .... see DragEnter method.
    dwEffect := DROPEFFECT_NONE;
    IsScrolling := false;
  end else
  begin
    ShiftState := KeysToShiftState(grfKeyState);
    dwEffect := GetValidDropEffect(ShiftState, pt, dwEffect);
    if Assigned(fOnDragOver) then
      fOnDragOver(self, ShiftState, pt, dwEffect);

    if (fScrollDirs <> []) and (dwEffect and DROPEFFECT_SCROLL <> 0) then
    begin
      IsScrolling := true;
      fScrollTimer.enabled := true
    end else
    begin
      IsScrolling := false;
      fScrollTimer.enabled := false;
    end;
  end;

  if (fDragImageHandle <> 0) and
      ((fLastPoint.x <> pt.x) or (fLastPoint.y <> pt.y)) then
  begin
    fLastPoint := pt;
    if IsScrolling then
      //fScrollTimer.enabled := true
    else with ClientPtToWindowPt(fTarget.handle,pt) do
      ImageList_DragMove(X,Y);
  end
  else
    fLastPoint := pt;
  RESULT := S_OK;
end;
// -----------------------------------------------------------------------------

function TDropTarget.DragLeave: HResult;
begin
  ClearData;
  fScrollTimer.enabled := false;

  if fDataObj <> nil then
  begin
    fDataObj._Release;
    fDataObj := nil;
  end;

  if (fDragImageHandle <> 0) then
    ImageList_DragLeave(fTarget.handle);

  if Assigned(fOnLeave) then fOnLeave(self);
  Result := S_OK;
end;
// -----------------------------------------------------------------------------

function TDropTarget.Drop(const dataObj: IDataObject; grfKeyState: Longint;
  pt: TPoint; var dwEffect: Longint): HResult;
var
  ShiftState: TShiftState;
begin
  RESULT := S_OK;

  if fDataObj = nil then
  begin
    dwEffect := DROPEFFECT_NONE;
    exit;
  end;

  fScrollTimer.enabled := false;

  if (fDragImageHandle <> 0) then
    ImageList_DragLeave(fTarget.handle);

  ShiftState := KeysToShiftState(grfKeyState);
  pt := fTarget.ScreenToClient(pt);
  dwEffect := GetValidDropEffect(ShiftState, pt, dwEffect);
  if (not fGetDataOnEnter) and (not DoGetData) then
    dwEffect := DROPEFFECT_NONE
  else if Assigned(fOnDrop) then
    fOnDrop(Self, ShiftState, pt, dwEffect);

  // clean up!
  ClearData;
  if fDataObj = nil then exit;
  fDataObj._Release;
  fDataObj := nil;
end;
// -----------------------------------------------------------------------------

constructor TDropTarget.Create( AOwner: TComponent );
var
  bm: TBitmap;
begin
   inherited Create( AOwner );
   fScrollTimer := TTimer.create(self);
   fScrollTimer.interval := 100;
   fScrollTimer.enabled := false;
   fScrollTimer.OnTimer := DoTargetScroll;
   _AddRef;
   fGetDataOnEnter := false;

   fImages := TImageList.create(self);
   //Create a blank image for fImages which we...
   //will use to hide any cursor 'embedded' in a drag image.
   //This avoids the possibility of two cursors showing.
   bm := TBitmap.create;
   with bm do
   begin
     height := 32;
     width := 32;
     Canvas.Brush.Color:=clWindow;
     Canvas.FillRect(Rect(0,0,31,31));
     fImages.AddMasked(bm,clWindow);
     free;
   end;
   fDataObj := nil;
   ShowImage := true;
end;
// -----------------------------------------------------------------------------

destructor TDropTarget.Destroy;
begin
  fImages.free;
  fScrollTimer.free;
  Unregister;
  inherited Destroy;
end;
// -----------------------------------------------------------------------------

procedure TDropTarget.SetTarget(Target: TWinControl);
begin
  if fTarget = Target then exit;
  Unregister;
  fTarget := Target;
end;
// -----------------------------------------------------------------------------

procedure TDropTarget.Register(Target: TWinControl);
begin
  if fTarget = Target then exit;
  if (fTarget <> nil) then Unregister;
  fTarget := target;
  if fTarget = nil then exit;

  //CoLockObjectExternal(self,true,false);
  if not RegisterDragDrop(fTarget.handle,self) = S_OK then
      raise Exception.create('Failed to Register '+ fTarget.name);
  fRegistered := true;
end;
// -----------------------------------------------------------------------------

procedure TDropTarget.Unregister;
begin
  fRegistered := false;
  if (fTarget = nil) or not fTarget.handleallocated then exit;

  if not RevokeDragDrop(fTarget.handle) = S_OK then
      raise Exception.create('Failed to Unregister '+ fTarget.name);

  //CoLockObjectExternal(self,false,false);
  fTarget := nil;
end;
// -----------------------------------------------------------------------------

function TDropTarget.GetValidDropEffect(ShiftState: TShiftState;
  pt: TPoint; dwEffect: LongInt): LongInt;
begin
  //dwEffect 'in' parameter = set of drop effects allowed by drop source.
  //Now filter out the effects disallowed by target...
  if not (dtCopy in fDragTypes) then
    dwEffect := dwEffect and not DROPEFFECT_COPY;
  if not (dtMove in fDragTypes) then
    dwEffect := dwEffect and not DROPEFFECT_MOVE;
  if not (dtLink in fDragTypes) then
    dwEffect := dwEffect and not DROPEFFECT_LINK;
  Result := dwEffect;

  //'Default' behaviour can be overriden by assigning OnGetDropEffect.
  if Assigned(fGetDropEffectEvent) then
    fGetDropEffectEvent(self, ShiftState, pt, Result)
  else
  begin
    //As we're only interested in ssShift & ssCtrl here
    //mouse buttons states are screened out ...
    ShiftState := ([ssShift, ssCtrl] * ShiftState);

    if (ShiftState = [ssShift, ssCtrl]) and
      (dwEffect and DROPEFFECT_LINK <> 0) then result := DROPEFFECT_LINK
    else if (ShiftState = [ssShift]) and
      (dwEffect and DROPEFFECT_MOVE <> 0) then result := DROPEFFECT_MOVE
    else if (dwEffect and DROPEFFECT_COPY <> 0) then result := DROPEFFECT_COPY
    else if (dwEffect and DROPEFFECT_MOVE <> 0) then result := DROPEFFECT_MOVE
    else if (dwEffect and DROPEFFECT_LINK <> 0) then result := DROPEFFECT_LINK
    else result := DROPEFFECT_NONE;
    //Add Scroll effect if necessary...
    if fScrollDirs = [] then exit;
    if (sdHorizontal in fScrollDirs) and
      ((pt.x < fTargetScrollMargin) or
          (pt.x>fTarget.ClientWidth - fTargetScrollMargin)) then
        result := result or integer(DROPEFFECT_SCROLL)
    else if (sdVertical in fScrollDirs) and
      ((pt.y < fTargetScrollMargin) or
          (pt.y>fTarget.ClientHeight - fTargetScrollMargin)) then
        result := result or integer(DROPEFFECT_SCROLL);
  end;
end;
// -----------------------------------------------------------------------------

procedure TDropTarget.DoTargetScroll(Sender: TObject);
begin
  with fTarget, fLastPoint do
  begin
    if (fDragImageHandle <> 0) then ImageList_DragLeave(handle);
    if (Y < fTargetScrollMargin) then
      Perform(WM_VSCROLL,SB_LINEUP,0)
    else if (Y>ClientHeight - fTargetScrollMargin) then
      Perform(WM_VSCROLL,SB_LINEDOWN,0);
    if (X < fTargetScrollMargin) then
      Perform(WM_HSCROLL,SB_LINEUP,0)
    else if (X > ClientWidth - fTargetScrollMargin) then
      Perform(WM_HSCROLL,SB_LINEDOWN,0);
    if (fDragImageHandle <> 0) then
      with ClientPtToWindowPt(handle,fLastPoint) do
        ImageList_DragEnter(handle,x,y);
  end;
end;
// -----------------------------------------------------------------------------

function TDropTarget.PasteFromClipboard: longint;
var
  Global: HGlobal;
  pEffect: ^DWORD;
begin
  if not ClipBoard.HasFormat(CF_PREFERREDDROPEFFECT) then
    result := DROPEFFECT_NONE
  else
  begin
    Global := Clipboard.GetAsHandle(CF_PREFERREDDROPEFFECT);
    pEffect := pointer(GlobalLock(Global)); // DROPEFFECT_COPY, DROPEFFECT_MOVE ...
    result := pEffect^;
    GlobalUnlock(Global);
  end;
end;

// -----------------------------------------------------------------------------
//			TDropFileTarget
// -----------------------------------------------------------------------------

constructor TDropFileTarget.Create( AOwner: TComponent );
begin
  inherited Create( AOwner );
  fFiles := TStringList.Create;
  fMappedNames := TStringList.Create;
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

destructor TDropFileTarget.Destroy;
begin
  fFiles.Free;
  fMappedNames.Free;
  inherited Destroy;
end;
// -----------------------------------------------------------------------------

function TDropFileTarget.PasteFromClipboard: longint;
var
  Global: HGlobal;
  Preferred: longint;
begin
  result  := DROPEFFECT_NONE;
  if not ClipBoard.HasFormat(CF_HDROP) then exit;
  Global := Clipboard.GetAsHandle(CF_HDROP);
  fFiles.clear;
  if not GetFilesFromHGlobal(Global,fFiles) then exit;
  Preferred := inherited PasteFromClipboard;
  //if no Preferred DropEffect then return copy else return Preferred ...
  if (Preferred = DROPEFFECT_NONE) then
    result := DROPEFFECT_COPY else
    result := Preferred;
end;
// -----------------------------------------------------------------------------

function TDropFileTarget.HasValidFormats: boolean;
begin
  result := (fDataObj.QueryGetData(HDropFormatEtc) = S_OK);
end;
// -----------------------------------------------------------------------------

procedure TDropFileTarget.ClearData;
begin
  fFiles.clear;
  fMappedNames.clear;
end;
// -----------------------------------------------------------------------------

function TDropFileTarget.DoGetData: boolean;
var
  medium: TStgMedium;
  pFilename: pChar;
  pFilenameW: PWideChar;
  sFilename: string;
begin
  ClearData;
  result := false;

  if (fDataObj.GetData(HDropFormatEtc, medium) <> S_OK) then
    exit;

  try
    if (medium.tymed = TYMED_HGLOBAL) and
       GetFilesFromHGlobal(medium.HGlobal,fFiles) then
      result := true else
      result := false;
  finally
    //Don't forget to clean-up!
    ReleaseStgMedium(medium);
  end;

  //OK, now see if file name mapping is also used ...
  if (fDataObj.GetData(fFileNameMapFormatEtc, medium) = S_OK) then
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
        if fFiles.count <> fMappedNames.count then
          fMappedNames.clear;
      finally
        GlobalUnlock(medium.HGlobal);
      end;
    end;
  finally
    ReleaseStgMedium(medium);
  end
  else if (fDataObj.GetData(fFileNameMapWFormatEtc, medium) = S_OK) then
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

// -----------------------------------------------------------------------------
//			TDropTextTarget
// -----------------------------------------------------------------------------

function TDropTextTarget.PasteFromClipboard: longint;
var
  Global: HGlobal;
  TextPtr: pChar;
begin
  result := DROPEFFECT_NONE;
  if not ClipBoard.HasFormat(CF_TEXT) then exit;
  Global := Clipboard.GetAsHandle(CF_TEXT);
  TextPtr := GlobalLock(Global);
  fText := TextPtr;
  GlobalUnlock(Global);
  result := DROPEFFECT_COPY;
end;
// -----------------------------------------------------------------------------

function TDropTextTarget.HasValidFormats: boolean;
begin
  result := (fDataObj.QueryGetData(TextFormatEtc) = S_OK);
end;
// -----------------------------------------------------------------------------

procedure TDropTextTarget.ClearData;
begin
  fText := '';
end;
// -----------------------------------------------------------------------------

function TDropTextTarget.DoGetData: boolean;
var
  medium: TStgMedium;
  cText: pchar;
begin
  result := false;
  if fText <> '' then
    result := true // already got it!
  else if (fDataObj.GetData(TextFormatEtc, medium) = S_OK) then
  begin
    try
      if (medium.tymed <> TYMED_HGLOBAL) then exit;
      cText := PChar(GlobalLock(medium.HGlobal));
      fText := cText;
      GlobalUnlock(medium.HGlobal);
      result := true;
    finally
      ReleaseStgMedium(medium);
    end;
  end
  else
    result := false;
end;

// -----------------------------------------------------------------------------
//			TDropDummy
//      This component is designed just to display drag images over the
//      registered TWincontrol but where no drop is desired (eg a TForm).
// -----------------------------------------------------------------------------

function TDropDummy.HasValidFormats: boolean;
begin
  result := true;
end;
// -----------------------------------------------------------------------------

procedure TDropDummy.ClearData;
begin
  //abstract method override
end;
// -----------------------------------------------------------------------------

function TDropDummy.DoGetData: boolean;
begin
  result := false;
end;
// -----------------------------------------------------------------------------
// -----------------------------------------------------------------------------

end.

