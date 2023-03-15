unit DropTarget;

  // -----------------------------------------------------------------------------
  // Project:         Drag and Drop Target Component
  // Component Names: TDropFileTarget, TDropTextTarget
  // Module:          DropTarget
  // Description:     Implements Dragging & Dropping of text and files
  //                  INTO your application FROM another.
  // Version:	        3.0
  // Date:            22-SEP-1998
  // Target:          Win32, Delphi 3 & 4
  // Author:          Angus Johnson, ajohnson@rpi.net.au
  // Copyright        ©1998 Angus Johnson

  // -----------------------------------------------------------------------------
  // You are free to use this source but please give me credit for my work.
  // If you make improvements or derive new components from this code,
  // I would very much like to see your improvements. FEEDBACK IS WELCOME.
  // -----------------------------------------------------------------------------

  // -----------------------------------------------------------------------------
  // NOTE 1:
  // This component implements the IDropTarget COM interface instead of using the
  // Windows API message WM_DROPFILES and DragAcceptFiles(). The latter approach only
  // displays a COPY drag operation (not drag MOVE) when dragging to the target window.
  // In other words the small plus symbol will always be present in the drag image
  // when over the target window irrespective of Ctrl & Shift keyboard states when
  // implementing WM_DROPFILES and DragAcceptFiles().
  // This component however by implementing the IDropTarget COM interface allows
  // either copy OR move feedback options during a drag op. (See demo.)
  // -----------------------------------------------------------------------------
  // NOTE 2:
  // This component uses my DropSource.pas unit for the declaration of the
  // TInterfacedComponent class.
  // -----------------------------------------------------------------------------

  // History:
  // dd/mm/yy  Version  Changes
  // --------  -------  ----------------------------------------
  // 22.09.98  3.0      * Shortcuts (links) for TDropFileTarget now enabled.
  //                    * TDropSource.DoEnumFormatEtc() no longer declared abstract.
  //                    * Bug fix where StgMediums weren't released. (oops!)
  //                    * TDropTarget.GetValidDropEffect() moved to
  //                      protected section and declared virtual.
  //                    * Some bugs still with NT4 :-)
  // 08.09.98  2.0      * Delphi 3 & 4 version - using IDropTarget COM interface.
  // xx.08.97  1.0      * Delphi 2 version - using WM_DROPFILES and DragAcceptFiles().
  // -----------------------------------------------------------------------------

  // PUBLISHED PROPERTIES:
  //   DragTypes: TDragTypes;  // [dtCopy, dtMove,dtLink]
  //   Enabled: boolean;
  //   Target: TWinControl
  //   GetDataOnEnter: boolean; // usually set to false -> so gets data on drop
  // EVENTS:
  //   OnEnter: TTargetEnterEvent - optional
  //   OnDragOver: TNotifyEvent - optional
  //   OnLeave: TNotifyEvent - optional
  //   OnDrop: TTargetDropEvent - essential
  //   OnGetDropEffect: TGetDropEffectEvent - optional

  // USAGE:
  // 1. Add this non-visual component to the form you wish to drag TO.
  // 2. Select the Target Component (eg: ListView, Listbox etc.). This
  //    is the component which will register the Drop Event (ie: the
  //    cursor changes to a valid drop cursor over this component.)
  //    Note: It doesn't HAVE to be the component which will display the
  //    dropped files although it does makes more visual sense if it is.
  // 3. Set enabled to true. (Under some situations it is desirable to
  //    temporarily turn this off. (See demo.)
  // 4. Assign an OnDrop event (ie: what to do when files or text are
  //    "dropped" on your component).
  // -----------------------------------------------------------------------------


interface

  uses
    Windows, ActiveX, Classes, Controls, ShlObj, ShellApi, SysUtils,
    ClipBrd, DropSource;

  type

  TGetDropEffectEvent = procedure(Sender: TObject;
    const grfKeyState: Longint; var dwEffect: LongInt) of Object;

  //Note: TInterfacedComponent declared in DropSource.pas
  TDropTarget = class(TInterfacedComponent, IDropTarget)
  private
    fDataObj: IDataObject;
    fDragTypes: TDragTypes;
    fEnabled: boolean;
    fTarget: TWinControl;
    fGetDataOnEnter: boolean;
    fGetDropEffectEvent: TGetDropEffectEvent;
    procedure GetData; Virtual; Abstract;
  protected

    // IDropTarget methods...
    function DragEnter(const DataObj: IDataObject; grfKeyState: Longint;
      pt: TPoint; var dwEffect: Longint): HRESULT; StdCall;
    function DragOver(grfKeyState: Longint; pt: TPoint;
      var dwEffect: Longint): HRESULT; StdCall;
    function DragLeave: HRESULT; StdCall;
    function Drop(const dataObj: IDataObject; grfKeyState: Longint; pt: TPoint;
      var dwEffect: Longint): HRESULT; StdCall;

    //New methods...
    function DoDragEnter(grfKeyState: Longint; pt: TPoint;
             var dwEffect: Longint): HRESULT; Virtual; Abstract;
    function DoDragOver(grfKeyState: Longint; pt: TPoint;
             var dwEffect: Longint): HRESULT; Virtual; Abstract;
    procedure DoDragLeave; Virtual; Abstract;
    function DoDrop(grfKeyState: Longint; pt: TPoint;
             var dwEffect: Longint): HRESULT; Virtual; Abstract;

    procedure SetEnabled(Enabl: boolean);
    procedure SetTarget(targ: TWinControl);
    procedure Notification(comp: TComponent; Operation: TOperation); override;
    function GetValidDropEffect(grfKeyState: Longint): LongInt; Virtual;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  published
    property Dragtypes: TDragTypes read fDragTypes write fDragTypes;
    property Enabled: Boolean read fEnabled write SetEnabled;
    property GetDataOnEnter: Boolean read fGetDataOnEnter write fGetDataOnEnter;
    property Target: TWinControl read fTarget write SetTarget;
    property OnGetDropEffect: TGetDropEffectEvent
      read fGetDropEffectEvent write fGetDropEffectEvent;
  end;

  TTargetFileEnterEvent = procedure(Sender: TObject;
    Files: TStrings) of Object;

  TTargetFileDropEvent = procedure(Sender: TObject;
    DragType: TDragType; Files: TStrings; Point: TPoint) of Object;

  TDropFileTarget = class(TDropTarget)
  private
    fFiles: TStrings;
    fEnter: TTargetFileEnterEvent;
    fDragOver: TNotifyEvent;
    fLeave: TNotifyEvent;
    fDrop: TTargetFileDropEvent;
    procedure GetData; override;
  protected
    function DoDragEnter(grfKeyState: Longint; pt: TPoint;
             var dwEffect: Longint): HRESULT; override;
    function DoDragOver(grfKeyState: Longint; pt: TPoint;
             var dwEffect: Longint): HRESULT; override;
    procedure DoDragLeave; override;
    function DoDrop(grfKeyState: Longint; pt: TPoint;
             var dwEffect: Longint): HRESULT; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  published
    property OnEnter: TTargetFileEnterEvent read fEnter write fEnter;
    property OnDragOver: TNotifyEvent read fDragOver write fDragOver;
    property OnLeave: TNotifyEvent read fLeave write fLeave;
    property OnDrop: TTargetFileDropEvent read fDrop write fDrop;
  end;

  TTargetTextEnterEvent = procedure(Sender: TObject;
    Text: String) of Object;

  TTargetTextDropEvent = procedure(Sender: TObject;
    DragType: TDragType; Text: String; Point: TPoint) of Object;

  TDropTextTarget = class(TDropTarget)
  private
    fText: String;
    fEnter: TTargetTextEnterEvent;
    fDragOver: TNotifyEvent;
    fLeave: TNotifyEvent;
    fDrop: TTargetTextDropEvent;
    procedure GetData; override;
  protected
    function DoDragEnter(grfKeyState: Longint; pt: TPoint;
             var dwEffect: Longint): HRESULT; override;
    function DoDragOver(grfKeyState: Longint; pt: TPoint;
             var dwEffect: Longint): HRESULT; override;
    procedure DoDragLeave; override;
    function DoDrop(grfKeyState: Longint; pt: TPoint;
             var dwEffect: Longint): HRESULT; override;
  published
    property OnEnter: TTargetTextEnterEvent read fEnter write fEnter;
    property OnDragOver: TNotifyEvent read fDragOver write fDragOver;
    property OnLeave: TNotifyEvent read fLeave write fLeave;
    property OnDrop: TTargetTextDropEvent read fDrop write fDrop;
  end;

procedure Register;

implementation

procedure Register;
begin
  RegisterComponents('Samples', [TDropFileTarget, TDropTextTarget]);
end;

// -----------------------------------------------------------------------------
//			TDropTarget
// -----------------------------------------------------------------------------

//******************* TDropTarget.DragEnter *************************
function TDropTarget.DragEnter(const dataObj: IDataObject; grfKeyState: Longint;
  pt: TPoint; var dwEffect: Longint): HRESULT;
begin
  fDataObj := dataObj;
  fDataObj._AddRef;
  dwEffect := GetValidDropEffect(grfKeyState);

  Result := DoDragEnter(grfKeyState,pt,dwEffect);

  if Result <> S_OK then dwEffect := DROPEFFECT_NONE;
end;

//******************* TDropTarget.DragOver *************************
function TDropTarget.DragOver(grfKeyState: Longint; pt: TPoint; var dwEffect: Longint): HResult;
begin
  if fDataObj = nil then dwEffect := DROPEFFECT_NONE
  else dwEffect := GetValidDropEffect(grfKeyState);
  Result := DoDragOver(grfKeyState,pt,dwEffect);
end;

//******************* TDropTarget.DragLeave *************************
function TDropTarget.DragLeave: HResult;
begin
  Result := S_OK;
  if fDataObj <> nil then
  begin
    fDataObj._Release;
    fDataObj := nil;
  end;
  DoDragLeave;
end;

//******************* TDropTarget.Drop *************************
function TDropTarget.Drop(const dataObj: IDataObject; grfKeyState: Longint;
  pt: TPoint; var dwEffect: Longint): HResult;
begin
  if fDataObj = nil then
  begin
    result := E_FAIL;
    exit;
  end;
  dwEffect := GetValidDropEffect(grfKeyState);

  Result :=  DoDrop(grfKeyState,pt,dwEffect);

  // clean up!
  if fDataObj = nil then exit;
  fDataObj._Release;
  fDataObj := nil;
end;

//******************* TDropTarget.GetValidDropEffect *************************
function TDropTarget.GetValidDropEffect(grfKeyState: Longint): LongInt;
begin
  //Default drop behaviour ... assume COPY if neither Shift nor Ctrl pressed...
  if (grfKeyState and MK_SHIFT <> 0) and (grfKeyState and MK_CONTROL <> 0) and
       (dtLink in fDragTypes) then result := DROPEFFECT_LINK
  else if (grfKeyState and MK_SHIFT <> 0) and (grfKeyState and MK_CONTROL = 0) and
       (dtMove in fDragTypes) then result := DROPEFFECT_MOVE
  else if (dtCopy in fDragTypes) then result := DROPEFFECT_COPY
  else if (dtMove in fDragTypes) then result := DROPEFFECT_MOVE
  else result := DROPEFFECT_NONE;
  //Default behaviour can be overridden (see Demo).
  if Assigned(fGetDropEffectEvent) then fGetDropEffectEvent(self, grfKeyState, result);
end;


//******************* TDropTarget.Create *************************
constructor TDropTarget.Create( AOwner: TComponent );
begin
   inherited Create( AOwner );
   _AddRef;
   fEnabled := true;
   DragTypes := [dtCopy, dtMove, dtLink]; //default - allows user choice.
   fGetDataOnEnter := false;
   fDataObj := nil;
end;

//******************* TDropTarget.Destroy *************************
destructor TDropTarget.Destroy;
begin
  SetEnabled(false);
  inherited Destroy;
end;

//******************* TDropTarget.SetTarget *************************
procedure TDropTarget.SetTarget(Targ: TWinControl);
begin
  if fTarget = Targ then exit;

  if assigned(fTarget) and fEnabled then
    RevokeDragDrop(fTarget.handle);

  fTarget := Targ;

  if assigned(fTarget) and fEnabled then
    RegisterDragDrop(fTarget.handle,self as IDroptarget);
end;

//******************* TDropTarget.Notification *************************
procedure TDropTarget.Notification(comp: TComponent; Operation: TOperation);
begin
  inherited Notification(comp, Operation);
  if (comp = fTarget) and (Operation = opRemove) then
  begin
    if fEnabled and fTarget.HandleAllocated then
      RevokeDragDrop(fTarget.handle);
    fTarget := nil;
  end;
end;

//******************* TDropTarget.SetEnabled *************************
procedure TDropTarget.SetEnabled(Enabl: boolean);
begin
  fEnabled := Enabl;
  if assigned(fTarget) then
    if fEnabled then RegisterDragDrop(fTarget.handle,self as IDroptarget)
    else RevokeDragDrop(fTarget.handle);
end;

// -----------------------------------------------------------------------------
//			TDropFileTarget
// -----------------------------------------------------------------------------

const
  HDropFormatEtc: TFormatEtc = (cfFormat: CF_HDROP;
    ptd: nil; dwAspect: DVASPECT_CONTENT; lindex: -1; tymed: TYMED_HGLOBAL);
// -----------------------------------------------------------------------------

//******************* TDropFileTarget.Create *************************
constructor TDropFileTarget.Create( AOwner: TComponent );
begin
   inherited Create( AOwner );
   fFiles := TStringList.Create;
end;

//******************* TDropFileTarget.Destroy *************************
destructor TDropFileTarget.Destroy;
begin
  fFiles.Free;
  inherited Destroy;
end;

//******************* TDropFileTarget.DoDragEnter *************************
function TDropFileTarget.DoDragEnter(grfKeyState: Longint;
  pt: TPoint; var dwEffect: Longint): HResult;
begin
  fFiles.clear;
  if not Assigned(fDataObj) then
  begin
    result := E_FAIL;
    exit;
  end;

  result := fDataObj.QueryGetData(HDropFormatEtc);
  if result <> S_OK then
  begin
    //I know - calling _Release isn't necessary (in Delphi 3&4) ..
    //Delphi is supposed to do this when the interface goes out of scope.
    //However, how can I test it? Should I act on blind faith?
    fDataObj._Release;
    fDataObj := nil;
  end;

  if Assigned(fEnter) and Assigned(fDataObj) then
  begin
    //It's generally more efficient to get files only if a drop occurs
    //rather than on entering a potential target window.
    //However - sometimes there is a good reason to get them here - see Demo.
    if fGetDataOnEnter then GetData;
    fEnter(self, fFiles);
  end;
end;

//******************* TDropFileTarget.DoDragOver *************************
function TDropFileTarget.DoDragOver(grfKeyState: Longint; pt: TPoint;
             var dwEffect: Longint): HRESULT;
begin
  //Keep code in this event to a minimum as this is called very often.
  if Assigned(fDragOver) and Assigned(fDataObj) then fDragOver(self);
  RESULT := S_OK;
end;

//******************* TDropFileTarget.DoDragLeave *************************
procedure TDropFileTarget.DoDragLeave;
begin
  fFiles.clear;
  if Assigned(fLeave) then fLeave(self);
end;

//******************* TDropFileTarget.DoDrop *************************
function TDropFileTarget.DoDrop(grfKeyState: Longint; pt: TPoint;
         var dwEffect: Longint): HRESULT;
begin
  //If Filenames were collected on Entering target
  //don't bother doing it again!
  if fFiles.count = 0 then GetData;

  if fFiles.count = 0 then
  begin
    RESULT := E_FAIL;
    exit;
  end;

  RESULT := S_OK;
  if Assigned(fDrop) then
    if dwEffect = DROPEFFECT_MOVE then
      fDrop(Self, dtMove, fFiles, pt)
    else if dwEffect = DROPEFFECT_COPY then
      fDrop(Self, dtCopy, fFiles, pt)
    else
      fDrop(Self, dtLink, fFiles, pt);
end;

//******************* TDropFileTarget.GetData *************************
procedure TDropFileTarget.GetData;
var
  medium: TStgMedium;
  pdf: PDropFiles;
  dropfiles: pchar;
begin
  fFiles.clear;
  if (fDataObj.GetData(HDropFormatEtc, medium) <> S_OK) or
                             (medium.tymed <> TYMED_HGLOBAL) then exit;
  try
    pdf := GlobalLock(medium.HGlobal);
    dropfiles := PChar(pdf);
    Inc(dropfiles,pdf^.pFiles);
    while (dropfiles[0]<>#0) do
    begin
      fFiles.Add(strPas(dropfiles));
      Inc(dropfiles,1+strlen(dropfiles));
    end;
    GlobalUnlock(medium.HGlobal);
  except
  end;
  //Don't forget to clean-up!
  ReleaseStgMedium(medium);
end;

// -----------------------------------------------------------------------------
//			TDropTextTarget
// -----------------------------------------------------------------------------

// -----------------------------------------------------------------------------
const
  TextFormatEtc: TFormatEtc = (cfFormat: CF_TEXT;
    ptd: nil; dwAspect: DVASPECT_CONTENT; lindex: -1; tymed: TYMED_HGLOBAL);
// -----------------------------------------------------------------------------

//******************* TDropTextTarget.DoDragEnter *************************
function TDropTextTarget.DoDragEnter(grfKeyState: Longint;
  pt: TPoint; var dwEffect: Longint): HResult;
begin
  fText := '';
  if not Assigned(fDataObj) then
  begin
    result := E_FAIL;
    exit;
  end;

  result := fDataObj.QueryGetData(TextFormatEtc);
  if result <> S_OK then
  begin
    //I know - calling _Release isn't necessary (in Delphi 3&4) ..
    //Delphi is supposed to do this when the interface goes out of scope.
    //However, how can I test it? Should I act on blind faith?
    fDataObj._Release;
    fDataObj := nil;
  end;

  if Assigned(fEnter) and Assigned(fDataObj) then
  begin
    //It's generally more efficient to get files only if a drop occurs
    //rather than on entering a potential target window.
    //However - sometimes there is a good reason to get them here - see Demo.
    if fGetDataOnEnter then GetData;
    fEnter(self, fText);
  end;
end;

//******************* TDropTextTarget.DoDragOver *************************
function TDropTextTarget.DoDragOver(grfKeyState: Longint; pt: TPoint;
             var dwEffect: Longint): HRESULT;
begin
  //Keep code in this event to a minimum as this is called very often.
  if Assigned(fDragOver) and Assigned(fDataObj) then fDragOver(self);
  RESULT := S_OK;
end;

//******************* TDropTextTarget.DoDragLeave *************************
procedure TDropTextTarget.DoDragLeave;
begin
  fText := '';
  if Assigned(fLeave) then fLeave(self);
end;

//******************* TDropTextTarget.DoDrop *************************
function TDropTextTarget.DoDrop(grfKeyState: Longint; pt: TPoint;
         var dwEffect: Longint): HRESULT;
begin
  //If Filenames were collected on Entering target
  //don't bother doing it again!
  if fText = '' then GetData;

  if fText = '' then
  begin
    RESULT := E_FAIL;
    exit;
  end;

  RESULT := S_OK;
  if Assigned(fDrop) then
    if dwEffect = DROPEFFECT_MOVE then
      fDrop(Self, dtMove, fText, pt)
    else if dwEffect = DROPEFFECT_COPY then
      fDrop(Self, dtCopy, fText, pt)
    else
      fDrop(Self, dtLink, fText, pt);
end;

//******************* TDropTextTarget.GetData *************************
procedure TDropTextTarget.GetData;
var
  medium: TStgMedium;
  cText: pchar;
begin
  fText := '';
  if (fDataObj.GetData(TextFormatEtc, medium) <> S_OK) or
                             (medium.tymed <> TYMED_HGLOBAL) then exit;
  try
    cText := PChar(GlobalLock(medium.HGlobal));
    fText := cText;
    GlobalUnlock(medium.HGlobal);
  except
  end;
  //Don't forget to clean-up!
  ReleaseStgMedium(medium);
end;

// -----------------------------------------------------------------------------
// -----------------------------------------------------------------------------

{
// Done in DropSource...

initialization
  OleInitialize(nil);
finalization
  OleUnInitialize;
}
end.
