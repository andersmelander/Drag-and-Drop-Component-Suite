unit DropTarget;

  // -----------------------------------------------------------------------------
  // Project:         Drag and Drop Target Component
  // Component Names: TDropFileTarget, TDropTextTarget, TDropURLTarget
  // Module:          DropTarget
  // Description:     Implements Dragging & Dropping of text, files and URLs
  //                  INTO your application FROM another.
  // Version:	       3.3
  // Date:            16-NOV-1998
  // Target:          Win32, Delphi 3 & 4, CB3
  // Authors:         Angus Johnson,   ajohnson@rpi.net.au
  //                  Anders Melander, anders@melander.dk
  //                                   http://www.melander.dk
  //                  Graham Wideman,  graham@sdsu.edu
  //                                   http://www.wideman-one.com
  // Copyright        ©1998 Angus Johnson, Anders Melander & Graham Wideman

  // -----------------------------------------------------------------------------
  // You are free to use this source but please give us credit for our work.
  // If you make improvements or derive new components from this code,
  // we would very much like to see your improvements. FEEDBACK IS WELCOME.
  // -----------------------------------------------------------------------------

  // NOTE:
  // These components use the DropSource.pas unit for the declaration of the
  // TInterfacedComponent class.
  // -----------------------------------------------------------------------------

  // History:
  // dd/mm/yy  Version  Changes
  // --------  -------  ----------------------------------------
  // 16.11.98  3.3      * Changes to TDropBMPSource & TDropBMPTarget modules only.
  // 22.10.98  3.2      * TDropURLTarget moved to another module.
  //                    * TDropTarget.fDataObj moved from private to protected section
  //                      of type declaration.
  // 01.10.98  3.1      * Major design changes including changes to published properties and events.
  //                      (The previous version attempted to unregister the drop target window
  //                      in TDropTarget.Notification method. However, the target TWinControl handle is 
  //                      destroyed prior to this method being called so this was never going to work
  //                      if the target TWinControl was destroyed before TDropTarget. One avenue we 
  //                      investigated was hooking the target TWinControl message handler using its
  //                      WindowProc method. Although this works, if any other component hooks the 
  //                      same TWinControl the order of hooking and unhooking becomes critical. As this 
  //                      is not under the controll of our component this approach has been abandoned.
  //                      The only really safe approach appears to be getting the component user to
  //                      manually unregister the target window prior to the deletion of the target   
  //                      TWinControl. As a design issue, I decided to get the user to manually 
  //                      register the target TWinControl as well as I thought this would be the
  //                      best was to remind of the need to unregister. This small inconvenience
  //                      is far outweighed by the added reliability of the component. It is a simple
  //                      step to register and unregister in the FormCreate and FormDestroy methods
  //                      respectively. If for some reason the component user wishes to temporarily
  //                      disable the TDropTarget capability then the unregister / register methods
  //                      can again be used. The Enabled property has been removed as a consequence.
  //                      The TargetWindow property has also been removed as the Target TWinControl
  //                      is now assigned when passed as a parameter in the register method.)
  //                    * Other design changes now make it MUCH easier to create descendant classes
  //                      of TDropTarget.
  //                    * TDropURLTarget added.
  // 22.09.98  3.0      * Shortcuts (links) for TDropFileTarget now enabled.
  //                    * TDropSource.DoEnumFormatEtc() no longer declared abstract.
  //                    * Bug fix where StgMediums weren't released. (oops!)
  //                    * TDropTarget.GetValidDropEffect() moved to
  //                      protected section and declared virtual.
  //                    * Some bugs still with NT4 :-)
  // 08.09.98  2.0      * Delphi 3 & 4 version - using IDropTarget COM interface.
  // xx.08.97  1.0      * Delphi 2 version - using WM_DROPFILES and DragAcceptFiles().
  // -----------------------------------------------------------------------------

  // BASIC USAGE: (See demo for more detailed examples)
  // 1. Add this non-visual component to the form you wish to drag TO.
  // 2. In the FormCreate method add ... TDropFileTarget1.register(Listview1);
  // 3. In the FormDestroy method add ... TDropFileTarget1.unregister;
  // 4. In the DropTarget OnDrop event handler process the dropped data. Eg ...
  //     procedure TFormURL.DropURLTarget1Drop(Sender: TObject; DragType: TDragType; Point: TPoint);
  //     begin
  //       edit1.text := DropURLTarget1.URL;
  //     end;
  // -----------------------------------------------------------------------------

interface

  uses
    Windows, ActiveX, Classes, Controls, ShlObj, ShellApi, SysUtils,
    ClipBrd, DropSource, Graphics;

  type

  TGetDropEffectEvent = procedure(Sender: TObject;
    const grfKeyState: Longint; var dwEffect: LongInt) of Object;

  TTargetOnEnterEvent = procedure(Sender: TObject; pt: TPoint) of Object;

  TTargetOnDropEvent = procedure(Sender: TObject;
                           DragType: TDragType; Point: TPoint) of Object;

  //Note: TInterfacedComponent declared in DropSource.pas
  TDropTarget = class(TInterfacedComponent, IDropTarget)
  private
    fDragTypes: TDragTypes;
    fRegistered: boolean;
    fTarget: TWinControl;
    fGetDataOnEnter: boolean;
    fOnEnter: TTargetOnEnterEvent;
    fOnDragOver: TTargetOnEnterEvent;
    fOnLeave: TNotifyEvent;
    fOnDrop: TTargetOnDropEvent;
    fGetDropEffectEvent: TGetDropEffectEvent;
  protected
    fDataObj: IDataObject;

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
    function GetValidDropEffect(grfKeyState: Longint): LongInt; Virtual;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure Register(Target:TWinControl);
    procedure Unregister;
  published
    property Dragtypes: TDragTypes read fDragTypes write fDragTypes;
    property GetDataOnEnter: Boolean read fGetDataOnEnter write fGetDataOnEnter;
    property OnEnter: TTargetOnEnterEvent read fOnEnter write fOnEnter;
    property OnDragOver: TTargetOnEnterEvent read fOnDragOver write fOnDragOver;
    property OnLeave: TNotifyEvent read fOnLeave write fOnLeave;
    property OnDrop: TTargetOnDropEvent read fOnDrop write fOnDrop;
    property OnGetDropEffect: TGetDropEffectEvent
      read fGetDropEffectEvent write fGetDropEffectEvent;
  end;


  TDropFileTarget = class(TDropTarget)
  private
    fFiles: TStrings;
  protected
    procedure ClearData; override;
    function DoGetData: boolean; override;
    function HasValidFormats: boolean; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    property Files: TStrings Read fFiles;
  end;

  TDropTextTarget = class(TDropTarget)
  private
    fText: String;
  protected
    procedure ClearData; override;
    function DoGetData: boolean; override;
    function HasValidFormats: boolean; override;
  public
    property Text: String Read fText Write fText;
  end;

const
  HDropFormatEtc: TFormatEtc = (cfFormat: CF_HDROP;
    ptd: nil; dwAspect: DVASPECT_CONTENT; lindex: -1; tymed: TYMED_HGLOBAL);
  TextFormatEtc: TFormatEtc = (cfFormat: CF_TEXT;
    ptd: nil; dwAspect: DVASPECT_CONTENT; lindex: -1; tymed: TYMED_HGLOBAL);

function GetFilesFromHGlobal(const HGlob: HGlobal; var Files: TStrings): boolean;
procedure Register;

implementation

procedure Register;
begin
  RegisterComponents('DragDrop',[TDropFileTarget, TDropTextTarget]);
end;

// -----------------------------------------------------------------------------
//			Miscellaneous functions ...
// -----------------------------------------------------------------------------

//******************* GetFilesFromHGlobal *************************
function GetFilesFromHGlobal(const HGlob: HGlobal; var Files: TStrings): boolean;
var
  DropFiles		: PDropFiles;
  Filename		: PChar;
  s			: string;
begin
  DropFiles := PDropFiles(GlobalLock(HGlob));
  try
    Filename := PChar(DropFiles) + DropFiles^.pFiles;
    while (Filename^ <> #0) do
    begin
      if (DropFiles^.fWide) then // -> NT4 compatability
      begin
        s := WideCharToString(PWideChar(Filename));
        inc(Filename, (Length(s) + 1) * 2);
      end else
      begin
        s := StrPas(Filename);
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

//******************* TDropTarget.DragEnter *************************
function TDropTarget.DragEnter(const dataObj: IDataObject; grfKeyState: Longint;
  pt: TPoint; var dwEffect: Longint): HRESULT;
begin

  ClearData;
  fDataObj := dataObj;
  fDataObj._AddRef;
  dwEffect := GetValidDropEffect(grfKeyState);

  //enum formats here ...
  if HasValidFormats then
    result := S_OK else
    result := E_FAIL;

  if (result <> S_OK) then
  begin
    fDataObj._Release;
    fDataObj := nil;
    dwEffect := DROPEFFECT_NONE;
    exit;
  end;

  //It's generally more efficient to get files only if a drop occurs
  //rather than on entering a potential target window.
  //However - sometimes there is a good reason to get them here - see Demo.
  if fGetDataOnEnter and (not DoGetData) then
    dwEffect := DROPEFFECT_NONE;

  if Assigned(fOnEnter) then
    fOnEnter(self, pt);
end;

//******************* TDropTarget.DragOver *************************
function TDropTarget.DragOver(grfKeyState: Longint; pt: TPoint;
             var dwEffect: Longint): HResult;
begin
  //Keep code in this event to a minimum as this is called very often.
  dwEffect := GetValidDropEffect(grfKeyState);
  if Assigned(fOnDragOver) then
    fOnDragOver(self, pt);
  RESULT := S_OK;
end;

//******************* TDropTarget.DragLeave *************************
function TDropTarget.DragLeave: HResult;
begin
  ClearData;
  if fDataObj <> nil then
  begin
    fDataObj._Release;
    fDataObj := nil;
  end;
  if Assigned(fOnLeave) then fOnLeave(self);
  Result := S_OK;
end;

//******************* TDropTarget.Drop *************************
function TDropTarget.Drop(const dataObj: IDataObject; grfKeyState: Longint;
  pt: TPoint; var dwEffect: Longint): HResult;
begin
  RESULT := S_OK;
  dwEffect := GetValidDropEffect(grfKeyState);

  if (not fGetDataOnEnter) and (not DoGetData) then
    dwEffect := DROPEFFECT_NONE;

  if Assigned(fOnDrop) then
    case dwEffect of
      DROPEFFECT_MOVE: fOnDrop(Self, dtMove, pt);
      DROPEFFECT_COPY: fOnDrop(Self, dtCopy, pt);
      DROPEFFECT_LINK: fOnDrop(Self, dtLink, pt);
    end;

  // clean up!
  ClearData;
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
  else if (dtLink in fDragTypes) then result := DROPEFFECT_LINK
  else result := DROPEFFECT_NONE;
  //Default behaviour can be overridden (see Demo).
  if Assigned(fGetDropEffectEvent) then fGetDropEffectEvent(self, grfKeyState, result);
end;


//******************* TDropTarget.Create *************************
constructor TDropTarget.Create( AOwner: TComponent );
begin
   inherited Create( AOwner );
   _AddRef;
   DragTypes := [dtCopy, dtMove, dtLink];
   fGetDataOnEnter := false;
   fDataObj := nil;
end;

//******************* TDropTarget.Destroy *************************
destructor TDropTarget.Destroy;
begin
  Unregister;
  inherited Destroy;
end;

//******************* TDropTarget.RegisterTarget *************************
procedure TDropTarget.Register(Target: TWinControl);
begin
  if fTarget = Target then
    exit;
  if (fTarget <> nil) then
    Unregister;
  fTarget := target;

  CoLockObjectExternal(self as IUnknown,true,false);
  if not RegisterDragDrop(fTarget.handle,self as IDroptarget) = S_OK then
      raise Exception.create('Failed to Register '+ fTarget.name);
  fRegistered := true;
end;

//******************* TDropTarget.UnregisterTarget *************************
procedure TDropTarget.Unregister;
begin
  fRegistered := false;
  if (fTarget = nil) then
    exit;
  if not RevokeDragDrop(fTarget.handle) = S_OK then
      raise Exception.create('Failed to Unregister '+ fTarget.name);
  CoLockObjectExternal(self as IUnknown,false,false);
  fTarget := nil;
end;

// -----------------------------------------------------------------------------
//			TDropFileTarget
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

//******************* TDropFileTarget.HasValidFormats *************************
function TDropFileTarget.HasValidFormats: boolean;
begin
  result := (fDataObj.QueryGetData(HDropFormatEtc) = S_OK);
end;

//******************* TDropFileTarget.ClearData *************************
procedure TDropFileTarget.ClearData;
begin
  fFiles.clear;
end;

//******************* TDropFileTarget.DoGetData *************************
function TDropFileTarget.DoGetData: boolean;
var
  medium: TStgMedium;
begin
  fFiles.clear;
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
end;


// -----------------------------------------------------------------------------
//			TDropTextTarget
// -----------------------------------------------------------------------------

//******************* TDropTextTarget.HasValidFormats *************************
function TDropTextTarget.HasValidFormats: boolean;
begin
  result := (fDataObj.QueryGetData(TextFormatEtc) = S_OK);
end;

//******************* TDropTextTarget.ClearData *************************
procedure TDropTextTarget.ClearData;
begin
  fText := '';
end;

//******************* TDropTextTarget.DoGetData *************************
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
// -----------------------------------------------------------------------------

{ // Done in DropSource...
initialization
  OleInitialize(nil);
finalization
  OleUnInitialize;
}
end.
