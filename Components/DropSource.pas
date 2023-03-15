UNIT dropsource;
  
  // -----------------------------------------------------------------------------
  // Project:         Drag and Drop Source Components
  // Component Names: TDropTextSource, TDropFileSource
  // Module:          DropSource
  // Description:     Implements Dragging & Dropping of text and files
  //                  FROM your application TO another.
  // Version:	        3.0
  // Date:            22-SEP-1998
  // Target:          Win32, Delphi 3 & 4
  // Authors:         Angus Johnson, ajohnson@rpi.net.au
  //                  Anders Melander, anders@melander.dk
  //                                   http://www.melander.dk
  // Copyright        ©1998 Angus Johnson & Anders Melander
  // -----------------------------------------------------------------------------
  // You are free to use this source but please give us credit for our work.
  // If you make improvements or derive new components from this code,
  // we would very much like to see your improvements. FEEDBACK IS WELCOME.
  // -----------------------------------------------------------------------------
  
  // History:
  // dd/mm/yy  Version  Changes
  // --------  -------  ----------------------------------------
  // 22.09.98  3.0      * Shortcuts (links) for TDropFileSource now enabled.
  //                    * Scrap files for TDropTextSource now enabled.
  //                    * TDropSource.DoEnumFormatEtc() no longer declared abstract.
  //                    * Some bugs still with NT4 :-)
  // 08.09.98  2.0      * No significant changes to this module
  //                      but the version was updated to coincide with the
  //                      new DropTarget module included with this demo.
  // 31.08.98  1.5      * Fixed a Delphi 4 bug!
  //                      (I cut and pasted the wrong line!)
  //                    * Demo code now MUCH tidier and easier to read (I think).
  // 19.08.98  1.4      * CopyToClipboard method added.
  //                    * Should now compile in Delphi 4. (see below)
  //                    * Another tidy up of the code.
  // 21.07.98  1.3      * Fixed a bug in Ver 1.2 where OnDrop event was never called.
  //                    * Now able to drag text to WordPad.
  //                    * Added OnFeedback Event.
  //                    * Added dtLink to TDragType enumeration but
  //                      still not able to get it to work.
  //                    * Code tidy up.
  // 19.07.98  1.2      * Changed TDragType enumeration type to a
  //                      TDragTypes set (AM's suggestion) ready to
  //                      implement dtLink later.
  //                    * Added known bugs to header and demo.
  // 17.07.98  1.1      * Reenabled end-user option to select either
  //                      Copy or Move operation while dragging.
  // 15.07.98  1.0      * Initial Delphi 3 component implementation of AM's
  //                      DropSource unit. I released a Delphi 2 D'n'D
  //                      component (TDragFilesSrc) over 12 months ago. However
  //                      with the significant changes in COM implementations
  //                      between Delphi versions I decided to use AM's code
  //                      as the springboard for my new Delphi 3 D'n'D components.
  //                      Thanks to Anders for the excellent start and
  //                      suggestions along the way!
  //
  // -----------------------------------------------------------------------------

  // TDropTextSource -
  //   Public
  //      ....
  //      Text: string;
  //      function Execute: TDragResult; //drDropCopy, drDropMove, drDropCancel ...
  //      function CopyToClipboard: boolean;
  //    published
  //      property DragTypes: TDragTypes;  // [dtCopy, dtMove, dtLink]
  //      property OnDrop: TDropEvent;
  //      property OnFeedback: TFeedbackEvent;
  //      ....

  // TDropFileSource -
  //   Public
  //      ....
  //      property Files: TStrings;
  //      function Execute: TDragResult; //drDropCopy, drDropMove, drDropCancel ...
  //      function CopyToClipboard: boolean;
  //    published
  //      property DragTypes: TDragTypes;  // [dtCopy, dtMove]
  //      property OnDrop: TDropEvent;
  //      property OnFeedback: TFeedbackEvent;
  //      ....

  //
  // Very brief examples of usage (see included demo for more detailed examples):
  //
  // TDropTextSource -
  // DropFileSource1.DragTypes := [dtCopy];
  // DropTextSource1.text := edit1.text;
  // if DropTextSource1.execute = drDropCopy then ShowMessage('It worked!');
  //
  // TDropFileSource -
  // DropFileSource1.DragTypes := [dtCopy, dtMove];
  // //  ie: let the user decide - Copy or Move.
  // //      Hold Ctrl down during drag -> Copy
  // //      Hold Shift down during drag -> Move
  // DropFileSource1.files.clear;
  // DropFileSource1.files.add('c:\autoexec.bat');
  // DropFileSource1.files.add('c:\config.sys');
  // res := DropFileSource1.execute;
  // if res = drDropCopy then ShowMessage('Files Copied!')
  // else if res = drDropMove then ShowMessage('Files Moved!');
  // -----------------------------------------------------------------------------

INTERFACE
  USES
    Windows, ActiveX, Classes, ShlObj, SysUtils, ClipBrd;

  CONST
    MaxFormats = 20;

  TYPE

    TInterfacedComponent = CLASS(TComponent, IUnknown)
    Private
      fRefCount: Integer;
    Protected
      FUNCTION QueryInterface(CONST IID: TGuid; OUT Obj): HRESULT;
                 {$ifdef VER110} reintroduce; {$endif} StdCall;
      FUNCTION _AddRef: Integer; StdCall;
      FUNCTION _Release: Integer; StdCall;
    Public
       PROPERTY RefCount: Integer Read fRefCount;
    END;

    TDragType = (dtCopy, dtMove, dtLink);
    TDragTypes = SET OF TDragType;

    TDragResult = (drDropCopy, drDropMove, drDropLink, drCancel, drOutMemory, drUnknown);

    TDropEvent = PROCEDURE(Sender: TObject; DragType: TDragType;
                 VAR ContinueDrop: Boolean) OF Object;
    TFeedbackEvent = PROCEDURE(Sender: TObject; Effect: LongInt) OF Object;

  TDropSource = CLASS(TInterfacedComponent, IDropSource, IDataObject)
  Private
    fDragTypes: TDragTypes;
    FeedbackEffect: LongInt;
    fDropEvent: TDropEvent;
    fFBEvent: TFeedBackEvent;
    fDataFormats: array[0..MaxFormats-1] of TFormatEtc;
    DataFormatsCount: integer;
  Protected
    // IDropSource implementation

    FUNCTION QueryContinueDrag(fEscapePressed: bool; grfKeyState: LongInt): HRESULT; StdCall;
    FUNCTION GiveFeedback(dwEffect: LongInt): HRESULT; StdCall;

    // IDataObject implementation
    FUNCTION GetData(CONST FormatEtcIn: TFormatEtc; OUT Medium: TStgMedium):HRESULT; StdCall;
    FUNCTION GetDataHere(CONST FormatEtc: TFormatEtc; OUT Medium: TStgMedium):HRESULT; StdCall;
    FUNCTION QueryGetData(CONST FormatEtc: TFormatEtc): HRESULT; StdCall;
    FUNCTION GetCanonicalFormatEtc(CONST FormatEtc: TFormatEtc;
             OUT FormatEtcOut: TFormatEtc): HRESULT; StdCall;
    FUNCTION SetData(CONST FormatEtc: TFormatEtc; VAR Medium: TStgMedium;
             fRelease: Bool): HRESULT; StdCall;
    FUNCTION EnumFormatEtc(dwDirection: LongInt; OUT EnumFormatEtc: IEnumFormatEtc): HRESULT; StdCall;
    FUNCTION dAdvise(CONST FormatEtc: TFormatEtc; advf: LongInt;
             CONST advsink: IAdviseSink; OUT dwConnection: LongInt): HRESULT; StdCall;
    FUNCTION dUnadvise(dwConnection: LongInt): HRESULT; StdCall;
    FUNCTION EnumdAdvise(OUT EnumAdvise: IEnumStatData): HRESULT; StdCall;

    //New functions...
    FUNCTION DoGetData(CONST FormatEtcIn: TFormatEtc; OUT Medium: TStgMedium):HRESULT; Virtual;
    FUNCTION DoGetDataHere(CONST FormatEtc: TFormatEtc; OUT Medium: TStgMedium):HRESULT; Virtual;
    FUNCTION DoQueryGetData(CONST FormatEtc: TFormatEtc): HRESULT; Virtual;
    FUNCTION DoEnumFormatEtc(dwDirection: LongInt; OUT EnumFormatEtc: IEnumFormatEtc): HRESULT; Virtual;

  Public
    CONSTRUCTOR Create(aowner: TComponent); Override;
    FUNCTION Execute: TDragResult;
    FUNCTION CopyToClipboard: boolean; Virtual;
  Published
     PROPERTY Dragtypes: TDragTypes Read fDragTypes Write fDragTypes;
     PROPERTY OnDrop: TDropEvent Read fDropEvent Write fDropEvent;
     PROPERTY OnFeedback: TFeedbackEvent Read fFBEvent Write fFBEvent;
  END;

  TDropTextSource = CLASS(TDropSource)
  Private
    fText: String;
  Protected
    FUNCTION DoGetData(CONST FormatEtcIn: TFormatEtc; OUT Medium: TStgMedium):HRESULT; Override;
    FUNCTION DoGetDataHere(CONST FormatEtc: TFormatEtc; OUT Medium: TStgMedium):HRESULT; Override;
  Public
    CONSTRUCTOR Create(aOwner: TComponent); Override;
    FUNCTION CopyToClipboard: boolean; Override;
    PROPERTY Text: String Read fText Write fText;
  END;

  TDropFileSource = CLASS(TDropSource)
  Private
    fFiles: TStrings;
  Protected
    FUNCTION DoGetData(CONST FormatEtcIn: TFormatEtc; OUT Medium: TStgMedium):HRESULT; Override;
    FUNCTION DoGetDataHere(CONST FormatEtc: TFormatEtc; OUT Medium: TStgMedium):HRESULT; Override;
  Public
    CONSTRUCTOR Create(aOwner: TComponent); Override;
    DESTRUCTOR Destroy; Override;
    FUNCTION CopyToClipboard: boolean; Override;
    PROPERTY Files: TStrings Read fFiles;
  END;

  PROCEDURE Register;

IMPLEMENTATION

  // -----------------------------------------------------------------------------
  //			Miscellaneous declarations and functions.
  // -----------------------------------------------------------------------------

  TYPE
    TMyDropFiles = PACKED RECORD
      dropfiles: TDropFiles;
      filenames: ARRAY[1..1] OF Char;
    END;
    pMyDropFiles = ^TMyDropFiles;


  //******************* GetSizeOfPidl *************************
  function GetSizeOfPidl(pidl: pItemIDList): integer;
  var
    ptr: PByte;
    pSHIt: pSHItemID;
    i: integer;
  begin
    result := 2;
    ptr := pointer(pidl);
    repeat
      pSHIt := pointer(ptr);
      i := pSHIt^.cb;
      inc(result,i);
      inc(ptr,i);
    until i = 0;
  end;

  //******************* FreePidl *************************
  procedure FreePidl(pidl: pItemIDList);
  var
    ShellMalloc: IMalloc;
  begin
    SHGetMalloc(ShellMalloc);
    ShellMalloc.free(pidl);
  end;

  //******************* GetShellFolderOfPath *************************
  function GetShellFolderOfPath(FolderPath: TFileName): IShellFolder;
  var
    DeskTopFolder: IShellFolder;
    PathPidl: pItemIDList;
    OlePath: Array[0..MAX_PATH] of WideChar;
    dummy,pdwAttributes: ULONG;
  begin
    result := nil;
    StringToWideChar( FolderPath, OlePath, MAX_PATH );
    try
      If not (SHGetDesktopFolder(DeskTopFolder) = NOERROR) then exit;
      if (DesktopFolder.ParseDisplayName(0,
            nil,OlePath,dummy,PathPidl,pdwAttributes) = NOERROR) and
            (pdwAttributes and SFGAO_FOLDER <> 0) then
        DesktopFolder.BindToObject(PathPidl,nil,IID_IShellFolder,pointer(result));
      FreePidl(PathPidl);
    except
    end;
  end;

  //******************* GetFullPIDLFromPath *************************
  function GetFullPIDLFromPath(Path: TFileName): pItemIDList;
  var
     DeskTopFolder: IShellFolder;
     OlePath: Array[0..MAX_PATH] of WideChar;
     dummy1,dummy2: ULONG;
  begin
    result := nil;
    StringToWideChar( Path, OlePath, MAX_PATH );
    try
      If (SHGetDesktopFolder(DeskTopFolder) = NOERROR) then
        DesktopFolder.ParseDisplayName(0,nil,OlePath,dummy1,result,dummy2);
    except
    end;
  end;

  //******************* GetSubPidl *************************
  function GetSubPidl(Folder: IShellFolder; Sub: TFilename): pItemIDList;
  var
    dummy1,dummy2: ULONG;
    OleFile: Array[0..MAX_PATH] of WideChar;
  begin
    result := nil;
    try
      StringToWideChar( Sub, OleFile, MAX_PATH );
      Folder.ParseDisplayName(0,nil,OleFile,dummy1,result,dummy2);
    except
    end;
  end;

  //See "Clipboard Formats for Shell Data Transfers" in Ole.hlp...
  //(Needed to drag links (shortcuts).)

  //******************* ConvertFilesToShellIDList *************************
  type
    POffsets = ^TOffsets;
    TOffsets = array[0..$FFFF] of UINT;

  function ConvertFilesToShellIDList(path: string; files: TStrings): HGlobal;
  var
    shf: IShellFolder;
    PathPidl, pidl: pItemIDList;
    Ida: PIDA;
    pOffset: POffsets;
    ptrByte: ^Byte;
    i, PathPidlSize, IdaSize, PreviousPidlSize: integer;
  begin
    result := 0;
    shf := GetShellFolderOfPath(path);
    if shf = nil then exit;
    //Calculate size of IDA structure ...
    IdaSize := (files.count + 2) * sizeof(UINT);

    //Add to IdaSize space for ALL pidls...
    PathPidl := GetFullPIDLFromPath(path);
    PathPidlSize := GetSizeOfPidl(PathPidl);
    IdaSize := IdaSize + PathPidlSize;
    for i := 0 to files.count-1 do
    begin
      pidl := GetSubPidl(shf,files[i]);
      IdaSize := IdaSize + GetSizeOfPidl(Pidl);
      FreePidl(pidl);
    end;

    //Allocate memory...
    Result := GlobalAlloc(GMEM_SHARE OR GMEM_ZEROINIT, IdaSize);
    IF (Result = 0) THEN
    BEGIN
      FreePidl(PathPidl);
      Exit;
    END;
    Ida := GlobalLock(Result);
    FillChar(Ida^,IdaSize,0);

    //Fill in offset and pidl data...
    Ida^.cidl := files.count; //cidl = file count
    pOffset := @(Ida^.aoffset); //otherwise I would have to turn off range checking.
    pOffset^[0] := (files.count+2)*sizeof(UINT); //offset of Path pidl

    ptrByte := pointer(Ida);
    inc(ptrByte,pOffset^[0]); //ptrByte now points to Path pidl
    move(PathPidl^, ptrByte^, PathPidlSize); //copy path pidl
    FreePidl(PathPidl);

    PreviousPidlSize := PathPidlSize;
    for i := 1 to files.count do
    begin
      pidl := GetSubPidl(shf,files[i-1]);
      pOffset^[i] := pOffset^[i-1] + UINT(PreviousPidlSize); //offset of pidl
      PreviousPidlSize := GetSizeOfPidl(Pidl);

      ptrByte := pointer(Ida);
      inc(ptrByte,pOffset^[i]); //ptrByte now points to current file pidl
      move(Pidl^, ptrByte^, PreviousPidlSize); //copy file pidl
                            //PreviousPidlSize = current pidl size here
      FreePidl(pidl);
    end;
    GlobalUnLock(Result);
  end;


  //******************* Register *************************
  PROCEDURE Register;
  BEGIN
    RegisterComponents('Samples', [TDropFileSource,TDropTextSource]);
  END;


  // -----------------------------------------------------------------------------
  //			TInterfacedComponent
  // -----------------------------------------------------------------------------

  // QueryInterface now returns HRESULT so should now compile in Delphi 4
  // as well. Thanks to 'Scotto the Unwise' - scottos@gtcom.net
  //******************* TInterfacedComponent.QueryInterface *************************
  FUNCTION TInterfacedComponent.QueryInterface(CONST IID: TGuid; OUT Obj): HRESULT;
  BEGIN
    IF GetInterface(IID, Obj) THEN result := 0 ELSE result := E_NOINTERFACE;
  END;

  //******************* TInterfacedComponent._AddRef *************************
  FUNCTION TInterfacedComponent._AddRef: Integer;
  BEGIN
    Inc(fRefCount);
    result := fRefCount;
  END;

  //******************* TInterfacedComponent._Release *************************
  FUNCTION TInterfacedComponent._Release: Integer;
  BEGIN
    Dec(fRefCount);
    IF fRefCount = 0 THEN
    BEGIN
      Destroy;
      result := 0;
      Exit;
    END;
    result := fRefCount;
  END;

  // -----------------------------------------------------------------------------
  //			TEnumFormatEtc
  // -----------------------------------------------------------------------------
  { TEnumFormatEtc - format enumerator for TDataObject }

  TYPE
    pFormatList = ^TFormatList;
    TFormatList = ARRAY[0..255] OF TFormatEtc;

  TEnumFormatEtc = CLASS(TInterfacedObject, IEnumFormatEtc)
  Private
    fFormatList: pFormatList;
    fFormatCount: Integer;
    fIndex: Integer;
  Public
    CONSTRUCTOR Create(FormatList: pFormatList; FormatCount, Index: Integer);
    { IEnumFormatEtc }
    FUNCTION Next(Celt: LongInt; OUT Elt; pCeltFetched: pLongInt): HRESULT; StdCall;
    FUNCTION Skip(Celt: LongInt): HRESULT; StdCall;
    FUNCTION Reset: HRESULT; StdCall;
    FUNCTION Clone(OUT Enum: IEnumFormatEtc): HRESULT; StdCall;
  END;

//******************* TEnumFormatEtc.Create *************************
  CONSTRUCTOR TEnumFormatEtc.Create(FormatList: pFormatList;
              FormatCount, Index: Integer);
  BEGIN
    INHERITED Create;
    fFormatList := FormatList;
    fFormatCount := FormatCount;
    fIndex := Index;
  END;

  //******************* TEnumFormatEtc.Next *************************
  FUNCTION TEnumFormatEtc.Next(Celt: LongInt; OUT Elt; pCeltFetched: pLongInt): HRESULT;
  VAR
    i: Integer;
  BEGIN
    i := 0;
    WHILE (i < Celt) AND (fIndex < fFormatCount) DO
    BEGIN
      TFormatList(Elt)[i] := fFormatList[fIndex];
      Inc(fIndex);
      Inc(i);
    END;
    IF pCeltFetched <> NIL THEN pCeltFetched^ := i;
    IF i = Celt THEN result := S_OK ELSE result := S_FALSE;
  END;

  //******************* TEnumFormatEtc.Skip *************************
  FUNCTION TEnumFormatEtc.Skip(Celt: LongInt): HRESULT;
  BEGIN
    IF Celt <= fFormatCount - fIndex THEN
    BEGIN
      fIndex := fIndex + Celt;
      result := S_OK;
    END ELSE
    BEGIN
      fIndex := fFormatCount;
      result := S_FALSE;
    END;
  END;

  //******************* TEnumFormatEtc.Reset *************************
  FUNCTION TEnumFormatEtc.ReSet: HRESULT;
  BEGIN
    fIndex := 0;
    result := S_OK;
  END;

  //******************* TEnumFormatEtc.Clone *************************
  FUNCTION TEnumFormatEtc.Clone(OUT Enum: IEnumFormatEtc): HRESULT;
  BEGIN
    enum := TEnumFormatEtc.Create(fFormatList, fFormatCount, fIndex);
    result := S_OK;
  END;

  // -----------------------------------------------------------------------------
  //			TDropSource
  // -----------------------------------------------------------------------------

  //******************* TDropSource.Create *************************
  CONSTRUCTOR TDropSource.Create(aOwner: TComponent);
  BEGIN
    INHERITED Create(aOwner);
    DragTypes := [dtCopy]; //default to Copy.
    //To avoid premature release ...
    _AddRef;
    DataFormatsCount := 0;
  END;

  //******************* TDropSource.Execute *************************
  FUNCTION TDropSource.Execute: TDragResult;
  VAR
    res: HRESULT;
    okeffect: LongInt;
    effect: LongInt;
  BEGIN
    result := drUnknown;
    okeffect := DROPEFFECT_NONE;
    IF dtCopy IN fDragTypes THEN okeffect := okeffect + DROPEFFECT_COPY;
    IF dtMove IN fDragTypes THEN okeffect := okeffect + DROPEFFECT_MOVE;
    IF dtLink IN fDragTypes THEN okeffect := okeffect + DROPEFFECT_LINK;

    res := DoDragDrop(Self AS IDataObject, Self AS IDropSource, okeffect, effect);
    CASE res OF
      DRAGDROP_S_DROP:   BEGIN
                           IF (okeffect AND effect <> 0) THEN
                           BEGIN
                             IF (effect AND DROPEFFECT_COPY <> 0) THEN
                               result := drDropCopy
                             ELSE IF (effect AND DROPEFFECT_MOVE <> 0) THEN
                               result := drDropMove
                             ELSE result := drDropLink;
                           END ELSE
                             result := drCancel;
                         END;
      DRAGDROP_S_CANCEL: result := drCancel;
      E_OUTOFMEMORY:     result := drOutMemory;
    END;
  END;

  //******************* TDropSource.CopyToClipboard *************************
  FUNCTION TDropSource.CopyToClipboard: boolean;
  BEGIN
    result := false;
  END;

  //******************* TDropSource.QueryContinueDrag *************************
  FUNCTION TDropSource.QueryContinueDrag(fEscapePressed: bool;
    grfKeyState: LongInt): HRESULT; StdCall;
  VAR
    ContinueDrop: Boolean;
    dragtype: TDragType;
  BEGIN
    IF fEscapePressed THEN
      result := DRAGDROP_S_CANCEL
    // will now allow drag and drop with either mouse button.
    ELSE IF (grfKeyState AND (MK_LBUTTON OR MK_RBUTTON) = 0) THEN
    BEGIN
      ContinueDrop := True;
      dragtype := dtCopy;
      IF (FeedbackEffect AND DROPEFFECT_COPY <> 0) THEN // DragType = dtCopy
      ELSE IF (FeedbackEffect AND DROPEFFECT_MOVE <> 0) THEN dragtype := dtMove
      ELSE IF (FeedbackEffect AND DROPEFFECT_LINK <> 0) THEN dragtype := dtLink
      ELSE ContinueDrop := False;

      //if a valid drop then do OnDrop event if assigned...
      IF ContinueDrop AND
        (((dragtype = dtCopy) AND (dtCopy IN dragtypes)) OR
        ((dragtype = dtMove) AND (dtMove IN dragtypes)) OR
        ((dragtype = dtLink) AND (dtLink IN dragtypes))) AND
        Assigned(OnDrop) THEN OnDrop(Self, dragtype, ContinueDrop);

      IF ContinueDrop THEN result := DRAGDROP_S_DROP
      ELSE result := DRAGDROP_S_CANCEL;
    END ELSE
      result := NOERROR;
  END;

  //******************* TDropSource.GiveFeedback *************************
  FUNCTION TDropSource.GiveFeedback(dwEffect: LongInt): HRESULT; StdCall;
  BEGIN
    FeedbackEffect := dwEffect;

    //NB: Use the OnFeedback event sparingly as it will effect performance...
    IF Assigned(OnFeedback) THEN OnFeedback(Self, dwEffect);

    result:=DRAGDROP_S_USEDEFAULTCURSORS;

  END;

  //******************* TDropSource.DoQueryGetData *************************
  FUNCTION TDropSource.DoQueryGetData(CONST FormatEtc: TFormatEtc): HRESULT;
  VAR
    i: integer;
  BEGIN
    result:= S_OK;
    for i := 0 to DataFormatsCount-1 do
      with fDataFormats[i] do
      begin
        IF (FormatEtc.cfFormat = cfFormat) and
           (FormatEtc.dwAspect = dwAspect) and
           (FormatEtc.tymed AND tymed <> 0) THEN exit;
      end;
    result:= E_INVALIDARG;
  END;

  //******************* TDropSource.DoEnumFormatEtc *************************
  FUNCTION TDropSource.DoEnumFormatEtc(dwDirection: LongInt;
           OUT EnumFormatEtc:IEnumFormatEtc): HRESULT;
  BEGIN
    IF (dwDirection = DATADIR_GET) THEN
    BEGIN
      EnumFormatEtc :=
        TEnumFormatEtc.Create(@fDataFormats, DataFormatsCount, 0);
      result := S_OK;
    END ELSE IF (dwDirection = DATADIR_SET) THEN
      result := E_NOTIMPL
    ELSE result := E_INVALIDARG;
  END;

  //******************* TDropSource.GetCanonicalFormatEtc *************************
  FUNCTION TDropSource.GetCanonicalFormatEtc(CONST FormatEtc: TFormatEtc;
           OUT FormatEtcOut: TFormatEtc): HRESULT;
  BEGIN
    result := DATA_S_SAMEFORMATETC;
  END;

  //******************* TDropSource.SetData *************************
  FUNCTION TDropSource.SetData(CONST FormatEtc: TFormatEtc; VAR Medium: TStgMedium;
           fRelease: Bool): HRESULT;
  BEGIN
    result := E_NOTIMPL;
  END;

  //******************* TDropSource.DAdvise *************************
  FUNCTION tdropsource.DAdvise(CONST FormatEtc: TFormatEtc; advf: LongInt;
           CONST advSink: IAdviseSink; OUT dwConnection: LongInt): HRESULT;
  BEGIN
    result := OLE_E_ADVISENOTSUPPORTED;
  END;

  //******************* TDropSource.DUnadvise *************************
  FUNCTION TDropSource.DUnadvise(dwConnection: LongInt): HRESULT;
  BEGIN
    result := OLE_E_ADVISENOTSUPPORTED;
  END;

  //******************* TDropSource.EnumDAdvise *************************
  FUNCTION tdropsource.EnumDAdvise(OUT EnumAdvise: IEnumStatData): HRESULT;
  BEGIN
    result := OLE_E_ADVISENOTSUPPORTED;
  END;

  //******************* TDropSource.GetData *************************
  FUNCTION tdropsource.GetData(CONST FormatEtcIn: TFormatEtc; OUT Medium: TStgMedium):HRESULT; StdCall;
  BEGIN
    result := DoGetData(FormatEtcIn, Medium);
  END;

  //******************* TDropSource.GetDataHere *************************
  FUNCTION TDropSource.GetDataHere(CONST FormatEtc: TFormatEtc;
           OUT Medium: TStgMedium):HRESULT; StdCall;
  BEGIN
    result := DoGetDataHere(FormatEtc, Medium);
  END;

  //******************* TDropSource.QueryGetData *************************
  FUNCTION TDropSource.QueryGetData(CONST FormatEtc: TFormatEtc): HRESULT; StdCall;
  BEGIN
    result := DoQueryGetData(FormatEtc);
  END;

  //******************* TDropSource.EnumFormatEtc *************************
  FUNCTION TDropSource.EnumFormatEtc(dwDirection: LongInt;
           OUT EnumFormatEtc:IEnumFormatEtc): HRESULT; StdCall;
  BEGIN
    result := DoEnumFormatEtc(dwDirection, EnumFormatEtc);
  END;

  //******************* TDropSource.DoGetData *************************
  FUNCTION TDropSource.DoGetData(CONST FormatEtcIn: TFormatEtc; OUT Medium: TStgMedium):HRESULT;
  BEGIN
    result := DV_E_FORMATETC;
  END;

  //******************* TDropSource.DoGetDataHere *************************
  FUNCTION TDropSource.DoGetDataHere(CONST FormatEtc: TFormatEtc;
           OUT Medium: TStgMedium):hresult;
  BEGIN
    result := DV_E_FORMATETC;
  END;

  // -----------------------------------------------------------------------------
  //			TDropTextSource
  // -----------------------------------------------------------------------------

  var
    CF_FILEGROUPDESCRIPTOR, CF_FILECONTENTS: UINT;

  //******************* TDropTextSource.Create *************************
  CONSTRUCTOR TDropTextSource.Create(aOwner: TComponent);
  BEGIN
    INHERITED Create(aOwner);
    fText := '';
    CF_FILECONTENTS := RegisterClipboardFormat(CFSTR_FILECONTENTS);
    CF_FILEGROUPDESCRIPTOR := RegisterClipboardFormat(CFSTR_FILEDESCRIPTOR);

    fDataFormats[0].cfFormat := CF_TEXT;
    fDataFormats[0].ptd := NIL;
    fDataFormats[0].dwAspect := DVASPECT_CONTENT;
    fDataFormats[0].lIndex := -1;
    fDataFormats[0].tymed := TYMED_HGLOBAL;

    fDataFormats[1].cfFormat := CF_FILEGROUPDESCRIPTOR;
    fDataFormats[1].ptd := NIL;
    fDataFormats[1].dwAspect := DVASPECT_CONTENT;
    fDataFormats[1].lIndex := -1;
    fDataFormats[1].tymed := TYMED_HGLOBAL;

    fDataFormats[2].cfFormat := CF_FILECONTENTS;
    fDataFormats[2].ptd := NIL;
    fDataFormats[2].dwAspect := DVASPECT_CONTENT;
    fDataFormats[2].lIndex := -1;
    fDataFormats[2].tymed := TYMED_HGLOBAL;

    DataFormatsCount := 3;
  END;

  // Adapted from Zbysek Hlinka, zhlinka@login.cz.
  //******************* TDropTextSource.CopyToClipboard *************************
  FUNCTION TDropTextSource.CopyToClipboard: boolean;
  VAR
    FormatEtcIn: TFormatEtc;
    Medium: TStgMedium;
  BEGIN
    FormatEtcIn.cfFormat := CF_TEXT;
    FormatEtcIn.dwAspect := DVASPECT_CONTENT;
    FormatEtcIn.tymed := TYMED_HGLOBAL;
    IF fText = '' then result := false
    ELSE IF GetData(formatetcIn,Medium) = S_OK THEN
    BEGIN
      Clipboard.SetAsHandle(CF_TEXT,Medium.hGlobal);
      result := true;
    END ELSE result := false;
  END;

  //******************* TDropTextSource.DoGetDataHere *************************
  FUNCTION TDropTextSource.DoGetDataHere(CONST FormatEtc: TFormatEtc; OUT Medium: TStgMedium):HRESULT;
  VAR
    pText: PChar;
  BEGIN
    IF (FormatEtc.cfFormat = CF_TEXT) AND
      (FormatEtc.dwAspect = DVASPECT_CONTENT) AND
      (FormatEtc.tymed = TYMED_HGLOBAL) AND (Medium.tymed = TYMED_HGLOBAL) THEN
    BEGIN
      IF (Medium.hGlobal = 0) THEN
      BEGIN
        result := E_OUTOFMEMORY;
        Exit;
      END;
      pText := PChar(GlobalLock(Medium.hGlobal));
      TRY
        StrCopy(pText, PChar(fText));
      FINALLY
        GlobalUnlock(Medium.hGlobal);
      END;
      Medium.UnkForRelease := NIL;
      result := S_OK;
    END ELSE
      result := INHERITED DoGetDataHere(FormatEtc, Medium);
  END;

  // Adapted from stefc@fabula.com
  //******************* TDropTextSource.DoGetData *************************
  FUNCTION TDropTextSource.DoGetData(CONST FormatEtcIn: TFormatEtc; OUT Medium: TStgMedium):HRESULT;
  var
    pFGD: PFileGroupDescriptor;
    pText: PChar;
  BEGIN
    //result := E_FAIL;

    Medium.tymed := 0;
    Medium.UnkForRelease := NIL;
    Medium.hGlobal := 0;

    IF (FormatEtcIn.cfFormat = CF_TEXT) AND
      (FormatEtcIn.dwAspect = DVASPECT_CONTENT) AND
      (FormatEtcIn.tymed AND TYMED_HGLOBAL <> 0) THEN
    BEGIN
      Medium.hGlobal := GlobalAlloc(GMEM_SHARE OR GHND, Length(fText)+1);
      IF (Medium.hGlobal = 0) THEN
      BEGIN
        result := E_OUTOFMEMORY;
        Exit;
      END;
      medium.tymed := TYMED_HGLOBAL;

      result := DoGetDataHere(FormatEtcIn, Medium);
      //if RenderTextAsHGlobal(fText,Medium.hGlobal) then result := S_OK;
    END
    ELSE IF (FormatEtcIn.cfFormat = CF_FILEGROUPDESCRIPTOR) AND
      (FormatEtcIn.dwAspect = DVASPECT_CONTENT) AND
      (FormatEtcIn.tymed AND TYMED_HGLOBAL <> 0) THEN
    BEGIN
      Medium.hGlobal := GlobalAlloc(GMEM_SHARE OR GHND, SizeOf(TFileGroupDescriptor));
      IF (Medium.hGlobal = 0) THEN
      BEGIN
        result := E_OUTOFMEMORY;
        Exit;
      END;
      medium.tymed := TYMED_HGLOBAL;
      pFGD := pointer(GlobalLock(Medium.hGlobal));
      with pFGD^ do
      begin
        cItems := 1;
        fgd[0].dwFlags := FD_LINKUI;
        fgd[0].cFileName := 'Text Scrap File.txt';
      end;
      GlobalUnlock(Medium.hGlobal);
      result := S_OK;
    END
    ELSE IF (FormatEtcIn.cfFormat = CF_FILECONTENTS) AND
      (FormatEtcIn.dwAspect = DVASPECT_CONTENT) AND
      (FormatEtcIn.tymed AND TYMED_HGLOBAL <> 0) THEN
    BEGIN
      Medium.hGlobal := GlobalAlloc(GMEM_SHARE OR GHND, Length(fText)+1);
      IF (Medium.hGlobal = 0) THEN
      BEGIN
        result := E_OUTOFMEMORY;
        Exit;
      END;
      medium.tymed := TYMED_HGLOBAL;

      pText := PChar(GlobalLock(Medium.hGlobal));
      StrCopy(pText, PChar(fText));
      
      //if RenderTextAsHGlobal(fText,Medium.hGlobal) then result := S_OK;

      GlobalUnlock(Medium.hGlobal);
      result := S_OK;
    END ELSE
      result := DV_E_FORMATETC;
  END;

  // -----------------------------------------------------------------------------
  //			TDropFileSource
  // -----------------------------------------------------------------------------

  var
    CF_IDLIST: UINT;

  //******************* TDropFileSource.Create *************************
  CONSTRUCTOR TDropFileSource.Create(aOwner: TComponent);
  BEGIN
    INHERITED Create(aOwner);
    fFiles := TStringList.Create;
    CF_IDLIST := RegisterClipboardFormat(CFSTR_SHELLIDLIST);

    fDataFormats[0].cfFormat := CF_HDROP;
    fDataFormats[0].ptd      := NIL;
    fDataFormats[0].dwAspect := DVASPECT_CONTENT;
    fDataFormats[0].lIndex   := -1;
    fDataFormats[0].tymed    := TYMED_HGLOBAL;

    fDataFormats[1].cfFormat := CF_IDLIST;
    fDataFormats[1].ptd      := NIL;
    fDataFormats[1].dwAspect := DVASPECT_CONTENT;
    fDataFormats[1].lIndex   := -1;
    fDataFormats[1].tymed    := TYMED_HGLOBAL;

    DataFormatsCount := 2;
  END;

  //******************* TDropFileSource.Destroy *************************
  DESTRUCTOR TDropFileSource.destroy;
  BEGIN
    fFiles.Free;
    INHERITED Destroy;
  END;

  // Adapted from Zbysek Hlinka, zhlinka@login.cz.
  //******************* TDropFileSource.CopyToClipboard *************************
  FUNCTION TDropFileSource.CopyToClipboard: boolean;
  VAR
    FormatEtcIn: TFormatEtc;
    Medium: TStgMedium;
  BEGIN
    FormatEtcIn.cfFormat := CF_HDROP;
    FormatEtcIn.dwAspect := DVASPECT_CONTENT;
    FormatEtcIn.tymed := TYMED_HGLOBAL;
    IF Files.count = 0 then result := false
    ELSE IF GetData(formatetcIn,Medium) = S_OK THEN
    BEGIN
      Clipboard.SetAsHandle(CF_HDROP,Medium.hGlobal);
      result := true;
    END ELSE result := false;
  END;

  //******************* TDropFileSource.DoGetDataHere *************************
  FUNCTION TDropFileSource.DoGetDataHere(CONST FormatEtc: TFormatEtc;
           OUT Medium: TStgMedium):HRESULT;
  VAR
    i: Integer;
    dropfiles: pMyDropFiles;
    pText: PChar;
  BEGIN
    IF (FormatEtc.cfFormat = CF_HDROP) AND
      ((FormatEtc.dwAspect = DVASPECT_CONTENT) OR (FormatEtc.dwAspect = DVASPECT_ICON)) AND
      (FormatEtc.tymed = TYMED_HGLOBAL) AND (Medium.tymed = TYMED_HGLOBAL) THEN
    BEGIN
      IF (Medium.hGlobal = 0) THEN
      BEGIN
        result:=E_OUTOFMEMORY;
        Exit;
      END;
      dropfiles := GlobalLock(Medium.hGlobal);
      TRY
        dropfiles^.dropfiles.pfiles := SizeOf(TDropFiles);
        dropfiles^.dropfiles.fwide := False;
        pText := @(dropfiles^.filenames);
        FOR i := 0 TO fFiles.Count-1 DO
        BEGIN
          StrPCopy(pText, fFiles[i]);
          Inc(pText, Length(fFiles[i])+1);
        END;
        pText^ := #0;
      FINALLY
        GlobalUnlock(Medium.hGlobal);
      END;
      Medium.UnkForRelease := NIL;
      result := S_OK;
    END ELSE
      result := INHERITED DoGetDataHere(FormatEtc, Medium);
  END;

  //******************* TDropFileSource.DoGetData *************************
  FUNCTION TDropFileSource.DoGetData(CONST FormatEtcIn: TFormatEtc;
           OUT Medium: TStgMedium):HRESULT;
  VAR
    i: Integer;
    strlength: Integer;
    tmpFilenames: TStringList;
  BEGIN
    Medium.tymed := 0;
    Medium.UnkForRelease := NIL;
    Medium.hGlobal := 0;

    IF fFiles.count = 0 then result := E_UNEXPECTED
    ELSE IF (FormatEtcIn.cfFormat = CF_HDROP) AND
      ((FormatEtcIn.dwAspect = DVASPECT_CONTENT) OR (FormatEtcIn.dwAspect = DVASPECT_ICON)) AND
      (FormatEtcIn.tymed AND TYMED_HGLOBAL <> 0) THEN
    BEGIN
      strlength := 0;
      FOR i := 0 TO fFiles.Count-1 DO
        Inc(strlength, Length(fFiles[i])+1);
      Medium.hGlobal := GlobalAlloc(GMEM_SHARE OR GMEM_ZEROINIT, SizeOf(TDropFiles)+strlength+1);
      IF (Medium.hGlobal = 0) THEN
      BEGIN
        result:=E_OUTOFMEMORY;
        Exit;
      END;
      Medium.tymed := TYMED_HGLOBAL;
      result := DoGetDataHere(FormatEtcIn, Medium);
      END
    ELSE IF (FormatEtcIn.cfFormat = CF_IDLIST) AND
      ((FormatEtcIn.dwAspect = DVASPECT_CONTENT) OR (FormatEtcIn.dwAspect = DVASPECT_ICON)) AND
      (FormatEtcIn.tymed AND TYMED_HGLOBAL <> 0) THEN
    BEGIN
      tmpFilenames := TStringList.create;
      Medium.tymed := TYMED_HGLOBAL;
      for i := 0 to fFiles.count-1 do
        tmpFilenames.add(extractfilename(fFiles[i]));
      Medium.hGlobal :=
          ConvertFilesToShellIDList(extractfilepath(fFiles[0]),tmpFilenames);
      if Medium.hGlobal = 0 then
        result:=E_OUTOFMEMORY else
        result := S_OK;
      tmpFilenames.free;
    END ELSE
      result := DV_E_FORMATETC;
  END;

  //********************************************
  //********************************************

INITIALIZATION
  OleInitialize(NIL);
FINALIZATION
  OleUninitialize;

END.
