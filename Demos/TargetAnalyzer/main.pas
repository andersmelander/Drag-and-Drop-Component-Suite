unit main;

interface

uses
  System.SysUtils, System.Variants, System.Classes, System.ImageList,
  WinApi.ActiveX, WinApi.Windows, WinApi.Messages,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.ComCtrls,
  Vcl.ImgList, Vcl.StdCtrls, Vcl.Menus,
  Vcl.Buttons,
  DragDrop, DropSource;

type
  TRequestedFormat = record
    Format: TClipFormat;
    Medium: integer;
  end;

const
  MSG_FIREFOX_DETECTED = WM_USER;

type
  TOnQueryGetData = procedure(const FormatEtc: TFormatEtc; var Result: HRESULT; var Handled: boolean) of object;

  // Hack: Replace the TDropEmptySource in DropSource unit witj our own specialized debug version
  TDropEmptySource = class(TCustomDropMultiSource, IDataObject)
  private
    FOnQueryGetData: TOnQueryGetData;
  protected
    function QueryGetData(const FormatEtc: TFormatEtc): HRESULT; stdcall;
  public
    property OnQueryGetData: TOnQueryGetData read FOnQueryGetData write FOnQueryGetData;
  end;

  TTraceKind = (tkSourceBegin, tkSourceEnd, tkTargetBegin, tkTargetEnd, tkInfo);

  TFormTarget = class(TForm)
    DataFormatAdapter1: TDataFormatAdapter;
    DropEmptySource1: TDropEmptySource;
    ListViewTrace: TListView;
    ImageList1: TImageList;
    Panel2: TPanel;
    ListViewFormats: TListView;
    Panel3: TPanel;
    Panel1: TPanel;
    Panel4: TPanel;
    CheckBoxDropFile: TCheckBox;
    PopupMenu1: TPopupMenu;
    Copytoclipboard1: TMenuItem;
    Splitter1: TSplitter;
    SpeedButton1: TSpeedButton;
    ButtonAbort: TButton;
    procedure Panel1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure DropEmptySource1Drop(Sender: TObject; DragType: TDragType;
      var ContinueDrop: Boolean);
    procedure DropEmptySource1Feedback(Sender: TObject; Effect: Integer;
      var UseDefaultCursors: Boolean);
    procedure DropEmptySource1AfterDrop(Sender: TObject;
      DragResult: TDragResult; Optimized: Boolean);
    procedure DropEmptySource1GetData(Sender: TObject;
      const FormatEtc: tagFORMATETC; out Medium: tagSTGMEDIUM;
      var Handled: Boolean);
    procedure DropEmptySource1SetData(Sender: TObject;
      const FormatEtc: tagFORMATETC; out Medium: tagSTGMEDIUM;
      var Handled: Boolean);
    procedure DropEmptySource1Paste(Sender: TObject; Action: TDragResult;
      DeleteOnPaste: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure Copytoclipboard1Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure ButtonAbortClick(Sender: TObject);
  private
    FStartTime: DWORD;
//    FDragContextClipboardFormatID: TClipFormat;
    FFormats: array of TRequestedFormat;
    FAbortCount: integer;
    FMozillaFirefox: boolean;
    procedure Trace(Kind: TTraceKind; const Action: string; const Details: string = '');
    procedure AddClipboardFormat(Format: TClipFormat; Medium: integer);
    procedure OnQueryGetData(const FormatEtc: TFormatEtc; var Result: HRESULT; var Handled: boolean);
    procedure MsgFirefoxDetected(var Msg: TMessage); message MSG_FIREFOX_DETECTED;
  public
  end;

var
  FormTarget: TFormTarget;

implementation

{$R *.dfm}

uses
  DragDropFile,
  DragDropFormats;

const
  sDragType: array[TDragType] of string =
    ('Copy', 'Move', 'Link');
  sDragResult: array[TDragResult] of string =
    ('Copy', 'Move', 'Link', 'Cancel', 'Out of Memory',
     'Asynchronous drag in progress', 'Unknown');
  sCopyResult: array[boolean] of string =
    ('Failed', 'Success');
  sMediaNames: array[0..6] of string =
    ('GlobalMem', 'File', 'IStream', 'IStorage', 'GDI', 'MetaFile', 'EnhMetaFile');

resourcestring
  sAnyMedia = '(any)';

resourcestring
  sMsgFirefoxDetectTitle = 'Mozilla Firefox detected';
  sMsgFirefoxWorkaround = 'Applying work around';
  sMsgFirefoxDetect = 'Due to a bug in Firefox, dropping from this application onto Firefox causes an endless loop in Firefox.'#13#13+
    'The application will now terminate.'#13#13+
    'Please press [Esc] to terminate.';

type
  TDragContextClipboardFormat = class(TCustomDWORDClipboardFormat)
  public
    function GetClipboardFormat: TClipFormat; override;
  end;

function GetMediaName(Media: integer): string;
var
  Mask: integer;
  i: integer;
  s: string;
begin
  if (Media = -1) then
  begin
    // Some targets (e.g. Araxis Merge pass -1 in FormatEtc.tymed. I assume this
    // could mean they don't care about the medium, but it's probably a bug on
    // their end.
    Result := sAnyMedia;
    exit;
  end;

  Result := '';
  Mask := 1;
  for i := Low(sMediaNames) to High(sMediaNames) do
  begin
    if (Media and Mask = Mask) then
    begin
      if (Result = '') then
        Result := sMediaNames[i]
      else
        Result := Result+','+sMediaNames[i];
    end;
    Mask := Mask shl 1;
  end;
  if (Media >= Mask) then
  begin
    s := Format('$%0.4x', [Media and not(Mask-1)]);
    if (Result = '') then
      Result := s
    else
      Result := Result+','+s;
  end;
end;

{ TDropEmptySource }

function TDropEmptySource.QueryGetData(const FormatEtc: TFormatEtc): HRESULT;
var
  Handled: boolean;
begin
  if (Assigned(FOnQueryGetData)) then
  begin
    Result:= DV_E_FORMATETC;
    Handled := False;
    FOnQueryGetData(FormatEtc, Result, Handled);
  end;

  if (not Handled) then
    Result := inherited QueryGetData(FormatEtc);
end;

procedure TFormTarget.AddClipboardFormat(Format: TClipFormat; Medium: integer);
var
  i: integer;
begin
  i := 0;
  while (i < Length(FFormats)) do
  begin
    if (FFormats[i].Format = 0) then
    begin
      FFormats[i].Format := Format;
      FFormats[i].Medium := Medium;
      break;
    end else
    if (FFormats[i].Format = Format) then
    begin
      FFormats[i].Medium := FFormats[i].Medium or Medium;
      break;
    end;
    inc(i);
  end;

  if (i >= Length(FFormats)) then
  begin
    SetLength(FFormats, Length(FFormats)+8);
    FFormats[i].Format := Format;
    FFormats[i].Medium := Medium;
  end;

  if (i = ListViewFormats.Items.Count) then
    with ListViewFormats.Items.Add do
    begin
      Caption := GetClipboardFormatNameStr(FFormats[i].Format);
      SubItems.Add(GetMediaName(FFormats[i].Medium));
    end
  else
    ListViewFormats.Items[i].SubItems[0] := GetMediaName(FFormats[i].Medium);
end;

procedure TFormTarget.ButtonAbortClick(Sender: TObject);
begin
  inc(FAbortCount);
end;

procedure TFormTarget.Copytoclipboard1Click(Sender: TObject);
var
  DragResult: boolean;
begin
  ListViewTrace.Items.Clear;
  ListViewFormats.Items.Clear;
  Setlength(FFormats, 0);

  DataFormatAdapter1.Enabled := CheckBoxDropFile.Checked;
//  if (DataFormatAdapter1.Enabled) then
//    TFileDataFormat(DataFormatAdapter1.DataFormat).Files.Text := Application.ExeName;

  FStartTime := GetTickCount;
  FAbortCount := 0;
  Trace(tkSourceBegin, 'DropSource.CopyToClipboard');
  DragResult := DropEmptySource1.CopyToClipboard;
  Trace(tkSourceEnd, 'DropSource.CopyToClipboard', sCopyResult[DragResult]);
end;

procedure TFormTarget.DropEmptySource1AfterDrop(Sender: TObject;
  DragResult: TDragResult; Optimized: Boolean);
const
  sOptimized: array[boolean] of string = ('', ' (optimized)');
begin
  Trace(tkSourceEnd, 'DoDragDrop returned', sDragResult[DragResult]+sOptimized[Optimized]);

  ButtonAbort.Visible := False;
  Panel1.ShowCaption := True;
  FMozillaFirefox := False;
end;

procedure TFormTarget.DropEmptySource1Drop(Sender: TObject; DragType: TDragType;
  var ContinueDrop: Boolean);
begin
  Trace(tkTargetEnd, 'IDropSource.QueryContinueDrag (Drop)', sDragType[DragType]);
end;

procedure TFormTarget.DropEmptySource1Feedback(Sender: TObject; Effect: Integer;
  var UseDefaultCursors: Boolean);
var
  DragType: TDragType;
begin
  DropEffectToDragType(Effect, DragType);
  Trace(tkTargetEnd, 'IDropSource.GiveFeedback', sDragType[DragType]);
end;

procedure TFormTarget.DropEmptySource1GetData(Sender: TObject;
  const FormatEtc: tagFORMATETC; out Medium: tagSTGMEDIUM;
  var Handled: Boolean);
(*
var
  DragContext: TClipboardFormat;
*)
begin
  AddClipboardFormat(FormatEtc.cfFormat, FormatEtc.tymed);

  Trace(tkTargetEnd, 'IDataObject.GetData', Format('%d: %s on %s', [FormatEtc.cfFormat, GetClipboardFormatNameStr(FormatEtc.cfFormat), GetMediaName(FormatEtc.tymed)]));

  // Work around the Delphi IDE's attempt to reject data from anything but the Explorer
(*
  if (FormatEtc.cfFormat = FDragContextClipboardFormatID) then
  begin
    DragContext := TDragContextClipboardFormat.Create;
    try
      DragContext.SetDataToMedium(FormatEtc, Medium);
    finally
      DragContext.Free;
    end;
    Handled := True;
  end;
*)
end;

procedure TFormTarget.DropEmptySource1Paste(Sender: TObject;
  Action: TDragResult; DeleteOnPaste: Boolean);
const
  sDeleteOnPaste: array[boolean] of string = ('', ' (Delete on Paste)');
begin
  Trace(tkTargetEnd, 'Paste from clipboard', sDragResult[Action]+sDeleteOnPaste[DeleteOnPaste]);
end;

procedure TFormTarget.DropEmptySource1SetData(Sender: TObject;
  const FormatEtc: tagFORMATETC; out Medium: tagSTGMEDIUM;
  var Handled: Boolean);
begin
  Trace(tkTargetEnd, 'IDataObject.SetData', GetClipboardFormatNameStr(FormatEtc.cfFormat));
end;

procedure TFormTarget.FormCreate(Sender: TObject);
begin
  DropEmptySource1.OnQueryGetData := OnQueryGetData;
//  FDragContextClipboardFormatID := RegisterClipboardFormat('DragContext');
end;

procedure TFormTarget.FormDestroy(Sender: TObject);
begin
  // Make sure we don't get any events during form destruction.
  DropEmptySource1.FlushClipboard;
end;

procedure TFormTarget.MsgFirefoxDetected(var Msg: TMessage);
begin
  Screen.Cursor := crDefault;
  MessageBox(Handle, PChar(sMsgFirefoxDetect), PChar(sMsgFirefoxDetectTitle), MB_ICONWARNING or MB_OK);
  Halt;
end;

procedure TFormTarget.OnQueryGetData(const FormatEtc: TFormatEtc;
  var Result: HRESULT; var Handled: boolean);
var
  Name: string;
begin
  AddClipboardFormat(FormatEtc.cfFormat, FormatEtc.tymed);

  Name := GetClipboardFormatNameStr(FormatEtc.cfFormat);

  Trace(tkTargetEnd, 'IDataObject.QueryGetData', Format('%d: %s on %s', [FormatEtc.cfFormat, Name, GetMediaName(FormatEtc.tymed)]));

  // Work around Firefox's endless QuerygetData loop
  if (not FMozillaFirefox) and (Name = 'Mozilla/IDataObjectCollectionFormat') then
  begin
    Trace(tkInfo, sMsgFirefoxDetectTitle, sMsgFirefoxWorkaround);
    // Note: The original Firefox appears to have been fixed, but Firefox is
    // still a bit slow to react to the E_ABORT.
    // FMozillaFirefox := True;
  end;

  Result := S_OK;
  Handled := True;

  if (FMozillaFirefox) or (FAbortCount > 0) or (GetAsyncKeyState(VK_ESCAPE) and $0001 = $0001) then
  begin
    Result := E_ABORT;

    // Halt if we have detected that the application ignored the previous abort
    // request and are requesting one of the Firefox formats.
    if (FMozillaFirefox) and (FAbortCount > 100) then
    begin
      PostMessage(Handle, MSG_FIREFOX_DETECTED, 0, 0);
      // Provide us with some means of getting out of Firefox's endless QuerygetData loop
      Self.Enabled := False;
      while (True) do
        Application.ProcessMessages;
    end;
    inc(FAbortCount);
  end;
end;

procedure TFormTarget.Panel1MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  DragResult: TDragResult;
begin
  if (DragDetectPlus(Sender as TWinControl)) then
  begin
    ListViewTrace.Items.Clear;
    ListViewFormats.Items.Clear;
    Setlength(FFormats, 0);

    DataFormatAdapter1.Enabled := CheckBoxDropFile.Checked;
    if (DataFormatAdapter1.Enabled) then
      TFileDataFormat(DataFormatAdapter1.DataFormat).Files.Text := Application.ExeName;

    FStartTime := GetTickCount;
    FAbortCount := 0;
    Trace(tkSourceBegin, 'DropSource.Execute (DoDragDrop)');

    Panel1.ShowCaption := False;
    ButtonAbort.Visible := True;

    DragResult := DropEmptySource1.Execute(True);

    Trace(tkSourceEnd, 'DropSource.Execute (DoDragDrop)', sDragResult[DragResult]);
  end;
end;

resourcestring
  sAbout = 'This application analyzes how a drop target interacts with a drop source.'+#13+#13+
    'The left pane lists the actions performed during a drag drop operation.'+#13+
    'The right pane lists the clipboard formats requested by the drop target.';

procedure TFormTarget.SpeedButton1Click(Sender: TObject);
begin
  ShowMessage(sAbout);
end;

procedure TFormTarget.Trace(Kind: TTraceKind; const Action, Details: string);
var
  Item: TListItem;
  Delta: DWORD;
begin
  Delta := GetTickCount-FStartTime;
  ListViewTrace.Items.BeginUpdate;
  try
    Item := ListViewTrace.Items.Add;
    Item.Caption := Format('%d.%.3d', [Delta div 1000, Delta mod 1000]);
    Item.ImageIndex := ord(Kind);
    Item.SubItems.Add(Action);
    Item.SubItems.Add(Details);
  finally
    ListViewTrace.Items.EndUpdate;
  end;
  Item.MakeVisible(False);
  ListViewTrace.Update;
end;

{ TDragContextClipboardFormat }

function TDragContextClipboardFormat.GetClipboardFormat: TClipFormat;
begin
  Result := RegisterClipboardFormat('DragContext');
end;

end.
