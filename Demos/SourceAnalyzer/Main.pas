unit Main;

interface

uses
  DragDrop, DropTarget,
  ActiveX,
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, ComCtrls, StdCtrls;

const
  MAX_DATA = 32768; // Max bytes to render in preview

type
  TOmnipotentDropTarget = class(TCustomDropMultiTarget)
  protected
    function DoGetData: boolean; override;
  public
    function HasValidFormats(ADataObject: IDataObject): boolean; override;
  end;

  TFormMain = class(TForm)
    Panel2: TPanel;
    Splitter1: TSplitter;
    StatusBar1: TStatusBar;
    MemoHexView: TRichEdit;
    ListViewDataFormats: TListView;
    procedure FormCreate(Sender: TObject);
    procedure ListViewDataFormatsDeletion(Sender: TObject;
      Item: TListItem);
    procedure ListViewDataFormatsSelectItem(Sender: TObject;
      Item: TListItem; Selected: Boolean);
    procedure FormDestroy(Sender: TObject);
  private
    FDataObject: IDataObject;
    procedure OnDrop(Sender: TObject; ShiftState: TShiftState;
      APoint: TPoint; var Effect: Longint);
    function DataToHexDump(const Data: string): string;
    function GetDataSize(const FormatEtc: TFormatEtc): integer;
  public
  end;

var
  FormMain: TFormMain;

implementation

{$R *.dfm}

uses
  DragDropFormats;

{ TOmnipotentDropTarget }

function TOmnipotentDropTarget.DoGetData: boolean;
begin
  Result := True;
end;

function TOmnipotentDropTarget.HasValidFormats(ADataObject: IDataObject): boolean;
begin
  Result := True;
end;

{ TFormMain }

procedure TFormMain.FormDestroy(Sender: TObject);
begin
  FDataObject := nil;
end;

procedure TFormMain.FormCreate(Sender: TObject);
begin
  with TOmnipotentDropTarget.Create(Self) do
  begin
    DragTypes := [dtCopy, dtLink];
    Target := Self;
    OnDrop := Self.OnDrop;
  end;
end;

function TFormMain.GetDataSize(const FormatEtc: TFormatEtc): integer;
var
  Medium: TStgMedium;
begin
  FillChar(Medium, SizeOf(Medium), 0);
  if (Succeeded(FDataObject.GetData(FormatEtc, Medium))) then
  begin
    try
      Result := GetMediumDataSize(Medium);
    finally
      ReleaseStgMedium(Medium);
    end;
  end else
    Result := -1;
end;

procedure TFormMain.OnDrop(Sender: TObject; ShiftState: TShiftState;
  APoint: TPoint; var Effect: Integer);
var
  GetNum, GotNum: longInt;
  FormatEnumerator: IEnumFormatEtc;
  Aspects, Aspect: integer;
  AspectNum: integer;
  Media, Medium: integer;
  MediumNum: integer;
  SourceFormatEtc: TFormatEtc;
  FormatEtc: PFormatEtc;
  Item: TListItem;
  s: string;
  Size: integer;
const
 ClipNames: array[CF_TEXT..CF_MAX-1] of string =
   ('CF_TEXT', 'CF_BITMAP', 'CF_METAFILEPICT', 'CF_SYLK', 'CF_DIF', 'CF_TIFF',
   'CF_OEMTEXT', 'CF_DIB', 'CF_PALETTE', 'CF_PENDATA', 'CF_RIFF', 'CF_WAVE',
   'CF_UNICODETEXT', 'CF_ENHMETAFILE', 'CF_HDROP', 'CF_LOCALE');
 MediaNames: array[0..7] of string =
   ('GlobalMem', 'File', 'IStream', 'IStorage', 'GDI', 'MetaFile', 'EnhMetaFile', 'Unknown');
 AspectNames: array[0..3] of string =
    ('Content', 'Thumbnail', 'Icon', 'Print');
begin
  ListViewDataFormats.Items.BeginUpdate;
  try
    ListViewDataFormats.Items.Clear;

    FDataObject := TCustomDropTarget(Sender).DataObject;

    if (FDataObject.EnumFormatEtc(DATADIR_GET, FormatEnumerator) <> S_OK) or
      (FormatEnumerator.Reset <> S_OK) then
    begin
      FDataObject := nil;
      exit;
    end;

    GetNum := 1; // Get one format at a time.

    // Enumerate all data formats offered by the drop source.
    while (FormatEnumerator.Next(GetNum, SourceFormatEtc, @GotNum) = S_OK) and
      (GetNum = GotNum) do
    begin
      Item := ListViewDataFormats.Items.Add;

      Item.Caption := IntToStr(SourceFormatEtc.cfFormat);

      if (SourceFormatEtc.cfFormat < CF_MAX) then
        Item.SubItems.Add(ClipNames[SourceFormatEtc.cfFormat])
      else
        Item.SubItems.Add(GetClipboardFormatNameStr(SourceFormatEtc.cfFormat));

      Aspects := SourceFormatEtc.dwAspect;
      AspectNum := 0;
      Aspect := $0001;
      s := '';
      while (Aspects >= Aspect) do
      begin
        if (Aspects and Aspect <> 0) then
        begin
          if (s <> '') then
            s := s+'+';
          s := s+AspectNames[AspectNum];
        end;
        inc(AspectNum);
        Aspect := Aspect shl 1;
      end;
      Item.SubItems.Add(s);

      s := '';
      Media := SourceFormatEtc.tymed;
      MediumNum := 0;
      Medium := $0001;
      while (Media >= Medium) do
      begin
        if (Media and Medium <> 0) then
        begin
          if (s <> '') then
            s := s+', ';
          s := s+MediaNames[MediumNum];
        end;
        inc(MediumNum);
        Medium := Medium shl 1;
      end;
      Item.SubItems.Add(s);

      Size := GetDataSize(SourceFormatEtc);
      if (Size > 0) then
        s := Format('%.0n', [Int(Size)])
      else
        s := '-';
      Item.SubItems.Add(s);

      New(FormatEtc);
      Item.Data := FormatEtc;
      FormatEtc^ := SourceFormatEtc;
    end;
  finally
    ListViewDataFormats.Items.EndUpdate;
  end;
end;

function TFormMain.DataToHexDump(const Data: string): string;
var
  i: integer;
  Offset: integer;
  Hex: string;
  ASCII: string;
  LineLength: integer;
  Size: integer;
begin
  Result := '';
  LineLength := 0;
  Hex := '';
  ASCII := '';
  Offset := 0;

  Size := Length(Data);
  if (Size > MAX_DATA) then
    Size := MAX_DATA;

  for i := 0 to Size-1 do
  begin
    Hex := Hex+IntToHex(ord(Data[i+1]), 2)+' ';
    if (Data[i+1] in [' '..#$7F]) then
      ASCII := ASCII+Data[i+1]
    else
      ASCII := ASCII+'.';
    inc(LineLength);
    if (LineLength = 16) or (i = Length(Data)-1) then
    begin
      Result := Result+Format('%.8x  %-48.48s  %-16.16s'+#13+#10, [Offset, Hex, ASCII]);
      inc(Offset, LineLength);
      LineLength := 0;
      Hex := '';
      ASCII := '';
    end;
  end;
end;

procedure TFormMain.ListViewDataFormatsDeletion(Sender: TObject;
  Item: TListItem);
begin
  if (Item.Data <> nil) then
  begin
    Dispose(Item.Data);
    Item.Data := nil;
  end;
end;

procedure TFormMain.ListViewDataFormatsSelectItem(Sender: TObject;
  Item: TListItem; Selected: Boolean);
var
//  Stream: IStream;
  ClipFormat: TRawClipboardFormat;
begin
  if (Selected) and (Item.Data <> nil) then
  begin
    if (PFormatEtc(Item.Data)^.tymed = TYMED_ISTREAM) then
    begin
//      Stream := IStream(PFormatEtc(Item.Data)^.
    end else
    if (PFormatEtc(Item.Data)^.tymed = TYMED_HGLOBAL) then
    begin
    end else
    begin
    end;

(*
    If (DropRecord[N].FormatEtc.tymed and TYMED_HGLOBAL) <> 0 then
    Begin
      FE := FDropRecord[N].FormatEtc;
      FE.tymed := TYMED_HGLOBAL;
      Rslt := dataObj.GetData(FE, FDropRecord[N].StgMedium);
      If Rslt = S_OK then
      Begin
        Handle := FDropRecord[N].StgMedium.hGlobal;
        FDropRecord[N].datasize := GlobalSize(Handle);
      end;
    end else
    If DropRecord[FmtNum].StgMedium.tymed = TYMED_ISTREAM then
    begin
      Stream := IStream(DropRecord[FmtNum].StgMedium.stm);
      Stream.Seek(0, STREAM_SEEK_SET, PLargeuint(nil)^);
      GetMem(Buffer, MaxDataSize);
      try
        Stream.Read(Buffer, MaxDataSize, @DataSize);
        Render(Buffer, DataSize);
      finally
        FreeMem(Buffer);
      end;
*)

    ClipFormat := TRawClipboardFormat.CreateFormatEtc(PFormatEtc(Item.Data)^);
    try
      ClipFormat.GetData(FDataObject);
      MemoHexView.Lines.Add(DataToHexDump(ClipFormat.AsString));
    finally
      ClipFormat.Free;
    end;
  end else
    MemoHexView.Lines.Clear;
end;

end.

