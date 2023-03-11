unit main;

interface

uses
  DragDrop, DropSource, DragDropFormats,
  ActiveX, Windows, Classes, Controls, Forms, StdCtrls, ComCtrls, ExtCtrls,
  Buttons;

type
  TFormMain = class(TForm)
    Timer1: TTimer;
    Panel2: TPanel;
    DropEmptySource1: TDropEmptySource;
    DataFormatAdapterSource: TDataFormatAdapter;
    ProgressBar1: TProgressBar;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    RadioButtonNormal: TRadioButton;
    RadioButtonAsync: TRadioButton;
    PaintBoxPie: TPaintBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    StatusBar1: TStatusBar;
    ButtonAbort: TSpeedButton;
    procedure Timer1Timer(Sender: TObject);
    procedure DropEmptySource1Drop(Sender: TObject; DragType: TDragType;
      var ContinueDrop: Boolean);
    procedure DropEmptySource1AfterDrop(Sender: TObject;
      DragResult: TDragResult; Optimized: Boolean);
    procedure DropEmptySource1GetData(Sender: TObject;
      const FormatEtc: tagFORMATETC; out Medium: tagSTGMEDIUM;
      var Handled: Boolean);
    procedure OnMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure ButtonAbortClick(Sender: TObject);
    procedure StatusBar1Resize(Sender: TObject);
  private
    Tick: integer;
    EvenOdd: boolean;
    DoAbort: boolean;
    procedure OnGetStream(Sender: TFileContentsStreamOnDemandClipboardFormat;
      Index: integer; out AStream: IStream);
    procedure OnProgress(Sender: TObject; Count, MaxCount: integer);
  public
  end;

var
  FormMain: TFormMain;

implementation

{$R *.DFM}

uses
  Messages,
  ShlObj,
  Graphics;

const
  TestFileSize = 1024*1024*10; // 10Mb

procedure TFormMain.FormCreate(Sender: TObject);
begin
  // Setup event handler to let a drop target request data from our drop source.
  (DataFormatAdapterSource.DataFormat as TVirtualFileStreamDataFormat).OnGetStream := OnGetStream;

  StatusBar1.ControlStyle := StatusBar1.ControlStyle +[csAcceptsControls];
(*
  // Reparent abort button to statusbar
  ButtonAbort.Top := 3;
  ButtonAbort.Height := StatusBar1.Height-4;
  ButtonAbort.Left := StatusBar1.Width-ButtonAbort.Width-1;
*)
end;

procedure TFormMain.FormDestroy(Sender: TObject);
begin
  // Reparent abort button to form
  Timer1.Enabled := False;
end;

procedure TFormMain.OnMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  StatusBar1.SimpleText := '';
  if DragDetectPlus(Handle, Point(X, Y)) then
  begin
    StatusBar1.SimpleText := 'Dragging data';

    // Transfer the file names to the data format. The content will be extracted
    // by the target on-demand.
    TVirtualFileStreamDataFormat(DataFormatAdapterSource.DataFormat).FileNames.Clear;
    TVirtualFileStreamDataFormat(DataFormatAdapterSource.DataFormat).FileNames.Add('big text file.txt');
    // Set the size and timestamp attributes of the filename we just added.
    with PFileDescriptor(TVirtualFileStreamDataFormat(DataFormatAdapterSource.DataFormat).FileDescriptors[0])^ do
    begin
      GetSystemTimeAsFileTime(ftLastWriteTime);
      nFileSizeLow := TestFileSize;
      nFileSizeHigh := 0; // I assume the test file doesn't grow beyond 4Gb...
      dwFlags := FD_WRITESTIME or FD_FILESIZE or FD_PROGRESSUI;
    end;

    // Determine if we should perform an async drag or a normal drag.
    if (RadioButtonAsync.Checked) then
    begin
      DoAbort := False;
      ButtonAbort.Visible := True;
      ButtonAbort.Enabled := True;

      // Create a thread to perform the drag...
      if (DropEmptySource1.Execute(True) = drAsync) then
        StatusBar1.SimpleText := 'Asynchronous drag in progress...'
      else
        StatusBar1.SimpleText := 'Asynchronous drag failed'
    end else
    begin
      // Perform a normal drag (in the main thread).
      DropEmptySource1.Execute;

      StatusBar1.SimpleText := 'Drop completed';
    end;
  end;
end;

procedure TFormMain.Timer1Timer(Sender: TObject);

  procedure DrawPie(Percent: integer);
  var
    Center: TPoint;
    Radial: TPoint;
    v: Double;
    Radius: integer;
  begin
    // Assume paintbox width is smaller than height.
    Radius := PaintBoxPie.Width div 2 - 10;
    Center := Point(PaintBoxPie.Width div 2, PaintBoxPie.Height div 2);
    v := Percent * Pi / 50; // Convert percent to radians.
    Radial.X := Center.X+trunc(Radius * Cos(v));
    Radial.Y := Center.Y-trunc(Radius * Sin(v));

    PaintBoxPie.Canvas.Brush.Style := bsSolid;
    PaintBoxPie.Canvas.Pen.Color := clGray;
    PaintBoxPie.Canvas.Pen.Style := psSolid;

    if (EvenOdd) then
      PaintBoxPie.Canvas.Brush.Color := clRed
    else
      PaintBoxPie.Canvas.Brush.Color := Color;
    PaintBoxPie.Canvas.Pie(Center.X-Radius, Center.Y-Radius,
      Center.X+Radius, Center.Y+Radius,
      Radial.X, Radial.Y,
      Center.X+Radius, Center.Y);

    if (Percent <> 0) then
    begin
      if not(EvenOdd) then
        PaintBoxPie.Canvas.Brush.Color := clRed
      else
        PaintBoxPie.Canvas.Brush.Color := Color;
      PaintBoxPie.Canvas.Pie(Center.X-Radius, Center.Y-Radius,
        Center.X+Radius, Center.Y+Radius,
        Center.X+Radius, Center.Y,
        Radial.X, Radial.Y);
    end;
  end;

begin
  // Update the pie to indicate that the application is responding to
  // messages (i.e. isn't blocked).
  Tick := (Tick + 10) mod 100;
  if (Tick = 0) then
    EvenOdd := not EvenOdd;

  // Draw an animated pie chart to show that application is responsive to events.
  DrawPie(Tick);
end;

procedure TFormMain.DropEmptySource1Drop(Sender: TObject;
  DragType: TDragType; var ContinueDrop: Boolean);
begin
  StatusBar1.SimpleText := 'Target processing drop';
end;

procedure TFormMain.DropEmptySource1AfterDrop(Sender: TObject;
  DragResult: TDragResult; Optimized: Boolean);
begin
  StatusBar1.SimpleText := 'Drop completed';
  ButtonAbort.Visible := False;
end;

procedure TFormMain.DropEmptySource1GetData(Sender: TObject;
  const FormatEtc: tagFORMATETC; out Medium: tagSTGMEDIUM;
  var Handled: Boolean);
begin
  StatusBar1.SimpleText := 'Target reading data';
end;

type
  TStreamProgressEvent = procedure(Sender: TObject; Count, MaxCount: integer) of object;

  // TFakeStream is a read-only stream which produces its contents on-the-run.
  // It is used for this demo so we can simulate transfer of very large and
  // arbitrary amounts of data without using any memory.
  TFakeStream = class(TStream)
  private
    FSize, FPosition, FMaxCount: Longint;
    FProgress: TStreamProgressEvent;
    FAbort: boolean;
  protected
  public
    constructor Create(ASize, AMaxCount: LongInt);
    function Read(var Buffer; Count: Longint): Longint; override;
    function Seek(Offset: Longint; Origin: Word): Longint; override;
    procedure SetSize(NewSize: Longint); override;
    function Write(const Buffer; Count: Longint): Longint; override;
    procedure Abort;
    property OnProgress: TStreamProgressEvent read FProgress write FProgress;
  end;

procedure TFakeStream.Abort;
begin
  FAbort := True;
end;

constructor TFakeStream.Create(ASize, AMaxCount: LongInt);
begin
  inherited Create;
  FSize := ASize;
  FMaxCount := AMaxCount;
end;

function TFakeStream.Read(var Buffer; Count: Integer): Longint;
begin
  Result := 0;
  if (FAbort) then
    exit;

  if (FPosition >= 0) and (Count >= 0) then
  begin
    Result := FSize - FPosition;
    if Result > 0 then
    begin
      if Result > Count then
        Result := Count;
      if Result > FMaxCount then
        Result := FMaxCount;
      FillChar(Buffer, Result, ord('X'));
      Inc(FPosition, Result);
      if Assigned(FProgress) then
        FProgress(Self, FPosition, FSize);
    end;
  end;
end;

function TFakeStream.Seek(Offset: Integer; Origin: Word): Longint;
begin
  case Origin of
    soFromBeginning: FPosition := Offset;
    soFromCurrent: Inc(FPosition, Offset);
    soFromEnd: FPosition := FSize + Offset;
  end;
  if Assigned(FProgress) then
    FProgress(Self, FPosition, FMaxCount);
  Result := FPosition;
end;

procedure TFakeStream.SetSize(NewSize: Integer);
begin
end;

function TFakeStream.Write(const Buffer; Count: Integer): Longint;
begin
  Result := 0;
end;

procedure TFormMain.OnGetStream(
  Sender: TFileContentsStreamOnDemandClipboardFormat; Index: integer;
  out AStream: IStream);
var
  Stream: TStream;
begin
  // Note: This method might be called in the context of the transfer thread.
  // See TFormMain.OnProgress for a comment on this.

  // This event handler is called by TFileContentsStreamOnDemandClipboardFormat
  // when the drop target requests data from the drop source (that's us).
  StatusBar1.SimpleText := 'Transfering data';

  // Create a stream which contains the data we will transfer...
  // In this case we just create a dummy stream which contains 10Mb of 'X'
  // characters. In order to provide smoth feedback through the progress bar
  // (and slow the transfer down a bit) the stream will only transfer up to 32K
  // at a time - Each time TStream.Read is called the progress bar is updated
  // via the stream's progress event.
  Stream := TFakeStream.Create(TestFileSize, 32*1024);
  try
    TFakeStream(Stream).OnProgress := OnProgress;
    // ...and return the stream back to the target as an IStream. Note that the
    // target is responsible for deleting the stream (via reference counting).
    AStream := TFixedStreamAdapter.Create(Stream, soOwned);
  except
    Stream.Free;
    raise;
  end;

  ProgressBar1.Position := 0;
end;

procedure TFormMain.OnProgress(Sender: TObject; Count, MaxCount: integer);
begin
  // Note that during an asyncronous transfer, the progress event handler is
  // being called in the context of the transfer thread. This means that this
  // event handler should adhere to all the normal thread safety rules (i.e.
  // don't call GDI or mess with non-thread safe objects).

  // Update progress bar to show how much data has been transfered so far.
  // This isn't really thread safe since it modifies the form, but so far it
  // hasn't crashed on me.
  ProgressBar1.Max := MaxCount;
  ProgressBar1.Position := Count;
  if (DoAbort) then
    TFakeStream(Sender).Abort;
end;

procedure TFormMain.ButtonAbortClick(Sender: TObject);
begin
  DoAbort := True;
  TButton(Sender).Enabled := False;
end;

procedure TFormMain.StatusBar1Resize(Sender: TObject);
begin
  // This is nescessary to get controls inside TStatusBar to honour Anchors.
  StatusBar1.Realign;
end;

end.

