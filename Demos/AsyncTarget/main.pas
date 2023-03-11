unit main;
(*
This application demonstrates how to receive data asyncronously via a stream.

The application uses a TDropEmptyTarget component and extends it with a
TDataFormatAdapter component.

Note: Asynchronous drop targets requires version 5.00 of shell32.dll and are
only supported on Windows 2000, Windows ME and later.
*)

interface

uses
  DragDrop, DropTarget, DragDropFormats,
  Messages,
  ActiveX, Windows, Classes, Controls, Forms, StdCtrls, ComCtrls, ExtCtrls;

const
  MSG_DROPSTART = WM_USER;
  MSG_DROPINIT = WM_USER+1;
  MSG_DROPPROGRESS = WM_USER+2;
  MSG_DROPDONE = WM_USER+3;

type
  TFormMain = class(TForm)
    Timer1: TTimer;
    StatusBar1: TStatusBar;
    DataFormatAdapterTarget: TDataFormatAdapter;
    ProgressBar1: TProgressBar;
    Panel5: TPanel;
    DropEmptyTarget1: TDropEmptyTarget;
    Panel2: TPanel;
    Panel3: TPanel;
    PaintBoxPie: TPaintBox;
    Panel4: TPanel;
    Label2: TLabel;
    Label3: TLabel;
    RadioButtonNormal: TRadioButton;
    RadioButtonAsync: TRadioButton;
    Label1: TLabel;
    PanelTarget: TPanel;
    procedure Timer1Timer(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure RadioButtonNormalClick(Sender: TObject);
    procedure RadioButtonAsyncClick(Sender: TObject);
    procedure DropEmptyTarget1Drop(Sender: TObject;
      ShiftState: TShiftState; APoint: TPoint; var Effect: Integer);
    procedure DropEmptyTarget1StartAsyncTransfer(Sender: TObject);
    procedure DropEmptyTarget1EndAsyncTransfer(Sender: TObject);
  private
    Tick: integer;
    EvenOdd: boolean;
    procedure MsgDropStart(var Msg: TMessage); message MSG_DROPSTART;
    procedure MsgDropInit(var Msg: TMessage); message MSG_DROPINIT;
    procedure MsgDropProgress(var Msg: TMessage); message MSG_DROPPROGRESS;
    procedure MsgDropEnd(var Msg: TMessage); message MSG_DROPDONE;
  public
  end;

var
  FormMain: TFormMain;

implementation

{$R *.DFM}

uses
  ShlObj,
  Graphics;

procedure TFormMain.FormCreate(Sender: TObject);
var
  FileName: string;
  InfoSize, Wnd: DWORD;
  VerBuf: Pointer;
  FI: PVSFixedFileInfo;
  VerSize: DWORD;
  Version: integer;
begin
  // Check for correct version of shell32.dll.

  Version := 0;
  // GetFileVersionInfo modifies the filename parameter data while parsing.
  // Copy the string const into a local variable to create a writeable copy.
  FileName := 'shell32.dll';
  InfoSize := GetFileVersionInfoSize(PChar(FileName), Wnd);
  if InfoSize <> 0 then
  begin
    GetMem(VerBuf, InfoSize);
    try
      if GetFileVersionInfo(PChar(FileName), Wnd, InfoSize, VerBuf) then
        if VerQueryValue(VerBuf, '\', Pointer(FI), VerSize) then
          Version := FI.dwFileVersionMS;
    finally
      FreeMem(VerBuf);
    end;
  end;

  if (Version < $00050000) then
    StatusBar1.Color := clRed;
end;

procedure TFormMain.FormDestroy(Sender: TObject);
begin
  Timer1.Enabled := False;
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

procedure TFormMain.RadioButtonNormalClick(Sender: TObject);
begin
  DropEmptyTarget1.AllowAsyncTransfer := False;
end;

procedure TFormMain.RadioButtonAsyncClick(Sender: TObject);
begin
  DropEmptyTarget1.AllowAsyncTransfer := True;
end;

procedure TFormMain.DropEmptyTarget1Drop(Sender: TObject;
  ShiftState: TShiftState; APoint: TPoint; var Effect: Integer);
var
  i: integer;
  Stream: IStream;
  StatStg: TStatStg;
  Size, Chunk: longInt;
  Buffer: pointer;
  Progress: integer;
begin
  (*
  ** Warning: When an asynchronous transfer is performed, this event is not
  ** executed in the context of the main thread and thus must adhere to all the
  ** usual thread safety rules (e.g. don't call GDI or mess with non-thread safe
  ** objects).
  **
  ** In this demo we solve that problem by posting a message to ourself whenever
  ** we need to have something executed by the main thread. The messages are
  ** received by the main thread, dispatched to our message handlers, and we can
  ** then do whatever it is we need to do. In this case we just need to update
  ** the status and progress bar.
  *)

  if (DropEmptyTarget1.AsyncTransfer) then
    SendMessage(Handle, MSG_DROPSTART, 0, 0)
  else
    StatusBar1.SimpleText := 'Drop completed - getting data...';

  // Data was dropped on us.
  // Transfer the file contents from the drop source via the stream we
  // get from TVirtualFileStreamDataFormat.
  if (TVirtualFileStreamDataFormat(DataFormatAdapterTarget.DataFormat).FileNames.Count > 0) then
  begin
    // For this demo we just read the data from the source, but ignore the
    // actual content.
    // For a more meaning full example, please see the VirtualFile demo.
    for i := 0 to TVirtualFileStreamDataFormat(DataFormatAdapterTarget.DataFormat).FileNames.Count-1 do
    begin
      // Get data stream from source.
      Stream := TVirtualFileStreamDataFormat(DataFormatAdapterTarget.DataFormat).FileContentsClipboardFormat.GetStream(i);
      if (Stream <> nil) then
      begin
        // Determine size of stream.
        Stream.Stat(StatStg, STATFLAG_NONAME);
        Size := StatStg.cbSize;
        if (DropEmptyTarget1.AsyncTransfer) then
          SendMessage(Handle, MSG_DROPINIT, Size, 0)
        else
        begin
          ProgressBar1.Max := Size;
          ProgressBar1.Position := 0;
          ProgressBar1.Update;
        end;
        Progress := 0;

        // Read data from source in 1Mb chunks.
        GetMem(Buffer, 1024*1024);
        try
          while (Size > 0) do
          begin
            Chunk := 1024*1024;
            if (Chunk > Size) then
              Chunk := Size;

            Stream.Read(Buffer, Chunk, @Chunk);

            if (Chunk = 0) then
              break;

            Inc(Progress, Chunk);

            if (DropEmptyTarget1.AsyncTransfer) then
              SendMessage(Handle, MSG_DROPPROGRESS, Progress, 0)
            else
            begin
              ProgressBar1.Position := Progress;
              ProgressBar1.Update;
            end;
            dec(Size, Chunk);
          end;
        finally
          FreeMem(Buffer);
        end;
      end;
    end;
  end;
  if (DropEmptyTarget1.AsyncTransfer) then
    SendMessage(Handle, MSG_DROPDONE, 0, 0)
  else
    StatusBar1.SimpleText := 'Drop completed.';
end;

procedure TFormMain.MsgDropEnd(var Msg: TMessage);
begin
  StatusBar1.SimpleText := 'Drop completed.';
end;

procedure TFormMain.MsgDropInit(var Msg: TMessage);
begin
  ProgressBar1.Max := Msg.WParam;
  ProgressBar1.Position := 0;
  ProgressBar1.Update;
end;

procedure TFormMain.MsgDropProgress(var Msg: TMessage);
begin
  ProgressBar1.Position := Msg.WParam;
  ProgressBar1.Update;
end;

procedure TFormMain.MsgDropStart(var Msg: TMessage);
begin
  StatusBar1.SimpleText := 'Drop completed - getting data...';
end;

procedure TFormMain.DropEmptyTarget1StartAsyncTransfer(Sender: TObject);
begin
  StatusBar1.SimpleText := 'Asynchronous transfer starting...';
end;

procedure TFormMain.DropEmptyTarget1EndAsyncTransfer(Sender: TObject);
begin
  StatusBar1.SimpleText := 'Asynchronous transfer ending...';
end;

end.

