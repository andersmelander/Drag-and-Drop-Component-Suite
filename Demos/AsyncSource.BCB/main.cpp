//---------------------------------------------------------------------------
// This demo application was contributed by Jonathan Arnold.
//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include <math.h>

#include "main.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "DragDrop"
#pragma link "DropSource"
#pragma link "DragDrop"
#pragma resource "*.dfm"
TFormMain *FormMain;

static const TestFileSize = 1024*1024*10; // 10Mb

//---------------------------------------------------------------------------
__fastcall TFormMain::TFormMain(TComponent* Owner): TForm(Owner)
{
  // Setup event handler to let a drop target request data from our drop source.
  TVirtualFileStreamDataFormat* tvfsdf =
    dynamic_cast<TVirtualFileStreamDataFormat*>(DataFormatAdapterSource->DataFormat);
  tvfsdf->OnGetStream = OnGetStream;
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::FormDestroy(TObject *Sender)
{
  Timer1->Enabled = false;
}
//---------------------------------------------------------------------------

void __fastcall TFormMain::DropEmptySource1AfterDrop(TObject *Sender,
  TDragResult DragResult, bool Optimized)
{
  StatusBar1->SimpleText = "Target processing drop";
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::DropEmptySource1Drop(TObject *Sender,
      TDragType DragType, bool &ContinueDrop)
{
  StatusBar1->SimpleText = "Drop completed";
  ButtonAbort->Visible = false;
}
//---------------------------------------------------------------------------
TFakeStream::TFakeStream(int _Size, int _MaxCount ) :
  TStream(), FSize(_Size), FMaxCount(_MaxCount )
{
}
//---------------------------------------------------------------------------

int __fastcall TFakeStream::Read( void* Buffer, int Count )
{
  int amount_read = 0;

  if (!FAbort && (FPosition >= 0) && (Count >= 0)) {

    amount_read = FSize - FPosition;

    if (amount_read > 0) {

      if (amount_read > Count)
        amount_read = Count;

      if (amount_read > FMaxCount)
        amount_read = FMaxCount;

      memset(Buffer, 'X', amount_read);

      FPosition += amount_read;

      if (FProgress)
        FProgress(this, FPosition, FSize);
    }
  }

  return amount_read;
}

int __fastcall TFakeStream::Seek(int Offset, Word Origin)
{
  switch(Origin) {
    case soFromBeginning:
      FPosition = Offset;
      break;
    case soFromCurrent:
      FPosition += Offset;
      break;
    case soFromEnd:
      FPosition = FSize + Offset;
      break;
  }

  if (FProgress)
    FProgress(this, FPosition, FMaxCount);

  return FPosition;
}

#if defined(RTLVersion)
  #if (RTLVersion < 14)
__int64 __fastcall TFakeStream::Seek(const __int64 Offset, TSeekOrigin Origin)
{
  return Seek((int)Offset, (Word)Origin);
}
  #endif
#endif"

void __fastcall TFakeStream::SetSize(int NewSize)
{
}

void __fastcall TFakeStream::SetSize(const __int64 NewSize)
{
}

int __fastcall TFakeStream::Write(const void* Buffer, int Count)
{
  return 0;
}

void __fastcall TFormMain::OnGetStream(TFileContentsStreamOnDemandClipboardFormat* Sender,
  int Index, _di_IStream& AStream)
{
  // Note: This method might be called in the context of the transfer thread.
  // See TFormMain.OnProgress for a comment on this.

  // This event handler is called by TFileContentsStreamOnDemandClipboardFormat
  // when the drop target requests data from the drop source (that's us).
  StatusBar1->SimpleText = "Transfering data";

  // Create a stream which contains the data we will transfer...
  // In this case we just create a dummy stream which contains 10Mb of 'X'
  // characters. In order to provide smooth feedback through the progress bar
  // (and slow the transfer down a bit) the stream will only transfer up to 32K
  // at a time - Each time TStream.Read is called the progress bar is updated
  // via the stream's progress event.
  TFakeStream* fStream = new TFakeStream(TestFileSize, 32*1024);
  try {
    fStream->OnProgress = OnProgress;
    // ...and return the stream back to the target as an IStream. Note that the
    // target is responsible for deleting the stream (via reference counting).
    AStream = *new TFixedStreamAdapter(fStream, soOwned);
  }
  catch (...) {
    delete fStream;
    throw;
  }

  ProgressBar1->Position = 0;
}

void __fastcall TFormMain::OnProgress(TObject* Sender, int Count, int MaxCount)
{
  // Note that during an asyncronous transfer, the progress event handler is
  // being called in the context of the transfer thread. This means that this
  // event handler should abide to all the normal thread safety rules (i.e.
  // don't call GDI or mess with non-thread safe objects).

  // Update progress bar to show how much data has been transfered so far.
  // This isn't really thread safe since it modifies the form, but so far it
  // hasn't crashed on me.
  ProgressBar1->Max = MaxCount;
  ProgressBar1->Position = Count;
  if (DoAbort)
    dynamic_cast<TFakeStream*>(Sender)->Abort();
}

void __fastcall TFormMain::ButtonAbortClick(TObject *Sender)
{
  DoAbort = true;
}
//---------------------------------------------------------------------------

void __fastcall TFormMain::OnMouseDown(TObject *Sender,
      TMouseButton Button, TShiftState Shift, int X, int Y)
{
  StatusBar1->SimpleText = "";
  if (DragDetectPlus(reinterpret_cast<THandle>(Handle), Point(X, Y))) {
    StatusBar1->SimpleText = "Dragging data";

    // Transfer the file names to the data format. The content will be extracted
    // by the target on-demand.
    TVirtualFileStreamDataFormat* tvfsdf =
      dynamic_cast<TVirtualFileStreamDataFormat*>(DataFormatAdapterSource->DataFormat);
    tvfsdf->FileNames->Clear();
    tvfsdf->FileNames->Add("big text file.txt");

    // Set the size and timestamp attributes of the filename we just added.
    PFileDescriptor pfile = reinterpret_cast<PFileDescriptor>(tvfsdf->FileDescriptors->Items[0]);
    GetSystemTimeAsFileTime(&pfile->ftLastWriteTime);
    pfile->nFileSizeLow = TestFileSize;
    pfile->nFileSizeHigh = 0; // I assume the test file doesn't grow beyond 4Gb...
    pfile->dwFlags = FD_WRITESTIME | FD_FILESIZE | ::FD_PROGRESSUI;

    // Determine if we should perform an async drag or a normal drag.
    if (RadioButtonAsync.Checked) {
      DoAbort = false;
      ButtonAbort->Visible := true;

      // Create a thread to perform the drag...
      if (DropEmptySource1->Execute(true) == drAsync) {
        StatusBar1->SimpleText = "Asynchronous drag in progress...";
      }
      else
      {
        StatusBar1->SimpleText = "Asynchronous drag failed";
      }
    }
    else
    {
      // Perform a normal drag (in the main thread).
      DropEmptySource1->Execute;

      StatusBar1->SimpleText = "Drop completed";
    }
  }
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::Timer1Timer(TObject *Sender)
{
  // Update the pie to indicate that the application is responding to
  // messages (i.e. isn't blocked).
  Tick = (Tick + 10) % 100;
  if (Tick == 0)
    EvenOdd = !EvenOdd;

  // Draw an animated pie chart to show that application is responsive to events.
  DrawPie(Tick);
}
//---------------------------------------------------------------------------
void TFormMain::DrawPie(int Percent)
{
  // Assume paintbox width is smaller than height.
  int Radius = (PaintBoxPie->Width / 2) - 10;
  TPoint Center(PaintBoxPie->Width / 2, PaintBoxPie->Height / 2);
  double v = Percent * M_PI / 50; // Convert percent to radians.
  TPoint Radial;
  Radial.x = Center.x+int(Radius * cos(v));
  Radial.y = Center.y-int(Radius * sin(v));

  PaintBoxPie->Canvas->Brush->Style = bsSolid;
  PaintBoxPie->Canvas->Pen->Color = clGray;
  PaintBoxPie->Canvas->Pen->Style = psSolid;

  if (EvenOdd)
    PaintBoxPie->Canvas->Brush->Color = clRed;
  else
    PaintBoxPie->Canvas->Brush->Color = Color;
  PaintBoxPie->Canvas->Pie(Center.x-Radius, Center.y-Radius,
    Center.x+Radius, Center.y+Radius,
    Radial.x, Radial.y,
    Center.x+Radius, Center.y);

  if (Percent != 0) {
    if (!EvenOdd)
      PaintBoxPie->Canvas->Brush->Color = clRed;
    else
      PaintBoxPie->Canvas->Brush->Color = Color;
    PaintBoxPie->Canvas->Pie(Center.x-Radius, Center.y-Radius,
    Center.x+Radius, Center.y+Radius,
    Center.x+Radius, Center.y,
    Radial.x, Radial.y);
  }
}


