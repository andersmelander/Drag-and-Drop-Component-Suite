//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Source.h"
#include "DragDropTimeOfDay.h"

//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "DragDrop"
#pragma link "DragDropText"
#pragma link "DropSource"
#pragma link "DragDrop"
#pragma resource "*.dfm"
TFormSource *FormSource;
//---------------------------------------------------------------------------
__fastcall TFormSource::TFormSource(TComponent* Owner): TForm(Owner)
{
  // Define and register our custom clipboard format.
  // This needs to be done for both the drop source and target.
  TimeDataFormatSource = new TGenericDataFormat(DropTextSource1);
  TimeDataFormatSource->AddFormat(sTimeOfDayName);

  randomize();
}
//---------------------------------------------------------------------------
void __fastcall TFormSource::PanelSourceMouseDown(TObject *Sender,
      TMouseButton Button, TShiftState Shift, int X, int Y)
{
  Timer1->Enabled = false;
  try {
    if (DragDetectPlus(reinterpret_cast<THandle>(Handle), Point(X, Y))) {
      // Transfer time as text. This is not nescessary and is only done to offer
      // maximum flexibility in case the user wishes to drag our data to some
      // other application (e->g-> a word processor).
      DropTextSource1->Text = PanelSource->Caption;

      // Store the current time in a structure. This structure is our custom
      // data format.
      TTimeOfDay TOD;
      DecodeTime(Now(), TOD.hours, TOD.minutes, TOD.seconds, TOD.milliseconds);
      TOD.color = PanelSource->Color;

      // Transfer the structure to the drop source data object and execute the drag->
      TimeDataFormatSource->SetDataHere(&TOD, sizeof(TOD));

      DropTextSource1->Execute();
    }
  }
  __finally {
    Timer1->Enabled = true;
  }
}
//---------------------------------------------------------------------------
void __fastcall TFormSource::Timer1Timer(TObject *Sender)
{
  PanelSource->Caption = FormatDateTime("hh:nn:ss.zzz", Now());
  PanelSource->Color = static_cast<TColor>(random(0xFFFFFF));
  PanelSource->Font->Color = static_cast<TColor>(!(PanelSource->Color) & 0xFFFFFF);
}
//---------------------------------------------------------------------------
