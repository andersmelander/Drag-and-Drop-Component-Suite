//---------------------------------------------------------------------------
// This demo application was contributed by Jonathan Arnold.
//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Target.h"
#include "DragDropTimeOfDay.h"

//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "DragDrop"
#pragma link "DragDropText"
#pragma link "DropTarget"
#pragma resource "*.dfm"
TFormTarget *FormTarget;
//---------------------------------------------------------------------------
__fastcall TFormTarget::TFormTarget(TComponent* Owner): TForm(Owner)
{
  // Define and register our custom clipboard format.
  // This needs to be done for both the drop source and target.
  TimeDataFormatTarget = new TGenericDataFormat(DropTextTarget1);
  TimeDataFormatTarget->AddFormat(sTimeOfDayName);
}
//---------------------------------------------------------------------------
void __fastcall TFormTarget::DropTextTarget1Drop(TObject *Sender,
  TShiftState ShiftState, TPoint &APoint, int &Effect)
{
  // Determine if we got our custom format.
  if (TimeDataFormatTarget->HasData()) {

    // Extract the dropped data into our custom struct.
    TTimeOfDay TOD;
    TimeDataFormatTarget->GetDataHere(&TOD, sizeof(TOD));

    // Convert the time-of-day info to a TDateTime so we can display it.
    TDateTime Time = EncodeTime(TOD.hours, TOD.minutes, TOD.seconds, TOD.milliseconds);

    // Display the data.
    PanelDest->Caption = FormatDateTime("hh:nn:ss.zzz", Time);
    PanelDest->Color = TOD.color;
    PanelDest->Font->Color = TColor(!(PanelDest->Color) & 0xFFFFFF);
  }
  else
    PanelDest->Caption = static_cast<TDropTextTarget*>(Sender)->Text;
}
//---------------------------------------------------------------------------
