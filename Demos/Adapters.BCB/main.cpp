//---------------------------------------------------------------------------
// This demo application was contributed by Jonathan Arnold.
//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "main.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "DragDrop"
#pragma link "DragDropText"
#pragma link "DropTarget"

// Note: In order to get the File and URL data format support linked into the
// application, we have to link in the appropiate units.
// If you forget to do this, you will get a run time error.
// The DragDropFile unit contains the TFileDataFormat class and the
// DragDropInternet unit contains the TURLDataFormat class.
#pragma link "DragDropFormats"
#pragma link "DragDropInternet"
#pragma link "DragDropFile"

#pragma link "DragDrop"
#pragma resource "*.dfm"
TFormMain *FormMain;
//---------------------------------------------------------------------------
__fastcall TFormMain::TFormMain(TComponent* Owner): TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::DropTextTarget1Drop(TObject *Sender,
  TShiftState ShiftState, TPoint &APoint, int &Effect)
{
  // Dropped a file onto the form
  MemoText->Lines->Text = DropTextTarget1->Text;

  // Check if we have a data format and if so...
  TFileDataFormat* fdf = dynamic_cast<TFileDataFormat*>(DataFormatAdapterFile->DataFormat);
  if (fdf) {
    // ...Extract the list of file(s)
    MemoFile->Lines = fdf->Files;
  }

  TURLDataFormat* udf = dynamic_cast<TURLDataFormat*>(DataFormatAdapterURL->DataFormat);
  if (udf) {
    // pick up the last URL
    MemoURL->Text = udf->URL;
  }
}
//---------------------------------------------------------------------------

