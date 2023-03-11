//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Main.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "DragDrop"
#pragma link "DragDropFile"
#pragma link "DropSource"
#pragma link "DropTarget"
#pragma resource "*.dfm"
TFormMain *FormMain;
//---------------------------------------------------------------------------
__fastcall TFormMain::TFormMain(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::DropFileTarget1Drop(TObject *Sender,
      TShiftState ShiftState, TPoint &APoint, int &Effect)
{
  ListView1->Items->Clear();
  int i;
  for (i = 0; i < DropFileTarget1->Files->Count; i++)
    ListView1->Items->Add()->Caption = DropFileTarget1->Files->Strings[i];
}
//---------------------------------------------------------------------------

void __fastcall TFormMain::ListView1MouseDown(TObject *Sender,
      TMouseButton Button, TShiftState Shift, int X, int Y)
{
  if (DragDetectPlus(THandle(ListView1->Handle), Point(X,Y)) &&
    (ListView1->SelCount > 0)) {
    DropFileSource1->Files->Clear();
    int i;
    DropFileSource1->ShowImage = false;
    for (i = 0; i < ListView1->Items->Count; i++)
      if (ListView1->Items->Item[i]->Selected) {
        DropFileSource1->Files->Add(ListView1->Items->Item[i]->Caption);
        // If a single item is being dragged we can use the list view to create
        // a nice drag image.
        if (ListView1->SelCount == 1) {
          TPoint p;
          ImageList1->Handle = (int)ListView_CreateDragImage(ListView1->Handle,
            i, &p);
          DropFileSource1->ShowImage = true;
          DropFileSource1->ImageHotSpotX = X-ListView1->Items->Item[i]->Left;
          DropFileSource1->ImageHotSpotY = Y-ListView1->Items->Item[i]->Top;
        }
      }

    DropFileSource1->Execute();
  }
}
//---------------------------------------------------------------------------

