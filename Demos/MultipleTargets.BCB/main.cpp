//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "main.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "DragDrop"
#pragma link "DragDropText"
#pragma link "DropSource"
#pragma link "DropTarget"
#pragma link "DropTarget"
#pragma resource "*.dfm"
TForm1 *Form1;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner): TForm(Owner)
{
}
//---------------------------------------------------------------------------
__fastcall TForm1::~TForm1()
{
  // Unregister all targets.
  // This is not strictly nescessary since the target component will perform
  // the unregistration automatically when it is destroyed. Feel free to skip
  // this step if you like.
  DropTextTarget1->Unregister(NULL);
}
//---------------------------------------------------------------------------
void __fastcall TForm1::FormMouseMove(TObject *Sender, TShiftState Shift,
  int X, int Y)
{

  for (int i = 0; i < ComponentCount; i++) {
    // Remove highlight from all TMemo controls.
    TMemo* memo = dynamic_cast<TMemo*>(Components[i]);
    if (memo)
      memo->Color = clWindow;
  }

  // Demo of TDropTarget->FindTarget:
  // Highlight the control under the cursor if it is a drop target.
  TWinControl* winctrl = dynamic_cast<TWinControl*>(Sender);
  TWinControl* Control = DropTextTarget1->FindTarget(winctrl->ClientToScreen(Point(X,Y)));
  TMemo* memo = dynamic_cast<TMemo*>(Control);

  if (memo)
    memo->Color = clLime;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::CheckBoxLeftClick(TObject *Sender)
{
  TCheckBox* checkbox = dynamic_cast<TCheckBox*>(Sender);
  // Register or unregister control as drop target according to users selection.
  if (checkbox->Checked)
    DropTextTarget1->Register(MemoLeft);
  else {
    // unregister & clear text
    MemoLeft->Lines->Clear();
    DropTextTarget1->Unregister(MemoLeft);
  }
}
//---------------------------------------------------------------------------
void __fastcall TForm1::CheckBoxRightClick(TObject *Sender)
{
  TCheckBox* checkbox = dynamic_cast<TCheckBox*>(Sender);
  // Register or unregister control as drop target according to users selection.
  if (checkbox->Checked)
    DropTextTarget1->Register(MemoRight);
  else {
    // unregister & clear text
    MemoRight->Lines->Clear();
    DropTextTarget1->Unregister(MemoRight);
  }
}
//---------------------------------------------------------------------------
void __fastcall TForm1::DropTextTarget1Drop(TObject *Sender,
  TShiftState ShiftState, TPoint &APoint, int &Effect)
{
  TDropTextTarget* dropTarget = dynamic_cast<TDropTextTarget*>(Sender);

  // Copy dragged text from target component into target control.
  TMemo* memo = dynamic_cast<TMemo*>(dropTarget->Target);

  memo->Lines->Add(dropTarget->Text);

  // Remove highlight.
  memo->Color = clWindow;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::DropTextTarget1Enter(TObject *Sender,
  TShiftState ShiftState, TPoint &APoint, int &Effect)
{
  TDropTextTarget* dropTarget = dynamic_cast<TDropTextTarget*>(Sender);

  // Copy dragged text from target component into target control.
  TMemo* memo = dynamic_cast<TMemo*>(dropTarget->Target);

  // Highlight the current drag target.
  // Use the TDropTarget->Target property to determine which control is
  // the current drop target:
  memo->Color = clRed;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::DropTextTarget1Leave(TObject *Sender)
{
  TDropTextTarget* dropTarget = dynamic_cast<TDropTextTarget*>(Sender);

  // Copy dragged text from target component into target control.
  TMemo* memo = dynamic_cast<TMemo*>(dropTarget->Target);

  // Remove highlight.
  memo->Color = clWindow;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::MemoSourceMouseDown(TObject *Sender,
      TMouseButton Button, TShiftState Shift, int X, int Y)
{
  // Wait for user to move cursor before we start the drag/drop.
  if (DragDetectPlus(reinterpret_cast<THandle>(Handle), Point(X,Y))) {
    DropTextSource1->Text = MemoSource->Lines->Text;
    DropTextSource1->Execute();
  }
}
//---------------------------------------------------------------------------

