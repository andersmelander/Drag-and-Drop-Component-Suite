//---------------------------------------------------------------------------
// This demo application was contributed by Jonathan Arnold.
//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "DropText.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "DragDrop"
#pragma link "DragDropText"
#pragma link "DropSource"
#pragma link "DropTarget"
#pragma resource "*.dfm"
TFormText *FormText;

// Custom Cursors defined in Cursors.Res (included in DropFile.pas):
const int
  crTextCopy = 107,
  crTextMove = 108,
  crTextNoAccept = 109;

//---------------------------------------------------------------------------
__fastcall TFormText::TFormText(TComponent* Owner): TForm(Owner)
{
  // Used for Bottom Text Drag example...
  // Hook edit window so we can intercept WM_LBUTTONDOWN messages!
  OldEdit2WindowProc = Edit2->WindowProc;
  Edit2->WindowProc = NewEdit2WindowProc;

  // Load custom cursors.
  Screen->Cursors[crTextCopy] = ::LoadCursor(HInstance, "CUR_DRAG_COPY_TEXT");
  Screen->Cursors[crTextMove] = ::LoadCursor(HInstance, "CUR_DRAG_MOVE_TEXT");
  Screen->Cursors[crTextNoAccept] = ::LoadCursor(HInstance, "CUR_DRAG_NOACCEPT_TEXT");
}
//---------------------------------------------------------------------------
__fastcall TFormText::~TFormText()
{
Edit2->WindowProc = OldEdit2WindowProc;
}
//---------------------------------------------------------------------------
void __fastcall TFormText::DropSource1Feedback(TObject *Sender, int Effect,
  bool &UseDefaultCursors)
{
  // Provide custom drop source feedback.
  // Note: To use the standard drag/drop cursors, just disable this event
  // handler or set UseDefaultCursors to True.
  UseDefaultCursors = false; // We want to use our own cursors.

  // Ignore the drag scroll flag.
  unsigned int effect = Effect & ~DROPEFFECT_SCROLL;
  switch (effect) {
    case DROPEFFECT_COPY:
      ::SetCursor(Screen->Cursors[crTextCopy]);
      break;
    case DROPEFFECT_MOVE:
      ::SetCursor(Screen->Cursors[crTextMove]);
      break;
    default:
      ::SetCursor(Screen->Cursors[crTextNoAccept]);
      break;
  }
}
//---------------------------------------------------------------------------
void __fastcall TFormText::DropTextTarget1Drop(TObject *Sender,
  TShiftState ShiftState, TPoint &APoint, int &Effect)
{
  // Text has been dropped onto our drop target. Copy the dropped text into the
  // edit control.
  Edit1->Text = DropTextTarget1->Text;
}
//---------------------------------------------------------------------------
void __fastcall TFormText::Edit1MouseDown(TObject *Sender,
  TMouseButton Button, TShiftState Shift, int X, int Y)
{
  if (Edit1->Text == "")
    return;

  // Wait for user to move cursor before we start the drag/drop.
  if (DragDetectPlus(reinterpret_cast<THandle>(Handle), Point(X,Y))) {
    StatusBar1->SimpleText = "";
    Edit1->SelLength = 0;

    // Copy the data into the drop source->
    DropSource1->Text = Edit1->Text;

    // Temporarily disable Edit1 as a target so we can"t drop on the same
    // control as we are dragging from.
    DropTextTarget1->DragTypes = TDragTypes();
    TDragResult Res;
    try {
      // OK, now we are all set to go. Let"s start the drag from Edit1...
      Res = DropSource1->Execute();
    }
    __finally {
      // Enable Edit1 as a drop target again.
      DropTextTarget1->DragTypes = TDragTypes() << dtCopy;
    }

    // Display the result of the drag operation.
    switch (Res) {
      case drDropCopy:
        StatusBar1->SimpleText = "Copied successfully";
        break;
      case drDropLink:
        StatusBar1->SimpleText = "Scrap file created successfully";
        break;
      case drCancel:
        StatusBar1->SimpleText = "Drop cancelled";
        break;
      case drOutMemory:
        StatusBar1->SimpleText = "Drop failed - out of memory";
        break;
      default:
        StatusBar1->SimpleText = "Drop failed - reason unknown";
    }
  }
}
//---------------------------------------------------------------------------
void __fastcall TFormText::ButtonClipboardClick(TObject *Sender)
{
  if (Edit1->Text != "") { // Copy data into drop source component and then...
    DropSource1->Text = Edit1->Text;

    // ...Copy the data to the clipboard.
    DropSource1->CopyToClipboard();

    StatusBar1->SimpleText = "Text copied to clipboard.";
  }
}
//---------------------------------------------------------------------------
void __fastcall TFormText::Edit2MouseMove(TObject *Sender,
  TShiftState Shift, int X, int Y)
{
  // This method just changes mouse cursor to crHandPoint if over selected text.
  if (MouseIsOverEdit2Selection(X))
    Edit2->Cursor = crHandPoint;
  else
    Edit2->Cursor = crDefault;
}
//---------------------------------------------------------------------------
void __fastcall TFormText::ButtonCloseClick(TObject *Sender)
{
  Close();
}
//----------------------------------------------------------------------------

void __fastcall TFormText::DropTextTarget2Drop(TObject *Sender,
  TShiftState ShiftState, TPoint &APoint, int &Effect)
{
  // Text has been dropped onto our drop target. Copy the dropped text into the
  // edit control.
  Edit2->Text = DropTextTarget2->Text;
}
//---------------------------------------------------------------------------
// The following methods are used for the BOTTOM Text Drop SOURCE and TARGET examples.
// The DropSource code is almost identical. However, the Edit2 control
// has been hooked to override the default WM_LBUTTONDOWN message handling.
// Using the normal OnMouseDown event method does not work for this example.
//----------------------------------------------------------------------------

// The new WindowProc for Edit2 which intercepts WM_LBUTTONDOWN messages
// before ANY OTHER processing...
void __fastcall TFormText::NewEdit2WindowProc(Messages::TMessage &Msg)
{
  TWMMouse* mouse_msg = reinterpret_cast<TWMMouse*>(&Msg);

  if (mouse_msg->Msg == WM_LBUTTONDOWN &&
    MouseIsOverEdit2Selection(mouse_msg->XPos)) {
    StartEdit2Drag(); // Just a private method
    Msg.Result = 0;
  }
  else
  { //Otherwise do everything as before...
    OldEdit2WindowProc(Msg);
  }
}
//---------------------------------------------------------------------------
bool __fastcall TFormText::MouseIsOverEdit2Selection(int XPos)
{
  if ((Edit2->SelLength > 0) && (Edit2->Focused())) {
    TSize size1, size2;

    // Create a Device Context which can be used to retrieve font metrics.
    HDC dc = GetDC(0);
    try { // Select the edit control's font into our DC.
      HFONT SavedFont = ::SelectObject(dc, Edit2->Font->Handle);
      try { // Get text before selection.
        AnsiString s1 = Edit2->Text.SubString(1, Edit2->SelStart);

        // Get text up to and including selection.
        AnsiString s2 = s1 + Edit2->SelText;

        // Get dimensions of text before selection and up to and including
        // selection.
        ::GetTextExtentPoint32(dc, s1.c_str(), s1.Length(), &size1);
	::GetTextExtentPoint32(dc, s2.c_str(), s2.Length(), &size2);
      }
      __finally {
        ::SelectObject(dc, SavedFont);
      }
    }
    __finally {
      ReleaseDC(0,dc);
    }

    return (XPos >= size1.cx) && (XPos <= size2.cx);
  }
  else
    return false;
}
//---------------------------------------------------------------------------
void __fastcall TFormText::StartEdit2Drag(void)
{
  StatusBar1->SimpleText = "";

  // Copy the data into the drop source and...
  DropSource1->Text = Edit2->SelText;

  // ..->Start the drag/drop.
  TDragResult Res = DropSource1->Execute();

  // Display the result of the drag operation.
  switch (Res) {
    case drDropCopy:
      StatusBar1->SimpleText = "Copied successfully";
      break;
    case drDropLink:
      StatusBar1->SimpleText = "Scrap file created successfully";
      break;
    case drCancel:
      StatusBar1->SimpleText = "Drop cancelled";
      break;
    case drOutMemory:
      StatusBar1->SimpleText = "Drop failed - out of memory";
      break;
    default:
      StatusBar1->SimpleText = "Drop failed - reason unknown";
      break;
  }
}
//---------------------------------------------------------------------------

