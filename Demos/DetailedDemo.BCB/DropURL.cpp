//---------------------------------------------------------------------------
// This demo application was contributed by Jonathan Arnold.
//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "DropURL.h"
#include <comobj.hpp> // For OleCheck()
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "DragDrop"
#pragma link "DragDropGraphics"
#pragma link "DragDropInternet"
#pragma link "DropSource"
#pragma link "DropTarget"
#pragma resource "*.dfm"
TFormURL *FormURL;
//---------------------------------------------------------------------------
__fastcall TFormURL::TFormURL(TComponent* Owner): TForm(Owner)
{
  // Note: This is an example of "manual" target registration. We could just
  // as well have assigned the TDropTarget.Target property at design-time to
  // register the drop targets.

  // Register the URL and BMP DropTarget controls.
  DropURLTarget1->Register(PanelURL);
  DropBMPTarget1->Register(PanelImageTarget);

  // This enables the dragged image to be visible
  // over the whole form, not just the above targets.
  DropDummy1->Register(this);
}
//---------------------------------------------------------------------------
__fastcall TFormURL::~TFormURL()
{
  // Note: This is an example of "manual" target unregistration. However,
  // Since the targets are automatically unregistered when they are destroyed,
  // it is not nescessary to do it manually.

  // UnRegister the DropTarget windows.
  DropURLTarget1->Unregister(PanelURL);
  DropBMPTarget1->Unregister(PanelImageTarget);
  DropDummy1->Unregister();
}
//------------------------------------------------------------------------------
// URL stuff ...
//------------------------------------------------------------------------------

//---------------------------------------------------------------------------
void __fastcall TFormURL::URLMouseDown(TObject *Sender,
  TMouseButton Button, TShiftState Shift, int X, int Y)
{
  // Wait for user to move cursor before we start the drag/drop.
  if (DragDetectPlus(reinterpret_cast<THandle>(Handle), Point(X,Y))) {
  	
    TWinControl* wc = static_cast<TWinControl*>(Sender);

    // This demonstrates how to create a drag image based on the source control.
    // Note: DropURLSource1->Images = ImageList1
    Graphics::TBitmap* DragImage = new Graphics::TBitmap();
    try {
      DragImage->Width = wc->Width;
      DragImage->Height = wc->Height;
      wc->PaintTo(DragImage->Canvas->Handle, 0, 0);
      ImageList1->Width = DragImage->Width;
      ImageList1->Height = DragImage->Height;
      ImageList1->Add(DragImage, 0);
    }
    __finally {
      delete DragImage;
    }

    DropURLSource1->ImageHotSpotX = X;
    DropURLSource1->ImageHotSpotY = Y;
    DropURLSource1->ImageIndex = 0;

    try {
      // Copy the data into the drop source.
      DropURLSource1->Title = PanelURL->Caption;
      DropURLSource1->URL = LabelURL->Caption;

      // Temporarily disable Edit1 as a drop target.
      DropURLTarget1->DragTypes = TDragTypes();
      try {
        DropURLSource1->Execute();
      }
      __finally {
        // Enable Edit1 as a drop target again.
        DropURLTarget1->DragTypes = TDragTypes() << dtLink;
      }
    }
    __finally {
      // Now that the drag has completed we don't need the image list anymore.
      ImageList1->Clear();
    }
  }
}

//---------------------------------------------------------------------------
void __fastcall TFormURL::DropURLTarget1Drop(TObject *Sender,
  TShiftState ShiftState, TPoint &APoint, int &Effect)
{
  // An URL has been dropped - Copy the URL and title from the drop target.
  PanelURL->Caption = DropURLTarget1->Title;
  LabelURL->Caption = DropURLTarget1->URL;
}

//------------------------------------------------------------------------------
// Bitmap stuff ...
//------------------------------------------------------------------------------

TColor GetTransparentColor(Graphics::TBitmap* bmp)
{
  if ((!bmp) || (bmp->Empty)) // Warning: depends on order of evaluation!
    return clWhite;
  else
    return bmp->TransparentColor;
}

//---------------------------------------------------------------------------
void __fastcall TFormURL::ImageMouseDown(TObject *Sender,
  TMouseButton Button, TShiftState Shift, int X, int Y)
{
  TImage* ic = dynamic_cast<TImage*>(Sender);

  if (!ic || (Button == mbRight) || !ic->Picture->Graphic)
    return;

  // Since the TImage component hasn't got a window handle, we must
  // use the TPanel behind it instead...
  // First convert the mouse coordinates from Image relative coordinates
  // to screen coordinates...
  TPoint p(X,Y);
  p = ic->ClientToScreen(p);
  // ..->and then back to to TPanel relative ones.
  p = ic->Parent->ScreenToClient(p);

  // Now that the coordinates are relative to the panel, we can use
  // the panel's window handle for DragDetectPlus:
  if (DragDetectPlus(reinterpret_cast<THandle>(ic->Parent->Handle), Point(X,Y))) {
    // Freeze clipboard contents if we have live data on it.
    // This is only nescessary because the TGraphic based data formats (such as
    // TBitmapDataFormat) doesn't support auto flushing.
    if (DropBMPSource1->LiveDataOnClipboard)
      DropBMPSource1->FlushClipboard();

    // Copy the data into the drop source.
    DropBMPSource1->Bitmap->Assign(ic->Picture->Graphic);

    // This just demonstrates dynamically allocating a drag image.
    // Note: DropBMPSource1->Images = ImageList1
    ImageList1->Width = DropBMPSource1->Bitmap->Width;
    ImageList1->Height = DropBMPSource1->Bitmap->Height;
    ImageList1->AddMasked(DropBMPSource1->Bitmap,
    GetTransparentColor(DropBMPSource1->Bitmap));
    DropBMPSource1->ImageIndex = 0;
    try {
      // Perform the drag.
      if (DropBMPSource1->Execute() == drDropMove) {
        // Clear the source image if image were drag-moved.
        ic->Picture->Graphic = NULL;
      }
    }
    __finally {
      // Now that the drag has completed we don't need the image list anymore.
      ImageList1->Clear();
    }
  }
}
//---------------------------------------------------------------------------
void __fastcall TFormURL::PopupMenu1Popup(TObject *Sender)
{
  TPopupMenu* popmenu = dynamic_cast<TPopupMenu*>(Sender);
  if (!popmenu)
    return;

  TComponent* PopupSource = popmenu->PopupComponent;

  // Enable cut and copy for source image unless it is empty.
  if ((PopupSource == ImageSource1) || (PopupSource == ImageSource2)) {
    MenuCopy->Enabled = (static_cast<TImage*>(PopupSource))->Picture->Graphic != NULL;
    MenuCut->Enabled = MenuCopy->Enabled;
    MenuPaste->Enabled = false;
  }
  else
  {
    // Enable paste for target image if the clipboard contains a bitmap.
    if ((PopupSource == ImageTarget)) {
      MenuCopy->Enabled = False;
      MenuCut->Enabled = False;
      // Open the clipboard as an IDataObject
      IDataObject* DataObject;
      OleCheck(OleGetClipboard(&DataObject));
      try {
        // Enable paste menu if the clipboard contains data in any of
        // the supported formats.
        MenuPaste->Enabled = DropBMPTarget1->HasValidFormats(DataObject);
      }
      __finally {
        DataObject = NULL;
      }
    }
  }
}
//---------------------------------------------------------------------------
void __fastcall TFormURL::DropBMPTarget1Drop(TObject *Sender,
  TShiftState ShiftState, TPoint &APoint, int &Effect)
{
  // An image has just been dropped on the target - Copy it to
  // our TImage component
  ImageTarget->Picture->Assign(DropBMPTarget1->Bitmap);
}
//---------------------------------------------------------------------------
void __fastcall TFormURL::DropBMPSource1Paste(TObject *Sender,
  TDragResult Action, bool DeleteOnPaste)
{
  // When the target signals that it has pasted the image (after a
  // CutToClipboard operation), we can safely delete the source image.
  // This is an example of a "Delete on paste" operation.
  if (DeleteOnPaste)
    PasteImage->Picture->Assign(NULL);
  StatusBar1->SimpleText = "Bitmap pasted from clipboard";
}
//---------------------------------------------------------------------------
void __fastcall TFormURL::MenuCopyOrCutClick(TObject *Sender)
{
  // Clear the current content of the clipboard.
  // This isn"t strictly nescessary, but can improve performance; If the drop
  // source has live data on the clipboard and the drop source data is modified,
  // the drop source will copy all its current data to the clipboard and then
  // disconnect itself from it. It does this in order to preserve the clipboard
  // data in the state it was in when the data were copied to the clipboard.
  // Since we are about to copy new data to the clipboard, we might as well save
  // the drop source all this unnescessary work.
  DropBMPSource1->EmptyClipboard();

  // Remember which TImage component the data originated from. This is used so
  // we can clear the image if a "delete on paste" is performed.
  PasteImage = (TImage*)((TPopupMenu*)((TMenuItem*)(Sender))->GetParentMenu())->PopupComponent;

  DropBMPSource1->Bitmap->Assign(PasteImage->Picture->Graphic);

  AnsiString op;
  if (Sender == MenuCut) {
    DropBMPSource1->CutToClipboard();
    op = "cut";
  }
  else
  {
    DropBMPSource1->CopyToClipboard();
    op = "copied";
  }
  StatusBar1->SimpleText = "Bitmap " + op + " to clipboard";
}
//---------------------------------------------------------------------------

void __fastcall TFormURL::MenuPasteClick(TObject *Sender)
{
  // PasteFromClipboard fires an OnDrop event, so we don't need to do
  // anything special here.
  DropBMPTarget1->PasteFromClipboard();
  StatusBar1->SimpleText = "Bitmap pasted from clipboard";
}
//---------------------------------------------------------------------------

void __fastcall TFormURL::ButtonCloseClick(TObject *Sender)
{
  Close();
}
//---------------------------------------------------------------------------

