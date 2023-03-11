//---------------------------------------------------------------------------

#ifndef DropURLH
#define DropURLH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include "DragDrop.hpp"
#include "DragDropGraphics.hpp"
#include "DragDropInternet.hpp"
#include "DropSource.hpp"
#include "DropTarget.hpp"
#include <ComCtrls.hpp>
#include <ExtCtrls.hpp>
#include <Graphics.hpp>
#include <ImgList.hpp>
#include <Menus.hpp>
//---------------------------------------------------------------------------
class TFormURL : public TForm
{
__published:	// IDE-managed Components
    TLabel *LabelURL;
    TPanel *Panel1;
    TButton *ButtonClose;
    TStatusBar *StatusBar1;
    TMemo *Memo2;
    TPanel *PanelImageTarget;
    TImage *ImageTarget;
    TPanel *Panel3;
    TMemo *Memo1;
    TPanel *PanelImageSource2;
    TImage *ImageSource2;
    TPanel *PanelImageSource1;
    TImage *ImageSource1;
    TPanel *PanelURL;
    TDropURLTarget *DropURLTarget1;
    TDropURLSource *DropURLSource1;
    TDropBMPSource *DropBMPSource1;
    TDropBMPTarget *DropBMPTarget1;
    TImageList *ImageList1;
    TDropDummy *DropDummy1;
    TPopupMenu *PopupMenu1;
    TMenuItem *MenuCopy;
    TMenuItem *MenuCut;
    TMenuItem *N1;
    TMenuItem *MenuPaste;
    void __fastcall URLMouseDown(TObject *Sender, TMouseButton Button,
          TShiftState Shift, int X, int Y);
    void __fastcall DropURLTarget1Drop(TObject *Sender,
          TShiftState ShiftState, TPoint &APoint, int &Effect);
    void __fastcall ImageMouseDown(TObject *Sender,
          TMouseButton Button, TShiftState Shift, int X, int Y);
    void __fastcall DropBMPSource1Paste(TObject *Sender,
          TDragResult Action, bool DeleteOnPaste);
    void __fastcall PopupMenu1Popup(TObject *Sender);
    void __fastcall DropBMPTarget1Drop(TObject *Sender,
          TShiftState ShiftState, TPoint &APoint, int &Effect);
    void __fastcall MenuCopyOrCutClick(TObject *Sender);
    void __fastcall MenuPasteClick(TObject *Sender);
    void __fastcall ButtonCloseClick(TObject *Sender);
private:	// User declarations
    TImage* PasteImage; // Remembers which TImage is the source of a copy/paste
public:		// User declarations
    __fastcall TFormURL(TComponent* Owner);
    __fastcall ~TFormURL();
};
//---------------------------------------------------------------------------
extern PACKAGE TFormURL *FormURL;
//---------------------------------------------------------------------------
#endif
