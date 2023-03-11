//---------------------------------------------------------------------------

#ifndef MainH
#define MainH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include "DragDrop.hpp"
#include "DragDropFile.hpp"
#include "DropSource.hpp"
#include "DropTarget.hpp"
#include <ComCtrls.hpp>
#include <ExtCtrls.hpp>
#include <ImgList.hpp>
//---------------------------------------------------------------------------
class TFormMain : public TForm
{
__published:	// IDE-managed Components
        TDropFileTarget *DropFileTarget1;
        TDropFileSource *DropFileSource1;
        TListView *ListView1;
        TPanel *Panel1;
        TLabel *Label1;
        TImageList *ImageList1;
        TImageList *ImageList2;
        void __fastcall DropFileTarget1Drop(TObject *Sender,
          TShiftState ShiftState, TPoint &APoint, int &Effect);
        void __fastcall ListView1MouseDown(TObject *Sender,
          TMouseButton Button, TShiftState Shift, int X, int Y);
private:	// User declarations
public:		// User declarations
        __fastcall TFormMain(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TFormMain *FormMain;
//---------------------------------------------------------------------------
#endif
