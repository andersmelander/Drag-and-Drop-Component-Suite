//---------------------------------------------------------------------------

#ifndef mainH
#define mainH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include "DragDrop.hpp"
#include "DragDropText.hpp"
#include "DropSource.hpp"
#include "DropTarget.hpp"
//---------------------------------------------------------------------------
class TForm1 : public TForm
{
__published:	// IDE-managed Components
    TLabel *Label1;
    TMemo *MemoLeft;
    TMemo *MemoRight;
    TCheckBox *CheckBoxLeft;
    TCheckBox *CheckBoxRight;
    TMemo *MemoSource;
    TDropTextTarget *DropTextTarget1;
    TDropTextSource *DropTextSource1;
    TDropDummy *DropDummy1;
    void __fastcall FormMouseMove(TObject *Sender, TShiftState Shift,
          int X, int Y);
    void __fastcall CheckBoxLeftClick(TObject *Sender);
    void __fastcall CheckBoxRightClick(TObject *Sender);
    void __fastcall DropTextTarget1Drop(TObject *Sender,
          TShiftState ShiftState, TPoint &APoint, int &Effect);
    void __fastcall DropTextTarget1Enter(TObject *Sender,
          TShiftState ShiftState, TPoint &APoint, int &Effect);
    void __fastcall DropTextTarget1Leave(TObject *Sender);
    void __fastcall MemoSourceMouseDown(TObject *Sender,
          TMouseButton Button, TShiftState Shift, int X, int Y);
private:	// User declarations
public:		// User declarations
    __fastcall TForm1(TComponent* Owner);
	__fastcall ~TForm1();
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif
