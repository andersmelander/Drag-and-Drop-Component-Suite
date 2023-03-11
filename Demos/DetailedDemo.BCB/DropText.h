//---------------------------------------------------------------------------

#ifndef DropTextH
#define DropTextH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include "DragDrop.hpp"
#include "DragDropText.hpp"
#include "DropSource.hpp"
#include "DropTarget.hpp"
#include <ComCtrls.hpp>
#include <ExtCtrls.hpp>
//---------------------------------------------------------------------------
class TFormText : public TForm
{
__published:	// IDE-managed Components
    TMemo *Memo1;
    TButton *ButtonClose;
    TEdit *Edit2;
    TStatusBar *StatusBar1;
    TMemo *Memo2;
    TEdit *Edit1;
    TButton *ButtonClipboard;
    TPanel *Panel1;
    TDropTextSource *DropSource1;
    TDropTextTarget *DropTextTarget1;
    TDropTextTarget *DropTextTarget2;
    TDropDummy *DropDummy1;
    void __fastcall DropSource1Feedback(TObject *Sender, int Effect,
          bool &UseDefaultCursors);
    void __fastcall DropTextTarget1Drop(TObject *Sender,
          TShiftState ShiftState, TPoint &APoint, int &Effect);
    void __fastcall Edit1MouseDown(TObject *Sender, TMouseButton Button,
          TShiftState Shift, int X, int Y);
    void __fastcall ButtonClipboardClick(TObject *Sender);
    void __fastcall Edit2MouseMove(TObject *Sender, TShiftState Shift,
          int X, int Y);
    void __fastcall ButtonCloseClick(TObject *Sender);
    void __fastcall DropTextTarget2Drop(TObject *Sender,
          TShiftState ShiftState, TPoint &APoint, int &Effect);
private:	// User declarations
	Controls::TWndMethod OldEdit2WindowProc;
	void __fastcall NewEdit2WindowProc(Messages::TMessage &Msg);
	bool __fastcall MouseIsOverEdit2Selection(int XPos);
	void __fastcall StartEdit2Drag(void);
public:		// User declarations
    __fastcall TFormText(TComponent* Owner);
	__fastcall ~TFormText();
};
//---------------------------------------------------------------------------
extern PACKAGE TFormText *FormText;
//---------------------------------------------------------------------------
#endif
